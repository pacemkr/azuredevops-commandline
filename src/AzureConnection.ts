import * as winston from 'winston';
import * as _ from 'lodash';
import { differenceInMilliseconds, parseISO, parse } from "date-fns";

import * as azdev from "azure-devops-node-api";
import * as lim from "azure-devops-node-api/interfaces/LocationsInterfaces";
import * as CoreApi from 'azure-devops-node-api/CoreApi';
import * as CoreInterfaces from 'azure-devops-node-api/interfaces/CoreInterfaces';
import * as WorkApi from 'azure-devops-node-api/WorkApi';
import * as witApi from 'azure-devops-node-api/WorkItemTrackingApi';
import * as witInterfaces from 'azure-devops-node-api/interfaces/WorkItemTrackingInterfaces';

import { Configuration } from "./Configuration"
import { WorkItem } from "./WorkItem";
import { BoardColumn } from 'azure-devops-node-api/interfaces/WorkInterfaces';
import { isNull } from 'lodash';

export class AzureConnection {

    constructor(public Project:string,
                public TeamName:string, 
                public BoardName:string) {        
        this.pbis = new Array<WorkItem>(0);
    }
 
    async connect(): Promise<void> {

        let orgUrl = Configuration.getInstance().Url;

        let token: string = Configuration.getInstance().Token;

        let authHandler = azdev.getPersonalAccessTokenHandler(token);
        this.azureConnection = new azdev.WebApi(orgUrl, authHandler);
        let connData: lim.ConnectionData = await this.azureConnection.connect();
        winston.info(`Hello ${connData.authenticatedUser.providerDisplayName}`);
        winston.info(`Connected to ${Configuration.getInstance().Url}`)
    }

    public async fetchPbis(query):Promise<void>{

        winston.info(`Calling query ${query}`);
        const witApi: witApi.IWorkItemTrackingApi = await this.azureConnection.getWorkItemTrackingApi();

        const azureQuery: witInterfaces.QueryHierarchyItem = await witApi.getQuery(this.teamContext.project, query);
        if (_.isNull(azureQuery)){
            winston.error(`Query '${azureQuery}' not found.`);
            throw new Error(`Query '${azureQuery}' not found.`);
        }

        let results: witInterfaces.WorkItemQueryResult = await witApi.queryById(azureQuery.id, this.teamContext, false);        
        let workItems = await this.extractWorkItems(results.workItems, witApi);

        winston.info(`Adding ${workItems.length} to total list`);
        this.pbis = _.concat(this.pbis, workItems);
    }
    
    public async getProject():Promise<void>{

        const coreApiObject: CoreApi.CoreApi = await this.azureConnection.getCoreApi();
        const project: CoreInterfaces.TeamProject = await coreApiObject.getProject(this.Project);

        if (_.isNull(project)){
            winston.error(`Project '${this.Project}' not found.`);
            throw new Error(`Project '${this.Project}' not found.`);
        }

        let team:CoreInterfaces.WebApiTeam;
        if (this.TeamName === undefined){
            winston.info(`No team found. Using default team '${project.defaultTeam.name}'`);
            team = project.defaultTeam;
        }
        else 
            team = await coreApiObject.getTeam(project.id, this.TeamName); 

        this.teamContext = {
            project: project.name,
            projectId: project.id,
            team: team.name,
            teamId: team.id
        };

        winston.info(`Connected to project '${this.teamContext.project}' with team '${this.teamContext.team}'`);
    }

    public async getBoardColumns():Promise<void>{
        if (this.teamContext === undefined){
            winston.error(`Team context not created.`);
            throw new Error(`Team context not created.`);
        }

        winston.info(`Fetching board '${this.BoardName}'`);
        const workApiObject: WorkApi.IWorkApi = await this.azureConnection.getWorkApi();
        const columnsName = await workApiObject.getBoardColumns(this.teamContext, this.BoardName);

        if (_.isNull(columnsName)){
            winston.error(`Board '${this.BoardName}' not found.`)
            throw new Error(`Board '${this.BoardName}' not found.`);
        }

        this.createWorkflow(columnsName);
        winston.info(`Workflow is: ${this.Workflow}`)
    }

    private async extractWorkItems(workItems:witInterfaces.WorkItemReference[], witApi: witApi.IWorkItemTrackingApi):Promise<Array<WorkItem>>{
        winston.info(`  Got ${workItems.length} work items`);
        if (workItems.length === 0)
            return;

        let tempUpdates: witInterfaces.WorkItemUpdate[];
        let updates: witInterfaces.WorkItemUpdate[];
        let start = new Date();
        let myWorkItems = new Array<WorkItem>(workItems.length);
        let workItem: witInterfaces.WorkItemReference;
        for (let i in workItems) {
            workItem = workItems[i];
            let id = workItem.id;
            let moarUpdates: witInterfaces.WorkItemUpdate[];
            updates = await witApi.getUpdates(workItem.id );
            if (updates.length === 200){
                let size = 200;
                tempUpdates = await witApi.getUpdates(workItem.id, 200, size);
                updates = _.concat(updates, tempUpdates);
                while (tempUpdates.length === 200){
                    size += 200;
                    tempUpdates = await witApi.getUpdates(workItem.id, 200, size);
                    updates = _.concat(updates, tempUpdates);
                }
            } 

            let workflowDates = this.extractColumnsNameAndDates(updates)

            myWorkItems[i] = new WorkItem(workItem.id, 
                                         this.generateLink(workItem.id),
                                         this.extractTitle(updates), 
                                         workflowDates, 
                                         this.extractProperties(workflowDates, updates));
        }
        let end = differenceInMilliseconds(new Date(), start);
        winston.info(`  It took ${end} milliseconds to extract the work items`);

        return myWorkItems;
    }

    private extractProperties(workflowDates:any[], updates:witInterfaces.WorkItemUpdate[]):Object{

        let properties = new Array<string>();

        let latestDate = new Date(1900,0, 1, 0, 0, 0, 0);
        let status = this.columnsName[0];
        _.forEach(this.columnsName, (name, i) => {
            let dateInThisColumn = workflowDates[i][name];
            if (dateInThisColumn !== undefined &&
                dateInThisColumn > latestDate ){
                latestDate = dateInThisColumn
                status = name;
            }
        });

        let workItemType = updates[0].fields["System.WorkItemType"].newValue;
        let effortField = _.findLast(updates, (x) => {
            if (x.fields !== undefined &&
                x.fields['Microsoft.VSTS.Scheduling.Effort'] !== undefined &&
                x.fields['Microsoft.VSTS.Scheduling.Effort'].newValue !== undefined)
                return true;
            else
                return false;
         });

        let effort = '';
        if (effortField !== undefined) 
            effort = effortField.fields['Microsoft.VSTS.Scheduling.Effort'].newValue;

        return {
            'Status':status,
            "Type":workItemType,
            "Effort": effort
        };
    }

    private extractColumnsNameAndDates(updates: witInterfaces.WorkItemUpdate[]): Array<any> {
        let workflowDates = new Array<any>();

        let columnsName = _.filter(updates, (x) => {
            if (x.fields === undefined)
                return false;

            return x.fields['System.BoardColumn'] !== undefined ||
                x.fields['System.BoardColumnDone'] !== undefined
        });

        let columnsNameMap = _.map(columnsName, (x) => {
            return {
                'ColumnName': this.getValue(x),
                'Date': parseISO(x.fields['System.ChangedDate'].newValue)
            };
        });

        _.forEach(this.columnsName, function (columnName, index) {
            // Find the latest date in which the work item went into the column
            let item = _.findLast(columnsNameMap, (x) => { return x.ColumnName == columnName });

            if (item !== undefined)
                workflowDates.push({ [columnName]: item.Date });
            else
                workflowDates.push({ [columnName]: undefined });

        });

        return workflowDates;
    }

    private getValue(update: witInterfaces.WorkItemUpdate): any {

        let column = update.fields['System.BoardColumn'];
        let subColumn = update.fields['System.BoardColumnDone']

        if (column !== undefined) {
            // Work item move to a new column
            this.lastColumn = column.newValue;
            
            if (subColumn === undefined || subColumn.oldValue === undefined) {    
                // 3 scenarios in this block
                //  1- This is the first column of the board
                //  2- Work item moved from columnsName without subcolumnsName
                //  3- Work item moved from a column without subcolumnsName to a column which didn't had subcolumnsName at that time but now has subcolumnsName
                if (this.isSplitColumn[column.newValue])
                    return `${column.newValue} - Doing`;
                else
                    return column.newValue;
            }
            else if (subColumn !== undefined) {
                // 2 scenarios in this block
                //  1- From columnsName with subcolumnsName
                //  2- From column with subcolumnsName to column without subcolumn
                // 3 - From column with subolumns to column with subcolumnsName
                if (this.isSplitColumn[column.newValue]) {
                    if (subColumn.newValue === true)
                        // Doing -> Done
                        return `${column.newValue} - Done`;
                    else
                        // Done -> Doing
                        return `${column.newValue} - Doing`;
                }
                else
                    return `${column.newValue}`;
            }
        }
        else if (column === undefined && subColumn !== undefined) {
            // Work Item moved from subcolumn to subcolumn 
            // within the same column
            if (subColumn.newValue === true)
                // Doing -> Done
                return `${this.lastColumn} - Done`;
            else
                // Done -> Doing
                return `${this.lastColumn} - Doing`;
        } else
            return undefined;
    }

    private createWorkflow(columnsName: BoardColumn[]): void {

        this.columnsName = new Array(0);
        this.isSplitColumn = new Map<string, boolean>();
        for (let column of columnsName) {
            this.isSplitColumn[column.name] = column.hasOwnProperty('isSplit') && column.isSplit;
            if (column.isSplit === true) {
                this.columnsName.push(`${column.name} - Doing`);
                this.columnsName.push(`${column.name} - Done`);
            }
            else
                this.columnsName.push(column.name);
        }
    }

    private generateLink(id:number):string{
        return `${Configuration.getInstance().Url}/${Configuration.getInstance().ProjectName}/_workitems/edit/${id}`
    }

    private extractTitle(updates: witInterfaces.WorkItemUpdate[]): string {

        let update = _.findLast(updates, (x) => { 
            if (x.fields === undefined)
                return false;

            return x.fields['System.Title'] !== undefined; 
        });
        
        if (update !== undefined)
            return update.fields['System.Title'].newValue;
        else
            return "";
    }

    public get Workflow(): Array<string> {
        return this.columnsName;
    }
    private columnsName: Array<string>;

    public get Headers():Array<string>{
        let headers = _.clone(this.Workflow);
        headers = _.concat(headers, Configuration.getInstance().Properties);
        return headers;            
    }

    public get Pbis(): Array<WorkItem> {
        return this.pbis;
    }
    private pbis: Array<WorkItem>;

    private isSplitColumn: Map<string, boolean>;
    private lastColumn: string;

    private azureConnection: azdev.WebApi = undefined;
    private teamContext: CoreInterfaces.TeamContext = undefined;
}