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

export class AzureConnection {

    constructor() {
    }
 
    async run(): Promise<void> {

        let orgUrl = Configuration.getInstance().Url;

        let token: string = Configuration.getInstance().Token;
        let projectName = Configuration.getInstance().ProjectName;

        let authHandler = azdev.getPersonalAccessTokenHandler(token);
        let connection = new azdev.WebApi(orgUrl, authHandler);
        let connData: lim.ConnectionData = await connection.connect();
        winston.info(`Hello ${connData.authenticatedUser.providerDisplayName}`);
        winston.info(`Connected to ${Configuration.getInstance().Url}`)

        const workApiObject: WorkApi.IWorkApi = await connection.getWorkApi();
        const coreApiObject: CoreApi.CoreApi = await connection.getCoreApi();
        const witApi: witApi.IWorkItemTrackingApi = await connection.getWorkItemTrackingApi();
        const project: CoreInterfaces.TeamProject = await coreApiObject.getProject(projectName);
        const teamContext: CoreInterfaces.TeamContext = {
            project: projectName,
            projectId: project.id,
            team: project.defaultTeam.name,
            teamId: project.defaultTeam.id
        };

        winston.info(`Connected to project ${project.name}`);

        winston.info(`Fetching board '${Configuration.getInstance().BoardName}'`);
        const columns = await workApiObject.getBoardColumns(teamContext, Configuration.getInstance().BoardName);
        this.createWorkflow(columns);
        winston.info(`Workflow is: ${this.Workflow}`)

        winston.info(`Calling query ${Configuration.getInstance().WipQuery}`);
        const wip: witInterfaces.QueryHierarchyItem = await witApi.getQuery(projectName, Configuration.getInstance().WipQuery);
        let results: witInterfaces.WorkItemQueryResult = await witApi.queryById(wip.id, teamContext, false);        
        this.wipWorkItems = await this.extractWorkItems(results.workItems, witApi);

        winston.info(`Calling query ${Configuration.getInstance().DoneQuery}`);
        const done: witInterfaces.QueryHierarchyItem = await witApi.getQuery(projectName, Configuration.getInstance().DoneQuery);
        results = await witApi.queryById(done.id, teamContext, false);
        this.doneWorkItems = await this.extractWorkItems(results.workItems, witApi);
    }

    private async extractWorkItems(workItems:witInterfaces.WorkItemReference[], witApi: witApi.IWorkItemTrackingApi):Promise<Array<WorkItem>>{
        winston.info(`  Got ${workItems.length} work items`);
        if (workItems.length === 0)
            return;

        let updates: witInterfaces.WorkItemUpdate[];
        let start = new Date();
        let myWorkItems = new Array<WorkItem>(workItems.length);
        let workItem: witInterfaces.WorkItemReference;
        for (let i in workItems) {
            workItem = workItems[i];
            let id = workItem.id;
            updates = await witApi.getUpdates(workItem.id);

            let workflowDates = this.extractColumnsAndDates(updates)

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

    private generateLink(id:number):string{
        return `${Configuration.getInstance().Url}/${Configuration.getInstance().ProjectName}/_workitems/edit/${id}`
    }

    private extractProperties(workflowDates:any[], updates:witInterfaces.WorkItemUpdate[]):Object{

        let properties = new Array<string>();
        let status:string;
        _.forEachRight(workflowDates, (x, index) => {
            if (x['Date'] != undefined)
                status = this.columns[index];
        })

        status = this.columns[0];


        let workItemType = updates[0].fields["System.WorkItemType"].newValue;

        return {
            'Status':status,
            "Work Item Type":workItemType
        };
    }

    private extractColumnsAndDates(updates: witInterfaces.WorkItemUpdate[]): Array<any> {
        let workflowDates = new Array<any>();

        let columns = _.filter(updates, (x) => {
            if (x.fields === undefined)
                return false;

            return x.fields['System.BoardColumn'] !== undefined ||
                x.fields['System.BoardColumnDone'] !== undefined
        });

        let columnsMap = _.map(columns, (x) => {
            return {
                'ColumnName': this.getValue(x),
                'Date': parseISO(x.fields['System.ChangedDate'].newValue)
            };
        });

        _.forEach(this.columns, function (columnName, index) {
            // Find the latest date in which the work item went into the column
            let item = _.findLast(columnsMap, (x) => { return x.ColumnName == columnName });

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
                //  2- Work item moved from columns without subcolumns
                //  3- Work item moved from a column without subcolumns to a column which didn't had subcolumns at that time but now has subcolumns
                if (this.isSplitColumn[column.newValue])
                    return `${column.newValue} - Doing`;
                else
                    return column.newValue;
            }
            else if (subColumn !== undefined) {
                // 2 scenarios in this block
                //  1- From columns with subcolumns
                //  2- From column with subcolumns to column without subcolumn
                // 3 - From column with subolumns to column with subcolumns
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

    public get Workflow(): Array<string> {
        return this.columns;
    }

    private createWorkflow(columns: BoardColumn[]): void {

        this.columns = new Array(0);
        this.isSplitColumn = new Map<string, boolean>();
        for (let column of columns) {
            this.isSplitColumn[column.name] = column.hasOwnProperty('isSplit') && column.isSplit;
            if (column.isSplit === true) {
                this.columns.push(`${column.name} - Doing`);
                this.columns.push(`${column.name} - Done`);
            }
            else
                this.columns.push(column.name);
        }
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

    public get WipWorkItems(): Array<WorkItem> {
        return this.wipWorkItems;
    }
    private wipWorkItems: Array<WorkItem>;

    public get DoneWorkItems(): Array<WorkItem> {
        return this.doneWorkItems;
    }
    private doneWorkItems: Array<WorkItem>;

    private columns: Array<string>;
    private isSplitColumn: Map<string, boolean>;
    private lastColumn: string;
}