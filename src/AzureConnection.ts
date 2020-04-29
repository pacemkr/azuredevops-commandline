import * as _ from 'lodash';
import {differenceInMilliseconds} from "date-fns";

import * as azdev from "azure-devops-node-api";
import * as lim from "azure-devops-node-api/interfaces/LocationsInterfaces";
import * as CoreApi from 'azure-devops-node-api/CoreApi';
import * as CoreInterfaces from 'azure-devops-node-api/interfaces/CoreInterfaces';
import * as WorkApi from 'azure-devops-node-api/WorkApi';
import * as witApi from 'azure-devops-node-api/WorkItemTrackingApi';
import * as witInterfaces from 'azure-devops-node-api/interfaces/WorkItemTrackingInterfaces';

import {Configuration} from "./Configuration"
import {WorkItem} from "./WorkItem";

export class AzureConnection {

    constructor(){
        this.workItems = new Array<WorkItem>();
    
    }
    async run(): Promise<void> {
        
        let orgUrl = Configuration.getInstance().Url;

        let token: string = Configuration.getInstance().Token;
        let projectName = Configuration.getInstance().ProjectName;

        let authHandler = azdev.getPersonalAccessTokenHandler(token);
        let connection = new azdev.WebApi(orgUrl, authHandler);
        let connData: lim.ConnectionData = await connection.connect();
        console.log(`Hello ${connData.authenticatedUser.providerDisplayName}`);

        const workApiObject: WorkApi.IWorkApi = await connection.getWorkApi();
        const coreApiObject: CoreApi.CoreApi = await connection.getCoreApi();
        const witApi: witApi.IWorkItemTrackingApi = await connection.getWorkItemTrackingApi();
        const project: CoreInterfaces.TeamProject = await coreApiObject.getProject(projectName );
        const teamContext: CoreInterfaces.TeamContext = {
            project: projectName ,
            projectId: project.id,
            team: project.defaultTeam.name,
            teamId: project.defaultTeam.id
        };

        const q: witInterfaces.QueryHierarchyItem[] = await witApi.getQueries(projectName );
        const wip: witInterfaces.QueryHierarchyItem = await witApi.getQuery(projectName, Configuration.getInstance().WipQuery);
        const results: witInterfaces.WorkItemQueryResult = await witApi.queryById(wip.id, teamContext, false);
        

        let updates: witInterfaces.WorkItemUpdate[];
        console.log(`Got ${results.workItems.length} work items`);
        let start = new Date();
        this.workItems = new Array<WorkItem>(results.workItems.length);
        let workItem:witInterfaces.WorkItemReference;
        for (let i in results.workItems) {
            workItem = results.workItems[i];
            updates = await witApi.getUpdates(workItem.id);
            
            let workFlowDates = this.extractColumnsAndDates(updates);
            let title = this.extractTitle(updates);
            let status = this.extractStatus(workFlowDates);
            
            this.workItems[i] = new WorkItem(workItem.id, title, workFlowDates, [status, "0"]);
        }
        let end = differenceInMilliseconds(new Date(), start);
        console.log(`It took ${end} milliseconds`);

        //const revisions: WorkItemTrackingInterfaces.WorkItem[] = await witApi.getRevisions(3);
        //console.log(`result length is ${results.workItems.length}`)        
    }
    extractStatus(workFlowDates: Array<any>):string {
          
        _.forEachRight(workFlowDates, (x, index) =>{
            if (x['Date'] != undefined)
                return Configuration.getInstance().Workflow[index];
        })
            
        return Configuration.getInstance().Workflow[0];
    }

    extractColumnsAndDates(updates:witInterfaces.WorkItemUpdate[]):Array<any>{
        let workflowDates = new Array<any>();

        let columns = _.filter(updates, (x) => { return x.fields['System.BoardColumn'] !== undefined });

        let columnsMap = _.map(columns, (x) => { return {
            'ColumnName':x.fields['System.BoardColumn'].newValue, 
            'Date':new Date(x.fields['System.ChangedDate'].newValue)
            }; 
        } );

        let workflow = Configuration.getInstance().Workflow;
        _.forEach(workflow, function(columnName, index){
            let item = _.find(columnsMap, (x) => {return x.ColumnName == columnName});
            
            if (item !== undefined){
                workflowDates.push({[columnName]:item.Date});                
            }                
            else
                workflowDates.push({[columnName]:""});

        });

        return workflowDates;
    }

    extractTitle(updates:witInterfaces.WorkItemUpdate[]):string{

        let update = _.findLast(updates, (x) => {return x.fields['System.Title'] !== undefined;} );
        if (update !== undefined)
            return update.fields['System.Title'].newValue;
        else
            return "";
    }
    
    public get WorkItems():Array<WorkItem>{
        return this.workItems;
    }

    private workItems:Array<WorkItem>;
}