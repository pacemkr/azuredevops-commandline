import * as winston from "winston";
import * as _ from 'lodash';

import {AzureConnection} from "./AzureConnection"
import {CsvExporter} from "./CsvExporter"
import {Configuration} from "./Configuration"
import { WorkItem } from './WorkItem';

winston.configure({
    level: 'info',
    format: winston.format.simple(),
    transports: [
        // new BrowserConsole(
        //     {
        //         format: winston.format.simple(),
        //         level: "debug",
        //     },
        // ),
        new winston.transports.Console()
    ],
})


let ac = new AzureConnection();
ac.connect().then(async function(result){
    await ac.getProject();
    await ac.getBoardColumns();
    await ac.fetchPbis();

    let headers = _.clone(ac.Workflow);
    headers = _.concat(headers, Configuration.getInstance().Properties);
 
    let csvExporter = new CsvExporter(
        headers, 
        Configuration.getInstance().CsvFilename);
    
    let workItems = _.concat(ac.WipWorkItems, ac.DoneWorkItems);

    csvExporter.export(workItems);    
}).catch(function(error){
    winston.error(error)
})
