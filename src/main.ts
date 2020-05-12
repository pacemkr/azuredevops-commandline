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
ac.run().then(function(result){
    let csvExporter = new CsvExporter(
        ac.Workflow, 
        Configuration.getInstance().CsvFilename);
    
    let workItems = _.concat(ac.WipWorkItems, ac.DoneWorkItems);

    csvExporter.export(workItems);    
}).catch(function(error){
    winston.error(error)
})
