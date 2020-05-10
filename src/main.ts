import * as winston from "winston";
import * as _ from 'lodash';

import {AzureConnection} from "./AzureConnection"
import {CsvExporter} from "./CsvExporter"
import {Configuration} from "./Configuration"
import { WorkItem } from './WorkItem';

winston.configure({
    transports: [
        // new BrowserConsole(
        //     {
        //         format: winston.format.simple(),
        //         level: "debug",
        //     },
        // ),
        new winston.transports.Console({
            format: winston.format.simple(),
            level:"debug"
        }),
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
    winston.log(error)
})
