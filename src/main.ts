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

    for (let query of Configuration.getInstance().Queries){
        await ac.fetchPbis(query);
    }

 
    let csvExporter = new CsvExporter(
        ac.Headers, 
        Configuration.getInstance().CsvFilename);
    
    csvExporter.export(ac.Pbis);    
}).catch(function(error){
    winston.error(error)
})
