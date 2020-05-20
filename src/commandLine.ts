import * as winston from "winston";
import * as _ from 'lodash';

import {AzureConnection} from "./AzureConnection"
import {CsvExporter} from "./CsvExporter"
import {Configuration} from "./Configuration"
import {Logging} from './Logging';

Logging.configure();

let ac = new AzureConnection(
    Configuration.getInstance().ProjectName,
    Configuration.getInstance().TeamName,
    Configuration.getInstance().BoardName);

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
