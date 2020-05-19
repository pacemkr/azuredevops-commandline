import * as express from 'express';
import * as json2csv from 'json2csv';
import * as _ from 'lodash';

import {AzureConnection} from "./AzureConnection"
import {CsvExporter} from "./CsvExporter"
import {Configuration} from "./Configuration"
import { WorkItem } from './WorkItem';

let app = express();

app.get('/node/:project/:teamName/:board', async (req, res) => {

    let ac = new AzureConnection(
        req.params.project, 
        req.params.teamName,
        req.params.board);

    await ac.connect();
        await ac.getProject();
        await ac.getBoardColumns();
    
        for (let query of Configuration.getInstance().Queries){
            await ac.fetchPbis(query);
        }
    
     
        var fields = ac.Headers;
        var d = new json2csv.Parser();
        var stuff = _.map(ac.Pbis, (x) => {
            return x.toObject();
        })
    d.preprocessFieldsInfo(fields);
    var s = d.parse(stuff);
    
        res.attachment('filename.csv');
        res.status(200).send(s);
});


app.listen(8080);