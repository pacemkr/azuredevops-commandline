import * as express from 'express';
import * as json2csv from 'json2csv';
import * as _ from 'lodash';

import {AzureConnection} from "./AzureConnection"
import {Configuration} from "./Configuration"
import { runInNewContext } from 'vm';

let app = express();

app.get('/node/:project/:teamName/:board', async (req, res, next) => {
try{


    let ac = new AzureConnection(
        req.params.project, 
        req.params.teamName,
        req.params.board);

    await ac.connect();
    await ac.getProject();
    await ac.getBoardColumns();
    
    for (let query of Configuration.getInstance().Queries)
        await ac.fetchPbis(query);
     
    
    let exportableObjects = _.map(ac.Pbis, (x) => {
        return x.toObject();
    });

    let jsonParser = new json2csv.Parser();
    jsonParser.preprocessFieldsInfo(ac.Headers);
    var dataInCsv = jsonParser.parse(exportableObjects);

    res.attachment(`${req.params.project}.csv`);
    res.status(200).send(dataInCsv);
}
catch (error){
        return next(error);
}
});


app.listen(8080);