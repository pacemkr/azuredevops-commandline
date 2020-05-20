import * as express from 'express';
import * as json2csv from 'json2csv';
import * as _ from 'lodash';
import * as winston from 'winston';
import {format} from 'date-fns';

import { AzureConnection } from "./AzureConnection"
import { Configuration } from "./Configuration"
import { Logging } from "./Logging"

Logging.configure();
let port = 8080;
winston.info(`Starting web server`);

let app = express();

app.get('/node/:project/:teamName/:board', async (req, res, next) => {
    winston.info(`Got request /${req.params.project}/${req.params.teamName}/${req.params.board}`);
    try {

        let ac = new AzureConnection(
            req.params.project,
            req.params.teamName,
            req.params.board);

        await ac.connect();
        await ac.getProject();
        await ac.getBoardColumns();

        // Build the queries based on the team name
        let queries = _.map(Configuration.getInstance().Queries, (query) => {
            return _.replace(query, "/", `/${req.params.teamName} - `);
        });

        for (let query of queries)
            await ac.fetchPbis(query);

        let exportableObjects = _.map(ac.Pbis, (x) => {
            return x.toObject();
        });

        let filename = createFilename(req.params.project, req.params.teamName);
        winston.info(`Exporting data to file ${filename}`);
        let jsonParser = new json2csv.Parser();
        jsonParser.preprocessFieldsInfo(ac.Headers);
        var dataInCsv = jsonParser.parse(exportableObjects);
        winston.info('Exporting done');

        res.attachment(`${filename}`);
        res.status(200).send(dataInCsv);
    }
    catch (error) {
        return next(error);
    }
});

function createFilename(project:string, teamName:string):string{
    let now = new Date();
    let nowFormatted = format(now, "yyyy_MM_dd_HH_MM");

    return `${project}_${teamName}_${nowFormatted}.csv`;
}

winston.info(`Web server started on port ${port}`);
app.listen(port);