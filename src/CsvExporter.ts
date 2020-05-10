import * as winston from 'winston';
import * as _ from 'lodash';
import * as path from "path";
import { writeToPath } from '@fast-csv/format';

import { WorkItem } from "./WorkItem";
import { Configuration } from "./Configuration"

export class CsvExporter {

    constructor(headers: Array<string>, public Filename: string) {
        this.headers = _.concat(["id", "link", "title"], headers);
    }

    export(workItems: Array<WorkItem>): void {
        let filepath = path.resolve(__dirname, this.Filename);
        winston.info(`Extracting to file ${filepath}`);

        // Convert the objects to exportable data
        let rows = _.map(workItems, (x) => {
            return x.toObject();
        })

        writeToPath(filepath, rows, {
            headers: this.headers,
            writeHeaders: true
        })
            .on('error', err => winston.error(err))
            .on('finish', () => winston.info('Extration done'));

    }

    private options: any;
    private headers: any;
}
