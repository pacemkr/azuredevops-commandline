
import * as _ from 'lodash';
import * as path from "path";
import { writeToPath } from '@fast-csv/format';

import { WorkItem } from "./WorkItem";
import { Configuration } from "./Configuration"

export class CsvExporter {

    constructor(headers: Array<string>, filename: String) {
        this.headers = _.concat(["id", "link", "title"], headers);
    }

    export(workItems: Array<WorkItem>): void {

        // Convert the objects to exportable data
        let rows = _.map(workItems, (x) => {
            return x.toObject();
        })

        writeToPath(path.resolve(__dirname, Configuration.getInstance().CsvFilename), rows, {
            headers: this.headers,
            writeHeaders: true
        })
            .on('error', err => console.error(err))
            .on('finish', () => console.log('Done writing.'));

    }

    private options: any;
    private headers: any;
}
