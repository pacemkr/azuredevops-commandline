import {Configuration} from "./Configuration";
import {format} from 'date-fns';

import * as _ from 'lodash';

export class WorkItem{

    constructor(public Id:number,
        public Link:string,
        public Title:string,
        public WorkflowDates:Array<string>,
        public Properties:Object){}


    public toObject():Object{
        // Create the object
        let data = {
            "id":this.Id,
            "link": this.Link,
            "title": this.Title,
        };

        // Set the dates on each column of the workflow
        for (let item of this.WorkflowDates){
            let key = Object.keys(item)[0];
            let value = item[ key ];
            if (value !== undefined)
                item[ key ] = format(value, Configuration.getInstance().DateFormat);

            Object.assign(data, item);
        }

        // Set the properties to the data object
        return Object.assign(data, this.Properties);
    }
}