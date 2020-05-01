import {Configuration} from "./Configuration";
import {format} from 'date-fns';

import * as _ from 'lodash';

export class WorkItem{

    constructor(public Id:number,
        public Title:string,
        public WorkflowDates:Array<string>,
        public Properties:Array<string>){}


    public toObject():Object{
        let t = {
            "id":this.Id,
            "link": `${Configuration.getInstance().Url}/${Configuration.getInstance().ProjectName}/_workitems/edit/${this.Id}`,
            "title": this.Title,
        };

        for (let item of this.WorkflowDates){
            let key = Object.keys(item)[0];
            let value = item[ key ];
            if (value !== undefined)
                item[ key ] = format(value, Configuration.getInstance().DateFormat);

            Object.assign(t, item);
        }

       return t;
    }
}