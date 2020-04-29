import {Configuration} from "./Configuration";

import {format} from "date-fns";
import * as _ from 'lodash';

export class WorkItem{

    constructor(public Id:number,
        public Title:string,
        public WorkflowDates:Array<any>,
        public Properties:Array<string>){}


    public toObject():Object{
        let t = {
            "id":this.Id,
            "link": `${Configuration.getInstance().Url}/${Configuration.getInstance().ProjectName}/_workitems/edit/${this.Id}`,
            "title": this.Title,
        };

        for (let item of this.WorkflowDates){
            let key = Object.keys(item)[0];
            if (item[ key ] != "")
                item[ key ] = format( item[ key ], Configuration.getInstance().DateFormat);

            Object.assign(t, item);
        }

       return t;
    }
}