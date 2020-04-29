

export class Configuration{

    public get ProjectName():string{
        return this.config.project.name as string;
    }

    public get DateFormat():string{
        return this.config.project.dateformat as string;
    }

    public get Url():string{
        return this.config.url as string;
    }

    public get CsvFilename():string{
        return this.config.project.csvFilename as string;
    }

    public get Token():string{
        return this.config.project.token as string;
    }

    public get WipQuery():string{
        return this.config.project.queries.wip as string;
    }

    public get DoneQuery():string{
        return this.config.project.queries.done as string;
    }

    public get Workflow():Array<string>{
        return this.config.project.workflow as Array<string>;
    }

    private constructor(){
        this.config = require('./config.json');
    }

    public static getInstance(): Configuration {
        if (!Configuration.instance) {
            Configuration.instance = new Configuration();
        }

        return Configuration.instance;
    }
    private config:any;

    private static instance: Configuration;
}