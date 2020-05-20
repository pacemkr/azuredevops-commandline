import * as winston from "winston";

export class Logging{

    public static configure():void{
        winston.configure({
            level: 'info',
            format: winston.format.simple(),
            transports: [
                // new BrowserConsole(
                //     {
                //         format: winston.format.simple(),
                //         level: "debug",
                //     },
                // ),
                new winston.transports.Console()
            ],
        })
        
    }

}