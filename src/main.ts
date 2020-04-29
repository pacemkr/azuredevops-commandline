import {AzureConnection} from "./AzureConnection"
import {CsvExporter} from "./CsvExporter"
import {Configuration} from "./Configuration"

let ac = new AzureConnection();
ac.run().then(function(result){
    let csvExporter = new CsvExporter(
        Configuration.getInstance().Workflow, 
        Configuration.getInstance().CsvFilename);
    
    csvExporter.export(ac.WorkItems);    
}).catch(function(error){
    console.log(error)
})
