# Summary
Code to extract data from different Agile planning tools

# Install and build
To install the Node modules required, type:

```npm install```

To build the application, type:

```npm run build ```

You will need TypeScript version >= 3.0

To update to latest version : 
```npm install -g typescript@latest```

# Configuration
To make it easy to understand how to configure the extraction, here's an example of a config file but instead of having values, each field has a short sentence to document it.

```
{
    "url": <The base URL where to extract data from> 
    "project": {
        "name": <The name of the project>,
        "token":<The token key>,
        "csvFilename":<The name of the csv file. Defaults to the name of project if none is given>,
        "dateformat":<The date format for dates in the csv file>,
        "queries":{
            "wip":<The query to fetch the Work In Progress items>,
            "done":"<The query to fetch the Done items>"
        },
        "properties":[ <An array of additional properties to extract>
        ]
    }
}
```

# Running the data extraction
To run the extraction, type

```node out/main.js ```
