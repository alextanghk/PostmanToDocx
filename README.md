# Convert Postman API Document Json

Converting Postman collection Json file to MS Word document(docx).

- [Description](#description)
- [Environment](#environment)
- [Supporting Schema](#supporting-schema)
- [Installation](#installation)
- [Usage](#usage)
- [Command](#command)

<br />

## Description

<br />

Postman is a very useful tool for API testing and documentation, however it didn't provide the option on exporting the API document to Word document, so I created my own module to serve this purpose. 


<br />

## Environment

<br />

Node JS Version: 12 or above

<br />

## Supporting Schema

<br />

- https://schema.postman.com/collection/json/v2.0.0/draft-07/collection.json
- https://schema.postman.com/collection/json/v2.1.0/draft-07/collection.json

<br />

## Installation 
````
$ npm i postman-to-docx -g // For cli
$ npm i postman-to-docx // For node project
````
<br />

## Usage
To use this function in your code:
<br />
````
const postmanToDocx = require("postman-to-docx");

// Promise return with true if success and throw Error when fail.
// postmanToDocx(schema: string, source: string, [outputPath: string])
postmanToDocx("2.1.0","/Souces/API.json","/Export"); 

````
<br />

## Command
You can also use cli to convert the json to docx file.
<br />
````
$ p2dx --sources=<FULL_PATH_OF_JSON> [--output=<PATH_OF_OUTPUT_LOCATION>, --schema=<SCHEMA_VERSION>]
選項：
      --version  顯示版本                                                 [布林]
  -s, --source   Full path of the json file.                       [字串] [必填]
  -o, --output   Output file path                           [字串] [預設值： ""]
  -c, --schema   Version of Postman Collection schema
                           [String] [可選值： "2.1.0", "2.0.0"] [預設值： "2.1.0"]
  -h, --help     顯示說明                                                 [布林]

範例：
  p2dx --source=/root/API.json          Export to same directory
  p2dx --source=/root/API.json          Export to same directory with schema
  --schema=2.1.0                        version
  p2dx --source=/root/API.json          Export to other directory
  --output=/Export/
  p2dx --source=/root/API.json          Export to other directory with
  --output=/Export/ --schema=2.1.0      schema version
````

