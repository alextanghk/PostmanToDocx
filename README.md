# Convert Postman API Document Json

CLI tool to convert Postman collection Json file to word document(docx).

- [Description](#description)
- [Environment](#environment)
- [Supporting Schema](#supporting-schema)
- [Usage](#usage)

<br />

## Description

<br />

Postman is a very useful tool for API testing and documentation, however it didn't provide the option on exporting the API document to Word document, so I create my own cli to serve this purpose. 


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


## Usage

<br />


````
$ npm i postman-to-doc -g
$ convert-postman --sources=<FULL_PATH_OF_JSON> [--output=<PATH_OF_OUTPUT_LOCATION>, --schema=<SCHEMA_VERSION>]
````