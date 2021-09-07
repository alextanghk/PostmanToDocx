# Convert Postman API Document Json

CLI tool to convert Postman collection(v2.1) Json file to word document(docx).

- [Description](#description)
- [Environment](#environment)
- [Usage](#usage)

## Description

Postman is a very useful tools for API testing and documentation, however it didn't provide the option on exporting the API document to Word document, so I create my own cli to serve this purpose. 

This tools will convert the collection (v2.1) json into a word document, you can combine two collection json to a array in single file.

<br />

## Environment

<br />

Node JS Version: 12 or above

<br />

## Usage

<br />


````
$ npm i postman-to-doc -g
$ convert-postman --sources=<FULL_PATH_OF_JSON> [--output=<PATH_OF_OUTPUT_LOCATION>]
````