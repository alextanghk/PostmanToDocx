#!/usr/bin/env node
"use strict"

const fs = require("fs");
const path = require("path");
const Ajv = require("ajv")
const yargs = require('yargs/yargs')
const { hideBin } = require('yargs/helpers');
const https = require('https');


const ConvertDocx = require("./ConvertDocx");
const parseSchema = require("./schema/schema");

const argv = yargs(hideBin(process.argv))
    .example([
        ['$0 --source=/root/API.json', "Export to same directory"],
        ['$0 --source=/root/API.json --schema=2.1.0', "Export to same directory with schema version"],
        ['$0 --source=/root/API.json --output=/Export/', "Export to other directory"],
        ['$0 --source=/root/API.json --output=/Export/ --schema=2.1.0', "Export to other directory with schema version"]
    ])
    .options({
        "source" : {
            alias:"s",
            string: true,
            demandOption: "Sources cannot be empty",
            describe: "Full path of the json file."
        },
        "output": {
            alias:"o",
            string: true,
            default: "",
            describe: "Output file path"
        },
        "schema": {
            alias:"c",
            string: true,
            default: "2.1.0",
            describe: "Version of Postman Collection schema",
            choices: ['2.1.0','2.0.0']
        }
    }).check((argv, option)=>{
        const source = argv.source;
        const output = argv.output;
        if (!fs.existsSync(source)) {
            throw new Error(`Cannot found source file from ${source}`);
        }
        if (output != "") {
            if (fs.existsSync(output) && !fs.lstatSync(output).isDirectory()) {
                throw new Error("Output value is not a directory");
                
            }
        }
        return true;
    })
    .help()
    .alias('help', 'h')
    .argv;


const source = argv.source;

console.log("Reading file from %s", source);

const basename = path.basename(source);
const output = argv.output == "" ? source.replace(basename,""): argv.output;

if (!fs.existsSync(output)) {
    console.log("Directory not exist, creating the folder");
    fs.mkdirSync(output,{recursive: true});
}

const schemaVersion = argv.schema;

// Validation Schema
const apiJson = require(source);

const schemaUrls = require("../schema.json");

if (
    schemaUrls[schemaVersion] === "" ||  
    schemaUrls[schemaVersion] === undefined ||
    schemaUrls[schemaVersion] === null
)
{
    throw new Error(`Do not support this schema version v${schemaVersion}`);
}
let p = new Promise((resolve, reject)=>{
    https.get(schemaUrls[schemaVersion], (res)=>{
        let body = "";
        res.on("data",(chunk)=>{
            body += chunk;
        });

        res.on("end",()=>{
            try {
                let schema = JSON.parse(body);
                resolve(schema);
            } catch(e) {
                reject(`Got error: ${e.message}`);
            }
        })
    }).on('error', (e) => {
        reject(`Got error: ${e.message}`);
    });
})

p.then((schema) =>{
    const ajv = new Ajv({
        strict: false
    });
    const validate = ajv.compile(schema);
    const json = (apiJson instanceof Array) ? apiJson: [apiJson];
    if (json.length < 1) {
        Promise.reject("Source cannot be empty");
    }
    for(let item of json) {
        if (!validate(item)) {
            Promise.reject(`Source schema invalid`);
        }
    }
    
    return json.map((section)=>{
        return parseSchema(section, schemaVersion);
    });
}).then((json)=>{
    return ConvertDocx(json, {
        output: path.join(output,basename.replace(".json",".docx"))
    });
}).then((result)=>{
    console.log("File saved!");
}).catch((err)=>{
    throw new Error(err.message);
})