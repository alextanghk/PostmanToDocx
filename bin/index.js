#!/usr/bin/env node
"use strict"

const fs = require("fs");
const util = require('util');
const path = require("path");
const ConvertDocx = require("./ConvertDocx");
const Ajv = require("ajv")
const yargs = require('yargs/yargs')
const { hideBin } = require('yargs/helpers');
const https = require('https');
const { strict } = require("yargs");
const argv = yargs(hideBin(process.argv))
    .example([
        ['$0 --source=/root/API.json', "Export to same directory"],
        ['$0 --source=/root/API.json --output=/Export/', "Export to other directory"]
    ])
    .options({
        "source" : {
            alias: "s",
            string: true,
            demandOption: "Sources cannot be empty",
            describe: "Full path of the json file."
        },
        "output": {
            alias: "o",
            string: true,
            default: "",
            describe: "Output file path"
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

// Validation Schema
const apiJson = require(source);

https.get("https://schema.postman.com/collection/json/v2.1.0/draft-07/collection.json", (res)=>{
    let body = "";
    res.on("data",(chunk)=>{
        body += chunk;
    });

    res.on("end",()=>{
        const schema = JSON.parse(body);
        const ajv = new Ajv({
            strict: false
        });
        const validate = ajv.compile(schema);
        if (!Array.isArray(apiJson)) 
        {
            if (!validate(apiJson)) {
                throw new Error("Only support Postman Collection v2.1.0");
            }
        } else {
            if (apiJson.length < 1) {
                throw new Error("Source cannot be empty");
            }
            for(let item of apiJson) {
                if (!validate(item)) {
                    throw new Error("Only support Postman Collection v2.1.0");
                }
            }
        }
        ConvertDocx(apiJson, {
            output: path.join(output,basename.replace(".json",".docx"))
        });
    })
})