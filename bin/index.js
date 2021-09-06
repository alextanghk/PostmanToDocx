#!/usr/bin/env node

const fs = require("fs");
const util = require('util');
const path = require("path");

const ConvertDocx = require("./ConvertDocx");

const yargs = require('yargs/yargs')
const { hideBin } = require('yargs/helpers');


const argv = yargs(hideBin(process.argv))
    .option('source',{
        alias: "s",
        string: true,
        default: "",
        describe: "Full path of the json file."
    })
    .option('output',{
        alias: "o",
        string: true,
        default: "",
        describe: "Output file path"
    })
    .help()
    .alias('help', 'h')
    .argv;


const source = argv.source;
const basename = path.basename(source);
const output = argv.output == "" ? source.replace(basename,""): argv.output;

let apiJson = require(source);
if (!Array.isArray(apiJson)) 
{
    apiJson = [apiJson];
}
ConvertDocx(path.join(output,basename.replace(".json",".docx")), apiJson);