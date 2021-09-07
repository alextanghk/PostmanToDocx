const fs = require("fs");
const path = require("path");
const Ajv = require("ajv")
const https = require('https');

const ConvertDocx = require("./bin/ConvertDocx");
const parseSchema = require("./bin/schema/schema");

function postmanToDocx(schemaVersion, source, target = "") {
    
    if (!fs.existsSync(source)) {
        return Promise.reject(`Cannot found source file from ${source}`);
    }

    const basename = path.basename(source);
    const output = target == "" ? source.replace(basename,""): target;

    if (fs.existsSync(output) && !fs.lstatSync(output).isDirectory()) {
        return Promise.reject("Output value is not a directory");
    }
    if (!fs.existsSync(output)) {
        fs.mkdirSync(output,{recursive: true});
    }
    const apiJson = require(source);
    const schemaUrls = require("./schema.json");

    if (
        schemaUrls[schemaVersion] === "" ||  
        schemaUrls[schemaVersion] === undefined ||
        schemaUrls[schemaVersion] === null
    )
    {
        return Promise.reject(`Do not support this schema version v${schemaVersion}`);
    }

    return new Promise((resolve, reject)=>{
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
    }).then((schema) =>{
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
    })
}

module.exports = postmanToDocx;