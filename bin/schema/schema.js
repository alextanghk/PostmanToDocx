const parse200 = require("./v2.0.0");
const parse210 = require("./v2.1.0");


function parseSchema(source, version = "2.1.0") {
    let result = {};
    switch(version) {
        case "2.0.0":
            result = parse200(source);
            break;
        case "2.1.0":
            result = parse210(source);
            break;
        default:
            throw new Error("Version not support");
            break;
    }
    return result;
}

module.exports = parseSchema;