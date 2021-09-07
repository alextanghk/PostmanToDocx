const _ = require("lodash");

function urlFormating(url) {
    const { protocol, host, path, query } = url;
    const result = `${protocol}://${host.join(".")}/${path.join("/")}${query !== undefined ? "?"+query.reduce((r,v,k)=>{ r = `${r}&${v.key}=${v.value}`; return r; },"") : ""}`;

    return result;
}

function formItem(item) {
    return {
        name: _.get(item,"name",""),
        description: _.get(item,"request.description",""),
        request: {
            method: _.get(item,"request.method",""),
            header: _.get(item,"request.header",[]),
            query: _.get(item,"request.url.query",[]),
            url: _.get(item,"request.url.raw",_.get(item,"request.url","")),
            auth: {
                type: _.get(item,"request.auth.type","NONE"),
            },
            body: _.get(item,"request.body.raw","")
        },
        response: _.get(item,"response",[]).map((response)=>{
            return {
                name: _.get(response,"name",""),
                status: _.get(response,"status",""),
                code: _.get(response,"code",""),
                body: _.get(response,"body",""),
                request: {
                    method: _.get(response,"originalRequest.method",""),
                    header: _.get(response,"originalRequest.header",[]),
                    query: _.get(item,"originalRequest.query",[]),
                    url: _.get(response,"originalRequest.url.raw",_.get(response,"originalRequest.url","")),
                    auth: {
                        type: _.get(item,"request.auth.type","NONE"),
                    },
                    body: _.get(response,"originalRequest.body.raw","")
                }
            }
        })
    }
}

function parse200(data) {
    let result = {
        info: {
            name: _.get(data,"info.name",""),
            description: _.get(data,"info.description",""),
        },
        item: _.get(data,"item",[]).map((item) =>{
            if (item.item !== undefined) {
                return {
                    name: _.get(item,"name",""),
                    description: _.get(item,"description",""),
                    item: _.get(item,"item",[]).map((sub)=>{
                        return formItem(sub);
                    })
                }
            } else {
                return formItem(item);
            }
        })
    };

    return result;
}

module.exports = parse200;