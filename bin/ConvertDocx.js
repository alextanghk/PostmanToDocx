const fs = require("fs");
const docx = require("docx");
const _ = require("lodash")

const { HeadingLevel, WidthType, BorderStyle, Document, Table, TableRow, TableCell, Paragraph, TextRun, Packer } = docx;

// Global Setting
const CellStyles = {
    margins: {
        top: 50,
        bottom: 50,
        left: 50,
        right: 50,
    }
}

const NoBorder = {
    top: { style: BorderStyle.NONE, size: 2, color: "000000" },
    left: { style: BorderStyle.NONE, size: 0, color: "000000" },
    bottom: { style: BorderStyle.NONE, size: 2, color: "000000" },
    right: { style: BorderStyle.NONE, size: 2, color: "000000" },
}

// Global Function
function urlFormating(url) {
    const { protocol, host, path, query } = url;
    const result = `${protocol}://${host.join(".")}/${path.join("/")}${query !== undefined ? "?"+query.reduce((r,v,k)=>{ r = `${r}&${v.key}=${v.value}`; return r; },"") : ""}`;

    return result;
}
function twoColumnRow(title, content) {
    const contentItem = (typeof content === "string") ? new Paragraph(content) : content;
    return new TableRow({
        children:[
            new TableCell({
                width: {
                    size: 20,
                    type: WidthType.PERCENTAGE
                },
                children:[new Paragraph({
                    children: [
                        new TextRun({
                            text: title,
                            bold: true
                        })
                    ]
                })],
                ...CellStyles
            }),
            new TableCell({
                width: {
                    size: 80,
                    type: WidthType.PERCENTAGE
                },
                children:[contentItem],
                ...CellStyles
            })
        ]
    })
}

function descriptionEle(description, options = {}) {
    const desc = description === undefined || description === null ? [] : description.split("\n");
    return new Paragraph({
        children: desc.length === 0 ? [new TextRun("")]: desc.map((t)=>{
            return new TextRun({
                text: t,
                break: 1
            })
        }),
        ...options
    })
}

// Intro function
function sectionIntro(info) {
    const { name, description = ""} = info;
    const desc = description === undefined ? [] : description.split("\n");
    return [new Paragraph({
        heading: HeadingLevel.HEADING_1,
        spacing: {
            after: 100
        },
        children: [
            new TextRun({
                text: name,
                bold: true
            })
        ],
        spacing: {
            after: 100
        }
    }),
        descriptionEle(description,{spacing: {
            after: 300
        }})
    ]
}

function itemIntro(item) {
    const { name, request:{ description = "" }={description: ""}} = item
    return [new Paragraph({
        heading: HeadingLevel.HEADING_2,
        spacing: {
            after: 100
        },
        border: {
            top: { value: "single", space:50, size: 2, color: "000000" },
        },
        children: [
            new TextRun({
                text: name,
                bold: true
            })
        ]
    }),
        descriptionEle(description)
    ]
}


function requestTable(request) {

    const { url, url: {query}, method, body } = request;

    const rows = [
        twoColumnRow("URL:",urlFormating(url)),
        twoColumnRow("Method:",method),
    ]

    if (request.auth !== undefined) {
        rows.push(twoColumnRow("Authorization:",request.auth.type))
    }

    if (query !== undefined && query.length > 0) {
        const queryRow = query.map((q, i)=>{
            const { key = "", value = "", description = "" } = q;

            return new TableRow({
                children:[new TableCell({
                    width: {
                        size: 20,
                        type: WidthType.PERCENTAGE
                    },
                    children:[new Paragraph(key)],
                    borders: NoBorder,
                    ...CellStyles
                }),
                new TableCell({
                    width: {
                        size: 20,
                        type: WidthType.PERCENTAGE
                    },
                    children:[new Paragraph(value === null ? "" : value)],
                    borders: NoBorder,
                    ...CellStyles
                }),
                new TableCell({
                    width: {
                        size: 60,
                        type: WidthType.PERCENTAGE
                    },
                    children:[new Paragraph(description === null ? "" : description)],
                    borders: NoBorder,
                    ...CellStyles
                })]
            })
        });
        const requestQuery = new Table({
            rows: [
                new TableRow({
                    children:[new TableCell({
                        width: {
                            size: 20,
                            type: WidthType.PERCENTAGE
                        },
                        children:[new Paragraph({
                            children: [
                                new TextRun({
                                    text: "Key",
                                    bold: true
                                })
                            ]
                        })],
                        borders: NoBorder,
                        ...CellStyles
                    }),
                    new TableCell({
                        width: {
                            size: 20,
                            type: WidthType.PERCENTAGE
                        },
                        children:[new Paragraph({
                            children: [
                                new TextRun({
                                    text: "Value",
                                    bold: true
                                })
                            ]
                        })],
                        borders: NoBorder,
                        ...CellStyles
                    }),
                    new TableCell({
                        width: {
                            size: 60,
                            type: WidthType.PERCENTAGE
                        },
                        children:[new Paragraph({
                            children: [
                                new TextRun({
                                    text: "Description",
                                    bold: true
                                })
                            ]
                        })],
                        borders: NoBorder,
                        ...CellStyles
                    })]
                }),
                ...queryRow
            ]
        });
        rows.push(twoColumnRow("Query:",requestQuery))
    }

    if (request.header.length > 0) {
        const header = new Table({
            rows: request.header.map((header, i)=>{
                return new TableRow({
                    children:[new TableCell({
                        width: {
                            size: 20,
                            type: WidthType.PERCENTAGE
                        },
                        children:[new Paragraph(header.key)],
                        borders: NoBorder,
                        ...CellStyles
                    }),
                    new TableCell({
                        width: {
                            size: 80,
                            type: WidthType.PERCENTAGE
                        },
                        children:[new Paragraph(header.value)],
                        borders: NoBorder,
                        ...CellStyles
                    })]
                })
            })
        });
        rows.push(twoColumnRow("Header:",header))
    }
    
    if (body !== undefined) {
        
        rows.push(twoColumnRow("Request Body:",_.get(body,"options.raw.language","")))
        rows.push(new TableRow({
            children: [
                new TableCell({
                    columnSpan: 2,
                    width: {
                        size: 100,
                        type: WidthType.PERCENTAGE
                    },
                    children: [new Paragraph(_.get(body,"raw",""))],
                    ...CellStyles
                })
            ]
        }) )
    }
    
    return rows;
}

function responseTable(response) {
    const { name, code, status, body, originalRequest } = response;
    const rows = requestTable(originalRequest);
    rows.unshift(twoColumnRow("Name:",name))
    rows.push(twoColumnRow("Status Code:",code+" " +status))
    rows.push(new TableRow({
        children: [
            new TableCell({
                columnSpan: 2,
                children: [new Paragraph({
                    children: [
                        new TextRun({
                            text: "Response Body:",
                            bold: true
                        })
                    ]
                })],
                ...CellStyles
            })
        ]
    }))
    rows.push(new TableRow({
        children: [
            new TableCell({
                columnSpan: 2,
                children: [new Paragraph(body)],
                ...CellStyles
            })
        ]
    }))
    rows.push(new TableRow({
        children: [
            new TableCell({
                columnSpan: 2,
                children: [new Paragraph("")],
                borders: NoBorder
            })
        ]
    }))
    return new Table({
        width: {
            size: 100,
            type: WidthType.PERCENTAGE
        },
        rows:rows
    })
}



function itemBody(item) {
    const { request, response=[] } = item;
    
    let result = itemIntro(item);
    if (request !== undefined)
    {
        result.push(new Paragraph({
            text: "Request:",
            heading: HeadingLevel.HEADING_3,
            spacing: {
                before: 200,
                after: 100
            }
        }));
        
        result.push(new Table({
            width: {
                size: 100,
                type: WidthType.PERCENTAGE
            },
            rows: requestTable(request)
        }))
    }
    if (response !== undefined && response.length > 0) {
        result.push(new Paragraph({
            text: "Response Example:",
            heading: HeadingLevel.HEADING_3,
            spacing: {
                before: 200,
                after: 100
            }
        }))

        result = result.concat(item.response.map(responseTable))
    }

    // Add break between item
    result.push(new Paragraph({
        text: "",
        spacing: {
            before: 200,
            after: 200
        }
    }))

    return result;
}

function ConvertDocx(output, json) {

    const myDoc = new Document({
        sections: json.map((page) => {

            const { item, info } = page

            const children = item.reduce((result,item,index)=>{

                if (item.item === undefined) {
                    result = result.concat(itemBody(item));
                } else {
                    result.push(new Paragraph({
                        text: item.name,
                        heading: HeadingLevel.HEADING_2,
                        spacing: {
                            before: 200
                        }
                    }))
                    const subItem = item.item.reduce((sr, sv, si) =>{
                        sr = sr.concat(itemBody(sv));
                        return sr;
                    },[])

                    result = result.concat(subItem);
                }
                return result;
            },[]);

            const sIntro = sectionIntro(info);
            return ({
                children: [
                    ...sIntro,
                    ...children
                ]
            })
        })
    })
    
    Packer.toBuffer(myDoc).then((buffer)=> {
        fs.writeFileSync(output,buffer)
    })
}
module.exports = ConvertDocx;