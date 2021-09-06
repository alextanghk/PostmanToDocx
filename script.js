const fs = require("fs");
const util = require('util');
const path = require("path");
const docx = require("docx");

const yargs = require('yargs/yargs')
const { hideBin } = require('yargs/helpers');
const { HeadingLevel, WidthType, BorderStyle } = require("docx");

const { Document, Table, TableRow, TableCell, VerticalAlign, Paragraph, TextRun, Packer, convertInchesToTwip } = docx;

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

const apiJson = require(source);

const CellStyles = {
    margins: {
        top: 50,
        bottom: 50,
        left: 50,
        right: 50,
    }
}

const RequestTable = (request) => {
    const rows = [
        new TableRow({
            children:[
                new TableCell({
                    width: {
                        size: 20,
                        type: WidthType.PERCENTAGE
                    },
                    children:[new Paragraph({
                        children: [
                            new TextRun({
                                text: "URL:",
                                bold: true
                            })
                        ]
                    })],
                    ...CellStyles
                }),
                new TableCell({
                    columnSpan: 2,
                    width: {
                        size: 80,
                        type: WidthType.PERCENTAGE
                    },
                    children:[new Paragraph(request.url.raw)],
                    ...CellStyles
                })
            ]
        }),
        new TableRow({
            children:[
                new TableCell({
                    width: {
                        size: 20,
                        type: WidthType.PERCENTAGE
                    },
                    children:[new Paragraph({
                        children: [
                            new TextRun({
                                text: "Method:",
                                bold: true
                            })
                        ]
                    })],
                    ...CellStyles
                }),
                new TableCell({
                    columnSpan: 2,
                    width: {
                        size: 80,
                        type: WidthType.PERCENTAGE
                    },
                    children:[new Paragraph(request.method)],
                    ...CellStyles
                })
            ]
        })
    ]

    if (request.auth !== undefined) {
        rows.push(new TableRow({
            children:[
                new TableCell({
                    width: {
                        size: 20,
                        type: WidthType.PERCENTAGE
                    },
                    children:[new Paragraph({
                        children: [
                            new TextRun({
                                text: "Authorization:",
                                bold: true
                            })
                        ]
                    })],
                    ...CellStyles
                }),
                new TableCell({
                    columnSpan: 2,
                    width: {
                        size: 80,
                        type: WidthType.PERCENTAGE
                    },
                    children:[new Paragraph(request.auth.type)],
                    ...CellStyles
                })
            ]
        }));
    }

    if (request.url.query !== undefined) {
        rows.push(new TableRow({
            children:[
                new TableCell({
                    width: {
                        size: 20,
                        type: WidthType.PERCENTAGE
                    },
                    children:[new Paragraph({
                        children: [
                            new TextRun({
                                text: "Queries:",
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
                    children:[new Table({
                        rows: request.url.query.map((query, i)=>{
                            return new TableRow({
                                children:[new TableCell({
                                    width: {
                                        size: 20,
                                        type: WidthType.PERCENTAGE
                                    },
                                    children:[new Paragraph(query.key)],
                                    borders: {
                                        top: { style: i == 0 ? BorderStyle.NONE : BorderStyle.SINGLE, size: 2, color: "000000" },
                                        left: { style: BorderStyle.NONE, size: 0, color: "000000" },
                                        bottom: { style: BorderStyle.NONE, size: 2, color: "000000" },
                                        right: { style: BorderStyle.SINGLE, size: 2, color: "000000" },
                                    },
                                    ...CellStyles
                                }),
                                new TableCell({
                                    width: {
                                        size: 80,
                                        type: WidthType.PERCENTAGE
                                    },
                                    children:[new Paragraph(query.description)],
                                    borders: {
                                        top: { style: i == 0 ? BorderStyle.NONE : BorderStyle.SINGLE, size: 2, color: "000000" },
                                        left: { style: BorderStyle.NONE, size: 0, color: "000000" },
                                        bottom: { style: BorderStyle.NONE, size: 2, color: "000000" },
                                        right: { style: BorderStyle.NONE, size: 0, color: "000000" },
                                    },
                                    ...CellStyles
                                })]
                            })
                        })
                    })]
                })
            ]
        }))
    }

    if (request.header.length > 0) {
        rows.push(new TableRow({
            children:[
                new TableCell({
                    width: {
                        size: 20,
                        type: WidthType.PERCENTAGE
                    },
                    children:[new Paragraph({
                        children: [
                            new TextRun({
                                text: "Headers:",
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
                    children:[new Table({
                        rows: request.header.map((header, i)=>{
                            return new TableRow({
                                children:[new TableCell({
                                    width: {
                                        size: 20,
                                        type: WidthType.PERCENTAGE
                                    },
                                    children:[new Paragraph(header.key)],
                                    borders: {
                                        top: { style: i == 0 ? BorderStyle.NONE : BorderStyle.SINGLE, size: 2, color: "000000" },
                                        left: { style: BorderStyle.NONE, size: 0, color: "000000" },
                                        bottom: { style: BorderStyle.NONE, size: 2, color: "000000" },
                                        right: { style: BorderStyle.SINGLE, size: 2, color: "000000" },
                                    },
                                    ...CellStyles
                                }),
                                new TableCell({
                                    width: {
                                        size: 80,
                                        type: WidthType.PERCENTAGE
                                    },
                                    children:[new Paragraph("<"+header.type+">")],
                                    borders: {
                                        top: { style: i == 0 ? BorderStyle.NONE : BorderStyle.SINGLE, size: 2, color: "000000" },
                                        left: { style: BorderStyle.NONE, size: 0, color: "000000" },
                                        bottom: { style: BorderStyle.NONE, size: 2, color: "000000" },
                                        right: { style: BorderStyle.NONE, size: 0, color: "000000" },
                                    },
                                    ...CellStyles
                                })]
                            })
                        })
                    })]
                })
            ]
        }))
    }
    
    if (request.body !== undefined) {
        rows.push(new TableRow({
            children:[
                new TableCell({
                    width: {
                        size: 20,
                        type: WidthType.PERCENTAGE
                    },
                    children:[new Paragraph({
                        children: [
                            new TextRun({
                                text: "Body:",
                                bold: true
                            })
                        ]
                    })],
                    ...CellStyles
                }),
                new TableCell({
                    columnSpan: 2,
                    width: {
                        size: 80,
                        type: WidthType.PERCENTAGE
                    },
                    children:[new Paragraph(request.body.options.raw.language)],
                    ...CellStyles
                })
            ]
        }))

        rows.push(new TableRow({
            children:[
                new TableCell({
                    columnSpan: 3,
                    children:[new Paragraph({
                        text: request.body.raw,
                    })],
                    ...CellStyles
                }),
            ]
        }))
    }
    

    
    return new Table({
        width: {
            size: 100,
            type: WidthType.PERCENTAGE
        },
        rows: rows
    })
}

const myDoc = new Document({
    sections: apiJson.map((page) => {
        const children = page.item.reduce((result,item,index)=>{
            result = result.concat([new Paragraph({
                heading: HeadingLevel.HEADING_2,
                spacing: {
                    after: 100
                },
                border: {
                    top: { value: "single", space:50, size: 2, color: "000000" },
                },
                children: [
                    new TextRun({
                        text: item.name,
                        bold: true
                    })
                ]
            }),new Paragraph({
                children: item.request.description === undefined ? [new TextRun("")]: item.request.description.split("\n").map((desc)=>{
                    return new TextRun({
                        text: desc,
                        break: 1
                    })
                })
            })])
            
            result.push(new Paragraph({
                text: "Request:",
                heading: HeadingLevel.HEADING_3,
                spacing: {
                    before: 200,
                    after: 100
                }
            }));
            result.push(RequestTable(item.request));

            if (item.response.length > 0) {
                result.push(new Paragraph({
                    text: "",
                    spacing: {
                        before: 200,
                        after: 200
                    }
                }))
    
                result.push(new Paragraph({
                    text: "Response Example:",
                    heading: HeadingLevel.HEADING_3,
                    spacing: {
                        before: 200,
                        after: 200
                    }
                }))
    
                result = result.concat(item.response.map((response)=>{
                    return (new Table({
                        width: {
                            size: 100,
                            type: WidthType.PERCENTAGE
                        },
                        rows:[
                            new TableRow({
                                children: [
                                    new TableCell({
                                        width: {
                                            width: 75,
                                            type: WidthType.PERCENTAGE
                                        },
                                        children: [new Paragraph({
                                            children: [
                                                new TextRun({
                                                    text: response.name,
                                                    bold: true
                                                })
                                            ]
                                        })],
                                        borders: {top: { style: BorderStyle.SINGLE, size: 2, color: "000000" }},
                                        ...CellStyles
                                    }),
                                    new TableCell({
                                        width: {
                                            width: 10,
                                            type: WidthType.PERCENTAGE
                                        },
                                        children: [
                                            new Paragraph({
                                                children: [
                                                    new TextRun({
                                                        text: "Status Code:",
                                                        bold: true
                                                    })
                                                ]
                                            })
                                        ],
                                        borders: {top: { style: BorderStyle.SINGLE, size: 2, color: "000000" }}
                                    }),
                                    new TableCell({
                                        width: {
                                            width: 15,
                                            type: WidthType.PERCENTAGE
                                        },
                                        children: [new Paragraph(response.code+" " +response.status)],
                                        borders: {top: { style: BorderStyle.SINGLE, size: 2, color: "000000" }},
                                        ...CellStyles
                                    })
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({
                                        columnSpan: 3,
                                        children: [new Paragraph(response.body)],
                                        ...CellStyles
                                    })
                                ]
                            }),
                            new TableRow({
                                children: [
                                    new TableCell({
                                        columnSpan: 3,
                                        children: [new Paragraph("")],
                                        borders: {
                                            right: { style: BorderStyle.NONE, size: 2, color: "000000" },
                                            bottom: { style: BorderStyle.NONE, size: 2, color: "000000" },
                                            left: { style: BorderStyle.NONE, size: 2, color: "000000" }
                                        }
                                    })
                                ]
                            })
                        ]
                    }))
                }))
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
        },[])
        
        children.unshift(
            new Paragraph({
                children:[
                    new TextRun({
                        text: page.info.name,
                        bold: true
                    })
                ],
                heading: HeadingLevel.HEADING_1,
                spacing: {
                    after: 100
                }
            }),
            new Paragraph({
                children: page.info.description === undefined ? [new TextRun("")]: page.info.description.split("\n").map((desc)=>{
                        return new TextRun({
                            text: desc,
                            break: 1
                        })
                    }),
                spacing: {
                    after: 500
                }
            })
        );
        return ({
            children: children
        })
    })
})

Packer.toBuffer(myDoc).then((buffer)=> {
    fs.writeFileSync(path.join(output,basename.replace(".json",".docx")),buffer)
})