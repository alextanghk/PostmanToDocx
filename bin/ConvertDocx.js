const fs = require("fs");
const util = require('util');
const path = require("path");
const docx = require("docx");

const { HeadingLevel, WidthType, BorderStyle, Document, Table, TableRow, TableCell, VerticalAlign, Paragraph, TextRun, Packer, convertInchesToTwip } = docx;
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

function myTableRow(title, content) {
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
                columnSpan: 2,
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

function RequestTable(request) {
    const rows = [
        myTableRow("URL:",request.url.raw),
        myTableRow("Method:",request.method),
    ]

    if (request.auth !== undefined) {
        rows.push(myTableRow("Authorization:",request.auth.type))
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
                                        ...NoBorder
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
                                        ...NoBorder
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
                                        ...NoBorder
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
                                        ...NoBorder
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
        rows.push(myTableRow("Body:",request.body.options.raw.language))

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
function ConvertDocx(output, json) {
    
    
    
    const myDoc = new Document({
        sections: json.map((page) => {
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
                                            borders: {top: { style: BorderStyle.SINGLE, size: 2, color: "000000" }},
                                            ...CellStyles
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
        fs.writeFileSync(output,buffer)
    })
}
module.exports = ConvertDocx;