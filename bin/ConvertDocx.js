const fs = require("fs");
const docx = require("docx");

const { HeadingLevel, WidthType, BorderStyle, Document, Table, TableRow, TableCell, Paragraph, TextRun, Packer } = docx;
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

function sectionIntro(name, description) {
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
    }),new Paragraph({
        children: desc.length === 0 ? [new TextRun("")]: desc.map((t)=>{
            return new TextRun({
                text: t,
                break: 1
            })
        }),
        spacing: {
            after: 300
        }
    })]
}

function requestTableRow(title, content) {
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

function requestTable(request) {
    const rows = [
        requestTableRow("URL:",request.url.raw),
        requestTableRow("Method:",request.method),
    ]

    if (request.auth !== undefined) {
        rows.push(requestTableRow("Authorization:",request.auth.type))
    }

    if (request.url.query !== undefined) {
        const query = new Table({
            rows: request.url.query.map((query, i)=>{
                return new TableRow({
                    children:[new TableCell({
                        width: {
                            size: 20,
                            type: WidthType.PERCENTAGE
                        },
                        children:[new Paragraph(query.key)],
                        borders: NoBorder,
                        ...CellStyles
                    }),
                    new TableCell({
                        width: {
                            size: 80,
                            type: WidthType.PERCENTAGE
                        },
                        children:[new Paragraph(query.description)],
                        borders: NoBorder,
                        ...CellStyles
                    })]
                })
            })
        });
        rows.push(requestTableRow("Queries:",query))
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
                        children:[new Paragraph("<"+header.type+">")],
                        borders: NoBorder,
                        ...CellStyles
                    })]
                })
            })
        });
        rows.push(requestTableRow("Header:",header))
    }
    
    if (request.body !== undefined) {
        rows.push(requestTableRow("Body:",request.body.options.raw.language))
    }
    
    return new Table({
        width: {
            size: 100,
            type: WidthType.PERCENTAGE
        },
        rows: rows
    })
}

function responseTable(response) {
    const { name, code, status, body } = response;
    return new Table({
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
                                    text: name,
                                    bold: true
                                })
                            ]
                        })],
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
                        ...CellStyles
                    }),
                    new TableCell({
                        width: {
                            width: 15,
                            type: WidthType.PERCENTAGE
                        },
                        children: [new Paragraph(code+" " +status)],
                        ...CellStyles
                    })
                ]
            }),
            new TableRow({
                children: [
                    new TableCell({
                        columnSpan: 3,
                        children: [new Paragraph(body)],
                        ...CellStyles
                    })
                ]
            }),
            new TableRow({
                children: [
                    new TableCell({
                        columnSpan: 3,
                        children: [new Paragraph("")],
                        borders: NoBorder
                    })
                ]
            })
        ]
    })
}

function itemIntro(name, description) {
    const desc = description === undefined ? [] : description.split("\n");
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
    }),new Paragraph({
        children: desc.length === 0 ? [new TextRun("")]: desc.map((t)=>{
            return new TextRun({
                text: t,
                break: 1
            })
        })
    })]
}

function ConvertDocx(output, json) {

    const myDoc = new Document({
        sections: json.map((page) => {
            const children = page.item.reduce((result,item,index)=>{
                result = result.concat(itemIntro(item.name,item.request.description))
                
                result.push(new Paragraph({
                    text: "Request:",
                    heading: HeadingLevel.HEADING_3,
                    spacing: {
                        before: 200,
                        after: 100
                    }
                }));
                result.push(requestTable(item.request));
    
                if (item.response.length > 0) {
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
            },[])
            const sIntro = sectionIntro(page.info.name, page.info.description);
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