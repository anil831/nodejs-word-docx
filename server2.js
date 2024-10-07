const fs = require('fs');
const docx = require('docx');

const { Document, Packer, Table, TableRow, TableCell, Paragraph,WidthType,VerticalAlign } = docx;



const doc = new Document({
    sections: [
        {
            children:[
                new Table({
                    rows: [
                      new TableRow({
                        children: [
                          new TableCell({
                            children: [new Paragraph({text:"Name of Applicant"})],
                            shading: { fill: "BBBBBB" },
                            margins: {
                                top: 100,
                                bottom: 100, 
                                left: 100, 
                                right: 100, 
                              },
                              width: {
                                size: 22, // Set cell width
                                type: WidthType.PERCENTAGE,
                              },
                              properties: {
                                columnSpan: 1, // Span one column
                              },
                          }),
                          new TableCell({
                            children: [new Paragraph("Anil kumar  Hanumantha Rao Tanneeru")],
                            margins: {
                                top: 100,
                                bottom: 100, 
                                left: 100, 
                                right: 100, 
                              },
                              width: {
                                size: 28, // Set cell width
                                type: WidthType.PERCENTAGE,
                              },
                              properties: {
                                columnSpan: 1, // Span one column
                              },
                          }),
                          new TableCell({
                            children: [new Paragraph("Applicant ID")],
                            shading: { fill: "BBBBBB" },
                            margins: {
                                top: 100,
                                bottom: 100, 
                                left: 100, 
                                right: 100, 
                              },
                              width: {
                                size: 22, // Set cell width
                                type: WidthType.PERCENTAGE,
                              },
                              properties: {
                                columnSpan: 1, // Span one column
                              },
                          }),
                          new TableCell({
                            children: [new Paragraph("12345821")],
                            margins: {
                                top: 100,
                                bottom: 100, 
                                left: 100, 
                                right: 100, 
                              },
                              width: {
                                size: 28, // Set cell width
                                type: WidthType.PERCENTAGE,
                              },
                              properties: {
                                columnSpan: 1, // Span one column
                              },
                          })
                        ]
                      }),
                      new TableRow({
                        children: [
                          new TableCell({
                            children: [new Paragraph("Employee ID")],
                            shading: { fill: "BBBBBB" },
                            margins: {
                                top: 100,
                                bottom: 100, 
                                left: 100, 
                                right: 100, 
                              },
                              width: {
                                size: 22, // Set cell width
                                type: WidthType.PERCENTAGE,
                              },
                              properties: {
                                columnSpan: 1, // Span one column
                              },
                          }),
                          new TableCell({
                            children: [new Paragraph("emp8484")],
                            margins: {
                                top: 100,
                                bottom: 100, 
                                left: 100, 
                                right: 100, 
                              },
                              width: {
                                size: 38, // Set cell width
                                type: WidthType.PERCENTAGE,
                              },
                              properties: {
                                columnSpan: 1, // Span one column
                              },
                          }),
                          new TableCell({
                            children: [new Paragraph("Date of Birth")],
                            shading: { fill: "BBBBBB" },
                            margins: {
                                top: 100,
                                bottom: 100, 
                                left: 100, 
                                right: 100, 
                              },
                              width: {
                                size: 22, // Set cell width
                                type: WidthType.PERCENTAGE,
                              },
                              properties: {
                                columnSpan: 1, // Span one column
                              },
                        }),
                          new TableCell({
                            children: [new Paragraph("01-12-1996")],
                            margins: {
                                top: 100,
                                bottom: 100, 
                                left: 100, 
                                right: 100, 
                              },
                              width: {
                                size: 28, // Set cell width
                                type: WidthType.PERCENTAGE,
                              },
                              properties: {
                                columnSpan: 1, // Span one column
                              },
                        })
                        ]
                      }),
                      new TableRow({
                        children: [
                          new TableCell({
                            children: [new Paragraph("Date of joining")],
                            shading: { fill: "BBBBBB" },
                            margins: {
                                top: 100,
                                bottom: 100, 
                                left: 100, 
                                right: 100, 
                              },
                              width: {
                                size: 22, // Set cell width
                                type: WidthType.PERCENTAGE,
                              },
                              properties: {
                                columnSpan: 1, // Span one column
                              },

                          }),
                          new TableCell({
                            children: [new Paragraph("01-12-2020")],
                            columnSpan:3,
                            margins: {
                                top: 100,
                                bottom: 100, 
                                left: 100, 
                                right: 100, 
                              },
                              width: {
                                size: 78, // Set cell width
                                type: WidthType.PERCENTAGE,
                              },
                              properties: {
                                columnSpan: 1, // Span one column
                              },
                          })
                        ]
                      })
                    ]
                  })
            ]
        }
    ]
})

// Save the document
Packer.toBuffer(doc).then(buffer => {
    fs.writeFileSync('output_table.docx', buffer);
    console.log('Document saved!');
  }).catch(error => {
    console.error('Error:', error);
  });


