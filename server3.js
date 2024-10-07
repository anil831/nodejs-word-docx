const fs = require('fs');
const docx = require('docx');
const pdf2img = require("pdf2img")
const { Document, Packer, Table, TableRow, TableCell, WidthType } = docx;

// Create a new document

pdf2img.setOptions({
  type:'jpg',
  size:1024,
        density:600,
        outputdir:filePath + "/",
  outputname:output,
  page:null,
  quality:100	
});
pdf2img.convert(input,function(err,info){
        if(err){
    resolve(true)	    
    console.log('error covnerting',err)	    
  }else{
    allJpegs.push(filePath) 	    
    resolve(true)	    
    console.log('converted successfully',info)	    
  }		
}) 

const doc = new Document({
    sections:[
        {
            children:[
                new Table({
                    rows: [
                      new TableRow({
                        children: [
                          new TableCell({
                            children: [new docx.Paragraph("Row 1, Cell 1")],
                            width: {
                              size: 33,
                              type: WidthType.PERCENTAGE,
                            },
                          }),
                          new TableCell({
                            children: [new docx.Paragraph("Row 1, Cell 2")],
                            width: {
                              size: 33,
                              type: WidthType.PERCENTAGE,
                            },
                          }),
                          new TableCell({
                            children: [new docx.Paragraph("Row 1, Cell 3")],
                            width: {
                              size: 33,
                              type: WidthType.PERCENTAGE,
                            },
                          }),
                        ],
                      }),
                      new TableRow({
                        children: [
                          new TableCell({
                            children: [new docx.Paragraph("Row 2, Cell 1 (Spanning Two Rows)")],
                            rowSpan: 2, // Simulate rowspan effect by merging cells
                            width: {
                              size: 33,
                              type: WidthType.PERCENTAGE,
                            },
                          }),
                          new TableCell({
                            children: [new docx.Paragraph("Row 2, Cell 2")],
                            width: {
                              size: 33,
                              type: WidthType.PERCENTAGE,
                            },
                          }),
                          new TableCell({
                            children: [new docx.Paragraph("Row 2, Cell 3")],
                            width: {
                              size: 33,
                              type: WidthType.PERCENTAGE,
                            },
                          }),
                        ],
                      }),
                      new TableRow({
                        children: [
                          new TableCell({
                            children: [new docx.Paragraph("Row 3, Cell 1")],
                            width: {
                              size: 33,
                              type: WidthType.PERCENTAGE,
                            },
                          }),
                          new TableCell({
                            children: [new docx.Paragraph("Row 3, Cell 2")],
                            width: {
                              size: 33,
                              type: WidthType.PERCENTAGE,
                            },
                          }),
                          new TableCell({
                            children: [new docx.Paragraph("Row 3, Cell 3")],
                            width: {
                              size: 33,
                              type: WidthType.PERCENTAGE,
                            },
                          }),
                        ],
                      }),
                    ],
                  })
            ]
        }
    ]
});


// Save the document
Packer.toBuffer(doc).then(buffer => {
    fs.writeFileSync('output_table3.docx', buffer);
    console.log('Document saved!');
  }).catch(error => {
    console.error('Error:', error);
  });

