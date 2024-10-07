const fs = require('fs');
const docx = require('docx');
const pdfPoppler = require('pdf-poppler');

const pdfPath = "C:/Users/Anil/Downloads/pdf-test.pdf";
const outputDir = __dirname; // Specify the directory where you want to save the images

const opts = {
  format: 'png', // Output format (you can also use 'jpg', 'jpeg', 'tiff', etc.)
  out_dir: outputDir,
  out_prefix: "signature", // Prefix for image files
  page: 1, // Page number to extract
};

pdfPoppler.convert(pdfPath, opts)
  .then(() => {

    const { Document, Packer, Paragraph, Table, TableRow, TableCell, WidthType, ImageRun, VerticalAlign, AlignmentType } = require('docx');
    // Create a new document

    const doc = new Document({
      sections: [
        {
          properties: {},
          children: [
            new Table({
              rows: [
                new TableRow({
                  children: [
                    new TableCell({
                      children: [new Paragraph("Row 1, Cell 1")],
                      width: {
                        size: 50,
                        type: WidthType.PERCENTAGE,
                      },
                    }),

                    new TableCell({
                      children: [
                        new Paragraph({
                          children: [
                            new ImageRun({
                              data: fs.readFileSync("C:/Users/Anil/Desktop/anil/learning/nodejs-docx/signature-1.png"), // Use the correct filename here
                              transformation: { width: 300, height: 300 },
                            })
                          ],
                          alignment: docx.AlignmentType.CENTER

                        }),
                      ],
                      width: {
                        size: 50,
                        type: WidthType.PERCENTAGE,
                      },
                      // Center the content within the TableCell
                      alignment: {
                        vertical: VerticalAlign.CENTER,
                        horizontal: AlignmentType.CENTER,
                      }
                    }),
                  ],
                }),
              ],
            }),
          ],
        },
      ],
    });
    // Save the document
    Packer.toBuffer(doc).then(buffer => {
      fs.writeFileSync('output_table5.docx', buffer);
      console.log('Document saved!');
    }).catch(error => {
      console.error('Error:', error);
    });

  })
  .catch((error) => {
    console.error('Error:', error);
  });
