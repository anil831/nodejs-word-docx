const fs = require('fs');
const docx = require('docx');
const path = require("path");
const { Document, Packer, Paragraph, Table, TableRow, TableCell, WidthType, ImageRun, VerticalAlign, AlignmentType } = require('docx');
var pdf2img = require('pdf-img-convert');


// if (
//     fs.existsSync(
//       "C:/cvws_new_uploads/case_uploads/230606A0190B/employment/647efd9f9f4dc5b9555e71e8/candidatedocs"
//     )
//   ) {
//     let files = fs.readdirSync(
//         "C:/cvws_new_uploads/case_uploads/230606A0190B/employment/647efd9f9f4dc5b9555e71e8/candidatedocs"
//     );



   
//   }


(async function () {
    try{
        pdfArray = await pdf2img.convert('https://www.africau.edu/images/default/sample.pdf');
        console.log("saving");
        for (i = 0; i < pdfArray.length; i++){
          fs.writeFile("output"+i+".png", pdfArray[i], function (error) {
            if (error) { console.error("Error: " + error); }
          }); //writeFile
        } // for
    }catch(error){
console.log(error);
    }
    
  })();