const { workerData, parentPort } = require("worker_threads");
const fs = require("fs");
const path = require("path");
const moment = require("moment");
const docx = require("docx");
// const pdfPoppler = require('pdf-poppler');
const pdf2pic = require("pdf2pic")

const {
  Document,
  Table,
  TableRow,
  TableCell,
  Paragraph,
  WidthType,
  AlignmentType,
  TextRun,
  BorderStyle,
  VerticalAlign,
  UnderlineType,
  ExternalHyperlink,
  InternalHyperlink,
  Bookmark,
  SymbolRun,
  convertInchesToTwip,
  ShadingType,
  PageBreak,
  ImageRun,
} = docx;

const getBlankLine = function () {
  const blankLine = new Paragraph("");
  return blankLine;
};

let blankLine = getBlankLine();

const getTableHeading = async function (heading, bookmarkid, docCount) {
  // Create a paragraph with the text run
  try {

    if (bookmarkid && docCount && docCount === 1) {
      const paragraph = new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new Bookmark({
            id: bookmarkid,
            children: [
              new TextRun({
                text: heading + " " + docCount,
                bold: true,
                font: "Calibri",
                size: 25,
              }),
            ],
          }),
        ],
        pageBreakBefore: true,
      });

      return paragraph;
    } else {
      const paragraph = new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({
            text: heading + " " + docCount,
            bold: true,
            font: "Calibri",
            size: 25,
          }),
        ],
      });

      return paragraph;
    }
  } catch (error) {
    console.log(error);
  }
};


let writeReportTitleTable = function () {
  return new Promise((resolve, reject) => {
    const reportTitleTable = new docx.Table({
      rows: [new docx.TableRow({
        children: [
          new docx.TableCell({
            width: {
              size: 9000,
              type: docx.WidthType.DXA,
            },
            shading: {
              fill: "DCDCDC",
              // type: docx.ShadingType.REVERSE_DIAGONAL_STRIPE,
              color: "auto",
            },
            children: [new docx.Paragraph({
              children: [
                new docx.TextRun({
                  text: "BACKGROUND VERIFICATION FINAL REPORT",
                  allCaps: true,
                  font: "Calibri",
                  bold: true,
                  size: 30
                })
              ],
              alignment: docx.AlignmentType.CENTER
            })]
          })
        ]

      })
      ]
    })
    resolve(reportTitleTable)
  })
}

let writeClientNameTable = function (clientName) {
  try {


    console.log("In write client name table client name is............................................................ ", clientName)
    return new Promise((resolve, reject) => {
      const clientNameTable = new docx.Table({
        rows: [new docx.TableRow({
          children: [
            new docx.TableCell({
              width: {
                size: 9000,
                type: docx.WidthType.DXA,
              },
              shading: {
                fill: "DCDCDC",
                // type: docx.ShadingType.REVERSE_DIAGONAL_STRIPE,
                color: "auto",
              },
              children: [new docx.Paragraph({
                children: [
                  new docx.TextRun({
                    text: clientName,
                    allCaps: true,
                    font: "Calibri",
                    bold: true,
                    size: 30
                  })
                ],
                alignment: docx.AlignmentType.CENTER
              })]
            }),
          ]

        })
        ]
      })
      console.log("Returning clientNameTable", clientNameTable)
      resolve(clientNameTable)
    })
  } catch (error) {
    console.log(error);
  }
}


const getApplicantTable = async function (caseDetails, personalDetails) {
  try {

    const dateOfBirth = getDateString(personalDetails.dateofbirth);
    const dateOfJoining = getDateString(personalDetails.dateofjoining);

    const applicantTable = new Table({
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new Bookmark({
                      id: "anchorForChapter1",
                      children: [
                        new TextRun({
                          text: "Name of Applicant",
                          bold: true,
                          font: "Calibri",
                          size: 22,
                        }),
                      ],
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                bottom: 100,
                left: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(22), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: caseDetails.candidateName,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                bottom: 100,
                left: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(28), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Applicant ID",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                bottom: 100,
                left: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(22), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: caseDetails.caseId,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                bottom: 100,
                left: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(28), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Employee ID",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                bottom: 100,
                left: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(22), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: personalDetails?.empid || workerData.caseDetails?.employeeId,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                bottom: 100,
                left: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(28), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Date of Birth",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                bottom: 100,
                left: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(22), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: dateOfBirth,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                bottom: 100,
                left: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(28), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Date of joining",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                bottom: 100,
                left: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(22), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: dateOfJoining,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              columnSpan: 3,
              margins: {
                top: 100,
                bottom: 100,
                left: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(78), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
      ],
    });
    return applicantTable;
  } catch (error) {
    console.log(error);
  }

};

const getCaseDetailsTable = async function (caseDetails, profileOrPackageName) {
  try {
    const dateOfInitiation = getDateString(caseDetails.initiationDate);
    const interimReportDate = getDateString(caseDetails.interimReportDate);
    const supplementaryReportDate = getDateString(caseDetails.supplementaryReportDate);
    const lastInsufficiencyClearedDate = getDateString(
      caseDetails.lastInsufficiencyClearedDate
    );

    const firstInsufficiencyRaisedDate = getDateString(
      caseDetails.firstInsufficiencyRaisedDate
    );
    const outputqcCompletionDate = getDateString(
      caseDetails.outputqcCompletionDate
    );
    const caseDetailsTable = new Table({
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Case Details",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              columnSpan: 4,
              width: {
                size: getWidthPercentage(100), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
              margins: {
                top: 100,
                bottom: 100,
                left: 100,
                right: 100,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Date of Initiation",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(22), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: dateOfInitiation,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(28), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),

            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Case Reference",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                    new TextRun({
                      break: true,
                    }),
                    new TextRun({
                      text: "Number/URN No.",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                  spacing: {
                    line: 300, // apply space between lines
                  },
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(22), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: VerticalAlign.CENTER,
                  children: [
                    new TextRun({
                      text: caseDetails.caseId,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(28), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Fresher/Lateral",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(22), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: workerData.caseDetails.subclient.name,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(28), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),

            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Client/Scope",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(22), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: VerticalAlign.CENTER,
                  children: [
                    new TextRun({
                      text: caseDetails.subclient?.client?.name
                        ? caseDetails.subclient?.client?.name
                        : caseDetails.subclient?.name,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(28), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Interim Report",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                    new TextRun({
                      break: true,
                    }),
                    new TextRun({
                      text: "Date",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                  spacing: {
                    line: 300,
                  },
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(22), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: interimReportDate,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(28), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),

            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Interim Report",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                    new TextRun({
                      break: true,
                    }),
                    new TextRun({
                      text: "SLA (No of Days)",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                  spacing: {
                    line: 300,
                  },
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(22), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: VerticalAlign.CENTER,
                  children: [
                    new TextRun({
                      text: "",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(28), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "CEA Initiation",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                    new TextRun({
                      break: true,
                    }),
                    new TextRun({
                      text: "Date",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                  spacing: {
                    line: 300,
                  },
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(22), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: firstInsufficiencyRaisedDate,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(28), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),

            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "CE Insufficiency",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                    new TextRun({
                      break: true,
                    }),
                    new TextRun({
                      text: "Clearance Date",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(22), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: VerticalAlign.CENTER,
                  children: [
                    new TextRun({
                      text: lastInsufficiencyClearedDate,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(28), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Final Report Date",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(22), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: outputqcCompletionDate,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(28), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),

            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Final Report SLA",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                    new TextRun({
                      break: true,
                    }),
                    new TextRun({
                      text: "(No of Days)",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                  spacing: {
                    line: 300,
                  },
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(22), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: VerticalAlign.CENTER,
                  children: [
                    new TextRun({
                      text: "",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(28), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Re-initiation Date",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(22), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(28), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),

            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Supplementary",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                    new TextRun({
                      break: true,
                    }),
                    new TextRun({
                      text: "Report Date",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(22), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: VerticalAlign.CENTER,
                  children: [
                    new TextRun({
                      text: supplementaryReportDate,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(28), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Supplementary",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                    new TextRun({
                      break: true,
                    }),
                    new TextRun({
                      text: "Report SLA",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                    new TextRun({
                      break: true,
                    }),
                    new TextRun({
                      text: "(No of Days)",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(22), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(28), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),

            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Colour Code",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                    new TextRun({
                      break: true,
                    }),
                    new TextRun({
                      text: "(Red/Amber/",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                    new TextRun({
                      break: true,
                    }),
                    new TextRun({
                      text: "Green/IRCEP/Stop",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                    new TextRun({
                      break: true,
                    }),
                    new TextRun({
                      text: "Case)",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                  spacing: {
                    line: 300,
                  },
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(22), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: VerticalAlign.CENTER,
                  children: [
                    new TextRun({
                      text: caseDetails.gardeName,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(28), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
      ],
    });

    return caseDetailsTable;
  } catch (error) {
    console.log(error);
  }
};

const getWidthPercentage = (percentage) => {
  try {

    // Calculate the width in twips (1/20th of a point) based on a percentage of the page width
    const pageWidthTwip = convertInchesToTwip(8.5); // Assuming standard letter size page width
    return (percentage / 100) * pageWidthTwip;
  } catch (error) {
    console.log(error);
  }

};

const getColorCodeTable = async function () {
  try {
    const colorCodeTable = new Table({
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new SymbolRun({
                      char: "25A0",
                      bold: true,
                      color: "#FF0000",
                      size: 100,
                    }),
                  ],
                }),
              ],
              width: {
                size: getWidthPercentage(5),
                type: WidthType.DXA,
              },
              borders: {
                right: {
                  style: BorderStyle.NONE,
                  color: "#FFFFFF",
                  space: 1,
                },
              },
              margins: {
                top: 100,
                bottom: 100,
                left: 100,
                right: 100,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Major", // Add two spaces before "Major",
                      font: 'Calibri',
                      size: 22
                    }),
                    new TextRun({
                      break: true, // This will insert a line break
                    }),
                    new TextRun({
                      text: "Discrepancy",
                      font: 'Calibri',
                      size: 22
                    }),
                    new TextRun({
                      break: true, // This will insert a line break
                    }),
                    new TextRun({
                      text: " (RED)",
                      font: 'Calibri',
                      size: 22
                    }),
                  ],
                  spacing: {
                    line: 300,
                  },
                }),
              ],
              width: {
                size: getWidthPercentage(30),
                type: WidthType.DXA,
              },
              borders: {
                left: {
                  style: BorderStyle.NONE,
                  color: "#FFFFFF",
                  space: 1,
                },
              },
              margins: {
                top: 100,
                bottom: 100,
                left: 100,
                right: 100,
              },
            }),

            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new SymbolRun({
                      char: "25A0",
                      bold: true,
                      color: "#FFBF00",
                      size: 100,
                    }),
                  ],
                }),
              ],
              width: {
                size: getWidthPercentage(5),
                type: WidthType.DXA,
              },
              borders: {
                right: {
                  style: BorderStyle.NONE,
                  color: "#FFFFFF",
                  space: 1,
                },
              },
              margins: {
                top: 100,
                bottom: 100,
                left: 100,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Inaccessible for verification / Unable to verify / ", // Add two spaces before "Major"
                      font: 'Calibri',
                      size: 22
                    }),
                    new TextRun({
                      break: true, // This will insert a line break
                    }),
                    new TextRun({
                      text: " Inputs required/ Minor Discrepancy",
                      font: 'Calibri',
                      size: 22
                    }),
                    new TextRun({
                      break: true, // This will insert a line break
                    }),
                    new TextRun({
                      text: " (AMBER)",
                      font: 'Calibri',
                      size: 22
                    }),
                  ],
                  spacing: {
                    line: 300,
                  },
                }),
              ],
              width: {
                size: getWidthPercentage(40),
                type: WidthType.DXA,
              },
              borders: {
                left: {
                  style: BorderStyle.NONE,
                  color: "#FFFFFF",
                  space: 1,
                },
              },
            }),

            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new SymbolRun({
                      char: "25A0",
                      bold: true,
                      color: "#00FF00",
                      size: 100,
                    }),
                  ],
                }),
              ],
              width: {
                size: getWidthPercentage(5),
                type: WidthType.DXA,
              },
              borders: {
                right: {
                  style: BorderStyle.NONE,
                  color: "#FFFFFF",
                  space: 1,
                },
              },
              margins: {
                top: 100,
                bottom: 100,
                left: 100,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "  Verified  ", // Add two spaces before "Major"
                      font: 'Calibri',
                      size: 22
                    }),
                    new TextRun({
                      break: true, // This will insert a line break
                    }),
                    new TextRun({
                      text: " Report",
                      font: 'Calibri',
                      size: 22
                    }),
                    new TextRun({
                      text: true, // This will insert a line break
                    }),
                    new TextRun({
                      text: " (GREEN)",
                      font: 'Calibri',
                      size: 22
                    }),
                  ],
                  spacing: {
                    line: 300,
                  },
                }),
              ],
              width: {
                size: getWidthPercentage(15),
                type: WidthType.DXA,
              },
              borders: {
                left: {
                  style: BorderStyle.NONE,
                  color: "#FFFFFF",
                  space: 1,
                },
              },
              margins: {
                top: 100,
                bottom: 100,
                left: 100,
                right: 100,
              },
            }),
          ],
        }),
      ],
    });

    return colorCodeTable;
  } catch (error) {
    console.log(error);
  }
};

const getBackGroundVerificationReportTable = async function () {
  try {

    const backGroundVerificationReportTable = new Table({
      borders: {
        top: { style: docx.BorderStyle.SINGLE, size: 6, color: "000000" },
        bottom: { style: docx.BorderStyle.SINGLE, size: 6, color: "000000" },
        left: { style: docx.BorderStyle.SINGLE, size: 6, color: "000000" },
        right: { style: docx.BorderStyle.SINGLE, size: 6, color: "000000" },
      },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "BACKGROUND VERIFICATION REPORT – INTERIM REPORT / FINAL REPORT / SUPPLEMENTARY REPORT",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              columnSpan: 4,
              width: {
                size: getWidthPercentage(100), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
              margins: {
                top: 100,
                bottom: 100,
                left: 100,
                right: 100,
              },
            }),
          ],
        }),
      ],
    });

    return backGroundVerificationReportTable;
  } catch (error) {
    console.log(error);
  }

};

const getExecutiveSummaryHeading = async function () {
  try {
    // Create a paragraph with the text run
    const paragraph = new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({
          text: "Executive Summary",
          bold: true,
          font: 'Calibri',
          size: 35,
          color: "0000FF",
          underline: {
            type: UnderlineType.SINGLE,
            color: "0000FF",
          },
        }),
      ],
    });

    return paragraph;
  } catch (error) {
    console.log(error);
  }
};

const getEducationalVerificationSummaryTable = async function (
  educationalVerificationDetails
) {

  try {


    const rows = [];
    for (let educationDetails of educationalVerificationDetails) {
      rows.push(
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: educationDetails.qualificationRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: educationDetails.universityRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),

            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: educationDetails.yopRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: educationDetails.status,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        })
      );
    }

    // link:`http://localhost:3000/localreports/reports/techMahindra/downloadAnnexures?component=employment&caseId=230606a0190b`,

    const caseDetailsTable = new Table({
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new InternalHyperlink({
                      children: [
                        new TextRun({
                          text: "Educational Verification",
                          style: "Hyperlink",
                          font: 'Calibri',
                          size: 22
                        }),
                      ],
                      anchor: "anchorForEducationalVerification",
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "University/ College Name",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "YOP",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Status",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        ...rows,
      ],
    });

    return caseDetailsTable;
  } catch (error) {
    console.log(error);
  }
};

const getEmploymentVerificationSummaryTable = async function (
  employmentVerificationDetails
) {
  try {

    const rows = [];
    for (let employmentDetails of employmentVerificationDetails) {
      const tenureFrom = getDateString(employmentDetails.dojbvfRhs);
      const tenureTo = getDateString(employmentDetails.dorbvfRhs);

      const dojDate = new Date(employmentDetails.dojbvfRhs);
      const lwdDate = new Date(employmentDetails.dorbvfRhs);

      const periodVerifiedInMonths =
        (lwdDate.getFullYear() - dojDate.getFullYear()) * 12 +
        lwdDate.getMonth() -
        dojDate.getMonth();

      rows.push(
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: employmentDetails.employer,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(20), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      //text: tenureFrom,
                      text: employmentDetails.dojbvfRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(20), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),

            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      //text: tenureTo,
                      text: employmentDetails.dorbvfRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(20), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: periodVerifiedInMonths,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(20), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: employmentDetails.status,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(20), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        })
      );
    }

    const caseDetailsTable = new Table({
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new InternalHyperlink({
                      children: [
                        new TextRun({
                          text: "Employment Verification",
                          style: "Hyperlink",
                          font: 'Calibri',
                          size: 22
                        }),
                      ],
                      anchor: "anchorForEmploymentVerification",
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(20), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Tenure (From)",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(20), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Tenure",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                    new TextRun({
                      break: true
                    }),
                    new TextRun({
                      text: "(To Period)",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(20), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Period Verified in Months",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(20), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Status",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(20), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        ...rows,
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Total Experience Verified",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },

              columnSpan: 4,
              width: {
                size: getWidthPercentage(80), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(20), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
      ],
    });

    return caseDetailsTable;
  } catch (error) {
    console.log(error);
  }
};

const getGapVerificationSummaryTable = async function (gapVfnsDetails) {
  try {

    const rows = [];
    for (let gapVfnsDetail of gapVfnsDetails) {

      rows.push(
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: gapVfnsDetail?.gaptypeRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(35), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: gapVfnsDetail?.tenureRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(35), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),

            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: gapVfnsDetail?.reasonRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(30), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
      )
    }

    const gapVerificationSummaryTable = new Table({
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new InternalHyperlink({
                      children: [
                        new TextRun({
                          text: "GAP Verification",
                          style: "Hyperlink",
                          font: 'Calibri',
                          size: 22
                        }),
                      ],
                      anchor: "anchorForGapVerification",
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(35), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Period",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(35), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Reason for Gap",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(30), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        ...rows,
      ],
    });

    return gapVerificationSummaryTable;
  } catch (error) {
    console.log(error);
  }
};

const getCourtRecordVerificationSummaryTable = async function (courtRecordDetails) {
  try {
    const rows = [];

    for (let courtRecordDetail of courtRecordDetails) {
      rows.push(
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Civil Proceedings",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(35), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: courtRecordDetail.status,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(65), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Criminal Proceedings",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(35), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: courtRecordDetail.status,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(65), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),)
    }
    const courtRecordVerificationSummaryTable = new Table({
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new InternalHyperlink({
                      children: [
                        new TextRun({
                          text: "Court Record Verification",
                          style: "Hyperlink",
                          font: 'Calibri',
                          size: 22
                        }),
                      ],
                      anchor: "anchorForCourtRecordVerification",
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(35), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Verification Status",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(35), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        ...rows
      ],
    });

    return courtRecordVerificationSummaryTable;
  } catch (error) {
    console.log(error);
  }
};

const getDatabaseVerificationDetailsTable = async function () {
  try {
    const databaseVerificationTable = new Table({
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.LEFT,
                  children: [
                    new TextRun({
                      text: "Database Verification",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              columnSpan: 3,
              width: {
                size: getWidthPercentage(100), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
              margins: {
                top: 100,
                bottom: 100,
                left: 100,
                right: 100,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new InternalHyperlink({
                      children: [
                        new TextRun({
                          text: "Executive Summary",
                          style: "Hyperlink",
                          font: 'Calibri',
                          size: 22
                        }),
                      ],
                      anchor: "anchorForChapter1",
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(30), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Database Verification",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(30), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Status",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(40), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "India Specific Regulatory &",
                      font: "Calibri",
                      size: 22,
                    }),
                    new TextRun({
                      break: true,
                    }),
                    new TextRun({
                      text: "Compliance Database",
                      font: "Calibri",
                      size: 22,
                    }),

                    new TextRun({
                      break: true,
                    }),
                    new TextRun({
                      break: true,
                    }),
                    new TextRun({
                      text: "(Bases Detail searches color",
                      font: "Calibri",
                      size: 22,
                    }),
                    new TextRun({
                      break: true,
                    }),
                    new TextRun({
                      text: "code to be mentioned in",
                      font: "Calibri",
                      size: 22,
                    }),
                    new TextRun({
                      break: true,
                    }),
                    new TextRun({
                      text: "executive Summary)",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                  spacing: {
                    line: 300,
                  },
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(30), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
              rowSpan: 4,
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Civil Litigation Database Checks – India",
                      font: "Calibri",
                      size: 22,
                    })
                  ],
                  spacing: {
                    line: 300,
                  },
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(30), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: workerData.DatabaseVerificationDetails?.length ? workerData.DatabaseVerificationDetails[0]?.civillitiRhs : "",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(40), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Credit and Reputational Risk Database Checks – India",
                      font: "Calibri",
                      size: 22,
                    })
                  ],
                  spacing: {
                    line: 300,
                  },
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(30), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: workerData.DatabaseVerificationDetails?.length ? workerData.DatabaseVerificationDetails[0]?.creditRhs : "",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(40), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Criminal Records Database Checks – India",
                      font: "Calibri",
                      size: 22,
                    })
                  ],
                  spacing: {
                    line: 300,
                  },
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(30), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: workerData.DatabaseVerificationDetails?.length ? workerData.DatabaseVerificationDetails[0]?.criminalRhs : "",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(40), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Directorship Watch List",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(30), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: workerData.DatabaseVerificationDetails?.length ? workerData.DatabaseVerificationDetails[0]?.directorshipRhs : "",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(40), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Database Global",
                      font: "Calibri",
                      size: 22,
                    }),
                    new TextRun({
                      break: true,
                    }),
                    new TextRun({
                      break: true,
                    }),
                    new TextRun({
                      text: "(Bases Detail searches color",
                      font: "Calibri",
                      size: 22,
                    }),
                    new TextRun({
                      break: true,
                    }),
                    new TextRun({
                      text: "code to be mentioned in",
                      font: "Calibri",
                      size: 22,
                    }),
                    new TextRun({
                      break: true,
                    }),
                    new TextRun({
                      text: "executive Summary)",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                  spacing: {
                    line: 300,
                  },
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(30), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
              rowSpan: 4,
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Compliance Authorities",
                      font: "Calibri",
                      size: 22,
                    }),
                    new TextRun({
                      break: true,
                    }),
                    new TextRun({
                      text: "Database Checks – Global",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                  spacing: {
                    line: 300,
                  },
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(30), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: workerData.DatabaseVerificationDetails?.length ? workerData.DatabaseVerificationDetails[0]?.complianceglobalRhs : "",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(40), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Regulatory Authorities",
                      font: "Calibri",
                      size: 22,
                    }),
                    new TextRun({
                      break: true,
                    }),
                    new TextRun({
                      text: "Database Checks – Global",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                  spacing: {
                    line: 300,
                  },
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(30), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: workerData.DatabaseVerificationDetails?.length ? workerData.DatabaseVerificationDetails[0]?.regulatoryglobalRhs : "",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(40), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Serious and Organized Crimes",
                      font: "Calibri",
                      size: 22,
                    }),
                    new TextRun({
                      break: true,
                    }),
                    new TextRun({
                      text: "Database Checks – Global",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                  spacing: {
                    line: 300,
                  },
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(30), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: workerData.DatabaseVerificationDetails?.length ? workerData.DatabaseVerificationDetails[0]?.seriousglobalRhs : "",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(40), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Web and Media Searches",
                      font: "Calibri",
                      size: 22,
                    }),
                    new TextRun({
                      break: true,
                    }),
                    new TextRun({
                      text: "- Global",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                  spacing: {
                    line: 300,
                  },
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(30), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: workerData.DatabaseVerificationDetails?.length ? workerData.DatabaseVerificationDetails[0]?.webglobalRhs : "",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(40), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Compliance Link (OFAC)",
                      font: "Calibri",
                      size: 22,
                    }),
                    new TextRun({
                      break: true,
                    }),
                    new TextRun({
                      text: "(Prohibited Parties)",
                      font: "Calibri",
                      size: 22,
                    }),
                    new TextRun({
                      break: true,
                    }),
                    new TextRun({
                      break: true,
                    }),
                    new TextRun({
                      text: "(Bases Detail searches color",
                      font: "Calibri",
                      size: 22,
                    }),
                    new TextRun({
                      break: true,
                    }),
                    new TextRun({
                      text: "code to be mentioned in",
                      font: "Calibri",
                      size: 22,
                    }),
                    new TextRun({
                      break: true,
                    }),
                    new TextRun({
                      text: "executive Summary)",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                  spacing: {
                    line: 300,
                  },
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(30), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "OFAC Specific Database",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(30), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: workerData.DatabaseVerificationDetails?.length ? workerData.DatabaseVerificationDetails[0]?.ofacRhs : "",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(40), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new InternalHyperlink({
                      children: [
                        new TextRun({
                          text: "Annexure Details",
                          style: "Hyperlink",
                          font: 'Calibri',
                          size: 22,
                          bold: true,
                        }),
                      ],
                      anchor: "gdbAnnexuresBmkId",
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              columnSpan: 2,
            }),
          ],
        })
      ],
    });

    
    if (
      workerData.DatabaseVerificationDetails?.length && workerData.DatabaseVerificationDetails[0] &&
      fs.existsSync(
        "/cvws_new_uploads/case_uploads/" +
        workerData.DatabaseVerificationDetails[0].case.caseId +
        "/techmgdb/" +
        convertToHexString(workerData.DatabaseVerificationDetails[0]._id) +
        "/proofofwork/"
      )
    ) {
      let files = fs.readdirSync(
        "/cvws_new_uploads/case_uploads/" +
        workerData.DatabaseVerificationDetails[0].case.caseId +
        "/techmgdb/" +
        convertToHexString(workerData.DatabaseVerificationDetails[0]._id) +
        "/proofofwork/"
      );
      let k = 0;
      for (let j = 0; j < files.length; j++) {
        if (path.extname(files[j]) != ".jpg") {
          continue;
        }

        const gdbAnnexure = new Paragraph({
          children: [
            new docx.PageBreak(),
            new Bookmark({
              id: k === 0 ? "gdbAnnexuresBmkId" : "",
              children: [
                new ImageRun({
                  data: fs.readFileSync(
                    "/cvws_new_uploads/case_uploads/" +
                    workerData.DatabaseVerificationDetails[0].case.caseId +
                    "/techmgdb/" +
                    convertToHexString(workerData.DatabaseVerificationDetails[0]._id) +
                    "/proofofwork/" +
                    files[j]
                  ),
                  transformation: {
                    height: 600,
                    width: 600,
                  },
                }),
              ]
            })
          ],
        });

        databaseVerificationTable.addChildElement(gdbAnnexure);
        k++;
      }
    }
    return databaseVerificationTable;
  } catch (error) {
    console.log(error);
  }
};

const getReferenceVerificationSummaryTable = async function (referenceDetails) {
  try {
    const rows = [];

    for (let referenceDetail of referenceDetails) {
      rows.push(
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: referenceDetail?.nameofreferenceRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(33), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: referenceDetail?.mode,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(33), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: referenceDetail?.status,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(34), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),)
    }
    const referenceVerificationTable = new Table({
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new InternalHyperlink({
                      children: [
                        new TextRun({
                          text: "Reference Verification",
                          style: "Hyperlink",
                          font: 'Calibri',
                          size: 22
                        }),
                      ],
                      anchor: "anchorForReferenceVerification",
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(33), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Written/ Verbal",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(33), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Status",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(34), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        ...rows
      ],
    });

    return referenceVerificationTable;
  } catch (error) {
    console.log(error);
  }
};

// address start

const getAddressVerificationSummaryTable = async function (
  addressVerificationDetails
) {
  try {
    let currentAddress;
    let permanentAddress;
    for (let item of addressVerificationDetails) {
      if (item.typeofaddress === "Present") {
        currentAddress = item;
      } else if (item.typeofaddress === " Permanent") {
        permanentAddress = item;
      }
    }

    const addressVerificationTable = new Table({
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new InternalHyperlink({
                      children: [
                        new TextRun({
                          text: "Address Verification",
                          style: "Hyperlink",
                          font: 'Calibri',
                          size: 22
                        }),
                      ],
                      anchor: "anchorForAddressVerification",
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(33), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Period of Stay",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(33), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Status",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(34), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Current Address",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(33), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: currentAddress?.tenureofstay,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(33), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: currentAddress?.status,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(34), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Permanent address",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(33), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: permanentAddress?.tenureofstay,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(33), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: permanentAddress?.status,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(34), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
      ],
    });

    return addressVerificationTable;
  } catch (error) {
    console.log(error);
  }
};
const getTechMPermanantAddressTable = async function (addressVerificationDetails) {
  try {
    let currentAddress;
    for (let item of addressVerificationDetails) {
      if (item.typeofaddress === " Permanent" || item.typeofaddress === " Present&Permanent") {
        currentAddress = item;
        break;
      }
    }
    const dateOfVerificationCompleted = getDateString(
      currentAddress?.verificationCompletionDate
    );

    const currentAddressTable = new Table({
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Details as per Subject’s Application Form",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              columnSpan: 2,
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Verification Remarks",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Address",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(20), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: currentAddress?.address,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(40), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: currentAddress?.addressRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(40), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Period of stay (From)",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(20), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: currentAddress?.tenurefrom,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(40), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: currentAddress?.tenurefromRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(40), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Period of stay (To)",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(20), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: currentAddress?.tenureto,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(40), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: currentAddress?.tenuretoRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(40), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Ownership status",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: currentAddress?.stayRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              columnSpan: 2,
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Date of verification",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: dateOfVerificationCompleted,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              columnSpan: 2,
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Any additional comments",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: currentAddress?.gradingComments,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              columnSpan: 2,
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Mode of Verification",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: currentAddress?.mode,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              columnSpan: 2,
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Results of Digital verification",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: currentAddress?.digitalsignRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              columnSpan: 2,
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Relationship with individual",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: currentAddress?.relationshipRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              columnSpan: 2,
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Verifier name",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: currentAddress?.nameofrespondentRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              columnSpan: 2,
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new InternalHyperlink({
                      children: [
                        new TextRun({
                          text: "Annexure Details",
                          style: "Hyperlink",
                          font: 'Calibri',
                          size: 22,
                          bold: true,
                        }),
                      ],
                      anchor: "permanantAddrAnnexuresBmkId",
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: currentAddress?.annexuredetailsRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              columnSpan: 2,
            }),
          ],
        }),
      ],
    });

    if (
      currentAddress &&
      fs.existsSync(
        "/cvws_new_uploads/case_uploads/" +
        currentAddress?.case.caseId +
        "/techmaddress/" +
        convertToHexString(currentAddress?._id) +
        "/proofofwork/"
      )
    ) {
      let files = fs.readdirSync(
        "/cvws_new_uploads/case_uploads/" +
        currentAddress?.case.caseId +
        "/techmaddress/" +
        convertToHexString(currentAddress?._id) +
        "/proofofwork/"
      );

      let k = 0;
      for (let j = 0; j < files.length; j++) {
        if (path.extname(files[j]) != ".jpg") {
          continue;
        }

        const addressAnnexure = new Paragraph({
          children: [
            new docx.PageBreak(),
            new Bookmark({
              id: k === 0 ? "permanantAddrAnnexuresBmkId" : "",
              children: [
                new ImageRun({
                  data: fs.readFileSync(
                    "/cvws_new_uploads/case_uploads/" +
                    currentAddress?.case.caseId +
                    "/techmaddress/" +
                    convertToHexString(currentAddress?._id) +
                    "/proofofwork/" +
                    files[j]
                  ),
                  transformation: {
                    height: 600,
                    width: 600,
                  },
                }),
              ]
            })
          ],
        });

        currentAddressTable.addChildElement(addressAnnexure);
        k++;
      }
    }

    return currentAddressTable;
  } catch (error) {
    console.log(error);
  }
};



const getTechMAddressVerificationSummaryTable = async function (
  addressVerificationDetails
) {
  try {
    const rows = [];
    let currentAddress;
    let permanentAddress;
    for (let item of addressVerificationDetails) {
      if (item.typeofaddress === "Present" || item.typeofaddress === " Present&Permanent") {
        currentAddress = item;
        rows.push(new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Current Address",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(33), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: currentAddress?.tenurefromRhs + " - " + currentAddress?.tenuretoRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(33), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: currentAddress?.status,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(34), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }));
      } 
  if (item.typeofaddress === " Permanent" || item.typeofaddress === " Present&Permanent") {
        permanentAddress = item;
        rows.push(new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Permanent address",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(33), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: permanentAddress?.tenurefromRhs + " - " + permanentAddress?.tenuretoRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(33), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: permanentAddress?.status,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(34), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }));
      }
    }

    const addressVerificationTable = new Table({
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new InternalHyperlink({
                      children: [
                        new TextRun({
                          text: "Address Verification",
                          style: "Hyperlink",
                          font: 'Calibri',
                          size: 22
                        }),
                      ],
                      anchor: "anchorForAddressVerification",
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(33), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Period of Stay",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(33), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Status",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(34), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        ...rows
      ],
    });

    return addressVerificationTable;
  } catch (error) {
    console.log(error);
  }
};

// address end

const getPassportInvestigationSummaryTable = async function (passportDetails) {
  try {
    const passportDetailsTable = new Table({
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new InternalHyperlink({
                      children: [
                        new TextRun({
                          text: "Passport Investigation",
                          style: "Hyperlink",
                          font: 'Calibri',
                          size: 22
                        }),
                      ],
                      anchor: "anchorForPassportInvestigation",
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(50), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Status",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(50), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Passport-MRZ",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(50), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: passportDetails?.status,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(50), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
      ],
    });

    return passportDetailsTable;
  } catch (error) {
    console.log(error);
  }
};

const getPanSummaryTable = async function (panDetails) {
  try {
    const panTable = new Table({
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new InternalHyperlink({
                      children: [
                        new TextRun({
                          text: "Permanent Account Number (PAN) Verification",
                          style: "Hyperlink",
                          font: 'Calibri',
                          size: 22
                        }),
                      ],
                      anchor: "anchorForPANVerification",
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(50), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Status",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(50), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "PAN Verification",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(50), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: panDetails?.status,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(50), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
      ],
    });

    return panTable;
  } catch (error) {
    console.log(error);
  }
};

// creditcheck start
const getCreditCheckSummaryTable = async function (creditCheckDetails) {
  try {
    const creditCheckTables = [];
    let i = 0;

    for (let item of creditCheckDetails) {
      const creditCheckTable = new Table({
        rows: [
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new InternalHyperlink({
                        children: [
                          new TextRun({
                            text: "Credit Check",
                            style: "Hyperlink",
                            font: 'Calibri',
                            size: 22
                          }),
                        ],
                        anchor: "anchorForCreditCheck",
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(50), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Status",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(50), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Credit Check through CIBIL",
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(50), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item?.status,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(50), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
            ],
          }),
        ],
      });

      creditCheckTables.push(creditCheckTable);
      if (creditCheckDetails.length > 1 && i < creditCheckDetails.length - 1) {
        let blankLine = getBlankLine();
        creditCheckTables.push(blankLine, blankLine);
      }
      i++;
    }

    return creditCheckTables;
  } catch (error) {
    console.log(error);
  }
};


const getTechMCreditCheckSummaryTable = async function (creditCheckDetails) {
  try {
    const creditCheckTables = [];
    let i = 0;

    for (let item of creditCheckDetails) {
      const creditCheckTable = new Table({
        rows: [
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new InternalHyperlink({
                        children: [
                          new TextRun({
                            text: "Credit Check",
                            style: "Hyperlink",
                            font: 'Calibri',
                            size: 22
                          }),
                        ],
                        anchor: "anchorForCreditCheck",
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(50), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Status",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(50), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Credit Check through CIBIL",
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(50), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item?.status,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(50), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
            ],
          }),
        ],
      });

      creditCheckTables.push(creditCheckTable);
      if (creditCheckDetails.length > 1 && i < creditCheckDetails.length - 1) {
        let blankLine = getBlankLine();
        creditCheckTables.push(blankLine, blankLine);
      }
      i++;
    }

    return creditCheckTables;
  } catch (error) {
    console.log(error);
  }
};


// creditcheck end

const getDrugTestSummaryTable = async function (
  drugTestFiveDetails,
  drugTestSevenDetails,
  drugTestTenDetails
) {
  try {
    const drugTestTable = new Table({
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new InternalHyperlink({
                      children: [
                        new TextRun({
                          text: "Drug Test Verification",
                          style: "Hyperlink",
                          font: 'Calibri',
                          size: 22
                        }),
                      ],
                      anchor: "anchorForDrugTestVerification",
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(50), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Status",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(50), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "5 Panel Drug Test (As per client requirement)",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(50), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: drugTestFiveDetails?.status,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(50), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "7 Panel Drug Test (As per client requirement)",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(50), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: drugTestSevenDetails?.status,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(50), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "10 Panel Drug Test (As per client requirement)",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(50), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: drugTestTenDetails?.status,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(50), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
      ],
    });

    return drugTestTable;
  } catch (error) {
    console.log(error);
  }
};

const getLoaSummaryTable = async function () {
  try {

    const loaTable = new Table({
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new InternalHyperlink({
                      children: [
                        new TextRun({
                          text: "LOA",
                          style: "Hyperlink",
                          font: 'Calibri',
                          size: 22
                        }),
                      ],
                      anchor: "anchorForLOA",
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(50), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Status",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(50), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new ExternalHyperlink({
                      children: [
                        new TextRun({
                          text: "LOA",
                          font: "Calibri",
                          size: 22,
                          style: "Hyperlink",
                          color: "0000FF",
                        }),
                      ],
                      link: ``,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(50), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Attached",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(50), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
      ],
    });

    return loaTable;
  } catch (error) {
    console.log(error);
  }
};

const getAadhaarSummaryTable = async function(aadhaarDeatils){
  try{
    const aadhaarSummaryTable = new Table({
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new InternalHyperlink({
                      children: [
                        new TextRun({
                          text: "Aadhaar Verification",
                          style: "Hyperlink",
                          font: 'Calibri',
                          size: 22
                        }),
                      ],
                      anchor: "anchorForAadhaarVerification",
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(50), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Status",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(50), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Aadhaar Verification",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(50), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: aadhaarDeatils?.status,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(50),
                type: WidthType.DXA,
              },
            }),
          ],
        }),
      ],
    });

    return aadhaarSummaryTable;
  } catch (error) {
    console.log(error);
  }

}

const getGDBSummaryTable = async function(gdbDeatils){
  try{
    const gdbSummaryTable = new Table({
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new InternalHyperlink({
                      children: [
                        new TextRun({
                          text: "Global Database Verification",
                          style: "Hyperlink",
                          font: 'Calibri',
                          size: 22
                        }),
                      ],
                      anchor: "anchorForGDBVerification",
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(50), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Status",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(50), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Global Database Verification",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(50), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: gdbDeatils?.status,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(50),
                type: WidthType.DXA,
              },
            }),
          ],
        }),
      ],
    });

    return gdbSummaryTable;
  } catch (error) {
    console.log(error);
  }

}

const getEducationalVerificationTable = async function (
  educationalVerificationDetails
) {
  try {
    const eductionalVerificationTables = [];
    let i = 0;

    for (let item of educationalVerificationDetails) {

      const educationalVerificationHeading = await getTableHeading("Educational Verification",
        "anchorForEducationalVerification",
        i + 1)
      eductionalVerificationTables.push(educationalVerificationHeading, blankLine, blankLine);

      const dateOfVerificationCompleted = getDateString(
        item.verificationCompletionDate
      );

      const educationalVerificationTable = new Table({
        rows: [
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Details as per Subject’s Application Form",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Verification Remarks",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                // width:{
                //     size:35,
                //     type:WidthType.PERCENTAGE
                // }
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Complete name of",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "Qualification/Degree",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "Attained",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                    spacing: {
                      line: 300,
                    },
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(20), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item?.qualification,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(40), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item?.qualificationRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(40), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Year of passing",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(20), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item?.yop,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(40), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item?.yopRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(40), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "University name",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(20), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.university,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(40), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.universityRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(40), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "School / College /",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "Institution attended",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "(full name)",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                    spacing: {
                      line: 300,
                    },
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(20), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.school,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(40), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.schoolRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(40), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "UGC/ AICTE Approval",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "Status",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                    spacing: {
                      line: 300,
                    },
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.aicteRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Suspicious Education",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "Database check",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "(Any Findings)",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                    spacing: {
                      line: 300,
                    },
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.educationdatabaseRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Additional comments",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.gradingComments,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Verifier name",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item?.verifiernameRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Verifier designation",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item?.verifierdesignationRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Department name",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item?.departmentnameRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Verified date",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: dateOfVerificationCompleted,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new InternalHyperlink({
                        children: [
                          new TextRun({
                            text: "Annexure Details",
                            style: "Hyperlink",
                            font: 'Calibri',
                            size: 22,
                            bold: true,
                          }),
                        ],
                        anchor: "educationAnnexuresBmkId" + i,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item?.annexuredetailsRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
        ],
      });
      eductionalVerificationTables.push(educationalVerificationTable);

      if (
        educationalVerificationDetails.length > 1 &&
        i < educationalVerificationDetails.length - 1
      ) {
        let blankLine = getBlankLine();
        eductionalVerificationTables.push(blankLine, blankLine);
      }

      if (
        fs.existsSync(
          "/cvws_new_uploads/case_uploads/" +
          educationalVerificationDetails[i].case.caseId +
          "/techmeducation/" +
          convertToHexString(educationalVerificationDetails[i]._id) +
          "/proofofwork/"
        )
      ) {
        let files = fs.readdirSync(
          "/cvws_new_uploads/case_uploads/" +
          educationalVerificationDetails[i].case.caseId +
          "/techmeducation/" +
          convertToHexString(educationalVerificationDetails[i]._id) +
          "/proofofwork/"
        );
        console.log("files ......", files);
        let k = 0;
        for (let j = 0; j < files.length; j++) {
          if (path.extname(files[j]) != ".jpg") {
            continue;
          }

          const educationAnnexure = new Paragraph({
            children: [
              new docx.PageBreak(),
              new Bookmark({
                id: k === 0 ? "educationAnnexuresBmkId" + i : "",
                children: [
                  new ImageRun({
                    data: fs.readFileSync(
                      "/cvws_new_uploads/case_uploads/" +
                      educationalVerificationDetails[i].case.caseId +
                      "/techmeducation/" +
                      convertToHexString(educationalVerificationDetails[i]._id) +
                      "/proofofwork/" +
                      files[j]
                    ),
                    transformation: {
                      height: 600,
                      width: 600,
                    },
                  }),
                ]
              })
            ],
          });

          educationalVerificationTable.addChildElement(educationAnnexure);
          k++;
        }
      }

      i++;
    }
    return eductionalVerificationTables;
  } catch (error) {
    console.log(error);
  }
};

const getEmploymentVerificationTable = async function (
  employmentVerificationDetails
) {
  try {
    const employmentVerificationTables = [];
    let i = 0;

    for (let item of employmentVerificationDetails) {
      const employmentVerificationHeading = await getTableHeading(
        "Employment Verification",
        "anchorForEmploymentVerification",
        i + 1
      );
      employmentVerificationTables.push(employmentVerificationHeading, blankLine, blankLine);

      const dateOfVerificationCompleted = getDateString(
        item.verificationCompletionDate
      );

      const employmentVerificationTable = new Table({
        rows: [
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Details as per Subject’s Application Form",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Verification Remarks",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Employer name",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(28), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item?.employer,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(36), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item?.employerRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(36), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Employee ID",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(20), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item?.empid,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(40), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item?.empidRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(40), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Suspicious Database",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "Check (Any findings)",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                    spacing: {
                      line: 300,
                    },
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item?.databaseRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "MCA (CIN)",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item?.mcaRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "EPFO",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item?.epfoRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "GST Filing Status",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "(Last Filing Date)",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                    spacing: {
                      line: 300,
                    },
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item?.gstRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "If Not In MCA/ EPFO",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "Site visit status",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "required",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "(Mandate site visit if",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "paid up capital less",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "than 5 lacs)",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                    spacing: {
                      line: 300,
                    },
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item?.notmcaRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Domain Creation Date",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "(Company Website",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "check, Internet",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "search)",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                    spacing: {
                      line: 300,
                    },
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item?.domaincreationRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Period of Employment",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "- As per BVF",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                    spacing: {
                      line: 300,
                    },
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(20), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item?.dojbvf + " - " + item?.dorbvf,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(40), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item?.dojbvfRhs + " - " + item?.dorbvfRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(40), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Period of Employment",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "- As per document",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                    spacing: {
                      line: 300,
                    },
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(20), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item?.dojdocument + " - " + item?.dordocument,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(40), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item?.dojdocumentRhs + " - " + item?.dordocumentRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(40), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Designation",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(20), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.designation,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(40), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.designationRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(40), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Remuneration :",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "CTC/Net (Any one)",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                    spacing: {
                      line: 300,
                    },
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(20), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item?.remuneration,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(40), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item?.remunerationRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(40), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Supervisor name &",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "Designation",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                    spacing: {
                      line: 300,
                    },
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(20), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item?.supervisor,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(40), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item?.supervisorRhs + " - " + item?.designationRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(40), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Reason for leaving",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(20), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.reasonforleaving,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(40), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.reasonforleavingRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(40), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Any issues pertaining",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "to the employment",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                    spacing: {
                      line: 300,
                    },
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item?.issuesRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Exit formalities",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "completed",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                    spacing: {
                      line: 300,
                    },
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item?.exitformalitiesRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Eligible for rehire",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item?.rehireRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Is the document authentic?",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item?.authenticRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Any additional comments",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(20), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.gradingComments,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(80), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
                columnSpan: 2
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Verifier name",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(20), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.verifiernameRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(80), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Verifier designation",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item?.verifierdesignationRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Department name",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(20), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item?.departmentnameRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(80), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Verified date",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(20), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: dateOfVerificationCompleted,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(80), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Mode of Verification",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(20), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.mode,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(80), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [

                      new InternalHyperlink({
                        children: [
                          new TextRun({
                            text: "Annexure Details",
                            style: "Hyperlink",
                            font: 'Calibri',
                            size: 22,
                            bold: true,
                          }),
                        ],
                        anchor: "employmentAnnexuresBmkId" + i,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(20), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item?.annexuredetailsRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(80), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
                columnSpan: 2,
              }),
            ],
          }),
        ],
      });

      employmentVerificationTables.push(employmentVerificationTable);

      if (
        employmentVerificationDetails.length > 1 &&
        i < employmentVerificationDetails.length - 1
      ) {
        let blankLine = getBlankLine();
        employmentVerificationTables.push(blankLine, blankLine);
      }

      if (
        fs.existsSync(
          "/cvws_new_uploads/case_uploads/" +
          employmentVerificationDetails[i].case.caseId +
          "/techmemployment/" +
          convertToHexString(employmentVerificationDetails[i]._id) +
          "/proofofwork/"
        )
      ) {
        let files = fs.readdirSync(
          "/cvws_new_uploads/case_uploads/" +
          employmentVerificationDetails[i].case.caseId +
          "/techmemployment/" +
          convertToHexString(employmentVerificationDetails[i]._id) +
          "/proofofwork/"
        );
        let k = 0;
        for (let j = 0; j < files.length; j++) {
          if (path.extname(files[j]) != ".jpg") {
            continue;
          }

          const employmentAnnexure = new Paragraph({
            children: [
              new docx.PageBreak(),
              new Bookmark({
                id: k === 0 ? "employmentAnnexuresBmkId" + i : "",
                children: [
                  new ImageRun({
                    data: fs.readFileSync(
                      "/cvws_new_uploads/case_uploads/" +
                      employmentVerificationDetails[i].case.caseId +
                      "/techmemployment/" +
                      convertToHexString(employmentVerificationDetails[i]._id) +
                      "/proofofwork/" +
                      files[j]
                    ),
                    transformation: {
                      height: 600,
                      width: 600,
                    },
                  }),
                ]
              })
            ],
          });

          employmentVerificationTable.addChildElement(employmentAnnexure);
          k++;
        }
      }

      i++;
    }

    return employmentVerificationTables;
  } catch (error) {
    console.log(error);
  }
};

const getGapJustificationTable = async function (gapVfnsDetails) {
  try {
    const gapJustificationTables = [];
    let i = 0;
    for (let item of gapVfnsDetails) {
      console.log(item);

      const gapJustificationHeading = await getTableHeading(
        "Gap Justification",
        "anchorForGapVerification",
        i + 1
      );

      gapJustificationTables.push(gapJustificationHeading, blankLine, blankLine);

      const gapJustificationTable = new Table({
        rows: [
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Gap Justification",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(20), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Remarks",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(80), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Tenure",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(20), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.tenureRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(80), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Reason",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(20), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.reasonRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(80), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new InternalHyperlink({
                        children: [
                          new TextRun({
                            text: "Annexure Details",
                            style: "Hyperlink",
                            font: 'Calibri',
                            size: 22,
                            bold: true,
                          }),
                        ],
                        anchor: "gapJustAnnexuresBmkId" + i,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(20), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item?.annexuredetailsRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(80), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
            ],
          }),
        ],
      });

      gapJustificationTables.push(gapJustificationTable);
      if (gapVfnsDetails.length > 1 && i < gapVfnsDetails.length - 1) {
        gapJustificationTables.push(blankLine, blankLine);
      }

      if (
        fs.existsSync(
          "/cvws_new_uploads/case_uploads/" +
          gapVfnsDetails[i].case.caseId +
          "/techmgapcheck/" +
          convertToHexString(gapVfnsDetails[i]._id) +
          "/proofofwork/"
        )
      ) {
        let files = fs.readdirSync(
          "/cvws_new_uploads/case_uploads/" +
          gapVfnsDetails[i].case.caseId +
          "/techmgapcheck/" +
          convertToHexString(gapVfnsDetails[i]._id) +
          "/proofofwork/"
        );
        let k = 0;
        for (let j = 0; j < files.length; j++) {
          if (path.extname(files[j]) != ".jpg") {
            continue;
          }

          const gapJustificationAnnexure = new Paragraph({
            children: [
              new docx.PageBreak(),
              new Bookmark({
                id: k === 0 ? "gapJustAnnexuresBmkId" + i : "",
                children: [
                  new ImageRun({
                    data: fs.readFileSync(
                      "/cvws_new_uploads/case_uploads/" +
                      gapVfnsDetails[i].case.caseId +
                      "/techmgapcheck/" +
                      convertToHexString(gapVfnsDetails[i]._id) +
                      "/proofofwork/" +
                      files[j]
                    ),
                    transformation: {
                      height: 600,
                      width: 600,
                    },
                  }),
                ]
              })
            ],
          });

          gapJustificationTable.addChildElement(gapJustificationAnnexure);
          k++;
        }
      }

      i++;
    }

    return gapJustificationTables;
  } catch (error) {
    console.log(error);
  }
};

const getReferenceVerificationTable = async function (referenceDetails) {
  try {
    const referenceVerificationTables = [];
    let i = 0;

    for (let item of referenceDetails) {

      const referenceVerificationHeading = await getTableHeading(
        "Reference Verification",
        "anchorForReferenceVerification",
        i + 1
      );
      referenceVerificationTables.push(referenceVerificationHeading, blankLine, blankLine);

      const dateOfVerificationCompleted = getDateString(
        item.verificationCompletionDate
      );

      const referenceTable = new Table({
        rows: [
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Details as per Subject’s Application Form",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Verification Remarks",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Name of the",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "reference",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                    spacing: {
                      line: 300, // apply space between lines
                    },
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(20), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item?.nameofreference,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(40), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item?.nameofreferenceRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(40), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Company name and",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "Designation of the",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "reference",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                    spacing: {
                      line: 300, // apply space between lines
                    },
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(20), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item?.designation,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(40), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item?.designationRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(40), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Contact number of",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "the reference",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                    spacing: {
                      line: 300, // apply space between lines
                    },
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(20), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.contactdetails,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(40), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.contactdetailsRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(40), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Association with the",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "subject (How do you",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "know the subject and",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "for how many years)",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                    spacing: {
                      line: 300, // apply space between lines
                    },
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(20), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.associationwithcandidateRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
                width: {
                  size: getWidthPercentage(80), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Communication skills",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "and interpersonal",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "skills",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "(Average/Good/Excell",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "ent)",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                    spacing: {
                      line: 300, // apply space between lines
                    },
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(20), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.communicationRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(80), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Personal/Professional",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "strengths",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                    spacing: {
                      line: 300, // apply space between lines
                    },
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(20), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.strengthsRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(80), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Personal/Professional",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "weakness",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                    spacing: {
                      line: 300, // apply space between lines
                    },
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(20), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.weaknessesRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(80), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Result oriented",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(20), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.resultRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(80), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Attitude towards work",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.attitudeRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Honest and Reliable",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text:
                          (item.honestyRhs ? item.honestyRhs : "-") +
                          " and " +
                          (item.reliabilityRhs ? item.reliabilityRhs : "-"),
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Additional comments",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.mode,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Verified date",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: dateOfVerificationCompleted,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new InternalHyperlink({
                        children: [
                          new TextRun({
                            text: "Annexure Details",
                            style: "Hyperlink",
                            font: 'Calibri',
                            size: 22,
                            bold: true,
                          }),
                        ],
                        anchor: "refAnnexuresBmkId" + i,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "",
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
        ],
      });

      referenceVerificationTables.push(referenceTable);
      if (referenceDetails.length > 1 && i < referenceDetails.length - 1) {
        let blankLine = getBlankLine();
        referenceVerificationTables.push(blankLine, blankLine);
      }

      if (
        fs.existsSync(
          "/cvws_new_uploads/case_uploads/" +
          referenceDetails[i].case.caseId +
          "/reference/" +
          convertToHexString(referenceDetails[i]._id) +
          "/proofofwork/"
        )
      ) {
        let files = fs.readdirSync(
          "/cvws_new_uploads/case_uploads/" +
          referenceDetails[i].case.caseId +
          "/reference/" +
          convertToHexString(referenceDetails[i]._id) +
          "/proofofwork/"
        );
        let k = 0;
        for (let j = 0; j < files.length; j++) {
          if (path.extname(files[j]) != ".jpg") {
            continue;
          }

          const referenceAnnexure = new Paragraph({
            children: [
              new docx.PageBreak(),
              new Bookmark({
                id: k === 0 ? "refAnnexuresBmkId" + i : "",
                children: [
                  new ImageRun({
                    data: fs.readFileSync(
                      "/cvws_new_uploads/case_uploads/" +
                      referenceDetails[i].case.caseId +
                      "/reference/" +
                      convertToHexString(referenceDetails[i]._id) +
                      "/proofofwork/" +
                      files[j]
                    ),
                    transformation: {
                      height: 600,
                      width: 600,
                    },
                  }),
                ]
              })
            ],
          });

          referenceTable.addChildElement(referenceAnnexure);
          k++;
        }
      }

      i++;
    }

    return referenceVerificationTables;
  } catch (error) {
    console.log(error);
  }
};

const getCourtRecordVerificationTable = async function (courtRecordDetails) {
  try {
    const courtRecordVerificationTables = [];
    let i = 0;
    for (let item of courtRecordDetails) {
      const courtRecordVerificationHeading = await getTableHeading(
        "Court Record Verification",
        "anchorForCourtRecordVerification",
        i + 1
      );
      courtRecordVerificationTables.push(courtRecordVerificationHeading, blankLine, blankLine);

      const dateOfVerificationCompleted = getDateString(
        item.dateofverificationRhs
	      
      );
      const timeOfVerificationCompleted = getTimeString(
        item.dateofverificationRhs
      );

      const courtRecordVerificationTable = new Table({
        rows: [
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Civil Proceedings (Civil and High Court)",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 3,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Includes Original Suit, Miscellaneous Suit, Execution and Arbitration Case",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 3,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Address",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(20), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.addresswithpin,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(40), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.addresswithpinRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(40), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "City name",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(20), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.city,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(40), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.cityRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(40), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Original suit (Civil",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "court)",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                    spacing: {
                      line: 300,
                    },
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.civilcourtRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Appeals (High court)",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.highcourtRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Criminal cases (CC)",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.civilcourtRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Private Compliant",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "Report (PCR) –",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "(Magistrate court)",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                    spacing: {
                      line: 300,
                    },
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.magistrateRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Criminal appeals",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "(Sessions court)",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                    spacing: {
                      line: 300,
                    },
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.sessionscourtRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Criminal appeals (High",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "court)",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                    spacing: {
                      line: 300,
                    },
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.highcourtRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Status",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.status,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Complete Details of",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "Case (Detail",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "information of",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "Sections found against",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "the associate)",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                    spacing: {
                      line: 300,
                    },
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "",
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Verifier name",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.verifiedbyRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Verifier designation",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "",
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Department name",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "",
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Verified Date & time",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text:
                          dateOfVerificationCompleted +
                          "   " +
                          timeOfVerificationCompleted,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new InternalHyperlink({
                        children: [
                          new TextRun({
                            text: "Annexure Details",
                            style: "Hyperlink",
                            font: 'Calibri',
                            size: 22,
                            bold: true,
                          }),
                        ],
                        anchor: "courtRecordAnnexuresBmkId" + i,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "",
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
        ],
      });

      courtRecordVerificationTables.push(courtRecordVerificationTable, blankLine, blankLine);

      const criminalRecordTable = new Table({
        rows: [
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Criminal Proceedings (Magistrate, Sessions and High Court)",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 3,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Includes Criminal Petitions, Criminal Appeal, Sessions Case, Special Sessions Case, Criminal Miscellaneous Petition and Criminal Revision Appeal.",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 3,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Address",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(20), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.addresswithpin,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(40), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.addresswithpinRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(40), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "City name",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(20), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.city,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(40), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.cityRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(40), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Original suit (Civil",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "court)",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                    spacing: {
                      line: 300,
                    },
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.civilcourtRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Appeals (High court)",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.highcourtRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Criminal cases (CC)",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.civilcourtRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Private Compliant",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "Report (PCR) –",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "(Magistrate court)",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                    spacing: {
                      line: 300,
                    },
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.magistrateRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Criminal appeals",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "(Sessions court)",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                    spacing: {
                      line: 300,
                    },
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.sessionscourtRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Criminal appeals (High",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "court)",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                    spacing: {
                      line: 300,
                    },
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.highcourtRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Status",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.status,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Complete Details of",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "Case (Detail",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "information of",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "Sections found against",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                      new TextRun({
                        break: true,
                      }),
                      new TextRun({
                        text: "the associate)",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                    spacing: {
                      line: 300,
                    },
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "",
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Verifier name",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.verifiedbyRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Verifier designation",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "",
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Department name",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "",
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Verified Date & time",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
			     
                          text:
                          dateOfVerificationCompleted +
                          "   " +
                          timeOfVerificationCompleted,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new InternalHyperlink({
                        children: [
                          new TextRun({
                            text: "Annexure Details",
                            style: "Hyperlink",
                            font: 'Calibri',
                            size: 22,
                            bold: true,
                          }),
                        ],
                        anchor: "courtRecordAnnexuresBmkId" + i,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "",
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
        ],
      });

      courtRecordVerificationTables.push(criminalRecordTable, blankLine, blankLine)

      if (courtRecordDetails.length > 1 && i < courtRecordDetails.length - 1) {
        let blankLine = getBlankLine();
        courtRecordVerificationTables.push(blankLine, blankLine);
      }

      if (
        fs.existsSync(
          "/cvws_new_uploads/case_uploads/" +
          courtRecordDetails[i].case.caseId +
          "/courtrecord/" +
          convertToHexString(courtRecordDetails[i]._id) +
          "/proofofwork/"
        )
      ) {
        let files = fs.readdirSync(
          "/cvws_new_uploads/case_uploads/" +
          courtRecordDetails[i].case.caseId +
          "/courtrecord/" +
          convertToHexString(courtRecordDetails[i]._id) +
          "/proofofwork/"
        );
        let k = 0;
        for (let j = 0; j < files.length; j++) {
          if (path.extname(files[j]) != ".jpg") {
            continue;
          }

          const courtRecordAnnexure = new Paragraph({
            children: [
              new docx.PageBreak(),
              new Bookmark({
                id: k === 0 ? "courtRecordAnnexuresBmkId" + i : "",
                children: [
                  new ImageRun({
                    data: fs.readFileSync(
                      "/cvws_new_uploads/case_uploads/" +
                      courtRecordDetails[i].case.caseId +
                      "/courtrecord/" +
                      convertToHexString(courtRecordDetails[i]._id) +
                      "/proofofwork/" +
                      files[j]
                    ),
                    transformation: {
                      height: 600,
                      width: 600,
                    },
                  }),
                ]
              })
            ],
          });
          //courtRecordVerificationTable.addChildElement(courtRecordAnnexure);
	criminalRecordTable.addChildElement(courtRecordAnnexure)
          k++;
        }
      }

      i++;
    }

    return courtRecordVerificationTables;
  } catch (error) {
    console.log(error);
  }
};


const getTechMCurrentAddressTable = async function (addressVerificationDetails) {
  try {
    let currentAddress;
    for (let item of addressVerificationDetails) {
      if (item.typeofaddress === "Present" || item.typeofaddress === " Present&Permanent") {
        currentAddress = item;
        break;
      }
    }
    const dateOfVerificationCompleted = getDateString(
      //currentAddress?.verificationCompletionDate
	currentAddress?.dateofverificationRhs
    );

    const currentAddressTable = new Table({
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Details as per Subject’s Application Form",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              columnSpan: 2,
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Verification Remarks",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Address",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(20), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: currentAddress?.address,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(40), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: currentAddress?.addressRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(40), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Period of stay (From)",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(20), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: currentAddress?.tenurefrom,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(40), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: currentAddress?.tenurefromRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(40), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Period of stay (To)",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(20), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: currentAddress?.tenureto,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(40), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: currentAddress?.tenuretoRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(40), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Ownership status",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: currentAddress?.stayRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              columnSpan: 2,
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Date of verification",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: dateOfVerificationCompleted,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              columnSpan: 2,
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Any additional comments",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: currentAddress?.gradingComments,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              columnSpan: 2,
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Mode of Verification",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: currentAddress?.mode,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              columnSpan: 2,
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Results of Digital verification",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: currentAddress?.digitalsignRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              columnSpan: 2,
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Relationship with individual",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: currentAddress?.relationshipRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              columnSpan: 2,
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Verifier name",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: currentAddress?.nameofrespondentRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              columnSpan: 2,
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new InternalHyperlink({
                      children: [
                        new TextRun({
                          text: "Annexure Details",
                          style: "Hyperlink",
                          font: 'Calibri',
                          size: 22,
                          bold: true,
                        }),
                      ],
                      anchor: "currentAddrAnnexuresBmkId",
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: currentAddress?.annexuredetailsRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              columnSpan: 2,
            }),
          ],
        }),
      ],
    });

    if (
      currentAddress &&
      fs.existsSync(
        "/cvws_new_uploads/case_uploads/" +
        currentAddress?.case.caseId +
        "/techmaddress/" +
        convertToHexString(currentAddress?._id) +
        "/proofofwork/"
      )
    ) {
      let files = fs.readdirSync(
        "/cvws_new_uploads/case_uploads/" +
        currentAddress?.case.caseId +
        "/techmaddress/" +
        convertToHexString(currentAddress?._id) +
        "/proofofwork/"
      );

      let k = 0;
      for (let j = 0; j < files.length; j++) {
        if (path.extname(files[j]) != ".jpg") {
          continue;
        }

        const addressAnnexure = new Paragraph({
          children: [
            new docx.PageBreak(),
            new Bookmark({
              id: k === 0 ? "currentAddrAnnexuresBmkId" : "",
              children: [
                new ImageRun({
                  data: fs.readFileSync(
                    "/cvws_new_uploads/case_uploads/" +
                    currentAddress?.case.caseId +
                    "/techmaddress/" +
                    convertToHexString(currentAddress?._id) +
                    "/proofofwork/" +
                    files[j]
                  ),
                  transformation: {
                    height: 600,
                    width: 600,
                  },
                }),
              ]
            })
          ],
        });


        currentAddressTable.addChildElement(addressAnnexure);
        k++;
      }
    }

    return currentAddressTable;
  } catch (error) {
    console.log(error);
  }
};

// address end
const getPassportInvestigationTable = async function (passportDetails) {
  try {
    const dateOfExpiryLhs = getDateString(passportDetails?.expirydate);
    const dateOfExpiryRhs = getDateString(passportDetails?.expirydateRhs);
    const dateOfBirth = getDateString(passportDetails?.dob);
    const dateOfBirthRhs = getDateString(passportDetails?.dobRhs);
    const dateOfIsuue = getDateString(passportDetails?.doi);
    const dateOfIssueRhs = getDateString(passportDetails?.doiRhs);
    const dateOfVerificationCompleted = getDateString(
      passportDetails?.verificationCompletionDate
    );

    const passportInvestigationTable = new Table({
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Details as per Subject’s Application Form",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              columnSpan: 2,
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Verification Remarks",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Date of birth",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(20), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: dateOfBirth,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: dateOfBirthRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Date of issue",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(20), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: dateOfIsuue,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: dateOfIssueRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Place of issue",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(20), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: passportDetails?.country,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: passportDetails?.countryRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Date of expiry",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(20), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: dateOfExpiryLhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: dateOfExpiryRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Passport number",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(20), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: passportDetails?.passportnumber,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: passportDetails?.passportnumberRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Verified date",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: dateOfVerificationCompleted,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              columnSpan: 2,
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Mode of Verification",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: passportDetails?.mode,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              columnSpan: 2,
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new InternalHyperlink({
                      children: [
                        new TextRun({
                          text: "Annexure Details",
                          style: "Hyperlink",
                          font: 'Calibri',
                          size: 22,
                          bold: true,
                        }),
                      ],
                      anchor: "passportAnnexuresBmkId",
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: passportDetails?.annexuredetailsRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              columnSpan: 2,
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Additional Comments",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: passportDetails?.gradingComments,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              columnSpan: 2,
            }),
          ],
        }),
      ],
    });

    if (
      passportDetails &&
      fs.existsSync(
        "/cvws_new_uploads/case_uploads/" +
        passportDetails.case.caseId +
        "/techmpassport/" +
        convertToHexString(passportDetails._id) +
        "/proofofwork/"
      )
    ) {
      let files = fs.readdirSync(
        "/cvws_new_uploads/case_uploads/" +
        passportDetails.case.caseId +
        "/techmpassport/" +
        convertToHexString(passportDetails._id) +
        "/proofofwork/"
      );
      let k = 0;
      for (let j = 0; j < files.length; j++) {
        if (path.extname(files[j]) != ".jpg") {
          continue;
        }

        const passportAnnexure = new Paragraph({
          children: [
            new docx.PageBreak(),
            new Bookmark({
              id: k === 0 ? "passportAnnexuresBmkId" : "",
              children: [
                new ImageRun({
                  data: fs.readFileSync(
                    "/cvws_new_uploads/case_uploads/" +
                    passportDetails.case.caseId +
                    "/techmpassport/" +
                    convertToHexString(passportDetails._id) +
                    "/proofofwork/" +
                    files[j]
                  ),
                  transformation: {
                    height: 600,
                    width: 600,
                  },
                }),
              ]
            })
          ],
        });


        passportInvestigationTable.addChildElement(passportAnnexure);
        k++;
      }
    }

    return passportInvestigationTable;
  } catch (error) {
    console.log(error);
  }
};

const getDrivingLicenseTable = async function (drivingLicenseDetails) {
  try {
    const dateOfVerificationCompleted = getDateString(
      drivingLicenseDetails?.verificationCompletionDate
    );

    const drivingLicenseTable = new Table({
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Details as per Subject’s Application Form",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              columnSpan: 2,
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Verification Remarks",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Address",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(20), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(40), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(40), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Name of the candidate",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(20), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(40), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(40), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Name as per driving license",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(20), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: drivingLicenseDetails?.nameasperid,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(40), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: drivingLicenseDetails?.nameasperidRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(40), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Date of birth",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(20), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(40), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(40), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Driving license number",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(20), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: drivingLicenseDetails?.idnumber,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(40), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: drivingLicenseDetails?.idnumberRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(40), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),

        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Mode of Verification",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: drivingLicenseDetails?.mode,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              columnSpan: 2,
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Date of verification",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: dateOfVerificationCompleted,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              columnSpan: 2,
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Status",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: drivingLicenseDetails?.status,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              columnSpan: 2,
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new InternalHyperlink({
                      children: [
                        new TextRun({
                          text: "Annexure Details",
                          style: "Hyperlink",
                          font: 'Calibri',
                          size: 22,
                          bold: true,
                        }),
                      ],
                      anchor: "DLAnnexuresBmkId",
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              columnSpan: 2,
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Additional comments",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: drivingLicenseDetails?.gradingComments,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              columnSpan: 2,
            }),
          ],
        }),
      ],
    });

    if (
      drivingLicenseDetails &&
      fs.existsSync(
        "/cvws_new_uploads/case_uploads/" +
        drivingLicenseDetails.case.caseId +
        "/identity/" +
        convertToHexString(drivingLicenseDetails._id) +
        "/proofofwork/"
      )
    ) {
      let files = fs.readdirSync(
        "/cvws_new_uploads/case_uploads/" +
        drivingLicenseDetails.case.caseId +
        "/identity/" +
        convertToHexString(drivingLicenseDetails._id) +
        "/proofofwork/"
      );

      let k = 0;
      for (let j = 0; j < files.length; j++) {
        if (path.extname(files[j]) != ".jpg") {
          continue;
        }

        const identityAnnexure = new Paragraph({
          children: [
            new docx.PageBreak(),
            new Bookmark({
              id: k === 0 ? "DLAnnexuresBmkId" : "",
              children: [
                new ImageRun({
                  data: fs.readFileSync(
                    "/cvws_new_uploads/case_uploads/" +
                    drivingLicenseDetails.case.caseId +
                    "/identity/" +
                    convertToHexString(drivingLicenseDetails._id) +
                    "/proofofwork/" +
                    files[j]
                  ),
                  transformation: {
                    height: 600,
                    width: 600,
                  },
                }),
              ]
            })
          ],
        });


        drivingLicenseTable.addChildElement(identityAnnexure);
        k++;
      }
    }

    return drivingLicenseTable;
  } catch (error) {
    console.log(error);
  }
};

const getPanVerificationTable = async function (panDetails) {
  try {
    const dateOfVerificationCompleted = getDateString(
      panDetails?.verificationCompletionDate
    );
    const dobLhs = getDateString(panDetails?.dob);
    const dobRhs = getDateString(panDetails?.dobRhs);

    const panVerificationTable = new Table({
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Details as per Subject’s Application Form",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              columnSpan: 2,
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Verification Remarks",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Name as per pan card",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(20), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: panDetails?.candidatename,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(40), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: panDetails?.candidatenameRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(40), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Date of birth",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(20), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: dobLhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(40), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: dobRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(40), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Pan Card number",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(20), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: panDetails?.pan,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(40), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: panDetails?.panRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(40), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Mode of Verification",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: panDetails?.mode,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              columnSpan: 2,
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Verified date",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: dateOfVerificationCompleted,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              columnSpan: 2,
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new InternalHyperlink({
                      children: [
                        new TextRun({
                          text: "Annexure Details",
                          style: "Hyperlink",
                          font: 'Calibri',
                          size: 22,
                          bold: true,
                        }),
                      ],
                      anchor: "panAnnexuresBmkId",
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              columnSpan: 2,
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Additional comments",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: panDetails?.gradingComments,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              columnSpan: 2,
            }),
          ],
        }),
      ],
    });

    if (
      panDetails &&
      fs.existsSync(
        "/cvws_new_uploads/case_uploads/" +
        panDetails.case.caseId +
        "/pan/" +
        convertToHexString(panDetails._id) +
        "/proofofwork/"
      )
    ) {
      let files = fs.readdirSync(
        "/cvws_new_uploads/case_uploads/" +
        panDetails.case.caseId +
        "/pan/" +
        convertToHexString(panDetails._id) +
        "/proofofwork/"
      );
      let k = 0;
      for (let j = 0; j < files.length; j++) {
        if (path.extname(files[j]) != ".jpg") {
          continue;
        }

        const panAnnexure = new Paragraph({
          children: [
            new docx.PageBreak(),
            new Bookmark({
              id: k === 0 ? "panAnnexuresBmkId" : "",
              children: [
                new ImageRun({
                  data: fs.readFileSync(
                    "/cvws_new_uploads/case_uploads/" +
                    panDetails.case.caseId +
                    "/pan/" +
                    convertToHexString(panDetails._id) +
                    "/proofofwork/" +
                    files[j]
                  ),
                  transformation: {
                    height: 600,
                    width: 600,
                  },
                }),
              ]
            })
          ],
        });

        panVerificationTable.addChildElement(panAnnexure);
        k++;
      }
    }

    return panVerificationTable;
  } catch (error) {
    console.log(error);
  }
};


const getTechMCreditCheckVerificationTable = async function (creditCheckDetails) {
  try {
    const creditCheckVerificationTables = [];
    let i = 0;
    for (let item of creditCheckDetails) {
      const creditCheckVerificationHeading = await getTableHeading(
        "Credit Check Verification",
        "anchorForCreditCheck",
        i + 1
      );

      creditCheckVerificationTables.push(creditCheckVerificationHeading, blankLine, blankLine);

      const dateOfVerificationCompleted = getDateString(
        item?.verificationCompletionDate
      );

      const creditCheckVerificationTable = new Table({
        rows: [
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Details as per Subject’s Application Form",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Verification Remarks",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Country name",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(20), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item?.countryname,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(40), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item?.countrynameRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(40), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
            ],
          }),

          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Status",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.status,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Credit Score",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item.creditscoreRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Mode of Verification",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item?.mode,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Verified date",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: dateOfVerificationCompleted,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new InternalHyperlink({
                        children: [
                          new TextRun({
                            text: "Annexure Details",
                            style: "Hyperlink",
                            font: 'Calibri',
                            size: 22,
                            bold: true,
                          }),
                        ],
                        anchor: "creditAnnexuresBmkId" + i,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item?.annexuredetailsRhs,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Additional comments",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: item?.gradingComments,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                columnSpan: 2,
              }),
            ],
          }),
        ],
      });

      creditCheckVerificationTables.push(creditCheckVerificationTable);

      if (creditCheckDetails.length > 1 && i < creditCheckDetails.length - 1) {
        creditCheckVerificationTables.push(blankLine, blankLine);
      }

      if (
        fs.existsSync(
          "/cvws_new_uploads/case_uploads/" +
          creditCheckDetails[i].case.caseId +
          "/techmcreditcheck/" +
          convertToHexString(creditCheckDetails[i]._id) +
          "/proofofwork/"
        )
      ) {
        let files = fs.readdirSync(
          "/cvws_new_uploads/case_uploads/" +
          creditCheckDetails[i].case.caseId +
          "/techmcreditcheck/" +
          convertToHexString(creditCheckDetails[i]._id) +
          "/proofofwork/"
        );

        let k = 0;
        for (let j = 0; j < files.length; j++) {
          if (path.extname(files[j]) != ".jpg") {
            continue;
          }

          const creditcheckAnnexure = new Paragraph({
            children: [
              new docx.PageBreak(),

              new Bookmark({
                id: k === 0 ? "creditAnnexuresBmkId" + i : "",
                children: [
                  new ImageRun({
                    data: fs.readFileSync(
                      "/cvws_new_uploads/case_uploads/" +
                      creditCheckDetails[i].case.caseId +
                      "/techmcreditcheck/" +
                      convertToHexString(creditCheckDetails[i]._id) +
                      "/proofofwork/" +
                      files[j]
                    ),
                    transformation: {
                      height: 600,
                      width: 600,
                    },
                  }),
                ]
              })
            ],
          });

          creditCheckVerificationTable.addChildElement(creditcheckAnnexure);
          k++;
        }
      }

      i++;
    }
    return creditCheckVerificationTables;
  } catch (error) {
    console.log(error);
  }
};
// creditcheck end
const getDrugTestVerificationTable = async function (drugTestFive) {

  try {
    const dateOfVerificationCompleted = getDateString(
      drugTestFive?.verificationCompletionDate
    );

    const drugTestVerificationTable = new Table({
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Substance",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(65), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Result",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(35), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Amphetamine",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(65), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: drugTestFive?.amphetamineRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(35), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Cocaine",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(65), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: drugTestFive?.cocaineRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(35), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Marijuana (Cannabinoids)",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(65), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: drugTestFive?.marijuanaRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(35), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Phencyclidine (PCP)",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(65), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: drugTestFive?.phencyclidineRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(35), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Opiates",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(65), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: drugTestFive?.opiatesRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(35), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Additional comments",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(65), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: drugTestFive?.gradingComments,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(35), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Verified date",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(65), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: dateOfVerificationCompleted,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(35), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Annexure Details",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(65), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: drugTestFive?.annexuredetailsRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(35), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
      ],
    });

    return drugTestVerificationTable;
  } catch (error) {
    console.log(error);
  }
};

const getDrugTestTenVerificationTable = async function (drugTestTen) {

  try {
    const dateOfVerificationCompleted = getDateString(
      drugTestTen?.verificationCompletionDate
    );

    const drugTestVerificationTable = new Table({
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Substance",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(65), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Result",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(35), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Amphetamine",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(65), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: drugTestTen?.amphetamineRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(35), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Cocaine",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(65), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: drugTestTen?.cocaineRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(35), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Barbiturates",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(65), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: drugTestTen?.barbituratesRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(35), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Phencyclidine (PCP)",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(65), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: drugTestTen?.phencyclidineRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(35), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Benzodiazepines",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(65), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: drugTestTen?.benzodiazepinesRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(35), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),

        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Methadone",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(65), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: drugTestTen?.methadoneRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(35), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),

        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Methaqualone - Quaaludes",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(65), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: drugTestTen?.methaqualonRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(35), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        //
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Additional comments",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(65), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: drugTestTen?.gradingComments,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(35), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Verified date",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(65), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: dateOfVerificationCompleted,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(35), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Annexure Details",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(65), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: drugTestTen?.annexuredetailsRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(35), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
      ],
    });

    return drugTestVerificationTable;
  } catch (error) {
    console.log(error);
  }
};

/*const getLoaTable = async function (caseDetails) {
  try{
  const loaTable = new Table({
    rows: [
      new TableRow({
        children: [
          new TableCell({
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({
                    text: "Annexure Details",
                    bold: true,
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
            ],
            shading: { fill: "CCCCCC" },
            margins: {
              top: 100,
              left: 100,
              bottom: 100,
              right: 100,
            },
            width: {
              size: getWidthPercentage(35), // Set the table width to 100% of the page width
              type: WidthType.DXA,
            },
          }),
          new TableCell({
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({
                    text: "Result",
                    font: "Calibri",
                    size: 22,
                  }),
                ],
              }),
            ],
            margins: {
              top: 100,
              left: 100,
              bottom: 100,
              right: 100,
            },
            width: {
              size: getWidthPercentage(65), // Set the table width to 100% of the page width
              type: WidthType.DXA,
            },
          }),
        ],
      }),
    ],
  });
  if (
    fs.existsSync(
      "/cvws_new_uploads/case_uploads/" +
        caseDetails.caseId +
        "/loa/"
    )
  ) {
    let files = fs.readdirSync(
      "/cvws_new_uploads/case_uploads/" +
        caseDetails.caseId +
        "/loa/"
    );

    let k=0;
    for (let j = 0; j < files.length; j++) {
      if (path.extname(files[j]) != ".jpg") {
        continue;
      }

      const loaAnnexure =  new Paragraph({
        children: [
          new docx.PageBreak(),
          new ImageRun({
            data: fs.readFileSync(
              "/cvws_new_uploads/case_uploads/" + caseDetails.caseId +"/loa/" + files[j]
            ),
            transformation: {
              height: 600,
              width: 600,
            },
          }),
        ],
      });

      loaTable.addChildElement(loaAnnexure);
      k++;
    }
  }
  return loaTable;
}catch(error){
  console.log(error);
}
}; */ // LOA

const getLoaTable = async function (caseDetails) {
  try {

    if (
      fs.existsSync(
        "/cvws_new_uploads/case_uploads/" +
        caseDetails.caseId +
        "/loa/"
      )
    ) {
      let files = fs.readdirSync(
        "/cvws_new_uploads/case_uploads/" +
        caseDetails.caseId +
        "/loa/"
      );
      let output_file_prefix = ""

      let pdfPath = "/cvws_new_uploads/case_uploads/" + caseDetails.caseId + "/loa/";
      for (let j = 0; j < files.length; j++) {
        if (path.extname(files[j]) == ".pdf") {
          output_file_prefix = files[j].split(".pdf")[0] ;
	  console.log("outputfile:", output_file_prefix)
          pdfPath += files[j]
          break;
        }
      }

      let outputDir = "/cvws_new_uploads/case_uploads/" +
        caseDetails.caseId +
        "/loa/"

     console.log("PDF Path:", pdfPath)
      const options = {
        density: 100,
        saveFilename: output_file_prefix,
        savePath: `${outputDir}`,
        format: "png",
        width: 600,
        height: 600
      };
	
try{
	
      const convert = pdf2pic.fromPath(pdfPath, options);
      

      await convert(1)
	fs.renameSync(outputDir + "/LOA.1.png", outputDir + "LOA.png")
	console.log("Converted to PNG")
}catch(err){
	console.log(err)
}
      const loaTable = new Table({
        rows: [
          new TableRow({
            children: [
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
			       new Bookmark({
                    id: "anchorForLOA",
                    children: [
                      new TextRun({
                        text: "Annexure Details",
                        bold: true,
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                    ],
                  }),
                ],
                shading: { fill: "CCCCCC" },
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(35), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
              new TableCell({
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Result",
                        font: "Calibri",
                        size: 22,
                      }),
                    ],
                  }),
                ],
                margins: {
                  top: 100,
                  left: 100,
                  bottom: 100,
                  right: 100,
                },
                width: {
                  size: getWidthPercentage(65), // Set the table width to 100% of the page width
                  type: WidthType.DXA,
                },
              }),
            ],
          }),
        ],
      });

      const loaAnnexure = new Paragraph({
        children: [
          new docx.PageBreak(),
          new ImageRun({
            data: fs.readFileSync(
              "/cvws_new_uploads/case_uploads/" + caseDetails.caseId + "/loa/" + output_file_prefix  + ".png"
            ),
            transformation: {
              height: 600,
              width: 600,
            },
          }),
        ],
      });

      loaTable.addChildElement(loaAnnexure);

      //fs.unlinkSync("/cvws_new_uploads/case_uploads/" + caseDetails.caseId + "/loa/" + output_file_prefix + ".jpg");
      return loaTable;
    }
  } catch (error) {
    console.log(error);
  }
};

const getAadhaarVerificationTable = async function (aadhaarDeatils) {
  try {
    const dobLhs = getDateString(aadhaarDeatils?.dob);
    const dobRhs = getDateString(aadhaarDeatils?.dobRhs);
    const verifiedOnLhs = getDateString(aadhaarDeatils?.verifiedon);
    const verifiedOnRhs = getDateString(aadhaarDeatils?.verifiedonRhs);
    
    const aadhaarVerificationTable = new Table({
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "COMPONENTS",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(33), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "INFORMATION PROVIDED",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(33), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "INFORMATION VERIFIED",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(34), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Candidate Name",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(33),
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: aadhaarDeatils?.candidatename,
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "FFFFFF" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(33), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: aadhaarDeatils?.candidatenameRhs,
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "FFFFFF" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(34), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Aadhaar Number",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(33),
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: aadhaarDeatils?.aadhaarnumber,
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "FFFFFF" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(33), 
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: aadhaarDeatils?.aadhaarnumberRhs,
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "FFFFFF" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(34), 
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Date of Birth",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(33),
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: dobLhs,
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "FFFFFF" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(33), 
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: dobRhs,
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "FFFFFF" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(34), 
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Verified By",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(33),
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: aadhaarDeatils?.verifiedby,
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "FFFFFF" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(33), 
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: aadhaarDeatils?.verifiedbyRhs,
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "FFFFFF" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(34), 
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Verified On",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(33),
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: verifiedOnLhs,
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "FFFFFF" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(33), 
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: verifiedOnRhs,
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "FFFFFF" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(34), 
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new InternalHyperlink({
                      children: [
                        new TextRun({
                          text: "Annexure Details",
                          style: "Hyperlink",
                          font: 'Calibri',
                          size: 22,
                          bold: true,
                        }),
                      ],
                      anchor: "aadhaarAnnexuresBmkId",
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(33),
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: '',
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "FFFFFF" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(33), 
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: '',
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "FFFFFF" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(34), 
                type: WidthType.DXA,
              },
            }),
          ],
        })
      ],
    });

    if (
      aadhaarDeatils &&
      fs.existsSync(
        "/cvws_new_uploads/case_uploads/" +
        aadhaarDeatils.case.caseId +
        "/aadhaarverification/" +
        convertToHexString(aadhaarDeatils._id) +
        "/proofofwork/"
      )
    ) {
      let files = fs.readdirSync(
        "/cvws_new_uploads/case_uploads/" +
        aadhaarDeatils.case.caseId +
        "/aadhaarverification/" +
        convertToHexString(aadhaarDeatils._id) +
        "/proofofwork/"
      );
      let k = 0;
      for (let j = 0; j < files.length; j++) {
        if (path.extname(files[j]) != ".jpg") {
          continue;
        }

        const aadhaarAnnexure = new Paragraph({
          children: [
            new docx.PageBreak(),
            new Bookmark({
              id: k === 0 ? "aadhaarAnnexuresBmkId" : "",
              children: [
                new ImageRun({
                  data: fs.readFileSync(
                    "/cvws_new_uploads/case_uploads/" +
                    aadhaarDeatils.case.caseId +
                    "/aadhaarverification/" +
                    convertToHexString(aadhaarDeatils._id) +
                    "/proofofwork/" +
                    files[j]
                  ),
                  transformation: {
                    height: 600,
                    width: 600,
                  },
                }),
              ]
            })
          ],
        });

        aadhaarVerificationTable.addChildElement(aadhaarAnnexure);
        k++;
      }
    }

    return aadhaarVerificationTable;
  } catch (error) {
    console.log(error);
  }
};
const pageBreakCode = new Paragraph({
  text: "",
  pageBreakBefore: true,
});


// Added by anil on 2/22/2024

const getCVAnalysisVerificationSummaryTable = async function (cvanalysisDetails) {
  try {
 
    const cvAnalysisSummaryTable = new Table({
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new InternalHyperlink({
                      children: [
                        new TextRun({
                          text: "CV Analysis Verification",
                          style: "Hyperlink",
                          font: 'Calibri',
                          size: 22
                        }),
                      ],
                      anchor: "anchorForCVAnalysis",
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(35), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Verification Status",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(35), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "CV Analysis",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(35), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.status,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(65), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        })
      ],
    });

    return cvAnalysisSummaryTable;
  } catch (error) {
    console.log(error);
  }
};

const getCVAnalysisVerificationDetailsTable = async function (
  cvanalysisDetails
) {
  try {
    const cvAnalysisDetailsTable = new Table({
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "COMPONENTS",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                })
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "INFORMATION PROVIDED",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "INFORMATION VERIFIED",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "VARIANCE",
                      bold: true,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                bottom: 100,
                right: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Candidate Name",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.candidatename,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.candidatenameRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.candidatenameRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Education 1 School",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.educationoneschoolRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.educationoneschoolcvRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.educationoneschoolvarianceRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Education 1 - Start Date",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.educationonestartdateRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.educationonestartdatecvRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.educationonestartdatevarianceRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Education 1 - End Date",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.educationoneenddateRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.educationoneenddatecvRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.educationoneenddatevarianceRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Education 1 - Major Area of Study",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.educationonemajorareaofstudyRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.educationonemajorareaofstudycvRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.educationonemajorareaofstudyvarianceRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Education 1 - Degree Completed",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.educationonedegreecompletedRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.educationonedegreecompletedcvRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.educationonedegreecompletedvarianceRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Education 1 - Degree Received",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.educationonedegreereceivedRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.educationonedegreereceivedcvRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.educationonedegreereceivedvarianceRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Education 1 - Degree Received Date",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.educationonedegreereceiveddateRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.educationonedegreereceiveddatecvRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.educationonedegreereceiveddatevarianceRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Education 2 - School",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.educationtwoschoolRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.educationtwoschoolcvRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.educationtwoschoolvarianceRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Education 2 - Start Date",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.educationtwostartdateRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.educationtwostartdatecvRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.educationtwostartdatevarianceRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Education 2 - End Date",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.educationtwoenddateRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.educationtwoenddatecvRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.educationtwoenddatevarianceRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Education 2 - Major Area of Study",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.educationtwomajorareaofstudyRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.educationtwomajorareaofstudycvRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.educationtwomajorareaofstudyvarianceRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Education 2 - Degree Completed",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.educationtwodegreecompletedRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.educationtwodegreecompletedcvRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.educationtwodegreecompletedvarianceRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Education 2 - Degree Received",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.educationtwodegreereceivedRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.educationtwodegreereceivedcvRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.educationtwodegreereceivedvarianceRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Education 2 - Degree Received Date",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.educationtwodegreereceiveddateRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.educationtwodegreereceiveddatecvRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.educationtwodegreereceiveddatevarianceRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Empoyer 1 Name",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employeroneRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employeronecvRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employeronevarianceRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Empoyer 1 Job Title",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employeronejobtitleRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employeronejobtitlecvRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employeronejobtitlevarianceRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Empoyer 1 Start Date",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employeronestartdateRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employeronestartdatecvRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employeronestartdatevarianceRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Empoyment 1 End Date",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employeroneenddateRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employeroneenddatecvRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employeroneenddatevarianceRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Empoyment 1 - Length of Employment",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employeronelengthofemploymentRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employeronelengthofemploymentcvRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employeronelengthofemploymentvarianceRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Empoyer 2 Name",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employertwoRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employertwocvRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employertwovarianceRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Empoyer 2 Job Title",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employertwojobtitleRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employertwojobtitlecvRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employertwojobtitlevarianceRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Empoyer 2 Start Date",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employertwostartdateRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employertwostartdatecvRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employertwostartdatevarianceRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Empoyment 2 End Date",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employertwoenddateRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employertwoenddatecvRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employertwoenddatevarianceRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Empoyment 2 - Length of Employment",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employertwolengthofemploymentRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employertwolengthofemploymentcvRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employertwolengthofemploymentvarianceRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Empoyer 3 Name",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employerthreeRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employerthreecvRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employerthreevarianceRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Empoyer 3 Job Title",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employerthreejobtitleRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employerthreejobtitlecvRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employerthreejobtitlevarianceRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Empoyer 3 Start Date",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employerthreestartdateRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employerthreestartdatecvRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employerthreestartdatevarianceRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Empoyment 3 End Date",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employerthreeenddateRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employerthreeenddatecvRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employerthreeenddatevarianceRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Empoyment 3 - Length of Employment",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employerthreelengthofemploymentRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employerthreelengthofemploymentcvRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employerthreelengthofemploymentvarianceRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Empoyer 4 Name",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employerfourRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employerfourcvRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employerfourvarianceRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Empoyer 4 Job Title",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employerfourjobtitleRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employerfourjobtitlecvRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employerfourjobtitlevarianceRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Empoyer 4 Start Date",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employerfourstartdateRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employerfourstartdatecvRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employerfourstartdatevarianceRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Empoyment 4 End Date",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employerfourenddateRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employerfourenddatecvRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employerfourenddatevarianceRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Empoyment 4 - Length of Employment",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employerfourlengthofemploymentRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employerfourlengthofemploymentcvRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employerfourlengthofemploymentvarianceRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Empoyer 5 Name",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employerfiveRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employerfivecvRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employerfivevarianceRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Empoyer 5 Job Title",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employerfivejobtitleRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employerfivejobtitlecvRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employerfivejobtitlevarianceRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Empoyer 5 Start Date",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employerfivestartdateRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employerfivestartdatecvRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employerfivestartdatevarianceRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Empoyment 5 End Date",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employerfiveenddateRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employerfiveenddatecvRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employerfiveenddatevarianceRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "Empoyment 5 - Length of Employment	5	5	No",
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employerfivelengthofemploymentRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employerfivelengthofemploymentcvRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: cvanalysisDetails?.employerfivelengthofemploymentvarianceRhs,
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new InternalHyperlink({
                      children: [
                        new TextRun({
                          text: "Annexure Details",
                          style: "Hyperlink",
                          font: 'Calibri',
                          size: 22,
                          bold: true,
                        }),
                      ],
                      anchor: "cvAnalysisAnnexuresBmkId",
                    }),
                  ],
                }),
              ],
              shading: { fill: "CCCCCC" },
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: '',
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: '',
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: '',
                      font: "Calibri",
                      size: 22,
                    }),
                  ],
                }),
              ],
              margins: {
                top: 100,
                left: 100,
                right: 100,
                bottom: 100,
              },
              width: {
                size: getWidthPercentage(25), // Set the table width to 100% of the page width
                type: WidthType.DXA,
              },
            }),
          ],
        })
      ],
    });

    if (
      cvanalysisDetails &&
      fs.existsSync(
        "/cvws_new_uploads/case_uploads/" +
        cvanalysisDetails.case.caseId +
        "/cvanalysis/" +
        convertToHexString(cvanalysisDetails._id) +
        "/proofofwork/"
      )
    ) {
      let files = fs.readdirSync(
        "/cvws_new_uploads/case_uploads/" +
        cvanalysisDetails.case.caseId +
        "/cvanalysis/" +
        convertToHexString(cvanalysisDetails._id) +
        "/proofofwork/"
      );
      let k = 0;
      for (let j = 0; j < files.length; j++) {
        if (path.extname(files[j]) != ".jpg") {
          continue;
        }

        const cvAnalysisAnnexure = new Paragraph({
          children: [
            new docx.PageBreak(),
            new Bookmark({
              id: k === 0 ? "cvAnalysisAnnexuresBmkId" : "",
              children: [
                new ImageRun({
                  data: fs.readFileSync(
                    "/cvws_new_uploads/case_uploads/" +
                    cvanalysisDetails.case.caseId +
                    "/cvanalysis/" +
                    convertToHexString(cvanalysisDetails._id) +
                    "/proofofwork/" +
                    files[j]
                  ),
                  transformation: {
                    height: 600,
                    width: 600,
                  },
                }),
              ]
            })
          ],
        });

        cvAnalysisDetailsTable.addChildElement(cvAnalysisAnnexure);
        k++;
      }
    }

    return cvAnalysisDetailsTable;
  } catch (error) {
    console.log(error);
  }
};


createReport();
async function createReport() {
  try {
    console.log("Started Creating Doc");

    let clientNameTable1 = await writeClientNameTable(workerData.caseDetails.subclient.client.name)

    let reportTitleTable = await writeReportTitleTable(workerData.caseDetails)


    const clientNameTable = await getApplicantTable(
      workerData.caseDetails,
      workerData.personalDetails
    );
    const caseDetailsTable = await getCaseDetailsTable(
      workerData.caseDetails,
      workerData.profileOrPackageName
    );
    const colorCodeTable = await getColorCodeTable();

    const backGroundVerificationReportTable = await getBackGroundVerificationReportTable();

    const executiveSummaryHeading = await getExecutiveSummaryHeading();

    const educationVerificationalSummaryTable =
      await getEducationalVerificationSummaryTable(
        workerData.educationalVerificationDetails
      );


    const employmentVerificationalSummaryTable =
      await getEmploymentVerificationSummaryTable(
        workerData.employmentVerificationDetails
      );
    const gapVerificationSummaryTable = await getGapVerificationSummaryTable(workerData.gapVfnsDetails);
    const courtRecordVerificationSummaryTable =
      await getCourtRecordVerificationSummaryTable(workerData.courtRecordDetails);
    // const databaseVerificationDetailsTable =
    //   await getDatabaseVerificationDetailsTable();
    const referenceVerificationSummaryTable =
      await getReferenceVerificationSummaryTable(workerData.referenceDetails);


    // Address start
    const addressVerificationSummaryTable =
      await getAddressVerificationSummaryTable(
        workerData.addressVerificationDetails
      );

    const techMaddressVerificationSummaryTable =
      await getTechMAddressVerificationSummaryTable(
        workerData.techMaddressVerificationDetails
      );

    // Address End




    const passportInvestigationSummaryTable =
      await getPassportInvestigationSummaryTable(workerData.passportDetails);
    const panSummaryTable = await getPanSummaryTable(workerData.panDetails);

    // creditcheck start
    const creditCheckSummaryTable = await getCreditCheckSummaryTable(
      workerData.creditCheckDetails
    );

    const techMcreditCheckSummaryTable = await getTechMCreditCheckSummaryTable(
      workerData.techMcreditCheckDetails
    );
    // creditcheck end


    const drugTestSummaryTable = await getDrugTestSummaryTable(
      workerData.drugTestFive,
      workerData.drugTestSeven,
      workerData.drugTestTen
    );
    const loaSummaryTable = await getLoaSummaryTable();
    const aadhaarSummaryTable = await getAadhaarSummaryTable(workerData.aadhaarDeatils);
    let gdbSummaryTable;
    if(workerData.DatabaseVerificationDetails?.length){
       gdbSummaryTable = await getGDBSummaryTable(workerData.DatabaseVerificationDetails[0]);

    }
    
    const educationalVerificationalTable =
      await getEducationalVerificationTable(
        workerData.educationalVerificationDetails
      );

    const employmentVerificationTable = await getEmploymentVerificationTable(
      workerData.employmentVerificationDetails
    );

    const gapJustificationTable = await getGapJustificationTable(
      workerData.gapVfnsDetails
    );
    const referenceVerificationTable = await getReferenceVerificationTable(
      workerData.referenceDetails
    );

    const courtRecordVerificationTable = await getCourtRecordVerificationTable(
      workerData.courtRecordDetails
    );



    const techMCurrentAddressHeading = await getTableHeading(
      "Current Address (Physical/Digital)",
      "anchorForAddressVerification",
      workerData.techMaddressVerificationDetails?.length ? 1:""

    );
    const techMCcurrentAddressTable = await getTechMCurrentAddressTable(
      workerData.techMaddressVerificationDetails
    );

    const istechMCurrentAddress = workerData.techMaddressVerificationDetails?.some(item => item.typeofaddress === "Present" || item.typeofaddress === " Present&Permanent");

const techMPermanantAddressHeading = await getTableHeading(
  "Permanent Address (Physical/Digital)",
  istechMCurrentAddress ? "":"anchorForAddressVerification",
  workerData.techMaddressVerificationDetails?.length ? 1:""

);

const techMPermanantAddressTable = await getTechMPermanantAddressTable(
  workerData.techMaddressVerificationDetails
);

const istechMPermanentAddress = workerData.techMaddressVerificationDetails?.some(item => item.typeofaddress === " Permanent" || item.typeofaddress === " Present&Permanent");





    // address end

    let passPortCount = 0;
    if (workerData.passportDetails) {
      passPortCount = 1;
    }
    const passportInvestigationHeading = await getTableHeading(
      "Passport Investigation",
      "anchorForPassportInvestigation",
      passPortCount
    );
    const passportVerificationTable = await getPassportInvestigationTable(
      workerData.passportDetails
    );

    let drivingLicenceCount = 0;

    if (workerData.drivingLicenseDetails) {
      drivingLicenceCount = 1;
    }
    const drivingLicenseHeading = await getTableHeading(
      "Driving License Verification",
      null,
      drivingLicenceCount
    );
    const drivingLicenseTable = await getDrivingLicenseTable(
      workerData.drivingLicenseDetails
    );

    let panDetailsCount = 0;
    if (workerData.panDetails) {
      panDetailsCount = 1;
    }
    const panVerificationHeading = await getTableHeading(
      "Permanent Account Number (PAN) Verification",
      "anchorForPANVerification",
      panDetailsCount
    );
    const panVerificationTable = await getPanVerificationTable(
      workerData.panDetails
    );


    const techMcreditCheckVerificationTable = await getTechMCreditCheckVerificationTable(
      workerData.techMcreditCheckDetails
    );

    // creditcheck end


    let drugTestFiveCount = 0;
    if (workerData.drugTestFive) {
      drugTestFiveCount = 1;
    }
    const drugTestVerificationHeading = await getTableHeading(
      "Drug Test Verification",
      "anchorForDrugTestVerification",
      drugTestFiveCount
    );

    const drugTestVerificationTable = await getDrugTestVerificationTable(
      workerData.drugTestFive
    );

    const drugTestTenVerificationTable = await getDrugTestTenVerificationTable(
      workerData.drugTestTen
    );

    const loaTableHeading = await getTableHeading("LOA", "anchorForLOA", "");
    const loaTable = await getLoaTable(workerData.caseDetails);

    let aadhaarDetailsCount = 0;
    if (workerData.aadhaarDeatils) {
      aadhaarDetailsCount = 1;
    }
    const aadhaarVerificationHeading = await getTableHeading(
      "Aadhaar Verification",
      "anchorForAadhaarVerification",
      aadhaarDetailsCount
    );
    const aadhaarVerificationTable = await getAadhaarVerificationTable(
      workerData.aadhaarDeatils
    );

    let gdbDetailsCount = 0;
    if(workerData.DatabaseVerificationDetails?.length){
      gdbDetailsCount = 1;
    }

    const gdbVerificationHeading = await getTableHeading(
      "Global Database Verification",
      "anchorForGDBVerification",
      gdbDetailsCount
    );

    const databaseVerificationDetailsTable =
    await getDatabaseVerificationDetailsTable();


    let blankLine = getBlankLine();

    let disclaimerSection = await writeDisclaimer()

   //Added by anil on 2/22/2024 cvanalysis start
   
   const cvanalysisSummaryTable =
   await getCVAnalysisVerificationSummaryTable(workerData.cvanalysisDetails);
   
   let cvAnalysisDetailsCount = 0;
   if (workerData.cvanalysisDetails) {
     cvAnalysisDetailsCount = 1;
   }
   const cvAnalysisHeading = await getTableHeading(
     "CV Analysis Verification",
     "anchorForCVAnalysis",
     cvAnalysisDetailsCount
   );
   const cvAnalysisTable = await getCVAnalysisVerificationDetailsTable(
     workerData?.cvanalysisDetails
   );

   //Added by anil on 2/22/2024 cvanalysis end

        //jan 23 2024 start ---------------------------------------------

        const childrenArray = [
          clientNameTable1, blankLine, blankLine, reportTitleTable, blankLine, clientNameTable,
          blankLine,
          blankLine,
          caseDetailsTable,
          blankLine,
          blankLine,
          colorCodeTable,
          blankLine,
          blankLine,
          backGroundVerificationReportTable,
          blankLine,
          blankLine,		
          executiveSummaryHeading,
          blankLine,
          blankLine,
        ]
    
        if (workerData.educationalVerificationDetails?.length) {
          childrenArray.push(educationVerificationalSummaryTable, blankLine,
            blankLine)
        }
    
        if (workerData.employmentVerificationDetails?.length) {
          childrenArray.push(
            employmentVerificationalSummaryTable,
            blankLine,
            blankLine)
        }
    
        if (workerData.gapVfnsDetails?.length) {
          childrenArray.push(
            gapVerificationSummaryTable,
            blankLine,
            blankLine)
        }
    
        if (workerData.courtRecordDetails?.length) {
          childrenArray.push(
            courtRecordVerificationSummaryTable,
            blankLine,
            blankLine)
        }
    
        if (workerData.referenceDetails?.length) {
            childrenArray.push(referenceVerificationSummaryTable, blankLine,
              blankLine)
        }
    
        if (istechMCurrentAddress || istechMPermanentAddress) {
          childrenArray.push(techMaddressVerificationSummaryTable,
            blankLine,blankLine)
        }
        if (workerData.passportDetails) {
          childrenArray.push(passportInvestigationSummaryTable,
            blankLine,
            blankLine);
        }
        if (workerData.panDetails) {
          childrenArray.push(
            panSummaryTable,
            blankLine,
            blankLine)
        }

        if(workerData.DatabaseVerificationDetails?.length){
          childrenArray.push(gdbSummaryTable,blankLine,blankLine);
        }
      
        if (workerData.aadhaarDeatils) {
          childrenArray.push(
            aadhaarSummaryTable,
            blankLine,
            blankLine)
        }
    
        if (workerData.techMcreditCheckDetails?.length) {
          childrenArray.push(...techMcreditCheckSummaryTable,
            blankLine,
            blankLine)
        }
    
        if(workerData.drugTestFive || workerData.drugTestTen){
          childrenArray.push(drugTestSummaryTable,
            blankLine,
            blankLine)
        }
    
        if(workerData?.cvanalysisDetails){
          childrenArray.push(
            cvanalysisSummaryTable,
            blankLine,
            blankLine)
        }
        
        if(workerData.caseDetails){
          childrenArray.push(loaSummaryTable,blankLine,blankLine);
        }
    
    
        //Jan 23 2024 end ---------------------------------------



    if (workerData.educationalVerificationDetails?.length) {
      childrenArray.push(...educationalVerificationalTable, blankLine,
        blankLine)
    }
    if (workerData.employmentVerificationDetails?.length) {
      childrenArray.push(
        ...employmentVerificationTable,
        blankLine,
        blankLine)
    }
    if (workerData.gapVfnsDetails?.length) {
      childrenArray.push(
        ...gapJustificationTable,
        blankLine,
        blankLine)
    }

  

    if (workerData.courtRecordDetails?.length) {
      childrenArray.push(
        ...courtRecordVerificationTable,
        blankLine,
        blankLine)
    }

    if (workerData.referenceDetails?.length) {
      childrenArray.push(...referenceVerificationTable, blankLine,
        blankLine)
    }

    if (istechMCurrentAddress) {
      childrenArray.push(techMCurrentAddressHeading,
        blankLine,
        blankLine,
        techMCcurrentAddressTable,
        blankLine,
        blankLine)
    }

if (istechMPermanentAddress) {
  childrenArray.push(techMPermanantAddressHeading,
    blankLine,
    blankLine,
    techMPermanantAddressTable,
    blankLine,
    blankLine)
}

    // address end
    if (workerData.passportDetails) {
      childrenArray.push(passportInvestigationHeading,
        blankLine,
        blankLine,
        passportVerificationTable,
        blankLine,
        blankLine);
    }

    if (workerData.drivingLicenseDetails) {
      childrenArray.push(drivingLicenseHeading,
        blankLine,
        blankLine,
        drivingLicenseTable,
        blankLine,
        blankLine);
    }

    if (workerData.panDetails) {
      childrenArray.push(
        panVerificationHeading,
        blankLine,
        blankLine,
        panVerificationTable,
        blankLine,
        blankLine)
    }

    if(workerData.DatabaseVerificationDetails?.length){
      childrenArray.push(pageBreakCode,gdbVerificationHeading,blankLine,blankLine, databaseVerificationDetailsTable,pageBreakCode);
    }
    if (workerData.aadhaarDeatils) {
      childrenArray.push(
        aadhaarVerificationHeading,
        blankLine,
        blankLine,
        aadhaarVerificationTable,
        blankLine,
        blankLine)
    }

    if (workerData.techMcreditCheckDetails?.length) {
      childrenArray.push(...techMcreditCheckVerificationTable,
        blankLine,
        blankLine)
    }
    // creditcheck end
    if (workerData.drugTestFive) {
      childrenArray.push(drugTestVerificationHeading,
        blankLine,
        blankLine,
        drugTestVerificationTable,
        blankLine,
        blankLine)
    }

    if (workerData.drugTestTen) {
      childrenArray.push(
        drugTestTenVerificationTable,
        blankLine,
        blankLine)
    }

    //added by anil on 2/22/2024 cvanalysis push start
    if (workerData.cvanalysisDetails) {
      childrenArray.push(
        cvAnalysisHeading,
        blankLine,
        blankLine,
        cvAnalysisTable,
        blankLine,
        blankLine)
    }
  //added by anil on 2/22/2024 cvanalysis push end
    const doc = new Document({
      sections: [
        {
          propertis: {},
          headers: {
            default: new docx.Header({
              children: [
                new docx.Paragraph({
                  children: [new docx.ImageRun({
                    data: fs.readFileSync('/cvws_new_uploads/verifacts_logo/Verifacts-Log.jpg'),
                    transformation: {
                      width: 150,
                      height: 80,
                    },
                  })],
                  alignment: docx.AlignmentType.CENTER
                }),
              ]
            })
          },
          footers: {
            default: new docx.Footer({
              children: [
                new docx.Paragraph({
                  children: [
                    new InternalHyperlink({
                      children: [
                        new TextRun({
                          text: "Back to summary",
                          style: "Hyperlink",
                          font: 'Calibri',
                          size: 22
                        }),
                      ],
                      anchor: "anchorForChapter1",
                    }),
                  ],
                }),
                new docx.Paragraph({
                  children: [
                    new docx.TextRun({
                      text: "CONFIDENTIAL",
                      font: "Calibri"
                    }),
                    new docx.TextRun({
                      children: ["\t\t\t\t Page ", docx.PageNumber.CURRENT, " of ", docx.PageNumber.TOTAL_PAGES],
                      font: "Calibri"
                    }),

                  ]
                }),
              ],
            }),
          },
          children: [

            ...childrenArray,

            loaTableHeading,
            blankLine,
            blankLine,
            loaTable,
            blankLine,
            blankLine,
            disclaimerSection
          ],
        },
      ],
    });

    console.log("Created Doc");
    const filePath = `/cvws_new_uploads/backgroundVerification/${workerData.caseDetails.caseId}.docx`;

    const folderPath = path.dirname(filePath);

    // Check if the folder and its parent folders exist, create them if they don't
    createFolderRecursive(folderPath);

    // Check if the file exists and create it if it doesn't
    createFileIfNotExists(filePath);

    docx.Packer.toBuffer(doc)
      .then((data) => {
        fs.writeFileSync(filePath, data);
        console.log("File Written ---now will delete the jpegs");
        console.log("Now downloading the file");
        parentPort.postMessage({ status: "Done" });
      })
      .catch((err) => {
        parentPort.postMessage({ status: "error" });
      });
  } catch (err) {
    console.log("error in createReport", err);
  }
}

function createFolderRecursive(dir) {
  try {
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true });
    }
  } catch (error) {
    console.log(error);
  }
}

function createFileIfNotExists(file) {
  try {
    if (!fs.existsSync(file)) {
      fs.writeFileSync(file, "");
    }
  } catch (error) {
    console.log(error);
  }
}

function getDateString(dateAndTime) {
  try {
    if (dateAndTime) {
      const date = new Date(dateAndTime);
      const dobDate = date.getDate().toString().padStart(2, "0");
      const dobMonth = date.toLocaleString("default", { month: "short" });
      const dobYear = date.getFullYear();
      return dobDate + " - " + dobMonth + " - " + dobYear;
    }
    return "";
  } catch (error) {
    console.log(error);
  }
}

function getTimeString(date) {
  try {
    if (date) {
      let hours = new Date(date).getHours();
      const amPm = hours >= 12 ? "PM" : "AM";
      hours = hours ? hours % 12 : 12;
      let minutes = new Date(date).getMinutes().toString().padStart(2, "0");
      let seconds = new Date(date).getSeconds().toString().padStart(2, "0");
      return (
        hours.toString().padStart(2, "0") +
        " : " +
        minutes +
        " : " +
        seconds +
        " " +
        amPm
      );
    }
    return "";
  } catch (error) {
    console.log(error);
  }

}



function convertToHexString(buffer) {
  try {
    return Array.prototype.map
      .call(buffer.id, (x) => ("00" + x.toString(16)).slice(-2))
      .join("");
  } catch (error) {
    console.log(error);
  }

}

//NEW 13-DEC-22
let writeDisclaimer = function () {
  try {
    return new Promise((resolve, reject) => {
      let disclaimerRootParagraph = new docx.Paragraph({
        children: [new docx.TextRun("")]
      })
      let disclaimerHeaderParagraph = new docx.Paragraph({
        children: [,
          new docx.PageBreak(),
          new docx.TextRun({
            text: "Restrictions and Limitations",
            font: "Calibri",
            size: 25,
            bold: true,
            underline: true
          })
        ],
        alignment: docx.AlignmentType.CENTER
      })
      let blankLine = new docx.Paragraph({
        children: [
          new docx.TextRun("")
        ]
      })
      let firstParagraph = new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: "Our reports and comments are confidential in nature and are meant only for the internal use of the client to make an assessment of the background of the applicant. They are not intended for publication or circulation or sharing with any other person including the applicant. Also, they are not to be reproduced or used for any other purpose, in whole or in part, without our prior written consent in each specific instance",
            font: "Calibri",
            size: 20
          })
        ]
      })

      let secondParagraph = new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: "We request you to recognize that we are not the source of the data gathered and our findings are based on the information made available to us; therefore, we cannot guarantee the accuracy of the information collected. Should additional information or documentation become available to us, which impacts the conclusions reached in our reports, we reserve the right to amend our findings in our report accordingly. ",
            font: "Calibri",
            size: 20
          })
        ]
      })
      let thirdParagraph = new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: "We expressly disclaim all responsibility or liability for any costs, damages, losses, liabilities, expenses incurred by anyone as a result of circulation, publication, reproduction or use of our reports contrary to the provisions of this paragraph. You will appreciate that due to factors beyond our control, it may be possible that we are unable to get all the necessary information. Because of the limitations mentioned above, the results of our work with respect to the background checks should be considered only as a guide. Our reports and comments should not be considered as a definitive pronouncement on the individual.",
            font: "Calibri",
            size: 20

          })
        ]
      })
      let fourthParagraph = new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: "Note – For any concerns on the report sent/uploaded kindly reach out to escalations@verifacts.co.in",
            font: "Calibri",
            size: 20

          })
        ]
      })

      let footerParagraph = new docx.Paragraph({
        children: [
          new docx.TextRun({
            text: "--END OF REPORT--",
            font: "Calibri",
            size: 20,
            bold: true
          })
        ],
        alignment: docx.AlignmentType.CENTER
      })
      disclaimerRootParagraph.addChildElement(disclaimerHeaderParagraph)
      disclaimerRootParagraph.addChildElement(blankLine)
      disclaimerRootParagraph.addChildElement(firstParagraph)
      disclaimerRootParagraph.addChildElement(blankLine)
      disclaimerRootParagraph.addChildElement(secondParagraph)
      disclaimerRootParagraph.addChildElement(blankLine)
      disclaimerRootParagraph.addChildElement(thirdParagraph)
      disclaimerRootParagraph.addChildElement(blankLine)
      disclaimerRootParagraph.addChildElement(fourthParagraph)
      disclaimerRootParagraph.addChildElement(blankLine)
      disclaimerRootParagraph.addChildElement(footerParagraph)
      resolve(disclaimerRootParagraph)

    })
  } catch (error) {
    console.log(error);
  }
}
