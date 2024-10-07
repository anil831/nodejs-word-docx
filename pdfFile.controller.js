const fs = require('fs')
const PDFDocument = require('pdfkit');
const { Worker } = require('worker_threads')
const moment = require('moment');
const path = require('path');
const { FootNoteReferenceRunAttributes, TextRun } = require('docx');
const { resolve } = require('path');
const ColorMaster = require('../../models/administration/color_master.model')
const ClientContractPackage = require('../../models/administration/client_contract_package.model')
const ClientContractProfile = require('../../models/administration/client_contract_profile.model')
const client = require("../../models/administration/client.model")
const subClient = require("../../models/administration/subclient.model")
const Case = require('../../models/uploads/case.model');
const PersonalDetailsData = require('../../models/data_entry/personal_details_data.model')
const User = require('../../models/administration/user.model')
const ExcelJS = require('exceljs');
const UserRole = require('../../models/administration/user_role.model');
const ComponentAccess = require('../../models/administration/component_access.model')
const employment = require('../../models/data_entry/employment.model')
const education = require('../../models/data_entry/education.model');
const address = require('../../models/data_entry/address.model');
const courtrecord = require('../../models/data_entry/courtrecord.model');
const criminalrecord = require('../../models/data_entry/criminalrecord.model');
const identity = require('../../models/data_entry/identity.model');
const creditcheck = require('../../models/data_entry/creditcheck.model');
const socialmedia = require('../../models/data_entry/socialmedia.model');
const globaldatabase = require('../../models/data_entry/globaldatabase.model');
const reference = require('../../models/data_entry/reference.model');
const refbasic = require('../../models/data_entry/refbasic.model');
const colorblindness = require('../../models/data_entry/colorblindness.model')
const drugtestfive = require('../../models/data_entry/drugtestfive.model');
const drugtestten = require('../../models/data_entry/drugtestten.model');
const passport = require('../../models/data_entry/passport.model');
const addresstelephone = require('../../models/data_entry/addresstelephone.model');
const addressonline = require('../../models/data_entry/addressonline.model');
const addresscomprehensive = require('../../models/data_entry/addresscomprehensive.model');
const addressbusiness = require('../../models/data_entry/addressbusiness.model');
const educationcomprehensive = require('../../models/data_entry/educationcomprehensive.model');
const cvanalysis = require('../../models/data_entry/cvanalysis.model')
const educationadvanced = require('../../models/data_entry/educationadvanced.model');
const drugtestsix = require('../../models/data_entry/drugtestsix.model');
const drugtestseven = require('../../models/data_entry/drugtestseven.model');
const drugtesteight = require('../../models/data_entry/drugtesteight.model');
const drugtestnine = require('../../models/data_entry/drugtestnine.model');
const facisl3 = require('../../models/data_entry/facisl3.model');
const credittrans = require('../../models/data_entry/credittrans.model');
const creditequifax = require('../../models/data_entry/creditequifax.model');
const empadvance = require('../../models/data_entry/empadvance.model');
const empbasic = require('../../models/data_entry/empbasic.model');
const vddadvance = require('../../models/data_entry/vddadvance.model');
const dlcheck = require('../../models/data_entry/dlcheck.model');
const voterid = require('../../models/data_entry/voterid.model');
const ofac = require('../../models/data_entry/ofac.model');
const physostan = require('../../models/data_entry/physostan.model')
const gapvfn = require('../../models/data_entry/gapvfn.model')
const sitecheck = require('../../models/data_entry/sitecheck.model')
const bankstmt = require('../../models/data_entry/bankstmt.model')
const directorshipcheck = require('../../models/data_entry/directorshipcheck.model')
const exitinterview = require('../../models/data_entry/exitinterview.model')
const uan = require('../../models/data_entry/uan.model')
const EPFO = require('../../models/data_entry/epfo.model')
const twentysixas = require('../../models/data_entry/twentysixas.model')
const vddpv = require('../../models/data_entry/vddpv.model')
const caconfirmation = require('../../models/data_entry/caconfirmation.model')
const vdddeclaration = require('../../models/data_entry/vdddeclaration.model')
const tcsvdd = require('../../models/data_entry/tcsvdd.model')
const formsixteen = require('../../models/data_entry/formsixteen.model')
const hl = require('../../models/data_entry/hl.model')
const gs = require('../../models/data_entry/gs.model')
//const socialmedia = require('../../models/data_entry/socialmedia.model')
const pdf2img = require('pdf2img');
const isPDF = require('is-pdf-valid');
const pdfkittable = require("pdfkit-table")

const component = require("../../models/administration/component.model");
const componentFields = require("../../models/administration/component_field.model");
const mongoose = require('mongoose');
const bankstmtModel = require('../../models/data_entry/bankstmt.model');
const subclientModel = require('../../models/administration/subclient.model');
const drugtesteightModel = require('../../models/data_entry/drugtesteight.model');
const glob = require('glob'); // added on 4/7/23

// const pdf2pic = require('pdf2pic');
// const PDFImage = require("pdf-image").PDFImage
// const { pdf } = require("pdf-to-img");
const pdfPoppler = require('pdf-poppler');
// Function to load the image and return a promise



exports.convertDataToPdf = async (req, res) => {
  const doc = new PDFDocument();
  let pdfImage;
  let pdfImageOptions;

  //fields _id,actualComponents(this field holds components name)  from cases collection
  const componentList = await Case.findOne({ caseId: req.params.caseId }, { '_id': 1, 'caseId': 1, 'actualComponents': 1, 'client': 1, 'subclient': 1, 'initiationDate': 1, 'reportDate': 1, 'firstInsufficiencyRaisedDate': 1, 'lastInsufficiencyClearedDate': 1, 'grade': 1 });

  console.log("componentList", componentList.caseId);

  // getting all the jpg files from the folder

  // const getJpgFiles = async () => {
  //   const uniqueComponents = [...new Set(componentList.actualComponents)];
  //   console.log("componentList", componentList.caseId);

  //   for (let component of uniqueComponents) {
  //     const modelPath = path.join(__dirname, `../../models/data_entry/${component}.model.js`);
  //     if (!fs.existsSync(modelPath)) {
  //       continue;
  //     }

  //     const model = require(`../../models/data_entry/${component}.model`);
  //     const modelData = await model.find({ case: componentList._id }).populate("component");
  //     console.log("modelData.length", modelData);

  //     for (let data of modelData) {
  //       const folderPath = `/cvws_new_uploads/case_uploads/${componentList.caseId}/${component}/${data._id}/proofofwork`;

  //       if (fs.existsSync(folderPath)) {
  //         const files = fs.readdirSync(folderPath);
  //         console.log("files:", files);

  //         const jpgFiles = files.filter(file => file.slice(-3).toLowerCase() === "jpg");
  //         console.log("jpgFiles:", jpgFiles);
  //       }
  //     }
  //   }
  // };


  const getJpgFiles = async () => {
    // getting all the jpg files and push all the jpg files in an array
    const jpgFiles = [];

    const uniqueComponents = [...new Set(componentList.actualComponents)];

    for (let component of uniqueComponents) {
      const modelPath = path.join(__dirname, `../../models/data_entry/${component}.model.js`);
      if (!fs.existsSync(modelPath)) {
        continue;
      }

      const model = require(`../../models/data_entry/${component}.model`);
      const modelData = await model.find({ case: componentList._id }).populate("component");

      for (let data of modelData) {
        const folderPath = `/cvws_new_uploads/case_uploads/${componentList.caseId}/${component}/${data._id}/proofofwork`;

        if (fs.existsSync(folderPath)) {
          const proofofworkfiles = fs.readdirSync(folderPath);

          const jpgFilesInFolder = proofofworkfiles.filter(file => file.slice(-4).toLowerCase() === ".jpg");

          for (let jpgFile of jpgFilesInFolder) {
            // const jpgPath = `${folderPath}/converted-${jpgFile.slice(0, -4)}jpg`;
            const jpgPath = `${folderPath}/${jpgFile}`;
            jpgFiles.push(jpgPath);
          }
        }
      }
    }

    return jpgFiles;
  };

  getJpgFiles()
    .then(jpgFiles => {
      console.log("JPG files:", jpgFiles);
    })
    .catch(error => {
      console.error("An error occurred:", error);
    });


  //getting the pdf and converting the pdf into image 
  const convertPdfToImage = async () => {
    const uniqueComponents = [...new Set(componentList.actualComponents)];
    console.log("componentList", componentList.caseId);

    for (let component of uniqueComponents) {
      const modelPath = path.join(__dirname, `../../models/data_entry/${component}.model.js`);
      if (!fs.existsSync(modelPath)) {
        continue;
      }

      const model = require(`../../models/data_entry/${component}.model`);
      const modelData = await model.find({ case: componentList._id }).populate("component");
      console.log("modelData.length", modelData);

      for (let data of modelData) {
        const folderPath = `/cvws_new_uploads/case_uploads/${componentList.caseId}/${component}/${data._id}/proofofwork`;

        if (fs.existsSync(folderPath)) {
          const proofofworkfiles = fs.readdirSync(folderPath);
          console.log("proofofworkfiles:", proofofworkfiles);

          const pdfFiles = proofofworkfiles.filter(file => file.slice(-3).toLowerCase() === "pdf");
          console.log("pdfFiles:", pdfFiles);

          for (let pdfFile of pdfFiles) {
            const pathFile = `${folderPath}/${pdfFile}`;
            const outputPath = folderPath;
            const outputFormat = 'jpg';

            const options = {
              format: outputFormat,
              out_dir: outputPath,
              out_prefix: 'converted',
            };

            pdfPoppler.convert(pathFile, options)
              .then(info => {
                console.log(`Report: ${pdfFile} - Converted to image`);
                console.log("Image paths:", info);
                console.log('PDF created successfully');
              })
              .catch(err => {
                console.error(`Error converting ${pdfFile}: ${err}`);
              });
          }
        }
      }
    }
  };

  convertPdfToImage()
    .catch(error => {
      console.error("An error occurred:", error);
    });


  console.log("componentList", componentList.caseId);



  let componentsData = {};
  for (let i = 0; i < componentList.actualComponents.length; i++) {
    componentsData[componentList.actualComponents[i]] = {};
  }

  // console.log("componentsData is :", componentsData);

  //get ObjectId and name of each component from component collection. 
  //using name match copy ObjectId into componentsData object
  const componentsIdList = await component.find({ name: { $in: componentList.actualComponents } },
    { '_id': 1, 'name': 1, 'displayName': 1 });

  // console.log("componentsIdList is: ", componentsIdList);

  for (let i = 0; i < componentsIdList.length; i++) {
    if (componentsData.hasOwnProperty(componentsIdList[i].name)) {
      componentsData[componentsIdList[i].name]['_id'] = componentsIdList[i]._id;
      componentsData[componentsIdList[i].name]['displayName'] = componentsIdList[i].displayName;
    }
  }
  console.log("componentsIdList is: ", componentsIdList);

  //use component ObjectId to get all field related to that particular component from 
  //componentFields collection
  for (let i = 0; i < componentsIdList.length; i++) {
    const convertedComponentId = mongoose.Types.ObjectId(componentsIdList[i]._id)

    // console.log("convertedComponentId is:", convertedComponentId)

    const componentFieldList = await componentFields.find({ component: convertedComponentId },
      { '_id': 0, 'name': 1, 'label': 1 });

    // console.log("componentFieldList is :", componentFieldList);

    let componentFieldNames = [];
    let componentLables = [];
    let fieldsAndLabels = [];

    //appendRhs to each filed and save final fields list into array and map to respective component
    for (let i = 0; i < componentFieldList.length; i++) {

      //below we are calling toObject() on top of mongo query response to only include fields that are returning from mongodb 
      //query not mongodb document properties or prototype properties
      const convertedFieldObject = componentFieldList[i].toObject();


      let originalComponentFields = [];
      let temp = { [convertedFieldObject.name]: 1 };
      originalComponentFields.push(temp);

      componentLables.push(convertedFieldObject.label);

      componentFieldNames = componentFieldNames.concat(originalComponentFields);

      originalComponentFields.forEach(field => {

        let modifiedField = Object.keys(field)[0] + 'Rhs';
        fieldsAndLabels.push({ [convertedFieldObject.label]: convertedFieldObject.label, lhsField: convertedFieldObject.name, rhsField: modifiedField })

        let temp = { [modifiedField]: 1 };
        componentFieldNames.push(temp);
      });

      // console.log("originalComponentFields is :", originalComponentFields);

    }

    if (componentsData.hasOwnProperty(componentsIdList[i].name)) {
      componentsData[componentsIdList[i].name]['fields'] = componentFieldNames;
      componentsData[componentsIdList[i].name]['labels'] = componentLables;
      componentsData[componentsIdList[i].name]['fieldsAndLabels'] = fieldsAndLabels;
    }

  }

  const caseId = mongoose.Types.ObjectId(componentList._id);

  const logoPath = "C:/Users/It-spare/Desktop/localreports/Pictur1-removebg-preview.png";
  const logoWidth = 140;
  const logoHeight = 50;
  const logoStartY = 20;


  // Calculate the X coordinate to center the image horizontally
  const centerX = (doc.page.width - logoWidth) / 2;

  // adding stamp 
  const stampPath = "C:/Users/It-spare/Desktop/vibe updated/localreports/download.png"

  const stampWidth = 75;
  const stampHeight = 75;
  const stampStartY = 670;
  const stampX = (doc.page.width - stampWidth) / 2;

  doc.image(logoPath, centerX, logoStartY, { width: logoWidth, height: logoHeight });

  doc.font('Helvetica').fontSize(10).text("For Verifacts Services Pvt. Ltd.", centerX + 180, logoStartY + 640, { width: logoWidth, height: logoHeight });
  doc.font('Helvetica').fontSize(10).text("Client Account Manager", centerX + 180, logoStartY + 730, { width: logoWidth, height: logoHeight });


  doc.image(stampPath, stampX + 160, stampStartY, { width: stampWidth, height: stampHeight })

  // getting client names
  const clientDoc = await client.findById(componentList.client, { 'name': 1, '_id': 0 })

  const clientX = 100;
  const clientY = 95;
  const clientCellWidth = 400;
  const clientCellPadding = 5;
  const pageWidth = doc.page.width;
  const pageHeight = doc.page.height;
  const clientTextWidth = doc.widthOfString(clientDoc['name'])
  const reportTextWidth = doc.widthOfString('BACKGROUND VERIFICATION FINAL REPORT');

  // Get the available width of the container
  containerWidth = 400

  // Calculate the starting position to center the text
  const clientStartX = clientX - 15 + (containerWidth - clientTextWidth) / 2
  const clientMaxCellHeight = doc.currentLineHeight() + clientCellPadding * 2

  doc.lineWidth(0.1).rect(clientX, clientY, clientCellWidth, clientMaxCellHeight).fillAndStroke('lightgray', 'black');
  doc.font('Helvetica-Bold').fontSize(10).fillColor('black').text(clientDoc.name.toUpperCase(), clientStartX, clientY + clientCellPadding)

  const reportX = 100;
  const reportY = 130;

  const reportStartX = reportX + 15 + (containerWidth - reportTextWidth) / 2;

  doc.lineWidth(0.1).rect(reportX, reportY, clientCellWidth, clientMaxCellHeight).fillAndStroke('lightgray', 'black');
  doc.font('Helvetica-Bold').fontSize(10).fillColor('black').text('BACKGROUND VERIFICATION FINAL REPORT', reportStartX, reportY + clientCellPadding)


  //  >-------------->
  const NameX = 100;
  const NameY = 160;
  const NameCellWidth = 100;
  const NameCellPadding = 3;

  const personalDetails = await PersonalDetailsData.findOne({ case: mongoose.Types.ObjectId(componentList._id) }, { 'candidatename': 1, 'dateofbirth': 1 })
  console.log('case id is:', componentList._id);

  if (personalDetails && personalDetails.dateofbirth) {
    const dateOfBirth = new Date(personalDetails.dateofbirth);
    const formattedDOB = dateOfBirth.toLocaleDateString('en-US', {
      day: 'numeric',
      month: 'long',
      year: 'numeric'
    });

    personalDetails.formattedDOB = formattedDOB;
  }
  console.log("personal details is :", personalDetails);

  // Set font properties
  doc.font('Helvetica').fontSize(8);

  const NameMaxCellHeight = doc.currentLineHeight() + NameCellPadding * 2
  // for column
  doc.rect(NameX, NameY, NameCellWidth, NameMaxCellHeight).fillAndStroke('lightgray', 'black');
  doc.fillColor('black').text('CANDIDATE NAME', NameX + NameCellPadding, NameY + NameCellPadding);

  // for rows
  doc.rect(NameX + 100, NameY, NameCellWidth + 200, NameMaxCellHeight).fillAndStroke('white', 'black');
  doc.fillColor('black').text(personalDetails.candidatename, NameX + NameCellWidth + NameCellPadding, NameY + NameCellPadding);

  const subClientX = 100;
  const subClientY = 174;
  const subClientCellWidth = 100;
  const subClientCellPadding = 3;

  // using subclient id from cases we are getting the data of names subclients models 
  const subClientDoc = await subClient.findById(componentList.subclient, { 'name': 1, '_id': 0 })
  // console.log("subclient id is", subClientDoc);

  // Set font properties
  doc.fontSize(8);

  const subClientMaxCellHeight = doc.currentLineHeight() + subClientCellPadding * 2
  // for column
  doc.rect(subClientX, subClientY, subClientCellWidth, subClientMaxCellHeight).fillAndStroke('lightgray', 'black');
  doc.fillColor('black').text('Sub-Client', subClientX + subClientCellPadding, subClientY + subClientCellPadding);

  // for rows
  doc.rect(subClientX + 100, subClientY, subClientCellWidth + 6, subClientMaxCellHeight).fillAndStroke('white', 'black');
  doc.fillColor('black').text(subClientDoc.name, subClientX + subClientCellWidth + subClientCellPadding, subClientY + subClientCellPadding);

  doc.rect(subClientX + 200, subClientY, subClientCellWidth + 6, subClientMaxCellHeight).fillAndStroke('lightgray', 'black');
  doc.fillColor('black').text('DATE OF BIRTH', subClientX + subClientCellPadding + subClientCellWidth * 2, subClientY + subClientCellPadding);

  doc.rect(subClientX + 300, subClientY, subClientCellWidth, subClientMaxCellHeight).fillAndStroke('white', 'black');
  doc.fillColor('black').text(personalDetails.formattedDOB, subClientX + subClientCellPadding + subClientCellWidth * 3, subClientY + subClientCellPadding);

  const InitiatedX = 100;
  const InitiatedY = 188;
  const InitiatedCellWidth = 100;
  const InitiatedCellPadding = 3;

  // for innitiation date
  const initiationdate = new Date(componentList['initiationDate'])
  const initiatedday = String(initiationdate.getDate()).padStart(2, '0');
  const initiatedmonth = initiationdate.toLocaleString('default', { month: 'long' });
  const initiatedyear = initiationdate.getFullYear();
  const dateOfInnitiation = `${initiatedday}-${initiatedmonth}-${initiatedyear}`;
  // console.log(dateOfInnitiation);


  // for report date
  const reportdate = new Date(componentList['reportDate'])
  const reportday = String(reportdate.getDate()).padStart(2, '0');
  const reportmonth = reportdate.toLocaleString('default', { month: 'long' });
  const reportyear = reportdate.getFullYear();
  const reportDate = `${reportday}-${reportmonth}-${reportyear}`;
  // console.log(reportDate);

  // Set font properties
  doc.fontSize(8);

  const InitiatedMaxCellHeight = doc.currentLineHeight() + InitiatedCellPadding * 2

  doc.rect(InitiatedX, InitiatedY, InitiatedCellWidth, InitiatedMaxCellHeight).fillAndStroke('lightgray', 'black');
  doc.fillColor('black').text('DATE INITIATED', InitiatedX + InitiatedCellPadding, InitiatedY + InitiatedCellPadding);

  doc.rect(InitiatedX + 100, InitiatedY, InitiatedCellWidth, InitiatedMaxCellHeight).fillAndStroke('white', 'black');
  doc.fillColor('black').text(dateOfInnitiation, InitiatedX + InitiatedCellWidth + InitiatedCellPadding, InitiatedY + InitiatedCellPadding);

  doc.rect(InitiatedX + 200, InitiatedY, InitiatedCellWidth, InitiatedMaxCellHeight).fillAndStroke('lightgray', 'black');
  doc.fillColor('black').text('DATE OF REPORT', InitiatedX + InitiatedCellPadding + InitiatedCellWidth * 2, InitiatedY + InitiatedCellPadding);

  doc.rect(InitiatedX + 300, InitiatedY, InitiatedCellWidth, InitiatedMaxCellHeight).fillAndStroke('white', 'black');
  doc.fillColor('black').text(reportDate, InitiatedX + InitiatedCellPadding + InitiatedCellWidth * 3, InitiatedY + InitiatedCellPadding);

  const verRefX = 100;
  const verRefY = 202;
  const verRefCellWidth = 100;
  const verRefCellPadding = 3;

  const refecrenceNo = componentList['caseId']

  // Set font properties
  doc.fontSize(7);

  const verRefMaxCellHeight = doc.currentLineHeight() + verRefCellPadding * 2

  doc.rect(verRefX, verRefY, verRefCellWidth, verRefMaxCellHeight).fillAndStroke('lightgray', 'black');
  doc.fillColor('black').text('VERIFACTS REFERENCE NO', verRefX + verRefCellPadding, verRefY + verRefCellPadding);

  doc.rect(verRefX + 100, verRefY, verRefCellWidth, verRefMaxCellHeight).fillAndStroke('white', 'black');
  doc.fillColor('black').text(refecrenceNo, verRefX + verRefCellWidth + verRefCellPadding, verRefY + verRefCellPadding);

  doc.rect(verRefX + 200, verRefY, verRefCellWidth, verRefMaxCellHeight).fillAndStroke('lightgray', 'black');
  doc.fillColor('black').text('CLIENT REFERENCE NO ', verRefX + verRefCellPadding + verRefCellWidth * 2, verRefY + verRefCellPadding);

  doc.rect(verRefX + 300, verRefY, verRefCellWidth, verRefMaxCellHeight).fillAndStroke('white', 'black');
  // doc.fillColor('black').text('CLIENT REFERENCE NO ', verRefX + verRefCellPadding + verRefCellWidth * 3, verRefY + verRefCellPadding);


  const insuffX = 100;
  const insuffY = 215;
  const insuffCellWidth = 100;
  const insuffCellPadding = 3;


  // for insuff raised date
  const insuffraiseddate = new Date(componentList['firstInsufficiencyRaisedDate'])
  const insuffraisedday = String(insuffraiseddate.getDate()).padStart(2, '0');
  const insuffraisedmonth = insuffraiseddate.toLocaleString('default', { month: 'long' });
  const insuffraisedyear = insuffraiseddate.getFullYear();
  const insuffRaisedDate = `${insuffraisedday}-${insuffraisedmonth}-${insuffraisedyear}`;
  // console.log(insuffRaisedDate);

  // for insuff cleared date
  const insuffCleareddate = new Date(componentList['lastInsufficiencyClearedDate'])
  const insuffClearedday = String(insuffCleareddate.getDate()).padStart(2, '0');
  const insuffClearedmonth = insuffCleareddate.toLocaleString('default', { month: 'long' });
  const insuffClearedyear = insuffCleareddate.getFullYear();
  const insuffClearedDate = `${insuffClearedday}-${insuffClearedmonth}-${insuffClearedyear}`;
  // console.log(insuffClearedDate);


  // Set font properties
  doc.fontSize(8);

  const insuffMaxCellHeight = doc.currentLineHeight() + insuffCellPadding * 2

  doc.rect(insuffX, insuffY, insuffCellWidth, insuffMaxCellHeight).fillAndStroke('lightgray', 'black');
  doc.fillColor('black').text('INSUFF RAISED ON', insuffX + insuffCellPadding, insuffY + insuffCellPadding);

  doc.rect(insuffX + 100, insuffY, insuffCellWidth, insuffMaxCellHeight).fillAndStroke('white', 'black');
  //doc.fillColor('black').text(insuffRaisedDate, insuffX + insuffCellPadding + insuffCellWidth, insuffY + insuffCellPadding);

  doc.rect(insuffX + 200, insuffY, insuffCellWidth, insuffMaxCellHeight).fillAndStroke('lightgray', 'black');
  doc.fillColor('black').text('INSUFF CLEARED ON', insuffX + insuffCellPadding + insuffCellWidth * 2, insuffY + insuffCellPadding);

  doc.rect(insuffX + 300, insuffY, insuffCellWidth, insuffMaxCellHeight).fillAndStroke('white', 'black');
  //doc.fillColor('black').text(insuffClearedDate, insuffX + insuffCellPadding + insuffCellWidth * 3, insuffY + insuffCellPadding);


  const levelX = 100;
  const levelY = 228;
  const levelCellWidth = 100;
  const levelCellPadding = 3;


  const colorGrade = await ColorMaster.findById(componentList.grade, { 'name': 1, '_id': 0, 'colorCode': 1 })
  console.log("color master ", componentList.grade1);


  // Set font properties
  doc.fontSize(8);

  const levelMaxCellHeight = doc.currentLineHeight() + levelCellPadding * 2

  doc.rect(levelX, levelY, levelCellWidth, levelMaxCellHeight).fillAndStroke('lightgray', 'black');
  doc.fillColor('black').text('LEVEL / ENTITY', levelX + levelCellPadding, levelY + levelCellPadding);

  doc.rect(levelX + 100, levelY, levelCellWidth, levelMaxCellHeight).fillAndStroke('white', 'black');
  // doc.fillColor('black').text('LEVEL / ENTITY', levelX + levelCellPadding + levelCellWidth, levelY + levelCellPadding);

  doc.rect(levelX + 200, levelY, levelCellWidth, levelMaxCellHeight).fillAndStroke('lightgray', 'black');
  doc.fillColor('black').text('COLOR CODE', levelX + levelCellPadding + levelCellWidth * 2, levelY + levelCellPadding);

  doc.rect(levelX + 300, levelY, levelCellWidth, levelMaxCellHeight).fillAndStroke('white', 'black');
  doc.fillColor(`${colorGrade?.colorCode ? colorGrade?.colorCode : "black"}`).text(colorGrade?.name, levelX + levelCellPadding + levelCellWidth * 3, levelY + levelCellPadding);

  // for discepency format 
  const discepencyX = 100;
  const discepencyY = 255;
  const discepencyCellWidth = 400 / 3;
  const discepencyCellPadding = 3;

  // Set font properties
  doc.fontSize(10);

  const discepencyMaxCellHeight = doc.currentLineHeight() + discepencyCellPadding * 2

  doc.rect(discepencyX, discepencyY, discepencyCellWidth, discepencyMaxCellHeight).fillAndStroke('white', 'red');
  doc.fillColor('red').text('Discrepancy', discepencyX + 32 + discepencyCellPadding, discepencyY + discepencyCellPadding);

  doc.rect(discepencyX + discepencyCellWidth, discepencyY, discepencyCellWidth, discepencyMaxCellHeight).fillAndStroke('white', 'orange');
  doc.fillColor('orange').text('Unable to Verify', discepencyX + 28 + discepencyCellPadding + discepencyCellWidth, discepencyY + discepencyCellPadding);

  doc.rect(discepencyX + 2 * discepencyCellWidth, discepencyY, discepencyCellWidth, discepencyMaxCellHeight).fillAndStroke('white', 'green');
  doc.fillColor('green').text('Clear Report', discepencyX + 32 + discepencyCellPadding + discepencyCellWidth * 2, discepencyY + discepencyCellPadding);


  //Executive summary

  const summaryX = 100;
  const summaryY = 285;
  const summaryCellWidth = 400;
  const summaryCellPadding = 3;

  // Set font properties
  doc.fontSize(10);

  const summaryMaxCellHeight = doc.currentLineHeight() + summaryCellPadding * 2

  doc.rect(summaryX, summaryY, summaryCellWidth, summaryMaxCellHeight).fillAndStroke('lightgray', 'black');
  doc.fillColor('black').text('EXECUTIVE SUMMARY', summaryX + 140 + summaryCellPadding, summaryY + summaryCellPadding);

  const checkX = 100;
  const checkY = 300;
  const checkCellWidth = 100;
  const checkCellPadding = 3;

  // Set font properties
  doc.fontSize(8);

  const checkMaxCellHeight = doc.currentLineHeight() + checkCellPadding * 2;

  doc.rect(checkX, checkY, checkCellWidth, checkMaxCellHeight).fillAndStroke('white', 'black');
  doc.fillColor('black').text('TYPES OF CHECK', checkX + 10 + checkCellPadding, checkY + checkCellPadding);

  doc.rect(checkX + 100, checkY, checkCellWidth + 100, levelMaxCellHeight).fillAndStroke('white', 'black');
  doc.fillColor('black').text('BRIEF DETAILS', checkX + 70 + checkCellPadding + checkCellWidth, checkY + checkCellPadding);

  doc.rect(checkX + 300, checkY, checkCellWidth, checkMaxCellHeight).fillAndStroke('white', 'black');
  doc.fillColor('black').text('BGC STATUS', checkX + 20 + checkCellWidth + checkCellPadding + checkCellWidth * 2, checkY + checkCellPadding);

  doc.font('Helvetica');
  // ----------------------------------------
  // add page numbers
  let pageNum = 2;

  // adding page number in summary page 
  doc.font('Helvetica').fillColor('Black').text('CONFIDENTIAL', centerX - 140, logoStartY + 730, { width: logoWidth, height: logoHeight });
  doc.font('Helvetica').fillColor('Black').text(`Page ${pageNum - 1}`, centerX + 50, logoStartY + 730, { width: logoWidth, height: logoHeight });



  for (let i = 0; i < componentsIdList.length; i++) {
    if (componentsData.hasOwnProperty(componentsIdList[i].name)) {

      if (componentsIdList[i].name === 'employment') {
        const employmentDoc = await employment.find({ case: mongoose.Types.ObjectId(componentList._id) }, { nameofemployer: 1, _id: 0, grade: 1 });

        componentsData[componentsIdList[i].name]['typeOfChecks'] = 'Employment Previous';

        const emparr = [];
        for (let i = 0; i < employmentDoc.length; i++) {
          const colorDoc = await ColorMaster.findById(employmentDoc[i].grade, { name: 1, colorCode: 1, _id: 0 })

          emparr.push({ ...colorDoc?.toObject(), briefDetails: employmentDoc[i].nameofemployer })
        }


        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;
      }


      else if (componentsIdList[i].name === 'education') {
        const educationDoc = await education.find({ case: mongoose.Types.ObjectId(componentList._id) }, { nameofuniversity: 1, _id: 0, grade: 1 })
        componentsData[componentsIdList[i].name]['typeOfChecks'] = 'Education';

        const emparr = [];
        for (let i = 0; i < educationDoc.length; i++) {
          const colorDoc = await ColorMaster.findById(educationDoc[i].grade, { name: 1, colorCode: 1, _id: 0 })

          emparr.push({ ...colorDoc?.toObject(), briefDetails: educationDoc[i].nameofuniversity })
        }

        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;
      }


      else if (componentsIdList[i].name === 'address') {
        const addressDoc = await address.find({ case: mongoose.Types.ObjectId(componentList._id) }, { address: 1, _id: 0, grade: 1 })
        componentsData[componentsIdList[i].name]['typeOfChecks'] = 'Address Present & Permanent';

        // getting the fields data of
        const emparr = [];
        for (let i = 0; i < addressDoc.length; i++) {
          const colorDoc = await ColorMaster.findById(addressDoc[i].grade, { name: 1, colorCode: 1, _id: 0 })
          emparr.push({ ...colorDoc?.toObject(), briefDetails: addressDoc[i].address })
        }
        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;

      }


      else if (componentsIdList[i].name === 'courtrecord') {
        const courtRecordDoc = await courtrecord.find({ case: mongoose.Types.ObjectId(componentList._id) }, { addresswithpin: 1, _id: 0, grade: 1 })
        componentsData[componentsIdList[i].name]['typeOfChecks'] = 'Court Record Present & Permanent';

        const emparr = [];
        for (let i = 0; i < courtRecordDoc.length; i++) {
          const colorDoc = await ColorMaster.findById(courtRecordDoc[i].grade, { name: 1, colorCode: 1, _id: 0 })
          emparr.push({ ...colorDoc?.toObject(), briefDetails: courtRecordDoc[i].addresswithpin })
        }

        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;
      }


      else if (componentsIdList[i].name === 'criminalrecord') {
        const criminalRecordDoc = await criminalrecord.find({ case: mongoose.Types.ObjectId(componentList._id) }, { fulladdress: 1, _id: 0, grade: 1 })
        componentsData[componentsIdList[i].name]['typeOfChecks'] = 'Criminal Record Check';

        const emparr = [];
        for (let i = 0; i < criminalRecordDoc.length; i++) {
          const colorDoc = await ColorMaster.findById(criminalRecordDoc[i].grade, { name: 1, colorCode: 1, _id: 0 })
          emparr.push({ ...colorDoc?.toObject(), briefDetails: criminalRecordDoc[i].fulladdress })
        }

        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;
      }


      else if (componentsIdList[i].name === 'identity') {
        const identityDoc = await identity.find({ case: mongoose.Types.ObjectId(componentList._id) }, { typeofid: 1, _id: 0, grade: 1 })
        componentsData[componentsIdList[i].name]['typeOfChecks'] = 'Identity Check';

        const emparr = [];
        for (let i = 0; i < identityDoc.length; i++) {
          const colorDoc = await ColorMaster.findById(identityDoc[i].grade, { name: 1, colorCode: 1, _id: 0 })
          emparr.push({ ...colorDoc?.toObject(), briefDetails: identityDoc[i].typeofid })
        }

        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;
      }


      else if (componentsIdList[i].name === 'creditcheck') {
        const creditcheckDoc = await creditcheck.find({ case: mongoose.Types.ObjectId(componentList._id) }, { typeofid: 1, _id: 0, grade: 1 })
        componentsData[componentsIdList[i].name]['typeOfChecks'] = 'Credit Check';

        const emparr = [];
        for (let i = 0; i < creditcheckDoc.length; i++) {
          const colorDoc = await ColorMaster.findById(creditcheckDoc[i].grade, { name: 1, colorCode: 1, _id: 0 })
          emparr.push({ ...colorDoc?.toObject(), briefDetails: creditcheckDoc[i].typeofid })
        }

        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;
      }


      else if (componentsIdList[i].name === "socialmedia") {
        const socialMediaDoc = await socialmedia.find({ case: mongoose.Types.ObjectId(componentList._id) }, { searchname: 1, _id: 0, grade: 1 })
        componentsData[componentsIdList[i].name]['typeOfChecks'] = 'Social Media Check';

        const emparr = [];
        for (let i = 0; i < socialMediaDoc.length; i++) {
          const colorDoc = await ColorMaster.findById(socialMediaDoc[i].grade, { name: 1, colorCode: 1, _id: 0 })
          emparr.push({ ...colorDoc?.toObject(), briefDetails: socialMediaDoc[i].searchname })
        }

        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;
      }


      else if (componentsIdList[i].name === "globaldatabase") {
        const globaldatabaseDoc = await globaldatabase.find({ case: mongoose.Types.ObjectId(componentList._id) }, { searchname: 1, _id: 0, grade: 1 })
        componentsData[componentsIdList[i].name]['typeOfChecks'] = 'Global Database';

        const emparr = [];
        for (let i = 0; i < globaldatabaseDoc.length; i++) {
          const colorDoc = await ColorMaster.findById(globaldatabaseDoc[i].grade, { name: 1, colorCode: 1, _id: 0 })
          emparr.push({ ...colorDoc?.toObject(), briefDetails: globaldatabaseDoc[i].searchname })
        }

        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;
      }


      else if (componentsIdList[i].name === "reference") {
        const referenceDoc = await reference.find({ case: mongoose.Types.ObjectId(componentList._id) }, { nameofreference: 1, _id: 0, grade: 1 })
        componentsData[componentsIdList[i].name]['typeOfChecks'] = 'Reference';

        const emparr = [];
        for (let i = 0; i < referenceDoc.length; i++) {
          const colorDoc = await ColorMaster.findById(referenceDoc[i].grade, { name: 1, colorCode: 1, _id: 0 })
          emparr.push({ ...colorDoc?.toObject(), briefDetails: referenceDoc[i].nameofreference })
        }

        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;
      }


      else if (componentsIdList[i].name === "drugtestfive") {
        const drugtestfiveDoc = await drugtestfive.find({ case: mongoose.Types.ObjectId(componentList._id) }, { nameofemployee: 1, _id: 0, grade: 1 })
        componentsData[componentsIdList[i].name]['typeOfChecks'] = 'Drug Test Five';

        const emparr = [];
        for (let i = 0; i < drugtestfiveDoc.length; i++) {
          const colorDoc = await ColorMaster.findById(drugtestfiveDoc[i].grade, { name: 1, colorCode: 1, _id: 0 })
          emparr.push({ ...colorDoc?.toObject(), briefDetails: drugtestfiveDoc[i].nameofemployee })
        }

        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;
      }


      else if (componentsIdList[i].name === "drugtestten") {
        const drugtesttenDoc = await drugtestten.find({ case: mongoose.Types.ObjectId(componentList._id) }, { nameofemployee: 1, _id: 0, grade: 1 })
        componentsData[componentsIdList[i].name]['typeOfChecks'] = 'Drug Test Ten';

        const emparr = [];
        for (let i = 0; i < drugtesttenDoc.length; i++) {
          const colorDoc = await ColorMaster.findById(drugtesttenDoc[i].grade, { name: 1, colorCode: 1, _id: 0 })
          emparr.push({ ...colorDoc?.toObject(), briefDetails: drugtesttenDoc[i].nameofemployee })
        }

        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;
      }


      else if (componentsIdList[i].name === "passport") {
        const passportDoc = await passport.find({ case: mongoose.Types.ObjectId(componentList._id) }, { givenname: 1, _id: 0, grade: 1 })
        componentsData[componentsIdList[i].name]['typeOfChecks'] = 'Passport Check';

        const emparr = [];
        for (let i = 0; i < passportDoc.length; i++) {
          const colorDoc = await ColorMaster.findById(passportDoc[i].grade, { name: 1, colorCode: 1, _id: 0 })
          emparr.push({ ...colorDoc?.toObject(), briefDetails: passportDoc[i].givenname })
        }

        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;
      }


      else if (componentsIdList[i].name === "addresstelephone") {
        const addresstelephoneDoc = await addresstelephone.find({ case: mongoose.Types.ObjectId(componentList._id) }, { address: 1, _id: 0, grade: 1 })
        componentsData[componentsIdList[i].name]['typeOfChecks'] = 'Address Check - Telephonic'; // ask

        const emparr = [];
        for (let i = 0; i < addresstelephoneDoc.length; i++) {
          const colorDoc = await ColorMaster.findById(addresstelephoneDoc[i].grade, { name: 1, colorCode: 1, _id: 0 })
          emparr.push({ ...colorDoc?.toObject(), briefDetails: addresstelephoneDoc[i].address })
        }

        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;
      }


      else if (componentsIdList[i].name === "addresscomprehensive") {
        const addresscomprehensiveDoc = await addresscomprehensive.find({ case: mongoose.Types.ObjectId(componentList._id) }, { fulladdress: 1, _id: 0, grade: 1 }) // ask for the data
        componentsData[componentsIdList[i].name]['typeOfChecks'] = 'Address Present & Permanent';

        const emparr = [];
        for (let i = 0; i < addresscomprehensiveDoc.length; i++) {
          const colorDoc = await ColorMaster.findById(addresscomprehensiveDoc[i].grade, { name: 1, colorCode: 1, _id: 0 })
          emparr.push({ ...colorDoc?.toObject(), briefDetails: addresscomprehensiveDoc[i].fulladdress })
        }

        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;
      }


      else if (componentsIdList[i].name === "educationcomprehensive") {
        const educationcomprehensiveDoc = await educationcomprehensive.find({ case: mongoose.Types.ObjectId(componentList._id) }, { nameofuniversity: 1, _id: 0, grade: 1 })
        componentsData[componentsIdList[i].name]['typeOfChecks'] = 'Education'; // ask

        const emparr = [];
        for (let i = 0; i < educationcomprehensiveDoc.length; i++) {
          const colorDoc = await ColorMaster.findById(educationcomprehensiveDoc[i].grade, { name: 1, colorCode: 1, _id: 0 })
          emparr.push({ ...colorDoc?.toObject(), briefDetails: educationcomprehensiveDoc[i].nameofuniversity })
        }

        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;
      }


      else if (componentsIdList[i].name === "educationadvanced") {
        const educationadvancedDoc = await educationadvanced.find({ case: mongoose.Types.ObjectId(componentList._id) }, { nameofuniverskty: 1, _id: 0, grade: 1 })
        componentsData[componentsIdList[i].name]['typeOfChecks'] = 'Education'; // ask

        const emparr = [];
        for (let i = 0; i < educationadvancedDoc.length; i++) {
          const colorDoc = await ColorMaster.findById(educationadvancedDoc[i].grade, { name: 1, colorCode: 1, _id: 0 })
          emparr.push({ ...colorDoc?.toObject(), briefDetails: educationadvancedDoc[i].nameofuniverskty })
        }

        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;
      }

      // no data for the collection of drug test six 

      else if (componentsIdList[i].name === "drugtestsix") {
        const drugtestsixDoc = await drugtestsix.find({ case: mongoose.Types.ObjectId(componentList._id) }, { _id: 0, grade: 1 }) // ask for field which is required
        componentsData[componentsIdList[i].name]['typeOfChecks'] = 'Drug Test 6 Panel';

        const emparr = [];
        for (let i = 0; i < drugtestsixDoc.length; i++) {
          const colorDoc = await ColorMaster.findById(drugtestsixDoc[i].grade, { name: 1, colorCode: 1, _id: 0 })
          emparr.push({ ...colorDoc?.toObject(), briefDetails: drugtestsixDoc[i].nameofuniverskty })
        }

        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;
      }


      else if (componentsIdList[i].name === "drugtestseven") {
        const drugtestsevenDoc = await drugtestseven.find({ case: mongoose.Types.ObjectId(componentList._id) }, { nameofemploybee: 1, _id: 0, grade: 1 })
        componentsData[componentsIdList[i].name]['typeOfChecks'] = 'Drug Test 7 Panel';

        const emparr = [];
        for (let i = 0; i < drugtestsevenDoc.length; i++) {
          const colorDoc = await ColorMaster.findById(drugtestsevenDoc[i].grade, { name: 1, colorCode: 1, _id: 0 })
          emparr.push({ ...colorDoc?.toObject(), briefDetails: drugtestsevenDoc[i].nameofemploybee })
        }

        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;
      }


      else if (componentsIdList[i].name === "drugtesteight") {
        const drugtesteightDoc = await drugtesteight.find({ case: mongoose.Types.ObjectId(componentList._id) }, { nameofemploybee: 1, _id: 0, grade: 1 })
        componentsData[componentsIdList[i].name]['typeOfChecks'] = 'Drug Test 8 Panel';

        const emparr = [];
        for (let i = 0; i < drugtesteightDoc.length; i++) {
          const colorDoc = await ColorMaster.findById(drugtesteightDoc[i].grade, { name: 1, colorCode: 1, _id: 0 })
          emparr.push({ ...colorDoc?.toObject(), briefDetails: drugtesteightDoc[i].nameofemploybee })
        }

        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;
      }


      else if (componentsIdList[i].name === "drugtestnine") {
        const drugtestnineDoc = await drugtestnine.find({ case: mongoose.Types.ObjectId(componentList._id) }, { nameofemployee: 1, _id: 0, grade: 1 })
        componentsData[componentsIdList[i].name]['typeOfChecks'] = 'Drug Test 9 Panel';

        const emparr = [];
        for (let i = 0; i < drugtestnineDoc.length; i++) {
          const colorDoc = await ColorMaster.findById(drugtestnineDoc[i].grade, { name: 1, colorCode: 1, _id: 0 })
          emparr.push({ ...colorDoc?.toObject(), briefDetails: drugtestnineDoc[i].nameofemployee })
        }

        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;
      }


      else if (componentsIdList[i].name === "facisl3") {
        const facisl3Doc = await facisl3.find({ case: mongoose.Types.ObjectId(componentList._id) }, { applicantname: 1, _id: 0, grade: 1 })
        componentsData[componentsIdList[i].name]['typeOfChecks'] = 'FACIS L3 Check';

        const emparr = [];
        for (let i = 0; i < facisl3Doc.length; i++) {
          const colorDoc = await ColorMaster.findById(facisl3Doc[i].grade, { name: 1, colorCode: 1, _id: 0 })
          emparr.push({ ...colorDoc?.toObject(), briefDetails: facisl3Doc[i].applicantname })
        }

        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;
      }


      else if (componentsIdList[i].name === "credittrans") {
        const credittransDoc = await credittrans.find({ case: mongoose.Types.ObjectId(componentList._id) }, { credittrans: 1, _id: 0, grade: 1 })
        componentsData[componentsIdList[i].name]['typeOfChecks'] = 'Credit Check - TransUnion'; // ask

        const emparr = [];
        for (let i = 0; i < credittransDoc.length; i++) {
          const colorDoc = await ColorMaster.findById(credittransDoc[i].grade, { name: 1, colorCode: 1, _id: 0 })
          emparr.push({ ...colorDoc?.toObject(), briefDetails: credittransDoc[i].credittrans })
        }

        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;
      }


      else if (componentsIdList[i].name === "creditequifax") {
        const creditequifaxDoc = await creditequifax.find({ case: mongoose.Types.ObjectId(componentList._id) }, { panname: 1, _id: 0, grade: 1 })
        componentsData[componentsIdList[i].name]['typeOfChecks'] = 'Credit Check - Equifax'; // ask

        const emparr = [];
        for (let i = 0; i < creditequifaxDoc.length; i++) {
          const colorDoc = await ColorMaster.findById(creditequifaxDoc[i].grade, { name: 1, colorCode: 1, _id: 0 })
          emparr.push({ ...colorDoc?.toObject(), briefDetails: creditequifaxDoc[i].panname })
        }

        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;
      }


      else if (componentsIdList[i].name === "empadvance") {
        const empadvanceDoc = await empadvance.find({ case: mongoose.Types.ObjectId(componentList._id) }, { nameofemployer: 1, _id: 0, grade: 1 })
        componentsData[componentsIdList[i].name]['typeOfChecks'] = 'Employment Previous'; // ask

        const emparr = [];
        for (let i = 0; i < empadvanceDoc.length; i++) {
          const colorDoc = await ColorMaster.findById(empadvanceDoc[i].grade, { name: 1, colorCode: 1, _id: 0 })
          emparr.push({ ...colorDoc?.toObject(), briefDetails: empadvanceDoc[i].nameofemployer })
        }

        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;
      }


      else if (componentsIdList[i].name === "empbasic") {
        const empbasicDoc = await empbasic.find({ case: mongoose.Types.ObjectId(componentList._id) }, { nameofemployer: 1, _id: 0, grade: 1 })
        componentsData[componentsIdList[i].name]['typeOfChecks'] = 'Employment Previous'; // ask

        const emparr = [];
        for (let i = 0; i < empbasicDoc.length; i++) {
          const colorDoc = await ColorMaster.findById(empbasicDoc[i].grade, { name: 1, colorCode: 1, _id: 0 })
          emparr.push({ ...colorDoc?.toObject(), briefDetails: empbasicDoc[i].nameofemployer })
        }

        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;
      }


      else if (componentsIdList[i].name === "vddadvance") {
        const vddadvanceDoc = await vddadvance.find({ case: mongoose.Types.ObjectId(componentList._id) }, { companyname: 1, _id: 0, grade: 1 })
        componentsData[componentsIdList[i].name]['typeOfChecks'] = 'Vendor Due Diligence';

        const emparr = [];
        for (let i = 0; i < vddadvanceDoc.length; i++) {
          const colorDoc = await ColorMaster.findById(vddadvanceDoc[i].grade, { name: 1, colorCode: 1, _id: 0 })
          emparr.push({ ...colorDoc?.toObject(), briefDetails: vddadvanceDoc[i].companyname })
        }

        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;
      }


      else if (componentsIdList[i].name === "dlcheck") {
        const dlcheckDoc = await dlcheck.find({ case: mongoose.Types.ObjectId(componentList._id) }, { nameperdl: 1, _id: 0, grade: 1 })
        componentsData[componentsIdList[i].name]['typeOfChecks'] = 'Driving License Check';

        const emparr = [];
        for (let i = 0; i < dlcheckDoc.length; i++) {
          const colorDoc = await ColorMaster.findById(dlcheckDoc[i].grade, { name: 1, colorCode: 1, _id: 0 })
          emparr.push({ ...colorDoc?.toObject(), briefDetails: dlcheckDoc[i].nameperdl })
        }

        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;
      }


      else if (componentsIdList[i].name === "voterid") {
        const voteridDoc = await voterid.find({ case: mongoose.Types.ObjectId(componentList._id) }, { epicname: 1, _id: 0, grade: 1 })
        componentsData[componentsIdList[i].name]['typeOfChecks'] = 'Voter ID Check';

        const emparr = [];
        for (let i = 0; i < voteridDoc.length; i++) {
          const colorDoc = await ColorMaster.findById(voteridDoc[i].grade, { name: 1, colorCode: 1, _id: 0 })
          emparr.push({ ...colorDoc?.toObject(), briefDetails: voteridDoc[i].epicname })
        }

        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;
      }


      else if (componentsIdList[i].name === "ofac") {
        const ofacDoc = await ofac.find({ case: mongoose.Types.ObjectId(componentList._id) }, { candname: 1, _id: 0, grade: 1 })
        componentsData[componentsIdList[i].name]['typeOfChecks'] = 'FAC Check';

        const emparr = [];
        for (let i = 0; i < ofacDoc.length; i++) {
          const colorDoc = await ColorMaster.findById(ofacDoc[i].grade, { name: 1, colorCode: 1, _id: 0 })
          emparr.push({ ...colorDoc?.toObject(), briefDetails: ofacDoc[i].candname })
        }

        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;
      }


      else if (componentsIdList[i].name === "gapvfn") {
        const gapvfnDoc = await gapvfn.find({ case: mongoose.Types.ObjectId(componentList._id) }, { tenureofgap: 1, _id: 0, grade: 1 })
        componentsData[componentsIdList[i].name]['typeOfChecks'] = 'GAP Verification';

        const emparr = [];
        for (let i = 0; i < gapvfnDoc.length; i++) {
          const colorDoc = await ColorMaster.findById(gapvfnDoc[i].grade, { name: 1, colorCode: 1, _id: 0 })
          emparr.push({ ...colorDoc?.toObject(), briefDetails: gapvfnDoc[i].tenureofgap })
        }

        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;
      }


      else if (componentsIdList[i].name === "bankstmt") {
        const bankstmtDoc = await bankstmt.find({ case: mongoose.Types.ObjectId(componentList._id) }, { nameofbank: 1, _id: 0, grade: 1 })
        componentsData[componentsIdList[i].name]['typeOfChecks'] = 'Bank Statement Verification';

        const emparr = [];
        for (let i = 0; i < bankstmtDoc.length; i++) {
          const colorDoc = await ColorMaster.findById(bankstmtDoc[i].grade, { name: 1, colorCode: 1, _id: 0 })
          emparr.push({ ...colorDoc?.toObject(), briefDetails: bankstmtDoc[i].nameofbank })
        }

        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;
      }


      else if (componentsIdList[i].name === "refbasic") {
        const refbasicDoc = await refbasic.find({ case: mongoose.Types.ObjectId(componentList._id) }, { name: 1, _id: 0, grade: 1 })
        componentsData[componentsIdList[i].name]['typeOfChecks'] = 'Reference Check'; // ask

        const emparr = [];
        for (let i = 0; i < refbasicDoc.length; i++) {
          const colorDoc = await ColorMaster.findById(refbasicDoc[i].grade, { name: 1, colorCode: 1, _id: 0 })
          emparr.push({ ...colorDoc?.toObject(), briefDetails: refbasicDoc[i].name })
        }

        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;
      }


      else if (componentsIdList[i].name === "directorshipcheck") {
        const directorshipcheckDoc = await directorshipcheck.find({ case: mongoose.Types.ObjectId(componentList._id) }, { directorname: 1, _id: 0, grade: 1 })
        componentsData[componentsIdList[i].name]['typeOfChecks'] = 'Directorship Check';

        const emparr = [];
        for (let i = 0; i < directorshipcheckDoc.length; i++) {
          const colorDoc = await ColorMaster.findById(directorshipcheckDoc[i].grade, { name: 1, colorCode: 1, _id: 0 })
          emparr.push({ ...colorDoc?.toObject(), briefDetails: directorshipcheckDoc[i].directorname })
        }

        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;
      }


      else if (componentsIdList[i].name === "addressonline") {
        const addressonlineDoc = await addressonline.find({ case: mongoose.Types.ObjectId(componentList._id) }, { fulladdwithpin: 1, _id: 0, grade: 1 })
        componentsData[componentsIdList[i].name]['typeOfChecks'] = 'Address Check - Online'; // ask

        const emparr = [];
        for (let i = 0; i < addressonlineDoc.length; i++) {
          const colorDoc = await ColorMaster.findById(addressonlineDoc[i].grade, { name: 1, colorCode: 1, _id: 0 })
          emparr.push({ ...colorDoc?.toObject(), briefDetails: addressonlineDoc[i].fulladdwithpin })
        }

        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;
      }


      else if (componentsIdList[i].name === "physostan") {
        const physostanDoc = await physostan.find({ case: mongoose.Types.ObjectId(componentList._id) }, { name: 1, _id: 0, grade: 1 })
        componentsData[componentsIdList[i].name]['typeOfChecks'] = 'physostan';

        const emparr = [];
        for (let i = 0; i < physostanDoc.length; i++) {
          const colorDoc = await ColorMaster.findById(physostanDoc[i].grade, { name: 1, colorCode: 1, _id: 0 })
          emparr.push({ ...colorDoc?.toObject(), briefDetails: physostanDoc[i].name })
        }

        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;
      }


      else if (componentsIdList[i].name === "sitecheck") {
        const sitecheckDoc = await sitecheck.find({ case: mongoose.Types.ObjectId(componentList._id) }, { name: 1, _id: 0, grade: 1 })
        componentsData[componentsIdList[i].name]['typeOfChecks'] = 'Site Check';

        const emparr = [];
        for (let i = 0; i < sitecheckDoc.length; i++) {
          const colorDoc = await ColorMaster.findById(sitecheckDoc[i].grade, { name: 1, colorCode: 1, _id: 0 })
          emparr.push({ ...colorDoc?.toObject(), briefDetails: sitecheckDoc[i].name })
        }

        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;
      }


      else if (componentsIdList[i].name === "uan") {
        const uanDoc = await uan.find({ case: mongoose.Types.ObjectId(componentList._id) }, { candidatename: 1, _id: 0, grade: 1 })
        componentsData[componentsIdList[i].name]['typeOfChecks'] = 'UAN Check';

        const emparr = [];
        for (let i = 0; i < uanDoc.length; i++) {
          const colorDoc = await ColorMaster.findById(uanDoc[i].grade, { name: 1, colorCode: 1, _id: 0 })
          emparr.push({ ...colorDoc?.toObject(), briefDetails: uanDoc[i].candidatename })
        }

        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;
      }


      else if (componentsIdList[i].name === "addressbusiness") {
        const addressbusinessDoc = await addressbusiness.find({ case: mongoose.Types.ObjectId(componentList._id) }, { city: 1, _id: 0, grade: 1 })
        componentsData[componentsIdList[i].name]['typeOfChecks'] = 'Address Verification - Business'; // ask

        const emparr = [];
        for (let i = 0; i < addressbusinessDoc.length; i++) {
          const colorDoc = await ColorMaster.findById(addressbusinessDoc[i].grade, { name: 1, colorCode: 1, _id: 0 })
          emparr.push({ ...colorDoc?.toObject(), briefDetails: addressbusinessDoc[i].city })
        }

        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;
      }


      else if (componentsIdList[i].name === "caconfirmation") {
        const caconfirmationDoc = await caconfirmation.find({ case: mongoose.Types.ObjectId(componentList._id) }, { companyname: 1, _id: 0, grade: 1 })
        componentsData[componentsIdList[i].name]['typeOfChecks'] = 'CA Confirmation';

        const emparr = [];
        for (let i = 0; i < caconfirmationDoc.length; i++) {
          const colorDoc = await ColorMaster.findById(caconfirmationDoc[i].grade, { name: 1, colorCode: 1, _id: 0 })
          emparr.push({ ...colorDoc?.toObject(), briefDetails: caconfirmationDoc[i].companyname })
        }


        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;
      }


      else if (componentsIdList[i].name === "colorblindness") {
        const colorblindnessDoc = await colorblindness.find({ case: mongoose.Types.ObjectId(componentList._id) }, { colorblindness: 1, _id: 0, grade: 1 })
        componentsData[componentsIdList[i].name]['typeOfChecks'] = 'Color Blindness Test';

        const emparr = [];
        for (let i = 0; i < colorblindnessDoc.length; i++) {
          const colorDoc = await ColorMaster.findById(colorblindnessDoc[i].grade, { name: 1, colorCode: 1, _id: 0 })
          emparr.push({ ...colorDoc?.toObject(), briefDetails: colorblindnessDoc[i].colorblindness })
        }

        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;
      }

      else if (componentsIdList[i].name === "exitinterview") {
        const exitinterviewDoc = await exitinterview.find({ case: mongoose.Types.ObjectId(componentList._id) }, { candidatename: 1, _id: 0, grade: 1 })
        componentsData[componentsIdList[i].name]['typeOfChecks'] = 'Exit Interview';

        const emparr = [];
        for (let i = 0; i < exitinterviewDoc.length; i++) {
          const colorDoc = await ColorMaster.findById(exitinterviewDoc[i].grade, { name: 1, colorCode: 1, _id: 0 })
          emparr.push({ ...colorDoc?.toObject(), briefDetails: exitinterviewDoc[i].candidatename })
        }

        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;
      }


      // no collection for epfo
      else if (componentsIdList[i].name === "epfo") {
        const epfoDoc = await EPFO.find({ case: mongoose.Types.ObjectId(componentList._id) }, { candidatename: 1, _id: 0, grade: 1 })
        componentsData[componentsIdList[i].name]['typeOfChecks'] = 'EPFO Check';

        const emparr = [];
        for (let i = 0; i < epfoDoc.length; i++) {
          const colorDoc = await ColorMaster.findById(epfoDoc[i].grade, { name: 1, colorCode: 1, _id: 0 })
          emparr.push({ ...colorDoc?.toObject(), briefDetails: epfoDoc[i].candidatename })
        }

        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;
      }


      else if (componentsIdList[i].name === "twentysixas") {
        const twentysixasDoc = await twentysixas.find({ case: mongoose.Types.ObjectId(componentList._id) }, { candidatename: 1, _id: 0, grade: 1 })
        componentsData[componentsIdList[i].name]['typeOfChecks'] = '26 AS';

        const emparr = [];
        for (let i = 0; i < twentysixasDoc.length; i++) {
          const colorDoc = await ColorMaster.findById(twentysixasDoc[i].grade, { name: 1, colorCode: 1, _id: 0 })
          emparr.push({ ...colorDoc?.toObject(), briefDetails: twentysixasDoc[i].candidatename })
        }

        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;
      }


      else if (componentsIdList[i].name === "vdddeclaration") {
        const vdddeclarationDoc = await vdddeclaration.find({ case: mongoose.Types.ObjectId(componentList._id) }, { familymemberoneoccupation: 1, _id: 0, grade: 1 })
        componentsData[componentsIdList[i].name]['typeOfChecks'] = 'VDD - Declaration';

        const emparr = [];
        for (let i = 0; i < vdddeclarationDoc.length; i++) {
          const colorDoc = await ColorMaster.findById(vdddeclarationDoc[i].grade, { name: 1, colorCode: 1, _id: 0 })
          emparr.push({ ...colorDoc?.toObject(), briefDetails: vdddeclarationDoc[i].familymemberoneoccupation })
        }

        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;
      }


      else if (componentsIdList[i].name === "tcsvdd") {
        const tcsvddDoc = await tcsvdd.find({ case: mongoose.Types.ObjectId(componentList._id) }, { companyname: 1, _id: 0, grade: 1 })
        componentsData[componentsIdList[i].name]['typeOfChecks'] = 'TCS - VDD';

        const emparr = [];
        for (let i = 0; i < tcsvddDoc.length; i++) {
          const colorDoc = await ColorMaster.findById(tcsvddDoc[i].grade, { name: 1, colorCode: 1, _id: 0 })
          emparr.push({ ...colorDoc?.toObject(), briefDetails: tcsvddDoc[i].companyname })
        }

        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;
      }


      else if (componentsIdList[i].name === "cvanalysis") {
        const cvanalysisDoc = await cvanalysis.find({ case: mongoose.Types.ObjectId(componentList._id) }, { companyname: 1, _id: 0, grade: 1 })
        componentsData[componentsIdList[i].name]['typeOfChecks'] = 'CV Analysis';

        const emparr = [];
        for (let i = 0; i < cvanalysisDoc.length; i++) {
          const colorDoc = await ColorMaster.findById(cvanalysisDoc[i].grade, { name: 1, colorCode: 1, _id: 0 })
          emparr.push({ ...colorDoc?.toObject(), briefDetails: cvanalysisDoc[i].companyname })
        }

        componentsData[componentsIdList[i].name]['DetailsAndStatus'] = emparr;
      }

      for (let j = 0; j < componentsData[componentsIdList[i].name].DetailsAndStatus?.length; j++) {

        let textHeight = doc.heightOfString(componentsData[componentsIdList[i].name].typeOfChecks, { width: checkCellWidth });

        const cellHeight = Math.max(textHeight, checkMaxCellHeight);

        const initialY = 313.5;

        let currentY = initialY;

        for (let k = 0; k < componentsIdList.length; k++) {
          const componentName = componentsIdList[k].name;

          // Check if the component exists in componentsData
          if (componentsData.hasOwnProperty(componentName)) {
            const componentData = componentsData[componentName];
            const detailsAndStatus = componentData.DetailsAndStatus;

            if (detailsAndStatus) {
              let maxDetailsStatusHeight = 0;

              for (let j = 0; j < detailsAndStatus.length; j++) {
                const text2Height = doc.heightOfString(detailsAndStatus[j].briefDetails, { width: checkCellWidth });
                const text3Height = doc.heightOfString(detailsAndStatus[j].name, { width: checkCellWidth });

                const detailsStatusHeight = Math.max(text2Height, text3Height);
                maxDetailsStatusHeight = Math.max(textHeight, detailsStatusHeight);
              }

              const totalCellHeight = Math.max(cellHeight, maxDetailsStatusHeight + checkCellPadding * 2);

              for (let j = 0; j < detailsAndStatus.length; j++) {
                const textX1 = checkX;
                const textX2 = checkX + checkCellWidth;
                const textX3 = checkX + 3 * checkCellWidth;

                const textY1 = currentY + checkCellPadding;
                const textY2 = currentY + checkCellPadding;
                const textY3 = currentY + checkCellPadding;


                doc.rect(textX1, currentY, checkCellWidth, totalCellHeight).fillAndStroke('white', 'black');
                doc.fillColor('black').text(componentData.typeOfChecks, textX1 + checkCellPadding, textY1 + checkCellPadding,
                  {
                    width: checkCellWidth - checkCellPadding * 2,
                    height: totalCellHeight - checkCellPadding * 2,
                    align: 'center',
                    valign: 'center',
                    lineBreak: false,
                  }
                );

                doc.rect(textX2, currentY, 2 * checkCellWidth, totalCellHeight).fillAndStroke('white', 'black');

                doc.fillColor('black').text(
                  detailsAndStatus[j].briefDetails,
                  textX2 + checkCellPadding * 2,
                  textY2 + checkCellPadding,
                  {
                    width: checkCellWidth * 1.8,
                    height: totalCellHeight - checkCellPadding * 2,
                    align: 'center',
                    valign: 'center',
                    lineBreak: false,
                  }
                );

                doc.rect(textX3, currentY, checkCellWidth, totalCellHeight).fillAndStroke('white', 'black');

                doc.fillColor(detailsAndStatus[j].colorCode).text(detailsAndStatus[j].name, textX3 + checkCellPadding, textY3 + checkCellPadding,
                  {
                    width: checkCellWidth - checkCellPadding * 2,
                    height: totalCellHeight - checkCellPadding * 2,
                    align: 'center',
                    valign: 'center',
                    lineBreak: false,
                  }
                );
              }


              // Increment the currentY position for the next component
              currentY += totalCellHeight;
            }
          }
        }


        currentY += cellHeight;
      }

    }

  }


  // <---------------------<
  console.log("actual components are: ", componentList.actualComponents);
  let index = 0;


  for (let component of componentList.actualComponents) {

    let queryFilter = { '_id': 0 };
    for (let field of componentsData[component].fields) {
      Object.assign(queryFilter, field)
    }

    // console.log("query filter is: ", queryFilter);

    let jpgFiles = [];

    if (component === "employment") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("employment") - index);
      const employmentData = await employment.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = employmentData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, employmentData[docIndex]._id);

    } else if (component === "education") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("education") - index);
      const educationData = await education.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = educationData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, educationData[docIndex]._id);

    } else if (component === "address") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("address") - index);
      const addressData = await address.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = addressData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, addressData[docIndex]._id);
      console.log("datat@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@",jpgFiles )


    } else if (component === "courtrecord") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("courtrecord") - index);
      const courtRecordData = await courtrecord.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = courtRecordData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, courtRecordData[docIndex]._id);

    } else if (component === "criminalrecord") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("criminalrecord") - index);
      const criminalRecordData = await criminalrecord.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = criminalRecordData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, criminalRecordData[docIndex]._id);


    } else if (component === "identity") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("identity") - index);
      const identityData = await identity.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = identityData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, identityData[docIndex]._id);

    } else if (component === "creditcheck") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("creditcheck") - index);
      const creditCheckData = await creditcheck.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = creditCheckData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, creditCheckData[docIndex]._id);


    } else if (component === "socialmedia") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("socialmedia") - index);
      const socialMediaData = await socialmedia.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = socialMediaData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, socialMediaData[docIndex]._id);


    } else if (component === "globaldatabase") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("globaldatabase") - index);
      const globaldatabaseData = await globaldatabase.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = globaldatabaseData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, globaldatabaseData[docIndex]._id);

    } else if (component === "reference") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("reference") - index);
      const referenceData = await reference.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = referenceData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, referenceData[docIndex]._id);

    } else if (component === "drugtestfive") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("drugtestfive") - index);
      const drugtestfiveData = await drugtestfive.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = drugtestfiveData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, drugtestfiveData[docIndex]._id);


    } else if (component === "drugtestten") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("drugtestten") - index);
      const drugtesttenData = await drugtestten.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = drugtesttenData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, drugtesttenData[docIndex]._id);


    } else if (component === "passport") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("passport") - index);
      const passportData = await passport.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = passportData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, passportData[docIndex]._id);


    } else if (component === "addresstelephone") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("addresstelephone") - index);
      const addresstelephoneData = await addresstelephone.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = addresstelephoneData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, addresstelephoneData[docIndex]._id);


    } else if (component === "addresscomprehensive") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("addresscomprehensive") - index);
      const addresscomprehensiveData = await addresscomprehensive.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = addresscomprehensiveData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, addresscomprehensiveData[docIndex]._id);


    } else if (component === "educationcomprehensive") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("educationcomprehensive") - index);
      const educationcomprehensiveData = await educationcomprehensive.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = educationcomprehensiveData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, educationcomprehensiveData[docIndex]._id);


    } else if (component === "educationadvanced") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("educationadvanced") - index);
      const educationadvancedData = await educationadvanced.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = educationadvancedData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, educationadvancedData[docIndex]._id);


    } else if (component === "drugtestsix") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("drugtestsix") - index);
      const drugtestsixData = await drugtestsix.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = drugtestsixData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, drugtestsixData[docIndex]._id);


    } else if (component === "drugtestseven") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("drugtestseven") - index);
      const drugtestsevenData = await drugtestseven.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = drugtestsevenData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, drugtestsevenData[docIndex]._id);


    } else if (component === "drugtesteight") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("drugtesteight") - index);
      const drugtesteightData = await drugtesteight.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = drugtesteightData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, drugtesteightData[docIndex]._id);


    } else if (component === "drugtestnine") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("drugtestnine") - index);
      const drugtestnineData = await drugtestnine.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = drugtestnineData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, drugtestnineData[docIndex]._id);


    } else if (component === "facisl3") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("facisl3") - index);
      const facisl3Data = await facisl3.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = facisl3Data[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, facisl3Data[docIndex]._id);


    } else if (component === "credittrans") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("credittrans") - index);
      const credittransData = await credittrans.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = credittransData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, credittransData[docIndex]._id);


    } else if (component === "creditequifax") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("creditequifax") - index);
      const creditequifaxData = await creditequifax.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = creditequifaxData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, creditequifaxData[docIndex]._id);


    } else if (component === "empadvance") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("empadvance") - index);
      const empadvanceData = await empadvance.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = empadvanceData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, empadvanceData[docIndex]._id);


    } else if (component === "empbasic") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("empbasic") - index);
      const empbasicData = await empbasic.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = empbasicData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, empbasicData[docIndex]._id);


    } else if (component === "vddadvance") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("vddadvance") - index);
      const vddadvanceData = await vddadvance.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = vddadvanceData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, vddadvanceData[docIndex]._id);


    } else if (component === "dlcheck") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("dlcheck") - index);
      const dlcheckData = await dlcheck.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = dlcheckData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, dlcheckData[docIndex]._id);


    } else if (component === "voterid") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("voterid") - index);
      const voteridData = await voterid.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = voteridData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, voteridData[docIndex]._id);


    } else if (component === "ofac") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("ofac") - index);
      const ofacData = await ofac.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = ofacData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, ofacData[docIndex]._id);


    } else if (component === "gapvfn") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("gapvfn") - index);
      const gapvfnData = await gapvfn.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = gapvfnData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, gapvfnData[docIndex]._id);

    } else if (component === "bankstmt") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("bankstmt") - index);
      const bankstmtData = await bankstmt.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = bankstmtData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, bankstmtData[docIndex]._id);


    } else if (component === "refbasic") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("refbasic") - index);
      const refbasicData = await refbasic.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = refbasicData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, refbasicData[docIndex]._id);


    } else if (component === "directorshipcheck") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("directorshipcheck") - index);
      const directorshipcheckData = await directorshipcheck.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = directorshipcheckData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, directorshipcheckData[docIndex]._id);


    } else if (component === "addressonline") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("addressonline") - index);
      const addressonlineData = await addressonline.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = addressonlineData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, addressonlineData[docIndex]._id);


    } else if (component === "physostan") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("physostan") - index);
      const physostanData = await physostan.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = physostanData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, physostanData[docIndex]._id);


    } else if (component === "sitecheck") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("sitecheck") - index);
      const sitecheckData = await sitecheck.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = sitecheckData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, sitecheckData[docIndex]._id);


    } else if (component === "uan") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("uan") - index);
      const uanData = await uan.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = uanData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, uanData[docIndex]._id);


    } else if (component === "addressbusiness") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("addressbusiness") - index);
      const addressbusinessData = await addressbusiness.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = addressbusinessData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, addressbusinessData[docIndex]._id);


    } else if (component === "caconfirmation") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("caconfirmation") - index);
      const caconfirmationData = await caconfirmation.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = caconfirmationData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, caconfirmationData[docIndex]._id);


    } else if (component === "colorblindness") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("colorblindness") - index);
      const colorblindnessnData = await colorblindness.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = colorblindnessnData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, colorblindnessnData[docIndex]._id);


    } else if (component === "exitinterview") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("exitinterview") - index);
      const exitinterviewData = await exitinterview.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = exitinterviewData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, exitinterviewData[docIndex]._id);


    } else if (component === "epfo") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("epfo") - index);
      const epfoData = await EPFO.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = epfoData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, epfoData[docIndex]._id);


    } else if (component === "twentysixas") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("twentysixas") - index);
      const twentysixasData = await twentysixas.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = twentysixasData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, twentysixasData[docIndex]._id);


    } else if (component === "vddpv") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("vddpv") - index);
      const vddpvData = await vddpv.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = vddpvData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, vddpvData[docIndex]._id);


    } else if (component === "vdddeclaration") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("vdddeclaration") - index);
      const vdddeclarationData = await vdddeclaration.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = vdddeclarationData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, vdddeclarationData[docIndex]._id);


    } else if (component === "tcsvdd") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("tcsvdd") - index);
      const tcsvddData = await tcsvdd.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = tcsvddData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, tcsvddData[docIndex]._id);


    } else if (component === "cvanalysis") {
      const docIndex = Math.abs(componentList.actualComponents.indexOf("cvanalysis") - index);
      const cvanalysisData = await cvanalysis.find({ case: caseId }, queryFilter);
      componentsData[component]['fieldsValue'] = cvanalysisData[docIndex].toObject();

      jpgFiles = await getJpgFiles(componentList, component, cvanalysisData[docIndex]._id);


    }

    doc.addPage();

    doc.image(logoPath, centerX, logoStartY, { width: logoWidth, height: logoHeight });
    // doc.addPage();
    // Add the page number to the document
    doc.font('Helvetica').fontSize(10).text(`Page ${pageNum}`, centerX + 50, logoStartY + 730, { width: logoWidth, height: logoHeight });

    pageNum++;


    // adding stamp on every component page
    doc.image(stampPath, stampX + 160, stampStartY, { width: stampWidth, height: stampHeight })
    doc.font('Helvetica').fontSize(10).text("For Verifacts Services Pvt. Ltd.", centerX + 180, logoStartY + 640, { width: logoWidth, height: logoHeight });
    doc.font('Helvetica').fontSize(10).text("Client Account Manager", centerX + 180, logoStartY + 730, { width: logoWidth, height: logoHeight });



    //adding "confidential" text
    doc.font('Helvetica').fillColor('Black').text('CONFIDENTIAL', centerX - 140, logoStartY + 730, { width: logoWidth, height: logoHeight });


    const fontSize = 10;

    const textWidth = doc.widthOfString(componentsData[component].displayName, { fontSize });
    const textHeight = doc.currentLineHeight();

    const textX = (pageWidth - textWidth) / 2;
    const textY = (pageHeight - textHeight) / 2;

    doc.font('Helvetica').fontSize(fontSize).text(componentsData[component].displayName, textX, 95);
    // Adding tables 
    const startX = 100;
    const startY = 120;
    const cellWidth = 130;
    const cellPadding = 5;

    // Set font properties
    doc.fontSize(9);

    // Calculate the maximum cell height
    const maxCellHeight = doc.currentLineHeight() + cellPadding * 3;
    // Draw the table headers

    doc.rect(startX, startY, cellWidth, maxCellHeight).fillAndStroke('lightgray', 'black');
    doc.fillColor('black').text('COMPONENTS', startX + cellPadding, startY + cellPadding);

    doc.rect(startX + cellWidth, startY, cellWidth, maxCellHeight).fillAndStroke('lightgray', 'black');
    doc.fillColor('black').text('INFORMATION PROVIDED', startX + cellWidth + cellPadding, startY + cellPadding);

    doc.rect(startX + 2 * cellWidth, startY, cellWidth, maxCellHeight).fillAndStroke('lightgray', 'black');
    doc.fillColor('black').text('INFORMATION VERIFIED', startX + 2 * cellWidth + cellPadding, startY + cellPadding);


    // Draw the table rows
    let currentY = startY + maxCellHeight;
    for (let i = 0; i < componentsData[component].labels.length; i++) {
      let lhsField;
      let rhsField;

      for (const item of componentsData[component].fieldsAndLabels) {
        if (item[componentsData[component].labels[i]] === componentsData[component].labels[i]) {
          lhsField = item.lhsField;
          rhsField = item.rhsField;
          break;

        }
      }


      let textHeight = doc.heightOfString(componentsData[component].labels[i], { width: cellWidth - cellPadding * 2 });
      const text2Height = doc.heightOfString(componentsData[component].fieldsValue[lhsField], { width: cellWidth - cellPadding * 2 });
      const text3Height = doc.heightOfString(componentsData[component].fieldsValue[rhsField], { width: cellWidth - cellPadding * 2 });

      textHeight = Math.max(Math.max(textHeight, text2Height), text3Height);

      // Check if the current cell height exceeds the maximum cell height
      const cellHeight = Math.max(textHeight + cellPadding * 2, maxCellHeight);

      doc.rect(startX, currentY, cellWidth, cellHeight).stroke();
      doc.fillColor('black').text(componentsData[component].labels[i], startX + cellPadding, currentY + cellPadding, { width: cellWidth - cellPadding * 2, height: cellHeight - cellPadding * 2, align: 'left' });


      doc.rect(startX + cellWidth, currentY, cellWidth, cellHeight).stroke();
      doc.fillColor('black').text(componentsData[component].fieldsValue[lhsField], startX + cellWidth + cellPadding, currentY + cellPadding, { width: cellWidth - cellPadding * 2, height: cellHeight - cellPadding * 2, align: 'left' });


      doc.rect(startX + 2 * cellWidth, currentY, cellWidth, cellHeight).stroke();
      doc.fillColor('black').text(componentsData[component].fieldsValue[rhsField], startX + 2 * cellWidth + cellPadding, currentY + cellPadding, { width: cellWidth - cellPadding * 2, height: cellHeight - cellPadding * 2, align: 'left' });

      currentY += cellHeight;

    }

    if (jpgFiles.length > 0) {
      for (let jpgFile of jpgFiles) {
        doc.addPage()
        doc.image(logoPath, centerX, logoStartY, { width: logoWidth, height: logoHeight });
        doc.image(jpgFile, {
          fit: [470, 550], // Adjust the width and height as needed
          align: 'center',
          valign: 'center',
        });

        doc.font('Helvetica').fillColor('Black').text('CONFIDENTIAL', centerX - 140, logoStartY + 730, { width: logoWidth, height: logoHeight });
        doc.font('Helvetica').fontSize(10).text("For Verifacts Services Pvt. Ltd.", centerX + 180, logoStartY + 640, { width: logoWidth, height: logoHeight });
        doc.font('Helvetica').fontSize(10).text("Client Account Manager", centerX + 180, logoStartY + 730, { width: logoWidth, height: logoHeight });
        doc.font('Helvetica').fontSize(10).text(`Page ${pageNum}`, centerX + 50, logoStartY + 730, { width: logoWidth, height: logoHeight });
        pageNum++;
        doc.image(stampPath, stampX + 160, stampStartY, { width: stampWidth, height: stampHeight })
      }

      // pageNum++;

      index++;
    }
  }

  


  doc.addPage();
  doc.image(logoPath, centerX, logoStartY, { width: logoWidth, height: logoHeight });

  //adding "confidential" text
  doc.font('Helvetica').fillColor('Black').text('CONFIDENTIAL', centerX - 140, logoStartY + 730, { width: logoWidth, height: logoHeight });


  // adding stamp on every component page
  doc.image(stampPath, stampX + 160, stampStartY, { width: stampWidth, height: stampHeight })
  doc.font('Helvetica').fontSize(10).text("For Verifacts Services Pvt. Ltd.", centerX + 180, logoStartY + 640, { width: logoWidth, height: logoHeight });
  doc.font('Helvetica').fontSize(10).text("Client Account Manager", centerX + 180, logoStartY + 730, { width: logoWidth, height: logoHeight });


  doc.font('Helvetica-Bold').fontSize(12).text('Restriction and Limitations', centerX - 165, logoStartY + 20 + logoHeight + 10, { align: 'center', underline: true });

  const conclusion = "Our reports and comments are confidential in nature and are meant only for the internal use of the client to make an assessment of the background of the applicant. They are not intended for publication or circulation or sharing with any other person including the applicant. Also, they are not to be reproduced or used for any other purpose, in whole or in part, without our prior written consent in each specific instance.\n\nWe request you to recognize that we are not the source of the data gathered and our findings are based on the information made available to us; therefore, we cannot guarantee the accuracy of the information collected. Should additional information or documentation become available to us, which impacts the conclusions reached in our reports, we reserve the right to amend our findings in our report accordingly.\n\nWe expressly disclaim all responsibility or liability for any costs, damages, losses, liabilities, expenses incurred by anyone as a result of circulation, publication, reproduction or use of our reports contrary to the provisions of this paragraph. You will appreciate that due to factors beyond our control, it may be possible that we are unable to get all the necessary information. Because of the limitations mentioned above, the results of our work with respect to the background checks should be considered only as a guide. Our reports and comments should not be considered as a definitive pronouncement on the individual.";

  // Split the text into paragraphs
  const paragraphs = conclusion.split('\n\n');

  // Adjust the spacing between paragraphs if needed
  const paragraphSpacing = 10;

  doc.font('Helvetica').fontSize(10);

  let startY = logoStartY + 50 + logoHeight + 10;

  paragraphs.forEach((paragraph) => {
    doc.text(paragraph, centerX - 170, startY, {
      align: 'justify'
    });
    startY += doc.heightOfString(paragraph, {
      width: 400, // Adjust the width as per your requirements
      align: 'justify'
    }) + paragraphSpacing;
  });

  doc.font('Helvetica-Bold').fontSize(12).text('--END OF REPORT--', centerX - 165, logoStartY + 280 + logoHeight + 10, { align: 'center', underline: true });

  // added page no.
  doc.font('Helvetica').fontSize(10).text(`Page ${pageNum}`, centerX + 50, logoStartY + 730, { width: logoWidth, height: logoHeight });

  // doc.addPage();
  // doc.drawImage(pdfImage, pdfImageOptions);

  try {

    // Finalize the PDF and create a writable stream
    const stream = doc.pipe(fs.createWriteStream('report.pdf'));
    // Set table position and cell dimensions

    doc.end();

    stream.on('finish', () => {
      console.log('PDF created successfully');
      res.download('report.pdf', 'report.pdf', (err) => {
        if (err) {
          console.log('Error downloading PDF:', err);
          res.sendStatus(500); // Send a response to indicate an error occurred during PDF download
        } else {
          // Delete the generated PDF file after download
          fs.unlinkSync('report.pdf');
        }
      });
    });

    stream.on('error', (err) => {
      console.log('Error creating PDF:', err);
      res.sendStatus(500); // Send a response to indicate an error occurred during PDF generation
    });
  } catch (error) {
    console.log('Error generating PDF:', error);
    res.sendStatus(500); // Send a response to indicate an error occurred during PDF generation
  }


}
const loadImage = (path) => {
  return new Promise((resolve, reject) => {
    fs.access(path, fs.constants.F_OK, (error) => {
      if (error) {
        reject(error); // Reject if the file doesn't exist or is inaccessible
      } else {
        resolve(); // Resolve if the file exists and is accessible
      }
    });
  });
};
