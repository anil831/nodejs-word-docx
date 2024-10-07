const Case = require('../../models/uploads/case.model');
const PersonalDetails = require('../../models/data_entry/personal_details_data.model');
const ClientContractProfiles = require("../../models/administration/client_contract_profile.model");
const ClientContractPackage=require("../../models/administration/client_contract_package.model");
const ColorMaster = require("../../models/administration/color_master.model");

const EducationBasic = require("../../models/data_entry/techmeducation.model");

const EmploymentBasic = require("../../models/data_entry/techmemployment.model");

// Address start
const Address = require("../../models/data_entry/address.model");
const TechmAddress = require("../../models/data_entry/techmaddress.model");
// Address end

const Pan = require("../../models/data_entry/pan.model");

const CreditCheck = require("../../models/data_entry/creditcheck.model");
const TechMCreditCheck = require("../../models/data_entry/techmcreditcheck.model");

const DrugTestFive = require("../../models/data_entry/techmdrugfive.model");

const DrugTestSeven = require("../../models/data_entry/drugtestseven.model");

const DrugTestTen = require("../../models/data_entry/techmdrugten.model");

const Gapvfns = require("../../models/data_entry/techmgapcheck.model");

const Reference = require("../../models/data_entry/reference.model");
const CourtRecord = require("../../models/data_entry/courtrecord.model");
const Passport = require("../../models/data_entry/techmpassport.model");
const Identities = require("../../models/data_entry/identity.model");

const { Worker } = require('worker_threads')
const gdb = require("../../models/data_entry/techmgdb.model");
const AadhaaarVerification=require("../../models/data_entry/aadhaarverification.model");	 

//added by anil on 2/22/2024

const CVAnalysis = require("../../models/data_entry/cvanalysis.model");

exports.backgroundVerificationReport =async (req,res) => {
   try{
    const caseId=req.params.caseId
    let getCaseDetails = function(){
        return new Promise((resolve,reject)=>{	 
           Case
           .findOne({caseId:caseId})
           .lean()	    
           .populate({path:"subclient",populate:{path:"client"}})
           .then(async data=>{
            if(data.grade){
               const gardeName = await getGradeDetails(data.grade)
               data.gardeName = gardeName

            }
            
              
              resolve(data)     	 
           })
           .catch(err=>{
            console.log(err);
              reject()	    
          })
       }) 	    
     }
     
     const getGradeDetails= async function(grade){

     return new Promise((resolve,reject)=>{
      ColorMaster.findById(grade,{"name":1,"_id":0}).lean()
      .then(data => resolve(data?.name))
      .catch((err)=>{
         console.log(err);
         reject()
      } )
     })
      
        
     }

     let getPersonalDetailsFromDb = function(caseObjectId){
        return new Promise((resolve,reject)=>{
         PersonalDetails
          .findOne({case:caseObjectId})
          .lean()	    
          .then(data=>{
         resolve(data)     
          })
          .catch(err=>{
            console.log(err);
         reject()     
          })	    
        })	    
     }

    let getDatabaseVerificationDetails = function(caseObjectId){
      return new Promise((resolve,reject)=>{
         gdb.find({case:caseObjectId}).lean().populate('case').then(async data =>{
            for(let item of data){
             const gardeName = await getGradeDetails(item.grade)
             item.status = gardeName           }
            resolve(data)
         }).catch(err => {
          console.log(err);
          reject()
         })
       })
    }

     let getClientContractProfileName= function(profile){
      return new Promise((resolve,reject)=>{
            ClientContractProfiles.findById(profile,{"_id":0,"name":1}).lean().then(data =>{
                     resolve(data?.name)
            }).catch(err =>{
               console.log(err);
               reject()
            })
      })
     }

     let getClientContractPackageName= function(package){
      return new Promise((resolve,reject)=>{
            ClientContractPackage.findById(package,{"_id":0,"name":1}).lean().then(data =>{
                     resolve(data?.name)
            }).catch(err =>{
               console.log(err);
               reject()
            })
      })
     }

     
     const getEducationalVerificationDetails = async function(caseObjectId){
      console.log("caseObjectId....",caseObjectId);
      return new Promise((resolve,reject)=>{
        EducationBasic.find({case:caseObjectId}).lean().populate('case').then(async data =>{
         console.log("edu data",data);
           for(let item of data){
            const gardeName = await getGradeDetails(item.grade)
            item.status = gardeName           }
           resolve(data)
        }).catch(err => {
         console.log(err);
         reject()
        })
      })
     }

     const getEmploymentVerificationDetails = async function(caseObjectId){
      return new Promise((resolve,reject)=>{
         EmploymentBasic.find({case:caseObjectId})
         .lean().populate('case')
         .then(async data =>{
            for(let item of data){
               const gardeName = await getGradeDetails(item.grade)
               item.status = gardeName  
            }
              resolve(data) 
            }).catch(err =>{
               console.log("err109",err);
               reject()
              } )
       })
     }



     // Address  start

     const getAddressVerificationDetails = async function(caseObjectId){
          return new Promise((resolve,reject)=>{
            Address.find({"case" : caseObjectId,typeofaddress:{$in:["Present"," Permanent"," Present&Permanent"]}})
            .lean().populate('case')
            .then(async data => {
               for(let item of data){
                  const gardeName = await getGradeDetails(item.grade)
                  item.status = gardeName  
               }
              resolve(data);
            }).catch(err => {
               console.log(err);
               reject();
            }) 
         })
     }

     const getTechMAddressVerificationDetails = async function(caseObjectId){
      return new Promise((resolve,reject)=>{
         TechmAddress.find({"case" : caseObjectId,typeofaddress:{$in:["Present"," Permanent"," Present&Permanent"]}})
        .lean().populate('case')
        .then(async data => {
           for(let item of data){
              const gardeName = await getGradeDetails(item.grade)
              item.status = gardeName  
           }
          resolve(data);
        }).catch(err => {
           console.log(err);
           reject();
        }) 
     })
 }
     // Adress end



     const getPanDetails = async function(caseObjectId){
      return new Promise ((resolve ,reject)=> {
         Pan.findOne({case : caseObjectId}).lean().populate('case').then(async data => {
            if(data){
               const gardeName = await getGradeDetails(data?.grade)
               data.status = gardeName  
            }
         
            resolve(data);
         }).catch(err => {
            console.log(err);
            reject()
         })
      })
     }


     // creditcheck start
     const getCreditCheckDetails = async function(caseObjectId){
      return new Promise ((resolve ,reject)=> {
         CreditCheck.find({case : caseObjectId}).lean().populate('case').then(async data => {
            for(let item of data){
               const gardeName = await getGradeDetails(item.grade)
               item.status = gardeName  
            }
              resolve(data) 
         
            resolve(data);
         }).catch(err => {
            console.log(err);
            reject()
         })
      })
     }

     const getTechMCreditCheckDetails = async function(caseObjectId){
      return new Promise ((resolve ,reject)=> {
         TechMCreditCheck.find({case : caseObjectId}).lean().populate('case').then(async data => {
            for(let item of data){
               const gardeName = await getGradeDetails(item.grade)
               item.status = gardeName  
            }
              resolve(data) 
         
            resolve(data);
         }).catch(err => {
            console.log(err);
            reject()
         })
      })
     }
     

     // creditcheck end
     
     const getDrugTestFiveDetails = async function(caseObjectId){
      return new Promise ((resolve ,reject)=> {
         DrugTestFive.findOne({case : caseObjectId}).lean().populate('case').then(async data => {
            if(data){
               const gardeName = await getGradeDetails(data?.grade)
               data.status = gardeName  
            }
         
            resolve(data);
         }).catch(err => {
            console.log(err);
            reject()
         })
      })
     }

     const getDrugTestSevenDetails = async function(caseObjectId){
      return new Promise ((resolve ,reject)=> {
         DrugTestSeven.findOne({case : caseObjectId}).lean().then(async data => {
            if(data){
               const gardeName = await getGradeDetails(data?.grade)
               data.status = gardeName  
            }
         
            resolve(data);
         }).catch(err => {
            console.log(err);
            reject()
         })
      })
     }

     const getDrugTestTenDetails = async function(caseObjectId){
      return new Promise ((resolve ,reject)=> {
         DrugTestTen.findOne({case : caseObjectId}).lean().then(async data => {
            if(data){
               const gardeName = await getGradeDetails(data?.grade)
               data.status = gardeName  
            }
         
            resolve(data);
         }).catch(err => {
            console.log(err);
            reject()
         })
      })
     }

     const gapVfns = async function(caseObjectId){
      return new Promise ((resolve ,reject)=> {
         Gapvfns.find({case : caseObjectId}).lean().populate('case').then(async data => {
            for(let item of data){
               const gardeName = await getGradeDetails(item.grade)
               item.status = gardeName  
            }    
            resolve(data);
         }).catch(err => {
            console.log(err);
            reject()
         })
      })
     }

     const getReferenceDetails = async function(caseObjectId){
      return new Promise ((resolve ,reject)=> {
         Reference.find({case : caseObjectId}).lean().populate('case').then(async data => {
            for(let item of data){
               const gardeName = await getGradeDetails(item.grade)
               item.status = gardeName  
            }    
            resolve(data);
         }).catch(err => {
            console.log(err);
            reject()
         })
      })
     }

     const getCourtRecordDetails = async function(caseObjectId){
      return new Promise ((resolve ,reject)=> {
         CourtRecord.find({case : caseObjectId}).lean()
         .populate({path:"personalDetailsData"}).populate('case')
         .then(async data => {
            for(let item of data){
               const gardeName = await getGradeDetails(item.grade)
               item.status = gardeName  
            }    
            resolve(data);
         }).catch(err => {
            console.log(err);
            reject()
         })
      })
     }

     const getPassportDetails = async function(caseObjectId){
      return new Promise ((resolve ,reject)=> {
         Passport.findOne({case : caseObjectId}).lean().populate('case').then(async data => {
            if(data){
               const gardeName = await getGradeDetails(data?.grade)
               data.status = gardeName  
            }
         
            resolve(data);
         }).catch(err => {
            console.log(err);
            reject()
         })
      })
     }

     const getDrivingLicenseDetails = async function(caseObjectId){
      return new Promise ((resolve ,reject)=> {
         Identities.findOne({case : caseObjectId,typeofid:{$in:[/DL/,/Driving/]}}).lean().populate('case').then(async data => {
            if(data){
               const gardeName = await getGradeDetails(data?.grade)
               data.status = gardeName  
            }
            resolve(data);
         }).catch(err => {
            console.log(err);
            reject()
         })
      })
     }


     const getAadhaarVerificationDetails = async function(caseObjectId){
      return new Promise ((resolve ,reject)=> {
         AadhaaarVerification.findOne({case : caseObjectId}).lean().populate('case').then(async data => {
            if(data){
               const gardeName = await getGradeDetails(data?.grade)
               data.status = gardeName  
            }
            resolve(data);
         }).catch(err => {
            console.log(err);
            reject()
         })
      })
     }


     const getCVAnalysisVerificationDetails = async function(caseObjectId){
      return new Promise ((resolve ,reject)=> {
         CVAnalysis.findOne({case : caseObjectId}).lean().populate('case').then(async data => {
            if(data){
               const gardeName = await getGradeDetails(data?.grade)
               data.status = gardeName  
            }
            resolve(data);
         }).catch(err => {
            console.log(err);
            reject()
         })
      })
     }
     let writeWordDocument = async function(caseDetails,personalDetails,profileOrPackageName,educationalVerificationDetails,
      employmentVerificationDetails,addressVerificationDetails,techMaddressVerificationDetails,panDetails,creditCheckDetails,techMcreditCheckDetails,drugTestFive,drugTestSeven,drugTestTen,
      gapVfnsDetails,referenceDetails,courtRecordDetails,passportDetails,drivingLicenseDetails,DatabaseVerificationDetails,aadhaarDeatils,cvanalysisDetails){
         return new Promise((resolve,reject)=>{
            let worker = new Worker("./controllers/reports/background_verification_report.js",{workerData:{caseDetails,personalDetails,profileOrPackageName,
               educationalVerificationDetails,employmentVerificationDetails,addressVerificationDetails,techMaddressVerificationDetails,panDetails,creditCheckDetails,techMcreditCheckDetails,
               drugTestFive,drugTestSeven,drugTestTen,gapVfnsDetails,referenceDetails,courtRecordDetails,passportDetails,drivingLicenseDetails,DatabaseVerificationDetails,aadhaarDeatils,cvanalysisDetails}}) 
            worker.on("message",resolve)

         })
     }


        let caseDetails = await getCaseDetails()
        const personalDetails = await getPersonalDetailsFromDb(caseDetails._id)
        const educationalVerificationDetails = await getEducationalVerificationDetails(caseDetails._id);
        const employmentVerificationDetails = await getEmploymentVerificationDetails(caseDetails._id);
        const DatabaseVerificationDetails = await getDatabaseVerificationDetails(caseDetails._id);

        

        //Address start

        const addressVerificationDetails = await getAddressVerificationDetails(caseDetails._id);
        const techMaddressVerificationDetails = await getTechMAddressVerificationDetails(caseDetails._id);

        // Address end

        const panDetails = await getPanDetails(caseDetails._id);

        // creditcheck start
        const creditCheckDetails = await getCreditCheckDetails(caseDetails._id);
        const techMcreditCheckDetails = await getTechMCreditCheckDetails(caseDetails._id);
       
        // creditcheck end

        

        const drugTestFive = await getDrugTestFiveDetails(caseDetails._id);
        const drugTestSeven = await getDrugTestSevenDetails(caseDetails._id);
        const drugTestTen = await getDrugTestTenDetails(caseDetails._id);
        const gapVfnsDetails = await gapVfns(caseDetails._id);
        const referenceDetails = await getReferenceDetails(caseDetails._id);
        const courtRecordDetails = await getCourtRecordDetails(caseDetails._id);
        const passportDetails  = await getPassportDetails(caseDetails._id);
        const drivingLicenseDetails = await getDrivingLicenseDetails(caseDetails._id);

        let profileOrPackageName="";
        if(caseDetails.profile){
         profileOrPackageName= await getClientContractProfileName(caseDetails.profile);
        }else if(caseDetails.package){
         profileOrPackageName= await getClientContractPackageName(caseDetails.package);
        }
        const aadhaarDeatils = await getAadhaarVerificationDetails(caseDetails._id);

        //Added by anil on 2/22/2024

        const cvanalysisDetails = await getCVAnalysisVerificationDetails(caseDetails._id);

        await writeWordDocument(caseDetails,personalDetails,profileOrPackageName,educationalVerificationDetails,employmentVerificationDetails,
         addressVerificationDetails,techMaddressVerificationDetails,panDetails,creditCheckDetails,techMcreditCheckDetails,drugTestFive,drugTestSeven,drugTestTen,gapVfnsDetails,referenceDetails,
         courtRecordDetails,passportDetails,drivingLicenseDetails,DatabaseVerificationDetails,aadhaarDeatils,cvanalysisDetails)

      return  res.download(`/cvws_new_uploads/backgroundVerification/${caseDetails.caseId}.docx`,(err)=>{
         if(err){
           console.log("Error Downloading ",err)	
                res.status(500).send({
               message:"Could not download  the file"	  
           })      		
         }	
          })
    }catch(error){
console.log("error",error);
return res.status(500).send(error)
    }
}



