const multer     = require('multer');
const XLSX       = require('xlsx');
const async = require('async');
const excelModel = require("../model/excelData");

const postXl = async (req,res)=>{
  // console.log("req.file::--" ,req.file);
  var workbook =  XLSX.readFile(req.file.path);
  var worksheet = workbook.Sheets[workbook.SheetNames[0]];
  var jsonData = XLSX.utils.sheet_to_json(worksheet);
  // console.log(jsonData);

  // Use async.eachSeries to process one candidate at a time, after reading from excel
  async.eachSeries(jsonData, async(data, callback)=>{
    try {
      // console.log(data)
      
      // Duplicate check has to be done for each row against the database using a mongo query
      const repEmail = await excelModel.findOne({email: data['Email']})
      if(!repEmail){
        const item = new excelModel({
          name: data['Name of the Candidate'],
          email: data['Email'],
          mobile_no: data['Mobile No.'],
          dob: data['Date of Birth'],
          woe: data['Work Experience'] ,
          re: data['Resume Title'] ,
          cl: data['Current Location'] ,
          pa: data['Postal Address'] ,
          ce: data['Current Employer'] ,
          cd: data['Current Designation'] ,
  
        });
        await item.save();
     }
    } catch (error) {
    console.log(error);
    res.status(500).json("UnSuccessFull");
    }
  })
  res.status(201).json("SuccessFull")
 
}

module.exports = { postXl};

