const  mongoose  = require("mongoose");

const excelSchema = new mongoose.Schema({
    name:String,
    email:String,
    mobile_no:Number, 
    dob:Date, 
    woe:String , 
    re:String ,
    cl:String , 
    pa:String , 
    ce:String ,
    cd:String

});

module.exports = mongoose.model('excelData',excelSchema);


