const xlsx= require("xlsx");
let wb =xlsx.readFile("data.xlsx");
let ws = wb.Sheets["Sheet1"];
let data= xlsx.utils.sheet_to_json(ws);


for (let i=0; i<data.length; i++)
{
   
  if(data[i].data1=="Required!!!!!")
   {
    const nodemailer = require("nodemailer");
    let transporter = nodemailer.createTransport({
        service:"gmail",
        auth: {
          user:"yourname1@gmail.com",//  generated ethereal user
          pass:"yourPass" , // generated ethereal password
        },
      });
      let mailoption = {
        from: '"your_name" <yourname2@gmail.com>', // sender address
        to: "youremail3@gmail.com",  //list of receivers 
        subject: "Hi world", // Subject line
        text: "Hello world?", // plain text body
        html: '<h1 style ="color :blue"> Hello World  </h1> <p>Look the file attached</p>', // html body
        attachments: [ 
          { 
            filename: 'Name_Of_File.pdf', 
            path: __dirname + '/Name_Of_File.pdf', 
            cid: 'uniq-Name_Of_File.pdf'
               }
        ] // File attached
          };
    transporter.sendMail(mailoption,function(err,info){
      if(err){
          console.log("error")
      }
      else{
          console.log("email sent ")
      }
    })
   }
   
   if(data[i].data2=="Required!!!!!")
    {
     const nodemailer = require("nodemailer");
     let transporter = nodemailer.createTransport({
         service:"gmail",
         auth: {
           user:"yourname1@gmail.com",//  generated ethereal user
           pass:"yourPass" , // generated ethereal password
         },
       });
     let mailoption = {
       from: '"your_name" <yourname2@gmail.com>', // sender address
       to: "youremail4@gmail.com",  //list of receivers 
       subject: "Hi world", // Subject line
       text: "Hello world?", // plain text body
       html: '<h1 style ="color :blue"> Hello World  </h1> <p>Look the file attached</p>', // html body
       attachments: [ 
        { 
          filename: 'Name_Of_File.pdf', 
          path: __dirname + '/Name_Of_File.pdf', 
          cid: 'uniq-Name_Of_File.pdf'
             }
       ] // File attached
     
         };
         transporter.sendMail(mailoption,function(err,info){
           if(err){
               console.log("error")
           }
           else{
               console.log("email sent ")
           }
         })
        }
}




