const express = require("express");
const cors = require("cors");
const bodyParser = require("body-parser");
const nodemailer = require("nodemailer");
const replace = require('replace-in-file');
var JSZip = require('jszip');
var Docxtemplater = require('docxtemplater');
var fs = require('fs');
var path = require('path');
var gutil = require('gulp-util');

const details = require("./details.json");

const app = express();
app.use(cors({ origin: "*" }));
app.use(bodyParser.json());

app.listen(3000, () => {
  console.log("The server started on port 3000 !!!!!!");
});


app.get("/", (req, res) => {
  res.send(
    "<h1 style='text-align: center'>Wellcome in Email Service</h1>"
  );
  sendMail(user, info => {
    console.log(`The mail has beed send 😃 and the id is ${info.messageId}`);
    console.log(user)
    res.send(info);
  });
  // fileRemove('/undefinedinvoiceTemplate.docx');

});

app.post("/sendmail", (req, res) => {
  let user = req.body;
  // console.log(user);
  getDoc({
    "senderName": user.invoice.sender.name,
    "senderCompany": user.invoice.sender.company,
    "senderCompanyAdress": user.invoice.sender.companyAdress,
    "senderCity": user.invoice.sender.city,
    "senderCountry": user.invoice.sender.country,
    "senderPhone": user.invoice.sender.phone,
    "recipientName": user.invoice.recipient.name,
    "recipientCompany": user.invoice.recipient.company,
    "recipientCompanyAdress": user.invoice.recipient.companyAdress,
    "recipientCity": user.invoice.recipient.city,
    "recipientCountry": user.invoice.recipient.country,
    "invoiceID": user.invoice.id,
    "created": user.invoice.created,
    "deadline": user.invoice.deadline,
    "description": user.invoice.description,
    "quantity": user.invoice.quantity,
    "salary": user.invoice.salary,
    "sum": user.totalGross,
    "taxT": user.taxSalary,
    "taxR": user.invoice.tax + '%',
    "Total": user.salaryNet
  },'invoiceTemplate.docx');
  
  console.log(user)
  sendMail(user, info => {
    console.log(`The mail has beed send 😃 and the id is ${info.messageId}`);
    // res.send(info);
  });
});

async function sendMail(user, callback) {
  // create reusable transporter object using the default SMTP transport
  console.log('Sending...')
  let transporter = nodemailer.createTransport({
    service: 'gmail',
    host: "smtp.gmail.com",
    port: 587,
    secure: false, // true for 465, false for other ports
    auth: {
      user: 'chanj3396@gmail.com',
      pass: 'hqcyfnwhhfmcumxw'
    },
    tls: {
      rejectUnauthorized: false
    }
  });

  let mailOptions = {
    from: '"Angular App Service"<example.gimail.com>', // sender address
    to: user.email, // list of receivers
    subject: "Invoice Attachment", // Subject line
    html: `<h1>Hi kurwa </h1><br>
    <h4>Thanks for joining us</h4>`,
    attachments: [{   // file on disk as an attachment
      filename: 'undefinedinvoiceTemplate.docx',
      path: './undefinedinvoiceTemplate.docx' // stream this file
    }]
  };

  // send mail with defined transport object
  let info = await transporter.sendMail(mailOptions);

  callback(info);
}

 function getDoc (data, file) {
  //  Cargo el docx como un binary 
  var content = fs.readFileSync(path.resolve(file), 'binary');
  var zip = new JSZip(content);
  var doc = new Docxtemplater();
  doc.loadZip(zip);
  // setea los valores de data ej: { first_name: 'John' , last_name: 'Doe'}
  doc.setData(data);
  try {
    // renderiza el documento (remplaza las ocurrencias como {first_name} by John, {last_name} by Doe, ...) 
    doc.render();
  } catch (error) {
    var e = {
      message: error.message,
      name: error.name,
      stack: error.stack,
      properties: error.properties,
    }
    console.log(JSON.stringify({ error: e })); throw error;
  } var buffer = doc.getZip().generate({ type: 'nodebuffer' });
   fs.writeFileSync(path.resolve('', data.doc + file), buffer);
    return path.resolve(__dirname.replace('', data.doc + file));
}

function fileRemove(path){
  fs.unlink(path, function (err) {            
    if (err) {                                                 
        console.error(err);  
        console.log('File not exist');                              
    }                                                          
   console.log('File has been Deleted');                           
  });  
}



