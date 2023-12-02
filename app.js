const XLSX = require("xlsx");
const dotenv = require("dotenv");
const nodemailer = require("nodemailer");

dotenv.config();

const workbook = XLSX.readFile("xxx.xlsx"); // Replace with your file name

let emailAddresses = [];

for (const sheetName of workbook.SheetNames) {
  const worksheet = workbook.Sheets[sheetName];
  const sheetData = XLSX.utils.sheet_to_json(worksheet);

  sheetData.forEach((student) => {
    if (student.Email) {
      emailAddresses.push(student.Email);
    }
  });
}

const transporter = nodemailer.createTransport({
  service: "Brevo", // Replace with your email service (e.g., Gmail, Outlook)
  auth: {
    user: process.env.BREVO_EMAIL,
    pass: process.env.BREVO_PASS,
  },
});

cnt = 0;

emailAddresses.forEach((email) => {
  console.log(`${++cnt}  ${email}`);
  const mailOptions = {
    from: "noreply@acu-ecpc.com", // Replace with your email
    to: email,
    subject: "",
    text: "",
  };

  transporter.sendMail(mailOptions, (error, info) => {
    if (error) {
      console.error("Error sending email:", error);
    } else {
      console.log("Email sent:", info.response);
    }
  });
});
