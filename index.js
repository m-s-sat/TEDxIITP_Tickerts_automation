const { google } = require('googleapis');
const nodemailer = require('nodemailer');
const fs = require('fs');
const path = require('path');
require('dotenv').config();

// Load service account credentials
const KEYFILEPATH = path.join(__dirname, 'credentials.json');
const SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly'];
const auth = new google.auth.GoogleAuth({
  keyFile: KEYFILEPATH,
  scopes: SCOPES,
});

const sheets = google.sheets({ version: 'v4', auth });

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const SHEET_NAME = 'Sheet1';

const TRACK_FILE = 'last_row.txt';

function getLastRow() {
  if (fs.existsSync(TRACK_FILE)) {
    return parseInt(fs.readFileSync(TRACK_FILE, 'utf8'));
  }
  return 1;
}

function saveLastRow(row) {
  fs.writeFileSync(TRACK_FILE, row.toString());
}


async function sendEmail(name, email) {
  let transporter = nodemailer.createTransport({
    service: 'gmail',
    auth: {
      user: process.env.EMAIL_USER,
      pass: process.env.EMAIL_PASS,
    },
  });

  const mailOptions = {
    from: process.env.EMAIL_USER,
    to: email,
    subject: 'Your Event Ticket',
    text: `Hello ${name},\n\nThank you for registering. Here is your ticket.\n\nRegards.`,
  };

  await transporter.sendMail(mailOptions);
  console.log(`Email sent to ${email}`);
}
async function checkSheet() {
  const lastRow = getLastRow();

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${SHEET_NAME}`,
  });

  const rows = res.data.values;
  if (!rows || rows.length <= lastRow) {
    console.log('No new rows.');
    return;
  }

  for (let i = lastRow; i < rows.length; i++) {
    const row = rows[i];
    const name = row[0];
    const email = row[1];
    try {
      await sendEmail(name, email);
    } catch (err) {
      console.error(`Failed to send email to ${email}:`, err);
    }
    saveLastRow(i + 1);
  }
}

// Run periodically
setInterval(checkSheet, 10000); // check every 10 seconds

console.log('Server is running...');
checkSheet();
