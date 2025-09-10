const { google } = require('googleapis');
const nodemailer = require('nodemailer');
const QRCode = require('qrcode');
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
const TICKETS_FILE = 'sent_tickets.json';

// Initialize tickets tracking file
function initializeTicketsFile() {
  if (!fs.existsSync(TICKETS_FILE)) {
    fs.writeFileSync(TICKETS_FILE, JSON.stringify([]));
  }
}

// Get last processed row
function getLastRow() {
  if (fs.existsSync(TRACK_FILE)) {
    const content = fs.readFileSync(TRACK_FILE, 'utf8');
    const parsed = parseInt(content);
    return isNaN(parsed) ? 1 : parsed;
  }
  return 1;
}

// Save last processed row
function saveLastRow(row) {
  try {
    fs.writeFileSync(TRACK_FILE, row.toString(), 'utf8');
    console.log(`Last row updated to ${row}`);
  } catch (error) {
    console.error('Error saving last row:', error);
  }
}

// Save ticket information to file
function saveTicketInfo(ticketData) {
  try {
    let tickets = [];
    if (fs.existsSync(TICKETS_FILE)) {
      const data = fs.readFileSync(TICKETS_FILE, 'utf8');
      tickets = JSON.parse(data);
    }
    
    tickets.push({
      ...ticketData,
      timestamp: new Date().toISOString()
    });
    
    fs.writeFileSync(TICKETS_FILE, JSON.stringify(tickets, null, 2));
    console.log(`Ticket info saved: ${ticketData.ticketId}`);
  } catch (error) {
    console.error('Error saving ticket info:', error);
  }
}

// Generate ticket ID based on session
function generateTicketId(session) {
  const prefix = session === 1 ? 'T7S1' : 'T7S2';
  const randomNum = Math.floor(10000 + Math.random() * 90000);
  return `${prefix}${randomNum}`;
}

// Generate QR code
async function generateQRCode(ticketId, recipientEmail, recipientName, session) {
  const qrData = {
    ticketId: ticketId,
    event: 'TEDxIITPatna - Kaleidoscopic Interludes',
    date: '14th SEP, 2025',
    venue: 'AUDITORIUM',
    name: recipientName,
    email: recipientEmail,
    session: `Session ${session}`,
    generatedAt: new Date().toISOString()
  };
  
  const qrString = JSON.stringify(qrData);
  
  try {
    const qrCodeDataUrl = await QRCode.toDataURL(qrString, {
      width: 200,
      margin: 2,
      color: {
        dark: '#000000',
        light: '#FFFFFF'
      }
    });
    return qrCodeDataUrl;
  } catch (error) {
    console.error('Error generating QR code:', error);
    throw error;
  }
}

// Create HTML for ticket
function createTicketHTML(ticketId, recipientName, recipientEmail, qrCodeDataUrl, session) {
  return `
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <title>TEDxIITPatna Ticket</title>
  <style>
    body { font-family: Arial, sans-serif; background: #f5f5f5; padding: 20px; }
    .ticket-container { max-width: 600px; margin: auto; background: white; border-radius: 10px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); overflow: hidden;}
    .header { background: #8B0000; color: white; text-align: center; padding: 20px; }
    .content { padding: 20px; }
    .qr-code { text-align: center; margin-top: 20px; }
    .qr-code img { width: 150px; }
    .footer { background: #f0f0f0; text-align: center; padding: 10px; font-size: 12px; }
  </style>
</head>
<body>
  <div class="ticket-container">
    <div class="header">
      <h1>TEDxIITPatna</h1>
      <h3>Session ${session}</h3>
    </div>
    <div class="content">
      <p><strong>Ticket ID:</strong> ${ticketId}</p>
      <p><strong>Name:</strong> ${recipientName}</p>
      <p><strong>Email:</strong> ${recipientEmail}</p>
      <p><strong>Date:</strong> 14th SEP, 2025</p>
      <p><strong>Venue:</strong> AUDITORIUM</p>
      <div class="qr-code">
        <img src="${qrCodeDataUrl}" alt="QR Code"/>
        <p>Scan at entrance</p>
      </div>
    </div>
    <div class="footer">
      TEDxIITPatna | x = independently organized TED event
    </div>
  </div>
</body>
</html>
  `;
}

// Send ticket email
async function sendTEDxTicket(name, email, session) {
  try {
    const ticketId = generateTicketId(session);
    const qrCodeDataUrl = await generateQRCode(ticketId, email, name, session);
    const htmlContent = createTicketHTML(ticketId, name, email, qrCodeDataUrl, session);
    
    let transporter = nodemailer.createTransport({
      service: 'outlook',
      auth: {
        user: process.env.EMAIL_USER,
        pass: process.env.EMAIL_PASS,
      },
    });

    const mailOptions = {
      from: {
        name: 'TEDxIITPatna',
        address: process.env.EMAIL_USER
      },
      to: email,
      subject: `🎫 Your TEDxIITPatna Ticket - ${ticketId}`,
      html: htmlContent,
      attachments: [
        {
          filename: `TEDxIITPatna-Ticket-${ticketId}.html`,
          content: htmlContent,
          contentType: 'text/html'
        }
      ]
    };

    const info = await transporter.sendMail(mailOptions);

    const ticketData = {
      ticketId: ticketId,
      name: name,
      email: email,
      session: `Session ${session}`,
      messageId: info.messageId,
      status: 'sent'
    };

    saveTicketInfo(ticketData);

    console.log(`✅ Ticket sent: ${email}, Session ${session}, ID ${ticketId}`);

    return { success: true, ticketId: ticketId };
  } catch (error) {
    console.error(`❌ Failed to send ticket to ${email} (Session ${session}):`, error);

    const ticketData = {
      ticketId: null,
      name: name,
      email: email,
      session: `Session ${session}`,
      status: 'failed',
      error: error.message
    };

    saveTicketInfo(ticketData);

    return { success: false, error: error.message };
  }
}

// Check Google Sheet for new registrations
// Check Google Sheet for new registrations
async function checkSheet() {
  try {
    const lastRow = getLastRow();

    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}`,
    });

    const rows = res.data.values;
    if (!rows || rows.length <= lastRow) {
      return;
    }

    console.log(`Processing ${rows.length - lastRow} new registration(s)...`);

    for (let i = lastRow; i < rows.length; i++) {
      const row = rows[i];
      const name = row[0];
      const email = row[1];
      const sessions = row[2];

      if (!name || !email || !sessions) {
        console.log(`Skipping row ${i + 1}: Missing data`);
        continue;
      }

      console.log(`Processing: ${name} (${email}) Sessions: ${sessions}`);

      const sessionList = sessions.split(',').map(s => s.trim());
      let allSessionsProcessed = true;

      // Inside the checkSheet() function, replace the session loop with:
      for (const session of sessionList) {
        const sessionNum = session === 'Session 1' ? 1 : session === 'Session 2' ? 2 : null;
        if (!sessionNum) {
          console.log(`Unknown session "${session}" for ${email}`);
          continue;
        }
        await sendTEDxTicket(name, email, sessionNum);
        await new Promise(resolve => setTimeout(resolve, 1000));
      }

      // Update last row AFTER processing all sessions for this row
      saveLastRow(i + 1);

      // Only update last row if ALL sessions were processed successfully
      if (allSessionsProcessed) {
        saveLastRow(i + 1);
      } else {
        console.log(`Not updating last row for ${email} due to failed session processing`);
        break; // Stop processing further rows to avoid skipping failed ones
      }
    }

    console.log('Finished processing new registrations.');
  } catch (error) {
    console.error('Error checking sheet:', error);
  }
}

// Display ticket statistics
function getTicketStats() {
  try {
    if (!fs.existsSync(TICKETS_FILE)) {
      return { total: 0, sent: 0, failed: 0 };
    }
    
    const data = fs.readFileSync(TICKETS_FILE, 'utf8');
    const tickets = JSON.parse(data);
    
    return {
      total: tickets.length,
      sent: tickets.filter(t => t.status === 'sent').length,
      failed: tickets.filter(t => t.status === 'failed').length
    };
  } catch (error) {
    console.error('Error reading stats:', error);
    return { total: 0, sent: 0, failed: 0 };
  }
}

function displayStats() {
  const stats = getTicketStats();
  console.log('\n📊 Ticket Statistics');
  console.log('Total:', stats.total);
  console.log('Sent:', stats.sent);
  console.log('Failed:', stats.failed);
  console.log('Success Rate:', stats.total > 0 ? `${Math.round((stats.sent / stats.total) * 100)}%` : 'N/A');
}

// Initialize the system
function initialize() {
  initializeTicketsFile();
  console.log('🎫 TEDxIITPatna Ticket System Started');
  displayStats();
}

initialize();
checkSheet();
setInterval(checkSheet, 10000);
setInterval(displayStats, 300000);

console.log('🚀 Server is running and monitoring for new registrations...');
