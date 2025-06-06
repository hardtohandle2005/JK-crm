// 🚀 Basic Setup (OneDrive-Ready)
const express = require('express');
const axios = require('axios');
const qs = require('qs');
const path = require('path');
const ExcelJS = require('exceljs');
require('dotenv').config();
const pdfParse = require('pdf-parse');
const fs = require('fs');               // For fs.createReadStream etc.
const fsp = require('fs').promises;     // For await fsp.readFile(...)
const sharp = require('sharp'); // add at top of your file


const app = express();
const port = 3000;
const fileUpload = require('express-fileupload');
app.use(fileUpload());


const puppeteer = require('puppeteer');
app.use(express.json());

async function generatePdfFromHtml(htmlContent) {
  const browser = await puppeteer.launch({
    headless: 'new', // For newer versions of Puppeteer
    args: ['--no-sandbox', '--disable-setuid-sandbox']
  });

  const page = await browser.newPage();
  await page.setContent(htmlContent, { waitUntil: 'networkidle0' });
  const pdfBuffer = await page.pdf({ format: 'A4', printBackground: true });

  await browser.close();
  return pdfBuffer;
}



app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'login.html'));
});

app.use(express.json());
app.use(express.urlencoded({ extended: true }));

const session = require('express-session');
const bcrypt = require('bcryptjs');

const otpStore = new Map(); // email => { otp, expiresAt }
const crypto = require('crypto');

global.tempAccessRequests = global.tempAccessRequests || {};



app.use(session({
  secret: 'jk-solar-secret',
  resave: false,
  saveUninitialized: false,
  cookie: {
    maxAge: 3600000, // 1 hour
    secure: false    // true only if using HTTPS
  }
}));

// 📂 OneDrive Setup – no local paths needed

// 🛠 Basic Route Setup
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});



// 📄 Excel file name for stock sheet (stored in OneDrive)
const stockSheetFileName = 'Stock Sheet.xlsx';

// ✅ Middlewares already set earlier — no need to duplicate



// ✅ Multer instance removed (OneDrive will be used instead)
// This block is deprecated and should be replaced with direct upload to OneDrive API

// Placeholder for OneDrive upload handling logic
// File naming logic will be reused in upload routes
const fieldNameToFileName = {
  aadharFront: 'aadharfront',
  aadharBack: 'aadharback',
  panCard: 'pancard',
  bill: 'bill',
  ownershipProof: 'ownershipproof',
  cancelCheque: 'cancelcheque',
  clientPhoto: 'clientphoto'
};


app.post('/submit-client', async (req, res) => {
  try {
    const data = req.body;
    const files = req.files;
    const clientName = `${data.name}-${data.kw}-${data.mobile}`.trim().toLowerCase().replace(/\s+/g, '_');

    if (data.photoData) {
      const base64Match = data.photoData.match(/^data:image\/(\w+);base64,(.+)$/);
      if (base64Match) {
        const extension = base64Match[1];
        const base64Data = base64Match[2];
        const buffer = Buffer.from(base64Data, 'base64');

        console.log('🧪 Base64 length:', base64Data.length);
        console.log('🧪 Detected extension:', extension);

        await uploadToOneDriveFolder(clientName, 'clientPhoto', buffer, `clientPhoto.${extension}`);
      } else {
        console.warn('⚠️ Invalid base64 format in photoData');
      }
    }

    console.log('📦 Sending files, please wait...');
    const uploadTasks = Object.entries(files).map(async ([field, fileObj]) => {
      const file = fileObj[0] || fileObj;
      const ext = path.extname(file.name);
      const customName = `${field.toLowerCase()}${ext}`;
      await uploadToOneDriveFolder(clientName, field, file.data, customName);

      const pretty = field.replace(/([A-Z])/g, ' $1').replace(/^./, s => s.toUpperCase());
      console.log(`📤 ${pretty} uploaded...`);
    });

        await Promise.all(uploadTasks);

    const client = {
      date: data.date,
      name: data.name,
      so: data.so,
      houseNo: data.houseNo,
      location: data.location,
      landmark: data.landmark,
      district: data.district,
      state: data.state,
      pin: data.pin,
      mobile: data.mobile,
      email: data.email,
      kw: data.kw,
      advance: data.advance,
      totalCost: data.totalCost,
      aadharFront: `uploads/${clientName}/aadharfront.png`,
      aadharBack: `uploads/${clientName}/aadharback.png`,
      panCard: `uploads/${clientName}/pancard.png`,
      bill: `uploads/${clientName}/bill.png`,
      ownershipProof: `uploads/${clientName}/ownershipproof.png`,
      cancelCheque: `uploads/${clientName}/cancelcheque.png`
    };

    const { workbook, token } = await getWorkbookFromOneDrive('TempData.xlsx');
    const sheet = workbook.getWorksheet('Client Data') || workbook.addWorksheet('Client Data');

    const row = [
      client.date, client.name, ` ${client.so || ''}`,client.houseNo, client.location,  client.landmark,
      client.district, client.state, client.pin, client.mobile, client.email, client.kw, client.advance, client.totalCost,
      client.aadharFront, client.aadharBack, client.panCard, client.bill,
      client.ownershipProof, client.cancelCheque
    ];



    sheet.addRow(row);
    await uploadWorkbookToOneDrive('TempData.xlsx', workbook, token);


    let tempAddress = '';
    let tempLocation = '';
    let tempCost = '';

try {
  const { workbook: tempBook } = await getWorkbookFromOneDrive('TempData.xlsx');
  const tempSheet = tempBook.getWorksheet('Client Data');

  tempSheet.eachRow((row, idx) => {
    const name = (row.getCell(2).value || '').toString().toLowerCase().trim();
    const mobile = (row.getCell(10).value || '').toString().trim();
    const kwValue = (row.getCell(12).value || '').toString().trim();

    if (
      name === (client.name || '').toLowerCase().trim() &&
      mobile === (client.mobile || '').trim() &&
      kwValue === (client.kw || '').toString().trim()
    ) {
      const hno      = row.getCell(4)?.value?.toString().trim() || '';
const location = row.getCell(5)?.value?.toString().trim() || '';
const landmark = row.getCell(6)?.value?.toString().trim() || '';
const district = row.getCell(7)?.value?.toString().trim() || '';
const state    = row.getCell(8)?.value?.toString().trim() || '';
const pin      = row.getCell(9)?.value?.toString().trim() || '';

tempAddress = `${hno}, ${location}, ${landmark}, ${district}, ${state} - ${pin}`;
tempLocation = `${district}, ${state}`;
tempCost = row.getCell(14)?.value?.toString().trim() || '';

    }
  });
} catch (err) {
  console.warn('⚠️ Could not fetch address from TempData:', err.message);
}

    let proposalData = null;
    try {
      const { workbook: proposalBook } = await getWorkbookFromOneDrive('proposal.xlsx');
      const proposalSheet = proposalBook.getWorksheet('Proposals');

      const clientKW = (client.kw || '').toString().trim();
      const clientMobile = (client.mobile || '').toString().trim();

      console.log(`🧪 Searching in proposal.xlsx for: KW="${clientKW}", Mobile="${clientMobile}"`);

      proposalSheet.eachRow((row, index) => {
  if (index === 1) return; // skip header

  const ref = (row.getCell(1).value || '').toString().trim(); // Ref: e.g., JK-05-9876543210
  const match = ref.match(/^JK-(\d+)-(\d{10})$/);  // Safely extract KW and Mobile

  if (!match) return;

  const refKW = match[1];            // '05'
  const refMobile = match[2];        // '9876543210'
  const cleanClientMobile = (client.mobile || '').replace(/\D/g, '').slice(-10);

  console.log(`🧪 Matching: RefKW(${refKW}) == ClientKW(${clientKW}), RefMob(${refMobile}) == ClientMob(${cleanClientMobile})`);

  if (refKW === clientKW && refMobile === cleanClientMobile) {
    console.log(`✅ Match found: Row ${index}`);

          proposalData = {
            name: row.getCell(8).value,
            location: tempLocation,
            address: tempAddress,
            mobile: row.getCell(9).value || '',
            price: row.getCell(10).value || '',
            email: client.email,
            kw: row.getCell(4).value || '',
            panelBrand: row.getCell(11).value || '',
            panelWp: row.getCell(12).value || '',
            inverterBrand: row.getCell(13).value || '',
            ref: row.getCell(1).value || '', 
            date: row.getCell(2).value || client.date
          };

          proposalData.panelQty = proposalData.kw && proposalData.panelWp
            ? Math.round((parseFloat(proposalData.kw) * 1000) / parseFloat(proposalData.panelWp))
            : '';
        }
      });

      if (!proposalData) {
        console.warn('🚫 No matching row found in proposal.xlsx');
      }

    } catch (err) {
      console.warn('⚠️ Proposal data not fetched:', err.message);
    }

    

    let so = data.so || '';

    try {
      const { workbook: tempBook } = await getWorkbookFromOneDrive('TempData.xlsx');
      const tempSheet = tempBook.getWorksheet('Client Data');

      tempSheet.eachRow((row, idx) => {
        const name = (row.getCell(2).value || '').toString().toLowerCase().trim();
        const mobile = (row.getCell(10).value || '').toString().trim();

        if (
          name === (client.name || '').toLowerCase().trim() &&
          mobile === (client.mobile || '').trim()
        ) {
          so = row.getCell(3)?.value?.toString() || '';
        }
      });
    } catch (e) {
      console.warn('⚠️ Could not fetch S/o from TempData.xlsx:', e.message);
    }

    let netSuccess = false;
    let poSuccess = false;

    if (proposalData) {
      try {
        const htmlPath = path.join(__dirname, 'public', 'net-metering-agreement.html');
        const htmlContent = await fs.promises.readFile(htmlPath, 'utf8');

        const today = new Date();
        const day = today.getDate();
        const month = today.toLocaleString('default', { month: 'long' });
        const year = today.getFullYear();

        const populatedHTML = htmlContent
          .replace(/{{name}}/g, `<u>${client.name || ''}</u>`)
          .replace(/{{location}}/g, `<u>${proposalData.location || ''}</u>`)
          .replace(/{{so}}/g, `<u>${so || ''}</u>`)
          .replace(/{{address}}/g, `<u>${proposalData.address || ''}</u>`)
          .replace(/{{mobile}}/g, `<u>${proposalData.mobile || ''}</u>`)
          .replace(/{{email}}/g, `<u>${proposalData.email || ''}</u>`)
          .replace(/{{kw}}/g, `<u>${proposalData.kw || ''}</u>`)
          .replace(/{{price}}/g, `<u>${tempCost || ''}</u>`)
          .replace(/{{date}}/g, `<u>${proposalData.date || ''}</u>`)
          .replace(/{{panelBrand}}/g, `<u>${proposalData.panelBrand || ''}</u>`)
          .replace(/{{panelWp}}/g, `<u>${proposalData.panelWp || ''}</u>`)
          .replace(/{{inverterBrand}}/g, `<u>${proposalData.inverterBrand || ''}</u>`)
          .replace(/{{day}}/g, `<u>${day}</u>`)
          .replace(/{{month}}/g, `<u>${month}</u>`)
          .replace(/{{year}}/g, `<u>${year}</u>`);

        const pdfBuffer = await generatePdfFromHtml(populatedHTML);
        await uploadToOneDriveFolder(clientName, 'net-metering', pdfBuffer, 'net-metering-agreement.pdf');
        console.log('📄 Net Metering Agreement created');
        netSuccess = true;
      } catch (err) {
        console.warn('⚠️ Net Metering PDF generation failed:', err.message);
      }

      try {
        const htmlPath = path.join(__dirname, 'public', 'purchase-order.html');
        const htmlContent = await fs.promises.readFile(htmlPath, 'utf8');

        const today = new Date();
        const day = today.getDate();
        const month = today.toLocaleString('default', { month: 'long' });
        const year = today.getFullYear();
        const todayDate = `${day}${getDateSuffix(day)} ${month}, ${year}`;

        const populatedHTML = htmlContent
          .replace(/{{name}}/g, client.name)
          .replace(/{{address}}/g, proposalData.address || '')
          .replace(/{{mobile}}/g, client.mobile)
          .replace(/{{kw}}/g, proposalData.kw || '')
          .replace(/{{ref}}/g, proposalData.ref || '')
          .replace(/{{price}}/g, tempCost || '')
          .replace(/{{today}}/g, todayDate)
          .replace(/{{proposalDate}}/g, proposalData.date || '')
          .replace(/{{panelBrand}}/g, proposalData.panelBrand || '')
          .replace(/{{panelWp}}/g, proposalData.panelWp || '')
          .replace(/{{inverterBrand}}/g, proposalData.inverterBrand || '')
          .replace(/{{panelQty}}/g, proposalData.panelQty || '');

        const pdfBuffer = await generatePdfFromHtml(populatedHTML);
        await uploadToOneDriveFolder(clientName, 'purchase-order', pdfBuffer, 'purchase-order.pdf');
        console.log('📃 Purchase Order PDF created');
        poSuccess = true;
      } catch (err) {
        console.warn('⚠️ Purchase Order PDF generation failed:', err.message);
      }
    }

    res.send(`
      <html>
        <head>
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <style>
            body {
              background: #f0f8ff;
              font-family: 'Segoe UI', sans-serif;
              display: flex;
              justify-content: center;
              align-items: center;
              height: 90vh;
            }
            .toast {
              background: #d7f8d7;
              padding: 25px 30px;
              border-radius: 12px;
              box-shadow: 0 4px 15px rgba(0,0,0,0.1);
              color: #206020;
              font-size: 18px;
              line-height: 1.6;
            }
            .toast span {
              display: block;
              margin-top: 10px;
            }
            .toast .error {
              color: #c0392b;
              font-weight: bold;
            }
          </style>
        </head>
        <body>
          <div class="toast">
            ✅ Client submitted successfully.
            ${netSuccess ? '<span>📄 Net Metering PDF also generated!</span>' : '<span class="error">📄 Net Metering not created. Data mismatched!</span>'}
            ${poSuccess ? '<span>📃 Purchase Order PDF also generated!</span>' : '<span class="error">📃 Purchase Order not created. Data mismatched!</span>'}
          </div>
        </body>
      </html>
    `);

    function getDateSuffix(day) {
      if (day >= 11 && day <= 13) return 'th';
      const lastDigit = day % 10;
      return ['st', 'nd', 'rd'][lastDigit - 1] || 'th';
    }
  } catch (err) {
    console.error('❌ Error in /submit-client:', err.message);
    res.status(500).send('⚠️ Could not save data to OneDrive.');
  }
});








// ✅ Search route to check if client folder exists in OneDrive
app.get('/search-client', async (req, res) => {
  const name = req.query.name?.trim();
  if (!name) {
    return res.status(400).json({ error: 'No name provided' });
  }

  const clientName = name.toLowerCase().replace(/\s+/g, '_');

  try {
    const token = await getAccessToken();
    const url = `https://graph.microsoft.com/v1.0/users/muninderpal@jk17.onmicrosoft.com/drive/root:/uploads/${clientName}`;
    
    await axios.get(url, {
      headers: { Authorization: `Bearer ${token}` }
    });

    res.json({ found: true });
  } catch (err) {
    if (err.response?.status === 404) {
      res.json({ found: false });
    } else {
      console.error('❌ Error checking OneDrive folder:', err.message);
      res.status(500).json({ error: 'Internal Server Error' });
    }
  }
});

app.get('/api/client-suggestions', async (req, res) => {
  const query = (req.query.query || '').toLowerCase().trim();
  if (!query) return res.json([]);

  try {
    const { workbook } = await getWorkbookFromOneDrive('TempData.xlsx');
    const sheet = workbook.getWorksheet('Client Data');
    const results = [];

    sheet.eachRow((row, i) => {
      if (i === 1) return; // skip header

      const name = (row.getCell(2).value || '').toString();
      const kw = (row.getCell(12).value || '').toString();
      const mobile = (row.getCell(10).value || '').toString();

      if (
        name.toLowerCase().includes(query) ||
        mobile.includes(query) ||
        kw.includes(query)
      ) {
        results.push({ name, kw, mobile });
      }
    });

    res.json(results.slice(0, 10)); // limit to top 10
  } catch (err) {
    console.error("❌ Suggestion error:", err.message);
    res.json([]);
  }
});

// ✅ Route to handle adding timeline event for a client via OneDrive
app.post('/add-timeline', express.json(), async (req, res) => {
  const { clientName, event, eventDate, eventDescription, status } = req.body;

  try {
    const { workbook, token } = await getWorkbookFromOneDrive('TempData.xlsx');
    let worksheet = workbook.getWorksheet('Timeline');

    // Create Timeline sheet if it doesn't exist
    if (!worksheet) {
      worksheet = workbook.addWorksheet('Timeline');
      worksheet.addRow(['Client Name', 'Event', 'Event Date', 'Event Description', 'Status']);
    }

    // Add the new timeline event
    worksheet.addRow([clientName, event, eventDate, eventDescription, status]);

    // Save back to OneDrive
    await uploadWorkbookToOneDrive('TempData.xlsx', workbook, token);

    res.send('✅ Timeline event added successfully to OneDrive!');
  } catch (error) {
    console.error('❌ Error adding timeline event:', error.message);
    res.status(500).send('❌ Error adding timeline event to OneDrive.');
  }
});


const mime = require('mime-types');

async function uploadToOneDriveFolder(clientName, field, buffer, fileName) {
  const folderPath = `uploads/${clientName}`;
  const fullPath = `${folderPath}/${fileName}`;

  const token = await getAccessToken();
  const contentType = mime.lookup(fileName) || 'application/octet-stream';

  const uploadUrl = `https://graph.microsoft.com/v1.0/users/muninderpal@jk17.onmicrosoft.com/drive/root:/${fullPath}:/content`;

  await axios.put(uploadUrl, buffer, {
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': contentType
    }
  });

  console.log(`📁 Uploaded ${fileName} to ${fullPath}`);
}



app.post('/submit-documents', async (req, res) => {
  const clientName = req.body.name?.trim().toLowerCase().replace(/\s+/g, '_');
  const files = req.files;

  if (!files || Object.keys(files).length === 0) {
    return res.status(400).send('No files were uploaded.');
  }

 try {
  for (let field in files) {
    const file = files[field][0] || files[field];
    const ext = path.extname(file.name);
    const customName = `${fieldNameToFileName[field] || field}${ext}`;
    await uploadToOneDriveFolder(clientName, field, file.data, customName);
  }

    console.log('📤 Uploaded files:', Object.keys(files));
    res.send('✅ Documents uploaded successfully to OneDrive!');
  } catch (err) {
    console.error('❌ Upload error:', err.message);
    res.status(500).send('❌ Failed to upload documents to OneDrive.');
  }
});



app.get('/file-status/:clientName', async (req, res) => {
  const clientNameRaw = req.params.clientName.trim().toLowerCase();
  const fileFields = [
    'AadharFront',
    'AadharBack',
    'PanCard',
    'Bill',
    'OwnershipProof',
    'CancelCheque',
    'PurchaseAgreement',
    'NetMeteringAgreement'
  ];

  try {
    console.log(`📂 Checking files for: ${clientNameRaw}`);

    let clientInfo = null;
    let fileStatus = [];

    const { workbook } = await getWorkbookFromOneDrive('TempData.xlsx');
    const worksheet = workbook.getWorksheet('Client Data');

    worksheet.eachRow((row, rowIndex) => {
      const name = (row.getCell(2).value || '').toString().trim(); // B
      const kw = (row.getCell(12).value || '').toString().trim(); // L
      const mobile = (row.getCell(10).value || '').toString().trim(); // J

      const fullNameFormat = `${name}-${kw}-${mobile}`.toLowerCase();
      console.log(`🔍 Excel: '${fullNameFormat}' vs Search: '${clientNameRaw}'`);

      if (fullNameFormat === clientNameRaw) {
        clientInfo = {
          name: name,
          address: row.getCell(4).value || '',     // D (H.No.)
          mobile: mobile,
          email: row.getCell(11).value || '',      // K
          kw: kw
        };

        // Loop over columns O to V (15 to 22)
        fileFields.forEach((field, index) => {
          const pathInExcel = row.getCell(15 + index).value?.toString().trim() || '';
          const label = field.replace(/([A-Z])/g, ' $1').replace(/^./, s => s.toUpperCase());

          fileStatus.push({
            file: field,
            label: label,
            exists: !!pathInExcel
          });
        });
        return;
      }
    });

    if (clientInfo) {
      return res.json({
        files: fileStatus,
        clientInfo: clientInfo
      });
    } else {
      return res.status(404).json({ error: 'Client not found' });
    }

  } catch (err) {
    console.error('❌ Error in /file-status route:', err);
    return res.status(500).json({ error: 'Internal Server Error' });
  }
});





app.get('/check-files', async (req, res) => {
  try {
    const token = await getAccessToken();
    const clientName = 'demo'; // 🔁 Replace with dynamic if needed

    const fileList = [
      `uploads/${clientName}/aadharfront.png`,
      `uploads/${clientName}/aadharback.png`,
      `uploads/${clientName}/pancard.png`,
      `uploads/${clientName}/bill.png`,
      `uploads/${clientName}/ownershipproof.png`,
      `uploads/${clientName}/cancelcheque.png`,
      `uploads/${clientName}/net-metering-agreement.pdf`,
      `uploads/${clientName}/purchase-order.pdf`
    ];

    const results = [];

    for (const path of fileList) {
      const encodedPath = encodeURIComponent(path);
      const url = `https://graph.microsoft.com/v1.0/users/muninderpal@jk17.onmicrosoft.com/drive/root:/${encodedPath}`;

      try {
        await axios.get(url, {
          headers: { Authorization: `Bearer ${token}` }
        });
        results.push({ file: path, exists: true });
      } catch (err) {
        results.push({ file: path, exists: false });
      }
    }

    // ✅ Extract PDF statuses
    const netMeteringExists = results.find(f => f.file.includes('net-metering-agreement.pdf'))?.exists;
    const purchaseOrderExists = results.find(f => f.file.includes('purchase-order.pdf'))?.exists;

    // ✅ Update Excel (U = Purchase, V = Net Metering)
    const { workbook } = await getWorkbookFromOneDrive('TempData.xlsx');
    const sheet = workbook.getWorksheet('Client Data');

    sheet.eachRow((row, i) => {
      const name = (row.getCell(2).value || '').toString().toLowerCase().trim();
      const mobile = (row.getCell(10).value || '').toString().trim();
      const kw = (row.getCell(12).value || '').toString().trim();

      if (name === 'simran' && mobile === '9650179900' && kw === '10') { // 🔁 Use dynamic match in real app
        row.getCell('U').value = purchaseOrderExists ? 'Yes' : '';
        row.getCell('V').value = netMeteringExists ? 'Yes' : '';
      }
    });

    await saveWorkbookToOneDrive(workbook, 'TempData.xlsx');

    res.send(`<pre>${JSON.stringify(results, null, 2)}</pre>`);
  } catch (err) {
    console.error('❌ Error in /check-files:', err.message);
    res.status(500).send('Failed to check file statuses.');
  }
});



// 📅 View timeline of a client from OneDrive Excel
app.get('/view-timeline/:clientName', async (req, res) => {
  const rawClientName = req.params.clientName || '';
  const clientName = rawClientName.trim().toLowerCase();

  console.log(`🔍 Looking for timeline for client: "${clientName}"`);

  try {
    const { workbook } = await getWorkbookFromOneDrive('TempData.xlsx');
    const worksheet = workbook.getWorksheet('Client Data');

    if (!worksheet) {
      return res.status(404).json({ error: 'Client Data sheet not found' });
    }

    const clientTimeline = [];

    worksheet.eachRow((row, rowNumber) => {
      const rowValues = row.values;

      const name = (rowValues[2] || '').toString().trim().toLowerCase();   // Name - B
      const kw = (rowValues[12] || '').toString().trim();                  // KW - L
      const mobile = (rowValues[10] || '').toString().trim();             // Mobile - J
      const expectedName = `${name}-${kw}-${mobile}`.toLowerCase();

      if (expectedName === clientName) {
        clientTimeline.push({
          appliedKW: rowValues[23] || '',                 // W
          appliedPMSurya: rowValues[24] || '',            // X
          dhbvnApplication: rowValues[25] || '',          // Y
          loadNameChange: rowValues[26] || ''             // Z
        });
      }
    });

    if (clientTimeline.length === 0) {
      return res.status(404).json({ error: 'No timeline data found for this client' });
    }

    res.json(clientTimeline);
  } catch (error) {
    console.error('❌ Error reading Excel from OneDrive:', error.message);
    res.status(500).json({ error: 'Internal Server Error' });
  }
});



// 🧾 Get client info (basic details) from OneDrive
app.get('/client-info/:clientName', async (req, res) => {
  const clientName = req.params.clientName.trim().toLowerCase();

  try {
    const { workbook } = await getWorkbookFromOneDrive('TempData.xlsx');
    const sheet = workbook.getWorksheet('Client Data');

    if (!sheet) {
      return res.status(404).json({ error: 'Client Data sheet not found' });
    }

    let clientInfo = null;

    sheet.eachRow((row) => {
      const rowClientName = row.getCell(2).value?.toString().trim().toLowerCase(); // Column B - Name

      if (rowClientName === clientName) {
        clientInfo = {
          name: row.getCell(2).value || '',     // B - Name
          address: row.getCell(4).value || '',  // D - Location
          mobile: row.getCell(10).value || '',  // J - Mob. No.
          email: row.getCell(11).value || '',   // K - Email ID
          kw: row.getCell(12).value || ''       // L - KW
        };
      }
    });

    if (clientInfo) {
      res.json(clientInfo);
    } else {
      res.status(404).json({ error: 'Client not found' });
    }
  } catch (err) {
    console.error('❌ Error reading client info from OneDrive:', err.message);
    res.status(500).json({ error: 'Internal Server Error' });
  }
});





app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// Serve static files (like your HTML, CSS, etc.)
app.use(express.static('public'));
setHeaders: (res, path) => {
    if (path.endsWith('.html')) {
      res.set('Cache-Control', 'no-store');
    }
  }


// 📝 Route to save timeline data to OneDrive Excel
app.post('/save-timeline', async (req, res) => {
  const {
    appliedKW,
    appliedOnPMSurya,
    applicationDHBVN,
    loadChangeReductionNewConnection,
    clientName
  } = req.body;

  if (!clientName) {
    console.warn('⛔ Client name missing in request');
    return res.status(400).json({ error: 'Client name missing' });
  }

  try {
    const { workbook, token } = await getWorkbookFromOneDrive('TempData.xlsx');
    const sheet = workbook.getWorksheet('Client Data');

    if (!sheet) {
      console.error('📄 Sheet "Client Data" not found');
      return res.status(404).json({ error: 'Client Data sheet not found' });
    }

    let found = false;

    sheet.eachRow((row, rowIndex) => {
      if (rowIndex === 1) return; // Skip header

      const name = (row.getCell(2).value || '').toString().trim();     // B
      const mobile = (row.getCell(10).value || '').toString().trim();  // J
      const kw = (row.getCell(12).value || '').toString().trim();      // L

      const constructedName = `${name}-${kw}-${mobile}`.toLowerCase().replace(/\s+/g, '_');

      console.log(`🔍 Comparing: Excel="${constructedName}" vs Incoming="${clientName.toLowerCase().trim()}"`);

      if (constructedName === clientName.toLowerCase().trim()) {
        if (appliedKW) row.getCell(23).value = appliedKW;                             // W
if (appliedOnPMSurya) row.getCell(24).value = appliedOnPMSurya;               // X
if (applicationDHBVN) row.getCell(25).value = applicationDHBVN;               // Y
if (loadChangeReductionNewConnection) row.getCell(26).value = loadChangeReductionNewConnection; // Z

        found = true;
      }
    });

    if (!found) {
      console.warn(`🚫 Client not found for name: ${clientName}`);
      return res.status(404).json({ error: 'Client not found in Excel' });
    }

    await uploadWorkbookToOneDrive('TempData.xlsx', workbook, token);
    console.log('✅ Timeline saved successfully for:', clientName);
    res.json({ message: '✅ Timeline data saved successfully to OneDrive!' });

  } catch (err) {
    console.error('❌ Error saving timeline data:', err.message);
    res.status(500).json({ error: 'Internal server error' });
  }
});




app.get('/application-timeline/:clientName', async (req, res) => {
  const clientName = req.params.clientName.toLowerCase().trim();

  try {
    const { workbook } = await getWorkbookFromOneDrive('TempData.xlsx');
    const sheet = workbook.getWorksheet('Client Data');

    if (!sheet) {
      return res.status(404).json({ error: 'Client Data sheet not found' });
    }

    let result = null;

    sheet.eachRow((row, rowIndex) => {
      if (rowIndex === 1) return; // Skip header

      const name = (row.getCell(2).value || '').toString().trim();    // B
      const mobile = (row.getCell(10).value || '').toString().trim(); // J ✅
      const kw = (row.getCell(12).value || '').toString().trim();     // L ✅

      const constructedName = `${name}-${kw}-${mobile}`.toLowerCase().replace(/\s+/g, '_');

      if (constructedName === clientName) {
        result = {
          appliedKW: row.getCell(23)?.value || '',   // W
          appliedOnPMSurya: row.getCell(24)?.value || '', // X
          applicationDHBVN: row.getCell(25)?.value || '', // Y
          loadChangeReductionNewConnection: row.getCell(26)?.value || '' // Z
        };
      }
    });

    if (result) {
      res.json(result);
    } else {
      res.status(404).json({ error: 'Client not found' });
    }

  } catch (err) {
    console.error('❌ Error loading timeline data from OneDrive:', err.message);
    res.status(500).json({ error: 'Internal server error' });
  }
});


//upload application odf to folder
app.post('/upload-application-pdf', async (req, res) => {
  const clientName = req.body.clientName?.toLowerCase().trim();
  const file = req.files?.pdfFile;

  if (!clientName || !file) {
    return res.status(400).send('Missing client name or file');
  }

  try {
    const ext = path.extname(file.name) || '.pdf';
    const buffer = file.data;
    const fileName = `application${ext}`;

    await uploadToOneDriveFolder(clientName, 'application', buffer, fileName);
    res.send('✅ Application PDF uploaded successfully');
  } catch (err) {
    console.error('❌ Error uploading application PDF:', err.message);
    res.status(500).send('❌ Failed to upload Application PDF');
  }
});

app.get('/check-application-pdf/:clientName', async (req, res) => {
  const clientName = req.params.clientName?.toLowerCase().trim();
  const filePath = `uploads/${clientName}/application.pdf`;

  try {
    const token = await getAccessToken();
    const url = `https://graph.microsoft.com/v1.0/users/muninderpal@jk17.onmicrosoft.com/drive/root:/${filePath}`;

    await axios.get(url, {
      headers: { Authorization: `Bearer ${token}` }
    });

    res.json({ exists: true });
  } catch (err) {
    if (err.response?.status === 404) {
      res.json({ exists: false });
    } else {
      console.error('❌ Error checking Application PDF:', err.message);
      res.status(500).json({ error: 'Internal Server Error' });
    }
  }
});

app.get('/preview-pdf/:clientName', async (req, res) => {
  const clientName = req.params.clientName.toLowerCase().trim();
  const filePath = `uploads/${clientName}/application.pdf`;

  try {
    const token = await getAccessToken();
    const fileUrl = `https://graph.microsoft.com/v1.0/users/muninderpal@jk17.onmicrosoft.com/drive/root:/${filePath}:/content`;

    const response = await axios.get(fileUrl, {
      headers: { Authorization: `Bearer ${token}` },
      responseType: 'stream'
    });

    res.setHeader('Content-Type', 'application/pdf');
    response.data.pipe(res);
  } catch (err) {
    console.error('❌ Preview error:', err.message);
    res.status(404).send('⚠️ PDF not found');
  }
});


const projectFields = [
  "Civil", "Earthing", "EarthingCable", "Panel", "Inverter", "ACDB",
  "DCDB", "ACCable", "DCCable", "LA", "NetMetering"
];

app.post('/save-project-status', express.urlencoded({ extended: true }), async (req, res) => {
  const clientName = req.body.clientName?.toLowerCase().trim();
  if (!clientName) return res.status(400).json({ error: 'Client name missing' });

  try {
    const { workbook, token } = await getWorkbookFromOneDrive('TempData.xlsx');
    const sheet = workbook.getWorksheet('Client Data');
    if (!sheet) return res.status(404).json({ error: 'Client Data sheet not found' });

    let updated = false;

    sheet.eachRow((row, rowIndex) => {
      if (rowIndex === 1) return; // skip header

      const name = (row.getCell(2).value || '').toString().trim();       // B
      const mobile = (row.getCell(10).value || '').toString().trim();    // J
      const kw = (row.getCell(12).value || '').toString().trim();        // L 

      const constructedName = `${name}-${kw}-${mobile}`.toLowerCase().replace(/\s+/g, '_');

      if (constructedName === clientName) {
        // Columns U (21) to AE (31)
        projectFields.forEach((field, index) => {
        row.getCell(27 + index).value = req.body[field] || '';
        });


        updated = true;

        // ✅ Check if first 10 steps are all marked 'yes' or '✅'
        const completedSteps = projectFields.slice(0, 10).map((_, i) => row.getCell(27 + i).value);
        console.log('📋 Completed Steps:', completedSteps);

        const allDone = completedSteps.every(step =>
          ['yes', '✅'].includes((step || '').toString().trim().toLowerCase())
        );

        if (allDone) {
          console.log(`✅ All 10 project steps completed for ${clientName}. Triggering PDF generation...`);
          generateWorkCompletionReport(clientName); // Auto PDF generation
        }
      }
    });

    if (!updated) {
      return res.status(404).json({ error: 'Client not found in Excel' });
    }

    await uploadWorkbookToOneDrive('TempData.xlsx', workbook, token);
    res.json({ message: '✅ Project status saved successfully to OneDrive' });

  } catch (err) {
    console.error('❌ Error saving project status:', err.message);
    res.status(500).json({ error: 'Internal Server Error' });
  }
});


async function generateWorkCompletionReport(clientName) {
  try {
    console.log(`🔍 Starting Work Completion Report generation for: ${clientName}`);

    const proposalFile = 'proposal.xlsx';
    const htmlTemplatePath = path.join(__dirname, 'work-completion-report.html');
    const saveFileName = 'work-completion-report.pdf';
    const folderPath = `uploads/${clientName}`;

    // ✂️ Extract KW and Mobile from clientName
    const parts = clientName.split('-');
    const clientKW = parts[1]?.trim();
    const clientMobile = parts[2]?.trim();

    if (!clientKW || !clientMobile) {
      return console.log('❌ Cannot extract KW and Mobile from clientName');
    }

    // 🧾 Load proposal data
    console.log('📑 Reading proposal.xlsx...');
    const { workbook: proposalBook } = await getWorkbookFromOneDrive(proposalFile);
    const proposalSheet = proposalBook.getWorksheet('Proposals');
    if (!proposalSheet) return console.log('❌ Proposals sheet not found');

    console.log('📄 Matching proposal row...');
    let matchedRow;
    proposalSheet.eachRow((row) => {
      const ref = (row.getCell(1).value || '').toString(); // Ref: JK-10-9650179900
      if (ref.includes(clientKW) && ref.includes(clientMobile)) {
        matchedRow = row;
      }
    });

    if (!matchedRow) return console.log('❌ No matching proposal row found');

    // 🧠 Extract values
    const panelWp = matchedRow.getCell(13).value || '';       // Panel Wp (Col M)
    const inverterBrand = matchedRow.getCell(14).value || ''; // Inverter Brand (Col N)
    const kw = matchedRow.getCell(4).value || '';             // KW (Col D)
    const name = matchedRow.getCell(8).value || '';           // To Whom (Col H)
    const address = matchedRow.getCell(5).value || '';        // Address (Col E)

    // 📄 Read application.pdf and extract values
    console.log('📥 Reading application.pdf from OneDrive...');
    const appPdfBuffer = await downloadFileFromOneDrive(`${folderPath}/application.pdf`);
    const pdfData = await pdfParse(appPdfBuffer);
    const text = pdfData.text;
    console.log('📄 Extracted PDF Text:\n', text.substring(0, 1000));


    const dhbvn     = text.match(/Operation Sub-?Division.*?\n(.*?)G\d+-\d+/)?.[1]?.trim() || '';
const applicationNo = text.match(/(G\d{2}-\d{3,}-\d{4}-\d{3,})/)?.[1]?.trim() || '';
const applicationDate = text.match(/(\d{2}\/\d{2}\/\d{4})/)?.[1] || '';


    // 🧩 Fill into HTML
    console.log('🧩 Injecting data into HTML...');
    let htmlContent = await fs.readFile(htmlTemplatePath, 'utf8');
    htmlContent = htmlContent
      .replace(/{{panelWp}}/g, panelWp)
      .replace(/{{inverterBrand}}/g, inverterBrand)
      .replace(/{{kw}}/g, kw)
      .replace(/{{name}}/g, name)
      .replace(/{{address}}/g, address)
      .replace(/{{applicationNo}}/g, applicationNo)
      .replace(/{{applicationDate}}/g, applicationDate)
      .replace(/{{dhbvn}}/g, dhbvn);

      // 👇 Add here before PDF generation
console.log('📤 applicationNo:', applicationNo);
console.log('📤 applicationDate:', applicationDate);
console.log('📤 dhbvn:', dhbvn);
console.log('📄 Final HTML preview:\n', htmlContent.substring(0, 500));

    // 🖨️ Generate PDF
    console.log('🖨️ Generating PDF with Puppeteer...');
    const browser = await puppeteer.launch({ headless: 'new' });
    const page = await browser.newPage();
    await page.setContent(htmlContent, { waitUntil: 'networkidle0' });
    const pdfBuffer = await page.pdf({ format: 'A4', printBackground: true });
    await browser.close();

    
    // ☁️ Upload to OneDrive
    console.log('☁️ Uploading work-completion-report.pdf to OneDrive...');
    await uploadFileToOneDrive(`${folderPath}/${saveFileName}`, pdfBuffer);
    console.log(`✅ Work Completion Report saved for ${clientName}`);
  } catch (err) {
    console.error('❌ Error generating Work Completion Report:', err.message);
  }
}


const streamToBuffer = async (stream) => {
  const chunks = [];
  for await (const chunk of stream) {
    chunks.push(chunk);
  }
  return Buffer.concat(chunks);
};

async function downloadFileFromOneDrive(filePath) {
  const accessToken = await getAccessToken();
  const client = require('@microsoft/microsoft-graph-client').Client.init({
    authProvider: (done) => done(null, accessToken)
  });

  const responseStream = await client
    .api(`/users/muninderpal@jk17.onmicrosoft.com/drive/root:/${filePath}:/content`)
    .getStream();

  return await streamToBuffer(responseStream);
}


const { Client } = require('@microsoft/microsoft-graph-client');
require('isomorphic-fetch');






app.get('/project-status/:clientName', async (req, res) => {
  const clientName = req.params.clientName.toLowerCase().trim();

  try {
    const { workbook } = await getWorkbookFromOneDrive('TempData.xlsx');
    const sheet = workbook.getWorksheet('Client Data');
    if (!sheet) return res.status(404).json({ error: 'Client Data sheet not found' });

    let result = null;

    sheet.eachRow((row, rowIndex) => {
      if (rowIndex === 1) return; // Skip header

      const name = (row.getCell(2).value || '').toString().trim();     // B
      const mobile = (row.getCell(10).value || '').toString().trim(); // J
      const kw = (row.getCell(12).value || '').toString().trim();     // L


      const constructedName = `${name}-${kw}-${mobile}`.toLowerCase().replace(/\s+/g, '_');

      if (constructedName === clientName) {
          result = {
  Civil: row.getCell(27)?.value || '',
  Earthing: row.getCell(28)?.value || '',
  EarthingCable: row.getCell(29)?.value || '',
  Panel: row.getCell(30)?.value || '',
  Inverter: row.getCell(31)?.value || '',
  ACDB: row.getCell(32)?.value || '',
  DCDB: row.getCell(33)?.value || '',
  ACCable: row.getCell(34)?.value || '',
  DCCable: row.getCell(35)?.value || '',
  LA: row.getCell(36)?.value || '',
  NetMetering: row.getCell(37)?.value || ''
};

      }
    });

    if (!result) return res.status(404).json({ error: 'Client not found' });
    res.json({ status: result });

  } catch (err) {
    console.error('❌ Error loading project status from OneDrive:', err.message);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// 💸 Payment timeline autofetch route from OneDrive
app.get('/payment-status/:clientName', async (req, res) => {
  const clientName = req.params.clientName.toLowerCase().trim();

  try {
    const { workbook } = await getWorkbookFromOneDrive('TempData.xlsx');
    const sheet = workbook.getWorksheet('Client Data');
    if (!sheet) return res.status(404).json({ error: 'Client Data sheet not found' });

    let result = null;

    sheet.eachRow((row, rowIndex) => {
      if (rowIndex === 1) return; // Skip header

      const name = (row.getCell(2).value || '').toString().trim();     // B
      const mobile = (row.getCell(10).value || '').toString().trim();  // J
      const kw = (row.getCell(12).value || '').toString().trim();      // L

      const constructedName = `${name}-${kw}-${mobile}`.toLowerCase().replace(/\s+/g, '_');

      if (constructedName === clientName) {
        result = {
          totalCost: row.getCell(14)?.value || '',   // N
          advance: row.getCell(13)?.value || '',     // M
          projectStatus: {
            Civil: row.getCell(27)?.value || '',       // AA
            NetMetering: row.getCell(37)?.value || ''  // AK
          },
          saved: {
            installment2: row.getCell(38)?.value || '',   // AL
            installment3: row.getCell(39)?.value || '',   // AM
            finalPayment: row.getCell(40)?.value || '',   // AN
            balance: row.getCell(41)?.value || ''         // AO
          }
        };
      }
    });

    if (!result) return res.status(404).json({ error: 'Client not found' });
    res.json(result);
  } catch (err) {
    console.error('❌ Error reading payment data from OneDrive:', err.message);
    res.status(500).json({ error: 'Internal Server Error' });
  }
});




// 💾 Save payment status to OneDrive Excel
app.post('/save-payment-status', express.urlencoded({ extended: true }), async (req, res) => {
  const clientName = req.body.clientName?.toLowerCase().trim();
  if (!clientName) return res.status(400).json({ error: 'Client name missing' });

  try {
    const { workbook, token } = await getWorkbookFromOneDrive('TempData.xlsx');
    const sheet = workbook.getWorksheet('Client Data');
    if (!sheet) return res.status(404).json({ error: 'Client Data sheet not found' });

    let updated = false;

    sheet.eachRow((row, rowIndex) => {
      if (rowIndex === 1) return; // Skip header

      const name = (row.getCell(2).value || '').toString().trim();     // B
      const mobile = (row.getCell(10).value || '').toString().trim();  // J
      const kw = (row.getCell(12).value || '').toString().trim();      // L

      const constructedName = `${name}-${kw}-${mobile}`.toLowerCase().replace(/\s+/g, '_');

      if (constructedName === clientName) {
        row.getCell(38).value = req.body.installment2 || '';   // AL
        row.getCell(39).value = req.body.installment3 || '';   // AM
        row.getCell(40).value = req.body.finalPayment || '';   // AN
        row.getCell(41).value = req.body.balance || '';        // AO
        updated = true;
      }
    });

    if (!updated) return res.status(404).json({ error: 'Client not found' });

    await uploadWorkbookToOneDrive('TempData.xlsx', workbook, token);
    res.json({ message: '✅ Payment status saved successfully to OneDrive' });
  } catch (err) {
    console.error('❌ Error saving payment status:', err.message);
    res.status(500).json({ error: 'Internal Server Error' });
  }
});




// 🧾 Load workbook from OneDrive
async function loadWorkbookFromOneDrive(fileName) {
  const { workbook, token } = await getWorkbookFromOneDrive(fileName);
  return { workbook, token };
}

// 💾 Save workbook back to OneDrive
async function saveWorkbookToOneDrive(fileName, workbook, token) {
  await uploadWorkbookToOneDrive(fileName, workbook, token);
}

// ✅ Find or create row for material in stock sheet
function findOrCreateMaterialRow(sheet, material) {
  for (let i = 3; i <= sheet.rowCount; i++) {
    const cell = sheet.getCell(`A${i}`);
    if (cell.value && cell.value.toString().toLowerCase() === material.toLowerCase()) {
      return sheet.getRow(i);
    }
  }
  const newRow = sheet.addRow([material]);
  return newRow;
}



// ✅ Stock In (OneDrive-powered)
app.post('/submit-stock-in', async (req, res) => {
  const { date, material, invoice, quantity } = req.body;
  if (!date || !material || !invoice || !quantity) {
    return res.status(400).send('Missing required fields');
  }

  try {
    const fileName = 'Stock Sheet.xlsx';
    const { workbook, token } = await loadWorkbookFromOneDrive(fileName);

    const monthNumber = ('0' + (new Date(date).getMonth() + 1)).slice(-2);
    const stockSheetName = `Stock ${monthNumber}`;
    const stockInSheetName = `Stock In ${monthNumber}`;

    const sheetIn = getOrCreateSheet(workbook, stockInSheetName);
    const stockSheet = getOrCreateSheet(workbook, stockSheetName);

    // Save to stock in sheet
    sheetIn.addRow([date, material, invoice, quantity]);

    // Ensure date columns exist
    createMissingDates(stockSheet, date);

    // Retry finding date columns
    let dateCols = findDateColumns(stockSheet, date);

    if (!dateCols) {
      await saveWorkbookToOneDrive(fileName, workbook, token);

      const retry = await loadWorkbookFromOneDrive(fileName);
      const retrySheet = retry.workbook.getWorksheet(stockSheetName);
      createMissingDates(retrySheet, date);
      dateCols = findDateColumns(retrySheet, date);
      if (!dateCols) {
        console.error('❌ Still cannot find columns after re-creating');
        return res.status(500).send('❌ Date columns missing.');
      }
    }

    const materialRow = findOrCreateMaterialRow(stockSheet, material);
    const { inCol } = dateCols;
    const existingQty = parseFloat(materialRow.getCell(inCol).value) || 0;
    materialRow.getCell(inCol).value = existingQty + parseFloat(quantity);

    await updateCurrentStock(workbook, stockSheet, material);
    await saveWorkbookToOneDrive(fileName, workbook, token);

    res.send('✅ Stock In recorded & Stock Sheet updated in OneDrive');
  } catch (err) {
    console.error('❌ Error during Stock In:', err.message);
    res.status(500).send('❌ Error writing to OneDrive Excel');
  }
});




// ✅ Stock Out (OneDrive-powered)
app.post('/submit-stock-out', async (req, res) => {
  const { date, material, quantity, remarks } = req.body;
  if (!date || !material || !quantity) {
    return res.status(400).send('Missing required fields');
  }

  try {
    const fileName = 'Stock Sheet.xlsx';
    const { workbook, token } = await loadWorkbookFromOneDrive(fileName);

    const monthNumber = ('0' + (new Date(date).getMonth() + 1)).slice(-2);
    const stockSheetName = `Stock ${monthNumber}`;
    const stockOutSheetName = `Stock Out ${monthNumber}`;

    const sheetOut = getOrCreateSheet(workbook, stockOutSheetName);
    const stockSheet = getOrCreateSheet(workbook, stockSheetName);

    // Save to stock out sheet
    sheetOut.addRow([date, material, quantity, remarks || '']);

    // Ensure date columns exist
    createMissingDates(stockSheet, date);

    // Retry finding date columns
    let dateCols = findDateColumns(stockSheet, date);

    if (!dateCols) {
      await saveWorkbookToOneDrive(fileName, workbook, token);

      const retry = await loadWorkbookFromOneDrive(fileName);
      const retrySheet = retry.workbook.getWorksheet(stockSheetName);
      createMissingDates(retrySheet, date);
      dateCols = findDateColumns(retrySheet, date);
      if (!dateCols) {
        console.error('❌ Still cannot find columns after re-creating');
        return res.status(500).send('❌ Date columns missing.');
      }
    }

    const materialRow = findOrCreateMaterialRow(stockSheet, material);
    const { outCol } = dateCols;

    const existingOut = parseFloat(materialRow.getCell(outCol).value) || 0;
    materialRow.getCell(outCol).value = existingOut + parseFloat(quantity);

    await updateCurrentStock(workbook, stockSheet, material);
    await saveWorkbookToOneDrive(fileName, workbook, token);

    res.send('✅ Stock Out recorded & Stock Sheet updated in OneDrive');
  } catch (err) {
    console.error('❌ Error during Stock Out:', err.message);
    res.status(500).send('❌ Error writing to OneDrive Excel');
  }
});




// ✅ Fix getOrCreateSheet
function getOrCreateSheet(workbook, sheetName) {
  let sheet = workbook.getWorksheet(sheetName);
  if (!sheet) {
    if (sheetName.includes('Stock In')) {
      sheet = workbook.addWorksheet(sheetName);
      sheet.addRow(['Date', 'Material', 'Invoice No.', 'Quantity']);
    } else if (sheetName.includes('Stock Out')) {
      sheet = workbook.addWorksheet(sheetName);
      sheet.addRow(['Date', 'Material', 'Quantity', 'Remarks']);
    } else if (sheetName.includes('Stock')) {
      sheet = workbook.addWorksheet(sheetName);
      setupStockSheet(sheet, new Date().getFullYear(), sheetName.split(' ')[1]); // ✨ correctly setup Stock sheet
    }
  }
  return sheet;
}





// ✅ Update Current Stock in Stock Sheet (Column C)
async function updateCurrentStock(workbook, stockSheet, material) {
  const materialRow = findOrCreateMaterialRow(stockSheet, material);

  const openingStock = parseFloat(materialRow.getCell(2).value) || 0; // Column B

  let totalIn = 0;
  let totalOut = 0;

  for (let col = 5; col <= stockSheet.columnCount; col += 3) {
    const inVal = parseFloat(materialRow.getCell(col).value) || 0;
    const outVal = parseFloat(materialRow.getCell(col + 1).value) || 0;

    totalIn += inVal;
    totalOut += outVal;
  }

  const currentStock = openingStock + totalIn - totalOut;
  materialRow.getCell(3).value = currentStock; // Column C (Current Stock)

  updateMinStockAndHighlight(stockSheet, materialRow);
  // ❌ No need to saveWorkbook here — save happens in route handler after this call
}

// ✅ Update Min Stock (10% of Opening Stock) and highlight red if needed
function updateMinStockAndHighlight(sheet, row) {
  const openingStock = parseFloat(row.getCell(2).value) || 0;  // Column B
  const closingStock = parseFloat(row.getCell(3).value) || 0;  // Column C

  const minStock = +(openingStock * 0.10).toFixed(2);
  row.getCell(4).value = minStock; // Column D

  const redFill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFF0000' }
  };

  const clearFill = {
    type: 'pattern',
    pattern: 'none'
  };

  if (closingStock <= minStock) {
    row.getCell(1).fill = redFill;  // Material
    row.getCell(4).fill = redFill;  // Min Stock
  } else {
    row.getCell(1).fill = clearFill;
    row.getCell(4).fill = clearFill;
  }
}

// ✅ Setup Stock Sheet (Initial columns and per-day date blocks)
function setupStockSheet(sheet, year, month) {
  sheet.getCell('A1').value = 'Material';
  sheet.getCell('B1').value = 'Opening Stock';
  sheet.getCell('C1').value = 'Current Stock';
  sheet.getCell('D1').value = 'Min Stock';

  // Row 2 left blank for A–D
  sheet.getCell('A2').value = '';
  sheet.getCell('B2').value = '';
  sheet.getCell('C2').value = '';
  sheet.getCell('D2').value = '';

  let daysInMonth = monthDays[parseInt(month)];
  if (parseInt(month) === 2 && isLeapYear(year)) {
    daysInMonth = 29;
  }

  let startCol = 5;

  for (let day = 1; day <= daysInMonth; day++) {
    const dateStr = `${day.toString().padStart(2, '0')}-${month}-${year}`;

    sheet.mergeCells(1, startCol, 1, startCol + 2);
    sheet.getCell(1, startCol).value = dateStr;
    sheet.getCell(1, startCol).alignment = { vertical: 'middle', horizontal: 'center' };

    sheet.getCell(2, startCol).value = 'In';
    sheet.getCell(2, startCol + 1).value = 'Out';
    sheet.getCell(2, startCol + 2).value = 'Remarks';

    startCol += 3;
  }
}



//find date column
function findDateColumns(sheet, targetDate) {
  const headerRow1 = sheet.getRow(1);
  const formattedTarget = `${String(new Date(targetDate).getDate()).padStart(2, '0')}-${String(new Date(targetDate).getMonth() + 1).padStart(2, '0')}-${new Date(targetDate).getFullYear()}`;

  for (let col = 5; col <= sheet.columnCount; col += 3) {
    const val = headerRow1.getCell(col).value;
    if (val && typeof val === 'string' && val.trim() === formattedTarget) {
      return {
        inCol: col,
        outCol: col + 1,
        remarksCol: col + 2
      };
    }
  }
  return null;
}



//find or create material row
function createMissingDates(sheet, targetDate) {
  const headerRow1 = sheet.getRow(1);
  const headerRow2 = sheet.getRow(2);

  const formatDate = (dateObj) =>
    `${String(dateObj.getDate()).padStart(2, '0')}-${String(dateObj.getMonth() + 1).padStart(2, '0')}-${dateObj.getFullYear()}`;

  const startCol = 5;
  const existingDates = new Set();

  for (let col = startCol; col <= sheet.columnCount; col += 3) {
    const val = headerRow1.getCell(col).value;
    if (val && typeof val === 'string' && !isNaN(new Date(val))) {
      existingDates.add(val);
    }
  }

  const target = new Date(targetDate);
  const targetMonth = target.getMonth();
  const targetYear = target.getFullYear();
  const daysInMonth = new Date(targetYear, targetMonth + 1, 0).getDate();

  let current = new Date(targetYear, targetMonth, 1);
  let insertCol = sheet.columnCount + 1;

  for (let day = 1; day <= daysInMonth; day++) {
    const dateStr = formatDate(current);
    if (!existingDates.has(dateStr)) {
      sheet.mergeCells(1, insertCol, 1, insertCol + 2);
      sheet.getCell(1, insertCol).value = dateStr;
      sheet.getCell(1, insertCol).alignment = { vertical: 'middle', horizontal: 'center' };

      headerRow2.getCell(insertCol).value = 'In';
      headerRow2.getCell(insertCol + 1).value = 'Out';
      headerRow2.getCell(insertCol + 2).value = 'Remarks';

      insertCol += 3;
    }
    current.setDate(current.getDate() + 1);
  }
}


//dates manually
const monthDays = {
  1: 31,
  2: 28, // We'll adjust for leap year separately
  3: 31,
  4: 30,
  5: 31,
  6: 30,
  7: 31,
  8: 31,
  9: 30,
  10: 31,
  11: 30,
  12: 31
};

// ✅ Helper to check leap year
function isLeapYear(year) {
  return (year % 4 === 0 && year % 100 !== 0) || (year % 400 === 0);
}

// ✅ Initialize monthly sheets if missing
async function initializeMonthlySheets(workbook, targetDate) {
  const month = ('0' + (targetDate.getMonth() + 1)).slice(-2); // "04"
  const year = targetDate.getFullYear();

  const stockSheetName = `Stock ${month}`;
  const stockInSheetName = `Stock In ${month}`;
  const stockOutSheetName = `Stock Out ${month}`;

  // Check if the sheets exist already
  let stockSheet = workbook.getWorksheet(stockSheetName);
  let stockInSheet = workbook.getWorksheet(stockInSheetName);
  let stockOutSheet = workbook.getWorksheet(stockOutSheetName);

  // If any missing, create
  if (!stockSheet) {
    stockSheet = workbook.addWorksheet(stockSheetName);
    setupStockSheet(stockSheet, year, month);
  }
  if (!stockInSheet) {
    stockInSheet = workbook.addWorksheet(stockInSheetName);
    stockInSheet.addRow(["Date", "Material", "Invoice No.", "Quantity"]);
  }
  if (!stockOutSheet) {
    stockOutSheet = workbook.addWorksheet(stockOutSheetName);
    stockOutSheet.addRow(["Date", "Material", "Quantity", "Remarks"]);
  }
}

function setupStockSheet(sheet, year, month) {
  const headers = ['Material', 'Opening Stock', 'Current Stock', 'Min Stock'];
  const blueGreyFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'D9D9D9' } };

  for (let i = 0; i < headers.length; i++) {
    const cell = sheet.getCell(1, i + 1);
    cell.value = headers[i];
    cell.font = { bold: true };
    cell.fill = blueGreyFill;
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    sheet.mergeCells(1, i + 1, 2, i + 1);
  }

  let daysInMonth = monthDays[parseInt(month)];
  if (parseInt(month) === 2 && isLeapYear(year)) daysInMonth = 29;

  let col = 5;
  for (let day = 1; day <= daysInMonth; day++) {
    const dateStr = `${String(day).padStart(2, '0')}-${month}-${year}`;

    sheet.mergeCells(1, col, 1, col + 2);
    const headerCell = sheet.getCell(1, col);
    headerCell.value = dateStr;
    headerCell.font = { bold: true };
    headerCell.alignment = { vertical: 'middle', horizontal: 'center' };

    const greenFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E2F0D9' } };
    const orangeFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F9CB9C' } };
    const blueGreyFillLight = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'D9D9D9' } };

    const inCell = sheet.getCell(2, col);
    const outCell = sheet.getCell(2, col + 1);
    const remarksCell = sheet.getCell(2, col + 2);

    inCell.value = 'In'; inCell.font = { bold: true }; inCell.fill = greenFill;
    inCell.alignment = { vertical: 'middle', horizontal: 'center' };

    outCell.value = 'Out'; outCell.font = { bold: true }; outCell.fill = orangeFill;
    outCell.alignment = { vertical: 'middle', horizontal: 'center' };

    remarksCell.value = 'Remarks'; remarksCell.font = { bold: true }; remarksCell.fill = blueGreyFillLight;
    remarksCell.alignment = { vertical: 'middle', horizontal: 'center' };

    col += 3;
  }

  // Optional Column Widths
  const widths = [20, 15, 15, 15];
  for (let i = 0; i < widths.length; i++) {
    sheet.getColumn(i + 1).width = widths[i];
  }
}



//search current stock
app.get('/search-current-stock', async (req, res) => {
  const { material, month } = req.query;
  if (!material || !month) return res.status(400).json({ error: 'Material and Month required' });

  try {
    const fileName = 'Stock Sheet.xlsx';
    const { workbook } = await loadWorkbookFromOneDrive(fileName);

    const sheet = workbook.getWorksheet(`Stock ${month}`);
    if (!sheet) return res.status(404).json({ error: `Stock ${month} Sheet not found` });

    for (let i = 3; i <= sheet.rowCount; i++) {
      const mat = sheet.getCell(`A${i}`).value;
      if (mat && mat.toString().toLowerCase() === material.toLowerCase()) {
        const currentStock = sheet.getCell(`C${i}`).value || 0; // Column C
        return res.json({ material: mat, currentStock });
      }
    }

    return res.status(404).json({ error: 'Material not found in selected month.' });
  } catch (err) {
    console.error('❌ Error in search-current-stock:', err.message);
    res.status(500).json({ error: 'Internal Server Error' });
  }
});


//search min. stock
app.get('/search-min-stock', async (req, res) => {
  const { month } = req.query;
  if (!month) return res.status(400).json({ error: 'Month required' });

  try {
    const fileName = 'Stock Sheet.xlsx';
    const { workbook } = await loadWorkbookFromOneDrive(fileName);

    const sheet = workbook.getWorksheet(`Stock ${month}`);
    if (!sheet) return res.status(404).json({ error: `Stock ${month} Sheet not found` });

    const result = [];

    for (let i = 3; i <= sheet.rowCount; i++) {
      const material = sheet.getCell(`A${i}`).value;
      const currentStock = parseFloat(sheet.getCell(`C${i}`).value) || 0;
      const minStock = parseFloat(sheet.getCell(`D${i}`).value) || 0;

      if (material && currentStock <= minStock) {
        result.push({ material, currentStock });
      }
    }

    res.json(result);
  } catch (err) {
    console.error('❌ Error in search-min-stock:', err.message);
    res.status(500).json({ error: 'Internal Server Error' });
  }
});


app.get('/search-stock-by-date', async (req, res) => {
  const { date } = req.query;
  if (!date) return res.status(400).json({ error: 'Date is required' });

  try {
    const fileName = 'Stock Sheet.xlsx';
    const month = ('0' + (new Date(date).getMonth() + 1)).slice(-2);

    const { workbook } = await loadWorkbookFromOneDrive(fileName);
    const sheet = workbook.getWorksheet(`Stock ${month}`);
    if (!sheet) return res.status(404).json({ error: `Stock ${month} Sheet not found` });

    const headerRow = sheet.getRow(1);
    const formattedDate = `${String(new Date(date).getDate()).padStart(2, '0')}-${month}-${new Date(date).getFullYear()}`;

    let foundCol = null;
    for (let col = 5; col <= sheet.columnCount; col += 3) {
      if (headerRow.getCell(col).value === formattedDate) {
        foundCol = col;
        break;
      }
    }

    if (!foundCol) return res.status(404).json({ error: 'Date not found in Sheet.' });

    const result = [];

    for (let i = 3; i <= sheet.rowCount; i++) {
      const material = sheet.getCell(`A${i}`).value;
      if (material) {
        const inQty = sheet.getCell(i, foundCol).value || 0;
        const outQty = sheet.getCell(i, foundCol + 1).value || 0;

        if (inQty !== 0 || outQty !== 0) {
          result.push({ material, in: inQty, out: outQty });
        }
      }
    }

    res.json(result);
  } catch (err) {
    console.error('❌ Error in /search-stock-by-date:', err.message);
    res.status(500).json({ error: 'Internal Server Error' });
  }
});

app.get('/api/getDashboardStats', async (req, res) => {
  try {
    res.set('Cache-Control', 'no-store'); // Ensure no caching

    const fileName = 'TempData.xlsx';
    const { workbook } = await getWorkbookFromOneDrive(fileName);
    const sheet = workbook.getWorksheet('Client Data');
    if (!sheet) return res.status(404).json({ error: 'Client Data sheet not found' });

    let totalSalesRevenue = 0;
    let totalBalance = 0;
    let plantsInstalled = 0;

    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // Skip header row

      const totalCost = parseFloat(row.getCell(8).value) || 0;    // Column H
      const balanceRaw = row.getCell(35).value;                   // Column AI

      totalSalesRevenue += totalCost;

      let balance = null;
      if (balanceRaw !== null && balanceRaw !== '' && balanceRaw !== '-' && balanceRaw !== '.') {
        if (typeof balanceRaw === 'string') {
          balance = parseFloat(balanceRaw.trim()) || 0;
        } else if (typeof balanceRaw === 'number') {
          balance = balanceRaw;
        }

        totalBalance += balance;

        if (balance === 0) {
          plantsInstalled += 1;
        }
      }
    });

    const totalPaymentReceived = totalSalesRevenue - totalBalance;

    res.json({ totalSalesRevenue, totalPaymentReceived, plantsInstalled });
  } catch (error) {
    console.error('❌ Error in getDashboardStats:', error.message);
    res.status(500).json({ error: 'Internal Server Error' });
  }
});





app.get('/api/getPieData', async (req, res) => {
  try {
    res.set('Cache-Control', 'no-store'); // Always live

    const fileName = 'TempData.xlsx';
    const { workbook } = await getWorkbookFromOneDrive(fileName);
    const sheet = workbook.getWorksheet('Client Data');
    if (!sheet) return res.status(404).json({ error: 'Client Data sheet not found' });

    let totalAmount = 0;   // Total Cost
    let totalBalance = 0;  // Remaining Balance

    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // Skip header

      const totalCost = parseFloat(row.getCell(8).value) || 0;    // Column H
      const balanceRaw = row.getCell(35).value;                   // Column AI

      totalAmount += totalCost;

      let balance = 0;
      if (balanceRaw !== null && balanceRaw !== '' && balanceRaw !== '-' && balanceRaw !== '.') {
        balance = typeof balanceRaw === 'string'
          ? parseFloat(balanceRaw.trim()) || 0
          : (typeof balanceRaw === 'number' ? balanceRaw : 0);
      }

      totalBalance += balance;
    });

    res.json({ totalAmount, totalBalance });
  } catch (error) {
    console.error('❌ Error in getPieData:', error.message);
    res.status(500).json({ error: 'Internal Server Error' });
  }
});


app.get('/api/getApplicationStatusData', async (req, res) => {
  try {
    res.set('Cache-Control', 'no-store'); // Live data only

    const fileName = 'TempData.xlsx';
    const { workbook } = await getWorkbookFromOneDrive(fileName);
    const sheet = workbook.getWorksheet('Client Data');
    if (!sheet) return res.status(404).json({ error: 'Client Data sheet not found' });

    let applicationApplied = 0;
    let applicationPending = 0;

    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // Skip header

      const applicationCell = row.getCell(19).value; // Column S

      if (applicationCell !== null && applicationCell !== '' && applicationCell.toString().toLowerCase() !== 'no') {
        applicationApplied += 1;
      } else {
        applicationPending += 1;
      }
    });

    res.json({ applied: applicationApplied, pending: applicationPending });
  } catch (error) {
    console.error('❌ Error in getApplicationStatusData:', error.message);
    res.status(500).json({ error: 'Internal Server Error' });
  }
});


app.get('/api/getBarGraphData', async (req, res) => {
  try {
    res.set('Cache-Control', 'no-store');

    const fileName = 'TempData.xlsx';
    const { workbook } = await getWorkbookFromOneDrive(fileName);
    const sheet = workbook.getWorksheet('Client Data');
    if (!sheet) return res.status(404).json({ error: 'Client Data sheet not found' });

    let totalCostSum = 0;
    let advanceSum = 0;
    let secondInstallmentReceivedSum = 0;
    let finalInstallmentReceivedSum = 0;

    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // Skip header

      const totalCost = parseFloat(row.getCell(8).value) || 0;     // H
      const advance = parseFloat(row.getCell(7).value) || 0;       // G
      const secondInstallment = parseFloat(row.getCell(32).value) || 0; // AF
      const finalInstallment = parseFloat(row.getCell(34).value) || 0;  // AH

      totalCostSum += totalCost;
      advanceSum += advance;
      secondInstallmentReceivedSum += secondInstallment;
      finalInstallmentReceivedSum += finalInstallment;
    });

    const sixtyPercentOfTotalCost = 0.6 * totalCostSum;

    const secondInstallmentDue = totalCostSum - (advanceSum + sixtyPercentOfTotalCost);
    const finalInstallmentDue = totalCostSum - (advanceSum + secondInstallmentReceivedSum);

    res.json({
      totalCostSum,
      advanceSum,
      secondInstallmentReceivedSum,
      secondInstallmentDue,
      finalInstallmentReceivedSum,
      finalInstallmentDue
    });
  } catch (error) {
    console.error('❌ Error in getBarGraphData:', error.message);
    res.status(500).json({ error: 'Internal Server Error' });
  }
});


app.get('/api/getPaymentsData', async (req, res) => {
  try {
    res.set('Cache-Control', 'no-store');

    const fileName = 'TempData.xlsx';
    const { workbook } = await getWorkbookFromOneDrive(fileName);
    const sheet = workbook.getWorksheet('Client Data');
    if (!sheet) return res.status(404).json({ error: 'Client Data sheet not found' });

    const payments = [];

    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // Skip header

      const customerName = row.getCell(2).value || '';
      const totalCost = parseFloat(row.getCell(8).value) || 0;
      const advance = parseFloat(row.getCell(7).value) || 0;
      const secondInstallment = row.getCell(32).value;
      const finalInstallment = row.getCell(34).value;
      const balance = parseFloat(row.getCell(35).value) || 0;

      if (balance !== 0) {
        payments.push({
          customerName,
          totalCost,
          advance,
          secondInstallment: (secondInstallment === null || secondInstallment === '' || secondInstallment === undefined) ? 'Due' : secondInstallment,
          finalInstallment: (finalInstallment === null || finalInstallment === '' || finalInstallment === undefined) ? 'Due' : finalInstallment,
          balance
        });
      }
    });

    res.json({ payments });
  } catch (error) {
    console.error('❌ Error fetching Payments Data:', error.message);
    res.status(500).json({ error: 'Internal Server Error' });
  }
});



app.post('/api/addTask', async (req, res) => {
  try {
    const { date, time, description } = req.body;
    const fileName = 'TempData.xlsx';

    const { workbook, token } = await getWorkbookFromOneDrive(fileName);
    const sheet = workbook.getWorksheet('Client Data');
    if (!sheet) return res.status(404).json({ error: 'Client Data sheet not found' });

    // Find first empty row in column AK (37)
    let rowToUse;
    sheet.eachRow((row, rowNumber) => {
      if (!row.getCell(37).value && !rowToUse) {
        rowToUse = rowNumber;
      }
    });

    if (!rowToUse) {
      rowToUse = sheet.lastRow.number + 1;
    }

    sheet.getRow(rowToUse).getCell(37).value = date;         // AK
    sheet.getRow(rowToUse).getCell(38).value = time;         // AL
    sheet.getRow(rowToUse).getCell(39).value = description;  // AM

    await uploadWorkbookToOneDrive(fileName, workbook, token);
    res.json({ success: true });

  } catch (error) {
    console.error('❌ Error adding task:', error.message);
    res.status(500).json({ error: 'Failed to add task' });
  }
});


app.get('/get-next-refno', async (req, res) => {
  const fileName = 'leads.xlsx';
  const workbook = new ExcelJS.Workbook();

  try {
    const { workbook } = await getWorkbookFromOneDrive(fileName);
    const sheet = workbook.getWorksheet(1);
    if (!sheet) return res.status(404).send('Sheet not found');

    let lastRef = null;

    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // Skip header
      const refCell = row.getCell(5).value; // Column E
      if (refCell && typeof refCell === 'string' && /^[A-Z]\d{4}$/.test(refCell)) {
        lastRef = refCell;
      }
    });

    let nextRef = 'A0001';
    if (lastRef) {
      const letter = lastRef.charAt(0);
      const number = parseInt(lastRef.slice(1));
      if (number < 9999) {
        nextRef = letter + (number + 1).toString().padStart(4, '0');
      } else {
        const nextChar = String.fromCharCode(letter.charCodeAt(0) + 1);
        nextRef = nextChar + '0001';
      }
    }

    res.send(nextRef);
  } catch (err) {
    console.error("❌ Error generating next ref no:", err.message);
    res.status(500).send('Error');
  }
});



app.post('/save-lead', async (req, res) => {
  const { date, name, address, mobile, refno, kw, reference } = req.body;
  const fileName = 'leads.xlsx';

  try {
    const token = await getAccessToken();

    // 📥 Download workbook from OneDrive
    const response = await axios.get(
      `https://graph.microsoft.com/v1.0/users/muninderpal@jk17.onmicrosoft.com/drive/root:/${fileName}:/content`,
      {
        headers: { Authorization: `Bearer ${token}` },
        responseType: 'arraybuffer'
      }
    );

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(response.data);

    let sheet = workbook.getWorksheet(1);

    // 🧾 If sheet doesn't exist, create it with headers
    if (!sheet) {
      sheet = workbook.addWorksheet('Leads');
      sheet.addRow(['Date', 'Consumer Name', 'Address', 'Mobile No.', 'Ref No.', 'KW', 'Reference']);
    }

    // ➕ Add new lead
    sheet.addRow([date, name, address, mobile, refno, kw, reference]);

    // 📤 Upload workbook back to OneDrive
    const buffer = await workbook.xlsx.writeBuffer();
    await axios.put(
      `https://graph.microsoft.com/v1.0/users/muninderpal@jk17.onmicrosoft.com/drive/root:/${fileName}:/content`,
      buffer,
      {
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        }
      }
    );

    res.sendStatus(200);
  } catch (err) {
    console.error('❌ Error saving lead:', err.message);
    res.status(500).send('Failed to save');
  }
});




app.get('/get-leads', async (req, res) => {
  const fileName = 'leads.xlsx';

  try {
    const { workbook } = await getWorkbookFromOneDrive(fileName);
    const sheet = workbook.getWorksheet(1);
    if (!sheet) return res.status(404).send('Sheet not found');

    const data = [];

    sheet.eachRow((row, rowNumber) => {
      const rowData = row.values.slice(1); // Remove blank first
      if (rowNumber === 1 || rowData[6]?.toString().toLowerCase() !== 'no') {
        data.push(rowData);
      }
    });

    res.json(data);
  } catch (err) {
    console.error('❌ Error reading leads:', err.message);
    res.status(500).send('Failed to read');
  }
});



app.post('/update-lead', async (req, res) => {
  const { field, rowIndex, value } = req.body;
  const fileName = 'leads.xlsx';

  try {
    const { workbook, token } = await getWorkbookFromOneDrive(fileName);
    const sheet = workbook.getWorksheet(1);
    if (!sheet) return res.status(404).send('Sheet not found');

    const fieldMap = {
      call: 8,
      proposal: 9,
      meeting: 10,
      reminder: 11,
      status: 12,
      final: 13
    };

    const row = sheet.getRow(rowIndex + 2); // Adjust for header and 0-index
    const colIndex = fieldMap[field];

    if (row && colIndex) {
      row.getCell(colIndex).value = value;
      row.commit();

      await uploadWorkbookToOneDrive(fileName, workbook, token);
      res.sendStatus(200);
    } else {
      res.status(400).send("Invalid row/column");
    }
  } catch (err) {
    console.error("❌ Error updating lead:", err.message);
    res.status(500).send("Update failed");
  }
});

app.post('/delete-lead', async (req, res) => {
  const { rowIndex } = req.body;
  const fileName = 'leads.xlsx';

  try {
    const { workbook, token } = await getWorkbookFromOneDrive(fileName);
    const sheet = workbook.getWorksheet(1);
    if (!sheet) return res.status(404).send('Sheet not found');

    const actualRow = rowIndex + 2; // Adjust for header

    sheet.spliceRows(actualRow, 1); // Delete 1 row
    await uploadWorkbookToOneDrive(fileName, workbook, token);

    res.sendStatus(200);
  } catch (err) {
    console.error("❌ Error deleting lead:", err.message);
    res.status(500).send("Delete failed");
  }
});


//neww proposall
app.post('/save-proposal', async (req, res) => {
  const data = req.body;
  const fileName = 'proposal.xlsx';

  try {
    let { workbook, token } = await getWorkbookFromOneDrive(fileName);
    let sheet = workbook.getWorksheet('Proposals');

    // If file or sheet is missing, create structure
    if (!sheet) {
      sheet = workbook.addWorksheet('Proposals');
      sheet.addRow([
        'Ref', 'Date', 'Subsidy', 'KW', 'Address', 'State', 'City', 
        'To Whom', 'Mobile', 'Price', 'Panel Brand', 'Panel Wp', 'Inverter Brand', 'sentBy', 'Battery Info', 'Battery Warranty', 'Battery Qty', 'Battery Name'
      ]);
    }

    // Auto-generate Ref No like JK-5-9999999999
    const lastRow = sheet.lastRow;
    let newNumber = 1;
    if (lastRow && lastRow.getCell(1).value && lastRow.getCell(1).value.toString().endsWith('P')) {
      const lastRef = lastRow.getCell(1).value.toString().replace('P', '');
      newNumber = parseInt(lastRef) + 1;
    }
    const newRef = `JK-${data.kw}-${data.mobile}`;

    sheet.addRow([
      newRef,
      data.date,
      data.subsidy,
      data.kw,
      data.address,
      data.state,
      data.city,
      data.toWhom,
      data.mobile,
      data.price,
      data.panelBrand,
      data.panelWp,
      data.inverterBrand,
      data.sentBy,
      data.batteryinfo,
      data.batterywarranty,
      data.batteryqty, 
      data.batteryname
    ]);

    await uploadWorkbookToOneDrive(fileName, workbook, token);
    res.json({ success: true, ref: newRef });

  } catch (err) {
    console.error('❌ Proposal save error:', err.message);
    res.status(500).json({ success: false, error: 'Proposal saving failed' });
  }
});



app.get('/get-proposal', async (req, res) => {
  const ref = req.query.ref;
  const fileName = 'proposal.xlsx';

  if (!ref) return res.status(400).json({ success: false, error: 'Missing reference number' });

  try {
    const { workbook } = await getWorkbookFromOneDrive(fileName);
    const sheet = workbook.getWorksheet('Proposals');
    if (!sheet) return res.status(404).json({ success: false, error: 'Proposals sheet not found' });

    let matchedRow;

    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // Skip header
      if (row.getCell(1).value === ref) {
        matchedRow = row;
      }
    });

    if (!matchedRow) return res.status(404).json({ success: false, error: 'Proposal not found' });

    const headers = sheet.getRow(1).values.slice(1); // Ignore index 0
    const values = matchedRow.values.slice(1);
    const proposalData = {};

    headers.forEach((header, i) => {
      proposalData[header.trim().toLowerCase().replace(/\s+/g, '')] = values[i];
    });

    res.json({ success: true, data: proposalData });

  } catch (err) {
    console.error('❌ Error reading proposal.xlsx from OneDrive:', err.message);
    res.status(500).json({ success: false, error: 'Error reading Excel from cloud' });
  }
});





app.get('/generate-pdf', async (req, res) => {
  const ref = req.query.ref;
  if (!ref) return res.status(400).send('Missing reference number');

  const previewURL = `https://jk-crm-0qal.onrender.com/proposal-preview.html?ref=${ref}`;

  try {
    const browser = await puppeteer.launch({
      headless: 'new',
      args: ['--no-sandbox', '--disable-setuid-sandbox'] // ✅ Required for Render or headless hosting
    });

    const page = await browser.newPage();
    await page.goto(previewURL, { waitUntil: 'networkidle0' });

    const pdfBuffer = await page.pdf({
      format: 'A4',
      printBackground: true,
      margin: { top: '10mm', bottom: '10mm', left: '10mm', right: '10mm' }
    });

    await browser.close();

    res.set({
      'Content-Type': 'application/pdf',
      'Content-Disposition': `attachment; filename="JK_Solar_Proposal_${ref}.pdf"`
    });

    res.send(pdfBuffer);
  } catch (err) {
    console.error('❌ PDF Generation Error:', err.message);
    res.status(500).send('Failed to generate PDF');
  }
});


app.get('/get-proposals', async (req, res) => {
  const fileName = 'proposal.xlsx';

  try {
    const { workbook } = await getWorkbookFromOneDrive(fileName);
    const sheet = workbook.getWorksheet('Proposals');

    if (!sheet) return res.json([]);

    const data = [];
    const headers = sheet.getRow(1).values.slice(1); // Skip index 0

    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // skip header

      const values = row.values.slice(1); // Remove first empty
      const rowData = {};

      headers.forEach((header, i) => {
        rowData[header.trim().toLowerCase().replace(/\s+/g, '')] = values[i];
      });

      data.push(rowData);
    });

    res.json(data);
  } catch (err) {
    console.error('❌ Error reading proposals from OneDrive:', err.message);
    res.status(500).json({ error: 'Failed to read proposal data.' });
  }
});



async function transferProposalToLeads() {
  const proposalFile = 'proposal.xlsx';
  const leadsFile = 'leads.xlsx';

  const { workbook: proposalWorkbook, token } = await getWorkbookFromOneDrive(proposalFile);
  const proposalSheet = proposalWorkbook.getWorksheet(1);

  const { workbook: leadsWorkbook } = await getWorkbookFromOneDrive(leadsFile);
  const leadsSheet = leadsWorkbook.getWorksheet(1);

  for (let i = 2; i <= proposalSheet.rowCount; i++) {
    const rowProposal = proposalSheet.getRow(i);
    const proposalKW = String(rowProposal.getCell('D').value).trim();
    const proposalName = String(rowProposal.getCell('H').value).trim();
    const proposalMobile = String(rowProposal.getCell('I').value).trim();

    let replaced = false;

    for (let j = 2; j <= leadsSheet.rowCount; j++) {
      const rowLead = leadsSheet.getRow(j);
      const leadKW = String(rowLead.getCell(6).value).trim();
      const leadName = String(rowLead.getCell(2).value).trim();
      const leadMobile = String(rowLead.getCell(4).value).trim();

      if (proposalKW === leadKW && proposalName === leadName && proposalMobile === leadMobile) {
        leadsSheet.spliceRows(j, 1);
        leadsSheet.insertRow(j, [
          rowProposal.getCell('B').value,
          rowProposal.getCell('H').value,
          rowProposal.getCell('E').value,
          rowProposal.getCell('I').value,
          rowProposal.getCell('A').value,
          rowProposal.getCell('D').value,
          '',
          '',
          'Sent'
        ]);
        console.log(`🔁 Replaced row ${j} for ${proposalName}, ${proposalMobile}, ${proposalKW}`);
        replaced = true;
        break;
      }
    }

    if (!replaced) {
      leadsSheet.addRow([
        rowProposal.getCell('B').value,
        rowProposal.getCell('H').value,
        rowProposal.getCell('E').value,
        rowProposal.getCell('I').value,
        rowProposal.getCell('A').value,
        rowProposal.getCell('D').value,
        '',
        '',
        'Sent'
      ]);
      console.log(`🆕 Added new row for ${proposalName}, ${proposalMobile}, ${proposalKW}`);
    }
  }

  await uploadWorkbookToOneDrive(leadsFile, leadsWorkbook, token);
  console.log('✅ Leads file updated successfully.');
}



app.get('/api/getNotes', async (req, res) => {
  try {
    const fileName = 'TempData.xlsx';
    const { workbook } = await getWorkbookFromOneDrive(fileName);
    const sheet = workbook.getWorksheet('Client Data');
    const notes = [];

    if (!sheet) return res.json({ notes: [] });

    for (let i = 2; i <= sheet.rowCount; i++) {
      const note = sheet.getRow(i).getCell(38).value; // Column AL
      if (note) {
        notes.push(note.toString());
      }
    }

    res.json({ notes });
  } catch (err) {
    console.error('❌ Error fetching notes:', err.message);
    res.status(500).json({ error: 'Failed to fetch notes' });
  }
});



app.post('/api/addNote', async (req, res) => {
  const { note } = req.body;
  const fileName = 'TempData.xlsx';

  try {
    const { workbook, token } = await getWorkbookFromOneDrive(fileName);
    const sheet = workbook.getWorksheet('Client Data');

    if (!sheet) return res.status(404).json({ error: 'Sheet not found' });

    let rowToWrite = sheet.actualRowCount + 1;

    // Find the next empty row in column AL (38)
    for (let i = 2; i <= sheet.rowCount + 100; i++) {
      const cell = sheet.getRow(i).getCell(38).value;
      if (!cell || cell === '') {
        rowToWrite = i;
        break;
      }
    }

    sheet.getRow(rowToWrite).getCell(38).value = note;

    await uploadWorkbookToOneDrive(fileName, workbook, token);
    res.sendStatus(200);
  } catch (err) {
    console.error('❌ Error adding note:', err.message);
    res.status(500).json({ error: 'Failed to add note' });
  }
});


app.post('/api/deleteNote', async (req, res) => {
  const { index } = req.body;
  const fileName = 'TempData.xlsx';

  try {
    const { workbook, token } = await getWorkbookFromOneDrive(fileName);
    const sheet = workbook.getWorksheet('Client Data');
    if (!sheet) return res.status(404).json({ error: 'Sheet not found' });

    let rowIndex = 0, count = 0;

    sheet.eachRow((row, i) => {
      if (row.getCell(38).value) {
        if (count === index) rowIndex = i;
        count++;
      }
    });

    if (rowIndex > 0) {
      sheet.getRow(rowIndex).getCell(38).value = null;
      await uploadWorkbookToOneDrive(fileName, workbook, token);
    }

    res.sendStatus(200);
  } catch (err) {
    console.error('❌ Error deleting note:', err.message);
    res.status(500).json({ error: 'Failed to delete note' });
  }
});


// 🎟️ Get Access Token using Client Credentials
async function getAccessToken() {
  const tokenUrl = `https://login.microsoftonline.com/785fd7e9-594d-4549-91b9-9372f7295962/oauth2/v2.0/token`;

  const data = qs.stringify({
    grant_type: 'client_credentials',
    client_id: '89a49313-0f16-44c3-9f71-cf96eab166ad',
    client_secret: 'IZ-8Q~GaHcwhQnrCCj~ZH_I_3bHsZYpoC1xm2aLk',
    scope: 'https://graph.microsoft.com/.default',
  });

  const response = await axios.post(tokenUrl, data, {
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
  });

   console.log("🎟️ Access Token:", response.data.access_token);
  return response.data.access_token;
}

// 📥 Fetch leads.xlsx from OneDrive and return as ExcelJS workbook
async function downloadExcelFromOneDrive() {
  try {
    const accessToken = await getAccessToken();

    // Update URL to include the correct user endpoint
    const fileUrl = "https://graph.microsoft.com/v1.0/users/muninderpal@jk17.onmicrosoft.com/drive/root:/leads.xlsx:/content";

    const response = await axios.get(fileUrl, {
      headers: { Authorization: `Bearer ${accessToken}` },
      responseType: 'arraybuffer',
    });

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(response.data);

    console.log("✅ leads.xlsx downloaded and loaded from OneDrive");
    return workbook;
  } catch (err) {
    console.error("❌ Error downloading file from OneDrive:", err.response?.data || err.message);
    throw err;
  }
}

async function getWorkbookFromOneDrive(fileName) {
  const token = await getAccessToken();

  // Updated URL to use users/<user-email> for app-only token
  const fileUrl = `https://graph.microsoft.com/v1.0/users/muninderpal@jk17.onmicrosoft.com/drive/root:/${fileName}:/content`;

  const response = await axios.get(fileUrl, {
    headers: { Authorization: `Bearer ${token}` },
    responseType: 'arraybuffer'
  });

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(response.data);

  return { workbook, token };
}

async function uploadWorkbookToOneDrive(fileName, workbook, token) {
  const buffer = await workbook.xlsx.writeBuffer();

  // Updated URL for app-only access
  const uploadUrl = `https://graph.microsoft.com/v1.0/users/muninderpal@jk17.onmicrosoft.com/drive/root:/${fileName}:/content`;

  await axios.put(uploadUrl, buffer, {
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }
  });

  console.log(`✅ Uploaded ${fileName} to OneDrive.`);
}



module.exports = {
  getAccessToken,
  getWorkbookFromOneDrive,
  uploadWorkbookToOneDrive
};

app.post('/api/deleteNote', async (req, res) => {
  const { index } = req.body;
  const fileName = 'TempData.xlsx';

  try {
    const { workbook, token } = await getWorkbookFromOneDrive(fileName);
    const sheet = workbook.getWorksheet('Client Data');
    if (!sheet) return res.status(404).json({ error: 'Sheet not found' });

    let rowIndex = 0, count = 0;

    sheet.eachRow((row, i) => {
      if (row.getCell(38).value) {
        if (count === index) rowIndex = i;
        count++;
      }
    });

    if (rowIndex > 0) {
      sheet.getRow(rowIndex).getCell(38).value = null;
      await uploadWorkbookToOneDrive(fileName, workbook, token);
    }

    res.sendStatus(200);
  } catch (err) {
    console.error('❌ Error deleting note:', err.message);
    res.status(500).json({ error: 'Failed to delete note' });
  }
});


// 🎟️ Get Access Token using Client Credentials
async function getAccessToken() {
  const tokenUrl = `https://login.microsoftonline.com/785fd7e9-594d-4549-91b9-9372f7295962/oauth2/v2.0/token`;

  const data = qs.stringify({
    grant_type: 'client_credentials',
    client_id: '89a49313-0f16-44c3-9f71-cf96eab166ad',
    client_secret: 'IZ-8Q~GaHcwhQnrCCj~ZH_I_3bHsZYpoC1xm2aLk',
    scope: 'https://graph.microsoft.com/.default',
  });

  const response = await axios.post(tokenUrl, data, {
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
  });

   console.log("🎟️ Access Token:", response.data.access_token);
  return response.data.access_token;
}

// 📥 Fetch leads.xlsx from OneDrive and return as ExcelJS workbook
async function downloadExcelFromOneDrive() {
  try {
    const accessToken = await getAccessToken();

    // Update URL to include the correct user endpoint
    const fileUrl = "https://graph.microsoft.com/v1.0/users/muninderpal@jk17.onmicrosoft.com/drive/root:/leads.xlsx:/content";

    const response = await axios.get(fileUrl, {
      headers: { Authorization: `Bearer ${accessToken}` },
      responseType: 'arraybuffer',
    });

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(response.data);

    console.log("✅ leads.xlsx downloaded and loaded from OneDrive");
    return workbook;
  } catch (err) {
    console.error("❌ Error downloading file from OneDrive:", err.response?.data || err.message);
    throw err;
  }
}

async function getWorkbookFromOneDrive(fileName) {
  const token = await getAccessToken();

  // Updated URL to use users/<user-email> for app-only token
  const fileUrl = `https://graph.microsoft.com/v1.0/users/muninderpal@jk17.onmicrosoft.com/drive/root:/${fileName}:/content`;

  const response = await axios.get(fileUrl, {
    headers: { Authorization: `Bearer ${token}` },
    responseType: 'arraybuffer'
  });

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(response.data);

  return { workbook, token };
}

async function uploadWorkbookToOneDrive(fileName, workbook, token) {
  const buffer = await workbook.xlsx.writeBuffer();

  // Updated URL for app-only access
  const uploadUrl = `https://graph.microsoft.com/v1.0/users/muninderpal@jk17.onmicrosoft.com/drive/root:/${fileName}:/content`;

  await axios.put(uploadUrl, buffer, {
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }
  });
console.log(`✅ Uploaded ${fileName} to OneDrive.`);
}

module.exports = {
  getAccessToken,
  getWorkbookFromOneDrive,
  uploadWorkbookToOneDrive
};




//Login page starts here
//Login page starts here
//Login page starts here
//Login page starts here

app.post('/login', async (req, res) => {
  try {
    const { username, password } = req.body;
    const token = await getAccessToken();

    const response = await axios.get(
      `https://graph.microsoft.com/v1.0/users/muninderpal@jk17.onmicrosoft.com/drive/root:/credentials.xlsx:/content`,
      {
        headers: { Authorization: `Bearer ${token}` },
        responseType: 'arraybuffer'
      }
    );

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(response.data);
    const sheet = workbook.getWorksheet('Users') || workbook.worksheets[0];

    let valid = false;

    sheet.eachRow((row, index) => {
      if (index === 1) return; // skip header
      const storedUsername = row.getCell(1).value?.toString().trim();
      const storedPassword = row.getCell(2).value?.toString().trim();
      const role = row.getCell(3).value?.toString().trim().toLowerCase() || '';
      const status = row.getCell(4).value?.toString().trim();

      if (
        storedUsername.toLowerCase() === username.toLowerCase() &&
        storedPassword === password &&
        status.toLowerCase() === 'active'
      ) {
        // ✅ Save both username and role in session
        req.session.user = {
          username: storedUsername,
          role: role
        };
        valid = true;
      }
    });

    if (valid) {
      res.json({ success: true, redirect: "index.html" });
    } else {
      res.json({ success: false, message: "Invalid username, password, or inactive user" });
    }

  } catch (err) {
    console.error("❌ Login error:", err.message);
    res.status(500).json({ success: false, message: "Server error during login" });
  }
});



app.use('/index.html', (req, res, next) => {
  if (req.session.user && req.session.user.role === 'Admin') {
    next();
  } else {
    res.redirect('/login.html');
  }
});

app.get('/logout', (req, res) => {
  req.session.destroy();
  res.redirect('/login.html');
});

app.get('/login.html', (req, res) => {
  res.sendFile(path.join(__dirname, 'login.html'));
});

app.get('/request-user.html', (req, res) => {
  res.sendFile(path.join(__dirname, 'request-user.html'));
});

app.get('/forgot-password.html', (req, res) => {
  res.sendFile(path.join(__dirname, 'forgot-password.html'));
});

app.get('/error.html', (req, res) => {
  res.sendFile(path.join(__dirname, 'error.html'));
});

//mail
async function sendGraphMail(toEmail, subject, htmlBody) {
  const accessToken = await getAccessToken();

  const res = await fetch('https://graph.microsoft.com/v1.0/users/MuninderPal@JK17.onmicrosoft.com/sendMail', {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      message: {
        subject: subject,
        body: {
          contentType: "HTML",
          content: htmlBody,
        },
        toRecipients: [
          {
            emailAddress: {
              address: toEmail,
            },
          },
        ],
      },
      saveToSentItems: true
    }),
  });

  if (!res.ok) {
    const errorText = await res.text();
    console.error("❌ Graph Mail Error:", errorText);
    throw new Error("Graph mail send failed");
  }

  console.log("✅ OTP Email sent to:", toEmail);
}


//email part
app.post('/send-otp', async (req, res) => {
  const { name, email, contact } = req.body;

  if (!name || !email || !contact) {
    return res.status(400).send("All fields are required");
  }

  const otp = Math.floor(100000 + Math.random() * 900000).toString();
  otpStore.set(email, { otp, expiresAt: Date.now() + 10 * 60 * 1000 }); // 10 min

  const html = `
    <h3>New CRM Access Request</h3>
    <p><strong>Name:</strong> ${name}</p>
    <p><strong>Email:</strong> ${email}</p>
    <p><strong>Contact:</strong> ${contact}</p>
    <p><strong>OTP:</strong> <b>${otp}</b></p>
  `;

  try {
    await sendGraphMail('MuninderPal@JK17.onmicrosoft.com', 'New User Request OTP', html);
    res.send("OTP sent to admin. Please ask admin for it.");
  } catch (err) {
    console.error("❌ Error sending OTP:", err.message);
    res.status(500).send("Failed to send OTP");
  }
});

app.post('/verify-otp', (req, res) => {
  const { email, otp } = req.body;
  const stored = otpStore.get(email);
  
  if (!stored) {
    return res.status(400).send("OTP expired or not found.");
  }

  if (stored.otp === otp) {
    otpStore.delete(email);
    requestOtpAttempts.delete(email);
    return res.send("✅ OTP verified. Please set username and password.");
  } else {
    let attempts = requestOtpAttempts.get(email) || 0;
    attempts += 1;
    requestOtpAttempts.set(email, attempts);

    if (attempts >= 3) {
      otpStore.delete(email);
      requestOtpAttempts.delete(email);
      return res.status(400).send("❌ OTP expired after 3 failed attempts.");
    }

    return res.status(400).send(`Incorrect OTP. Attempt ${attempts}/3.`);
  }
});







async function getWorkbookFromOneDrive(fileName) {
  const token = await getAccessToken();

  try {
    // Try to get metadata and file ID
    const meta = await axios.get(
      `https://graph.microsoft.com/v1.0/users/MuninderPal@JK17.onmicrosoft.com/drive/root:/${fileName}`,
      {
        headers: { Authorization: `Bearer ${token}` }
      }
    );

    const fileId = meta.data.id;

    const response = await axios.get(
      `https://graph.microsoft.com/v1.0/users/MuninderPal@JK17.onmicrosoft.com/drive/items/${fileId}/content`,
      {
        headers: { Authorization: `Bearer ${token}` },
        responseType: 'arraybuffer'
      }
    );

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(response.data);
    return { workbook, fileId, token };
  } catch (err) {
    console.warn(`⚠️ Workbook "${fileName}" not found in OneDrive.`);
    return { workbook: null, fileId: null, token };
  }
}








//forget pass
app.post('/forgot-password/send-otp', async (req, res) => {
  const { username } = req.body;
  const token = await getAccessToken();

  try {
    // Load Excel file
    const file = await axios.get(
      'https://graph.microsoft.com/v1.0/users/muninderpal@jk17.onmicrosoft.com/drive/root:/credentials.xlsx:/content',
      {
        headers: { Authorization: `Bearer ${token}` },
        responseType: 'arraybuffer'
      }
    );
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(file.data);
    const sheet = workbook.getWorksheet('Users') || workbook.worksheets[0];

    let userExists = false;
    sheet.eachRow((row, i) => {
      if (i === 1) return;
      const name = row.getCell(1).value?.toString().trim().toLowerCase();
      if (name === username.toLowerCase()) {
        userExists = true;
      }
    });

    if (!userExists) return res.status(404).send('User not found');

    // Generate OTP
    const otp = Math.floor(100000 + Math.random() * 900000).toString();
    otpStore.set(username.toLowerCase(), otp);

    // Compose HTML email
    const html = `
      <h3>Password Reset Request</h3>
      <p><strong>Username:</strong> ${username}</p>
      <p><strong>OTP:</strong> <b>${otp}</b></p>
    `;

    // Send Email
    await sendGraphMail('MuninderPal@JK17.onmicrosoft.com', 'JK SOLAR CRM - Password Reset OTP', html);

    res.send("✅ OTP sent to admin.");
  } catch (err) {
    console.error('OTP error:', err.message);
    res.status(500).send('OTP send failed');
  }
});





// Keep track of attempts
const otpAttempts = new Map();

app.post('/forgot-password/verify-otp', (req, res) => {
  const { username, otp } = req.body;
  const userKey = username.toLowerCase();
  const storedOtp = otpStore.get(userKey);

  if (!storedOtp) {
    return res.status(400).send("OTP expired or not found.");
  }

  // Check OTP
  if (storedOtp === otp) {
    otpStore.delete(userKey);
    otpAttempts.delete(userKey); // reset attempts
    return res.send("✅ OTP verified. Set new password.");
  } else {
    let attempts = otpAttempts.get(userKey) || 0;
    attempts += 1;
    otpAttempts.set(userKey, attempts);

    if (attempts >= 3) {
      otpStore.delete(userKey);
      otpAttempts.delete(userKey);
      return res.status(400).send("❌ OTP expired after 3 failed attempts.");
    }

    return res.status(400).send(`Incorrect OTP. Attempt ${attempts}/3.`);
  }
});


app.post('/forgot-password/reset', async (req, res) => {
  const { username, newPassword } = req.body;
  const token = await getAccessToken();

  try {
    const file = await axios.get(
      'https://graph.microsoft.com/v1.0/users/muninderpal@jk17.onmicrosoft.com/drive/root:/credentials.xlsx:/content',
      {
        headers: { Authorization: `Bearer ${token}` },
        responseType: 'arraybuffer',
      }
    );

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(file.data);
    const sheet = workbook.getWorksheet('Users') || workbook.worksheets[0];

    let updated = false;

    sheet.eachRow((row, i) => {
      if (i === 1) return;
      const user = row.getCell(1).value?.toString().trim();
      if (user.toLowerCase() === username.toLowerCase()) {
        row.getCell(2).value = newPassword;
        updated = true;
      }
    });

    if (!updated) return res.status(404).send("User not found");

    const buffer = await workbook.xlsx.writeBuffer();
    await axios.put(
      'https://graph.microsoft.com/v1.0/users/muninderpal@jk17.onmicrosoft.com/drive/root:/credentials.xlsx:/content',
      buffer,
      {
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        },
      }
    );

    otpStore.delete(username.toLowerCase());
    res.send("✅ Password updated successfully");
  } catch (err) {
    console.error("Reset error:", err.message);
    res.status(500).send("Could not reset password");
  }
});







app.post('/request-access', async (req, res) => {
  const { name, email, contact } = req.body;
  if (!name || !email || !contact) {
    return res.status(400).send('Missing fields');
  }

  const token = crypto.randomBytes(16).toString('hex');
  global.pendingUsers = global.pendingUsers || {};
  global.pendingUsers[token] = { name, email, contact };

  const approvalLink = `http://192.168.1.12:3000/grant-access.html?token=${token}`;
 const denialLink = `http://192.168.1.12:3000/deny-access/${token}`;

  const htmlContent = `
    <h3>New CRM User Request</h3>
    <p><strong>Name:</strong> ${name}<br/>
       <strong>Email:</strong> ${email}<br/>
       <strong>Contact:</strong> ${contact}</p>

    <p><strong>Action:</strong></p>

    <a href="${approvalLink}" 
       style="padding:10px 18px; background:#4caf50; color:#fff; text-decoration:none; border-radius:6px;">
       ✅ Grant Access
    </a><br/><br/>

    <a href="${denialLink}" 
       style="padding:10px 18px; background:#f44336; color:#fff; text-decoration:none; border-radius:6px;">
       ❌ Deny Access
    </a>
  `;

  await sendGraphMail(
    'muninderpal@jk17.onmicrosoft.com',
    '📥 CRM Access Request',
    htmlContent
  );

  res.send('Access request sent successfully!');
});



// ✅ Route: Admin GETs the grant access form with checkboxes (full access for now)
// ✅ Show Grant Access HTML form
app.get('/grant-access/:token', (req, res) => {
  const { token } = req.params;
  const user = global.pendingUsers?.[token];
  if (!user) return res.status(404).send('⚠️ Invalid or expired link.');

  res.send(`
  <!DOCTYPE html>
  <html><head><title>Grant Access</title></head>
  <body>
    <h2>Select Modules</h2>
    <form method="POST" action="/save-access">
      <input type="hidden" name="token" value="${token}">
      <label><input type="checkbox" name="modules" value="addclient"> Add Client</label><br>
      <label><input type="checkbox" name="modules" value="searchclient"> Search Client</label><br>
      <label><input type="checkbox" name="modules" value="stocksheet"> Stock Sheet</label><br>
      <label><input type="checkbox" name="modules" value="leads"> Leads</label><br>
      <label><input type="checkbox" name="modules" value="proposal"> Proposal</label><br>
      <label><input type="checkbox" name="modules" value="dashboard"> Dashboard</label><br><br>
      <button type="submit">Submit Access</button>
    </form>
  </body>
  </html>
`);
});

// ✅ Save access and notify user to create credentials
app.post('/save-access', async (req, res) => {
  const { token, modules } = req.body;
  const user = global.pendingUsers?.[token];
  if (!user) return res.status(400).send('⚠️ Token expired or invalid');
  delete global.pendingUsers[token];

  const fileName = 'credentials.xlsx';
  const { workbook, token: driveToken } = await getWorkbookFromOneDrive(fileName);
  const sheet = workbook.worksheets[0];

  const accessList = Array.isArray(modules) ? modules.join(',') : modules;
  sheet.addRow([
    'not-set',        // Username
    'not-set',        // Password
    accessList,       // Role (modules granted)
    'Pending',        // Status
    user.email,
    user.contact
  ]);

  const buffer = await workbook.xlsx.writeBuffer();
await uploadWorkbookToOneDrive(fileName, buffer, driveToken);

async function uploadWorkbookToOneDrive(fileName, buffer, token) {
  const uploadUrl = `https://graph.microsoft.com/v1.0/users/muninderpal@jk17.onmicrosoft.com/drive/root:/${fileName}:/content`;

  await axios.put(uploadUrl, buffer, {
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }
  });

  console.log(`✅ Uploaded ${fileName} to OneDrive.`);
}

  const html = `
    <p>✅ Your access to JK CRM has been approved!</p>
    <p>You may now create your login credentials:</p>
    <a href="http://192.168.1.12:3000/create-user.html?email=${encodeURIComponent(user.email)}&token=${token}" 
       style="padding:10px 18px; background:#4caf50; color:#fff; text-decoration:none; border-radius:6px;">
       Set Username & Password
    </a>`;

  await sendGraphMail(user.email, 'Access Granted – Set CRM Login', html);

  res.send('✅ Access granted, User has been notified to set credentials.');

});



app.get('/deny-access/:token', (req, res) => {
  const { token } = req.params;
  if (global.pendingUsers) delete global.pendingUsers[token];
  res.send("❌ Access denied. The request has been discarded.");
});



app.get('/grant-access.html', (req, res) => {
  res.sendFile(path.join(__dirname, 'grant-access.html'));
}); 

// Serve the create-user form
app.get('/create-user.html', (req, res) => {
  res.sendFile(path.join(__dirname, 'create-user.html'));
});

// Handle create-user form submission
app.post('/create-user', async (req, res) => {
  const { email, contact, username, password } = req.body;

  if (!email || !contact || !username || !password) {
    return res.status(400).send('⚠️ Missing fields');
  }

  try {
    const fileName = 'credentials.xlsx';
    const { workbook, token } = await getWorkbookFromOneDrive(fileName);
    const sheet = workbook.getWorksheet('Credentials') || workbook.worksheets[0];

    let found = false;

    sheet.eachRow((row, rowNumber) => {
      const sheetEmail = (row.getCell(5).value || '').toString().trim().toLowerCase();   // Column E
      const sheetContact = (row.getCell(6).value || '').toString().trim();               // Column F

      if (
        sheetEmail === email.trim().toLowerCase() &&
        sheetContact === contact.trim()
      ) {
        console.log(`✅ Match at row ${rowNumber}`);
        row.getCell(1).value = username;  // Column A
        row.getCell(2).value = password;  // Column B
        row.getCell(4).value = 'Active';  // Column D
        found = true;
      }
    });

    if (!found) {
      console.log('❌ No matching row found.');
      return res.status(404).send('❌ Email and contact not found in same row.');
    }

    // Save to OneDrive
    const buffer = await workbook.xlsx.writeBuffer();
    await axios.put(
      `https://graph.microsoft.com/v1.0/users/muninderpal@jk17.onmicrosoft.com/drive/root:/${fileName}:/content`,
      buffer,
      {
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        }
      }
    );

    res.send('✅ Your CRM account has been created. You can now log in!');
  } catch (err) {
    console.error('❌ Error during user creation:', err);
    res.status(500).send('Server error during user creation.');
  }
});




app.get('/get-modules', async (req, res) => {
  const { username } = req.query;
  if (!username) return res.json({ success: false, error: "Username missing" });

  try {
    const { workbook } = await getWorkbookFromOneDrive('credentials.xlsx');
    const sheet = workbook.getWorksheet('Sheet1');
    let modules = null;

    sheet.eachRow((row, rowNumber) => {
      const userCell = row.getCell(4).value; // Username is in column 4
      if (userCell && userCell.toString().trim().toLowerCase() === username.trim().toLowerCase()) {
        const roleCell = row.getCell(3).value; // Modules are in column 3 (Role)
        if (roleCell) {
          modules = roleCell.toString().split(',').map(m => m.trim());
        }
      }
    });

    if (modules) {
      res.json({ success: true, modules });
    } else {
      res.json({ success: false, error: "User not found" });
    }
  } catch (err) {
    console.error("Module fetch error:", err);
    res.status(500).json({ success: false, error: "Server error" });
  }
});

app.get('/api/user-modules', (req, res) => {
  if (!req.session.user) {
    console.log("⚠️ No user in session");
    return res.status(401).json({ error: 'Not logged in' });
  }

  res.json({ modules: req.session.user.role });
});

app.get('/api/check-proposal-match', async (req, res) => {
  try {
    const { mobile, kw } = req.query;
    const cleanMobile = mobile.toString().replace(/\D/g, '').slice(-10);
    const clientKW = kw.toString().trim();

    const { workbook: proposalBook } = await getWorkbookFromOneDrive('proposal.xlsx');
    const proposalSheet = proposalBook.getWorksheet('Proposals');

    let matchFound = false;

    proposalSheet.eachRow((row, index) => {
      if (index === 1) return; // skip header
      const ref = (row.getCell(1).value || '').toString().trim(); // Ref: JK-10-9876543210
      const match = ref.match(/^JK-(\d+)-(\d{10})$/);
      if (!match) return;

      const refKW = match[1];
      const refMobile = match[2];

      if (refKW === clientKW && refMobile === cleanMobile) {
        matchFound = true;
      }
    });

    if (!matchFound) {
      return res.json({ match: false, reason: 'not_in_proposal' });
    }

    // ✅ Check TempData for duplicate mobile
    const { workbook: tempBook } = await getWorkbookFromOneDrive('TempData.xlsx');
    const tempSheet = tempBook.getWorksheet('Client Data');

    let duplicate = false;
    tempSheet.eachRow((row, idx) => {
      const tempMobile = (row.getCell(10).value || '').toString().replace(/\D/g, '').slice(-10);
      if (tempMobile === cleanMobile) {
        duplicate = true;
      }
    });

    if (duplicate) {
      return res.json({ match: false, reason: 'already_exists' });
    }

    res.json({ match: true });

  } catch (err) {
    console.error('❌ Error in check-proposal-match:', err.message);
    res.status(500).json({ error: 'Internal Server Error' });
  }
});





//netmeteringpdf
async function generateNetMeteringPDF(clientName, data) {
  const htmlPath = path.join(__dirname, 'net-metering-agreement.html');
  const html = await fs.readFile(htmlPath, 'utf8');

  const today = new Date();
  const formatted = {
    day: today.getDate(),
    month: today.toLocaleString('default', { month: 'long' }),
    year: today.getFullYear()
  };

  // 🔍 Fetch S/o from TempData.xlsx if not already present in data
  if (!data.so) {
    try {
      const { workbook: tempWorkbook } = await getWorkbookFromOneDrive('TempData.xlsx');
      const tempSheet = tempWorkbook.getWorksheet('Client Data');

      tempSheet.eachRow((row, i) => {
        if (i === 1) return; // Skip header
        const tempName = (row.getCell(2).value || '').toString().trim().toLowerCase();
        const tempMobile = (row.getCell(10).value || '').toString().trim();
        const tempKW = (row.getCell(12).value || '').toString().trim();

        if (
          tempName === data.name.toLowerCase().trim() &&
          tempMobile === data.mobile &&
          tempKW === data.kw
        ) {
          data.so = row.getCell(36).value || ''; // Assuming S/o is in column G
        }
      });
    } catch (err) {
      console.error('⚠️ Failed to fetch S/o from TempData.xlsx:', err.message);
    }
  }

  let filled = html
    .replace(/{{date}}/g, `${formatted.day}-${formatted.month}-${formatted.year}`)
    .replace(/{{toWhom}}/g, data.towhom || '')
    .replace(/{{so}}/g, data.so || '')
    .replace(/{{mobile}}/g, data.mobile || '')
    .replace(/{{kw}}/g, data.kw || '')
    .replace(/{{address}}/g, data.address || '')
    .replace(/{{city}}/g, data.city || '')
    .replace(/{{state}}/g, data.state || '');

  const savePath = path.join(__dirname, 'uploads', clientName, 'Net Meter Agreement.pdf');

  const browser = await puppeteer.launch({ headless: 'new' });
  const page = await browser.newPage();
  await page.setContent(filled, { waitUntil: 'networkidle0' });
  await page.pdf({ path: savePath, format: 'A4', printBackground: true });
  await browser.close();
}




//dhvn list page

app.get('/api/dhbvn-list', async (req, res) => {
  try {
    const { workbook } = await getWorkbookFromOneDrive('TempData.xlsx');
    const sheet = workbook.getWorksheet('Client Data');
    let data = [];

    sheet.eachRow((row, idx) => {
      if (idx === 1) return; // skip header

      const name = row.getCell(2).value?.toString().trim() || '';
      const mobile = row.getCell(10).value?.toString().trim() || '';
      const kw = row.getCell(12).value?.toString().trim() || '';

      const hno = row.getCell(4).value?.toString().trim() || '';
      const loc = row.getCell(5).value?.toString().trim() || '';
      const landmark = row.getCell(6).value?.toString().trim() || '';
      const dist = row.getCell(7).value?.toString().trim() || '';
      const state = row.getCell(8).value?.toString().trim() || '';
      const pin = row.getCell(9).value?.toString().trim() || '';

      const address = `${hno}, ${loc}, ${landmark}, ${dist}, ${state} - ${pin}`;

      const ba = (row.getCell(53).value || '').toString().trim();

      if (ba === '') {
        data.push({ name, kw, mobile, address });
      }
    });

    res.json(data);
  } catch (err) {
    console.error('❌ Error in DHBVN list API:', err);
    res.status(500).send('Server error');
  }
});

async function uploadBufferToOneDrive(fileName, buffer, token) {
  const uploadUrl = `https://graph.microsoft.com/v1.0/users/muninderpal@jk17.onmicrosoft.com/drive/root:/${fileName}:/content`;

  await axios.put(uploadUrl, buffer, {
    headers: {
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Authorization': `Bearer ${token}`
    }
  });
}



app.post('/submit-dhbvn-form', async (req, res) => {
  const data = req.body;
  try {
    const { workbook, token } = await getWorkbookFromOneDrive('TempData.xlsx');
    if (!workbook) return res.status(500).send('❌ Could not load TempData.xlsx');

    const sheet = workbook.getWorksheet('Client Data');
    if (!sheet) return res.status(500).send('❌ Sheet not found');

    const inputKW = (data.kw || '').toString().trim().replace(/\.0$/, '');
    const inputMobile = (data.mobile || '').toString().trim().replace(/\D/g, '').slice(-10);

    let updated = false;

sheet.eachRow((row, rowIndex) => {
  const rowKW = (row.getCell(12).value || '').toString().trim().replace(/\.0$/, '');
  const rowMobile = (row.getCell(10).value || '').toString().trim().replace(/\D/g, '').slice(-10);

  const inputKW = (data.kw || '').toString().trim().replace(/\.0$/, '');
  const inputMobile = (data.mobile || '').toString().trim().replace(/\D/g, '').slice(-10);

  if (rowKW === inputKW && rowMobile === inputMobile) {
    const startCol = 43; // AQ = col 43
    const values = [
      String(data.circle || '').trim(),
      String(data.subdivision || '').trim(),
      String(data.category || '').trim(),
      String(data.firmType || '').trim(),
      String(data.identityProof || '').trim(),
      String(data.aadharNo || '').trim(),
      String(data.accountNo || '').trim(),
      String(data.proposedCapacity || '').trim(),
      String(data.voltageSupply || '').trim(),
      String(data.voltageOutput || '').trim(),
      new Date().toLocaleDateString()
    ];

    values.forEach((val, i) => {
      const cell = row.getCell(startCol + i);
      cell.value = val;
    });

    row.commit();  // 🔐 Necessary to save changes
    updated = true;
  }
});


    if (!updated) {
      return res.status(404).send('❌ KW and Mobile not found in any row.');
    }

    await saveWorkbookToOneDrive('TempData.xlsx', workbook, token);
    res.send('✅ DHBVN form saved and updated!');
  } catch (err) {
    console.error('❌ DHBVN Save Error:', err);
    res.status(500).send('❌ DHBVN Save Error: ' + err.message);
  }
});







//login dhbvn

app.post('/login-dhbvn', async (req, res) => {
  
console.log('📥 Received /login-dhbvn POST request');

  res.setHeader('Content-Type', 'text/plain');
  res.setHeader('Transfer-Encoding', 'chunked');

  const writeLog = msg => res.write(`${msg}\n`);

  const { kw, mobile } = req.body || {};
  if (!kw || !mobile) {
    writeLog("❌ Missing KW or Mobile in request.");
    return res.end();
  }

  writeLog(`🔍 Received KW: ${kw}, Mobile: ${mobile}`);

  try {
    writeLog("🚀 Launching browser...");
    const browser = await puppeteer.launch({
      headless: false,  // ✅ Streaming-compatible headless mode
      slowMo: 100,        // 👁️ Visible delay between steps
      args: ['--start-maximized']
    });

    const pages = await browser.pages();
    const page = pages[0];             // ✅ Use the first opened tab
    await page.bringToFront();         // ✅ Ensure it's visible in taskbar

   

   writeLog("🌐 Opening DHBVN login page...");
await page.goto('https://esolarconn.dhbvn.org.in/userDashboard/Home.aspx', {
  waitUntil: 'domcontentloaded'
});

await new Promise(res => setTimeout(res, 2000));

writeLog("⌨️ Typing username & password...");
await page.type('#txtUserName', 'JKSOLAR1');
await page.type('#txtPassword', '*Munyz001');

await new Promise(res => setTimeout(res, 1000));

writeLog("🔐 Clicking login...");
await Promise.all([
  page.waitForNavigation({ waitUntil: 'domcontentloaded' }),
  page.evaluate(() => {
    const btn = [...document.querySelectorAll('button,input[type="submit"]')]
      .find(b => b.innerText?.trim() === 'Login' || b.value?.trim() === 'Login');
    btn?.click();
  })
]);

writeLog("✅ Logged in. Opening new tab for NetMetering form...");

// ✅ SESSION IS NOW ACTIVE - we can safely open new tab
const newPage = await browser.newPage();
await newPage.goto('https://esolarconn.dhbvn.org.in/userDashboard/NetMeteringOnlineApplyNewCon.aspx#no-back-button', {
  waitUntil: 'domcontentloaded'
});

await new Promise(res => setTimeout(res, 2000));

writeLog("📥 Selecting 'New connection' from dropdown...");

// open dropdown (visual)
const box = await newPage.evaluate(() => {
  const el = document.querySelector('#ddlNewConnReConn');
  const r = el.getBoundingClientRect();
  return { x: r.left + 10, y: r.top + 10 };
});
await newPage.mouse.click(box.x, box.y);
await new Promise(res => setTimeout(res, 500));

// now select "New connection" option
await newPage.evaluate(() => {
  const dropdown = document.querySelector('#ddlNewConnReConn');
  dropdown.value = '1'; // value for "New connection"
  dropdown.dispatchEvent(new Event('change', { bubbles: true }));
});

await new Promise(res => setTimeout(res, 1500));
writeLog("✅ 'New connection' selected successfully.");




const { kw, mobile } = req.body;
const { workbook } = await getWorkbookFromOneDrive('TempData.xlsx');
const sheet = workbook.getWorksheet('Client Data');

let data = null;

sheet.eachRow((row, i) => {
  const rowKW = (row.getCell(12).value || '').toString().trim();
  const rowMobile = (row.getCell(10).value || '').toString().trim();
  if (rowKW === kw && rowMobile === mobile) {
    data = {
      hno: (row.getCell(4)?.value || '').toString().trim(),       // D
  address: (row.getCell(5)?.value || '').toString().trim(),   // E
  landmark: (row.getCell(6)?.value || '').toString().trim(),  // F
  district: (row.getCell(7)?.value || '').toString().trim(),  // G
  pincode: (row.getCell(9)?.value || '').toString().trim(),   // I
  email: (row.getCell(11)?.value || '').toString().trim(),     // H
      fatherName: (row.getCell(3)?.value || '').toString(),
      circle: (row.getCell(43)?.value || '').toString(),
      subdivision: (row.getCell(44)?.value || '').toString(),
      category: (row.getCell(45)?.value || '').toString(),
      applicantType: (row.getCell(46)?.value || '').toString(),
      identityProof: (row.getCell(47)?.value || '').toString(),
      proposedCapacity: (row.getCell(50)?.value || '').toString(),
      aadharNo: (row.getCell(48)?.value || '').toString(),
      accountNo: (row.getCell(49)?.value || '').toString(),
      voltageSupply: (row.getCell(51)?.value || '').toString(),
      voltageOutput: (row.getCell(52)?.value || '').toString(),
      aadharPath: (row.getCell(15)?.value || '').toString().trim()  // Column O
    };
  }
});

if (!data) {
  writeLog("❌ Could not find matching KW & Mobile in TempData.");
  return res.end("❌ KW & Mobile not found in Excel");
}






// 🧾 Fill CIRCLE
await newPage.select('#ddlCircle', data.circle);
writeLog(`✅ Circle selected: ${data.circle}`);

// 🧾 Fill SUBDIVISION
await newPage.select('#ddlSubDivision', data.subdivision);
writeLog(`✅ Subdivision selected: ${data.subdivision}`);

// 🧾 Fill CATEGORY
await newPage.select('#ddlCategoryOfConnection', data.category);
writeLog(`✅ Category selected: ${data.category}`);

// ✅ Applicant Type Selection Logic
const applicantValue = (() => {
  const val = (data.applicantType || '').toLowerCase();
  if (val.includes('individual')) return 'Individual';
  if (val.includes('gov')) return 'Govt Department or Institution';
  if (val.includes('institution') || val.includes('private')) return 'Private Institution or establishment';
  return '0'; // fallback to ---Select---
})();

await newPage.select('#ddlProfession', applicantValue);
writeLog(`✅ Applicant Type selected: ${applicantValue}`);


// 🧾 Fill FIXED "M/s" for firm
await newPage.select('#ddlFirstName', '4');
writeLog(`✅ Firm Prefix set to M/S`);




// 🧾 Fill FATHER/HUSBAND NAME
await newPage.type('#inpFathersName', String(data.fatherName || ''));
writeLog(`✅ Father/Husband Name: ${data.fatherName || '⚠️ Empty'}`);


// 🧾 Fill ID PROOF
await newPage.type('#inpIdentityName', String(data.identityProof || ''));
writeLog(`✅ Identity Proof: ${data.identityProof || '⚠️ Empty'}`);

// 🧾 Fill ACCOUNT NO
await newPage.type('#txtAccountNumber', String(data.accountNo || ''));
writeLog(`✅ Account Number: ${data.accountNo || '⚠️ Empty'}`);

// 🧾 Fill CONNECTED LOAD (KW)
await newPage.type('#txtConnectionLoad', String(data.proposedCapacity || ''));
writeLog(`✅ Connected Load (KW): ${data.proposedCapacity || '⚠️ Empty'}`);

// 🧾 Fill SITE ADDRESS FIELDS
await newPage.type('#txtFullAddress', String(data.hno || ''));
await newPage.type('#txtStreetName', String(data.address || ''));
await newPage.type('#txtConnLandMark', String(data.landmark || ''));
await newPage.type('#txtCity', String(data.district || ''));
await newPage.type('#txtPinCode', String(data.pincode || ''));
writeLog(`✅ Site Address filled.`);

// 🧾 Fill EMAIL & MOBILE// 🧾 Clear and fill Email field
await newPage.click('#inpEmailID'); // Focus the field
await newPage.keyboard.down('Control');
await newPage.keyboard.press('A'); // Select all
await newPage.keyboard.up('Control');
await newPage.keyboard.press('Backspace'); // Clear
await newPage.type('#inpEmailID', String(data.email || ''));
writeLog(`✅ Email filled: ${data.email || '⚠️ Empty'}`);


// 🧾 Fill AADHAR NUMBER
await newPage.type('#txtAddressAdharNo', String(data.aadharNo || ''));
writeLog(`✅ Aadhar Number: ${data.aadharNo || '⚠️ Empty'}`);

// 🧾 Voltage at which Supply has been given
await newPage.evaluate((val) => {
  const select = document.querySelector('#ddlVoltageSupply');
  if (!select) return;

  const index = (val === '1') ? 1 : (val === '2') ? 2 : 0;
  select.selectedIndex = index;
  select.dispatchEvent(new Event('change'));
}, String(data.voltageSupply || '0'));
writeLog(`✅ Voltage Supply selected index: ${data.voltageSupply || '0'}`);

// 🧾 Voltage at Output of Solar Inverter
await newPage.evaluate((val) => {
  const select = document.querySelector('#ddlVoltageOutput');
  if (!select) return;

  const index = (val === '1') ? 1 : (val === '2') ? 2 : 0;
  select.selectedIndex = index;
  select.dispatchEvent(new Event('change'));
}, String(data.voltageOutput || '0'));
writeLog(`✅ Voltage Output selected index: ${data.voltageOutput || '0'}`);


// 🧾 Fill PROPOSED CAPACITY
await newPage.type('#txtAcCapacity', String(data.proposedCapacity || ''));
writeLog(`✅ Proposed Capacity: ${data.proposedCapacity || '⚠️ Empty'}`);

// 🧾 Select Vendor: "JK SOLAR PRIVATE LIMITED"
await newPage.evaluate(() => {
  const select = document.querySelector('#ddlVendor');
  if (!select) return;

  const options = [...select.options];
  const match = options.find(opt =>
    opt.textContent.trim().toLowerCase() === 'jk solar private limited'
  );

  if (match) {
    select.value = match.value;
    select.dispatchEvent(new Event('change')); // Trigger onchange if needed
  }
});
writeLog("✅ Vendor selected: JK SOLAR PRIVATE LIMITED");






res.end();
  } catch (err) {
    writeLog("❌ Error: " + err.message);
    res.end();
  }
});
async function downloadFromOneDrive(filePath, user = 'muninderpal@jk17.onmicrosoft.com') {
  const accessToken = await getAccessToken();
  const url = `https://graph.microsoft.com/v1.0/users/${user}/drive/root:/${filePath}:/content`;
  const response = await axios.get(url, {
    headers: { Authorization: `Bearer ${accessToken}` },
    responseType: 'arraybuffer'
  });
  return response.data;
}


  
// 🚀 Start server
app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
});
