// 🚀 Basic Setup (OneDrive-Ready)
const express = require('express');
const axios = require('axios');
const qs = require('qs');
const path = require('path');
const ExcelJS = require('exceljs');
require('dotenv').config();

const app = express();
const port = 3000;
const fileUpload = require('express-fileupload');
app.use(fileUpload());


const puppeteer = require('puppeteer');

app.use(express.static('public')); // serves HTML/CSS/JS
app.use(express.json());
app.use(express.urlencoded({ extended: true }));


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
  purchaseAgreement: 'purchaseagreement',
  netMeteringAgreement: 'netmeteringagreement',
  clientPhoto: 'clientphoto'
};



app.post('/submit-client', async (req, res) => {
  try {
    const data = req.body;
    const files = req.files;
    const clientName = data.name.trim().toLowerCase().replace(/\s+/g, '_');

    // 🖼️ Upload client photo (base64)
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

    // 📤 Upload document files (Aadhar, PAN, etc.) — parallel with logs
    const uploadTasks = Object.entries(files).map(async ([field, fileObj]) => {
      const file = fileObj[0] || fileObj;
      const ext = path.extname(file.name);
      const customName = `${field.toLowerCase()}${ext}`;

      await uploadToOneDriveFolder(clientName, field, file.data, customName);

      const pretty = field.replace(/([A-Z])/g, ' $1').replace(/^./, s => s.toUpperCase());
      console.log(`📤 ${pretty} uploaded...`);
    });

    await Promise.all(uploadTasks);

    // 📄 Prepare client data with OneDrive paths
    const client = {
      date: data.date,
      name: data.name,
      address: data.address,
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
      cancelCheque: `uploads/${clientName}/cancelcheque.png`,
      purchaseAgreement: `uploads/${clientName}/purchaseagreement.png`,
      netMeteringAgreement: `uploads/${clientName}/netmeteringagreement.png`
    };

    // 📊 Save client data to OneDrive Excel
    const { workbook, token } = await getWorkbookFromOneDrive('TempData.xlsx');
    const sheet = workbook.getWorksheet('Client Data') || workbook.addWorksheet('Client Data');

    const newRow = sheet.addRow([
      client.date, client.name, client.address, client.mobile, client.email, client.kw,
      client.advance, client.totalCost, client.aadharFront, client.aadharBack, client.panCard, client.bill,
      client.ownershipProof, client.cancelCheque, client.purchaseAgreement, client.netMeteringAgreement
    ]);

    newRow.hidden = false;

    await uploadWorkbookToOneDrive('TempData.xlsx', workbook, token);
    console.log('✅ Client data saved to OneDrive Excel.');

    res.send(`
  <html>
    <head>
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <style>
        body {
          font-family: 'Segoe UI', sans-serif;
          margin: 0;
          padding: 0;
          display: flex;
          justify-content: center;
          align-items: center;
          height: 80vh;
          background-color: #f2f9ff;
        }
        .success {
          font-size: 24px;
          color: #2d7a2d;
          background: #e8fce8;
          padding: 20px 30px;
          border-radius: 12px;
          box-shadow: 0 4px 15px rgba(0, 0, 0, 0.1);
          text-align: center;
        }
        @media (max-width: 768px) {
          .success {
            font-size: 28px;
          }
        }
      </style>
    </head>
    <body>
      <div class="success">✅ Client submitted successfully and all files saved .</div>
    </body>
  </html>
`);

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
  const clientName = req.params.clientName.trim().toLowerCase();
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
    console.log(`📂 Checking files for: ${clientName}`);

    let clientInfo = null;
    let fileStatus = [];

    const { workbook } = await getWorkbookFromOneDrive('TempData.xlsx');
    const worksheet = workbook.getWorksheet('Client Data');

    worksheet.eachRow((row, rowIndex) => {
      const rowClientName = (row.getCell(2).value || '').toString().trim().toLowerCase();

      console.log(`🔍 Excel: '${rowClientName}' vs Search: '${clientName}'`);

      if (rowClientName === clientName) {
        clientInfo = {
          name: row.getCell(2).value || '',
          address: row.getCell(3).value || '',
          mobile: row.getCell(4).value || '',
          email: row.getCell(5).value || '',
          kw: row.getCell(6).value || ''
        };

        // Loop over columns I to P (9 to 16)
        fileFields.forEach((field, index) => {
          const pathInExcel = row.getCell(9 + index).value?.toString().trim() || '';
          const label = field.replace(/([A-Z])/g, ' $1').replace(/^./, s => s.toUpperCase());

          fileStatus.push({
            file: field,
            label: label,
            exists: !!pathInExcel  // We assume true if there's a path (skip fs.existsSync)
          });
        });
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

    const fileList = [
      'uploads/demo/aadharfront.png',
      'uploads/demo/pancard.png',
      'uploads/demo/bill.png'
    ];

    const results = [];

    for (const path of fileList) {
      const encodedPath = encodeURIComponent(path);
      const url = `https://graph.microsoft.com/v1.0/users/muninderpal@jk17.onmicrosoft.com/drive/root:/${encodedPath}`;
      try {
        await axios.get(url, { headers: { Authorization: `Bearer ${token}` } });
        results.push({ file: path, exists: true });
      } catch (err) {
        results.push({ file: path, exists: false });
      }
    }

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
      const clientNameFromRow = rowValues[2]; // Column B = Name
      const event = rowValues[3];             // Column C = Event

      if (
        clientNameFromRow &&
        clientNameFromRow.toLowerCase().trim() === clientName &&
        event && event.trim() !== ''
      ) {
        clientTimeline.push({
          event: event,
          eventDate: rowValues[4],         // D
          eventDescription: rowValues[5],  // E
          status: rowValues[6]             // F
        });
      }
    });

    if (clientTimeline.length === 0) {
      return res.status(404).json({ error: 'No timeline events found for this client' });
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
      const rowClientName = row.getCell(2).value?.toString().trim().toLowerCase(); // Column B

      if (rowClientName === clientName) {
        clientInfo = {
          name: row.getCell(2).value || '',     // B
          address: row.getCell(3).value || '',  // C
          mobile: row.getCell(4).value || '',   // D
          email: row.getCell(5).value || '',    // E
          kw: row.getCell(6).value || ''        // F
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

  if (!clientName) return res.status(400).json({ error: 'Client name missing' });

  try {
    const { workbook, token } = await getWorkbookFromOneDrive('TempData.xlsx');
    const sheet = workbook.getWorksheet('Client Data');

    if (!sheet) return res.status(404).json({ error: 'Client Data sheet not found' });

    let found = false;

    sheet.eachRow((row) => {
      const nameCell = row.getCell(2).value?.toString().toLowerCase().trim(); // Column B

      if (nameCell === clientName.toLowerCase().trim()) {
        row.getCell(17).value = appliedKW;                             // Column Q
        row.getCell(18).value = appliedOnPMSurya;                      // Column R
        row.getCell(19).value = applicationDHBVN;                      // Column S
        row.getCell(20).value = loadChangeReductionNewConnection;     // Column T
        found = true;
      }
    });

    if (!found) return res.status(404).json({ error: 'Client not found in Excel' });

    await uploadWorkbookToOneDrive('TempData.xlsx', workbook, token);

    res.json({ message: '✅ Timeline data saved successfully to OneDrive!' });
  } catch (err) {
    console.error('❌ Error saving timeline data:', err.message);
    res.status(500).json({ error: 'Internal server error' });
  }
});



// Route to fetch the application timeline for a specific client from OneDrive
app.get('/application-timeline/:clientName', async (req, res) => {
  const clientName = req.params.clientName.toLowerCase().trim();

  try {
    const { workbook } = await getWorkbookFromOneDrive('TempData.xlsx');
    const sheet = workbook.getWorksheet('Client Data');

    if (!sheet) {
      return res.status(404).json({ error: 'Client Data sheet not found' });
    }

    let result = null;

    sheet.eachRow((row) => {
      const nameCell = row.getCell(2).value?.toString().toLowerCase().trim(); // Column B

      if (nameCell === clientName) {
        result = {
          appliedKW: row.getCell(17)?.value || '',
          appliedOnPMSurya: row.getCell(18)?.value || '',
          applicationDHBVN: row.getCell(19)?.value || '',
          loadChangeReductionNewConnection: row.getCell(20)?.value || ''
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


//project status timeline excel
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

    sheet.eachRow((row) => {
      const nameCell = row.getCell(2).value?.toString().toLowerCase().trim(); // Column B
      if (nameCell === clientName) {
        // Columns U (21) to AE (31)
        projectFields.forEach((field, index) => {
          row.getCell(21 + index).value = req.body[field] || '';
        });
        updated = true;
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


app.get('/project-status/:clientName', async (req, res) => {
  const clientName = req.params.clientName.toLowerCase().trim();

  try {
    const { workbook } = await getWorkbookFromOneDrive('TempData.xlsx');
    const sheet = workbook.getWorksheet('Client Data');
    if (!sheet) return res.status(404).json({ error: 'Client Data sheet not found' });

    let result = null;

    sheet.eachRow(row => {
      const name = row.getCell(2).value?.toString().toLowerCase().trim(); // Column B
      if (name === clientName) {
        result = {
          Civil: row.getCell(21)?.value || '',
          Earthing: row.getCell(22)?.value || '',
          EarthingCable: row.getCell(23)?.value || '',
          Panel: row.getCell(24)?.value || '',
          Inverter: row.getCell(25)?.value || '',
          ACDB: row.getCell(26)?.value || '',
          DCDB: row.getCell(27)?.value || '',
          ACCable: row.getCell(28)?.value || '',
          DCCable: row.getCell(29)?.value || '',
          LA: row.getCell(30)?.value || '',
          NetMetering: row.getCell(31)?.value || ''
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

    sheet.eachRow((row) => {
      const name = row.getCell(2).value?.toString().toLowerCase().trim(); // Column B
      if (name === clientName) {
        result = {
          totalCost: row.getCell(8)?.value || '',   // Column H
          advance: row.getCell(7)?.value || '',     // Column G
          projectStatus: {
            Civil: row.getCell(21)?.value || '',       // Column U
            NetMetering: row.getCell(31)?.value || ''  // Column AE
          },
          saved: {
            installment2: row.getCell(32)?.value || '',   // AF
            installment3: row.getCell(33)?.value || '',   // AG
            finalPayment: row.getCell(34)?.value || '',   // AH
            balance: row.getCell(35)?.value || ''         // AI
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

    sheet.eachRow(row => {
      const nameCell = row.getCell(2).value?.toString().toLowerCase().trim(); // Column B
      if (nameCell === clientName) {
        row.getCell(32).value = req.body.installment2 || '';   // AF
        row.getCell(33).value = req.body.installment3 || '';   // AG
        row.getCell(34).value = req.body.finalPayment || '';   // AH
        row.getCell(35).value = req.body.balance || '';        // AI
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
        'To Whom', 'Mobile', 'Price', 'Panel Brand', 'Panel Wp', 'Inverter Brand'
      ]);
    }

    // Auto-generate Ref No like 01P, 02P...
    const lastRow = sheet.lastRow;
    let newNumber = 1;
    if (lastRow && lastRow.getCell(1).value && lastRow.getCell(1).value.toString().endsWith('P')) {
      const lastRef = lastRow.getCell(1).value.toString().replace('P', '');
      newNumber = parseInt(lastRef) + 1;
    }
    const newRef = `${String(newNumber).padStart(2, '0')}P`;

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
      data.inverterBrand
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

// 🚀 Start server
app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
});
