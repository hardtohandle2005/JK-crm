const path = require("path");
const axios = require("axios");
const ExcelJS = require("exceljs");
const mime = require("mime-types");
const qs = require("qs");
const formidable = require("formidable");
const fs = require("fs");

const CLIENT_ID = "8a127620-3295-4498-8701-5102725dd17e";
const CLIENT_SECRET = "AE98Q~cqDbVt8TsSHS9Oaa18oNRinfLNtScBb.b";
const TENANT_ID = "785fd7e9-594d-4549-91b9-9372f7295962";
const ONEDRIVE_USER = "MuninderPal@JK17.onmicrosoft.com";

module.exports = async function (context, req) {
  const form = formidable({ multiples: true });

  form.parse(req, async (err, fields, files) => {
    if (err) {
      context.log.error("❌ Form parsing error:", err);
      context.res = {
        status: 500,
        body: "Form parsing error"
      };
      return;
    }

    try {
      const data = fields;
      const clientName = data.name.trim().toLowerCase().replace(/\s+/g, "_");

      // 🖼️ Upload client photo (base64)
      if (data.photoData) {
        const base64Match = data.photoData.match(/^data:image\/(\w+);base64,(.+)$/);
        if (base64Match) {
          const extension = base64Match[1];
          const base64Data = base64Match[2];
          const buffer = Buffer.from(base64Data, "base64");

          context.log(`🖼️ Base64 photo detected: .${extension} (${base64Data.length} bytes)`);

          await uploadToOneDriveFolder(clientName, "clientPhoto", buffer, `clientPhoto.${extension}`);
          context.log("📤 clientPhoto uploaded.");
        } else {
          context.log.warn("⚠️ Invalid base64 format in photoData");
        }
      }

      // 📤 Upload all files (with pretty logging)
      const uploadTasks = Object.entries(files).map(async ([field, fileObj]) => {
        const file = fileObj[0] || fileObj;
        const ext = path.extname(file.originalFilename || file.name);
        const customName = `${field.toLowerCase()}${ext}`;
        const buffer = fs.readFileSync(file.filepath || file.path);

        await uploadToOneDriveFolder(clientName, field, buffer, customName);

        const pretty = field.replace(/([A-Z])/g, " $1").replace(/^./, s => s.toUpperCase());
        context.log(`📤 ${pretty} uploaded...`);
      });

      await Promise.all(uploadTasks);

      // 📊 Save to Excel
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

      const { workbook, token } = await getWorkbookFromOneDrive("TempData.xlsx");
      const sheet = workbook.getWorksheet("Client Data") || workbook.addWorksheet("Client Data");

      sheet.addRow([
        client.date, client.name, client.address, client.mobile, client.email, client.kw,
        client.advance, client.totalCost, client.aadharFront, client.aadharBack, client.panCard, client.bill,
        client.ownershipProof, client.cancelCheque, client.purchaseAgreement, client.netMeteringAgreement
      ]);

      await uploadWorkbookToOneDrive("TempData.xlsx", workbook, token);
      context.log("✅ Client data saved to OneDrive Excel");

      context.res = {
        status: 200,
        headers: { "Content-Type": "text/html" },
        body: `
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
      <div class="success">✅ Client submitted successfully and all files saved.</div>
    </body>
  </html>
`
      };
    } catch (err) {
      context.log.error("❌ Error in /submit-client:", err.message);
      context.res = {
        status: 500,
        body: "⚠️ Could not save data to OneDrive."
      };
    }
  });
};

// --- Helpers ---

async function getAccessToken() {
  const url = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;

  const { data } = await axios.post(url, qs.stringify({
    client_id: CLIENT_ID,
    scope: "https://graph.microsoft.com/.default",
    client_secret: CLIENT_SECRET,
    grant_type: "client_credentials"
  }), {
    headers: { "Content-Type": "application/x-www-form-urlencoded" }
  });

  return data.access_token;
}

async function getWorkbookFromOneDrive(fileName) {
  const token = await getAccessToken();
  const url = `https://graph.microsoft.com/v1.0/users/${ONEDRIVE_USER}/drive/root:/${fileName}:/content`;

  const response = await axios.get(url, {
    headers: { Authorization: `Bearer ${token}` },
    responseType: "arraybuffer"
  });

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(response.data);

  return { workbook, token };
}

async function uploadWorkbookToOneDrive(fileName, workbook, token) {
  const buffer = await workbook.xlsx.writeBuffer();
  const url = `https://graph.microsoft.com/v1.0/users/${ONEDRIVE_USER}/drive/root:/${fileName}:/content`;

  await axios.put(url, buffer, {
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    }
  });
}

async function uploadToOneDriveFolder(clientName, field, buffer, fileName) {
  const token = await getAccessToken();
  const fullPath = `uploads/${clientName}/${fileName}`;
  const url = `https://graph.microsoft.com/v1.0/users/${ONEDRIVE_USER}/drive/root:/${fullPath}:/content`;
  const contentType = mime.lookup(fileName) || "application/octet-stream";

  await axios.put(url, buffer, {
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": contentType
    }
  });
}
