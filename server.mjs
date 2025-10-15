// import express from 'express';
// import fetch from 'node-fetch';
// import cors from 'cors';
// import e from 'express';
// import dotenv from 'dotenv';


// const app = express();
// const port = 3000;

// // Enable CORS for all origins
// app.use(cors());

// // Middleware to parse JSON bodies
// app.use(express.json());

// // Endpoint to fetch Auth Token
// app.post('/auth-token', async (req, res) => { 
//   // environment variables
//   const tenantId = process.env.TENANT_ID;
//   const clientId = process.env.CLIENT_ID;  
//   const clientSecret = process.env.CLIENT_SECRET;  

//   const scope = "https://analysis.windows.net/powerbi/api/.default";
//   const authUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

//   try {
//     const response = await fetch(authUrl, {
//       method: 'POST',
//       headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
//       body: new URLSearchParams({
//         grant_type: 'client_credentials',
//         client_id: clientId,
//         client_secret: clientSecret,
//         scope: scope
//       })
//     });

//     if (!response.ok) {
//       throw new Error('Failed to fetch Auth Token');
//     }

//     const data = await response.json();
//     res.json(data); // Send back the auth token to the frontend
//   } catch (error) {
//     console.error('Error fetching Auth Token:', error);
//     res.status(500).json({ error: 'Failed to fetch Auth Token' });
//   }
// });

// // Endpoint to fetch Embed Token
// app.post('/embed-token', async (req, res) => {
//   const groupId = "f0795f87-1ddd-47e8-8d54-088db38f6507";
//   const reportId = "3ea11afe-6e16-498f-8acd-6df601280226";
//   const powerBIUrl = `https://api.powerbi.com/v1.0/myorg/groups/${groupId}/reports/${reportId}/GenerateToken`;

//   // Auth token is sent from the frontend
//   const authToken = req.body.authToken;

//   try {
//     const response = await fetch(powerBIUrl, {
//       method: 'POST',
//       headers: {
//         'Content-Type': 'application/json',
//         Authorization: `Bearer ${authToken}`
//       },
//       body: JSON.stringify({ accessLevel: 'View' }) 
//     });

//     if (!response.ok) {
//       throw new Error('Failed to fetch Embed Token');
//     }

//     const data = await response.json();
//     res.json(data); // Send back the embed token and embed URL to the frontend
//   } catch (error) {
//     console.error('Error fetching Embed Token:', error);
//     res.status(500).json({ error: 'Failed to fetch Embed Token' });
//   }
// });

// // Start the server
// app.listen(port, () => {
//   console.log(`Server running on http://localhost:${port}`);
// });



import express from 'express';
import fetch from 'node-fetch';
import cors from 'cors';
import dotenv from 'dotenv';
import XLSX from 'xlsx'; // 📦 npm install xlsx

dotenv.config();

const app = express();
const port = 3000;

app.use(cors());
app.use(express.json());

// --------------------------------------------------
// 🔹 Auth Token Endpoint
// --------------------------------------------------
app.post('/auth-token', async (req, res) => {
  const tenantId = process.env.TENANT_ID;
  const clientId = process.env.CLIENT_ID;
  const clientSecret = process.env.CLIENT_SECRET;
  const scope = "https://analysis.windows.net/powerbi/api/.default";
  const authUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

  try {
    const response = await fetch(authUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        grant_type: 'client_credentials',
        client_id: clientId,
        client_secret: clientSecret,
        scope
      })
    });

    if (!response.ok) throw new Error('Failed to fetch Auth Token');
    const data = await response.json();
    res.json(data);
  } catch (error) {
    console.error('Error fetching Auth Token:', error);
    res.status(500).json({ error: 'Failed to fetch Auth Token' });
  }
});

// --------------------------------------------------
// 🔹 Embed Token Endpoint
// --------------------------------------------------
app.post('/embed-token', async (req, res) => {
  const groupId = "f0795f87-1ddd-47e8-8d54-088db38f6507";
  const reportId = "3ea11afe-6e16-498f-8acd-6df601280226";
  const powerBIUrl = `https://api.powerbi.com/v1.0/myorg/groups/${groupId}/reports/${reportId}/GenerateToken`;
  const authToken = req.body.authToken;

  try {
    const response = await fetch(powerBIUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${authToken}`
      },
      body: JSON.stringify({ accessLevel: 'View' })
    });

    if (!response.ok) throw new Error('Failed to fetch Embed Token');
    const data = await response.json();
    res.json(data);
  } catch (error) {
    console.error('Error fetching Embed Token:', error);
    res.status(500).json({ error: 'Failed to fetch Embed Token' });
  }
});

// --------------------------------------------------
// 🔹 Excel Export Endpoint (Org Details only)
// --------------------------------------------------
app.post('/export', async (req, res) => {
  try {
    const tenantId = process.env.TENANT_ID;
    const clientId = process.env.CLIENT_ID;
    const clientSecret = process.env.CLIENT_SECRET;
    const datasetId = "867dc48c-c22e-4ed8-90ec-a16952dfcbf0";

    // 1️⃣ Get Power BI Access Token
    const tokenResp = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        grant_type: "client_credentials",
        client_id: clientId,
        client_secret: clientSecret,
        scope: "https://analysis.windows.net/powerbi/api/.default"
      })
    });
    const tokenData = await tokenResp.json();
    const accessToken = tokenData.access_token;

    // 2️⃣ Build DAX query (hardcoded for Org Details)
    const daxQuery = {
      queries: [
        {
          query: `
            EVALUATE
            TOPN(
              200,
              'Org Details'
            )
          `
        }
      ]
    };

    // 3️⃣ Execute DAX query
    const queryResp = await fetch(`https://api.powerbi.com/v1.0/myorg/datasets/${datasetId}/executeQueries`, {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${accessToken}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify(daxQuery)
    });

    if (!queryResp.ok) {
      const errText = await queryResp.text();
      throw new Error(`Power BI query failed: ${errText}`);
    }

    const result = await queryResp.json();
    const rows = result?.results?.[0]?.tables?.[0]?.rows || [];

    if (!rows.length) {
      res.status(404).json({ message: "No data returned from Org Details." });
      return;
    }

    // 4️⃣ Convert to Excel
    const worksheet = XLSX.utils.json_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "OrgDetails");
    const buffer = XLSX.write(workbook, { bookType: "xlsx", type: "buffer" });

    // 5️⃣ Send file to client
    res.setHeader("Content-Disposition", "attachment; filename=OrgDetails_Export.xlsx");
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.send(buffer);
  } catch (err) {
    console.error("Export error:", err);
    res.status(500).json({ error: err.message });
  }
});

// --------------------------------------------------
// 🔹 Start Server
// --------------------------------------------------
app.listen(port, () => {
  console.log(`🚀 Server running on http://localhost:${port}`);
});
