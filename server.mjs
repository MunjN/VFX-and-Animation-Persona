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



import express from "express";
import fetch from "node-fetch";
import cors from "cors";
import dotenv from "dotenv";
import XLSX from "xlsx";

dotenv.config();

const app = express();
const port = 3000;

app.use(cors());
app.use(express.json());

// Auth token route (unchanged)
app.post("/auth-token", async (req, res) => {
  const tenantId = process.env.TENANT_ID;
  const clientId = process.env.CLIENT_ID;
  const clientSecret = process.env.CLIENT_SECRET;
  const scope = "https://analysis.windows.net/powerbi/api/.default";
  const authUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

  try {
    const response = await fetch(authUrl, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        grant_type: "client_credentials",
        client_id: clientId,
        client_secret: clientSecret,
        scope,
      }),
    });

    const data = await response.json();
    if (!response.ok) throw new Error(JSON.stringify(data));
    res.json(data);
  } catch (error) {
    console.error("Auth Token error:", error);
    res.status(500).json({ error: "Failed to fetch Auth Token" });
  }
});

// Embed token route (unchanged)
app.post("/embed-token", async (req, res) => {
  const groupId = "f0795f87-1ddd-47e8-8d54-088db38f6507";
  const reportId = "3ea11afe-6e16-498f-8acd-6df601280226";
  const powerBIUrl = `https://api.powerbi.com/v1.0/myorg/groups/${groupId}/reports/${reportId}/GenerateToken`;
  const authToken = req.body.authToken;

  try {
    const response = await fetch(powerBIUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${authToken}`,
      },
      body: JSON.stringify({ accessLevel: "View" }),
    });

    const data = await response.json();
    if (!response.ok) throw new Error(JSON.stringify(data));
    res.json(data);
  } catch (error) {
    console.error("Embed Token error:", error);
    res.status(500).json({ error: "Failed to fetch Embed Token" });
  }
});

// =========================================================
// 🔹 Premium Export Route (ExportToFile API)
// =========================================================
app.post("/export-visual", async (req, res) => {
  const tenantId = process.env.TENANT_ID;
  const clientId = process.env.CLIENT_ID;
  const clientSecret = process.env.CLIENT_SECRET;
  const groupId = "f0795f87-1ddd-47e8-8d54-088db38f6507";
  const reportId = "3ea11afe-6e16-498f-8acd-6df601280226";

  try {
    // 1️⃣ Get Power BI Access Token
    const tokenResp = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        grant_type: "client_credentials",
        client_id: clientId,
        client_secret: clientSecret,
        scope: "https://analysis.windows.net/powerbi/api/.default",
      }),
    });

    const tokenData = await tokenResp.json();
    const accessToken = tokenData.access_token;
    if (!accessToken) throw new Error("No access token returned.");

    // 2️⃣ Start Export Job (Org Details page)
    const startResp = await fetch(
      `https://api.powerbi.com/v1.0/myorg/groups/${groupId}/reports/${reportId}/ExportTo`,
      {
        method: "POST",
        headers: {
          "Authorization": `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          format: "XLSX",
          powerBIReportConfiguration: {
            pages: [
              {
                // You can also add visualName if you only want one visual
                pageName: "Org Details", // <-- change to your actual Org Details page name
              },
            ],
          },
        }),
      }
    );

    const startData = await startResp.json();
    if (!startResp.ok) throw new Error(JSON.stringify(startData));

    const exportId = startData.id;
    console.log("Export job started:", exportId);

    // 3️⃣ Poll until export job completes
    let status = "Running";
    let exportResult = null;

    while (status === "Running" || status === "NotStarted") {
      await new Promise((r) => setTimeout(r, 3000));
      const pollResp = await fetch(
        `https://api.powerbi.com/v1.0/myorg/groups/${groupId}/reports/${reportId}/exports/${exportId}`,
        { headers: { Authorization: `Bearer ${accessToken}` } }
      );
      exportResult = await pollResp.json();
      status = exportResult.status;
      console.log("Export status:", status);
    }

    if (status !== "Succeeded") throw new Error(`Export failed: ${status}`);

    // 4️⃣ Download exported file
    const fileResp = await fetch(exportResult.resourceLocation, {
      headers: { Authorization: `Bearer ${accessToken}` },
    });
    const buffer = await fileResp.arrayBuffer();

    // 5️⃣ Return file to client
    res.setHeader("Content-Disposition", "attachment; filename=OrgDetails_Export.xlsx");
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.send(Buffer.from(buffer));
  } catch (err) {
    console.error("Premium export error:", err);
    res.status(500).json({ error: err.message });
  }
});

app.listen(port, () => console.log(`🚀 Server running on http://localhost:${port}`));
