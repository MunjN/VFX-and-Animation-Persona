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

// =========================================================
// 🔹 Auth Token Route
// =========================================================
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

// =========================================================
// 🔹 Embed Token Route
// =========================================================
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
// 🔹 Dynamic Filtered Export to Excel Route
// =========================================================
app.post("/export-to-excel", async (req, res) => {
  const tenantId = process.env.TENANT_ID;
  const clientId = process.env.CLIENT_ID;
  const clientSecret = process.env.CLIENT_SECRET;
  const datasetId = "867dc48c-c22e-4ed8-90ec-a16952dfcbf0";
  const groupId = "f0795f87-1ddd-47e8-8d54-088db38f6507";
  const filters = req.body.filters || [];

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

    // 2️⃣ Build DAX Query (null-safe)
    let filterConditions = "";

    if (Array.isArray(filters) && filters.length > 0) {
      const validFilters = [];

      for (const f of filters) {
        try {
          if (
            !f.target ||
            typeof f.target.table !== "string" ||
            typeof f.target.column !== "string" ||
            !Array.isArray(f.values)
          ) {
            console.log("⚠️ Skipping invalid filter:", f);
            continue;
          }

          const table = f.target.table.trim();
          const column = f.target.column.trim();
          if (!table || !column) continue;

          const cleanValues = f.values.filter((v) => v !== null && v !== undefined);
          if (cleanValues.length === 0) continue;

      const colRef = `'${table}'[${column}]`;
const vals = cleanValues
  .map((v) => `"${String(v).replace(/"/g, '""')}"`)
  .join(", ");

let clause = `${colRef} IN {${vals}}`;
if (f.operator && f.operator.toLowerCase() === "notin") {
  clause = `NOT(${colRef} IN {${vals}})`;
}


          validFilters.push(clause);
        } catch (e) {
          console.log("⚠️ Filter parse error, skipping:", e);
        }
      }

      if (validFilters.length > 0) {
        filterConditions = `, ${validFilters.join(", ")}`;
      }
    }

    const daxQuery = `
      EVALUATE
      TOPN(
        200,
        SELECTCOLUMNS(
          CALCULATETABLE(
            'UNPIVOTED_FX_DATA'
            ${filterConditions}
          ),
          "ORG_ID", 'UNPIVOTED_FX_DATA'[ORG_ID],
          "ORG_NAME", 'UNPIVOTED_FX_DATA'[ORG_NAME],
          "CATEGORY", 'UNPIVOTED_FX_DATA'[CATEGORY],
          "ATTRIBUTE", 'UNPIVOTED_FX_DATA'[ATTRIBUTE],
          "VALUE", 'UNPIVOTED_FX_DATA'[DISPLAY_VALUE]
        )
      )
    `;

    console.log("🧠 Final DAX Query:\n", daxQuery);

    // 3️⃣ Execute DAX Query
    const execResp = await fetch(
      `https://api.powerbi.com/v1.0/myorg/groups/${groupId}/datasets/${datasetId}/executeQueries`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({ queries: [{ query: daxQuery }] }),
      }
    );

    const execData = await execResp.json();
    if (!execResp.ok) throw new Error(JSON.stringify(execData));

    const tableData = execData.results[0].tables[0];
    if (!tableData || !tableData.columns || !tableData.rows) {
      throw new Error("No data returned from Power BI.");
    }

    const columns = tableData.columns.map((c) => c.name);
    const rows = tableData.rows;

    // 4️⃣ Convert JSON to Excel
    const sheetData = [columns, ...rows];
    const worksheet = XLSX.utils.aoa_to_sheet(sheetData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "OrgDetails");

    const buffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });

    // 5️⃣ Send file to client
    res.setHeader("Content-Disposition", "attachment; filename=OrgDetails_Export.xlsx");
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.send(buffer);
  } catch (err) {
    console.error("❌ Export to Excel error:", err);
    res.status(500).json({ error: err.message });
  }
});

// =========================================================
// 🚀 Start Server
// =========================================================
app.listen(port, () => console.log(`🚀 Server running on http://localhost:${port}`));
