/*************************************
 * server.js (Heroku-ready, Polling Example + BasicAuth)
 *************************************/
const express = require("express");
const bodyParser = require("body-parser");
const cors = require("cors");
const path = require("path");
const http = require("http");
const { google } = require("googleapis");
const WebSocket = require("ws");

const PORT = process.env.PORT || 3000;
const SPREADSHEET_ID = "1VzLZaq2kiGzB4QXtlvsDiF9G_h7hENyCb0_rg1lxaJE";
const API_KEY = "AIzaSyCJMJHHiar0P2e6jdWm0HJdGldAaE3b05I";
const EXCLUDED_SHEETS = [
  "データまとめ",
  "営業所分類",
  "来場数内訳",
  "リーダー",
  "NPKK営業所分類",
  "来場数内訳_NPKK",
  "抜粋",
  "宿泊∩未来場",
  "キャンセルまとめ",
  "履歴"
];

const app = express();
app.use(cors());
app.use(bodyParser.json());

// Basic認証
function basicAuth(req, res, next) {
  const authHeader = req.headers["authorization"];
  if (!authHeader) {
    res.setHeader("WWW-Authenticate", 'Basic realm="Restricted"');
    return res.status(401).send("Authentication required");
  }
  const base64 = authHeader.split(" ")[1];
  const decoded = Buffer.from(base64, "base64").toString();
  const [user, pass] = decoded.split(":");
  if (user === "test" && pass === "1771") {
    next();
  } else {
    res.setHeader("WWW-Authenticate", 'Basic realm="Restricted"');
    return res.status(401).send("Invalid credentials");
  }
}
app.use(basicAuth);
app.use(express.static(path.join(__dirname, "public")));

async function getSpreadsheetInfo() {
  const sheets = google.sheets({ version: "v4" });
  try {
    const resp = await sheets.spreadsheets.get({
      spreadsheetId: SPREADSHEET_ID,
      key: API_KEY
    });
    const title = resp.data.properties.title;
    let sheetNames = resp.data.sheets.map(s => s.properties.title);
    sheetNames = sheetNames.filter(n => !EXCLUDED_SHEETS.includes(n));
    return { title, sheetNames };
  } catch (err) {
    console.error("[getSpreadsheetInfo] Error:", err.message);
    return { title: "スプレッドシート", sheetNames: [] };
  }
}

async function getMultipleSheetsData(sheetNames) {
  if (!sheetNames || sheetNames.length === 0) return {};
  const sheets = google.sheets({ version: "v4" });
  const ranges = sheetNames.map(n => `${n}!B2:O`);
  try {
    const resp = await sheets.spreadsheets.values.batchGet({
      spreadsheetId: SPREADSHEET_ID,
      ranges,
      key: API_KEY
    });
    const valueRanges = resp.data.valueRanges || [];
    const result = {};
    valueRanges.forEach(vr => {
      const [rawSheetName] = vr.range.split("!");
      const sheetName = rawSheetName.replace(/^'/, "").replace(/'$/, "");
      const rawData = vr.values || [];
      const parsed = rawData.slice(1).map(row => {
        const no          = row[0] || "";
        const rawArr      = (row[1] || "").toUpperCase();
        const rawDep      = (row[2] || "").toUpperCase();
        const drName      = row[3] || "";
        const honorific   = row[4] || "";
        const furigana    = row[5] || "";
        const facility    = row[6] || "";
        const remarks     = row[7] || "";
        const arrTime     = row[9] || "";
        const depTime     = row[10] || "";
        const facilityLocation = row[11] || "";
        const region      = row[12] || "";
        const rawCancel   = (row[13] || "").toUpperCase();
        return {
          no,
          arrival: (rawArr === "TRUE") ? "true" : "false",
          departure: (rawDep === "TRUE") ? "true" : "false",
          drName,
          honorific,
          furigana,
          facility,
          remarks,
          arrivalTime: arrTime,
          departureTime: depTime,
          facilityLocation,
          region,
          canceledByO: (rawCancel === "TRUE")
        };
      });
      result[sheetName] = parsed;
    });
    return result;
  } catch (err) {
    console.error("[getMultipleSheetsData] error:", err.message);
    return {};
  }
}

// /info エンドポイント
app.get("/info", async (req, res) => {
  console.log("[GET /info]");
  const info = await getSpreadsheetInfo();
  res.json(info);
});

// /batch-data エンドポイント
app.get("/batch-data", async (req, res) => {
  console.log("[GET /batch-data]");
  try {
    const { sheetNames } = await getSpreadsheetInfo();
    const dataObj = await getMultipleSheetsData(sheetNames);
    res.json(dataObj);
  } catch (err) {
    console.error("[GET /batch-data] error:", err.message);
    res.status(500).send("Error fetching batch data");
  }
});

// /summary-data エンドポイント (A1:B16 を取得)
app.get("/summary-data", async (req, res) => {
  console.log("[GET /summary-data]");
  try {
    const sheets = google.sheets({ version: "v4" });
    const range = `'来場数内訳'!A1:B16`;
    const resp = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range,
      key: API_KEY
    });
    let tableData = resp.data.values || [];
    // 各セルの trim のみを実施し、元の構造を保持
    tableData = tableData.map(row => row.map(cell => cell.trim()));
    res.json(tableData);
  } catch (err) {
    console.error("[GET /summary-data] error:", err.message);
    res.status(500).send("Error fetching summary data");
  }
});

const httpServer = http.createServer(app);
const wss = new WebSocket.Server({ server: httpServer });
httpServer.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
