/*************************************
 * server.js (Heroku-ready, polling approach)
 *************************************/
const express = require("express");
const bodyParser = require("body-parser");
const cors = require("cors");
const path = require("path");
const http = require("http");
const { google } = require("googleapis");
const WebSocket = require("ws");

// ★ Heroku環境 or ローカル
const PORT = process.env.PORT || 3000;

// スプレッドシート設定
const SPREADSHEET_ID = "1hMKTxFPIliPWGDt4auIbRq8wGJYCu_-0DKXujqvV4gw";
const API_KEY = "AIzaSyCJMJHHiar0P2e6jdWm0HJdGldAaE3b05I";

const EXCLUDED_SHEETS = [
  "データまとめ",
  "営業所分類",
  "来場数内訳",
  "リーダー",
  "履歴"
];

const app = express();
app.use(cors());
app.use(bodyParser.json());

// 静的ファイル
app.use(express.static(path.join(__dirname, "public")));

// Google Sheets 関数
async function getSpreadsheetInfo() {
  const sheets = google.sheets({ version: "v4" });
  try {
    const resp = await sheets.spreadsheets.get({
      spreadsheetId: SPREADSHEET_ID,
      key: API_KEY,
    });
    const title = resp.data.properties.title;
    let sheetNames = resp.data.sheets.map(s => s.properties.title);
    sheetNames = sheetNames.filter(name => !EXCLUDED_SHEETS.includes(name));
    return { title, sheetNames };
  } catch (err) {
    console.error("[getSpreadsheetInfo] Error:", err.message);
    return { title: "スプレッドシート", sheetNames: [] };
  }
}

async function getMultipleSheetsData(sheetNames) {
  if (!sheetNames || sheetNames.length === 0) return {};
  const sheets = google.sheets({ version: "v4" });
  const ranges = sheetNames.map(name => `${name}!B2:O`);
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

      let parsed = rawData.slice(1).map(row => {
        // B=0, C=1, D=2, ... N=12
        const no = row[0] || "";
        const rawArr  = (row[1] || "").toUpperCase();
        const rawDep  = (row[2] || "").toUpperCase();
        const drName  = row[3] || "";
        const gBlank  = row[4] || "";
        const furigana= row[5] || "";
        const facility= row[6] || "";
        const remarks = row[7] || "";
        const arrTime = row[9]  || "";
        const depTime = row[10] || "";
        const region  = row[11] || "";
        const rawCancel = (row[12] || "").toUpperCase();

        return {
          no,
          arrival    : rawArr==="TRUE" ? "true":"false",
          departure  : rawDep==="TRUE" ? "true":"false",
          drName,
          gBlank,
          furigana,
          facility,
          remarks,
          arrivalTime: arrTime,
          departureTime: depTime,
          region,
          canceledByO: rawCancel==="TRUE"
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

async function getSummaryData() {
  const sheets = google.sheets({ version: "v4" });
  const range = `'来場数内訳'!A1:E13`;
  try {
    const resp = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range,
      key: API_KEY
    });
    let tableData = resp.data.values || [];
    tableData = tableData
      .map(rowArr => rowArr.filter(cell => cell && cell.trim()))
      .filter(r => r.length>0);
    return tableData;
  } catch (err) {
    console.error("[getSummaryData] error:", err.message);
    return [];
  }
}

// ルート
app.get("/info", async (req, res) => {
  console.log("[GET /info]");
  const info = await getSpreadsheetInfo();
  res.json(await info);
});

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

app.get("/summary-data", async (req, res) => {
  console.log("[GET /summary-data]");
  try {
    const table = await getSummaryData();
    res.json(table);
  } catch (err) {
    console.error("[GET /summary-data] error:", err.message);
    res.status(500).send("Error fetching summary data");
  }
});

/** 
 * ★ Webhookは使用しない → 削除 or 無効化
 * app.post("/api/sheetUpdate", ...) 
 * 代わりにポーリング 
 */

// HTTP + WebSocket (もし不要ならWebSocketも削除可)
const httpServer = http.createServer(app);
const wss = new WebSocket.Server({ server: httpServer });

httpServer.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
