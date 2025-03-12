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

// 1) HerokuのPORT or ローカル3000
const PORT = process.env.PORT || 3000;

// 2) スプレッドシート設定
const SPREADSHEET_ID = "1VzLZaq2kiGzB4QXtlvsDiF9G_h7hENyCb0_rg1lxaJE";
const API_KEY = "AIzaSyCJMJHHiar0P2e6jdWm0HJdGldAaE3b05I";

// 除外シート
const EXCLUDED_SHEETS = [
  "データまとめ",
  "営業所分類",
  "来場数内訳",
  "リーダー",
  "NPKK営業所分類",
  "来場数内訳_NPKK",
  "履歴",
  "抜粋"
];

const app = express();
app.use(cors());
app.use(bodyParser.json());

/** ★ Basic認証 (ID=0126, PW=0126) を全エンドポイントに適用 */
function basicAuth(req, res, next) {
  const authHeader = req.headers["authorization"];
  if (!authHeader) {
    res.setHeader("WWW-Authenticate", 'Basic realm="Restricted"');
    return res.status(401).send("Authentication required");
  }
  // Basic 認証のフォーマット: "Basic base64(<user>:<pass>)"
  const base64Credentials = authHeader.split(" ")[1];
  const decoded = Buffer.from(base64Credentials, "base64").toString();
  const [user, pass] = decoded.split(":");

  if (user === "test" && pass === "1771") {
    next(); // 認証OK
  } else {
    res.setHeader("WWW-Authenticate", 'Basic realm="Restricted"');
    return res.status(401).send("Invalid credentials");
  }
}
app.use(basicAuth);

/** 
 * 3) 静的ファイル (index.htmlなど)
 *    Basic認証が通った後に配信される
 */
app.use(express.static(path.join(__dirname, "public")));

// 4) Google Sheets 関数
async function getSpreadsheetInfo() {
  const sheets = google.sheets({ version: "v4" });
  try {
    const resp = await sheets.spreadsheets.get({
      spreadsheetId: SPREADSHEET_ID,
      key: API_KEY,
    });
    const title = resp.data.properties.title;
    let sheetNames = resp.data.sheets.map(s => s.properties.title);
    // 除外
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
        // B=0, C=1, D=2, E=3 ... N=12
        const no        = row[0] || "";
        const rawArr    = (row[1]||"").toUpperCase();
        const rawDep    = (row[2]||"").toUpperCase();
        const drName    = row[3] || "";
        const gBlank    = row[4] || "";
        const furigana  = row[5] || "";
        const facility  = row[6] || "";
        const remarks   = row[7] || "";
        const arrTime   = row[9] || "";
        const depTime   = row[10]|| "";
        const region    = row[11]|| "";
        const rawCancel = (row[12]||"").toUpperCase();
        return {
          no,
          arrival    : (rawArr==="TRUE")?"true":"false",
          departure  : (rawDep==="TRUE")?"true":"false",
          drName,
          gBlank,
          furigana,
          facility,
          remarks,
          arrivalTime: arrTime,
          departureTime: depTime,
          region,
          canceledByO: (rawCancel==="TRUE")
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

// ★ 修正箇所 ★
// getSummaryData の取得範囲をA1:E16に変更し、フィルタ処理を削除して各セルのtrimのみを行う
async function getSummaryData() {
  const sheets = google.sheets({ version: "v4" });
  const range = `'来場数内訳'!A1:E16`;  // 変更：16行目まで取得
  try {
    const resp = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range,
      key: API_KEY
    });
    let tableData = resp.data.values || [];
    // 変更：フィルタ処理ではなく、各セルのtrimのみを実施
    tableData = tableData.map(row => row.map(cell => cell.trim()));
    return tableData;
  } catch (err) {
    console.error("[getSummaryData] error:", err.message);
    return [];
  }
}

// 5) ルート
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
    console.error("[getSummaryData] error:", err.message);
    res.status(500).send("Error fetching summary data");
  }
});

/** 
 * Webhook不要の場合、コメントアウト or 削除
 * 
app.post("/api/sheetUpdate", (req, res) => {
  ...
});
*/

const httpServer = http.createServer(app);
const wss = new WebSocket.Server({ server: httpServer });

httpServer.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
