/*************************************
 * server.js
 *************************************/
const express = require("express");
const bodyParser = require("body-parser");
const cors = require("cors");
const path = require("path");
const http = require("http");
const WebSocket = require("ws");
const { google } = require("googleapis");

// ====== Herokuで単一ポート ======
const PORT = process.env.PORT || 3000;

// Basic認証 ID/PW
const BASIC_USER = "abc";
const BASIC_PASS = "ABC";

// Google Sheets API設定
const SPREADSHEET_ID = "1pMus9SiW5B26B9rc6lzdG3rc0IkA4lEMkUAwWVqWm4I";
const API_KEY = 'AIzaSyCJMJHHiar0P2e6jdWm0HJdGldAaE3b05I'; // 実際のキーに

// 除外するシート
const EXCLUDED_SHEETS = [
  "データまとめ",
  "営業所分類",
  "来場数内訳", // タブには表示しないがサブテーブル表示で使う
  "リーダー",
  "履歴"
];

const app = express();

// Basic認証ミドルウェア
app.use((req, res, next) => {
  const authHeader = req.headers["authorization"];
  if (!authHeader) {
    res.setHeader("WWW-Authenticate", 'Basic realm="Restricted Area"');
    return res.status(401).send("Authentication required");
  }
  const base64Credentials = authHeader.split(" ")[1];
  const decoded = Buffer.from(base64Credentials, "base64").toString();
  const [user, pass] = decoded.split(":");
  if (user === BASIC_USER && pass === BASIC_PASS) {
    next();
  } else {
    res.setHeader("WWW-Authenticate", 'Basic realm="Restricted Area"');
    return res.status(401).send("Invalid credentials");
  }
});

app.use(cors());
app.use(bodyParser.json());

// 静的ファイル (public)
app.use(express.static(path.join(__dirname, "public")));

// ====== Google Sheets 関数 ======
async function getSpreadsheetInfo() {
  const sheets = google.sheets({ version: "v4" });
  try {
    const resp = await sheets.spreadsheets.get({
      spreadsheetId: SPREADSHEET_ID,
      key: API_KEY
    });
    const title = resp.data.properties.title;
    let sheetNames = resp.data.sheets.map(s => s.properties.title);
    // 除外シートを削除
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

      let parsed = rawData.slice(1)
        .filter(row => {
          if (!row) return false;
          const isAllEmpty = row.every(cell => !cell || !cell.trim?.());
          return !isAllEmpty;
        })
        .map(row => {
          // B=0, C=1, D=2, E=3, F=4, G=5, H=6, I=7, ...
          const colC = row[1] || "";
          const colF = row[4] || "";
          const colH = row[6] || "";
          return {
            no: row[0] || "",
            arrival: row[2]?.toUpperCase() === "TRUE" ? "true" : "false",
            departure: row[3]?.toUpperCase() === "TRUE" ? "true" : "false",
            drName: row[4] || "",
            gBlank: row[5] || "",
            furigana: row[6] || "",
            facility: row[7] || "",
            remarks: row[8] || "",
            arrivalTime: row[10] || "",
            departureTime: row[11] || "",
            region: row[12] || "",
            canceledByO: row[13]?.toUpperCase() === "TRUE",

            colC, colF, colH
          };
        });

      // 全体シート => C,F,Hいずれか空なら除外
      if (sheetName === "全体(あいうえお順)") {
        parsed = parsed.filter(r => {
          if (!r.colC.trim() || !r.colF.trim() || !r.colH.trim()) {
            return false;
          }
          return true;
        });
      }

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
  const range = `'来場数内訳'!B2:E12`;
  try {
    const resp = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range,
      key: API_KEY
    });
    let tableData = resp.data.values || [];
    // 空セル除外
    tableData = tableData.map(rowArr => {
      const filtered = rowArr.filter(cell => cell && cell.trim());
      return filtered;
    }).filter(r => r.length > 0);
    return tableData;
  } catch (err) {
    console.error("[getSummaryData] error:", err.message);
    return [];
  }
}

// ====== ルート ======
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

// Webhook
app.post("/api/sheetUpdate", (req, res) => {
  console.log("[POST /api/sheetUpdate] body=", req.body);
  wss.clients.forEach(client => {
    if (client.readyState === WebSocket.OPEN) {
      client.send(JSON.stringify({
        type: "sheetUpdated",
        sheetName: req.body.sheetName,
        range: req.body.range,
        newValue: req.body.newValue
      }));
    }
  });
  res.status(200).json({ status: "ok" });
});

// 単一ポートで HTTP + WebSocket
const httpServer = http.createServer(app);
const wss = new WebSocket.Server({ server: httpServer });

// リッスン
httpServer.listen(PORT, () => {
  console.log(`Server & WS running on port ${PORT}`);
});
