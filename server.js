/*************************************
 * server.js
 *************************************/
const express = require('express');
const bodyParser = require('body-parser');
const { google } = require('googleapis');
const WebSocket = require('ws');
const cors = require('cors');
const path = require('path');

// ========== 設定 ==========
const app = express();
const PORT = 3000;

// Basic認証用のID/PW
const BASIC_USER = 'abc';
const BASIC_PASS = 'ABC';

// 1) CORS, JSONパース
app.use(cors());
app.use(bodyParser.json());

// 2) Basic認証ミドルウェア
app.use((req, res, next) => {
  const authHeader = req.headers['authorization'];
  if (!authHeader) {
    // 認証ヘッダが無い => 401
    res.setHeader('WWW-Authenticate', 'Basic realm="Restricted Area"');
    return res.status(401).send('Authentication required');
  }
  const base64Credentials = authHeader.split(' ')[1];
  const decoded = Buffer.from(base64Credentials, 'base64').toString();
  const [user, pass] = decoded.split(':');

  if (user === BASIC_USER && pass === BASIC_PASS) {
    next();
  } else {
    res.setHeader('WWW-Authenticate', 'Basic realm="Restricted Area"');
    return res.status(401).send('Invalid credentials');
  }
});

// ========== WebSocket ==========
const wss = new WebSocket.Server({ port: 8081 });

// ========== スプレッドシート設定 ========== //
const SPREADSHEET_ID = '1pMus9SiW5B26B9rc6lzdG3rc0IkA4lEMkUAwWVqWm4I';
const API_KEY = 'AIzaSyCJMJHHiar0P2e6jdWm0HJdGldAaE3b05I';

// 除外するシート
const EXCLUDED_SHEETS = [
  'データまとめ',
  '営業所分類',
  '来場数内訳',  // ←このシートはタブに表示しない
  'リーダー',
  '履歴'
];

// 静的ファイル配信
app.use(express.static(path.join(__dirname, 'public')));

// ====== スプレッドシート関数群 ======
async function getSpreadsheetInfo() {
  const sheets = google.sheets({ version: 'v4' });
  try {
    const resp = await sheets.spreadsheets.get({
      spreadsheetId: SPREADSHEET_ID,
      key: API_KEY,
    });
    const title = resp.data.properties.title;
    let sheetNames = resp.data.sheets.map(s => s.properties.title);
    // 除外シートを削除
    sheetNames = sheetNames.filter(name => !EXCLUDED_SHEETS.includes(name));
    return { title, sheetNames };
  } catch (err) {
    console.error('[getSpreadsheetInfo] Error:', err.message);
    return { title: 'スプレッドシート', sheetNames: [] };
  }
}

/**
 * 各シート(B2:O)をまとめて取得
 * "全体"シートのみ row[1], row[4], row[6] (C,F,H列相当) がすべて非空でないと表示しない
 */
async function getMultipleSheetsData(sheetNames) {
  if (!sheetNames || sheetNames.length === 0) return {};

  const sheets = google.sheets({ version: 'v4' });
  const ranges = sheetNames.map(name => `${name}!B2:O`);
  try {
    const resp = await sheets.spreadsheets.values.batchGet({
      spreadsheetId: SPREADSHEET_ID,
      ranges,
      key: API_KEY,
    });
    const valueRanges = resp.data.valueRanges || [];
    const result = {};

    valueRanges.forEach(vr => {
      const [rawSheetName] = vr.range.split('!');
      const sheetName = rawSheetName.replace(/^'/, '').replace(/'$/, '');
      const rawData = vr.values || [];

      // parse rows
      let parsed = rawData.slice(1)  // skip header
        .filter(row => {
          if (!row) return false;
          const isAllEmpty = row.every(cell => !cell || !cell.trim?.());
          return !isAllEmpty;
        })
        .map(row => ({
          no: row[0] || '',
          arrival: row[2]?.toUpperCase() === 'TRUE' ? 'true' : 'false',
          departure: row[3]?.toUpperCase() === 'TRUE' ? 'true' : 'false',
          drName: row[4] || '',
          gBlank: row[5] || '',
          furigana: row[6] || '',
          facility: row[7] || '',
          remarks: row[8] || '',
          arrivalTime: row[10] || '',
          departureTime: row[11] || '',
          region: row[12] || '',
          canceledByO: row[13]?.toUpperCase() === 'TRUE',

          // 下記が "C,F,H列" に相当(※B2:Oマッピングで B=0, C=1, D=2, E=3, F=4, G=5, H=6)
          colC: row[1] || '',  // C列
          colF: row[4] || '',  // F列 → 既に drNameに使っている row[4], ただし下で再利用
          colH: row[6] || ''   // H列 → 既に furiganaに使っている row[6]
        }));

      // "全体"シートの場合のみ、C,F,H列が全て空の場合は非表示(=除外)
      if (sheetName === '全体') {
        parsed = parsed.filter(r => {
          // C, F, H 列に入力がない => skip
          // r.colC, r.colF, r.colH がいずれか空なら除外、と解釈
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
    console.error('[batchGet] error:', err.message);
    return {};
  }
}

// 「来場数内訳」シート (B2:E12)
async function getSummaryData() {
  const sheets = google.sheets({ version: 'v4' });
  const range = `'来場数内訳'!B2:E12`;
  try {
    const resp = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range,
      key: API_KEY,
    });
    let tableData = resp.data.values || [];
    // 空セル削除
    tableData = tableData.map(rowArr => {
      const filtered = rowArr.filter(cell => cell && cell.trim());
      return filtered;
    }).filter(r => r.length > 0);
    return tableData;
  } catch (err) {
    console.error('[getSummaryData] error:', err.message);
    return [];
  }
}

// ====== ルート ======
app.get('/info', async (req, res) => {
  console.log('[GET /info]');
  const info = await getSpreadsheetInfo();
  res.json(await info);
});

app.get('/batch-data', async (req, res) => {
  console.log('[GET /batch-data]');
  try {
    const { sheetNames } = await getSpreadsheetInfo();
    const dataObj = await getMultipleSheetsData(sheetNames);
    res.json(dataObj);
  } catch (err) {
    console.error('[GET /batch-data] error:', err.message);
    res.status(500).send('Error fetching batch data');
  }
});

app.get('/summary-data', async (req, res) => {
  console.log('[GET /summary-data]');
  try {
    const table = await getSummaryData();
    res.json(table);
  } catch (err) {
    console.error('[GET /summary-data] error:', err.message);
    res.status(500).send('Error fetching summary data');
  }
});

// Webhook
app.post('/api/sheetUpdate', (req, res) => {
  const { sheetName, range, newValue } = req.body;
  console.log('[POST /api/sheetUpdate] body=', req.body);

  wss.clients.forEach(client => {
    if (client.readyState === WebSocket.OPEN) {
      client.send(JSON.stringify({
        type: 'sheetUpdated',
        sheetName,
        range,
        newValue
      }));
    }
  });

  res.status(200).json({ status: 'ok' });
});

app.listen(PORT, () => {
  console.log(`Server on http://localhost:${PORT}`);
  console.log(`WebSocket on ws://localhost:8081`);
});
