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
  "抜粋",
  "履歴"
];

const app = express();
app.use(cors());
app.use(bodyParser.json());

// Basic認証 (ID=0126, PW=0126)
function basicAuth(req, res, next) {
  const authHeader = req.headers["authorization"];
  if (!authHeader) {
    res.setHeader("WWW-Authenticate", 'Basic realm="Restricted"');
    return res.status(401).send("Authentication required");
  }
  const base64 = authHeader.split(" ")[1];
  const decoded = Buffer.from(base64,"base64").toString();
  const [user, pass] = decoded.split(":");
  if (user==="test" && pass==="1771"){
    next();
  } else {
    res.setHeader("WWW-Authenticate",'Basic realm="Restricted"');
    return res.status(401).send("Invalid credentials");
  }
}
app.use(basicAuth);

// 静的ファイル
app.use(express.static(path.join(__dirname,"public")));

// Google Sheets
async function getSpreadsheetInfo(){
  const sheets=google.sheets({version:"v4"});
  try{
    const resp=await sheets.spreadsheets.get({
      spreadsheetId:SPREADSHEET_ID,
      key:API_KEY
    });
    const title=resp.data.properties.title;
    let sheetNames=resp.data.sheets.map(s=>s.properties.title);
    // 除外
    sheetNames=sheetNames.filter(n=>!EXCLUDED_SHEETS.includes(n));
    return {title, sheetNames};
  } catch(err){
    console.error("[getSpreadsheetInfo] Error:", err.message);
    return {title:"スプレッドシート", sheetNames:[]};
  }
}

async function getMultipleSheetsData(sheetNames){
  if(!sheetNames||sheetNames.length===0) return {};
  const sheets=google.sheets({version:"v4"});
  const ranges=sheetNames.map(n=>`${n}!B2:O`);
  try{
    const resp=await sheets.spreadsheets.values.batchGet({
      spreadsheetId:SPREADSHEET_ID,
      ranges,
      key:API_KEY
    });
    const valueRanges=resp.data.valueRanges||[];
    const result={};

    valueRanges.forEach(vr=>{
      const [rawSheetName]=vr.range.split("!");
      const sheetName=rawSheetName.replace(/^'/,"").replace(/'$/,"");
      const rawData=vr.values||[];
      const parsed=rawData.slice(1).map(row=>{
        // スプレッドシートのデータ
        // B[0]=No、C[1]=来場、D[2]=退場、E[3]=DR氏名、F[4]=敬称、G[5]=フリガナ
        // H[6]=施設名、I[7]=備考、K[9]=来場時間、L[10]=退場時間、M[11]=施設所在地、N[12]=地域区分、O[13]=キャンセル
        const no          = row[0]||"";
        const rawArr      = (row[1]||"").toUpperCase();
        const rawDep      = (row[2]||"").toUpperCase();
        const drName      = row[3]||"";
        const honorific   = row[4]||"";  // 敬称
        const furigana    = row[5]||"";
        const facility    = row[6]||"";
        const remarks     = row[7]||"";
        const arrTime     = row[9]||"";
        const depTime     = row[10]||"";
        const facilityLocation = row[11]||"";  // 施設所在地
        const region      = row[12]||"";
        const rawCancel   = (row[13]||"").toUpperCase();
        
        return {
          no,
          arrival:(rawArr==="TRUE")?"true":"false",
          departure:(rawDep==="TRUE")?"true":"false",
          drName,
          honorific,
          furigana,
          facility,
          remarks,
          arrivalTime:arrTime,
          departureTime:depTime,
          facilityLocation,
          region,
          canceledByO:(rawCancel==="TRUE")
        };
      });
      result[sheetName]=parsed;
    });
    return result;
  }catch(err){
    console.error("[getMultipleSheetsData] error:",err.message);
    return {};
  }
}

// /info
app.get("/info",async(req,res)=>{
  console.log("[GET /info]");
  const info=await getSpreadsheetInfo();
  res.json(info);
});

// /batch-data
app.get("/batch-data",async(req,res)=>{
  console.log("[GET /batch-data]");
  try{
    const {sheetNames}=await getSpreadsheetInfo();
    const dataObj=await getMultipleSheetsData(sheetNames);
    res.json(dataObj);
  }catch(err){
    console.error("[GET /batch-data] error:",err.message);
    res.status(500).send("Error fetching batch data");
  }
});

// /summary-data -> A,B の 1~9
app.get("/summary-data", async(req,res)=>{
  console.log("[GET /summary-data]");
  try{
    const sheets=google.sheets({version:"v4"});
    // 要件: A,B の1~9
    const range=`'来場数内訳'!A1:B16`;
    const resp=await sheets.spreadsheets.values.get({
      spreadsheetId:SPREADSHEET_ID,
      range,
      key:API_KEY
    });
    let tableData=resp.data.values||[];
    // 空セル除去等が必要なら: 
    // tableData=tableData.map(r=>r.filter(c=>c&&c.trim())).filter(r=>r.length>0);
    res.json(tableData);
  }catch(err){
    console.error("[GET /summary-data] error:",err.message);
    res.status(500).send("Error fetching summary data");
  }
});

const httpServer=http.createServer(app);
const wss=new WebSocket.Server({ server:httpServer });
httpServer.listen(PORT,()=>{
  console.log(`Server running on port ${PORT}`);
});