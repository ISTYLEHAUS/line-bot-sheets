const express = require('express');
const axios = require('axios');
const { google } = require('googleapis');
const app = express();

app.use(express.json());

const LINE_TOKEN = process.env.LINE_CHANNEL_ACCESS_TOKEN;
const SHEET_ID = process.env.SHEET_ID;
const SHEET_NAME = process.env.SHEET_NAME;
const GOOGLE_CREDENTIALS = JSON.parse(process.env.GOOGLE_CREDENTIALS);

async function getSheet() {
  const auth = new google.auth.GoogleAuth({
    credentials: GOOGLE_CREDENTIALS,
    scopes: ['https://www.googleapis.com/auth/spreadsheets']
  });
  const sheets = google.sheets({ version: 'v4', auth });
  return sheets;
}

app.get('/', (req, res) => res.send('OK'));

app.post('/webhook', async (req, res) => {
  res.sendStatus(200);

  const events = req.body.events || [];

  for (const event of events) {
    const replyToken = event.replyToken;

    if (event.type === 'postback') {
      const data = JSON.parse(event.postback.data);
      const sheets = await getSheet();
      await sheets.spreadsheets.values.append({
        spreadsheetId: SHEET_ID,
        range: `${SHEET_NAME}!A:C`,
        valueInputOption: 'USER_ENTERED',
        requestBody: {
          values: [[new Date().toLocaleString('th-TH'), data.text, data.status]]
        }
      });
      await replyMessage(replyToken, `✅ บันทึกแล้ว\nข้อความ: ${data.text}\nสถานะ: ${data.status}`);
      continue;
    }

    if (event.type === 'message' && event.message.type === 'text') {
      await replyWithButtons(replyToken, event.message.text);
    }
  }
});

async function replyMessage(replyToken, text) {
  await axios.post('https://api.line.me/v2/bot/message/reply', {
    replyToken,
    messages: [{ type: 'text', text }]
  }, { headers: { Authorization: `Bearer ${LINE_TOKEN}` } });
}

async function replyWithButtons(replyToken, text) {
  await axios.post('https://api.line.me/v2/bot/message/reply', {
    replyToken,
    messages: [{
      type: 'text',
      text:
