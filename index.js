const express = require('express');
const axios = require('axios');
const { google } = require('googleapis');
const app = express();

app.use(express.json());

const LINE_TOKEN = process.env.LINE_CHANNEL_ACCESS_TOKEN;
const SHEET_ID = process.env.SHEET_ID;
const SHEET_NAME = process.env.SHEET_NAME || 'AutoBot';
const GOOGLE_CREDENTIALS = JSON.parse(process.env.GOOGLE_CREDENTIALS);

const MONTH_FOLDERS = {
  "05": process.env.FOLDER_MAY,
  "06": process.env.FOLDER_JUN,
  "07": process.env.FOLDER_JUL,
  "08": process.env.FOLDER_AUG,
  "09": process.env.FOLDER_SEP,
  "10": process.env.FOLDER_OCT,
  "11": process.env.FOLDER_NOV,
  "12": process.env.FOLDER_DEC
};

const CATEGORIES = [
  "Production", "Raw Material", "Ads", "Salary", "Shipping",
  "Account", "Service", "Sales Revenue (Transfer)", "Other Income", "Office & Admin"
];

const sessions = {};

async function getSheets() {
  const auth = new google.auth.GoogleAuth({
    credentials: GOOGLE_CREDENTIALS,
    scopes: ['https://www.googleapis.com/auth/spreadsheets']
  });
  return google.sheets({ version: 'v4', auth });
}

async function getDrive() {
  const auth = new google.auth.GoogleAuth({
    credentials: GOOGLE_CREDENTIALS,
    scopes: ['https://www.googleapis.com/auth/drive']
  });
  return google.drive({ version: 'v3', auth });
}

app.get('/', (req, res) => res.send('OK'));

app.post('/webhook', async (req, res) => {
  res.sendStatus(200);
  const events = req.body.events || [];
  for (const event of events) {
    try { await handleEvent(event); }
    catch (err) { console.error('handleEvent error:', err); }
  }
});

async function handleEvent(event) {
  const userId = event.source.userId;
  const replyToken = event.replyToken;

  if (event.type === 'postback') {
    const data = JSON.parse(event.postback.data);
    if (data.action === 'select_category') {
      sessions[userId].category = data.category;
      await replyConfirmation(replyToken, sessions[userId]);
      return;
    }
    if (data.action === 'confirm') { await processEntry(userId, replyToken); return; }
    if (data.action === 'edit') { await replyEditOptions(replyToken); return; }
    if (data.action === 'edit_category') { await replyCategoryButtons(replyToken); return; }
    if (data.action === 'edit_detail') {
      sessions[userId].waitingFor = 'detail';
      await replyText(replyToken, 'พิมพ์รายละเอียดใหม่ได้เลยครับ');
      return;
    }
    if (data.action === 'edit_date') {
      sessions[userId].waitingFor = 'date';
      await replyText(replyToken, 'พิมพ์วันที่ใหม่ได้เลยครับ\nเช่น 15/06');
      return;
    }
  }

  if (event.type === 'message' && event.message.type === 'image') {
    if (sessions[userId] && sessions[userId].imageIds) {
      sessions[userId].imageIds.push(event.message.id);
      await replyText(replyToken, `รับรูปที่ ${sessions[userId].imageIds.length} แล้วครับ ส่งรูปเพิ่มได้ หรือพิมพ์รายละเอียดได้เลย`);
      return;
    }
    sessions[userId] = { imageIds: [event.message.id], detail: null, category: null, date: new Date(), waitingFor: null };
    await replyText(replyToken, 'รับรูปสลิปแล้วครับ พิมพ์รายละเอียดได้เลย');
    return;
  }

  if (event.type === 'message' && event.message.type === 'text') {
    const text = event.message.text.trim();
    if (sessions[userId] && sessions[userId].waitingFor === 'detail') {
      sessions[userId].detail = text;
      sessions[userId].waitingFor = null;
      await replyConfirmation(replyToken, sessions[userId]);
      return;
    }
    if (sessions[userId] && sessions[userId].waitingFor === 'date') {
      sessions[userId].date = parseDate(text);
      sessions[userId].waitingFor = null;
      await replyConfirmation(replyToken, sessions[userId]);
      return;
    }
    if (sessions[userId] && sessions[userId].imageIds) {
      sessions[userId].detail = text;
      await replyCategoryButtons(replyToken);
      return;
    }
    await replyText(replyToken, 'ส่งรูปสลิปมาก่อนได้เลยครับ แล้วพิมพ์รายละเอียดตามมา');
  }
}

async function processEntry(userId, replyToken) {
  const session = sessions[userId];
  if (!session) { await replyText(replyToken, 'ไม่พบข้อมูล กรุณาเริ่มใหม่ครับ'); return; }

  try {
    const sheets = await getSheets();

    // 1. หาแถวแรกที่ column A ว่าง
    const checkResult = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: `${SHEET_NAME}!A:A`
    });
    const aValues = checkResult.data.values || [];
    let targetRow = 2;
    for (let i = 1; i < aValues.length; i++) {
      if (!aValues[i] || !aValues[i][0]) {
        targetRow = i + 1;
        break;
      }
      targetRow = i + 2;
    }

    // 2. เขียน A, C, D แยกกัน ไม่แตะ B เพื่อให้ formula คำนวณเอง
    const dateStr = formatDateSheet(session.date);
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: {
        valueInputOption: 'USER_ENTERED',
        data: [
          { range: `${SHEET_NAME}!A${targetRow}`, values: [[dateStr]] },
          { range: `${SHEET_NAME}!C${targetRow}`, values: [[session.category]] },
          { range: `${SHEET_NAME}!D${targetRow}`, values: [[session.detail]] }
        ]
      }
    });

    // 3. รอให้ formula คำนวณ B
    await sleep(3000);

    // 4. อ่านค่า B จริงจาก sheet
    const bResult = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: `${SHEET_NAME}!B${targetRow}`
    });
    const docNumber = bResult.data.values && bResult.data.values[0] ? bResult.data.values[0][0] : null;

    if (!docNumber) {
      await replyText(replyToken, `บันทึกแล้ว (แถว ${targetRow}) แต่ formula B ยังไม่คำนวณ กรุณาเช็ค Sheet ครับ`);
      delete sessions[userId];
      return;
    }

    const monthNum = ('0' + (session.date.getMonth() + 1)).slice(-2);
    const folderId = MONTH_FOLDERS[monthNum];

    if (!folderId) {
      await replyText(replyToken, `บันทึกแล้ว (${docNumber}) แต่ไม่พบโฟลเดอร์เดือน ${monthNum}`);
      delete sessions[userId];
      return;
    }

    // 5. สร้างโฟลเดอร์ชื่อ "06-RM-025"
    const folderName = `${monthNum}-${docNumber}`;
    const drive = await getDrive();
    const folderResponse = await drive.files.create({
      requestBody: { name: folderName, mimeType: 'application/vnd.google-apps.folder', parents: [folderId] },
      fields: 'id, webViewLink'
    });
    const newFolderId = folderResponse.data.id;
    const folderUrl = folderResponse.data.webViewLink;

    // 6. อัปโหลดรูปสลิป
    for (let i = 0; i < session.imageIds.length; i++) {
      const imageStream = await downloadLineImage(session.imageIds[i]);
      await drive.files.create({
        requestBody: { name: `slip_${i + 1}.jpg`, parents: [newFolderId] },
        media: { mimeType: 'image/jpeg', body: imageStream }
      });
    }

    // 7. ใส่ลิงก์ลง column M
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `${SHEET_NAME}!M${targetRow}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [[`=HYPERLINK("${folderUrl}","${folderName}")`]] }
    });

    delete sessions[userId];

    await replyText(replyToken,
      `✅ บันทึกแล้วครับ\n` +
      `เลขเอกสาร: ${docNumber}\n` +
      `ประเภท: ${session.category}\n` +
      `รายละเอียด: ${session.detail}\n` +
      `โฟลเดอร์: ${folderName}\n` +
      `รูปสลิป: ${session.imageIds.length} รูป`
    );

  } catch (err) {
    console.error('processEntry error:', err);
    await replyText(replyToken, `เกิดข้อผิดพลาด: ${err.message}`);
    delete sessions[userId];
  }
}

async function downloadLineImage(messageId) {
  const response = await axios.get(
    `https://api-data.line.me/v2/bot/message/${messageId}/content`,
    { headers: { Authorization: `Bearer ${LINE_TOKEN}` }, responseType: 'stream' }
  );
  return response.data;
}

async function replyCategoryButtons(replyToken) {
  const items = CATEGORIES.map(cat => ({
    type: 'action',
    action: {
      type: 'postback',
      label: cat.length > 20 ? cat.substring(0, 20) : cat,
      data: JSON.stringify({ action: 'select_category', category: cat }),
      displayText: cat
    }
  }));
  await sendReply(replyToken, [{ type: 'text', text: 'เลือกประเภทรายการได้เลยครับ', quickReply: { items } }]);
}

async function replyConfirmation(replyToken, session) {
  const dateStr = formatDateSheet(session.date);
  const text =
    `📋 ตรวจสอบอีกรอบ\n` +
    `ประเภท: ${session.category}\n` +
    `รายละเอียด: ${session.detail}\n` +
    `วันที่: ${dateStr}\n` +
    `รูปสลิป: ${session.imageIds.length} รูป\n\nถูกต้องไหมครับ?`;
  await sendReply(replyToken, [{
    type: 'text', text,
    quickReply: { items: [
      { type: 'action', action: { type: 'postback', label: '✅ ยืนยัน', data: JSON.stringify({ action: 'confirm' }), displayText: '✅ ยืนยัน' } },
      { type: 'action', action: { type: 'postback', label: '❌ แก้ไข', data: JSON.stringify({ action: 'edit' }), displayText: '❌ แก้ไข' } }
    ]}
  }]);
}

async function replyEditOptions(replyToken) {
  await sendReply(replyToken, [{
    type: 'text', text: 'อยากแก้ไขอะไรครับ?',
    quickReply: { items: [
      { type: 'action', action: { type: 'postback', label: 'เปลี่ยนประเภท', data: JSON.stringify({ action: 'edit_category' }), displayText: 'เปลี่ยนประเภท' } },
      { type: 'action', action: { type: 'postback', label: 'เปลี่ยนรายละเอียด', data: JSON.stringify({ action: 'edit_detail' }), displayText: 'เปลี่ยนรายละเอียด' } },
      { type: 'action', action: { type: 'postback', label: 'เปลี่ยนวันที่', data: JSON.stringify({ action: 'edit_date' }), displayText: 'เปลี่ยนวันที่' } }
    ]}
  }]);
}

async function replyText(replyToken, text) {
  await sendReply(replyToken, [{ type: 'text', text }]);
}

async function sendReply(replyToken, messages) {
  await axios.post('https://api.line.me/v2/bot/message/reply',
    { replyToken, messages },
    { headers: { Authorization: `Bearer ${LINE_TOKEN}` } }
  );
}

// format วันที่เป็น 19-Jun (เหมือน Sheet เดิม)
function formatDateSheet(date) {
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return `${date.getDate()}-${months[date.getMonth()]}`;
}

function parseDate(text) {
  const parts = text.split('/');
  if (parts.length === 2) {
    const d = parseInt(parts[0]);
    const m = parseInt(parts[1]) - 1;
    const date = new Date();
    date.setDate(d);
    date.setMonth(m);
    return date;
  }
  return new Date();
}

function sleep(ms) { return new Promise(resolve => setTimeout(resolve, ms)); }

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
