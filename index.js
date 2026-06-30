const express = require('express');
const axios = require('axios');
const { google } = require('googleapis');
const app = express();

app.use(express.json());

const LINE_TOKEN = process.env.LINE_CHANNEL_ACCESS_TOKEN;
const SHEET_ID = process.env.SHEET_ID;
const SHEET_NAME = process.env.SHEET_NAME || 'AutoBot';
const SESSION_SHEET = 'Sessions';
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

async function getSheets() {
  const auth = new google.auth.GoogleAuth({
    credentials: GOOGLE_CREDENTIALS,
    scopes: ['https://www.googleapis.com/auth/spreadsheets']
  });
  return google.sheets({ version: 'v4', auth });
}

// ===== เรียก Apps Script =====
async function callAppsScript(payload) {
  const appsScriptUrl = process.env.APPS_SCRIPT_URL;
  const bodyStr = JSON.stringify(payload);

  // node-fetch follows redirect correctly with POST
  const { default: fetch } = await import('node-fetch');

  const res = await fetch(appsScriptUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: bodyStr,
    redirect: 'follow'
  });

  const text = await res.text();
  console.log('[callAppsScript] status:', res.status, 'body:', text.substring(0, 200));

  try {
    return JSON.parse(text);
  } catch(e) {
    throw new Error('Apps Script returned non-JSON: ' + text.substring(0, 300));
  }
}

// ===== Session Management (เก็บใน Google Sheet) =====
async function getSession(userId) {
  try {
    const sheets = await getSheets();
    const result = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: `${SESSION_SHEET}!A:B`
    });
    const rows = result.data.values || [];
    for (const row of rows) {
      if (row[0] === userId) {
        return JSON.parse(row[1]);
      }
    }
    return null;
  } catch (err) {
    console.error('getSession error:', err);
    return null;
  }
}

async function saveSession(userId, session) {
  try {
    const sheets = await getSheets();
    const result = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: `${SESSION_SHEET}!A:A`
    });
    const rows = result.data.values || [];
    let rowIndex = -1;
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][0] === userId) { rowIndex = i + 1; break; }
    }

    const sessionStr = JSON.stringify(session);
    if (rowIndex > 0) {
      await sheets.spreadsheets.values.update({
        spreadsheetId: SHEET_ID,
        range: `${SESSION_SHEET}!A${rowIndex}:B${rowIndex}`,
        valueInputOption: 'RAW',
        requestBody: { values: [[userId, sessionStr]] }
      });
    } else {
      await sheets.spreadsheets.values.append({
        spreadsheetId: SHEET_ID,
        range: `${SESSION_SHEET}!A:B`,
        valueInputOption: 'RAW',
        requestBody: { values: [[userId, sessionStr]] }
      });
    }
  } catch (err) {
    console.error('saveSession error:', err);
  }
}

async function deleteSession(userId) {
  try {
    const sheets = await getSheets();
    const result = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: `${SESSION_SHEET}!A:A`
    });
    const rows = result.data.values || [];
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][0] === userId) {
        await sheets.spreadsheets.values.clear({
          spreadsheetId: SHEET_ID,
          range: `${SESSION_SHEET}!A${i + 1}:B${i + 1}`
        });
        break;
      }
    }
  } catch (err) {
    console.error('deleteSession error:', err);
  }
}

// ===== Endpoints =====
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
    const session = await getSession(userId);

    if (data.action === 'select_category') {
      if (!session) {
        await replyText(replyToken, 'เซสชันหมดอายุแล้วครับ กรุณาส่งรูปสลิปใหม่อีกครั้ง');
        return;
      }
      session.category = data.category;
      await saveSession(userId, session);
      await replyConfirmation(replyToken, session);
      return;
    }

    if (data.action === 'confirm') { await processEntry(userId, replyToken); return; }
    if (data.action === 'edit') { await replyEditOptions(replyToken); return; }
    if (data.action === 'edit_category') { await replyCategoryButtons(replyToken); return; }

    if (data.action === 'edit_detail') {
      if (session) { session.waitingFor = 'detail'; await saveSession(userId, session); }
      await replyText(replyToken, 'พิมพ์รายละเอียดใหม่ได้เลยครับ');
      return;
    }

    if (data.action === 'edit_date') {
      if (session) { session.waitingFor = 'date'; await saveSession(userId, session); }
      await replyText(replyToken, 'พิมพ์วันที่ใหม่ได้เลยครับ เช่น 15/06');
      return;
    }

    if (data.action === 'edit_amount') {
      if (session) { session.waitingFor = 'amount'; await saveSession(userId, session); }
      await replyText(replyToken, 'พิมพ์ยอดเงินใหม่ได้เลยครับ เช่น 1500 หรือ 1500.50');
      return;
    }
  }

  if (event.type === 'message' && event.message.type === 'image') {
    let session = await getSession(userId);
    if (session && session.imageIds) {
      session.imageIds.push(event.message.id);
      await saveSession(userId, session);
      await replyText(replyToken, `รับรูปที่ ${session.imageIds.length} แล้วครับ ส่งรูปเพิ่มได้ หรือพิมพ์รายละเอียดได้เลย`);
      return;
    }
    const newSession = {
      imageIds: [event.message.id],
      detail: null,
      amount: null,
      category: null,
      date: new Date().toISOString(),
      waitingFor: null
    };
    await saveSession(userId, newSession);
    await replyText(replyToken, 'รับรูปสลิปแล้วครับ พิมพ์รายละเอียดได้เลย');
    return;
  }

  if (event.type === 'message' && event.message.type === 'text') {
    const text = event.message.text.trim();
    const session = await getSession(userId);

    if (session && session.waitingFor === 'detail') {
      session.detail = text;
      session.waitingFor = null;
      await saveSession(userId, session);
      await replyConfirmation(replyToken, session);
      return;
    }

    if (session && session.waitingFor === 'amount') {
      const amount = parseAmount(text);
      if (amount === null) {
        await replyText(replyToken, '⚠️ กรุณาพิมพ์ยอดเงินเป็นตัวเลขเท่านั้นครับ เช่น 1500 หรือ 1500.50');
        return;
      }
      session.amount = amount;
      session.waitingFor = null;
      await saveSession(userId, session);
      await replyConfirmation(replyToken, session);
      return;
    }

    if (session && session.waitingFor === 'date') {
      session.date = parseDate(text).toISOString();
      session.waitingFor = null;
      await saveSession(userId, session);
      await replyConfirmation(replyToken, session);
      return;
    }

    if (session && session.imageIds && session.detail === null) {
      session.detail = text;
      session.waitingFor = 'amount';
      await saveSession(userId, session);
      await replyText(replyToken, 'ใส่ยอดเงินด้วยครับ เช่น 1500 หรือ 1500.50');
      return;
    }

    if (session && session.imageIds && session.amount === null) {
      const amount = parseAmount(text);
      if (amount === null) {
        await replyText(replyToken, '⚠️ กรุณาพิมพ์ยอดเงินเป็นตัวเลขเท่านั้นครับ เช่น 1500 หรือ 1500.50');
        return;
      }
      session.amount = amount;
      await saveSession(userId, session);
      await replyCategoryButtons(replyToken);
      return;
    }

    await replyText(replyToken, 'ส่งรูปสลิปมาก่อนได้เลยครับ แล้วพิมพ์รายละเอียดตามมา');
  }
}

async function processEntry(userId, replyToken) {
  const session = await getSession(userId);
  if (!session) { await replyText(replyToken, 'ไม่พบข้อมูล กรุณาเริ่มใหม่ครับ'); return; }

  try {
    const sheets = await getSheets();
    const sessionDate = new Date(session.date);

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

    // 2. เขียน A, C, D, H แยกกัน ไม่แตะ B
    const dateStr = formatDateSheet(sessionDate);
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: {
        valueInputOption: 'USER_ENTERED',
        data: [
          { range: `${SHEET_NAME}!A${targetRow}`, values: [[dateStr]] },
          { range: `${SHEET_NAME}!C${targetRow}`, values: [[session.category]] },
          { range: `${SHEET_NAME}!D${targetRow}`, values: [[session.detail]] },
          { range: `${SHEET_NAME}!H${targetRow}`, values: [[session.amount]] }
        ]
      }
    });

    // 3. รอให้ formula คำนวณ B
    await sleep(3000);

    // 4. อ่านค่า B จริง
    const bResult = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: `${SHEET_NAME}!B${targetRow}`
    });
    const docNumber = bResult.data.values && bResult.data.values[0] ? bResult.data.values[0][0] : null;

    if (!docNumber) {
      await replyText(replyToken, `บันทึกแล้ว (แถว ${targetRow}) แต่ formula B ยังไม่คำนวณ กรุณาเช็ค Sheet ครับ`);
      await deleteSession(userId);
      return;
    }

    const monthNum = ('0' + (sessionDate.getMonth() + 1)).slice(-2);
    const folderId = MONTH_FOLDERS[monthNum];

    if (!folderId) {
      await replyText(replyToken, `บันทึกแล้ว (${docNumber}) แต่ไม่พบโฟลเดอร์เดือน ${monthNum} — เช็ค env FOLDER_${monthNum} ครับ`);
      await deleteSession(userId);
      return;
    }

    // 5. สร้างโฟลเดอร์ผ่าน Apps Script
    const folderName = `${monthNum}-${docNumber}`;
    console.log(`[createFolder] folderName=${folderName} parentFolderId=${folderId}`);

    const createFolderData = await callAppsScript({
      action: 'createFolder',
      folderName: folderName,
      parentFolderId: folderId
    });
    console.log('[createFolder] response:', JSON.stringify(createFolderData));

    if (!createFolderData.success) {
      await replyText(replyToken, `บันทึกแล้ว (${docNumber}) แต่สร้างโฟลเดอร์ไม่สำเร็จ: ${createFolderData.error}`);
      await deleteSession(userId);
      return;
    }

    const newFolderId = createFolderData.folderId;
    const folderUrl = createFolderData.folderUrl;

    // 6. อัปโหลดรูปสลิปผ่าน Apps Script
    for (let i = 0; i < session.imageIds.length; i++) {
      const imageResponse = await axios.get(
        `https://api-data.line.me/v2/bot/message/${session.imageIds[i]}/content`,
        { headers: { Authorization: `Bearer ${LINE_TOKEN}` }, responseType: 'arraybuffer' }
      );
      const base64Data = Buffer.from(imageResponse.data).toString('base64');

      console.log(`[uploadImage] slip_${i + 1}.jpg → folderId=${newFolderId}`);
      const uploadData = await callAppsScript({
        action: 'uploadImage',
        base64Data: base64Data,
        folderId: newFolderId,
        fileName: `slip_${i + 1}.jpg`
      });
      console.log(`[uploadImage] response:`, JSON.stringify(uploadData));
    }

    // 7. ใส่ลิงก์ลง column M
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `${SHEET_NAME}!M${targetRow}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [[`=HYPERLINK("${folderUrl}","${folderName}")`]] }
    });

    await deleteSession(userId);

    const safeMsg = [
      '✅ บันทึกแล้วครับ',
      'เลขเอกสาร: ' + String(docNumber || '').slice(0, 50),
      'ประเภท: ' + String(session.category || '').slice(0, 50),
      'รายละเอียด: ' + String(session.detail || '').slice(0, 100),
      'ยอดเงิน: ' + Number(session.amount || 0).toLocaleString('th-TH') + ' บาท',
      'โฟลเดอร์: ' + String(folderName || '').slice(0, 50),
      'รูปสลิป: ' + session.imageIds.length + ' รูป'
    ].join('\n');
    await replyText(replyToken, safeMsg);

  } catch (err) {
    console.error('processEntry error:', err);
    const errMsg = err && err.message ? String(err.message).slice(0, 150) : 'unknown';
    await replyText(replyToken, 'เกิดข้อผิดพลาด: ' + errMsg).catch(() => {});
    await deleteSession(userId);
  }
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
  const sessionDate = new Date(session.date);
  const dateStr = formatDateSheet(sessionDate);
  const amountStr = session.amount !== null && session.amount !== undefined
    ? Number(session.amount).toLocaleString('th-TH', { minimumFractionDigits: 0, maximumFractionDigits: 2 })
    : '-';
  const text =
    `📋 ตรวจสอบอีกรอบ\n` +
    `ประเภท: ${session.category}\n` +
    `รายละเอียด: ${session.detail}\n` +
    `ยอดเงิน: ${amountStr} บาท\n` +
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
      { type: 'action', action: { type: 'postback', label: 'เปลี่ยนยอดเงิน', data: JSON.stringify({ action: 'edit_amount' }), displayText: 'เปลี่ยนยอดเงิน' } },
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

function formatDateSheet(date) {
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return `${date.getDate()}-${months[date.getMonth()]}`;
}

function parseAmount(text) {
  const cleaned = text.replace(/,/g, '').trim();
  if (!/^\d+(\.\d{1,2})?$/.test(cleaned)) return null;
  const num = parseFloat(cleaned);
  if (isNaN(num) || num <= 0) return null;
  return num;
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
