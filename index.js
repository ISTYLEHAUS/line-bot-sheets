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

// ===== à¹€à¸£à¸µà¸¢à¸ Apps Script à¸žà¸£à¹‰à¸­à¸¡à¸ˆà¸±à¸”à¸à¸²à¸£ redirect =====
async function callAppsScript(payload) {
  const appsScriptUrl = process.env.APPS_SCRIPT_URL;

  // Apps Script Web App redirect POST â†’ URL à¹ƒà¸«à¸¡à¹ˆà¸—à¸µà¹ˆà¸£à¸±à¸šà¹à¸„à¹ˆ GET
  // à¸§à¸´à¸˜à¸µà¸—à¸µà¹ˆà¹ƒà¸Šà¹‰à¸‡à¸²à¸™à¹„à¸”à¹‰à¸ˆà¸£à¸´à¸‡à¸„à¸·à¸­à¸ªà¹ˆà¸‡à¹€à¸›à¹‡à¸™ GET à¸žà¸£à¹‰à¸­à¸¡ payload à¹ƒà¸™ query string
  // à¹à¸•à¹ˆà¸£à¸¹à¸› base64 à¹ƒà¸«à¸à¹ˆà¹€à¸à¸´à¸™ URL limit â€” à¸ˆà¸¶à¸‡à¹à¸¢à¸ 2 à¸à¸£à¸“à¸µ

  const bodyStr = JSON.stringify(payload);

  try {
    // Step 1: POST à¹„à¸›à¸—à¸µà¹ˆ URL à¹€à¸”à¸´à¸¡ à¹„à¸¡à¹ˆ follow redirect
    const res = await axios.post(appsScriptUrl, bodyStr, {
      maxRedirects: 0,
      validateStatus: (s) => s < 500,
      headers: { 'Content-Type': 'text/plain' }
    });

    // Step 2: à¸–à¹‰à¸²à¹„à¸”à¹‰ redirect (301/302) à¹ƒà¸«à¹‰ POST à¸‹à¹‰à¸³à¹„à¸›à¸—à¸µà¹ˆ URL à¹ƒà¸«à¸¡à¹ˆ
    if ((res.status === 301 || res.status === 302) && res.headers.location) {
      const redirectUrl = res.headers.location;
      console.log('[callAppsScript] redirect to:', redirectUrl);
      const res2 = await axios.post(redirectUrl, bodyStr, {
        headers: { 'Content-Type': 'text/plain' },
        maxRedirects: 0,
        validateStatus: (s) => s < 500
      });
      return typeof res2.data === 'string' ? JSON.parse(res2.data) : res2.data;
    }

    return typeof res.data === 'string' ? JSON.parse(res.data) : res.data;

  } catch (err) {
    if (err.response && (err.response.status === 301 || err.response.status === 302) && err.response.headers.location) {
      const redirectUrl = err.response.headers.location;
      console.log('[callAppsScript] catch redirect to:', redirectUrl);
      const res2 = await axios.post(redirectUrl, bodyStr, {
        headers: { 'Content-Type': 'text/plain' },
        maxRedirects: 0,
        validateStatus: (s) => s < 500
      });
      return typeof res2.data === 'string' ? JSON.parse(res2.data) : res2.data;
    }
    throw err;
  }
}

// ===== Session Management (à¹€à¸à¹‡à¸šà¹ƒà¸™ Google Sheet) =====
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
        await replyText(replyToken, 'à¹€à¸‹à¸ªà¸Šà¸±à¸™à¸«à¸¡à¸”à¸­à¸²à¸¢à¸¸à¹à¸¥à¹‰à¸§à¸„à¸£à¸±à¸š à¸à¸£à¸¸à¸“à¸²à¸ªà¹ˆà¸‡à¸£à¸¹à¸›à¸ªà¸¥à¸´à¸›à¹ƒà¸«à¸¡à¹ˆà¸­à¸µà¸à¸„à¸£à¸±à¹‰à¸‡');
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
      await replyText(replyToken, 'à¸žà¸´à¸¡à¸žà¹Œà¸£à¸²à¸¢à¸¥à¸°à¹€à¸­à¸µà¸¢à¸”à¹ƒà¸«à¸¡à¹ˆà¹„à¸”à¹‰à¹€à¸¥à¸¢à¸„à¸£à¸±à¸š');
      return;
    }

    if (data.action === 'edit_date') {
      if (session) { session.waitingFor = 'date'; await saveSession(userId, session); }
      await replyText(replyToken, 'à¸žà¸´à¸¡à¸žà¹Œà¸§à¸±à¸™à¸—à¸µà¹ˆà¹ƒà¸«à¸¡à¹ˆà¹„à¸”à¹‰à¹€à¸¥à¸¢à¸„à¸£à¸±à¸š à¹€à¸Šà¹ˆà¸™ 15/06');
      return;
    }
  }

  if (event.type === 'message' && event.message.type === 'image') {
    let session = await getSession(userId);
    if (session && session.imageIds) {
      session.imageIds.push(event.message.id);
      await saveSession(userId, session);
      await replyText(replyToken, `à¸£à¸±à¸šà¸£à¸¹à¸›à¸—à¸µà¹ˆ ${session.imageIds.length} à¹à¸¥à¹‰à¸§à¸„à¸£à¸±à¸š à¸ªà¹ˆà¸‡à¸£à¸¹à¸›à¹€à¸žà¸´à¹ˆà¸¡à¹„à¸”à¹‰ à¸«à¸£à¸·à¸­à¸žà¸´à¸¡à¸žà¹Œà¸£à¸²à¸¢à¸¥à¸°à¹€à¸­à¸µà¸¢à¸”à¹„à¸”à¹‰à¹€à¸¥à¸¢`);
      return;
    }
    const newSession = {
      imageIds: [event.message.id],
      detail: null,
      category: null,
      date: new Date().toISOString(),
      waitingFor: null
    };
    await saveSession(userId, newSession);
    await replyText(replyToken, 'à¸£à¸±à¸šà¸£à¸¹à¸›à¸ªà¸¥à¸´à¸›à¹à¸¥à¹‰à¸§à¸„à¸£à¸±à¸š à¸žà¸´à¸¡à¸žà¹Œà¸£à¸²à¸¢à¸¥à¸°à¹€à¸­à¸µà¸¢à¸”à¹„à¸”à¹‰à¹€à¸¥à¸¢');
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

    if (session && session.waitingFor === 'date') {
      session.date = parseDate(text).toISOString();
      session.waitingFor = null;
      await saveSession(userId, session);
      await replyConfirmation(replyToken, session);
      return;
    }

    if (session && session.imageIds) {
      session.detail = text;
      await saveSession(userId, session);
      await replyCategoryButtons(replyToken);
      return;
    }

    await replyText(replyToken, 'à¸ªà¹ˆà¸‡à¸£à¸¹à¸›à¸ªà¸¥à¸´à¸›à¸¡à¸²à¸à¹ˆà¸­à¸™à¹„à¸”à¹‰à¹€à¸¥à¸¢à¸„à¸£à¸±à¸š à¹à¸¥à¹‰à¸§à¸žà¸´à¸¡à¸žà¹Œà¸£à¸²à¸¢à¸¥à¸°à¹€à¸­à¸µà¸¢à¸”à¸•à¸²à¸¡à¸¡à¸²');
  }
}

async function processEntry(userId, replyToken) {
  const session = await getSession(userId);
  if (!session) { await replyText(replyToken, 'à¹„à¸¡à¹ˆà¸žà¸šà¸‚à¹‰à¸­à¸¡à¸¹à¸¥ à¸à¸£à¸¸à¸“à¸²à¹€à¸£à¸´à¹ˆà¸¡à¹ƒà¸«à¸¡à¹ˆà¸„à¸£à¸±à¸š'); return; }

  try {
    const sheets = await getSheets();
    const sessionDate = new Date(session.date);

    // 1. à¸«à¸²à¹à¸–à¸§à¹à¸£à¸à¸—à¸µà¹ˆ column A à¸§à¹ˆà¸²à¸‡
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

    // 2. à¹€à¸‚à¸µà¸¢à¸™ A, C, D à¹à¸¢à¸à¸à¸±à¸™ à¹„à¸¡à¹ˆà¹à¸•à¸° B
    const dateStr = formatDateSheet(sessionDate);
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

    // 3. à¸£à¸­à¹ƒà¸«à¹‰ formula à¸„à¸³à¸™à¸§à¸“ B
    await sleep(3000);

    // 4. à¸­à¹ˆà¸²à¸™à¸„à¹ˆà¸² B à¸ˆà¸£à¸´à¸‡
    const bResult = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: `${SHEET_NAME}!B${targetRow}`
    });
    const docNumber = bResult.data.values && bResult.data.values[0] ? bResult.data.values[0][0] : null;

    if (!docNumber) {
      await replyText(replyToken, `à¸šà¸±à¸™à¸—à¸¶à¸à¹à¸¥à¹‰à¸§ (à¹à¸–à¸§ ${targetRow}) à¹à¸•à¹ˆ formula B à¸¢à¸±à¸‡à¹„à¸¡à¹ˆà¸„à¸³à¸™à¸§à¸“ à¸à¸£à¸¸à¸“à¸²à¹€à¸Šà¹‡à¸„ Sheet à¸„à¸£à¸±à¸š`);
      await deleteSession(userId);
      return;
    }

    const monthNum = ('0' + (sessionDate.getMonth() + 1)).slice(-2);
    const folderId = MONTH_FOLDERS[monthNum];

    if (!folderId) {
      await replyText(replyToken, `à¸šà¸±à¸™à¸—à¸¶à¸à¹à¸¥à¹‰à¸§ (${docNumber}) à¹à¸•à¹ˆà¹„à¸¡à¹ˆà¸žà¸šà¹‚à¸Ÿà¸¥à¹€à¸”à¸­à¸£à¹Œà¹€à¸”à¸·à¸­à¸™ ${monthNum} â€” à¹€à¸Šà¹‡à¸„ env FOLDER_${monthNum} à¸„à¸£à¸±à¸š`);
      await deleteSession(userId);
      return;
    }

    // 5. à¸ªà¸£à¹‰à¸²à¸‡à¹‚à¸Ÿà¸¥à¹€à¸”à¸­à¸£à¹Œà¸œà¹ˆà¸²à¸™ Apps Script
    const folderName = `${monthNum}-${docNumber}`;
    console.log(`[createFolder] folderName=${folderName} parentFolderId=${folderId}`);

    const createFolderData = await callAppsScript({
      action: 'createFolder',
      folderName: folderName,
      parentFolderId: folderId
    });
    console.log('[createFolder] response:', JSON.stringify(createFolderData));

    if (!createFolderData.success) {
      await replyText(replyToken, `à¸šà¸±à¸™à¸—à¸¶à¸à¹à¸¥à¹‰à¸§ (${docNumber}) à¹à¸•à¹ˆà¸ªà¸£à¹‰à¸²à¸‡à¹‚à¸Ÿà¸¥à¹€à¸”à¸­à¸£à¹Œà¹„à¸¡à¹ˆà¸ªà¸³à¹€à¸£à¹‡à¸ˆ: ${createFolderData.error}`);
      await deleteSession(userId);
      return;
    }

    const newFolderId = createFolderData.folderId;
    const folderUrl = createFolderData.folderUrl;

    // 6. à¸­à¸±à¸›à¹‚à¸«à¸¥à¸”à¸£à¸¹à¸›à¸ªà¸¥à¸´à¸›à¸œà¹ˆà¸²à¸™ Apps Script
    for (let i = 0; i < session.imageIds.length; i++) {
      const imageResponse = await axios.get(
        `https://api-data.line.me/v2/bot/message/${session.imageIds[i]}/content`,
        { headers: { Authorization: `Bearer ${LINE_TOKEN}` }, responseType: 'arraybuffer' }
      );
      const base64Data = Buffer.from(imageResponse.data).toString('base64');

      console.log(`[uploadImage] slip_${i + 1}.jpg â†’ folderId=${newFolderId}`);
      const uploadData = await callAppsScript({
        action: 'uploadImage',
        base64Data: base64Data,
        folderId: newFolderId,
        fileName: `slip_${i + 1}.jpg`
      });
      console.log(`[uploadImage] response:`, JSON.stringify(uploadData));
    }

    // 7. à¹ƒà¸ªà¹ˆà¸¥à¸´à¸‡à¸à¹Œà¸¥à¸‡ column M
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `${SHEET_NAME}!M${targetRow}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [[`=HYPERLINK("${folderUrl}","${folderName}")`]] }
    });

    await deleteSession(userId);

    await replyText(replyToken,
      `âœ… à¸šà¸±à¸™à¸—à¸¶à¸à¹à¸¥à¹‰à¸§à¸„à¸£à¸±à¸š\n` +
      `à¹€à¸¥à¸‚à¹€à¸­à¸à¸ªà¸²à¸£: ${docNumber}\n` +
      `à¸›à¸£à¸°à¹€à¸ à¸—: ${session.category}\n` +
      `à¸£à¸²à¸¢à¸¥à¸°à¹€à¸­à¸µà¸¢à¸”: ${session.detail}\n` +
      `à¹‚à¸Ÿà¸¥à¹€à¸”à¸­à¸£à¹Œ: ${folderName}\n` +
      `à¸£à¸¹à¸›à¸ªà¸¥à¸´à¸›: ${session.imageIds.length} à¸£à¸¹à¸›`
    );

  } catch (err) {
    console.error('processEntry error:', err);
    await replyText(replyToken, `à¹€à¸à¸´à¸”à¸‚à¹‰à¸­à¸œà¸´à¸”à¸žà¸¥à¸²à¸”: ${err.message}`);
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
  await sendReply(replyToken, [{ type: 'text', text: 'à¹€à¸¥à¸·à¸­à¸à¸›à¸£à¸°à¹€à¸ à¸—à¸£à¸²à¸¢à¸à¸²à¸£à¹„à¸”à¹‰à¹€à¸¥à¸¢à¸„à¸£à¸±à¸š', quickReply: { items } }]);
}

async function replyConfirmation(replyToken, session) {
  const sessionDate = new Date(session.date);
  const dateStr = formatDateSheet(sessionDate);
  const text =
    `ðŸ“‹ à¸•à¸£à¸§à¸ˆà¸ªà¸­à¸šà¸­à¸µà¸à¸£à¸­à¸š\n` +
    `à¸›à¸£à¸°à¹€à¸ à¸—: ${session.category}\n` +
    `à¸£à¸²à¸¢à¸¥à¸°à¹€à¸­à¸µà¸¢à¸”: ${session.detail}\n` +
    `à¸§à¸±à¸™à¸—à¸µà¹ˆ: ${dateStr}\n` +
    `à¸£à¸¹à¸›à¸ªà¸¥à¸´à¸›: ${session.imageIds.length} à¸£à¸¹à¸›\n\nà¸–à¸¹à¸à¸•à¹‰à¸­à¸‡à¹„à¸«à¸¡à¸„à¸£à¸±à¸š?`;
  await sendReply(replyToken, [{
    type: 'text', text,
    quickReply: { items: [
      { type: 'action', action: { type: 'postback', label: 'âœ… à¸¢à¸·à¸™à¸¢à¸±à¸™', data: JSON.stringify({ action: 'confirm' }), displayText: 'âœ… à¸¢à¸·à¸™à¸¢à¸±à¸™' } },
      { type: 'action', action: { type: 'postback', label: 'âŒ à¹à¸à¹‰à¹„à¸‚', data: JSON.stringify({ action: 'edit' }), displayText: 'âŒ à¹à¸à¹‰à¹„à¸‚' } }
    ]}
  }]);
}

async function replyEditOptions(replyToken) {
  await sendReply(replyToken, [{
    type: 'text', text: 'à¸­à¸¢à¸²à¸à¹à¸à¹‰à¹„à¸‚à¸­à¸°à¹„à¸£à¸„à¸£à¸±à¸š?',
    quickReply: { items: [
      { type: 'action', action: { type: 'postback', label: 'à¹€à¸›à¸¥à¸µà¹ˆà¸¢à¸™à¸›à¸£à¸°à¹€à¸ à¸—', data: JSON.stringify({ action: 'edit_category' }), displayText: 'à¹€à¸›à¸¥à¸µà¹ˆà¸¢à¸™à¸›à¸£à¸°à¹€à¸ à¸—' } },
      { type: 'action', action: { type: 'postback', label: 'à¹€à¸›à¸¥à¸µà¹ˆà¸¢à¸™à¸£à¸²à¸¢à¸¥à¸°à¹€à¸­à¸µà¸¢à¸”', data: JSON.stringify({ action: 'edit_detail' }), displayText: 'à¹€à¸›à¸¥à¸µà¹ˆà¸¢à¸™à¸£à¸²à¸¢à¸¥à¸°à¹€à¸­à¸µà¸¢à¸”' } },
      { type: 'action', action: { type: 'postback', label: 'à¹€à¸›à¸¥à¸µà¹ˆà¸¢à¸™à¸§à¸±à¸™à¸—à¸µà¹ˆ', data: JSON.stringify({ action: 'edit_date' }), displayText: 'à¹€à¸›à¸¥à¸µà¹ˆà¸¢à¸™à¸§à¸±à¸™à¸—à¸µà¹ˆ' } }
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
