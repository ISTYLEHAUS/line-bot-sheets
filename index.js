const express = require('express');
const axios = require('axios');
const { google } = require('googleapis');
const app = express();

app.use(express.json());

// ===== Config =====
const LINE_TOKEN = process.env.LINE_CHANNEL_ACCESS_TOKEN;
const SHEET_ID = process.env.SHEET_ID;
const SHEET_NAME = process.env.SHEET_NAME || 'AutoBot';
const GOOGLE_CREDENTIALS = JSON.parse(process.env.GOOGLE_CREDENTIALS);

// MONTH_FOLDERS - Folder ID ของแต่ละเดือนใน Google Drive
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

// 10 Category
const CATEGORIES = [
  "Production",
  "Raw Material",
  "Ads",
  "Salary",
  "Shipping",
  "Account",
  "Service",
  "Sales Revenue (Transfer)",
  "Other Income",
  "Office & Admin"
];

// Session เก็บสถานะการคุยของแต่ละ user (in-memory)
const sessions = {};

// ===== Google Sheets =====
async function getSheets() {
  const auth = new google.auth.GoogleAuth({
    credentials: GOOGLE_CREDENTIALS,
    scopes: ['https://www.googleapis.com/auth/spreadsheets']
  });
  return google.sheets({ version: 'v4', auth });
}

// ===== Google Drive =====
async function getDrive() {
  const auth = new google.auth.GoogleAuth({
    credentials: GOOGLE_CREDENTIALS,
    scopes: ['https://www.googleapis.com/auth/drive']
  });
  return google.drive({ version: 'v3', auth });
}

// ===== Endpoints =====
app.get('/', (req, res) => res.send('OK'));

app.post('/webhook', async (req, res) => {
  res.sendStatus(200); // ตอบ LINE ทันที

  const events = req.body.events || [];

  for (const event of events) {
    try {
      await handleEvent(event);
    } catch (err) {
      console.error('handleEvent error:', err);
    }
  }
});

// ===== Event Handler =====
async function handleEvent(event) {
  const userId = event.source.userId;
  const replyToken = event.replyToken;

  // ----- กดปุ่ม Postback -----
  if (event.type === 'postback') {
    const data = JSON.parse(event.postback.data);

    // เลือก category
    if (data.action === 'select_category') {
      sessions[userId].category = data.category;
      await replyConfirmation(replyToken, sessions[userId]);
      return;
    }

    // ยืนยัน
    if (data.action === 'confirm') {
      await processEntry(userId, replyToken);
      return;
    }

    // แก้ไข
    if (data.action === 'edit') {
      await replyEditOptions(replyToken);
      return;
    }

    // แก้ไขประเภท
    if (data.action === 'edit_category') {
      await replyCategoryButtons(replyToken);
      return;
    }

    // แก้ไขรายละเอียด
    if (data.action === 'edit_detail') {
      sessions[userId].waitingFor = 'detail';
      await replyText(replyToken, '📝 พิมพ์รายละเอียดใหม่ได้เลยครับ');
      return;
    }

    // แก้ไขวันที่
    if (data.action === 'edit_date') {
      sessions[userId].waitingFor = 'date';
      await replyText(replyToken, '📅 พิมพ์วันที่ใหม่ได้เลยครับ\nเช่น 15/06 หรือ 15 มิ.ย.');
      return;
    }
  }

  // ----- ส่งรูปภาพ -----
  if (event.type === 'message' && event.message.type === 'image') {
    // ถ้ามี session อยู่แล้วและกำลังรอรูปเพิ่ม
    if (sessions[userId] && sessions[userId].imageIds) {
      sessions[userId].imageIds.push(event.message.id);
      await replyText(replyToken, `รับรูปที่ ${sessions[userId].imageIds.length} แล้วครับ ✅`);
      return;
    }
    // เริ่ม session ใหม่ด้วยรูป
    sessions[userId] = {
      imageIds: [event.message.id],
      detail: null,
      category: null,
      date: new Date(),
      waitingFor: null
    };
    await replyText(replyToken, '📷 รับรูปสลิปแล้วครับ\nส่งรูปเพิ่มได้เลย หรือพิมพ์รายละเอียดได้เลยครับ');
    return;
  }

  // ----- ส่งข้อความ -----
  if (event.type === 'message' && event.message.type === 'text') {
    const text = event.message.text.trim();

    // กรณีกำลังรอ input แก้ไข
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

    // ถ้ามี session อยู่แล้ว (ส่งรูปมาแล้ว) → รับรายละเอียด
    if (sessions[userId] && sessions[userId].imageIds) {
      sessions[userId].detail = text;
      await replyCategoryButtons(replyToken);
      return;
    }

    // ถ้าไม่มี session → แจ้งให้ส่งรูปก่อน
    await replyText(replyToken, '📷 ส่งรูปสลิปมาก่อนได้เลยครับ แล้วพิมพ์รายละเอียดตามมา');
  }
}

// ===== Process Entry (บันทึกจริง) =====
async function processEntry(userId, replyToken) {
  const session = sessions[userId];
  if (!session) {
    await replyText(replyToken, '❌ ไม่พบข้อมูล กรุณาเริ่มใหม่ครับ');
    return;
  }

  try {
    const sheets = await getSheets();

    // 1. เขียน column A (วันที่) + C (category) + D (รายละเอียด) ลง Sheet
    const dateStr = formatDate(session.date);
    await sheets.spreadsheets.values.append({
      spreadsheetId: SHEET_ID,
      range: `${SHEET_NAME}!A:D`,
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [[dateStr, '', session.category, session.detail]]
        // B ว่างไว้ให้ formula คำนวณเอง
      }
    });

    // 2. อ่าน B จริงจาก sheet (รอ formula คำนวณ)
    await sleep(2000); // รอ 2 วินาทีให้ formula คำนวณ

    const result = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: `${SHEET_NAME}!A:D`
    });

    const rows = result.data.values || [];
    const lastRow = rows[rows.length - 1];
    const docNumber = lastRow[1]; // column B

    if (!docNumber) {
      await replyText(replyToken, '⚠️ บันทึกข้อมูลแล้ว แต่ไม่สามารถอ่านเลขเอกสารได้ กรุณาเช็ค Sheet ครับ');
      delete sessions[userId];
      return;
    }

    // 3. หาเดือนจากวันที่
    const monthNum = ('0' + (session.date.getMonth() + 1)).slice(-2);
    const folderId = MONTH_FOLDERS[monthNum];

    if (!folderId) {
      await replyText(replyToken, `⚠️ บันทึกแล้ว (${docNumber}) แต่ไม่พบโฟลเดอร์เดือน ${monthNum} กรุณาเพิ่ม Folder ID ครับ`);
      delete sessions[userId];
      return;
    }

    // 4. สร้างโฟลเดอร์ชื่อ "06-RM-025"
    const folderName = `${monthNum}-${docNumber}`;
    const drive = await getDrive();

    const folderResponse = await drive.files.create({
      requestBody: {
        name: folderName,
        mimeType: 'application/vnd.google-apps.folder',
        parents: [folderId]
      },
      fields: 'id, webViewLink'
    });

    const newFolderId = folderResponse.data.id;
    const folderUrl = folderResponse.data.webViewLink;

    // 5. อัปโหลดรูปสลิปทั้งหมดเข้าโฟลเดอร์
    for (let i = 0; i < session.imageIds.length; i++) {
      const imageBuffer = await downloadLineImage(session.imageIds[i]);
      await drive.files.create({
        requestBody: {
          name: `slip_${i + 1}.jpg`,
          parents: [newFolderId]
        },
        media: {
          mimeType: 'image/jpeg',
          body: imageBuffer
        }
      });
    }

    // 6. ใส่ลิงก์โฟลเดอร์ลง column M ของ sheet
    // หาแถวล่าสุดก่อน
    const lastRowNum = rows.length + 1;
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `${SHEET_NAME}!M${lastRowNum}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [[`=HYPERLINK("${folderUrl}","${folderName}")`]]
      }
    });

    // 7. เคลียร์ session
    delete sessions[userId];

    // 8. ตอบกลับ
    await replyText(replyToken,
      `✅ บันทึกแล้วครับ\n` +
      `เลขเอกสาร: ${docNumber}\n` +
      `ประเภท: ${session.category}\n` +
      `รายละเอียด: ${session.detail}\n` +
      `โฟลเดอร์: ${folderName}\n` +
      `อัปโหลดรูป: ${session.imageIds.length} รูป`
    );

  } catch (err) {
    console.error('processEntry error:', err);
    await replyText(replyToken, `❌ เกิดข้อผิดพลาด: ${err.message}`);
    delete sessions[userId];
  }
}

// ===== ดาวน์โหลดรูปจาก LINE =====
async function downloadLineImage(messageId) {
  const response = await axios.get(
    `https://api-data.line.me/v2/bot/message/${messageId}/content`,
    {
      headers: { Authorization: `Bearer ${LINE_TOKEN}` },
      responseType: 'stream'
    }
  );
  return response.data;
}

// ===== Reply Functions =====
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

  await sendReply(replyToken, [{
    type: 'text',
    text: '📁 เลือกประเภทรายการได้เลยครับ 👇',
    quickReply: { items }
  }]);
}

async function replyConfirmation(replyToken, session) {
  const dateStr = formatDate(session.date);
  const text =
    `📋 ตรวจสอบอีกรอบ\n` +
    `ประเภท: ${session.category}\n` +
    `รายละเอียด: ${session.detail}\n` +
    `วันที่: ${dateStr}\n` +
    `รูปสลิป: ${session.imageIds.length} รูป\n\n` +
    `ถูกต้องไหมครับ?`;

  await sendReply(replyToken, [{
    type: 'text',
    text,
    quickReply: {
      items: [
        {
          type: 'action',
          action: {
            type: 'postback',
            label: '✅ ยืนยัน',
            data: JSON.stringify({ action: 'confirm' }),
            displayText: '✅ ยืนยัน'
          }
        },
        {
          type: 'action',
          action: {
            type: 'postback',
            label: '❌ แก้ไข',
            data: JSON.stringify({ action: 'edit' }),
            displayText: '❌ แก้ไข'
          }
        }
      ]
    }
  }]);
}

async function replyEditOptions(replyToken) {
  await sendReply(replyToken, [{
    type: 'text',
    text: '✏️ อยากแก้ไขอะไรครับ?',
    quickReply: {
      items: [
        {
          type: 'action',
          action: {
            type: 'postback',
            label: '📁 เปลี่ยนประเภท',
            data: JSON.stringify({ action: 'edit_category' }),
            displayText: 'เปลี่ยนประเภท'
          }
        },
        {
          type: 'action',
          action: {
            type: 'postback',
            label: '📝 เปลี่ยนรายละเอียด',
            data: JSON.stringify({ action: 'edit_detail' }),
            displayText: 'เปลี่ยนรายละเอียด'
          }
        },
        {
          type: 'action',
          action: {
            type: 'postback',
            label: '📅 เปลี่ยนวันที่',
            data: JSON.stringify({ action: 'edit_date' }),
            displayText: 'เปลี่ยนวันที่'
          }
        }
      ]
    }
  }]);
}

async function replyText(replyToken, text) {
  await sendReply(replyToken, [{ type: 'text', text }]);
}

async function sendReply(replyToken, messages) {
  await axios.post('https://api.line.me/v2/bot/message/reply', {
    replyToken,
    messages
  }, {
    headers: { Authorization: `Bearer ${LINE_TOKEN}` }
  });
}

// ===== Helpers =====
function formatDate(date) {
  const months = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.',
    'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
  return `${date.getDate()} ${months[date.getMonth()]} ${date.getFullYear() + 543}`;
}

function parseDate(text) {
  // รองรับรูปแบบ 15/06 หรือ 15 มิ.ย.
  const parts = text.split('/');
  if (parts.length === 2) {
    const d = parseInt(parts[0]);
    const m = parseInt(parts[1]) - 1;
    const date = new Date();
    date.setDate(d);
    date.setMonth(m);
    return date;
  }
  return new Date(); // ถ้า parse ไม่ได้ใช้วันนี้
}

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

// ===== Start Server =====
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
