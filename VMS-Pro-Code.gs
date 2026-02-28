// ============================================================
// VMS Pro — Google Apps Script Backend
// วาง Code นี้ใน Extensions > Apps Script แล้ว Deploy เป็น Web App
// ============================================================

// ---- CONFIG: แก้ค่าเหล่านี้ก่อน Deploy ----
const CONFIG = {
  SPREADSHEET_ID: '15_heJ1TW1mycodEIh-gWRcRQ_KpVGrnXZdwoh4wSwGk',    // ID ของ Google Sheet
  CALENDAR_ID:    'c_896db41e6cf6bb93d0232de5e07fb2f174389fb601ec2e5810a8f6fabdd545ee@group.calendar.google.com',      // เช่น xxx@gmail.com หรือ calendar ID
  SENDER_NAME:    'ระบบจัดการยานพาหนะ VMS Pro',
  COMPANY_NAME:   '  COMPANY_NAME:   'สำนักงานสถาบันฯ - งานระบบกายภาพและโสตทัศนูปกรณ์ ',
',
  APP_URL:        'https://yourname.github.io/vms/', // URL ของ Web App
};

// ---- Sheet Names ----
const SHEETS = {
  bookings: 'Bookings',
  vehicles: 'Vehicles',
  drivers:  'Drivers',
  users:    'Users',
  trips:    'Trips',
  logs:     'EmailLogs',
};

// ============================================================
// WEB APP ENTRY POINT (CORS-friendly)
// ============================================================
function doPost(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  try {
    const payload = JSON.parse(e.postData.contents);
    let result = {};
    switch(payload.action) {
      case 'write':       result = writeData(payload.table, payload.data); break;
      case 'read':        result = readData(payload.table); break;
      case 'sendEmail':   result = sendEmailNotification(payload.to, payload.subject, payload.body, payload.data); break;
      case 'syncCal':     result = createCalendarEvent(payload.event); break;
      case 'testConn':    result = { ok: true, message: 'Connected!' }; break;
      default:            result = { ok: false, error: 'Unknown action' };
    }
    output.setContent(JSON.stringify({ ok: true, ...result }));
  } catch(err) {
    output.setContent(JSON.stringify({ ok: false, error: err.message }));
  }
  return output;
}

function doGet(e) {
  const table = e.parameter.table;
  if(table) {
    const data = readData(table);
    return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput('VMS Pro API v2.0 — OK');
}

// ============================================================
// GOOGLE SHEETS READ / WRITE
// ============================================================
function getSheet(name) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName(name);
  if(!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function writeData(tableName, rows) {
  const sheetName = SHEETS[tableName] || tableName;
  const sheet = getSheet(sheetName);
  if(!rows || rows.length === 0) return { written: 0 };
  
  // Clear and rewrite (Full sync)
  sheet.clearContents();
  const headers = Object.keys(rows[0]);
  sheet.appendRow(headers);
  rows.forEach(row => sheet.appendRow(headers.map(h => {
    const val = row[h];
    return (typeof val === 'object') ? JSON.stringify(val) : (val ?? '');
  })));
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, headers.length);
  return { written: rows.length };
}

function readData(tableName) {
  const sheetName = SHEETS[tableName] || tableName;
  const sheet = getSheet(sheetName);
  const values = sheet.getDataRange().getValues();
  if(values.length < 2) return { data: [] };
  const headers = values[0];
  const data = values.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
  return { data };
}

// ============================================================
// EMAIL NOTIFICATIONS (Gmail API)
// ============================================================
function sendEmailNotification(to, subject, bodyText, data) {
  try {
    if(!to || !to.includes('@')) return { sent: false, reason: 'Invalid email' };
    
    // Check duplicate (prevent double-send)
    if(data && isEmailAlreadySent(data.bookingId, data.eventType)) {
      return { sent: false, reason: 'Duplicate — already sent' };
    }
    
    const htmlBody = buildEmailTemplate(subject, bodyText, data);
    GmailApp.sendEmail(to, subject, bodyText, {
      htmlBody: htmlBody,
      name: CONFIG.SENDER_NAME,
    });
    
    // Log sent email
    if(data) logEmail(data.bookingId, data.eventType, to);
    return { sent: true };
  } catch(e) {
    Logger.log('Email error: ' + e.message);
    return { sent: false, error: e.message };
  }
}

function isEmailAlreadySent(bookingId, eventType) {
  if(!bookingId || !eventType) return false;
  const sheet = getSheet(SHEETS.logs);
  const values = sheet.getDataRange().getValues();
  return values.some(row => row[0] === bookingId && row[1] === eventType);
}

function logEmail(bookingId, eventType, to) {
  const sheet = getSheet(SHEETS.logs);
  sheet.appendRow([bookingId, eventType, to, new Date().toISOString(), 'Sent']);
}

function buildEmailTemplate(title, bodyText, data) {
  const d = data || {};
  return `<!DOCTYPE html>
<html>
<head>
<style>
  body{font-family:'Helvetica Neue',Arial,sans-serif;line-height:1.6;color:#333;background:#f5f5f5;margin:0;padding:20px}
  .wrap{max-width:600px;margin:0 auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,.1)}
  .head{background:linear-gradient(135deg,#0ea5e9,#0284c7);color:#fff;padding:28px 32px;text-align:center}
  .head h1{font-size:20px;margin:0 0 6px;font-weight:700}
  .head p{margin:0;opacity:.85;font-size:14px}
  .logo{font-size:36px;margin-bottom:12px}
  .body{padding:28px 32px}
  .info-box{background:#f8fafc;border-radius:8px;padding:18px;margin:18px 0;border-left:4px solid #0ea5e9}
  .info-row{display:flex;gap:16px;margin-bottom:10px;align-items:flex-start}
  .info-row .icon{font-size:16px;flex-shrink:0;width:22px}
  .info-row .label{color:#64748b;font-size:13px;width:100px;flex-shrink:0;font-weight:600}
  .info-row .val{color:#0f172a;font-size:14px;font-weight:700}
  .status-chip{display:inline-block;padding:6px 16px;border-radius:20px;font-weight:700;font-size:14px;margin-bottom:16px}
  .status-approved{background:#dcfce7;color:#16a34a}
  .status-dispatched{background:#cffafe;color:#0e7490}
  .status-rejected{background:#fee2e2;color:#dc2626}
  .btn{display:inline-block;background:#0ea5e9;color:#fff;padding:12px 28px;border-radius:8px;text-decoration:none;font-weight:700;font-size:14px;margin-top:20px}
  .footer{background:#f1f5f9;padding:16px 32px;text-align:center;font-size:12px;color:#94a3b8}
  hr{border:none;border-top:1px solid #e2e8f0;margin:20px 0}
</style>
</head>
<body>
<div class="wrap">
  <div class="head">
    <div class="logo">🚗</div>
    <h1>${title}</h1>
    <p>${CONFIG.COMPANY_NAME} — ระบบจัดการยานพาหนะ VMS Pro</p>
  </div>
  <div class="body">
    ${d.status ? `<span class="status-chip status-${d.status==='อนุมัติ'||d.status==='จัดรถแล้ว'?'dispatched':d.status==='ไม่อนุมัติ'?'rejected':'approved'}">${getStatusEmoji(d.status)} ${d.status}</span>` : ''}
    <p>${bodyText}</p>
    ${d.bookingId ? `
    <div class="info-box">
      <div class="info-row"><span class="icon">📋</span><span class="label">เลขที่คำขอ</span><span class="val">${d.bookingId}</span></div>
      ${d.destination?`<div class="info-row"><span class="icon">📍</span><span class="label">ปลายทาง</span><span class="val">${d.destination}</span></div>`:''}
      ${d.datetime?`<div class="info-row"><span class="icon">🕒</span><span class="label">วัน-เวลา</span><span class="val">${d.datetime}</span></div>`:''}
      ${d.plate?`<div class="info-row"><span class="icon">🚗</span><span class="label">ทะเบียนรถ</span><span class="val">${d.plate} (${d.vehicleType||''})</span></div>`:''}
      ${d.driverName?`<div class="info-row"><span class="icon">👨‍✈️</span><span class="label">พนักงานขับรถ</span><span class="val">${d.driverName}</span></div>`:''}
      ${d.driverPhone?`<div class="info-row"><span class="icon">📞</span><span class="label">เบอร์ติดต่อ</span><span class="val"><a href="tel:${d.driverPhone}">${d.driverPhone}</a></span></div>`:''}
      ${d.meetPoint?`<div class="info-row"><span class="icon">🏁</span><span class="label">จุดนัดพบ</span><span class="val">${d.meetPoint}</span></div>`:''}
      ${d.rejectReason?`<div class="info-row"><span class="icon">❌</span><span class="label">เหตุผล</span><span class="val" style="color:#dc2626">${d.rejectReason}</span></div>`:''}
    </div>
    <a href="${CONFIG.APP_URL}?id=${d.bookingId}" class="btn">🌐 ดูรายละเอียดบนเว็บไซต์</a>
    ` : ''}
  </div>
  <div class="footer">
    ${CONFIG.COMPANY_NAME} — ระบบจัดการยานพาหนะอัตโนมัติ<br>
    อีเมลนี้ส่งโดยระบบอัตโนมัติ กรุณาอย่าตอบกลับ<br>
    <small>© 2025 VMS Pro — Powered by Google Apps Script</small>
  </div>
</div>
</body>
</html>`;
}

function getStatusEmoji(s) {
  const m = {'อนุมัติ':'✅','จัดรถแล้ว':'🚗','ไม่อนุมัติ':'❌','รออนุมัติ':'⏳','เสร็จสิ้น':'🏁','กำลังเดินทาง':'🛣️'};
  return m[s]||'📋';
}

// ============================================================
// GOOGLE CALENDAR INTEGRATION
// ============================================================
function createCalendarEvent(evt) {
  try {
    const cal = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID) || CalendarApp.getDefaultCalendar();
    const start = new Date(evt.start);
    const end = new Date(evt.end);
    const event = cal.createEvent(
      `🚗 [${evt.bookingId}] ${evt.requester} → ${evt.destination}`,
      start, end,
      {
        description: [
          `คำขอ: ${evt.bookingId}`,
          `ผู้ขอ: ${evt.requester}`,
          `ปลายทาง: ${evt.destination}`,
          `ทะเบียนรถ: ${evt.plate}`,
          `คนขับ: ${evt.driverName} (${evt.driverPhone})`,
          `วัตถุประสงค์: ${evt.purpose}`,
        ].join('\n'),
        guests: evt.guestEmails || '',
        sendInvites: true,
        colorId: '2', // Green = จัดรถแล้ว
      }
    );
    // Add reminder 60 minutes before
    event.addEmailReminder(60);
    event.addPopupReminder(30);
    return { ok: true, eventId: event.getId(), eventUrl: event.getEditEventUrl() };
  } catch(e) {
    Logger.log('Calendar error: ' + e.message);
    return { ok: false, error: e.message };
  }
}

// ============================================================
// AUTO TRIGGERS (ตั้งค่าใน Triggers ของ Apps Script)
// ============================================================

// trigger ทุกวัน 08:00 — ตรวจสอบการแจ้งเตือน
function dailyMorningCheck() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  checkVehicleExpiry(ss);
  checkTodayTrips(ss);
}

function checkVehicleExpiry(ss) {
  const sheet = ss.getSheetByName(SHEETS.vehicles);
  if(!sheet) return;
  const data = sheet.getDataRange().getValues();
  const today = new Date();
  const headers = data[0];
  const prbIdx = headers.indexOf('prb');
  const insIdx = headers.indexOf('ins');
  const plateIdx = headers.indexOf('plate');
  
  data.slice(1).forEach(row => {
    const plate = row[plateIdx];
    [['พรบ.', row[prbIdx]], ['ประกัน', row[insIdx]]].forEach(([type, dateStr]) => {
      if(!dateStr) return;
      const expDate = new Date(dateStr);
      const daysLeft = Math.ceil((expDate - today) / 86400000);
      if(daysLeft <= 30 && daysLeft >= 0) {
        const admins = getAdminEmails(ss);
        admins.forEach(email => {
          sendEmailNotification(email, 
            `⚠️ แจ้งเตือน: ${type} รถ ${plate} หมดอายุใน ${daysLeft} วัน`,
            `กรุณาดำเนินการต่ออายุ ${type} สำหรับรถทะเบียน ${plate} ที่จะหมดอายุในวันที่ ${dateStr}`,
            { bookingId: `MAINT-${plate}`, eventType: `${type}_warn_${dateStr}` }
          );
        });
      }
    });
  });
}

function checkTodayTrips(ss) {
  const sheet = ss.getSheetByName(SHEETS.bookings);
  if(!sheet) return;
  const data = sheet.getDataRange().getValues();
  const today = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
  const headers = data[0];
  const startIdx = headers.indexOf('start');
  const statusIdx = headers.indexOf('status');
  const driverEmailIdx = headers.indexOf('driverEmail');
  const idIdx = headers.indexOf('id');
  const toIdx = headers.indexOf('to');
  
  data.slice(1).forEach(row => {
    const startStr = String(row[startIdx]);
    if(startStr.startsWith(today) && row[statusIdx] === 'จัดรถแล้ว') {
      const driverEmail = row[driverEmailIdx];
      if(driverEmail) {
        sendEmailNotification(driverEmail,
          `🚗 แจ้งเตือน: มีงานวันนี้ — ${row[idIdx]}`,
          `คุณมีงานขับรถวันนี้ไปยัง ${row[toIdx]} กรุณาตรวจสอบรายละเอียดในระบบ`,
          { bookingId: row[idIdx], eventType: 'morning_reminder' }
        );
      }
    }
  });
}

function getAdminEmails(ss) {
  const sheet = ss.getSheetByName(SHEETS.users);
  if(!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const roleIdx = headers.indexOf('role');
  const emailIdx = headers.indexOf('email');
  return data.slice(1)
    .filter(row => ['admin','manager'].includes(row[roleIdx]))
    .map(row => row[emailIdx]);
}

// ============================================================
// onEdit TRIGGER — Auto-send email when status changes in Sheet
// ============================================================
function onEditTrigger(e) {
  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();
  if(sheet.getName() !== SHEETS.bookings || row <= 1) return;
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const statusCol = headers.indexOf('status') + 1;
  const emailSentCol = headers.indexOf('emailSent') + 1;
  
  if(col !== statusCol) return;
  
  const newStatus = e.value;
  const emailSent = sheet.getRange(row, emailSentCol).getValue();
  if(emailSent === newStatus) return; // Already sent
  
  // Get full row data
  const rowData = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
  const obj = {};
  headers.forEach((h, i) => obj[h] = rowData[i]);
  
  const emailData = {
    bookingId: obj.id,
    destination: obj.to,
    datetime: obj.start,
    plate: obj.plate || '',
    driverName: obj.driverName || '',
    driverPhone: obj.driverPhone || '',
    status: newStatus,
    eventType: newStatus,
  };
  
  if(newStatus === 'อนุมัติ') {
    sendEmailNotification(obj.requesterEmail || '',
      `[อนุมัติ] คำขอใช้รถ ${obj.id}`,
      `คำขอของคุณเดินทางไป ${obj.to} ได้รับการอนุมัติแล้ว รอแอดมินจัดรถ`,
      emailData);
  } else if(newStatus === 'จัดรถแล้ว') {
    sendEmailNotification(obj.requesterEmail || '',
      `[จัดรถแล้ว] ${obj.id} — ${obj.plate}`,
      `มีรถรอรับคุณแล้ว ทะเบียน ${obj.plate} คนขับ ${obj.driverName}`,
      emailData);
  } else if(newStatus === 'ไม่อนุมัติ') {
    sendEmailNotification(obj.requesterEmail || '',
      `[ไม่อนุมัติ] คำขอ ${obj.id}`,
      `คำขอของคุณไม่ได้รับการอนุมัติ เหตุผล: ${obj.rejectReason || '—'}`,
      { ...emailData, rejectReason: obj.rejectReason });
  }
  
  // Mark as sent
  sheet.getRange(row, emailSentCol).setValue(newStatus);
  sheet.getRange(row, emailSentCol).setBackground(newStatus==='ไม่อนุมัติ'?'#fee2e2':'#dcfce7');
}
