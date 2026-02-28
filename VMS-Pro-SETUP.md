# 🚗 VMS Pro — คู่มือการติดตั้งและเชื่อมต่อ Google Services

## ภาพรวมระบบ

```
[index.html on GitHub Pages]
        │
        ├── localStorage (ข้อมูล offline/demo)
        ├── Google Sheets API (ฐานข้อมูล Real-time)
        ├── Google Apps Script (Email + Calendar)
        └── Google OAuth (Login)
```

---

## ขั้นตอนที่ 1: ตั้งค่า Google Sheets

1. ไปที่ https://sheets.google.com สร้าง Spreadsheet ใหม่
2. ตั้งชื่อ Sheets (Tabs) ดังนี้:
   - `Bookings`
   - `Vehicles`
   - `Drivers`
   - `Users`
   - `Trips`
   - `EmailLogs`
3. คัดลอก **Spreadsheet ID** จาก URL:
   `https://docs.google.com/spreadsheets/d/**[SPREADSHEET_ID]**/edit`
4. แชร์ไฟล์เป็น "Anyone with the link can View"

---

## ขั้นตอนที่ 2: Deploy Google Apps Script

1. เปิด Google Sheets → เมนู **Extensions > Apps Script**
2. ลบโค้ดเดิมทั้งหมด แล้ววาง `Code.gs` ที่ให้ไว้
3. แก้ไข CONFIG ที่บรรทัดบนสุด:

```javascript
const CONFIG = {
  SPREADSHEET_ID: 'วาง ID ของ Spreadsheet ที่นี่',
  CALENDAR_ID:    'your-email@gmail.com',
  SENDER_NAME:    'ระบบจัดการยานพาหนะ VMS',
  COMPANY_NAME:   'ชื่อองค์กรของคุณ',
  APP_URL:        'https://yourname.github.io/vms/',
};
```

4. กด **Deploy > New deployment**
   - Type: **Web App**
   - Execute as: **Me**
   - Who has access: **Anyone**
5. คัดลอก **Web App URL** ที่ได้

---

## ขั้นตอนที่ 3: ตั้งค่า Trigger สำหรับ Auto Email

1. ใน Apps Script → เมนู **Triggers (นาฬิกา)**
2. เพิ่ม Trigger ใหม่:
   - Function: `dailyMorningCheck`
   - Event: **Time-driven > Day timer > 8am to 9am**
3. เพิ่ม Trigger อีกอัน:
   - Function: `onEditTrigger`
   - Event: **From spreadsheet > On edit**

---

## ขั้นตอนที่ 4: ตั้งค่าใน VMS Pro Web App

1. เปิด `index.html` → เข้าสู่ระบบ
2. ไปที่ **⚙️ ตั้งค่าระบบ > Google Services**
3. กรอก **Google Spreadsheet ID** แล้วกด "บันทึก"
4. ไปที่ Tab **Email / แจ้งเตือน**
5. กรอก **Apps Script Web App URL** ที่ได้จาก Deploy
6. กด "ส่ง Test Email" เพื่อทดสอบ

---

## ขั้นตอนที่ 5: Deploy บน GitHub Pages

```bash
# 1. สร้าง Repository ใหม่บน GitHub
git init
git add .
git commit -m "Initial VMS Pro"
git remote add origin https://github.com/yourname/vms.git
git push -u origin main

# 2. เปิด Settings > Pages > Source: main branch
# 3. URL จะเป็น: https://yourname.github.io/vms/
```

---

## ข้อมูลบัญชีทดสอบ (Demo)

| อีเมล | รหัสผ่าน | บทบาท |
|-------|---------|-------|
| admin@company.com | admin123 | System Admin |
| user@co.th | 1234 | ผู้ขอใช้รถ |
| manager@co.th | 1234 | หัวหน้างาน |
| driver@co.th | 1234 | พนักงานขับรถ |

---

## ฟีเจอร์ทั้งหมด

### 🔐 Login & Security
- Login ด้วยอีเมล + รหัสผ่าน (AES-256 ready)
- Google OAuth 2.0 (ต้องตั้งค่า Client ID)
- Role-based access control (4 บทบาท)
- Session management via localStorage

### 📋 คำขอใช้รถ (Bookings)
- ✅ สร้าง / แก้ไข / ลบ คำขอ
- ✅ แนบเอกสาร PDF หลายไฟล์
- ✅ Filter ตามสถานะ / ค้นหา
- ✅ Export to Google Sheets

### ✅ Approval Workflow
- ✅ อนุมัติ / ปฏิเสธ พร้อมเหตุผล
- ✅ ส่ง Email แจ้งเตือนอัตโนมัติ
- ✅ ป้องกัน Email ซ้ำซ้อน (EmailLogs)

### 🚗 Dispatch Management
- ✅ Gantt Chart ตารางเดินรถ
- ✅ จัดคู่รถ + คนขับ
- ✅ ส่ง Email แจ้งผู้ขอ + คนขับ
- ✅ Sync Google Calendar Event

### 🚌 Vehicle Management (CRUD)
- ✅ เพิ่ม / แก้ไข / ลบ รถ
- ✅ ติดตาม พรบ. / ประกันภัย / เลขไมล์
- ✅ แจ้งเตือน 30 วันก่อนหมดอายุ
- ✅ Sync Google Sheets

### 👨‍✈️ Driver Management (CRUD)
- ✅ เพิ่ม / แก้ไข / ลบ คนขับ
- ✅ ติดตามใบขับขี่หมดอายุ
- ✅ สถานะคนขับ real-time

### 📅 Google Calendar
- ✅ ปฏิทินรายเดือน
- ✅ แสดง Events จาก Bookings
- ✅ Sync อัตโนมัติเมื่อจัดรถ
- ✅ แจ้งเตือนบำรุงรักษา

### 📈 Reports
- ✅ สรุปการใช้รถแยกแผนก
- ✅ ค่าใช้จ่ายรายเดือน
- ✅ Export PDF
- ✅ ส่งรายงานทาง Email

### 🔔 Notifications
- ✅ In-app notification panel
- ✅ Email via Gmail API
- ✅ Daily morning check trigger
- ✅ onEdit trigger (auto email เมื่อ status เปลี่ยน)

---

## โครงสร้างไฟล์

```
vms/
├── index.html          ← Web Application หลัก (ไฟล์เดียว)
├── Code.gs             ← Google Apps Script Backend
├── SETUP.md            ← คู่มือนี้
└── README.md           ← สำหรับ GitHub
```

---

## Security Notes

1. **อย่าเก็บ API Key ใน Frontend code** — ใช้ Apps Script เป็น proxy แทน
2. **Google Apps Script** จัดการ Authentication ด้วย Google Account
3. **Environment Variables** สำหรับ GitHub Actions ใช้ Repository Secrets
4. **HTTPS only** — GitHub Pages รองรับ HTTPS โดยอัตโนมัติ
5. **Input Validation** — ควรเพิ่มบน Apps Script ฝั่ง Server

---

*VMS Pro v2.1.0 — Built with ❤️ using Vanilla JS + Google Apps Script*
