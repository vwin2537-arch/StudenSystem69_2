# 📋 ระบบรายงานการปฏิบัติงาน อุทยานแห่งชาติเอราวัณ

> **📌 เอกสารนี้เป็น AI Context Document** — ใช้แนบให้ AI อ่านเพื่อทำความเข้าใจโปรเจคก่อนเริ่มแชทใหม่  
> อ่านไฟล์นี้แล้วจะเข้าใจสถาปัตยกรรม, ฟังก์ชันทั้งหมด, กฎธุรกิจ และสถานะปัจจุบันของโปรเจคได้ทันที

---

## 🎯 ภาพรวมโปรเจค

ระบบ Web App สำหรับเจ้าหน้าที่อุทยานแห่งชาติเอราวัณ ใช้วางแผนวันปฏิบัติงาน, เช็คชื่อเข้า-ออกงาน, รายงานผลการปฏิบัติงาน พร้อมระบบ Admin สำหรับบริหารจัดการ

**ผู้ใช้งาน 2 กลุ่ม:**
- **Admin** — บริหารจัดการเจ้าหน้าที่, อนุมัติแผน/คำร้อง, ดู Dashboard วิเคราะห์ข้อมูล
- **User (เจ้าหน้าที่)** — ส่งแผนวันทำงาน, เช็คอิน-เช็คเอาท์, รายงานผลงาน, ขอสลับวัน

---

## 🛠 Tech Stack

| ส่วน | เทคโนโลยี |
|---|---|
| **Backend** | Google Apps Script (GAS) |
| **Database** | Google Sheets (5 ชีต) |
| **File Storage** | Google Drive |
| **Frontend** | HTML5 SPA (Single Page App) ใน GAS HtmlService |
| **CSS Framework** | Tailwind CSS (CDN) |
| **Charts** | Chart.js (CDN) |
| **Tables** | jQuery DataTables (CDN) |
| **Dialogs** | SweetAlert2 (CDN) |
| **Font** | Kanit (Google Fonts) |
| **Theme Colors** | `forest` (green), `earth` (brown) — defined in tailwind.config |

---

## 📁 โครงสร้างไฟล์

```
├── Code.gs          — Backend: ฟังก์ชันทั้งหมดของ Google Apps Script (Web App)
├── LineBotCode.gs   — LINE Bot: รายงานเช้า/เย็นอัตโนมัติ + webhook (แยก GAS Project)
├── index.html       — Frontend: HTML + CSS + JavaScript (SPA ทั้งหมดในไฟล์เดียว)
├── Draft.md         — เอกสาร draft / brainstorm
└── README.md        — ไฟล์นี้ (AI Context Document)
```

> ⚠️ เวลา deploy ไปที่ Google Apps Script ต้องใช้ชื่อไฟล์ HTML ว่า `index` (ไม่มีนามสกุล)  
> GAS จะ serve ผ่าน `doGet()` → `HtmlService.createHtmlOutputFromFile('index')`

---

## ⚙️ Configuration (Code.gs บรรทัด 1-15)

```js
var CONFIG = {
  SPREADSHEET_ID: '1yRrEV40jMMQCzycruoj8lIM0NPXTRhxXxUfvsaTO2uM',
  DRIVE_FOLDER_ID: '1kch--KBTd15dXHVIKutMAFy7Hj5ARtBo',
  DEFAULT_ADMIN: { username: 'admin', password: 'admin1234', name: 'ผู้ดูแลระบบ', role: 'Admin' },
  DEV_MODE: true,  // true = ข้ามการตรวจระยะ GPS (ยังบันทึกพิกัดปกติ) — เปลี่ยนเป็น false ตอน production
  CHECKIN_LOCATION: { lat: 14.37462, lng: 99.14541, radiusMeters: 1000, name: 'อุทยานแห่งชาติเอราวัณ' },
  SHEETS: { ... }
}
```

- **DEV_MODE: true** → ข้ามการเช็คระยะทาง GPS (เหมาะสำหรับทดสอบ)
- **CHECKIN_LOCATION** → พิกัดอุทยานเอราวัณ, รัศมี 1000 เมตร
- Google Sheet และ Drive Folder ถูก hardcode ไว้ใน CONFIG

---

## 🗄 โครงสร้างฐานข้อมูล (Google Sheets)

### ชีต `Users`
| คอลัมน์ | ประเภท | ค่าที่เป็นไปได้ |
|---|---|---|
| ID | string | `U{timestamp}` |
| Username | string | ชื่อผู้ใช้ |
| Password | string | รหัสผ่าน (plain text) |
| Name | string | ชื่อจริง |
| Role | string | `Admin`, `User` |
| Status | string | `Active`, `Pending`, `Unregistered` |

**User Status Flow:** `Unregistered` → (ลงทะเบียนรหัสผ่าน) → `Pending` → (Admin อนุมัติ) → `Active`

---

### ชีต `WorkCycles`
| คอลัมน์ | ประเภท | หมายเหตุ |
|---|---|---|
| CycleID | string | `CYC{timestamp}` |
| UserID | string | FK → Users.ID |
| Name | string | ชื่อเจ้าหน้าที่ |
| Start_Date | string | `yyyy-MM-dd` |
| End_Date | string | `yyyy-MM-dd` (Start + 89 วัน หากไม่กำหนด) |
| Required_Work_Days | number | 30 (fixed) |
| Status | string | `Active`, `Completed`, `Cancelled` |

---

### ชีต `WorkPlans`
| คอลัมน์ | ประเภท | หมายเหตุ |
|---|---|---|
| PlanID | string | `PLN{timestamp}{idx}` |
| Submission_ID | string | `SUB{timestamp}` — กลุ่มแผนที่ส่งพร้อมกัน |
| CycleID | string | FK → WorkCycles.CycleID |
| UserID | string | FK → Users.ID |
| Name | string | ชื่อเจ้าหน้าที่ |
| Plan_Date | string | `yyyy-MM-dd` |
| Plan_Status | string | `Pending`, `Approved`, `Rejected`, `Swapped_Out` |
| Notes | string | หมายเหตุ |
| Completed_LogID | string | FK → AttendanceLog.LogID (เมื่อเช็คอินแล้ว) |
| Created_At | string | datetime |
| Submitted_At | string | datetime |
| Approved_At | string | datetime |
| Day_Type | string | `Full`, `Half` |

**Plan_Status Values:**
- `Pending` — รออนุมัติจาก Admin
- `Approved` — อนุมัติแล้ว (เช็คอินได้)
- `Rejected` — ไม่อนุมัติ
- `Swapped_Out` — ถูกยกเลิกเนื่องจากสลับวัน

---

### ชีต `AttendanceLog`
| คอลัมน์ | ประเภท | หมายเหตุ |
|---|---|---|
| LogID | string | `LOG{timestamp}` |
| Date | string | `yyyy-MM-dd` |
| Name | string | ชื่อเจ้าหน้าที่ |
| Time_In | string | `yyyy-MM-dd HH:mm:ss` |
| Time_Out | string | `yyyy-MM-dd HH:mm:ss` (ว่างจนกว่าจะรายงานผล) |
| Task_Report | string | รายงานงานที่ทำ |
| Photo_URL | string | URL โฟลเดอร์รูปภาพใน Google Drive |
| Status | string | `On_Time`, `Late`, `Completed`, `Late_Report` |
| Latitude | number | พิกัดละติจูดตอนเช็คอิน |
| Longitude | number | พิกัดลองจิจูดตอนเช็คอิน |
| Distance_m | number | ระยะห่างจากจุดลงเวลา (เมตร) |
| Selfie_URL | string | URL รูปเซลฟี่ใน Google Drive |

**Status Values:**
- `On_Time` — เข้างานตรงเวลา (≤08:15) ยังไม่รายงาน
- `Late` — เข้างานสาย (>08:15) ยังไม่รายงาน
- `Completed` — รายงานผลตรงเวลา (≤17:00)
- `Late_Report` — รายงานผลล่าช้า (>17:00)

---

### ชีต `ScheduleRequests`
| คอลัมน์ | ประเภท | หมายเหตุ |
|---|---|---|
| ReqID | string | `REQ{timestamp}` |
| CycleID | string | FK → WorkCycles.CycleID |
| UserID | string | FK → Users.ID |
| Name | string | ชื่อเจ้าหน้าที่ |
| Original_Date | string | วันเดิม |
| Requested_Date | string | วันที่ต้องการเปลี่ยน |
| Reason | string | เหตุผล |
| Status | string | `Pending`, `Approved`, `Rejected` |
| Created_At | string | datetime |
| Decision_At | string | datetime |
| Request_Type | string | `Swap` (สลับวัน), `Half_Day` (ครึ่งวัน) |

---

## 🔧 Backend API (Code.gs) — ฟังก์ชันทั้งหมด

### Auth & Session
| ฟังก์ชัน | ผู้เรียก | หมายเหตุ |
|---|---|---|
| `login(username, password)` | Frontend | คืน token + user object |
| `loginAndGetData(username, password)` | Frontend | login แล้วดึง appData ทีเดียว |
| `restoreSessionAndGetData(token)` | Frontend | ใช้ตอน reload หน้า |
| `logout(token)` | Frontend | ลบ session |
| `getSessionUser(token)` | Internal | อ่าน session จาก PropertiesService |
| `validateSession_(token)` | Internal | throw error ถ้า session หมดอายุ |

**Session Storage:** ใช้ `PropertiesService.getUserProperties()` เก็บ key `session_{token}` → JSON user object

---

### Data Fetching
| ฟังก์ชัน | Role | คืนข้อมูล |
|---|---|---|
| `getUserAppData(token)` | User | activeCycle, plans, approvedDates, logs, todayLog, canCheckIn, canCheckOut |
| `getAdminAppData(token)` | Admin | ข้อมูลครบทั้งหมด + analytics (ดูหัวข้อ Admin Analytics) |

### Admin Analytics (ใน `getAdminAppData`)
ฟังก์ชันนี้คำนวณ analytics เพิ่มเติม:
- `todayExpectedCount` — จำนวนคนที่มีแผน Approved วันนี้
- `todayAvgTime` — เวลาเช็คอินเฉลี่ย/เร็วสุด/ช้าสุด วันนี้
- `dailyStats` — array 14 วัน: `{date, total, onTime, late}`
- `weekdayStats` — array 7 วัน (อา-ส): `{day, onTime, late, total}`
- `perUserStats` — array ต่อ User: `{name, username, totalDays, onTime, late, rate, avgCheckIn}` เรียงตาม rate
- `cycleProgress` — ต่อรอบงาน Active: `{name, completed, required, progressPct, remainDays, behindSchedule}`
- `weekComparison` — `{thisWeek, lastWeek, totalDiff, onTimeDiff, lateDiff}`
- `recentActivity` — 10 กิจกรรมล่าสุด: `{icon, name, date, time, detail}`

---

### Work Cycles
| ฟังก์ชัน | Role | หมายเหตุ |
|---|---|---|
| `createWorkCycle(token, userId, userName, startDate, endDateCustom)` | Admin | สร้างรอบงาน (default 90 วัน) |

---

### Work Plans
| ฟังก์ชัน | Role | หมายเหตุ |
|---|---|---|
| `createWorkPlan(token, cycleId, dates)` | User | ส่งแผน (dates = array วันที่) ตรวจ Pending ก่อน |
| `updateWorkPlanApprovalStatus(token, submissionId, newStatus)` | Admin | อนุมัติ/ปฏิเสธแผนทั้ง Submission |

---

### Schedule Requests
| ฟังก์ชัน | Role | หมายเหตุ |
|---|---|---|
| `createScheduleChangeRequest(token, cycleId, originalDate, requestedDate, reason, requestType)` | User | ยื่นคำร้อง (Swap หรือ Half_Day) |
| `updateScheduleRequestStatus(token, reqId, newStatus)` | Admin | อนุมัติ → ปรับแผนอัตโนมัติ |

**เมื่ออนุมัติ `Swap`:** วันเดิม → `Swapped_Out`, เพิ่มวันใหม่เป็น `Approved Full`  
**เมื่ออนุมัติ `Half_Day`:** วันเดิม → `Day_Type = Half`, เพิ่มวันชดเชยเป็น `Approved Half`

---

### Check In / Check Out
| ฟังก์ชัน | Role | หมายเหตุ |
|---|---|---|
| `checkIn(token, latitude, longitude, selfieData)` | User | บันทึกเวลาเข้า + อัปโหลดเซลฟี่ |
| `checkOut(token, taskReport, photoDataArray)` | User | บันทึกเวลาออก + อัปโหลดรูป + รายงานผล |

**เงื่อนไข Check In:**
1. เวลา ≥ 08:05 น. (485 นาที)
2. ต้องมีรูปเซลฟี่
3. ต้องมีพิกัด GPS
4. ต้องอยู่ในรัศมี 1000 เมตร (ข้ามถ้า `DEV_MODE: true`)
5. วันนี้ต้องมีแผน `Approved`
6. ยังไม่เคยเช็คอินวันนี้
7. **ตรงเวลา:** ≤ 08:15 (495 นาที) → status `On_Time`; >08:15 → status `Late`

**เงื่อนไข Check Out:**
1. เวลา ≥ 16:00 น. (960 นาที)
2. ต้องมีบันทึกเช็คอินวันนี้ที่ยังไม่มี Time_Out
3. ต้องแนบรูปภาพอย่างน้อย 3 รูป (ตรวจที่ Frontend)
4. รายงานตรงเวลา: ≤ 17:00 (1020 นาที) → `Completed`; >17:00 → `Late_Report`

---

### Photo Management
| ฟังก์ชัน | หมายเหตุ |
|---|---|
| `uploadSelfie_(userName, dateStr, selfieData)` | อัปโหลดรูปเซลฟี่ → Drive, คืน thumbnail URL |
| `uploadPhotos_(userName, dateStr, photoDataArray)` | อัปโหลดรูปงาน → Drive, คืน folder URL |

**โครงสร้างโฟลเดอร์ Drive:**  
`ROOT_FOLDER / {yyyy-MM} / {yyyy-MM-dd} / {ชื่อ} / selfie_*.jpg + รูปงาน`

---

### User Management
| ฟังก์ชัน | Role | หมายเหตุ |
|---|---|---|
| `addUser(token, username, name)` | Admin | สร้าง User ใหม่ (Status: Unregistered) |
| `getUnregisteredUsers()` | Public | รายชื่อ Unregistered สำหรับหน้าลงทะเบียน |
| `registerUser(userId, password)` | Public | ลงทะเบียนรหัสผ่าน → Status: Pending |
| `approveUser(token, userId)` | Admin | อนุมัติ → Status: Active |
| `rejectUser(token, userId)` | Admin | ปฏิเสธ → Status: Unregistered (ล้าง Password) |

---

### Helper Functions
| ฟังก์ชัน | หมายเหตุ |
|---|---|
| `getSheetData_(ss, sheetName)` | อ่านทุก row คืนเป็น array of objects (header เป็น key) |
| `formatDate_(date)` | → `yyyy-MM-dd` |
| `formatTime_(date)` | → `HH:mm:ss` |
| `formatDateTime_(date)` | → `yyyy-MM-dd HH:mm:ss` |
| `calculateDistance_(lat1, lng1, lat2, lng2)` | Haversine formula → เมตร |
| `ensureSetup_()` | สร้างชีตและ Admin เริ่มต้นถ้ายังไม่มี |
| `updatePlanCompletedLog_(ss, userId, date, logId)` | อัปเดท Completed_LogID ใน WorkPlans หลังเช็คอิน |

---

## 🖥 Frontend (index.html) — โครงสร้าง

### Global State
```js
var APP = {
  token: null,      // session token
  user: null,       // { id, name, role }
  data: null        // ข้อมูลทั้งหมดจาก appData
};
```

### หน้าหลัก (Login Flow)
1. โหลดหน้า → `initializeApp()` → แสดง login form
2. ถ้ามี localStorage token → `restoreSessionAndGetData(token)` → ข้ามหน้า login
3. Login สำเร็จ → `APP.token`, `APP.user`, `APP.data` → `renderApp()`

### Function หลัก Frontend
| ฟังก์ชัน | หมายเหตุ |
|---|---|
| `renderApp()` | ตรวจ role → render tabs ของ Admin หรือ User |
| `switchTab(tabId)` | เปลี่ยนแท็บ → เรียก render function ที่เหมาะสม |
| `loadAdminData()` | reload `getAdminAppData` แล้ว re-render |
| `formatThaiDate(dateStr)` | แปลง `yyyy-MM-dd` → Thai date string |
| `formatDateLocal(d)` | แปลง Date object → `yyyy-MM-dd` ตาม local timezone |
| `startRealtimeClock(mode)` | นาฬิกา real-time + countdown ถึงเวลาเปิด/ปิด |
| `doCheckIn()` | เปิด Swal ถ่ายเซลฟี่ → `checkIn()` |
| `doCheckOut()` | เปิด Swal กรอกรายงาน + แนบรูป → `checkOut()` |
| `compressImage(file, maxDim, quality, callback)` | บีบอัดรูปก่อนอัปโหลด (max 1200px, 70%) |

---

## 📱 แท็บ Admin

| Tab ID | ฟังก์ชัน render | เนื้อหา |
|---|---|---|
| `dashboard` | `renderDashboard(area)` | KPI cards, 4 charts, ranking, cycle progress, activity feed |
| `manage_cycles` | `renderCycles(area)` | สร้างรอบงาน, รายการรอบงานทั้งหมด |
| `approve_plans` | `renderApprovePlans(area)` | รายการแผนรออนุมัติ |
| `approve_requests` | `renderApproveRequests(area)` | รายการคำร้องรออนุมัติ |
| `reports` | `renderReports(area)` | รายงานการเข้างาน (DataTable + รูป) |
| `manage_users` | `renderManageUsers(area)` | จัดการผู้ใช้ + อนุมัติ/ปฏิเสธ |

---

## 📱 แท็บ User (เจ้าหน้าที่)

| Tab ID | ฟังก์ชัน render | เนื้อหา |
|---|---|---|
| `user_dashboard` | `renderUserDashboard(area)` | สถานะวันนี้, ความคืบหน้ารอบ, ประวัติล่าสุด |
| `work_plan` | `renderWorkPlan(area)` | ปฏิทินเลือกวันทำงาน (30 วัน) + ส่งแผน |
| `checkin` | `renderCheckIn(area)` | นาฬิกา real-time + ปุ่มเช็คอิน + ประวัติ |
| `checkout` | `renderCheckOut(area)` | กรอกรายงาน + แนบรูป + ปุ่มเช็คเอาท์ |
| `requests` | `renderRequests(area)` | ยื่นคำร้องสลับวัน/ครึ่งวัน + ประวัติคำร้อง |

---

## 📊 Admin Dashboard — รายละเอียด (`renderDashboard`)

Dashboard ใหม่ (เขียนใหม่ทั้งหมด ณ April 2026) ประกอบด้วย:

1. **Header** — ทักทาย Admin + วันที่ + ปุ่มรีเฟรช
2. **Alert Bar** — แจ้งงานรอดำเนินการ (คลิกได้ → switchTab)
3. **KPI Row 1 (วันนี้)** — มาปฏิบัติงาน (x/y + progress bar), ตรงเวลา (+trend ▲▼), สาย (+trend), เวลาเฉลี่ย
4. **KPI Row 2 (ภาพรวม)** — เจ้าหน้าที่ Active, รอบงาน, บันทึกทั้งหมด, อัตราตรงเวลา (+ progress bar)
5. **กราฟ 4 ชิ้น:**
   - Line chart: แนวโน้ม 14 วันล่าสุด (ตรงเวลา/สาย/รวม)
   - Stacked bar: วิเคราะห์ตามวันในสัปดาห์ (จ-อา)
   - Stacked bar: สถิติรายเดือน
   - Doughnut: สัดส่วนตรงเวลา/สาย + สรุปเปรียบเทียบสัปดาห์
6. **Ranking Table** — สถิติรายบุคคล, badge (🟢🟡🔴), progress bar อัตรา%, เวลาเฉลี่ย
7. **Cycle Progress** — progress bar ต่อรอบ + ป้ายช้ากว่า/ตามกำหนด
8. **Bottom Grid (2/3 + 1/3)** — ตารางผู้มาวันนี้ + Activity Feed 10 รายการ

---

## 🗓 ปฏิทิน (Work Plan Calendar)

- แสดงเดือนปัจจุบัน เลื่อนไปข้างหน้าได้ในรอบงาน
- วันสีต่างๆ: `selected` (เขียว), `approved` (เขียวเข้ม), `pending` (เหลือง), `half-day` (ฟ้า), `holiday` (แดงประ), `disabled` (จาง)
- วันหยุดนักขัตฤกษ์ไทย hardcode ใน JS (`thaiHolidays` object) — ยังเลือกได้ถ้าต้องการ
- ต้องเลือก **ครบ 30 วัน** จึงจะส่งแผนได้
- ไม่สามารถเลือกวันที่เช็คอินแล้ว หรือวันที่มีแผน Approved แล้ว

---

## 📏 กฎธุรกิจสำคัญ

| กฎ | รายละเอียด |
|---|---|
| รอบงาน | 1 รอบ = 90 วัน, Required = 30 วัน |
| ส่งแผน | ต้องเลือกครบ 30 วัน, ห้ามส่งซ้ำถ้ายังมี Pending อยู่ |
| เช็คอิน | เฉพาะวันที่ Plan_Status = Approved, ตั้งแต่ 08:05 น. |
| ตรงเวลา | ≤ 08:15 น. = On_Time; > 08:15 น. = Late |
| เช็คเอาท์ | ตั้งแต่ 16:00 น., ≤ 17:00 น. = Completed, > 17:00 น. = Late_Report |
| รูปถ่าย | ต้องแนบ ≥ 3 รูป ตอนรายงานผล |
| เซลฟี่ | บังคับถ่ายตอนเช็คอิน |
| GPS | ต้องอยู่ในรัศมี 1000 ม. (ข้ามถ้า DEV_MODE=true) |
| รูปบีบอัด | max 1200px, quality 70% ก่อน upload |

---

## 👤 User Status Flow

```
[Admin สร้าง] → Unregistered
     ↓ (User ลงทะเบียนรหัสผ่าน)
  Pending
     ↓ (Admin อนุมัติ)
  Active  ←→  (Admin ปฏิเสธ) → Unregistered
```

---

## 🔑 บัญชีเริ่มต้น

| Username | Password | Role |
|---|---|---|
| `admin` | `admin1234` | Admin |

> ⚠️ เปลี่ยนรหัสผ่าน Admin หลัง deploy จริง

---

## 🚀 วิธี Deploy

1. ไปที่ [script.google.com](https://script.google.com) → สร้างโปรเจกต์ใหม่
2. วางโค้ดจาก `Code.gs` ลงในไฟล์ `Code.gs`
3. สร้างไฟล์ HTML ใหม่ชื่อ `index` (ไม่มีนามสกุล) วางโค้ดจาก `index.html`
4. **Deploy > New deployment > Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
5. Copy URL → เปิดในเบราว์เซอร์
6. Allow permissions (Google Sheets + Drive)

---

## 🔄 Git Repository

- **GitHub:** `https://github.com/vwin2537-arch/StudenSystem69_2`
- **Branch:** `main`
- **Remote:** `origin`

```bash
# commit และ push
git add Code.gs index.html
git commit -m "message"
git push
```

> หมายเหตุ: GAS ไม่ sync กับ Git อัตโนมัติ — ต้อง copy-paste โค้ดไปที่ script.google.com ด้วยตนเอง

---

## 📝 ประวัติการแก้ไขล่าสุด (Session Log)

### April 2026 — Professional Admin Dashboard
**แก้ไขไฟล์:** `Code.gs`, `index.html`

**Code.gs — เพิ่มใน `getAdminAppData`:**
- `todayExpectedCount` — คนที่มีแผน Approved วันนี้
- `todayAvgTime` — เวลาเช็คอินเฉลี่ย/เร็วสุด/ช้าสุดวันนี้
- `dailyStats` — 14 วันล่าสุด (array)
- `weekdayStats` — วิเคราะห์ตามวัน จ-อา
- `perUserStats` — Engagement Score รายบุคคล (วิเคราะห์ Time_In / Time_Out อิสระ, นับวันขาด, เต็ม 100/วัน)
- `cycleProgress` — ความคืบหน้ารอบงาน + ตรวจว่าช้ากว่ากำหนดหรือไม่
- `weekComparison` — เทียบสัปดาห์นี้ vs สัปดาห์ก่อน
- `recentActivity` — 10 กิจกรรมล่าสุด (checkin/checkout/plan/request)

**index.html — `renderDashboard` เขียนใหม่ทั้งหมด:**
- KPI cards พร้อม progress bar + trend indicator (▲▼)
- เพิ่มกราฟ Line (แนวโน้ม 14 วัน) และ Bar (วิเคราะห์รายวันสัปดาห์)
- Ranking Table: TOP 3 cards + ตาราง 13 คอลัมน์ (score, แผน, มา, ขาด, เช้าตรง/สาย, รายงานตรง/สาย/ไม่ส่ง, เฉลี่ยเข้า, ระดับ)
- Cycle Progress cards
- Activity Feed timeline
- Alert bar คลิกได้ (switchTab)

---

### April 28, 2026 — Fix: ส่งรายงานไม่ได้เพราะ Google Drive Error
**แก้ไขไฟล์:** `Code.gs`

**อาการ:** User 1 คนส่งรายงานเย็นไม่ได้ ขึ้น "Exception: ข้อผิดพลาดของบริการ: ไดรฟ์"

**สาเหตุ:** `checkOut` ไม่มี try-catch ครอบ `uploadPhotos_()` — เมื่อ Drive ขัดข้อง (ไฟล์ใหญ่/transient error) ก็ crash ทั้ง function → ส่งรายงานไม่ได้เลย (เทียบกับ `checkIn` ที่มี try-catch แล้ว)

**แก้ไข:**
- `checkOut`: เพิ่ม try-catch ครอบ `uploadPhotos_` — ถ้าอัปโหลดรูปไม่ได้ยังบันทึกรายงานได้ (บันทึก error ไว้ใน Photo_URL)
- `uploadPhotos_`: เพิ่ม try-catch รอบแต่ละรูป — รูปใดเสียก็ข้ามไป ไม่ block รูปอื่น

---

### April 27, 2026 — Rewrite Engagement Score Ranking System
**แก้ไขไฟล์:** `Code.gs`, `index.html`

**ปัญหาเดิม:** ระบบจัดอันดับใช้ `Status` จาก AttendanceLog ซึ่งถูก checkout เขียนทับ → คนเช้าสายแต่ส่งรายงานตรงเวลาถูกนับว่า "ตรงเวลา" + ไม่นับรายงานเย็น + ไม่นับคนขาด

**Engagement Score ใหม่ (ต่อวัน เต็ม 100):**
| เกณฑ์ | คะแนน | เงื่อนไข |
|---|---|---|
| มาทำงาน | 30 | มี check-in record |
| เช้าตรงเวลา | 30 | `Time_In` ≤ 08:15 |
| ส่งรายงาน | 20 | มี `Time_Out` |
| รายงานตรงเวลา | 20 | `Time_Out` ≤ 17:00 |
| ขาดงาน | 0 | มีแผนแต่ไม่มา |

**สูตร:** `engagementScore = totalPoints / (plannedDays × 100) × 100`

**Backend:** วิเคราะห์ `Time_In` / `Time_Out` อิสระจากกัน (ไม่พึ่ง Status), นับ plannedDays จาก WorkPlans (Approved, ≤ today)
**Frontend:** TOP 3 highlight cards + ตาราง 13 คอลัมน์ + คำอธิบายเกณฑ์ด้านล่าง

---

### April 26, 2026 — Bug Fix: Schedule Swap ไม่มีผลใน WorkPlans
**แก้ไขไฟล์:** `Code.gs`

**สาเหตุ:** ฟังก์ชัน `updateScheduleRequestStatus` ใช้ `getDataRange().getValues()` แบบ raw ซึ่ง Google Sheets คืน Date เป็น JS Date object แต่เปรียบเทียบด้วย `===` (reference comparison) → ไม่เจอแถวที่ตรงกัน → วันเดิมไม่ถูกเปลี่ยนเป็น `Swapped_Out`

**แก้ไข:**
- เพิ่ม normalize Date→string ก่อนเปรียบเทียบ (`instanceof Date ? formatDate_() : String()`)
- เพิ่ม `String()` wrapper รอบ UserID/CycleID เพื่อป้องกัน type mismatch
- เขียน date ลง appendRow เป็น string `yyyy-MM-dd` แทน Date object

---

## 🤖 LINE Bot (`LineBotCode.gs`)

> ⚠️ ไฟล์นี้ **deploy แยก GAS Project** จาก Web App หลัก แต่ใช้ Google Sheet ID เดียวกัน

**Functions:**
| ฟังก์ชัน | Trigger | หมายเหตุ |
|---|---|---|
| `sendMorningReport()` | 08:30 ทุกวัน | สรุปเช้า: ใครมา/ไม่มา/สาย |
| `sendEveningReport()` | 17:30 ทุกวัน | สรุปเย็น: ใครส่ง/ยังไม่ส่งรายงาน |
| `doPost(e)` | LINE Webhook | ตอบ "ทดสอบ" → ส่งรายงานทั้ง 2 ช่วง |
| `setupTriggers()` | รันครั้งเดียว | ตั้ง time-based trigger |

**Morning Report ดึงข้อมูลจาก:** `WorkPlans` (Plan_Date = today, Plan_Status = Approved) + `AttendanceLog` (Date = today)  
**Evening Report ดึงข้อมูลจาก:** `AttendanceLog` (Time_Out มี/ไม่มี, Status = Completed/Late_Report)

**Config ที่ต้องตั้ง:**
- `BOT_CONFIG.LINE_CHANNEL_ACCESS_TOKEN` — จาก LINE Developers Console
- `BOT_CONFIG.LINE_GROUP_ID` — ได้จากเชิญ bot เข้ากลุ่ม (bot จะ reply Group ID)

---

## 🔍 จุดที่ต้องระวังเมื่อแก้ไขโค้ด

1. **`getSheetData_`** — คืน array of objects; Date objects ถูกแปลงเป็น string อัตโนมัติ
2. **`session_` key** — เก็บใน `PropertiesService.getUserProperties()` ไม่ใช่ Script Properties
3. **Chart.js** — ใช้ CDN, ต้องรอ `setTimeout(..., 150)` ให้ DOM พร้อมก่อน render
4. **DEV_MODE** — ตอนนี้เป็น `true` (ข้ามเช็ค GPS) อย่าลืมเปลี่ยนก่อน production
5. **`formatDate_`** — ใช้ `Session.getScriptTimeZone()` (Asia/Bangkok) — อย่าใช้ `new Date().toISOString()`
6. **Frontend date** — ใช้ `formatDateLocal()` (local TZ) ไม่ใช่ `toISOString()` เพื่อป้องกัน off-by-one
7. **Photo compression** — ทำที่ Frontend ก่อน base64 → ส่งไป GAS
8. **Submission_ID** — Admin อนุมัติ/ปฏิเสธ ทีละ Submission (หลายแถวพร้อมกัน) ไม่ใช่ทีละ Plan
9. **`getDataRange().getValues()` vs `getSheetData_()`** — raw values คืน Date object; ห้ามเปรียบเทียบ Date กับ Date ด้วย `===` ต้อง normalize เป็น string ก่อนเสมอ
10. **`Status` ใน AttendanceLog ถูก checkout เขียนทับ** — checkout เปลี่ยน `On_Time`→`Completed` หรือ `Late`→`Late_Report` ทำให้หาย ข้อมูลเช้าสาย/ตรงเวลาจริงต้องดูจาก `Time_In` เท่านั้น
