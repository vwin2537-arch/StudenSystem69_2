# คู่มือตั้งค่า LINE Bot แจ้งเตือนระบบเช็คชื่อ

## สิ่งที่ต้องเตรียม
- บัญชี LINE (ใช้กับ LINE Developers)
- Google Account (ใช้สร้าง Apps Script ตัวใหม่)

---

## ขั้นตอนที่ 1: สร้าง LINE Bot ใน LINE Developers

1. เข้า https://developers.line.biz/ → Login ด้วย LINE
2. กด **Create a new provider** → ตั้งชื่อ เช่น `อช.เอราวัณ`
3. กด **Create a Messaging API channel**
   - Channel name: `บอทแจ้งเตือน อช.เอราวัณ` (หรือชื่อที่ต้องการ)
   - Channel description: `บอทสรุปการเช็คชื่อประจำวัน`
   - Category: `Organization`
   - Subcategory: `Government`
4. เข้าหน้า Channel ที่สร้าง → ไปที่แท็บ **Messaging API**
5. เลื่อนลงหา **Channel access token (long-lived)** → กด **Issue**
6. **Copy token นี้ไว้** → จะใช้ใส่ในโค้ด

### ตั้งค่าเพิ่มเติมใน Messaging API:
- **Auto-reply messages** → ปิด (Disabled)
- **Greeting messages** → ปิด (Disabled)
- **Allow bot to join group chats** → เปิด (Enabled)

---

## ขั้นตอนที่ 2: สร้าง Google Apps Script ตัวใหม่

> ⚠️ ใช้ **Script ตัวใหม่แยกจากตัวเดิม** — Webhook URL ต้องเป็นของ Script ตัวใหม่นี้

1. เข้า https://script.google.com/ → กด **New project**
2. ตั้งชื่อโปรเจค: `LINE Bot แจ้งเตือน อช.เอราวัณ`
3. ลบโค้ดเดิมทั้งหมดในไฟล์ `Code.gs`
4. Copy โค้ดจากไฟล์ `LineBotCode.gs` ในโปรเจคนี้ วางลงไปแทน
5. แก้ไข 2 ค่าใน `BOT_CONFIG`:
   ```javascript
   LINE_CHANNEL_ACCESS_TOKEN: 'วาง_token_ที่_copy_จากขั้นตอนที่_1'
   LINE_GROUP_ID: 'จะได้ในขั้นตอนที่ 4'
   ```

---

## ขั้นตอนที่ 3: Deploy Script ตัวใหม่ + ตั้ง Webhook

### Deploy เป็น Web App:
1. กด **Deploy** → **New deployment**
2. กดไอคอนเฟือง ⚙️ → เลือก **Web app**
3. ตั้งค่า:
   - Description: `LINE Bot Webhook`
   - Execute as: **Me**
   - Who has access: **Anyone**
4. กด **Deploy**
5. **Copy URL ที่ได้** (จะเป็น `https://script.google.com/macros/s/xxx/exec`)

### ตั้ง Webhook ใน LINE Developers:
1. กลับไปหน้า LINE Developers → แท็บ **Messaging API**
2. หาช่อง **Webhook URL** → กด **Edit**
3. วาง URL จากขั้นตอนด้านบน
4. กด **Update** → กด **Verify** (ต้องขึ้น ✅ Success)
5. เปิด **Use webhook** → ON

---

## ขั้นตอนที่ 4: หา Group ID

1. เปิดแอป LINE → สร้างกลุ่ม หรือเข้ากลุ่มที่ต้องการให้บอทส่งข้อความ
2. เข้าไปที่กลุ่ม → กดเชิญสมาชิก → ค้นหาชื่อบอทที่สร้าง → เพิ่มเข้ากลุ่ม
3. **บอทจะส่งข้อความอัตโนมัติ** แสดง Group ID ของกลุ่มนั้น
4. Copy Group ID → ไปแก้ในโค้ด:
   ```javascript
   LINE_GROUP_ID: 'Cxxxxxxxxxxxxxxxxxxxxxxxxxx'
   ```
5. กด **Deploy** → **Manage deployments** → **Edit (ไอคอนดินสอ)** → เลือก **New version** → **Deploy**

> 💡 หมายเหตุ: ทุกครั้งที่แก้โค้ด ต้อง Deploy version ใหม่เสมอ!

---

## ขั้นตอนที่ 5: ตั้ง Trigger (แจ้งเตือนอัตโนมัติ)

1. ในหน้า Apps Script Editor
2. ไปที่เมนูด้านบน → เลือกฟังก์ชัน `setupTriggers` จาก dropdown
3. กดปุ่ม **▶ Run**
4. อนุญาต permissions ที่ขอ (Google Sheet, External URL)
5. ระบบจะตั้ง Trigger อัตโนมัติ:
   - **08:30** → ส่งรายงานช่วงเช้า
   - **17:30** → ส่งรายงานช่วงเย็น

> ⏰ หมายเหตุ: `nearMinute(30)` อาจคลาดเคลื่อน ±15 นาที (ข้อจำกัดของ GAS Time Trigger)

---

## ขั้นตอนที่ 6: ทดสอบ

1. เข้าไปในกลุ่ม LINE ที่เพิ่มบอทแล้ว
2. พิมพ์ข้อความ: **ทดสอบ**
3. บอทจะตอบกลับด้วยรายงานตัวอย่างทั้ง 2 ช่วง (เช้า + เย็น)
4. ข้อความอื่นๆ ที่ไม่ใช่ "ทดสอบ" → บอทจะไม่ตอบ (ตามที่ออกแบบไว้)

---

## สรุปโครงสร้าง

| รายการ | รายละเอียด |
|--------|-----------|
| Script เดิม | ระบบเช็คชื่อหลัก (Code.gs + index.html) |
| Script ใหม่ | LINE Bot แจ้งเตือน (LineBotCode.gs) |
| Webhook URL | ใช้ลิงก์ของ **Script ตัวใหม่** |
| ข้อมูล | อ่านจาก **Google Sheet เดียวกัน** |
| Trigger | เช้า 08:30 + เย็น 17:30 (ตั้งผ่าน `setupTriggers`) |
| คำสั่งทดสอบ | พิมพ์ "ทดสอบ" ในกลุ่ม LINE |

---

## FAQ

**Q: แก้โค้ดแล้วบอทไม่เปลี่ยน?**
A: ต้อง Deploy version ใหม่ทุกครั้ง (Manage deployments → Edit → New version → Deploy)

**Q: บอทไม่ตอบเลย?**
A: ตรวจสอบ 3 สิ่ง:
1. Webhook URL ถูกต้อง + Verify ผ่าน
2. Use webhook เปิดอยู่
3. Channel Access Token ถูกต้อง

**Q: Trigger ไม่ทำงาน?**
A: ไปที่ Apps Script → ไอคอนนาฬิกา (Triggers) → ตรวจว่ามี 2 triggers อยู่
