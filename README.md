# Akara Resources — HR Policy Web App
## วิธีใช้งาน (Setup Guide)

---

## โครงสร้างโฟลเดอร์

```
hr-policy-app/
├── index.html                ← หน้าหลัก (เปิดไฟล์นี้)
├── generate-script.py        ← Script สร้าง HTML อัตโนมัติ
├── assets/
│   ├── style.css             ← Akara theme (ใช้ร่วมกันทุกหน้า)
│   └── Akara_Logo.jpg        ← ⚠️ วางโลโก้ไว้ที่นี่
├── content/
│   ├── work-rules/           ← ⚠️ วาง PDF Work Rules 59 ไฟล์ที่นี่
│   └── welfare/              ← ⚠️ วาง PDF Welfare 30 ไฟล์ที่นี่
├── work-rules/
│   └── index.html            ← สารบัญ Work Rules (auto-generated)
└── welfare/
    └── index.html            ← สารบัญ Welfare (auto-generated)
```

---

## ขั้นตอน

### 1. ติดตั้ง Python Library
```bash
pip install PyMuPDF
```

### 2. วางไฟล์

**โลโก้:**
```
assets/Akara_Logo.jpg
```

**PDF Work Rules (59 ไฟล์) — ตั้งชื่อให้เรียงลำดับ:**
```
content/work-rules/01_บทนำ.pdf
content/work-rules/02_เวลาทำงาน.pdf
content/work-rules/03_การลา.pdf
...
content/work-rules/59_บทสรุป.pdf
```

**PDF Welfare (30 ไฟล์):**
```
content/welfare/01_ประกันสุขภาพ.pdf
content/welfare/02_กองทุนสำรองเลี้ยงชีพ.pdf
...
content/welfare/30_สวัสดิการอื่นๆ.pdf
```

> 💡 **สำคัญ:** ตั้งชื่อไฟล์ขึ้นต้นด้วยตัวเลข เช่น `01_`, `02_`
> เพื่อให้ script เรียงลำดับถูกต้อง

### 3. รัน Script
```bash
cd hr-policy-app
python generate-script.py
```

Script จะ:
- อ่านข้อความจาก PDF ทุกไฟล์
- สร้างหน้า HTML สำหรับแต่ละบท
- อัปเดตหน้าสารบัญอัตโนมัติ

### 4. เปิดดูใน VS Code
- ติดตั้ง Extension: **Live Server**
- คลิกขวาที่ `index.html` → **Open with Live Server**

---

## PDF ที่แยกข้อความไม่ได้ (Scanned PDF)

ถ้า PDF เป็นไฟล์สแกน (ภาพถ่าย) ต้องใช้ OCR ก่อน:
```bash
pip install pytesseract pillow
# + ติดตั้ง Tesseract: https://github.com/tesseract-ocr/tesseract
```

แก้ไข `generate-script.py` บรรทัด `extract_text_from_pdf()` ตามต้องการ

---

## ปรับแต่ง Theme

แก้ไขสีใน `assets/style.css`:
```css
:root {
  --blue: #2E5BA8;    /* สีหลัก Akara */
  --gold: #F5C518;    /* สีทอง */
}
```

---

## Deploy

เมื่อต้องการแชร์ให้พนักงาน สามารถ:
1. **Upload ไฟล์ทั้งโฟลเดอร์** ขึ้น Web Server หรือ SharePoint
2. **ใช้ GitHub Pages** (ฟรี) สำหรับ hosting
3. **วางไว้บน Network Drive** แล้วเปิดด้วย Browser ได้เลย
