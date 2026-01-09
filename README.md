# ERP → ใบเช็คสต็อก (React + Vite)

แอปนี้ทำหน้าที่:
- อัปโหลดไฟล์ Excel ที่ export จาก ERP ได้หลายไฟล์พร้อมกัน
- ใช้ 2 คอลัมน์หลัก: **SKU** และ **สต็อกพร้อมขายในคลัง**
- แยก SKU เป็น **รุ่น / สี / ไซส์** (รูปแบบ: รุ่น-สีไซส์ เช่น `T009-ขาวS`)
- สร้างไฟล์ Excel output เป็นฟอร์ม “ใบเช็คสต็อก” (S, M, L, XL, XXL, 3XL, 4XL) และเติมช่องที่ไม่มีเป็น 0

## 1) ติดตั้งและรันบน Mac
ต้องมี Node.js (แนะนำ v18+)

```bash
npm install
npm run dev
```

แล้วเปิดลิงก์ที่ขึ้นใน Terminal (เช่น http://localhost:5173)

## 2) Build เพื่ออัปขึ้นโฮส (เช่น Vercel / Netlify)
```bash
npm run build
npm run preview
```

ไฟล์สำหรับ deploy จะอยู่ในโฟลเดอร์ `dist/`

## 3) วิธี Deploy ขึ้น Vercel (แบบง่าย)
1) สร้าง repo ใน GitHub แล้ว push โปรเจกต์นี้ขึ้นไป
2) เข้า Vercel → New Project → import repo
3) Framework: Vite
4) Build Command: `npm run build`
5) Output Directory: `dist`

## 4) หมายเหตุเรื่องคอลัมน์ใน ERP
แอปจะพยายามหา column จากชื่อหัวตารางก่อน ถ้าไม่เจอจะ fallback:
- SKU = คอลัมน์ B
- Stock ready = คอลัมน์ I

## 5) Error Report
ถ้า SKU บางรายการ parse ไม่ได้ (เช่น ไม่มี `-` หรือไม่มีไซส์ท้ายสุด)
แอปจะให้ดาวน์โหลดไฟล์ `errors_YYYYMMDD.csv`
