# check-edi-Quality

เครื่องมือตรวจสอบเอกสารขาเข้าเรือก่อนจัดทำใบขนสินค้าขาเข้า (Thailand ocean import) —
เทียบ **MANIFEST** (รายงานจากสายเรือ, .xls/.xlsx) กับ **ENTER** (ฟอร์ม AMENDMENT ของ Quality, .pdf)
ทีละ B/L ย่อย แล้วสร้างรายงาน `EDI.xlsx` แบบไฮไลต์สีบอกจุดที่ต้องแก้ก่อนยื่นใบขน

> **ข้อจำกัด:** หน้าเว็บ (และตัวอ่าน ENTER ในโปรแกรม) รองรับเฉพาะ **ฟอร์ม AMENDMENT ของ Quality**
> ถ้าเจอฟอร์มของเจ้าอื่นหน้าเว็บจะแจ้งว่าไม่ใช่ฟอร์ม Quality ไฟล์ CNTRS ไม่นำมาเทียบ

## Check EDI Quality บนเว็บ (ให้เพื่อนใช้ได้ ไม่ต้องติดตั้ง)

https://rattanao.github.io/check-edi-Quality/

1. ลากไฟล์ **MANIFEST (.xls/.xlsx)** และ **ENTER ฟอร์ม AMENDMENT ของ Quality (.pdf ได้หลายไฟล์)** มาวางพร้อมกัน
   (ตั้งชื่ออะไรก็ได้ — แยกประเภทจากนามสกุล; ไฟล์ที่ชื่อมี `CNTRS` จะถูกข้าม)
2. ใส่ SHED NO. ที่แจ้ง (ถ้ามี) แล้วกด **ตรวจสอบ**
3. ดูผลบนหน้าเว็บ (ตารางสีเหมือน EDI.xlsx, ค้นหา/กรอง "แสดงเฉพาะที่ไม่ตรง") หรือกด **ดาวน์โหลด EDI.xlsx**

ไฟล์ถูกประมวลผลในเบราว์เซอร์ของผู้ใช้เท่านั้น ไม่มีการอัปโหลดไปที่ใด

โค้ดหน้าเว็บอยู่ที่ `docs/` (`index.html` + `core.js` — พอร์ตจาก `build_edi.py` เฉพาะฟอร์ม Quality,
ฐานเดียวกับ [Check EDI AGN](https://rattanao.github.io/check-edi-AGN/))

## ตรวจอะไรบ้าง

- CONSIGNEE, CONTAINER NO., STATUS (CY / LCL / LCL-CFS), TOTAL PACKAGE, PACKAGING
- GROSS WEIGHT, MEASUREMENT
- MARKS, DESCRIPTION — เทียบทีละคู่ ไม่นับช่องว่าง/วรรคตอน แต่คำที่เกิน/ขาดต้องเตือน (สีส้ม)
- REEFER TEMP, DG (CLASS/UN), TRANSIT/TRANSHIPMENT — พบฝั่งใดก็เตือนเสมอ
- SHED NO., ประเทศปลายทาง, CARGO MOVEMENT, TAX ID เทียบ NOTIFY
- ยอดรวมทั้งฉบับเทียบกับยอดที่แต่ละเอกสารระบุเอง และ B/L ที่มีใน ENTER แต่ตกหล่นจาก MANIFEST (วิกฤต)

สี: แดง ⚠ = ไม่ตรง · ส้ม = ต้องให้คนยืนยัน · เหลือง = ไม่มี ENTER เทียบ · เขียว = ผ่าน (คอลัมน์สรุป)

## ใช้งานแบบโปรแกรม Python (ในเครื่อง)

```bash
pip install -r requirements.txt
python build_edi.py --manifest MANIFEST.xls --enter ENTER.pdf --outdir .
```

หรือวาง MANIFEST + ENTER ไว้ในโฟลเดอร์ `input/` แล้วรัน `python build_edi.py` (หาไฟล์ให้อัตโนมัติ)

## วิธีเปิดเว็บให้เพื่อนใช้ (GitHub Pages)

1. ขึ้น GitHub: repo → **Settings → Pages**
2. **Build and deployment → Source:** *Deploy from a branch*
3. **Branch:** `main` และโฟลเดอร์ `/ (root)` → กด **Save**
4. รอ ~1 นาที ลิงก์จะขึ้นที่หัวหน้า Pages: `https://rattanao.github.io/check-edi-Quality/`
   (หน้าแรกของ repo จะเด้งไป `docs/` เอง)
5. ส่งลิงก์ให้เพื่อนได้เลย — แก้โค้ดแล้ว `git push` หน้าเว็บอัปเดตเองภายในไม่กี่นาที

## หมายเหตุเรื่องข้อมูลลูกค้า

`input/`, `output/` และไฟล์ `.xls/.xlsx/.pdf` ถูก git-ignore ไว้ — repo นี้เก็บเฉพาะโปรแกรม
ไม่เก็บเอกสารจริงของลูกค้า
