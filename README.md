# Check EDI Freight Links

โปรแกรมตรวจเอกสารขาเข้าเรือก่อนทำใบขนสินค้าขาเข้า — เทียบ **MANIFEST** (`.xls`/`.xlsx`
ของสายเรือ) กับ **ENTER** (`.pdf` ของลูกค้า/forwarder รวมข้อความสีแดงที่พิมพ์เพิ่มบน PDF)
แล้วออกรายงาน **`EDI.xlsx`** แบบใส่สี ทีละ B/L:

- CONSIGNEE, CONTAINER NO., STATUS (CY / LCL / LCL-CFS), TAX ID vs NOTIFY
- TOTAL PACKAGE + ชนิดหีบห่อ, GROSS WEIGHT, MEASUREMENT
- MARKS, DESCRIPTION
- REEFER TEMP, DG (CLASS/UN), TRANSIT/TRANSHIPMENT (เจอเมื่อไหร่ขึ้นแดงทันที)
- SHED NO., ประเทศปลายทาง, CARGO MOVEMENT
- แถว TOTAL เทียบยอดรวมกับที่เอกสารแต่ละฉบับระบุเอง
- แถววิกฤตสำหรับ B/L ที่มีใน ENTER แต่หายไปจาก MANIFEST

**ความหมายของสี:** 🟥 แดง = ไม่ตรง · 🟧 ส้ม = ต้องให้คนตรวจยืนยัน · 🟨 เหลือง = ไม่มี ENTER ให้เทียบ · 🟩 เขียว = ผ่าน

รองรับ ENTER หลายแบบฟอร์ม (Layout A–M) รวมแบบ FREIGHT LINKS EXPRESS "IMPORT CARGO MANIFEST"
(หน้าแรกเป็นจดหมายปะหน้า ช่องติ๊ก [ X ] และข้อความสีแดงที่พิมพ์เพิ่ม = STATUS ของทุก B/L)
— รายละเอียดอยู่ใน [`references/format-notes.md`](references/format-notes.md)

---

## ใช้ออนไลน์ (ง่ายสุด ไม่ต้องติดตั้ง)

เปิด **<https://rattanao.github.io/Freight-links/>** → ลากไฟล์ MANIFEST + ENTER (ฟอร์ม IMPORT CARGO MANIFEST ของ Freight Links) ใส่ → ใส่ SHED NO. (ถ้ามี) → กด **ตรวจสอบ** → กด **ดาวน์โหลด EDI.xlsx**

ไฟล์ถูกประมวลผลในเบราว์เซอร์ของผู้ใช้เท่านั้น ไม่มีการอัปโหลดไปที่ใด · คู่มือ: <https://rattanao.github.io/Freight-links/guide.html>

---

## วิธีที่ 1 — ใช้แบบดับเบิลคลิก (ไม่ต้องมี Claude)

### ติดตั้งครั้งแรก (ทำครั้งเดียว)

1. ติดตั้ง **Python 3** จาก <https://www.python.org/downloads/>
   - หน้าแรกของตัวติดตั้ง **ติ๊ก ☑ "Add python.exe to PATH"** ก่อนกด Install
2. ดาวน์โหลดโปรแกรมนี้: หน้า repo กดปุ่มเขียว **Code → Download ZIP**
   แล้วแตกไฟล์ไว้ที่ไหนก็ได้ (เช่น `Documents\Check EDI Freight Links`)
   - หรือถ้ามี git: `git clone https://github.com/Rattanao/Freight-links.git`

### ใช้งานทุกครั้ง

1. ลบไฟล์เก่าในโฟลเดอร์ `input\` ออก แล้ววางไฟล์ใหม่ 2 ไฟล์:
   - MANIFEST (`.xls` หรือ `.xlsx`) — ชื่อขึ้นต้นด้วย `MANIFEST` จะดีที่สุด
   - ENTER (`.pdf`) — ชื่อขึ้นต้นด้วย `ENTER` จะดีที่สุด
2. ดับเบิลคลิก **`run-check-edi.bat`**
   - ครั้งแรกจะติดตั้งไลบรารีให้เอง (ต้องต่ออินเทอร์เน็ต)
   - ถาม SHED NO. ของงาน — พิมพ์เลข เช่น `0141` แล้ว Enter, หรือกด Enter ข้ามถ้าไม่รู้
3. เปิดผลที่ **`output\EDI.xlsx`** — หน้าจอดำจะสรุปจุดที่ต้องตรวจให้ด้วย
   - ถ้าเปิด `EDI.xlsx` ค้างไว้ใน Excel ต้องปิดก่อนรันใหม่ ไม่งั้นบันทึกทับไม่ได้

### ใช้แบบพิมพ์คำสั่งเอง

```bash
pip install -r requirements.txt
python scripts/build_edi.py --manifest "input/MANIFEST.xls" --enter "input/ENTER.pdf" --outdir output --shed 0141
```

---

## วิธีที่ 2 — ใช้เป็น Skill ใน Claude Code

ให้ Claude ตรวจและอธิบายผลให้ (และแก้โปรแกรมเองได้ถ้าเจอแบบฟอร์ม ENTER ใหม่)

1. ติดตั้ง Python และไลบรารีตามวิธีที่ 1 (`pip install -r requirements.txt`)
2. คัดลอกโฟลเดอร์นี้ไปไว้ที่โฟลเดอร์ skills ของ Claude Code:

   ```bash
   git clone https://github.com/Rattanao/Freight-links.git ~/.claude/skills/check-edi-freight-links
   ```

   (Windows: `C:\Users\<ชื่อผู้ใช้>\.claude\skills\check-edi-freight-links\` — ต้องมีไฟล์ `SKILL.md` อยู่ในโฟลเดอร์นี้ตรงๆ)
3. เปิด Claude Code ใหม่ ไปที่โฟลเดอร์งานที่มี MANIFEST + ENTER แล้วพิมพ์ เช่น
   `ตรวจ EDI ให้หน่อย SHED NO. 0141` หรือ `เทียบ MANIFEST กับ ENTER`

### อัปเดตเวอร์ชันใหม่

- ดาวน์โหลด ZIP: โหลดใหม่แล้วแตกทับ (เก็บโฟลเดอร์ `input\` ไว้ได้)
- git: `git pull` ในโฟลเดอร์ที่ clone ไว้

---

## โครงสร้างไฟล์

```
Freight-links/
├── run-check-edi.bat           ← ดับเบิลคลิกเพื่อรัน
├── requirements.txt            ← ไลบรารี Python ที่ต้องใช้
├── input/                      ← วาง MANIFEST + ENTER ตรงนี้ (ไม่ถูกอัปโหลดขึ้น GitHub)
├── output/                     ← EDI.xlsx ออกมาที่นี่
├── scripts/build_edi.py        ← ตัวโปรแกรม
├── SKILL.md                    ← คำสั่งสำหรับ Claude Code + กติกาของรายงาน
└── references/format-notes.md  ← ตัวอย่างรูปแบบ MANIFEST / ENTER ที่รองรับ
```

## ถ้าเจอปัญหา

| อาการ | วิธีแก้ |
|---|---|
| `Python was not found` | ติดตั้ง Python ใหม่ และติ๊ก "Add python.exe to PATH" |
| `PermissionError: [Errno 13]` | ปิดไฟล์ `EDI.xlsx` ใน Excel แล้วรันใหม่ |
| `ไม่พบไฟล์ MANIFEST / ENTER` | ตรวจว่าวางไฟล์ไว้ใน `input\` และนามสกุลถูก (.xls/.xlsx และ .pdf) |
| ฝั่ง ENTER ว่างทั้งแถว / ขึ้นเหลืองเยอะผิดปกติ | ENTER เป็นแบบฟอร์มใหม่ที่ยังไม่รองรับ — ส่งไฟล์ให้ Rattana หรือให้ Claude Code (วิธีที่ 2) เพิ่มแบบฟอร์มให้ |

> ผลลัพธ์เป็นตัวช่วยตรวจ ไม่ใช่ตัวตัดสิน — ก่อนยื่นใบขนต้องให้คนตรวจทุกจุดที่ขึ้นแดง/ส้มเสมอ
