#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_edi.py — เทียบ MANIFEST (สายเรือ) กับ ENTER (เอกสารลูกค้า) แล้วสร้าง EDI.xlsx
ก่อนจัดทำใบขนสินค้าขาเข้า (Thailand ocean import).

ใช้งาน:
    python build_edi.py --manifest MANIFEST.xls --enter ENTER.pdf --outdir .
    python build_edi.py                                   # หาไฟล์ .xls/.xlsx และ .pdf ในโฟลเดอร์
                                                            # input/ (หรือโฟลเดอร์ปัจจุบัน) อัตโนมัติ

รูปแบบไฟล์ที่รองรับ (ดู references/format-notes.md ในสกิลนี้สำหรับรายละเอียด):
  MANIFEST.xls  — รายงานอิสระ (free-form) 1 บล็อกต่อ 1 B/L ย่อย
                  col0=B/L (แถวหัว) / "S :" "C :" "N :" | col5=marks+container
                  col10=จำนวนหีบห่อ+STATUS+รายละเอียดสินค้า(+DG/REEFER/TRANSIT ถ้ามี)
                  col16=น้ำหนัก(แถวหัว)/ปริมาตร(MTQ)
                  ปรับเลขคอลัมน์ในฟังก์ชัน parse_manifest() ถ้าพบว่าไฟล์คาร์เรียร์อื่นเรียงคอลัมน์ต่างไป
  ENTER.pdf     — เอกสารยืนยัน/แก้ไขจากลูกค้า 1 บล็อกต่อ 1 B/L (MARK&NO. / NO. OF PKGS /
                  DESCRIPTION OF GOODS / CNTR : / B/L NO. : / ชื่อผู้รับ) — อาจมีข้อความ
                  สีแดงที่พิมพ์เพิ่ม (FreeText annotation) เหนือแต่ละบล็อก เช่น STATUS ที่แก้ไข,
                  INTRANSIT/TRANSHIPMENT, DG CLASS/UN, REEFER TEMP — ดึงอัตโนมัติตามตำแหน่ง y
                  ในหน้า (ครึ่งบนของหน้า = บล็อกแรก, ครึ่งล่าง = บล็อกที่สอง)

จุดที่ตรวจ: CONSIGNEE, CONTAINER NO., STATUS (CY/LCL/LCL-CFS), TOTAL PACKAGE, PACKAGING
(ชนิดบรรจุภัณฑ์), GROSS WEIGHT, MEASUREMENT, MARKS, DESCRIPTION, REEFER TEMP, DG (CLASS/UN),
TRANSIT/TRANSHIPMENT, และยอดรวมทั้งฉบับเทียบกับยอดที่เอกสารต้นฉบับระบุเอง.

ต้องมี: openpyxl, xlrd>=2.0, pymupdf
"""
from __future__ import annotations
import argparse, glob, os, re, sys
from collections import Counter

import xlrd, pymupdf, openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ═══════════════════════════ regex / ตัวช่วยที่ใช้ร่วมกัน ═══════════════════════════
CONT_RE    = re.compile(r'[A-Z]{4}\d{6,7}')                      # เลขตู้ ISO 6346
TRANSIT_RE = re.compile(r'IN[\s-]?TRANSIT|TRANS[\s-]?S?HIP(?:MENT)?|TRANSIT\s*PORT', re.I)
DG_RE      = re.compile(r'\bCLASS\s*[0-9]|\bUN\s*[- ]?\d{3,4}\b', re.I)


def find_dg(*txts):
    for t in txts:
        if t and DG_RE.search(t):
            cl = re.search(r'CLASS\s*([0-9](?:\.[0-9])?)', t, re.I)
            un = re.search(r'UN\s*[- ]?\s*(\d{3,4})', t, re.I)
            return ' '.join(x for x in ((f'CLASS {cl.group(1)}' if cl else ''),
                                        (f'UN {un.group(1)}' if un else '')) if x) or t.strip()
    return ''


def find_temp(*txts):
    for t in txts:
        if not t:
            continue
        if re.search(r'REEFER|TEMP|อุณหภูมิ', t, re.I):
            m = re.search(r"(-?\d{1,2}(?:\.\d)?)\s*['°]?\s*(?:deg\.?)?\s*C\b", t, re.I)
            return (m.group(1).lstrip('+') + '°C') if m else re.sub(r'\s+', ' ', t).strip()
    return ''


def extract_transit_phrase(txt):
    """ดึงเฉพาะใจความสำคัญ 'INTRANSIT TO <ปลายทาง>' / 'TRANSHIPMENT TO <ปลายทาง>' ออกจากประโยคเต็ม —
       บางเอกสารเขียนยาว เช่น 'IN TRANSIT CARGO FROM HOCHIMINH PORT, VIETNAM VIA LAEM CHABANG PORT
       BY SEA TO MYWADDY, MYANMAR BY TRUCK' ซึ่งมีรายละเอียดเส้นทางที่ไม่จำเป็นต้องขึ้นจอ ตัดเหลือแค่
       คำสำคัญ + ปลายทางจริง (คำ TO ตัวสุดท้ายก่อนเจอ VIA/BY/FROM หรือจบประโยค)."""
    if not txt:
        return ''
    m = re.search(r'(IN\s*TRANSIT|TRANS\s?S?HIP(?:MENT)?)', txt, re.I)
    if not m:
        return re.sub(r'\s+', ' ', txt).strip()
    kw = 'INTRANSIT' if 'TRANSIT' in m.group(1).upper().replace(' ', '') else 'TRANSHIPMENT'
    rest = txt[m.end():]
    tos = list(re.finditer(r'\bTO\s+([A-Za-z][A-Za-z,.\s]*?)(?=\s+(?:VIA|BY|FROM)\b|["\.]|$)', rest, re.I))
    if tos:
        dest = re.sub(r'\s+', ' ', tos[-1].group(1)).strip(' ,.')
        return f'{kw} TO {dest}'
    return kw


PKG_KIND = {'CS': 'CASE', 'PX': 'PALLET', 'PL': 'PALLET', 'CT': 'CARTON', 'CTN': 'CARTON',
            'BX': 'BOX', 'PK': 'PACKAGE', 'PKG': 'PACKAGE', 'SX': 'SET', 'DR': 'DRUM',
            'RO': 'ROLL', 'RL': 'ROLL', 'BG': 'BAG', 'BE': 'BALE', 'BL': 'BALE',
            'CR': 'CRATE', 'UN': 'UNIT', 'PC': 'PIECE'}


def canon_pkg(s):
    """ชนิดบรรจุภัณฑ์มาตรฐาน — ตัดตัวเลข/รหัสย่อ/'(s)' ออก ('5 PX (PALLET(s))' -> 'PALLET')."""
    if not s:
        return ''
    u = s.upper()
    for w in ('WOODEN CASE', 'PALLET', 'CARTON', 'PACKAGE', 'CASE', 'BOX', 'SET', 'DRUM',
              'ROLL', 'BAG', 'BALE', 'CRATE', 'UNIT', 'PIECE', 'SKID', 'BUNDLE'):
        if w in u:
            return 'CASE' if w == 'WOODEN CASE' else w
    m = re.match(r'^\s*[\d,]*\s*([A-Z]{2,3})\b', u)
    if m and m.group(1) in PKG_KIND:
        return PKG_KIND[m.group(1)]
    m = re.search(r'\(([A-Z]+?)\(S\)\)', u) or re.search(r'\(([A-Z ]+?)\)', u)
    if m:
        return (m.group(1).strip().rstrip('S')) or m.group(1).strip()
    return re.sub(r'[\d,()]', '', u).strip()


def canon_status(s):
    """คืน CY / LCL / LCL/CFS.
       CY,CY/CY,FCL,ลากตู้=CY | LCL,ขน(ส่ง)=LCL | CFS,LCL/CFS,เปิดตู้=LCL/CFS.
       ปรับ 3 เงื่อนไข if แรกได้ถ้าลูกค้าใช้คำไทยอื่นสำหรับ STATUS."""
    if not s:
        return ''
    th = str(s)
    if 'เปิดตู้' in th:
        return 'LCL/CFS'
    if 'ลากตู้' in th:
        return 'CY'
    if 'ขน' in th:
        return 'LCL'
    u = re.sub(r'STATUS', '', th, flags=re.I)
    u = re.sub(r'[^A-Za-z/]', '', u).upper()
    if not u:
        return th.strip()
    if 'CFS' in u:
        return 'LCL/CFS'
    if u == 'LCL':
        return 'LCL'
    if 'LCL' in u and 'CFS' in u:
        return 'LCL/CFS'
    if 'CY' in u or 'FCL' in u:
        return 'CY'
    if 'LCL' in u:
        return 'LCL'
    return u


def num(s):
    # \d นำหน้าบังคับให้ต้องมีตัวเลขจริงอย่างน้อย 1 ตัว — ถ้าใช้ [\d,]+ เฉยๆ สตริงที่มีแต่ "," ลอยๆ
    # (ไม่มีเลขเลย) จะแมตช์ติดแล้ว float('') พังตอน replace(',', '') เหลือสตริงว่าง
    m = re.search(r'-?\d[\d,]*\.?\d*', str(s or ''))
    return float(m.group(0).replace(',', '')) if m else None


def pkgnum(s):
    m = re.search(r'\d[\d,]*', str(s or ''))
    return int(m.group(0).replace(',', '')) if m else None


def norm_co(s):
    s = re.sub(r'[^A-Z0-9 ]', ' ', str(s or '').upper())
    s = re.sub(r'\b(CO|LTD|COMPANY|LIMITED|PUBLIC|CORP|CORPORATION|INC|GROUP|THE)\b', ' ', s)
    return re.sub(r'\s+', ' ', s).strip()


def comatch(a, b):
    """ตรงกันเมื่อ: ชื่อเหมือน / ฝั่งสั้นเป็นส่วนย่อยของฝั่งยาว (≥2 คำ) / คำซ้ำ ≥2."""
    sa, sb = set(a.split()), set(b.split())
    if not sa or not sb:
        return False
    if sa == sb:
        return True
    inter = sa & sb
    if inter and (sa <= sb or sb <= sa) and min(len(sa), len(sb)) >= 2:
        return True
    return len(inter) >= 2


_DTL_STOP = {'MADE', 'KOREA', 'THAILAND', 'ATTACHED', 'SHEET', 'PER', 'NOS', 'NUMBER', 'MARK',
             'MARKS', 'PLT', 'PLTS', 'BANGKOK', 'GROSS', 'NET', 'WEIGHT', 'MADEIN', 'AND',
             'THE', 'FOR', 'OF', 'PALLET', 'PALLETS', 'CARTON', 'CARTONS', 'CASE', 'CASES',
             'BOX', 'BOXES', 'PACKAGE', 'PACKAGES', 'SET', 'SETS', 'WOODEN', 'DRUM', 'ROLL',
             'BAG', 'BALE', 'PCS', 'PIECE', 'CTNO', 'CNO', 'PONO', 'PO', 'INV', 'ITEM',
             'COLOR', 'QTY', 'SIZE', 'DIA', 'CONSIGNEE', 'SHIPPER', 'NAME', 'PRODUCT', 'LOT'}


def toks(s):
    s = re.sub(r'[^A-Z0-9 ]', ' ', str(s or '').upper())
    return {w for w in s.split() if len(w) > 2 and w not in _DTL_STOP and not w.isdigit()}


def same_detail(a, b):
    """เทียบ MARKS/DESCRIPTION แบบเข้ม — ต่างกันแม้แต่นิดเดียวต้องเตือน.
       ยอมให้ต่างได้เฉพาะช่องว่าง/เครื่องหมายวรรคตอน (กัน MANIFEST ตัดคำขึ้นบรรทัดใหม่).
       True = ตรงกัน ; False = ไม่ตรง (เตือน) ; None = ไม่มีข้อความทั้งคู่."""
    na = re.sub(r'[^A-Z0-9]', '', str(a or '').upper())
    nb = re.sub(r'[^A-Z0-9]', '', str(b or '').upper())
    if not na and not nb:
        return None
    return na == nb


def same_bag(a, b):
    """เทียบชุดคำแบบไม่สนใจลำดับ — ใช้เฉพาะตอนเทียบข้อความรวม MARKS+DESCRIPTION ของ ENTER
       เลย์เอาต์ที่ไม่มีคอลัมน์แยก (combined_text) กับ MARKS+DESCRIPTION ของ MANIFEST ที่ถูกแบ่ง
       เป็น 2 คอลัมน์ตามตำแหน่งการตัดบรรทัดของ Excel ไม่ใช่ลำดับจริงของเอกสารต้นฉบับ — ยังต้องมี
       คำครบเหมือนกันทุกคำ (นับซ้ำด้วย) เพียงแต่ไม่สนใจว่าใครมาก่อนมาหลัง.
       True = ตรงกัน ; False = ไม่ตรง (ให้ผู้ตรวจสอบคนดูเอง) ; None = ไม่มีข้อความทั้งคู่."""
    wa = sorted(re.sub(r'[^A-Z0-9]', ' ', str(a or '').upper()).split())
    wb = sorted(re.sub(r'[^A-Z0-9]', ' ', str(b or '').upper()).split())
    if not wa and not wb:
        return None
    return wa == wb


# ═══════════════════════════ 1) MANIFEST.xls ═══════════════════════════
def detect_bl_prefix(rows):
    """เดา prefix ของเลข B/L จากคอลัมน์ 0 (เช่น 'HASL' ของ Heung-A) โดยหาสตริงรูปแบบ
       ตัวอักษรนำ+ตัวเลข ที่ซ้ำกันบ่อยที่สุด ไม่ต้องแก้โค้ดเองเมื่อเปลี่ยนสายเรือ.
       คืน compiled regex ที่แมตช์ B/L ของไฟล์นี้."""
    cand = []
    for row in rows:
        v = row[0] if row else ''
        if isinstance(v, float):
            continue
        s = str(v).strip()
        m = re.match(r'^([A-Z]{3,8})[A-Z0-9]{5,}$', s)
        if m:
            cand.append(m.group(1))
    if cand:
        # ใช้ longest-common-prefix ของตัวอักษรนำทุกตัวที่เจอ แทนตัวที่ซ้ำมากที่สุดตัวเดียว —
        # สายเรือเดียวกันมักมีรหัสบริการย่อยต่อท้าย prefix จริงหลายแบบ (เช่นของ Heung-A: HASLK,
        # HASLC, HASLJ, HASLS ล้วนใช้ prefix จริงร่วมกันแค่ 'HASL') ถ้าเลือกตัวที่ฮิตสุดตัวเดียว
        # (regex เดิม) B/L ที่ใช้รหัสย่อยอื่นจะไม่แมตช์และหลุดจากการตรวจไปเงียบๆ ทั้งกลุ่ม
        lcp = re.match(r'^[A-Z]*', os.path.commonprefix(cand)).group(0)
        prefix = lcp if len(lcp) >= 3 else Counter(cand).most_common(1)[0][0]
        return re.compile(rf'^{re.escape(prefix)}[A-Z0-9]{{5,}}$')
    return re.compile(r'^[A-Z]{3,8}[A-Z0-9]{6,}$')   # fallback ทั่วไป


def _cell(row, i):
    if i >= len(row):
        return ''
    v = row[i]
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    # แทนที่ (ไม่ใช่ลบทิ้ง) อักขระควบคุม/บรรทัดใหม่ในเซลล์ด้วยช่องว่าง — กันคำสองบรรทัดใน
    # เซลล์ที่ตัดคำ (word-wrap) ถูกเชื่อมติดกัน เช่น 'DP-14\nMODEL' กลายเป็น 'DP-14MODEL'
    # (same_detail() เทียบแบบตัดช่องว่างทิ้งอยู่แล้วจึงไม่กระทบ, แต่ same_bag() เทียบทีละคำต้องการช่องว่างนี้)
    return re.sub(r'[\x00-\x1f]+', ' ', str(v)).strip()


def parse_manifest(path):
    """คืน (records:dict[bl->rec], vessel:str, declared:dict).
       ปรับเลขคอลัมน์ตรงนี้ถ้าไฟล์คาร์เรียร์อื่นจัดหน้าต่างจากที่คอมเมนต์ไว้ด้านบน."""
    book = xlrd.open_workbook(path, ragged_rows=True)
    sh = book.sheet_by_index(0)
    rows = [sh.row_values(r) for r in range(sh.nrows)]
    BL_RE = detect_bl_prefix(rows)

    vessel = ''
    declared = {}   # ยอดรวมที่ MANIFEST ระบุเอง (GRAND TOTAL: ...) ถ้ามี
    MAN, i, n = {}, 0, len(rows)
    END = re.compile(r'^(PORT TOTAL|GRAND TOTAL|TOTAL\b|PortOfDischarge)', re.I)
    ST_TOK = ('CY/CY', 'CY-CY', 'LCL/CFS', 'CFS/CFS', 'FCL/CFS', 'CY', 'LCL', 'CFS', 'FCL')

    for ridx, row in enumerate(rows):
        v0 = _cell(row, 0)
        if v0.upper() == 'VESSEL & VOYAGE' and not vessel:
            for c in row[1:]:
                cs = str(c).strip()
                if cs:
                    vessel = cs
                    break
        row_txt = ' '.join(_cell(row, k) for k in range(len(row)))
        if re.search(r'GRAND TOTAL', row_txt, re.I):
            joined = ' '.join(' '.join(_cell(r2, k) for k in range(len(r2)))
                               for r2 in rows[ridx:ridx + 4])
            mpk = re.search(r'TOTAL PACK\s*:\s*([\d,]+)', joined, re.I)
            mgw = re.search(r'TOTAL GROSSWEIGHT\s*:\s*([\d,.]+)', joined, re.I)
            mme = re.search(r'TOTAL MEASUREMENT\s*:\s*([\d,.]+)', joined, re.I)
            if mpk: declared['pkg'] = int(mpk.group(1).replace(',', ''))
            if mgw: declared['gw'] = float(mgw.group(1).replace(',', ''))
            if mme: declared['meas'] = float(mme.group(1).replace(',', ''))

    while i < n:
        bl = _cell(rows[i], 0)
        if not BL_RE.match(bl):
            i += 1
            continue
        rec = {'bl': bl, 'pkg_hdr': _cell(rows[i], 10), 'gw': _cell(rows[i], 16),
               'cons': '', 'notify': '', 'cont': [], 'status': '', 'meas': '',
               'marks': [], 'desc': [], 'transit': '', 'dg': '', 'reefer': ''}
        h5 = _cell(rows[i], 5)
        if h5 and not re.match(r'^\d+\.\s*([A-Z]{4}\d{6,7}|\s*$)', h5):
            rec['marks'].append(h5)
        transit_open = False   # true ตั้งแต่พบคำ transit ในบล็อกนี้ครั้งแรก — ข้อความบรรทัดต่อจากนั้น
                                # (คำอธิบายท่า/ประเทศปลายทางที่พิมพ์ต่อลงมาหลายบรรทัด) ให้ต่อเข้า transit
                                # แทนที่จะตกไปเป็น description
        j = i + 1
        while j < n:
            a = _cell(rows[j], 0)
            if BL_RE.match(a):
                break
            if END.match(a):
                j += 1
                break
            if a.startswith('C :'):
                rec['cons'] = a[3:].strip()
            elif a.startswith('N :'):
                rec['notify'] = a[3:].strip()
            elif a.startswith('S :'):
                pass
            elif rec['cons'] and not rec['notify'] and not a.startswith(('1.', '2.', '3.')):
                rec['cons'] += ' ' + a
            elif rec['notify'] and not a.startswith(('1.', '2.', '3.')):
                rec['notify'] += ' ' + a
            c5 = _cell(rows[j], 5)
            if c5:
                mm = re.match(r'^\d+\.\s*([A-Z]{4}\d{6,7})', c5)
                if mm:
                    rec['cont'].append(mm.group(1))
                elif not re.match(r'^\d+\.\s*$', c5):
                    rec['marks'].append(c5)
            for cc in (_cell(rows[j], 5), _cell(rows[j], 10)):
                if cc and DG_RE.search(cc) and not rec['dg']:
                    rec['dg'] = find_dg(cc)
                if cc and re.search(r'REEFER|TEMP|อุณหภูมิ', cc, re.I) and not rec['reefer']:
                    rec['reefer'] = find_temp(cc)
            c10 = _cell(rows[j], 10)
            if c10:
                u = c10.upper()
                # DG/REEFER/TRANSIT ถูกดึงไปเก็บเป็นช่องแยกไว้แล้ว (ใช้ตรวจ/ไฮไลต์เฉพาะจุด) แต่ "ห้าม
                # ตัดออกจาก DESCRIPTION" ทุกกรณี — ต้องเห็น DESCRIPTION (MANIFEST) เหมือนข้อความจริงที่
                # อยู่ในคอลัมน์ "NO. OF PKGS DESCRIPTION OF GOODS" ของ MANIFEST ทุกตัวอักษร ส่วนที่ตัด
                # ทิ้งจริง ๆ มีแค่ตัวเลข/รหัสโครงสร้างที่ไม่ใช่ "รายละเอียดสินค้า" เลย (STATUS, รหัสชนิดตู้)
                if DG_RE.search(c10) or re.search(r'REEFER|TEMP|อุณหภูมิ', c10, re.I):
                    rec['desc'].append(c10)
                elif transit_open or TRANSIT_RE.search(u):
                    # ประโยค transit มักพิมพ์ต่อกันหลายบรรทัด (เช่น "IN TRANSIT CARGO FROM ...
                    # PORT," / "VIETNAM VIA ... TO" / "MYANMAR BY TRUCK") และมักเป็นเนื้อหาท้ายสุด
                    # ของบล็อกก่อนขึ้น B/L ถัดไป — เมื่อเจอคำ transit ครั้งแรกแล้ว ให้ถือว่าบรรทัด
                    # ที่เหลือในบล็อกนี้เป็นส่วนต่อของประโยคเดียวกันทั้งหมด ไม่ใช่ description ใหม่
                    rec['transit'] = (rec['transit'] + ' ' + c10).strip()
                    transit_open = True
                    rec['desc'].append(c10)
                elif u in ST_TOK and not rec['status']:
                    rec['status'] = c10
                elif re.match(r'^\d{0,2}[A-Z]\d[A-Z0-9]', u) or 'PART CNTR' in u \
                        or re.match(r'^[\d,]+\s*(PK|CS|PX|CT|BX|SX|BE|DR|RO|BG)\b', u):
                    pass
                else:
                    rec['desc'].append(c10)
            c16 = _cell(rows[j], 16)
            if c16 and 'MTQ' in c16.upper() and not rec['meas']:
                rec['meas'] = c16
            j += 1
        rec['marks'] = ' '.join(rec['marks'])
        rec['desc'] = ' '.join(rec['desc'])
        rec['cont'] = sorted(set(rec['cont']))
        MAN[bl] = rec
        i = j
    return MAN, vessel, declared, BL_RE


# ═══════════════════════════ 2) ENTER.pdf ═══════════════════════════
JUNK = {'MARK&NO.', 'PACKAGES', 'DESCRIPTION OF GOODS', 'NO. OF PKGS'}
_PAGELN = re.compile(r'^Page \d+ of \d+$')
PKG_WORDS = ('WOODEN CASE', 'PALLET', 'CARTON', 'PACKAGE', 'CASE', 'BOX', 'SET', 'DRUM',
             'ROLL', 'BAG', 'BALE', 'CRATE', 'UNIT', 'PIECE')


def parse_enter(path, bl_re):
    """คืน (records:dict[bl->rec], declared:dict).
       ข้อความสีแดงที่พิมพ์เพิ่ม (FreeText annotation) ถูกผูกกับ B/L ที่ใกล้ที่สุด "ด้านล่าง"
       ตำแหน่ง y ของมันในหน้าเดียวกัน (แต่ละหน้ามักมี 2 บล็อก บนกับล่าง)."""
    doc = pymupdf.open(path)

    # 1) ดึงข้อความพิมพ์เพิ่ม (FreeText) ผูกกับ B/L ตามตำแหน่ง y
    ent_annot = {}
    annot_raw = []   # (page_idx, ข้อความดิบเต็ม) ของทุก FreeText annotation — ใช้กรองไม่ให้ปนเข้า
                      # marks/desc ทีหลัง เก็บเลขหน้าไว้ด้วยเพื่อกรองแค่บรรทัดในหน้าเดียวกัน (ดูจุดใช้งาน
                      # ด้านล่าง — คำสั้นๆ ที่บังเอิญไปซ้ำกับคำในเนื้อความ annotation ของ B/L อื่นคนละหน้า
                      # ไม่ควรโดนกรองทิ้งไปด้วย)
    for p in range(doc.page_count):
        pg = doc[p]
        anchors = []
        for b in pg.get_text('dict')['blocks']:
            for l in b.get('lines', []):
                t = ''.join(s['text'] for s in l['spans']).strip()
                if bl_re.match(t):
                    anchors.append((l['bbox'][1], t))
        anchors.sort()
        if not anchors:
            continue
        def _apply_annot(bl, txt):
            rec = ent_annot.setdefault(bl, {'status': '', 'transit': '', 'dg': '', 'reefer': ''})
            if DG_RE.search(txt):
                rec['dg'] = find_dg(txt)
            elif re.search(r'REEFER|TEMP|อุณหภูมิ', txt, re.I):
                rec['reefer'] = find_temp(txt)
            elif TRANSIT_RE.search(txt):
                rec['transit'] = txt
            else:
                rec['status'] = (rec['status'] + ' ' + txt).strip()

        # บางครั้งลูกค้าพิมพ์ข้อความเดียวครอบคลุมทั้งช่วง B/L ย่อย เช่น
        # "HASLK01260707682-T    STATUS :เปิดตู้เข้าโกดัง" หมายถึงตั้งแต่ B/L ฐาน (682) ถึงตัวอักษร
        # ต่อท้าย T (682A,682B,...,682T) ทั้งหมด ไม่ใช่แค่บล็อกที่อยู่ใกล้ตำแหน่งนั้นที่สุด — บางครั้งก็
        # เขียนกลับด้าน คือคำสั่งมาก่อน แล้วค่อยตามด้วย "ฐาน-ตัวอักษรจบ" ท้ายสุด เช่น
        # "เปิดตู้เข้าโกดัง HASLK01260606400-S" (หมายถึง 400 ถึง 400S ทั้งหมด) จับทั้งสองทิศ
        RANGE_RE = re.compile(r'^([A-Z0-9]{6,}?)-([A-Z])\s+(.+)$', re.S)
        RANGE_RE2 = re.compile(r'^(.+?)\s+([A-Z0-9]{6,}?)-([A-Z])$', re.S)

        for a in (pg.annots() or []):
            if a.type[1] != 'FreeText':
                continue
            txt = re.sub(r'\s+', ' ', a.info.get('content', '')).strip()
            if not txt:
                continue
            rm = RANGE_RE.match(txt)
            rm2 = None if rm else RANGE_RE2.match(txt)
            if rm or (rm2 and bl_re.match(rm2.group(2))):
                # เก็บไว้กรองแค่ส่วนคำสั่ง (rest) ไม่เอาทั้งข้อความดิบที่ขึ้นต้นด้วยเลข B/L จริง —
                # ถ้าเอาทั้งก้อนไปกรอง บรรทัดเลข B/L ตัวจริงในเอกสาร (เช่น 'HASLK...682' เฉยๆ) จะโดน
                # กรองทิ้งไปด้วยเพราะมันเป็นสับสตริงของข้อความช่วงนี้พอดี (ไม่ว่าจะอยู่หัวหรือท้ายข้อความ)
                if rm:
                    base, end_letter, rest = rm.group(1), rm.group(2), rm.group(3)
                else:
                    rest, base, end_letter = rm2.group(1), rm2.group(2), rm2.group(3)
                annot_raw.append((p, rest))
                for c in [''] + [chr(x) for x in range(ord('A'), ord(end_letter) + 1)]:
                    _apply_annot(base + c, rest)
                continue
            annot_raw.append((p, txt))
            ay = a.rect.y0
            cand = [(y, bl) for y, bl in anchors if y > ay - 5] or anchors
            bl = min(cand)[1]
            _apply_annot(bl, txt)

    # 2) ดึงยอดรวมที่เอกสารระบุเอง (มักอยู่หน้าสุดท้าย: "PACKAGES :", "KGS :", "CBM :" — บาง
    #    เลย์เอาต์ค่าตัวเลขมาก่อนป้ายกำกับของมันเอง เช่น "...KGS : 36,190.15 91.255 CBM :" จึงลองจับ
    #    ทั้งบล็อกก่อน แล้วค่อย fallback ไปจับทีละฟิลด์ถ้ารูปแบบไม่ตรง)
    declared = {}
    full_text = '\n'.join(doc[p].get_text() for p in range(doc.page_count))
    # บางเลย์เอาต์ (ไม่มีคอลัมน์ MARK&NO./DESCRIPTION OF GOODS แยก) เขียนยอดรวมแบบ
    # 'GRAND TOTAL' ตามด้วยตัวเลขก่อนป้ายกำกับของมันเองทุกฟิลด์ ('102 PACKAGES ... KGS ... CBM')
    # — ลองจับรูปแบบนี้ก่อนเพราะเฉพาะเจาะจงกว่า แล้วค่อย fallback ไปแบบอื่น
    # NUM ต้องมีเลขอย่างน้อย 1 ตัว (ไม่ใช่แค่ "." ลอยๆ) และ .{0,200}? จำกัดระยะไม่ให้ข้ามบล็อก
    # B/L อื่นที่อยู่ไกลออกไปหลายหน้า — เอกสาร ENTER บางฉบับเป็นชุดจดหมายแก้ไขหลาย B/L ต่อกัน
    # (ไม่มียอดรวมทั้งฉบับจริงๆ) ถ้าปล่อยให้ข้ามได้ไม่จำกัดจะจับเลขจากคนละบล็อกมาปนกัน
    NUM = r'\d[\d,.]*'
    gt = re.search(rf'GRAND\s+TOTAL\s*({NUM})\s*PACKAGES\s*({NUM})\s*KGS.{{0,200}}?({NUM})\s*CBM',
                    full_text, re.I | re.S)
    blk = None if gt else re.search(
        rf'PACKAGES\s*:?\s*({NUM}).{{0,200}}?KGS\s*:?\s*({NUM})\s+({NUM})\s*CBM', full_text, re.I | re.S)
    if gt:
        declared['pkg'] = int(float(gt.group(1).replace(',', '')))
        declared['gw'] = float(gt.group(2).replace(',', ''))
        declared['meas'] = float(gt.group(3).replace(',', ''))
    elif blk:
        declared['pkg'] = int(float(blk.group(1).replace(',', '')))
        declared['gw'] = float(blk.group(2).replace(',', ''))
        declared['meas'] = float(blk.group(3).replace(',', ''))
    else:
        # ต้องเจอครบทั้ง 3 ฟิลด์ถึงจะถือว่าเป็นยอดรวมทั้งฉบับจริง — ถ้าเจอแค่ฟิลด์เดียวลอยๆ
        # มีโอกาสสูงว่าเป็นตัวเลขของ B/L ย่อยตัวใดตัวหนึ่ง ไม่ใช่ยอดรวมทั้งฉบับ
        mpk = re.search(rf'PACKAGES\s*:\s*({NUM})', full_text, re.I)
        mgw = re.search(rf'KGS\s*:\s*({NUM})', full_text, re.I)
        mme = re.search(rf'CBM\s*:\s*({NUM})', full_text, re.I)
        if mpk and mgw and mme:
            declared['pkg'] = int(float(mpk.group(1).replace(',', '')))
            declared['gw'] = float(mgw.group(1).replace(',', ''))
            declared['meas'] = float(mme.group(1).replace(',', ''))

    # 3) ดึง marks/pkgs/desc/gw/meas/consignee/container ต่อบล็อก B/L
    # ใช้ get_text('dict') แทน get_text() ธรรมดา เพื่อเก็บพิกัด x0 ของแต่ละบรรทัดไว้คู่กับข้อความ —
    # เลย์เอาต์แบบที่ 2 (ไม่มี anchor 'B/L NO. :' เว้นวรรค) ยังอาจมีคอลัมน์ MARKS & NOS /
    # DESCRIPTIONS OF GOODS แยกจริงตามภาพ (แค่ anchor คนละแบบ) — x0 บอกได้ชัดว่าบรรทัดไหนอยู่
    # คอลัมน์ซ้าย (MARKS) กับคอลัมน์ขวา (DESCRIPTION) โดยไม่ต้องเดาจากลำดับบรรทัด
    lines = []
    xs = []
    # กรองด้วยข้อความ annotation ดิบ (annot_raw) เทียบทั้งสองทาง เพราะข้อความช่วง B/L (เช่น
    # "HASLK...-T    STATUS :...") ถูกเก็บใน ent_annot เป็นแค่ส่วน "rest" (ตัดคำนำหน้าช่วงออกแล้ว)
    # ซึ่งสั้นกว่าบรรทัดจริงในหน้า PDF — เทียบแบบ substring ทั้งสองทิศจึงจะจับได้ครบ
    # สำคัญ: กรองแค่ "หน้าเดียวกัน" กับ annotation นั้นเท่านั้น (ไม่ใช่ทั้งเอกสาร) — ไม่งั้นคำสั้นๆ ที่
    # บังเอิญซ้ำกับข้อความ annotation ของ B/L อื่นคนละหน้า (เช่น ประเทศปลายทาง "THAILAND" ที่เป็น
    # บรรทัดจริงของบล็อกหนึ่ง แต่ก็บังเอิญเป็นส่วนหนึ่งของประโยค annotation ยาวๆ ของอีกบล็อกในหน้าอื่น)
    # จะโดนกรองทิ้งไปทั้งเอกสารทั้งที่ไม่เกี่ยวข้องกันเลย
    annot_texts_by_page = {}
    for p, txt in annot_raw:
        annot_texts_by_page.setdefault(p, set()).add(txt)
    for p in range(doc.page_count):
        annot_texts = annot_texts_by_page.get(p, set())
        for b in doc[p].get_text('dict')['blocks']:
            for l in b.get('lines', []):
                s = ''.join(sp['text'] for sp in l['spans']).rstrip()
                st = s.strip()
                if not st or st in JUNK or _PAGELN.match(st):
                    continue
                # บรรทัดที่เป็นเลข B/L จริง (ตรง bl_re) ห้ามโดนกรองทิ้งเด็ดขาด แม้จะบังเอิญเป็น
                # substring ของข้อความ annotation อื่นในหน้าเดียวกัน (เช่น annotation ช่วง B/L ที่พิมพ์
                # กลับด้าน "<คำสั่ง> HASLK...400-S" ซึ่งมี "HASLK...400" ฝังอยู่ท้ายข้อความพอดี) — บรรทัด
                # นี้เป็นจุดยึด (anchor) ของขอบเขตบล็อกที่ codeต้องใช้ตัดบล็อก ถ้าหายไปบล็อกทั้งก้อนจะ
                # เพี้ยน/หายไปทั้งบล็อก
                if not bl_re.match(st) and \
                        any(st == v or st in v or (len(v) >= 8 and v in st) for v in annot_texts):
                    continue                                  # ตัดข้อความ FreeText ที่ปนมาท้ายหน้า
                lines.append(s)
                xs.append(l['bbox'][0])

    # anchor คำว่า 'B/L NO. :' / 'CONSIGNEE :' ฯลฯ เป็นแค่ป้ายกำกับสั้น ๆ ที่บังเอิญไปปรากฏในแบบฟอร์ม
    # อื่นที่ไม่ใช่ layout A/B จริงได้ง่าย (เอกสารรวมจดหมายจากหลาย forwarder มักมีคำว่า "B/L NO." ซ้ำ
    # กันในหลายสิบหน้าโดยที่โครงสร้างจริงไม่ใช่ layout A/B เลยสักหน้า) — นับจำนวน anchor ที่เจอเฉย ๆ
    # จึงหลอกได้ง่าย ต้องดูอัตราความสำเร็จ (กี่ anchor ที่ parse ออกมาเป็น B/L ที่ถูกต้องจริง) แทน —
    # layout ที่ใช่จริงทั้งฉบับควรสำเร็จแทบทุก anchor ที่เจอ ไม่ใช่แค่ไม่กี่เปอร์เซ็นต์
    LAYOUT_TRUST_MIN_N, LAYOUT_TRUST_MIN_RATE = 3, 0.5

    def _trusted(n_success, n_anchor):
        return n_success >= LAYOUT_TRUST_MIN_N and n_anchor > 0 and (n_success / n_anchor) >= LAYOUT_TRUST_MIN_RATE

    ENT = {}
    ENT_A = {}
    idx = [k for k, l in enumerate(lines) if l.strip() == 'B/L NO. :']
    if idx:
        # เลย์เอาต์เดิม: ค่า B/L มาก่อนป้ายกำกับ 'B/L NO. :' ของมันเอง (มีคอลัมน์ MARK&NO./
        # DESCRIPTION OF GOODS แยกกันชัดเจน — ดู references/format-notes.md)
        prev = 0
        for k in idx:
            blk = lines[prev:k + 2]
            prev = k + 2
            cons = blk[-1].strip()
            try:
                bi = len(blk) - 1 - blk[::-1].index('B/L NO. :')
            except ValueError:
                continue
            bl = blk[bi - 1].strip()
            if not bl_re.match(bl):
                continue
            try:
                ci = len(blk) - 1 - blk[::-1].index('CNTR :')
            except ValueError:
                continue
            cm = CONT_RE.search(blk[ci - 1])
            cont = cm.group(0) if cm else ''
            gi = next((x for x in range(len(blk) - 1, -1, -1) if blk[x].strip().startswith('G.W')), None)
            mi = next((x for x in range(len(blk) - 1, -1, -1) if blk[x].strip().startswith('MEAS')), None)
            gw = num(blk[gi]) if gi is not None else None
            meas = num(blk[mi]) if mi is not None else None
            pkgtype, pt_i = '', None
            for x in range((gi or len(blk)) - 1, -1, -1):
                if blk[x].strip().upper() in PKG_WORDS:
                    pkgtype, pt_i = blk[x].strip(), x
                    break
            cnt, cnt_i = None, None
            for x in range((pt_i or gi or len(blk)) - 1, -1, -1):
                if re.fullmatch(r'[\d,]+', blk[x].strip()):
                    cnt = int(blk[x].strip().replace(',', ''))
                    cnt_i = x
                    break
            desc = ' '.join(l.strip() for l in blk[cnt_i + 1:pt_i] if l.strip()) if (cnt_i is not None and pt_i) else ''
            marks = ' '.join(l.strip() for l in blk[:cnt_i]
                              if l.strip() and l.strip() not in JUNK) if cnt_i is not None else ''
            ENT_A[bl] = {'bl': bl, 'cons': cons, 'cont': cont, 'pkgs': cnt, 'pkgtype': pkgtype,
                         'gw': gw, 'meas': meas, 'marks': marks, 'desc': desc,
                         'status_raw': ent_annot.get(bl, {}).get('status', ''),
                         'transit_raw': ent_annot.get(bl, {}).get('transit', ''),
                         'dg': ent_annot.get(bl, {}).get('dg', '') or find_dg(marks, desc),
                         'reefer': ent_annot.get(bl, {}).get('reefer', '') or find_temp(marks, desc)}
        # ไม่ return ที่นี่แม้ layout A จะแมตช์ได้บ้าง — เอกสาร ENTER บางฉบับเป็นการรวมจดหมายจาก
        # freight forwarder หลายรายในไฟล์เดียว แต่ละรายใช้ layout ต่างกัน (บางหน้า layout A, บางหน้า
        # layout B, บางหน้าไม่ตรงทั้งคู่) ต้องลองทุกวิธีทั่วทั้งเอกสารแล้วรวมผลลัพธ์ ไม่ใช่หยุดแค่วิธีแรก
        # ที่เจอ B/L สักตัว มิฉะนั้น B/L ส่วนใหญ่ในเอกสารผสมแบบนี้จะหลุดจากการตรวจไปเงียบ ๆ ทั้งหมด
    layout_a_trusted = _trusted(len(ENT_A), len(idx))
    for bl, rec in ENT_A.items():
        rec['low_confidence'] = not layout_a_trusted
        ENT[bl] = rec

    # เลย์เอาต์แบบที่ 2: ไม่มีคอลัมน์ MARK&NO./DESCRIPTION OF GOODS แยกกัน — แต่ละบล็อกคือ
    # [จำนวน+ชนิดหีบห่อ] / [น้ำหนัก KGS] / [ปริมาตร CBM] / ข้อความอิสระ (marks+desc ปนกัน) /
    # 'CONSIGNEE :' + ชื่อ / 'CONT.NO.:' + เลขตู้ / 'B/L NO.:' + เลข B/L (ป้ายกำกับมาก่อนค่าเสมอ)
    anchor_re = re.compile(r'^B\s*/\s*L\s*NO\.?\s*:?\s*$', re.I)
    cons_re = re.compile(r'^CONSIGNEE\s*:?\s*$', re.I)
    contlbl_re = re.compile(r'^CONT\.?\s*NO\.?\s*:?\s*$', re.I)
    pkg_re = re.compile(r'^([\d,]+)\s+([A-Za-z].*)$')
    idx2 = [k for k, l in enumerate(lines) if anchor_re.match(l.strip())]
    ENT_B = {}
    prev = 0
    for k in idx2:
        blk = lines[prev:k + 2]
        blk_xs = xs[prev:k + 2]
        prev = k + 2
        if not blk:
            continue
        bl = blk[-1].strip()
        if not bl_re.match(bl):
            continue
        cont_li = next((x for x in range(len(blk) - 1, -1, -1) if contlbl_re.match(blk[x].strip())), None)
        cont = ''
        if cont_li is not None and cont_li + 1 < len(blk):
            cm = CONT_RE.search(blk[cont_li + 1])
            cont = cm.group(0) if cm else blk[cont_li + 1].strip()
        cons_li = next((x for x in range(len(blk) - 1, -1, -1) if cons_re.match(blk[x].strip())), None)
        # ชื่อผู้รับบางรายยาวจนตัดขึ้นบรรทัดที่ 2 ในฟอร์ม (เช่น "...(THAILAND)" / "PUBLIC COMPANY
        # LIMITED." คนละบรรทัด) — รวมทุกบรรทัดตั้งแต่หลังป้าย 'CONSIGNEE :' ไปจนถึงป้าย 'CONT.NO.:'
        # ถัดไป ไม่ใช่หยิบมาแค่บรรทัดแรกบรรทัดเดียว
        cons_end = cont_li if (cont_li is not None and cons_li is not None and cont_li > cons_li) \
            else (cons_li + 2 if cons_li is not None else None)
        cons = ' '.join(l.strip() for l in blk[cons_li + 1:cons_end] if l.strip()) \
            if cons_li is not None and cons_end is not None else ''
        end_i = min(x for x in (cons_li, cont_li) if x is not None) if (cons_li is not None or cont_li is not None) \
            else max(len(blk) - 2, 0)
        cnt, pkgtype, gw, meas, pkg_i = None, '', None, None, None
        for x in range(end_i):
            m0 = pkg_re.match(blk[x].strip())
            if m0:
                pkg_i = x
                cnt = int(m0.group(1).replace(',', ''))
                pkgtype = re.sub(r'\bPLTS?\b', 'PALLET', m0.group(2).strip(), flags=re.I)
                pkgtype = re.sub(r'\bPKGS?\b', 'PACKAGE', pkgtype, flags=re.I)
                if x + 1 < end_i:
                    gw = num(blk[x + 1])
                if x + 2 < end_i:
                    meas = num(blk[x + 2])
                break
        desc_start = pkg_i + 3 if pkg_i is not None else 0
        # เอกสารนี้บางครั้งก็ยังมีคอลัมน์ MARKS & NOS (ซ้าย) กับ DESCRIPTIONS OF GOODS (ขวา) แยกจริง
        # ตามภาพ (แค่ anchor 'B/L NO.:' เป็นแบบไม่เว้นวรรค คนละจุดกับเลย์เอาต์แบบที่ 1) — ใช้ x0 ของ
        # แต่ละบรรทัด (บันทึกไว้คู่กับ lines ตอนดึงข้อความ) แยกว่าบรรทัดไหนอยู่คอลัมน์ไหน แทนการเดา
        # จากลำดับบรรทัด/จำนวนบรรทัด ซึ่งกลับกันได้ระหว่างบล็อก (ดู format-notes.md)
        LEFT_X_MAX = 150   # x0 < 150 = คอลัมน์ MARKS & NOS (ซ้าย, ชิดขอบเดียวกับป้าย CONSIGNEE/CONT.NO./B/L NO.)
        body = [(t.strip(), x) for t, x in zip(blk[desc_start:end_i], blk_xs[desc_start:end_i])
                if t.strip() and t.strip() not in JUNK]
        marks = ' '.join(t for t, x in body if x < LEFT_X_MAX)
        desc = ' '.join(t for t, x in body if x >= LEFT_X_MAX)
        combined = f'{marks} {desc}'.strip()
        # ยังตั้ง combined_text=True ไว้ (เทียบแบบ word-bag/ส้ม ไม่ใช่แดง) เพราะแม้แยกคอลัมน์ในเอกสาร
        # ENTER เองได้แม่นแล้ว แต่จุดตัด MARKS/DESCRIPTION ของ MANIFEST อาจไม่ตรงกับจุดตัดนี้เป๊ะ
        ENT_B[bl] = {'bl': bl, 'cons': cons, 'cont': cont, 'pkgs': cnt, 'pkgtype': pkgtype,
                     'gw': gw, 'meas': meas, 'marks': marks, 'desc': desc, 'combined_text': True,
                     'status_raw': ent_annot.get(bl, {}).get('status', ''),
                     'transit_raw': ent_annot.get(bl, {}).get('transit', ''),
                     'dg': ent_annot.get(bl, {}).get('dg', '') or find_dg(combined),
                     'reefer': ent_annot.get(bl, {}).get('reefer', '') or find_temp(combined)}
    layout_b_trusted = _trusted(len(ENT_B), len(idx2))
    for bl, rec in ENT_B.items():
        rec['low_confidence'] = not layout_b_trusted
        ENT[bl] = rec   # layout B ทับ layout A ถ้าเจอ B/L เดียวกันทั้งสองวิธี (โครงสร้างเฉพาะกว่า A)

    # เติมด้วยตัวอ่านสำรองทั่วไปสำหรับ B/L ที่ layout A/B ที่รู้จักหาไม่เจอ (ทั้งเอกสาร ไม่ใช่แค่ตอน
    # ENT ว่างเปล่าทั้งหมด) — เอกสารผสมหลาย layout ต้องพึ่ง fallback เฉพาะหน้าที่ไม่ตรงทั้ง A และ B
    for bl, rec in parse_enter_generic(doc, bl_re, ent_annot).items():
        if bl not in ENT:
            ENT[bl] = rec
    return ENT, declared


# ═══════════════════ 2b) ENTER.pdf — fallback ทั่วไปเมื่อไม่ตรงทั้ง layout A และ B ═══════════════════
_GEN_PKG_WORD = r'(?:PALLETS?|PLTS?|CARTONS?|CTNS?|CASES?|BOXES?|PACKAGES?|PKGS?|SETS?|ROLLS?|DRUMS?|BAGS?|BALES?|CRATES?|UNITS?|PIECES?)'
# (?<![A-Za-z0-9]) กันแมตช์เลขที่ฝังอยู่กลางโทเคนอื่น (เช่น ท่อนท้ายของเลข B/L "...S21260701163" เอง
# บังเอิญตามด้วยคำหีบห่อของบล็อกถัดไปพอดีเพราะต่อบรรทัดกันด้วย \n ซึ่ง \s ก็นับเป็นช่องว่าง) — ต้อง
# เป็นตัวเลขที่ "เริ่มต้นใหม่จริง ๆ" (หน้ามันไม่ใช่ตัวอักษร/ตัวเลข) เท่านั้นถึงจะนับเป็นค่าจริง
_GEN_GW_RE = re.compile(r'(?<![A-Za-z0-9])([\d][\d,]*\.?\d*)\s*KGS?\.?\b', re.I)
_GEN_MEAS_RE = re.compile(r'(?<![A-Za-z0-9])([\d][\d,]*\.?\d*)\s*(?:CBM|M3\.?|MTQ)\b', re.I)
_GEN_PKGCNT_RE = re.compile(rf'(?<![A-Za-z0-9])([\d]{{1,6}}(?:,\d{{3}})*)\s+{_GEN_PKG_WORD}', re.I)
_GEN_CONS_LBL = re.compile(r'^CONSIGNEE\b', re.I)
_GEN_STATUS_LBL = re.compile(r'^STATUS\b', re.I)
# บรรทัดป้ายกำกับ/หัวคอลัมน์/ข้อความมาตรฐานที่พบซ้ำในแบบฟอร์มต่าง ๆ (ไม่ใช่เนื้อหาสินค้าจริง) — ตัด
# ออกจาก DESCRIPTION ที่รวบรวมแบบทั่วไป กันไม่ให้ป้ายฟอร์มมาปนกับคำอธิบายสินค้าจริง
_GEN_NOISE_RE = re.compile(
    r'^(BILL OF LADING|COPY|NON.?NEGOTIABLE|SHIPPER.?S LOAD,? COUNT (?:AND|&) WEIGHT|SAID TO CONTAIN|'
    r'FREIGHT (?:PREPAID|COLLECT)|PARTICULARS FURNISHED BY MERCHANT|AS CARRIER|ALL TERMS,? CONDITIONS.*|'
    r'AND EXCEPTIONS AS PER.*|ORIGINAL BILL OF LADING.*|PAGE \d+ OF \d+|AMEND(?:MENT)?( MANIFEST)?|'
    r'WE WOULD LIKE TO REQUEST.*|B\s*/?\s*L\s*(?:NO\.?|CHANGE NO\.?)\s*:?.*|CNTR\.?\s*NO\.?\s*:?.*|'
    r'CONT\.?\s*NO\.?\s*:?.*|D\s*/\s*O\s*NO\.?.*|CONSIGNEE\s*:?.*|NOTIFY PARTY\s*:?.*|SAME AS CONSIGNEE|'
    r'STATUS\s*:?.*|VESSEL\s*:?.*|ETD\s*:?.*|ETA\s*:?.*|DATE\s*:?.*|TO\s*:?|FROM\s*:?|ATTN\s*:?.*|'
    r'IMPORT DEPT\.?|MARKS?\s*(?:AND|&)\s*NUMBERS?|MARK&NO\.?|QUANTITY(?: AND KIND)?|OF PACKAGE|'
    r'DESCRIPTION(?: OF (?:PACKAGE|GOODS))?|AND GOODS|GROSS WEIGHTS?|MEASUREMENT|M3\.?|KIND OF PACKAGES.*|'
    r'NO\.? OF (?:CONTAI\s*NERS? )?(?:OR )?PKGS.*|SEAL NO\.?|CONTAINER NO\.?|PLACE OF (?:RECEIPT|LOADING|DELIVERY)|'
    r'PORT OF (?:LOADING|DISCHARGE)|FINAL DESTINATION.*|PRE CARRIAGE BY|OCEAN VESSEL|VOY NO\.?|TOTAL|'
    r'G\.W[\d\s.,]*KGS?\.?|\*{2,}.*\*{2,}|FM-CSI.*|SHED\s*:?.*|FEEDER\s*:?.*|M B/L\s*:?.*|MASTER JOB NO\.?\s*:?.*|'
    r'COMPANY|AGENT|AGN|EXCH\. RATE\s*:?.*|TRANSIT PORT\s*:?.*)$', re.I)


def parse_enter_generic(doc, bl_re, ent_annot):
    """ตัวอ่านสำรองแบบทั่วไป (best-effort) — ใช้เมื่อ ENTER.pdf ไม่ตรงกับ layout A หรือ B ที่รู้จัก
       (พบได้บ่อยเมื่อเอกสารเป็นการรวมจดหมายแก้ไขจาก freight forwarder หลายรายที่ใช้แบบฟอร์มต่างกัน
       ในไฟล์เดียว). หาเลข B/L จากทุกบรรทัดในทุกหน้า (ไม่ใช้ anchor ตายตัว) แล้วดึงฟิลด์รอบ ๆ ด้วย
       regex ทั่วไป — แม่นยำน้อยกว่า layout A/B มาก จึงทำเครื่องหมาย generic_layout=True ให้ build_report
       ลดระดับทุกจุดที่ไม่ตรงกันเป็น "ส้ม = ต้องตรวจสอบด้วยคน" แทนที่จะฟันธง "แดง = ผิดแน่นอน"."""
    bl_search = re.compile(bl_re.pattern.strip('^$'))

    # หน้าเดียวกันบางครั้งมีเลข B/L ถูกตัดขึ้นบรรทัดใหม่กลางคัน (เช่น "HASLS2126070" / "2092" คนละ
    # บรรทัด) ทำให้ regex จับได้แค่ท่อนสั้น ๆ ที่ไม่ใช่เลข B/L จริง — เดาความยาวเลข B/L ที่ถูกต้องจาก
    # ความยาวที่พบบ่อยที่สุดในเอกสารก่อน (เลขจริงมักยาวเท่ากันแทบทั้งฉบับ ต่างกันแค่ตัวอักษรต่อท้าย
    # sub-B/L 0-2 ตัว) แล้วค่อยกรองท่อนสั้นผิดปกติทิ้งในรอบจริง กัน false-positive "มีใน ENTER แต่ไม่มี
    # ใน MANIFEST" จากขยะการตัดบรรทัด ไม่ใช่ B/L ที่ตกหล่นจริง
    page_raw = []
    all_lens = []
    for p in range(doc.page_count):
        pg = doc[p]
        raw = []
        for b in pg.get_text('dict')['blocks']:
            for l in b.get('lines', []):
                t = ''.join(s['text'] for s in l['spans']).strip()
                if t:
                    raw.append((l['bbox'][1], l['bbox'][0], t))
        raw.sort(key=lambda x: (round(x[0], 1), x[1]))
        page_raw.append(raw)
        for _, _, t in raw:
            m = bl_search.search(t)
            if m:
                all_lens.append(len(m.group(0)))
    base_len = Counter(all_lens).most_common(1)[0][0] if all_lens else 0
    valid_len = set(range(base_len, base_len + 3)) if base_len else None   # ฐาน / +1-2 ตัวอักษร sub-B/L

    ENT = {}
    for p in range(doc.page_count):
        raw = page_raw[p]
        if not raw:
            continue
        anchors = [(i, m.group(0)) for i, (_, _, t) in enumerate(raw)
                   for m in [bl_search.search(t)] if m]
        if not anchors:
            continue
        for k, (ai, bl) in enumerate(anchors):
            if not bl_re.match(bl):
                continue
            if valid_len and len(bl) not in valid_len:
                continue   # ท่อนสั้นผิดปกติ — น่าจะเป็นเลข B/L ที่ถูกตัดขึ้นบรรทัดใหม่กลางคัน ไม่ใช่ของจริง
            lo = 0 if k == 0 else (anchors[k - 1][0] + ai) // 2 + 1
            hi = (len(raw) - 1) if k == len(anchors) - 1 else (ai + anchors[k + 1][0]) // 2
            blk = [t for _, _, t in raw[lo:hi + 1]]
            blk_text = '\n'.join(blk)

            cons = ''
            ci = next((i for i, t in enumerate(blk) if _GEN_CONS_LBL.match(t.strip())), None)
            if ci is not None:
                after = re.sub(r'^CONSIGNEE\s*:?\s*', '', blk[ci].strip(), flags=re.I)
                nxt = blk[ci + 1].strip() if ci + 1 < len(blk) else ''
                prv = blk[ci - 1].strip() if ci > 0 else ''
                cons = after or nxt or prv

            status_raw = ''
            si = next((i for i, t in enumerate(blk) if _GEN_STATUS_LBL.match(t.strip())), None)
            if si is not None:
                after = re.sub(r'^STATUS\s*:?\s*', '', blk[si].strip(), flags=re.I)
                status_raw = after or (blk[si + 1].strip() if si + 1 < len(blk) else '')

            cm = CONT_RE.search(blk_text)
            cont = cm.group(0) if cm else ''
            gm = _GEN_GW_RE.search(blk_text)
            gw = num(gm.group(1)) if gm else None
            mm2 = _GEN_MEAS_RE.search(blk_text)
            meas = num(mm2.group(1)) if mm2 else None
            pm = _GEN_PKGCNT_RE.search(blk_text)
            cnt = pkgnum(pm.group(1)) if pm else None
            pkgtype = pm.group(0) if pm else ''

            desc = ' '.join(re.sub(r'\s+', ' ', t).strip() for t in blk
                             if t.strip() and not _GEN_NOISE_RE.match(t.strip())
                             and t.strip() != bl and not bl_re.match(t.strip()))

            rec = ent_annot.get(bl, {})
            prev = ENT.get(bl, {})
            ENT[bl] = {
                'bl': bl,
                'cons': prev.get('cons') or cons,
                'cont': prev.get('cont') or cont,
                'pkgs': prev.get('pkgs') if prev.get('pkgs') is not None else cnt,
                'pkgtype': prev.get('pkgtype') or pkgtype,
                'gw': prev.get('gw') if prev.get('gw') is not None else gw,
                'meas': prev.get('meas') if prev.get('meas') is not None else meas,
                'marks': prev.get('marks') or desc,
                'desc': prev.get('desc') or desc,
                'combined_text': True,
                'generic_layout': True,
                'status_raw': prev.get('status_raw') or rec.get('status', '') or status_raw,
                'transit_raw': prev.get('transit_raw') or rec.get('transit', '') or (blk_text if TRANSIT_RE.search(blk_text) else ''),
                'dg': prev.get('dg') or rec.get('dg', '') or find_dg(blk_text),
                'reefer': prev.get('reefer') or rec.get('reefer', '') or find_temp(blk_text),
            }
    return ENT


# ═══════════════════════════ 3) สร้างรายงาน EDI.xlsx ═══════════════════════════
FN, FS = 'Aptos Narrow', 9
WHITE, GREEN, RED, ORANGE, YEL, NAVY, DKRED = 'FFFFFF', 'C6EFCE', 'FDD7D7', 'FCE4D6', 'FFF2CC', '1F4E78', 'B00000'
f_ok = PatternFill('solid', start_color=WHITE, end_color=WHITE)     # ไม่ผิด = พื้นขาว
f_grn = PatternFill('solid', start_color=GREEN, end_color=GREEN)    # ช่องสรุป "ผ่าน" = เขียว
f_red = PatternFill('solid', start_color=RED, end_color=RED)
f_org = PatternFill('solid', start_color=ORANGE, end_color=ORANGE)
f_yel = PatternFill('solid', start_color=YEL, end_color=YEL)
f_hdr = PatternFill('solid', start_color=NAVY, end_color=NAVY)
F_base = Font(name=FN, size=FS)
F_hdr = Font(name=FN, size=FS, bold=True, color='FFFFFF')
F_bad = Font(name=FN, size=FS, bold=True, color=DKRED)
F_ttl = Font(name=FN, size=FS, bold=True)
thin = Side(style='thin', color='BFBFBF')
BORD = Border(left=thin, right=thin, top=thin, bottom=thin)
WARN, OK = '⚠', '✔'

HEAD = ['สรุปผลเร่งด่วน', 'B/L NO.',
        'CNEE. (ENTER)', 'CNEE. (MANIFEST)',
        'CNTRS NO. (ENTER)', 'CNTRS NO. (MANIFEST)',
        'STATUS (ENTER)', 'STATUS (MANIFEST)',
        'TOTAL PKG (ENTER)', 'PKG. (ENTER)',
        'TOTAL PKG (MANIFEST)', 'PKG. (MANIFEST)',
        'G.W. KGM (ENTER)', 'G.W. KGM (MANIFEST)',
        'MEAS. MTQ (ENTER)', 'MEAS. MTQ (MANIFEST)',
        'MARKS (ENTER)', 'MARKS (MANIFEST)',
        'DESC. (ENTER)', 'DESC. (MANIFEST)',
        'REEFER TEMP', 'DG', 'TRANSIT/TRANSHIPMENT',
        'หมายเหตุ / จุดที่ไม่ตรงกัน']
NC = len(HEAD)
NUMCOL = {13, 14, 15, 16}    # G.W./MEAS. (ENTER+MANIFEST) = ชิดขวา ; ที่เหลือชิดซ้าย
WIDE = {17, 18, 19, 20}      # MARKS / DESC. ให้กว้างอ่านได้เต็ม


def build_report(MAN, ENT, vessel, man_declared, ent_declared, outpath):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'EDI'

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=NC)
    ws.cell(1, 1, 'รายงานตรวจสอบเอกสารขาเข้าเรือ  EDI  —  เทียบ ENTER (เอกสารลูกค้า) กับ MANIFEST (สายเรือ)  |  '
                   'จุดสำคัญที่มีผลกับใบขนสินค้าขาเข้า').font = F_ttl
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=NC)
    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=NC)
    for j, h in enumerate(HEAD, 1):
        c = ws.cell(4, j, h)
        c.font = F_hdr
        c.fill = f_hdr
        c.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
    ws.freeze_panes = 'C5'   # ตรึงคอลัมน์ สรุปผลเร่งด่วน + B/L NO.
    ws.row_dimensions[4].height = 30

    r = 5
    issue_rows = 0
    crit = []
    TOT = {'pe': 0, 'pm': 0, 'ge': 0.0, 'gm': 0.0, 'me': 0.0, 'mm': 0.0,
           'pke': {}, 'pkm': {}, 'cte': {}, 'ctm': {}, 'st_bad': 0, 'transit': 0, 'dg': 0, 'reefer': 0}
    sub = sorted(MAN, key=lambda b: (len(b), b))

    for bl in sub:
        m = MAN[bl]
        e = ENT.get(bl)
        notes, urgent, cs = [], [], {}

        def mk(col, ok):
            cs[col] = 'ok' if ok else 'diff'

        if not e:
            _pm, _gm, _mm = pkgnum(m['pkg_hdr']), num(m['gw']), num(m['meas'])
            vals = [f'{WARN} ไม่มีเอกสาร ENTER เทียบ',
                    f'{WARN} {bl}', '—', m['cons'] or m['notify'], '—', ';'.join(m['cont']),
                    '—', canon_status(m['status']) or m['status'] or '(ว่าง)',
                    '—', '—', _pm if _pm is not None else (m['pkg_hdr'] or '-'), canon_pkg(m['pkg_hdr']) or '-',
                    '—', f'{_gm:,.3f}' if _gm is not None else '-',
                    '—', f'{_mm:,.3f}' if _mm is not None else '-',
                    '—', m['marks'], '—', m['desc'], '-', '-',
                    extract_transit_phrase(m['transit']) or '-',
                    'ไม่พบ B/L นี้ในเอกสาร ENTER']
            for j, v in enumerate(vals, 1):
                c = ws.cell(r, j, v)
                c.fill = f_yel
                c.border = BORD
                c.font = F_bad if j in (1, 2, 24) else F_base
                c.alignment = Alignment(wrap_text=True, vertical='top',
                                         horizontal=('right' if j in (13, 14, 15, 16) else 'left'))
            if _pm: TOT['pm'] += _pm
            if _gm: TOT['gm'] += _gm
            if _mm: TOT['mm'] += _mm
            r += 1
            issue_rows += 1
            crit.append(f'{bl}: ไม่มีเอกสาร ENTER')
            continue

        # CONSIGNEE (เทียบ ENTER กับ consignee + notify ของ MANIFEST)
        ne = norm_co(e['cons'])
        mc, mn = norm_co(m['cons']), norm_co(m['notify'])
        hit_c = bool(ne and mc and comatch(ne, mc))
        hit_n = bool(ne and mn and comatch(ne, mn))
        cons_ok = bool(hit_c or hit_n)
        mk(3, cons_ok); mk(4, cons_ok)
        if not cons_ok:
            notes.append(f"CONSIGNEE ไม่ตรง (ENTER: {e['cons']} / MANIFEST C:{m['cons']} N:{m['notify']})")
        elif hit_n and not hit_c:
            cs[3] = 'warn'
            notes.append(f"CONSIGNEE(MANIFEST) ช่อง C: = '{m['cons']}' ไม่ครบ — ชื่อเต็มตรงกับ Notify")

        # CONTAINER
        ecm = CONT_RE.search(e['cont'] or '')
        e_c = ecm.group(0) if ecm else ''
        cont_ok = bool(e_c and e_c in m['cont'])
        mk(5, cont_ok); mk(6, cont_ok)
        if not cont_ok:
            notes.append(f"CONTAINER ไม่ตรง (ENTER: {e_c or '-'} / MANIFEST: {';'.join(m['cont']) or '-'})")

        # STATUS
        se_raw, sm_raw = e['status_raw'] or '', m['status'] or ''
        se, sm = canon_status(se_raw), canon_status(sm_raw)
        st_ok = (se == sm) if (se and sm) else True
        mk(7, st_ok); mk(8, st_ok)
        disp_e = (se_raw + (f'  (= {se})' if se and se_raw.strip() != se else '')) if se_raw else '(ไม่ระบุใน ENTER)'
        disp_m = (sm_raw + (f'  (= {sm})' if sm and sm_raw != sm else '')) if sm_raw else '(ว่าง)'
        if not st_ok:
            TOT['st_bad'] += 1
            notes.append(f"STATUS ไม่ตรง (ENTER {se_raw} = {se} / MANIFEST {sm_raw} = {sm})")

        # TOTAL PACKAGE
        pe, pm = e['pkgs'], pkgnum(m['pkg_hdr'])
        pk_ok = (pe is not None and pm is not None and pe == pm)
        mk(9, pk_ok); mk(11, pk_ok)
        if not pk_ok:
            notes.append(f"TOTAL PACKAGE ไม่ตรง (ENTER {pe if pe is not None else '-'} / "
                         f"MANIFEST {pm if pm is not None else '(ไม่มี)'})")

        # PACKAGING (ชนิดบรรจุภัณฑ์)
        ke, km = canon_pkg(e['pkgtype']), canon_pkg(m['pkg_hdr'])
        pg_ok = bool(ke and km and (ke == km or ke in km or km in ke))
        mk(10, pg_ok); mk(12, pg_ok)
        if not pg_ok:
            notes.append(f"PACKAGING ไม่ตรง (ENTER {ke or e['pkgtype'] or '-'} / "
                         f"MANIFEST {km or m['pkg_hdr'] or '(ไม่มี)'})")

        # GROSS WEIGHT
        ge, gm = e['gw'], num(m['gw'])
        gw_ok = (ge is not None and gm is not None and round(ge, 2) == round(gm, 2))
        if (gm in (0, None)) and ge:
            gw_ok = False
        mk(13, gw_ok); mk(14, gw_ok)
        if not gw_ok:
            extra = ' — MANIFEST = 0.000 (ค่าหาย) แต่ ENTER มีค่าจริง' if (gm in (0, None) and ge) else ''
            notes.append(f"GROSS WEIGHT ไม่ตรง (ENTER {ge} / MANIFEST {gm}){extra}")

        # MEASUREMENT
        me_, mm = e['meas'], num(m['meas'])
        ms_ok = (me_ is not None and mm is not None and round(me_, 3) == round(mm, 3))
        if (mm in (0, None)) and me_:
            ms_ok = False
        mk(15, ms_ok); mk(16, ms_ok)
        if not ms_ok:
            extra = ' — MANIFEST = 0.000 (ค่าหาย) แต่ ENTER มีค่าจริง' if (mm in (0, None) and me_) else ''
            notes.append(f"MEASUREMENT ไม่ตรง (ENTER {me_} / MANIFEST {mm}){extra}")

        # MARKS / DESCRIPTION — เตือนทุก B/L ที่ข้อมูลไม่เหมือนกัน (เข้มงวด ไม่ตัดทอนข้อความ)
        if e.get('combined_text'):
            # ENTER เอกสารนี้แยกคอลัมน์ MARKS & NOS / DESCRIPTIONS OF GOODS ตามตำแหน่ง x0 ในหน้า PDF
            # ได้แล้ว (e['marks'] / e['desc'] จริง ไม่ใช่บล็อกเดียวกันซ้ำสองช่อง) แต่จุดตัด MARKS/DESC
            # ของ MANIFEST มาจากตำแหน่งตัดบรรทัดของ Excel ซึ่งอาจไม่ตรงกับจุดตัดของ ENTER เป๊ะ — เทียบ
            # ทั้งก้อน (marks+desc รวมกัน) ด้วย same_detail (ตัวหนังสือ/ตัวเลขล้วน ไม่สนใจช่องว่าง/
            # เครื่องหมายวรรคตอนเลย ไม่จับผิดกรณี PDF/Excel แทรกช่องว่างกลางคำผิดที่ เช่น "CO." เป็น
            # "C O.") แต่ยังต้องเตือนทุกครั้งที่เนื้อหาจริงต่างกัน (คำ/ประโยคเพิ่มหรือขาดหาย) ทุก B/L —
            # เทียบทีละคอลัมน์ (MARKS กับ MARKS, DESC กับ DESC) ไม่ใช่เหมารวมเทียบทั้งก้อนแล้วเตือนทั้ง
            # 4 ช่องพร้อมกัน เพราะถ้าคอลัมน์หนึ่งเหมือนกันเป๊ะอยู่แล้ว (เช่น MARKS ตรงกันทุกตัวอักษร)
            # แต่ดันโดนเตือนไปด้วยเพราะอีกคอลัมน์ต่าง คนตรวจจะงงว่าช่องที่เหมือนกันมันผิดตรงไหน
            if same_detail(e['marks'], m['marks']) is False:
                cs[17] = cs[18] = 'warn'
                notes.append(f"MARKS ไม่ตรงกัน (ไม่นับช่องว่าง) — (ENTER: {e['marks'] or '(ว่าง)'} / "
                             f"MANIFEST: {m['marks'] or '(ว่าง)'})")
            else:
                mk(17, True); mk(18, True)
            if same_detail(e['desc'], m['desc']) is False:
                cs[19] = cs[20] = 'warn'
                notes.append(f"DESCRIPTION ไม่ตรงกัน (ไม่นับช่องว่าง) — (ENTER: {e['desc'] or '(ว่าง)'} / "
                             f"MANIFEST: {m['desc'] or '(ว่าง)'})")
            else:
                mk(19, True); mk(20, True)
        else:
            if same_detail(e['marks'], m['marks']) is False:
                mk(17, False); mk(18, False)
                notes.append(f"MARKS ไม่เหมือนกัน (ENTER: {e['marks'] or '(ว่าง)'} / MANIFEST: {m['marks'] or '(ว่าง)'})")
            if same_detail(e['desc'], m['desc']) is False:
                mk(19, False); mk(20, False)
                notes.append(f"DESCRIPTION ไม่เหมือนกัน (ENTER: {e['desc'] or '(ว่าง)'} / MANIFEST: {m['desc'] or '(ว่าง)'})")

        # REEFER TEMP / DG (CLASS/UN) — พบที่ฝั่งใดก็ต้องเตือนเสมอ (มีผลทางกฎหมาย)
        re_e, re_m = e['reefer'] or '', m['reefer'] or ''
        dg_e, dg_m = e['dg'] or '', m['dg'] or ''
        reefer = ' | '.join(p for p in ((f'ENTER: {re_e}' if re_e else ''),
                                        (f'MANIFEST: {re_m}' if re_m else '')) if p) or '-'
        dg = ' | '.join(p for p in ((f'ENTER: {dg_e}' if dg_e else ''),
                                    (f'MANIFEST: {dg_m}' if dg_m else '')) if p) or '-'
        if re_e or re_m:
            mk(21, False)
            TOT['reefer'] += 1
            tail = 'อีกฝั่งไม่ระบุ ; ' if bool(re_e) != bool(re_m) else ''
            notes.append(f"REEFER TEMP {reefer} — {tail}ต้องสำแดงตู้เย็น/ยืนยันอุณหภูมิในใบขน")
            urgent.append(f"{WARN} REEFER {re_e or re_m}")
        else:
            mk(21, True)
        if dg_e or dg_m:
            mk(22, False)
            TOT['dg'] += 1
            tail = 'อีกฝั่งไม่ระบุ ; ' if bool(dg_e) != bool(dg_m) else ''
            notes.append(f"DG {dg} — {tail}ต้องสำแดงวัตถุอันตราย (CLASS/UN) + แนบ MSDS/DGD")
            urgent.append(f"{WARN} DG {dg_e or dg_m}")
        else:
            mk(22, True)

        # TRANSIT / TRANSHIPMENT — โชว์เฉพาะใจความ "INTRANSIT TO x" / "TRANSHIPMENT TO x" สั้น ๆ
        # (ไม่เอาป้าย ENTER:/MANIFEST: หรือรายละเอียดเส้นทางอ้อมมาปนในคอลัมน์)
        te, tm = e['transit_raw'] or '', m['transit'] or ''
        te_p, tm_p = extract_transit_phrase(te), extract_transit_phrase(tm)
        parts = [p for p in (te_p, tm_p) if p]
        transit_disp = ' / '.join(dict.fromkeys(parts)) if parts else '-'
        if parts:
            mk(23, False)
            TOT['transit'] += 1
            notes.append(f"พบคำ transit/transhipment — {transit_disp} (ไม่ใช่ LOCAL, ตรวจประเภทใบขน: ถ่ายลำ/ผ่านแดน)")
            urgent.append(f"{WARN} {transit_disp}")
        else:
            mk(23, True)

        # เอกสาร ENTER ที่ parse ด้วย fallback ทั่วไป (layout ที่ไม่รู้จัก) หรือด้วย layout A/B ที่รู้จัก
        # แต่พบ anchor น้อยเกินกว่าจะเชื่อว่าเป็น layout หลักของทั้งฉบับจริง (low_confidence) แม่นยำน้อย
        # กว่ามาก — ลดระดับทุกจุดที่ไม่ตรงกัน (ยกเว้น DG/REEFER/TRANSIT ที่ต้องเตือนเสมอเมื่อพบคำเหล่านี้
        # ไม่ว่าจะมั่นใจแค่ไหน) จาก "แดง=ผิดแน่นอน" เป็น "ส้ม=ต้องตรวจสอบด้วยคน" แทน
        if e.get('generic_layout') or e.get('low_confidence'):
            for c in list(cs):
                if cs[c] == 'diff' and c not in (21, 22, 23):
                    cs[c] = 'warn'
            notes.append('ENTER เอกสารนี้ดึงข้อมูลด้วยความมั่นใจต่ำ (layout ไม่ชัดเจน/พบ anchor น้อยเกินไป) — '
                          'ตรวจสอบทุกช่องที่ไฮไลต์ส้มด้วยคน')

        diff = [c for c, s in cs.items() if s == 'diff']
        warn_cols = [c for c, s in cs.items() if s == 'warn']
        if diff or warn_cols:
            issue_rows += 1
        LBL = {3: 'CONSIGNEE', 4: 'CONSIGNEE', 5: 'CONTAINER', 6: 'CONTAINER', 7: 'STATUS', 8: 'STATUS',
               9: 'PACKAGE', 11: 'PACKAGE', 10: 'PACKAGING', 12: 'PACKAGING', 13: 'G.W.', 14: 'G.W.',
               15: 'MEAS.', 16: 'MEAS.', 17: 'MARKS', 18: 'MARKS', 19: 'DESC', 20: 'DESC',
               21: 'REEFER', 22: 'DG', 23: 'TRANSIT'}
        if diff:
            tags = sorted({LBL[c] for c in diff if c not in (21, 22, 23)})
            if tags:
                urgent.insert(0, f"{WARN} ไม่ตรง: " + ', '.join(tags))
                crit.append(f"{bl}: ไม่ตรง " + ', '.join(tags))
        elif warn_cols:
            # ไม่มีจุดที่ฟันธงว่าผิด (แดง) แต่มีจุดที่ต้องให้คนยืนยัน (ส้ม เช่น ดึงข้อมูลด้วยความมั่นใจ
            # ต่ำ) — ห้ามให้คอลัมน์สรุปผลเร่งด่วนขึ้นว่า "ผ่าน" เฉย ๆ ไม่งั้นคนตรวจจะข้ามแถวนี้ไปทั้งที่
            # ยังมีช่องส้มให้ตรวจอยู่
            wtags = sorted({LBL[c] for c in warn_cols if c in LBL})
            urgent.insert(0, f"{WARN} ต้องตรวจสอบ: " + ', '.join(wtags) if wtags else f'{WARN} ต้องตรวจสอบด้วยคน')
        if urgent and not urgent[0].startswith(WARN):
            urgent[0] = f'{WARN} {urgent[0]}'
        urgent_txt = ' ; '.join(dict.fromkeys(urgent)) if urgent else f'{OK} ผ่าน'

        if pe is not None: TOT['pe'] += pe
        if pm is not None: TOT['pm'] += pm
        if ge is not None: TOT['ge'] += ge
        if gm is not None: TOT['gm'] += gm
        if me_ is not None: TOT['me'] += me_
        if mm is not None: TOT['mm'] += mm
        if ke: TOT['pke'][ke] = TOT['pke'].get(ke, 0) + (pe or 0)
        if km: TOT['pkm'][km] = TOT['pkm'].get(km, 0) + (pm or 0)
        if e_c: TOT['cte'][e_c] = TOT['cte'].get(e_c, 0) + 1
        for x in m['cont']: TOT['ctm'][x] = TOT['ctm'].get(x, 0) + 1

        row = [urgent_txt, bl,
               e['cons'], m['cons'] or m['notify'],
               e_c, ';'.join(m['cont']),
               disp_e, disp_m,
               pe if pe is not None else '-', ke or e['pkgtype'] or '-',
               pm if pm is not None else '(ไม่มี)', km or m['pkg_hdr'] or '(ไม่มี)',
               f'{ge:,.3f}' if ge is not None else '-', f'{gm:,.3f}' if gm is not None else '-',
               f'{me_:,.3f}' if me_ is not None else '-', f'{mm:,.3f}' if mm is not None else '-',
               e['marks'], m['marks'],
               e['desc'], m['desc'],
               reefer, dg, transit_disp,
               ' ; '.join(notes) if notes else '-']
        has_diff = bool(diff)
        has_warn = any(s == 'warn' for s in cs.values())
        for j, v in enumerate(row, 1):
            c = ws.cell(r, j, v)
            s = cs.get(j, 'ok')
            c.fill = {'ok': f_ok, 'diff': f_red, 'warn': f_org}[s]
            c.font = F_bad if s in ('diff', 'warn') else F_base
            c.alignment = Alignment(wrap_text=True, vertical='top',
                                     horizontal=('right' if j in NUMCOL else 'left'))
            c.border = BORD
            if j == 1:
                if has_diff:
                    c.font = F_bad; c.fill = f_red
                elif has_warn:
                    c.font = F_bad; c.fill = f_org
                else:
                    c.font = Font(name=FN, size=FS, bold=True, color='006100'); c.fill = f_grn
            if j == 24:
                c.font = F_bad if (has_diff or has_warn) else F_base
                c.fill = f_red if has_diff else (f_org if has_warn else f_ok)
                if v in (None, '-', ''):
                    c.fill = f_ok
                elif (has_diff or has_warn) and not str(v).startswith(WARN):
                    c.value = f'{WARN} {v}'
        for j, s in cs.items():
            if s in ('diff', 'warn'):
                cc = ws.cell(r, j)
                if cc.value not in (None, '') and not str(cc.value).startswith(WARN):
                    cc.value = f'{WARN} {cc.value}'
        r += 1

    # B/L ที่มีอยู่ใน ENTER แต่ไม่มีใน MANIFEST เลย — วิกฤต: สินค้าอาจตกหล่นจากใบขนทั้งรายการ
    # (ตรงข้ามกับกรณี "ไม่มี ENTER เทียบ" ด้านบน ซึ่งอย่างน้อยยังมีของอยู่ใน MANIFEST)
    extra_bls = sorted((b for b in ENT if b not in MAN), key=lambda b: (len(b), b))
    for bl in extra_bls:
        e = ENT[bl]
        pe, ge, me_ = e['pkgs'], e['gw'], e['meas']
        # แถวที่มาจาก fallback ทั่วไป หรือ layout A/B ที่พบ anchor น้อยเกินไปจะเชื่อได้เต็มร้อย มีโอกาส
        # เป็นขยะจากการอ่านข้อความผิด (เช่น เลข B/L ที่ตัดขึ้นบรรทัดใหม่กลางคันแล้วจับได้แค่บางส่วน)
        # มากกว่าจะเป็น B/L ที่ตกหล่นจริง — ทำเป็นส้ม "ต้องตรวจสอบด้วยคน" แทนแดง "วิกฤต" เพื่อไม่ให้
        # ตื่นตระหนกเกินจริงจากความไม่แม่นยำของตัวอ่านเอง (ต่างจาก layout A/B ที่พบ anchor มากพอ ซึ่ง
        # ความมั่นใจสูงกว่ามากจึงยังคงแดง)
        is_generic = bool(e.get('generic_layout') or e.get('low_confidence'))
        headline = (f'{WARN} พบข้อความคล้ายเลข B/L นี้ใน ENTER (ความมั่นใจต่ำ) แต่ไม่มีใน MANIFEST — '
                    f'ตรวจสอบด้วยคนว่าเป็น B/L จริงหรือขยะจากการอ่านข้อความ' if is_generic
                    else f'{WARN} มีใน ENTER แต่ไม่มีใน MANIFEST — เสี่ยงตกหล่นจากใบขน')
        note = (f'{WARN} B/L นี้ไม่พบใน MANIFEST — ดึงข้อมูลมาด้วยความมั่นใจต่ำ ตรวจสอบด้วยคนก่อนแจ้งสายเรือ/ลูกค้า'
                if is_generic else
                f'{WARN} B/L นี้ไม่มีอยู่ใน MANIFEST เลย — ของอาจตกหล่นจากใบขนทั้งรายการ ต้องแจ้งสายเรือ/ลูกค้าด่วน')
        vals = [headline,
                f'{WARN} {bl}', e['cons'], '—', e['cont'] or '', '—',
                e['status_raw'] or '(ไม่ระบุใน ENTER)', '—',
                pe if pe is not None else '-', canon_pkg(e['pkgtype']) or e['pkgtype'] or '-', '—', '—',
                f'{ge:,.3f}' if ge is not None else '-', '—',
                f'{me_:,.3f}' if me_ is not None else '-', '—',
                e['marks'], '—', e['desc'], '—',
                e['reefer'] or '-', e['dg'] or '-',
                extract_transit_phrase(e['transit_raw']) or '-',
                note]
        fill = f_org if is_generic else f_red
        for j, v in enumerate(vals, 1):
            c = ws.cell(r, j, v)
            c.fill = fill
            c.border = BORD
            c.font = F_bad
            c.alignment = Alignment(wrap_text=True, vertical='top',
                                     horizontal=('right' if j in (13, 14, 15, 16) else 'left'))
        if pe: TOT['pe'] += pe
        if ge: TOT['ge'] += ge
        if me_: TOT['me'] += me_
        r += 1
        issue_rows += 1
        crit.insert(0, f'{bl}: ' + ('⚠ อยู่ใน ENTER (layout ไม่รู้จัก) แต่ไม่มีใน MANIFEST — ตรวจสอบด้วยคน' if is_generic
                                     else '⚠⚠ อยู่ใน ENTER แต่ไม่มีใน MANIFEST เลย (วิกฤต)'))

    data_last = r - 1

    # ───────── แถวยอดรวม ─────────
    pe_ok = TOT['pe'] == TOT['pm']
    ge_ok = round(TOT['ge'], 2) == round(TOT['gm'], 2)
    me_ok = round(TOT['me'], 3) == round(TOT['mm'], 3)
    _pk = lambda dd: ' / '.join(f'{k} {v}' for k, v in sorted(dd.items())) or '-'
    _ct = lambda dd: ', '.join(sorted(dd)) or '-'          # ตู้เหมือนกันนับเป็น 1
    na = len(sub)
    _all_ok = (pe_ok and ge_ok and me_ok and not TOT['st_bad'] and not TOT['transit']
               and not TOT['dg'] and not TOT['reefer'])

    decl_bits = []
    if ent_declared:
        decl_bits.append('ENTER: ' + ' / '.join(
            f"{ent_declared[k]:,.3f}" if isinstance(ent_declared[k], float) else str(ent_declared[k])
            for k in ('pkg', 'gw', 'meas') if k in ent_declared))
    if man_declared:
        decl_bits.append('MANIFEST(GRAND TOTAL): ' + ' / '.join(
            f"{man_declared[k]:,.3f}" if isinstance(man_declared[k], float) else str(man_declared[k])
            for k in ('pkg', 'gw', 'meas') if k in man_declared))
    decl_txt = ('ยอดที่ระบุในเอกสาร — ' + '  |  '.join(decl_bits) + '  ||  ') if decl_bits else ''

    trow = [(f'{OK} ยอดรวม ENTER = MANIFEST' if _all_ok else f'{WARN} ยอดรวมไม่ตรง — ตรวจด่วน'),
            'TOTAL / ยอดรวม',
            f'{na} B/L ย่อย', f'{na} B/L ย่อย',
            f"{len(TOT['cte'])} ตู้: {_ct(TOT['cte'])}", f"{len(TOT['ctm'])} ตู้: {_ct(TOT['ctm'])}",
            f"STATUS ไม่ตรง {TOT['st_bad']} B/L", f"STATUS ไม่ตรง {TOT['st_bad']} B/L",
            TOT['pe'], _pk(TOT['pke']),
            TOT['pm'], _pk(TOT['pkm']),
            f"{TOT['ge']:,.3f}", f"{TOT['gm']:,.3f}",
            f"{TOT['me']:,.3f}", f"{TOT['mm']:,.3f}",
            '', '', '', '',
            f"{TOT['reefer']} B/L", f"{TOT['dg']} B/L", f"{TOT['transit']} B/L",
            (decl_txt + f"ตรวจยอดรวมจากรายการข้างบน: PACKAGE {'ตรง' if pe_ok else 'ไม่ตรง'}, "
                        f"GROSS WEIGHT {'ตรง' if ge_ok else 'ไม่ตรง'}, MEASUREMENT {'ตรง' if me_ok else 'ไม่ตรง'}")]
    topb = Border(left=thin, right=thin, bottom=thin, top=Side(style='medium', color=NAVY))
    for j, v in enumerate(trow, 1):
        c = ws.cell(r, j, v)
        c.border = topb
        c.alignment = Alignment(wrap_text=True, vertical='center',
                                 horizontal=('right' if j in NUMCOL else 'left'))
        if j == 2:
            c.font = Font(name=FN, size=FS, bold=True, color='FFFFFF'); c.fill = f_hdr
        else:
            c.font = Font(name=FN, size=FS, bold=True); c.fill = f_ok
    for cols, ok in ((9, pe_ok), (11, pe_ok), (13, ge_ok), (14, ge_ok), (15, me_ok), (16, me_ok)):
        if not ok:
            cc = ws.cell(r, cols)
            cc.fill = f_red; cc.font = Font(name=FN, size=FS, bold=True, color=DKRED)
            if not str(cc.value).startswith(WARN): cc.value = f'{WARN} {cc.value}'
    if TOT['st_bad']:
        for cc in (ws.cell(r, 7), ws.cell(r, 8)):
            cc.fill = f_red; cc.font = Font(name=FN, size=FS, bold=True, color=DKRED); cc.value = f'{WARN} {cc.value}'
    for cc, cnt in ((ws.cell(r, 21), TOT['reefer']), (ws.cell(r, 22), TOT['dg']), (ws.cell(r, 23), TOT['transit'])):
        if cnt:
            cc.fill = f_red; cc.font = Font(name=FN, size=FS, bold=True, color=DKRED); cc.value = f'{WARN} {cc.value}'
    ws.cell(r, 1).fill = f_grn if _all_ok else f_red
    ws.cell(r, 1).font = Font(name=FN, size=FS, bold=True, color=('006100' if _all_ok else DKRED))
    if not (pe_ok and ge_ok and me_ok):
        ws.cell(r, 24).fill = f_red
        ws.cell(r, 24).font = Font(name=FN, size=FS, bold=True, color=DKRED)
    ws.row_dimensions[r].height = 99
    r += 1

    # ───────── แถวสรุป (row 2) + legend (row 3) ─────────
    master_bl = min(sub, key=len) if sub else ''
    c2 = ws.cell(2, 1, f'Vessel {vessel or "(ไม่ระบุ)"} | Master B/L {master_bl}')
    c2.font = Font(name=FN, size=FS, bold=True, color='000000')      # แถว 2 = สีดำ
    c2.alignment = Alignment(wrap_text=True, vertical='center')
    c3 = ws.cell(3, 1, "สี: ขาว=ไม่ผิด | เขียว=ผ่าน (ช่องสรุป) | แดง ⚠=ไม่ตรง/จุดเร่งด่วน | ส้ม=ต้องยืนยัน | "
                        "เหลือง=ไม่มี ENTER เทียบ    ||    "
                        "STATUS: CY,CY/CY,FCL,ลากตู้=CY ; LCL,ขน(ส่ง)=LCL ; CFS,LCL/CFS,เปิดตู้=LCL/CFS")
    c3.font = Font(name=FN, size=FS, bold=True, color=DKRED)         # แถวที่ 3 = สีแดง
    c3.alignment = Alignment(wrap_text=True, vertical='center')
    ws.row_dimensions[2].height = 18
    ws.row_dimensions[3].height = 32

    width = [(32 if (j + 1) in WIDE else max(len(h) + 2, 9)) for j, h in enumerate(HEAD)]
    width[1] = 14   # B = B/L NO.
    for j, w in enumerate(width, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.auto_filter.ref = f'A4:{get_column_letter(NC)}{data_last}'
    for rr in range(5, data_last + 1):
        ws.row_dimensions[rr].height = 26
    wb.save(outpath)
    return {'bl': len(sub), 'issue_rows': issue_rows, 'st_bad': TOT['st_bad'],
            'transit': TOT['transit'], 'dg': TOT['dg'], 'reefer': TOT['reefer'], 'crit': crit}


# ═══════════════════════════ main ═══════════════════════════
def _find_one(patterns, folder):
    for pat in patterns:
        hits = glob.glob(os.path.join(folder, pat))
        if hits:
            return hits[0]
    return None


def main(argv=None):
    for s in (sys.stdout, sys.stderr):
        try:
            s.reconfigure(encoding='utf-8', errors='replace')
        except Exception:
            pass
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--manifest', default=None, help='พาธไฟล์ MANIFEST .xls/.xlsx (ค่าเริ่มต้น: หาให้อัตโนมัติ)')
    ap.add_argument('--enter', default=None, help='พาธไฟล์ ENTER .pdf (ค่าเริ่มต้น: หาให้อัตโนมัติ)')
    ap.add_argument('--indir', default='.', help='โฟลเดอร์ที่ใช้หาไฟล์อัตโนมัติ (ค่าเริ่มต้น: โฟลเดอร์ปัจจุบัน)')
    ap.add_argument('--outdir', default=None, help='โฟลเดอร์ผลลัพธ์ (ค่าเริ่มต้น: โฟลเดอร์เดียวกับ --indir)')
    a = ap.parse_args(argv)

    indir = a.indir
    search_dirs = [os.path.join(indir, 'input'), indir]
    manifest = a.manifest or next((_find_one(['MANIFEST*.xls', 'MANIFEST*.xlsx', '*.xls', '*.xlsx'], d)
                                    for d in search_dirs if _find_one(['MANIFEST*.xls', 'MANIFEST*.xlsx', '*.xls', '*.xlsx'], d)), None)
    enter = a.enter or next((_find_one(['ENTER*.pdf', '*.pdf'], d)
                              for d in search_dirs if _find_one(['ENTER*.pdf', '*.pdf'], d)), None)
    if not manifest or not os.path.exists(manifest):
        sys.exit(f'!! ไม่พบไฟล์ MANIFEST (.xls/.xlsx) — ระบุด้วย --manifest หรือวางไว้ในโฟลเดอร์นี้/input/')
    if not enter or not os.path.exists(enter):
        sys.exit(f'!! ไม่พบไฟล์ ENTER (.pdf) — ระบุด้วย --enter หรือวางไว้ในโฟลเดอร์นี้/input/')
    outdir = a.outdir or os.path.dirname(manifest) or '.'
    os.makedirs(outdir, exist_ok=True)
    outpath = os.path.join(outdir, 'EDI.xlsx')

    print(f'MANIFEST : {manifest}')
    print(f'ENTER    : {enter}')

    MAN, vessel, man_declared, bl_re = parse_manifest(manifest)
    ENT, ent_declared = parse_enter(enter, bl_re)

    if not MAN:
        sys.exit('!! อ่าน B/L จาก MANIFEST ไม่ได้เลย — ตรวจว่ารูปแบบไฟล์ตรงกับที่ parse_manifest() คาดไว้หรือไม่')

    stats = build_report(MAN, ENT, vessel, man_declared, ent_declared, outpath)

    print('=' * 70)
    print(f'บันทึกแล้ว: {outpath}')
    print(f'B/L ทั้งหมด {stats["bl"]} | จุดต้องตรวจ {stats["issue_rows"]} แถว | '
          f'STATUS ไม่ตรง {stats["st_bad"]} | DG {stats["dg"]} | REEFER {stats["reefer"]} | TRANSIT {stats["transit"]}')
    missing = [b for b in MAN if b not in ENT]
    if missing:
        print(f'⚠️  B/L ที่ไม่มีเอกสาร ENTER เทียบ ({len(missing)}): {", ".join(missing)}')
    extra = [b for b in ENT if b not in MAN]
    if extra:
        print(f'⚠️  B/L ที่มีใน ENTER แต่ไม่มีใน MANIFEST ({len(extra)}): {", ".join(extra)}')
    if stats['crit']:
        print('\nจุดที่ต้องตรวจสอบ:')
        for x in stats['crit']:
            print('  -', x)
    print('=' * 70)


if __name__ == '__main__':
    main()
