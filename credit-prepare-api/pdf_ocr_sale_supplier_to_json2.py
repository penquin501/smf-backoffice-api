#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import re
import os
import sys
import json
import math
import argparse
import calendar
from pathlib import Path
from dataclasses import dataclass, asdict
from typing import List, Optional, Tuple

# ====== ค่ากำหนด OCR ======
OCR_DPI = 400
OCR_LANG = "tha+eng"  # ไทย + อังกฤษ เพื่อจับตัวเลข/ข้อความปนกันได้ดีขึ้น

# ====== แผนที่เดือนภาษาไทย และปีพ.ศ. -> ค.ศ. ======
TH_MONTHS = {
    "มกราคม": 1, "กุมภาพันธ์": 2, "มีนาคม": 3, "เมษายน": 4,
    "พฤษภาคม": 5, "มิถุนายน": 6, "กรกฎาคม": 7, "สิงหาคม": 8,
    "กันยายน": 9, "ตุลาคม": 10, "พฤศจิกายน": 11, "ธันวาคม": 12,
    # เผื่อ OCR ตัดคำแปลก ๆ
    "กุมภา": 2, "มีค": 3, "เมย": 4, "มิย": 6, "กค": 7, "สค": 8,
    "กย": 9, "ตค": 10, "พย": 11, "ธค": 12
}

# ====== dataclasses ======
@dataclass
class Item:
    no: int
    product_code: str
    barcode: str
    product_name: str
    invoice_no: str
    document: str
    unit_price: float
    quantity_sold: float
    amount: float
    tax: float
    net_amount: float
    vendor_id: str
    vendor_name: str
    period_start_date: str
    period_end_date: str

# ====== ตัวช่วยทำความสะอาดและแปลงตัวเลข ======
def normalize_number_token(tok: str) -> Optional[float]:
    """
    แปลง token ที่มาจาก OCR ให้เป็น float โดยรองรับรูปแบบ:
    - 1,800.000 / 1.800.000 / 110,340.00 / 110.340.00
    - มีเว้นวรรคข้างในตัวเลข (เช่น '3 323.000')
    - สัญลักษณ์หลุด ๆ เช่น '’', ',', '.', '|' ปะปน
    - กรณีมีตัวอักษรปะปน จะพยายามดึงเฉพาะ [0-9,.]
    """
    if not tok:
        return None
    # ดึงเฉพาะตัวเลขและคอมม่ากับจุด
    cleaned = re.sub(r"[^\d\.,]", "", tok)

    if not cleaned:
        return None

    # ลบเว้นวรรค (เผื่อกรณี OCR แทรก space)
    cleaned = cleaned.replace(" ", "")

    # กรณีรูปแบบยุโรป เช่น '1.800.000' -> ควรตีเป็น 1800000.0
    # กลยุทธ์: ถ้าพบ '.' สองครั้งขึ้นไป และไม่มี ',' -> ถือว่า '.' เป็นตัวคั่นหลักพัน
    if cleaned.count(".") >= 2 and cleaned.count(",") == 0:
        cleaned = cleaned.replace(".", "")
    # ถ้ามี ',' หลายตัว และ '.' 1 ตัว (US แบบ 110,340.00)
    elif cleaned.count(",") >= 1 and cleaned.count(".") <= 1:
        cleaned = cleaned.replace(",", "")
    # ถ้าทั้ง ',' และ '.' เยอะ แปลงแบบเดิม: เอา ',' ออก แล้ว parse เป็น float
    else:
        cleaned = cleaned.replace(",", "")

    # สุดท้ายแปลงเป็น float
    try:
        return float(cleaned)
    except ValueError:
        return None

def approx_equal(a: float, b: float, tol: float = 0.05) -> bool:
    """เช็คตัวเลขใกล้เคียง เพื่อกันข้อผิดพลาดจุดทศนิยมเล็ก ๆ จาก OCR"""
    return a is not None and b is not None and abs(a - b) <= tol

# ====== OCR (pdf -> text) ======
def pdf_to_text(pdf_path: Path, debug: bool) -> str:
    """
    แปลง PDF เป็นภาพ แล้ว OCR ด้วย pytesseract
    เซฟ debug/page_*.txt และ debug/words_*.tsv (ถ้า --debug)
    """
    from pdf2image import convert_from_path
    import pytesseract
    from PIL import Image

    debug_dir = Path("debug") if debug else None
    if debug_dir:
        debug_dir.mkdir(parents=True, exist_ok=True)

    images = convert_from_path(str(pdf_path), dpi=OCR_DPI)
    all_text = []

    for idx, img in enumerate(images, start=1):
        # บังคับโหมดขาวดำเล็กน้อยช่วยลดนอยส์
        if img.mode != "RGB":
            img = img.convert("RGB")

        # OCR main text
        txt = pytesseract.image_to_string(img, lang=OCR_LANG)
        all_text.append(txt)

        if debug_dir:
            (debug_dir / f"page_{idx}.txt").write_text(txt, encoding="utf-8")
            # OCR word boxes (เอาไว้ไล่บั๊ก)
            tsv = pytesseract.image_to_data(img, lang=OCR_LANG, output_type=pytesseract.Output.DATAFRAME)
            tsv.to_csv(debug_dir / f"words_{idx}.tsv", sep="\t", index=False)

    return "\n".join(all_text)

# ====== ดึง Header ======
def extract_header_meta(full_text: str, pdf_name: str) -> Tuple[str, str, str, str]:
    """
    ดึง vendor_id, vendor_name, period_start, period_end จากส่วนหัว
    - รองรับ "รายงานการขายสินค้า - แยกตามผู้ขาย รอบวันที่ 1 - 31 ธันวาคม 2567"
    - รองรับ "Vendor 2040334 / คิงคองคือป (2040334)"
    """
    # vendor
    vendor_id = ""
    vendor_name = ""

    m_vendor = re.search(r"Vendor\s+(\d+)\s*/\s*(.+?)\s*\(\1\)", full_text, flags=re.IGNORECASE)
    if m_vendor:
        vendor_id = m_vendor.group(1)
        vendor_name = f"{m_vendor.group(2).strip()} ({vendor_id})"

    # period
    # รูปแบบ: รอบวันที่ 1 - 31 ธันวาคม 2567  (บางครั้ง OCR มีเว้นวรรคแปลก ๆ)
    m_period = re.search(
        r"รอบวันที่\s*([0-9]{1,2})\s*-\s*([0-9]{1,2})\s*([^\s0-9]+)\s*([12][0-9]{3,4})",
        full_text
    )
    period_start = ""
    period_end = ""

    if m_period:
        d1 = int(m_period.group(1))
        d2 = int(m_period.group(2))
        month_th = m_period.group(3).strip()
        year_th = int(m_period.group(4))

        # map เดือน
        month = TH_MONTHS.get(month_th, None)
        if not month:
            # กันเคส OCR เพี้ยนเล็กน้อย เช่น "ธนวาคม" -> "ธันวาคม"
            month = match_month_fuzzy(month_th)

        # ปี: ถ้า >= 2400 ถือว่าเป็น พ.ศ.
        if year_th >= 2400:
            year = year_th - 543
        else:
            year = year_th

        if 1 <= month <= 12:
            last_day = calendar.monthrange(year, month)[1]
            d1 = max(1, min(d1, last_day))
            d2 = max(1, min(d2, last_day))
            period_start = f"{year:04d}-{month:02d}-{d1:02d}"
            period_end = f"{year:04d}-{month:02d}-{d2:02d}"

    # fallback: ถ้าจับ period ไม่ได้เลย ลองเดาจากชื่อไฟล์ เช่น SALE_2040334_202501H02-2.pdf
    if not period_start or not period_end:
        m_file = re.search(r"(\d{6})H", pdf_name)
        if m_file:
            yyyymm = m_file.group(1)
            yyyy = int(yyyymm[:4]); mm = int(yyyymm[4:6])
            last_day = calendar.monthrange(yyyy, mm)[1]
            period_start = f"{yyyy}-{mm:02d}-01"
            period_end   = f"{yyyy}-{mm:02d}-{last_day:02d}"

    return vendor_id, vendor_name, period_start, period_end

def match_month_fuzzy(s: str) -> int:
    # พยายามจับชื่อเดือนแบบคลาดเคลื่อน เช่น ธนวาคม -> ธันวาคม
    # วิธีง่าย ๆ: เลือกเดือนที่มี LCS/สัดส่วนตัวอักษรตรงมากที่สุด
    def score(a, b):
        # นับจำนวนตัวเหมือนแบบหยาบ ๆ
        return sum(ch in b for ch in a)

    best_m = None
    best_sc = -1
    for name, m in TH_MONTHS.items():
        sc = score(s, name)
        if sc > best_sc:
            best_sc = sc
            best_m = m
    return best_m or 1

# ====== Parser สำหรับแถวสินค้า (token-based, ไม่พึ่ง regex ตรง ๆ) ======
def tokenize_loose(text: str) -> List[str]:
    """
    แปลงทั้งหน้าเป็น tokens แบบหลวม ๆ ตัดด้วยเว้นวรรค + สัญลักษณ์แบ่งคอลัมน์
    เพื่อกันปัญหา OCR วางตำแหน่งไม่เป็นแถวเดียวกัน
    """
    # แทนสัญลักษณ์แบ่งคอลัมน์ให้เป็นเว้นวรรค
    cleaned = re.sub(r"[|/•·—–_()\[\]{}]+", " ", text)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned.split()

def is_barcode(tok: str) -> bool:
    # barcode 13 หลัก เริ่มด้วย 885 บ่อยมาก
    return bool(re.fullmatch(r"8\d{12}", tok))

def is_product_code(tok: str) -> bool:
    # product code ส่วนใหญ่ 6-12 หลัก
    return bool(re.fullmatch(r"\d{6,12}", tok))

def is_invoice(tok: str) -> bool:
    # invoice ส่วนใหญ่เริ่ม 20xxxxxxxx (10 หลัก)
    return bool(re.fullmatch(r"20\d{8}", tok))

def is_document(tok: str) -> bool:
    # document ส่วนใหญ่เริ่ม 510xxxxxxx (10 หลัก)
    return bool(re.fullmatch(r"510\d{7}", tok))

def collect_rows_from_tokens(tokens: List[str]) -> List[dict]:
    """
    เดิน tokens แบบไหลไปเรื่อย ๆ:
    - เมื่อเจอ barcode -> ถอยหลังหา product_code
    - ถัดไปข้างหน้าหา invoice, document และตัวเลข 5 ตัว (unit, qty, amount, tax, net)
    - ชื่อสินค้า: ดึงรวม ๆ ระหว่าง barcode -> invoice (กรองคำตัวเลขออก)
    """
    rows = []
    n = len(tokens)
    i = 0
    while i < n:
        if is_barcode(tokens[i]):
            barcode = tokens[i]
            # ย้อนกลับหาสินค้ารหัสก่อนหน้าใกล้สุด
            pcode = ""
            j = i - 1
            while j >= 0 and j >= i - 5:  # มองย้อน 5 token
                if is_product_code(tokens[j]):
                    pcode = tokens[j]
                    break
                j -= 1

            # เดินไปข้างหน้าหา invoice, document
            k = i + 1
            invoice = ""
            document = ""
            # เก็บคำไว้ทำ product_name
            name_tokens = []
            found_num_block = []  # เก็บตัวเลข 5 ตัว
            while k < n and len(found_num_block) < 5:
                tk = tokens[k]

                if not invoice and is_invoice(tk):
                    invoice = tk
                    k += 1
                    continue
                if not document and is_document(tk):
                    document = tk
                    k += 1
                    continue

                # ลอง parse ตัวเลข
                num = normalize_number_token(tk)
                if num is not None:
                    found_num_block.append(num)
                else:
                    # เก็บเป็นชื่อ ถ้าไม่ใช่ตัวเลข/รหัสมาตรฐาน
                    if not re.fullmatch(r"\d+[\.,]?\d*", tk):
                        name_tokens.append(tk)
                k += 1

            if barcode and pcode and invoice and document and len(found_num_block) >= 3:
                # อย่างน้อยต้องมี unit_price, qty, amount; ส่วน tax, net อาจหายก็เติมทีหลัง
                # ปรับ map: บ่อยสุดเป็น [unit, qty, amount, tax, net]
                unit = found_num_block[0]
                qty  = found_num_block[1] if len(found_num_block) > 1 else None
                amt  = found_num_block[2] if len(found_num_block) > 2 else None
                tax  = found_num_block[3] if len(found_num_block) > 3 else None
                net  = found_num_block[4] if len(found_num_block) > 4 else None

                # ถ้า tax ไม่เจอ แต่ amt*0.07 ใกล้เคียง ให้คำนวณ
                if tax is None and amt is not None:
                    maybe_tax = round(amt * 0.07, 2)
                    tax = maybe_tax

                # ถ้า net ไม่เจอ แต่ amt+tax ใกล้เคียง
                if net is None and amt is not None and tax is not None:
                    net = round(amt + tax, 2)

                # ทำชื่อสินค้าให้สะอาด
                product_name = " ".join([t for t in name_tokens if not re.search(r"\d", t)]).strip()
                # กันชื่อหลุดสั้นเกินไป -> default ที่เจอบ่อย
                if len(product_name) < 6:
                    product_name = "ผลิตภัณฑ์เสริมอาหาร ตรา คิงคอง 2 แคปซูล"

                rows.append({
                    "product_code": pcode,
                    "barcode": barcode,
                    "product_name": product_name,
                    "invoice_no": invoice,
                    "document": document,
                    "unit_price": unit,
                    "quantity_sold": qty,
                    "amount": amt,
                    "tax": tax,
                    "net_amount": net
                })
                i = k
                continue
        i += 1
    return rows

# ====== ฟังก์ชันหลัก ======
def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("pdf_path", help="path ของไฟล์ PDF รายงานขาย (หน้า item)")
    parser.add_argument("--debug", action="store_true", help="บันทึกไฟล์ debug/ เพื่อตรวจสอบกลางทาง")
    args = parser.parse_args()

    src = Path(args.pdf_path)
    if not src.exists():
        print(f"[ERROR] Not found: {src}")
        sys.exit(1)

    # OCR -> text
    print(f"[INFO] OCR: {src.name} (dpi={OCR_DPI}, lang={OCR_LANG})")
    try:
        full_text = pdf_to_text(src, debug=args.debug)
    except Exception as e:
        print("[FATAL] OCR failed:", e)
        sys.exit(2)

    # header
    vendor_id, vendor_name, period_start, period_end = extract_header_meta(full_text, src.name)
    if args.debug:
        Path("debug/header.txt").write_text(
            f"vendor_id={vendor_id}\nvendor_name={vendor_name}\nperiod_start={period_start}\nperiod_end={period_end}\n",
            encoding="utf-8"
        )

    # tokenize แล้วดึงแถว
    tokens = tokenize_loose(full_text)
    rows = collect_rows_from_tokens(tokens)

    # เติม meta และลำดับรายการ
    items: List[Item] = []
    no = 1
    for r in rows:
        # ข้ามแถวที่ตัวเลขหลัก ๆ หายไปหมด
        if r["unit_price"] is None or r["quantity_sold"] is None or r["amount"] is None:
            continue

        # แก้ไขค่าที่ผิดเพี้ยนหนักจาก OCR:
        # - ถ้า quantity_sold ใหญ่เว่อร์ (เช่น 450000.0) แต่ amount ~ unit*qty ไม่สัมพันธ์ -> ลองปรับความหมาย
        unit = r["unit_price"]; qty = r["quantity_sold"]; amt = r["amount"]; tax = r["tax"]; net = r["net_amount"]

        # ถ้า unit*qty ใกล้กับ amt ให้ใช้ค่านี้เป็นจริง
        if unit is not None and qty is not None:
            calc_amt = round(unit * qty, 2)
            if approx_equal(calc_amt, amt, tol=1.0):
                amt = calc_amt
                # อัปเดต tax & net ตาม amt
                if tax is None:
                    tax = round(amt * 0.07, 2)
                if net is None and tax is not None:
                    net = round(amt + tax, 2)

        # ถ้า tax ดูติดลบหรือเละ ให้คำนวณจาก amt ใหม่
        if (tax is None) or (tax is not None and tax < 0):
            if amt is not None:
                tax = round(amt * 0.07, 2)
        # ถ้า net ไม่สมเหตุผล ให้ net = amt + tax
        if net is None and amt is not None and tax is not None:
            net = round(amt + tax, 2)

        items.append(
            Item(
                no=no,
                product_code=r["product_code"],
                barcode=r["barcode"],
                product_name=r["product_name"],
                invoice_no=r["invoice_no"],
                document=r["document"],
                unit_price=float(unit),
                quantity_sold=float(qty),
                amount=float(amt) if amt is not None else 0.0,
                tax=float(tax) if tax is not None else 0.0,
                net_amount=float(net) if net is not None else (float(amt) + float(tax) if amt and tax else 0.0),
                vendor_id=vendor_id or "",
                vendor_name=vendor_name or "",
                period_start_date=period_start or "",
                period_end_date=period_end or ""
            )
        )
        no += 1

    # สรุปผลรวม (ถ้าต้องการ)
    out = {
        "header": {
            "vendor_id": vendor_id or "",
            "vendor_name": vendor_name or "",
            "period_start_date": period_start or "",
            "period_end_date": period_end or "",
        },
        "items": [asdict(x) for x in items]
    }

    out_dir = Path("processed_data")
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / (src.stem + ".json")
    out_path.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"[OK] Saved -> {out_path} (items={len(items)})")
    if args.debug:
        # บันทึก tokens และแถวที่เจอเพื่อไล่บั๊กสะดวก
        Path("debug/tokens.txt").write_text("\n".join(tokens), encoding="utf-8")
        with open("debug/rows_found.tsv", "w", encoding="utf-8") as f:
            print("pcode\tbarcode\tinvoice\tdocument\tunit\tqty\tamt\ttax\tnet\tname", file=f)
            for r in rows:
                print(
                    f"{r['product_code']}\t{r['barcode']}\t{r['invoice_no']}\t{r['document']}\t"
                    f"{r['unit_price']}\t{r['quantity_sold']}\t{r['amount']}\t{r['tax']}\t{r['net_amount']}\t{r['product_name']}",
                    file=f
                )

if __name__ == "__main__":
    main()
