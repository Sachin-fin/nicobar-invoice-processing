import os
import sys
import re
import shutil
import datetime
import pymupdf
import openpyxl

from config import EXPENSES_DIR, REGISTER_PATH, VENDOR_CODES_PATH

def clean_name(s):
    return re.sub(r'[\/\\*?:"<>|]', '-', str(s)).strip()

def parse_date(txt):
    m = re.search(r'(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})', txt)
    if m:
        d, mth, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if y < 100: y += 2000
        if 1 <= mth <= 12 and 1 <= d <= 31:
            return datetime.date(y, mth, d)
    return None

def main():
    temp_dir = EXPENSES_DIR / "_temp_gmail"
    if not temp_dir.exists():
        print("No _temp_gmail folder found.")
        return

    # Load known vendors and GSTINs
    known_gstins = {}
    if REGISTER_PATH.exists():
        try:
            wb = openpyxl.load_workbook(REGISTER_PATH, data_only=True, read_only=True)
            ws = wb["Invoice Register"]
            for r in ws.iter_rows(min_row=5, values_only=True):
                if r and len(r) > 7:
                    gstin = r[7]
                    sup = r[6]
                    if gstin and sup:
                        known_gstins[str(gstin).strip().upper()] = str(sup).strip()
        except Exception as e:
            print(f"Warning reading register: {e}")

    filed_count = 0
    pdf_files = [f for f in os.listdir(temp_dir) if f.lower().endswith(('.pdf', '.png', '.jpg', '.jpeg'))]
    print(f"Organizing {len(pdf_files)} downloaded attachments...")

    for f in pdf_files:
        p = temp_dir / f
        try:
            doc = pymupdf.open(p)
            txt = '\n'.join(page.get_text() for page in doc)
            
            # 1. Determine Vendor Name
            gstins = re.findall(r'\b([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1})\b', txt)
            sup = None
            for g in gstins:
                if 'AAFCN6567R' not in g and g in known_gstins:
                    sup = known_gstins[g]
                    break
            
            if not sup:
                u_txt = txt.upper()
                if 'XTO10X' in u_txt:
                    sup = 'XTO10X Technologies Private Limited'
                elif "NIVA'S DESIGNS" in u_txt or 'NIVAS DESIGNS' in u_txt or 'MACRAMEDESIGNS' in u_txt:
                    sup = "Niva's Designs"
                elif 'MAHINDRA' in u_txt and 'FINANCIAL' in u_txt:
                    sup = 'Mahindra & Mahindra Financial Services Limited'
                elif 'IPINFO' in u_txt:
                    sup = 'IPinfo Inc'
                elif 'A & S CREATION' in u_txt or '07ABNFA6509L' in u_txt:
                    sup = 'A & S Creations'
                elif 'CUSHMAN' in u_txt:
                    sup = 'Cushman & Wakefield India Private Limited'
                elif 'AMBIENCE FACILITIES' in u_txt:
                    sup = 'Ambience Facilities Management Private Limited'
                elif 'AMBIENCE DEVELOPERS' in u_txt:
                    sup = 'Ambience Developers & Infrastructure Pvt Ltd'
                elif 'PREZOTECH' in u_txt:
                    sup = 'Prezotech Solutions Pvt Ltd'
                elif 'CLAY CRAFT' in u_txt:
                    sup = 'Clay Craft India Limited'
                elif 'STITCH' in u_txt:
                    sup = 'Stitch-9'
                elif 'MKR APP' in u_txt:
                    sup = 'MKR APP Private Limited'
                else:
                    sup = 'General Expenses'

            # 2. Extract Invoice No
            inv_no = None
            m_inv = re.search(r'(?:Invoice\s*No\.?|Invoice\s*number|Document\s*No\.?|Bill\s*No\.?)[.:\s]+([A-Za-z0-9\-_/]+)', txt, re.IGNORECASE)
            if m_inv:
                inv_no = m_inv.group(1).strip()
            if not inv_no:
                # Use clean prefix of original filename
                clean_f = re.sub(r'\s*\d{2}\.\d{2}\.\d{2,4}.*$', '', f)
                inv_no = clean_f if clean_f else 'INV'

            # 3. Extract Invoice Date
            dt = parse_date(txt)
            if not dt:
                dt_m = re.search(r'(\d{2}\.\d{2}\.\d{2,4})', f)
                if dt_m:
                    dt = parse_date(dt_m.group(1))
            if not dt:
                dt = datetime.date.today()
            dt_str = dt.strftime('%d.%m.%y')

            # 4. Construct Final Path: <EXPENSES_DIR>\<Vendor Name>\<Invoice No> <dd.mm.yy>.pdf
            inv_no_clean = clean_name(inv_no).replace('/', '-')
            sup_clean = clean_name(sup)
            vendor_dir = EXPENSES_DIR / sup_clean
            vendor_dir.mkdir(parents=True, exist_ok=True)

            target_name = f"{inv_no_clean} {dt_str}.pdf"
            target_path = vendor_dir / target_name

            if target_path.exists():
                # Disambiguate if same name exists
                target_name = f"{inv_no_clean}_{int(datetime.datetime.now().timestamp())%1000} {dt_str}.pdf"
                target_path = vendor_dir / target_name

            doc.close()
            shutil.move(p, target_path)
            filed_count += 1
            print(f"Filed: [{sup_clean}] -> {target_name}")

        except Exception as e:
            print(f"Error filing {f}: {e}")

    # Remove empty temp folder
    try:
        if not os.listdir(temp_dir):
            os.rmdir(temp_dir)
    except Exception:
        pass

    print(f"\nCompleted: {filed_count} invoices filed into their Vendor Name folders with '<invoice no> <dd.mm.yy>.pdf' format.")

if __name__ == "__main__":
    main()
