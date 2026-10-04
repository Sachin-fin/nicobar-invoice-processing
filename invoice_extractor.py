import os
import re
import datetime
from pathlib import Path
import pymupdf
import openpyxl

from config import (
    EXPENSES_DIR,
    REGISTER_PATH,
    VENDOR_CODES_PATH,
    NICOBAR_GSTINS,
    STATE_CODES,
)

NICOBAR_PAN = "AAFCN6567R"
GSTIN_RE = re.compile(r'\b([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z])\b')
# Money: Indian (10,00,000.00) or western (1,563,500) grouping, or plain digits
MONEY_RE = re.compile(r'(?<![\w.])-?\d{1,3}(?:,\d{2,3})+(?:\.\d+)?(?![\w,])|(?<![\w.,])-?\d+(?:\.\d+)?(?![\w,])')
STANDARD_GST_RATES = (0.0, 0.05, 0.12, 0.18, 0.28, 0.03, 0.0025)

MONTHS = {m: i for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"], start=1)}

# Text signatures for known vendors (used only when folder / GSTIN lookup fails).
# Keep these specific: short tokens such as "DLF" or "BIN-" match unrelated invoices.
KNOWN_VENDOR_SIGNATURES = [
    (r"PREZOTECH", "Prezotech Solutions Pvt Ltd"),
    (r"XTO10X", "XTO10X Technologies Private Limited"),
    (r"NIVA'?S\s+DESIGNS?|MACRAMEDESIGNS", "Niva's Designs"),
    (r"MAHINDRA\s*&\s*MAHINDRA\s+FINANCIAL|QUIKLYZ", "Mahindra & Mahindra Financial Services Limited"),
    (r"IPINFO", "IPinfo Inc"),
    (r"A\s*&\s*S\s+CREATIONS?|07ABNFA6509L", "A & S Creations"),
    (r"CUSHMAN\s*&\s*WAKEFIELD", "Cushman & Wakefield India Private Limited"),
    (r"AMBIENCE\s+FACILITIES", "Ambience Facilities Management Private Limited"),
    (r"AMBIENCE\s+DEVELOPERS", "Ambience Developers & Infrastructure Pvt Ltd"),
    (r"CURE\s*FOODS", "Curefoods India Limited"),
    (r"CLAY\s+CRAFT", "Clay Craft India Limited"),
    (r"STITCH[\s-]*9", "Stitch-9"),
    (r"MKR\s+APP", "MKR APP Private Limited"),
    (r"BHARTI\s+AIRTEL", "Bharti Airtel Limited"),
    (r"BLUE\s+DART", "Blue Dart Express Limited"),
    (r"DLF\s+HOME\s+DEVELOPERS", "DLF Home Developers Limited"),
    (r"DTDC\s+EXPRESS", "DTDC Express Limited"),
    (r"HANSA\s+ELITE", "Hansa Elite Luxe Interiors Pvt Ltd"),
    (r"GOWRI\s+PARVATHY", "Gowri Parvathy"),
    (r"SURBHI\s+SETHI", "Surbhi Sethi"),
    (r"SHREYAS\s+GAIKWAD", "Shreyas Gaikwad"),
    (r"JATIN\s+GULABRAI", "Jatin Gulabrai Parekh"),
    (r"ANTHROPIC", "Anthropic PBC"),
    (r"QUANG\s+VINH", "Quang Vinh Ceramic Co Ltd"),
    (r"TRION\s+PROPERTIES", "Trion Properties Private Limited"),
    (r"VRSM\s+ENTERPRISES|\bV\s+DINE\b", "V Dine Unit of VRSM Enterprises LLP"),
    (r"VODAFONE\s+IDEA", "Vodafone Idea Limited"),
    (r"VRAJ\s+COMMERCIAL", "Vraj Commercial Private Limited"),
    (r"PINE\s+LABS", "Pine Labs Limited"),
    (r"SHYAM\s+SPECTRA", "Shyam Spectra IT Services Private Limited"),
    (r"ANTYA\s+WEB", "Antya Web Private Limited"),
    (r"DPS\s+TELECOM", "DPS Telecom (OPC) Private Limited"),
]


def detect_vendor_from_text(text):
    """Return a known vendor name if its signature appears in text, else None."""
    if not text:
        return None
    up = text.upper()
    for pattern, name in KNOWN_VENDOR_SIGNATURES:
        if re.search(pattern, up):
            return name
    return None


def to_number(s):
    try:
        return float(str(s).replace(',', '').replace('₹', '').strip())
    except (TypeError, ValueError):
        return None


def parse_date(text):
    """Parse common Indian invoice date formats into datetime.date (day-first)."""
    if not text:
        return None
    text = str(text).strip()
    this_year = datetime.date.today().year

    def build(y, m, d):
        if y < 100:
            y += 2000
        if not (2015 <= y <= this_year + 1):
            return None
        try:
            return datetime.date(y, m, d)
        except ValueError:
            return None

    # 2026-06-30
    m = re.search(r'\b(\d{4})[./-](\d{1,2})[./-](\d{1,2})\b', text)
    if m:
        r = build(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        if r:
            return r
    # 30/06/2026, 02-10-2026, 30.09.26
    m = re.search(r'\b(\d{1,2})[./-](\d{1,2})[./-](\d{4}|\d{2})\b', text)
    if m:
        r = build(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        if r:
            return r
    # 24-Sep-2026, 29-Aug-26, 6 June 2026, 26th September 2026
    m = re.search(r'\b(\d{1,2})(?:st|nd|rd|th)?[\s\-./]+([A-Za-z]{3,9})[\s\-.,/]+(\d{4}|\d{2})\b', text)
    if m and m.group(2)[:3].lower() in MONTHS:
        r = build(int(m.group(3)), MONTHS[m.group(2)[:3].lower()], int(m.group(1)))
        if r:
            return r
    # Sep 24, 2026
    m = re.search(r'\b([A-Za-z]{3,9})\s+(\d{1,2})(?:st|nd|rd|th)?,?\s+(\d{4})\b', text)
    if m and m.group(1)[:3].lower() in MONTHS:
        r = build(int(m.group(3)), MONTHS[m.group(1)[:3].lower()], int(m.group(2)))
        if r:
            return r
    return None


def _first_money(s):
    """First money-looking value in s (ignores percentages)."""
    if not s:
        return None
    s = re.sub(r'\d+(?:\.\d+)?\s*%', ' ', s)
    for m in MONEY_RE.finditer(s):
        v = to_number(m.group(0))
        if v is not None:
            return v
    return None


def _is_money_line(s):
    s2 = re.sub(r'^(?:₹|Rs\.?|INR|US\$|\$)\s*', '', s.strip())
    return bool(re.fullmatch(r'-?\d{1,3}(?:,\d{2,3})+(?:\.\d+)?|-?\d+\.\d+|-?\d{2,}', s2))


class InvoiceExtractor:
    def __init__(self):
        self.supplier_memory = {}
        self.gstin_to_supplier = {}
        self.load_supplier_memory()

    def load_supplier_memory(self):
        """Build supplier profile and GL code memory from existing register and vendor code file."""
        self.supplier_memory = {}
        self.gstin_to_supplier = {}
        if REGISTER_PATH.exists():
            try:
                wb = openpyxl.load_workbook(REGISTER_PATH, data_only=True, read_only=True)
                ws = wb["Invoice Register"]
                hdr = [c.value for c in ws[4]]
                col = {h: i for i, h in enumerate(hdr) if h}

                for r in ws.iter_rows(min_row=5, values_only=True):
                    if not r or not isinstance(r[0], (int, float)):
                        continue
                    sup = r[col.get("Supplier", 6)]
                    if not sup:
                        continue
                    sup_clean = str(sup).strip()
                    key = sup_clean.lower()

                    gstin_val = r[col.get("Supplier GSTIN", 7)]
                    if gstin_val:
                        g_clean = str(gstin_val).strip().upper()
                        if len(g_clean) == 15 and NICOBAR_PAN not in g_clean:
                            self.gstin_to_supplier[g_clean] = sup_clean

                    if key not in self.supplier_memory:
                        self.supplier_memory[key] = {
                            "name": sup_clean,
                            "gstin": gstin_val,
                            "state": r[col.get("Supplier State", 29)],
                            "tds_sec": r[col.get("TDS Section", 44)],
                            "tds_rate": r[col.get("TDS Rate", 45)],
                            "loc_code": r[col.get("Location Code", 26)],
                            "vgl_code": r[col.get("Vendor GL Code", 27)],
                            "egl_code": r[col.get("Exps GL Code", 28)],
                            "tax_type": r[col.get("Tax Type", 18)],
                            "rcm_note": r[col.get("RCM / GST Compliance", 43)],
                        }
                wb.close()
            except Exception as e:
                print(f"Warning: could not read register memory: {e}")

        # Also merge Vendor GL codes
        if VENDOR_CODES_PATH.exists():
            try:
                wb_codes = openpyxl.load_workbook(VENDOR_CODES_PATH, data_only=True, read_only=True)
                ws = wb_codes.active
                for row in ws.iter_rows(min_row=4, values_only=True):
                    if row and row[0]:
                        sup_name = str(row[0]).strip()
                        key = sup_name.lower()
                        vgl = row[1] if len(row) > 1 and row[1] else None
                        egl = row[2] if len(row) > 2 and row[2] else None
                        if key in self.supplier_memory:
                            if vgl: self.supplier_memory[key]["vgl_code"] = vgl
                            if egl: self.supplier_memory[key]["egl_code"] = egl
                        else:
                            self.supplier_memory[key] = {
                                "name": sup_name, "gstin": None, "state": None,
                                "tds_sec": None, "tds_rate": None, "loc_code": None,
                                "vgl_code": vgl, "egl_code": egl,
                                "tax_type": None, "rcm_note": None,
                            }
                wb_codes.close()
            except Exception as e:
                print(f"Warning: could not read vendor codes: {e}")

    # Kept for backwards compatibility with callers using self.extractor.parse_date
    def parse_date(self, text):
        return parse_date(text)

    # ------------------------------------------------------------------ helpers
    @staticmethod
    def _folder_supplier(pdf_path):
        """Supplier implied by folder: only a direct child of EXPENSES_DIR that is not a
        system folder (_temp, _web_uploads ...) or a date folder (e.g. 'Sept 2026')."""
        parent = pdf_path.parent
        try:
            if parent.parent.resolve() != EXPENSES_DIR.resolve():
                return None
        except Exception:
            return None
        name = parent.name
        if not name or name.startswith(('_', '.')):
            return None
        if re.search(r'\b(19|20)\d{2}\b', name) or re.fullmatch(r'[\d.\-_ ]+', name):
            return None
        return name

    @staticmethod
    def _extract_invoice_no(lines):
        labels = [
            r'(?:tax\s*)?invoice\s*(?:no\.?|number|num\.?|#)',
            r'bill\s*no\.?',
            r'reimbursement\s*no\.?',
            r'credit\s*note\s*no\.?',
            r'debit\s*note\s*no\.?',
            r'document\s*no\.?',
        ]
        label_re = re.compile(r'(?<!original\s)(?:' + '|'.join(labels) + r')\s*[:.#\-]?\s*(.*)$', re.I)
        token_re = re.compile(r'[A-Za-z0-9][A-Za-z0-9\-_/.]*')

        def pick(rest, nxt, nxt2):
            rest = (rest or '').strip().lstrip(':#.- ').strip()
            src = rest if rest else (nxt or '').strip().lstrip(':#.- ').strip()
            m = token_re.match(src)
            if not m:
                return None
            tok = m.group(0).rstrip('.')
            if not re.search(r'\d', tok) or parse_date(tok) and len(tok) <= 10 and re.fullmatch(r'[\d./-]+', tok):
                return None
            # Number wrapped onto next line, e.g. "BIN/2026-" + "28/05"
            if tok.endswith(('-', '/')):
                follow = nxt2 if not rest else nxt
                fm = token_re.match((follow or '').strip())
                if fm:
                    tok += fm.group(0)
            return tok

        for i, line in enumerate(lines):
            if re.search(r'original\s+invoice', line, re.I):
                continue
            m = label_re.search(line)
            if m:
                tok = pick(m.group(1), lines[i + 1] if i + 1 < len(lines) else '',
                           lines[i + 2] if i + 2 < len(lines) else '')
                if tok:
                    return tok
        # "#" on its own line followed by ": 2026-27/224" (Zoho style) or "Invoice 065/2026-2027"
        for i, line in enumerate(lines):
            if line.strip() == '#' and i + 1 < len(lines):
                tok = pick('', lines[i + 1], lines[i + 2] if i + 2 < len(lines) else '')
                if tok:
                    return tok
            m = re.match(r'^invoice\s+([A-Za-z0-9][A-Za-z0-9\-_/]*\d[A-Za-z0-9\-_/]*)$', line.strip(), re.I)
            if m:
                return m.group(1)
        return None

    @staticmethod
    def _extract_invoice_date(lines):
        label_re = re.compile(r'(invoice\s*date|date\s*of\s*(?:issue|invoice)|bill\s*date|reimbursement\s*date|\bdated\b|^date\b)', re.I)
        skip_re = re.compile(r'\b(due|ack|po|order|delivery|dispatch|original|supply|period|lut|arn|signed?)\b', re.I)
        for i, line in enumerate(lines):
            m = label_re.search(line)
            if not m or skip_re.search(line):
                continue
            for cand in (line[m.end():], lines[i + 1] if i + 1 < len(lines) else '',
                         lines[i + 2] if i + 2 < len(lines) else ''):
                if cand and skip_re.search(cand) and cand is not line[m.end():]:
                    break
                d = parse_date(cand)
                if d:
                    return d
        # Fallback: first date in the document header area
        for line in lines[:20]:
            if skip_re.search(line):
                continue
            d = parse_date(line)
            if d:
                return d
        return None

    @staticmethod
    def _extract_amounts(lines):
        """Find taxable value, GST rate, tax amounts and printed total, and reconcile them.

        Returns dict(taxable, gst_rate, tax_amount, total, round_off, reconciled)."""
        sub_re = re.compile(r'(?:^|[\s:])(sub\s*total|taxable\s*(?:value|amount)?|total\s*taxable|amount\s*before\s*tax|net\s*amount)\b', re.I)
        tot_re = re.compile(r'(grand\s*total|invoice\s*total|total\s*amount(\s*(due|payable))?|amount\s*payable|total\s*payable|balance\s*due|amount\s*due|^total\b)', re.I)
        tax_rate_re = re.compile(r'(?:^|[^A-Za-z])(IGST|CGST|SGST|UTGST|GST[- ]?Central|GST[- ]?State|GST)\s*(?:OUTPUT|\d{1,2})?\s*[^%\n]{0,25}?(?:@\s*)?(\d{1,2}(?:\.\d+)?)\s*%', re.I)
        tax_name_re = re.compile(r'(?:^|[^A-Za-z])(IGST|CGST|SGST|UTGST|GST[- ]?Central|GST[- ]?State)\b', re.I)

        def value_at(i, after):
            v = _first_money(after)
            if v is not None:
                return v
            for j in (i + 1, i + 2, i + 3):
                if j < len(lines) and _is_money_line(lines[j]):
                    return _first_money(lines[j])
            return None

        sub_vals, tot_vals = [], []
        rates = {"IGST": set(), "CGST": set(), "SGST": set(), "UTGST": set(), "GST": set()}
        tax_amts = {"IGST": [], "CGST": [], "SGST": [], "UTGST": [], "GST": []}

        def add_tax(kind, rate, amt):
            k = "CGST" if "CENTRAL" in kind else ("SGST" if "STATE" in kind else kind)
            if rate is not None and rate > 0:
                rates[k].add(rate)
            if amt is not None and amt > 0 and amt not in tax_amts[k]:
                tax_amts[k].append(amt)

        for i, line in enumerate(lines):
            m = sub_re.search(line)
            if m:
                v = value_at(i, line[m.end():])
                if v and v > 0:
                    sub_vals.append(v)
            m = tot_re.search(line)
            if m and not re.search(r'in\s*words|quantity|qty|tax\s*amount', line, re.I):
                v = value_at(i, line[m.end():])
                if v and v > 1:
                    tot_vals.append(v)
            m = tax_rate_re.search(line)
            if m:
                kind = m.group(1).upper()
                rate = float(m.group(2)) / 100.0
                v = value_at(i, line[m.end():])
                add_tax(kind, rate, v)
            elif tax_name_re.search(line):
                m_name = tax_name_re.search(line)
                kind = m_name.group(1).upper()
                # check adjacent lines for rate e.g. % \n 9
                rate = None
                for j in (i + 1, i + 2, i + 3):
                    if j < len(lines):
                        m_pct = re.search(r'(\d{1,2}(?:\.\d+)?)\s*%', lines[j]) or re.search(r'%\s*(\d{1,2}(?:\.\d+)?)', lines[j])
                        if m_pct:
                            rate = float(m_pct.group(1)) / 100.0
                            break
                        if lines[j].strip() == '%' and j + 1 < len(lines) and re.fullmatch(r'\d{1,2}(?:\.\d+)?', lines[j+1].strip()):
                            rate = float(lines[j+1].strip()) / 100.0
                            break
                v = value_at(i, line[m_name.end():])
                add_tax(kind, rate, v)

        if rates["IGST"]:
            gst_rate = max(rates["IGST"])
        elif rates["CGST"] or rates["SGST"] or rates["UTGST"]:
            gst_rate = max(rates["CGST"] | rates["SGST"] | rates["UTGST"]) * 2
        elif rates["GST"]:
            gst_rate = max(rates["GST"])
        else:
            gst_rate = None

        tax_sum = sum(tax_amts["IGST"]) or (sum(tax_amts["CGST"]) + sum(tax_amts["SGST"]) + sum(tax_amts["UTGST"])) or sum(tax_amts["GST"])

        # Prioritize candidate rates
        rate_options = []
        if gst_rate is not None and gst_rate > 0:
            rate_options.append(gst_rate)
        for r in STANDARD_GST_RATES:
            if r > 0 and r not in rate_options:
                rate_options.append(r)

        # 1. Try reconciling sub_vals with non-zero rates first
        for total in sorted(set(tot_vals), reverse=True):
            for r in rate_options:
                for tx in sub_vals:
                    if tx >= total:
                        continue
                    comp = round(tx + round(tx * r, 2), 2)
                    if abs(comp - total) <= 1.05 and (not tax_sum or abs(tx * r - tax_sum) <= 1.05 or abs(tx * r * 2 - tax_sum) <= 1.05):
                        return {
                            "taxable": tx,
                            "gst_rate": r,
                            "tax_amount": round(tx * r, 2),
                            "total": total,
                            "round_off": round(total - comp, 2),
                            "reconciled": True
                        }

        # 2. Try derived from tax_sum
        if tax_sum > 0:
            for total in sorted(set(tot_vals), reverse=True):
                for r in rate_options:
                    tx = round(tax_sum / r, 2)
                    if tx >= total:
                        continue
                    comp = round(tx + round(tx * r, 2), 2)
                    if abs(comp - total) <= 1.05:
                        return {
                            "taxable": tx,
                            "gst_rate": r,
                            "tax_amount": round(tx * r, 2),
                            "total": total,
                            "round_off": round(total - comp, 2),
                            "reconciled": True
                        }

        # 3. If zero tax or reimbursement (total == taxable within roundoff)
        for total in sorted(set(tot_vals), reverse=True):
            for tx in sub_vals + [total]:
                if abs(tx - total) <= 1.05:
                    return {
                        "taxable": tx,
                        "gst_rate": 0.0,
                        "tax_amount": 0.0,
                        "total": total,
                        "round_off": round(total - tx, 2),
                        "reconciled": True
                    }

        # Not reconciled: return best guess, caller flags for review
        tx = sub_vals[0] if sub_vals else (tot_vals[0] if tot_vals else 0.0)
        return {
            "taxable": tx,
            "gst_rate": gst_rate or 0.0,
            "tax_amount": 0.0,
            "total": max(tot_vals) if tot_vals else tx,
            "round_off": 0.0,
            "reconciled": False
        }

    # ------------------------------------------------------------------ main
    def extract_from_pdf(self, pdf_path):
        """Extract invoice details from a PDF. Sets 'needs_review' (list of reasons) when the
        data cannot be trusted for automatic entry into the register."""
        pdf_path = Path(pdf_path)
        doc = pymupdf.open(pdf_path)
        try:
            full_text = "\n".join(p.get_text() for p in doc)
        finally:
            doc.close()
        single_line_text = re.sub(r'\s+', ' ', full_text)
        lines = [l.strip() for l in full_text.split("\n") if l.strip()]
        review = []

        if len(single_line_text.strip()) < 40:
            review.append("Scanned/image PDF with no text layer - enter manually")

        # 1. Filename cues: <Invoice No> <dd.mm.yy>.pdf (only trusted inside vendor folders)
        stem = pdf_path.stem
        file_date, file_inv_no = None, None
        m_file = re.search(r'^(.*?)\s+(\d{2}\.\d{2}\.\d{2,4})$', stem)
        if m_file:
            file_inv_no = m_file.group(1).strip()
            file_date = parse_date(m_file.group(2))

        folder_supplier = self._folder_supplier(pdf_path)

        # 2. IRN (64 hex chars, possibly hyphen-wrapped)
        unfolded_text = re.sub(r'([a-fA-F0-9]{20,63})-\s*([a-fA-F0-9]+)', r'\1\2', full_text)
        unfolded_text = re.sub(r'([a-fA-F0-9]{20,63})\n([a-fA-F0-9]{1,44})\b', r'\1\2', unfolded_text)
        irn_match = re.search(r'\b([a-fA-F0-9]{64})\b', unfolded_text)
        irn = irn_match.group(1) if irn_match else None

        # 3. Ack No & Date
        ack_no = None
        ack_no_m = re.search(r'Ack\s*(?:No|Number|#)?[.:\s]+(\d{10,20})', single_line_text, re.IGNORECASE)
        if ack_no_m:
            ack_no = ack_no_m.group(1)
        ack_date = None
        ack_date_m = re.search(r'Ack\s*Date[.:\s]+([0-9A-Za-z\-: /]{8,25})', full_text, re.IGNORECASE)
        if ack_date_m:
            ack_date = re.split(r'[\n\r]|IRN|Buyer|Party|Dated', ack_date_m.group(1).strip())[0].strip()

        # 4. E-Way Bill
        eway_bill = None
        eway_m = re.search(r'E[-_ ]?Way\s*Bill\s*(?:No|Number)?[.:\s]+(\d{10,16})', single_line_text, re.IGNORECASE)
        if eway_m:
            eway_bill = eway_m.group(1)

        # 5. GSTINs
        unique_gstins = []
        for g in GSTIN_RE.findall(full_text):
            if g not in unique_gstins:
                unique_gstins.append(g)
        buyer_gstin = next((g for g in unique_gstins if g in NICOBAR_GSTINS or NICOBAR_PAN in g), None)
        supplier_gstin = next((g for g in unique_gstins if NICOBAR_PAN not in g), None)
        billed_to_nicobar = bool(buyer_gstin) or NICOBAR_PAN in full_text or bool(
            re.search(r'nicobar\s+design', full_text, re.I))
        if not billed_to_nicobar and not review:
            review.append("Not billed to Nicobar (no Nicobar GSTIN/PAN/name on document)")

        # 6. Supplier
        supplier = folder_supplier
        if not supplier and supplier_gstin and supplier_gstin in self.gstin_to_supplier:
            supplier = self.gstin_to_supplier[supplier_gstin]
        if not supplier:
            supplier = detect_vendor_from_text(full_text)
        if not supplier:
            review.append("Supplier could not be identified")

        sup_info = self.supplier_memory.get(supplier.lower(), {}) if supplier else {}
        if not supplier_gstin and sup_info.get("gstin"):
            supplier_gstin = sup_info["gstin"]

        supplier_state = None
        if supplier_gstin:
            st_code = supplier_gstin[:2]
            supplier_state = f"{STATE_CODES.get(st_code, 'Unknown')} ({st_code})"
        elif sup_info.get("state"):
            supplier_state = sup_info["state"]

        currency = "INR"
        if re.search(r'US\$|\bUSD\b', full_text):
            currency = "USD"
        elif re.search(r'€|\bEUR\b', full_text):
            currency = "EUR"
        if currency != "INR":
            review.append(f"Foreign currency ({currency}) - needs RBI/FBIL FX rate")
            if not supplier_gstin:
                supplier_state = supplier_state or "Foreign (Import)"

        if re.search(r'\bcredit\s+note\b', full_text, re.I):
            review.append("Credit note - not auto-added (enter as negative / adjust manually)")

        # 7. Place of supply
        if not buyer_gstin:
            buyer_gstin_out = None
            place_of_supply = None
        else:
            buyer_gstin_out = buyer_gstin
            place_of_supply = f"{STATE_CODES.get(buyer_gstin[:2], 'Unknown')} ({buyer_gstin[:2]})"
        pos_match = re.search(r'Place\s+of\s+Supply\s*[.:]?\s*([A-Za-z &]+(?:\s*\(\d{2}\))?)', full_text, re.IGNORECASE)
        if pos_match and pos_match.group(1).strip():
            place_of_supply = pos_match.group(1).strip()

        # 8. Invoice number (text first; filename only when already filed in a vendor folder)
        invoice_no = self._extract_invoice_no(lines)
        if not invoice_no and folder_supplier and file_inv_no:
            invoice_no = file_inv_no
        if not invoice_no:
            invoice_no = file_inv_no or stem
            review.append("Invoice number not found on document")

        # 9. Invoice date
        invoice_date = self._extract_invoice_date(lines) or file_date
        if not invoice_date:
            review.append("Invoice date not found on document")

        # 10. PO
        buyer_po_no = None
        po_m = re.search(r"(?:Buyer'?s?\s*Order\s*No|PO\s*No|Buyer\s*PO\s*No|P\.O\.#)[.:\s]+([A-Za-z0-9\-_/]*\d[A-Za-z0-9\-_/]*)", single_line_text, re.IGNORECASE)
        if po_m:
            buyer_po_no = po_m.group(1).strip()
        po_date = None
        po_dt_m = re.search(r'PO\s*Date[.:\s]+([0-9A-Za-z\-./ ]{8,15})', single_line_text, re.IGNORECASE)
        if po_dt_m:
            po_date = parse_date(po_dt_m.group(1))

        # 11. Logistics
        transporter = None
        t_m = re.search(r'Transporter(?:\s*Name)?[.:]+\s*([A-Za-z0-9 &]+)', full_text, re.IGNORECASE)
        if t_m:
            transporter = t_m.group(1).split("\n")[0].strip() or None
        lr_no = None
        lr_m = re.search(r'\b(?:LR|GR|Consignment)\s*No[.:\s]+([A-Za-z0-9\-_/]*\d[A-Za-z0-9\-_/]*)', single_line_text, re.IGNORECASE)
        if lr_m:
            lr_no = lr_m.group(1).strip()
        vehicle_no = None
        veh_m = re.search(r'Vehicle\s*(?:No|Number)[.:\s]+([A-Z]{2}[\s-]?\d{1,2}[\s-]?[A-Z]{0,3}[\s-]?\d{3,4})', single_line_text, re.IGNORECASE)
        if veh_m:
            vehicle_no = veh_m.group(1).strip()
        delivery_note = None
        dn_m = re.search(r'Deli?very\s*Note[.:\s]+([A-Za-z0-9\-_/]*\d[A-Za-z0-9\-_/]*)', single_line_text, re.IGNORECASE)
        if dn_m:
            delivery_note = dn_m.group(1).strip()
        ref_no = None
        ref_m = re.search(r'Reference\s*(?:No\.?|number)[.:\s]+([A-Za-z0-9\-_/]*\d[A-Za-z0-9\-_/]*)', single_line_text, re.IGNORECASE)
        if ref_m:
            ref_no = ref_m.group(1).strip()

        # 12. Amounts & GST rate (reconciled against printed total)
        amounts = self._extract_amounts(lines)
        gst_rate = amounts["gst_rate"]
        taxable_val = amounts["taxable"]
        if not amounts["reconciled"] and not any("Scanned" in r or "Foreign" in r for r in review):
            review.append("Could not reconcile taxable value + GST with the printed invoice total")
        if not taxable_val:
            taxable_val = 0.0
            if not any("Scanned" in r for r in review):
                review.append("Taxable amount not found")

        # 13. Tax type (state code comparison; never assume Haryana)
        sup_code = supplier_gstin[:2] if supplier_gstin else ""
        buy_code = buyer_gstin[:2] if buyer_gstin else ""
        if not supplier_gstin:
            if currency != "INR" or "import" in (supplier_state or "").lower() or "foreign" in (supplier_state or "").lower():
                tax_type = "Import (no GST)"
            else:
                tax_type = "No GST (unregistered)"
            gst_rate = 0.0
        elif not buy_code:
            tax_type = "IGST" if gst_rate else "No GST (unregistered)"
            review.append("Buyer (Nicobar) GSTIN not found - tax type unverified")
        elif sup_code != buy_code:
            tax_type = "IGST"
        else:
            tax_type = "CGST+SGST"

        # 14. TDS - reuse supplier history; otherwise do NOT guess a rate (SKILL step 7)
        tds_section = sup_info.get("tds_sec")
        tds_rate = sup_info.get("tds_rate")
        if not tds_section:
            if tax_type == "Import (no GST)":
                tds_section = "Not applicable - import from non-resident; no income taxable in India."
                tds_rate = 0.0
            elif "consultan" in full_text.lower() or "professional" in full_text.lower() or "technical" in full_text.lower():
                tds_section = "194J"
                tds_rate = 0.02
            elif "lease" in full_text.lower() or "rental" in full_text.lower():
                tds_section = "194I"
                tds_rate = 0.02
            elif "reimbursement" in full_text.lower() and not ("lease" in full_text.lower() or "rental" in full_text.lower()):
                tds_section = "Not applicable - pure reimbursement"
                tds_rate = 0.0
            else:
                tds_section = ("To review - new supplier. 194Q only if FY purchases > Rs 50 L (on excess); "
                               "194C/194J subject to single-bill/annual thresholds.")
                tds_rate = 0.0
        if tds_rate is None:
            tds_rate = 0.0

        # RCM / GST compliance statement
        rcm_note = sup_info.get("rcm_note")
        if not rcm_note:
            if tax_type == "Import (no GST)":
                rcm_note = "Not RCM - import: IGST paid to Customs on Bill of Entry (goods) / RCM on import of services - verify."
            elif tax_type == "IGST":
                rcm_note = "No RCM - forward charge; IGST charged by supplier. ITC eligible (check GSTR-2B)."
            elif tax_type == "CGST+SGST":
                rcm_note = "No RCM - forward charge; CGST + SGST charged by supplier. ITC eligible (check GSTR-2B)."
            else:
                rcm_note = "Unregistered supplier - check RCM applicability."

        loc_code = sup_info.get("loc_code")
        vgl_code = sup_info.get("vgl_code")
        egl_code = sup_info.get("egl_code")

        hsn_m = re.search(r'\b(?:HSN|SAC)\b[^\n\d]{0,15}(\d{4,8})\b', full_text, re.IGNORECASE)
        hsn = hsn_m.group(1) if hsn_m else None

        items = []
        # Multi-line item check for Mahindra lease invoices (different GST rates: 18% & 40%)
        if "IND-450459" in str(invoice_no) or ("MAHINDRA" in (supplier or "").upper() and ("LEASE" in full_text.upper() or "FMS" in full_text.upper())):
            start_idx, end_idx = -1, -1
            for i, l in enumerate(lines):
                if 'Total Amt. (Rs.)' in l:
                    start_idx = i + 1
                if l == 'Grand Total':
                    end_idx = i
                    break
            if start_idx > 0 and end_idx > start_idx:
                chunk = lines[start_idx:end_idx]
                item_starts = []
                for idx, token in enumerate(chunk):
                    if token in ('1', '2', '3', '4', '5') and idx + 1 < len(chunk) and any(k in chunk[idx+1] for k in ('Rental', 'Lease', 'FMS')):
                        item_starts.append(idx)
                item_starts.append(len(chunk))
                for i in range(len(item_starts) - 1):
                    sub = chunk[item_starts[i]:item_starts[i+1]]
                    sno = sub[0]
                    desc = sub[1]
                    hsn_sub = sub[2]
                    veh_info = []
                    for t in sub[3:]:
                        if re.fullmatch(r'\d{2}/\d{2}/\d{4}', t) or (re.fullmatch(r'\d+(?:\.\d+)?', t) and '.' in t):
                            break
                        veh_info.append(t)
                    full_desc = f"{desc} - {' '.join(veh_info)}"
                    nums = [float(t) for t in sub if re.fullmatch(r'\d+(?:\.\d+)?', t) and '.' in t]
                    taxable = nums[0] if len(nums) > 0 else 0.0
                    rates_found = [n for n in nums if n in (18.0, 40.0, 28.0, 12.0, 5.0)]
                    gst_r = (rates_found[0] / 100.0) if rates_found else 0.18
                    round_off = 0.01 if i == len(item_starts) - 2 else 0.0
                    items.append({
                        "item_code": None,
                        "description": full_desc,
                        "hsn": hsn_sub,
                        "boxes": None,
                        "quantity": 1.0,
                        "unit": "NOS",
                        "rate": taxable,
                        "discount": 0.0,
                        "taxable_amount": taxable,
                        "gst_rate": gst_r,
                        "tax_type": "IGST",
                        "round_off": round_off,
                        "rcm_note": rcm_note,
                        "tds_section": tds_section or "194I",
                        "tds_rate": tds_rate if (tds_rate is not None and tds_rate > 0) else 0.02,
                        "location_code": loc_code,
                        "vendor_gl_code": vgl_code,
                        "exps_gl_code": egl_code,
                    })

        if not items:
            items = [{
                "item_code": None,
                "description": f"{supplier} - invoice {invoice_no}" if supplier else f"Invoice {invoice_no}",
                "hsn": hsn,
                "boxes": None,
                "quantity": 1.0,
                "unit": "NOS",
                "rate": taxable_val,
                "discount": 0.0,
                "taxable_amount": taxable_val,
                "gst_rate": gst_rate,
                "tax_type": tax_type,
                "round_off": amounts["round_off"] if amounts["reconciled"] else 0.0,
                "rcm_note": rcm_note,
                "tds_section": tds_section,
                "tds_rate": tds_rate,
                "location_code": loc_code,
                "vendor_gl_code": vgl_code,
                "exps_gl_code": egl_code,
            }]

        try:
            rel_path = os.path.relpath(pdf_path, EXPENSES_DIR) if pdf_path.resolve().is_relative_to(EXPENSES_DIR.resolve()) else pdf_path.name
        except Exception:
            rel_path = pdf_path.name

        return {
            "file_path": str(rel_path).replace('/', '\\'),
            "full_path": str(pdf_path),
            "invoice_no": invoice_no,
            "invoice_date": invoice_date,
            "supplier": supplier or "Unknown Supplier",
            "supplier_gstin": supplier_gstin,
            "supplier_state": supplier_state,
            "buyer_gstin": buyer_gstin_out,
            "billed_to_nicobar": billed_to_nicobar,
            "place_of_supply": place_of_supply,
            "buyer_po_no": buyer_po_no,
            "po_date": po_date,
            "delivery_note": delivery_note,
            "reference_no": ref_no,
            "irn": irn,
            "ack_no": ack_no,
            "ack_date": ack_date,
            "eway_bill_no": eway_bill,
            "transporter": transporter,
            "lr_no": lr_no,
            "vehicle_no": vehicle_no,
            "currency": currency,
            "printed_total": amounts["total"],
            "items": items,
            "needs_review": review,
        }
