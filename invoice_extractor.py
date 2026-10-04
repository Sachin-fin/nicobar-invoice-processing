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
    IGNORED_FILES,
    IGNORED_DIRS,
)

class InvoiceExtractor:
    def __init__(self):
        self.supplier_memory = {}
        self.gstin_to_supplier = {}
        self.load_supplier_memory()

    def load_supplier_memory(self):
        """Build supplier profile and GL code memory from existing register and vendor code file."""
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
                        if len(g_clean) == 15 and "AAFCN6567R" not in g_clean:
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
                                "name": sup_name,
                                "gstin": None,
                                "state": None,
                                "tds_sec": None,
                                "tds_rate": 0,
                                "loc_code": None,
                                "vgl_code": vgl,
                                "egl_code": egl,
                                "tax_type": None,
                                "rcm_note": None,
                            }
            except Exception as e:
                print(f"Warning: could not read vendor codes: {e}")

    def clean_text(self, text):
        return re.sub(r'[ \t]+', ' ', text).strip()

    def parse_date(self, text):
        """Parse various date formats into datetime.date."""
        if not text:
            return None
        text = str(text).strip()
        patterns = [
            r'(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})',
            r'(\d{1,2})(?:st|nd|rd|th)?\s+([A-Za-z]+)\s+(\d{2,4})',
            r'([A-Za-z]+)\s+(\d{1,2})(?:st|nd|rd|th)?,?\s+(\d{2,4})',
            r'(\d{4})[./-](\d{1,2})[./-](\d{1,2})',
        ]
        for pat in patterns:
            m = re.search(pat, text)
            if m:
                groups = m.groups()
                # Case 1: dd-mm-yy(yy)
                if len(groups) == 3 and groups[0].isdigit() and groups[1].isdigit():
                    d, mth, y = int(groups[0]), int(groups[1]), int(groups[2])
                    if y < 100: y += 2000
                    if 1 <= mth <= 12 and 1 <= d <= 31:
                        try: return datetime.date(y, mth, d)
                        except ValueError: pass
                # Case 2: yyyy-mm-dd
                elif len(groups) == 3 and len(groups[0]) == 4 and groups[0].isdigit():
                    y, mth, d = int(groups[0]), int(groups[1]), int(groups[2])
                    if 1 <= mth <= 12 and 1 <= d <= 31:
                        try: return datetime.date(y, mth, d)
                        except ValueError: pass
                # Case 3: dd Month yyyy
                elif len(groups) == 3:
                    for fmt in ("%d %B %Y", "%d %b %Y", "%B %d %Y", "%b %d %Y", "%d %b %y"):
                        try:
                            clean_str = f"{groups[0]} {groups[1]} {groups[2]}"
                            return datetime.datetime.strptime(clean_str, fmt).date()
                        except Exception:
                            pass
        return None

    def extract_from_pdf(self, pdf_path):
        """Extract invoice details and line items from a PDF file."""
        pdf_path = Path(pdf_path)
        doc = pymupdf.open(pdf_path)
        pages_text = [p.get_text() for p in doc]
        full_text = "\n".join(pages_text)
        single_line_text = re.sub(r'\s+', ' ', full_text)

        # 1. Filename cues: <Invoice No> <dd.mm.yy>.pdf
        stem = pdf_path.stem
        file_date = None
        file_inv_no = stem
        m_file = re.search(r'^(.*?)\s+(\d{2}\.\d{2}\.\d{2,4})$', stem)
        if m_file:
            file_inv_no = m_file.group(1).strip()
            file_date = self.parse_date(m_file.group(2))

        # Supplier from parent folder
        folder_supplier = pdf_path.parent.name if pdf_path.parent.name != EXPENSES_DIR.name else None

        # 2. Extract IRN (64 hex characters)
        # Handle hyphenated breaks across lines
        unfolded_text = re.sub(r'([a-fA-F0-9]{20,63})-\s*([a-fA-F0-9]+)', r'\1\2', full_text)
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
            ack_date_str = ack_date_m.group(1).strip()
            # Clean up trailing tokens
            ack_date_str = re.split(r'[\n\r]|IRN|Buyer|Party|Dated', ack_date_str)[0].strip()
            ack_date = ack_date_str

        # 4. E-Way Bill No
        eway_bill = None
        eway_m = re.search(r'E[-_ ]?Way\s*Bill\s*(?:No|Number)?[.:\s]+(\d{10,16})', single_line_text, re.IGNORECASE)
        if eway_m:
            eway_bill = eway_m.group(1)

        # 5. GSTINs
        gstins = re.findall(r'\b([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1})\b', full_text)
        # Deduplicate while preserving order
        unique_gstins = []
        for g in gstins:
            if g not in unique_gstins:
                unique_gstins.append(g)

        buyer_gstin = None
        supplier_gstin = None

        # Check Nicobar GSTINs
        for g in unique_gstins:
            if g in NICOBAR_GSTINS or "AAFCN6567R" in g:
                if not buyer_gstin:
                    buyer_gstin = g

        for g in unique_gstins:
            if g != buyer_gstin:
                if not supplier_gstin:
                    supplier_gstin = g

        # 6. Supplier Name & State
        supplier = folder_supplier
        if not supplier or "_temp" in str(supplier).lower() or supplier == "Unknown_Supplier":
            if supplier_gstin and supplier_gstin in self.gstin_to_supplier:
                supplier = self.gstin_to_supplier[supplier_gstin]
            else:
                for g in unique_gstins:
                    if g in self.gstin_to_supplier:
                        supplier = self.gstin_to_supplier[g]
                        supplier_gstin = g
                        break

            if not supplier:
                u_txt = full_text.upper()
                if "PREZOTECH" in u_txt: supplier = "Prezotech Solutions Pvt Ltd"
                elif "GOWRI" in u_txt or "PARVATHY" in u_txt: supplier = "Gowri Parvathy"
                elif "A & S CREATION" in u_txt or "07ABNFA6509L" in u_txt: supplier = "A & S Creations"
                elif "CUSHMAN" in u_txt or "APIPL" in u_txt: supplier = "Cushman & Wakefield India Private Limited"
                elif "AMBIENCE FACILITIES" in u_txt: supplier = "Ambience Facilities Management Private Limited"
                elif "AMBIENCE DEVELOPERS" in u_txt: supplier = "Ambience Developers & Infrastructure Pvt Ltd"
                elif "NIVA" in u_txt or "MACRAMEDESIGNS" in u_txt: supplier = "Niva's Designs"
                elif "CUREFOODS" in u_txt or "BIN-" in u_txt: supplier = "Curefoods India Limited"
                elif "MAHINDRA" in u_txt or "QUIKLYZ" in u_txt: supplier = "Mahindra & Mahindra Financial Services Limited"
                elif "XTO10X" in u_txt: supplier = "XTO10X Technologies Private Limited"
                elif "CLAY CRAFT" in u_txt: supplier = "Clay Craft India Limited"
                elif "STITCH" in u_txt: supplier = "Stitch-9"
                elif "MKR APP" in u_txt: supplier = "MKR APP Private Limited"
                elif "AIRTEL" in u_txt or "BHARTI" in u_txt: supplier = "Bharti Airtel Limited"
                elif "BLUE DART" in u_txt: supplier = "Blue Dart Express Limited"
                elif "DLF" in u_txt: supplier = "DLF Home Developers Limited"
                elif "DTDC" in u_txt: supplier = "DTDC Express Limited"
                elif "HANSA ELITE" in u_txt: supplier = "Hansa Elite Luxe Interiors Pvt Ltd"
                elif "IPINFO" in u_txt: supplier = "IPinfo Inc"
                elif "SIMRAN" in u_txt: supplier = "Simran"
                elif "SURBHI SETHI" in u_txt: supplier = "Surbhi Sethi"
                elif "SHREYAS GAIKWAD" in u_txt: supplier = "Shreyas Gaikwad"
                elif "JATIN GULABRAI" in u_txt: supplier = "Jatin Gulabrai Parekh"
                elif "ANTHROPIC" in u_txt: supplier = "Anthropic PBC"
                elif "QUANG VINH" in u_txt: supplier = "Quang Vinh Ceramic Co Ltd"
                elif "TRION PROPERTIES" in u_txt: supplier = "Trion Properties Private Limited"
                elif "V DINE" in u_txt: supplier = "V Dine Unit of VRSM Enterprises LLP"
                elif "VODAFONE" in u_txt or "VI BUSINESS" in u_txt: supplier = "Vodafone Idea Limited"
                elif "VRAJ COMMERCIAL" in u_txt: supplier = "Vraj Commercial Private Limited"

        supplier_state = None
        if supplier_gstin:
            st_code = supplier_gstin[:2]
            st_name = STATE_CODES.get(st_code, "Unknown")
            supplier_state = f"{st_name} ({st_code})"

        # Check memory for supplier defaults
        sup_info = {}
        if supplier and supplier.lower() in self.supplier_memory:
            sup_info = self.supplier_memory[supplier.lower()]
            if not supplier_gstin and sup_info.get("gstin"):
                supplier_gstin = sup_info["gstin"]
            if not supplier_state and sup_info.get("state"):
                supplier_state = sup_info["state"]

        # If foreign supplier without GSTIN
        if not supplier_gstin:
            if "vietnam" in full_text.lower() or (supplier and "vietnam" in supplier.lower()):
                supplier_state = "Vietnam (Import)"
            elif "anthropic" in full_text.lower() or (supplier and "anthropic" in supplier.lower()):
                supplier_state = "USA (Import)"

        # 7. Buyer GSTIN and Place of Supply
        if not buyer_gstin:
            buyer_gstin = "06AAFCN6567R1ZL"  # Default Nicobar HQ

        pos_match = re.search(r'Place of Supply[.:\s]+([A-Za-z &]+(?:\s*\(\d{2}\))?)', full_text, re.IGNORECASE)
        if pos_match:
            place_of_supply = pos_match.group(1).strip()
        else:
            b_code = buyer_gstin[:2]
            place_of_supply = f"{STATE_CODES.get(b_code, 'Haryana')} ({b_code})"

        # 8. Invoice Number
        invoice_no = None
        inv_patterns = [
            r'Invoice\s*No\.?[.:\s]+([A-Za-z0-9\-_/]+)',
            r'Tax\s*Invoice\s*No\.?[.:\s]+([A-Za-z0-9\-_/]+)',
            r'Bill\s*No\.?[.:\s]+([A-Za-z0-9\-_/]+)',
            r'Invoice\s*#\s*[:.\s]+([A-Za-z0-9\-_/]+)',
        ]
        for pat in inv_patterns:
            m = re.search(pat, full_text, re.IGNORECASE)
            if m:
                candidate = m.group(1).strip()
                if len(candidate) >= 3 and not candidate.lower().startswith(("date", "mode", "term")):
                    invoice_no = candidate
                    break

        if not invoice_no and file_inv_no:
            # Fallback to file name invoice no with common slash restoration
            invoice_no = file_inv_no.replace("-", "/")

        # 9. Invoice Date
        invoice_date = None
        dt_patterns = [
            r'Dated?[.:\s]+([0-9A-Za-z\-./ ]{8,20})',
            r'Invoice\s*Date[.:\s]+([0-9A-Za-z\-./ ]{8,20})',
            r'Date\s*of\s*Invoice[.:\s]+([0-9A-Za-z\-./ ]{8,20})',
        ]
        for pat in dt_patterns:
            m = re.search(pat, full_text, re.IGNORECASE)
            if m:
                d = self.parse_date(m.group(1))
                if d:
                    invoice_date = d
                    break
        if not invoice_date and file_date:
            invoice_date = file_date
        if not invoice_date:
            invoice_date = datetime.date.today()

        # 10. Purchase Order (PO No. & Date)
        buyer_po_no = None
        po_date = None
        po_m = re.search(r'(?:Buyer\'?s?\s*Order\s*No|PO\s*No|Buyer\s*PO\s*No)[.:\s]+([A-Za-z0-9\-_/]+)', single_line_text, re.IGNORECASE)
        if po_m:
            buyer_po_no = po_m.group(1).strip()

        po_dt_m = re.search(r'PO\s*Date[.:\s]+([0-9A-Za-z\-./ ]{8,15})', single_line_text, re.IGNORECASE)
        if po_dt_m:
            po_date = self.parse_date(po_dt_m.group(1))

        # 11. Logistics: Transporter, LR, Vehicle
        transporter = None
        t_m = re.search(r'Transporter?(?: Name)?[.:\s]+([A-Za-z0-9 &]+)', full_text, re.IGNORECASE)
        if t_m:
            transporter = t_m.group(1).split("\n")[0].strip()

        lr_no = None
        lr_m = re.search(r'(?:LR|GR|Consignment)\s*No[.:\s]+([A-Za-z0-9\-_/]+)', single_line_text, re.IGNORECASE)
        if lr_m:
            lr_no = lr_m.group(1).strip()

        vehicle_no = None
        veh_m = re.search(r'Vehicle\s*(?:No|Number)[.:\s]+([A-Z0-9\-_/ ]{6,15})', single_line_text, re.IGNORECASE)
        if veh_m:
            vehicle_no = veh_m.group(1).strip()

        delivery_note = None
        dn_m = re.search(r'Delie?very\s*Note[.:\s]+([A-Za-z0-9\-_/]+)', single_line_text, re.IGNORECASE)
        if dn_m:
            delivery_note = dn_m.group(1).strip()

        ref_no = None
        ref_m = re.search(r'Reference\s*No\.?[.:\s]+([A-Za-z0-9\-_/]+)', single_line_text, re.IGNORECASE)
        if ref_m:
            ref_no = ref_m.group(1).strip()

        # 12. Tax Type
        sup_code = supplier_gstin[:2] if supplier_gstin else ""
        buy_code = buyer_gstin[:2] if buyer_gstin else ""
        
        if not supplier_gstin:
            if "import" in (supplier_state or "").lower():
                tax_type = "Import (no GST)"
                gst_rate = 0.0
            else:
                tax_type = "No GST (unregistered)"
                gst_rate = 0.0
        elif sup_code != buy_code:
            tax_type = "IGST"
            gst_rate = 0.18 # Default standard if not detected
        else:
            tax_type = "CGST+SGST"
            gst_rate = 0.18

        # 13. Line Items & Amounts
        # Look for total / summary amounts
        total_m = re.search(r'(?:Grand\s*Total|Invoice\s*Total|Total\s*Amount|Total)[.:\s₹Rs]+([\d,]+(?:\.\d{1,2})?)', full_text, re.IGNORECASE)
        round_off_m = re.search(r'Round\s*Off[.:\s₹Rs]+(-?[\d,]+(?:\.\d{1,2})?)', full_text, re.IGNORECASE)
        
        round_off = 0.0
        if round_off_m:
            try:
                round_off = float(round_off_m.group(1).replace(',', ''))
            except Exception:
                round_off = 0.0

        # TDS Section & Rate
        tds_section = sup_info.get("tds_sec")
        tds_rate = sup_info.get("tds_rate") if sup_info.get("tds_rate") is not None else 0.001
        
        if not tds_section:
            if tax_type == "Import (no GST)":
                tds_section = "Not applicable – import of goods from non-resident, no income taxable in India (old 195 / 194Q N/A)."
                tds_rate = 0.0
            elif "contract" in full_text.lower() or "service" in full_text.lower() or "maintenance" in full_text.lower():
                tds_section = "Sec 393(1) – contract/service (old 194C), 2% (company). Confirm with CA."
                tds_rate = 0.02
            else:
                tds_section = "Sec 393(1) – purchase of goods (old 194Q), code 1031. Apply only if FY purchases from this supplier > ₹50 L (on excess)."
                tds_rate = 0.001

        # RCM / GST Compliance statement
        rcm_note = sup_info.get("rcm_note")
        if not rcm_note:
            if tax_type == "Import (no GST)":
                rcm_note = "Not RCM – import of goods: IGST + BCD paid to Customs on Bill of Entry; claim IGST ITC from BoE."
            elif tax_type == "IGST":
                rcm_note = "No RCM – forward charge; goods/services, IGST charged by supplier. ITC eligible (check GSTR-2B)."
            else:
                rcm_note = "No RCM – forward charge; CGST + SGST charged by supplier. ITC eligible (check GSTR-2B)."

        # Codes
        loc_code = sup_info.get("loc_code")
        vgl_code = sup_info.get("vgl_code")
        egl_code = sup_info.get("egl_code")

        # Parse Line Items (extract lines or generate primary line)
        items = []
        # Attempt table extraction
        table_lines = self._extract_table_lines(doc, full_text)
        if table_lines:
            for itm in table_lines:
                itm["tax_type"] = tax_type
                itm["round_off"] = round_off if len(items) == len(table_lines) - 1 else 0.0
                itm["rcm_note"] = rcm_note
                itm["tds_section"] = tds_section
                itm["tds_rate"] = tds_rate
                itm["location_code"] = loc_code
                itm["vendor_gl_code"] = vgl_code
                itm["exps_gl_code"] = egl_code
                items.append(itm)
        else:
            # Fallback to single line summary
            taxable_val = 0.0
            tax_m = re.search(r'Taxable\s*(?:Amount|Value)[.:\s₹Rs]+([\d,]+(?:\.\d{1,2})?)', full_text, re.IGNORECASE)
            if tax_m:
                try: taxable_val = float(tax_m.group(1).replace(',', ''))
                except Exception: pass
            
            if taxable_val == 0.0 and total_m:
                try:
                    tot = float(total_m.group(1).replace(',', ''))
                    taxable_val = round(tot / (1.0 + gst_rate), 2)
                except Exception:
                    taxable_val = 0.0

            desc = f"Expense / Services - {supplier}" if supplier else "Expense Item"
            hsn_m = re.search(r'\b(HSN|SAC)[/\s.:]+(\d{4,8})\b', full_text, re.IGNORECASE)
            hsn = hsn_m.group(2) if hsn_m else None

            items.append({
                "item_code": None,
                "description": desc,
                "hsn": hsn,
                "boxes": None,
                "quantity": 1.0,
                "unit": "NOS",
                "rate": taxable_val,
                "discount": 0.0,
                "taxable_amount": taxable_val,
                "gst_rate": gst_rate,
                "tax_type": tax_type,
                "round_off": round_off,
                "rcm_note": rcm_note,
                "tds_section": tds_section,
                "tds_rate": tds_rate,
                "location_code": loc_code,
                "vendor_gl_code": vgl_code,
                "exps_gl_code": egl_code,
            })

        rel_path = os.path.relpath(pdf_path, EXPENSES_DIR) if pdf_path.is_relative_to(EXPENSES_DIR) else pdf_path.name

        return {
            "file_path": str(rel_path).replace('/', '\\'),
            "full_path": str(pdf_path),
            "invoice_no": invoice_no or stem,
            "invoice_date": invoice_date,
            "supplier": supplier or "Unknown Supplier",
            "supplier_gstin": supplier_gstin,
            "supplier_state": supplier_state,
            "buyer_gstin": buyer_gstin,
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
            "items": items,
        }

    def _extract_table_lines(self, doc, full_text):
        """Helper to find structured items if present."""
        items = []
        # Check for multi-line pattern: Style Code / Item Code, Description, HSN, Qty, Rate, Taxable
        pattern = re.compile(
            r'(\d+)\.\s+([A-Za-z0-9\-_]+)?\s*[\r\n]+'
            r'(\d{4,8})\s+([A-Za-z0-9\-_]+)?\s*[\r\n]+'
            r'([\d,]+(?:\.\d+)?)\s*([A-Za-z]+)\s*[\r\n]+'
            r'([\d,]+(?:\.\d+)?)\s*[\r\n]+'
            r'([\d,]+(?:\.\d+)?)\s*[\r\n]+'
            r'([\d,]+(?:\.\d+)?)',
            re.MULTILINE
        )
        matches = list(pattern.finditer(full_text))
        if matches:
            for m in matches:
                hsn = m.group(3)
                qty = float(m.group(5).replace(',', ''))
                unit = m.group(6)
                rate = float(m.group(7).replace(',', ''))
                taxable = float(m.group(9).replace(',', ''))
                items.append({
                    "item_code": m.group(2) or m.group(4),
                    "description": "Item " + (m.group(2) or ""),
                    "hsn": hsn,
                    "boxes": None,
                    "quantity": qty,
                    "unit": unit,
                    "rate": rate,
                    "discount": 0.0,
                    "taxable_amount": taxable,
                    "gst_rate": 0.18,
                })
        return items
