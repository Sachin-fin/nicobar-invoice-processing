import os
import sys
import re
import shutil
import zipfile
import subprocess
import datetime
import xml.etree.ElementTree as ET
from pathlib import Path
import openpyxl

from config import (
    EXPENSES_DIR,
    REGISTER_PATH,
    VENDOR_CODES_PATH,
    DASHBOARD_HTML_PATH,
    BUILD_DASHBOARD_SCRIPT,
)

NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
ET.register_namespace('', NS_MAIN)

def col_letter(col_idx):
    """Convert 1-based column index to Excel column letters (A, B, ..., Z, AA, ...)."""
    result = ""
    while col_idx > 0:
        col_idx, rem = divmod(col_idx - 1, 26)
        result = chr(65 + rem) + result
    return result

def normalize_str(s):
    if not s:
        return ""
    return re.sub(r'[\s\-/_#.]+', '', str(s)).lower()

class ExcelManager:
    def __init__(self):
        self.register_path = REGISTER_PATH
        self.codes_path = VENDOR_CODES_PATH
        self._cached_state = None
        self._cached_mtime = 0

    def get_register_state(self, force_reload=False):
        """Read current register state (data rows, column mapping, KPIs) with mtime caching."""
        if not self.register_path.exists():
            return {"error": "Register file not found"}

        current_mtime = self.register_path.stat().st_mtime
        if not force_reload and self._cached_state is not None and self._cached_mtime == current_mtime:
            return self._cached_state

        wb = openpyxl.load_workbook(self.register_path, data_only=True, read_only=True)
        ws = wb["Invoice Register"]
        hdr = [c.value for c in ws[4]]
        col = {h: i for i, h in enumerate(hdr) if h}

        rows = []
        registered_files = set()
        registered_invoices = []
        registered_irns = set()
        max_sno = 0

        tot_taxable = 0.0
        tot_igst = 0.0
        tot_cgst = 0.0
        tot_sgst = 0.0
        tot_invoice = 0.0
        tot_tds = 0.0
        tot_net = 0.0
        unique_suppliers = set()

        for r_vals in ws.iter_rows(min_row=5, values_only=True):
            if not r_vals:
                continue
            sno_idx = col.get("S.No.", 0)
            if sno_idx >= len(r_vals):
                continue
            sno_val = r_vals[sno_idx]
            if not isinstance(sno_val, (int, float)):
                continue

            sno = int(sno_val)
            if sno > max_sno:
                max_sno = sno

            def gv(h, default=None):
                idx = col.get(h)
                if idx is not None and idx < len(r_vals):
                    v = r_vals[idx]
                    return v if v is not None else default
                return default

            fn = gv("Folder / File")
            inv_no = gv("Invoice No.")
            inv_dt = gv("Invoice Date")
            sup = gv("Supplier")
            gstin = gv("Supplier GSTIN")
            desc = gv("Description")
            qty = gv("Quantity")
            rate = gv("Rate (₹)")
            taxable = gv("Taxable Amount (₹)", 0.0)
            gst_rate = gv("GST Rate", 0.0)
            tax_type = gv("Tax Type", "")
            igst = gv("IGST (₹)", 0.0)
            cgst = gv("CGST (₹)", 0.0)
            sgst = gv("SGST (₹)", 0.0)
            inv_total = gv("Invoice Total (₹)", 0.0)
            tds_amt = gv("TDS Amount (₹)", 0.0)
            net_pay = gv("Net Payable (₹)", 0.0)
            irn = gv("IRN")

            tot_taxable += float(taxable) if isinstance(taxable, (int, float)) else 0.0
            tot_igst += float(igst) if isinstance(igst, (int, float)) else 0.0
            tot_cgst += float(cgst) if isinstance(cgst, (int, float)) else 0.0
            tot_sgst += float(sgst) if isinstance(sgst, (int, float)) else 0.0
            tot_invoice += float(inv_total) if isinstance(inv_total, (int, float)) else 0.0
            tot_tds += float(tds_amt) if isinstance(tds_amt, (int, float)) else 0.0
            tot_net += float(net_pay) if isinstance(net_pay, (int, float)) else 0.0

            if sup:
                unique_suppliers.add(str(sup).strip())

            if fn:
                registered_files.add(str(fn).strip().lower().replace('/', '\\'))
            if inv_no and sup:
                registered_invoices.append({
                    "sno": sno,
                    "inv_no_norm": normalize_str(inv_no),
                    "sup_norm": normalize_str(sup),
                    "gstin_norm": normalize_str(gstin) if gstin else "",
                    "inv_no": str(inv_no),
                    "supplier": str(sup),
                })
            if irn and str(irn).strip():
                registered_irns.add(str(irn).strip().lower())

            dt_str = ""
            if isinstance(inv_dt, datetime.datetime):
                dt_str = inv_dt.strftime("%d-%m-%Y")
            elif isinstance(inv_dt, datetime.date):
                dt_str = inv_dt.strftime("%d-%m-%Y")
            elif inv_dt:
                dt_str = str(inv_dt)

            rows.append({
                "sno": sno,
                "file": str(fn) if fn else "",
                "invoice_no": str(inv_no) if inv_no else "",
                "invoice_date": dt_str,
                "supplier": str(sup) if sup else "",
                "description": str(desc) if desc else "",
                "taxable": round(float(taxable), 2) if isinstance(taxable, (int, float)) else 0.0,
                "gst_rate": f"{round(float(gst_rate)*100)}%" if isinstance(gst_rate, (int, float)) else str(gst_rate),
                "tax_type": str(tax_type),
                "igst": round(float(igst), 2) if isinstance(igst, (int, float)) else 0.0,
                "cgst": round(float(cgst), 2) if isinstance(cgst, (int, float)) else 0.0,
                "sgst": round(float(sgst), 2) if isinstance(sgst, (int, float)) else 0.0,
                "total": round(float(inv_total), 2) if isinstance(inv_total, (int, float)) else 0.0,
                "tds": round(float(tds_amt), 2) if isinstance(tds_amt, (int, float)) else 0.0,
                "net": round(float(net_pay), 2) if isinstance(net_pay, (int, float)) else 0.0,
            })

        state_data = {
            "max_sno": max_sno,
            "total_rows": len(rows),
            "rows": rows,
            "registered_files": list(registered_files),
            "registered_invoices": registered_invoices,
            "registered_irns": list(registered_irns),
            "kpis": {
                "total_invoices": max_sno,
                "total_lines": len(rows),
                "total_suppliers": len(unique_suppliers),
                "total_taxable": round(tot_taxable, 2),
                "total_igst": round(tot_igst, 2),
                "total_cgst": round(tot_cgst, 2),
                "total_sgst": round(tot_sgst, 2),
                "total_invoice": round(tot_invoice, 2),
                "total_tds": round(tot_tds, 2),
                "total_net": round(tot_net, 2),
            }
        }
        self._cached_state = state_data
        self._cached_mtime = current_mtime
        return state_data

    def check_duplicate(self, invoice_data, state=None):
        """Check duplicate against register rules in SKILL.md Step 4."""
        if state is None:
            state = self.get_register_state()

        inv_no = invoice_data.get("invoice_no")
        supplier = invoice_data.get("supplier")
        supplier_gstin = invoice_data.get("supplier_gstin")
        irn = invoice_data.get("irn")

        # 1. Match IRN if present
        if irn and str(irn).strip():
            irn_clean = str(irn).strip().lower()
            if irn_clean in state.get("registered_irns", []):
                return True, f"Duplicate IRN: {irn[:16]}... already registered", None

        # 2. Match normalized Invoice No + Supplier or Supplier GSTIN
        if inv_no:
            inv_norm = normalize_str(inv_no)
            sup_norm = normalize_str(supplier)
            gstin_norm = normalize_str(supplier_gstin)

            for item in state.get("registered_invoices", []):
                if item["inv_no_norm"] == inv_norm:
                    if (sup_norm and sup_norm in item["sup_norm"]) or \
                       (item["sup_norm"] and item["sup_norm"] in sup_norm) or \
                       (gstin_norm and item["gstin_norm"] and gstin_norm == item["gstin_norm"]):
                        return True, f"Duplicate of Invoice '{inv_no}' from {supplier} (already at S.No. {item['sno']})", item["sno"]

        return False, "Not duplicate", None

    def add_invoices(self, invoices):
        """
        Add new invoices directly into 1. Expense_Invoice_Register.xlsx.
        Updates sheet2.xml directly inside the zip archive to preserve all charts and drawings 100%!
        """
        if not invoices:
            return {"status": "success", "message": "No invoices to add", "added_count": 0}

        # Backup register first
        backup_dir = EXPENSES_DIR / "_backup"
        backup_dir.mkdir(exist_ok=True)
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_file = backup_dir / f"Expense_Invoice_Register_{timestamp}.xlsx"
        shutil.copy2(self.register_path, backup_file)

        # Get state
        state = self.get_register_state()
        curr_sno = state.get("max_sno", 0)

        # Build list of new lines to insert
        lines_to_insert = []
        new_suppliers = set()

        for inv in invoices:
            is_dup, reason, sno = self.check_duplicate(inv, state)
            if is_dup:
                print(f"Skipping duplicate: {reason}")
                continue

            curr_sno += 1
            sup_name = inv.get("supplier", "Unknown")
            new_suppliers.add(sup_name)

            for itm in inv.get("items", []):
                lines_to_insert.append({
                    "sno": curr_sno,
                    "file_path": inv.get("file_path", ""),
                    "invoice_no": inv.get("invoice_no", ""),
                    "invoice_date": inv.get("invoice_date"),
                    "supplier": sup_name,
                    "supplier_gstin": inv.get("supplier_gstin"),
                    "item_code": itm.get("item_code"),
                    "description": itm.get("description", ""),
                    "hsn": itm.get("hsn"),
                    "boxes": itm.get("boxes"),
                    "quantity": itm.get("quantity", 1),
                    "unit": itm.get("unit", "NOS"),
                    "rate": itm.get("rate", 0),
                    "discount": itm.get("discount", 0),
                    "taxable_amount": itm.get("taxable_amount", 0),
                    "gst_rate": itm.get("gst_rate", 0.18),
                    "tax_type": itm.get("tax_type", "IGST"),
                    "round_off": itm.get("round_off", 0),
                    "location_code": itm.get("location_code"),
                    "vendor_gl_code": itm.get("vendor_gl_code"),
                    "exps_gl_code": itm.get("exps_gl_code"),
                    "supplier_state": inv.get("supplier_state"),
                    "buyer_gstin": inv.get("buyer_gstin"),
                    "place_of_supply": inv.get("place_of_supply"),
                    "buyer_po_no": inv.get("buyer_po_no"),
                    "po_date": inv.get("po_date"),
                    "delivery_note": inv.get("delivery_note"),
                    "reference_no": inv.get("reference_no"),
                    "irn": inv.get("irn"),
                    "ack_no": inv.get("ack_no"),
                    "ack_date": inv.get("ack_date"),
                    "eway_bill_no": inv.get("eway_bill_no"),
                    "transporter": inv.get("transporter"),
                    "lr_no": inv.get("lr_no"),
                    "vehicle_no": inv.get("vehicle_no"),
                    "rcm_note": itm.get("rcm_note"),
                    "tds_section": itm.get("tds_section"),
                    "tds_rate": itm.get("tds_rate", 0.001),
                })

        if not lines_to_insert:
            return {"status": "skipped", "message": "All invoices were duplicates or empty", "added_count": 0}

        # Modify sheet2.xml
        self._patch_sheet2_xml(lines_to_insert)

        # Update 3.Vendor_GL_Codes.xlsx with any new suppliers
        self._update_vendor_codes(new_suppliers)

        # Rebuild dashboard HTML
        self._refresh_dashboard()

        return {
            "status": "success",
            "message": f"Successfully added {len(lines_to_insert)} lines ({len(invoices)} invoices) to register",
            "added_count": len(lines_to_insert),
            "new_sno_max": curr_sno,
        }

    def _patch_sheet2_xml(self, lines_to_insert):
        """Directly insert rows into xl/worksheets/sheet2.xml preserving drawings/charts."""
        temp_zip = EXPENSES_DIR / "temp_patch_reg.xlsx"
        num_new = len(lines_to_insert)

        with zipfile.ZipFile(self.register_path, 'r') as zin, zipfile.ZipFile(temp_zip, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == 'xl/worksheets/sheet2.xml':
                    raw_xml = zin.read(item.filename)
                    modified_xml = self._insert_rows_into_sheet_xml(raw_xml, lines_to_insert)
                    zout.writestr(item, modified_xml)
                else:
                    zout.writestr(item, zin.read(item.filename))

        # Replace register with temp
        shutil.move(temp_zip, self.register_path)

    def _insert_rows_into_sheet_xml(self, raw_xml, lines_to_insert):
        root = ET.fromstring(raw_xml)
        ns = {'ns': NS_MAIN}
        sheet_data = root.find('ns:sheetData', ns)

        # Locate template row (e.g. 121), the blank row (122), and total row (123)
        rows_by_num = {}
        for r in sheet_data.findall('ns:row', ns):
            r_num = int(r.attrib.get('r', 0))
            rows_by_num[r_num] = r

        max_existing_data_row = 121
        for r_num in sorted(rows_by_num.keys()):
            if r_num < 122:
                max_existing_data_row = r_num

        total_row_num = 123
        for r_num in sorted(rows_by_num.keys()):
            if r_num > max_existing_data_row:
                # Find row containing TOTAL
                r_elem = rows_by_num[r_num]
                for c in r_elem.findall('ns:c', ns):
                    f = c.find('ns:f', ns)
                    if f is not None and 'SUM' in (f.text or ''):
                        total_row_num = r_num
                        break
                if total_row_num == r_num:
                    break

        blank_row_num = total_row_num - 1
        num_new = len(lines_to_insert)

        # Shift all rows >= blank_row_num by num_new
        for r_num in sorted(rows_by_num.keys(), reverse=True):
            if r_num >= blank_row_num:
                row_elem = rows_by_num[r_num]
                new_r_num = r_num + num_new
                row_elem.attrib['r'] = str(new_r_num)
                # Update cell coordinates
                for c in row_elem.findall('ns:c', ns):
                    coord = c.attrib.get('r', '')
                    col_letters = re.match(r'([A-Za-z]+)', coord).group(1)
                    c.attrib['r'] = f"{col_letters}{new_r_num}"
                    # If this is the TOTAL row, update SUM formulas
                    f = c.find('ns:f', ns)
                    if f is not None and f.text and 'SUM(' in f.text:
                        # e.g. SUM(Q5:Q122) -> SUM(Q5:Q{max_existing_data_row + num_new})
                        m = re.match(r'SUM\(([A-Za-z]+)\d+:([A-Za-z]+)\d+\)', f.text)
                        if m:
                            c1, c2 = m.group(1), m.group(2)
                            f.text = f"SUM({c1}5:{c2}{max_existing_data_row + num_new})"

        # Generate new rows starting from blank_row_num (which was 122)
        template_row = rows_by_num.get(max_existing_data_row)
        styles = self._extract_styles_from_row(template_row)

        new_rows_elements = []
        for idx, line in enumerate(lines_to_insert):
            target_row_idx = max_existing_data_row + 1 + idx
            r_elem = self._create_row_element(target_row_idx, line, styles)
            new_rows_elements.append(r_elem)

        # Insert new rows into sheetData in proper order
        all_rows = []
        for r in sheet_data.findall('ns:row', ns):
            all_rows.append(r)
            sheet_data.remove(r)

        all_rows.extend(new_rows_elements)
        all_rows.sort(key=lambda r: int(r.attrib['r']))
        for r in all_rows:
            sheet_data.append(r)

        # Update autoFilter ref if present
        auto_filter = root.find('ns:autoFilter', ns)
        if auto_filter is not None:
            auto_filter.attrib['ref'] = f"A4:AT{max_existing_data_row + num_new}"

        return ET.tostring(root, encoding='utf-8', xml_declaration=True)

    def _extract_styles_from_row(self, row_elem):
        styles = {}
        if row_elem is not None:
            ns = {'ns': NS_MAIN}
            for c in row_elem.findall('ns:c', ns):
                coord = c.attrib.get('r', '')
                m = re.match(r'([A-Za-z]+)', coord)
                if m:
                    col_ltr = m.group(1)
                    styles[col_ltr] = c.attrib.get('s', '47')
        return styles

    def _create_row_element(self, row_num, line, styles):
        ns_uri = NS_MAIN
        row_elem = ET.Element(f'{{{ns_uri}}}row', {
            'r': str(row_num),
            'customFormat': 'false',
            'ht': '14.25',
            'hidden': 'false',
            'customHeight': 'true',
            'outlineLevel': '0',
            'collapsed': 'false',
        })

        # Date to excel serial number
        inv_dt = line.get("invoice_date")
        if isinstance(inv_dt, datetime.date):
            epoch = datetime.date(1899, 12, 30)
            date_serial = (inv_dt - epoch).days
        else:
            date_serial = 46295

        col_defs = [
            ("A", line.get("sno"), "n", None),
            ("B", line.get("file_path"), "inlineStr", None),
            ("C", line.get("invoice_no"), "inlineStr", None),
            ("D", date_serial, "n", None),
            ("E", None, "n", f"YEAR(D{row_num})"),
            ("F", None, "str", f'TEXT(D{row_num},"mmm")'),
            ("G", line.get("supplier"), "inlineStr", None),
            ("H", line.get("supplier_gstin"), "inlineStr", None),
            ("I", line.get("item_code"), "inlineStr", None),
            ("J", line.get("description"), "inlineStr", None),
            ("K", line.get("hsn"), "inlineStr", None),
            ("L", line.get("boxes"), "n", None),
            ("M", line.get("quantity", 1), "n", None),
            ("N", line.get("unit", "NOS"), "inlineStr", None),
            ("O", line.get("rate", 0), "n", None),
            ("P", line.get("discount", 0), "n", None),
            ("Q", None, "n", f"M{row_num}*O{row_num}-P{row_num}"),
            ("R", line.get("gst_rate", 0.18), "n", None),
            ("S", line.get("tax_type", "IGST"), "inlineStr", None),
            ("T", None, "n", f'IF(S{row_num}="IGST",ROUND(Q{row_num}*R{row_num},2),0)'),
            ("U", None, "n", f'IF(S{row_num}="CGST+SGST",ROUND(Q{row_num}*R{row_num}/2,2),0)'),
            ("V", None, "n", f"U{row_num}"),
            ("W", line.get("round_off", 0), "n", None),
            ("X", None, "n", f"Q{row_num}+T{row_num}+U{row_num}+V{row_num}+W{row_num}"),
            ("Y", None, "n", f"ROUND(Q{row_num}*AT{row_num},0)"),
            ("Z", None, "n", f"X{row_num}-Y{row_num}"),
            ("AA", line.get("location_code"), "inlineStr", None),
            ("AB", line.get("vendor_gl_code"), "inlineStr", None),
            ("AC", line.get("exps_gl_code"), "inlineStr", None),
            ("AD", line.get("supplier_state"), "inlineStr", None),
            ("AE", line.get("buyer_gstin"), "inlineStr", None),
            ("AF", line.get("place_of_supply"), "inlineStr", None),
            ("AG", line.get("buyer_po_no"), "inlineStr", None),
            ("AH", line.get("po_date"), "inlineStr", None),
            ("AI", line.get("delivery_note"), "inlineStr", None),
            ("AJ", line.get("reference_no"), "inlineStr", None),
            ("AK", line.get("irn"), "inlineStr", None),
            ("AL", line.get("ack_no"), "inlineStr", None),
            ("AM", line.get("ack_date"), "inlineStr", None),
            ("AN", line.get("eway_bill_no"), "inlineStr", None),
            ("AO", line.get("transporter"), "inlineStr", None),
            ("AP", line.get("lr_no"), "inlineStr", None),
            ("AQ", line.get("vehicle_no"), "inlineStr", None),
            ("AR", line.get("rcm_note"), "inlineStr", None),
            ("AS", line.get("tds_section"), "inlineStr", None),
            ("AT", line.get("tds_rate", 0.001), "n", None),
        ]

        for col_ltr, val, val_type, formula in col_defs:
            c_elem = ET.SubElement(row_elem, f'{{{ns_uri}}}c', {
                'r': f"{col_ltr}{row_num}",
                's': styles.get(col_ltr, '47'),
            })

            if formula:
                f_elem = ET.SubElement(c_elem, f'{{{ns_uri}}}f', {'aca': 'false'})
                f_elem.text = formula
                c_elem.attrib['t'] = val_type
            elif val is not None and str(val).strip() != '':
                if val_type == 'inlineStr':
                    c_elem.attrib['t'] = 'inlineStr'
                    is_elem = ET.SubElement(c_elem, f'{{{ns_uri}}}is')
                    t_elem = ET.SubElement(is_elem, f'{{{ns_uri}}}t')
                    t_elem.text = str(val)
                elif val_type == 'n':
                    c_elem.attrib['t'] = 'n'
                    v_elem = ET.SubElement(c_elem, f'{{{ns_uri}}}v')
                    v_elem.text = str(val)

        return row_elem

    def _update_vendor_codes(self, supplier_names):
        """Append any new suppliers to 3.Vendor_GL_Codes.xlsx with codes blank."""
        if not self.codes_path.exists() or not supplier_names:
            return

        try:
            wb = openpyxl.load_workbook(self.codes_path)
            ws = wb.active
            existing = set()
            for r in range(4, ws.max_row + 1):
                val = ws.cell(row=r, column=1).value
                if val:
                    existing.add(str(val).strip().lower())

            added = False
            for sup in supplier_names:
                if sup and sup.strip().lower() not in existing:
                    ws.append([sup.strip(), None, None, None])
                    existing.add(sup.strip().lower())
                    added = True

            if added:
                wb.save(self.codes_path)
        except Exception as e:
            print(f"Warning: could not update vendor codes: {e}")

    def _refresh_dashboard(self):
        """Execute build_dashboard.py to update Expense_Dashboard.html."""
        if BUILD_DASHBOARD_SCRIPT.exists() and DASHBOARD_HTML_PATH.exists():
            try:
                cmd = [
                    str(Path(sys.executable)),
                    str(BUILD_DASHBOARD_SCRIPT),
                    str(self.register_path),
                    str(self.codes_path),
                    str(DASHBOARD_HTML_PATH),
                    str(DASHBOARD_HTML_PATH),
                ]
                subprocess.run(cmd, check=True, capture_output=True)
            except Exception as e:
                print(f"Dashboard build warning: {e}")
