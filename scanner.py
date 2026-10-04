import os
import sys
from pathlib import Path
from config import EXPENSES_DIR, IGNORED_FILES, IGNORED_DIRS
from invoice_extractor import InvoiceExtractor
from excel_manager import ExcelManager

class InvoiceScanner:
    def __init__(self):
        self.extractor = InvoiceExtractor()
        self.excel_mgr = ExcelManager()

    def find_unprocessed_pdfs(self):
        """Fast scan Expenses folder (depth <= 2) and return list of unrecorded PDFs."""
        state = self.excel_mgr.get_register_state()
        reg_files = set(state.get("registered_files", []))

        unprocessed = []
        skip_folders = {'_tools', '_docs', '_backup', '_web_uploads', '.venv', '__pycache__'}
        
        try:
            for entry in os.scandir(EXPENSES_DIR):
                if entry.is_dir() and entry.name not in skip_folders and not entry.name.startswith('.'):
                    try:
                        for sub in os.scandir(entry.path):
                            if sub.is_file() and sub.name.lower().endswith('.pdf') and sub.name.lower() not in IGNORED_FILES:
                                full_p = Path(sub.path)
                                rel_p = full_p.relative_to(EXPENSES_DIR)
                                norm_path = str(rel_p).strip().lower().replace('/', '\\')
                                if norm_path not in reg_files:
                                    unprocessed.append(full_p)
                            elif sub.is_dir() and not sub.name.startswith('.'):
                                for sub2 in os.scandir(sub.path):
                                    if sub2.is_file() and sub2.name.lower().endswith('.pdf') and sub2.name.lower() not in IGNORED_FILES:
                                        full_p = Path(sub2.path)
                                        rel_p = full_p.relative_to(EXPENSES_DIR)
                                        norm_path = str(rel_p).strip().lower().replace('/', '\\')
                                        if norm_path not in reg_files:
                                            unprocessed.append(full_p)
                    except PermissionError:
                        pass
                elif entry.is_file() and entry.name.lower().endswith('.pdf') and entry.name.lower() not in IGNORED_FILES:
                    full_p = Path(entry.path)
                    rel_p = full_p.relative_to(EXPENSES_DIR)
                    norm_path = str(rel_p).strip().lower().replace('/', '\\')
                    if norm_path not in reg_files:
                        unprocessed.append(full_p)
        except Exception as e:
            print(f"Scan error: {e}")

        return unprocessed, state

    def scan_and_process(self, dry_run=False):
        """Scan Expenses folder, extract unrecorded invoices, and add them to register."""
        unprocessed_pdfs, state = self.find_unprocessed_pdfs()
        
        results = {
            "total_found": len(unprocessed_pdfs),
            "processed": [],
            "skipped_duplicates": [],
            "errors": [],
            "dry_run": dry_run,
        }

        if not unprocessed_pdfs:
            results["message"] = "All invoices in Expenses directory are already registered. No new files to process."
            return results

        invoices_to_add = []
        for pdf in unprocessed_pdfs:
            try:
                inv_data = self.extractor.extract_from_pdf(pdf)
                is_dup, reason, sno = self.excel_mgr.check_duplicate(inv_data, state)
                if is_dup:
                    results["skipped_duplicates"].append({
                        "file": str(pdf.name),
                        "reason": reason,
                        "sno": sno,
                    })
                else:
                    invoices_to_add.append(inv_data)
                    results["processed"].append({
                        "file": str(pdf.name),
                        "invoice_no": inv_data.get("invoice_no"),
                        "supplier": inv_data.get("supplier"),
                        "items_count": len(inv_data.get("items", [])),
                    })
            except Exception as e:
                results["errors"].append({
                    "file": str(pdf.name),
                    "error": str(e),
                })

        if not dry_run and invoices_to_add:
            commit_res = self.excel_mgr.add_invoices(invoices_to_add)
            results["commit"] = commit_res

        return results

    def process_single_pdf(self, pdf_path, dry_run=False):
        """Process an uploaded or selected invoice PDF file."""
        pdf_path = Path(pdf_path)
        if not pdf_path.exists():
            return {"status": "error", "message": f"File not found: {pdf_path}"}

        inv_data = self.extractor.extract_from_pdf(pdf_path)
        state = self.excel_mgr.get_register_state()
        is_dup, reason, sno = self.excel_mgr.check_duplicate(inv_data, state)

        res = {
            "invoice": inv_data,
            "is_duplicate": is_dup,
            "duplicate_reason": reason,
            "duplicate_sno": sno,
        }

        if not is_dup and not dry_run:
            commit_res = self.excel_mgr.add_invoices([inv_data])
            res["commit"] = commit_res

        return res
