import os
import sys
import re
import json
import time
import email
from email.header import decode_header
import imaplib
import datetime
import shutil
from pathlib import Path
import pymupdf
from dotenv import load_dotenv

from config import (
    EXPENSES_DIR,
    FALLBACK_DIR,
    REGISTER_PATH,
    GMAIL_STATE_PATH,
    NICOBAR_GSTINS,
)
from invoice_extractor import InvoiceExtractor
from excel_manager import ExcelManager

load_dotenv()

def sanitize_filename(name):
    return re.sub(r'[\\/*?:"<>|]', '-', str(name)).strip()

class GmailIntake:
    def __init__(self):
        self.mailbox = os.getenv("GMAIL_USER", "invoice@nicobar.com")
        self.app_password = os.getenv("GMAIL_APP_PASSWORD", "")
        self.imap_server = os.getenv("IMAP_SERVER", "imap.gmail.com")
        self.state_file = GMAIL_STATE_PATH if GMAIL_STATE_PATH.exists() else (Path(__file__).parent / "_gmail_scan_state.json")
        self.extractor = InvoiceExtractor()
        self.excel_mgr = ExcelManager()

    def get_state(self):
        if self.state_file.exists():
            try:
                with open(self.state_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                print(f"Error reading gmail state: {e}")
        return {
            "mailbox": self.mailbox,
            "last_scan_epoch": int(time.time()) - (7 * 86400),
            "last_scan_utc": "",
            "last_scan_ist": "",
            "pending_messages": [],
            "ignored_files": ["Cowork_Invoice_to_Dynamics365_Process.pdf"]
        }

    def save_state(self, state):
        try:
            with open(self.state_file, 'w', encoding='utf-8') as f:
                json.dump(state, f, indent=2)
        except Exception as e:
            print(f"Error saving gmail state: {e}")

    def decode_str(self, header_value):
        if not header_value:
            return ""
        decoded_parts = decode_header(header_value)
        result = []
        for text, enc in decoded_parts:
            if isinstance(text, bytes):
                try:
                    result.append(text.decode(enc or 'utf-8', errors='ignore'))
                except Exception:
                    result.append(text.decode('latin1', errors='ignore'))
            else:
                result.append(str(text))
        return "".join(result).strip()

    def connect(self):
        if not self.app_password:
            raise ValueError(
                "Missing GMAIL_APP_PASSWORD. Please configure your 16-character Google App Password in .env "
                "or via the Settings panel in the Web UI."
            )
        mail = imaplib.IMAP4_SSL(self.imap_server, 993)
        mail.login(self.mailbox, self.app_password)
        return mail

    def scan_and_auto_process(self, dry_run=False, max_emails=50):
        """
        Connect to invoice@nicobar.com, fetch new invoice attachments,
        file them into supplier folders, and auto-add to the Excel register.
        """
        state = self.get_state()
        last_epoch = state.get("last_scan_epoch", 0)

        # Convert epoch to IMAP date (e.g. 01-Oct-2026)
        if last_epoch > 0:
            since_dt = datetime.datetime.fromtimestamp(last_epoch)
            since_str = since_dt.strftime("%d-%b-%Y")
            search_query = f'SINCE {since_str}'
        else:
            search_query = 'ALL'

        print(f"Connecting to {self.mailbox} on {self.imap_server}...")
        mail = self.connect()
        mail.select("INBOX")

        print(f"Searching messages with query: {search_query}")
        status, data = mail.search(None, search_query)
        if status != "OK" or not data or not data[0]:
            mail.logout()
            return {
                "status": "success",
                "message": "No new emails found since last watermark.",
                "downloaded_invoices": [],
                "added_to_register": 0,
            }

        msg_ids = data[0].split()
        print(f"Found {len(msg_ids)} candidate emails. Processing up to {max_emails}...")
        msg_ids = msg_ids[-max_emails:] # Most recent first

        downloaded_invoices = []
        skipped_non_nicobar = []
        errors = []
        highest_epoch_seen = last_epoch

        temp_dir = EXPENSES_DIR / "_temp_gmail"
        temp_dir.mkdir(parents=True, exist_ok=True)

        for mid in reversed(msg_ids):
            try:
                res, msg_data = mail.fetch(mid, '(RFC822)')
                if res != 'OK':
                    continue

                raw_email = msg_data[0][1]
                msg = email.message_from_bytes(raw_email)

                subject = self.decode_str(msg.get("Subject", ""))
                sender = self.decode_str(msg.get("From", ""))
                date_header = msg.get("Date", "")

                # Parse date to epoch
                msg_dt = email.utils.parsedate_to_datetime(date_header) if date_header else datetime.datetime.now()
                msg_epoch = int(msg_dt.timestamp())
                if msg_epoch > highest_epoch_seen:
                    highest_epoch_seen = msg_epoch

                # Check attachments
                for part in msg.walk():
                    if part.get_content_maintype() == 'multipart':
                        continue
                    if part.get('Content-Disposition') is None:
                        continue

                    filename = part.get_filename()
                    if not filename:
                        continue
                    filename = self.decode_str(filename)
                    fn_lower = filename.lower()

                    # Skip email inline signatures, tracking icons, and statements
                    if fn_lower.endswith(('.dat', '.bin')):
                        continue
                    if 'image00' in fn_lower or fn_lower.startswith('~wrd') or fn_lower == 'gl_26-27.pdf':
                        continue

                    if fn_lower.endswith(('.pdf', '.png', '.jpg', '.jpeg')):
                        payload = part.get_payload(decode=True)
                        if not payload:
                            continue

                        # Skip tiny images (email icons / logos)
                        if fn_lower.endswith(('.png', '.jpg', '.jpeg')) and len(payload) < 80 * 1024:
                            continue

                        # Check file size (e.g. > 15MB)
                        if len(payload) > 15 * 1024 * 1024:
                            state.setdefault("pending_messages", []).append({
                                "id": mid.decode('utf-8', errors='ignore'),
                                "subject": subject,
                                "sender": sender,
                                "size_bytes": len(payload),
                                "reason": "Attachment too large (>15MB) - saved to pending",
                            })
                            continue

                        temp_file = temp_dir / filename
                        with open(temp_file, 'wb') as f:
                            f.write(payload)

                        # Extract details using InvoiceExtractor
                        try:
                            inv_data = self.extractor.extract_from_pdf(temp_file)
                            
                            # Step 0 Filter: Only invoices billed to Nicobar (PAN AAFCN6567R or Nicobar GSTIN)
                            buyer_g = inv_data.get("buyer_gstin", "")
                            is_nicobar = (
                                "AAFCN6567R" in buyer_g
                                or buyer_g in NICOBAR_GSTINS
                                or "nicobar" in inv_data.get("full_path", "").lower()
                                or "nicobar" in subject.lower()
                            )

                            if not is_nicobar:
                                skipped_non_nicobar.append({
                                    "filename": filename,
                                    "subject": subject,
                                    "reason": "Not billed to Nicobar (sales invoice or quotation)"
                                })
                                temp_file.unlink(missing_ok=True)
                                continue

                            # Step 1 Filing: Determine true Vendor Name & format <Vendor>\<Invoice No> <dd.mm.yy>.pdf
                            sup_detected = inv_data.get("supplier")
                            if not sup_detected or "_temp" in str(sup_detected).lower() or sup_detected == "Unknown_Supplier":
                                sup_gstin = inv_data.get("supplier_gstin")
                                if sup_gstin and sup_gstin.upper() in self.extractor.gstin_to_supplier:
                                    sup_detected = self.extractor.gstin_to_supplier[sup_gstin.upper()]
                                
                                # Inspect raw text and subject
                                if not sup_detected or "_temp" in str(sup_detected).lower():
                                    try:
                                        t_doc = pymupdf.open(temp_file)
                                        raw_t = " ".join(p.get_text() for p in t_doc).upper()
                                        t_doc.close()
                                    except Exception:
                                        raw_t = ""

                                    u_subj = subject.upper()
                                    u_fn = filename.upper()
                                    combo = f"{raw_t} {u_subj} {u_fn}"

                                    if 'PREZOTECH' in combo: sup_detected = 'Prezotech Solutions Pvt Ltd'
                                    elif 'GOWRI' in combo or 'PARVATHY' in combo: sup_detected = 'Gowri Parvathy'
                                    elif 'XTO10X' in combo: sup_detected = 'XTO10X Technologies Private Limited'
                                    elif 'NIVA' in combo or 'MACRAMEDESIGNS' in combo: sup_detected = "Niva's Designs"
                                    elif 'MAHINDRA' in combo or 'QUIKLYZ' in combo: sup_detected = 'Mahindra & Mahindra Financial Services Limited'
                                    elif 'IPINFO' in combo: sup_detected = 'IPinfo Inc'
                                    elif 'A & S CREATION' in combo or '07ABNFA6509L' in combo: sup_detected = 'A & S Creations'
                                    elif 'CUSHMAN' in combo or 'APIPL' in combo or '07AABCC3804L' in combo: sup_detected = 'Cushman & Wakefield India Private Limited'
                                    elif 'AMBIENCE FACILITIES' in combo or '06AABCA0558A' in combo: sup_detected = 'Ambience Facilities Management Private Limited'
                                    elif 'AMBIENCE DEVELOPERS' in combo or '06AABCA0557B' in combo or 'ELECT' in u_fn: sup_detected = 'Ambience Developers & Infrastructure Pvt Ltd'
                                    elif 'CUREFOODS' in combo or 'BIN-' in combo: sup_detected = 'Curefoods India Limited'
                                    elif 'CLAY CRAFT' in combo or '08AAACC1260F' in combo: sup_detected = 'Clay Craft India Limited'
                                    elif 'STITCH' in combo: sup_detected = 'Stitch-9'
                                    elif 'MKR APP' in combo: sup_detected = 'MKR APP Private Limited'
                                    elif 'BHARTI' in combo or 'AIRTEL' in combo: sup_detected = 'Bharti Airtel Limited'
                                    elif 'BLUE DART' in combo: sup_detected = 'Blue Dart Express Limited'
                                    elif 'DLF' in combo: sup_detected = 'DLF Home Developers Limited'
                                    elif 'DTDC' in combo: sup_detected = 'DTDC Express Limited'
                                    elif 'HANSA ELITE' in combo: sup_detected = 'Hansa Elite Luxe Interiors Pvt Ltd'
                                    elif 'SIMRAN' in combo: sup_detected = 'Simran'
                                    elif 'SURBHI SETHI' in combo: sup_detected = 'Surbhi Sethi'
                                    elif 'SHREYAS GAIKWAD' in combo: sup_detected = 'Shreyas Gaikwad'
                                    elif 'JATIN GULABRAI' in combo: sup_detected = 'Jatin Gulabrai Parekh'
                                    elif 'ANTHROPIC' in combo: sup_detected = 'Anthropic PBC'
                                    elif 'QUANG VINH' in combo: sup_detected = 'Quang Vinh Ceramic Co Ltd'
                                    elif 'TRION PROPERTIES' in combo: sup_detected = 'Trion Properties Private Limited'
                                    elif 'V DINE' in combo: sup_detected = 'V Dine Unit of VRSM Enterprises LLP'
                                    elif 'VODAFONE' in combo or 'VI BUSINESS' in combo: sup_detected = 'Vodafone Idea Limited'
                                    elif 'VRAJ COMMERCIAL' in combo: sup_detected = 'Vraj Commercial Private Limited'

                            if not sup_detected or "_temp" in str(sup_detected).lower():
                                # Check sender header
                                if sender:
                                    m_snd = re.match(r'^"?([^"<]+)"?', sender)
                                    if m_snd and not m_snd.group(1).lower().startswith(("invoice", "scan", "accounts")):
                                        sup_detected = m_snd.group(1).strip()
                                    else:
                                        sup_detected = "General Expenses"
                                else:
                                    sup_detected = "General Expenses"

                            sup_clean = sanitize_filename(sup_detected)
                            
                            # Determine clean Invoice No
                            inv_no = inv_data.get("invoice_no")
                            if not inv_no or inv_no in ("INV", "Tax Invoice", "Unknown"):
                                # Look for invoice no patterns in text
                                try:
                                    t_doc = pymupdf.open(temp_file)
                                    r_text = " ".join(p.get_text() for p in t_doc)
                                    t_doc.close()
                                    m_num = re.search(r'(?:Invoice\s*(?:No\.?|number|#)|Bill\s*No\.?|Document\s*No\.?)[.:\s]+([A-Za-z0-9\-_/]+)', r_text, re.IGNORECASE)
                                    if m_num:
                                        inv_no = m_num.group(1).strip()
                                except Exception:
                                    pass

                            if not inv_no or inv_no in ("INV", "Tax Invoice"):
                                # Use file stem
                                clean_stem = re.sub(r'\s*\d{2}\.\d{2}\.\d{2,4}.*$', '', Path(filename).stem)
                                inv_no = clean_stem if clean_stem else Path(filename).stem

                            inv_no_clean = sanitize_filename(inv_no).replace('/', '-')
                            
                            # Determine Invoice Date
                            inv_dt = inv_data.get("invoice_date")
                            if not inv_dt or not isinstance(inv_dt, (datetime.date, datetime.datetime)):
                                # Try parsing from email msg_dt
                                inv_dt = msg_dt.date() if msg_dt else datetime.date.today()

                            dt_part = inv_dt.strftime("%d.%m.%y")

                            # Vendor target folder under EXPENSES_DIR (e.g. D:\Google _invoice Processing\<Vendor Name>)
                            target_folder = EXPENSES_DIR / sup_clean
                            target_folder.mkdir(parents=True, exist_ok=True)

                            final_filename = f"{inv_no_clean} {dt_part}.pdf"
                            final_path = target_folder / final_filename

                            if final_path.exists() and final_path.stat().st_size != temp_file.stat().st_size:
                                final_filename = f"{inv_no_clean}_{int(time.time())%1000} {dt_part}.pdf"
                                final_path = target_folder / final_filename

                            temp_file.replace(final_path)

                            # Mirror copy to FALLBACK_DIR if different
                            if FALLBACK_DIR.exists() and FALLBACK_DIR != EXPENSES_DIR:
                                try:
                                    fb_folder = FALLBACK_DIR / sup_clean
                                    fb_folder.mkdir(parents=True, exist_ok=True)
                                    shutil.copy2(final_path, fb_folder / final_filename)
                                except Exception:
                                    pass

                            # Re-extract with final filed path
                            final_inv_data = self.extractor.extract_from_pdf(final_path)
                            final_inv_data["supplier"] = sup_detected
                            downloaded_invoices.append(final_inv_data)

                        except Exception as parse_err:
                            errors.append({"file": filename, "error": str(parse_err)})

            except Exception as e:
                errors.append({"mid": str(mid), "error": str(e)})

        mail.logout()

        # Clean up empty temp_dir
        try:
            if temp_dir.exists() and not os.listdir(temp_dir):
                temp_dir.rmdir()
        except Exception:
            pass

        # Step 2 & 3: Duplicate Check and Auto-Register Entry
        added_count = 0
        duplicate_skipped = []
        if downloaded_invoices and not dry_run:
            register_state = self.excel_mgr.get_register_state()
            valid_to_add = []
            for inv in downloaded_invoices:
                is_dup, reason, sno = self.excel_mgr.check_duplicate(inv, register_state)
                if is_dup:
                    duplicate_skipped.append({"invoice": inv.get("invoice_no"), "reason": reason, "sno": sno})
                else:
                    valid_to_add.append(inv)

            if valid_to_add:
                commit_res = self.excel_mgr.add_invoices(valid_to_add)
                added_count = commit_res.get("added_count", 0)

        # Update watermark in state
        now_dt = datetime.datetime.now(datetime.timezone.utc)
        ist_dt = now_dt + datetime.timedelta(hours=5, minutes=30)
        state["last_scan_epoch"] = highest_epoch_seen if highest_epoch_seen > last_epoch else int(time.time())
        state["last_scan_utc"] = now_dt.strftime("%Y-%m-%dT%H:%M:%SZ")
        state["last_scan_ist"] = ist_dt.strftime("%Y-%m-%d %H:%M:%S IST")
        self.save_state(state)

        return {
            "status": "success",
            "mailbox": self.mailbox,
            "emails_checked": len(msg_ids),
            "downloaded_count": len(downloaded_invoices),
            "added_to_register": added_count,
            "duplicate_skipped": duplicate_skipped,
            "skipped_non_nicobar": len(skipped_non_nicobar),
            "errors": errors,
            "last_scan_ist": state["last_scan_ist"],
        }
