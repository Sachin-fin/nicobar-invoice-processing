import os
import sys
import re
import shutil
import datetime
from pathlib import Path
import pymupdf
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')

from config import EXPENSES_DIR, FALLBACK_DIR, REGISTER_PATH

def parse_date(txt):
    if not txt: return None
    patterns = [
        r'(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})',
        r'(\d{1,2})(?:st|nd|rd|th)?\s+([A-Za-z]+)\s+(\d{2,4})',
        r'(\d{4})[./-](\d{1,2})[./-](\d{1,2})',
    ]
    for pat in patterns:
        m = re.search(pat, txt)
        if m:
            g = m.groups()
            if len(g) == 3 and g[0].isdigit() and g[1].isdigit():
                d, mth, y = int(g[0]), int(g[1]), int(g[2])
                if y < 100: y += 2000
                if 1 <= mth <= 12 and 1 <= d <= 31:
                    try: return datetime.date(y, mth, d)
                    except ValueError: pass
            elif len(g) == 3 and len(g[0]) == 4 and g[0].isdigit():
                y, mth, d = int(g[0]), int(g[1]), int(g[2])
                if 1 <= mth <= 12 and 1 <= d <= 31:
                    try: return datetime.date(y, mth, d)
                    except ValueError: pass
            elif len(g) == 3:
                for fmt in ("%d %B %Y", "%d %b %Y"):
                    try: return datetime.datetime.strptime(f"{g[0]} {g[1]} {g[2]}", fmt).date()
                    except Exception: pass
    return None

def clean_name(s):
    return re.sub(r'[\/\\*?:"<>|]', '-', str(s)).strip()

def main():
    base_dirs = [EXPENSES_DIR]
    if FALLBACK_DIR.exists() and FALLBACK_DIR != EXPENSES_DIR:
        base_dirs.append(FALLBACK_DIR)

    # Specific known canonical cleanups
    for b_dir in base_dirs:
        # General expenses cleanup
        gen_dir = b_dir / "General Expenses"
        if gen_dir.exists():
            for f in list(gen_dir.iterdir()):
                f.unlink(missing_ok=True)
            try: gen_dir.rmdir()
            except Exception: pass

        # MKR APP non-invoice statement cleanup
        mkr_file = b_dir / "MKR APP Private Limited" / "GL_26-27 04.10.26.pdf"
        if mkr_file.exists():
            mkr_file.unlink(missing_ok=True)

        # Gowri Parvathy duplicate
        gowri_dup = b_dir / "Gowri Parvathy" / "Nicobar invoice 04.10.26.pdf"
        if gowri_dup.exists():
            gowri_dup.unlink(missing_ok=True)

        # Curefoods duplicate
        cf_dup = b_dir / "Curefoods India Limited" / "BIN-2026- 02.10.26.pdf"
        if cf_dup.exists():
            cf_dup.unlink(missing_ok=True)

        # Cushman & Wakefield duplicates
        cw_dir = b_dir / "Cushman & Wakefield India Private Limited"
        if cw_dir.exists():
            for unwanted in [
                "APIPL_CUS_INV-2026-09-000008616 04.10.26.pdf",
                "APIPL_CUS_INV-2026-09-000008616_557 04.10.26.pdf",
                "MH-2026-09-2376 26.09.76.pdf",
                "MH-2026-09-2376_61 26.09.76.pdf",
            ]:
                (cw_dir / unwanted).unlink(missing_ok=True)

        # Mahindra duplicates
        mm_dir = b_dir / "Mahindra & Mahindra Financial Services Limited"
        if mm_dir.exists():
            for unwanted in [
                "NICOBAR DESIGN PRIVATE LIMITED_Oct 26_FMS Rental_Lease Rental_IND-450459 02.10.26.pdf",
                "NICOBAR DESIGN PRIVATE LIMITED_Oct 26_Reimbursement_231742 02.10.26.pdf",
            ]:
                (mm_dir / unwanted).unlink(missing_ok=True)

        # Niva's duplicates
        niva_dir = b_dir / "Niva's Designs"
        if niva_dir.exists():
            for unwanted in [
                "CN_Nicobar_Delhi_29-08-26 04.10.26.pdf",
                "ND-26-27-043 04.10.26.pdf",
                "ND-26-27-043A 29.08.26.pdf",
            ]:
                (niva_dir / unwanted).unlink(missing_ok=True)

        # IPinfo duplicate
        ip_dir = b_dir / "IPinfo Inc"
        if ip_dir.exists():
            (ip_dir / "Invoice-3MNDSIX0-0003 (1) 04.10.26.pdf").unlink(missing_ok=True)

        # Prezotech duplicate
        prezo_dir = b_dir / "Prezotech Solutions Pvt Ltd"
        if prezo_dir.exists():
            (prezo_dir / "INV-26-27-0901_61 30.06.26.pdf").unlink(missing_ok=True)

        # A & S duplicate
        as_dir = b_dir / "A & S Creations"
        if as_dir.exists():
            (as_dir / "Invoice 065 - Nicobar - September 26 04.10.26.pdf").unlink(missing_ok=True)

        # Ambience Developers duplicates
        amb_dev = b_dir / "Ambience Developers & Infrastructure Pvt Ltd"
        if amb_dev.exists():
            (amb_dev / "EAD-09-32-26-27 04.10.26.pdf").unlink(missing_ok=True)
            (amb_dev / "SFE-09-32-26-27 04.10.26.pdf").unlink(missing_ok=True)

        # Ambience Facilities duplicates
        amb_fac = b_dir / "Ambience Facilities Management Private Limited"
        if amb_fac.exists():
            (amb_fac / "CAF-10-34-26-27 04.10.26.pdf").unlink(missing_ok=True)

    print("Cleanup and organization complete.")

if __name__ == "__main__":
    main()
