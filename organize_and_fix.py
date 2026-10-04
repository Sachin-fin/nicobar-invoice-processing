import os
import sys
import re
import shutil
import datetime
from pathlib import Path
import pymupdf
import openpyxl

from config import EXPENSES_DIR, REGISTER_PATH

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
    dirs_to_check = [
        Path(r"D:\Claude_Sales\Expenses"),
        Path(r"D:\Google _invoice Processing"),
    ]
    
    # Specific known relocations for the downloaded batch
    relocations = [
        ("2026-27-224_61 01.10.26.pdf", "XTO10X Technologies Private Limited", "2026-27-224 01.10.26.pdf"),
        ("IND-450459 01.10.26.pdf", "Mahindra & Mahindra Financial Services Limited", "IND-450459 01.10.26.pdf"),
        ("NICOBAR DESIGN PRIVATE LIMITED_Oct 26_Reimbursement_231742_61 02.10.26.pdf", "Mahindra & Mahindra Financial Services Limited", "231742 01.10.26.pdf"),
        ("Invoice 065 - Nicobar - September 26_61 04.10.26.pdf", "A & S Creations", "065-2026-2027 26.09.26.pdf"),
        ("ND-26-27-043_61 04.10.26.pdf", "Niva's Designs", "ND-26-27-043 29.07.26.pdf"),
        ("ND-26-27-043A_61 29.07.26.pdf", "Niva's Designs", "ND-26-27-043A 29.07.26.pdf"),
        ("CN_Nicobar_Delhi_29-08-26_61 04.10.26.pdf", "Niva's Designs", "CN-Nicobar-Delhi-29-08-26 29.08.26.pdf"),
        ("CUREFOODS 02.10.26.pdf", "Curefoods India Limited", "BIN-2026-28-05 02.10.26.pdf"),
    ]

    for base_dir in dirs_to_check:
        temp_dir = base_dir / "_temp_gmail"
        if not temp_dir.exists():
            continue
        print(f"Checking {temp_dir}...")
        for src_name, vendor_folder, target_name in relocations:
            src_path = temp_dir / src_name
            if src_path.exists():
                dest_dir = base_dir / vendor_folder
                dest_dir.mkdir(parents=True, exist_ok=True)
                dest_path = dest_dir / target_name
                shutil.move(src_path, dest_path)
                print(f"Relocated: [{vendor_folder}] -> {target_name}")

        remaining = list(temp_dir.iterdir())
        if not remaining:
            temp_dir.rmdir()
            print(f"Removed empty {temp_dir}")
        else:
            print(f"Remaining in {temp_dir}: {[f.name for f in remaining]}")

    for src_name, vendor_folder, target_name in relocations:
        src_path = temp_dir / src_name
        if src_path.exists():
            dest_dir = EXPENSES_DIR / vendor_folder
            dest_dir.mkdir(parents=True, exist_ok=True)
            dest_path = dest_dir / target_name
            shutil.move(src_path, dest_path)
            print(f"Relocated: [{vendor_folder}] -> {target_name}")

    # Remove temp folder if empty
    if temp_dir.exists():
        remaining = list(temp_dir.iterdir())
        if not remaining:
            temp_dir.rmdir()
            print("Removed empty _temp_gmail directory.")
        else:
            print(f"Remaining in _temp_gmail: {[f.name for f in remaining]}")

if __name__ == "__main__":
    main()
