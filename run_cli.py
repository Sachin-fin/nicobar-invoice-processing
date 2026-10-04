#!/usr/bin/env python3
"""
CLI runner for Nicobar Invoice Automation System.
Usage:
  python run_cli.py --status
  python run_cli.py --scan [--dry-run]
  python run_cli.py --file <path_to_invoice.pdf> [--dry-run]
"""

import sys
import argparse
from pathlib import Path

# Fix Windows console UTF-8 output
sys.stdout.reconfigure(encoding='utf-8')

from scanner import InvoiceScanner
from excel_manager import ExcelManager
from gmail_intake import GmailIntake

def print_kpis(kpis):
    print("=" * 60)
    print("  NICOBAR DESIGN - EXPENSE INVOICE REGISTER STATUS")
    print("=" * 60)
    print(f"  Total Invoices Recorded : {kpis.get('total_invoices', 0)}")
    print(f"  Total Line Items        : {kpis.get('total_lines', 0)}")
    print(f"  Active Suppliers        : {kpis.get('total_suppliers', 0)}")
    print(f"  Total Taxable Value     : ₹{kpis.get('total_taxable', 0):,.2f}")
    print(f"  Total IGST              : ₹{kpis.get('total_igst', 0):,.2f}")
    print(f"  Total CGST              : ₹{kpis.get('total_cgst', 0):,.2f}")
    print(f"  Total SGST              : ₹{kpis.get('total_sgst', 0):,.2f}")
    print(f"  Total Invoice Value     : ₹{kpis.get('total_invoice', 0):,.2f}")
    print(f"  Total TDS Deducted      : ₹{kpis.get('total_tds', 0):,.2f}")
    print(f"  Net Payable             : ₹{kpis.get('total_net', 0):,.2f}")
    print("=" * 60)

def main():
    parser = argparse.ArgumentParser(description="Nicobar Invoice Automation CLI")
    parser.add_argument("--status", action="store_true", help="Show register summary and KPIs")
    parser.add_argument("--scan", action="store_true", help="Scan Expenses directory for new invoices")
    parser.add_argument("--sync-gmail", action="store_true", help="Fetch new invoices from invoice@nicobar.com and auto-record")
    parser.add_argument("--file", type=str, help="Process a single invoice PDF")
    parser.add_argument("--dry-run", action="store_true", help="Perform extraction without modifying Excel")

    args = parser.parse_args()
    scanner = InvoiceScanner()
    excel_mgr = ExcelManager()
    gmail = GmailIntake()

    if args.sync_gmail:
        print("Connecting to invoice@nicobar.com to check new incoming invoice emails...")
        try:
            res = gmail.scan_and_auto_process(dry_run=args.dry_run)
            print(f"Status: {res.get('status')}")
            print(f"Emails checked: {res.get('emails_checked', 0)}")
            print(f"Invoices downloaded & filed: {res.get('downloaded_count', 0)}")
            print(f"Added to Excel register: {res.get('added_to_register', 0)}")
            if res.get('duplicate_skipped'):
                print(f"Skipped duplicates: {len(res['duplicate_skipped'])}")
            print(f"Last scan watermark: {res.get('last_scan_ist')}")
        except Exception as e:
            print(f"Gmail sync error: {e}")

    elif args.status or (not args.scan and not args.file and not args.sync_gmail):
        state = excel_mgr.get_register_state()
        print_kpis(state.get("kpis", {}))
        unprocessed, _ = scanner.find_unprocessed_pdfs()
        print(f"\nUnprocessed PDFs in Expenses directory: {len(unprocessed)}")
        for p in unprocessed[:10]:
            print(f"  - {p.name}")
        if len(unprocessed) > 10:
            print(f"  ... and {len(unprocessed) - 10} more.")

    elif args.scan:
        print("Scanning D:\\Claude_Sales\\Expenses for new invoice PDFs...")
        res = scanner.scan_and_process(dry_run=args.dry_run)
        print(f"\nDiscovered: {res['total_found']} candidate files")
        if res.get("skipped_duplicates"):
            print(f"Skipped Duplicates ({len(res['skipped_duplicates'])}):")
            for dup in res["skipped_duplicates"]:
                print(f"  - {dup['file']}: {dup['reason']}")
        if res.get("processed"):
            print(f"Processed ({len(res['processed'])}):")
            for p in res["processed"]:
                print(f"  - {p['file']} | Inv #{p['invoice_no']} | {p['supplier']} ({p['items_count']} lines)")
        if res.get("commit"):
            print(f"\nResult: {res['commit'].get('message')}")

    elif args.file:
        file_path = Path(args.file)
        if not file_path.exists():
            sys.exit(f"File not found: {file_path}")
        print(f"Processing: {file_path.name}...")
        res = scanner.process_single_pdf(file_path, dry_run=args.dry_run)
        if res.get("is_duplicate"):
            print(f"DUPLICATE DETECTED: {res.get('duplicate_reason')}")
        else:
            inv = res.get("invoice", {})
            print(f"Extracted Invoice #{inv.get('invoice_no')} dated {inv.get('invoice_date')}")
            print(f"Supplier: {inv.get('supplier')} | GSTIN: {inv.get('supplier_gstin')} | State: {inv.get('supplier_state')}")
            print(f"Buyer GSTIN: {inv.get('buyer_gstin')} | Place of Supply: {inv.get('place_of_supply')}")
            print(f"IRN: {inv.get('irn')}")
            print(f"Lines ({len(inv.get('items', []))}):")
            for i, itm in enumerate(inv.get('items', []), 1):
                print(f"  {i}. {itm.get('description')} | HSN {itm.get('hsn')} | Qty {itm.get('quantity')} | Rate ₹{itm.get('rate')} | Taxable ₹{itm.get('taxable_amount')} | {itm.get('tax_type')}")
            if res.get("commit"):
                print(f"\nResult: {res['commit'].get('message')}")

if __name__ == "__main__":
    main()
