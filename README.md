# Nicobar Design - Expense Invoice Processing System

Complete automated software solution for extracting purchase/expense invoice details and updating `1. Expense_Invoice_Register.xlsx` according to the rules in [SKILL.md](file:///d:/Google%20_invoice%20Processing/SKILL.md).

---

## 🚀 Quick Start (Running the Software)

### Method 1: Web Application Dashboard (Recommended)
Double-click `start_software.bat` or run:
```powershell
.venv\Scripts\python app.py
```
Then open your browser at: **[http://localhost:5050](http://localhost:5050)**

#### Web Dashboard Features:
- **Executive KPIs**: Real-time summary of Invoices, Gross Value, Taxable Value, IGST/CGST/SGST, TDS, and Net Payable.
- **Scan Expenses Folder**: One-click scan of `D:\Claude_Sales\Expenses` to detect and record new invoices.
- **Upload & Interactive Verification**: Drag-and-drop any invoice PDF, review all extracted fields (Supplier, GSTIN, PO, IRN, line items, taxes, TDS), and commit to Excel with one click.
- **Duplicate Prevention**: Prevents double-entry by checking normalized invoice numbers, suppliers, and IRN.
- **Direct Launchers**: One-click buttons to open `1. Expense_Invoice_Register.xlsx`, `3.Vendor_GL_Codes.xlsx`, and `Expense_Dashboard.html`.

---

### Method 2: Command-Line Interface (CLI)

- **View Current Status & KPIs**:
  ```powershell
  .venv\Scripts\python run_cli.py --status
  ```
- **Scan & Process All Invoices**:
  ```powershell
  .venv\Scripts\python run_cli.py --scan
  ```
- **Dry-run Scan (Preview without modifying Excel)**:
  ```powershell
  .venv\Scripts\python run_cli.py --scan --dry-run
  ```
- **Process a Single PDF**:
  ```powershell
  .venv\Scripts\python run_cli.py --file "D:\Claude_Sales\Expenses\SupplierName\Invoice.pdf"
  ```

---

## 🛠️ Architecture & Core Components

| File | Purpose |
|------|---------|
| [app.py](file:///d:/Google%20_invoice%20Processing/app.py) | Full-featured Flask Web Application & REST API backend. |
| [invoice_extractor.py](file:///d:/Google%20_invoice%20Processing/invoice_extractor.py) | High-precision PDF extraction engine using PyMuPDF & regex for Indian GST invoices. |
| [excel_manager.py](file:///d:/Google%20_invoice%20Processing/excel_manager.py) | Direct OpenXML spreadsheet updater that preserves formulas, formatting, and dashboard charts. |
| [scanner.py](file:///d:/Google%20_invoice%20Processing/scanner.py) | Folder scanner, duplicate validator, and pipeline coordinator. |
| [run_cli.py](file:///d:/Google%20_invoice%20Processing/run_cli.py) | Command-line runner for batch tasks. |
| [config.py](file:///d:/Google%20_invoice%20Processing/config.py) | Configuration paths, Nicobar GSTINs, and state code mappings. |
| [templates/index.html](file:///d:/Google%20_invoice%20Processing/templates/index.html) | Modern dark-mode dashboard UI. |
| [static/styles.css](file:///d:/Google%20_invoice%20Processing/static/styles.css) | Custom responsive CSS styling. |
| [static/app.js](file:///d:/Google%20_invoice%20Processing/static/app.js) | Client-side reactive interface controller. |
| [start_software.bat](file:///d:/Google%20_invoice%20Processing/start_software.bat) | One-click launcher for Windows. |
