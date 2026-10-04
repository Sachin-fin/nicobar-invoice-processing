---
name: "invoice-register-entry"
description: "Use when extracting purchase/expense invoice details (PDFs in D:\\Google _invoice Processing) and adding them to 1. Expense_Invoice_Register.xlsx for Nicobar Design."
---

# Invoice Register Entry

Add invoice details to Sachin's existing Excel register, always following the column layout he has set in that file. Never rebuild or redesign the register. Never add the same invoice twice.

## Files

- Invoices: one folder per supplier under `D:\Google _invoice Processing`, file name `<Invoice No> <dd.mm.yy>.pdf` (for example `Clay Craft India Limited\MN-10076-26-27 12.09.26.pdf`). New files may also be dropped in the root or in date folders.
- Register: `D:\Google _invoice Processing\1. Expense_Invoice_Register.xlsx` (if not found, use the file whose name contains "Expense_Invoice_Register"), sheet `Invoice Register`, header row 4.
- Codes: `D:\Google _invoice Processing\3.Vendor_GL_Codes.xlsx`.
- Not invoices: anything under `_tools` or `_docs`, `Cowork_Invoice_to_Dynamics365_Process.pdf`, and files listed in `ignored_files` in `_gmail_scan_state.json`.

## Steps

1. **Find new invoices.** List the Expenses folder recursively. Skip any PDF whose path is already in the register's "Folder / File" column and anything listed under "Not invoices" (unless Sachin names specific files).
2. **Read the register's current layout first.** Stage the register, then load it with openpyxl twice: once normally for formulas, once with `data_only=True` for values (read only – never save with openpyxl, it breaks the Dashboard charts). The header row (currently row 4) is the source of truth. Use its exact column names, order and count. Find columns by header name, never by a fixed letter. Note each column's number format, font colour (blue = input, black = formula) and the formula in the last data row.
3. **Extract every field those headers ask for** from each invoice PDF: invoice no./date, supplier name, GSTIN and state, buyer GSTIN, place of supply, PO no./date, IRN, Ack no./date, E-Way Bill, transporter, LR, vehicle, item code, description, HSN, boxes, quantity, unit, rate, discount, GST rate, tax type, round-off, RCM / GST compliance note, TDS section and rate. Add one row per invoice line item. Enter extra charges (for example repacking) as their own row. If a header asks for something that isn't on the invoice, leave the cell blank and mention it.
4. **Duplicate check. Ignore repeat invoices.** Before adding an invoice, compare it with the invoices already in the register and with the other new PDFs in the same run. Treat it as a duplicate if any of these match:
   - the same Invoice No. and the same Supplier (or Supplier GSTIN). Normalise both first: ignore case, spaces and punctuation.
   - the same IRN, when one is present.

   A duplicate is never added, even if the file name is different. Don't move or delete the duplicate PDF. List it in the reply as "Skipped – duplicate of <invoice no.> (already at S.No. x)". Several rows with the same invoice number are normal (one per line item).
5. **Add rows above the blank row before TOTAL** using LibreOffice UNO (insert whole rows on "Invoice Register", `copyRange` from the previous data row, then set the input cells). Continue the S.No. sequence. Formula columns (Year, Month, Taxable Amount, IGST, CGST, SGST, Invoice Total, TDS Amount, Net Payable) must stay formulas copied from the row above – never hardcode them. In UNO `setFormula()` the argument separator is `;`. Extend the TOTAL row's SUM ranges, the autofilter and `_xlnm._FilterDatabase` to cover the new rows; keep the Note rows below TOTAL.
6. **Tax logic.** Compare the supplier's state code (first two digits of the Supplier GSTIN) with the state code of the Nicobar GSTIN the invoice is billed to (Place of Supply). Different → IGST; same → CGST+SGST. Nicobar has GSTINs in several states (06, 07, 09, 22, 24, 27, 36 …) – never assume Haryana. Import invoices with no GST → Tax Type "Import (no GST)" at 0%. Unregistered suppliers → "No GST (unregistered)". Round Off = printed invoice total − computed total, on the invoice's last line.
7. **TDS.** Fill TDS Section (with the reason) and TDS Rate. 194Q (0.1% on goods) only when this FY's purchases from the supplier exceed ₹50 L (on the excess); otherwise rate 0 and say why. Respect single-bill/annual thresholds for 194C/194J. Never change TDS on existing rows; list doubts in the reply.
8. **Codes.** Fill Location Code, Vendor GL Code and Exps GL Code from 3.Vendor_GL_Codes.xlsx or earlier rows for the same supplier; never invent codes. Add any new supplier name to 3.Vendor_GL_Codes.xlsx with the codes blank.
9. **Foreign currency.** Convert to ₹ at the RBI/FBIL reference rate for the invoice date. Write Rate (₹) as a formula such as `=USD_rate*FX`. Add a Note row below TOTAL with the currency, original amount and FX rate, and flag the rate in the reply so Sachin can replace it with the customs or bank rate.
10. **Verify.** Re-open the saved file with `data_only=True`: there must be no error values (#REF!, #NAME?, #VALUE!, #DIV/0!) and the Dashboard check "vendor total = invoice total" must say OK. Check that each new invoice's computed total matches the total printed on the PDF. Report any mismatch rather than forcing it.
11. **Save back.** If the register is open in Excel (the stage call reports saveBackBlocked / open_in_another_app), ask Sachin to save and close it first. Commit the file back to the same path using the staged mtime as `expectedMtimeMs`; if rejected, re-stage and redo – never force.
12. **Reply briefly.** List the invoices added (number, supplier, ₹ total, TDS, net payable), duplicates skipped, codes missing, and any blanks, mismatches or exchange rates assumed.
