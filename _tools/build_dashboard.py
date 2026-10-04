#!/usr/bin/env python3
"""Refresh the embedded data (const D=...) in Expense_Dashboard.html from the register.

Usage: python3 build_dashboard.py <register.xlsx> <Vendor_GL_Codes.xlsx> <in.html> <out.html> [asof dd-mm-yyyy]

Reads cached values (the register must have been recalculated/saved by LibreOffice or Excel).
Only the line starting with 'const D=' is replaced; the page design is untouched.
Columns are located by header name (header row 4 of 'Invoice Register').
"""
import sys, json, re, datetime, warnings
import openpyxl
warnings.filterwarnings('ignore')

reg_p, codes_p, html_in, html_out = sys.argv[1:5]
asof = sys.argv[5] if len(sys.argv) > 5 else datetime.date.today().strftime('%d-%m-%Y')

wb = openpyxl.load_workbook(reg_p, data_only=True)
ws = wb['Invoice Register']
hdr = [c.value for c in ws[4]]
col = {h: i for i, h in enumerate(hdr) if h}
need = {'file': 'Folder / File', 'no': 'Invoice No.', 'date': 'Invoice Date', 'sup': 'Supplier', 'desc': 'Description',
        'tt': 'Tax Type', 'tax': 'Taxable Amount (₹)', 'igst': 'IGST (₹)', 'cgst': 'CGST (₹)', 'sgst': 'SGST (₹)',
        'tot': 'Invoice Total (₹)', 'tds': 'TDS Amount (₹)', 'net': 'Net Payable (₹)', 'loc': 'Location Code',
        'vgl': 'Vendor GL Code', 'egl': 'Exps GL Code', 'bg': 'Buyer GSTIN', 'rcm': 'RCM / GST Compliance'}
missing = [v for v in need.values() if v not in col]
if missing:
    sys.exit('Missing register columns: %s' % missing)

def num(v):
    return round(float(v), 2) if isinstance(v, (int, float)) else 0.0

def txt(v):
    if v is None or v == '':
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()

rows = []
for r in ws.iter_rows(min_row=5, values_only=True):
    if not isinstance(r[0], (int, float)):  # S.No. present only on data rows
        continue
    d = r[col['Invoice Date']]
    if isinstance(d, datetime.datetime):
        d = d.date()
    if not isinstance(d, datetime.date):
        sys.exit('Row S.No. %s has no valid Invoice Date' % r[0])
    o = {}
    for k, h in need.items():
        v = r[col[h]]
        o[k] = num(v) if k in ('tax', 'igst', 'cgst', 'sgst', 'tot', 'tds', 'net') else txt(v)
    o['date'] = d.isoformat()
    rows.append(o)

# GL descriptions from the code list (and keep any already used in the register)
gldesc = {}
cw = openpyxl.load_workbook(codes_p, data_only=True)
for sh in cw.worksheets:
    hr = None
    for i, row in enumerate(sh.iter_rows(values_only=True), 1):
        if hr is None:
            if row and any(isinstance(c, str) and c.strip().lower().startswith('gl code') for c in row):
                hr = [str(c or '').strip().lower() for c in row]
                gi = next(j for j, c in enumerate(hr) if c.startswith('gl code'))
                di = next((j for j, c in enumerate(hr) if c.startswith('gl description')), None)
            continue
        g = txt(row[gi]) if gi < len(row) else None
        if g and di is not None and di < len(row) and row[di]:
            gldesc.setdefault(g, re.sub(r'\s*\(.*\)\s*$', '', str(row[di])).strip())

# Emails Not Scanned sheet -> list
emails = []
scan = ''
if 'Emails Not Scanned' in wb.sheetnames:
    es = wb['Emails Not Scanned']
    a2 = es['A2'].value or ''
    scan = a2.split(' Emails below')[0].strip()
    wbl = openpyxl.load_workbook(reg_p)  # for hyperlinks
    esl = wbl['Emails Not Scanned']
    heads = None
    for rr in range(1, es.max_row + 1):
        vals = [es.cell(rr, c).value for c in range(1, 14)]
        if vals[0] == '#':
            heads = [str(v) for v in vals]
            continue
        if heads is None or not isinstance(vals[0], (int, float)):
            continue
        dt = vals[1]
        if isinstance(dt, datetime.datetime):
            dt = dt.strftime('%d-%m-%Y %H:%M')
        link = None
        hl = esl.cell(rr, 12).hyperlink
        if hl is not None:
            link = (hl.target or '') + ('#' + hl.location.replace('%7C', '|') if hl.location else '')
        size = vals[6]
        emails.append({'n': int(vals[0]), 'date': txt(dt) or '—', 'frm': txt(vals[2]), 'subj': txt(vals[5]),
                       'size': round(size, 1) if isinstance(size, (int, float)) else '—', 'status': txt(vals[7]),
                       'reason': txt(vals[8]), 'action': txt(vals[9]), 'link': link or None})

D = {'rows': rows, 'emails': emails, 'asof': asof, 'scan': scan, 'gldesc': gldesc}
html = open(html_in, encoding='utf-8').read()
i = html.index('const D=')
j = html.index('\n', i)
html = html[:i] + 'const D=' + json.dumps(D, ensure_ascii=False, separators=(',', ':')).replace('</', '<\\/') + ';' + html[j:]
open(html_out, 'w', encoding='utf-8').write(html)
print('rows', len(rows), 'emails', len(emails), 'gl codes', len(gldesc),
      'invoice total', round(sum(r['tot'] for r in rows), 2))
