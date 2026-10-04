let currentInvoices = [];
let pendingExtractedInvoice = null;

function formatCurrency(val) {
  if (val === null || val === undefined || isNaN(val)) return '₹0';
  return '₹' + Number(val).toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function showBanner(type, message) {
  const b = document.getElementById('status-banner');
  b.className = `status-banner ${type}`;
  b.innerHTML = `<span>${message}</span><button style="background:none;border:none;color:inherit;cursor:pointer;font-size:1.1rem;" onclick="this.parentElement.style.display='none'">&times;</button>`;
  b.style.display = 'flex';
  setTimeout(() => {
    b.style.display = 'none';
  }, 10000);
}

async function fetchStatus() {
  try {
    const res = await fetch('/api/status');
    const data = await res.json();
    if (data.status === 'success') {
      const k = data.kpis;
      document.getElementById('kpi-invoices').innerText = k.total_invoices || 0;
      document.getElementById('kpi-lines').innerText = `${k.total_lines || 0} line items recorded`;
      document.getElementById('kpi-total').innerText = formatCurrency(k.total_invoice);
      document.getElementById('kpi-taxable').innerText = formatCurrency(k.total_taxable);
      document.getElementById('kpi-suppliers').innerText = `${k.total_suppliers || 0} active suppliers`;
      
      const totalGST = (k.total_igst || 0) + (k.total_cgst || 0) + (k.total_sgst || 0);
      document.getElementById('kpi-gst').innerText = formatCurrency(totalGST);
      document.getElementById('kpi-gst-breakdown').innerText = `IGST: ${formatCurrency(k.total_igst)} | CGST+SGST: ${formatCurrency((k.total_cgst || 0) + (k.total_sgst || 0))}`;
      
      document.getElementById('kpi-tds').innerText = formatCurrency(k.total_tds);
      document.getElementById('kpi-net').innerText = formatCurrency(k.total_net);

      if (data.unprocessed_count > 0) {
        showBanner('warning', `Found ${data.unprocessed_count} new invoice PDF(s) in Expenses folder ready to be processed.`);
      }
    }
  } catch (e) {
    console.error('Error fetching status:', e);
  }
}

async function fetchInvoices(searchQuery = '') {
  try {
    const url = searchQuery ? `/api/invoices?q=${encodeURIComponent(searchQuery)}` : '/api/invoices';
    const res = await fetch(url);
    const data = await res.json();
    if (data.status === 'success') {
      currentInvoices = data.rows;
      renderTable(data.rows);
      document.getElementById('table-meta').innerText = `Showing ${data.rows.length} recorded entries`;
    }
  } catch (e) {
    console.error('Error fetching invoices:', e);
  }
}

function renderTable(rows) {
  const tbody = document.getElementById('invoice-tbody');
  if (!rows || rows.length === 0) {
    tbody.innerHTML = `<tr><td colspan="10" style="text-align:center;padding:32px;color:var(--text-dim);">No invoices found matching query</td></tr>`;
    return;
  }

  tbody.innerHTML = rows.map(r => {
    let badgeClass = 'badge-igst';
    if (r.tax_type === 'CGST+SGST') badgeClass = 'badge-cgst';
    else if (r.tax_type.includes('Import')) badgeClass = 'badge-import';
    else if (r.tax_type.includes('unregistered')) badgeClass = 'badge-unregistered';

    return `
      <tr>
        <td class="num-cell" style="font-weight:600;color:var(--accent-blue);">${r.sno}</td>
        <td>${r.invoice_date || '-'}</td>
        <td style="font-weight:600;color:#fff;">${r.invoice_no || '-'}</td>
        <td>${r.supplier || '-'}</td>
        <td style="max-width:240px;overflow:hidden;text-overflow:ellipsis;" title="${r.description || ''}">${r.description || '-'}</td>
        <td class="text-right num-cell">${formatCurrency(r.taxable)}</td>
        <td><span class="badge ${badgeClass}">${r.tax_type || 'N/A'}</span></td>
        <td class="text-right num-cell" style="font-weight:600;color:#fff;">${formatCurrency(r.total)}</td>
        <td class="text-right num-cell" style="color:var(--accent-amber);">${formatCurrency(r.tds)}</td>
        <td class="text-right num-cell" style="font-weight:600;color:var(--accent-emerald);">${formatCurrency(r.net)}</td>
      </tr>
    `;
  }).join('');
}

let searchTimeout = null;
function handleSearch(val) {
  clearTimeout(searchTimeout);
  searchTimeout = setTimeout(() => {
    fetchInvoices(val);
  }, 250);
}

async function triggerFolderScan() {
  const btn = document.getElementById('scan-btn');
  const originalHtml = btn.innerHTML;
  btn.innerHTML = `<svg width="18" height="18" class="spin" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15"/></svg> Scanning...`;
  btn.disabled = true;

  try {
    const res = await fetch('/api/scan', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ dry_run: false })
    });
    const data = await res.json();

    if (data.status === 'success') {
      const r = data.result;
      if (r.processed && r.processed.length > 0) {
        showBanner('success', `Success: Processed and added ${r.processed.length} new invoices into Excel register!`);
      } else if (r.skipped_duplicates && r.skipped_duplicates.length > 0) {
        showBanner('warning', `Scan complete: All ${r.skipped_duplicates.length} discovered invoices were already in register (skipped duplicates).`);
      } else {
        showBanner('success', 'Scan complete: All files are up to date in the register.');
      }
      fetchStatus();
      fetchInvoices();
    } else {
      showBanner('error', `Scan error: ${data.message}`);
    }
  } catch (e) {
    showBanner('error', `Network error during scan: ${e}`);
  } finally {
    btn.innerHTML = originalHtml;
    btn.disabled = false;
  }
}

async function handleFileUpload(e) {
  const file = e.target.files[0];
  if (!file) return;

  const formData = new FormData();
  formData.append('file', file);
  formData.append('auto_commit', 'false');

  showBanner('warning', `Extracting invoice data from ${file.name}...`);

  try {
    const res = await fetch('/api/upload', {
      method: 'POST',
      body: formData
    });
    const data = await res.json();
    if (data.status === 'success') {
      openReviewModal(data.extracted);
    } else {
      showBanner('error', `Failed to extract: ${data.message}`);
    }
  } catch (err) {
    showBanner('error', `Upload error: ${err}`);
  }
  e.target.value = '';
}

function openReviewModal(extractedData) {
  const inv = extractedData.invoice;
  pendingExtractedInvoice = inv;

  const isDup = extractedData.is_duplicate;
  const modal = document.getElementById('review-modal');
  document.getElementById('modal-title').innerText = isDup ? 'Invoice Duplicate Detected' : 'Review Extracted Invoice';
  document.getElementById('modal-subtitle').innerText = isDup ? extractedData.duplicate_reason : `Ready to add to Excel register`;

  const commitBtn = document.getElementById('modal-commit-btn');
  if (isDup) {
    commitBtn.style.display = 'none';
  } else {
    commitBtn.style.display = 'inline-flex';
  }

  let itemsHtml = (inv.items || []).map((itm, i) => `
    <div style="background:rgba(255,255,255,0.03);padding:12px;border-radius:8px;margin-bottom:8px;border:1px solid var(--border-color);">
      <div style="display:flex;justify-content:space-between;margin-bottom:6px;">
        <span style="font-weight:600;color:#fff;">Line ${i+1}: ${itm.description || 'Item'}</span>
        <span class="badge badge-igst">${itm.tax_type}</span>
      </div>
      <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:8px;font-size:0.8rem;">
        <div><span style="color:var(--text-dim);">HSN:</span> ${itm.hsn || '-'}</div>
        <div><span style="color:var(--text-dim);">Qty:</span> ${itm.quantity} ${itm.unit}</div>
        <div><span style="color:var(--text-dim);">Rate:</span> ${formatCurrency(itm.rate)}</div>
        <div><span style="color:var(--text-dim);">Taxable:</span> ${formatCurrency(itm.taxable_amount)}</div>
      </div>
    </div>
  `).join('');

  document.getElementById('modal-body').innerHTML = `
    ${isDup ? `<div class="status-banner warning" style="display:flex;margin-bottom:16px;">⚠️ ${extractedData.duplicate_reason}</div>` : ''}
    <div class="review-grid">
      <div class="review-item"><div class="review-label">Invoice Number</div><div class="review-val">${inv.invoice_no || '-'}</div></div>
      <div class="review-item"><div class="review-label">Invoice Date</div><div class="review-val">${inv.invoice_date || '-'}</div></div>
      <div class="review-item"><div class="review-label">Supplier</div><div class="review-val">${inv.supplier || '-'}</div></div>
      <div class="review-item"><div class="review-label">Supplier GSTIN & State</div><div class="review-val">${inv.supplier_gstin || '-'} (${inv.supplier_state || '-'})</div></div>
      <div class="review-item"><div class="review-label">Buyer GSTIN & Place of Supply</div><div class="review-val">${inv.buyer_gstin || '-'} (${inv.place_of_supply || '-'})</div></div>
      <div class="review-item"><div class="review-label">IRN</div><div class="review-val" style="font-size:0.75rem;word-break:break-all;">${inv.irn || 'None'}</div></div>
      <div class="review-item"><div class="review-label">Ack No & Date</div><div class="review-val">${inv.ack_no || '-'} / ${inv.ack_date || '-'}</div></div>
      <div class="review-item"><div class="review-label">E-Way Bill</div><div class="review-val">${inv.eway_bill_no || 'None'}</div></div>
    </div>
    <h4 style="margin:16px 0 10px;font-size:0.9rem;text-transform:uppercase;color:var(--text-dim);">Line Items (${inv.items.length})</h4>
    ${itemsHtml}
  `;

  modal.style.display = 'flex';
}

function closeModal() {
  document.getElementById('review-modal').style.display = 'none';
  pendingExtractedInvoice = null;
}

async function commitModalInvoice() {
  if (!pendingExtractedInvoice) return;
  const btn = document.getElementById('modal-commit-btn');
  btn.innerText = 'Writing to Excel...';
  btn.disabled = true;

  try {
    const res = await fetch('/api/commit_extracted', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ invoice: pendingExtractedInvoice })
    });
    const data = await res.json();
    if (data.status === 'success') {
      showBanner('success', `Added to register! ${data.commit.message}`);
      closeModal();
      fetchStatus();
      fetchInvoices();
    } else {
      showBanner('error', `Failed to commit: ${data.message}`);
    }
  } catch (e) {
    showBanner('error', `Error committing: ${e}`);
  } finally {
    btn.innerText = 'Commit to Excel Register';
    btn.disabled = false;
  }
}

async function openFile(target) {
  try {
    const res = await fetch('/api/open_file', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ target })
    });
    const data = await res.json();
    if (data.status === 'success') {
      showBanner('success', data.message);
    } else {
      showBanner('error', data.message);
    }
  } catch (e) {
    console.error(e);
  }
}

// Drag & drop support
const dropZone = document.getElementById('drop-zone');
['dragenter', 'dragover'].forEach(eventName => {
  dropZone.addEventListener(eventName, e => {
    e.preventDefault();
    dropZone.style.borderColor = 'var(--accent-blue)';
    dropZone.style.background = 'rgba(59, 130, 246, 0.1)';
  }, false);
});

['dragleave', 'drop'].forEach(eventName => {
  dropZone.addEventListener(eventName, e => {
    e.preventDefault();
    dropZone.style.borderColor = 'rgba(255, 255, 255, 0.15)';
    dropZone.style.background = 'rgba(17, 24, 39, 0.5)';
  }, false);
});

dropZone.addEventListener('drop', e => {
  const dt = e.dataTransfer;
  const files = dt.files;
  if (files.length > 0) {
    const file = files[0];
    if (file.name.toLowerCase().endsWith('.pdf')) {
      const fakeEvent = { target: { files: [file], value: '' } };
      handleFileUpload(fakeEvent);
    } else {
      showBanner('error', 'Only PDF invoice files are supported.');
    }
  }
});

// Gmail Intake Integration
async function checkGmailStatus() {
  try {
    const res = await fetch('/api/gmail/status');
    const data = await res.json();
    if (data.status === 'success') {
      const infoEl = document.getElementById('gmail-last-sync-info');
      if (infoEl) {
        infoEl.innerHTML = `<strong>Status:</strong> ${data.has_password ? '✅ App Password Configured' : '⚠️ Password Not Set'}<br><strong>Last Scan:</strong> ${data.last_scan_ist || 'Never'}<br><strong>Pending Large Emails:</strong> ${data.pending_count || 0}`;
      }
      if (!data.has_password) {
        const syncBtn = document.getElementById('gmail-sync-btn');
        if (syncBtn) {
          syncBtn.title = 'Click to configure Google App Password';
        }
      }
    }
  } catch (e) {
    console.error('Error fetching gmail status:', e);
  }
}

function openGmailModal() {
  document.getElementById('gmail-modal').style.display = 'flex';
  checkGmailStatus();
}

function closeGmailModal() {
  document.getElementById('gmail-modal').style.display = 'none';
}

async function saveGmailSettings() {
  const mailbox = document.getElementById('gmail-mailbox-input').value.trim();
  const pwd = document.getElementById('gmail-password-input').value.trim();

  if (!pwd) {
    showBanner('warning', 'Please enter a 16-character Google App Password.');
    return;
  }

  showBanner('warning', 'Testing Gmail connection and saving credentials...');

  try {
    const res = await fetch('/api/gmail/config', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ mailbox: mailbox, app_password: pwd })
    });
    const data = await res.json();
    if (data.status === 'success') {
      showBanner('success', data.message);
      closeGmailModal();
      checkGmailStatus();
    } else {
      showBanner('error', `Failed to save: ${data.message}`);
    }
  } catch (e) {
    showBanner('error', `Error saving Gmail settings: ${e}`);
  }
}

async function triggerGmailSync() {
  const btn = document.getElementById('gmail-sync-btn');
  const origHtml = btn.innerHTML;
  btn.innerHTML = `<svg width="18" height="18" class="spin" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15"/></svg> Connecting Gmail...`;
  btn.disabled = true;

  showBanner('warning', 'Connecting to invoice@nicobar.com to check new invoice emails...');

  try {
    const res = await fetch('/api/gmail/sync', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ dry_run: false })
    });
    const data = await res.json();

    if (data.status === 'success') {
      const r = data.result;
      if (r.downloaded_count > 0) {
        showBanner('success', `Gmail Sync: Downloaded ${r.downloaded_count} invoices, recorded ${r.added_to_register} new lines into Excel register!`);
      } else {
        showBanner('success', `Gmail Sync Complete: ${r.message || 'No new invoice emails found since last watermark.'}`);
      }
      fetchStatus();
      fetchInvoices();
      checkGmailStatus();
    } else {
      if (data.message && data.message.includes('Missing GMAIL_APP_PASSWORD')) {
        showBanner('warning', 'Please enter your Google App Password to link invoice@nicobar.com.');
        openGmailModal();
      } else {
        showBanner('error', `Gmail sync failed: ${data.message}`);
      }
    }
  } catch (e) {
    showBanner('error', `Network error during Gmail sync: ${e}`);
  } finally {
    btn.innerHTML = origHtml;
    btn.disabled = false;
  }
}

// Initial load
window.addEventListener('DOMContentLoaded', () => {
  fetchStatus();
  fetchInvoices();
  checkGmailStatus();
});
