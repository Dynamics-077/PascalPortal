/* ============================================================
   Pascal Press Sales Rep Portal — Shared JS
   Phase 1: Internal Sales Rep Portal
   - Auth: Microsoft Entra ID (Azure AD) SSO
   - User: Sales Agents placing orders on behalf of customers
   ============================================================ */

'use strict';

/* ---- Active nav highlighting ------------------------------ */
(function () {
  const page = location.pathname.split('/').pop() || 'dashboard.html';
  document.querySelectorAll('.nav-item').forEach(el => {
    const href = el.getAttribute('href') || '';
    if (href && page.includes(href.replace('.html', '').replace('#', ''))) {
      if (!href.includes('#')) el.classList.add('active');
    }
  });
})();

/* ---- Toast ------------------------------------------------ */
function showToast(message, type = 'default', duration = 3500) {
  let container = document.getElementById('toast-container');
  if (!container) {
    container = document.createElement('div');
    container.id = 'toast-container';
    document.body.appendChild(container);
  }
  const icons = { success: '✓', error: '✕', warning: '⚠', default: 'ℹ' };
  const toast = document.createElement('div');
  toast.className = `toast ${type}`;
  toast.innerHTML = `<span>${icons[type] || icons.default}</span><span>${message}</span>`;
  container.appendChild(toast);
  setTimeout(() => {
    toast.style.animation = 'none';
    toast.style.opacity = '0';
    toast.style.transform = 'translateX(30px)';
    toast.style.transition = 'all .25s ease';
    setTimeout(() => toast.remove(), 280);
  }, duration);
}

/* ---- Modal helpers ---------------------------------------- */
function openModal(id) {
  const el = document.getElementById(id);
  if (el) el.classList.add('open');
}
function closeModal(id) {
  const el = document.getElementById(id);
  if (el) el.classList.remove('open');
}
document.addEventListener('click', e => {
  if (e.target.classList.contains('modal-overlay')) {
    e.target.classList.remove('open');
  }
});
document.addEventListener('keydown', e => {
  if (e.key === 'Escape') {
    document.querySelectorAll('.modal-overlay.open').forEach(m => m.classList.remove('open'));
  }
});

/* ---- Tab switching ---------------------------------------- */
function initTabs(containerSelector) {
  const containers = document.querySelectorAll(containerSelector || '[data-tabs]');
  containers.forEach(container => {
    const buttons = container.querySelectorAll('.tab-btn');
    const panels  = container.querySelectorAll('.tab-panel');
    buttons.forEach((btn, i) => {
      btn.addEventListener('click', () => {
        buttons.forEach(b => b.classList.remove('active'));
        panels.forEach(p => p.classList.remove('active'));
        btn.classList.add('active');
        if (panels[i]) panels[i].classList.add('active');
      });
    });
    if (buttons[0]) buttons[0].classList.add('active');
    if (panels[0])  panels[0].classList.add('active');
  });
}
document.addEventListener('DOMContentLoaded', () => {
  initTabs();
  // Pre-fill topbar chip from MSAL-stored salesrep (written by index.html on SSO login)
  const cached = getSalesRep();
  if (cached && cached.name) setTopbarUser(cached.name);
});

/* ---- Logout (Entra ID: clear session + redirect) ---------- */
function logout() {
  // 1. Clear local memory and session storage tokens
  localStorage.removeItem('pp_salesrep');
  localStorage.removeItem('pp_rep_name');
  localStorage.removeItem('pp_token');
  sessionStorage.clear();

  showToast('Signing out...', 'default', 1000);

  // 2. Redirect locally to the Sign In page immediately
  setTimeout(() => { 
    window.location.href = 'index.html'; 
  }, 800);
}

/* ============================================================
   MOCK DATA — Phase 1: Sales Rep Portal
   Replace all with real fetch() calls to:
   - Microsoft Entra ID (MSAL) for auth
   - D365 F&O OData/REST endpoints via Node.js middleware
   ============================================================ */

const MOCK = {
  salesRep: {
    id:           '',
    name:         'Sales Rep',
    firstName:    'Rep',
    email:        '',
    phone:        '',
    territory:    '',
    d365WorkerId: '',
    entraOid:     '',
    targets:     { monthly: 0, ytd: 0 },
    performance: { monthly: 0, ytd: 0 }
  },
  customers:        [],
  salesOrders:      [],
  dashboardSummary: {
    ordersToday:      0,
    ordersThisWeek:   0,
    ordersThisMonth:  0,
    revenueThisMonth: 0,
    pendingApprovals: 0,
    monthlyTarget:    0,
    ytdRevenue:       0,
    ytdTarget:        0,
  }
};

/* ---- Customer search helper (rep's territory) ------------ */
function searchCustomers(query) {
  /* TODO: GET /api/customers/search?q={query} — filtered by rep's territory */
  if (typeof DB !== 'undefined' && DB.ready && typeof DB.searchCustomers === 'function') {
    return DB.searchCustomers(query);
  }
  if (!query || query.length < 1) return MOCK.customers;
  const q = query.toLowerCase();
  return MOCK.customers.filter(c =>
    c.name.toLowerCase().includes(q) ||
    c.id.toLowerCase().includes(q) ||
    c.contactPerson.toLowerCase().includes(q)
  ).slice(0, 10);
}

/* ---- Get logged-in sales rep ----------------------------- */
function getSalesRep() {
  const stored = localStorage.getItem('pp_salesrep');
  if (stored) {
    try { return JSON.parse(stored); } catch(e) {}
  }
  return MOCK.salesRep;
}

/* ---- Currency formatter ----------------------------------- */
function fmtCurrency(val) {
  return new Intl.NumberFormat('en-AU', { style: 'currency', currency: 'AUD' }).format(val || 0);
}

/* ---- Date formatter --------------------------------------- */
function fmtDate(str) {
  if (!str) return '—';
  const d = new Date(str);
  return d.toLocaleDateString('en-AU', { day: '2-digit', month: 'short', year: 'numeric' });
}

/* ---- Status badge HTML ------------------------------------ */
function statusBadge(status) {
  // Matches SharePoint SalesOrderHeader Status choices exactly
  const map = {
    'Draft':       'badge-pending',
    'In Progress': 'badge-open',
    'Submit':      'badge-confirmed',
    'Complete':    'badge-paid',
    'Shipped':     'badge-shipped',
    'Invoiced':    'badge-partial',
    'Cancelled':   'badge-cancelled',
    // Legacy / fallback
    'Open':        'badge-open',
    'Confirmed':   'badge-confirmed',
    'On Hold':     'badge-overdue',
    'Overdue':     'badge-overdue',
  };
  const cls = map[status] || 'badge-open';
  return `<span class="badge ${cls}">${status}</span>`;
}

/* ---- Calculate line net amount ---------------------------- */
function calcNet(price, qty, discount) {
  return price * qty * (1 - (discount || 0) / 100);
}

/* ---- Generate a draft Sales Order ID (matches server format) */
function nextOrderId() {
  const date = new Date().toISOString().slice(0, 10).replace(/-/g, '');
  const rand = Math.floor(Math.random() * 9000) + 1000; // 1000-9999, matches PA rand(1000,9999)
  return `PAS-${date}-${rand}`;
}

/* ============================================================
   EXCEL EXPORT — powered by SheetJS
   ============================================================ */

function _downloadXLSX(workbook, filename) {
  if (typeof XLSX === 'undefined') {
    showToast('Excel library not loaded — check internet connection', 'error');
    return;
  }
  XLSX.writeFile(workbook, filename);
  showToast(`✓ Downloaded: ${filename}`, 'success');
}

function _setColWidths(ws, widths) {
  ws['!cols'] = widths.map(w => ({ wch: w }));
}

function _buildSheet(headers, rows) {
  const data = [headers, ...rows];
  return XLSX.utils.aoa_to_sheet(data);
}

/* Export Sales Orders (rep-centric: includes Customer column) */
function exportOrders(orders) {
  orders = orders || MOCK.salesOrders;
  const rep = getSalesRep();

  const headers = [
    'Order ID', 'Date Placed', 'Customer Name', 'Customer Ref (PO#)',
    'Status', 'Delivery Date', 'Subtotal (ex GST)', 'GST (10%)',
    'Total (incl GST)', 'Warehouse'
  ];

  const rows = orders.map(o => {
    const gst   = o.total * 0.1;
    const total = o.total + gst;
    return [
      o.id, o.date, o.customerName || '', o.reference || '',
      o.status, o.deliveryDate || '',
      o.total, parseFloat(gst.toFixed(2)), parseFloat(total.toFixed(2)), 'Sydney (SYD)'
    ];
  });

  const totalAmt = orders.reduce((s, o) => s + o.total, 0);
  rows.push([]);
  rows.push(['', '', '', '', '', 'TOTAL', parseFloat(totalAmt.toFixed(2)),
    parseFloat((totalAmt * 0.1).toFixed(2)), parseFloat((totalAmt * 1.1).toFixed(2)), '']);

  const ws = _buildSheet(headers, rows);
  _setColWidths(ws, [14, 14, 28, 18, 12, 14, 18, 12, 18, 14]);

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Sales Orders');

  const summaryData = [
    ['Pascal Press — Sales Order Export'],
    [''],
    ['Sales Rep',     rep.name],
    ['Territory',     rep.territory],
    ['Export Date',   new Date().toLocaleDateString('en-AU')],
    ['Orders',        orders.length],
    ['Total (ex GST)', parseFloat(totalAmt.toFixed(2))],
    ['Total (incl GST)', parseFloat((totalAmt * 1.1).toFixed(2))],
  ];
  const wsSummary = XLSX.utils.aoa_to_sheet(summaryData);
  _setColWidths(wsSummary, [22, 28]);
  XLSX.utils.book_append_sheet(wb, wsSummary, 'Summary');

  const date = new Date().toISOString().slice(0, 10);
  _downloadXLSX(wb, `PascalPress_Orders_${rep.id}_${date}.xlsx`);
}

function exportOrderLines(order) {
  if (!order || !order.lines || order.lines.length === 0) {
    showToast('No line items to export for this order', 'warning');
    return;
  }
  const rep = getSalesRep();

  const infoData = [
    ['Pascal Press — Order Detail Export'],
    [''],
    ['Order ID',       order.id],
    ['Customer',       order.customerName || ''],
    ['Sales Rep',      rep.name],
    ['Customer Ref',   order.reference || ''],
    ['Order Date',     order.date],
    ['Delivery Date',  order.deliveryDate || ''],
    ['Status',         order.status],
  ];
  const wsInfo = XLSX.utils.aoa_to_sheet(infoData);
  _setColWidths(wsInfo, [20, 30]);

  const lineHeaders = [
    'Line #', 'ISBN', 'Product Name', 'Category',
    'Qty', 'Unit', 'Unit Price', 'Discount %', 'Net Amount', 'Delivery Type'
  ];
  const lineRows = order.lines.map(l => [
    l.lineNo, l.itemNo, l.name, l.category || '',
    l.qty, l.unit, l.price, l.discount || 0,
    parseFloat(calcNet(l.price, l.qty, l.discount).toFixed(2)), l.deliveryType
  ]);

  const lineTotal = order.lines.reduce((s, l) => s + calcNet(l.price, l.qty, l.discount), 0);
  lineRows.push([]);
  lineRows.push(['', '', '', '', order.lines.reduce((s,l)=>s+l.qty,0), '', '', 'TOTAL', parseFloat(lineTotal.toFixed(2)), '']);

  const wsLines = _buildSheet(lineHeaders, lineRows);
  _setColWidths(wsLines, [8, 18, 38, 14, 8, 8, 12, 12, 14, 14]);

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, wsLines, 'Order Lines');
  XLSX.utils.book_append_sheet(wb, wsInfo,  'Order Info');
  _downloadXLSX(wb, `PascalPress_Order_${order.id}.xlsx`);
}

/* ---- Topbar user chip ------------------------------------- */
function setTopbarUser(name) {
  if (!name) return;
  const nameEl   = document.getElementById('topbarUserName');
  const avatarEl = document.getElementById('topbarAvatar');
  if (!nameEl) return;
  nameEl.textContent = name;
  if (avatarEl) {
    avatarEl.textContent = name.split(' ').map(w => w[0]).join('').slice(0, 2).toUpperCase();
  }
  // Cache so next page load shows name immediately
  try { localStorage.setItem('pp_rep_name', name); } catch (_) {}
}

/* ---- Expose globals --------------------------------------- */
window.PP = {
  MOCK, searchCustomers, getSalesRep, setTopbarUser,
  fmtCurrency, fmtDate, statusBadge, calcNet, nextOrderId,
  showToast, openModal, closeModal,
  exportOrders, exportOrderLines,
};
