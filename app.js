const SHEET_ID    = '1vIERVGUheXWkMS155VWfBEuCrUV4qXGYSUM9mIdppfc';
const CSV_URL     = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=csv`;
const SHIP_URL    = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=csv&gid=1459899540`;

// ── Print Logger config ────────────────────────────────────────
// Paste your Google Apps Script Web App URL here after deploying:
const SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbz1LuTt6ySUIXR_Rp3f8EyE4nV1XNW3hmK8YqirpxN_6HwzXFvvtyuVl_7Pj8_eNICs/exec';
// List your printers here:
const PRINTERS = ['Bottle 1', 'Bottle 2', 'Mug 1', 'Travel Bottle 1'];

function todayStr() {
  const d = new Date();
  return `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`;
}
function withDate(note) { return note ? `[${todayStr()}] ${note}` : ''; }

// ── Known owners (always appear in dropdowns) ─────────────────
const KNOWN_OWNERS = ['Daan','Gerrit','Jim','Mees','Skip'];
const OWNER_ALIASES = { 'geertjan': 'Gerrit', 'gerrit': 'Gerrit' };
function normOwner(name) { return OWNER_ALIASES[(name||'').toLowerCase()] || name; }

// ── Auth ──────────────────────────────────────────────────────
const ALLOWED_EMAILS = [
  'daan@izybottles.com',
  'jim@izybottles.com',
  'biessenlevi@gmail.com',
  'sharon@orderchamp.com',
  'ivan@izybottles.com',
  'geertjan@izybottles.com',
  'mees@izybottles.com',
];

// Emails that only see the Active Queue tab
const ACTIVE_QUEUE_ONLY = ['ivan@izybottles.com'];

let currentUser   = null;
let sleeveRows    = [];
let sleeveLoaded  = false;
let mockupRows    = [];
let mockupLoaded  = false;
let mkTypeFilter  = '';

// Multi-select state
const aqSelected = new Set(); // indices into allRows
const svSelected = new Set(); // indices into sleeveRows
const mkSelected = new Set(); // indices into mockupRows

function updateSelectionBar() {
  const total = aqSelected.size + svSelected.size + mkSelected.size;
  const bar   = document.getElementById('selection-bar');
  if (total === 0) { bar.style.display = 'none'; return; }
  bar.style.display = 'flex';
  document.getElementById('selection-bar-text').textContent = `${total} row${total !== 1 ? 's' : ''} selected`;
}

function clearSelection() {
  aqSelected.clear(); svSelected.clear(); mkSelected.clear();
  document.querySelectorAll('.row-select').forEach(cb => { cb.checked = false; });
  updateSelectionBar();
}

function handleCredentialResponse(response) {
  const payload = parseJwt(response.credential);
  const email = (payload.email || '').toLowerCase();
  if (!ALLOWED_EMAILS.includes(email)) {
    document.getElementById('login-error').textContent = 'Access denied for: ' + payload.email;
    return;
  }
  currentUser = { name: payload.given_name || payload.name, email: payload.email };
  localStorage.setItem('izy_user', JSON.stringify(currentUser));
  showApp();
}

function parseJwt(token) {
  const base64 = token.split('.')[1].replace(/-/g, '+').replace(/_/g, '/');
  return JSON.parse(atob(base64));
}

function showApp() {
  document.getElementById('login-overlay').style.display = 'none';
  document.getElementById('user-name').textContent = currentUser.name;

  // Restrict tabs for limited users
  if (ACTIVE_QUEUE_ONLY.includes((currentUser.email || '').toLowerCase())) {
    document.querySelectorAll('.tab-btn:not([data-tab="active-queue"])').forEach(b => b.style.display = 'none');
    activateTab('active-queue');
  }

  refreshData();
}

function signOut() {
  localStorage.removeItem('izy_user');
  currentUser = null;
  location.reload();
}

// ─────────────────────────────────────────────────────────────
let allRows      = [];
let printLogRows = []; // append-only log from PrintLog sheet
let shippedRows  = []; // confirmed matched shipping rows
let reviewRows   = []; // unmatched / low-confidence rows for review
let charts       = {};
let invoiceData  = []; // Moneybird invoices

// ── Moneybird helpers ─────────────────────────────────────────
async function loadInvoices() {
  try {
    const data = await fetch(SCRIPT_URL + '?sheet=invoices&t=' + Date.now()).then(r => r.json());
    invoiceData = data.invoices || [];
  } catch(_) { invoiceData = []; }
}

function normCompany(s) {
  return (s || '').toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]/g, '');
}

function matchInvoice(companyName) {
  if (!invoiceData.length || !companyName) return null;
  const jobNorm = normCompany(companyName);
  if (jobNorm.length < 3) return null;
  const matches = invoiceData.filter(inv => {
    const n = normCompany(inv.company);
    return n.includes(jobNorm) || jobNorm.includes(n);
  });
  if (!matches.length) return null;
  const priority = { late: 0, open: 1, draft: 2, paid: 3 };
  matches.sort((a, b) => {
    const pd = (priority[a.state] ?? 4) - (priority[b.state] ?? 4);
    return pd !== 0 ? pd : (b.date || '').localeCompare(a.date || '');
  });
  return matches[0];
}

function invoiceBadge(inv) {
  if (!inv) return '<span class="inv-badge inv-none">No invoice</span>';
  const labels = { draft: 'Draft', open: 'Sent', late: 'Overdue', paid: 'Paid' };
  const cls    = { draft: 'inv-draft', open: 'inv-open', late: 'inv-late', paid: 'inv-paid' };
  return `<span class="inv-badge ${cls[inv.state] || 'inv-none'}">${labels[inv.state] || inv.state}</span>`;
}

// ── Fetch & Parse ─────────────────────────────────────────────
async function fetchData() {
  const res  = await fetch(CSV_URL + '&t=' + Date.now());
  const text = await res.text();
  return parseCSV(text);
}

function colLetter(i) {
  // Convert 0-based column index to spreadsheet letter (A, B, ..., Z, AA, ...)
  let s = '', n = i + 1;
  while (n > 0) { n--; s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26); }
  return s;
}

function parseCSV(text) {
  const lines   = text.trim().split('\n');
  const headers = splitLine(lines[0]).map(h => h.replace(/"/g, '').trim());
  return lines.slice(1).map((line, idx) => {
    const vals = splitLine(line);
    const row  = {};
    headers.forEach((h, i) => {
      const v = (vals[i] || '').replace(/"/g, '').trim();
      row[h] = v;
      row['_col_' + colLetter(i)] = v; // also accessible by column letter e.g. _col_Q
    });
    row['_sheetRow'] = idx + 2;
    return row;
  });
}

function splitLine(line) {
  const result = []; let cur = ''; let q = false;
  for (const ch of line) {
    if (ch === '"') q = !q;
    else if (ch === ',' && !q) { result.push(cur); cur = ''; }
    else cur += ch;
  }
  result.push(cur);
  return result;
}

// ── Helpers ───────────────────────────────────────────────────
function get(row, key) { return (row[key] || '').trim(); }
function getCI(row, keyword) {
  const k = Object.keys(row).find(k => !k.startsWith('_') && k.toLowerCase().includes(keyword.toLowerCase()));
  return k ? (row[k] || '').trim() : '';
}
function num(row, key) { return parseInt(row[key]) || 0; }

function parseDate(str) {
  if (!str || str === '—' || str === 'Geen deadline') return null;
  const [d, m, y] = str.split('/');
  if (!d || !m || !y) return null;
  return new Date(+y, +m - 1, +d);
}

function daysFrom(date) {
  if (!date) return null;
  return Math.round((date - new Date()) / 86400000);
}

const TYPE_COLORS = {
  'bottle':         { bg: '#dbeafe', text: '#1d4ed8' },
  'mug':            { bg: '#fce7f3', text: '#be185d' },
  'travel bottle':  { bg: '#dcfce7', text: '#15803d' },
  'tumbler':        { bg: '#fef3c7', text: '#b45309' },
};

function typeBadge(type) {
  if (!type) return '—';
  const key = type.toLowerCase();
  const c = Object.entries(TYPE_COLORS).find(([k]) => key.includes(k));
  const { bg, text } = c ? c[1] : { bg: '#f1f5f9', text: '#475569' };
  return `<span style="background:${bg};color:${text};border-radius:4px;padding:1px 7px;font-size:11px;font-weight:600;white-space:nowrap;">${type}</span>`;
}

function badge(status) {
  const s = (status || '').toLowerCase();
  if (s === 'shipped')          return `<span class="badge b-shipped">Shipped</span>`;
  if (s === 'to print')         return `<span class="badge b-to-print">To Print</span>`;
  if (s === 'to sleeve')           return `<span class="badge b-to-print">To Sleeve</span>`;
  if (s === 'to make')             return `<span class="badge b-to-print">To make</span>`;
  if (s === 'made')                return `<span class="badge b-progress">Made</span>`;
  if (s === 'feedback/revisions')  return `<span class="badge b-waiting">Feedback/revisions</span>`;
  if (s === 'waiting')             return `<span class="badge b-waiting">Waiting</span>`;
  if (s === 'ordered')             return `<span class="badge b-progress">Ordered</span>`;
  if (s === 'finished')            return `<span class="badge b-shipped">Finished</span>`;
  if (s.includes('progress'))      return `<span class="badge b-progress">In Progress</span>`;
  if (s === 'ready to ship')       return `<span class="badge b-ready-ship">Ready to Ship</span>`;
  if (s === 'done')                return `<span class="badge b-shipped">Done</span>`;
  if (!status)                  return '';
  return `<span class="badge b-default">${status}</span>`;
}

function daysCell(days) {
  if (days === null) return '—';
  if (days < 0)  return `<span class="cell-danger">${Math.abs(days)}d late</span>`;
  if (days <= 3) return `<span class="cell-warn">${days}d left</span>`;
  return `<span class="cell-ok">${days}d left</span>`;
}

function isActive(row) {
  return get(row, 'Status').toLowerCase() !== 'shipped';
}

function isOverdue(row) {
  const d = parseDate(get(row, 'Deadline'));
  return d && d < new Date() && isActive(row);
}

// ── Tab switching ─────────────────────────────────────────────
function activateTab(tabName) {
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  document.querySelectorAll('.tab-content').forEach(s => s.classList.remove('active'));
  const btn = document.querySelector(`.tab-btn[data-tab="${tabName}"]`);
  if (btn) btn.classList.add('active');
  const content = document.getElementById('tab-' + tabName);
  if (content) content.classList.add('active');
  if (tabName === 'shipping') loadShipping();
  if (tabName === 'add-job') { populateAddJobOwners(); if (!mockupLoaded) loadMockups().then(populateApprovedMockups); else populateApprovedMockups(); }
  if (tabName === 'sleeves') { populateSleeveOwners(); loadSleeves(); }
  if (tabName === 'mockups') loadMockups();
  if (tabName === 'stock') loadStock();
  if (tabName === 'calendar') loadCalendar();
}

document.querySelectorAll('.tab-btn').forEach(btn => {
  btn.addEventListener('click', () => activateTab(btn.dataset.tab));
});

// On mobile, always start on Active Queue
if (window.innerWidth <= 768) activateTab('active-queue');

// ── Stats ─────────────────────────────────────────────────────
function renderStats(rows) {
  const shipped    = rows.filter(r => get(r,'Status').toLowerCase() === 'shipped').length;
  const inProgress = rows.filter(r => isActive(r)).length;
  const stillTotal = rows.reduce((s, r) => s + num(r, 'Quantity still to print'), 0);
  const bottles    = rows.reduce((s, r) => s + num(r, 'Quantity'), 0);
  const totalFaulty   = rows.reduce((s, r) => s + num(r, 'Faulty prints'), 0);
  const totalPrinted  = rows.reduce((s, r) => s + num(r, 'Quantity printed'), 0);
  const faultyPct     = totalPrinted > 0 ? ((totalFaulty / totalPrinted) * 100).toFixed(1) + '%' : '—';
  const needSleeve    = rows.filter(r => get(r,'Status').toLowerCase() === 'waiting').length;
  const overdue       = rows.filter(r => isOverdue(r)).length;

  document.getElementById('s-total').textContent       = rows.length;
  document.getElementById('s-shipped').textContent     = shipped;
  document.getElementById('s-in-progress').textContent = inProgress;
  document.getElementById('s-still').textContent       = stillTotal.toLocaleString();
  document.getElementById('s-bottles').textContent     = bottles.toLocaleString();
  document.getElementById('s-faulty').textContent      = faultyPct;
  document.getElementById('s-sleeve').textContent      = needSleeve;
  document.getElementById('s-overdue').textContent     = overdue;
}

// ── Charts ────────────────────────────────────────────────────
function destroyChart(id) {
  if (charts[id]) { charts[id].destroy(); delete charts[id]; }
}

const PALETTE = ['#0077ff','#28a745','#fd7e14','#dc3545','#6f42c1','#f0c22b','#20c997','#e83e8c','#17a2b8','#343a40'];

function renderCharts(rows) {
  // Status distribution
  const statusCounts = {};
  rows.forEach(r => { const s = get(r,'Status') || 'Unknown'; statusCounts[s] = (statusCounts[s]||0)+1; });
  destroyChart('status');
  charts.status = new Chart(document.getElementById('chart-status'), {
    type: 'doughnut',
    data: {
      labels: Object.keys(statusCounts),
      datasets: [{ data: Object.values(statusCounts), backgroundColor: PALETTE }]
    },
    options: { plugins: { legend: { position: 'bottom' } }, cutout: '60%' }
  });

  // Jobs by owner
  const ownerCounts = {};
  rows.forEach(r => { const o = get(r,'Owner') || '—'; ownerCounts[o] = (ownerCounts[o]||0)+1; });
  destroyChart('owner');
  charts.owner = new Chart(document.getElementById('chart-owner'), {
    type: 'bar',
    data: {
      labels: Object.keys(ownerCounts),
      datasets: [{ label: 'Jobs', data: Object.values(ownerCounts), backgroundColor: PALETTE }]
    },
    options: { plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true } } }
  });

  // Monthly jobs added
  const monthly = {};
  rows.forEach(r => {
    const d = parseDate(get(r,'Date added ') || get(r,'Date added'));
    if (!d) return;
    const key = `${d.getFullYear()}/${String(d.getMonth()+1).padStart(2,'0')}`;
    monthly[key] = (monthly[key]||0)+1;
  });
  const sortedMonths = Object.keys(monthly).sort();
  destroyChart('monthly');
  charts.monthly = new Chart(document.getElementById('chart-monthly'), {
    type: 'line',
    data: {
      labels: sortedMonths,
      datasets: [{ label: 'Jobs Added', data: sortedMonths.map(k => monthly[k]), borderColor: '#0077ff', backgroundColor: 'rgba(0,119,255,0.1)', fill: true, tension: 0.3 }]
    },
    options: { plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true } } }
  });

  // Jobs by product type
  const typeCounts = {};
  rows.forEach(r => { const t = get(r,'Soort') || 'Unknown'; typeCounts[t] = (typeCounts[t]||0)+1; });
  const typeEntries = Object.entries(typeCounts).sort((a,b) => b[1]-a[1]);
  const typeColors = typeEntries.map(([t]) => {
    const key = t.toLowerCase();
    const c = Object.entries(TYPE_COLORS).find(([k]) => key.includes(k));
    return c ? c[1].bg.replace(')', ',0.9)').replace('rgb','rgba') : '#e2e8f0';
  });
  destroyChart('colors');
  charts.colors = new Chart(document.getElementById('chart-colors'), {
    type: 'doughnut',
    data: {
      labels: typeEntries.map(e => e[0]),
      datasets: [{ data: typeEntries.map(e => e[1]), backgroundColor: PALETTE }]
    },
    options: { plugins: { legend: { position: 'bottom' } }, cutout: '55%' }
  });

  // Top 10 companies by volume
  const companyVol = {};
  rows.forEach(r => { const c = get(r,'Name_Company'); if (c) companyVol[c] = (companyVol[c]||0)+num(r,'Quantity'); });
  const top10 = Object.entries(companyVol).sort((a,b) => b[1]-a[1]).slice(0,10);
  destroyChart('companies');
  charts.companies = new Chart(document.getElementById('chart-companies'), {
    type: 'bar',
    data: {
      labels: top10.map(e => e[0]),
      datasets: [{ label: 'Items', data: top10.map(e => e[1]), backgroundColor: '#0077ff' }]
    },
    options: { plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true } }, indexAxis: 'y' }
  });

  // Bottles still to print by owner
  const stillByOwner = {};
  rows.forEach(r => { const o = get(r,'Owner') || '—'; stillByOwner[o] = (stillByOwner[o]||0)+num(r,'Quantity still to print'); });
  destroyChart('still-owner');
  charts['still-owner'] = new Chart(document.getElementById('chart-still-owner'), {
    type: 'bar',
    data: {
      labels: Object.keys(stillByOwner),
      datasets: [{ label: 'Still to Print', data: Object.values(stillByOwner), backgroundColor: '#fd7e14' }]
    },
    options: { plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true } } }
  });
}

// ── Date range helpers ────────────────────────────────────────
function getDateRange(period) {
  const now   = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());

  if (period === 'today') {
    return { from: today, to: new Date(today.getTime() + 86399999) };
  }
  if (period === 'week') {
    const day  = today.getDay(); // 0=Sun
    const mon  = new Date(today); mon.setDate(today.getDate() - ((day + 6) % 7));
    const sun  = new Date(mon);   sun.setDate(mon.getDate() + 6);
    return { from: mon, to: new Date(sun.getTime() + 86399999) };
  }
  if (period === 'month') {
    const from = new Date(today.getFullYear(), today.getMonth(), 1);
    const to   = new Date(today.getFullYear(), today.getMonth() + 1, 0, 23, 59, 59);
    return { from, to };
  }
  if (period === 'year') {
    const from = new Date(today.getFullYear(), 0, 1);
    const to   = new Date(today.getFullYear(), 11, 31, 23, 59, 59);
    return { from, to };
  }
  return null;
}

function inDateRange(row, dateField, period, customFrom, customTo) {
  if (!period) return true;
  const d = parseDate(get(row, dateField));
  if (!d) return false;

  if (period === 'custom') {
    const f = customFrom ? new Date(customFrom) : null;
    const t = customTo   ? new Date(customTo + 'T23:59:59') : null;
    if (f && d < f) return false;
    if (t && d > t) return false;
    return true;
  }

  const range = getDateRange(period);
  if (!range) return true;
  return d >= range.from && d <= range.to;
}

// ── All Jobs Table ────────────────────────────────────────────
let ajSort = { key: 'Priority', asc: true };

function populateAllJobsFilters(rows) {
  const statuses = [...new Set(rows.map(r => get(r,'Status')).filter(Boolean))].sort();
  const owners   = [...new Set(rows.map(r => get(r,'Owner')).filter(Boolean))].sort();
  const colors   = [...new Set(rows.map(r => get(r,'Bottle color')).filter(Boolean))].sort();
  const types    = [...new Set(rows.map(r => get(r,'Soort')).filter(Boolean))].sort();

  fill('aj-status', statuses);
  fill('aj-owner',  owners);
  fill('aj-color',  colors);
  fill('aj-type',   types);
}

function fill(id, options) {
  const sel = document.getElementById(id);
  const existing = [...sel.options].map(o => o.value);
  options.forEach(v => { if (!existing.includes(v)) { const o = document.createElement('option'); o.value = o.textContent = v; sel.appendChild(o); }});
}

function getAJFiltered() {
  const search     = document.getElementById('aj-search').value.toLowerCase();
  const status     = document.getElementById('aj-status').value;
  const owner      = document.getElementById('aj-owner').value;
  const color      = document.getElementById('aj-color').value;
  const sleeve     = document.getElementById('aj-sleeve').value;
  const type       = document.getElementById('aj-type').value;
  const dateField  = document.getElementById('aj-date-field').value;
  const period     = document.getElementById('aj-period').value;
  const dateFrom   = document.getElementById('aj-date-from').value;
  const dateTo     = document.getElementById('aj-date-to').value;
  const onlyActive = document.getElementById('aj-only-active').checked;
  const onlyFaulty = document.getElementById('aj-only-faulty').checked;

  return allRows.filter(r => {
    if (search && !(get(r,'Name_Company').toLowerCase().includes(search) || get(r,'Name_Print').toLowerCase().includes(search))) return false;
    if (status && get(r,'Status') !== status) return false;
    if (owner  && get(r,'Owner')  !== owner)  return false;
    if (color  && get(r,'Bottle color') !== color) return false;
    if (sleeve && getCI(r,'sleeve') !== sleeve) return false;
    if (type   && get(r,'Soort') !== type)    return false;
    if (period && !inDateRange(r, dateField, period, dateFrom, dateTo)) return false;
    if (onlyActive && !isActive(r))           return false;
    if (onlyFaulty && num(r,'Faulty prints') === 0) return false;
    return true;
  });
}

function renderAllJobs() {
  let rows = getAJFiltered();

  // Sort
  rows.sort((a, b) => {
    let va = get(a, ajSort.key), vb = get(b, ajSort.key);
    const na = parseFloat(va), nb = parseFloat(vb);
    if (!isNaN(na) && !isNaN(nb)) { va = na; vb = nb; }
    if (va < vb) return ajSort.asc ? -1 : 1;
    if (va > vb) return ajSort.asc ?  1 : -1;
    return 0;
  });

  document.getElementById('aj-count').textContent = `${rows.length} jobs`;

  document.getElementById('aj-body').innerHTML = rows.length === 0
    ? '<tr><td colspan="13">No results.</td></tr>'
    : rows.map(r => {
        const still = num(r,'Quantity still to print');
        const faulty = num(r,'Faulty prints');
        const overdue = isOverdue(r);
        return `<tr class="${!isActive(r) ? 'row-shipped' : ''} ${overdue ? 'row-overdue' : ''}">
          <td><strong>${get(r,'Name_Company')}</strong></td>
          <td class="print-name">${get(r,'Name_Print') || '—'}</td>
          <td>${badge(get(r,'Status'))}</td>
          <td>${get(r,'Owner') || '—'}</td>
          <td>${get(r,'Deadline') || '—'}</td>
          <td>${typeBadge(get(r,'Soort'))}</td>
          <td>${get(r,'Bottle color') || '—'}</td>
          <td>${num(r,'Quantity') || '—'}</td>
          <td>${num(r,'Quantity printed ') || num(r,'Quantity printed') || '—'}</td>
          <td class="${still > 0 ? 'cell-danger' : ''}">${still > 0 ? still : '—'}</td>
          <td class="${faulty > 0 ? 'cell-warn' : ''}">${faulty > 0 ? faulty : '—'}</td>
          <td>${getCI(r,'sleeve') || '—'}</td>
        </tr>`;
      }).join('');
}

function clearFilters() {
  ['aj-search','aj-date-from','aj-date-to'].forEach(id => document.getElementById(id).value = '');
  ['aj-status','aj-owner','aj-color','aj-sleeve','aj-type','aj-period','aj-date-field'].forEach(id => document.getElementById(id).selectedIndex = 0);
  document.getElementById('aj-only-active').checked = false;
  document.getElementById('aj-only-faulty').checked = false;
  document.getElementById('aj-custom-range').style.display = 'none';
  renderAllJobs();
}

// Sortable headers
document.querySelectorAll('#all-jobs-table thead th[data-sort]').forEach(th => {
  th.addEventListener('click', () => {
    const key = th.dataset.sort;
    if (ajSort.key === key) ajSort.asc = !ajSort.asc;
    else { ajSort.key = key; ajSort.asc = true; }
    renderAllJobs();
  });
});

document.getElementById('overview-period').addEventListener('change', function() {
  document.getElementById('overview-custom-range').style.display = this.value === 'custom' ? 'inline-flex' : 'none';
  const overviewRows = getOverviewRows();
  renderStats(overviewRows);
  renderCharts(overviewRows);
});
['overview-date-from','overview-date-to'].forEach(id => {
  document.getElementById(id).addEventListener('change', () => {
    const overviewRows = getOverviewRows();
    renderStats(overviewRows);
    renderCharts(overviewRows);
  });
});

['aj-search','aj-status','aj-owner','aj-color','aj-sleeve','aj-type','aj-date-field','aj-period','aj-date-from','aj-date-to','aj-only-active','aj-only-faulty'].forEach(id => {
  const el = document.getElementById(id);
  el.addEventListener(id.startsWith('aj-only') ? 'change' : 'input', renderAllJobs);
  el.addEventListener('change', renderAllJobs);
});

// Show/hide custom range inputs for All Jobs
document.getElementById('aj-period').addEventListener('change', function() {
  const wrap = document.getElementById('aj-custom-range');
  wrap.style.display = this.value === 'custom' ? 'inline-flex' : 'none';
});

// Show/hide custom range inputs for Active Queue
document.getElementById('aq-period').addEventListener('change', function() {
  const wrap = document.getElementById('aq-custom-range');
  wrap.style.display = this.value === 'custom' ? 'inline-flex' : 'none';
});

// ── Active Queue ──────────────────────────────────────────────
const AQ_SECTIONS = [
  { label: 'Bottles',        colors: TYPE_COLORS['bottle'],        match: s => s === 'bottle' },
  { label: 'Mugs',           colors: TYPE_COLORS['mug'],           match: s => s === 'mug' },
  { label: 'Travel Bottles', colors: TYPE_COLORS['travel bottle'], match: s => s === 'travel bottle' },
  { label: 'Tumblers',       colors: TYPE_COLORS['tumbler'],       match: s => s === 'tumbler' },
  { label: 'Samples',        colors: { bg: '#f1f5f9', text: '#475569' }, match: s => s.includes('sample') },
  { label: 'Other',          colors: { bg: '#f8fafc', text: '#64748b' }, match: s => !['bottle','mug','travel','tumbler','sample'].some(kw => s.includes(kw)) },
];

function statusSortOrder(r) {
  const s = get(r,'Status').toLowerCase();
  if (s.includes('progress')) return 0;
  if (s === 'to print')       return 1;
  if (s === 'waiting')        return 2;
  if (s === 'ready to ship')  return 3;
  return 4;
}

let aqTypeFilter = '';

function renderActiveQueue() {
  const search   = document.getElementById('aq-search').value.toLowerCase();
  const status   = document.getElementById('aq-status').value;
  const period   = document.getElementById('aq-period').value;
  const dateFrom = document.getElementById('aq-date-from').value;
  const dateTo   = document.getElementById('aq-date-to').value;

  const filtered = allRows.filter(r => {
    if (!isActive(r)) return false;
    if (search && !(get(r,'Name_Company').toLowerCase().includes(search) || get(r,'Name_Print').toLowerCase().includes(search))) return false;
    if (status && get(r,'Status') !== status) return false;
    if (period && !inDateRange(r, 'Deadline', period, dateFrom, dateTo)) return false;
    return true;
  });

  // Compute effective display status per row
  // Auto-correct stale status when still=0 but sheet wasn't updated (e.g. quantity entered manually)
  const getDisplayStatus = r => {
    const still = num(r,'Quantity still to print');
    const status = get(r,'Status').toLowerCase();
    if (still <= 0 && (status === 'to print' || status === 'printing progress' || status === '')) {
      const needsSleeve = (get(r,'To sleeve?') || getCI(r,'sleeve')).toLowerCase() === 'yes';
      return needsSleeve ? 'Waiting' : 'Ready to Ship';
    }
    return get(r,'Status');
  };

  const readyRows = filtered.filter(r => getDisplayStatus(r).toLowerCase() === 'ready to ship');
  const activeRows = filtered.filter(r => getDisplayStatus(r).toLowerCase() !== 'ready to ship');

  // Type filter pills (exclude Ready to Ship from counts)
  const tabsEl = document.getElementById('aq-type-tabs');
  tabsEl.innerHTML = `<button class="aq-type-tab${aqTypeFilter === '' ? ' active' : ''}" data-type="">All <span class="aq-tab-count">${activeRows.length}</span></button>` +
    AQ_SECTIONS.map(s => {
      const n = activeRows.filter(r => s.match(get(r,'Soort').toLowerCase())).length;
      if (!n) return '';
      return `<button class="aq-type-tab${aqTypeFilter === s.label ? ' active' : ''}" data-type="${s.label}" style="--tc:${s.colors.text};--tb:${s.colors.bg}">${s.label} <span class="aq-tab-count">${n}</span></button>`;
    }).join('');

  const html = AQ_SECTIONS.filter(s => !aqTypeFilter || s.label === aqTypeFilter).map(section => {
    const rows = activeRows
      .filter(r => section.match(get(r,'Soort').toLowerCase()))
      .sort((a,b) => statusSortOrder(a) - statusSortOrder(b));

    if (rows.length === 0) return '';

    const c = section.colors;
    const rowsHtml = rows.map(r => {
      const d     = parseDate(get(r,'Deadline'));
      const days  = daysFrom(d);
      const still = num(r,'Quantity still to print');
      const displayStatus = getDisplayStatus(r);
      const idx   = allRows.indexOf(r);
      const sleeveVal = (get(r,'To sleeve?') || getCI(r,'sleeve')).toLowerCase();
      const sleeveBtn = sleeveVal !== 'yes' ? '' :
        displayStatus.toLowerCase() === 'ready to ship'
          ? `<button class="btn-sleeve sleeved" data-rowidx="${idx}">✓ Sleeved</button>`
          : `<button class="btn-sleeve" data-rowidx="${idx}">✕ Sleeve</button>`;
      const actionBtns = `<button class="btn-log" data-rowidx="${idx}">✏️ Log</button><button class="btn-log aq-btn-edit" data-rowidx="${idx}" style="background:var(--blue-dim);color:var(--blue);">✎ Edit</button>${sleeveBtn}<button class="btn-reset" data-rowidx="${idx}">↺ Reset</button>`;

      const aqFileUrls = (getCI(r,'file') || getCI(r,'design') || '').split(/[\n,]/).map(u => u.trim()).filter(Boolean);
      const aqFileLink = aqFileUrls.length
        ? aqFileUrls.map((u,i) => `<a href="${u}" target="_blank" rel="noopener" style="color:var(--blue);font-size:12px;text-decoration:none;" title="Open file">📎${aqFileUrls.length > 1 ? ' File '+(i+1) : ' File'}</a>`).join(' ')
        : '';

      const inv  = matchInvoice(get(r,'Name_Company'));
      const card = `<div class="aq-card${isOverdue(r) ? ' overdue' : ''}${aqSelected.has(idx) ? ' row-selected' : ''}" style="--tc:${c.text};--tb:${c.bg}">
        <div class="aq-card-top">
          <div class="aq-card-left">
            <label class="row-check-wrap" onclick="event.stopPropagation()"><input type="checkbox" class="row-select aq-select" data-rowidx="${idx}" ${aqSelected.has(idx) ? 'checked' : ''} /></label>
            <span class="aq-company">${get(r,'Name_Company')}</span>
          </div>
          <div class="aq-badges">${badge(displayStatus)}${invoiceBadge(inv)}</div>
        </div>
        ${get(r,'Name_Print') ? `<div class="aq-print-name">${get(r,'Name_Print')}</div>` : ''}
        ${aqFileLink ? `<div style="margin:4px 0 2px;">${aqFileLink}</div>` : ''}
        <div class="aq-meta">
          <div class="aq-meta-item"><span class="aq-meta-label">Deadline</span><span>${get(r,'Deadline') || '—'}</span></div>
          <div class="aq-meta-item"><span class="aq-meta-label">Days left</span>${daysCell(days)}</div>
          <div class="aq-meta-item"><span class="aq-meta-label">Qty</span><span>${num(r,'Quantity') || '—'}</span></div>
          ${still > 0 ? `<div class="aq-meta-item"><span class="aq-meta-label">Still to print</span><span class="cell-danger">${still}</span></div>` : ''}
          ${get(r,'Bottle color') ? `<div class="aq-meta-item"><span class="aq-meta-label">Color</span><span>${get(r,'Bottle color')}</span></div>` : ''}
          ${get(r,'Lid') ? `<div class="aq-meta-item"><span class="aq-meta-label">Lid</span><span>${get(r,'Lid')}</span></div>` : ''}
        </div>
        <div class="aq-card-actions">${actionBtns}</div>
      </div>`;

      const row = `<tr class="${isOverdue(r) ? 'row-overdue' : ''}${aqSelected.has(idx) ? ' row-selected' : ''}">
        <td><input type="checkbox" class="row-select aq-select" data-rowidx="${idx}" ${aqSelected.has(idx) ? 'checked' : ''} /></td>
        <td><strong>${get(r,'Name_Company')}</strong></td>
        <td class="print-name">${get(r,'Name_Print') || '—'}</td>
        <td>${badge(displayStatus)}</td>
        <td>${invoiceBadge(inv)}</td>
        <td>${typeBadge(get(r,'Soort'))}</td>
        <td>${get(r,'Deadline') || '—'}</td>
        <td>${get(r,'Bottle color') || '—'}</td>
        <td>${get(r,'Lid') || '—'}</td>
        <td>${num(r,'Quantity') || '—'}</td>
        <td class="${still > 0 ? 'cell-danger' : ''}">${still > 0 ? still : '—'}</td>
        <td>${daysCell(days)}</td>
        <td>${aqFileUrls.length ? aqFileUrls.map((u,i) => `<a href="${u}" target="_blank" rel="noopener" style="color:var(--blue);text-decoration:none;">📎${aqFileUrls.length > 1 ? (i+1) : ''}</a>`).join(' ') : '—'}</td>
        <td style="white-space:nowrap">${actionBtns}</td>
      </tr>`;

      return { card, row };
    });

    return `
      <div class="aq-section">
        <div class="aq-section-title" style="background:${c.bg};color:${c.text};">
          <span style="font-size:13px;font-weight:700;">${section.label}</span>
          <span class="aq-section-count" style="color:${c.text};opacity:0.7;">${rows.length} job${rows.length !== 1 ? 's' : ''}</span>
        </div>
        <div class="aq-cards">${rowsHtml.map(x => x.card).join('')}</div>
        <div class="aq-table-wrap table-wrap">
          <table>
            <thead style="--th-bg:${c.text};--th-bg-img:none;"><tr>
              <th></th><th>Company</th><th>Print Name</th><th>Status</th><th>Invoice</th>
              <th>Type</th><th>Deadline</th><th>Color</th><th>Lid</th>
              <th>Qty</th><th>Still to Print</th><th>Days Left</th><th>Files</th><th>Actions</th>
            </tr></thead>
            <tbody>${rowsHtml.map(x => x.row).join('')}</tbody>
          </table>
        </div>
      </div>`;
  }).join('');

  // Ready to Ship section
  const rtsFiltered = aqTypeFilter
    ? readyRows.filter(r => AQ_SECTIONS.find(s => s.label === aqTypeFilter)?.match(get(r,'Soort').toLowerCase()))
    : readyRows;

  const rtsHtml = rtsFiltered.length ? (() => {
    const rowsHtml = rtsFiltered.map(r => {
      const d     = parseDate(get(r,'Deadline'));
      const days  = daysFrom(d);
      const idx   = allRows.indexOf(r);
      const sleeveVal = (get(r,'To sleeve?') || getCI(r,'sleeve')).toLowerCase();
      const sleeveBtn = sleeveVal !== 'yes' ? '' :
        `<button class="btn-sleeve sleeved" data-rowidx="${idx}">✓ Sleeved</button>`;
      const actionBtns = `<button class="btn-log" data-rowidx="${idx}">✏️ Log</button><button class="btn-log aq-btn-edit" data-rowidx="${idx}" style="background:var(--blue-dim);color:var(--blue);">✎ Edit</button>${sleeveBtn}<button class="btn-ship" data-rowidx="${idx}" style="background:#15803d;color:#fff;border:none;border-radius:var(--radius-sm);padding:5px 12px;font-size:12px;font-weight:600;cursor:pointer;">✓ Ship</button><button class="btn-reset" data-rowidx="${idx}">↺ Reset</button>`;
      const aqFileUrls = (getCI(r,'file') || getCI(r,'design') || '').split(/[\n,]/).map(u => u.trim()).filter(Boolean);
      const aqFileLink = aqFileUrls.length
        ? aqFileUrls.map((u,i) => `<a href="${u}" target="_blank" rel="noopener" style="color:var(--blue);font-size:12px;text-decoration:none;">📎${aqFileUrls.length > 1 ? ' File '+(i+1) : ' File'}</a>`).join(' ')
        : '';
      const card = `<div class="aq-card${aqSelected.has(idx) ? ' row-selected' : ''}" style="--tc:#15803d;--tb:#dcfce7">
        <div class="aq-card-top">
          <div class="aq-card-left">
            <label class="row-check-wrap" onclick="event.stopPropagation()"><input type="checkbox" class="row-select aq-select" data-rowidx="${idx}" ${aqSelected.has(idx) ? 'checked' : ''} /></label>
            <span class="aq-company">${get(r,'Name_Company')}</span>
          </div>
          ${badge('Ready to Ship')}
        </div>
        ${get(r,'Name_Print') ? `<div class="aq-print-name">${get(r,'Name_Print')}</div>` : ''}
        ${aqFileLink ? `<div style="margin:4px 0 2px;">${aqFileLink}</div>` : ''}
        <div class="aq-meta">
          <div class="aq-meta-item"><span class="aq-meta-label">Type</span>${typeBadge(get(r,'Soort'))}</div>
          ${get(r,'Deadline') ? `<div class="aq-meta-item"><span class="aq-meta-label">Deadline</span><span>${get(r,'Deadline')}</span></div>` : ''}
          ${get(r,'Bottle color') ? `<div class="aq-meta-item"><span class="aq-meta-label">Color</span><span>${get(r,'Bottle color')}</span></div>` : ''}
          ${get(r,'Lid') ? `<div class="aq-meta-item"><span class="aq-meta-label">Lid</span><span>${get(r,'Lid')}</span></div>` : ''}
        </div>
        <div class="aq-card-actions">${actionBtns}</div>
      </div>`;
      const inv   = matchInvoice(get(r,'Name_Company'));
      const still = num(r,'Quantity still to print');
      const row = `<tr class="${aqSelected.has(idx) ? 'row-selected' : ''}">
        <td><input type="checkbox" class="row-select aq-select" data-rowidx="${idx}" ${aqSelected.has(idx) ? 'checked' : ''} /></td>
        <td><strong>${get(r,'Name_Company')}</strong></td>
        <td class="print-name">${get(r,'Name_Print') || '—'}</td>
        <td>${badge('Ready to Ship')}</td>
        <td>${invoiceBadge(inv)}</td>
        <td>${typeBadge(get(r,'Soort'))}</td>
        <td>${get(r,'Deadline') || '—'}</td>
        <td>${get(r,'Bottle color') || '—'}</td>
        <td>${get(r,'Lid') || '—'}</td>
        <td>${num(r,'Quantity') || '—'}</td>
        <td class="${still > 0 ? 'cell-danger' : ''}">${still > 0 ? still : '—'}</td>
        <td>${daysCell(days)}</td>
        <td style="white-space:nowrap">${actionBtns}</td>
      </tr>`;
      return { card, row };
    });
    return `<div class="aq-section">
      <div class="aq-section-title" style="background:#dcfce7;color:#15803d;">
        <span style="font-size:13px;font-weight:700;">✓ Ready to Ship</span>
        <span class="aq-section-count" style="color:#15803d;opacity:0.7;">${rtsFiltered.length} job${rtsFiltered.length !== 1 ? 's' : ''}</span>
      </div>
      <div class="aq-cards">${rowsHtml.map(x => x.card).join('')}</div>
      <div class="aq-table-wrap aq-rts-table table-wrap">
        <table>
          <thead style="background:#15803d;background-image:none;"><tr>
            <th></th><th>Company</th><th>Print Name</th><th>Status</th><th>Invoice</th>
            <th>Type</th><th>Deadline</th><th>Color</th><th>Lid</th>
            <th>Qty</th><th>Still to Print</th><th>Days Left</th><th>Actions</th>
          </tr></thead>
          <tbody>${rowsHtml.map(x => x.row).join('')}</tbody>
        </table>
      </div>
    </div>`;
  })() : '';

  document.getElementById('aq-sections').innerHTML = (html + rtsHtml) ||
    '<p style="color:#94a3b8;padding:20px;">No active jobs match the current filters.</p>';
}

['aq-search','aq-status','aq-period','aq-date-from','aq-date-to'].forEach(id => {
  document.getElementById(id).addEventListener('input', renderActiveQueue);
  document.getElementById(id).addEventListener('change', renderActiveQueue);
});

document.getElementById('aq-type-tabs').addEventListener('click', e => {
  const btn = e.target.closest('.aq-type-tab');
  if (btn) { aqTypeFilter = btn.dataset.type; renderActiveQueue(); }
});

// Log + Sleeve + Reset buttons — event delegation on section container
document.getElementById('tab-active-queue').addEventListener('click', function(e) {
  const selectCb  = e.target.closest('.aq-select');
  const logBtn    = e.target.closest('.btn-log');
  const sleeveBtn = e.target.closest('.btn-sleeve');
  const resetBtn  = e.target.closest('.btn-reset');
  const shipBtn   = e.target.closest('.btn-ship');
  if (selectCb && e.target.tagName === 'INPUT') {
    const idx = parseInt(selectCb.dataset.rowidx);
    if (selectCb.checked) aqSelected.add(idx); else aqSelected.delete(idx);
    updateSelectionBar();
    return;
  }
  const aqEditBtn = e.target.closest('.aq-btn-edit');
  if (logBtn && !aqEditBtn)   openPrintModal(parseInt(logBtn.dataset.rowidx));
  if (aqEditBtn) openEditJobModal(parseInt(aqEditBtn.dataset.rowidx), 'active');
  if (resetBtn) resetJob(parseInt(resetBtn.dataset.rowidx));
  if (shipBtn)  shipJob(parseInt(shipBtn.dataset.rowidx), shipBtn);
  if (sleeveBtn) {
    if (sleeveBtn.classList.contains('sleeved')) unsleeveJob(parseInt(sleeveBtn.dataset.rowidx), sleeveBtn);
    else markSleeved(parseInt(sleeveBtn.dataset.rowidx), sleeveBtn);
  }
});

// ── By Company ────────────────────────────────────────────────

// ── Reports ───────────────────────────────────────────────────
function renderReports() {
  const today    = new Date(); today.setHours(0,0,0,0);
  const in7      = new Date(today); in7.setDate(today.getDate()+7);
  const rpOwner  = document.getElementById('rp-owner').value;
  const rpPeriod = document.getElementById('rp-period').value;
  const rpFrom   = document.getElementById('rp-date-from').value;
  const rpTo     = document.getElementById('rp-date-to').value;

  // Pre-filter rows by owner + period (on Date added)
  const base = allRows.filter(r => {
    if (rpOwner  && get(r,'Owner') !== rpOwner) return false;
    if (rpPeriod && !inDateRange(r, 'Date added ', rpPeriod, rpFrom, rpTo)) return false;
    return true;
  });

  // Overdue
  const overdue = base.filter(r => isOverdue(r)).sort((a,b) => {
    const da = parseDate(get(a,'Deadline')), db = parseDate(get(b,'Deadline'));
    return (da||0) - (db||0);
  });
  document.getElementById('rep-overdue').innerHTML = overdue.length === 0
    ? '<tr><td colspan="7" class="cell-ok">No overdue jobs!</td></tr>'
    : overdue.map(r => {
        const d = parseDate(get(r,'Deadline'));
        const late = d ? Math.round((today - d)/86400000) : '?';
        return `<tr><td>${get(r,'Name_Company')}</td>
          <td>${badge(get(r,'Status'))}</td><td>${get(r,'Owner')}</td>
          <td>${get(r,'Deadline')}</td><td class="cell-danger">${late}d</td>
          <td>${num(r,'Quantity')}</td></tr>`;
      }).join('');

  // Faulty prints
  const faulty = base.filter(r => num(r,'Faulty prints') > 0).sort((a,b) => num(b,'Faulty prints') - num(a,'Faulty prints'));
  document.getElementById('rep-faulty').innerHTML = faulty.length === 0
    ? '<tr><td colspan="7">No faulty prints recorded.</td></tr>'
    : faulty.map(r => {
        const f = num(r,'Faulty prints'), q = num(r,'Quantity');
        const pct = q ? ((f/q)*100).toFixed(1)+'%' : '—';
        return `<tr><td>${get(r,'Name_Company')}</td>
          <td class="print-name">${get(r,'Name_Print') || '—'}</td><td>${get(r,'Owner')}</td>
          <td>${q}</td><td class="cell-danger">${f}</td><td class="cell-warn">${pct}</td></tr>`;
      }).join('');

  // Needs sleeving
  const sleeve = base.filter(r => getCI(r,'sleeve').toLowerCase() === 'yes' && get(r,'Gesleeved?') !== 'Yes' && isActive(r));
  document.getElementById('rep-sleeve').innerHTML = sleeve.length === 0
    ? '<tr><td colspan="6" class="cell-ok">All sleeveable jobs are done!</td></tr>'
    : sleeve.map(r => `<tr><td>${get(r,'Name_Company')}</td>
        <td class="print-name">${get(r,'Name_Print') || '—'}</td><td>${badge(get(r,'Status'))}</td>
        <td>${get(r,'Owner')}</td><td>${num(r,'Quantity')}</td></tr>`).join('');

  // To Print
  const toPrint = base.filter(r => get(r,'Status').toLowerCase() === 'to print')
    .sort((a,b) => num(a,'Priority') - num(b,'Priority'));
  document.getElementById('rep-to-print').innerHTML = toPrint.length === 0
    ? '<tr><td colspan="8">No jobs queued to print.</td></tr>'
    : toPrint.map(r => `<tr><td>${get(r,'Name_Company')}</td>
        <td class="print-name">${get(r,'Name_Print') || '—'}</td><td>${get(r,'Owner')}</td>
        <td>${get(r,'Deadline') || '—'}</td><td>${typeBadge(get(r,'Soort'))}</td>
        <td>${get(r,'Bottle color') || '—'}</td><td>${num(r,'Quantity')}</td></tr>`).join('');

  // Waiting
  const waiting = base.filter(r => get(r,'Status').toLowerCase() === 'waiting')
    .sort((a,b) => num(a,'Priority') - num(b,'Priority'));
  document.getElementById('rep-waiting').innerHTML = waiting.length === 0
    ? '<tr><td colspan="6">No waiting jobs.</td></tr>'
    : waiting.map(r => `<tr><td>${get(r,'Name_Company')}</td>
        <td class="print-name">${get(r,'Name_Print') || '—'}</td><td>${get(r,'Owner')}</td>
        <td>${get(r,'Deadline') || '—'}</td><td>${num(r,'Quantity')}</td></tr>`).join('');

  // Print ready, not shipped
  const readyToShip = base.filter(r => {
    const printReady = get(r,'Printing ready?') === 'Yes';
    const sleeveOk   = getCI(r,'sleeve').toLowerCase() !== 'yes' || get(r,'Gesleeved?') === 'Yes';
    return printReady && sleeveOk && isActive(r);
  });
  document.getElementById('rep-ready-ship').innerHTML = readyToShip.length === 0
    ? '<tr><td colspan="7">Nothing ready for shipment.</td></tr>'
    : readyToShip.map(r => `<tr><td>${get(r,'Name_Company')}</td>
        <td class="print-name">${get(r,'Name_Print') || '—'}</td><td>${get(r,'Owner')}</td>
        <td>${typeBadge(get(r,'Soort'))}</td>
        <td>${num(r,'Quantity')}</td><td>${getCI(r,'sleeve').toLowerCase() === 'yes' ? '✅' : '—'}</td></tr>`).join('');

  // Deadlines this week
  const thisWeek = base.filter(r => {
    const d = parseDate(get(r,'Deadline'));
    return d && d >= today && d <= in7 && isActive(r);
  }).sort((a,b) => parseDate(get(a,'Deadline')) - parseDate(get(b,'Deadline')));
  document.getElementById('rep-this-week').innerHTML = thisWeek.length === 0
    ? '<tr><td colspan="7">No deadlines in the next 7 days.</td></tr>'
    : thisWeek.map(r => {
        const d = parseDate(get(r,'Deadline'));
        return `<tr><td>${get(r,'Name_Company')}</td>
          <td>${badge(get(r,'Status'))}</td><td>${get(r,'Owner')}</td>
          <td>${get(r,'Deadline')}</td><td>${daysCell(daysFrom(d))}</td>
          <td>${num(r,'Quantity')}</td></tr>`;
      }).join('');

  // Recently shipped (last 30 days)
  const ago30 = new Date(today); ago30.setDate(today.getDate()-30);
  const recentShipped = base.filter(r => {
    const d = parseDate(get(r,'Shipped'));
    return d && d >= ago30 && get(r,'Status').toLowerCase() === 'shipped';
  }).sort((a,b) => parseDate(get(b,'Shipped')) - parseDate(get(a,'Shipped')));
  document.getElementById('rep-recent-shipped').innerHTML = recentShipped.length === 0
    ? '<tr><td colspan="6">No recent shipments.</td></tr>'
    : recentShipped.map(r => `<tr><td>${get(r,'Name_Company')}</td>
        <td class="print-name">${get(r,'Name_Print') || '—'}</td><td>${get(r,'Owner')}</td>
        <td>${get(r,'Shipped') || '—'}</td><td>${num(r,'Quantity')}</td></tr>`).join('');
}

// ── Printed Report ────────────────────────────────────────────
function filterPrintLog() {
  const period = document.getElementById('pr-period').value;
  const from   = document.getElementById('pr-date-from').value;
  const to     = document.getElementById('pr-date-to').value;
  const owner  = document.getElementById('pr-owner').value;

  return printLogRows.filter(r => {
    if (owner && r['Owner'] !== owner) return false;
    if (!period) return true;
    const d = parseDate(r['Date']);
    if (!d) return false;
    if (period === 'custom') {
      const f = from ? new Date(from) : null;
      const t = to   ? new Date(to + 'T23:59:59') : null;
      if (f && d < f) return false;
      if (t && d > t) return false;
      return true;
    }
    const range = getDateRange(period);
    if (!range) return true;
    return d >= range.from && d <= range.to;
  }).sort((a, b) => (parseDate(b['Date']) || 0) - (parseDate(a['Date']) || 0));
}

function renderPrintedReport() {
  if (printLogRows.length === 0) {
    document.getElementById('pr-summary').innerHTML =
      '<div class="pr-stat" style="flex:1;text-align:center;color:var(--text-2);font-size:13px;padding:16px;">' +
      'Nog geen printlog-regels gevonden. Vul een aantal in via het print-log icoon en laad de pagina opnieuw.' +
      '</div>';
    document.getElementById('rep-printed').innerHTML =
      '<tr><td colspan="7" style="text-align:center;color:var(--text-2);padding:20px;">Geen data</td></tr>';
    return;
  }

  const rows = filterPrintLog();

  const totalEntries = rows.length;
  const totalItems   = rows.reduce((s, r) => s + (parseInt(r['Quantity']) || 0), 0);
  const byOwner      = {};
  rows.forEach(r => {
    const o = r['Owner'] || 'Onbekend';
    byOwner[o] = (byOwner[o] || 0) + (parseInt(r['Quantity']) || 0);
  });

  document.getElementById('pr-summary').innerHTML =
    `<div class="pr-stat"><div class="pr-stat-value">${totalEntries}</div><div class="pr-stat-label">Print-runs</div></div>` +
    `<div class="pr-stat"><div class="pr-stat-value">${totalItems.toLocaleString()}</div><div class="pr-stat-label">Stuks geprint</div></div>` +
    Object.entries(byOwner).sort((a, b) => b[1] - a[1]).map(([o, q]) =>
      `<div class="pr-stat"><div class="pr-stat-value" style="font-size:20px;">${q.toLocaleString()}</div><div class="pr-stat-label">${o}</div></div>`
    ).join('');

  document.getElementById('rep-printed').innerHTML = rows.length === 0
    ? '<tr><td colspan="7" style="text-align:center;color:var(--text-2);padding:20px;">Geen geprinte jobs voor deze periode.</td></tr>'
    : rows.map(r => `<tr>
        <td>${r['Company'] || '—'}</td>
        <td class="print-name">${r['Print Name'] || '—'}</td>
        <td>${r['Owner'] || '—'}</td>
        <td>${typeBadge(r['Type'] || '')}</td>
        <td>${r['Date'] || '—'}</td>
        <td>${parseInt(r['Quantity']) || 0}</td>
      </tr>`).join('');
}

// ── Populate filter dropdowns ─────────────────────────────────
function populateSecondaryFilters(rows) {
  const active   = rows.filter(r => isActive(r));
  const statuses = [...new Set(active.map(r => get(r,'Status')).filter(Boolean))].sort();
  const owners   = [...new Set(rows.map(r => get(r,'Owner')).filter(Boolean))].sort();
  fill('aq-status', statuses);
  fill('rp-owner',  owners);
  // pr-owner comes from printLogRows, not allRows
  const logOwners = [...new Set(printLogRows.map(r => r['Owner']).filter(Boolean))].sort();
  fill('pr-owner', logOwners.length ? logOwners : owners);
}

['rp-owner','rp-period','rp-date-from','rp-date-to'].forEach(id => {
  document.getElementById(id).addEventListener('input', renderReports);
  document.getElementById(id).addEventListener('change', renderReports);
});
document.getElementById('rp-period').addEventListener('change', function() {
  document.getElementById('rp-custom-range').style.display = this.value === 'custom' ? 'inline-flex' : 'none';
});

// Printed report event listeners
document.getElementById('pr-period').addEventListener('change', function() {
  document.getElementById('pr-custom-range').style.display = this.value === 'custom' ? 'inline-flex' : 'none';
  renderPrintedReport();
});
['pr-date-from','pr-date-to','pr-owner'].forEach(id => {
  document.getElementById(id).addEventListener('change', renderPrintedReport);
});

// ── Shipping History ──────────────────────────────────────────

function trackingUrl(awb, carrier) {
  if (!awb) return null;
  const c = (carrier || '').toLowerCase();
  if (c.includes('dhl'))   return `https://www.dhl.com/nl-nl/home/tracking.html?tracking-id=${awb}`;
  if (c.includes('dpd'))   return `https://tracking.dpd.de/status/nl_NL/parcel/${awb}`;
  if (c.includes('ups'))   return `https://www.ups.com/track?tracknum=${awb}`;
  if (c.includes('fedex')) return `https://www.fedex.com/apps/fedextrack/?tracknumbers=${awb}`;
  if (c.includes('bpost') || c.includes('post')) return `https://track.bpost.cloud/btr/web/#/search?itemCode=${awb}`;
  // Fallback: detect by AWB format
  if (/^1Z/i.test(awb))   return `https://www.ups.com/track?tracknum=${awb}`;
  if (/^JVGL/i.test(awb)) return `https://www.dhl.com/nl-nl/home/tracking.html?tracking-id=${awb}`;
  if (/^\d{13}$/.test(awb)) return `https://tracking.dpd.de/status/nl_NL/parcel/${awb}`;
  return null;
}

// Parse YYYY-MM-DD dates from the courier export
function parseDateISO(str) {
  if (!str) return null;
  const parts = str.split('-');
  if (parts.length !== 3) return null;
  const d = new Date(+parts[0], +parts[1] - 1, +parts[2]);
  return isNaN(d) ? null : d;
}

// Normalize a company name for fuzzy matching:
// lowercase, strip legal suffixes & common words, remove non-alphanumeric
function normalizeName(name) {
  return (name || '').toLowerCase()
    .replace(/\b(gmbh|b\.v\.|bv|nv|ltd|llc|inc|bvba|vzw|srl|sarl|ag|sa|plc|co\.|exhibition|exhibitions|hall|winkel|shop|store|museum|stichting|foundation|group|groep|international|intl)\b/g, '')
    .replace(/[^a-z0-9]/g, '')
    .trim();
}

function calcConfidence(matchMethod, daysDiff) {
  if (matchMethod === 'reference') return 100;
  if (daysDiff <= 3)  return 95;
  if (daysDiff <= 7)  return 90;
  if (daysDiff <= 14) return 80;
  if (daysDiff <= 30) return 70;
  if (daysDiff <= 60) return 45; // possible match — goes to review
  return 0;
}

function confidenceBadge(score) {
  if (score === 100) return `<span class="badge b-shipped" style="font-size:10px;">100% ref</span>`;
  if (score >= 90)   return `<span class="badge b-shipped" style="font-size:10px;">${score}%</span>`;
  if (score >= 70)   return `<span class="badge b-waiting" style="font-size:10px;">${score}%</span>`;
  return               `<span class="badge b-progress" style="font-size:10px;">${score}%</span>`;
}

function buildShippedRows(mainRows, shipRows) {
  const shippedJobs = mainRows.filter(r => get(r,'Status').toLowerCase() === 'shipped');

  const jobByPriority = {};
  shippedJobs.forEach(r => { if (get(r,'Priority')) jobByPriority[get(r,'Priority')] = r; });

  const jobsByNorm = {};
  shippedJobs.forEach(r => {
    const norm = normalizeName(get(r,'Name_Company'));
    if (!norm || norm.length < 4) return;
    if (!jobsByNorm[norm]) jobsByNorm[norm] = [];
    jobsByNorm[norm].push(r);
  });

  // Find candidates by name, within a given max day window
  function findCandidates(recipientName, shipDate, maxDays) {
    const normShip = normalizeName(recipientName);
    if (!normShip || normShip.length < 4) return [];
    const candidates = [];
    for (const [normJob, jobs] of Object.entries(jobsByNorm)) {
      if (normShip.includes(normJob) || normJob.includes(normShip)) {
        jobs.forEach(job => {
          const jobShipped = parseDate(getCI(job,'shipped'));
          // If no shipped date on job, allow name match with a high daysDiff (sorted last)
          const daysDiff = jobShipped ? Math.abs((shipDate - jobShipped) / 86400000) : 999;
          if (daysDiff <= maxDays || !jobShipped) candidates.push({ job, normJob, daysDiff });
        });
      }
    }
    candidates.sort((a, b) => a.daysDiff - b.daysDiff || b.normJob.length - a.normJob.length);
    return candidates;
  }

  function gs(s, key) { return getCI(s, key); } // shorthand for shipping row field access

  function buildRow(s, job, matchMethod, daysDiff, shipDate) {
    const addedDate  = parseDate(get(job,'Date added ') || get(job,'Date added'));
    const turnaround = addedDate ? Math.round((shipDate - addedDate) / 86400000) : null;
    const priceRaw   = gs(s,'prijs').replace(',','.').replace(/[^0-9.]/g,'');
    const plaats     = gs(s,'plaats');
    const land       = gs(s,'land');
    return {
      ordernummer:  gs(s,'ordernummer'),
      priority:     get(job,'Priority'),
      company:      get(job,'Name_Company'),
      recipient:    gs(s,'bedrijfsnaam'),
      printName:    get(job,'Name_Print'),
      owner:        get(job,'Owner'),
      dateAdded:    addedDate,
      shipDate,
      turnaround,
      type:         get(job,'Soort'),
      color:        get(job,'Bottle color'),
      quantity:     num(job,'Quantity'),
      faulty:       num(job,'Faulty prints'),
      carrier:      gs(s,'vervoerder'),
      destination:  [plaats, land].filter(Boolean).join(', '),
      country:      land,
      price:        parseFloat(priceRaw) || null,
      tracking:     gs(s,'awb') || gs(s,'tracking'),
      status:       gs(s,'bericht'),
      matchMethod,
      confidence:   calcConfidence(matchMethod, daysDiff),
    };
  }

  const matched        = [];
  const matchedJobKeys = new Set(); // Priority keys of jobs that got a shipment match

  shipRows.forEach(s => {
    const shipDate = parseDate(getCI(s,'datum'));
    if (!shipDate) return;

    // 1. Referentie = Priority (100% reliable)
    const ref = getCI(s,'referentie').trim();
    if (ref && /^\d+$/.test(ref) && jobByPriority[ref]) {
      matched.push(buildRow(s, jobByPriority[ref], 'reference', 0, shipDate));
      matchedJobKeys.add(ref);
      return;
    }

    // 2. Name match within 30 days → confirmed
    const close = findCandidates(getCI(s,'bedrijfsnaam'), shipDate, 30);
    if (close.length) {
      matched.push(buildRow(s, close[0].job, 'name', close[0].daysDiff, shipDate));
      matchedJobKeys.add(get(close[0].job, 'Priority'));
      return;
    }
  });

  // Shipped jobs in the workfile with no matching shipment record
  const review = shippedJobs
    .filter(job => !matchedJobKeys.has(get(job,'Priority')))
    .map(job => ({
      priority:  get(job,'Priority'),
      company:   get(job,'Name_Company'),
      printName: get(job,'Name_Print'),
      owner:     get(job,'Owner'),
      dateAdded: parseDate(get(job,'Date added ') || get(job,'Date added')),
      deadline:  parseDate(get(job,'Deadline')),
      shippedDate: parseDate(get(job,'Shipped')),
      quantity:  num(job,'Quantity'),
    }));

  // Merge multiple package rows for the same shipment into one
  const shipmentMap = new Map();
  matched.forEach(r => {
    // Group key: ordernummer if present, else company + date
    const key = r.ordernummer
      ? r.ordernummer
      : normalizeName(r.company) + '|' + r.shipDate.toDateString();
    if (shipmentMap.has(key)) {
      const existing = shipmentMap.get(key);
      existing.price    = (existing.price || 0) + (r.price || 0);
      existing.boxes    = (existing.boxes || 1) + 1;
      // Keep best tracking/carrier if not already set
      if (!existing.tracking && r.tracking) existing.tracking = r.tracking;
    } else {
      shipmentMap.set(key, { ...r, boxes: 1 });
    }
  });
  const deduped = Array.from(shipmentMap.values());

  return {
    matched: deduped.sort((a,b) => b.shipDate - a.shipDate),
    review:  review.sort((a,b) => (b.shippedDate||0) - (a.shippedDate||0)),
  };
}

function turnaroundColor(days) {
  if (days === null) return '—';
  if (days <= 7)  return `<span class="cell-ok">${days}d</span>`;
  if (days <= 21) return `<span style="color:#f97316;font-weight:600;">${days}d</span>`;
  return `<span class="cell-danger">${days}d</span>`;
}

function fmtDate(d) {
  if (!d) return '—';
  return `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`;
}

function renderShipping() {
  const search       = document.getElementById('sh-search').value.toLowerCase();
  const owner        = document.getElementById('sh-owner').value;
  const period       = document.getElementById('sh-period').value;
  const dateFrom     = document.getElementById('sh-date-from').value;
  const dateTo       = document.getElementById('sh-date-to').value;
  const globalPeriod = document.getElementById('sh-global-period')?.value || '';

  // Filter
  const rows = shippedRows.filter(r => {
    if (globalPeriod) {
      if (globalPeriod === 'quarter') {
        const cutoff = new Date(); cutoff.setMonth(cutoff.getMonth() - 3);
        if (r.shipDate < cutoff) return false;
      } else {
        const range = getDateRange(globalPeriod);
        if (range && (r.shipDate < range.from || r.shipDate > range.to)) return false;
      }
    }
    if (search && !(r.company.toLowerCase().includes(search) || r.printName.toLowerCase().includes(search))) return false;
    if (owner && r.owner !== owner) return false;
    if (period) {
      if (period === 'custom') {
        const f = dateFrom ? new Date(dateFrom) : null;
        const t = dateTo   ? new Date(dateTo + 'T23:59:59') : null;
        if (f && r.shipDate < f) return false;
        if (t && r.shipDate > t) return false;
      } else {
        const range = getDateRange(period);
        if (range && (r.shipDate < range.from || r.shipDate > range.to)) return false;
      }
    }
    return true;
  });

  // Stats
  const totalBottles   = rows.reduce((s,r) => s + r.quantity, 0);
  const withTurnaround = rows.filter(r => r.turnaround !== null && r.turnaround >= 0);
  const avgDays        = withTurnaround.length ? Math.round(withTurnaround.reduce((s,r) => s + r.turnaround, 0) / withTurnaround.length) : null;
  const fastest        = withTurnaround.length ? Math.min(...withTurnaround.map(r => r.turnaround)) : null;
  const slowest        = withTurnaround.length ? Math.max(...withTurnaround.map(r => r.turnaround)) : null;
  const companies      = new Set(rows.map(r => r.company)).size;
  const totalCost      = rows.reduce((s,r) => s + (r.price || 0), 0);
  const byRef          = rows.filter(r => r.matchMethod === 'reference').length;

  document.getElementById('sh-count').textContent    = rows.length;
  document.getElementById('sh-bottles').textContent  = totalBottles.toLocaleString();
  document.getElementById('sh-avg-days').textContent = avgDays !== null ? avgDays + 'd' : '—';
  document.getElementById('sh-fastest').textContent  = fastest !== null ? fastest + 'd' : '—';
  document.getElementById('sh-slowest').textContent  = slowest !== null ? slowest + 'd' : '—';
  document.getElementById('sh-companies').textContent = companies;
  document.getElementById('sh-cost').textContent     = '€' + totalCost.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  document.getElementById('sh-by-ref').textContent   = byRef + ' / ' + rows.length;

  // Charts
  // Weekly shipments
  const weekly = {};
  rows.forEach(r => {
    const d = r.shipDate;
    const mon = new Date(d); mon.setDate(d.getDate() - ((d.getDay()+6)%7));
    const key = `${mon.getFullYear()}-W${String(Math.ceil((mon.getDate())/7)).padStart(2,'0')}-${String(mon.getMonth()+1).padStart(2,'0')}-${String(mon.getDate()).padStart(2,'0')}`;
    weekly[key] = (weekly[key]||0) + 1;
  });
  const weekKeys = Object.keys(weekly).sort();
  destroyChart('ship-weekly');
  charts['ship-weekly'] = new Chart(document.getElementById('chart-ship-weekly'), {
    type: 'bar',
    data: { labels: weekKeys.map(k => k.slice(k.indexOf('-W')+2)), datasets: [{ label: 'Shipments', data: weekKeys.map(k => weekly[k]), backgroundColor: '#22c55e' }] },
    options: { plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true } } }
  });

  // Monthly bottles shipped
  const monthly = {};
  rows.forEach(r => {
    const key = `${r.shipDate.getFullYear()}/${String(r.shipDate.getMonth()+1).padStart(2,'0')}`;
    monthly[key] = (monthly[key]||0) + r.quantity;
  });
  const monthKeys = Object.keys(monthly).sort();
  destroyChart('ship-monthly');
  charts['ship-monthly'] = new Chart(document.getElementById('chart-ship-monthly'), {
    type: 'bar',
    data: { labels: monthKeys, datasets: [{ label: 'Items', data: monthKeys.map(k => monthly[k]), backgroundColor: '#3b82f6' }] },
    options: { plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true } } }
  });

  // Avg turnaround per month
  const taMonth = {};
  withTurnaround.filter(r => {
    if (!period) return true;
    if (period === 'custom') {
      const f = dateFrom ? new Date(dateFrom) : null;
      const t = dateTo ? new Date(dateTo + 'T23:59:59') : null;
      if (f && r.shipDate < f) return false;
      if (t && r.shipDate > t) return false;
      return true;
    }
    const range = getDateRange(period);
    return !range || (r.shipDate >= range.from && r.shipDate <= range.to);
  }).forEach(r => {
    const key = `${r.shipDate.getFullYear()}/${String(r.shipDate.getMonth()+1).padStart(2,'0')}`;
    if (!taMonth[key]) taMonth[key] = [];
    taMonth[key].push(r.turnaround);
  });
  const taMonthKeys = Object.keys(taMonth).sort();
  destroyChart('turnaround-monthly');
  charts['turnaround-monthly'] = new Chart(document.getElementById('chart-turnaround-monthly'), {
    type: 'line',
    data: {
      labels: taMonthKeys,
      datasets: [{ label: 'Avg days', data: taMonthKeys.map(k => Math.round(taMonth[k].reduce((a,b)=>a+b,0)/taMonth[k].length)), borderColor: '#f97316', backgroundColor: 'rgba(249,115,22,0.1)', fill: true, tension: 0.3 }]
    },
    options: { plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true } } }
  });

  // Avg turnaround per owner
  const taOwner = {};
  withTurnaround.forEach(r => {
    if (!taOwner[r.owner]) taOwner[r.owner] = [];
    taOwner[r.owner].push(r.turnaround);
  });
  destroyChart('turnaround-owner');
  charts['turnaround-owner'] = new Chart(document.getElementById('chart-turnaround-owner'), {
    type: 'bar',
    data: {
      labels: Object.keys(taOwner),
      datasets: [{ label: 'Avg days', data: Object.values(taOwner).map(arr => Math.round(arr.reduce((a,b)=>a+b,0)/arr.length)), backgroundColor: PALETTE }]
    },
    options: { plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true } } }
  });

  // Carrier breakdown
  const carriers = {};
  rows.forEach(r => { const c = r.carrier || 'Unknown'; carriers[c] = (carriers[c]||0)+1; });
  const carrierEntries = Object.entries(carriers).sort((a,b) => b[1]-a[1]);
  destroyChart('ship-carrier');
  charts['ship-carrier'] = new Chart(document.getElementById('chart-ship-carrier'), {
    type: 'doughnut',
    data: { labels: carrierEntries.map(e=>e[0]), datasets: [{ data: carrierEntries.map(e=>e[1]), backgroundColor: PALETTE }] },
    options: { plugins: { legend: { position: 'bottom' } }, cutout: '55%' }
  });

  // Shipping cost per month
  const costMonth = {};
  rows.filter(r => r.price).forEach(r => {
    const key = `${r.shipDate.getFullYear()}/${String(r.shipDate.getMonth()+1).padStart(2,'0')}`;
    costMonth[key] = (costMonth[key]||0) + r.price;
  });
  const costKeys = Object.keys(costMonth).sort();
  destroyChart('ship-cost-monthly');
  charts['ship-cost-monthly'] = new Chart(document.getElementById('chart-ship-cost'), {
    type: 'bar',
    data: { labels: costKeys, datasets: [{ label: '€ Cost', data: costKeys.map(k => +costMonth[k].toFixed(2)), backgroundColor: '#a855f7' }] },
    options: { plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true } } }
  });

  // Slowest / fastest tables
  const sorted = [...withTurnaround].sort((a,b) => b.turnaround - a.turnaround);
  const row10 = r => `<tr><td>${r.company}</td><td>${r.owner}</td><td>${turnaroundColor(r.turnaround)}</td><td>${fmtDate(r.shipDate)}</td></tr>`;
  document.getElementById('sh-slowest-list').innerHTML = sorted.slice(0,10).map(row10).join('');
  document.getElementById('sh-fastest-list').innerHTML = [...sorted].reverse().slice(0,10).map(row10).join('');

  // Main table
  document.getElementById('sh-body').innerHTML = rows.length === 0
    ? '<tr><td colspan="14">No matched shipments found.</td></tr>'
    : rows.map(r => `<tr>
        <td><strong>${r.company}</strong></td>
        <td class="print-name" title="${r.recipient}">${r.recipient || '—'}</td>
        <td>${r.owner || '—'}</td>
        <td>${typeBadge(r.type)}</td>
        <td>${fmtDate(r.dateAdded)}</td>
        <td>${fmtDate(r.shipDate)}</td>
        <td>${turnaroundColor(r.turnaround)}</td>
        <td>${r.destination || '—'}</td>
        <td>${r.carrier ? r.carrier.split(' ')[0] : '—'}${r.boxes > 1 ? ` <span style="font-size:11px;color:var(--text-3);">(${r.boxes} boxes)</span>` : ''}</td>
        <td>${r.price ? '€'+r.price.toFixed(2) : '—'}</td>
        <td>${r.quantity || '—'}</td>
        <td>${confidenceBadge(r.confidence)}</td>
        <td>${trackingUrl(r.tracking, r.carrier) ? `<a href="${trackingUrl(r.tracking, r.carrier)}" target="_blank" rel="noopener" style="color:var(--blue);text-decoration:none;font-size:16px;" title="${r.tracking}">📦</a>` : (r.tracking ? `<span title="${r.tracking}" style="cursor:default;opacity:0.5;">📦</span>` : '—')}</td>
      </tr>`).join('');
}

['sh-search','sh-owner','sh-period','sh-date-from','sh-date-to'].forEach(id => {
  document.getElementById(id).addEventListener('input', renderShipping);
  document.getElementById(id).addEventListener('change', renderShipping);
});
document.getElementById('sh-global-period').addEventListener('change', renderShipping);
document.getElementById('sh-period').addEventListener('change', function() {
  document.getElementById('sh-custom-range').style.display = this.value === 'custom' ? 'inline-flex' : 'none';
});

// ── Review Section ─────────────────────────────────────────────
const DISMISSED_KEY = 'izy_review_dismissed';

function getDismissed() {
  try { return new Set(JSON.parse(localStorage.getItem(DISMISSED_KEY) || '[]')); }
  catch { return new Set(); }
}

function saveDismissed(set) {
  localStorage.setItem(DISMISSED_KEY, JSON.stringify([...set]));
}

function dismissJob(priority) {
  const d = getDismissed();
  d.add(priority);
  saveDismissed(d);
  renderReview();
}

function restoreDismissed() {
  saveDismissed(new Set());
  renderReview();
}

function renderReview() {
  const countEl = document.getElementById('review-count');
  const tbody   = document.getElementById('review-body');
  if (!countEl || !tbody) return;

  const dismissed = getDismissed();
  const visible   = reviewRows.filter(r => !dismissed.has(r.priority));

  countEl.textContent = visible.length;
  document.getElementById('review-section').style.display = reviewRows.length ? 'block' : 'none';

  const restoreBtn = document.getElementById('review-restore');
  if (restoreBtn) restoreBtn.style.display = dismissed.size ? 'inline-block' : 'none';

  tbody.innerHTML = visible.length === 0
    ? '<tr><td colspan="8" style="color:#94a3b8;text-align:center;">All clear — no unmatched shipped jobs.</td></tr>'
    : visible.map(r => `<tr>
        <td><strong>${r.company}</strong></td>
        <td class="print-name">${r.printName || '—'}</td>
        <td>${r.owner || '—'}</td>
        <td>${fmtDate(r.dateAdded)}</td>
        <td>${fmtDate(r.shippedDate)}</td>
        <td>${r.quantity ? r.quantity.toLocaleString() : '—'}</td>
        <td><button onclick="dismissJob('${r.priority}')" style="background:none;border:1px solid #e2e8f0;border-radius:6px;padding:2px 8px;font-size:11px;color:#94a3b8;cursor:pointer;" title="Remove from this list">✕ Dismiss</button></td>
      </tr>`).join('');
}

// ── Sleeve Queue ──────────────────────────────────────────────

let svTypeFilter = '';
let svCompletedOpen = false;
const SV_COMPLETED_STATUSES = ['ordered', 'done', 'finished'];

function sleeveSortOrder(r) {
  const s = get(r,'Status').toLowerCase();
  if (s.includes('progress')) return 0;
  if (s === 'to sleeve')      return 1;
  if (s === 'done')           return 2;
  return 3;
}

function populateSleeveOwners() {
  const owners = [...new Set([...KNOWN_OWNERS, ...allRows.map(r => get(r,'Owner')).filter(Boolean)])].sort();
  const sel = document.getElementById('sv-owner');
  sel.innerHTML = '<option value="">Select owner…</option>' +
    owners.map(o => `<option value="${o}">${o}</option>`).join('');
}

function populateSleeveStatusFilter() {
  const statuses = [...new Set(sleeveRows.map(r => get(r,'Status')).filter(Boolean))].sort();
  fill('sv-status', statuses);
}

async function loadSleeves() {
  if (sleeveLoaded && sleeveRows.length) { renderSleeves(); return; }
  document.getElementById('sv-sections').innerHTML =
    '<p style="color:#94a3b8;padding:20px;">Loading sleeve data…</p>';
  try {
    const data = await fetch(SCRIPT_URL + '?sheet=sleeves&t=' + Date.now()).then(r => r.json());
    sleeveRows   = (data.rows || []).filter(r => get(r,'Name_Company') || get(r,'Status'));
    sleeveLoaded = true;
    populateSleeveStatusFilter();
    renderSleeves();
  } catch (err) {
    document.getElementById('sv-sections').innerHTML =
      '<p style="color:var(--red);padding:20px;">Error loading sleeve data — check your connection.</p>';
  }
}

function refreshSleeves() {
  sleeveLoaded = false;
  loadSleeves();
}

function renderSleeves() {
  const search = document.getElementById('sv-search').value.toLowerCase();
  const status = document.getElementById('sv-status').value;

  const isCompleted = r => SV_COMPLETED_STATUSES.includes(get(r,'Status').toLowerCase());

  const filtered = sleeveRows.filter(r => {
    if (isCompleted(r)) return false; // handled separately
    if (status && get(r,'Status') !== status) return false;
    if (search && !(
      get(r,'Name_Company').toLowerCase().includes(search) ||
      (get(r,'Name_Print') || '').toLowerCase().includes(search)
    )) return false;
    return true;
  });

  const completedFiltered = sleeveRows.filter(r => {
    if (!isCompleted(r)) return false;
    if (status && get(r,'Status') !== status) return false;
    if (search && !(
      get(r,'Name_Company').toLowerCase().includes(search) ||
      (get(r,'Name_Print') || '').toLowerCase().includes(search)
    )) return false;
    return true;
  });

  // Type filter pills (reuse AQ_SECTIONS)
  const tabsEl = document.getElementById('sv-type-tabs');
  tabsEl.innerHTML =
    `<button class="aq-type-tab${svTypeFilter === '' ? ' active' : ''}" data-svtype="">All <span class="aq-tab-count">${filtered.length}</span></button>` +
    AQ_SECTIONS.map(s => {
      const n = filtered.filter(r => s.match(get(r,'Soort').toLowerCase())).length;
      if (!n) return '';
      return `<button class="aq-type-tab${svTypeFilter === s.label ? ' active' : ''}" data-svtype="${s.label}" style="--tc:${s.colors.text};--tb:${s.colors.bg}">${s.label} <span class="aq-tab-count">${n}</span></button>`;
    }).join('');

  const html = AQ_SECTIONS.filter(s => !svTypeFilter || s.label === svTypeFilter).map(section => {
    const rows = filtered
      .filter(r => section.match(get(r,'Soort').toLowerCase()))
      .sort((a,b) => sleeveSortOrder(a) - sleeveSortOrder(b));

    if (rows.length === 0) return '';

    const c = section.colors;
    const rowsHtml = rows.map(r => {
      const idx    = sleeveRows.indexOf(r);
      const isDone = get(r,'Status').toLowerCase() === 'done';

      const actionBtns = `<button class="btn-log sv-btn-log" data-svidx="${idx}">✏️ Update</button><button class="btn-log sv-btn-edit" data-svidx="${idx}" style="background:var(--blue-dim);color:var(--blue);">✎ Edit</button>`;
      const svChk = `<label class="row-check-wrap" onclick="event.stopPropagation()"><input type="checkbox" class="row-select sv-select" data-svidx="${idx}" ${svSelected.has(idx) ? 'checked' : ''} /></label>`;

      // Support multiple file URLs (newline or comma separated)
      const rawFileUrls = (getCI(r,'file') || '').split(/[\n,]/).map(u => u.trim()).filter(Boolean);
      const fileLinks = rawFileUrls.length
        ? rawFileUrls.map((u,i) => `<a href="${u}" target="_blank" rel="noopener" style="color:var(--blue);font-size:12px;text-decoration:none;" title="Open file">📎 File${rawFileUrls.length > 1 ? ' '+(i+1) : ''}</a>`).join(' ')
        : '';

      const card = `<div class="aq-card${isDone ? ' sv-card-done' : ''}${svSelected.has(idx) ? ' row-selected' : ''}" style="--tc:${c.text};--tb:${c.bg}">
        <div class="aq-card-top">
          <div class="aq-card-left">
            ${svChk}
            <span class="aq-company">${get(r,'Name_Company')}</span>
          </div>
          ${badge(get(r,'Status'))}
        </div>
        ${get(r,'Name_Print') ? `<div class="aq-print-name">${get(r,'Name_Print')}</div>` : ''}
        ${fileLinks ? `<div style="margin:4px 0 2px;display:flex;flex-wrap:wrap;gap:6px;">${fileLinks}</div>` : ''}
        <div class="aq-meta">
          ${get(r,'Bottle color') ? `<div class="aq-meta-item"><span class="aq-meta-label">Color</span><span>${get(r,'Bottle color')}</span></div>` : ''}
          ${get(r,'Lid') ? `<div class="aq-meta-item"><span class="aq-meta-label">Lid</span><span>${get(r,'Lid')}</span></div>` : ''}
          ${get(r,'Owner') ? `<div class="aq-meta-item"><span class="aq-meta-label">Owner</span><span>${get(r,'Owner')}</span></div>` : ''}
          ${get(r,'Deadline') ? `<div class="aq-meta-item"><span class="aq-meta-label">Deadline</span><span>${get(r,'Deadline')}</span></div>` : ''}
          ${get(r,'Notes') ? `<div class="aq-meta-item aq-meta-notes"><span class="aq-meta-label">Notes</span><span class="notes-cell" title="${(get(r,'Notes')).replace(/"/g,"'")}">${get(r,'Notes')}</span></div>` : ''}
        </div>
        <div class="aq-card-actions">${actionBtns}</div>
      </div>`;

      const fileCell = rawFileUrls.length
        ? rawFileUrls.map((u,i) => `<a href="${u}" target="_blank" rel="noopener" style="color:var(--blue);text-decoration:none;">📎${rawFileUrls.length > 1 ? (i+1) : ''}</a>`).join(' ')
        : '—';

      const row = `<tr class="${isDone ? 'row-shipped' : ''}${svSelected.has(idx) ? ' row-selected' : ''}">
        <td><input type="checkbox" class="row-select sv-select" data-svidx="${idx}" ${svSelected.has(idx) ? 'checked' : ''} /></td>
        <td><strong>${get(r,'Name_Company')}</strong></td>
        <td>${badge(get(r,'Status'))}</td>
        <td>${typeBadge(get(r,'Soort'))}</td>
        <td>${get(r,'Bottle color') || '—'}</td>
        <td>${get(r,'Lid') || '—'}</td>
        <td>${get(r,'Owner') || '—'}</td>
        <td>${get(r,'Deadline') || '—'}</td>
        <td>${fileCell}</td>
        <td class="notes-cell" title="${(get(r,'Notes') || '').replace(/"/g,'&quot;')}">${get(r,'Notes') || '—'}</td>
        <td style="white-space:nowrap">${actionBtns}</td>
      </tr>`;

      return { card, row };
    });

    return `
      <div class="aq-section">
        <div class="aq-section-title" style="background:${c.bg};color:${c.text};">
          <span style="font-size:13px;font-weight:700;">${section.label}</span>
          <span class="aq-section-count" style="color:${c.text};opacity:0.7;">${rows.length} job${rows.length !== 1 ? 's' : ''}</span>
        </div>
        <div class="aq-cards">${rowsHtml.map(x => x.card).join('')}</div>
        <div class="aq-table-wrap table-wrap">
          <table>
            <thead><tr>
              <th></th><th>Company</th><th>Status</th>
              <th>Type</th><th>Color</th><th>Lid</th><th>Owner</th><th>Deadline</th><th>Files</th><th>Notes</th><th></th>
            </tr></thead>
            <tbody>${rowsHtml.map(x => x.row).join('')}</tbody>
          </table>
        </div>
      </div>`;
  }).join('');

  // Build completed (ordered/done/finished) collapsible section
  let completedHtml = '';
  if (completedFiltered.length > 0) {
    const rowsHtml = completedFiltered.map(r => {
      const idx = sleeveRows.indexOf(r);
      const actionBtns = `<button class="btn-log sv-btn-log" data-svidx="${idx}">✏️ Update</button><button class="btn-log sv-btn-edit" data-svidx="${idx}" style="background:var(--blue-dim);color:var(--blue);">✎ Edit</button>`;
      const svChk = `<label class="row-check-wrap" onclick="event.stopPropagation()"><input type="checkbox" class="row-select sv-select" data-svidx="${idx}" ${svSelected.has(idx) ? 'checked' : ''} /></label>`;
      const rawFileUrls = (getCI(r,'file') || '').split(/[\n,]/).map(u => u.trim()).filter(Boolean);
      const fileLinks = rawFileUrls.length
        ? rawFileUrls.map((u,i) => `<a href="${u}" target="_blank" rel="noopener" style="color:var(--blue);font-size:12px;text-decoration:none;">📎 File${rawFileUrls.length > 1 ? ' '+(i+1) : ''}</a>`).join(' ')
        : '';
      const fileCell = rawFileUrls.length
        ? rawFileUrls.map((u,i) => `<a href="${u}" target="_blank" rel="noopener" style="color:var(--blue);text-decoration:none;">📎${rawFileUrls.length > 1 ? (i+1) : ''}</a>`).join(' ')
        : '—';
      const card = `<div class="aq-card sv-card-done${svSelected.has(idx) ? ' row-selected' : ''}" style="--tc:#64748b;--tb:#f1f5f9">
        <div class="aq-card-top">
          <div class="aq-card-left">${svChk}<span class="aq-company">${get(r,'Name_Company')}</span></div>
          ${badge(get(r,'Status'))}
        </div>
        ${get(r,'Name_Print') ? `<div class="aq-print-name">${get(r,'Name_Print')}</div>` : ''}
        ${fileLinks ? `<div style="margin:4px 0 2px;display:flex;flex-wrap:wrap;gap:6px;">${fileLinks}</div>` : ''}
        <div class="aq-meta">
          ${get(r,'Bottle color') ? `<div class="aq-meta-item"><span class="aq-meta-label">Color</span><span>${get(r,'Bottle color')}</span></div>` : ''}
          ${get(r,'Lid') ? `<div class="aq-meta-item"><span class="aq-meta-label">Lid</span><span>${get(r,'Lid')}</span></div>` : ''}
          ${get(r,'Owner') ? `<div class="aq-meta-item"><span class="aq-meta-label">Owner</span><span>${get(r,'Owner')}</span></div>` : ''}
          ${get(r,'Deadline') ? `<div class="aq-meta-item"><span class="aq-meta-label">Deadline</span><span>${get(r,'Deadline')}</span></div>` : ''}
        </div>
        <div class="aq-card-actions">${actionBtns}</div>
      </div>`;
      const row = `<tr class="row-shipped${svSelected.has(idx) ? ' row-selected' : ''}">
        <td><input type="checkbox" class="row-select sv-select" data-svidx="${idx}" ${svSelected.has(idx) ? 'checked' : ''} /></td>
        <td><strong>${get(r,'Name_Company')}</strong></td>
        <td>${badge(get(r,'Status'))}</td>
        <td>${typeBadge(get(r,'Soort'))}</td>
        <td>${get(r,'Bottle color') || '—'}</td>
        <td>${get(r,'Lid') || '—'}</td>
        <td>${get(r,'Owner') || '—'}</td>
        <td>${get(r,'Deadline') || '—'}</td>
        <td>${fileCell}</td>
        <td class="notes-cell" title="${(get(r,'Notes') || '').replace(/"/g,'&quot;')}">${get(r,'Notes') || '—'}</td>
        <td style="white-space:nowrap">${actionBtns}</td>
      </tr>`;
      return { card, row };
    });

    completedHtml = `<div class="sv-completed-section" style="margin-top:16px;">
      <button class="sv-completed-toggle" onclick="svCompletedOpen=!svCompletedOpen;renderSleeves();" style="width:100%;display:flex;align-items:center;justify-content:space-between;padding:10px 16px;background:var(--hover);border:1px solid var(--border);border-radius:var(--radius);cursor:pointer;font-size:13px;font-weight:600;color:var(--text-2);">
        <span>📦 Ordered &amp; Finished <span style="font-weight:400;opacity:0.7;">(${completedFiltered.length} job${completedFiltered.length !== 1 ? 's' : ''})</span></span>
        <span>${svCompletedOpen ? '▲ Collapse' : '▼ Expand'}</span>
      </button>
      ${svCompletedOpen ? `
        <div class="aq-cards" style="margin-top:10px;">${rowsHtml.map(x => x.card).join('')}</div>
        <div class="aq-table-wrap table-wrap" style="margin-top:8px;">
          <table>
            <thead><tr>
              <th></th><th>Company</th><th>Status</th>
              <th>Type</th><th>Color</th><th>Lid</th><th>Owner</th><th>Deadline</th><th>Files</th><th>Notes</th><th></th>
            </tr></thead>
            <tbody>${rowsHtml.map(x => x.row).join('')}</tbody>
          </table>
        </div>` : ''}
    </div>`;
  }

  document.getElementById('sv-sections').innerHTML = (html ||
    '<p style="color:#94a3b8;padding:20px;">No sleeve jobs match the current filters.</p>') + completedHtml;
}

['sv-search','sv-status','sv-hide-done'].forEach(id => {
  const el = document.getElementById(id);
  el.addEventListener(id === 'sv-hide-done' ? 'change' : 'input', renderSleeves);
  el.addEventListener('change', renderSleeves);
});

document.getElementById('sv-type-tabs').addEventListener('click', e => {
  const btn = e.target.closest('.aq-type-tab');
  if (btn) { svTypeFilter = btn.dataset.svtype; renderSleeves(); }
});

// Sleeve tab button delegation
document.getElementById('tab-sleeves').addEventListener('click', function(e) {
  const selectCb = e.target.closest('.sv-select');
  if (selectCb && e.target.tagName === 'INPUT') {
    const idx = parseInt(selectCb.dataset.svidx);
    if (selectCb.checked) svSelected.add(idx); else svSelected.delete(idx);
    updateSelectionBar();
    return;
  }
  const logBtn  = e.target.closest('.sv-btn-log');
  if (logBtn) openSleeveModal(parseInt(logBtn.dataset.svidx));
  const editBtn = e.target.closest('.sv-btn-edit');
  if (editBtn) openEditJobModal(parseInt(editBtn.dataset.svidx), 'sleeve');
});

// ── Sleeve Modal ──────────────────────────────────────────────

let sleeveModalJob = null;

function addSleeveFileRow() {
  const container = document.getElementById('sv-modal-files');
  const row = document.createElement('div');
  row.className = 'sv-file-row';
  row.innerHTML = `<input type="file" accept="*/*" /><button type="button" class="btn-remove-file" title="Remove">✕</button>`;
  row.querySelector('.btn-remove-file').addEventListener('click', () => row.remove());
  container.appendChild(row);
}

function openSleeveModal(rowIdx) {
  sleeveModalJob = sleeveRows[rowIdx];
  if (!sleeveModalJob) return;

  document.getElementById('sv-modal-job-info').innerHTML = `
    <div class="modal-job-card">
      <div><span class="modal-label">Job</span><strong>#${get(sleeveModalJob,'Priority')} — ${get(sleeveModalJob,'Name_Company')}</strong></div>
      <div><span class="modal-label">Type</span>${typeBadge(get(sleeveModalJob,'Soort'))}</div>
      ${get(sleeveModalJob,'Bottle color') ? `<div><span class="modal-label">Color</span>${get(sleeveModalJob,'Bottle color')}</div>` : ''}
      ${get(sleeveModalJob,'Lid') ? `<div><span class="modal-label">Lid</span>${get(sleeveModalJob,'Lid')}</div>` : ''}
      ${get(sleeveModalJob,'Deadline') ? `<div><span class="modal-label">Deadline</span>${get(sleeveModalJob,'Deadline')}</div>` : ''}
    </div>`;

  // Pre-select current status
  const sel = document.getElementById('sv-modal-status-select');
  const currentStatus = get(sleeveModalJob,'Status');
  sel.value = currentStatus || '';

  // Reset file list to one empty row
  document.getElementById('sv-modal-files').innerHTML = '';
  addSleeveFileRow();
  document.getElementById('sv-modal-file-status').textContent = '';
  document.getElementById('sv-modal-status').textContent = '';

  document.getElementById('sleeve-modal-overlay').style.display = 'flex';
  sel.focus();
}

function closeSleeveModal() {
  document.getElementById('sleeve-modal-overlay').style.display = 'none';
  document.getElementById('sv-modal-files').innerHTML = '';
  document.getElementById('sv-modal-file-status').textContent = '';
  document.getElementById('sv-modal-status-select').value = '';
  document.getElementById('sv-modal-status').textContent = '';
  sleeveModalJob = null;
}

document.getElementById('sleeve-modal-overlay').addEventListener('click', function(e) {
  if (e.target === this) closeSleeveModal();
});

document.getElementById('sv-modal-add-file').addEventListener('click', addSleeveFileRow);

async function submitSleeveUpdate() {
  const statusEl     = document.getElementById('sv-modal-status');
  const submitBtn    = document.getElementById('sv-modal-submit');
  const fileStatusEl = document.getElementById('sv-modal-file-status');
  const chosenStatus = document.getElementById('sv-modal-status-select').value;

  if (!chosenStatus) {
    statusEl.textContent = 'Please select a status.';
    statusEl.className   = 'modal-error';
    return;
  }

  submitBtn.disabled    = true;
  submitBtn.textContent = 'Submitting…';
  statusEl.textContent  = '';

  // Collect and encode all selected files
  const fileRows = document.getElementById('sv-modal-files').querySelectorAll('.sv-file-row input[type="file"]');
  const sleeveFiles = [];
  const hasFiles = Array.from(fileRows).some(inp => inp.files && inp.files[0]);
  if (hasFiles) {
    fileStatusEl.textContent = 'Reading files…';
    for (const inp of fileRows) {
      if (inp.files && inp.files[0]) {
        try {
          const f = await readFileAsBase64(inp.files[0]);
          sleeveFiles.push({ base64: f.data, mime: f.mime, name: f.name });
        } catch (_) {}
      }
    }
    fileStatusEl.textContent = '';
  }

  try {
    await fetch(SCRIPT_URL, {
      method: 'POST',
      mode:   'no-cors',
      body:   JSON.stringify({
        action:      'update_sleeve',
        sheetRow:    sleeveModalJob['_sheetRow'],
        status:      chosenStatus,
        sleeveFiles,
        changedBy:   currentUser?.email,
      }),
    });
    // Apply to all other selected sleeve rows (status only, no files)
    const otherSvIdxs = [...svSelected].filter(i => sleeveRows[i] !== sleeveModalJob);
    for (const i of otherSvIdxs) {
      const r = sleeveRows[i];
      if (!r) continue;
      await fetch(SCRIPT_URL, { method: 'POST', mode: 'no-cors', body: JSON.stringify({
        action: 'update_sleeve', sheetRow: r['_sheetRow'], status: chosenStatus, changedBy: currentUser?.email,
      })});
    }
    svSelected.clear();
    updateSelectionBar();

    submitBtn.textContent = 'Submit';
    submitBtn.disabled    = false;
    statusEl.textContent  = '✅ Saved!';
    statusEl.className    = 'modal-success';
    setTimeout(() => { closeSleeveModal(); sleeveLoaded = false; loadSleeves(); }, 1800);
  } catch (err) {
    submitBtn.textContent = 'Submit';
    submitBtn.disabled    = false;
    statusEl.textContent  = '❌ Error: ' + err.message;
    statusEl.className    = 'modal-error';
  }
}

async function markSleeveDone(rowIdx) {
  const job = sleeveRows[rowIdx];
  if (!job) return;
  if (!confirm(`Mark "${get(job,'Name_Company')} #${get(job,'Priority')}" as Done?`)) return;
  try {
    await fetch(SCRIPT_URL, {
      method: 'POST',
      mode:   'no-cors',
      body:   JSON.stringify({
        action:    'update_sleeve',
        sheetRow:  job['_sheetRow'],
        status:    'Done',
        changedBy: currentUser?.email,
      }),
    });
    sleeveLoaded = false;
    loadSleeves();
  } catch (err) {
    alert('Could not update: ' + err.message);
  }
}

async function resetSleeveJob(rowIdx) {
  const job = sleeveRows[rowIdx];
  if (!job) return;
  if (!confirm(`Reset "${get(job,'Name_Company')} #${get(job,'Priority')}" back to "To make"?`)) return;
  try {
    await fetch(SCRIPT_URL, {
      method: 'POST',
      mode:   'no-cors',
      body:   JSON.stringify({
        action:    'update_sleeve',
        sheetRow:  job['_sheetRow'],
        status:    'To make',
        changedBy: currentUser?.email,
      }),
    });
    sleeveLoaded = false;
    loadSleeves();
  } catch (err) {
    alert('Could not reset: ' + err.message);
  }
}

// ── Add Sleeve Job ────────────────────────────────────────────

function populateApprovedMockupsSv() {
  const sel = document.getElementById('sv-mockup-select');
  if (!sel) return;
  const approved = mockupRows.filter(r => get(r,'Status').toLowerCase() === 'approved');
  sel.innerHTML = '<option value="">— Select an approved mockup to pre-fill —</option>' +
    approved.map(r => {
      const label = [get(r,'Name_Company'), get(r,'Soort'), get(r,'Bottle color')].filter(Boolean).join(' · ');
      return `<option value="${mockupRows.indexOf(r)}">${label}</option>`;
    }).join('');
}

document.getElementById('sv-mockup-select').addEventListener('change', function() {
  const idx = parseInt(this.value);
  const filesDiv = document.getElementById('sv-mockup-files');

  // Clear on empty selection
  if (isNaN(idx)) { filesDiv.style.display = 'none'; filesDiv.innerHTML = ''; return; }
  const r = mockupRows[idx];
  if (!r) return;

  const setVal = (id, val) => { const el = document.getElementById(id); if (el && val !== undefined && val !== null) el.value = val; };
  setVal('sv-soort',        get(r,'Soort'));
  setVal('sv-company',      get(r,'Name_Company'));
  setVal('sv-bottle-color', get(r,'Bottle color') || get(r,'Color'));
  setVal('sv-lid-color',    get(r,'Lid'));
  setVal('sv-owner',        normOwner(get(r,'Owner')));
  setVal('sv-deadline',     get(r,'Deadline'));
  setVal('sv-quantity',     get(r,'Quantity'));
  setVal('sv-notes',        get(r,'Notes'));

  // Show Drive file links from the mockup (can't pre-load into file inputs due to browser security)
  const rawUrls = (getCI(r,'file') || '').split(/[\n,]/).map(u => u.trim()).filter(Boolean);
  if (rawUrls.length) {
    filesDiv.innerHTML = '<div style="font-size:12px;color:#94a3b8;margin-bottom:4px;">Files from mockup:</div>' +
      rawUrls.map((u, i) => `<a href="${u}" target="_blank" rel="noopener" style="display:inline-flex;align-items:center;gap:4px;color:var(--blue);font-size:13px;text-decoration:none;margin-right:10px;margin-bottom:4px;">📎 File${rawUrls.length > 1 ? ' '+(i+1) : ''}</a>`).join('');
    filesDiv.style.display = 'block';
  } else {
    filesDiv.style.display = 'none';
    filesDiv.innerHTML = '';
  }

  document.getElementById('sv-company').dispatchEvent(new Event('input'));
});

function toggleAddSleeveForm() {
  const overlay = document.getElementById('add-sleeve-modal-overlay');
  overlay.style.display = 'flex';
  populateSleeveOwners();
  if (!mockupLoaded) loadMockups().then(populateApprovedMockupsSv); else populateApprovedMockupsSv();
}
function closeAddSleeveModal() {
  document.getElementById('add-sleeve-modal-overlay').style.display = 'none';
  document.getElementById('sv-files-list').innerHTML = '';
  document.getElementById('sv-mockup-files').style.display = 'none';
  document.getElementById('sv-mockup-files').innerHTML = '';
  document.getElementById('sv-mockup-select').value = '';
  addSvFileRow();
}

// Sleeve add-form: multi-file rows
function addSvFileRow() {
  const container = document.getElementById('sv-files-list');
  const row = document.createElement('div');
  row.className = 'sv-file-row';
  row.innerHTML = `<input type="file" accept="*/*" /><button type="button" class="btn-remove-file" title="Remove">✕</button>`;
  row.querySelector('.btn-remove-file').addEventListener('click', () => row.remove());
  container.appendChild(row);
}
document.getElementById('sv-add-file')?.addEventListener('click', addSvFileRow);
if (document.getElementById('sv-files-list')) addSvFileRow();

function readFileAsBase64(file, onProgress) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    if (onProgress) reader.onprogress = e => { if (e.lengthComputable) onProgress(e.loaded / e.total); };
    reader.onload  = e => resolve({ data: e.target.result, name: file.name, mime: file.type || 'application/octet-stream' });
    reader.onerror = () => reject(new Error('Could not read file'));
    reader.readAsDataURL(file);
  });
}

async function postWithProgress(url, body, onProgress) {
  // Apps Script redirects POST→GET with XHR, so we must use fetch no-cors.
  // Real upload progress isn't available; animate smoothly toward 95% instead.
  let p = 0;
  const timer = onProgress ? setInterval(() => {
    p = Math.min(0.93, p + (0.93 - p) * 0.12);
    onProgress(p);
  }, 250) : null;
  try {
    await fetch(url, { method: 'POST', mode: 'no-cors', body });
  } finally {
    if (timer) clearInterval(timer);
  }
}

document.getElementById('sv-submit').addEventListener('click', async function() {
  const soort     = document.getElementById('sv-soort').value;
  const company   = document.getElementById('sv-company').value.trim();
  const printName = '';
  const quantity  = document.getElementById('sv-quantity').value;
  const deadline     = document.getElementById('sv-deadline').value;
  const owner        = document.getElementById('sv-owner').value;
  const bottleColor  = document.getElementById('sv-bottle-color').value.trim();
  const lidColor     = document.getElementById('sv-lid-color').value.trim();
  const notes        = withDate(document.getElementById('sv-notes').value.trim());
  const statusEl  = document.getElementById('sv-form-status');

  if (!soort || !company || !quantity) {
    statusEl.className   = 'form-status error';
    statusEl.textContent = 'Please fill in all required fields.';
    return;
  }

  this.disabled        = true;
  statusEl.className   = 'form-status';
  statusEl.textContent = '';

  const svProgressWrap  = document.getElementById('sv-upload-progress');
  const svProgressFill  = document.getElementById('sv-upload-fill');
  const svProgressLabel = document.getElementById('sv-upload-label');
  const svSetProgress = (pct, label) => {
    if (svProgressWrap)  svProgressWrap.style.display = 'block';
    if (svProgressFill)  svProgressFill.style.width   = Math.round(pct * 100) + '%';
    if (svProgressLabel) svProgressLabel.textContent  = label;
  };
  const svHideProgress = () => {
    if (svProgressWrap) svProgressWrap.style.display = 'none';
    if (svProgressFill) svProgressFill.style.width   = '0%';
  };

  const svFileInputs = document.getElementById('sv-files-list').querySelectorAll('input[type="file"]');
  const designFiles  = [];
  const svFilesToRead = [...svFileInputs].filter(inp => inp.files && inp.files[0]);
  for (let i = 0; i < svFilesToRead.length; i++) {
    const baseProgress = i / svFilesToRead.length;
    svSetProgress(baseProgress * 0.4, `Reading file ${i + 1} of ${svFilesToRead.length}…`);
    try {
      const f = await readFileAsBase64(svFilesToRead[i].files[0], p => svSetProgress((baseProgress + p / svFilesToRead.length) * 0.4, `Reading file ${i + 1} of ${svFilesToRead.length}… ${Math.round(p * 100)}%`));
      designFiles.push({ base64: f.data, mime: f.mime, name: f.name });
    } catch (_) {}
  }

  svSetProgress(svFilesToRead.length ? 0.4 : 0, 'Uploading…');

  try {
    await postWithProgress(
      SCRIPT_URL,
      JSON.stringify({
        action:    'add_sleeve_job',
        soort, company, printName,
        quantity:  parseInt(quantity),
        deadline, owner, bottleColor, lidColor, notes,
        designFiles,
        changedBy: currentUser?.email,
      }),
      p => svSetProgress((svFilesToRead.length ? 0.4 : 0) + p * (svFilesToRead.length ? 0.6 : 1), `Uploading… ${Math.round(p * 100)}%`)
    );
    svSetProgress(1, 'Done!');
    statusEl.className   = 'form-status success';
    statusEl.textContent = '✓ Sleeve job added!';
    document.getElementById('add-sleeve-form').reset();
    setTimeout(() => { svHideProgress(); closeAddSleeveModal(); statusEl.textContent = ''; }, 1500);
    sleeveLoaded = false;
    loadSleeves();
  } catch (err) {
    svHideProgress();
    statusEl.className   = 'form-status error';
    statusEl.textContent = 'Error: ' + err.message;
  }
  this.disabled = false;
});

// ── Edit Job Modal (shared for Sleeve & Mockup) ───────────────

let editJobType = null; // 'sleeve' or 'mockup'
let editJobRow  = null; // the row object

function openEditJobModal(rowIdx, type) {
  const rows = type === 'sleeve' ? sleeveRows : type === 'active' ? allRows : mockupRows;
  editJobRow  = rows[rowIdx];
  editJobType = type;
  if (!editJobRow) return;

  document.getElementById('edit-job-modal-title').textContent =
    type === 'sleeve' ? 'Edit Sleeve Job' : type === 'active' ? 'Edit Active Job' : 'Edit Mockup Job';

  document.getElementById('ej-company').value      = get(editJobRow, 'Name_Company') || '';
  document.getElementById('ej-print-name').value  = get(editJobRow, 'Name_Print')   || '';
  document.getElementById('ej-soort').value        = get(editJobRow, 'Soort')        || '';
  document.getElementById('ej-bottle-color').value = get(editJobRow, 'Bottle color') || '';
  document.getElementById('ej-lid-color').value    = get(editJobRow, 'Lid')          || '';
  document.getElementById('ej-quantity').value     = get(editJobRow, 'Quantity')     || '';
  document.getElementById('ej-notes-display').textContent = get(editJobRow, 'Notes') || '—';
  document.getElementById('ej-new-note').value = '';

  // Show active-queue-only fields
  const activeFields = document.getElementById('ej-active-fields');
  activeFields.style.display = type === 'active' ? '' : 'none';
  if (type === 'active') {
    const sleeveVal = (get(editJobRow,'To sleeve?') || getCI(editJobRow,'sleeve') || 'No');
    const sleeveToggle = document.getElementById('ej-tosleeve');
    sleeveToggle.dataset.value = sleeveVal;
    sleeveToggle.querySelectorAll('.sleeve-opt').forEach(b => b.classList.toggle('active', b.dataset.opt === sleeveVal));
    document.getElementById('ej-ship-contact').value = get(editJobRow,'Contactpersoon')  || '';
    document.getElementById('ej-ship-phone').value   = get(editJobRow,'Telefoonnummer')  || '';
    document.getElementById('ej-ship-email').value   = get(editJobRow,'E-mailadres')     || '';
    document.getElementById('ej-ship-street').value  = get(editJobRow,'Straat')          || '';
    document.getElementById('ej-ship-number').value  = get(editJobRow,'Huisnummer')      || '';
    document.getElementById('ej-ship-zipcode').value = get(editJobRow,'Postcode')        || '';
    document.getElementById('ej-ship-city').value    = get(editJobRow,'Plaats')          || '';
    document.getElementById('ej-ship-country').value = get(editJobRow,'Land')            || '';
  }

  // Convert DD/MM/YYYY deadline to YYYY-MM-DD for date input
  const dl = get(editJobRow, 'Deadline') || '';
  if (dl && dl.includes('/')) {
    const [d, m, y] = dl.split('/');
    document.getElementById('ej-deadline').value = `${y}-${m.padStart(2,'0')}-${d.padStart(2,'0')}`;
  } else {
    document.getElementById('ej-deadline').value = dl;
  }

  // Populate owner dropdown
  const owners = [...new Set([...KNOWN_OWNERS, ...allRows.map(r => get(r,'Owner')).filter(Boolean)])].sort();
  const ownerSel = document.getElementById('ej-owner');
  ownerSel.innerHTML = '<option value="">Select owner…</option>' +
    owners.map(o => `<option value="${o}">${o}</option>`).join('');
  ownerSel.value = get(editJobRow, 'Owner') || '';

  document.getElementById('ej-status').textContent = '';
  document.getElementById('edit-job-modal-overlay').style.display = 'flex';
}

function addEjFileRow() {
  const container = document.getElementById('ej-files-list');
  const row = document.createElement('div');
  row.className = 'sv-file-row';
  row.innerHTML = `<input type="file" accept="*/*" /><button type="button" class="btn-remove-file" title="Remove">✕</button>`;
  row.querySelector('.btn-remove-file').addEventListener('click', () => row.remove());
  container.appendChild(row);
}
document.getElementById('ej-add-file')?.addEventListener('click', addEjFileRow);

function closeEditJobModal() {
  document.getElementById('edit-job-modal-overlay').style.display = 'none';
  document.getElementById('ej-files-list').innerHTML = '';
  editJobRow = null; editJobType = null;
}

document.getElementById('edit-job-modal-overlay').addEventListener('click', function(e) {
  if (e.target === this) closeEditJobModal();
});
document.getElementById('add-sleeve-modal-overlay').addEventListener('click', function(e) {
  if (e.target === this) closeAddSleeveModal();
});
document.getElementById('add-mockup-modal-overlay').addEventListener('click', function(e) {
  if (e.target === this) closeAddMockupModal();
});

async function submitEditJob() {
  const company = document.getElementById('ej-company').value.trim();
  if (!company) {
    document.getElementById('ej-status').textContent = 'Company is required.';
    return;
  }

  const btn = document.getElementById('ej-submit');
  btn.disabled = true;
  document.getElementById('ej-status').textContent = 'Saving…';

  // Collect any new files
  const ejFileInputs = document.getElementById('ej-files-list').querySelectorAll('input[type="file"]');
  const designFiles = [];
  for (const inp of ejFileInputs) {
    if (inp.files && inp.files[0]) {
      try {
        const f = await readFileAsBase64(inp.files[0]);
        designFiles.push({ base64: f.data, mime: f.mime, name: f.name });
      } catch (_) {}
    }
  }

  // Convert YYYY-MM-DD back to DD/MM/YYYY for the sheet
  const rawDeadline = document.getElementById('ej-deadline').value;
  let deadline = rawDeadline;
  if (rawDeadline && rawDeadline.includes('-')) {
    const [y, m, d] = rawDeadline.split('-');
    deadline = `${d}/${m}/${y}`;
  }

  try {
    await fetch(SCRIPT_URL, {
      method: 'POST',
      mode:   'no-cors',
      body:   JSON.stringify({
        action:      editJobType === 'sleeve' ? 'edit_sleeve_job' : editJobType === 'active' ? 'edit_active_job' : 'edit_mockup_job',
        sheetRow:    editJobRow['_sheetRow'],
        company,
        printName:   document.getElementById('ej-print-name').value.trim(),
        soort:       document.getElementById('ej-soort').value,
        bottleColor: document.getElementById('ej-bottle-color').value,
        lidColor:    document.getElementById('ej-lid-color').value,
        quantity:    document.getElementById('ej-quantity').value || '',
        deadline,
        owner:       document.getElementById('ej-owner').value,
        notes:       (() => {
          const existing = document.getElementById('ej-notes-display').textContent.trim();
          const existingVal = existing === '—' ? '' : existing;
          const newNote = document.getElementById('ej-new-note').value.trim();
          if (!newNote) return existingVal;
          return existingVal ? `${existingVal}\n${withDate(newNote)}` : withDate(newNote);
        })(),
        ...(editJobType === 'active' ? {
          tosleeve:    document.getElementById('ej-tosleeve').dataset.value || 'No',
          shipContact: document.getElementById('ej-ship-contact').value.trim(),
          shipPhone:   document.getElementById('ej-ship-phone').value.trim(),
          shipEmail:   document.getElementById('ej-ship-email').value.trim(),
          shipStreet:  document.getElementById('ej-ship-street').value.trim(),
          shipNumber:  document.getElementById('ej-ship-number').value.trim(),
          shipZipcode: document.getElementById('ej-ship-zipcode').value.trim(),
          shipCity:    document.getElementById('ej-ship-city').value.trim(),
          shipCountry: document.getElementById('ej-ship-country').value.trim().toUpperCase(),
        } : {}),
        designFiles,
        changedBy:   currentUser?.email,
      }),
    });
    document.getElementById('ej-status').textContent = '✅ Saved!';
    setTimeout(() => {
      closeEditJobModal();
      if (editJobType === 'sleeve') { sleeveLoaded = false; loadSleeves(); }
      else if (editJobType === 'mockup') { mockupLoaded = false; loadMockups(); }
      else { refreshData(); }
    }, 1200);
  } catch (err) {
    document.getElementById('ej-status').textContent = '❌ Error: ' + err.message;
    btn.disabled = false;
  }
}

// ── Add Mockup Job ────────────────────────────────────────────

function populateMockupOwners() {
  const owners = [...new Set([...KNOWN_OWNERS, ...allRows.map(r => get(r,'Owner')).filter(Boolean)])].sort();
  const sel = document.getElementById('mk-form-owner');
  sel.innerHTML = '<option value="">Select owner…</option>' +
    owners.map(o => `<option value="${o}">${o}</option>`).join('');
}

function toggleAddMockupForm() {
  const overlay = document.getElementById('add-mockup-modal-overlay');
  overlay.style.display = 'flex';
  populateMockupOwners();
}
function closeAddMockupModal() {
  document.getElementById('add-mockup-modal-overlay').style.display = 'none';
  document.getElementById('mk-files-list').innerHTML = '';
  addMkFileRow();
}

// Mockup add-form: multi-file rows
function addMkFileRow() {
  const container = document.getElementById('mk-files-list');
  const row = document.createElement('div');
  row.className = 'sv-file-row';
  row.innerHTML = `<input type="file" accept="*/*" /><button type="button" class="btn-remove-file" title="Remove">✕</button>`;
  row.querySelector('.btn-remove-file').addEventListener('click', () => row.remove());
  container.appendChild(row);
}
document.getElementById('mk-add-file')?.addEventListener('click', addMkFileRow);
if (document.getElementById('mk-files-list')) addMkFileRow();

document.getElementById('mk-submit').addEventListener('click', async function() {
  const soort     = document.getElementById('mk-soort').value;
  const company   = document.getElementById('mk-company').value.trim();
  const quantity  = document.getElementById('mk-quantity').value;
  const deadline     = document.getElementById('mk-deadline').value;
  const owner        = document.getElementById('mk-form-owner').value;
  const bottleColor  = document.getElementById('mk-bottle-color').value.trim();
  const lidColor     = document.getElementById('mk-lid-color').value.trim();
  const notes        = withDate(document.getElementById('mk-notes').value.trim());
  const statusEl  = document.getElementById('mk-form-status');

  if (!soort || !company) {
    statusEl.className   = 'form-status error';
    statusEl.textContent = 'Please fill in all required fields.';
    return;
  }

  this.disabled        = true;
  statusEl.className   = 'form-status';
  statusEl.textContent = '';

  const mkProgressWrap  = document.getElementById('mk-upload-progress');
  const mkProgressFill  = document.getElementById('mk-upload-fill');
  const mkProgressLabel = document.getElementById('mk-upload-label');
  const mkSetProgress = (pct, label) => {
    if (mkProgressWrap)  mkProgressWrap.style.display = 'block';
    if (mkProgressFill)  mkProgressFill.style.width   = Math.round(pct * 100) + '%';
    if (mkProgressLabel) mkProgressLabel.textContent  = label;
  };
  const mkHideProgress = () => {
    if (mkProgressWrap) mkProgressWrap.style.display = 'none';
    if (mkProgressFill) mkProgressFill.style.width   = '0%';
  };

  const mkFileInputs = document.getElementById('mk-files-list').querySelectorAll('input[type="file"]');
  const designFiles  = [];
  const mkFilesToRead = [...mkFileInputs].filter(inp => inp.files && inp.files[0]);
  for (let i = 0; i < mkFilesToRead.length; i++) {
    const baseProgress = i / mkFilesToRead.length;
    mkSetProgress(baseProgress * 0.4, `Reading file ${i + 1} of ${mkFilesToRead.length}…`);
    try {
      const f = await readFileAsBase64(mkFilesToRead[i].files[0], p => mkSetProgress((baseProgress + p / mkFilesToRead.length) * 0.4, `Reading file ${i + 1} of ${mkFilesToRead.length}… ${Math.round(p * 100)}%`));
      designFiles.push({ base64: f.data, mime: f.mime, name: f.name });
    } catch (_) {}
  }

  mkSetProgress(mkFilesToRead.length ? 0.4 : 0, 'Uploading…');

  try {
    await postWithProgress(
      SCRIPT_URL,
      JSON.stringify({
        action:    'add_mockup_job',
        soort, company, printName: '',
        quantity:  quantity ? parseInt(quantity) : '',
        deadline, owner, bottleColor, lidColor, notes,
        designFiles,
        changedBy: currentUser?.email,
      }),
      p => mkSetProgress((mkFilesToRead.length ? 0.4 : 0) + p * (mkFilesToRead.length ? 0.6 : 1), `Uploading… ${Math.round(p * 100)}%`)
    );
    mkSetProgress(1, 'Done!');
    statusEl.className   = 'form-status success';
    statusEl.textContent = '✓ Mockup job added!';
    document.getElementById('add-mockup-form').reset();
    setTimeout(() => { mkHideProgress(); closeAddMockupModal(); statusEl.textContent = ''; }, 1500);
    mockupLoaded = false;
    loadMockups();
  } catch (err) {
    mkHideProgress();
    statusEl.className   = 'form-status error';
    statusEl.textContent = 'Error: ' + err.message;
  }
  this.disabled = false;
});

// ── Mockups Tab ───────────────────────────────────────────────

function populateMockupStatusFilter() {
  const statuses = [...new Set(mockupRows.map(r => get(r,'Status')).filter(Boolean))].sort();
  fill('mk-status', statuses);
}

async function loadMockups() {
  if (mockupLoaded && mockupRows.length) { renderMockups(); return; }
  document.getElementById('mk-sections').innerHTML =
    '<p style="color:#94a3b8;padding:20px;">Loading mockup data…</p>';
  try {
    const data = await fetch(SCRIPT_URL + '?sheet=mockups&t=' + Date.now()).then(r => r.json());
    mockupRows   = (data.rows || []).filter(r => get(r,'Name_Company') || get(r,'Status'));
    mockupLoaded = true;
    populateMockupStatusFilter();
    renderMockups();
  } catch (err) {
    document.getElementById('mk-sections').innerHTML =
      '<p style="color:var(--red);padding:20px;">Error loading mockup data — check your connection.</p>';
  }
}

function refreshMockups() {
  mockupLoaded = false;
  loadMockups();
}

function renderMockups() {
  const search   = document.getElementById('mk-search').value.toLowerCase();
  const status   = document.getElementById('mk-status').value;
  const hideDone = document.getElementById('mk-hide-done').checked;

  const oneDayMs = 24 * 60 * 60 * 1000;
  const filtered = mockupRows.filter(r => {
    const st = get(r,'Status').toLowerCase();
    // Auto-hide approved mockups after 1 day
    if (st === 'approved') {
      const approvedStr = getCI(r, 'approved');
      if (approvedStr) {
        const approvedDate = new Date(approvedStr.split(' ')[0].split('/').reverse().join('-'));
        if (!isNaN(approvedDate) && (Date.now() - approvedDate.getTime()) >= oneDayMs) return false;
      }
    }
    if (hideDone && ['approved','finished'].includes(st)) return false;
    if (status && get(r,'Status') !== status) return false;
    if (search && !(
      get(r,'Name_Company').toLowerCase().includes(search) ||
      (get(r,'Name_Print') || '').toLowerCase().includes(search)
    )) return false;
    return true;
  });

  const tabsEl = document.getElementById('mk-type-tabs');
  tabsEl.innerHTML =
    `<button class="aq-type-tab${mkTypeFilter === '' ? ' active' : ''}" data-mktype="">All <span class="aq-tab-count">${filtered.length}</span></button>` +
    AQ_SECTIONS.map(s => {
      const n = filtered.filter(r => s.match(get(r,'Soort').toLowerCase())).length;
      if (!n) return '';
      return `<button class="aq-type-tab${mkTypeFilter === s.label ? ' active' : ''}" data-mktype="${s.label}" style="--tc:${s.colors.text};--tb:${s.colors.bg}">${s.label} <span class="aq-tab-count">${n}</span></button>`;
    }).join('');

  const html = AQ_SECTIONS.filter(s => !mkTypeFilter || s.label === mkTypeFilter).map(section => {
    const rows = filtered
      .filter(r => section.match(get(r,'Soort').toLowerCase()))
      .sort((a,b) => (parseInt(get(a,'Priority'))||0) - (parseInt(get(b,'Priority'))||0));

    if (rows.length === 0) return '';

    const c = section.colors;
    const rowsHtml = rows.map(r => {
      const idx    = mockupRows.indexOf(r);
      const isDone = ['approved','finished'].includes(get(r,'Status').toLowerCase());
      const actionBtns = `<button class="btn-log mk-btn-log" data-mkidx="${idx}">✏️ Update</button><button class="btn-log mk-btn-edit" data-mkidx="${idx}" style="background:var(--blue-dim);color:var(--blue);">✎ Edit</button>`;
      const mkChk = `<label class="row-check-wrap" onclick="event.stopPropagation()"><input type="checkbox" class="row-select mk-select" data-mkidx="${idx}" ${mkSelected.has(idx) ? 'checked' : ''} /></label>`;

      const rawFileUrls = (getCI(r,'file') || '').split(/[\n,]/).map(u => u.trim()).filter(Boolean);
      const fileLinks = rawFileUrls.length
        ? rawFileUrls.map((u,i) => `<a href="${u}" target="_blank" rel="noopener" style="color:var(--blue);font-size:12px;text-decoration:none;" title="Open file">📎 File${rawFileUrls.length > 1 ? ' '+(i+1) : ''}</a>`).join(' ')
        : '';

      const card = `<div class="aq-card${isDone ? ' sv-card-done' : ''}${mkSelected.has(idx) ? ' row-selected' : ''}" style="--tc:${c.text};--tb:${c.bg}">
        <div class="aq-card-top">
          <div class="aq-card-left">
            ${mkChk}
            <span class="aq-company">${get(r,'Name_Company')}</span>
          </div>
          ${badge(get(r,'Status'))}
        </div>
        ${get(r,'Name_Print') ? `<div class="aq-print-name">${get(r,'Name_Print')}</div>` : ''}
        ${fileLinks ? `<div style="margin:4px 0 2px;display:flex;flex-wrap:wrap;gap:6px;">${fileLinks}</div>` : ''}
        <div class="aq-meta">
          ${get(r,'Bottle color') ? `<div class="aq-meta-item"><span class="aq-meta-label">Color</span><span>${get(r,'Bottle color')}</span></div>` : ''}
          ${get(r,'Lid') ? `<div class="aq-meta-item"><span class="aq-meta-label">Lid</span><span>${get(r,'Lid')}</span></div>` : ''}
          ${get(r,'Owner') ? `<div class="aq-meta-item"><span class="aq-meta-label">Owner</span><span>${get(r,'Owner')}</span></div>` : ''}
          ${get(r,'Deadline') ? `<div class="aq-meta-item"><span class="aq-meta-label">Deadline</span><span>${get(r,'Deadline')}</span></div>` : ''}
          ${get(r,'Notes') ? `<div class="aq-meta-item aq-meta-notes"><span class="aq-meta-label">Notes</span><span class="notes-cell" title="${(get(r,'Notes')).replace(/"/g,"'")}">${get(r,'Notes')}</span></div>` : ''}
        </div>
        <div class="aq-card-actions">${actionBtns}</div>
      </div>`;

      const fileCell = rawFileUrls.length
        ? rawFileUrls.map((u,i) => `<a href="${u}" target="_blank" rel="noopener" style="color:var(--blue);text-decoration:none;">📎${rawFileUrls.length > 1 ? (i+1) : ''}</a>`).join(' ')
        : '—';

      const row = `<tr class="${isDone ? 'row-shipped' : ''}${mkSelected.has(idx) ? ' row-selected' : ''}">
        <td><input type="checkbox" class="row-select mk-select" data-mkidx="${idx}" ${mkSelected.has(idx) ? 'checked' : ''} /></td>
        <td><strong>${get(r,'Name_Company')}</strong></td>
        <td>${badge(get(r,'Status'))}</td>
        <td>${typeBadge(get(r,'Soort'))}</td>
        <td>${get(r,'Bottle color') || '—'}</td>
        <td>${get(r,'Lid') || '—'}</td>
        <td>${get(r,'Owner') || '—'}</td>
        <td>${get(r,'Deadline') || '—'}</td>
        <td>${fileCell}</td>
        <td class="notes-cell" title="${(get(r,'Notes') || '').replace(/"/g,'&quot;')}">${get(r,'Notes') || '—'}</td>
        <td style="white-space:nowrap">${actionBtns}</td>
      </tr>`;

      return { card, row };
    });

    return `
      <div class="aq-section">
        <div class="aq-section-title" style="background:${c.bg};color:${c.text};">
          <span style="font-size:13px;font-weight:700;">${section.label}</span>
          <span class="aq-section-count" style="color:${c.text};opacity:0.7;">${rows.length} job${rows.length !== 1 ? 's' : ''}</span>
        </div>
        <div class="aq-cards">${rowsHtml.map(x => x.card).join('')}</div>
        <div class="aq-table-wrap table-wrap">
          <table>
            <thead><tr>
              <th></th><th>Company</th><th>Status</th>
              <th>Type</th><th>Color</th><th>Lid</th><th>Owner</th><th>Deadline</th><th>Files</th><th>Notes</th><th></th>
            </tr></thead>
            <tbody>${rowsHtml.map(x => x.row).join('')}</tbody>
          </table>
        </div>
      </div>`;
  }).join('');

  document.getElementById('mk-sections').innerHTML = html ||
    '<p style="color:#94a3b8;padding:20px;">No mockup jobs match the current filters.</p>';
}

['mk-search','mk-status','mk-hide-done'].forEach(id => {
  const el = document.getElementById(id);
  el.addEventListener(id === 'mk-hide-done' ? 'change' : 'input', renderMockups);
  el.addEventListener('change', renderMockups);
});

document.getElementById('mk-type-tabs').addEventListener('click', e => {
  const btn = e.target.closest('.aq-type-tab');
  if (btn) { mkTypeFilter = btn.dataset.mktype; renderMockups(); }
});

document.getElementById('tab-mockups').addEventListener('click', function(e) {
  const selectCb = e.target.closest('.mk-select');
  if (selectCb && e.target.tagName === 'INPUT') {
    const idx = parseInt(selectCb.dataset.mkidx);
    if (selectCb.checked) mkSelected.add(idx); else mkSelected.delete(idx);
    updateSelectionBar();
    return;
  }
  const logBtn  = e.target.closest('.mk-btn-log');
  if (logBtn) openMockupModal(parseInt(logBtn.dataset.mkidx));
  const editBtn = e.target.closest('.mk-btn-edit');
  if (editBtn) openEditJobModal(parseInt(editBtn.dataset.mkidx), 'mockup');
});

// ── Mockup Modal ──────────────────────────────────────────────

let mockupModalJob = null;

function addMockupFileRow() {
  const container = document.getElementById('mk-modal-files');
  const row = document.createElement('div');
  row.className = 'sv-file-row';
  row.innerHTML = `<input type="file" accept="*/*" /><button type="button" class="btn-remove-file" title="Remove">✕</button>`;
  row.querySelector('.btn-remove-file').addEventListener('click', () => row.remove());
  container.appendChild(row);
}

function openMockupModal(rowIdx) {
  mockupModalJob = mockupRows[rowIdx];
  if (!mockupModalJob) return;

  document.getElementById('mk-modal-job-info').innerHTML = `
    <div class="modal-job-card">
      <div><span class="modal-label">Job</span><strong>#${get(mockupModalJob,'Priority')} — ${get(mockupModalJob,'Name_Company')}</strong></div>
      <div><span class="modal-label">Type</span>${typeBadge(get(mockupModalJob,'Soort'))}</div>
      ${get(mockupModalJob,'Bottle color') ? `<div><span class="modal-label">Color</span>${get(mockupModalJob,'Bottle color')}</div>` : ''}
      ${get(mockupModalJob,'Lid') ? `<div><span class="modal-label">Lid</span>${get(mockupModalJob,'Lid')}</div>` : ''}
      ${get(mockupModalJob,'Deadline') ? `<div><span class="modal-label">Deadline</span>${get(mockupModalJob,'Deadline')}</div>` : ''}
    </div>`;

  const sel = document.getElementById('mk-modal-status-select');
  sel.value = get(mockupModalJob,'Status') || '';

  document.getElementById('mk-modal-files').innerHTML = '';
  addMockupFileRow();
  document.getElementById('mk-modal-file-status').textContent = '';
  document.getElementById('mk-modal-status').textContent = '';

  document.getElementById('mockup-modal-overlay').style.display = 'flex';
  sel.focus();
}

function closeMockupModal() {
  document.getElementById('mockup-modal-overlay').style.display = 'none';
  document.getElementById('mk-modal-files').innerHTML = '';
  document.getElementById('mk-modal-file-status').textContent = '';
  document.getElementById('mk-modal-status-select').value = '';
  document.getElementById('mk-modal-status').textContent = '';
  mockupModalJob = null;
}

document.getElementById('mockup-modal-overlay').addEventListener('click', function(e) {
  if (e.target === this) closeMockupModal();
});

document.getElementById('mk-modal-add-file').addEventListener('click', addMockupFileRow);

async function submitMockupUpdate() {
  const statusEl     = document.getElementById('mk-modal-status');
  const submitBtn    = document.getElementById('mk-modal-submit');
  const fileStatusEl = document.getElementById('mk-modal-file-status');
  const chosenStatus = document.getElementById('mk-modal-status-select').value;

  if (!chosenStatus) {
    statusEl.textContent = 'Please select a status.';
    statusEl.className   = 'modal-error';
    return;
  }

  submitBtn.disabled    = true;
  submitBtn.textContent = 'Submitting…';
  statusEl.textContent  = '';

  const fileRows    = document.getElementById('mk-modal-files').querySelectorAll('.sv-file-row input[type="file"]');
  const mockupFiles = [];
  const hasFiles    = Array.from(fileRows).some(inp => inp.files && inp.files[0]);
  if (hasFiles) {
    fileStatusEl.textContent = 'Reading files…';
    for (const inp of fileRows) {
      if (inp.files && inp.files[0]) {
        try {
          const f = await readFileAsBase64(inp.files[0]);
          mockupFiles.push({ base64: f.data, mime: f.mime, name: f.name });
        } catch (_) {}
      }
    }
    fileStatusEl.textContent = '';
  }

  try {
    await fetch(SCRIPT_URL, {
      method: 'POST',
      mode:   'no-cors',
      body:   JSON.stringify({
        action:      'update_mockup',
        sheetRow:    mockupModalJob['_sheetRow'],
        status:      chosenStatus,
        mockupFiles,
        changedBy:   currentUser?.email,
      }),
    });
    // Apply to all other selected mockup rows (status only, no files)
    const otherMkIdxs = [...mkSelected].filter(i => mockupRows[i] !== mockupModalJob);
    for (const i of otherMkIdxs) {
      const r = mockupRows[i];
      if (!r) continue;
      await fetch(SCRIPT_URL, { method: 'POST', mode: 'no-cors', body: JSON.stringify({
        action: 'update_mockup', sheetRow: r['_sheetRow'], status: chosenStatus, changedBy: currentUser?.email,
      })});
    }
    mkSelected.clear();
    updateSelectionBar();

    statusEl.textContent = '✓ Updated!';
    statusEl.className   = 'modal-success';
    mockupLoaded = false;
    setTimeout(() => { closeMockupModal(); loadMockups(); }, 800);
  } catch (err) {
    statusEl.textContent = 'Error: ' + err.message;
    statusEl.className   = 'modal-error';
  }
  submitBtn.disabled    = false;
  submitBtn.textContent = 'Submit';
}

// ── Cache helpers ─────────────────────────────────────────────
const MAIN_CACHE_KEY = 'izy_main_cache_v2';

function saveMainCache(rows) {
  try { localStorage.setItem(MAIN_CACHE_KEY, JSON.stringify({ rows, ts: Date.now() })); } catch (_) {}
}
function loadMainCache() {
  try {
    const raw = localStorage.getItem(MAIN_CACHE_KEY);
    if (!raw) return null;
    const { rows } = JSON.parse(raw);
    return rows || null;
  } catch (_) { return null; }
}

// ── Render main tabs (everything except shipping) ──────────────
function getOverviewRows() {
  const period = document.getElementById('overview-period')?.value || 'all';
  if (period === 'all') return allRows;

  const now   = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());

  if (period === 'today') {
    const end = new Date(today.getTime() + 86399999);
    return allRows.filter(r => { const d = parseDate(get(r,'Date added')); return d && d >= today && d <= end; });
  }
  if (period === 'yesterday') {
    const yStart = new Date(today); yStart.setDate(today.getDate() - 1);
    const yEnd   = new Date(yStart.getTime() + 86399999);
    return allRows.filter(r => { const d = parseDate(get(r,'Date added')); return d && d >= yStart && d <= yEnd; });
  }
  if (period === 'custom') {
    const from = document.getElementById('overview-date-from')?.value;
    const to   = document.getElementById('overview-date-to')?.value;
    const f = from ? new Date(from) : null;
    const t = to   ? new Date(to + 'T23:59:59') : null;
    return allRows.filter(r => {
      const d = parseDate(get(r,'Date added'));
      if (!d) return false;
      if (f && d < f) return false;
      if (t && d > t) return false;
      return true;
    });
  }

  const cutoff = new Date();
  if (period === 'week')    cutoff.setDate(now.getDate() - 7);
  if (period === 'month')   cutoff.setMonth(now.getMonth() - 1);
  if (period === 'quarter') cutoff.setMonth(now.getMonth() - 3);
  if (period === 'year')    { cutoff.setMonth(0, 1); cutoff.setHours(0, 0, 0, 0); }
  return allRows.filter(r => {
    const d = parseDate(get(r, 'Date added'));
    return d && d >= cutoff;
  });
}

function renderMain() {
  const overviewRows = getOverviewRows();
  renderStats(overviewRows);
  renderCharts(overviewRows);
  populateAllJobsFilters(allRows);
  renderAllJobs();
  populateSecondaryFilters(allRows);
  renderActiveQueue();
  renderReports();
  renderPrintedReport();
}

// ── Init & Refresh ────────────────────────────────────────────
let shippingLoaded = false;

async function loadShipping() {
  if (shippingLoaded && shippedRows.length) return;
  document.getElementById('sh-count').textContent = '…';
  try {
    const shipParsed = await fetch(SCRIPT_URL + '?sheet=shipping&t=' + Date.now()).then(r => r.json()).then(d => d.rows || []);
    const shipResult = buildShippedRows(allRows, shipParsed);
    shippedRows = shipResult.matched;
    reviewRows  = shipResult.review;
    shippingLoaded = true;
    fill('sh-owner', [...new Set(shippedRows.map(r => r.owner).filter(Boolean))].sort());
    renderShipping();
    renderReview();
  } catch (err) {
    document.getElementById('sh-count').textContent = 'Error';
  }
}

// ── Calendar ───────────────────────────────────────────────────
let calendarData   = [];   // raw rows from Calendar sheet
let calWeekOffset  = 0;    // 0 = current week, -1 = prev, +1 = next
let calEditRow     = null; // row being edited

const WORKER_COLORS = ['#3b82f6','#ec4899','#22c55e','#f97316','#a855f7','#eab308','#14b8a6'];
const workerColor = (() => {
  const map = {};
  let idx = 0;
  return (name) => {
    if (!name) return '#94a3b8';
    if (!map[name]) map[name] = WORKER_COLORS[idx++ % WORKER_COLORS.length];
    return map[name];
  };
})();

async function loadCalendar() {
  const grid = document.getElementById('cal-week-grid');
  if (!grid) return;
  grid.innerHTML = '<div style="color:var(--text-3);font-style:italic;padding:20px 0;">Loading calendar…</div>';
  try {
    const raw = await fetch(SCRIPT_URL + '?sheet=calendar&raw=1&t=' + Date.now()).then(r => r.json());
    const rows = raw.raw || [];
    // Find header row
    let hdrIdx = rows.findIndex(r => String(r[0]).toLowerCase().trim() === 'date' && String(r[2]).toLowerCase().trim() === 'who');
    if (hdrIdx < 0) { grid.innerHTML = '<div style="color:var(--red)">Could not read Calendar sheet headers.</div>'; return; }
    const hdr = rows[hdrIdx].map(h => String(h).trim().toLowerCase());
    const ci  = (kw) => hdr.findIndex(h => h.includes(kw));
    const cols = {
      date: ci('date'), day: ci('day'), who: ci('who'),
      start: ci('start'), end: ci('end'), hoursTotal: ci('hours total'),
      hoursPrint: ci('hours to print'), expected: ci('expected'),
    };
    // Google Sheets stores time values as Date objects anchored to 1899-12-30 (its epoch).
    // This helper extracts HH:MM from those, or passes plain strings through unchanged.
    const parseSheetTime = (val) => {
      if (!val) return '';
      const s = String(val);
      if (s.includes('1899-12-3')) {
        // Extract UTC HH:MM — Sheets time fractions are in UTC
        const d = new Date(val);
        return String(d.getUTCHours()).padStart(2,'0') + ':' + String(d.getUTCMinutes()).padStart(2,'0');
      }
      if (/^\d{1,2}:\d{2}/.test(s)) return s.substring(0,5); // already HH:MM
      return '';
    };

    // Parse data rows (skip header rows and blank rows)
    calendarData = [];
    for (let i = hdrIdx + 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r[cols.date]) continue;
      const dateVal = r[cols.date];
      if (String(dateVal).toLowerCase().trim() === 'date') continue; // repeated header
      const dateObj = new Date(dateVal);
      if (isNaN(dateObj)) continue;
      // Normalise to midnight local time (dates in sheet are 23:00 UTC = midnight Belgium)
      const d = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate());
      calendarData.push({
        _row:       i + 1,          // 1-based sheet row
        date:       d,
        day:        String(r[cols.day] ?? ''),
        who:        String(r[cols.who] ?? '').trim(),
        startTime:  parseSheetTime(r[cols.start]),
        endTime:    parseSheetTime(r[cols.end]),
        hoursTotal: parseFloat(r[cols.hoursTotal]) || 0,
        hoursPrint: parseFloat(r[cols.hoursPrint]) || 0,
        expected:   parseInt(r[cols.expected])     || 0,
      });
    }
    // Populate worker autocomplete from existing "who" values
    const workers = [...new Set(calendarData.map(r => r.who).filter(Boolean))].sort();
    const dl = document.getElementById('cal-who-list');
    if (dl) dl.innerHTML = workers.map(w => `<option value="${w}">`).join('');

    renderCalendar();
  } catch(err) {
    grid.innerHTML = '<div style="color:var(--red)">Error loading calendar.</div>';
  }
}

function getWeekDays(weekOffset) {
  const now   = new Date();
  const day   = now.getDay(); // 0=Sun
  const mon   = new Date(now);
  mon.setDate(now.getDate() - (day === 0 ? 6 : day - 1) + weekOffset * 7);
  mon.setHours(0, 0, 0, 0);
  return Array.from({ length: 7 }, (_, i) => {
    const d = new Date(mon);
    d.setDate(mon.getDate() + i);
    return d;
  });
}

// Store entries indexed for onclick access (avoids JSON-in-HTML escaping issues)
const _calEntryMap = {};

function renderCalendar() {
  const grid    = document.getElementById('cal-week-grid');
  const label   = document.getElementById('cal-week-label');
  const summary = document.getElementById('cal-summary');
  if (!grid) return;

  const days      = getWeekDays(calWeekOffset);
  const today     = new Date(); today.setHours(0,0,0,0);
  const fmt       = (d) => d.toLocaleDateString('en-GB', { day: 'numeric', month: 'short' });
  const fmtFull   = (d) => d.toLocaleDateString('en-GB', { weekday: 'short', day: 'numeric', month: 'short' });
  const DAY_NAMES = ['Mon','Tue','Wed','Thu','Fri','Sat','Sun'];

  label.textContent = `${fmt(days[0])} – ${fmt(days[6])} ${days[0].getFullYear()}`;

  // Group all entries by date string
  const byDate = {};
  calendarData.forEach(r => {
    const key = r.date.toDateString();
    if (!byDate[key]) byDate[key] = [];
    byDate[key].push(r);
  });

  // Group PrintLog actual output by date
  const logByDate = {};
  printLogRows.forEach(r => {
    const d = parseDate(r['Date']);
    if (!d) return;
    const key = d.toDateString();
    if (!logByDate[key]) logByDate[key] = [];
    logByDate[key].push(r);
  });

  // Week totals (all entries for all days)
  const allWeekEntries = days.flatMap(d => byDate[d.toDateString()] || []);
  const totalHours     = allWeekEntries.reduce((s, r) => s + r.hoursTotal, 0);
  const totalExpected  = allWeekEntries.reduce((s, r) => s + r.expected,   0);
  const workers        = [...new Set(allWeekEntries.map(r => r.who).filter(Boolean))];
  summary.innerHTML = [
    `<span class="cal-summary-stat"><strong>${totalHours.toFixed(1)}h</strong> this week</span>`,
    totalExpected ? `<span class="cal-summary-stat"><strong>${totalExpected}</strong> expected products</span>` : '',
    workers.length ? `<span class="cal-summary-stat"><strong>${workers.join(', ')}</strong></span>` : '',
  ].filter(Boolean).join('<span style="color:var(--border)">|</span>');

  // Clear entry map and rebuild
  Object.keys(_calEntryMap).forEach(k => delete _calEntryMap[k]);

  grid.innerHTML = days.map((d, i) => {
    const entries   = byDate[d.toDateString()] || [];
    const isToday   = d.toDateString() === today.toDateString();
    const isPast    = d < today;
    const isWeekend = i >= 5;
    const classes   = ['cal-day-card', isToday ? 'today' : '', isPast && !isToday ? 'past' : '', isWeekend ? 'weekend' : ''].filter(Boolean).join(' ');
    const dateLabel = fmtFull(d);

    let body;
    const filled = entries.filter(e => e.who || e.hoursTotal > 0);
    if (!filled.length) {
      const key = `${d.toDateString()}_0`;
      _calEntryMap[key] = { ref: { sheetRow: entries[0]?._row, date: dateLabel }, entry: entries[0] || {} };
      body = `<div class="cal-day-empty" onclick="_calClick('${key}')">No schedule — click to add</div>`;
    } else {
      body = filled.map((entry, ei) => {
        const key = `${d.toDateString()}_${ei}`;
        _calEntryMap[key] = { ref: { sheetRow: entry._row, date: dateLabel }, entry };
        const wc = workerColor(entry.who);
        const hrs = entry.hoursTotal > 0 ? entry.hoursTotal.toFixed(entry.hoursTotal % 1 === 0 ? 0 : 2) : null;
        const hpr = entry.hoursPrint  > 0 ? entry.hoursPrint.toFixed(entry.hoursPrint  % 1 === 0 ? 0 : 2) : null;
        return `<div class="cal-worker-row" onclick="_calClick('${key}')">
          <div class="cal-who"><span class="cal-who-dot" style="background:${wc}"></span><strong>${entry.who || '—'}</strong></div>
          ${entry.startTime && entry.endTime ? `<div class="cal-hours">🕐 ${entry.startTime} – ${entry.endTime}</div>` : ''}
          ${hrs ? `<div class="cal-hours">${hrs}h${hpr ? `, ${hpr}h printing` : ''}</div>` : ''}
          ${entry.expected > 0 ? `<div class="cal-expected">~${entry.expected} products</div>` : ''}
          <div class="cal-edit-link">✏️ Edit</div>
        </div>`;
      }).join('<div class="cal-worker-divider"></div>');
    }

    // Actual output from PrintLog for past days
    let actualHtml = '';
    if (isPast || isToday) {
      const logEntries = logByDate[d.toDateString()] || [];
      if (logEntries.length > 0) {
        const byOwner = {};
        logEntries.forEach(e => {
          const o = e['Owner'] || 'Unknown';
          byOwner[o] = (byOwner[o] || 0) + (parseInt(e['Quantity']) || 0);
        });
        const total = Object.values(byOwner).reduce((s, n) => s + n, 0);
        const ownerLines = Object.entries(byOwner)
          .map(([o, q]) => `<span style="color:${workerColor(o)};font-weight:600;">${o}</span> ${q}`)
          .join(', ');
        actualHtml = `<div class="cal-actual-output">
          <div class="cal-actual-label">✅ Actual output</div>
          <div class="cal-actual-total">${total} printed</div>
          <div class="cal-actual-detail">${ownerLines}</div>
        </div>`;
      }
    }

    return `<div class="${classes}">
      <div class="cal-day-header">
        <span class="cal-day-name">${DAY_NAMES[i]}</span>
        <span class="cal-day-date">${d.getDate()}</span>
      </div>
      <div class="cal-day-body">${body}${actualHtml}</div>
    </div>`;
  }).join('');
}

function _calClick(key) {
  const item = _calEntryMap[key];
  if (item) openCalendarEdit(item.ref, item.entry);
}

function calNav(dir) {
  if (dir === 0) calWeekOffset = 0;
  else calWeekOffset += dir;
  renderCalendar();
}

function openCalendarEdit(ref, entry) {
  calEditRow = ref;
  document.getElementById('cal-edit-title').textContent = ref.date || 'Edit Day';
  document.getElementById('cal-who').value          = entry.who        || '';
  document.getElementById('cal-start').value        = entry.startTime  || '';
  document.getElementById('cal-end').value          = entry.endTime    || '';
  document.getElementById('cal-hours-print').value  = entry.hoursPrint > 0 ? entry.hoursPrint : '';
  document.getElementById('cal-expected').value     = entry.expected   > 0 ? entry.expected   : '';
  const st = document.getElementById('cal-edit-status');
  st.style.display = 'none'; st.textContent = '';
  document.getElementById('cal-edit-submit').disabled = false;
  document.getElementById('cal-edit-overlay').style.display = 'flex';
}

function closeCalendarEdit() {
  document.getElementById('cal-edit-overlay').style.display = 'none';
}

document.getElementById('cal-edit-overlay').addEventListener('click', function(e) {
  if (e.target === this) closeCalendarEdit();
});

async function submitCalendarEdit() {
  const who      = document.getElementById('cal-who').value.trim();
  const start    = document.getElementById('cal-start').value;
  const end      = document.getElementById('cal-end').value;
  const hPrint   = document.getElementById('cal-hours-print').value;
  const expected = document.getElementById('cal-expected').value;
  const statusEl = document.getElementById('cal-edit-status');
  const btn      = document.getElementById('cal-edit-submit');

  if (!calEditRow?.sheetRow) {
    statusEl.className = 'form-status error'; statusEl.style.display = '';
    statusEl.textContent = 'Could not identify the sheet row. Try reloading the calendar.';
    return;
  }

  btn.disabled = true; btn.textContent = 'Saving…';
  statusEl.style.display = 'none';

  try {
    await fetch(SCRIPT_URL, {
      method: 'POST', mode: 'no-cors',
      body: JSON.stringify({
        action: 'update_calendar',
        sheetRow:        calEditRow.sheetRow,
        who, startTime:  start, endTime: end,
        hoursToPrint:    hPrint   || 0,
        expectedProducts: expected || 0,
        changedBy:       currentUser?.email,
      }),
    });
    btn.textContent = 'Save';
    btn.disabled = false;
    statusEl.className = 'form-status success'; statusEl.style.display = '';
    statusEl.textContent = 'Saved!';
    setTimeout(async () => { closeCalendarEdit(); await loadCalendar(); }, 800);
  } catch(err) {
    statusEl.className = 'form-status error'; statusEl.style.display = '';
    statusEl.textContent = 'Error saving. Try again.';
    btn.disabled = false; btn.textContent = 'Save';
  }
}

// ── Stock ──────────────────────────────────────────────────────
let stockRows = [];

async function loadStock() {
  const container = document.getElementById('stock-content');
  if (!container) return;
  container.innerHTML = '<div class="loading-msg">Loading stock…</div>';
  try {
    const raw = await fetch(SCRIPT_URL + '?sheet=stock&t=' + Date.now()).then(r => r.json());
    stockRows = raw.rows || [];
    renderStock();
  } catch (err) {
    container.innerHTML = '<div class="loading-msg">Error loading stock.</div>';
  }
}

const STOCK_TYPE_CONFIG = {
  'Bottle':        { icon: '🍶', bg: '#dbeafe', text: '#1d4ed8', bar: '#3b82f6' },
  'Mug':           { icon: '☕', bg: '#fce7f3', text: '#be185d', bar: '#ec4899' },
  'Travel Bottle': { icon: '🧋', bg: '#dcfce7', text: '#15803d', bar: '#22c55e' },
  'Tumbler':       { icon: '🥤', bg: '#fef3c7', text: '#b45309', bar: '#f59e0b' },
  'Bottle lids':   { icon: '🔩', bg: '#ede9fe', text: '#6d28d9', bar: '#8b5cf6' },
  'Mug lids':      { icon: '🔩', bg: '#f3e8ff', text: '#7e22ce', bar: '#a855f7' },
};

const COLOR_SWATCHES = {
  'white': '#f8fafc', 'off white': '#f5f0e8', 'cream': '#fefce8', 'beige': '#e8d9c0',
  'black': '#1e293b', 'dark grey': '#374151', 'grey': '#9ca3af', 'gray': '#9ca3af', 'silver': '#cbd5e1',
  'red': '#ef4444', 'dark red': '#991b1b', 'coral': '#fb7185', 'pink': '#ec4899', 'light pink': '#f9a8d4', 'rose': '#f43f5e',
  'blue': '#3b82f6', 'dark blue': '#1e40af', 'navy': '#1e3a5f', 'light blue': '#93c5fd', 'sky blue': '#7dd3fc',
  'green': '#22c55e', 'dark green': '#15803d', 'mint': '#6ee7b7', 'teal': '#14b8a6', 'olive': '#84cc16',
  'yellow': '#eab308', 'gold': '#ca8a04', 'orange': '#f97316', 'amber': '#f59e0b',
  'purple': '#a855f7', 'violet': '#7c3aed', 'lavender': '#c4b5fd',
  'brown': '#92400e', 'tan': '#d4a574',
  'transparent': 'rgba(0,0,0,0.05)',
};

function colorSwatch(colorName) {
  const key = (colorName || '').toLowerCase().trim();
  const hex = COLOR_SWATCHES[key];
  if (!hex) return `<span class="stock-swatch stock-swatch-unknown" title="${colorName}"></span>`;
  const needsBorder = ['#f8fafc','#f5f0e8','#fefce8','#e8d9c0','rgba(0,0,0,0.05)'].includes(hex);
  return `<span class="stock-swatch" style="background:${hex};${needsBorder ? 'border:1.5px solid #cbd5e1;' : ''}" title="${colorName}"></span>`;
}

function renderStock() {
  const container = document.getElementById('stock-content');
  if (!container) return;
  if (!stockRows.length) { container.innerHTML = '<div class="stock-empty">No stock data found.</div>'; return; }

  const groups = {};
  stockRows.forEach(r => {
    const type = r['Type'] || 'Unknown';
    if (!groups[type]) groups[type] = [];
    groups[type].push(r);
  });

  // Summary stats
  const PRODUCT_TYPES = ['Bottle', 'Mug', 'Travel Bottle', 'Tumbler'];
  const LID_TYPES     = ['Bottle lids', 'Mug lids'];
  const sumQty = (types) => stockRows
    .filter(r => types.includes(r['Type']))
    .reduce((s, r) => s + (parseInt(r['Quantity'])||0), 0);
  const productsTotal = sumQty(PRODUCT_TYPES);
  const lidsTotal     = sumQty(LID_TYPES);
  const lowCount  = stockRows.filter(r => { const q = parseInt(r['Quantity'])||0; return q > 0 && q < 100; }).length;
  const outCount  = stockRows.filter(r => (parseInt(r['Quantity'])||0) === 0).length;

  const summaryHtml = `
    <div class="stock-summary">
      <div class="stock-stat">
        <div class="stock-stat-value">${productsTotal.toLocaleString()}</div>
        <div class="stock-stat-label">Products in stock</div>
      </div>
      <div class="stock-stat">
        <div class="stock-stat-value">${lidsTotal.toLocaleString()}</div>
        <div class="stock-stat-label">Spare lids in stock</div>
      </div>
      <div class="stock-stat ${lowCount > 0 ? 'warning' : ''}">
        <div class="stock-stat-value">${lowCount}</div>
        <div class="stock-stat-label">Low stock</div>
      </div>
      <div class="stock-stat ${outCount > 0 ? 'danger' : ''}">
        <div class="stock-stat-value">${outCount}</div>
        <div class="stock-stat-label">Out of stock</div>
      </div>
    </div>`;

  const buildGroupHtml = (type) => {
    if (!groups[type]) return '';
    const grpRows   = groups[type];
    const cfg       = STOCK_TYPE_CONFIG[type] || { icon: '📦', bg: '#f1f5f9', text: '#475569', bar: '#94a3b8' };
    const typeTotal = grpRows.reduce((s, r) => s + (parseInt(r['Quantity'])||0), 0);
    const maxQty    = Math.max(...grpRows.map(r => parseInt(r['Quantity'])||0), 1);

    const rowsHtml = grpRows.map(r => {
      const qty      = parseInt(r['Quantity']) || 0;
      const pct      = Math.round((qty / maxQty) * 100);
      const qClass   = qty === 0 ? 'danger' : qty < 100 ? 'warning' : 'ok';
      const barColor = qty === 0 ? '#ef4444' : qty < 100 ? '#f97316' : cfg.bar;
      return `<div class="stock-row">
        ${colorSwatch(r['Color'])}
        <span class="stock-color-name">${r['Color'] || '—'}</span>
        <div class="stock-bar-wrap">
          <div class="stock-bar" style="width:${pct}%;background:${barColor}"></div>
        </div>
        <span class="stock-qty ${qClass}">${qty.toLocaleString()}</span>
        ${r['Levering'] ? `<span class="stock-levering">${r['Levering']}</span>` : ''}
        <button class="stock-add-btn" onclick="openDeliveryModal('${type}','${r['Color']}')" title="Add delivery">+</button>
      </div>`;
    }).join('');

    return `<div class="stock-group">
      <div class="stock-group-header" style="background:${cfg.bg};color:${cfg.text}">
        <span class="stock-group-icon">${cfg.icon}</span>
        <span class="stock-group-name">${type}</span>
        <span class="stock-group-total">${typeTotal.toLocaleString()}</span>
      </div>
      <div class="stock-rows">${rowsHtml}</div>
    </div>`;
  };

  const productRow = ['Bottle', 'Mug', 'Travel Bottle', 'Tumbler'].map(buildGroupHtml).join('');
  const lidRow     = ['Bottle lids', 'Mug lids'].map(buildGroupHtml).join('');

  container.innerHTML = summaryHtml
    + `<div class="stock-groups stock-groups-products">${productRow}</div>`
    + `<div class="stock-groups stock-groups-lids">${lidRow}</div>`;
}

function openDeliveryModal(preType, preColor) {
  if (!stockRows.length) { alert('Load the Stock tab first.'); return; }

  const types = [...new Set(stockRows.map(r => r['Type']).filter(Boolean))];
  const typeSelect  = document.getElementById('del-type');
  const colorSelect = document.getElementById('del-color');
  typeSelect.innerHTML = types.map(t => `<option value="${t}">${t}</option>`).join('');
  if (preType) typeSelect.value = preType;

  const updateColors = (selectColor) => {
    const chosen = typeSelect.value;
    const colors = stockRows.filter(r => r['Type'] === chosen).map(r => r['Color']).filter(Boolean);
    colorSelect.innerHTML = colors.map(c => `<option value="${c}">${c}</option>`).join('');
    if (selectColor) colorSelect.value = selectColor;
  };
  typeSelect.onchange = () => updateColors();
  updateColors(preColor);

  document.getElementById('del-quantity').value = '';
  document.getElementById('del-note').value = '';
  const st = document.getElementById('del-status');
  st.style.display = 'none'; st.textContent = '';
  document.getElementById('del-submit').disabled = false;
  document.getElementById('delivery-modal-overlay').style.display = 'flex';
}

function closeDeliveryModal() {
  document.getElementById('delivery-modal-overlay').style.display = 'none';
}

document.getElementById('delivery-modal-overlay').addEventListener('click', function(e) {
  if (e.target === this) closeDeliveryModal();
});

async function submitDelivery() {
  const type     = document.getElementById('del-type').value;
  const color    = document.getElementById('del-color').value;
  const quantity = parseInt(document.getElementById('del-quantity').value);
  const note     = document.getElementById('del-note').value.trim();
  const statusEl = document.getElementById('del-status');
  const btn      = document.getElementById('del-submit');

  if (!type || !color || !quantity || quantity < 1) {
    statusEl.className = 'form-status error'; statusEl.style.display = '';
    statusEl.textContent = 'Please fill in type, color and quantity.'; return;
  }

  btn.disabled = true; btn.textContent = 'Saving…';
  statusEl.style.display = 'none';

  try {
    await fetch(SCRIPT_URL, {
      method: 'POST', mode: 'no-cors',
      body: JSON.stringify({ action: 'add_stock', type, color, quantity, note, changedBy: currentUser?.email }),
    });
    statusEl.className = 'form-status success'; statusEl.style.display = '';
    statusEl.textContent = `Added ${quantity}× ${color} ${type} to stock.`;
    btn.textContent = 'Add to Stock';
    // Refresh stock display after short delay
    setTimeout(async () => { closeDeliveryModal(); await loadStock(); }, 1200);
  } catch (err) {
    statusEl.className = 'form-status error'; statusEl.style.display = '';
    statusEl.textContent = 'Error saving. Try again.';
    btn.disabled = false; btn.textContent = 'Add to Stock';
  }
}

async function refreshData() {
  clearSelection();
  const lu = document.getElementById('last-updated');
  lu.textContent = 'Updating…';

  // Show cached data instantly while fresh data loads
  const cached = loadMainCache();
  if (cached && cached.length > 0) {
    allRows = cached;
    try { renderMain(); } catch (_) {}
  } else {
    // No cache — make clear something is happening
    document.querySelectorAll('.stat-value').forEach(el => { el.textContent = '…'; });
    lu.textContent = 'Loading data… (first load may take 15–30s)';
  }

  try {
    // Use Promise.race for timeout — works in all browsers without AbortController quirks
    const fetchPromise = fetch(SCRIPT_URL + '?t=' + Date.now()).then(r => r.json());
    const timeoutPromise = new Promise((_, reject) =>
      setTimeout(() => reject(new Error('timeout')), 35000)
    );

    const [fetchData] = await Promise.all([
      Promise.race([fetchPromise, timeoutPromise]),
      loadInvoices(),
    ]);
    const parsed  = (fetchData.rows || []).filter(r => get(r,'Name_Company') && get(r,'Priority') && get(r,'Priority') !== '0')
      .map(r => { if (r['Owner']) r['Owner'] = normOwner(r['Owner']); return r; });
    printLogRows = fetchData.printLog || [];

    if (parsed.length > 0) {
      allRows = parsed;
      saveMainCache(allRows);
      try { renderMain(); } catch (_) {}
    } else if (!cached || !cached.length) {
      lu.textContent = '⚠️ No data — click ↻ Refresh';
      return;
    }

    shippingLoaded = false;
    if (document.getElementById('tab-shipping').classList.contains('active')) {
      loadShipping();
    }

    const now = new Date().toLocaleString();
    lu.textContent = now;
    document.getElementById('last-updated-footer').textContent = now;

  } catch (err) {
    if (err.message === 'timeout') {
      lu.textContent = cached && cached.length
        ? '⚠️ Slow connection — showing cached data'
        : '⚠️ Timed out — click ↻ Refresh to retry';
    } else {
      lu.textContent = cached && cached.length
        ? '⚠️ Offline — showing cached data'
        : '⚠️ Error loading — click ↻ Refresh';
    }
  }
}

// ── Print Logger Modal ────────────────────────────────────────
let modalJob = null;

function openPrintModal(rowIdx) {
  if (!SCRIPT_URL) {
    alert('Print logging is not configured yet.\nAsk your admin to set up the SCRIPT_URL in app.js.');
    return;
  }
  modalJob = allRows[rowIdx];
  if (!modalJob) return;

  // Populate job info
  document.getElementById('modal-job-info').innerHTML = `
    <div class="modal-job-card">
      <div><span class="modal-label">Job</span><strong>#${get(modalJob,'Priority')} — ${get(modalJob,'Name_Company')}</strong></div>
      <div><span class="modal-label">Print</span>${get(modalJob,'Name_Print') || '—'}</div>
      <div><span class="modal-label">Type</span>${typeBadge(get(modalJob,'Soort'))}</div>
      <div><span class="modal-label">Color</span>${get(modalJob,'Bottle color') || '—'}</div>
      <div><span class="modal-label">Total Qty</span>${num(modalJob,'Quantity')}</div>
      <div><span class="modal-label">Already Printed</span>${num(modalJob,'Quantity printed ') || num(modalJob,'Quantity printed') || 0}</div>
    </div>`;

  // Pre-fill: session input starts at 0; existing total shown in job info above
  document.getElementById('modal-printed').value = '';
  document.getElementById('modal-faulty').value  = num(modalJob,'Faulty prints') || 0;

  // Printer field — shown only for bottles; auto-assigned for other types
  const soort = get(modalJob,'Soort').toLowerCase();
  const printerField = document.getElementById('modal-printer-field');
  const sel = document.getElementById('modal-printer');
  const isBottle = soort.includes('bottle') && !soort.includes('travel');

  if (isBottle) {
    printerField.style.display = '';
    sel.innerHTML = '<option value="">Select printer...</option>' +
      ['Bottle 1','Bottle 2'].map(p => `<option value="${p}">${p}</option>`).join('');
    const cur = get(modalJob,'Printer');
    if (cur) sel.value = cur;
  } else {
    printerField.style.display = 'none';
    // Auto-assign printer based on type
    if (soort.includes('travel')) sel.value = 'Travel Bottle 1';
    else if (soort.includes('mug') || soort.includes('tumbler')) sel.value = 'Mug 1';
    else sel.value = '';
  }

  // Reset photo
  document.getElementById('modal-photo').value = '';
  document.getElementById('modal-preview').style.display = 'none';
  document.getElementById('modal-status').textContent = '';

  // Phone camera QR — show on desktop only
  const phoneSection = document.getElementById('modal-phone-section');
  phoneSection.style.display = window.innerWidth > 768 ? '' : 'none';
  document.getElementById('modal-qr-wrap').style.display = 'none';
  document.getElementById('modal-qr-status').textContent = '';
  window._phonePhotoUrl = null;
  clearInterval(window._photoSessionInterval);

  document.getElementById('print-modal-overlay').style.display = 'flex';
  document.getElementById('modal-printed').focus();
}

function closeModal() {
  clearInterval(window._photoSessionInterval);
  window._phonePhotoUrl = null;
  document.getElementById('print-modal-overlay').style.display = 'none';
  modalJob = null;
}

function startPhonePhotoSession() {
  const sessionKey = 'izy_' + Date.now();
  const priority   = get(modalJob, 'Priority');
  const company    = encodeURIComponent(get(modalJob, 'Name_Company'));
  const base       = location.href.replace(/[?#].*$/, '').replace(/\/[^/]*$/, '/');
  const photoUrl   = `${base}photo.html?session=${sessionKey}&priority=${priority}&company=${company}`;
  const qrSrc      = `https://api.qrserver.com/v1/create-qr-code/?size=200x200&color=1d4ed8&bgcolor=ffffff&data=${encodeURIComponent(photoUrl)}`;

  document.getElementById('modal-qr-img').src = qrSrc;
  document.getElementById('modal-qr-wrap').style.display = 'block';
  document.getElementById('modal-qr-status').textContent = '⏳ Waiting for photo…';
  document.getElementById('modal-qr-status').style.color = 'var(--text-3)';

  clearInterval(window._photoSessionInterval);
  window._photoSessionInterval = setInterval(async () => {
    try {
      const res  = await fetch(SCRIPT_URL + '?photosession=' + sessionKey + '&t=' + Date.now());
      const data = await res.json();
      if (data.photoUrl) {
        clearInterval(window._photoSessionInterval);
        window._phonePhotoUrl = data.photoUrl;
        const preview = document.getElementById('modal-preview');
        preview.src = data.photoUrl;
        preview.style.display = 'block';
        document.getElementById('modal-qr-status').textContent = '✅ Photo received!';
        document.getElementById('modal-qr-status').style.color = '#22c55e';
      }
    } catch (_) {}
  }, 3000);
}

// Close on overlay click
document.getElementById('print-modal-overlay').addEventListener('click', function(e) {
  if (e.target === this) closeModal();
});

function handlePhotoPreview(input) {
  const preview = document.getElementById('modal-preview');
  if (input.files && input.files[0]) {
    const reader = new FileReader();
    reader.onload = e => {
      preview.src = e.target.result;
      preview.style.display = 'block';
    };
    reader.readAsDataURL(input.files[0]);
  } else {
    preview.style.display = 'none';
  }
}

function compressImage(file, maxWidth = 1400) {
  return new Promise(resolve => {
    const reader = new FileReader();
    reader.onload = e => {
      const img = new Image();
      img.onload = () => {
        const scale  = Math.min(1, maxWidth / img.width);
        const canvas = document.createElement('canvas');
        canvas.width  = Math.round(img.width  * scale);
        canvas.height = Math.round(img.height * scale);
        canvas.getContext('2d').drawImage(img, 0, 0, canvas.width, canvas.height);
        resolve(canvas.toDataURL('image/jpeg', 0.82));
      };
      img.src = e.target.result;
    };
    reader.readAsDataURL(file);
  });
}

async function markSleeved(rowIdx, btn) {
  const job = allRows[rowIdx];
  if (!job) return;
  const label = get(job,'Name_Company') + ' #' + get(job,'Priority');
  if (!confirm(`Mark "${label}" as sleeved?\nThis will set the status to "Ready to Ship".`)) return;
  // Immediately flip button to green
  if (btn) {
    btn.textContent = '✓ Sleeved';
    btn.classList.add('sleeved');
    btn.disabled = true;
  }
  try {
    await fetch(SCRIPT_URL, {
      method: 'POST',
      mode:   'no-cors',
      body:   JSON.stringify({ sheetRow: job['_sheetRow'], status: 'Ready to Ship', changedBy: currentUser?.email }),
    });
    refreshData();
  } catch (err) {
    // Revert button on failure
    if (btn) { btn.textContent = '✕ Sleeve'; btn.classList.remove('sleeved'); btn.disabled = false; }
    alert('Could not update: ' + err.message);
  }
}

async function unsleeveJob(rowIdx, btn) {
  const job = allRows[rowIdx];
  if (!job) return;
  const label = get(job,'Name_Company') + ' #' + get(job,'Priority');
  if (!confirm(`Undo sleeving for "${label}"?\nStatus will go back to "Waiting".`)) return;
  if (btn) { btn.textContent = '✕ Sleeve'; btn.classList.remove('sleeved'); btn.disabled = false; }
  try {
    await fetch(SCRIPT_URL, {
      method: 'POST',
      mode:   'no-cors',
      body:   JSON.stringify({ sheetRow: job['_sheetRow'], status: 'Waiting', changedBy: currentUser?.email }),
    });
    refreshData();
  } catch (err) {
    if (btn) { btn.textContent = '✓ Sleeved'; btn.classList.add('sleeved'); btn.disabled = true; }
    alert('Could not update: ' + err.message);
  }
}

let shipModalRowIdx  = null;
let shipSelectedRate = null;

const PKG_DEFAULTS  = { length: 42, width: 42, height: 31 };
const PALLET_DEFAULTS = { length: 120, width: 80, height: '' };

function applyPkgTypeDefaults(row) {
  const type = row.querySelector('.pkg-type').value;
  const d = type === 'PALLET' ? PALLET_DEFAULTS : PKG_DEFAULTS;
  row.querySelector('.pkg-length').value = d.length;
  row.querySelector('.pkg-width').value  = d.width;
  row.querySelector('.pkg-height').value = d.height;
}

function addShipPackageRow() {
  const list = document.getElementById('ship-packages-list');
  const row  = document.createElement('div');
  row.className = 'ship-pkg-row';
  row.innerHTML = `
    <div class="form-group"><label>Type</label><select class="pkg-type"><option value="PACKAGE">Package</option><option value="PALLET">Pallet</option></select></div>
    <div class="form-group"><label>L (cm)</label><input type="number" class="pkg-length" value="42" min="1" /></div>
    <div class="form-group"><label>W (cm)</label><input type="number" class="pkg-width"  value="42" min="1" /></div>
    <div class="form-group"><label>H (cm)</label><input type="number" class="pkg-height" value="31" min="1" /></div>
    <div class="form-group"><label>kg</label><input type="number" class="pkg-weight" value="12" min="0.1" step="0.1" /></div>
    <button type="button" class="btn-remove-file ship-pkg-remove" title="Remove">✕</button>`;
  row.querySelector('.pkg-type').addEventListener('change', () => { applyPkgTypeDefaults(row); resetShipRates(); });
  row.querySelector('.ship-pkg-remove').addEventListener('click', () => { row.remove(); shipSelectedRate = null; resetShipRates(); });
  list.appendChild(row);
}

function resetShipRates() {
  document.getElementById('ship-rates-list').style.display = 'none';
  document.getElementById('ship-rates-list').innerHTML = '';
  document.getElementById('ship-rates-status').textContent = '';
  shipSelectedRate = null;
}

function getShipPackages() {
  return [...document.getElementById('ship-packages-list').querySelectorAll('.ship-pkg-row')].map(row => ({
    type:   row.querySelector('.pkg-type').value || 'PACKAGE',
    length: parseFloat(row.querySelector('.pkg-length').value) || 42,
    width:  parseFloat(row.querySelector('.pkg-width').value)  || 42,
    height: parseFloat(row.querySelector('.pkg-height').value) || 0,
    weight: parseFloat(row.querySelector('.pkg-weight').value) || 12,
  }));
}

async function loadShipRates() {
  const zipcode = document.getElementById('ship-zipcode').value.trim();
  const city    = document.getElementById('ship-city').value.trim();
  const country = document.getElementById('ship-country').value.trim().toUpperCase() || 'NL';
  if (!zipcode || !city) {
    document.getElementById('ship-rates-status').textContent = 'Fill in postcode and city first.';
    return;
  }
  const statusEl = document.getElementById('ship-rates-status');
  const listEl   = document.getElementById('ship-rates-list');
  const btn      = document.getElementById('ship-rates-btn');
  btn.disabled   = true;
  btn.textContent = 'Loading…';
  statusEl.textContent = '';
  listEl.style.display = 'none';
  shipSelectedRate = null;

  const pkgs = getShipPackages();
  const params = new URLSearchParams({
    action:    'get_ship_rates',
    owner:     get(allRows[shipModalRowIdx], 'Owner') || '',
    rcZipcode: zipcode, rcCity: city, rcCountry: country,
    pkgsJson:  JSON.stringify(pkgs),
    t:         Date.now(),
  });

  try {
    const data = await fetch(SCRIPT_URL + '?' + params.toString()).then(r => r.json());
    if (data.error) throw new Error(data.error);
    if (data._debug) console.log('CheapCargo raw response:', data._debug);
    const rates = data.rates || [];
    if (!rates.length) { statusEl.textContent = 'No rates available. Check console for details.'; console.log('CheapCargo debug:', data._debug); btn.disabled = false; btn.textContent = '🔍 Get rates'; return; }

    listEl.innerHTML = rates.map((r, i) => `
      <label class="ship-rate-option">
        <input type="radio" name="ship-rate" value="${i}" />
        <div class="ship-rate-info">
          <div class="ship-rate-name">${r.carrierName} <span class="ship-rate-service">— ${r.serviceLevel}</span></div>
          <div class="ship-rate-dates">Pickup: ${r.pickup || '—'} · Delivery: ${r.delivery || '—'}</div>
        </div>
        <div class="ship-rate-price">€${parseFloat(r.price).toFixed(2)}</div>
      </label>`).join('');

    listEl.querySelectorAll('input[type="radio"]').forEach(radio => {
      radio.addEventListener('change', () => { shipSelectedRate = rates[parseInt(radio.value)]; });
    });

    // Auto-select cheapest
    const firstRadio = listEl.querySelector('input[type="radio"]');
    if (firstRadio) { firstRadio.checked = true; shipSelectedRate = rates[0]; }

    listEl.style.display = 'block';
    statusEl.textContent = '';
  } catch(err) {
    statusEl.textContent = 'Error: ' + err.message;
  }
  btn.disabled = false;
  btn.textContent = '🔍 Get rates';
}

function shipJob(rowIdx) {
  const job = allRows[rowIdx];
  if (!job) return;
  shipModalRowIdx  = rowIdx;
  shipSelectedRate = null;

  document.getElementById('ship-job-info').innerHTML =
    `<strong>#${get(job,'Priority')} — ${get(job,'Name_Company')}</strong> &nbsp;·&nbsp; ${get(job,'Soort') || ''} &nbsp;·&nbsp; Qty: ${get(job,'Quantity') || '—'}`;

  document.getElementById('ship-company').value = get(job,'Name_Company');
  document.getElementById('ship-contact').value = get(job,'Contactpersoon')  || '';
  document.getElementById('ship-phone').value   = get(job,'Telefoonnummer')  || '';
  document.getElementById('ship-email').value   = get(job,'E-mailadres')     || '';
  document.getElementById('ship-street').value  = get(job,'Straat')          || '';
  document.getElementById('ship-number').value  = get(job,'Huisnummer') || '';
  document.getElementById('ship-zipcode').value = get(job,'Postcode')        || '';
  document.getElementById('ship-city').value    = get(job,'Plaats')          || '';
  document.getElementById('ship-country').value = get(job,'Land')            || 'NL';

  // Reset packages — one default row
  document.getElementById('ship-packages-list').innerHTML = '';
  addShipPackageRow();

  resetShipRates();
  document.getElementById('ship-status').textContent  = '';
  document.getElementById('ship-status').className    = 'form-status';
  document.getElementById('ship-result').style.display = 'none';
  document.getElementById('ship-submit').disabled     = false;
  document.getElementById('ship-submit').textContent  = '📦 Book Shipment';
  document.getElementById('ship-manual').disabled     = false;
  document.getElementById('ship-manual').textContent  = '✓ Mark as Shipped';

  document.getElementById('ship-modal-overlay').style.display = 'flex';
}

function closeShipModal() {
  document.getElementById('ship-modal-overlay').style.display = 'none';
  shipModalRowIdx = null;
}

async function markShippedManually() {
  const job = allRows[shipModalRowIdx];
  if (!job) return;
  const btn = document.getElementById('ship-manual');
  btn.disabled = true;
  btn.textContent = 'Saving…';
  try {
    const params = new URLSearchParams({ action: 'mark_shipped', sheetRow: job['_sheetRow'], t: Date.now() });
    const data = await fetch(SCRIPT_URL + '?' + params.toString()).then(r => r.json());
    if (data.error) throw new Error(data.error);
    btn.textContent = '✓ Marked shipped';
    setTimeout(() => { closeShipModal(); refreshData(); }, 1500);
  } catch(err) {
    btn.disabled = false;
    btn.textContent = '✓ Mark as Shipped';
    document.getElementById('ship-status').textContent = 'Error: ' + err.message;
    document.getElementById('ship-status').className = 'form-status error';
  }
}

document.getElementById('ship-modal-overlay').addEventListener('click', function(e) {
  if (e.target === this) closeShipModal();
});

async function submitShipment() {
  const job = allRows[shipModalRowIdx];
  if (!job) return;

  const get_ = id => document.getElementById(id).value.trim();
  const company = get_('ship-company');
  const contact = get_('ship-contact');
  const phone   = get_('ship-phone');
  const email   = get_('ship-email');
  const street  = get_('ship-street');
  const number  = get_('ship-number');
  const zipcode = get_('ship-zipcode');
  const city    = get_('ship-city');
  const country = get_('ship-country').toUpperCase();

  const statusEl = document.getElementById('ship-status');
  if (!company || !contact || !phone || !street || !number || !zipcode || !city || !country) {
    statusEl.textContent = 'Please fill in all required receiver fields.';
    statusEl.className   = 'form-status error';
    return;
  }
  if (!shipSelectedRate) {
    statusEl.textContent = 'Please get rates and select a carrier first.';
    statusEl.className   = 'form-status error';
    return;
  }

  const submitBtn = document.getElementById('ship-submit');
  submitBtn.disabled    = true;
  submitBtn.textContent = 'Booking…';
  statusEl.textContent  = 'Contacting CheapCargo…';
  statusEl.className    = 'form-status';

  const pkgs = getShipPackages();
  try {
    const params = new URLSearchParams({
      action:    'book_shipment',
      sheetRow:  job['_sheetRow'],
      owner:     get(job,'Owner') || '',
      reference: get(job,'Priority'),
      rcCompany: company,
      rcContact: contact,
      rcPhone:   phone,
      rcEmail:   email,
      rcStreet:  street,
      rcNumber:  number,
      rcZipcode: zipcode,
      rcCity:    city,
      rcCountry: country,
      rateId:      shipSelectedRate.id           || '',
      ratePrice:   shipSelectedRate.price        || '',
      ratePickup:  shipSelectedRate.pickup       || '',
      rateDelivery:shipSelectedRate.delivery     || '',
      rateService: shipSelectedRate.serviceLevel || '',
      quantity:    get(job,'Quantity') || '',
      pkgsJson:    JSON.stringify(pkgs),
      t:           Date.now(),
    });
    const resp = await fetch(SCRIPT_URL + '?' + params.toString());

    const result = await resp.json();
    if (result.error) throw new Error(result.error);

    statusEl.textContent = '✓ Shipment booked!';
    statusEl.className   = 'form-status success';

    const resultEl = document.getElementById('ship-result');
    resultEl.style.display = 'block';
    resultEl.innerHTML =
      `<div class="ship-result-title">✓ Shipment booked</div>` +
      `<div class="ship-result-row"><span>Order</span><strong>CC-${result.orderNumber}</strong></div>` +
      `<div class="ship-result-row"><span>Carrier</span><strong>${result.carrier || '—'}</strong></div>` +
      `<div class="ship-result-row"><span>AWB</span><strong>${result.awb || '—'}</strong></div>` +
      (result.trackAndTrace ? `<div class="ship-result-row"><span>Tracking</span><a href="${result.trackAndTrace}" target="_blank" rel="noopener">Track shipment →</a></div>` : '') +
      (result.labelUrl ? `<div class="ship-result-row" style="margin-top:10px;"><a href="${result.labelUrl}" target="_blank" rel="noopener" class="btn btn-secondary" style="width:100%;text-align:center;">🖨 Download shipping label</a></div>` : '');

    submitBtn.textContent = '✓ Booked';
    refreshData();
  } catch (err) {
    statusEl.textContent = 'Error: ' + err.message;
    statusEl.className   = 'form-status error';
    submitBtn.disabled    = false;
    submitBtn.textContent = '📦 Book Shipment';
  }
}

async function resetJob(rowIdx) {
  const job = allRows[rowIdx];
  if (!job) return;
  const label = get(job,'Name_Company') + ' #' + get(job,'Priority');
  if (!confirm(`Reset "${label}"?\nThis will clear quantity printed, faulty prints and printer, and set status back to "To Print".`)) return;
  try {
    await fetch(SCRIPT_URL, {
      method: 'POST',
      mode:   'no-cors',
      body:   JSON.stringify({
        sheetRow:        job['_sheetRow'],
        status:          'To Print',
        quantityPrinted: 0,
        faultyPrints:    0,
        printer:         '',
        changedBy:       currentUser?.email,
      }),
    });
    refreshData();
  } catch (err) {
    alert('Could not reset: ' + err.message);
  }
}

async function submitPrintUpdate() {
  const statusEl = document.getElementById('modal-status');
  const submitBtn = document.getElementById('modal-submit');
  const sessionPrinted = parseInt(document.getElementById('modal-printed').value);
  const faulty  = parseInt(document.getElementById('modal-faulty').value)  || 0;
  const printer = document.getElementById('modal-printer').value;
  const photoFile = document.getElementById('modal-photo').files[0];

  if (!sessionPrinted && sessionPrinted !== 0) { statusEl.textContent = 'Please enter quantity printed.'; statusEl.className = 'modal-error'; return; }

  const alreadyPrinted = num(modalJob,'Quantity printed ') || num(modalJob,'Quantity printed') || 0;
  const printed = alreadyPrinted + sessionPrinted;

  const quantity = num(modalJob,'Quantity') || 0;
  const stillLeft = quantity - printed;
  const needsSleeve = (get(modalJob,'To sleeve?') || getCI(modalJob,'sleeve')).toLowerCase() === 'yes';
  const autoStatus = stillLeft <= 0 ? (needsSleeve ? 'Waiting' : 'Ready to ship') : null;

  submitBtn.disabled = true;
  submitBtn.textContent = 'Submitting…';
  statusEl.textContent = '';

  try {
    let imageBase64 = null;
    let imageMime   = null;
    if (photoFile) {
      imageBase64 = await compressImage(photoFile);
      imageMime   = 'image/jpeg';
    }

    const payload = {
      sheetRow:        modalJob['_sheetRow'],
      priority:        get(modalJob,'Priority'),
      soort:           get(modalJob,'Soort'),
      quantityPrinted: printed,
      sessionPrinted:  sessionPrinted, // just this session — for the PrintLog
      faultyPrints:    faulty,
      printer:         printer,
      ...(autoStatus ? { status: autoStatus } : {}),
      imageBase64,
      imageMime,
      phonePhotoUrl:   window._phonePhotoUrl || null,
      changedBy:       currentUser?.email,
    };

    await fetch(SCRIPT_URL, {
      method: 'POST',
      mode:   'no-cors',
      body:   JSON.stringify(payload),
    });

    submitBtn.textContent = 'Submit';
    submitBtn.disabled = false;
    statusEl.textContent = '✅ Saved! The sheet will update within a few seconds.';
    statusEl.className = 'modal-success';

    // Apply to all other selected AQ rows (no photo for bulk)
    const otherAqIdxs = [...aqSelected].filter(i => allRows[i] !== modalJob);
    for (const i of otherAqIdxs) {
      const r = allRows[i];
      if (!r) continue;
      const alreadyP  = num(r,'Quantity printed ') || num(r,'Quantity printed') || 0;
      const totalP    = alreadyP + sessionPrinted;
      const qty       = num(r,'Quantity') || 0;
      const left      = qty - totalP;
      const needsSl   = (get(r,'To sleeve?') || getCI(r,'sleeve')).toLowerCase() === 'yes';
      const autoSt    = left <= 0 ? (needsSl ? 'Waiting' : 'Ready to ship') : null;
      await fetch(SCRIPT_URL, { method: 'POST', mode: 'no-cors', body: JSON.stringify({
        sheetRow: r['_sheetRow'], priority: get(r,'Priority'), soort: get(r,'Soort'),
        quantityPrinted: totalP, sessionPrinted, faultyPrints: faulty, printer,
        ...(autoSt ? { status: autoSt } : {}), changedBy: currentUser?.email,
      })});
    }
    aqSelected.clear();
    updateSelectionBar();

    // Close after 2 seconds and refresh data
    setTimeout(() => { closeModal(); refreshData(); }, 2200);

  } catch (err) {
    submitBtn.textContent = 'Submit';
    submitBtn.disabled = false;
    statusEl.textContent = '❌ Error: ' + err.message;
    statusEl.className = 'modal-error';
  }
}

// ── Add Job ────────────────────────────────────────────────────
function populateApprovedMockups() {
  const sel = document.getElementById('nj-mockup-select');
  if (!sel) return;
  const approved = mockupRows.filter(r => get(r,'Status').toLowerCase() === 'approved');
  sel.innerHTML = '<option value="">— Select an approved mockup to pre-fill —</option>' +
    approved.map(r => {
      const label = [get(r,'Name_Company'), get(r,'Soort'), get(r,'Bottle color')].filter(Boolean).join(' · ');
      return `<option value="${mockupRows.indexOf(r)}">${label}</option>`;
    }).join('');
}

document.getElementById('nj-mockup-select').addEventListener('change', function() {
  const idx = parseInt(this.value);
  if (isNaN(idx)) return;
  const r = mockupRows[idx];
  if (!r) return;
  const setVal = (id, val) => { const el = document.getElementById(id); if (el && val) el.value = val; };
  setVal('nj-soort',    get(r,'Soort'));
  setVal('nj-company',  get(r,'Name_Company'));
  setVal('nj-color',    get(r,'Bottle color') || get(r,'Color'));
  setVal('nj-lid',      get(r,'Lid'));
  setVal('nj-owner',    normOwner(get(r,'Owner')));
  setVal('nj-deadline', get(r,'Deadline'));
  setVal('nj-quantity', get(r,'Quantity'));
  // Trigger priority hint update
  document.getElementById('nj-soort').dispatchEvent(new Event('change'));
  // Update progress bar
  document.getElementById('nj-company').dispatchEvent(new Event('input'));
});

function populateAddJobOwners() {
  const owners = [...new Set([...KNOWN_OWNERS, ...allRows.map(r => get(r,'Owner')).filter(Boolean)])].sort();
  const sel = document.getElementById('nj-owner');
  sel.innerHTML = '<option value="">Select owner…</option>' +
    owners.map(o => `<option value="${o}">${o}</option>`).join('');
}

document.getElementById('nj-soort').addEventListener('change', function() {
  const soort = this.value.toLowerCase();
  const hint  = document.getElementById('nj-priority-hint');
  if (!soort || !allRows.length) { hint.textContent = ''; return; }
  const max = Math.max(0, ...allRows
    .filter(r => get(r,'Soort').toLowerCase() === soort)
    .map(r => parseInt(get(r,'Priority')) || 0));
  hint.textContent = max > 0 ? `(max: ${max})` : '';
});

//// Design files: dynamic multi-file rows
function addNjFileRow() {
  const container = document.getElementById('nj-files-list');
  const row = document.createElement('div');
  row.className = 'sv-file-row';
  row.innerHTML = `<input type="file" accept="*/*" /><button type="button" class="btn-remove-file" title="Remove">✕</button>`;
  row.querySelector('.btn-remove-file').addEventListener('click', () => row.remove());
  container.appendChild(row);
}
document.getElementById('nj-add-file').addEventListener('click', addNjFileRow);
// Start with one row when the form is first shown
addNjFileRow();

// Mockup file: update label display
document.getElementById('nj-mockup').addEventListener('change', function() {
  const label = document.getElementById('nj-mockup-label');
  const nameEl = document.getElementById('nj-mockup-name');
  if (this.files && this.files[0]) {
    nameEl.textContent = this.files[0].name;
    label.classList.add('has-file');
  } else {
    nameEl.textContent = 'Click to upload or drag & drop';
    label.classList.remove('has-file');
  }
});

// Need Mockup + To Sleeve segmented toggles
document.getElementById('nj-needmockup').addEventListener('click', function(e) {
  const opt = e.target.closest('.sleeve-opt');
  if (opt) this.dataset.value = opt.dataset.opt;
});
document.getElementById('nj-tosleeve').addEventListener('click', function(e) {
  const opt = e.target.closest('.sleeve-opt');
  if (opt) this.dataset.value = opt.dataset.opt;
});
document.getElementById('ej-tosleeve').addEventListener('click', function(e) {
  const opt = e.target.closest('.sleeve-opt');
  if (!opt) return;
  this.dataset.value = opt.dataset.opt;
  this.querySelectorAll('.sleeve-opt').forEach(b => b.classList.toggle('active', b === opt));
});

document.getElementById('nj-submit').addEventListener('click', async function() {
  const soort     = document.getElementById('nj-soort').value;
  const company   = document.getElementById('nj-company').value.trim();
  const printName = document.getElementById('nj-print-name').value.trim();
  const quantity  = document.getElementById('nj-quantity').value;
  const color     = document.getElementById('nj-color').value;
  const lid       = document.getElementById('nj-lid').value;
  const deadline  = document.getElementById('nj-deadline').value;
  const owner     = document.getElementById('nj-owner').value;
  const needmockup = document.getElementById('nj-needmockup').dataset.value;
  const tosleeve   = document.getElementById('nj-tosleeve').dataset.value;
  const notes     = withDate(document.getElementById('nj-notes').value.trim());
  const shipContact = document.getElementById('nj-ship-contact').value.trim();
  const shipPhone   = document.getElementById('nj-ship-phone').value.trim();
  const shipEmail   = document.getElementById('nj-ship-email').value.trim();
  const shipStreet  = document.getElementById('nj-ship-street').value.trim();
  const shipNumber  = document.getElementById('nj-ship-number').value.trim();
  const shipZipcode = document.getElementById('nj-ship-zipcode').value.trim();
  const shipCity    = document.getElementById('nj-ship-city').value.trim();
  const shipCountry = document.getElementById('nj-ship-country').value.trim().toUpperCase() || '';
  const mockupFile = document.getElementById('nj-mockup').files[0];
  const designFileInputs = document.getElementById('nj-files-list').querySelectorAll('input[type="file"]');
  const statusEl  = document.getElementById('nj-status');

  if (!soort || !company || !printName || !quantity) {
    statusEl.className = 'form-status error';
    statusEl.textContent = 'Please fill in all required fields.';
    return;
  }

  this.disabled = true;
  statusEl.className = 'form-status';
  statusEl.textContent = '';

  // Progress helpers (null-safe in case HTML hasn't refreshed yet)
  const progressWrap  = document.getElementById('nj-upload-progress');
  const progressFill  = document.getElementById('nj-upload-fill');
  const progressLabel = document.getElementById('nj-upload-label');
  const setProgress = (pct, label) => {
    if (progressWrap)  progressWrap.style.display  = 'block';
    if (progressFill)  progressFill.style.width    = Math.round(pct * 100) + '%';
    if (progressLabel) progressLabel.textContent   = label;
  };
  const hideProgress = () => {
    if (progressWrap) progressWrap.style.display = 'none';
    if (progressFill) progressFill.style.width   = '0%';
  };

  // Collect files to read
  const filesToRead = [];
  if (mockupFile) filesToRead.push({ file: mockupFile, role: 'mockup' });
  for (const inp of designFileInputs) { if (inp.files && inp.files[0]) filesToRead.push({ file: inp.files[0], role: 'design' }); }

  // Read all files with progress
  let mockupBase64 = null;
  const designFiles = [];
  for (let i = 0; i < filesToRead.length; i++) {
    const { file, role } = filesToRead[i];
    const baseProgress = i / filesToRead.length;
    setProgress(baseProgress * 0.4, `Reading file ${i + 1} of ${filesToRead.length}…`);
    try {
      const f = await readFileAsBase64(file, p => setProgress((baseProgress + p / filesToRead.length) * 0.4, `Reading file ${i + 1} of ${filesToRead.length}… ${Math.round(p * 100)}%`));
      if (role === 'mockup') mockupBase64 = f.data;
      else designFiles.push({ base64: f.data, mime: f.mime, name: f.name });
    } catch (_) {}
  }

  if (filesToRead.length === 0) setProgress(0, 'Uploading…');
  else setProgress(0.4, 'Uploading…');

  try {
    await postWithProgress(
      SCRIPT_URL,
      JSON.stringify({
        action:    'add_job',
        soort, company, printName,
        quantity:  parseInt(quantity),
        color, lid, deadline, owner, tosleeve, needmockup, notes,
        mockupBase64, designFiles,
        shipContact, shipPhone, shipEmail,
        shipStreet, shipNumber, shipZipcode, shipCity, shipCountry,
        changedBy: currentUser?.email,
        status:    'To Print',
      }),
      p => setProgress(0.4 + p * 0.55, `Uploading… ${Math.round(p * 100)}%`)
    );
    setProgress(1, 'Done!');

    sleeveLoaded = false;
    if (needmockup === 'Yes') mockupLoaded = false;

    statusEl.className = 'form-status success';
    const extras = [tosleeve === 'Yes' ? 'sleeve job' : '', needmockup === 'Yes' ? 'mockup job' : ''].filter(Boolean);
    statusEl.textContent = extras.length ? `✓ Print job added + ${extras.join(' & ')} created!` : '✓ Print job added!';
    setTimeout(() => hideProgress(), 1500);
    document.getElementById('add-job-form').reset();
    document.getElementById('nj-priority-hint').textContent = '';
    document.getElementById('nj-mockup-label').classList.remove('has-file');
    document.getElementById('nj-mockup-name').textContent = 'Click to upload image';
    document.getElementById('nj-files-list').innerHTML = '';
    addNjFileRow();
    document.getElementById('nj-needmockup').dataset.value = 'No';
    document.getElementById('nj-tosleeve').dataset.value = 'No';
    setTimeout(() => { statusEl.textContent = ''; }, 4000);
    refreshData();
  } catch (err) {
    hideProgress();
    statusEl.className = 'form-status error';
    statusEl.textContent = 'Error: ' + err.message;
  }
  this.disabled = false;
});

// ── Form completion progress bars ────────────────────────────
function setupFormProgress(fieldIds, fillId, labelId, total) {
  const fill  = document.getElementById(fillId);
  const label = document.getElementById(labelId);
  if (!fill || !label) return;
  function update() {
    const done = fieldIds.filter(id => {
      const el = document.getElementById(id);
      return el && el.value && el.value.trim() !== '';
    }).length;
    const pct = Math.round((done / total) * 100);
    fill.style.width = pct + '%';
    fill.classList.toggle('complete', done === total);
    label.textContent = done + ' / ' + total + ' required fields';
  }
  fieldIds.forEach(id => {
    const el = document.getElementById(id);
    if (el) el.addEventListener('input', update);
    if (el) el.addEventListener('change', update);
  });
  update();
}

setupFormProgress(['nj-soort','nj-quantity','nj-company','nj-print-name'], 'nj-progress-fill', 'nj-progress-label', 4);
setupFormProgress(['mk-soort','mk-company'], 'mk-progress-fill', 'mk-progress-label', 2);
setupFormProgress(['sv-soort','sv-quantity','sv-company'], 'sv-progress-fill', 'sv-progress-label', 3);

// ── Boot: check auth before loading data ─────────────────────
const _stored = localStorage.getItem('izy_user');
if (_stored) {
  try {
    currentUser = JSON.parse(_stored);
    showApp();
  } catch (_) {
    localStorage.removeItem('izy_user');
  }
}
