const SHEET_ID    = '1vIERVGUheXWkMS155VWfBEuCrUV4qXGYSUM9mIdppfc';
const CSV_URL     = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=csv`;
const SHIP_URL    = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=csv&gid=1459899540`;

// ── Print Logger config ────────────────────────────────────────
// Paste your Google Apps Script Web App URL here after deploying:
const SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbz1LuTt6ySUIXR_Rp3f8EyE4nV1XNW3hmK8YqirpxN_6HwzXFvvtyuVl_7Pj8_eNICs/exec';
// List your printers here:
const PRINTERS = ['Bottle 1', 'Bottle 2', 'Mug 1', 'Travel Bottle 1'];

// ── Auth ──────────────────────────────────────────────────────
const ALLOWED_EMAILS = [
  'daan@izybottles.com',
  'jim@izybottles.com',
  'biessenlevi@gmail.com',
  'sharon@orderchamp.com',
  'ivan@izybottles.com',
];

// Emails that only see the Active Queue tab
const ACTIVE_QUEUE_ONLY = ['ivan@izybottles.com'];

let currentUser  = null;
let sleeveRows   = [];
let sleeveLoaded = false;

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
let shippedRows  = []; // confirmed matched shipping rows
let reviewRows   = []; // unmatched / low-confidence rows for review
let charts       = {};

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
  if (s === 'to sleeve')        return `<span class="badge b-to-print">To Sleeve</span>`;
  if (s === 'waiting')          return `<span class="badge b-waiting">Waiting</span>`;
  if (s.includes('progress'))   return `<span class="badge b-progress">In Progress</span>`;
  if (s === 'ready to ship')    return `<span class="badge b-ready-ship">Ready to Ship</span>`;
  if (s === 'done')             return `<span class="badge b-shipped">Done</span>`;
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
  if (tabName === 'add-job') populateAddJobOwners();
  if (tabName === 'sleeves') { populateSleeveOwners(); loadSleeves(); }
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
          <td>${get(r,'Priority')}</td>
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

document.getElementById('overview-period').addEventListener('change', () => {
  const overviewRows = getOverviewRows();
  renderStats(overviewRows);
  renderCharts(overviewRows);
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

  // Type filter pills
  const tabsEl = document.getElementById('aq-type-tabs');
  tabsEl.innerHTML = `<button class="aq-type-tab${aqTypeFilter === '' ? ' active' : ''}" data-type="">All <span class="aq-tab-count">${filtered.length}</span></button>` +
    AQ_SECTIONS.map(s => {
      const n = filtered.filter(r => s.match(get(r,'Soort').toLowerCase())).length;
      if (!n) return '';
      return `<button class="aq-type-tab${aqTypeFilter === s.label ? ' active' : ''}" data-type="${s.label}" style="--tc:${s.colors.text};--tb:${s.colors.bg}">${s.label} <span class="aq-tab-count">${n}</span></button>`;
    }).join('');

  const html = AQ_SECTIONS.filter(s => !aqTypeFilter || s.label === aqTypeFilter).map(section => {
    const rows = filtered
      .filter(r => section.match(get(r,'Soort').toLowerCase()))
      .sort((a,b) => statusSortOrder(a) - statusSortOrder(b));

    if (rows.length === 0) return '';

    const c = section.colors;
    const rowsHtml = rows.map(r => {
      const d     = parseDate(get(r,'Deadline'));
      const days  = daysFrom(d);
      const still = num(r,'Quantity still to print');
      const idx   = allRows.indexOf(r);
      const sleeveVal = (get(r,'To sleeve?') || getCI(r,'sleeve')).toLowerCase();
      const sleeveBtn = sleeveVal !== 'yes' ? '' :
        get(r,'Status').toLowerCase() === 'ready to ship'
          ? `<button class="btn-sleeve sleeved" data-rowidx="${idx}">✓ Sleeved</button>`
          : `<button class="btn-sleeve" data-rowidx="${idx}">✕ Sleeve</button>`;
      const actionBtns = `<button class="btn-log" data-rowidx="${idx}">✏️ Log</button>${sleeveBtn}<button class="btn-reset" data-rowidx="${idx}">↺ Reset</button>`;

      const card = `<div class="aq-card${isOverdue(r) ? ' overdue' : ''}" style="--tc:${c.text};--tb:${c.bg}">
        <div class="aq-card-top">
          <div class="aq-card-left">
            <span class="aq-prio">#${get(r,'Priority')}</span>
            <span class="aq-company">${get(r,'Name_Company')}</span>
          </div>
          ${badge(get(r,'Status'))}
        </div>
        ${get(r,'Name_Print') ? `<div class="aq-print-name">${get(r,'Name_Print')}</div>` : ''}
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

      const row = `<tr class="${isOverdue(r) ? 'row-overdue' : ''}">
        <td>${get(r,'Priority')}</td>
        <td><strong>${get(r,'Name_Company')}</strong></td>
        <td class="print-name">${get(r,'Name_Print') || '—'}</td>
        <td>${badge(get(r,'Status'))}</td>
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

    return `
      <div class="aq-section">
        <div class="aq-section-title">
          <span style="background:${c.bg};color:${c.text};border-radius:6px;padding:4px 14px;font-size:13px;font-weight:700;">${section.label}</span>
          <span class="aq-section-count">${rows.length} job${rows.length !== 1 ? 's' : ''}</span>
        </div>
        <div class="aq-cards">${rowsHtml.map(x => x.card).join('')}</div>
        <div class="aq-table-wrap table-wrap">
          <table>
            <thead><tr>
              <th>#</th><th>Company</th><th>Print Name</th><th>Status</th>
              <th>Type</th><th>Deadline</th><th>Color</th><th>Lid</th>
              <th>Qty</th><th>Still to Print</th><th>Days Left</th><th></th>
            </tr></thead>
            <tbody>${rowsHtml.map(x => x.row).join('')}</tbody>
          </table>
        </div>
      </div>`;
  }).join('');

  document.getElementById('aq-sections').innerHTML = html ||
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
  const logBtn    = e.target.closest('.btn-log');
  const sleeveBtn = e.target.closest('.btn-sleeve');
  const resetBtn  = e.target.closest('.btn-reset');
  if (logBtn)   openPrintModal(parseInt(logBtn.dataset.rowidx));
  if (resetBtn) resetJob(parseInt(resetBtn.dataset.rowidx));
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
        return `<tr><td>${get(r,'Priority')}</td><td>${get(r,'Name_Company')}</td>
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
        return `<tr><td>${get(r,'Priority')}</td><td>${get(r,'Name_Company')}</td>
          <td class="print-name">${get(r,'Name_Print') || '—'}</td><td>${get(r,'Owner')}</td>
          <td>${q}</td><td class="cell-danger">${f}</td><td class="cell-warn">${pct}</td></tr>`;
      }).join('');

  // Needs sleeving
  const sleeve = base.filter(r => getCI(r,'sleeve').toLowerCase() === 'yes' && get(r,'Gesleeved?') !== 'Yes' && isActive(r));
  document.getElementById('rep-sleeve').innerHTML = sleeve.length === 0
    ? '<tr><td colspan="6" class="cell-ok">All sleeveable jobs are done!</td></tr>'
    : sleeve.map(r => `<tr><td>${get(r,'Priority')}</td><td>${get(r,'Name_Company')}</td>
        <td class="print-name">${get(r,'Name_Print') || '—'}</td><td>${badge(get(r,'Status'))}</td>
        <td>${get(r,'Owner')}</td><td>${num(r,'Quantity')}</td></tr>`).join('');

  // To Print
  const toPrint = base.filter(r => get(r,'Status').toLowerCase() === 'to print')
    .sort((a,b) => num(a,'Priority') - num(b,'Priority'));
  document.getElementById('rep-to-print').innerHTML = toPrint.length === 0
    ? '<tr><td colspan="8">No jobs queued to print.</td></tr>'
    : toPrint.map(r => `<tr><td>${get(r,'Priority')}</td><td>${get(r,'Name_Company')}</td>
        <td class="print-name">${get(r,'Name_Print') || '—'}</td><td>${get(r,'Owner')}</td>
        <td>${get(r,'Deadline') || '—'}</td><td>${typeBadge(get(r,'Soort'))}</td>
        <td>${get(r,'Bottle color') || '—'}</td><td>${num(r,'Quantity')}</td></tr>`).join('');

  // Waiting
  const waiting = base.filter(r => get(r,'Status').toLowerCase() === 'waiting')
    .sort((a,b) => num(a,'Priority') - num(b,'Priority'));
  document.getElementById('rep-waiting').innerHTML = waiting.length === 0
    ? '<tr><td colspan="6">No waiting jobs.</td></tr>'
    : waiting.map(r => `<tr><td>${get(r,'Priority')}</td><td>${get(r,'Name_Company')}</td>
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
    : readyToShip.map(r => `<tr><td>${get(r,'Priority')}</td><td>${get(r,'Name_Company')}</td>
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
        return `<tr><td>${get(r,'Priority')}</td><td>${get(r,'Name_Company')}</td>
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
    : recentShipped.map(r => `<tr><td>${get(r,'Priority')}</td><td>${get(r,'Name_Company')}</td>
        <td class="print-name">${get(r,'Name_Print') || '—'}</td><td>${get(r,'Owner')}</td>
        <td>${get(r,'Shipped') || '—'}</td><td>${num(r,'Quantity')}</td></tr>`).join('');
}

// ── Populate filter dropdowns ─────────────────────────────────
function populateSecondaryFilters(rows) {
  const active   = rows.filter(r => isActive(r));
  const statuses = [...new Set(active.map(r => get(r,'Status')).filter(Boolean))].sort();
  const owners   = [...new Set(rows.map(r => get(r,'Owner')).filter(Boolean))].sort();
  fill('aq-status', statuses);
  fill('rp-owner',  owners);
}

['rp-owner','rp-period','rp-date-from','rp-date-to'].forEach(id => {
  document.getElementById(id).addEventListener('input', renderReports);
  document.getElementById(id).addEventListener('change', renderReports);
});
document.getElementById('rp-period').addEventListener('change', function() {
  document.getElementById('rp-custom-range').style.display = this.value === 'custom' ? 'inline-flex' : 'none';
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
  const row10 = r => `<tr><td>${r.priority}</td><td>${r.company}</td><td>${r.owner}</td><td>${turnaroundColor(r.turnaround)}</td><td>${fmtDate(r.shipDate)}</td></tr>`;
  document.getElementById('sh-slowest-list').innerHTML = sorted.slice(0,10).map(row10).join('');
  document.getElementById('sh-fastest-list').innerHTML = [...sorted].reverse().slice(0,10).map(row10).join('');

  // Main table
  document.getElementById('sh-body').innerHTML = rows.length === 0
    ? '<tr><td colspan="14">No matched shipments found.</td></tr>'
    : rows.map(r => `<tr>
        <td>${r.priority || '—'}</td>
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
        <td>${r.priority || '—'}</td>
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

function sleeveSortOrder(r) {
  const s = get(r,'Status').toLowerCase();
  if (s.includes('progress')) return 0;
  if (s === 'to sleeve')      return 1;
  if (s === 'done')           return 2;
  return 3;
}

function populateSleeveOwners() {
  const owners = [...new Set(allRows.map(r => get(r,'Owner')).filter(Boolean))].sort();
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
  const search   = document.getElementById('sv-search').value.toLowerCase();
  const status   = document.getElementById('sv-status').value;
  const hideDone = document.getElementById('sv-hide-done').checked;

  const filtered = sleeveRows.filter(r => {
    if (hideDone && get(r,'Status').toLowerCase() === 'done') return false;
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
      const idx            = sleeveRows.indexOf(r);
      const alreadySleeved = parseInt(getCI(r,'sleeved')) || 0;
      const qty            = num(r,'Quantity');
      const stillToSleeve  = Math.max(0, qty - alreadySleeved);
      const d              = parseDate(get(r,'Deadline'));
      const days           = daysFrom(d);
      const isDone         = get(r,'Status').toLowerCase() === 'done';

      const actionBtns = isDone
        ? `<button class="btn-reset sv-btn-reset" data-svidx="${idx}">↺ Reset</button>`
        : `<button class="btn-log sv-btn-log" data-svidx="${idx}">✏️ Log</button><button class="btn-sleeve sv-btn-done" data-svidx="${idx}">✓ Done</button><button class="btn-reset sv-btn-reset" data-svidx="${idx}">↺ Reset</button>`;

      const card = `<div class="aq-card${isDone ? ' sv-card-done' : ''}" style="--tc:${c.text};--tb:${c.bg}">
        <div class="aq-card-top">
          <div class="aq-card-left">
            <span class="aq-prio">#${get(r,'Priority')}</span>
            <span class="aq-company">${get(r,'Name_Company')}</span>
          </div>
          ${badge(get(r,'Status'))}
        </div>
        ${get(r,'Name_Print') ? `<div class="aq-print-name">${get(r,'Name_Print')}</div>` : ''}
        <div class="aq-meta">
          <div class="aq-meta-item"><span class="aq-meta-label">Total Qty</span><span>${qty || '—'}</span></div>
          <div class="aq-meta-item"><span class="aq-meta-label">Sleeved</span><span>${alreadySleeved}</span></div>
          ${stillToSleeve > 0 ? `<div class="aq-meta-item"><span class="aq-meta-label">Still to sleeve</span><span class="cell-danger">${stillToSleeve}</span></div>` : ''}
          ${get(r,'Deadline') ? `<div class="aq-meta-item"><span class="aq-meta-label">Deadline</span><span>${get(r,'Deadline')}</span></div>` : ''}
          ${days !== null ? `<div class="aq-meta-item"><span class="aq-meta-label">Days left</span>${daysCell(days)}</div>` : ''}
        </div>
        <div class="aq-card-actions">${actionBtns}</div>
      </div>`;

      const row = `<tr class="${isDone ? 'row-shipped' : ''}">
        <td>${get(r,'Priority')}</td>
        <td><strong>${get(r,'Name_Company')}</strong></td>
        <td class="print-name">${get(r,'Name_Print') || '—'}</td>
        <td>${badge(get(r,'Status'))}</td>
        <td>${typeBadge(get(r,'Soort'))}</td>
        <td>${get(r,'Deadline') || '—'}</td>
        <td>${qty || '—'}</td>
        <td>${alreadySleeved || 0}</td>
        <td class="${stillToSleeve > 0 ? 'cell-danger' : ''}">${stillToSleeve > 0 ? stillToSleeve : '—'}</td>
        <td>${daysCell(days)}</td>
        <td style="white-space:nowrap">${actionBtns}</td>
      </tr>`;

      return { card, row };
    });

    return `
      <div class="aq-section">
        <div class="aq-section-title">
          <span style="background:${c.bg};color:${c.text};border-radius:6px;padding:4px 14px;font-size:13px;font-weight:700;">${section.label}</span>
          <span class="aq-section-count">${rows.length} job${rows.length !== 1 ? 's' : ''}</span>
        </div>
        <div class="aq-cards">${rowsHtml.map(x => x.card).join('')}</div>
        <div class="aq-table-wrap table-wrap">
          <table>
            <thead><tr>
              <th>#</th><th>Company</th><th>Product</th><th>Status</th>
              <th>Type</th><th>Deadline</th><th>Qty</th><th>Sleeved</th>
              <th>Still to Sleeve</th><th>Days Left</th><th></th>
            </tr></thead>
            <tbody>${rowsHtml.map(x => x.row).join('')}</tbody>
          </table>
        </div>
      </div>`;
  }).join('');

  document.getElementById('sv-sections').innerHTML = html ||
    '<p style="color:#94a3b8;padding:20px;">No sleeve jobs match the current filters.</p>';
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
  const logBtn  = e.target.closest('.sv-btn-log');
  const doneBtn = e.target.closest('.sv-btn-done');
  const resetBtn = e.target.closest('.sv-btn-reset');
  if (logBtn)   openSleeveModal(parseInt(logBtn.dataset.svidx));
  if (doneBtn)  markSleeveDone(parseInt(doneBtn.dataset.svidx));
  if (resetBtn) resetSleeveJob(parseInt(resetBtn.dataset.svidx));
});

// ── Sleeve Modal ──────────────────────────────────────────────

let sleeveModalJob = null;

function openSleeveModal(rowIdx) {
  sleeveModalJob = sleeveRows[rowIdx];
  if (!sleeveModalJob) return;

  const alreadySleeved = parseInt(getCI(sleeveModalJob,'sleeved')) || 0;
  const qty = num(sleeveModalJob,'Quantity');

  document.getElementById('sv-modal-job-info').innerHTML = `
    <div class="modal-job-card">
      <div><span class="modal-label">Job</span><strong>#${get(sleeveModalJob,'Priority')} — ${get(sleeveModalJob,'Name_Company')}</strong></div>
      <div><span class="modal-label">Product</span>${get(sleeveModalJob,'Name_Print') || '—'}</div>
      <div><span class="modal-label">Type</span>${typeBadge(get(sleeveModalJob,'Soort'))}</div>
      <div><span class="modal-label">Total Qty</span>${qty}</div>
      <div><span class="modal-label">Already Sleeved</span>${alreadySleeved}</div>
      ${qty - alreadySleeved > 0 ? `<div><span class="modal-label">Still to Sleeve</span><span class="cell-danger">${qty - alreadySleeved}</span></div>` : ''}
    </div>`;

  document.getElementById('sv-modal-sleeved').value = '';
  document.getElementById('sv-modal-status').textContent = '';

  document.getElementById('sleeve-modal-overlay').style.display = 'flex';
  document.getElementById('sv-modal-sleeved').focus();
}

function closeSleeveModal() {
  document.getElementById('sleeve-modal-overlay').style.display = 'none';
  sleeveModalJob = null;
}

document.getElementById('sleeve-modal-overlay').addEventListener('click', function(e) {
  if (e.target === this) closeSleeveModal();
});

async function submitSleeveUpdate() {
  const statusEl  = document.getElementById('sv-modal-status');
  const submitBtn = document.getElementById('sv-modal-submit');
  const sessionSleeved = parseInt(document.getElementById('sv-modal-sleeved').value);

  if (isNaN(sessionSleeved) || sessionSleeved < 0) {
    statusEl.textContent = 'Please enter a valid quantity.';
    statusEl.className   = 'modal-error';
    return;
  }

  const alreadySleeved = parseInt(getCI(sleeveModalJob,'sleeved')) || 0;
  const totalSleeved   = alreadySleeved + sessionSleeved;
  const qty            = num(sleeveModalJob,'Quantity');
  const autoStatus     = totalSleeved >= qty ? 'Done' : 'In Progress';

  submitBtn.disabled    = true;
  submitBtn.textContent = 'Submitting…';
  statusEl.textContent  = '';

  try {
    await fetch(SCRIPT_URL, {
      method: 'POST',
      mode:   'no-cors',
      body:   JSON.stringify({
        action:           'update_sleeve',
        sheetRow:         sleeveModalJob['_sheetRow'],
        quantitySleeved:  totalSleeved,
        status:           autoStatus,
        changedBy:        currentUser?.email,
      }),
    });
    submitBtn.textContent = 'Submit';
    submitBtn.disabled    = false;
    statusEl.textContent  = '✅ Saved! The sheet will update within a few seconds.';
    statusEl.className    = 'modal-success';
    setTimeout(() => { closeSleeveModal(); sleeveLoaded = false; loadSleeves(); }, 2200);
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
  if (!confirm(`Reset "${get(job,'Name_Company')} #${get(job,'Priority')}" back to "To Sleeve"?`)) return;
  try {
    await fetch(SCRIPT_URL, {
      method: 'POST',
      mode:   'no-cors',
      body:   JSON.stringify({
        action:           'update_sleeve',
        sheetRow:         job['_sheetRow'],
        quantitySleeved:  0,
        status:           'To Sleeve',
        changedBy:        currentUser?.email,
      }),
    });
    sleeveLoaded = false;
    loadSleeves();
  } catch (err) {
    alert('Could not reset: ' + err.message);
  }
}

// ── Add Sleeve Job ────────────────────────────────────────────

function toggleAddSleeveForm() {
  const wrap = document.getElementById('add-sleeve-form-wrap');
  const isVisible = wrap.style.display !== 'none';
  wrap.style.display = isVisible ? 'none' : 'block';
  if (!isVisible) populateSleeveOwners();
}

document.getElementById('sv-submit').addEventListener('click', async function() {
  const soort     = document.getElementById('sv-soort').value;
  const company   = document.getElementById('sv-company').value.trim();
  const printName = document.getElementById('sv-print-name').value.trim();
  const quantity  = document.getElementById('sv-quantity').value;
  const deadline  = document.getElementById('sv-deadline').value;
  const owner     = document.getElementById('sv-owner').value;
  const notes     = document.getElementById('sv-notes').value.trim();
  const statusEl  = document.getElementById('sv-form-status');

  if (!soort || !company || !printName || !quantity) {
    statusEl.className   = 'form-status error';
    statusEl.textContent = 'Please fill in all required fields.';
    return;
  }

  this.disabled        = true;
  statusEl.className   = 'form-status';
  statusEl.textContent = 'Saving…';

  try {
    await fetch(SCRIPT_URL, {
      method: 'POST',
      mode:   'no-cors',
      body:   JSON.stringify({
        action:    'add_sleeve_job',
        soort, company, printName,
        quantity:  parseInt(quantity),
        deadline, owner, notes,
        changedBy: currentUser?.email,
      }),
    });
    statusEl.className   = 'form-status success';
    statusEl.textContent = '✓ Sleeve job added!';
    document.getElementById('add-sleeve-form').reset();
    setTimeout(() => { statusEl.textContent = ''; }, 4000);
    sleeveLoaded = false;
    loadSleeves();
  } catch (err) {
    statusEl.className   = 'form-status error';
    statusEl.textContent = 'Error: ' + err.message;
  }
  this.disabled = false;
});

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
  const now = new Date();
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

async function refreshData() {
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
    const fetchPromise = fetch(SCRIPT_URL + '?t=' + Date.now()).then(r => r.json()).then(d => d.rows || []);
    const timeoutPromise = new Promise((_, reject) =>
      setTimeout(() => reject(new Error('timeout')), 35000)
    );

    const parsed  = (await Promise.race([fetchPromise, timeoutPromise])).filter(r => get(r,'Name_Company') && get(r,'Priority') && get(r,'Priority') !== '0');

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
function populateAddJobOwners() {
  const owners = [...new Set(allRows.map(r => get(r,'Owner')).filter(Boolean))].sort();
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

// To Sleeve segmented toggle
document.getElementById('nj-tosleeve').addEventListener('click', function(e) {
  const opt = e.target.closest('.sleeve-opt');
  if (opt) this.dataset.value = opt.dataset.opt;
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
  const tosleeve  = document.getElementById('nj-tosleeve').dataset.value;
  const notes     = document.getElementById('nj-notes').value.trim();
  const mockupFile = document.getElementById('nj-mockup').files[0];
  const statusEl  = document.getElementById('nj-status');

  if (!soort || !company || !printName || !quantity) {
    statusEl.className = 'form-status error';
    statusEl.textContent = 'Please fill in all required fields.';
    return;
  }

  this.disabled = true;
  statusEl.className = 'form-status';
  statusEl.textContent = 'Saving…';

  // Read mockup as base64 if provided
  let mockupBase64 = null;
  if (mockupFile) {
    mockupBase64 = await new Promise(resolve => {
      const reader = new FileReader();
      reader.onload = e => resolve(e.target.result);
      reader.readAsDataURL(mockupFile);
    });
  }

  try {
    await fetch(SCRIPT_URL, {
      method: 'POST',
      mode:   'no-cors',
      body:   JSON.stringify({
        action:    'add_job',
        soort,
        company,
        printName,
        quantity:  parseInt(quantity),
        color,
        lid,
        deadline,
        owner,
        tosleeve,
        notes,
        mockupBase64,
        changedBy: currentUser?.email,
        status:    'To Print',
      }),
    });
    // If To Sleeve = Yes, also create a sleeve job automatically
    if (tosleeve === 'Yes') {
      await fetch(SCRIPT_URL, {
        method: 'POST',
        mode:   'no-cors',
        body:   JSON.stringify({
          action:    'add_sleeve_job',
          soort:     soort,
          company:   company,
          printName: printName,
          quantity:  parseInt(quantity),
          deadline:  deadline,
          owner:     owner,
          notes:     notes,
          changedBy: currentUser?.email,
        }),
      });
      sleeveLoaded = false; // force reload next time Sleeves tab is opened
    }

    statusEl.className = 'form-status success';
    statusEl.textContent = tosleeve === 'Yes' ? '✓ Print job added + sleeve job created!' : '✓ Print job added!';
    document.getElementById('add-job-form').reset();
    document.getElementById('nj-priority-hint').textContent = '';
    document.getElementById('nj-mockup-label').classList.remove('has-file');
    document.getElementById('nj-mockup-name').textContent = 'Click to upload or drag & drop';
    document.getElementById('nj-tosleeve').dataset.value = 'No';
    setTimeout(() => { statusEl.textContent = ''; }, 4000);
    refreshData();
  } catch (err) {
    statusEl.className = 'form-status error';
    statusEl.textContent = 'Error: ' + err.message;
  }
  this.disabled = false;
});

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
