const SHEET_ID    = '1vIERVGUheXWkMS155VWfBEuCrUV4qXGYSUM9mIdppfc';
const CSV_URL     = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=csv`;
const SHIP_URL    = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=csv&gid=1459899540`;

// ── Print Logger config ────────────────────────────────────────
// Paste your Google Apps Script Web App URL here after deploying:
const SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbz1LuTt6ySUIXR_Rp3f8EyE4nV1XNW3hmK8YqirpxN_6HwzXFvvtyuVl_7Pj8_eNICs/exec';
// List your printers here:
const PRINTERS = ['Bottle 1', 'Bottle 2', 'Mug 1', 'Travel Bottle 1'];

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
  if (s === 'waiting')          return `<span class="badge b-waiting">Waiting</span>`;
  if (s.includes('progress'))   return `<span class="badge b-progress">In Progress</span>`;
  if (s === 'ready to ship')    return `<span class="badge b-ready-ship">Ready to Ship</span>`;
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
document.querySelectorAll('.tab-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(s => s.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById('tab-' + btn.dataset.tab).classList.add('active');
    if (btn.dataset.tab === 'shipping') loadShipping();
  });
});

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
    if (sleeve && get(r,'To sleeve?') !== sleeve) return false;
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
          <td>${get(r,'To sleeve?') || '—'}</td>
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

  const html = AQ_SECTIONS.map(section => {
    const rows = filtered
      .filter(r => section.match(get(r,'Soort').toLowerCase()))
      .sort((a,b) => statusSortOrder(a) - statusSortOrder(b));

    if (rows.length === 0) return '';

    const c = section.colors;
    return `
      <div class="aq-section">
        <div class="aq-section-title">
          <span style="background:${c.bg};color:${c.text};border-radius:6px;padding:4px 14px;font-size:13px;font-weight:700;">${section.label}</span>
          <span class="aq-section-count">${rows.length} job${rows.length !== 1 ? 's' : ''}</span>
        </div>
        <div class="table-wrap">
          <table>
            <thead><tr>
              <th>#</th><th>Company</th><th>Print Name</th><th>Status</th>
              <th>Type</th><th>Deadline</th><th>Color</th><th>Lid Color</th>
              <th>Qty</th><th>Still to Print</th><th>Days Left</th><th></th>
            </tr></thead>
            <tbody>${rows.map(r => {
              const d     = parseDate(get(r,'Deadline'));
              const days  = daysFrom(d);
              const still = num(r,'Quantity still to print');
              return `<tr class="${isOverdue(r) ? 'row-overdue' : ''}">
                <td>${get(r,'Priority')}</td>
                <td><strong>${get(r,'Name_Company')}</strong></td>
                <td class="print-name">${get(r,'Name_Print') || '—'}</td>
                <td>${badge(get(r,'Status'))}</td>
                <td>${typeBadge(get(r,'Soort'))}</td>
                <td>${get(r,'Deadline') || '—'}</td>
                <td>${get(r,'Bottle color') || '—'}</td>
                <td>${r['_col_Q'] || '—'}</td>
                <td>${num(r,'Quantity') || '—'}</td>
                <td class="${still > 0 ? 'cell-danger' : ''}">${still > 0 ? still : '—'}</td>
                <td>${daysCell(days)}</td>
                <td>
                  <button class="btn-log" data-rowidx="${allRows.indexOf(r)}">✏️ Log</button>
                  ${(() => { const st = get(r,'Status').toLowerCase(); if (st === 'waiting') return `<button class="btn-sleeve" data-rowidx="${allRows.indexOf(r)}">✕ Sleeve</button>`; if (st === 'ready to ship') return `<button class="btn-sleeve sleeved" data-rowidx="${allRows.indexOf(r)}" disabled>✓ Sleeved</button>`; return ''; })()}
                </td>
              </tr>`;
            }).join('')}</tbody>
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

// Log + Sleeve buttons — event delegation on section container
document.getElementById('tab-active-queue').addEventListener('click', function(e) {
  const logBtn    = e.target.closest('.btn-log');
  const sleeveBtn = e.target.closest('.btn-sleeve');
  if (logBtn)    openPrintModal(parseInt(logBtn.dataset.rowidx));
  if (sleeveBtn && !sleeveBtn.classList.contains('sleeved')) markSleeved(parseInt(sleeveBtn.dataset.rowidx), sleeveBtn);
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
  const sleeve = base.filter(r => get(r,'To sleeve?') === 'Yes' && get(r,'Gesleeved?') !== 'Yes' && isActive(r));
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
    const sleeveOk   = get(r,'To sleeve?') !== 'Yes' || get(r,'Gesleeved?') === 'Yes';
    return printReady && sleeveOk && isActive(r);
  });
  document.getElementById('rep-ready-ship').innerHTML = readyToShip.length === 0
    ? '<tr><td colspan="7">Nothing ready for shipment.</td></tr>'
    : readyToShip.map(r => `<tr><td>${get(r,'Priority')}</td><td>${get(r,'Name_Company')}</td>
        <td class="print-name">${get(r,'Name_Print') || '—'}</td><td>${get(r,'Owner')}</td>
        <td>${typeBadge(get(r,'Soort'))}</td>
        <td>${num(r,'Quantity')}</td><td>${get(r,'To sleeve?') === 'Yes' ? '✅' : '—'}</td></tr>`).join('');

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
          const jobShipped = parseDate(get(job,'Shipped'));
          if (!jobShipped) return;
          const daysDiff = Math.abs((shipDate - jobShipped) / 86400000);
          if (daysDiff <= maxDays) candidates.push({ job, normJob, daysDiff });
        });
      }
    }
    candidates.sort((a, b) => a.daysDiff - b.daysDiff || b.normJob.length - a.normJob.length);
    return candidates;
  }

  function buildRow(s, job, matchMethod, daysDiff, shipDate) {
    const addedDate  = parseDate(get(job,'Date added ') || get(job,'Date added'));
    const turnaround = addedDate ? Math.round((shipDate - addedDate) / 86400000) : null;
    const priceRaw   = get(s,'Prijs').replace(',','.').replace(/[^0-9.]/g,'');
    return {
      ordernummer:  get(s,'Ordernummer'),
      priority:     get(job,'Priority'),
      company:      get(job,'Name_Company'),
      recipient:    get(s,'Ontvanger bedrijfsnaam'),
      printName:    get(job,'Name_Print'),
      owner:        get(job,'Owner'),
      dateAdded:    addedDate,
      shipDate,
      turnaround,
      type:         get(job,'Soort'),
      color:        get(job,'Bottle color'),
      quantity:     num(job,'Quantity'),
      faulty:       num(job,'Faulty prints'),
      carrier:      get(s,'Vervoerder'),
      destination:  [get(s,'Ontvanger plaats'), get(s,'Ontvanger land')].filter(Boolean).join(', '),
      country:      get(s,'Ontvanger land'),
      price:        parseFloat(get(s,'Prijs').replace(',','.').replace(/[^0-9.]/g,'')) || null,
      tracking:     get(s,'AWB'),
      status:       get(s,'Bericht'),
      matchMethod,
      confidence:   calcConfidence(matchMethod, daysDiff),
    };
  }

  const matched        = [];
  const matchedJobKeys = new Set(); // Priority keys of jobs that got a shipment match

  shipRows.forEach(s => {
    const shipDate = parseDateISO(get(s,'Datum'));
    if (!shipDate) return;

    // 1. Referentie = Priority (100% reliable)
    const ref = get(s,'Referentie').trim();
    if (ref && /^\d+$/.test(ref) && jobByPriority[ref]) {
      matched.push(buildRow(s, jobByPriority[ref], 'reference', 0, shipDate));
      matchedJobKeys.add(ref);
      return;
    }

    // 2. Name match within 30 days → confirmed
    const close = findCandidates(get(s,'Ontvanger bedrijfsnaam'), shipDate, 30);
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

  return {
    matched: matched.sort((a,b) => b.shipDate - a.shipDate),
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
  const search   = document.getElementById('sh-search').value.toLowerCase();
  const owner    = document.getElementById('sh-owner').value;
  const period   = document.getElementById('sh-period').value;
  const dateFrom = document.getElementById('sh-date-from').value;
  const dateTo   = document.getElementById('sh-date-to').value;

  // Filter
  const rows = shippedRows.filter(r => {
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
        <td>${r.carrier ? r.carrier.split(' ')[0] : '—'}</td>
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
  if (period === 'year')    cutoff.setFullYear(now.getFullYear() - 1);
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
    const shipRaw    = await fetch(SHIP_URL + '&t=' + Date.now()).then(r => r.text());
    const shipParsed = parseCSV(shipRaw);
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
    const fetchPromise = fetch(CSV_URL + '&t=' + Date.now()).then(r => r.text());
    const timeoutPromise = new Promise((_, reject) =>
      setTimeout(() => reject(new Error('timeout')), 35000)
    );

    const mainRaw = await Promise.race([fetchPromise, timeoutPromise]);
    const parsed  = parseCSV(mainRaw).filter(r => get(r,'Name_Company') && get(r,'Priority') && get(r,'Priority') !== '0');

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

  // Pre-fill current values
  document.getElementById('modal-printed').value = num(modalJob,'Quantity printed ') || num(modalJob,'Quantity printed') || '';
  document.getElementById('modal-faulty').value  = num(modalJob,'Faulty prints') || 0;

  // Populate printers
  const sel = document.getElementById('modal-printer');
  sel.innerHTML = '<option value="">Select printer...</option>' +
    PRINTERS.map(p => `<option value="${p}">${p}</option>`).join('');
  const cur = get(modalJob,'Printer');
  if (cur) sel.value = cur;

  // Reset photo
  document.getElementById('modal-photo').value = '';
  document.getElementById('modal-preview').style.display = 'none';
  document.getElementById('modal-status').textContent = '';

  document.getElementById('print-modal-overlay').style.display = 'flex';
  document.getElementById('modal-printed').focus();
}

function closeModal() {
  document.getElementById('print-modal-overlay').style.display = 'none';
  modalJob = null;
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
      body:   JSON.stringify({ sheetRow: job['_sheetRow'], status: 'Ready to Ship' }),
    });
    refreshData();
  } catch (err) {
    // Revert button on failure
    if (btn) { btn.textContent = '✕ Sleeve'; btn.classList.remove('sleeved'); btn.disabled = false; }
    alert('Could not update: ' + err.message);
  }
}

async function submitPrintUpdate() {
  const statusEl = document.getElementById('modal-status');
  const submitBtn = document.getElementById('modal-submit');
  const printed = parseInt(document.getElementById('modal-printed').value);
  const faulty  = parseInt(document.getElementById('modal-faulty').value)  || 0;
  const printer = document.getElementById('modal-printer').value;
  const photoFile = document.getElementById('modal-photo').files[0];

  if (!printed && printed !== 0) { statusEl.textContent = 'Please enter quantity printed.'; statusEl.className = 'modal-error'; return; }

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
      imageBase64,
      imageMime,
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

refreshData();
