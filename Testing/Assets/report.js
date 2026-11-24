var currentPage = 1;
var rowsPerPage = 100;
var currentSortKey = 'Name';
var currentSortDir = 'asc';
var searchTimer = null;
var rows = [];          // pre-computed row records (built once)
var visibleRows = [];   // rows after search + sort (recomputed only when those change)
var collator = new Intl.Collator();
var ownedRows = [];     // row records for the cards you own
var missingRows = [];   // row records for the cards you are missing
var showMissing = false;

// The raw data behind the current view (used by the Excel export)
function activeData() {
  var d = showMissing ? (typeof missingData !== 'undefined' ? missingData : []) : cardData;
  return Array.isArray(d) ? d : (d ? [d] : []);
}

function findColumn(key) {
  return columns.filter(function(c) { return c.key === key; })[0];
}

function escapeHtml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// Plain-text version of a value (HTML tags stripped) for searching and sorting
function plainValue(col, item) {
  var v = item[col.key];
  v = (v === null || v === undefined) ? '' : String(v);
  return col.type === 'html' ? v.replace(/<[^>]*>/g, '') : v;
}

function cellHtml(col, item) {
  var v = item[col.key];
  if (v === null || v === undefined || v === '') return '';
  if (col.delta) {
    var n = parseFloat(String(v).replace(/[^0-9.-]/g, ''));
    if (isNaN(n)) return '<span class="change change-new">' + escapeHtml(v) + '</span>';
    var cls = n > 0 ? 'up' : (n < 0 ? 'down' : 'flat');
    var sign = n > 0 ? '+' : (n < 0 ? '-' : '');
    return '<span class="change change-' + cls + '">' + sign + '$' + Math.abs(n).toFixed(2) + '</span>';
  }
  if (col.type === 'number' && isNaN(parseFloat(String(v).replace(/[^0-9.-]/g, '')))) {
    return '<span class="no-price">' + escapeHtml(v) + '</span>';
  }
  if (col.badge) {
    var slug = String(v).toLowerCase().replace(/[^a-z]/g, '');
    return '<span class="badge badge-' + slug + '">' + escapeHtml(v) + '</span>';
  }
  var text = col.type === 'html' ? v : escapeHtml(v);
  return (col.prefix || '') + text;
}

// Plain text, numbers and the search string are computed ONCE here, so
// searching and sorting thousands of cards stays instant.
function makeRows(data) {
  var searchCols = columns.filter(function(c) { return c.searchable; });
  if (data && !Array.isArray(data)) data = [data];
  return (data || []).filter(Boolean).map(function(item) {
    var r = { item: item, plain: {}, num: {}, search: '', noPrice: false };
    columns.forEach(function(col) {
      var p = plainValue(col, item);
      r.plain[col.key] = p;
      if (col.type === 'number' && p === 'No price') r.noPrice = true;
      if (col.type === 'number' || col.total) r.num[col.key] = parseFloat(p.replace(/[^0-9.-]/g, '')) || 0;
    });
    r.search = searchCols.map(function(c) { return r.plain[c.key]; }).join('\n').toLowerCase();
    return r;
  });
}

function buildHeader() {
  var row = document.getElementById('headerRow');
  row.innerHTML = '';
  columns.forEach(function(col) {
    var th = document.createElement('th');
    th.textContent = col.header;
    th.setAttribute('data-key', col.key);
    th.onclick = function() { onSort(th); };
    row.appendChild(th);
  });

  var names = columns.filter(function(c) { return c.searchable; })
                     .map(function(c) { return c.header; });
  document.getElementById('searchInput').placeholder = 'Search by ' + names.join(' or ') + '...';
}

// For every column flagged Total: the sum of ALL cards, plus (only while a search is
// active) the sum of the cards matching the search, across all pages.
// For every column flagged DeltaTotal: the gain/loss (change per copy x Qty), green or red.
function renderTotals() {
  var box = document.getElementById('totals');
  box.innerHTML = '';
  var filtered = visibleRows.length !== rows.length;

  function addPill(label, col, list) {
    var sum = 0;
    list.forEach(function(r) { sum += r.num[col.key] || 0; });
    var span = document.createElement('span');
    span.textContent = label + ' ' + col.header + ': ' + (col.prefix || '') +
      sum.toLocaleString('en-AU', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    box.appendChild(span);
  }

  // Profit/loss pill: sum of (change per copy x Qty), green for gain, red for loss
  function addChangePill(label, col, list) {
    var hasQty = !!findColumn('Qty');
    var sum = 0, any = false;
    list.forEach(function(r) {
      var n = parseFloat(String(r.plain[col.key]).replace(/[^0-9.-]/g, ''));
      if (isNaN(n)) return;                       // blank or 'New'
      sum += n * (hasQty ? (r.num['Qty'] || 0) : 1);
      any = true;
    });
    if (!any) return;                             // no baseline yet (first run)
    sum = Math.round(sum * 100) / 100;
    var span = document.createElement('span');
    span.className = sum > 0 ? 'gain' : (sum < 0 ? 'loss' : '');
    span.textContent = label + ' ' + col.header + ': ' + (sum > 0 ? '+' : (sum < 0 ? '-' : '')) + '$' +
      Math.abs(sum).toLocaleString('en-AU', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    box.appendChild(span);
  }

  columns.filter(function(c) { return c.total; }).forEach(function(col) {
    addPill(showMissing ? 'Missing' : 'Total', col, rows);
    if (filtered) addPill('Showing', col, visibleRows);
  });

  columns.filter(function(c) { return c.deltaTotal; }).forEach(function(col) {
    addChangePill('Total', col, rows);
    if (filtered) addChangePill('Showing', col, visibleRows);
  });

  var missing = rows.filter(function(r) { return r.noPrice; }).length;
  if (missing > 0) {
    var warn = document.createElement('span');
    warn.className = 'warn';
    warn.textContent = 'No price: ' + missing.toLocaleString() + (missing === 1 ? ' card' : ' cards') + ' (counted as $0.00)';
    box.appendChild(warn);
  }
}

function messageRow(text, color) {
  return '<tr><td colspan="' + columns.length + '" style="text-align:center; color:' + color + ';">' + text + '</td></tr>';
}

// Re-filters and re-sorts. Called only when the search text or sort changes.
function refreshRows() {
  var query = document.getElementById('searchInput').value.toLowerCase();
  visibleRows = query ? rows.filter(function(r) { return r.search.indexOf(query) !== -1; }) : rows.slice();

  if (currentSortKey) {
    var key = currentSortKey;
    var numeric = findColumn(key).type === 'number';
    var dir = currentSortDir === 'asc' ? 1 : -1;
    visibleRows.sort(function(a, b) {
      return numeric ? (a.num[key] - b.num[key]) * dir : collator.compare(a.plain[key], b.plain[key]) * dir;
    });
  }
}

// Paints only the current page (cheap, even with thousands of cards)
function renderTable() {
  hidePreview();
  var tbody = document.getElementById('tableBody');
  var info = document.getElementById('pageInfo');

  if (!Array.isArray(cardData)) {
    tbody.innerHTML = messageRow('Error: Data format corrupted.', 'red');
    return;
  }

  if (visibleRows.length === 0) {
    tbody.innerHTML = messageRow('No matching cards found.', '#777');
    info.textContent = '';
    renderControls(1);
    return;
  }

  var totalPages = Math.max(1, Math.ceil(visibleRows.length / rowsPerPage));
  if (currentPage > totalPages) currentPage = totalPages;

  var start = (currentPage - 1) * rowsPerPage;
  var pageRows = visibleRows.slice(start, start + rowsPerPage);
  tbody.innerHTML = pageRows.map(function(r) {
    var high = highValueKey && (r.num[highValueKey] || 0) > highValueMin;
    return '<tr' + (high ? ' class="high-value"' : '') + '>' + columns.map(function(col) { return '<td data-key="' + col.key + '">' + cellHtml(col, r.item) + '</td>'; }).join('') + '</tr>';
  }).join('');

  info.textContent = 'Showing ' + (start + 1).toLocaleString() + '-' + (start + pageRows.length).toLocaleString() +
    ' of ' + visibleRows.length.toLocaleString() + ' cards' +
    (visibleRows.length !== rows.length ? ' (filtered from ' + rows.length.toLocaleString() + ')' : '');

  renderControls(totalPages);
}

// ---------- Owned / missing cards tick box ----------
function initViewToggle() {
  var box = document.getElementById('missingToggle');
  var wrap = document.getElementById('viewToggle');
  var count = document.getElementById('missingCount');
  var status = (typeof missingStatus !== 'undefined') ? missingStatus : 'Missing card data was not included in this report.';
  if (status) {
    box.disabled = true;
    wrap.className = 'view-toggle disabled';
    wrap.title = status;
    count.textContent = '';
  } else {
    count.textContent = '(' + missingRows.length.toLocaleString() + ')';
    wrap.title = 'Swap the table between the cards you own and the cards you are missing from your sets';
  }
}

function onViewChange() {
  showMissing = document.getElementById('missingToggle').checked;
  rows = showMissing ? missingRows : ownedRows;
  document.getElementById('reportTable').className = showMissing ? 'missing-view' : '';
  var note = document.getElementById('viewNote');
  if (showMissing) {
    var skipped = (typeof missingSkipped !== 'undefined' && missingSkipped.length) ? missingSkipped : [];
    var updated = (typeof missingUpdated !== 'undefined' && /^\d{8}$/.test(missingUpdated))
      ? ' Missing card data and prices were last downloaded ' + fmtDate(parseStamp(missingUpdated)) + '; use Refresh missing data to update them.' : '';
    note.textContent = 'Showing cards you do not own from the sets you own at least one card of. Price is for one copy, so the total is the cost of one of each.' + updated +
      (skipped.length ? ' Not listed (too large or could not be downloaded): ' + skipped.join(', ') + '.' : '');
    note.style.display = '';
  } else {
    note.style.display = 'none';
  }
  currentPage = 1;
  refreshRows();
  renderTotals();
  renderTable();
}

function applySearch() {
  currentPage = 1;
  refreshRows();
  renderTotals();
  renderTable();
}

// Debounced so typing stays smooth with large lists
function onSearchChange() {
  clearTimeout(searchTimer);
  searchTimer = setTimeout(applySearch, 200);
}

function clearSearch() {
  var input = document.getElementById('searchInput');
  clearTimeout(searchTimer);
  input.value = '';
  input.focus();
  applySearch();
}

function onPageSizeChange() {
  rowsPerPage = parseInt(document.getElementById('pageSize').value, 10) || 10;
  currentPage = 1;
  renderTable();
}

function onSort(headerElement) {
  var key = headerElement.getAttribute('data-key');
  if (currentSortKey === key) {
    currentSortDir = currentSortDir === 'asc' ? 'desc' : 'asc';
  } else {
    currentSortKey = key;
    currentSortDir = 'asc';
  }

  var ths = document.querySelectorAll('#headerRow th');
  for (var i = 0; i < ths.length; i++) { ths[i].className = ''; }
  headerElement.className = currentSortDir;

  currentPage = 1;
  refreshRows();
  renderTable();
}

// Windowed pagination: first, last and 2 either side of the current page
function renderControls(totalPages) {
  var container = document.getElementById('paginationControls');
  container.innerHTML = '';
  if (totalPages <= 1) return;

  function addButton(label, page, disabled, active) {
    var btn = document.createElement('button');
    btn.className = 'pagination-btn' + (active ? ' active' : '');
    btn.innerText = label;
    btn.disabled = !!disabled;
    btn.onclick = function() { currentPage = page; renderTable(); };
    container.appendChild(btn);
  }

  addButton('Prev', currentPage - 1, currentPage === 1, false);

  var previous = 0;
  for (var p = 1; p <= totalPages; p++) {
    if (p === 1 || p === totalPages || Math.abs(p - currentPage) <= 2) {
      if (p - previous > 1) {
        var gap = document.createElement('span');
        gap.className = 'ellipsis';
        gap.innerText = '...';
        container.appendChild(gap);
      }
      addButton(p, p, false, p === currentPage);
      previous = p;
    }
  }

  addButton('Next', currentPage + 1, currentPage === totalPages, false);
}

// ---------- Set completion pills ----------
// One pill per set in the collection: unique cards owned (by set#number, so extra copies and foil/normal
// duplicates count once) out of the set's total from Scryfall, drawn as a progress bar with the percentage inside.
// Always based on the whole collection (not the search box). Clicking a pill searches for that set.
function renderSetProgress() {
  var box = document.getElementById('setProgress');
  var panel = document.getElementById('setPanel');
  var info = (typeof setInfo !== 'undefined' && setInfo) ? setInfo : {};
  var owned = {};
  ownedRows.forEach(function(r) {
    var cn = r.plain['collector_number'] || '';
    var code = cn.split('#')[0].toLowerCase();
    if (!code) return;
    if (!owned[code]) owned[code] = {};
    owned[code][cn.toLowerCase()] = true;
  });

  var list = Object.keys(owned).map(function(code) {
    var have = Object.keys(owned[code]).length;
    var meta = info[code] || null;
    var total = (meta && meta.total > 0) ? meta.total : 0;
    var exact = total ? Math.min(have / total * 100, 100) : -1;
    return { code: code, name: meta ? meta.name : code.toUpperCase(), have: have, total: total, exact: exact };
  });
  list.sort(function(a, b) { return (b.exact - a.exact) || (b.have - a.have) || collator.compare(a.code, b.code); });

  box.innerHTML = '';
  if (list.length === 0) { panel.style.display = 'none'; return; }
  panel.style.display = '';
  document.getElementById('setSummary').textContent = 'Set completion (' + list.length + (list.length === 1 ? ' set)' : ' sets)');

  list.forEach(function(s) {
    var CODE = s.code.toUpperCase();
    var pill = document.createElement('div');
    var label = document.createElement('span');
    label.className = 'set-label';
    if (s.total) {
      var pct = Math.round(s.exact);
      var pctText = (pct === 0 && s.have > 0) ? '<1%' : pct + '%';
      var fill = document.createElement('div');
      fill.className = 'set-fill';
      fill.style.width = (s.have > 0 ? Math.max(s.exact, 1.5) : 0) + '%';
      pill.className = 'set-pill' + (s.have >= s.total ? ' complete' : '');
      pill.appendChild(fill);
      label.textContent = CODE + ' \u00B7 ' + s.have + '/' + s.total + ' \u00B7 ' + pctText;
      pill.title = s.name + ' (' + CODE + '): ' + s.have + ' of ' + s.total + ' cards. Click to search this set.';
    } else {
      pill.className = 'set-pill no-total';
      label.textContent = CODE + ' \u00B7 ' + s.have + ' cards';
      pill.title = s.name + ' (' + CODE + '): ' + s.have + ' cards (set size not available). Click to search this set.';
    }
    pill.appendChild(label);
    pill.onclick = function() {
      document.getElementById('searchInput').value = s.code + '#';
      applySearch();
    };
    box.appendChild(pill);
  });
}

// ---------- Collection value history: line chart + "since last week" figure ----------
// valueHistory holds one {date, usd} entry per day; values are shown in AUD at today's rate,
// so exchange-rate moves do not appear as changes in the value of the collection.
function parseStamp(s) {
  s = String(s);
  return new Date(parseInt(s.substr(0, 4), 10), parseInt(s.substr(4, 2), 10) - 1, parseInt(s.substr(6, 2), 10));
}

function fmtDate(d) {
  return d.toLocaleDateString('en-AU', { day: 'numeric', month: 'short', year: 'numeric' });
}

function fmtMoney(n) {
  return '$' + n.toLocaleString('en-AU', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function signOf(n) { return n > 0 ? '+' : (n < 0 ? '-' : ''); }

function renderHistoryInner() {
  var chart = document.getElementById('historyChart');
  var deltaBox = document.getElementById('historyDelta');

  // Accept a flat array (normal), or unwrap {"value":[...]} just in case; then keep only well-formed entries.
  var list = Array.isArray(valueHistory) ? valueHistory
           : (valueHistory && Array.isArray(valueHistory.value) ? valueHistory.value : []);
  var pts = list.filter(function(e) { return e && /^\d{8}$/.test(String(e.date)) && isFinite(parseFloat(e.usd)); })
    .map(function(e) { return { d: parseStamp(e.date), v: parseFloat(e.usd) * audRate }; })
    .sort(function(a, b) { return a.d - b.d; });

  if (pts.length === 0) {
    deltaBox.textContent = '';
    chart.innerHTML = '<div class="history-note">No value history saved yet. It is recorded each day the script runs with every card loading successfully.</div>';
    return;
  }

  if (pts.length < 2) {
    deltaBox.textContent = '';
    chart.innerHTML = '<div class="history-note">Today\'s value (' + fmtMoney(pts[0].v) + ') is saved. The chart appears once the script has run on a second day.</div>';
    return;
  }

  // Change since a week ago (or since the first saved day if there is less than a week of history)
  var last = pts[pts.length - 1];
  var target = new Date(last.d);
  target.setDate(target.getDate() - 7);
  var base = null;
  for (var i = 0; i < pts.length - 1; i++) { if (pts[i].d <= target) base = pts[i]; }
  var label = 'Since last week';
  if (!base) { base = pts[0]; label = 'Since ' + fmtDate(base.d); }
  var diff = Math.round((last.v - base.v) * 100) / 100;
  var pct = base.v > 0 ? (diff / base.v * 100) : 0;
  deltaBox.className = 'history-delta ' + (diff > 0 ? 'gain' : (diff < 0 ? 'loss' : ''));
  deltaBox.textContent = label + ': ' + signOf(diff) + fmtMoney(Math.abs(diff)) +
    ' (' + signOf(pct) + Math.abs(pct).toFixed(1) + '%)';

  // Line chart drawn as inline SVG (no libraries)
  var W = 800, H = 230, padL = 78, padR = 18, padT = 16, padB = 32;
  var tMin = pts[0].d.getTime(), tMax = last.d.getTime();
  var vals = pts.map(function(p) { return p.v; });
  var vMin = Math.min.apply(null, vals), vMax = Math.max.apply(null, vals);
  var pad = (vMax - vMin) * 0.08 || Math.max(vMax * 0.02, 1);
  vMin -= pad; vMax += pad;
  function x(t) { return padL + (tMax === tMin ? 0 : (t - tMin) / (tMax - tMin)) * (W - padL - padR); }
  function y(v) { return padT + (1 - (v - vMin) / (vMax - vMin)) * (H - padT - padB); }

  var line = pts.map(function(p) { return x(p.d.getTime()).toFixed(1) + ',' + y(p.v).toFixed(1); }).join(' ');
  var area = 'M' + x(tMin).toFixed(1) + ',' + (H - padB) + ' L' + line.replace(/ /g, ' L') +
             ' L' + x(tMax).toFixed(1) + ',' + (H - padB) + ' Z';

  var svg = '<svg viewBox="0 0 ' + W + ' ' + H + '" role="img" aria-label="Collection value over time">';
  for (var k = 0; k <= 2; k++) {
    var gv = vMin + (vMax - vMin) * k / 2, gy = y(gv);
    svg += '<line class="h-grid" x1="' + padL + '" y1="' + gy.toFixed(1) + '" x2="' + (W - padR) + '" y2="' + gy.toFixed(1) + '"/>' +
           '<text class="h-label" x="' + (padL - 8) + '" y="' + (gy + 4).toFixed(1) + '" text-anchor="end">$' +
           gv.toLocaleString('en-AU', { maximumFractionDigits: 0 }) + '</text>';
  }
  svg += '<path class="h-area" d="' + area + '"/><polyline class="h-line" points="' + line + '"/>';
  pts.forEach(function(p, idx) {
    if (pts.length > 60 && idx !== pts.length - 1) return;     // keep long histories uncluttered
    svg += '<circle class="h-dot" cx="' + x(p.d.getTime()).toFixed(1) + '" cy="' + y(p.v).toFixed(1) + '" r="3.5"><title>' +
           fmtDate(p.d) + ': ' + fmtMoney(p.v) + '</title></circle>';
  });
  svg += '<text class="h-label" x="' + padL + '" y="' + (H - 8) + '" text-anchor="start">' + fmtDate(pts[0].d) + '</text>' +
         '<text class="h-label" x="' + (W - padR) + '" y="' + (H - 8) + '" text-anchor="end">' + fmtDate(last.d) + '</text></svg>';
  chart.innerHTML = svg;
}

// Wrapper: a problem in the chart can never stop the rest of the report from loading
function renderHistory() {
  try { renderHistoryInner(); }
  catch (e) {
    document.getElementById('historyChart').innerHTML =
      '<div class="history-note">Chart error: ' + escapeHtml(e.message) + '</div>';
  }
}

// Launches the local PowerShell script via the tcgrefresh:// protocol.
// The script rebuilds this file; reload the page (F5) when its window closes.
function refreshData() {
  window.location.href = 'tcgrefresh://run';
}

// Same script, but it also re-downloads every set's cards (the missing cards cache) from Scryfall.
// Use it when you want fresh missing-card prices or new cards; reload the page (F5) when its window closes.
function refreshMissingData() {
  window.location.href = 'tcgrefreshmissing://run';
}

// ---------- Excel (.xlsx) export: no external libraries needed ----------
function crc32(bytes) {
  var table = crc32.table;
  if (!table) {
    table = crc32.table = [];
    for (var n = 0; n < 256; n++) {
      var c = n;
      for (var k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      table[n] = c >>> 0;
    }
  }
  var crc = 0xFFFFFFFF;
  for (var i = 0; i < bytes.length; i++) crc = table[(crc ^ bytes[i]) & 0xFF] ^ (crc >>> 8);
  return (crc ^ 0xFFFFFFFF) >>> 0;
}

function u16(v) { return [v & 255, (v >>> 8) & 255]; }
function u32(v) { return [v & 255, (v >>> 8) & 255, (v >>> 16) & 255, (v >>> 24) & 255]; }

// Builds an uncompressed .zip container from [{name, data}] (data is a string)
function buildZip(files) {
  var enc = new TextEncoder();
  var parts = [], central = [], offset = 0;
  files.forEach(function(f) {
    var nameBytes = enc.encode(f.name);
    var data = enc.encode(f.data);
    var crc = crc32(data);
    var local = [].concat(u32(0x04034b50), u16(20), u16(0x0800), u16(0), u16(0), u16(0x21),
      u32(crc), u32(data.length), u32(data.length), u16(nameBytes.length), u16(0));
    parts.push(new Uint8Array(local), nameBytes, data);
    central.push(new Uint8Array([].concat(u32(0x02014b50), u16(20), u16(20), u16(0x0800), u16(0), u16(0), u16(0x21),
      u32(crc), u32(data.length), u32(data.length), u16(nameBytes.length), u16(0), u16(0), u16(0), u16(0), u32(0), u32(offset))), nameBytes);
    offset += local.length + nameBytes.length + data.length;
  });
  var centralSize = 0;
  central.forEach(function(c) { centralSize += c.length; });
  var end = new Uint8Array([].concat(u32(0x06054b50), u16(0), u16(0), u16(files.length), u16(files.length),
    u32(centralSize), u32(offset), u16(0)));
  return parts.concat(central, [end]);
}

function xmlEscape(s) {
  return String(s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function colLetter(i) {
  var s = '';
  i++;
  while (i > 0) { var m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); }
  return s;
}

function textCell(ref, text, style) {
  return '<c r="' + ref + '" t="inlineStr"' + (style ? ' s="' + style + '"' : '') +
    '><is><t xml:space="preserve">' + xmlEscape(text) + '</t></is></c>';
}

function numberCell(ref, num) {
  return '<c r="' + ref + '" s="2"><v>' + num + '</v></c>';
}

function linkCell(ref, url, text) {
  var f = 'HYPERLINK("' + url.replace(/"/g, '""') + '","' + text.replace(/"/g, '""') + '")';
  return '<c r="' + ref + '" t="str"><f>' + xmlEscape(f) + '</f><v>' + xmlEscape(text) + '</v></c>';
}

// Builds the Excel file from ALL rows in cardData (ignores search/paging)
function buildXlsxBlob() {
  var header = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  var widths = columns.map(function(col) { return col.header.length; });
  var rows = [];

  var headerCells = columns.map(function(col, c) { return textCell(colLetter(c) + '1', col.header, 1); });
  rows.push('<row r="1">' + headerCells.join('') + '</row>');

  var data = activeData();
  data.forEach(function(item, r) {
    var rowNum = r + 2;
    var cells = columns.map(function(col, c) {
      var ref = colLetter(c) + rowNum;
      var text = plainValue(col, item);
      if (text.length > widths[c]) widths[c] = Math.min(text.length, 60);
      if (col.type === 'number') {
        var num = parseFloat(text.replace(/[^0-9.-]/g, ''));
        return isNaN(num) ? textCell(ref, text) : numberCell(ref, num);
      }
      if (col.type === 'html') {
        var m = /<a\s[^>]*href\s*=\s*['"]([^'"]*)['"]/i.exec(String(item[col.key] || ''));
        if (m && text) return linkCell(ref, m[1], text);
      }
      return textCell(ref, text);
    });
    rows.push('<row r="' + rowNum + '">' + cells.join('') + '</row>');
  });

  var cols = widths.map(function(w, i) {
    return '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + (w + 3) + '" customWidth="1"/>';
  }).join('');
  var lastRef = colLetter(columns.length - 1) + (data.length + 1);

  var sheet = header + '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>' +
    '<cols>' + cols + '</cols><sheetData>' + rows.join('') + '</sheetData>' +
    '<autoFilter ref="A1:' + lastRef + '"/></worksheet>';

  var styles = header + '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>' +
    '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +
    '<xf numFmtId="4" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>' +
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';

  var files = [
    { name: '[Content_Types].xml', data: header + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>' },
    { name: '_rels/.rels', data: header + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>' },
    { name: 'xl/workbook.xml', data: header + '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Cards" sheetId="1" r:id="rId1"/></sheets></workbook>' },
    { name: 'xl/_rels/workbook.xml.rels', data: header + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>' },
    { name: 'xl/styles.xml', data: styles },
    { name: 'xl/worksheets/sheet1.xml', data: sheet }
  ];

  return new Blob(buildZip(files), { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

function exportToExcel() {
  if (activeData().length === 0) return;
  var blob = buildXlsxBlob();
  var url = URL.createObjectURL(blob);
  var a = document.createElement('a');
  a.href = url;
  a.download = showMissing ? 'Scryfall_Missing_Cards.xlsx' : 'Scryfall_Price_Report.xlsx';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(function() { URL.revokeObjectURL(url); }, 1000);
}

// ---------- Card image popup: the picture is loaded from Scryfall's image servers on hover (nothing stored locally) ----------
var preview = null;
var previewImg = null;

function hidePreview() {
  if (preview) preview.style.display = 'none';
}

function movePreview(e) {
  var w = 300, h = 418, gap = 20;
  var x = e.clientX + gap;
  if (x + w > window.innerWidth - 8) x = e.clientX - gap - w;     // flip to the left near the right edge
  var y = e.clientY - h / 2;
  y = Math.max(8, Math.min(y, window.innerHeight - h - 8));       // keep fully on screen
  preview.style.left = Math.max(8, x) + 'px';
  preview.style.top = y + 'px';
}

function initPreview() {
  preview = document.createElement('div');
  preview.id = 'cardPreview';
  previewImg = document.createElement('img');
  previewImg.alt = '';
  previewImg.onerror = hidePreview;
  preview.appendChild(previewImg);
  document.body.appendChild(preview);

  var body = document.getElementById('tableBody');
  body.addEventListener('mouseover', function(e) {
    var link = e.target.closest ? e.target.closest('a[data-img]') : null;
    var src = link ? link.getAttribute('data-img') : '';
    if (!src) return;
    if (previewImg.getAttribute('src') !== src) previewImg.setAttribute('src', src);
    preview.style.display = 'block';
    movePreview(e);
  });
  body.addEventListener('mousemove', function(e) {
    if (preview.style.display === 'block') movePreview(e);
  });
  body.addEventListener('mouseout', function(e) {
    if (e.target.closest && e.target.closest('a[data-img]')) hidePreview();
  });
}

buildHeader();
var sortTh = document.querySelector('#headerRow th[data-key="' + currentSortKey + '"]');
if (sortTh) sortTh.className = currentSortDir;
ownedRows = makeRows(cardData);
missingRows = makeRows(typeof missingData !== 'undefined' ? missingData : []);
rows = ownedRows;
initViewToggle();
refreshRows();
renderTotals();
renderSetProgress();
renderHistory();
renderTable();
initPreview();
