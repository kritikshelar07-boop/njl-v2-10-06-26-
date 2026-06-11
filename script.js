// ── CONFIGURATION & STATE ────────────────────────────────────────────────────
// Dynamically routes to the correct API endpoint wherever it is hosted
const API_BASE = window.location.origin + "/api";

let currentPage     = 1;
const PAGE_SIZE     = 100;
// Dropdown filter selections (pipe-separated when sent to server)
const dropdownSelections = {
  category:          new Set(),
  subcategory:       new Set(),
  product_group:     new Set(),
  sub_product_group: new Set(),
  sku_status:        new Set(),
};

// Paste-input filter selections (Vendor, Warehouse, Serial, HUID, Location)
let selectedSerials    = new Set();
let selectedHuids      = new Set();
let selectedWarehouses = new Set();
let selectedLocations  = new Set();
let selectedVendors    = new Set();
let globalQuery     = '';   // raw input string
let globalTerms     = [];   // parsed array of lowercase terms
let sortMode        = 'default';
let totalItemsAll   = 0;

// DATA STATE VARIABLES
let currentData = [];
const selectedForDownload = new Map();

window.handleImgErr = function(el) {
  el.outerHTML = '<div class="img-placeholder"><div class="gem-icon">💎</div><div class="no-img">No Image</div></div>';
};

// ── INIT ─────────────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', async function init() {
  try {
    const statRes = await fetch(`${API_BASE}/stats`);
    const statData = await statRes.json();
    totalItemsAll = statData.total;

    document.getElementById('statTotal').textContent = totalItemsAll.toLocaleString();
    document.getElementById('countTotal').textContent = totalItemsAll.toLocaleString();

    // Populate dropdown filters (instant — data already in server cache)
    await Promise.all([
      populateDropdown('category'),
      populateDropdown('subcategory'),
      populateDropdown('product_group'),
      populateDropdown('sub_product_group'),
      populateDropdown('sku_status'),
    ]);
    try {
      await render();
    } catch (renderErr) {
      console.error("render() failed during init:", renderErr);
      document.getElementById('cardGrid').innerHTML =
        `<div style="grid-column:1/-1;color:#c0392b;padding:20px;">⚠️ Failed to load inventory: ${renderErr.message}</div>`;
    }
    // Pre-load dup-HUID data so badge shows count immediately
    loadDupData();
  } catch (error) {
    console.error("Failed to connect to the backend.", error);
    document.getElementById('cardGrid').innerHTML = '<div style="grid-column: 1/-1; color: red; padding: 20px;">⚠️ Could not connect to the server. Unable to Stream the database.</div>';
  }
});

// ── SELECTION LOGIC & VISIBILITY TOGGLE ──────────────────────────────────────
function updateSelectionUI() {
  const count = selectedForDownload.size;
  const fab = document.getElementById('selectionActionBar');
  const fabCount = document.getElementById('sabCountVal');

  // Trigger floating action bar visibility
  if (count > 0) {
    fabCount.innerText = count;
    fab.classList.add('visible');
  } else {
    fab.classList.remove('visible');
  }

  // Sync "Select All" checkbox based on current page selections
  const checkboxes = document.querySelectorAll('.card-checkbox');
  if (checkboxes.length > 0) {
    const allChecked = Array.from(checkboxes).every(cb => cb.checked);
    document.getElementById('selectAllPage').checked = allChecked;
  } else {
    document.getElementById('selectAllPage').checked = false;
  }
}

window.toggleSelection = function(checkbox) {
  const id = checkbox.value;
  if (checkbox.checked) {
    const rowData = currentData.find(r => String(r.SERIALNUMBER) === id);
    if (rowData) selectedForDownload.set(id, rowData);
  } else {
    selectedForDownload.delete(id);
  }
  updateSelectionUI();
};

window.toggleSelectAllPage = function(selectAllCb) {
  const isChecked = selectAllCb.checked;
  const checkboxes = document.querySelectorAll('.card-checkbox');

  checkboxes.forEach(cb => {
    cb.checked = isChecked;
    const id = cb.value;
    if (isChecked) {
      const rowData = currentData.find(r => String(r.SERIALNUMBER) === id);
      if (rowData) selectedForDownload.set(id, rowData);
    } else {
      selectedForDownload.delete(id);
    }
  });
  updateSelectionUI();
};

window.clearAllSelections = function() {
  selectedForDownload.clear();
  updateSelectionUI();

  // Uncheck all rendered checkboxes on the current page
  const checkboxes = document.querySelectorAll('.card-checkbox');
  checkboxes.forEach(cb => cb.checked = false);
  document.getElementById('selectAllPage').checked = false;
};

window.downloadSelectedExcel = function() {
  if (selectedForDownload.size === 0) return;

  const dataToExport = Array.from(selectedForDownload.values());

  try {
    const ws = XLSX.utils.json_to_sheet(dataToExport);

    // Apply auto-ish column widths based on headers
    const headers = Object.keys(dataToExport[0] || {});
    ws['!cols'] = headers.map(h => ({ wch: Math.max(String(h).length + 4, 18) }));

    // Bold Headers
    for (let c = 0; c < headers.length; c++) {
      const addr = XLSX.utils.encode_cell({ r: 0, c });
      if (ws[addr]) ws[addr].s = { font: { bold: true } };
    }

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Selected_Products");

    // File Name Generation
    const today = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, `Indriya_Selected_Products_${today}.xlsx`);

  } catch (err) {
    console.error("Error exporting to Excel:", err);
    alert("An error occurred while generating the Excel file.");
  }
};

// ── DROPDOWN FILTERS ─────────────────────────────────────────────────────────
const DROPDOWN_LABEL = {
  category:          'Category',
  subcategory:       'Subcategory',
  product_group:     'Product Group',
  sub_product_group: 'Sub-Product Group',
  sku_status:        'SKU Status',
};

// Full cached list per type (for client-side search filtering)
const dropdownCache = {};

async function populateDropdown(type, searchQ = '') {
  const listEl  = document.getElementById(`${type}List`);
  if (!listEl) return;

  // Fetch from server (fast — returns pre-built cache)
  if (!dropdownCache[type] || searchQ === '') {
    try {
      const res  = await fetch(`${API_BASE}/dropdown-values?type=${type}&q=${encodeURIComponent(searchQ)}`);
      const data = await res.json();
      if (!searchQ) dropdownCache[type] = data.values; // cache full list
      renderDropdownList(type, searchQ ? data.values : dropdownCache[type], listEl);
    } catch (e) { console.error(`Dropdown fetch failed for ${type}`, e); }
  } else {
    // Client-side filter on cached data (instant)
    const ql = searchQ.toLowerCase();
    const filtered = dropdownCache[type].filter(([v]) => v.toLowerCase().includes(ql));
    renderDropdownList(type, filtered, listEl);
  }
}

function renderDropdownList(type, entries, listEl) {
  const sel = dropdownSelections[type];
  listEl.innerHTML = '';
  entries.forEach(([val, count]) => {
    const safeId = `dd-${type}-${val.replace(/[^a-zA-Z0-9]/g, '_')}`;
    const div = document.createElement('div');
    div.className = 'filter-item' + (sel.has(val) ? ' selected' : '');
    div.innerHTML = `
      <input type="checkbox" id="${safeId}" ${sel.has(val) ? 'checked' : ''}/>
      <label class="filter-item-label" for="${safeId}" title="${escHtml(val)}">${escHtml(val)}</label>
      <span class="filter-count">${count.toLocaleString()}</span>`;
    div.querySelector('input').addEventListener('change', e => {
      if (e.target.checked) sel.add(val); else sel.delete(val);
      div.classList.toggle('selected', sel.has(val));
      updateDropdownChips(type);
      currentPage = 1;
      render();
    });
    listEl.appendChild(div);
  });
}

window.filterDropdownList = function(type) {
  const inputId = `${type}Search`;
  const el = document.getElementById(inputId);
  const q = el ? el.value : '';
  populateDropdown(type, q);
};

function updateDropdownChips(type) {
  const set  = dropdownSelections[type];
  const wrap = document.getElementById(`${type}-chips`);
  if (!wrap) return;
  wrap.innerHTML = '';
  set.forEach(v => {
    const chip = document.createElement('span');
    chip.className = 'chip';
    chip.title = v;
    chip.innerHTML = `${v.length > 14 ? v.slice(0,12)+'…' : escHtml(v)} <span class="remove">✕</span>`;
    chip.addEventListener('click', () => {
      set.delete(v);
      updateDropdownChips(type);
      populateDropdown(type);
      currentPage = 1;
      render();
    });
    wrap.appendChild(chip);
  });
}

window.clearDropdown = function(type) {
  dropdownSelections[type].clear();
  updateDropdownChips(type);
  populateDropdown(type);
  currentPage = 1;
  render();
};

// ── PASTE-INPUT FILTERS (Vendor, Warehouse, Serial, HUID, Location) ──────────
// No dropdown list needed. User types or pastes values into a textarea.
// Values split on newline/comma, deduplicated, shown as chips, applied instantly.

const PASTE_FILTER_SETS = {
  vendor:    () => selectedVendors,
  warehouse: () => selectedWarehouses,
  serial:    () => selectedSerials,
  huid:      () => selectedHuids,
  location:  () => selectedLocations,
};

function getPasteInputEl(type) {
  return document.getElementById(type + 'PasteInput');
}

function getPasteHintEl(type) {
  return document.getElementById(type + 'PasteHint');
}

function tokeniseInput(raw) {
  // Strategy: split on explicit separators first (newline/tab/comma/semicolon/pipe).
  // Then, for any token that contains a space, split it further on whitespace —
  // because ERP codes (serials, HUIDs, vendor accounts) are never multi-word.
  // This handles: "VC-0119 VC-0120" pasted on one line, AND newline-separated lists.
  const primary = raw.split(/[\n\r\t,;|]+/).map(t => t.trim()).filter(Boolean);
  const result = [];
  primary.forEach(tok => {
    if (tok.includes(' ')) {
      // secondary split on whitespace — each part is its own code
      tok.split(/\s+/).map(t => t.trim()).filter(Boolean).forEach(t => result.push(t));
    } else {
      result.push(tok);
    }
  });
  return result;
}

let pasteInputTimeouts = {};
const _pasteInFlight = {};   // flag: a paste event just fired, suppress oninput

// onpaste fires BEFORE the value is written to the DOM.
// Set a flag, defer by one tick so the browser commits the value, then apply.
window.handlePasteFilterPaste = function(e, type) {
  _pasteInFlight[type] = true;
  clearTimeout(pasteInputTimeouts[type]);
  setTimeout(() => {
    _pasteInFlight[type] = false;
    _applyPasteTokens(type, true);
  }, 0);
};

// oninput fires for every keystroke AND right after a paste (redundantly).
// If a paste is already being handled, do nothing — avoid double-apply / race.
// For normal typing: commit only on Enter/comma (handled by onkeydown),
// but also keep a 400ms debounce as a fallback for single-value typed input.
window.handlePasteFilterInput = function(e, type) {
  if (_pasteInFlight[type]) return;   // paste owns this update
  clearTimeout(pasteInputTimeouts[type]);
  pasteInputTimeouts[type] = setTimeout(() => _applyPasteTokens(type, true), 400);
};

// onkeydown: Enter or comma commits whatever is typed right now.
window.handlePasteFilterKeydown = function(e, type) {
  if (e.key === 'Enter' || e.key === ',') {
    e.preventDefault();
    clearTimeout(pasteInputTimeouts[type]);
    _applyPasteTokens(type, true);
  }
};

function _applyPasteTokens(type, clearAfterApply) {
  const inputEl = getPasteInputEl(type);
  const hintEl  = getPasteHintEl(type);
  const set     = PASTE_FILTER_SETS[type]();
  const raw     = inputEl ? inputEl.value : '';
  const tokens  = tokeniseInput(raw);

  if (tokens.length === 0) return;

  let added = 0;
  const seenToks = new Set();
  tokens.forEach(tok => {
    const norm = tok.trim().toUpperCase();
    if (!norm) return;
    if (seenToks.has(norm)) return;
    seenToks.add(norm);
    if (!set.has(norm)) { set.add(norm); added++; }
  });

  // Always clear the textarea after applying — no stale text left behind.
  if (inputEl) inputEl.value = '';

  if (hintEl) {
    if (added > 0) {
      hintEl.textContent = '✓ ' + added + ' value' + (added > 1 ? 's' : '') + ' added';
      hintEl.className = 'paste-filter-hint paste-filter-hint--ok';
    } else {
      hintEl.textContent = 'Already active';
      hintEl.className = 'paste-filter-hint paste-filter-hint--warn';
    }
    clearTimeout(hintEl._t);
    hintEl._t = setTimeout(() => {
      if (hintEl) { hintEl.textContent = ''; hintEl.className = 'paste-filter-hint'; }
    }, 2500);
  }

  // Only re-render if something actually changed.
  if (added > 0) {
    updatePasteChips(type);
    currentPage = 1;
    render();
  }
}

function updatePasteChips(type) {
  const set  = PASTE_FILTER_SETS[type]();
  const wrapId =
    type === 'serial'    ? 'serial-chips' :
    type === 'huid'      ? 'huid-chips' :
    type === 'warehouse' ? 'warehouse-chips' :
    type === 'location'  ? 'location-chips' :
    'vendor-chips';
  const wrap = document.getElementById(wrapId);
  if (!wrap) return;
  wrap.innerHTML = '';
  set.forEach(v => {
    const c = document.createElement('span');
    c.className = 'chip';
    c.title = v;
    c.innerHTML = (v.length > 14 ? v.slice(0, 12) + '…' : escHtml(v)) + ' <span class="remove">✕</span>';
    c.addEventListener('click', () => {
      set.delete(v);
      updatePasteChips(type);
      currentPage = 1;
      render();
    });
    wrap.appendChild(c);
  });
}

window.clearPasteFilter = function(type) {
  PASTE_FILTER_SETS[type]().clear();
  const inputEl = getPasteInputEl(type);
  const hintEl  = getPasteHintEl(type);
  if (inputEl) inputEl.value = '';
  if (hintEl)  { hintEl.textContent = ''; hintEl.className = 'paste-filter-hint'; }
  updatePasteChips(type);
  currentPage = 1;
  render();
};

// Backward compat alias (called from selectDupHuid and resetAll)
function updateChips(type) { updatePasteChips(type); }

// ── BULK PASTE FOR DROPDOWN FILTERS (category, subcategory, etc.) ────────────
async function handleDropdownPaste(e, type) {
  const pasted = (e.clipboardData || window.clipboardData).getData('text');
  const tokens = pasted.split(/[\n\r\t,]+/).map(t => t.trim()).filter(Boolean);
  if (tokens.length < 2) return;

  e.preventDefault();

  if (!dropdownCache[type]) {
    try {
      const res  = await fetch(`${API_BASE}/dropdown-values?type=${type}&q=`);
      const data = await res.json();
      dropdownCache[type] = data.values;
    } catch (err) {
      console.error('Dropdown bulk paste: fetch failed', err);
      return;
    }
  }

  const allValues = dropdownCache[type].map(([v]) => v);
  const allLower  = allValues.map(v => v.toLowerCase());
  const sel       = dropdownSelections[type];
  let matched = 0, unmatched = [];

  tokens.forEach(tok => {
    const idx = allLower.indexOf(tok.toLowerCase());
    if (idx !== -1) { sel.add(allValues[idx]); matched++; }
    else unmatched.push(tok);
  });

  const inputEl = document.getElementById(`${type}Search`);
  if (inputEl) inputEl.value = '';

  updateDropdownChips(type);
  populateDropdown(type);
  currentPage = 1;
  render();
}


// ── GLOBAL SEARCH ─────────────────────────────────────────────────────────────
function parseTerms(raw) {
  return raw.split(/[,\s]+/).map(t => t.trim().toLowerCase()).filter(Boolean);
}

function buildServerQuery() {
  return globalTerms.join(' ');
}

function committedTerms() {
  if (globalTerms.length === 0) return [];
  return /[,\s]$/.test(globalQuery) ? globalTerms : globalTerms.slice(0, -1);
}

let searchTimeout;
document.getElementById('globalSearch')?.addEventListener('input', e => {
  globalQuery = e.target.value;
  globalTerms = parseTerms(globalQuery);
  document.getElementById('searchClear').classList.toggle('visible', globalQuery.trim().length > 0);
  renderSearchTermPills();
  clearTimeout(searchTimeout);
  searchTimeout = setTimeout(() => {
    currentPage = 1;
    render();
  }, 250);
});

document.getElementById('searchClear')?.addEventListener('click', () => {
  document.getElementById('globalSearch').value = '';
  globalQuery = '';
  globalTerms = [];
  document.getElementById('searchClear').classList.remove('visible');
  renderSearchTermPills();
  currentPage = 1;
  render();
});

// ── SORT ──────────────────────────────────────────────────────────────────────
window.applySort = function() {
  sortMode = document.getElementById('sortSelect').value;
  currentPage = 1;
  render();
}

// ── MAIN RENDER ───────────────────────────────────────────────────────────────
async function render() {
  const grid = document.getElementById('cardGrid');
  grid.style.opacity = '0.4';
  grid.style.pointerEvents = 'none';

  const serialsParam    = Array.from(selectedSerials).join(',');
  const huidsParam      = Array.from(selectedHuids).join(',');
  const warehousesParam = Array.from(selectedWarehouses).join(',');
  const locationsParam  = Array.from(selectedLocations).join(',');
  const vendorsParam    = Array.from(selectedVendors).join(',');

  const url = new URL(`${API_BASE}/inventory`);
  url.searchParams.append('page', currentPage);
  url.searchParams.append('page_size', PAGE_SIZE);
  url.searchParams.append('sort', sortMode);
  const serverQ = buildServerQuery();
  if (serverQ)        url.searchParams.append('q',          serverQ);
  if (serialsParam)   url.searchParams.append('serials',    serialsParam);
  if (huidsParam)     url.searchParams.append('huids',      huidsParam);
  if (warehousesParam) url.searchParams.append('warehouses', warehousesParam);
  if (locationsParam) url.searchParams.append('locations',  locationsParam);
  if (vendorsParam)   url.searchParams.append('vendors',    vendorsParam);

  for (const [key, set] of Object.entries(dropdownSelections)) {
    if (set.size) {
      let paramName = key + 's';
      if (key === 'category')          { paramName = 'categories'; }
      else if (key === 'subcategory')  { paramName = 'subcategories'; }
      else if (key === 'product_group') { paramName = 'product_groups'; }
      else if (key === 'sub_product_group') { paramName = 'sub_prod_grps'; }
      else if (key === 'sku_status')   { paramName = 'sku_statuses'; }
      url.searchParams.append(paramName, Array.from(set).join('|'));
    }
  }

  try {
    const res = await fetch(url);
    if (!res.ok) {
      const errText = await res.text().catch(() => `HTTP ${res.status}`);
      throw new Error(`Server returned ${res.status}: ${errText}`);
    }
    const responseData = await res.json();
    const total = responseData.total_filtered;
    const pageData = responseData.data;

    currentData = pageData;
    const empty = document.getElementById('emptyState');

    grid.style.opacity = '';
    grid.style.pointerEvents = '';

    if (total === 0) {
      grid.innerHTML = '';
      empty.style.display = 'block';
    } else {
      empty.style.display = 'none';
      grid.innerHTML = pageData.map(renderCard).join('');
    }

    renderPagination(total, currentPage);

    const startItem = total === 0 ? 0 : ((currentPage - 1) * PAGE_SIZE) + 1;
    const endItem = Math.min(currentPage * PAGE_SIZE, total);

    document.getElementById('countShowing').textContent = `${startItem}–${endItem}`;
    document.getElementById('countTotal').textContent   = total.toLocaleString();
    document.getElementById('statShowing').textContent  = total.toLocaleString();
    document.getElementById('statPage').textContent     = currentPage;
    document.getElementById('badgeCount').textContent   = total.toLocaleString();

    updateSelectionUI();

  } catch (e) {
    grid.style.opacity = '';
    grid.style.pointerEvents = '';
    grid.innerHTML = `<div style="grid-column:1/-1;color:#c0392b;padding:20px;">⚠️ Failed to load inventory: ${e.message}</div>`;
    console.error("Error fetching inventory data", e);
  }
}

// ── SEARCH TERM PILLS ────────────────────────────────────────────────────────
function renderSearchTermPills() {
  let wrap = document.getElementById('searchTermPills');
  if (!wrap) return;
  wrap.innerHTML = '';

  const pillTerms = committedTerms();
  pillTerms.forEach(term => {
    const pill = document.createElement('span');
    pill.className = 'search-term-pill';
    pill.innerHTML = `${escHtml(term)} <span class="remove" onclick="removeSearchTerm('${escHtml(term)}')">✕</span>`;
    wrap.appendChild(pill);
  });
  wrap.style.display = pillTerms.length ? 'flex' : 'none';
}

window.removeSearchTerm = function(term) {
  const updated = globalTerms.filter(t => t.toLowerCase() !== term.toLowerCase());
  globalTerms = updated;
  globalQuery = updated.join(' ');
  document.getElementById('globalSearch').value = globalQuery;
  document.getElementById('searchClear').classList.toggle('visible', updated.length > 0);
  renderSearchTermPills();
  currentPage = 1;
  render();
};

// ── HIGHLIGHT ─────────────────────────────────────────────────────────────────
function hl(text) {
  if (!text && text !== 0) return '—';
  let safe = String(text).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  if (!globalTerms.length) return safe;
  globalTerms.forEach(term => {
    const escaped = term.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');
    safe = safe.replace(new RegExp('(' + escaped + ')', 'gi'), '<mark>$1</mark>');
  });
  return safe;
}

// ── RENDER CARD ───────────────────────────────────────────────────────────────
function renderCard(r) {
  const img = r.Image_Link ? String(r.Image_Link).trim() : '';
  const imgHtml = img
    ? `<img src="${img}" alt="${r.SERIALNUMBER}" loading="lazy" onerror="window.handleImgErr(this)">`
    : `<div class="img-placeholder"><div class="gem-icon">💎</div><div class="no-img">No Image</div></div>`;

  const cat    = r.Category          || '';
  const sub    = r.Subcategory       || '';
  const pg     = r.Product_Group     || '';
  const spg    = r.Sub_Product_Group || '';
  const huid   = r.HUID              || '';
  const status = (r.PWC_SKUSTATUS !== undefined && r.PWC_SKUSTATUS !== '' && r.PWC_SKUSTATUS !== 0) ? r.PWC_SKUSTATUS : '';
  const avail  = (r.AVAILABLE_PHYSICAL !== undefined && r.AVAILABLE_PHYSICAL !== null && r.AVAILABLE_PHYSICAL !== '') ? r.AVAILABLE_PHYSICAL : '';

  const isChecked = selectedForDownload.has(String(r.SERIALNUMBER)) ? 'checked' : '';

  return `
<div class="card">
  <div class="card-img-wrap">
    <input type="checkbox" class="card-checkbox" value="${r.SERIALNUMBER}" ${isChecked} onchange="toggleSelection(this)" title="Select to Download">

    ${imgHtml}
    ${cat ? `<span class="cat-badge">${hl(cat)}</span>` : ''}
  </div>
  <div class="card-body">
    <div class="serial-row">
      <div class="serial" title="${r.SERIALNUMBER}">${hl(r.SERIALNUMBER)}</div>
      <div class="itemid-tag" title="${r.ITEMID}">${hl(r.ITEMID)}</div>
    </div>
    <div class="gold-divider"></div>
    <div class="details-grid">
      <div class="detail-item">
        <div class="detail-key">Gross Qty</div>
        <div class="detail-val gold-val">${r.GROSSQTY !== '' && r.GROSSQTY !== 0 ? r.GROSSQTY : '<span class="na">—</span>'}</div>
      </div>
      <div class="detail-item">
        <div class="detail-key">Net Weight</div>
        <div class="detail-val gold-val">${r.NETWEIGHT !== '' && r.NETWEIGHT !== 0 ? r.NETWEIGHT+' g' : '<span class="na">—</span>'}</div>
      </div>
      <div class="detail-item">
        <div class="detail-key">Avail. Qty</div>
        <div class="detail-val highlight">${avail !== '' ? avail : '<span class="na">—</span>'}</div>
      </div>
      <div class="detail-item">
        <div class="detail-key">SKU Status</div>
        <div class="detail-val">${status ? `<span class="sku-status">${hl(status)}</span>` : '<span class="na">—</span>'}</div>
      </div>
      <div class="detail-item">
        <div class="detail-key">Vendor Acct</div>
        <div class="detail-val highlight" title="${r.VENDACCOUNT||''}">${hl(r.VENDACCOUNT)}</div>
      </div>
      <div class="detail-item">
        <div class="detail-key">HUID</div>
        <div class="detail-val" title="${huid}">${hl(huid)}</div>
      </div>
      <div class="detail-item" style="grid-column:1/-1">
        <div class="detail-key">Warehouse</div>
        <div class="detail-val highlight" title="${r.WAREHOUSE||''}${r.WAREHOUSE_NAME ? ' · ' + r.WAREHOUSE_NAME : ''}">${hl(r.WAREHOUSE || '—')}${r.WAREHOUSE_NAME ? `<span class="warehouse-name"> &nbsp;&nbsp;${hl(r.WAREHOUSE_NAME)}</span>` : ''}</div>
      </div>
      ${r.LOCATION ? `<div class="detail-item" style="grid-column:1/-1">
        <div class="detail-key">Location</div>
        <div class="detail-val location-val" title="${r.LOCATION}">${hl(r.LOCATION)}</div>
      </div>` : ''}
      ${cat ? `<div class="detail-item" style="grid-column:1/-1">
        <div class="detail-key">Category</div>
        <div class="detail-val">${hl(cat)}</div>
      </div>` : ''}
      ${sub ? `<div class="detail-item" style="grid-column:1/-1">
        <div class="detail-key">Subcategory</div>
        <div class="detail-val">${hl(sub)}</div>
      </div>` : ''}
    </div>
    ${(pg||spg) ? `<div class="tags-row">
      ${pg  ? `<span class="tag" title="${pg}">${hl(pg)}</span>`  : ''}
      ${spg ? `<span class="tag" title="${spg}">${hl(spg)}</span>` : ''}
    </div>` : ''}
    ${renderMatchHint(r)}
  </div>
</div>`;
}

// ── MATCH HINT ────────────────────────────────────────────────────────────────
function renderMatchHint(r) {
  if (!globalTerms.length) return '';
  const hiddenFields = {
    'Product Group':     r.Product_Group     || '',
    'Sub-Product Group': r.Sub_Product_Group || '',
    'SKU Status':        r.PWC_SKUSTATUS     || '',
    'Vendor Acct':       r.VENDACCOUNT       || '',
    'Location':          r.LOCATION          || '',
  };
  const matched = [];
  for (const [label, val] of Object.entries(hiddenFields)) {
    const lower = val.toLowerCase();
    if (globalTerms.some(t => lower.includes(t))) {
      matched.push(`${label}: <strong>${escHtml(val)}</strong>`);
    }
  }
  if (!matched.length) return '';
  return `<div class="match-hint">🔍 Matched in — ${matched.join(' · ')}</div>`;
}

// ── PAGINATION ────────────────────────────────────────────────────────────────
function renderPagination(total, cur) {
  const pages = Math.ceil(total / PAGE_SIZE);
  const el = document.getElementById('pagination');
  if (pages <= 1) { el.innerHTML = ''; return; }

  let html = `<button class="page-btn" ${cur === 1 ? 'disabled' : ''} onclick="goPage(${cur - 1})">&#8592;</button>`;
  let prev = null;
  for (let i = 1; i <= pages; i++) {
    if (i === 1 || i === pages || (i >= cur - 2 && i <= cur + 2)) {
      if (prev !== null && i - prev > 1) html += `<span class="page-dots">…</span>`;
      html += `<button class="page-btn ${i === cur ? 'active' : ''}" onclick="goPage(${i})">${i}</button>`;
      prev = i;
    }
  }
  html += `<button class="page-btn" ${cur === pages ? 'disabled' : ''} onclick="goPage(${cur + 1})">&#8594;</button>`;
  el.innerHTML = html;
}

window.goPage = function(p) {
  currentPage = p;
  render();
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

// ── CLEAR / RESET ─────────────────────────────────────────────────────────────
// Backward-compat alias — HTML still calls clearFilter for serial/huid/warehouse/location
window.clearFilter = function(type) { clearPasteFilter(type); };

// ── DUPLICATE HUID ────────────────────────────────────────────────────────────
let dupData       = [];   
let dupFiltered   = [];   
const DUP_PAGE_SZ = 200;  

async function loadDupData() {
  if (dupData.length > 0) return;
  try {
    const res  = await fetch(`${API_BASE}/duplicate-huids?page_size=500`);
    const json = await res.json();
    dupData = json.data || [];
    document.getElementById('dupBadgeCount').textContent = dupData.length.toLocaleString();
  } catch (e) {
    console.error('Failed to load duplicate HUIDs', e);
    dupData = [];
    document.getElementById('dupBadgeCount').textContent = '!';
  }
}

window.openDupModal = function() {
  document.getElementById('dupModalOverlay').style.display = 'flex';
  document.getElementById('dupModalSearch').value = '';
  document.body.style.overflow = 'hidden';
  dupFiltered = dupData;
  renderDupTable(dupFiltered);
  if (dupData.length === 0) {
    loadDupData().then(() => {
      dupFiltered = dupData;
      renderDupTable(dupFiltered);
    });
  }
}

window.closeDupModal = function() {
  document.getElementById('dupModalOverlay').style.display = 'none';
  document.body.style.overflow = '';
}

window.closeDupModalOnOverlay = function(e) {
  if (e.target === document.getElementById('dupModalOverlay')) closeDupModal();
}

window.filterDupModal = function() {
  const q = document.getElementById('dupModalSearch').value.trim().toLowerCase();
  dupFiltered = q ? dupData.filter(d => d.huid.toLowerCase().includes(q)) : dupData;
  renderDupTable(dupFiltered);
}

function renderDupTable(rows) {
  const tbody = document.getElementById('dupTableBody');
  const summary = document.getElementById('dupModalSummary');
  const pageInfo = document.getElementById('dupModalPageInfo');

  const showing = Math.min(rows.length, DUP_PAGE_SZ);

  summary.innerHTML = `<strong>${rows.length.toLocaleString()}</strong> duplicate HUID group${rows.length !== 1 ? 's' : ''}`;
  pageInfo.textContent = rows.length > DUP_PAGE_SZ
    ? `Showing first ${DUP_PAGE_SZ} — refine search to narrow down`
    : '';

  if (rows.length === 0) {
    tbody.innerHTML = `<tr><td colspan="4" class="dup-empty">
      <div class="icon">✅</div>
      <div>No duplicates match your search</div>
    </td></tr>`;
    return;
  }

  tbody.innerHTML = rows.slice(0, DUP_PAGE_SZ).map(d => {
    const MAX_PILLS = 5;
    const pills = d.skus.slice(0, MAX_PILLS)
      .map(s => `<span class="sku-pill">${escHtml(s)}</span>`)
      .join('');
    const more = d.skus.length > MAX_PILLS
      ? `<span class="sku-pill-more">+${d.skus.length - MAX_PILLS} more</span>`
      : '';

    return `<tr onclick="selectDupHuid('${escHtml(d.huid)}')">
      <td>
        <div class="dup-huid-val">
          ${escHtml(d.huid)}
          <span class="click-hint">▶ filter</span>
        </div>
      </td>
      <td><span class="dup-count-badge">${d.sku_count}</span></td>
      <td><span class="row-count-val">${d.row_count.toLocaleString()}</span></td>
      <td><div class="sku-pill-wrap">${pills}${more}</div></td>
    </tr>`;
  }).join('');
}

window.selectDupHuid = function(huid) {
  selectedHuids.clear();
  selectedHuids.add(huid);
  updatePasteChips('huid');
  // Reflect in textarea so user sees what's active
  const inputEl = getPasteInputEl('huid');
  if (inputEl) inputEl.value = '';
  currentPage = 1;
  render();
  closeDupModal();
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

function escHtml(str) {
  return String(str)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;')
    .replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

window.downloadDupExcel = function() {
  const btn = document.getElementById('dlBtn');
  if (!dupData || !dupData.length) {
    alert('No duplicate HUID data loaded yet. Please wait a moment and try again.');
    return;
  }

  if (typeof XLSX === 'undefined') {
    alert('Excel library failed to load. Please check your network or adblocker.');
    return;
  }

  const source = (dupFiltered && dupFiltered.length > 0) ? dupFiltered : dupData;

  btn.classList.add('loading');
  btn.textContent = 'Preparing…';

  try {
    const summaryRows = source.map(d => ({
      'HUID':               d.huid || 'Unknown',
      'SKU Count':          d.sku_count || 0,
      'Total Items (Rows)': d.row_count || 0,
      'Mapped Item IDs':    (d.skus || []).join(', '),
    }));

    const expandedRows = [];
    source.forEach(d => {
      (d.skus || []).forEach((sku, idx) => {
        expandedRows.push({
          'HUID':               d.huid || 'Unknown',
          'Item ID (SKU)':      sku || 'Unknown',
          'SKU Index':          idx + 1,
          'Total SKUs on HUID': d.sku_count || 0,
          'Total Items (Rows)': d.row_count || 0,
        });
      });
    });

    const wb = XLSX.utils.book_new();

    const wsSummary = XLSX.utils.json_to_sheet(summaryRows);
    styleSheet(wsSummary);
    XLSX.utils.book_append_sheet(wb, wsSummary, 'Duplicate HUID Summary');

    if (expandedRows.length > 0) {
      const wsExpanded = XLSX.utils.json_to_sheet(expandedRows);
      styleSheet(wsExpanded);
      XLSX.utils.book_append_sheet(wb, wsExpanded, 'HUID-SKU Pairs');
    }

    const today = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, `Duplicate_HUID_Report_${today}.xlsx`);

  } catch (err) {
    console.error('Excel export failed:', err);
    alert(`Export failed: ${err.message}. See console for details.`);
  } finally {
    btn.classList.remove('loading');
    btn.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5">
      <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
      <polyline points="7 10 12 15 17 10"/>
      <line x1="12" y1="15" x2="12" y2="3"/>
    </svg> Download Data`;
  }
}

function styleSheet(ws) {
  if (!ws || !ws['!ref']) return;
  try {
    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let c = range.s.c; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r: 0, c });
      if (ws[addr]) ws[addr].s = { font: { bold: true } };
    }
    const headers = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r: 0, c });
      headers.push(ws[addr] && ws[addr].v !== undefined ? ws[addr].v : '');
    }
    ws['!cols'] = headers.map(h => ({ wch: Math.max(String(h).length + 4, 18) }));
  } catch (err) {
    console.warn("Failed to apply styles, but proceeding with export:", err);
  }
}

document.addEventListener('keydown', e => {
  if (e.key === 'Escape' && document.getElementById('dupModalOverlay').style.display !== 'none') {
    closeDupModal();
  }
});


// ── FULL CATALOGUE DOWNLOAD ───────────────────────────────────────────────────
window.downloadCatalogue = async function() {
  const btn = document.getElementById('dlCatalogueBtn');
  const total = document.getElementById('countTotal')?.textContent?.replace(/,/g,'') || '?';

  btn.disabled = true;
  btn.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" width="15" height="15" style="animation:spin 1s linear infinite"><path d="M21 12a9 9 0 1 1-18 0 9 9 0 0 1 18 0"/></svg> Preparing ${Number(total).toLocaleString()} rows…`;

  try {
    const res = await fetch(`${API_BASE}/download-catalogue`);
    if (!res.ok) throw new Error(`Server error ${res.status}`);

    const blob = await res.blob();
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    // Use filename from Content-Disposition if present, else fallback
    const cd   = res.headers.get('Content-Disposition') || '';
    const match = cd.match(/filename="([^"]+)"/);
    a.download = match ? match[1] : `Indriya_Catalogue_${new Date().toISOString().slice(0,10)}.csv`;
    a.href = url;
    a.click();
    URL.revokeObjectURL(url);
  } catch (err) {
    console.error('Catalogue download failed:', err);
    alert('Download failed: ' + err.message);
  } finally {
    btn.disabled = false;
    btn.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" width="15" height="15"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg> Download Catalogue`;
  }
};

// ── SELECTIVE (FILTERED) DOWNLOAD ─────────────────────────────────────────────
window.downloadSelective = async function() {
  const btn = document.getElementById('dlSelectiveBtn');
  const total = document.getElementById('countTotal')?.textContent?.replace(/,/g, '') || '0';

  btn.disabled = true;
  btn.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" width="15" height="15" style="animation:spin 1s linear infinite"><path d="M21 12a9 9 0 1 1-18 0 9 9 0 0 1 18 0"/></svg> Preparing ${Number(total).toLocaleString()} rows…`;

  try {
    // Build the same filter URL as render() but point to the selective download endpoint
    const url = new URL(`${API_BASE}/download-selective`);
    url.searchParams.append('sort', sortMode);

    const serverQ = buildServerQuery();
    if (serverQ) url.searchParams.append('q', serverQ);

    const serialsParam    = Array.from(selectedSerials).join(',');
    const huidsParam      = Array.from(selectedHuids).join(',');
    const warehousesParam = Array.from(selectedWarehouses).join(',');
    const locationsParam  = Array.from(selectedLocations).join(',');
    const vendorsParam    = Array.from(selectedVendors).join(',');

    if (serialsParam)    url.searchParams.append('serials',    serialsParam);
    if (huidsParam)      url.searchParams.append('huids',      huidsParam);
    if (warehousesParam) url.searchParams.append('warehouses', warehousesParam);
    if (locationsParam)  url.searchParams.append('locations',  locationsParam);
    if (vendorsParam)    url.searchParams.append('vendors',    vendorsParam);

    for (const [key, set] of Object.entries(dropdownSelections)) {
      if (set.size) {
        let paramName = key + 's';
        if (key === 'category')           { paramName = 'categories'; }
        else if (key === 'subcategory')   { paramName = 'subcategories'; }
        else if (key === 'product_group') { paramName = 'product_groups'; }
        else if (key === 'sub_product_group') { paramName = 'sub_prod_grps'; }
        else if (key === 'sku_status')    { paramName = 'sku_statuses'; }
        url.searchParams.append(paramName, Array.from(set).join('|'));
      }
    }

    const res = await fetch(url);
    if (!res.ok) throw new Error(`Server error ${res.status}`);

    const blob = await res.blob();
    const blobUrl = URL.createObjectURL(blob);
    const a = document.createElement('a');
    const cd = res.headers.get('Content-Disposition') || '';
    const match = cd.match(/filename="([^"]+)"/);
    a.download = match ? match[1] : `Indriya_Selective_${new Date().toISOString().slice(0, 10)}.csv`;
    a.href = blobUrl;
    a.click();
    URL.revokeObjectURL(blobUrl);
  } catch (err) {
    console.error('Selective download failed:', err);
    alert('Download failed: ' + err.message);
  } finally {
    btn.disabled = false;
    btn.innerHTML = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" width="15" height="15"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg> Download Selective`;
  }
};

window.resetAll = function() {
  // Clear all paste-input filter sets
  ['serial','huid','warehouse','location','vendor'].forEach(type => {
    clearPasteFilter(type);
  });

  // Clear dropdown selections
  Object.keys(dropdownSelections).forEach(type => {
    dropdownSelections[type].clear();
    updateDropdownChips(type);
    populateDropdown(type);
  });

  sortMode = 'default';
  document.getElementById('globalSearch').value = '';
  globalQuery = '';
  globalTerms = [];
  document.getElementById('searchClear').classList.remove('visible');
  renderSearchTermPills();
  document.getElementById('sortSelect').value = 'default';

  currentPage = 1;
  render();
}
