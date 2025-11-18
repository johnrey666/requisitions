/* ============================================================= */
/*  Raw Material Requisition – HYBRID CLOUD SYNC                */
/*  FINAL VERSION: Unit Fixed + "Type" Column Added             */
/*  Qty/Unit → Unit → Qty/Pack → Pack Unit → Type (Perfect!)    */
/* ============================================================= */

let masterData = [], requisitionRows = [], uploadedFileName = '';
let db = null, sortField = null, sortAsc = true, searchQuery = '';
const DB_NAME = 'RequisitionDB', STORE_NAME = 'data', DB_VERSION = 3;
const itemsPerPage = 8;
let currentPage = 1;

const SHEETS_WEB_APP_URL = "https://script.google.com/macros/s/AKfycbxDh5b_lKJyDv2nAQT_my-V7Syhm81RWIcRQqn7-hgkgZeMM7gT-mup4JJ3OcKZQcYJQg/exec";

let isSyncEnabled = false;
let lastSyncTime = null;
let syncInProgress = false;

// DOM elements
let fileInput, uploadStatus, clearFileBtn, categorySelect, skuSelect, skuCodeDisplay;
let addBtn, exportBtn, clearBtn, tbody, prevBtn, nextBtn, pageInfo, searchInput;
let syncBtn, syncStatus, configBtn;

/* -------------------- Wait for Libraries -------------------- */
function waitForLibs() {
  return new Promise(resolve => {
    let attempts = 0;
    const maxAttempts = 100;
    const check = () => {
      attempts++;
      const xlsxLoaded = typeof XLSX !== 'undefined';
      const saveAsLoaded = typeof saveAs !== 'undefined';
      if (xlsxLoaded && saveAsLoaded) {
        console.log('%cXLSX & saveAs loaded!', 'color:green; font-weight:bold');
        const status = document.getElementById('libStatus');
        if (status) {
          status.textContent = 'Ready';
          status.style.background = '#d4edda';
          status.style.color = '#155724';
        }
        resolve(true);
      } else if (attempts >= maxAttempts) {
        alert('CRITICAL: XLSX library failed to load.\nPlease refresh the page.');
        resolve(false);
      } else {
        setTimeout(check, 50);
      }
    };
    check();
  });
}

/* -------------------- IndexedDB -------------------- */
function initDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = e => {
      db = e.target.result;
      if (db.objectStoreNames.contains(STORE_NAME)) db.deleteObjectStore(STORE_NAME);
      db.createObjectStore(STORE_NAME, { keyPath: 'id' });
    };
    req.onsuccess = e => { db = e.target.result; resolve(); };
    req.onerror = e => reject(e.target.error);
  });
}

async function saveAll() {
  if (!db) return;
  const tx = db.transaction(STORE_NAME, 'readwrite');
  await tx.objectStore(STORE_NAME).put({ 
    id: 1, 
    data: requisitionRows, 
    fileName: uploadedFileName, 
    master: masterData,
    lastModified: new Date().toISOString()
  });
  
  if (isSyncEnabled) await syncToCloudSilent();
  return tx.done;
}

async function loadAll() {
  if (!db) return;
  const req = db.transaction(STORE_NAME).objectStore(STORE_NAME).get(1);
  return new Promise(resolve => {
    req.onsuccess = () => {
      if (req.result) {
        requisitionRows = req.result.data || [];
        uploadedFileName = req.result.fileName || '';
        masterData = req.result.master || [];
        if (uploadedFileName) {
          uploadStatus.textContent = `Using: ${uploadedFileName}`;
          clearFileBtn.style.display = 'inline-block';
        }
        if (masterData.length) populateCategories();
      }
      renderPage();
      loadSyncConfig();
      resolve();
    };
  });
}

/* -------------------- Google Sheets Sync (100% ORIGINAL) -------------------- */
function loadSyncConfig() {
  isSyncEnabled = localStorage.getItem('syncEnabled') === 'true';
  lastSyncTime = localStorage.getItem('lastSyncTime');
  updateSyncUI();
}

function updateSyncUI() {
  if (!syncBtn || !syncStatus) return;
  
  if (isSyncEnabled) {
    syncBtn.innerHTML = '<i class="fas fa-cloud-upload-alt"></i> Synced';
    syncBtn.style.background = '#28a745';
    syncBtn.title = 'Click to disable Google Sheets sync';
    
    if (lastSyncTime) {
      const date = new Date(lastSyncTime);
      syncStatus.textContent = `Auto-sync ON • Last: ${date.toLocaleTimeString()}`;
      syncStatus.style.color = '#28a745';
    } else {
      syncStatus.textContent = 'Auto-sync ON';
      syncStatus.style.color = '#28a745';
    }
  } else {
    syncBtn.innerHTML = '<i class="fas fa-cloud-slash"></i> Offline';
    syncBtn.style.background = '#6c757d';
    syncBtn.title = 'Click to enable Google Sheets sync';
    syncStatus.textContent = 'Local only';
    syncStatus.style.color = '#6c757d';
  }
}

async function toggleSync() {
  isSyncEnabled = !isSyncEnabled;
  localStorage.setItem('syncEnabled', isSyncEnabled);
  updateSyncUI();

  if (isSyncEnabled) {
    syncStatus.textContent = 'Syncing...';
    try {
      await syncToCloud();
      alert('Google Sheets sync ENABLED!\n\nAll changes now auto-backup to your sheet.');
    } catch (e) {
      alert('Sync enabled, but first backup failed. Will retry automatically.');
    }
  } else {
    alert('Google Sheets sync disabled.\nData stays in browser only.');
  }
}

async function syncToCloudSilent() {
  if (!isSyncEnabled || syncInProgress) return;
  syncInProgress = true;
  try {
    const payload = {
      requisitionRows,
      masterData,
      uploadedFileName,
      lastModified: new Date().toISOString(),
      device: navigator.userAgent.substring(0, 80)
    };

    await fetch(SHEETS_WEB_APP_URL, {
      method: 'POST',
      body: JSON.stringify(payload),
      headers: { 'Content-Type': 'text/plain' }    });

    lastSyncTime = new Date().toISOString();
    localStorage.setItem('lastSyncTime', lastSyncTime);
    updateSyncUI();
  } catch (err) {
    console.warn('Silent sync failed:', err);
    if (syncStatus) {
      syncStatus.textContent = 'Offline - saved locally';
      syncStatus.style.color = '#ffc107';
    }
  } finally {
    syncInProgress = false;
  }
}

async function syncToCloud() {
  if (syncInProgress) return;
  syncInProgress = true;
  syncStatus.textContent = 'Syncing...';
  syncStatus.style.color = '#ffc107';

  try {
    const payload = {
      requisitionRows,
      masterData,
      uploadedFileName,
      lastModified: new Date().toISOString()
    };

    const response = await fetch(SHEETS_WEB_APP_URL, {
      method: 'POST',
      body: JSON.stringify(payload),
      headers: { 'Content-Type': 'text/plain;charset=utf-8' }
    });

    if (!response.ok) throw new Error('Network error');

    lastSyncTime = new Date().toISOString();
    localStorage.setItem('lastSyncTime', lastSyncTime);
    updateSyncUI();
  } catch (err) {
    syncStatus.textContent = 'Sync failed';
    syncStatus.style.color = '#dc3545';
    throw err;
  } finally {
    syncInProgress = false;
  }
}

async function restoreFromCloud() {
  if (!confirm('Restore from Google Sheets?\n\nThis will replace all local data.')) return;

  syncStatus.textContent = 'Restoring...';
  syncStatus.style.color = '#ffc107';

  try {
    const response = await fetch(SHEETS_WEB_APP_URL);
    if (!response.ok) throw new Error('Failed to reach sheet');

    const text = await response.text();
    if (!text || text.includes('error')) {
      alert('No backup found in Google Sheets yet.\nMake a change with sync ON first.');
      return;
    }

    const data = JSON.parse(text);
    requisitionRows = data.requisitionRows || [];
    masterData = data.masterData || [];
    uploadedFileName = data.uploadedFileName || '';

    await saveAll();
    if (uploadedFileName) {
      uploadStatus.textContent = `Using: ${uploadedFileName}`;
      clearFileBtn.style.display = 'inline-block';
    }
    if (masterData.length) populateCategories();
    renderPage();

    alert('Successfully restored from Google Sheets!');
    updateSyncUI();
  } catch (err) {
    alert('Restore failed: ' + err.message);
    syncStatus.textContent = 'Restore failed';
    syncStatus.style.color = '#dc3545';
  }
}

/* -------------------- DOM Ready -------------------- */
document.addEventListener('DOMContentLoaded', async () => {
  await waitForLibs();
  fileInput = document.getElementById('masterFile');
  uploadStatus = document.getElementById('uploadStatus');
  clearFileBtn = document.getElementById('clearFileBtn');
  categorySelect = document.getElementById('categorySelect');
  skuSelect = document.getElementById('skuSelect');
  skuCodeDisplay = document.getElementById('skuCodeDisplay');
  addBtn = document.getElementById('addBtn');
  exportBtn = document.getElementById('exportBtn');
  clearBtn = document.getElementById('clearBtn');
  tbody = document.getElementById('reqBody');
  prevBtn = document.getElementById('prevBtn');
  nextBtn = document.getElementById('nextBtn');
  pageInfo = document.getElementById('pageInfo');
  searchInput = document.getElementById('searchInput');
  syncBtn = document.getElementById('syncBtn');
  syncStatus = document.getElementById('syncStatus');
  configBtn = document.getElementById('configBtn');

  await initDB();
  await loadAll();

  fileInput.addEventListener('change', handleFileUpload);
  clearFileBtn.addEventListener('click', clearAll);
  categorySelect.addEventListener('change', handleCategoryChange);
  skuSelect.addEventListener('change', handleSkuChange);
  addBtn.addEventListener('click', handleAddSku);
  prevBtn.addEventListener('click', () => { currentPage = Math.max(1, currentPage - 1); renderPage(); });
  nextBtn.addEventListener('click', () => { currentPage++; renderPage(); });
  exportBtn.addEventListener('click', handleExportAll);
  clearBtn.addEventListener('click', clearAll);
  searchInput.addEventListener('input', () => {
    searchQuery = searchInput.value.trim().toLowerCase();
    currentPage = 1;
    renderPage();
  });

  syncBtn.addEventListener('click', toggleSync);
  syncBtn.addEventListener('contextmenu', e => { e.preventDefault(); restoreFromCloud(); });

  configBtn.addEventListener('click', () => {
    alert(`GOOGLE SHEETS BACKUP\n━━━━━━━━━━━━━━━━━━━━━━━━━━\nWeb App URL:\n${SHEETS_WEB_APP_URL}\n\nStatus: ${isSyncEnabled ? 'AUTO-SYNC ENABLED' : 'OFFLINE MODE'}\n${lastSyncTime ? 'Last sync: ' + new Date(lastSyncTime).toLocaleString() : ''}\n\nClick the cloud button to toggle sync.\nRight-click it to restore from sheet.`);
  });

  const darkBtn = document.getElementById('darkModeBtn');
  darkBtn.addEventListener('click', () => {
    document.body.classList.toggle('dark-mode');
    const isDark = document.body.classList.contains('dark-mode');
    localStorage.setItem('darkMode', isDark);
    darkBtn.innerHTML = isDark ? '<i class="fas fa-sun"></i>' : '<i class="fas fa-moon"></i>';
  });
  if (localStorage.getItem('darkMode') === 'true') {
    document.body.classList.add('dark-mode');
    darkBtn.innerHTML = '<i class="fas fa-sun"></i>';
  }

  document.getElementById('printBtn').addEventListener('click', () => window.print());
});

/* -------------------- File Upload – NOW WITH "TYPE" COLUMN -------------------- */
async function handleFileUpload(e) {
  const file = e.target.files[0];
  if (!file) return;

  if (typeof XLSX === 'undefined') {
    uploadStatus.textContent = 'ERROR: XLSX library not loaded';
    uploadStatus.style.color = 'red';
    alert('Please refresh the page and try again.');
    return;
  }

  uploadStatus.textContent = 'Reading file...';
  uploadStatus.style.color = '';

  const reader = new FileReader();
  reader.onload = async ev => {
    try {
      const data = new Uint8Array(ev.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
      if (rows.length < 2) throw new Error('File has no data rows.');

      const headerRow = rows[0].map(h => h.toString().trim().toLowerCase());

      const col = { 
        category: headerRow.findIndex(h => h.includes('category')),
        skuCode: headerRow.findIndex(h => h.includes('sku') && h.includes('code')),
        skuName: headerRow.findIndex(h => h.includes('sku') && !h.includes('code') && !h.includes('quantity')),
        qtyPerUnit: headerRow.findIndex(h => h.includes('quantity') && h.includes('per') && h.includes('unit') && !h.includes('pack')),
        unit: headerRow.findIndex(h => h === 'unit' && !h.includes('2') && !h.includes('4') && !h.includes('pack') && !h.includes('batch')),
        qtyPerPack: headerRow.findIndex(h => h.includes('quantity') && h.includes('per') && h.includes('pack')),
        unit2: headerRow.findIndex(h => h === 'unit2' || (h.includes('unit') && headerRow[headerRow.indexOf(h)-1]?.includes('pack'))),
        raw: headerRow.findIndex(h => h.includes('raw') && h.includes('material')),
        qtyBatch: headerRow.findIndex(h => h.includes('quantity') && h.includes('batch')),
        unit4: headerRow.findIndex(h => h.includes('unit') && (h.includes('4') || headerRow[headerRow.indexOf(h)-1]?.includes('batch'))),
        type: headerRow.findIndex(h => h === 'type' || h.includes('type'))
      };

      masterData = rows.slice(1)
        .map(r => ({
          'CATEGORY': (r[col.category] || '').toString().trim(),
          'SKU CODE': (r[col.skuCode] || '').toString().trim(),
          'SKU': (r[col.skuName] || '').toString().trim(),
          'QUANTITY PER UNIT': (r[col.qtyPerUnit] || '').toString().trim(),
          'UNIT': (r[col.unit] || '').toString().trim(),
          'QUANTITY PER PACK': (r[col.qtyPerPack] || '').toString().trim(),
          'UNIT2': (r[col.unit2] || '').toString().trim(),
          'RAW MATERIAL': (r[col.raw] || '').toString().trim(),
          'QUANTITY/BATCH': (r[col.qtyBatch] || '').toString().trim(),
          'UNIT4': (r[col.unit4] || '').toString().trim(),
          'TYPE': (r[col.type] || '').toString().trim()
        }))
        .filter(r => r['CATEGORY'] && r['SKU CODE'] && r['SKU']);

      if (masterData.length === 0) throw new Error('No valid data rows found.');

      uploadedFileName = file.name;
      uploadStatus.textContent = `Loaded: ${uploadedFileName} (${masterData.length} items)`;
      uploadStatus.style.color = 'green';
      clearFileBtn.style.display = 'inline-block';
      populateCategories();
      await saveAll();

    } catch (err) {
      console.error('Upload error:', err);
      uploadStatus.textContent = 'ERROR: ' + err.message;
      uploadStatus.style.color = 'red';
    }
  };
  reader.onerror = () => {
    uploadStatus.textContent = 'Failed to read file.';
    uploadStatus.style.color = 'red';
  };
  reader.readAsArrayBuffer(file);
}

function populateCategories() {
  const cats = [...new Set(masterData.map(r => r['CATEGORY']).filter(Boolean))].sort();
  categorySelect.innerHTML = '<option value="">-- Category --</option>';
  cats.forEach(c => categorySelect.add(new Option(c, c)));
  categorySelect.disabled = cats.length === 0;
}

function handleCategoryChange() {
  const cat = categorySelect.value;
  skuSelect.innerHTML = '<option value="">-- SKU --</option>';
  skuSelect.disabled = true;
  skuCodeDisplay.value = '';
  addBtn.disabled = true;
  if (!cat) return;

  const map = new Map();
  masterData
    .filter(r => r['CATEGORY'] === cat && r['SKU'] && r['SKU CODE'])
    .forEach(r => {
      const sku = r['SKU'].trim();
      const code = r['SKU CODE'].trim();
      if (sku && code && !map.has(sku)) map.set(sku, code);
    });

  Array.from(map).sort((a, b) => a[0].localeCompare(b[0]))
    .forEach(([name, code]) => skuSelect.add(new Option(name, code)));

  skuSelect.disabled = false;
}

function handleSkuChange() {
  skuCodeDisplay.value = skuSelect.value;
  addBtn.disabled = !skuSelect.value;
}

async function handleAddSku() {
  const code = skuSelect.value;
  const name = skuSelect.selectedOptions[0].text;
  const cat = categorySelect.value;

  const skuInfo = masterData.find(r => r['SKU CODE'] === code && r['SKU'] === name);
  if (!skuInfo) return alert('SKU not found in master data.');

  const mats = masterData
    .filter(r => r['SKU CODE'] === code && r['SKU'] === name && r['RAW MATERIAL'])
    .map(r => ({ 
      name: r['RAW MATERIAL'], 
      qty: r['QUANTITY/BATCH'], 
      unit: r['UNIT4'],
      type: r['TYPE'] || ''
    }));

  if (!mats.length) return alert('No raw materials found for this SKU.');

  requisitionRows.push({ 
    skuCode: code, 
    skuName: name, 
    category: cat, 
    qtyNeeded: 1, 
    supplier: '', 
    qtyPerUnit: skuInfo['QUANTITY PER UNIT'] || '',
    unit: skuInfo['UNIT'] || '',
    qtyPerPack: skuInfo['QUANTITY PER PACK'] || '',
    unit2: skuInfo['UNIT2'] || '',
    materials: mats 
  });

  currentPage = Math.ceil(requisitionRows.length / itemsPerPage);
  await saveAll();
  renderPage();
  skuSelect.value = ''; 
  skuCodeDisplay.value = ''; 
  addBtn.disabled = true;
}

function renderPage() {
  tbody.innerHTML = '';
  let items = [...requisitionRows];

  if (searchQuery) {
    items = items.filter(i =>
      i.skuCode.toLowerCase().includes(searchQuery) ||
      i.skuName.toLowerCase().includes(searchQuery) ||
      i.category.toLowerCase().includes(searchQuery)
    );
  }

  if (sortField) {
    items.sort((a, b) => {
      const A = (a[sortField] || '').toString().toLowerCase();
      const B = (b[sortField] || '').toString().toLowerCase();
      return (A < B ? -1 : A > B ? 1 : 0) * (sortAsc ? 1 : -1);
    });
  }

  const start = (currentPage - 1) * itemsPerPage;
  const end = start + itemsPerPage;
  const pageItems = items.slice(start, end);

  if (!pageItems.length) {
    tbody.innerHTML = `<tr><td colspan="12" style="text-align:center;padding:20px;color:#888;">
      ${searchQuery ? 'No results found.' : 'No SKUs added yet.'}</td></tr>`;
  }

  pageItems.forEach((item, i) => {
    const idx = start + i, rowId = `d-${idx}`;
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><button class="toggle-btn" data-target="${rowId}"><i class="fa-solid fa-chevron-down"></i></button></td>
      <td>${item.skuCode}</td>
      <td>${item.skuName}</td>
      <td>${item.category}</td>
      <td><input type="number" class="qty-input" min="1" max="99" value="${item.qtyNeeded}" data-idx="${idx}"></td>
      <td><input type="text" class="supplier-input" placeholder="Supplier" value="${item.supplier}" data-idx="${idx}"></td>
      <td style="text-align:center;font-weight:600;color:#1976d2;">${item.qtyPerUnit || '-'}</td>
      <td style="text-align:center;">${item.unit || '-'}</td>
      <td style="text-align:center;font-weight:600;color:#d32f2f;">${item.qtyPerPack || '-'}</td>
      <td style="text-align:center;">${item.unit2 || '-'}</td>
      <td>${item.materials.length}</td>
      <td><button class="remove-btn" data-idx="${idx}"><i class="fas fa-trash-alt"></i></button></td>
    `;
    tbody.appendChild(tr);

    const det = document.createElement('tr'); 
    det.id = rowId;
    det.innerHTML = `<td colspan="12" style="padding:0;"><div class="collapse-anim-wrapper"><div class="collapse-content"><table class="inner-table"><thead><tr><th>Raw Material</th><th>Qty/Batch</th><th>Unit</th><th>Type</th><th>Total Req</th></tr></thead><tbody>
      ${item.materials.map(m => {
        const total = (parseFloat(m.qty) || 0) * item.qtyNeeded;
        return `<tr><td><strong>${m.name}</strong></td><td>${m.qty}</td><td>${m.unit}</td><td>${m.type || '-'}</td><td>${total} ${m.unit}</td></tr>`;
      }).join('')}
    </tbody></table></div></div></td>`;
    tbody.appendChild(det);
  });

  document.querySelectorAll('th.sortable').forEach(th => {
    th.onclick = () => {
      const f = th.dataset.sort;
      if (sortField === f) sortAsc = !sortAsc;
      else { sortField = f; sortAsc = true; }
      currentPage = 1; updateSortIcons(); renderPage();
    };
  });

  updateSortIcons();
  setupCollapse(); 
  setupInputs(); 
  setupRemove();
  setupPagination(items.length);
}

function updateSortIcons() {
  document.querySelectorAll('th.sortable').forEach(th => {
    th.classList.remove('sort-asc', 'sort-desc');
    const icon = th.querySelector('.sort-icon');
    if (th.dataset.sort === sortField) {
      th.classList.add(sortAsc ? 'sort-asc' : 'sort-desc');
      icon.className = sortAsc ? 'fas fa-sort-up sort-icon' : 'fas fa-sort-down sort-icon';
    } else {
      icon.className = 'fas fa-sort sort-icon';
    }
  });
}

function setupCollapse() {
  document.querySelectorAll('.toggle-btn').forEach(b => {
    b.onclick = () => {
      const r = document.getElementById(b.dataset.target);
      const w = r.querySelector('.collapse-anim-wrapper');
      const i = b.querySelector('i');
      if (w.classList.contains('open')) {
        w.style.height = w.scrollHeight + 'px';
        requestAnimationFrame(() => w.style.height = '0px');
        w.classList.remove('open');
        i.classList.replace('fa-chevron-up', 'fa-chevron-down');
      } else {
        w.classList.add('open');
        w.style.height = w.scrollHeight + 'px';
        i.classList.replace('fa-chevron-down', 'fa-chevron-up');
        w.addEventListener('transitionend', () => {
          if (w.classList.contains('open')) w.style.height = 'auto';
        }, { once: true });
      }
    };
  });
}

function setupInputs() {
  tbody.addEventListener('change', async e => {
    if (e.target.classList.contains('qty-input')) {
      const idx = +e.target.dataset.idx;
      requisitionRows[idx].qtyNeeded = Math.max(1, Math.min(99, +e.target.value || 1));
      await saveAll(); 
      renderPage();
    }
    if (e.target.classList.contains('supplier-input')) {
      requisitionRows[+e.target.dataset.idx].supplier = e.target.value.trim();
      await saveAll();
    }
  });
}

function setupRemove() {
  tbody.addEventListener('click', async e => {
    if (e.target.closest('.remove-btn')) {
      const idx = +e.target.closest('.remove-btn').dataset.idx;
      if (confirm(`Remove ${requisitionRows[idx].skuName}?`)) {
        requisitionRows.splice(idx, 1);
        await saveAll(); 
        renderPage();
      }
    }
  });
}

function setupPagination(total) {
  const pages = Math.ceil(total / itemsPerPage) || 1;
  pageInfo.textContent = `${currentPage} / ${pages}`;
  prevBtn.disabled = currentPage === 1;
  nextBtn.disabled = currentPage === pages;
  exportBtn.disabled = requisitionRows.length === 0;
}

async function clearAll() {
  if (!confirm('Clear all data and uploaded file?')) return;
  requisitionRows = []; masterData = []; uploadedFileName = '';
  sortField = null; searchQuery = ''; searchInput.value = ''; currentPage = 1;
  await saveAll();
  uploadStatus.textContent = 'Upload MasterTable.xlsx'; 
  uploadStatus.style.color = '';
  clearFileBtn.style.display = 'none';
  categorySelect.innerHTML = '<option value="">-- Category --</option>'; 
  categorySelect.disabled = true;
  skuSelect.innerHTML = '<option value="">-- SKU --</option>'; 
  skuSelect.disabled = true;
  skuCodeDisplay.value = '';
  addBtn.disabled = true;
  renderPage();
}

function handleExportAll() {
  if (typeof XLSX === 'undefined' || typeof saveAs === 'undefined') {
    alert("Export libraries not loaded. Please refresh.");
    return;
  }

  let items = [...requisitionRows];
  if (searchQuery) {
    items = items.filter(i =>
      i.skuCode.toLowerCase().includes(searchQuery) ||
      i.skuName.toLowerCase().includes(searchQuery) ||
      i.category.toLowerCase().includes(searchQuery)
    );
  }
  if (sortField) {
    items.sort((a, b) => {
      const A = (a[sortField] || '').toString().toLowerCase();
      const B = (b[sortField] || '').toString().toLowerCase();
      return (A < B ? -1 : A > B ? 1 : 0) * (sortAsc ? 1 : -1);
    });
  }

  if (items.length === 0) {
    alert("No data to export.");
    return;
  }

  const data = [
    ['RAW MATERIAL REQUISITION'],
    ['Generated', new Date().toLocaleString('en-PH')],
    ['File', uploadedFileName || 'None'],
    [], 
    ['SKU Code', 'SKU', 'Category', 'Qty Needed', 'Supplier', 'Qty/Unit', 'Unit', 'Qty/Pack', 'Pack Unit', 'Raw Material', 'Qty/Batch', 'Unit', 'Type', 'Total Req']
  ];

  items.forEach(i => {
    i.materials.forEach(m => {
      const totalQty = (parseFloat(m.qty) || 0) * i.qtyNeeded;
      const total = totalQty + (m.unit ? ' ' + m.unit : '');
      data.push([
        i.skuCode, i.skuName, i.category, i.qtyNeeded, i.supplier,
        i.qtyPerUnit || '', i.unit || '', i.qtyPerPack || '', i.unit2 || '',
        m.name, m.qty, m.unit, m.type || '', total
      ]);
    });
  });

  const ws = XLSX.utils.aoa_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Requisition');
  ws['!cols'] = [
    { wch: 12 }, { wch: 30 }, { wch: 15 }, { wch: 10 },
    { wch: 20 }, { wch: 10 }, { wch: 8 }, { wch: 10 }, { wch: 10 },
    { wch: 35 }, { wch: 12 }, { wch: 8 }, { wch: 12 }, { wch: 15 }
  ];
  const fileName = `Requisition_${new Date().toISOString().slice(0,10).replace(/-/g,'')}.xlsx`;
  XLSX.writeFile(wb, fileName);
}