/* ============================================================= */
/*  Raw Material Requisition – Production Ready (Nov 17, 2025)   */
/*  Handles any sheet name, any column order, any casing        */
/* ============================================================= */

let masterData = [], requisitionRows = [], uploadedFileName = '';
let db = null, sortField = null, sortAsc = true, searchQuery = '';
const DB_NAME = 'RequisitionDB', STORE_NAME = 'data', DB_VERSION = 2;
const itemsPerPage = 9;
let currentPage = 1;

// DOM elements
let fileInput, uploadStatus, clearFileBtn, categorySelect, skuSelect, skuCodeDisplay;
let addBtn, exportBtn, clearBtn, tbody, prevBtn, nextBtn, pageInfo, searchInput;

/* -------------------- Wait for Libraries -------------------- */
function waitForLibs() {
  return new Promise(resolve => {
    const check = () => {
      if (typeof XLSX !== 'undefined' && typeof saveAs !== 'undefined') {
        console.log('%cXLSX & saveAs loaded', 'color:green');
        const status = document.getElementById('libStatus');
        if (status) status.textContent = 'XLSX Check | saveAs Check';
        resolve();
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
  await tx.objectStore(STORE_NAME).put({ id: 1, data: requisitionRows, fileName: uploadedFileName, master: masterData });
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
      resolve();
    };
  });
}

/* -------------------- DOM Ready -------------------- */
document.addEventListener('DOMContentLoaded', async () => {
  await waitForLibs();

  // Cache DOM
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

  await initDB();
  await loadAll();

  // Event Listeners
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

  // Dark Mode – Fixed context
  const darkBtn = document.getElementById('darkModeBtn');
  darkBtn.addEventListener('click', function () {
    document.body.classList.toggle('dark-mode');
    const isDark = document.body.classList.contains('dark-mode');
    localStorage.setItem('darkMode', isDark);
    this.innerHTML = isDark ? '<i class="fas fa-sun"></i>' : '<i class="fas fa-moon"></i>';
  });

  if (localStorage.getItem('darkMode') === 'true') {
    document.body.classList.add('dark-mode');
    darkBtn.innerHTML = '<i class="fas fa-sun"></i>';
  }

  document.getElementById('printBtn').addEventListener('click', () => window.print());
});

/* -------------------- File Upload (Robust) -------------------- */
async function handleFileUpload(e) {
  const file = e.target.files[0];
  if (!file) return;

  uploadStatus.textContent = 'Reading file...';
  uploadStatus.style.color = '';

  const reader = new FileReader();
  reader.onload = async ev => {
    try {
      const data = new Uint8Array(ev.target.result);
      const workbook = XLSX.read(data, { type: 'array' });

      const sheetNames = workbook.SheetNames;
      if (!sheetNames.length) throw new Error('No sheets found in file.');

      const sheet = workbook.Sheets[sheetNames[0]];
      console.log('%cUsing sheet: ' + sheetNames[0], 'color:cyan');

      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
      if (rows.length < 2) throw new Error('File has no data rows.');

      const headerRow = rows[0].map(h => h.toString().trim().toLowerCase());
      console.log('%cHeaders detected:', 'color:orange', headerRow);

      // Map column indexes
      const col = { category: -1, skuCode: -1, skuName: -1, raw: -1, qty: -1, unit: -1 };
      headerRow.forEach((h, i) => {
        if (h.includes('category')) col.category = i;
        if (h.includes('sku') && h.includes('code')) col.skuCode = i;
        if (h.includes('sku') && !h.includes('code')) col.skuName = i;
        if (h.includes('raw') && h.includes('material')) col.raw = i;
        if (/qty|quantity/.test(h)) col.qty = i;
        if (h.includes('unit')) col.unit = i;
      });

      if (col.category === -1 || col.skuCode === -1 || col.skuName === -1) {
        throw new Error(`Required columns missing. Need: CATEGORY, SKU, SKU CODE.\nFound: ${headerRow.join(', ')}`);
      }

      masterData = rows.slice(1)
        .map(r => ({
          'CATEGORY': (r[col.category] || '').toString().trim(),
          'SKU CODE': (r[col.skuCode] || '').toString().trim(),
          'SKU': (r[col.skuName] || '').toString().trim(),
          'RAW MATERIAL': (r[col.raw] || '').toString().trim(),
          'QUANTITY/BATCH': (r[col.qty] || '').toString().trim(),
          'UNIT4': (r[col.unit] || '').toString().trim()
        }))
        .filter(r => r['CATEGORY'] && r['SKU CODE'] && r['SKU']);

      if (masterData.length === 0) throw new Error('No valid data rows found after filtering.');

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
      masterData = [];
      categorySelect.disabled = true;
    }
  };

  reader.onerror = () => {
    uploadStatus.textContent = 'Failed to read file.';
    uploadStatus.style.color = 'red';
  };

  reader.readAsArrayBuffer(file);
}

/* -------------------- Category & SKU -------------------- */
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

/* -------------------- Add SKU -------------------- */
async function handleAddSku() {
  const code = skuSelect.value;
  const name = skuSelect.selectedOptions[0].text;
  const cat = categorySelect.value;

  const mats = masterData
    .filter(r => r['SKU CODE'] === code && r['SKU'] === name && r['RAW MATERIAL'])
    .map(r => ({ name: r['RAW MATERIAL'], qty: r['QUANTITY/BATCH'], unit: r['UNIT4'] }));

  if (!mats.length) return alert('No raw materials found for this SKU.');

  requisitionRows.push({ skuCode: code, skuName: name, category: cat, qtyNeeded: 1, supplier: '', materials: mats });
  currentPage = Math.ceil(requisitionRows.length / itemsPerPage);
  await saveAll();
  renderPage();
  skuSelect.value = ''; skuCodeDisplay.value = ''; addBtn.disabled = true;
}

/* -------------------- Render Page -------------------- */
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
    tbody.innerHTML = `<tr><td colspan="8" style="text-align:center;padding:20px;color:#888;">
      ${searchQuery ? 'No results.' : 'No SKUs added.'}</td></tr>`;
  }

  pageItems.forEach((item, i) => {
    const idx = start + i, rowId = `d-${idx}`;
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><button class="toggle-btn" data-target="${rowId}"><i class="fa-solid fa-chevron-down"></i></button></td>
      <td>${item.skuCode}</td><td>${item.skuName}</td><td>${item.category}</td>
      <td><input type="number" class="qty-input" min="1" max="99" value="${item.qtyNeeded}" data-idx="${idx}"></td>
      <td><input type="text" class="supplier-input" placeholder="Supplier" value="${item.supplier}" data-idx="${idx}"></td>
      <td>${item.materials.length}</td>
      <td><button class="remove-btn" data-idx="${idx}"><i class="fas fa-trash-alt"></i></button></td>
    `;
    tbody.appendChild(tr);

    const det = document.createElement('tr'); det.id = rowId;
    det.innerHTML = `<td colspan="8" style="padding:0;"><div class="collapse-anim-wrapper"><div class="collapse-content"><table class="inner-table"><thead><tr><th>Raw Material</th><th>Qty/Batch</th><th>Unit</th><th>Total Req</th></tr></thead><tbody>
      ${item.materials.map(m => {
        const total = (parseFloat(m.qty) || 0) * item.qtyNeeded;
        return `<tr><td><strong>${m.name}</strong></td><td>${m.qty}</td><td>${m.unit}</td><td>${total} ${m.unit}</td></tr>`;
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
  setupCollapse(); setupInputs(); setupRemove();
  setupPagination(items.length);
}

/* -------------------- Sort Icons -------------------- */
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

/* -------------------- Collapse Animation -------------------- */
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

/* -------------------- Input Handlers -------------------- */
function setupInputs() {
  tbody.addEventListener('change', async e => {
    if (e.target.classList.contains('qty-input')) {
      const idx = +e.target.dataset.idx;
      requisitionRows[idx].qtyNeeded = Math.max(1, Math.min(99, +e.target.value || 1));
      await saveAll(); renderPage();
    }
    if (e.target.classList.contains('supplier-input')) {
      requisitionRows[+e.target.dataset.idx].supplier = e.target.value.trim();
      await saveAll();
    }
  });
}

/* -------------------- Remove Row -------------------- */
function setupRemove() {
  tbody.addEventListener('click', async e => {
    if (e.target.closest('.remove-btn')) {
      const idx = +e.target.closest('.remove-btn').dataset.idx;
      if (confirm(`Remove ${requisitionRows[idx].skuName}?`)) {
        requisitionRows.splice(idx, 1);
        await saveAll(); renderPage();
      }
    }
  });
}

/* -------------------- Pagination -------------------- */
function setupPagination(total) {
  const pages = Math.ceil(total / itemsPerPage) || 1;
  pageInfo.textContent = `${currentPage} / ${pages}`;
  prevBtn.disabled = currentPage === 1;
  nextBtn.disabled = currentPage === pages;
  exportBtn.disabled = requisitionRows.length === 0;
}

/* -------------------- Clear All -------------------- */
async function clearAll() {
  if (!confirm('Clear all data and uploaded file?')) return;
  requisitionRows = []; masterData = []; uploadedFileName = '';
  sortField = null; searchQuery = ''; searchInput.value = ''; currentPage = 1;
  await saveAll();
  uploadStatus.textContent = 'Upload MasterTable.xlsx'; clearFileBtn.style.display = 'none';
  categorySelect.innerHTML = '<option value="">-- Category --</option>'; categorySelect.disabled = true;
  skuSelect.innerHTML = '<option value="">-- SKU --</option>'; skuSelect.disabled = true;
  renderPage();
}

/* -------------------- Export All -------------------- */
function handleExportAll() {
  if (typeof XLSX === 'undefined' || typeof saveAs === 'undefined') {
    alert("Export library not loaded. Please refresh.");
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
    ['SKU Code', 'SKU', 'Category', 'Qty Needed', 'Supplier', 'Raw Material', 'Qty/Batch', 'Unit', 'Total Req']
  ];

  items.forEach(i => {
    i.materials.forEach(m => {
      const totalQty = (parseFloat(m.qty) || 0) * i.qtyNeeded;
      const total = totalQty + (m.unit ? ' ' + m.unit : '');
      data.push([
        i.skuCode, i.skuName, i.category, i.qtyNeeded, i.supplier,
        m.name, m.qty, m.unit, total
      ]);
    });
  });

  try {
    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Requisition');
    ws['!cols'] = [
      { wch: 12 }, { wch: 30 }, { wch: 15 }, { wch: 10 },
      { wch: 20 }, { wch: 35 }, { wch: 12 }, { wch: 8 }, { wch: 15 }
    ];
    const fileName = `Requisition_${new Date().toISOString().slice(0,10).replace(/-/g,'')}.xlsx`;
    XLSX.writeFile(wb, fileName);
  } catch (err) {
    console.error(err);
    alert("Export failed: " + err.message);
  }
}