/**
 * LIFEINNO Daraz Helper - Popup Script
 */

let allProducts = [];
let activeProduct = null;

document.addEventListener('DOMContentLoaded', () => {
  // Init UI Elements
  const fileInput = document.getElementById('fileInput');
  const searchBar = document.getElementById('search-bar');
  const btnClearSearch = document.getElementById('btn-clear-search');
  const btnReloadFile = document.getElementById('btn-reload-file');
  const btnBackToList = document.getElementById('btn-back-to-list');
  const btnSidebar = document.getElementById('btn-sidebar');
  const uploadStatus = document.getElementById('upload-status');

  // Load from Storage on open
  chrome.storage.local.get(['lifeinno_products', 'lifeinno_filename'], (data) => {
    if (data.lifeinno_products && data.lifeinno_products.length > 0) {
      allProducts = data.lifeinno_products;
      const filename = data.lifeinno_filename || 'excel_file.xlsx';
      showLoadedState(allProducts.length, filename);
    } else {
      showView('view-upload');
    }
  });

  // Handle file uploads
  if (fileInput) {
    fileInput.addEventListener('change', handleFileUpload);
  }

  // Search input listeners
  if (searchBar) {
    searchBar.addEventListener('input', (e) => {
      const query = e.target.value;
      if (btnClearSearch) {
        btnClearSearch.style.display = query ? 'block' : 'none';
      }
      renderProductsList(query);
    });
  }

  if (btnClearSearch) {
    btnClearSearch.addEventListener('click', () => {
      searchBar.value = '';
      btnClearSearch.style.display = 'none';
      searchBar.focus();
      renderProductsList('');
    });
  }

  // Change file button - resets upload view
  if (btnReloadFile) {
    btnReloadFile.addEventListener('click', () => {
      showView('view-upload');
      if (uploadStatus) {
        uploadStatus.className = 'status-msg info';
        uploadStatus.textContent = 'No file loaded. Please load your product sheet.';
      }
    });
  }

  // Back button
  if (btnBackToList) {
    btnBackToList.addEventListener('click', () => {
      showView('view-list');
      activeProduct = null;
    });
  }

  // Open as Sidebar button (Popup only)
  if (btnSidebar) {
    btnSidebar.addEventListener('click', () => {
      chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
        if (tabs && tabs[0]) {
          chrome.runtime.sendMessage({ action: "open_sidebar", tabId: tabs[0].id });
        } else {
          showToast('Could not open sidebar: no active tab.');
        }
      });
    });
  }

  // Setup event delegation for product card clicks, copy clicks, and fill clicks
  document.addEventListener('click', handlePageClicks);
});

// Switch view panels safely
function showView(viewId) {
  const views = ['view-upload', 'view-list', 'view-detail'];
  views.forEach(id => {
    const el = document.getElementById(id);
    if (el) {
      el.style.display = (id === viewId) ? 'block' : 'none';
    }
  });
}

// Convert actual excel upload
function handleFileUpload(e) {
  const file = e.target.files[0];
  if (!file) return;

  const uploadStatus = document.getElementById('upload-status');
  if (uploadStatus) {
    uploadStatus.className = 'status-msg warning';
    uploadStatus.textContent = 'Parsing Excel sheet...';
  }

  const reader = new FileReader();
  reader.onload = function(evt) {
    try {
      const data = new Uint8Array(evt.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      
      const parsed = parseExcel(workbook);
      
      if (!parsed || parsed.length === 0) {
        throw new Error('No valid products found check columns A, B, E, G, H!');
      }

      // Save to chrome.storage
      allProducts = parsed;
      chrome.storage.local.set({
        'lifeinno_products': parsed,
        'lifeinno_filename': file.name
      }, () => {
        showLoadedState(parsed.length, file.name);
      });

    } catch (err) {
      console.error(err);
      if (uploadStatus) {
        uploadStatus.className = 'status-msg error';
        uploadStatus.textContent = '❌ Error: ' + err.message;
      }
    }
  };
  reader.readAsArrayBuffer(file);
}

// Spreadsheet reader with exact specified logic
function parseExcel(workbook) {
  const ws = workbook.Sheets[workbook.SheetNames[0]];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  
  const products = [];
  let current = null;

  for (let i = 1; i < raw.length; i++) {
    const row = raw[i];
    if (!row || row.length === 0) continue;

    const serial = row[0];
    const title = row[1];
    const size = row[2];
    const color = row[3];
    const pack = row[4];
    const basePrice = row[5];
    const actualPrice = row[6];
    const discountedPrice = row[7];

    // If Serial number exists, it is a new product
    if (serial !== null && serial !== undefined && String(serial).trim() !== '') {
      if (current) products.push(current);
      current = {
        serial: String(serial).trim(),
        title: title ? String(title).trim() : '—',
        size: size ? String(size).trim() : '—',
        color: color ? String(color).trim() : '—',
        basePrice: basePrice ? String(basePrice).trim() : '—',
        variations: []
      };
    }

    // Add variation to current product is pack slot name exists
    if (current && pack !== null && pack !== undefined && String(pack).trim() !== '') {
      current.variations.push({
        pack: String(pack).trim(),
        actualPrice: actualPrice !== null && actualPrice !== undefined ? String(actualPrice).trim() : '—',
        discountedPrice: discountedPrice !== null && discountedPrice !== undefined ? String(discountedPrice).trim() : '0'
      });
    }
  }
  
  if (current) products.push(current);
  return products;
}

// Configures and mounts products catalog page
function showLoadedState(count, filename) {
  const optLabel = document.getElementById('loaded-count');
  if (optLabel) {
    optLabel.textContent = `✅ ${count} products loaded ("${filename}")`;
  }
  renderProductsList('');
  showView('view-list');
}

// Draws scrollable products cards list
function renderProductsList(query) {
  const container = document.getElementById('products-scroll-container');
  if (!container) return;

  container.innerHTML = '';
  
  const filtered = allProducts.filter(p => {
    if (!query) return true;
    const q = query.toLowerCase();
    return (
      p.serial.toLowerCase().includes(q) ||
      p.title.toLowerCase().includes(q) ||
      p.size.toLowerCase().includes(q) ||
      p.color.toLowerCase().includes(q)
    );
  });

  if (filtered.length === 0) {
    container.innerHTML = `
      <div class="empty-state">
        <p>No products match "${query}"</p>
      </div>
    `;
    return;
  }

  filtered.forEach((p, idx) => {
    const card = document.createElement('div');
    card.className = 'product-card';
    card.dataset.serial = p.serial;
    
    card.innerHTML = `
      <div class="card-row">
        <span class="card-serial">#${p.serial}</span>
        <span class="card-vars">${p.variations.length} packs</span>
      </div>
      <h3 class="card-title">${p.title}</h3>
      <div class="card-footer">
        <span>Dim: ${p.size}</span>
        <span class="dot">|</span>
        <span>Col: ${p.color}</span>
      </div>
    `;
    container.appendChild(card);
  });
}

// Single click handler for delegating all views interactions
function handlePageClicks(e) {
  // 1. Click on product card
  const productCard = e.target.closest('.product-card');
  if (productCard) {
    const serial = productCard.dataset.serial;
    const prod = allProducts.find(p => p.serial === serial);
    if (prod) {
      openProductDetail(prod);
    }
    return;
  }

  // 2. Click on copy button
  const copyBtn = e.target.closest('.copy-btn');
  if (copyBtn) {
    const field = copyBtn.dataset.field;
    let value = '';

    if (copyBtn.dataset.customValue) {
      value = copyBtn.dataset.customValue;
    } else {
      value = getProductFieldValue(field);
    }

    copyTextToClipboard(value, copyBtn);
    return;
  }

  // 3. Click on fill button
  const fillBtn = e.target.closest('.fill-btn');
  if (fillBtn) {
    const field = fillBtn.dataset.field;
    let value = '';

    if (fillBtn.dataset.customValue) {
      value = fillBtn.dataset.customValue;
    } else {
      value = getProductFieldValue(field);
    }

    fillFieldInActiveTab(value, fillBtn);
    return;
  }
}

// Product detail drawer setup
function openProductDetail(product) {
  activeProduct = product;
  
  // Set badges and textual data
  document.getElementById('detail-serial-badge').textContent = `SN: ${product.serial}`;
  document.getElementById('val-title').textContent = product.title || '—';
  document.getElementById('val-size').textContent = product.size || '—';
  document.getElementById('val-color').textContent = product.color || '—';
  
  // Base Price showing as PKR string
  const base = product.basePrice;
  document.getElementById('val-base-price').textContent = base !== '—' && base ? `PKR ${base}` : '—';

  // Logistic calculations based on lowest variation
  const dims = getDimensions(product);
  document.getElementById('val-dimensions').textContent = `${dims.length} × ${dims.width} × ${dims.height}`;
  document.getElementById('val-weight').textContent = dims.weight;

  document.getElementById('val-dim-l').textContent = dims.length;
  document.getElementById('val-dim-w').textContent = dims.width;
  document.getElementById('val-dim-h').textContent = dims.height;

  // Bind individual logistics copy/fill custom parameters
  updateIndividualDimButtons(dims);

  // Variations list rendering
  const varBox = document.getElementById('variations-container');
  varBox.innerHTML = '';

  if (product.variations.length === 0) {
    varBox.innerHTML = `<div class="info-note">No variations found.</div>`;
  } else {
    product.variations.forEach(v => {
      const row = document.createElement('div');
      row.className = 'variation-card';
      row.innerHTML = `
        <div class="var-header">
          <span class="var-name">${v.pack}</span>
        </div>
        <div class="var-grid">
          <div class="var-price-cell">
            <span class="var-price-label">Actual Price</span>
            <div class="price-row">
              <span class="price-val">PKR ${v.actualPrice}</span>
              <button class="copy-btn nano-btn" data-custom-value="PKR ${v.actualPrice}" title="Copy Actual Price">📋</button>
              <button class="fill-btn nano-btn" data-custom-value="PKR ${v.actualPrice}" title="Submit Actual Price">⚡</button>
            </div>
          </div>
          <div class="var-price-cell">
            <span class="var-price-label">Special Price (Sale)</span>
            <div class="price-row highlight">
              <span class="price-val">PKR ${v.discountedPrice}</span>
              <button class="copy-btn nano-btn" data-custom-value="PKR ${v.discountedPrice}" title="Copy Sale Price">📋</button>
              <button class="fill-btn nano-btn" data-custom-value="PKR ${v.discountedPrice}" title="Submit Sale Price">⚡</button>
            </div>
          </div>
        </div>
      `;
      varBox.appendChild(row);
    });
  }

  showView('view-detail');
}

// Logic provided in specification
function getDimensions(product) {
  const smallestPack = product.variations[0];
  const priceAttr = smallestPack ? smallestPack.discountedPrice : '0';
  const price = parseFloat(priceAttr.replace(/[^0-9.]/g, '')) || 0;

  if (price <= 2000) {
    return { length: '20', width: '10', height: '5', weight: '300 grams', lwh: '20 x 10 x 5 cm' };
  } else {
    return { length: '30', width: '20', height: '10', weight: '700 grams', lwh: '30 x 20 x 10 cm' };
  }
}

// Binds custom values on the sub-component copy/fill targets
function updateIndividualDimButtons(dims) {
  const btns = document.querySelectorAll('[data-field="dim-l"], [data-field="dim-w"], [data-field="dim-h"]');
  btns.forEach(btn => {
    const f = btn.dataset.field;
    if (f === 'dim-l') btn.setAttribute('data-custom-value', dims.length);
    if (f === 'dim-w') btn.setAttribute('data-custom-value', dims.width);
    if (f === 'dim-h') btn.setAttribute('data-custom-value', dims.height);
  });
}

// Resolves text strings for dynamic copy actions based on active product
function getProductFieldValue(field) {
  if (!activeProduct) return '';

  switch (field) {
    case 'title':
      return activeProduct.title || '';
    case 'size':
      return activeProduct.size || '';
    case 'color':
      return activeProduct.color || '';
    case 'basePrice':
      return activeProduct.basePrice !== '—' ? `PKR ${activeProduct.basePrice}` : '';
    case 'dimensions':
      const dimsStr = getDimensions(activeProduct);
      return `${dimsStr.length} x ${dimsStr.width} x ${dimsStr.height}`;
    case 'weight':
      return getDimensions(activeProduct).weight;
    case 'warranty':
      return 'Seller Warranty — 1 Month';
    case 'returnPolicy':
      return `LIFEINNO Return Policy
If a product from Daraz is not working, you may return it under the following conditions:
• Product must be unused and in original condition
• Damaged or burnt items are not accepted
• Original packaging required
• Parcel opening video required
For more details Contact us: 03357714860 or IM Chat Please.`;
    default:
      return '';
  }
}

// Clipboard copier standard procedure
function copyTextToClipboard(text, btnElement) {
  if (!text) return;
  navigator.clipboard.writeText(text).then(() => {
    // Set last copied in storage for background "fill" floating helpers
    chrome.storage.local.set({ "lastCopiedValue": text });

    // Toast and Button feedback
    showToast('Copied to Clipboard!');
    
    if (btnElement) {
      const origText = btnElement.textContent;
      btnElement.textContent = '✅ Copied!';
      btnElement.classList.add('copied');
      setTimeout(() => {
        btnElement.textContent = origText;
        btnElement.classList.remove('copied');
      }, 1200);
    }
  }).catch(err => {
    console.error('Could not copy text: ', err);
    showToast('Failed to copy');
  });
}

// Form fills injecting trigger to focus webpage inputs
function fillFieldInActiveTab(value, btnElement) {
  if (!value) return;

  // Set as last copied to sync
  chrome.storage.local.set({ "lastCopiedValue": value });

  chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
    if (tabs && tabs[0]) {
      chrome.tabs.sendMessage(tabs[0].id, {
        type: 'FILL_FIELD',
        value: value
      }, (response) => {
        // Handle response or generic feedback
        if (chrome.runtime.lastError) {
          console.warn('Tab messaging error (is content.js loaded on seller.daraz.pk?):', chrome.runtime.lastError);
          showToast('⚠️ Click on an active input on Daraz first!');
        } else {
          showToast('⚡ Injected into active field!');
        }
      });
    } else {
      showToast('⚠️ No active tab detected');
    }
  });

  if (btnElement) {
    const origText = btnElement.textContent;
    btnElement.textContent = '⚡ Filled!';
    btnElement.style.backgroundColor = '#10B981';
    btnElement.style.color = '#FFF';
    setTimeout(() => {
      btnElement.textContent = origText;
      btnElement.style.backgroundColor = '';
      btnElement.style.color = '';
    }, 1000);
  }
}

// Utility Toast visual indicator
function showToast(message) {
  const toast = document.getElementById('toast-notification');
  if (!toast) return;

  toast.textContent = message;
  toast.className = 'toast-visible';
  setTimeout(() => {
    toast.className = 'toast-hidden';
  }, 2200);
}
