// ═══════════════════════════════════════════════════════
//  The Forbidden Style — Google Apps Script
//  Постави в: script.google.com → нов проект
//  После: Deploy → New deployment → Web App
//         Execute as: Me | Who can access: Anyone
// ═══════════════════════════════════════════════════════

// ─── НАСТРОЙКИ ───────────────────────────────────────
const OWNER_EMAIL   = 'YOUR_EMAIL@gmail.com';   // твоят имейл за поръчки
const SHEET_NAME_PRODUCTS = 'Products';
const SHEET_NAME_ORDERS   = 'Orders';
const STORE_NAME    = 'The Forbidden Style';
// ─────────────────────────────────────────────────────

// ── doGet: връща JSON данни (продукти или поръчки) ───
function doGet(e) {
  const action = e.parameter.action;

  if (action === 'checkAuth') {
    return checkAuth(e.parameter.pw);
  }

  if (action === 'getProducts') {
    const isAdmin = e.parameter.admin === '1';
    if (isAdmin && !isValidToken(e.parameter.token)) {
      return jsonResponse({ ok: false, error: 'Unauthorized' });
    }
    return getProductsJSON(!isAdmin); // публично = само активните
  }
  if (action === 'getOrders') {
    if (e.parameter.admin === '1' && !isValidToken(e.parameter.token)) {
      return jsonResponse({ ok: false, error: 'Unauthorized' });
    }
    return getOrdersJSON();
  }

  // Default: публична витрина — само активните продукти
  return getProductsJSON(true);
}

// ── doPost: приема поръчки и актуализации ────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.action === 'placeOrder') {
      // Поръчките идват от обикновени клиенти — не изискват admin токен
      return placeOrder(data);
    }

    // Всички останали действия променят данни и изискват валиден admin токен
    if (!isValidToken(data.token)) {
      return jsonResponse({ ok: false, error: 'Unauthorized' });
    }

    if (data.action === 'toggleCategory') {
      return toggleCategoryActive(data);
    }

    if (data.action === 'updateProduct') {
      return updateProduct(data);
    }
    if (data.action === 'addProduct') {
      return addProduct(data);
    }

    return jsonResponse({ ok: false, error: 'Unknown action' });
  } catch(err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

// ════════════════════════════════════════════════════
//  AUTH — паролата живее само тук, в Script Properties
//  (НЕ в кода, НЕ в GitHub). Виж README за настройка.
// ════════════════════════════════════════════════════

function checkAuth(password) {
  const props = PropertiesService.getScriptProperties();
  const correctPassword = props.getProperty('ADMIN_PASSWORD');

  if (!correctPassword) {
    return jsonResponse({ ok: false, error: 'Паролата не е настроена в Script Properties.' });
  }

  if (password === correctPassword) {
    const token = generateToken();
    // Запазваме токена за 8 часа валидност
    const cache = CacheService.getScriptCache();
    cache.put('token_' + token, 'valid', 8 * 60 * 60); // 8 часа в секунди
    return jsonResponse({ ok: true, token: token });
  }

  return jsonResponse({ ok: false, error: 'Грешна парола' });
}

function isValidToken(token) {
  if (!token) return false;
  const cache = CacheService.getScriptCache();
  return cache.get('token_' + token) === 'valid';
}

function generateToken() {
  return Utilities.getUuid();
}

// ════════════════════════════════════════════════════
//  PRODUCTS
// ════════════════════════════════════════════════════

function getProductsJSON(onlyActive) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_PRODUCTS);

  if (!sheet) {
    // Ако листът не съществува, създай го с примерни данни
    setupProductsSheet();
    return jsonResponse({ products: [] });
  }

  const rows = sheet.getDataRange().getValues();
  const headers = rows[0]; // id, category, name, desc, price, sizes_json, qty, active
  const products = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row[0]) continue; // skip empty rows

    const product = {
      id:       row[0],
      category: row[1],
      name:     row[2],
      desc:     row[3],
      price:    parseFloat(row[4]) || 0,
      sizes:    [],
      qty:      parseInt(row[6]) || 0,
      active:   row[7] === true || row[7] === 'TRUE' || row[7] === 1 || row[7] === '1'
    };

    // sizes_json is stored as JSON string in column F
    try {
      if (row[5]) product.sizes = JSON.parse(row[5]);
    } catch(e) {}

    // Публичната витрина вижда само активните (тикнатите) продукти
    if (onlyActive && !product.active) continue;

    products.push(product);
  }

  return jsonResponse({ products });
}

// Включва или изключва всички продукти от дадена категория наведнъж
function toggleCategoryActive(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_PRODUCTS);
  if (!sheet) return jsonResponse({ ok: false, error: 'Sheet not found' });

  const rows = sheet.getDataRange().getValues();
  const newValue = data.active === true ? 'TRUE' : 'FALSE';
  let count = 0;

  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    if (rows[i][1] === data.category) {
      sheet.getRange(i + 1, 8).setValue(newValue); // колона H = active
      count++;
    }
  }

  return jsonResponse({ ok: true, updated: count });
}

function updateProduct(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_PRODUCTS);
  if (!sheet) return jsonResponse({ ok: false, error: 'Sheet not found' });

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] == data.id) {
      // Update price
      if (data.price !== undefined) sheet.getRange(i+1, 5).setValue(data.price);
      // Update sizes JSON
      if (data.sizes !== null && data.sizes !== undefined) {
        sheet.getRange(i+1, 6).setValue(JSON.stringify(data.sizes));
      }
      // Update qty
      if (data.qty !== undefined && data.qty !== null) {
        sheet.getRange(i+1, 7).setValue(data.qty);
      }
      // Update visibility toggle (controls whether product shows on storefront)
      if (data.active !== undefined && data.active !== null) {
        sheet.getRange(i+1, 8).setValue(data.active === true ? 'TRUE' : 'FALSE');
      }
      return jsonResponse({ ok: true });
    }
  }
  return jsonResponse({ ok: false, error: 'Product not found' });
}

function addProduct(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME_PRODUCTS);
  if (!sheet) sheet = setupProductsSheet();

  sheet.appendRow([
    data.id,
    data.category,
    data.name,
    data.desc || '',
    data.price,
    data.sizes ? JSON.stringify(data.sizes) : '[]',
    data.qty || 0,
    data.active === true ? 'TRUE' : 'FALSE'
  ]);

  return jsonResponse({ ok: true });
}

// ════════════════════════════════════════════════════
//  ORDERS
// ════════════════════════════════════════════════════

function placeOrder(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME_ORDERS);
  if (!sheet) sheet = setupOrdersSheet();

  const itemsSummary = data.items.map(i =>
    `${i.name}${i.size ? ' ('+i.size+')' : ''} ×${i.qty} = лв ${(i.price*i.qty).toFixed(2)}`
  ).join('\n');

  // Save to sheet
  sheet.appendRow([
    new Date(),
    data.fullName,
    data.phone,
    data.email || '',
    data.courier,
    data.city,
    data.office,
    itemsSummary,
    'лв ' + data.total,
    'Нова'
  ]);

  // Бел.: наличността вече НЕ се намалява автоматично при поръчка —
  // ти управляваш бройките и видимостта ръчно от admin панела.

  // Send emails
  sendOwnerEmail(data, itemsSummary);
  if (data.email) sendClientEmail(data);

  return jsonResponse({ ok: true });
}

function getOrdersJSON() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_ORDERS);
  if (!sheet) return jsonResponse({ orders: [] });

  const rows = sheet.getDataRange().getValues();
  const orders = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row[1]) continue;
    orders.push({
      date:    row[0] ? new Date(row[0]).toISOString() : '',
      name:    row[1],
      phone:   row[2],
      email:   row[3],
      courier: row[4],
      city:    row[5],
      office:  row[6],
      itemsRaw: row[7],
      items:   parseOrderItems(row[7]),
      total:   String(row[8]).replace('лв ', ''),
      status:  row[9] || 'Нова'
    });
  }

  return jsonResponse({ orders });
}

function parseOrderItems(raw) {
  if (!raw) return [];
  return String(raw).split('\n').map(line => {
    const match = line.match(/^(.+?)(?:\s\((.+?)\))?\s×(\d+)/);
    if (match) return { name: match[1], size: match[2] || '', qty: parseInt(match[3]) };
    return { name: line, size: '', qty: 1 };
  });
}

// ════════════════════════════════════════════════════
//  EMAIL TEMPLATES
// ════════════════════════════════════════════════════

function sendOwnerEmail(data, itemsSummary) {
  const subject = `🛒 Нова поръчка — ${data.fullName} — лв ${data.total}`;
  const body = `
════════════════════════════
  НОВА ПОРЪЧКА — ${STORE_NAME}
════════════════════════════

👤 КЛИЕНТ:
   Име:      ${data.fullName}
   Телефон:  ${data.phone}
   Имейл:    ${data.email || '(не е посочен)'}

📦 ДОСТАВКА:
   Куриер:   ${data.courier}
   Град:     ${data.city}
   Офис:     ${data.office}

🛍️ ПРОДУКТИ:
${itemsSummary}

──────────────────────────────
   ОБЩО: лв ${data.total}
──────────────────────────────

Дата: ${new Date().toLocaleString('bg-BG')}
`;

  GmailApp.sendEmail(OWNER_EMAIL, subject, body);
}

function sendClientEmail(data) {
  const itemsHTML = data.items.map(i =>
    `<tr>
      <td style="padding:8px 12px;border-bottom:1px solid #eee;">${i.name}${i.size ? ' ('+i.size+')' : ''}</td>
      <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:center;">×${i.qty}</td>
      <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:right;">лв ${(i.price*i.qty).toFixed(2)}</td>
    </tr>`
  ).join('');

  const htmlBody = `
  <div style="font-family:Georgia,serif;max-width:560px;margin:0 auto;background:#0a0a0a;color:#f0ece4;padding:40px 30px;">
    <h1 style="font-size:2rem;font-weight:300;letter-spacing:0.05em;color:#c9a96e;margin-bottom:4px;">The Forbidden Style</h1>
    <p style="color:#888;font-size:0.8rem;letter-spacing:0.15em;text-transform:uppercase;margin-bottom:30px;">Потвърждение за поръчка</p>

    <p style="margin-bottom:20px;">Здравей, <strong>${data.fullName}</strong>!</p>
    <p style="color:#aaa;line-height:1.7;margin-bottom:30px;">
      Получихме твоята поръчка и скоро ще се свържем с теб.
      Доставката ще бъде изпратена чрез <strong style="color:#c9a96e;">${data.courier}</strong> до офис в <strong>${data.city}</strong>.
    </p>

    <table style="width:100%;border-collapse:collapse;margin-bottom:20px;background:#111;border-radius:4px;overflow:hidden;">
      <thead>
        <tr style="background:#1a1a1a;">
          <th style="padding:10px 12px;text-align:left;font-size:0.7rem;letter-spacing:0.1em;text-transform:uppercase;color:#888;">Продукт</th>
          <th style="padding:10px 12px;text-align:center;font-size:0.7rem;letter-spacing:0.1em;text-transform:uppercase;color:#888;">Бр.</th>
          <th style="padding:10px 12px;text-align:right;font-size:0.7rem;letter-spacing:0.1em;text-transform:uppercase;color:#888;">Цена</th>
        </tr>
      </thead>
      <tbody>${itemsHTML}</tbody>
      <tfoot>
        <tr style="background:#1a1a1a;">
          <td colspan="2" style="padding:10px 12px;font-size:0.8rem;letter-spacing:0.1em;text-transform:uppercase;color:#888;">Общо</td>
          <td style="padding:10px 12px;text-align:right;color:#c9a96e;font-size:1.1rem;">лв ${data.total}</td>
        </tr>
      </tfoot>
    </table>

    <p style="color:#aaa;font-size:0.82rem;line-height:1.7;border-top:1px solid #2a2a2a;padding-top:20px;">
      При въпроси се свържи с нас на <a href="tel:0876127997" style="color:#c9a96e;">0876 127 997</a>
    </p>

    <p style="color:#555;font-size:0.72rem;margin-top:30px;">© 2026 The Forbidden Style</p>
  </div>`;

  GmailApp.sendEmail(
    data.email,
    `✓ Поръчката ти в ${STORE_NAME} е получена`,
    `Поръчката ти е получена. Ще се свържем с теб скоро. Общо: лв ${data.total}`,
    { htmlBody }
  );
}

// ════════════════════════════════════════════════════
//  SETUP SHEETS (first run)
// ════════════════════════════════════════════════════

function setupProductsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.insertSheet(SHEET_NAME_PRODUCTS);
  sheet.appendRow(['id', 'category', 'name', 'desc', 'price', 'sizes_json', 'qty', 'active']);

  // Demo products — "active" решава дали продуктът се вижда на сайта.
  // Наличността (qty) е само информативна за теб, не влияе на видимостта.
  sheet.appendRow(['c1', 'clothes', 'Oversize Тениска "Shadow"', 'Тежък памук 280g/m², oversize fit', 45, JSON.stringify([{size:'S',qty:0},{size:'M',qty:0},{size:'L',qty:0},{size:'XL',qty:0},{size:'XXL',qty:0}]), 0, 'FALSE']);
  sheet.appendRow(['b1', 'bracelets', 'SmartBand X1', 'Фитнес трекер, пулс, крачки, 7 дни батерия', 55, '[]', 8, 'TRUE']);
  sheet.appendRow(['b2', 'bracelets', 'SmartBand Pro', 'AMOLED дисплей, SpO2, GPS, 5ATM', 89, '[]', 3, 'TRUE']);
  sheet.appendRow(['s1', 'speakers', 'BoomPod Mini', '10W, Bluetooth 5.3, IPX7, 12ч батерия', 65, '[]', 12, 'TRUE']);
  sheet.appendRow(['s2', 'speakers', 'BoomPod Max', '20W stereo, LED ринг, 18ч батерия', 95, '[]', 0, 'FALSE']);

  // Format header
  sheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#c9a96e');
  return sheet;
}

function setupOrdersSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.insertSheet(SHEET_NAME_ORDERS);
  sheet.appendRow(['Дата', 'Имe', 'Телефон', 'Имейл', 'Куриер', 'Град', 'Офис', 'Продукти', 'Сума', 'Статус']);
  sheet.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#c9a96e');
  return sheet;
}

// ── Helper ─────────────────────────────────────────
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
