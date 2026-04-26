// ═══════════════════════════════════════════════════════════
// STORELITE MASTER REGISTRY — Apps Script
// Deploy this on your MASTER REGISTRY Google Sheet ONLY
// (Sheet ID: 1K6jYaOrnmMLw_0_N5EvrgE5jEOpbIsWZjyCM8Yf6OLg)
//
// This script manages the Stores tab — add/edit/toggle stores.
// It does NOT handle individual store orders or products.
// Each store has its OWN Apps Script for that (updated-apps-script.js).
// ═══════════════════════════════════════════════════════════

var TEMPLATE_SHEET_ID = "PUT_YOUR_TEMPLATE_SHEET_ID_HERE";
var SITE_URL = "https://storelite.vercel.app";

function doGet(e) {
  var p = e.parameter;
  var action = p.action || '';
  
  // ═══ ADMIN AUTH ═══
  
  // Verify admin password server-side
  if (action === 'verifyAdmin' && p.pwd) {
    var result = verifyAdminPassword(p.pwd);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  // ═══ ADMIN ACTIONS (from admin.html) ═══
  // All write actions require admin token
  var token = p.token || '';
  
  // Add a new row to Stores tab
  if (action === 'addStoreRow') {
    if (!validateToken(token)) return jsonOut({error: 'Unauthorized'});
    addStoreRow(p);
    return ok('Store added');
  }
  
  // Update an existing row in Stores tab
  if (action === 'updateStoreRow' && p.row) {
    if (!validateToken(token)) return jsonOut({error: 'Unauthorized'});
    updateStoreRow(p);
    return ok('Store updated');
  }
  
  // Update a single cell (e.g. toggle Active)
  if (action === 'updateRegistry' && p.row && p.column) {
    if (!validateToken(token)) return jsonOut({error: 'Unauthorized'});
    updateRegistryCell(parseInt(p.row), p.column, p.value || '');
    return ok('Cell updated');
  }
  
  // ═══ ONBOARDING ACTIONS (from landing page signup) ═══
  
  // Auto-create store from signup form
  if (action === 'createStore') {
    var result = createStore(p);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  // Check if slug is available
  if (action === 'checkSlug') {
    var available = isSlugAvailable(p.slug || '');
    return ContentService.createTextOutput(JSON.stringify({available: available}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  return ok('Registry endpoint active');
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    if (data.action === 'verifyAdmin') {
      var result = verifyAdminPassword(data.pwd || '');
      return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    }
    var token = data.token || '';
    if (data.action === 'addStoreRow') { if(!validateToken(token))return jsonOut({error:'Unauthorized'}); addStoreRow(data); return ok('Store added'); }
    if (data.action === 'updateStoreRow') { if(!validateToken(token))return jsonOut({error:'Unauthorized'}); updateStoreRow(data); return ok('Store updated'); }
    if (data.action === 'updateRegistry') { if(!validateToken(token))return jsonOut({error:'Unauthorized'}); updateRegistryCell(parseInt(data.row), data.column, data.value || ''); return ok('Cell updated'); }
    if (data.action === 'createStore') {
      var result = createStore(data);
      return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    }
    return ok('Unknown action');
  } catch (err) { return ok('Error: ' + err); }
}

function ok(msg) { return ContentService.createTextOutput(msg); }
function jsonOut(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }

// ═══════════════════════════════════════
// ADMIN AUTH — Server-side password check
// ═══════════════════════════════════════
// HOW TO SET YOUR PASSWORD:
// 1. In your Master Registry sheet, create a tab called "Settings"
// 2. Row 1: Key | Value (headers)
// 3. Row 2: AdminPassword | YOUR_STRONG_PASSWORD_HERE
// The password is NEVER sent to the browser. Only validated here.

function verifyAdminPassword(pwd) {
  if (!pwd) return {success: false, error: 'Password required'};
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settingsSheet = ss.getSheetByName('Settings');
  
  if (!settingsSheet) {
    // Fallback: no Settings tab, create one with default
    settingsSheet = ss.insertSheet('Settings');
    settingsSheet.getRange(1, 1, 1, 2).setValues([['Key', 'Value']]);
    settingsSheet.getRange(2, 1, 1, 2).setValues([['AdminPassword', 'changeme123']]);
    settingsSheet.getRange(1, 1, 1, 2).setFontWeight('bold');
  }
  
  // Read password from Settings tab
  var data = settingsSheet.getDataRange().getValues();
  var storedPwd = '';
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === 'adminpassword') {
      storedPwd = String(data[i][1]).trim();
      break;
    }
  }
  
  if (!storedPwd) return {success: false, error: 'Admin password not configured in Settings tab'};
  
  // Compare using SHA256 hash to prevent timing attacks
  var inputHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pwd);
  var storedHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, storedPwd);
  
  var match = true;
  for (var i = 0; i < inputHash.length; i++) {
    if (inputHash[i] !== storedHash[i]) match = false;
  }
  
  if (match) {
    // Generate a session token (valid for 2 hours)
    var token = generateToken();
    storeToken(token);
    return {success: true, token: token};
  }
  
  return {success: false, error: 'Wrong password'};
}

function generateToken() {
  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    Utilities.getUuid() + Date.now() + Math.random()
  );
  return bytes.map(function(b) {
    return ('0' + (b & 0xFF).toString(16)).slice(-2);
  }).join('').substring(0, 32);
}

function storeToken(token) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settingsSheet = ss.getSheetByName('Settings');
  if (!settingsSheet) return;
  
  var data = settingsSheet.getDataRange().getValues();
  var found = false;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === 'admintoken') {
      settingsSheet.getRange(i + 1, 2).setValue(token + '|' + Date.now());
      found = true;
      break;
    }
  }
  if (!found) {
    settingsSheet.appendRow(['AdminToken', token + '|' + Date.now()]);
  }
}

function validateToken(token) {
  if (!token) return false;
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settingsSheet = ss.getSheetByName('Settings');
  if (!settingsSheet) return false;
  
  var data = settingsSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === 'admintoken') {
      var parts = String(data[i][1]).trim().split('|');
      var storedToken = parts[0];
      var timestamp = parseInt(parts[1]) || 0;
      
      // Token expires after 2 hours
      if (Date.now() - timestamp > 2 * 60 * 60 * 1000) return false;
      
      return storedToken === token;
    }
  }
  return false;
}

// ═══════════════════════════════════════
// ADMIN: ADD STORE ROW
// Adds a new row to the Stores tab
// ═══════════════════════════════════════
function addStoreRow(p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateStoresTab(ss);
  var headers = getHeaders(sheet);
  
  var row = [];
  headers.forEach(function(h) {
    // Match header to parameter (case-insensitive)
    var key = h.toLowerCase().replace(/\s+/g, '');
    var val = findParam(p, key) || '';
    row.push(val);
  });
  
  sheet.appendRow(row);
  
  // Color the new row
  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 1, 1, headers.length).setBackground('#e8faed');
}

// ═══════════════════════════════════════
// ADMIN: UPDATE STORE ROW
// Updates specific columns in an existing row
// ═══════════════════════════════════════
function updateStoreRow(p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Stores');
  if (!sheet) return;
  
  var rowNum = parseInt(p.row || p.Row);
  if (!rowNum || rowNum < 2) return;
  
  var headers = getHeaders(sheet);
  
  headers.forEach(function(h, i) {
    var key = h.toLowerCase().replace(/\s+/g, '');
    var val = findParam(p, key);
    if (val !== undefined && val !== null && val !== '') {
      sheet.getRange(rowNum, i + 1).setValue(val);
    }
  });
}

// ═══════════════════════════════════════
// ADMIN: UPDATE SINGLE CELL
// Used for toggling Active, quick edits
// ═══════════════════════════════════════
function updateRegistryCell(rowNum, columnName, value) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Stores');
  if (!sheet || rowNum < 2) return;
  
  var headers = getHeaders(sheet);
  var colName = columnName.toLowerCase().replace(/\s+/g, '');
  
  for (var i = 0; i < headers.length; i++) {
    if (headers[i].toLowerCase().replace(/\s+/g, '') === colName) {
      sheet.getRange(rowNum, i + 1).setValue(value);
      
      // Color coding for Active column
      if (colName === 'active') {
        var bg = value.toLowerCase() === 'yes' ? '#e8faed' : '#fef2f2';
        sheet.getRange(rowNum, i + 1).setBackground(bg);
      }
      break;
    }
  }
}

// ═══════════════════════════════════════
// ONBOARDING: AUTO-CREATE STORE
// Called from landing page signup form
// ═══════════════════════════════════════
function createStore(p) {
  var shopName = p.shopName || p.ShopName || p.shopname || '';
  var ownerName = p.ownerName || p.OwnerName || p.name || '';
  var phone = p.phone || p.OwnerPhone || '';
  var address = p.address || '';
  var shopType = p.shopType || p.ShopType || p.type || 'Grocery';
  var upi = p.upi || '';
  var plan = p.plan || p.Plan || 'Free';
  var city = p.city || p.City || extractCity(address);
  
  if (!shopName || !ownerName || !phone) {
    return {error: 'Shop name, owner name, and phone are required'};
  }
  
  var slug = generateSlug(shopName, city);
  var baseSlug = slug;
  var counter = 1;
  while (!isSlugAvailable(slug)) { slug = baseSlug + '-' + counter; counter++; }
  
  var pin = String(Math.floor(1000 + Math.random() * 9000));
  
  // Copy template sheet
  var newSheet;
  try {
    var template = DriveApp.getFileById(TEMPLATE_SHEET_ID);
    var copy = template.makeCopy(shopName + ' — StoreLite Store');
    newSheet = SpreadsheetApp.openById(copy.getId());
    copy.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (err) {
    return {error: 'Failed to create sheet: ' + String(err)};
  }
  
  var newSheetId = newSheet.getId();
  
  // Populate Config tab
  try {
    var configSheet = newSheet.getSheetByName('Config');
    if (configSheet) {
      var updates = {
        'ShopName': shopName, 'Phone': '91' + phone, 'WhatsApp': '91' + phone,
        'Address': address, 'ShopType': shopType, 'City': city, 'UPI': upi,
        'DashboardPIN': pin, 'Currency': '₹', 'MinOrder': '199',
        'DeliveryFee': '30', 'FreeDelivery': '999'
      };
      var configData = configSheet.getDataRange().getValues();
      var configMap = {};
      for (var i = 0; i < configData.length; i++) {
        configMap[String(configData[i][0]).trim().toLowerCase()] = i + 1;
      }
      for (var key in updates) {
        if (updates[key]) {
          var lk = key.toLowerCase();
          if (configMap[lk]) configSheet.getRange(configMap[lk], 2).setValue(updates[key]);
          else configSheet.appendRow([key, updates[key]]);
        }
      }
    }
  } catch (err) {}
  
  // Add to registry
  var storeUrl = SITE_URL + '/?store=' + slug;
  var dashUrl = SITE_URL + '/dashboard.html?store=' + slug;
  var expiry = '';
  if (plan !== 'Free') {
    var d = new Date(); d.setDate(d.getDate() + 30);
    expiry = d.toISOString().split('T')[0];
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var storesSheet = getOrCreateStoresTab(ss);
  
  storesSheet.appendRow([slug, newSheetId, ownerName, phone, shopName, shopType, city, plan, expiry, 'Yes', storeUrl, dashUrl]);
  storesSheet.getRange(storesSheet.getLastRow(), 1, 1, 12).setBackground('#e8faed');
  
  // Send welcome email
  if (p.email) {
    try { sendWelcomeEmail(p.email, shopName, ownerName, storeUrl, dashUrl, pin, newSheetId); } catch (err) {}
  }
  
  return { success: true, slug: slug, sheetId: newSheetId, storeUrl: storeUrl, dashboardUrl: dashUrl, pin: pin };
}

// ═══════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════
function getOrCreateStoresTab(ss) {
  var sheet = ss.getSheetByName('Stores');
  if (!sheet) {
    sheet = ss.insertSheet('Stores');
    sheet.getRange(1, 1, 1, 12).setValues([['Slug', 'SheetID', 'OwnerName', 'OwnerPhone', 'ShopName', 'ShopType', 'City', 'Plan', 'PlanExpiry', 'Active', 'URL', 'DashBoardURL']]);
    sheet.getRange(1, 1, 1, 12).setFontWeight('bold').setBackground('#0c831f').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getHeaders(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h) { return String(h).trim(); });
}

function findParam(p, key) {
  // Try exact match, then case variations
  if (p[key] !== undefined) return p[key];
  var upper = key.charAt(0).toUpperCase() + key.slice(1);
  if (p[upper] !== undefined) return p[upper];
  // Try all keys
  for (var k in p) {
    if (k.toLowerCase().replace(/\s+/g, '') === key) return p[k];
  }
  return undefined;
}

function generateSlug(name, city) {
  return (name + (city ? '-' + city : '')).toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '').substring(0, 40) || 'store';
}

function extractCity(address) {
  if (!address) return '';
  var parts = address.split(',').map(function(s) { return s.trim(); });
  return parts.length >= 2 ? parts[parts.length - 2] || parts[0] : parts[0] || '';
}

function isSlugAvailable(slug) {
  if (!slug) return false;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Stores');
  if (!sheet) return true;
  var data = sheet.getDataRange().getValues();
  slug = slug.toLowerCase();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === slug) return false;
  }
  return true;
}

function sendWelcomeEmail(email, shopName, ownerName, storeUrl, dashUrl, pin, sheetId) {
  var html = '<!DOCTYPE html><html><body style="margin:0;padding:0;font-family:-apple-system,sans-serif;background:#f5f5f5">'
    + '<div style="max-width:520px;margin:0 auto;background:#fff">'
    + '<div style="background:#0c831f;padding:28px 20px;text-align:center"><div style="font-size:36px">🏬</div><div style="color:#fff;font-size:22px;font-weight:700">Welcome to StoreLite!</div></div>'
    + '<div style="padding:24px"><div style="font-size:15px">Hi <strong>' + ownerName + '</strong>,</div>'
    + '<div style="font-size:14px;color:#666;margin-top:8px;line-height:1.7">Your store <strong>' + shopName + '</strong> is live!</div>'
    + '<div style="margin:16px 0;padding:16px;background:#e8faed;border-radius:8px">'
    + '<div style="font-size:12px;color:#666;margin-bottom:4px">Store:</div><a href="' + storeUrl + '" style="color:#0c831f;font-weight:700">' + storeUrl + '</a>'
    + '<div style="font-size:12px;color:#666;margin:12px 0 4px">Dashboard:</div><a href="' + dashUrl + '" style="color:#0c831f;font-weight:700">' + dashUrl + '</a>'
    + '<div style="font-size:12px;color:#666;margin:12px 0 4px">PIN:</div><div style="font-size:24px;font-weight:700;color:#0c831f;letter-spacing:4px">' + pin + '</div>'
    + '<div style="font-size:12px;color:#666;margin:12px 0 4px">Google Sheet:</div><a href="https://docs.google.com/spreadsheets/d/' + sheetId + '" style="color:#0c831f">Open Sheet</a>'
    + '</div></div>'
    + '<div style="padding:16px 24px;text-align:center;border-top:1px solid #eee;font-size:11px;color:#bbb">StoreLite</div></div></body></html>';
  MailApp.sendEmail({ to: email, subject: '🏬 Your Store is Live — ' + shopName, htmlBody: html });
}
