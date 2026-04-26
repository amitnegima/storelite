function doGet(e) {
  var p = e.parameter;
  var action = p.action || '';
  
  if (action === 'updateStatus' && p.orderId && p.newStatus) {
    updateOrderStatus(p.orderId, p.newStatus);
    return ok('Status updated');
  }
  if (action === 'updateConfig' && p.key) {
    updateConfig(p.key, p.value || '');
    return ok('Config updated');
  }
  if (action === 'addProduct' && p.name) {
    addProduct(p);
    return ok('Product added');
  }
  if (action === 'updateProduct' && p.row) {
    updateProduct(p);
    return ok('Product updated');
  }
  if (action === 'deleteProduct' && p.row) {
    deleteProduct(parseInt(p.row));
    return ok('Product deleted');
  }
  if (p.orderId && !action) {
    saveOrder(p);
    return ok('OK');
  }
  return ok('Endpoint active');
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    if (data.action === 'updateConfig') { updateConfig(data.key, data.value || ''); return ok('Config updated'); }
    if (data.action === 'addProduct') { addProduct(data); return ok('Product added'); }
    if (data.action === 'updateProduct') { updateProduct(data); return ok('Product updated'); }
    if (data.action === 'deleteProduct') { deleteProduct(parseInt(data.row)); return ok('Product deleted'); }
    saveOrder(data);
    return ok('OK');
  } catch (err) { return ok('Error: ' + err); }
}

function ok(msg) { return ContentService.createTextOutput(msg); }

// ═══ ORDERS ═══
function saveOrder(p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Orders');
  if (!sheet) {
    sheet = ss.insertSheet('Orders');
    sheet.getRange(1, 1, 1, 10).setValues([['Order ID', 'Date & Time', 'Mode', 'Customer Name', 'Phone', 'Email', 'Address', 'Items', 'Total', 'Status']]);
    sheet.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#0c831f').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([p.orderId||'', p.date||new Date().toLocaleString('en-IN',{timeZone:'Asia/Kolkata'}), (p.mode||'pickup').toUpperCase(), p.name||'', p.phone||'', p.email||'', p.address||'', p.items||'', p.total||0, p.status||'New']);
  sheet.getRange(sheet.getLastRow(), 10).setBackground('#e8faed');
}

function updateOrderStatus(orderId, newStatus) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Orders');
  if (!sheet) return;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === orderId) {
      sheet.getRange(i+1, 10).setValue(newStatus);
      var colors = {'New':'#e8faed','Confirmed':'#eff6ff','Packed':'#fff7ed','Delivered':'#f3f4f6','Picked Up':'#f3f4f6','Cancelled':'#fef2f2'};
      sheet.getRange(i+1, 10).setBackground(colors[newStatus]||'#ffffff');
      var email = data[i][5];
      if (email && newStatus === 'Confirmed') sendConfirmedEmail(email, orderId, data[i][3], data[i][4], data[i][2], data[i][6], data[i][7], data[i][8]);
      if (email && newStatus === 'Delivered') sendDeliveredEmail(email, orderId, data[i][3], data[i][8]);
      break;
    }
  }
}

// ═══ CONFIG ═══
function updateConfig(key, value) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Config');
  if (!sheet) return;
  var data = sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === String(key).trim().toLowerCase()) {
      sheet.getRange(i+1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}

// ═══ PRODUCTS ═══
function getProductHeaders(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h) { return String(h).trim().toLowerCase(); });
}

function addProduct(p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Products') || ss.getSheetByName('Sheet1');
  if (!sheet) {
    sheet = ss.insertSheet('Products');
    sheet.getRange(1,1,1,10).setValues([['Name','Category','Price','MRP','Unit','Image','Description','Type','Stock','Badge']]);
    sheet.getRange(1,1,1,10).setFontWeight('bold').setBackground('#0c831f').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  var headers = getProductHeaders(sheet);
  var row = [];
  headers.forEach(function(h) { row.push(p[h] || p[h.charAt(0).toUpperCase()+h.slice(1)] || ''); });
  sheet.appendRow(row);
}

function updateProduct(p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Products') || ss.getSheetByName('Sheet1');
  if (!sheet) return;
  var rowNum = parseInt(p.row);
  if (rowNum < 2) return;
  var headers = getProductHeaders(sheet);
  headers.forEach(function(h, i) {
    var val = p[h] || p[h.charAt(0).toUpperCase()+h.slice(1)];
    if (val !== undefined && val !== null) sheet.getRange(rowNum, i+1).setValue(val);
  });
}

function deleteProduct(rowNum) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Products') || ss.getSheetByName('Sheet1');
  if (!sheet || rowNum < 2) return;
  sheet.deleteRow(rowNum);
}

// ═══ EMAILS ═══
function sendConfirmedEmail(email, orderId, name, phone, mode, address, items, total) {
  try {
    var shopName = 'Grocery Grocery & Confectionery Shop';
    var modeLabel = (mode||'').toUpperCase()==='DELIVERY' ? '🚚 Home Delivery' : '🏪 Store Pickup';
    var itemRows = (items||'').split(',').map(function(item){return '<tr><td style="padding:10px 14px;border-bottom:1px solid #f0f0f0;font-size:14px;color:#333">'+item.trim()+'</td></tr>';}).join('');
    var html = '<!DOCTYPE html><html><body style="margin:0;padding:0;font-family:-apple-system,sans-serif;background:#f5f5f5"><div style="max-width:520px;margin:0 auto;background:#fff"><div style="background:#0c831f;padding:28px 20px;text-align:center"><div style="font-size:36px">✅</div><div style="color:#fff;font-size:22px;font-weight:700">Order Confirmed!</div><div style="color:rgba(255,255,255,.75);font-size:13px;margin-top:6px">'+shopName+'</div></div><div style="background:#e8faed;padding:14px 20px;text-align:center"><span style="font-size:18px;font-weight:700;color:#0c831f">'+orderId+'</span></div><div style="padding:20px 24px"><div style="font-size:15px">Hi <strong>'+(name||'Customer')+'</strong>, your order is confirmed!</div></div><div style="padding:0 24px"><table style="width:100%;background:#fafafa;border-radius:8px">'+itemRows+'</table></div><div style="margin:16px 24px;padding:16px;background:#0c831f;border-radius:10px;text-align:center"><div style="font-size:28px;font-weight:700;color:#fff">₹'+total+'</div></div><div style="padding:16px 24px;text-align:center;border-top:1px solid #eee;font-size:11px;color:#bbb">'+shopName+'</div></div></body></html>';
    MailApp.sendEmail({to:email, subject:'✅ Order Confirmed — '+orderId, htmlBody:html});
  } catch(err) { console.log('Email error: '+err); }
}

function sendDeliveredEmail(email, orderId, name, total) {
  try {
    var shopName = 'Grocery Grocery & Confectionery Shop';
    var html = '<!DOCTYPE html><html><body style="margin:0;padding:0;font-family:-apple-system,sans-serif;background:#f5f5f5"><div style="max-width:520px;margin:0 auto;background:#fff"><div style="background:#0c831f;padding:28px 20px;text-align:center"><div style="font-size:36px">🎉</div><div style="color:#fff;font-size:22px;font-weight:700">Order Delivered!</div></div><div style="padding:28px 24px;text-align:center"><div style="font-size:16px;font-weight:600">Thank you, '+(name||'Customer')+'! 🙏</div><div style="font-size:14px;color:#666;margin-top:10px">Your order <strong>'+orderId+'</strong> worth <strong>₹'+total+'</strong> has been delivered.</div></div><div style="padding:16px 24px;text-align:center;border-top:1px solid #eee;font-size:11px;color:#bbb">'+shopName+'</div></div></body></html>';
    MailApp.sendEmail({to:email, subject:'🎉 Delivered — '+orderId, htmlBody:html});
  } catch(err) { console.log('Email error: '+err); }
}
