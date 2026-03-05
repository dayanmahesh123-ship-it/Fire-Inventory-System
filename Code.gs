/* ================================================================
   🔥 FIRE PROTECTION & DETECTION — IMS
   Complete Code.gs — Fixed & Optimized
   ================================================================ */

var SHEET_NAME = 'FP_Inventory';
var ISSUE_SHEET = 'FP_Issues';

var SYSTEM_CONFIG = {
  HEADERS: [
    'ID','ItemName','Category','SubCategory','Size',
    'Unit','Quantity','MinThreshold','Location',
    'TechnicalSpecs','LastUpdated','Notes','Status'
  ],
  ISSUE_HEADERS: [
    'IssueID','ItemID','ItemName','Category','SubCategory',
    'Size','Unit','QtyIssued','IssuedTo',
    'IssueDate','IssueTime','IssueLocation',
    'ProjectName','Description','IssuedBy','Timestamp'
  ],
  CATEGORIES: {
    'Valve Station': [
      'Alarm Valve','Gate Valve','Check Valve','Butterfly Valve',
      'Pressure Gauge','Strainer','Trim Set','Retarding Chamber',
      'Water Motor Gong','OS&Y Valve','Pressure Relief Valve','Test & Drain Valve'
    ],
    'Fire Detection & Alarm': [
      'Smoke Detector','Heat Detector','Multi-Sensor Detector','Beam Detector',
      'Aspirating Smoke Detector','Duct Detector','Flame Detector',
      'Carbon Monoxide Detector','Manual Call Point','Fire Alarm Control Panel',
      'Repeater Panel','Network Card','Graphic Annunciator','Sounder',
      'Beacon Strobe','Sounder Beacon','Input Module','Output Module',
      'Loop Card','Isolator Module','Zone Module','Relay Module',
      'Response Indicator','End of Line Device','Detector Base',
      'Weatherproof Housing','Fire Telephone','Power Supply Unit',
      'Battery - Standby','Cable - Fire Rated','Cable - Standard',
      'Junction Box','Conduit & Fittings'
    ],
    'Suppression & Piping': [
      'Black Steel Pipe','Galvanized Pipe','CPVC Pipe',
      'Sprinkler Head - Pendent','Sprinkler Head - Upright',
      'Sprinkler Head - Sidewall','Sprinkler Head - Concealed',
      'Flow Switch','Pressure Switch','Flexible Connection',
      'Fire Hose Reel','Nozzle','Deluge Valve','Pre-Action Valve',
      'Dry Pipe Valve','Fire Hose Cabinet','Landing Valve','Breeching Inlet'
    ],
    'Fittings & Joints': [
      'Elbow 90°','Elbow 45°','Tee','Reducer','Coupling',
      'Mechanical Tee','Flange','Union','Nipple','Cap',
      'Bushing','Cross','Adaptor','Flange Adaptor'
    ],
    'Support & Hardware': [
      'Anchor Bolt','U-Clamp','Threaded Rod','Hex Nut','Hex Bolt',
      'Flat Washer','Spring Washer','Gasket','Hanger Rod','Pipe Clamp',
      'Spring Nut','Channel Bracket','Drop Rod','Beam Clamp',
      'C-Clamp','Toggle Bolt'
    ],
    'Gas Suppression': [
      'FM200 Cylinder','Novec 1230 Cylinder','CO2 Cylinder',
      'Inert Gas Cylinder (IG-541)','Discharge Nozzle',
      'Actuation Device - Electric','Actuation Device - Pneumatic',
      'Abort Switch','Release Panel','Pressure Switch - HP',
      'Flexible Discharge Hose','Cylinder Bracket','Selector Valve',
      'Check Valve - HP','Warning Sign','Door Holder Release',
      'Hooter - Gas Release','Pressure Relief Vent'
    ],
    'Fire Extinguishers': [
      'ABC Dry Powder','CO2 Extinguisher','Foam Extinguisher',
      'Water Extinguisher','Wet Chemical','Clean Agent Portable',
      'Wheeled Extinguisher','Extinguisher Cabinet','Extinguisher Stand',
      'Extinguisher Sign','Extinguisher Bracket','Spare Cartridge'
    ],
    'Emergency & Safety': [
      'Emergency Light','Exit Sign - Illuminated','Emergency Exit Light',
      'Fire Door Hardware','Door Closer - Fire Rated','Panic Bar',
      'Fire Rated Sealant','Firestop Collar','Firestop Pillow',
      'Firestop Mortar','Firestop Wrap Strip','Fire Blanket',
      'Fire Damper','Smoke Damper','Smoke Curtain','First Aid Kit',
      'Safety Sign','Fire Action Notice','Assembly Point Sign'
    ]
  },
  UNITS: ['Nos','Meters','Sets','Boxes','Pairs','Rolls','Kg','Lengths','Liters','Cans'],
  SIZES: [
    '1/2"','3/4"','1"','1-1/4"','1-1/2"','2"','2-1/2"','3"','4"','5"','6"','8"','10"','12"',
    'M6','M8','M10','M12','M16','M20',
    '1kg','2kg','4kg','6kg','9kg','25kg','50kg',
    '1L','2L','6L','9L','120L','180L','200L',
    'Standard','N/A'
  ]
};


/* ================================================================
   doGet — Web App Entry Point
   ================================================================ */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Fire Protection & Detection — IMS')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}


/* ================================================================
   HELPER FUNCTIONS
   ================================================================ */

function getOrCreateSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    var h = SYSTEM_CONFIG.HEADERS;
    sh.getRange(1, 1, 1, h.length).setValues([h]);
    sh.getRange(1, 1, 1, h.length)
      .setBackground('#7f1d1d')
      .setFontColor('#fff')
      .setFontWeight('bold')
      .setFontSize(10)
      .setHorizontalAlignment('center');
    sh.setFrozenRows(1);
    [160,220,180,190,80,70,80,100,130,260,170,220,70].forEach(function(w, i) {
      sh.setColumnWidth(i + 1, w);
    });
  }
  return sh;
}

function getOrCreateIssueSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(ISSUE_SHEET);
  if (!sh) {
    sh = ss.insertSheet(ISSUE_SHEET);
    var h = SYSTEM_CONFIG.ISSUE_HEADERS;
    sh.getRange(1, 1, 1, h.length).setValues([h]);
    sh.getRange(1, 1, 1, h.length)
      .setBackground('#1e3a5f')
      .setFontColor('#fff')
      .setFontWeight('bold')
      .setFontSize(10)
      .setHorizontalAlignment('center');
    sh.setFrozenRows(1);
    [160,160,220,160,170,80,70,90,180,120,100,180,180,280,180,170].forEach(function(w, i) {
      sh.setColumnWidth(i + 1, w);
    });
  }
  return sh;
}

function generateId_() {
  return 'FP-' + Date.now().toString(36).toUpperCase() + '-' + Math.random().toString(36).substr(2, 6).toUpperCase();
}

function generateIssueId_() {
  return 'ISS-' + Date.now().toString(36).toUpperCase() + '-' + Math.random().toString(36).substr(2, 6).toUpperCase();
}

function rowToObject_(row) {
  var o = {};
  SYSTEM_CONFIG.HEADERS.forEach(function(h, i) { o[h] = row[i]; });
  o.Quantity = Number(o.Quantity) || 0;
  o.MinThreshold = Number(o.MinThreshold) || 0;
  o.isLowStock = o.MinThreshold > 0 && o.Quantity <= o.MinThreshold;
  return o;
}

function issueRowToObject_(row) {
  var o = {};
  SYSTEM_CONFIG.ISSUE_HEADERS.forEach(function(h, i) { o[h] = row[i]; });
  o.QtyIssued = Number(o.QtyIssued) || 0;
  return o;
}

function computeStatus_(q, t) {
  if (t <= 0) return 'OK';
  if (q <= 0) return 'OUT';
  if (q <= t) return 'LOW';
  return 'OK';
}

function sanitize_(v) {
  return String(v || '').replace(/[<>]/g, '').trim();
}

function findRowById_(sh, id) {
  var last = sh.getLastRow();
  if (last <= 1) return -1;
  var d = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < d.length; i++) {
    if (String(d[i][0]).trim() === String(id).trim()) return i + 2;
  }
  return -1;
}

function validateItemData_(d) {
  if (!d || typeof d !== 'object') return { valid: false, message: 'Invalid data.' };
  if (!d.ItemName || String(d.ItemName).trim().length < 2) return { valid: false, message: 'Item name required (min 2 chars).' };
  if (!d.Category || !SYSTEM_CONFIG.CATEGORIES[d.Category]) return { valid: false, message: 'Invalid category.' };
  if (isNaN(Number(d.Quantity)) || Number(d.Quantity) < 0) return { valid: false, message: 'Invalid quantity.' };
  if (isNaN(Number(d.MinThreshold)) || Number(d.MinThreshold) < 0) return { valid: false, message: 'Invalid threshold.' };
  return { valid: true };
}

function buildStats_(items) {
  var cats = Object.keys(SYSTEM_CONFIG.CATEGORIES);
  var bk = {};

  cats.forEach(function(c) {
    var ci = items.filter(function(i) { return i.Category === c; });
    bk[c] = {
      count: ci.length,
      totalQty: ci.reduce(function(s, i) { return s + i.Quantity; }, 0),
      lowStock: ci.filter(function(i) { return i.isLowStock; }).length
    };
  });

  return {
    totalItems: items.length,
    totalQuantity: items.reduce(function(s, i) { return s + i.Quantity; }, 0),
    lowStockCount: items.filter(function(i) { return i.isLowStock; }).length,
    outOfStockCount: items.filter(function(i) { return i.Quantity <= 0 && i.MinThreshold > 0; }).length,
    breakdown: bk,
    lowStockItems: items.filter(function(i) { return i.isLowStock; })
  };
}


/* ================================================================
   PUBLIC API — INITIAL DATA LOAD
   ================================================================ */

function getConfig() {
  return JSON.parse(JSON.stringify(SYSTEM_CONFIG));
}

function getInitialData() {
  try {
    var config = JSON.parse(JSON.stringify(SYSTEM_CONFIG));
    var items = getAllItems();
    var stats = buildStats_(items);
    var emailConfig = getEmailConfig();
    var issues = getAllIssues();

    return {
      config: config,
      items: items,
      stats: stats,
      emailConfig: emailConfig,
      issues: issues
    };
  } catch (e) {
    Logger.log('getInitialData ERROR: ' + e.message);
    throw new Error('Failed to load data: ' + e.message);
  }
}


/* ================================================================
   PUBLIC API — INVENTORY CRUD
   ================================================================ */

function getAllItems() {
  try {
    var sh = getOrCreateSheet_();
    var last = sh.getLastRow();
    if (last <= 1) return [];
    return sh.getRange(2, 1, last - 1, SYSTEM_CONFIG.HEADERS.length)
      .getValues()
      .filter(function(r) { return r[0] && String(r[0]).trim() !== ''; })
      .map(rowToObject_);
  } catch (e) {
    Logger.log('getAllItems ERROR: ' + e.message);
    return [];
  }
}

function addItem(d) {
  var v = validateItemData_(d);
  if (!v.valid) return { success: false, message: v.message };

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var sh = getOrCreateSheet_();
    var id = generateId_();
    var now = new Date().toISOString();
    var q = Math.floor(Number(d.Quantity) || 0);
    var t = Math.floor(Number(d.MinThreshold) || 0);

    var row = [
      id,
      sanitize_(d.ItemName),
      d.Category,
      sanitize_(d.SubCategory),
      sanitize_(d.Size),
      d.Unit || 'Nos',
      q, t,
      sanitize_(d.Location),
      sanitize_(d.TechnicalSpecs),
      now,
      sanitize_(d.Notes),
      computeStatus_(q, t)
    ];

    sh.appendRow(row);
    Logger.log('Item added: ' + id + ' - ' + d.ItemName);
    return { success: true, item: rowToObject_(row), message: 'Item added successfully!' };
  } catch (e) {
    Logger.log('addItem ERROR: ' + e.message);
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

function updateItem(d) {
  if (!d || !d.ID) return { success: false, message: 'No ID provided.' };
  var v = validateItemData_(d);
  if (!v.valid) return { success: false, message: v.message };

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var sh = getOrCreateSheet_();
    var rn = findRowById_(sh, d.ID);
    if (rn === -1) return { success: false, message: 'Item not found.' };

    var now = new Date().toISOString();
    var q = Math.floor(Number(d.Quantity) || 0);
    var t = Math.floor(Number(d.MinThreshold) || 0);

    var row = [
      d.ID,
      sanitize_(d.ItemName),
      d.Category,
      sanitize_(d.SubCategory),
      sanitize_(d.Size),
      d.Unit || 'Nos',
      q, t,
      sanitize_(d.Location),
      sanitize_(d.TechnicalSpecs),
      now,
      sanitize_(d.Notes),
      computeStatus_(q, t)
    ];

    sh.getRange(rn, 1, 1, SYSTEM_CONFIG.HEADERS.length).setValues([row]);
    Logger.log('Item updated: ' + d.ID);
    return { success: true, item: rowToObject_(row), message: 'Item updated successfully!' };
  } catch (e) {
    Logger.log('updateItem ERROR: ' + e.message);
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteItem(id) {
  if (!id) return { success: false, message: 'No ID provided.' };

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var sh = getOrCreateSheet_();
    var rn = findRowById_(sh, id);
    if (rn === -1) return { success: false, message: 'Item not found.' };

    sh.deleteRow(rn);
    Logger.log('Item deleted: ' + id);
    return { success: true, message: 'Item deleted successfully!' };
  } catch (e) {
    Logger.log('deleteItem ERROR: ' + e.message);
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

function adjustQuantity(id, delta) {
  if (!id) return { success: false, message: 'No ID provided.' };
  delta = Number(delta) || 0;
  if (delta === 0) return { success: false, message: 'Zero adjustment.' };

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var sh = getOrCreateSheet_();
    var rn = findRowById_(sh, id);
    if (rn === -1) return { success: false, message: 'Item not found.' };

    var cur = Number(sh.getRange(rn, 7).getValue()) || 0;
    var nq = Math.max(0, cur + delta);
    var t = Number(sh.getRange(rn, 8).getValue()) || 0;
    var st = computeStatus_(nq, t);

    sh.getRange(rn, 7).setValue(nq);
    sh.getRange(rn, 11).setValue(new Date().toISOString());
    sh.getRange(rn, 13).setValue(st);

    Logger.log('Qty adjusted: ' + id + ' | ' + cur + ' → ' + nq);
    return { success: true, newQuantity: nq, status: st };
  } catch (e) {
    Logger.log('adjustQuantity ERROR: ' + e.message);
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

function bulkAddItems(arr) {
  if (!arr || !arr.length) return { success: false, message: 'No items to add.' };

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var sh = getOrCreateSheet_();
    var now = new Date().toISOString();
    var used = {};

    var rows = arr.map(function(d) {
      var id;
      do { id = generateId_(); } while (used[id]);
      used[id] = true;

      var q = Math.floor(Number(d.Quantity) || 0);
      var t = Math.floor(Number(d.MinThreshold) || 0);

      return [
        id,
        sanitize_(d.ItemName),
        d.Category || '',
        sanitize_(d.SubCategory),
        sanitize_(d.Size),
        d.Unit || 'Nos',
        q, t,
        sanitize_(d.Location),
        sanitize_(d.TechnicalSpecs),
        now,
        sanitize_(d.Notes),
        computeStatus_(q, t)
      ];
    });

    if (rows.length) {
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, SYSTEM_CONFIG.HEADERS.length).setValues(rows);
    }

    Logger.log('Bulk added: ' + rows.length + ' items');
    return { success: true, count: rows.length, message: rows.length + ' items added successfully!' };
  } catch (e) {
    Logger.log('bulkAddItems ERROR: ' + e.message);
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

function getDashboardStats() {
  return buildStats_(getAllItems());
}


/* ================================================================
   📋 ISSUE SYSTEM
   ================================================================ */

function issueItem(d) {
  Logger.log('issueItem called with: ' + JSON.stringify(d));

  if (!d || typeof d !== 'object') return { success: false, message: 'Invalid data.' };
  if (!d.ItemID) return { success: false, message: 'No item selected.' };
  if (!d.IssuedTo || String(d.IssuedTo).trim().length < 2) return { success: false, message: 'Receiver name required (min 2 chars).' };

  var qtyToIssue = Math.floor(Number(d.QtyIssued) || 0);
  if (qtyToIssue <= 0) return { success: false, message: 'Quantity must be at least 1.' };
  if (!d.IssueDate) return { success: false, message: 'Issue date is required.' };

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    // 1) Find item in inventory
    var invSheet = getOrCreateSheet_();
    var rowNum = findRowById_(invSheet, d.ItemID);
    if (rowNum === -1) {
      return { success: false, message: 'Item not found in inventory (ID: ' + d.ItemID + ').' };
    }

    // 2) Check available stock
    var curQty = Number(invSheet.getRange(rowNum, 7).getValue()) || 0;
    if (qtyToIssue > curQty) {
      return { success: false, message: 'Insufficient stock! Available: ' + curQty + ', Requested: ' + qtyToIssue };
    }

    // 3) Deduct quantity from inventory
    var newQty = curQty - qtyToIssue;
    var thr = Number(invSheet.getRange(rowNum, 8).getValue()) || 0;
    var status = computeStatus_(newQty, thr);

    invSheet.getRange(rowNum, 7).setValue(newQty);
    invSheet.getRange(rowNum, 11).setValue(new Date().toISOString());
    invSheet.getRange(rowNum, 13).setValue(status);

    // 4) Create issue record
    var issSheet = getOrCreateIssueSheet_();
    var issId = generateIssueId_();
    var timestamp = new Date().toISOString();

    var issRow = [
      issId,
      String(d.ItemID),
      sanitize_(d.ItemName),
      String(d.Category || ''),
      sanitize_(d.SubCategory),
      sanitize_(d.Size),
      String(d.Unit || 'Nos'),
      qtyToIssue,
      sanitize_(d.IssuedTo),
      String(d.IssueDate),
      String(d.IssueTime || ''),
      sanitize_(d.IssueLocation),
      sanitize_(d.ProjectName),
      sanitize_(d.Description),
      sanitize_(d.IssuedBy),
      timestamp
    ];

    issSheet.appendRow(issRow);

    Logger.log('Issue SUCCESS: ' + issId + ' | Qty: ' + qtyToIssue + ' | To: ' + d.IssuedTo);

    return {
      success: true,
      issue: issueRowToObject_(issRow),
      newQuantity: newQty,
      status: status,
      message: qtyToIssue + ' x ' + (d.ItemName || 'item') + ' issued to ' + d.IssuedTo + ' successfully!'
    };

  } catch (e) {
    Logger.log('issueItem ERROR: ' + e.message);
    return { success: false, message: 'Issue failed: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

function getAllIssues() {
  try {
    var sh = getOrCreateIssueSheet_();
    var last = sh.getLastRow();
    if (last <= 1) return [];
    return sh.getRange(2, 1, last - 1, SYSTEM_CONFIG.ISSUE_HEADERS.length)
      .getValues()
      .filter(function(r) { return r[0] && String(r[0]).trim() !== ''; })
      .map(issueRowToObject_)
      .reverse();
  } catch (e) {
    Logger.log('getAllIssues ERROR: ' + e.message);
    return [];
  }
}

function getIssuesByItem(itemId) {
  return getAllIssues().filter(function(h) { return h.ItemID === itemId; });
}

function getIssueStats() {
  var issues = getAllIssues();
  var today = new Date().toISOString().split('T')[0];
  var todayIss = issues.filter(function(h) { return h.IssueDate === today; });

  var wk = new Date();
  wk.setDate(wk.getDate() - 7);
  var weekIss = issues.filter(function(h) { return h.IssueDate >= wk.toISOString().split('T')[0]; });

  return {
    totalIssues: issues.length,
    todayCount: todayIss.length,
    todayQty: todayIss.reduce(function(s, h) { return s + h.QtyIssued; }, 0),
    weekCount: weekIss.length,
    weekQty: weekIss.reduce(function(s, h) { return s + h.QtyIssued; }, 0),
    totalQtyIssued: issues.reduce(function(s, h) { return s + h.QtyIssued; }, 0)
  };
}

function deleteIssue(issId) {
  if (!issId) return { success: false, message: 'No Issue ID.' };

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var sh = getOrCreateIssueSheet_();
    var last = sh.getLastRow();
    if (last <= 1) return { success: false, message: 'No issues found.' };

    var d = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < d.length; i++) {
      if (String(d[i][0]).trim() === String(issId).trim()) {
        sh.deleteRow(i + 2);
        Logger.log('Issue deleted: ' + issId);
        return { success: true, message: 'Issue record deleted!' };
      }
    }
    return { success: false, message: 'Issue not found.' };
  } catch (e) {
    Logger.log('deleteIssue ERROR: ' + e.message);
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}


/* ================================================================
   📧 EMAIL SYSTEM
   ================================================================ */

function getEmailConfig() {
  try {
    var r = PropertiesService.getScriptProperties().getProperty('EMAIL_CONFIG');
    if (r) return JSON.parse(r);
    return {
      recipients: '',
      autoAlertEnabled: false,
      alertTime: '08:00',
      includeOkItems: false,
      lastSent: null
    };
  } catch (e) {
    return {
      recipients: '',
      autoAlertEnabled: false,
      alertTime: '08:00',
      includeOkItems: false,
      lastSent: null
    };
  }
}

function saveEmailConfig(c) {
  try {
    var recs = String(c.recipients || '').trim();
    if (recs) {
      var ea = recs.split(',');
      for (var i = 0; i < ea.length; i++) {
        if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(ea[i].trim())) {
          return { success: false, message: 'Invalid email: ' + ea[i].trim() };
        }
      }
    }

    var ec = {
      recipients: recs,
      autoAlertEnabled: !!c.autoAlertEnabled,
      alertTime: c.alertTime || '08:00',
      includeOkItems: !!c.includeOkItems,
      lastSent: c.lastSent || null
    };

    PropertiesService.getScriptProperties().setProperty('EMAIL_CONFIG', JSON.stringify(ec));

    if (ec.autoAlertEnabled && ec.recipients) {
      setupDailyTrigger_();
    } else {
      removeDailyTrigger_();
    }

    return { success: true, config: ec, message: 'Email settings saved!' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function sendLowStockAlert(rec) {
  try {
    var c = getEmailConfig();
    var to = rec || c.recipients;
    if (!to) return { success: false, message: 'No email address configured.' };

    var items = getAllItems();
    var low = items.filter(function(i) { return i.isLowStock; });
    if (!low.length) return { success: false, message: 'No low stock items!' };

    var html = '<h2 style="color:#b91c1c;">⚠️ Low Stock Alert</h2>';
    html += '<p style="color:#666;">Generated: ' + new Date().toLocaleString() + '</p>';
    html += '<table border="1" cellpadding="8" style="border-collapse:collapse;font-family:Arial;font-size:13px;width:100%;">';
    html += '<tr style="background:#7f1d1d;color:#fff;"><th>Item</th><th>Category</th><th>Location</th><th>Current Qty</th><th>Min Required</th><th>Status</th></tr>';

    low.forEach(function(i) {
      var out = i.Quantity <= 0;
      html += '<tr style="background:' + (out ? '#fee2e2' : '#fef3c7') + ';">';
      html += '<td><b>' + i.ItemName + '</b></td>';
      html += '<td>' + i.Category + '</td>';
      html += '<td>' + (i.Location || '-') + '</td>';
      html += '<td style="text-align:center;color:' + (out ? 'red' : '#b45309') + ';font-weight:bold;">' + i.Quantity + '</td>';
      html += '<td style="text-align:center;">' + i.MinThreshold + '</td>';
      html += '<td style="text-align:center;font-weight:bold;color:' + (out ? 'red' : '#b45309') + ';">' + (out ? '🔴 OUT' : '⚠️ LOW') + '</td>';
      html += '</tr>';
    });

    html += '</table>';
    html += '<p style="color:#999;font-size:11px;margin-top:16px;">Fire Protection & Detection IMS — Automated Alert</p>';

    MailApp.sendEmail({
      to: to,
      subject: '⚠️ Fire IMS — Low Stock Alert (' + low.length + ' items)',
      htmlBody: html,
      name: 'Fire Protection IMS'
    });

    c.lastSent = new Date().toISOString();
    PropertiesService.getScriptProperties().setProperty('EMAIL_CONFIG', JSON.stringify(c));

    return { success: true, message: 'Low stock alert sent to ' + to };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function sendInventoryReport(rec) {
  try {
    var c = getEmailConfig();
    var to = rec || c.recipients;
    if (!to) return { success: false, message: 'No email address configured.' };

    var items = getAllItems();
    var st = buildStats_(items);

    var html = '<h2 style="color:#1e3a5f;">📊 Inventory Report</h2>';
    html += '<p>Total Items: <b>' + st.totalItems + '</b> | Total Qty: <b>' + st.totalQuantity + '</b> | Low Stock: <b style="color:#b45309;">' + st.lowStockCount + '</b> | Out of Stock: <b style="color:red;">' + st.outOfStockCount + '</b></p>';
    html += '<table border="1" cellpadding="6" style="border-collapse:collapse;font-family:Arial;font-size:12px;width:100%;">';
    html += '<tr style="background:#1e293b;color:#fff;"><th>Item</th><th>Category</th><th>Size</th><th>Qty</th><th>Min</th><th>Unit</th><th>Location</th><th>Status</th></tr>';

    items.forEach(function(i) {
      var out = i.Quantity <= 0 && i.MinThreshold > 0;
      var low = i.isLowStock && !out;
      var bg = out ? '#fee2e2' : low ? '#fef3c7' : '#fff';
      html += '<tr style="background:' + bg + ';">';
      html += '<td>' + i.ItemName + '</td><td>' + i.Category + '</td><td>' + (i.Size || '-') + '</td>';
      html += '<td style="text-align:center;font-weight:bold;color:' + (out ? 'red' : low ? '#b45309' : '#000') + ';">' + i.Quantity + '</td>';
      html += '<td style="text-align:center;">' + i.MinThreshold + '</td>';
      html += '<td>' + i.Unit + '</td><td>' + (i.Location || '-') + '</td>';
      html += '<td style="text-align:center;">' + (out ? '🔴 OUT' : low ? '⚠️ LOW' : '✅ OK') + '</td></tr>';
    });

    html += '</table>';
    html += '<p style="color:#999;font-size:11px;margin-top:16px;">Generated: ' + new Date().toLocaleString() + '</p>';

    MailApp.sendEmail({
      to: to,
      subject: '📊 Fire IMS — Full Inventory Report',
      htmlBody: html,
      name: 'Fire Protection IMS'
    });

    return { success: true, message: 'Inventory report sent to ' + to };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function sendTestEmail(em) {
  try {
    if (!em || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(em.trim())) {
      return { success: false, message: 'Invalid email address.' };
    }
    MailApp.sendEmail({
      to: em.trim(),
      subject: '✅ Fire IMS — Test Email OK',
      htmlBody: '<h2 style="color:green;">✅ Email Configuration Working!</h2><p>This is a test email from Fire Protection & Detection IMS.</p><p>Recipient: ' + em + '</p><p style="color:#999;font-size:11px;">Sent: ' + new Date().toLocaleString() + '</p>',
      name: 'Fire Protection IMS'
    });
    return { success: true, message: 'Test email sent to ' + em };
  } catch (e) {
    return { success: false, message: e.message };
  }
}


/* ================================================================
   AUTO TRIGGER
   ================================================================ */

function autoLowStockAlert() {
  try {
    var c = getEmailConfig();
    if (c.autoAlertEnabled && c.recipients) {
      var low = getAllItems().filter(function(i) { return i.isLowStock; });
      if (low.length) sendLowStockAlert();
    }
  } catch (e) {
    Logger.log('autoLowStockAlert ERROR: ' + e.message);
  }
}

function setupDailyTrigger_() {
  removeDailyTrigger_();
  var hour = parseInt((getEmailConfig().alertTime || '08:00').split(':')[0]) || 8;
  ScriptApp.newTrigger('autoLowStockAlert')
    .timeBased()
    .everyDays(1)
    .atHour(hour)
    .create();
  Logger.log('Daily trigger set for hour: ' + hour);
}

function removeDailyTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'autoLowStockAlert') {
      ScriptApp.deleteTrigger(t);
    }
  });
}


/* ================================================================
   🌱 SEED SAMPLE DATA
   ================================================================ */

function seedSampleData() {
  return bulkAddItems([
    { ItemName:'Alarm Check Valve', Category:'Valve Station', SubCategory:'Alarm Valve', Size:'4"', Unit:'Nos', Quantity:3, MinThreshold:1, Location:'Warehouse-A', TechnicalSpecs:'UL/FM 175PSI', Notes:'' },
    { ItemName:'OS&Y Gate Valve', Category:'Valve Station', SubCategory:'Gate Valve', Size:'4"', Unit:'Nos', Quantity:5, MinThreshold:2, Location:'Warehouse-A', TechnicalSpecs:'UL 200PSI', Notes:'' },
    { ItemName:'Butterfly Valve', Category:'Valve Station', SubCategory:'Butterfly Valve', Size:'4"', Unit:'Nos', Quantity:6, MinThreshold:2, Location:'Warehouse-A', TechnicalSpecs:'UL, Gear', Notes:'' },
    { ItemName:'Pressure Gauge 300PSI', Category:'Valve Station', SubCategory:'Pressure Gauge', Size:'1/4"', Unit:'Nos', Quantity:8, MinThreshold:3, Location:'Warehouse-B', TechnicalSpecs:'Glycerin', Notes:'' },
    { ItemName:'Smoke Detector', Category:'Fire Detection & Alarm', SubCategory:'Smoke Detector', Size:'Standard', Unit:'Nos', Quantity:45, MinThreshold:10, Location:'Warehouse-B', TechnicalSpecs:'Addressable 24VDC', Notes:'' },
    { ItemName:'Heat Detector ROR', Category:'Fire Detection & Alarm', SubCategory:'Heat Detector', Size:'Standard', Unit:'Nos', Quantity:30, MinThreshold:8, Location:'Warehouse-B', TechnicalSpecs:'57C+ROR', Notes:'' },
    { ItemName:'Manual Call Point', Category:'Fire Detection & Alarm', SubCategory:'Manual Call Point', Size:'Standard', Unit:'Nos', Quantity:12, MinThreshold:5, Location:'Warehouse-B', TechnicalSpecs:'Break-glass', Notes:'' },
    { ItemName:'Fire Alarm Panel', Category:'Fire Detection & Alarm', SubCategory:'Fire Alarm Control Panel', Size:'N/A', Unit:'Nos', Quantity:2, MinThreshold:1, Location:'Warehouse-B', TechnicalSpecs:'2-Loop 256pts', Notes:'' },
    { ItemName:'Sounder Red', Category:'Fire Detection & Alarm', SubCategory:'Sounder', Size:'Standard', Unit:'Nos', Quantity:20, MinThreshold:5, Location:'Warehouse-B', TechnicalSpecs:'100dB 24VDC', Notes:'' },
    { ItemName:'Detector Base', Category:'Fire Detection & Alarm', SubCategory:'Detector Base', Size:'Standard', Unit:'Nos', Quantity:50, MinThreshold:15, Location:'Warehouse-B', TechnicalSpecs:'Loop in/out', Notes:'' },
    { ItemName:'Fire Cable 1.5mm', Category:'Fire Detection & Alarm', SubCategory:'Cable - Fire Rated', Size:'1.5mm', Unit:'Rolls', Quantity:8, MinThreshold:3, Location:'Cable Store', TechnicalSpecs:'2hr Red 100m', Notes:'' },
    { ItemName:'BS Pipe 4"', Category:'Suppression & Piping', SubCategory:'Black Steel Pipe', Size:'4"', Unit:'Meters', Quantity:120, MinThreshold:30, Location:'Pipe Yard', TechnicalSpecs:'ASTM A53', Notes:'' },
    { ItemName:'BS Pipe 2"', Category:'Suppression & Piping', SubCategory:'Black Steel Pipe', Size:'2"', Unit:'Meters', Quantity:200, MinThreshold:50, Location:'Pipe Yard', TechnicalSpecs:'ASTM A53', Notes:'' },
    { ItemName:'Pendent Sprinkler K80', Category:'Suppression & Piping', SubCategory:'Sprinkler Head - Pendent', Size:'1/2"', Unit:'Nos', Quantity:80, MinThreshold:20, Location:'Warehouse-A', TechnicalSpecs:'K80 68C Chrome', Notes:'' },
    { ItemName:'Flow Switch 4"', Category:'Suppression & Piping', SubCategory:'Flow Switch', Size:'4"', Unit:'Nos', Quantity:5, MinThreshold:2, Location:'Warehouse-A', TechnicalSpecs:'Vane DPDT UL', Notes:'' },
    { ItemName:'Grooved Elbow 90 4"', Category:'Fittings & Joints', SubCategory:'Elbow 90°', Size:'4"', Unit:'Nos', Quantity:25, MinThreshold:10, Location:'Warehouse-A', TechnicalSpecs:'Ductile Iron', Notes:'' },
    { ItemName:'Rigid Coupling 4"', Category:'Fittings & Joints', SubCategory:'Coupling', Size:'4"', Unit:'Nos', Quantity:50, MinThreshold:15, Location:'Warehouse-A', TechnicalSpecs:'Victaulic', Notes:'' },
    { ItemName:'Mechanical Tee 4x1"', Category:'Fittings & Joints', SubCategory:'Mechanical Tee', Size:'4"', Unit:'Nos', Quantity:20, MinThreshold:8, Location:'Warehouse-A', TechnicalSpecs:'Grooved UL', Notes:'' },
    { ItemName:'Anchor Bolt M10', Category:'Support & Hardware', SubCategory:'Anchor Bolt', Size:'M10', Unit:'Nos', Quantity:5, MinThreshold:50, Location:'Warehouse-B', TechnicalSpecs:'Hilti SS304', Notes:'CRITICAL' },
    { ItemName:'Anchor Bolt M12', Category:'Support & Hardware', SubCategory:'Anchor Bolt', Size:'M12', Unit:'Nos', Quantity:30, MinThreshold:60, Location:'Warehouse-B', TechnicalSpecs:'Wedge HDG', Notes:'' },
    { ItemName:'U-Clamp 4"', Category:'Support & Hardware', SubCategory:'U-Clamp', Size:'4"', Unit:'Nos', Quantity:15, MinThreshold:30, Location:'Warehouse-B', TechnicalSpecs:'HDG', Notes:'' },
    { ItemName:'Hex Nut M10 Box', Category:'Support & Hardware', SubCategory:'Hex Nut', Size:'M10', Unit:'Boxes', Quantity:1, MinThreshold:3, Location:'Warehouse-B', TechnicalSpecs:'100/box', Notes:'LOW' },
    { ItemName:'Gasket EPDM 4"', Category:'Support & Hardware', SubCategory:'Gasket', Size:'4"', Unit:'Nos', Quantity:10, MinThreshold:25, Location:'Warehouse-B', TechnicalSpecs:'EPDM Grade E', Notes:'' },
    { ItemName:'FM200 Cylinder 180L', Category:'Gas Suppression', SubCategory:'FM200 Cylinder', Size:'180L', Unit:'Nos', Quantity:2, MinThreshold:1, Location:'Gas Store', TechnicalSpecs:'HFC-227ea 42bar', Notes:'' },
    { ItemName:'Gas Release Panel', Category:'Gas Suppression', SubCategory:'Release Panel', Size:'N/A', Unit:'Nos', Quantity:2, MinThreshold:1, Location:'Gas Store', TechnicalSpecs:'EN12094', Notes:'' },
    { ItemName:'ABC Extinguisher 6kg', Category:'Fire Extinguishers', SubCategory:'ABC Dry Powder', Size:'6kg', Unit:'Nos', Quantity:15, MinThreshold:5, Location:'Warehouse-C', TechnicalSpecs:'MAP', Notes:'' },
    { ItemName:'CO2 Extinguisher 5kg', Category:'Fire Extinguishers', SubCategory:'CO2 Extinguisher', Size:'5kg', Unit:'Nos', Quantity:8, MinThreshold:3, Location:'Warehouse-C', TechnicalSpecs:'Alloy steel', Notes:'' },
    { ItemName:'Emergency Light', Category:'Emergency & Safety', SubCategory:'Emergency Light', Size:'Standard', Unit:'Nos', Quantity:12, MinThreshold:4, Location:'Warehouse-C', TechnicalSpecs:'LED 3hr', Notes:'' },
    { ItemName:'Exit Sign LED', Category:'Emergency & Safety', SubCategory:'Exit Sign - Illuminated', Size:'Standard', Unit:'Nos', Quantity:10, MinThreshold:4, Location:'Warehouse-C', TechnicalSpecs:'Battery backup', Notes:'' },
    { ItemName:'Fire Sealant', Category:'Emergency & Safety', SubCategory:'Fire Rated Sealant', Size:'310ml', Unit:'Cans', Quantity:20, MinThreshold:8, Location:'Warehouse-C', TechnicalSpecs:'Intumescent 4hr', Notes:'' }
  ]);
}
