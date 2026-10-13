/**
 * ====================================================================
 * PATIL GROUP CRM - CENTRALIZED CLOUD BACKEND (GOOGLE APPS SCRIPT)
 * ====================================================================
 * Single reliable source of truth for all links (patilgroup.com/crm,
 * Vercel app), laptops, and sales reps.
 * 
 * FEATURES:
 * 1. REAL-TIME STORAGE ENGINE:
 *    - Chunked key-value storage engine (<40,000 chars per cell).
 *    - Human-readable tabs: Orders, Users, Coordination Logs, Lost Orders, Sales Rep Activity.
 * 2. MULTI-LAPTOP CENTRALIZATION:
 *    - Atomic lead upsert (action=save_order): Prevents overwriting leads from other laptops.
 *    - Smart-merge on full array set: Merges local and remote entries by ID.
 *    - Dynamic column expansion: Guarantees no bounds crashes (>26 columns).
 *    - Fast batch retrieval (action=get_all): Fetches all tabs in 1 request.
 * 3. TRANSPARENCY:
 *    - Returns connected spreadsheet URL & ID in responses so you can view the active sheet.
 * ====================================================================
 */

// OPTIONAL: If your Apps Script is standalone (not created via Extensions > Apps Script inside your Google Sheet),
// you can paste your target Google Sheet ID below:
// Example: var HARDCODED_SPREADSHEET_ID = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms";
var HARDCODED_SPREADSHEET_ID = "";

/* ================= WEB APP ROUTING ================= */

function doGet(e) {
  try {
    var params = e ? e.parameter : {};
    var action = params.action;
    var key = params.key;

    // Trigger fresh start via URL: ?action=start_fresh
    if (action === 'start_fresh' || action === 'clear_all') {
      var resetRes = startFreshWipeAllData();
      return responseJSON(resetRes);
    }

    var ss = getSpreadsheet();
    var storageSheet = getOrCreateSheet(ss, '_Storage');

    // Fast batch retrieval for all CRM data across all tabs in one single request
    if (action === 'get_all' || key === 'all') {
      var allOrders = getStoredValue(storageSheet, 'crm:orders') || [];
      return responseJSON({
        status: 'success',
        users: getStoredValue(storageSheet, 'crm:users') || getAuthorizedUsersList(),
        orders: allOrders,
        coordination: getStoredValue(storageSheet, 'crm:coordination') || [],
        lostOrders: getStoredValue(storageSheet, 'crm:lostorders') || [],
        activity: getStoredValue(storageSheet, 'crm:activity') || {},
        wipeEpoch: getStoredValue(storageSheet, 'crm:wipe_epoch') || 0,
        spreadsheetUrl: ss.getUrl(),
        spreadsheetId: ss.getId(),
        spreadsheetName: ss.getName(),
        ordersCount: Array.isArray(allOrders) ? allOrders.length : 0
      });
    }

    // Health check and sheet info
    if (!key || action === 'info') {
      var infoOrders = getStoredValue(storageSheet, 'crm:orders') || [];
      return responseJSON({
        status: 'online',
        service: 'Patil Group CRM Sync API (Centralized Engine)',
        spreadsheetUrl: ss.getUrl(),
        spreadsheetId: ss.getId(),
        spreadsheetName: ss.getName(),
        ordersCount: Array.isArray(infoOrders) ? infoOrders.length : 0,
        timestamp: new Date().toISOString()
      });
    }

    var val = getStoredValue(storageSheet, key);
    return responseJSON({
      status: 'success',
      key: key,
      value: val !== null ? val : [],
      spreadsheetUrl: ss.getUrl()
    });
  } catch (err) {
    return responseJSON({ status: 'error', message: err.toString() });
  }
}

function doPost(e) {
  try {
    var postData = {};
    if (e && e.postData && e.postData.contents) {
      try {
        postData = JSON.parse(e.postData.contents);
      } catch (pErr) {
        postData = {};
      }
    }

    var action = postData.action || (e && e.parameter ? e.parameter.action : 'set');
    var key = postData.key || (e && e.parameter ? e.parameter.key : null);
    var value = postData.value;
    if (value === undefined && e && e.parameter && e.parameter.value) {
      try {
        value = JSON.parse(e.parameter.value);
      } catch (pvErr) {
        value = e.parameter.value;
      }
    }

    var ss = getSpreadsheet();
    var storageSheet = getOrCreateSheet(ss, '_Storage');

    // 1. FRESH START / CLEAR ALL
    if (action === 'start_fresh' || action === 'clear_all') {
      var freshRes = startFreshWipeAllData();
      return responseJSON(freshRes);
    }

    // 2. ATOMIC LEAD UPSERT (Single Lead Add/Edit)
    // Ensures leads added from one laptop never overwrite leads entered on another laptop
    if (action === 'save_order' || action === 'upsert_order') {
      var singleOrder = postData.order;
      if (!singleOrder && e && e.parameter && e.parameter.order) {
        try { singleOrder = JSON.parse(e.parameter.order); } catch (oe) { singleOrder = null; }
      }
      if (!singleOrder || !singleOrder.id) {
        return responseJSON({ status: 'error', message: 'Missing order data or order.id' });
      }

      var currentOrders = getStoredValue(storageSheet, 'crm:orders') || [];
      if (!Array.isArray(currentOrders)) currentOrders = [];

      var foundIdx = -1;
      for (var oi = 0; oi < currentOrders.length; oi++) {
        if (currentOrders[oi] && currentOrders[oi].id === singleOrder.id) {
          foundIdx = oi;
          break;
        }
      }

      if (foundIdx >= 0) {
        currentOrders[foundIdx] = singleOrder;
      } else {
        currentOrders.unshift(singleOrder); // Prepend so latest lead appears first
      }

      setStoredValue(storageSheet, 'crm:is_wiped', false);
      setStoredValue(storageSheet, 'crm:orders', currentOrders);
      updateOrdersSheet(ss, currentOrders);

      return responseJSON({
        status: 'success',
        action: 'save_order',
        orderId: singleOrder.id,
        count: currentOrders.length,
        orders: currentOrders,
        spreadsheetUrl: ss.getUrl()
      });
    }

    // 3. ATOMIC LEAD DELETION
    if (action === 'delete_order') {
      var delId = postData.id || postData.orderId || (e && e.parameter ? (e.parameter.id || e.parameter.orderId) : null);
      if (!delId) {
        return responseJSON({ status: 'error', message: 'Missing order id to delete' });
      }

      var curOrdersForDel = getStoredValue(storageSheet, 'crm:orders') || [];
      if (!Array.isArray(curOrdersForDel)) curOrdersForDel = [];

      var filteredOrders = curOrdersForDel.filter(function(o) { return o && o.id !== delId; });
      setStoredValue(storageSheet, 'crm:orders', filteredOrders);
      updateOrdersSheet(ss, filteredOrders);

      return responseJSON({
        status: 'success',
        action: 'delete_order',
        deletedId: delId,
        count: filteredOrders.length,
        orders: filteredOrders,
        spreadsheetUrl: ss.getUrl()
      });
    }

    if (!key) {
      return responseJSON({ status: 'error', message: 'Missing key parameter' });
    }

    // 4. STORAGE KEY DELETION
    if (action === 'delete' && value === undefined) {
      deleteStoredKey(storageSheet, key);
      return responseJSON({ status: 'success', action: 'delete', key: key });
    }

    // 5. SMART-MERGE FOR ORDERS (Prevents multi-laptop race conditions)
    var finalVal = value;
    if (key === 'crm:orders' && Array.isArray(value)) {
      var isWiped = getStoredValue(storageSheet, 'crm:is_wiped');
      if (isWiped === true) {
        // Database was wiped: prevent client background loops from resurrecting old orders
        finalVal = [];
      } else {
        var isReplace = (postData.mode === 'replace' || (e && e.parameter && e.parameter.mode === 'replace'));
        if (!isReplace) {
          var existingInSheet = getStoredValue(storageSheet, 'crm:orders') || [];
          if (Array.isArray(existingInSheet) && existingInSheet.length > 0) {
            var orderMap = {};
            // Keep all existing orders from sheet
            for (var ei = 0; ei < existingInSheet.length; ei++) {
              if (existingInSheet[ei] && existingInSheet[ei].id) {
                orderMap[existingInSheet[ei].id] = existingInSheet[ei];
              }
            }
            // Merge/update with incoming orders from this client
            for (var vi = 0; vi < value.length; vi++) {
              if (value[vi] && value[vi].id) {
                orderMap[value[vi].id] = value[vi];
              }
            }
            finalVal = Object.keys(orderMap).map(function(k) { return orderMap[k]; });
          }
        }
      }
    }

    // Save data into chunked key-value storage engine
    setStoredValue(storageSheet, key, finalVal);

    // Format and synchronize data into human-readable tabs
    syncToReadableTabs(ss, key, finalVal);

    return responseJSON({
      status: 'success',
      key: key,
      count: Array.isArray(finalVal) ? finalVal.length : null,
      orders: key === 'crm:orders' ? finalVal : undefined,
      spreadsheetUrl: ss.getUrl()
    });
  } catch (err) {
    return responseJSON({ status: 'error', message: err.toString() });
  }
}

/* ================= 1-CLICK FRESH START: WIPE ALL OLD ROWS ================= */
function startFreshWipeAllData() {
  var ss = getSpreadsheet();
  var allowedSheets = ['Orders', 'Users', 'Coordination Logs', 'Lost Orders', 'Sales Rep Activity', '_Storage'];

  // 1. Ensure allowed sheets exist
  for (var i = 0; i < allowedSheets.length; i++) {
    getOrCreateSheet(ss, allowedSheets[i]);
  }

  // 2. Delete any extra/unwanted sheets (e.g. Sheet1, old test sheets)
  var allSheets = ss.getSheets();
  for (var j = 0; j < allSheets.length; j++) {
    var sheet = allSheets[j];
    var sheetName = sheet.getName();
    if (allowedSheets.indexOf(sheetName) === -1) {
      if (ss.getSheets().length > 1) {
        ss.deleteSheet(sheet);
      }
    }
  }

  // 3. Reset _Storage engine to clean empty arrays
  var storageSheet = ss.getSheetByName('_Storage');
  var defaultUsers = getAuthorizedUsersList();
  var wipeEpoch = new Date().getTime();

  setStoredValue(storageSheet, 'crm:wipe_epoch', wipeEpoch);
  setStoredValue(storageSheet, 'crm:is_wiped', true);
  setStoredValue(storageSheet, 'crm:orders', []);
  setStoredValue(storageSheet, 'crm:lostorders', []);
  setStoredValue(storageSheet, 'crm:coordination', []);
  setStoredValue(storageSheet, 'crm:activity', {});
  setStoredValue(storageSheet, 'crm:users', defaultUsers);

  // 4. Reset human-readable tabs: clear all rows & re-create headers only
  updateOrdersSheet(ss, []);
  updateLostOrdersSheet(ss, []);
  updateCoordinationSheet(ss, []);
  updateActivitySheet(ss, {});
  updateUsersSheet(ss, defaultUsers);

  // 5. Order tabs neatly and set active sheet to Orders
  for (var m = 0; m < allowedSheets.length; m++) {
    var s = ss.getSheetByName(allowedSheets[m]);
    if (s) {
      ss.setActiveSheet(s);
      ss.moveActiveSheet(m + 1);
    }
  }
  ss.setActiveSheet(ss.getSheetByName('Orders'));

  Logger.log('Success: All old rows deleted. Google Sheets is now 100% fresh.');
  return {
    status: 'success',
    message: 'All old rows permanently deleted. Ready for fresh data entry.',
    ordersCount: 0,
    spreadsheetUrl: ss.getUrl()
  };
}

/* ================= SPREADSHEET HELPER ================= */

function getSpreadsheet() {
  if (HARDCODED_SPREADSHEET_ID && HARDCODED_SPREADSHEET_ID.trim().length > 10) {
    try {
      return SpreadsheetApp.openById(HARDCODED_SPREADSHEET_ID.trim());
    } catch (e) {
      Logger.log("Could not open HARDCODED_SPREADSHEET_ID: " + e.toString());
    }
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) return ss;
  } catch (e) {}

  var props = PropertiesService.getScriptProperties();
  var ssId = props.getProperty('SPREADSHEET_ID');

  if (ssId) {
    try {
      return SpreadsheetApp.openById(ssId);
    } catch (e) {}
  }

  var newSs = SpreadsheetApp.create('Patil Group CRM Data');
  props.setProperty('SPREADSHEET_ID', newSs.getId());
  return newSs;
}

/* ================= GRID BOUNDS PROTECTION ================= */

/**
 * Automatically expands sheet columns and rows to prevent Google Sheets
 * coordinates outside bounds exceptions (e.g. 34 columns on default 26-column sheets).
 */
function ensureSheetDimensions(sheet, minRows, minCols) {
  var maxCols = sheet.getMaxColumns();
  if (maxCols < minCols) {
    sheet.insertColumnsAfter(maxCols, minCols - maxCols);
  }
  var maxRows = sheet.getMaxRows();
  if (maxRows < minRows) {
    sheet.insertRowsAfter(maxRows, minRows - maxRows);
  }
}

/* ================= KEY-VALUE STORAGE ENGINE ================= */

function getStoredValue(sheet, key) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      var raw = data[i].slice(1).join('');
      if (!raw) return null;
      try {
        return JSON.parse(raw);
      } catch (e) {
        return raw;
      }
    }
  }
  return null;
}

function setStoredValue(sheet, key, value) {
  var strValue = (typeof value === 'string') ? value : JSON.stringify(value);
  var chunkSize = 40000;
  var chunks = [];
  for (var c = 0; c < strValue.length; c += chunkSize) {
    chunks.push(strValue.substring(c, c + chunkSize));
  }
  if (chunks.length === 0) chunks.push('');

  var data = sheet.getDataRange().getValues();
  if (data.length === 0 || data[0][0] !== 'Key') {
    sheet.clear();
    ensureSheetDimensions(sheet, 10, chunks.length + 3);
    sheet.appendRow(['Key', 'Chunk_1']);
    data = sheet.getDataRange().getValues();
  }

  var requiredColumns = chunks.length + 2;
  ensureSheetDimensions(sheet, Math.max(10, sheet.getLastRow() + 2), requiredColumns);

  var rowToUpdate = -1;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      rowToUpdate = i + 1;
      break;
    }
  }

  if (rowToUpdate === -1) {
    rowToUpdate = sheet.getLastRow() + 1;
    sheet.getRange(rowToUpdate, 1).setValue(key);
  } else {
    var lastCol = Math.max(1, sheet.getLastColumn() - 1);
    sheet.getRange(rowToUpdate, 2, 1, lastCol).clearContent();
  }

  sheet.getRange(rowToUpdate, 2, 1, chunks.length).setValues([chunks]);
}

function deleteStoredKey(sheet, key) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
}

/* ================= HUMAN-READABLE TAB SYNC ================= */

function syncToReadableTabs(ss, key, value) {
  try {
    if (key === 'crm:users' && Array.isArray(value)) {
      updateUsersSheet(ss, value);
    } else if (key === 'crm:orders' && Array.isArray(value)) {
      updateOrdersSheet(ss, value);
    } else if (key === 'crm:coordination' && Array.isArray(value)) {
      updateCoordinationSheet(ss, value);
    } else if (key === 'crm:lostorders' && Array.isArray(value)) {
      updateLostOrdersSheet(ss, value);
    } else if (key === 'crm:activity' && typeof value === 'object' && value !== null) {
      updateActivitySheet(ss, value);
    }
  } catch (e) {
    Logger.log('Sync warning (' + key + '): ' + e.toString());
  }
}

/* ================= 1. USERS TAB ================= */

function updateUsersSheet(ss, users) {
  var sheet = getOrCreateSheet(ss, 'Users');
  sheet.clear();

  var headers = [
    "Username / Email",
    "Full Name",
    "Role",
    "Role Title",
    "Access / Auth Status",
    "Last Synced Date"
  ];
  var rows = [headers];
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT+5:30", "yyyy-MM-dd HH:mm");

  for (var i = 0; i < users.length; i++) {
    var u = users[i];
    var status = "Google OAuth Authorized";

    if (u.role === 'ceo') {
      status = 'Protected (Admin)';
    } else if (u.passwordDeleted) {
      status = 'Access Removed / Deleted';
    } else if (u.mustChangePassword) {
      status = 'Password Reset Required';
    }

    rows.push([
      u.username || '',
      u.name || '',
      u.role === 'ceo' ? 'CEO / Manager' : 'Sales Rep',
      u.roleTitle || (u.role === 'ceo' ? 'CEO / Manager' : 'Sales Rep'),
      status,
      today
    ]);
  }

  ensureSheetDimensions(sheet, rows.length + 5, headers.length + 2);
  sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
  formatHeaderRow(sheet, headers.length);
  autoFormatTable(sheet, rows.length, headers.length);
}

/* ================= 2. ORDERS TAB (CENTRALIZED) ================= */

function updateOrdersSheet(ss, orders) {
  var sheet = getOrCreateSheet(ss, 'Orders');
  sheet.clear();

  var headers = [
    "Offer / Order ID",
    "Lead Received Date",
    "Client Name / Company",
    "Client POC",
    "POC Designation",
    "Client Mobile",
    "Client Email",
    "Product(s) of Interest",
    "Product Other Desc",
    "Project Name & Description",
    "Project Category / Type",
    "Project Other Desc",
    "Location / Region",
    "Estimated Value (INR CR)",
    "Priority",
    "Current Deal Stage",
    "Stage Last Changed Date",
    "Sales Rep",
    "Created By (Email)",
    "Last Updated By (Email)",
    "Last Updated Timestamp",
    "Next Follow-up Action",
    "Type of Action (Dependency)",
    "Action Planned Start Date",
    "Action Planned End Date",
    "Next Follow-up Date",
    "Delay Status / Overdue Days",
    "Latest Action Delay Reason",
    "Latest Action Delay Type",
    "Lead Source Type",
    "Lead Source Name",
    "Lead Source Detail / Path",
    "Total Offer Cycle Time (days)",
    "Closed / Decision Date",
    "Reason Won / Lost",
    "Latest Update Note",
    "Updates Count"
  ];

  var rows = [headers];
  var todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT+5:30", "yyyy-MM-dd");

  for (var i = 0; i < orders.length; i++) {
    var o = orders[i];

    var actionType = '';
    var actionHist = o.actionHistory || [];
    if (actionHist.length > 0) {
      var lastAct = actionHist[actionHist.length - 1];
      actionType = lastAct.actionType || '';
    }
    if (!actionType && o.actionType) {
      actionType = o.actionType;
    }

    var delayReason = '';
    var delayType = '';
    for (var h = actionHist.length - 1; h >= 0; h--) {
      if (actionHist[h].delayReason || actionHist[h].delayType) {
        delayReason = actionHist[h].delayReason || '';
        delayType = actionHist[h].delayType || '';
        break;
      }
    }

    var overdueDays = 0;
    var isClosed = o.dealStage && (o.dealStage.indexOf('9.') === 0 || o.dealStage.indexOf('10.') === 0);
    if (!isClosed) {
      if (o.actionEndDate && o.actionEndDate < todayStr) {
        overdueDays = Math.max(overdueDays, computeDaysBetween(o.actionEndDate, todayStr));
      }
      if (o.followUpDate && o.followUpDate < todayStr) {
        overdueDays = Math.max(overdueDays, computeDaysBetween(o.followUpDate, todayStr));
      }
    }
    var delayStatus = isClosed ? 'Closed' : (overdueDays > 0 ? ('Overdue (' + overdueDays + 'd)') : 'On Track');

    var sourceDetail = '';
    if (o.leadSourceName === 'Physical Visit / Scouting' && Array.isArray(o.leadSourceVisitPath) && o.leadSourceVisitPath.length > 0) {
      sourceDetail = 'Physical Visit > ' + o.leadSourceVisitPath.join(' > ');
    } else if (o.leadSourceName === 'Others') {
      sourceDetail = o.leadSourceOther ? ('Others: ' + o.leadSourceOther) : 'Others';
    } else {
      sourceDetail = o.leadSourceName || '';
    }

    var cycleEnd = o.closedDate || todayStr;
    var cycleDays = (o.dateLogged && cycleEnd) ? computeDaysBetween(o.dateLogged, cycleEnd) : '';

    var latestUpdate = '';
    var updatesCount = 0;
    if (Array.isArray(o.updates) && o.updates.length > 0) {
      updatesCount = o.updates.length;
      var lastUp = o.updates[o.updates.length - 1];
      latestUpdate = (lastUp.date || '') + (lastUp.by ? ' (' + lastUp.by + ')' : '') + ': ' + (lastUp.text || '');
    }

    rows.push([
      o.id || '',
      o.dateLogged || '',
      o.clientName || '',
      o.clientPOC || '',
      o.clientPOCDesignation || '',
      o.clientMobile || '',
      o.clientEmail || '',
      o.product || '',
      o.productOtherDesc || '',
      o.projectName || '',
      o.projectType || '',
      o.projectOtherDesc || '',
      o.location || '',
      o.estimatedValue !== undefined ? o.estimatedValue : '',
      o.priority || '',
      o.dealStage || '',
      o.stageChangedDate || '',
      o.salesPerson || '',
      o.createdBy || '',
      o.lastUpdatedBy || '',
      o.lastUpdatedAt || '',
      o.nextAction || '',
      actionType,
      o.actionStartDate || '',
      o.actionEndDate || '',
      o.followUpDate || '',
      delayStatus,
      delayReason,
      delayType,
      o.leadSourceType || '',
      o.leadSourceName || '',
      sourceDetail,
      cycleDays,
      o.closedDate || '',
      o.reasonLostWon || '',
      latestUpdate,
      updatesCount
    ]);
  }

  // Ensure sheet has at least 36 columns so coordinates never crash on default 26 columns
  ensureSheetDimensions(sheet, Math.max(20, rows.length + 5), Math.max(36, headers.length));
  sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
  formatHeaderRow(sheet, headers.length);
  autoFormatTable(sheet, rows.length, headers.length);
}

/* ================= 3. COORDINATION LOGS TAB ================= */

function updateCoordinationSheet(ss, logs) {
  var sheet = getOrCreateSheet(ss, 'Coordination Logs');
  sheet.clear();

  var headers = [
    "Log ID",
    "Date Logged",
    "Linked Offer ID",
    "Client Name",
    "Project Name",
    "Sales Rep / Assigned Email",
    "Created By (Email)",
    "Category / Topic",
    "Technical Solution / Activity",
    "Plant Fulfillment Status",
    "Support Needed from Management",
    "Target Completion Date",
    "Support Completion Date"
  ];

  var rows = [headers];
  for (var i = 0; i < logs.length; i++) {
    var l = logs[i];
    rows.push([
      l.id || '',
      l.dateLogged || l.date || '',
      l.orderId || '',
      l.clientName || '',
      l.projectName || '',
      l.salesPerson || '',
      l.createdBy || '',
      l.category || '',
      l.activity || l.details || '',
      l.fulfillment || '',
      l.support || '',
      l.targetDate || '',
      l.supportDate || ''
    ]);
  }

  ensureSheetDimensions(sheet, Math.max(20, rows.length + 5), Math.max(16, headers.length));
  sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
  formatHeaderRow(sheet, headers.length);
  autoFormatTable(sheet, rows.length, headers.length);
}

/* ================= 4. LOST ORDERS TAB ================= */

function updateLostOrdersSheet(ss, lost) {
  var sheet = getOrCreateSheet(ss, 'Lost Orders');
  sheet.clear();

  var headers = [
    "Entry ID",
    "Linked Offer ID",
    "Date Analysed",
    "Client Name",
    "Project Name",
    "Estimated Value (INR CR)",
    "Sales Rep / Assigned Email",
    "Created By (Email)",
    "Last Updated By (Email)",
    "Competitor Who Won",
    "Implementation Status",
    "Primary Root Cause",
    "Plant / Production Improvements Required",
    "Action Owner",
    "Target Start Date",
    "Target End Date",
    "Total Planned Days",
    "Actual End Date",
    "Latest Improvement Note"
  ];

  var rows = [headers];
  for (var i = 0; i < lost.length; i++) {
    var l = lost[i];

    var latestNote = '';
    if (Array.isArray(l.notes) && l.notes.length > 0) {
      var lastN = l.notes[l.notes.length - 1];
      latestNote = (lastN.date || '') + (lastN.by ? ' (' + lastN.by + ')' : '') + ': ' + (lastN.text || '');
    }

    rows.push([
      l.id || '',
      l.linkedOrderId || '',
      l.dateAnalysed || '',
      l.clientName || '',
      l.projectName || '',
      l.estimatedValue !== undefined ? l.estimatedValue : '',
      l.salesPerson || '',
      l.createdBy || '',
      l.lastUpdatedBy || '',
      l.competitorWhoWon || '',
      l.status || 'Not started',
      l.rootCause || '',
      l.improvementsRequired || '',
      l.actionOwner || '',
      l.targetStart || '',
      l.targetEnd || l.targetDate || '',
      l.plannedDays !== undefined ? l.plannedDays : '',
      l.actualEndDate || l.completedDate || '',
      latestNote
    ]);
  }

  ensureSheetDimensions(sheet, Math.max(20, rows.length + 5), Math.max(20, headers.length));
  sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
  formatHeaderRow(sheet, headers.length);
  autoFormatTable(sheet, rows.length, headers.length);
}

/* ================= 5. SALES REP ACTIVITY TAB ================= */

function updateActivitySheet(ss, activityObj) {
  var sheet = getOrCreateSheet(ss, 'Sales Rep Activity');
  sheet.clear();

  var headers = [
    "Username / Email",
    "Field Presence Status",
    "Last Seen Timestamp",
    "Current Location Label",
    "Latitude",
    "Longitude"
  ];

  var rows = [headers];
  var keys = Object.keys(activityObj);
  var now = new Date().getTime();

  for (var i = 0; i < keys.length; i++) {
    var act = activityObj[keys[i]];
    var lastSeenMs = act.lastSeen ? Number(act.lastSeen) : 0;
    var lastSeenStr = lastSeenMs ? Utilities.formatDate(new Date(lastSeenMs), Session.getScriptTimeZone() || "GMT+5:30", "yyyy-MM-dd HH:mm:ss") : 'No Activity';

    var isActive = lastSeenMs > 0 && (now - lastSeenMs) < (60 * 60 * 1000);
    var minutesAgo = lastSeenMs > 0 ? Math.round((now - lastSeenMs) / 60000) : null;
    var statusText = isActive ? 'Active in field' : (minutesAgo !== null ? ('Inactive (' + minutesAgo + 'm ago)') : 'Inactive');

    rows.push([
      act.username || keys[i],
      statusText,
      lastSeenStr,
      act.locationLabel || 'Not shared',
      act.lat !== undefined ? act.lat : '',
      act.lng !== undefined ? act.lng : ''
    ]);
  }

  ensureSheetDimensions(sheet, Math.max(20, rows.length + 5), Math.max(10, headers.length));
  sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
  formatHeaderRow(sheet, headers.length);
  autoFormatTable(sheet, rows.length, headers.length);
}

/* ================= UTILITY & FORMATTING FUNCTIONS ================= */

function formatHeaderRow(sheet, colCount) {
  var range = sheet.getRange(1, 1, 1, colCount);
  range.setBackground('#79282d');
  range.setFontColor('#ffffff');
  range.setFontWeight('bold');
  range.setFontFamily('Arial');
  range.setFontSize(10);
  range.setWrap(false);
  sheet.setFrozenRows(1);
}

function autoFormatTable(sheet, rowCount, colCount) {
  if (rowCount <= 1) return;
  var dataRange = sheet.getRange(2, 1, rowCount - 1, colCount);
  dataRange.setFontFamily('Arial');
  dataRange.setFontSize(9);
  dataRange.setVerticalAlignment('middle');
  sheet.setRowHeights(2, rowCount - 1, 24);
}

function computeDaysBetween(startDateStr, endDateStr) {
  try {
    var s = new Date(startDateStr);
    var e = new Date(endDateStr);
    var diffTime = e.getTime() - s.getTime();
    return Math.round(diffTime / (1000 * 3600 * 24));
  } catch (err) {
    return 0;
  }
}

function getOrCreateSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function doOptions(e) {
  return ContentService.createTextOutput('')
    .setMimeType(ContentService.MimeType.TEXT);
}

function responseJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ================= AUTHORIZED TEAM ROSTER ================= */

function getAuthorizedUsersList() {
  return [
    { username: 'siddharth@patilgroup.com', role: 'ceo', name: 'Siddharth Sharma', roleTitle: 'CEO / Manager' },
    { username: 'vitrang@patilgroup.com', role: 'ceo', name: 'Vitrang G Sompura', roleTitle: 'Head Of Precast' },
    { username: 'francisyakobu@patilgroup.com', role: 'ceo', name: 'Francis Yakobu', roleTitle: 'Sales Manager' },
    { username: 'srikar@patilgroup.com', role: 'ceo', name: 'Srikar', roleTitle: 'Digital Engineer' },
    { username: 'krishnamraju@patilgroup.com', role: 'rep', name: 'GRK Raju', roleTitle: 'Regional Manager' },
    { username: 'saurabhsamya@patilgroup.com', role: 'rep', name: 'Saurabh Samya', roleTitle: 'Regional Manager' },
    { username: 'ramchandra.s@patilgroup.com', role: 'rep', name: 'Ramchandra', roleTitle: 'Engineer Trainee' },
    { username: 'rahul.s@patilgroup.com', role: 'rep', name: 'Rahul', roleTitle: 'Engineer Trainee' }
  ];
}
