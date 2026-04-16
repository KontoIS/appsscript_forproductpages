// ============================================================
//  Konto — Product Sales & Attendees  |  Google Apps Script
//
//  Pulls your Konto product sales list and attendee data
//  directly into Google Sheets via the Konto REST API.
//
//  API docs: https://konto.is/api/v1/document
//  Auth:     POST with username + api_key as form fields
// ============================================================

// ------------------------------------------------------------
//  ✏️  CONFIGURATION — fill in your credentials here
// ------------------------------------------------------------
var CONFIG = {
  BASE_URL:  "https://konto.is/api/v1",  // production endpoint
  USERNAME:  "YOUR_USERNAME_HERE",        // your Konto username
  API_KEY:   "YOUR_API_KEY_HERE",         // your Konto API key
  PAGE_SIZE: 100                          // records per page (max recommended: 100)
};
// ------------------------------------------------------------

var SHEET_SUMMARY   = "Summary";
var SHEET_ATTENDEES = "Attendees";

// ============================================================
//  MENU  — adds "🔄 Konto" to the Google Sheets menu bar
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🔄 Konto")
    .addItem("Refresh All Data",         "refreshAll")
    .addSeparator()
    .addItem("Refresh Sales Count only", "refreshSalesCount")
    .addItem("Pick a Product Sale…",     "showProductPicker")
    .addToUi();
}

// ============================================================
//  REFRESH ALL
//  Refreshes the sales count and re-loads attendees for the
//  last selected product (or prompts to pick one if none set).
// ============================================================
function refreshAll() {
  refreshSalesCount();
  var sheet = getOrCreateSheet(SHEET_ATTENDEES);
  var guid  = sheet.getRange("B2").getValue().toString().trim();
  if (guid) {
    fetchAndWriteAttendees(guid);
  } else {
    showProductPicker();
  }
}

// ============================================================
//  PRODUCT PICKER
//  Step 1 — shows a numbered list of all your product sales
//  Step 2 — prompts the user to enter the number of their choice
//  Step 3 — fetches and writes attendees for the chosen product
//
//  Uses only native Sheets dialogs (no HTML service / no
//  google.script.run callbacks) to avoid PERMISSION_DENIED.
// ============================================================
function showProductPicker() {
  var ui       = SpreadsheetApp.getUi();
  var products = fetchAllProductSales();

  if (products.length === 0) {
    ui.alert("No product sales found. Check your credentials in the CONFIG block.");
    return;
  }

  // Build numbered list e.g. "1.  Test event on FEB 17th"
  var lines = products.map(function(p, i) {
    var label = p.name || p.title || p.description || ("Product #" + (i + 1));
    return (i + 1) + ".  " + label;
  });

  // Step 1: show the list
  ui.alert(
    "Your Product Sales (" + products.length + " total)",
    lines.join("\n") + "\n\nNote the number of the product you want, then click OK.",
    ui.ButtonSet.OK
  );

  // Step 2: ask for the number
  var resp = ui.prompt(
    "Pick a Product Sale",
    "Enter the number (1–" + products.length + "):",
    ui.ButtonSet.OK_CANCEL
  );

  if (resp.getSelectedButton() !== ui.Button.OK) return;

  var num = parseInt(resp.getResponseText().trim(), 10);
  if (isNaN(num) || num < 1 || num > products.length) {
    ui.alert("Invalid number — please try again.");
    return;
  }

  var chosen = products[num - 1];
  var guid   = chosen.guid || "";
  if (!guid) {
    ui.alert("This product has no GUID — cannot load attendees.");
    return;
  }

  // Step 3: fetch attendees
  fetchAndWriteAttendees(guid);
}

// ============================================================
//  FETCH ATTENDEES
//  Fetches all pages of attendees for the given product GUID
//  and writes them to the Attendees sheet.
// ============================================================
function fetchAndWriteAttendees(guid) {
  var sheet = getOrCreateSheet(SHEET_ATTENDEES);
  setupAttendeesHeader(sheet);
  sheet.getRange("B2").setValue(guid); // persist so Refresh All can re-use it

  var allAttendees = [];
  var page = 1;

  while (true) {
    try {
      var data = apiPost("/get-product-sale-attends", {
        guid:  guid,
        page:  page,
        limit: CONFIG.PAGE_SIZE
      });
      var rows = extractArray(data);
      if (rows.length === 0) break;
      allAttendees = allAttendees.concat(rows);
      if (rows.length < CONFIG.PAGE_SIZE) break; // reached last page
      page++;
    } catch (e) {
      logError(sheet, "get-product-sale-attends (page " + page + ")", e);
      break;
    }
  }

  writeAttendeesToSheet(sheet, allAttendees, guid);
}

// ============================================================
//  FETCH ALL PRODUCT SALES  (paginates automatically)
// ============================================================
function fetchAllProductSales() {
  var all  = [];
  var page = 1;

  while (true) {
    try {
      var data = apiPost("/get-product-sales", { page: page, limit: CONFIG.PAGE_SIZE });
      var rows = extractArray(data);
      if (rows.length === 0) break;
      all = all.concat(rows);
      if (rows.length < CONFIG.PAGE_SIZE) break;
      page++;
    } catch (e) {
      Logger.log("fetchAllProductSales error: " + e.message);
      break;
    }
  }

  return all;
}

// ============================================================
//  SALES COUNT
//  Writes total number of product sales to the Summary sheet.
// ============================================================
function refreshSalesCount() {
  var sheet = getOrCreateSheet(SHEET_SUMMARY);

  try {
    var data = apiPost("/get-count-product-sales", {});

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["Field", "Value", "Last Updated"]);
      sheet.getRange(1, 1, 1, 3)
           .setFontWeight("bold").setBackground("#1a73e8").setFontColor("#ffffff");
    }

    // API returns count in data.result
    var countValue = data.result !== undefined ? data.result
                   : data.count  !== undefined ? data.count
                   : data.total  !== undefined ? data.total
                   : JSON.stringify(data);

    upsertRow(sheet, "Total Product Sales", [countValue, new Date()]);

  } catch (e) {
    logError(sheet, "get-count-product-sales", e);
  }
}

// ============================================================
//  WRITE ATTENDEES TO SHEET  (merge mode)
//
//  Safe to run repeatedly without losing manual edits:
//
//  • API columns (standard + custom fields) are refreshed
//    in-place for rows that already exist, matched by
//    invoice_number.
//  • Any columns the user has added to the RIGHT of the API
//    columns are never touched — notes, checkboxes, formulas
//    all survive a sync.
//  • New registrations from the API are appended at the bottom.
//  • The raw `fields` array is expanded — each unique label
//    becomes its own column (highlighted in lighter blue).
//
//  Merge key: invoice_number  (unique per registration)
// ============================================================
function writeAttendeesToSheet(sheet, attendees, guid) {
  sheet.getRange("D2").setValue(new Date()); // timestamp

  if (attendees.length === 0) {
    SpreadsheetApp.getActive().toast("No attendees found.", "Konto", 4);
    return;
  }

  // ── 1. Build column definitions ───────────────────────────

  // Standard API fields (strip the raw `fields` array — we expand it below)
  var standardHeaders = Object.keys(attendees[0]).filter(function(k) {
    return k !== "fields";
  });

  // Collect all unique custom field labels, in order of first appearance
  var customLabelsSeen = {};
  var customLabels = [];
  attendees.forEach(function(a) {
    if (Array.isArray(a.fields)) {
      a.fields.forEach(function(f) {
        if (f.label && !customLabelsSeen[f.label]) {
          customLabelsSeen[f.label] = true;
          customLabels.push(f.label);
        }
      });
    }
  });

  // Full set of API-owned headers
  var apiHeaders = standardHeaders.concat(customLabels);
  var apiColCount = apiHeaders.length;

  // Index of the merge key within apiHeaders
  var keyCol = apiHeaders.indexOf("invoice_number");
  if (keyCol === -1) { keyCol = 0; } // fallback to first column

  // ── 2. Read what's already in the sheet ──────────────────

  var lastRow    = sheet.getLastRow();
  var totalCols  = Math.max(sheet.getLastColumn(), apiColCount);
  var existingHeaders = [];
  var rowIndex   = {}; // invoice_number → sheet row number (1-based)

  if (lastRow >= 4) {
    // Read existing header row to know the current column order
    existingHeaders = sheet.getRange(4, 1, 1, totalCols).getValues()[0];
  }

  if (lastRow >= 5) {
    // Map every existing invoice_number to its row number
    var keySheetCol = existingHeaders.indexOf("invoice_number");
    if (keySheetCol === -1) { keySheetCol = 0; }
    var keyData = sheet.getRange(5, keySheetCol + 1, lastRow - 4, 1).getValues();
    keyData.forEach(function(r, i) {
      if (r[0] !== "") { rowIndex[String(r[0])] = 5 + i; }
    });
  }

  // ── 3. Write / update header row ─────────────────────────

  // Merge: keep any user-added columns that come after the API columns
  var userHeaders = [];
  if (existingHeaders.length > apiColCount) {
    userHeaders = existingHeaders.slice(apiColCount).filter(function(h) {
      return h !== "";
    });
  }
  var allHeaders = apiHeaders.concat(userHeaders);

  var headerRange = sheet.getRange(4, 1, 1, allHeaders.length);
  headerRange.setValues([allHeaders]);
  headerRange.setFontWeight("bold").setBackground("#1a73e8").setFontColor("#ffffff");

  // Lighter blue for custom field columns
  if (customLabels.length > 0) {
    sheet.getRange(4, standardHeaders.length + 1, 1, customLabels.length)
         .setBackground("#4a90d9");
  }

  // ── 4. Build a helper: attendee object → API value array ──

  function buildApiRow(a) {
    var row = standardHeaders.map(function(h) {
      var v = a[h];
      return v === null || v === undefined ? "" : v;
    });
    if (customLabels.length > 0) {
      var fieldMap = {};
      if (Array.isArray(a.fields)) {
        a.fields.forEach(function(f) {
          if (f.label) { fieldMap[f.label] = f.value !== undefined ? f.value : ""; }
        });
      }
      customLabels.forEach(function(label) {
        row.push(fieldMap[label] !== undefined ? fieldMap[label] : "");
      });
    }
    return row;
  }

  // ── 5. Merge: update existing rows, collect new ones ──────

  var newAttendees = [];

  attendees.forEach(function(a) {
    var key    = String(a.invoice_number || a[standardHeaders[0]] || "");
    var apiRow = buildApiRow(a);

    if (rowIndex[key] !== undefined) {
      // Row already exists — update only the API columns, leave user columns alone
      sheet.getRange(rowIndex[key], 1, 1, apiColCount).setValues([apiRow]);
    } else {
      // New registration — queue for append
      newAttendees.push(apiRow);
    }
  });

  // ── 6. Append new rows ────────────────────────────────────

  if (newAttendees.length > 0) {
    var appendStart = Math.max(lastRow + 1, 5);
    sheet.getRange(appendStart, 1, newAttendees.length, apiColCount)
         .setValues(newAttendees);
  }

  // ── 7. Auto-resize API columns only ──────────────────────

  for (var i = 1; i <= apiColCount; i++) { sheet.autoResizeColumn(i); }

  // ── 8. Toast ─────────────────────────────────────────────

  var updatedCount = attendees.length - newAttendees.length;
  var customNote   = customLabels.length > 0
    ? " · " + customLabels.length + " custom field" + (customLabels.length > 1 ? "s" : "")
    : "";
  SpreadsheetApp.getActive().toast(
    "✅ " + updatedCount + " updated · " + newAttendees.length + " new" + customNote,
    "Konto", 6
  );
}

// ============================================================
//  API HELPER
//  All Konto API calls use POST with username + api_key
//  passed as multipart form fields (same as curl -F).
// ============================================================
function apiPost(path, params) {
  // Always include credentials
  var payload = {
    username: CONFIG.USERNAME,
    api_key:  CONFIG.API_KEY
  };
  // Merge in any additional params
  for (var k in params) { payload[k] = String(params[k]); }

  var options = {
    method:             "post",
    payload:            payload,   // UrlFetchApp sends this as multipart/form-data
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(CONFIG.BASE_URL + path, options);
  var code     = response.getResponseCode();
  var body     = response.getContentText();

  if (code < 200 || code >= 300) {
    throw new Error("HTTP " + code + " — " + body.substring(0, 200));
  }

  var json = JSON.parse(body);
  if (json.status === false) {
    throw new Error(json.message || "API returned status: false");
  }

  return json;
}

// ============================================================
//  UTILITY HELPERS
// ============================================================

// Extracts the result array from various API response shapes
function extractArray(data) {
  if (Array.isArray(data))           return data;
  if (Array.isArray(data.result))    return data.result;
  if (Array.isArray(data.attendees)) return data.attendees;
  if (Array.isArray(data.data))      return data.data;
  if (Array.isArray(data.items))     return data.items;
  return [];
}

// Sets up the fixed header rows on the Attendees sheet
function setupAttendeesHeader(sheet) {
  sheet.getRange("A1").setValue("Konto — Product Sale Attendees")
       .setFontWeight("bold").setFontSize(13);
  sheet.getRange("A2").setValue("Active product sale GUID:").setFontWeight("bold");
  sheet.getRange("B2").setBackground("#fff2cc");
  sheet.getRange("C2").setValue("Last updated:").setFontStyle("italic").setFontColor("#888888");
}

// Returns the named sheet, creating it if it doesn't exist
function getOrCreateSheet(name) {
  var ss    = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(name);
  if (!sheet) { sheet = ss.insertSheet(name); }
  return sheet;
}

// Updates an existing labelled row or appends a new one
function upsertRow(sheet, label, values) {
  var data     = sheet.getDataRange().getValues();
  var rowIndex = -1;
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === label) { rowIndex = i + 1; break; }
  }
  if (rowIndex === -1) {
    sheet.appendRow([label].concat(values));
  } else {
    sheet.getRange(rowIndex, 1, 1, 1 + values.length)
         .setValues([[label].concat(values)]);
  }
}

// Logs an error to the sheet and shows a toast notification
function logError(sheet, context, err) {
  Logger.log("[Konto Error] " + context + ": " + err.message);
  sheet.getRange("A1").setValue("⚠️ Error (" + context + "): " + err.message)
       .setFontColor("red");
  SpreadsheetApp.getActive().toast("⚠️ Error: " + err.message, "Konto", 8);
}
