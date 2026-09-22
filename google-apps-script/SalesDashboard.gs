/**
 * ============================================================================
 *  SALES PERFORMANCE DASHBOARD — Google Apps Script
 * ============================================================================
 *  This script gives you TWO ways to see the dashboard — pick either or use
 *  both:
 *
 *  MODE A — a tab inside the spreadsheet (no deployment needed)
 *  --------------------------------------------------------------
 *   1. Open your Google Sheet (the one with the "crm" and "target" tabs).
 *   2. Extensions → Apps Script, paste this whole file in, save.
 *   3. Run the function "buildDashboard" once from the editor (▶) and
 *      accept the authorization prompt.
 *   4. Reload the spreadsheet — a "📊 Sales Dashboard" menu appears. Use
 *      "Build / Refresh Dashboard" any time crm/target data changes.
 *
 *  MODE B — a real, standalone web page (a "script page" / Web App)
 *  --------------------------------------------------------------
 *   1. Do steps 1–2 from Mode A (paste the script in, save).
 *   2. Click "Deploy" (top-right) → "New deployment".
 *   3. Click the gear icon next to "Select type" → choose "Web app".
 *   4. Fill in:
 *        Description:      Sales Dashboard
 *        Execute as:       Me (your account)
 *        Who has access:   Anyone with the link (or "Anyone within your
 *                           organization" if you're on Google Workspace)
 *   5. Click "Deploy" → authorize when prompted.
 *   6. Copy the "Web app URL" you're given — open it in any browser to see
 *      a live, colorful, auto-refreshing dashboard page (no login into the
 *      spreadsheet required for viewers who have the link).
 *   7. Whenever you edit crm/target data, just reload that URL — it always
 *      recomputes live from the current sheet contents.
 *   NOTE: if you edit the script later, repeat Deploy → "Manage deployments"
 *   → pencil icon → "New version" → Deploy, so the URL picks up the change.
 *
 *  WHAT IT COMPUTES (same for both modes)
 *  ---------------------------------------
 *   1. Store target vs achievement vs balance, per revenue type
 *   2. Working days: total / consumed / balance for the current month
 *   3. Salesperson performance per revenue type (every name found in "crm")
 *   4. Incentive valuation — AMC & Bundle slabs, per salesperson
 *   5. Store KPI parameters — target / achieved / balance / achieved
 *      value / achieved GP, from the "target" KPI grid
 *   6. Invoice valuation — total & unique invoices, per revenue type
 *   7. Date-wise sales revenue analysis with per-revenue-type daily target
 *      (= balance for that type ÷ balance working days)
 *
 *  DATA RULES APPLIED (fixed, per your spec)
 *  ------------------------------------------
 *   • Only rows where TranType (case-insensitive) = "SALES" are counted.
 *   • TaxableAmount = achieved value. GrossProfit = gross profit metric.
 *   • ItemDescription "Display"/"Battery"/"Glass" categories require BOTH
 *     the part keyword AND "iphone" in the same description.
 *   • SubCategory containing "AMC" = an AMC row (unique invoices counted).
 *   • SchemeCode starting with "DISC" (case-insensitive) = a bundle row
 *     (unique invoices counted).
 *
 *  CONFIG — edit these to match your sheet names / calendar if needed.
 * ============================================================================
 */

var CONFIG = {
  CRM_SHEET_NAME: "crm",
  TARGET_SHEET_NAME: "target",
  DASHBOARD_SHEET_NAME: "Dashboard",
  // Leave blank to always use the spreadsheet this script is bound to.
  // Only set this if you deploy the Web App from a *different* Apps Script
  // project than the one bound to your sheet (Extensions → Apps Script
  // always binds it for you, so you normally don't need to touch this).
  SPREADSHEET_ID: "",
  // Day-of-month numbers treated as holidays / non-working days this month.
  // Edit this list to match your store's actual holiday calendar.
  HOLIDAYS: [1, 15, 28],
  // A salesperson only *receives* their earned service + accessory incentive
  // once their repair revenue reaches this % of their repair target.
  // Repair revenue itself never earns incentive — it is purely a gate.
  REPAIR_INCENTIVE_GATE_PCT: 80,
  // Seconds the rendered page is cached so normal loads are fast. The
  // "Refresh" button always bypasses this and recomputes from the sheet.
  CACHE_SECONDS: 90,
};

/* ============================================================================
 * MENU
 * ========================================================================= */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📊 Sales Dashboard")
    .addItem("Build / Refresh Dashboard (sheet tab)", "buildDashboard")
    .addItem("How do I get the Web App page?", "showWebAppHelp_")
    .addToUi();
}

function showWebAppHelp_() {
  var html = HtmlService.createHtmlOutput(
    "<div style='font:14px Arial,sans-serif;line-height:1.5;padding:4px 8px'>" +
      "<p><b>To open the dashboard as a standalone web page:</b></p>" +
      "<ol>" +
      "<li>In the Apps Script editor, click <b>Deploy</b> (top-right) → <b>New deployment</b>.</li>" +
      "<li>Click the gear icon next to \"Select type\" → choose <b>Web app</b>.</li>" +
      "<li>Execute as: <b>Me</b>. Who has access: <b>Anyone with the link</b>.</li>" +
      "<li>Click <b>Deploy</b>, authorize, then copy the Web app URL.</li>" +
      "<li>Open that URL in any browser — it always shows live data from this sheet.</li>" +
      "</ol></div>",
  ).setWidth(420).setHeight(280);
  SpreadsheetApp.getUi().showModalDialog(html, "Web App Dashboard — Setup");
}

/* ============================================================================
 * ENTRY POINTS
 * ========================================================================= */

/** Resolves the spreadsheet regardless of whether we're in a UI context (Mode A) or an HTTP request (Mode B). */
function getSpreadsheet_() {
  if (CONFIG.SPREADSHEET_ID) return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  return SpreadsheetApp.getActiveSpreadsheet();
}

/** Shared pipeline: read crm + target sheets, parse, and compute the dashboard data object. */
function loadDashboardData_() {
  var ss = getSpreadsheet_();
  var crmSheet = ss.getSheetByName(CONFIG.CRM_SHEET_NAME);
  var targetSheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);

  if (!crmSheet) throw new Error('Sheet "' + CONFIG.CRM_SHEET_NAME + '" not found.');
  if (!targetSheet) throw new Error('Sheet "' + CONFIG.TARGET_SHEET_NAME + '" not found.');

  var crmValues = crmSheet.getDataRange().getValues();
  var targetValues = targetSheet.getDataRange().getValues();

  var allRows = parseCrmRows(crmValues);
  var target = parseTargetSheet(targetValues);
  return { ss: ss, data: computeDashboard(allRows, target) };
}

/** MODE A: writes/refreshes the "Dashboard" tab inside the spreadsheet. */
function buildDashboard() {
  var loaded = loadDashboardData_();
  writeDashboard(loaded.ss, loaded.data);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Dashboard rebuilt: " + loaded.data.store.totalInvoices + " SALES rows, " +
      loaded.data.store.uniqueInvoices + " unique invoices.",
    "Sales Dashboard",
    6,
  );
}

/**
 * MODE B: the Web App entry point. Apps Script calls this automatically
 * whenever someone opens the deployed Web App URL (see the setup steps in
 * the header comment above — Deploy → New deployment → Web app).
 *
 * Fast path: the fully-rendered HTML is cached for CONFIG.CACHE_SECONDS so
 * repeat visits are near-instant. Opening the URL with "?refresh=1" (what
 * the on-page Refresh button does) skips the cache and recomputes.
 */
function doGet(e) {
  var wantsFresh = !!(e && e.parameter && e.parameter.refresh);
  var cache = CacheService.getScriptCache();
  var CACHE_KEY = "dashboard_html_v3";

  var html = null;
  if (!wantsFresh) {
    try { html = cache.get(CACHE_KEY); } catch (err) { html = null; }
  }

  if (!html) {
    var loaded = loadDashboardData_();
    html = renderDashboardHtml_(loaded.data);
    // Cache entries are capped at 100 KB; skip caching if the page is bigger.
    if (html.length < 100000) {
      try { cache.put(CACHE_KEY, html, CONFIG.CACHE_SECONDS); } catch (err2) { /* non-fatal */ }
    }
  }

  return HtmlService.createHtmlOutput(html)
    .setTitle("Sales Performance Dashboard")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ============================================================================
 * SMALL HELPERS
 * ========================================================================= */

function num_(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return isFinite(v) ? v : 0;
  var s = String(v).replace(/[₹,\s]/g, "");
  var n = parseFloat(s);
  return isFinite(n) ? n : 0;
}

function str_(v) {
  return v === null || v === undefined ? "" : String(v).trim();
}

function pct_(part, whole) {
  if (!whole || whole <= 0) return 0;
  return Math.round((part / whole) * 1000) / 10;
}

function safeDiv_(a, b) {
  return b > 0 ? a / b : 0;
}

function titleCase_(key) {
  return String(key)
    .split(/\s+/)
    .filter(function (w) { return w.length > 0; })
    .map(function (w) { return w.charAt(0) + w.slice(1).toLowerCase(); })
    .join(" ");
}

/** Uppercase key used to merge employee spellings across sheets. */
function normalizePerson_(name) {
  return str_(name).toUpperCase().replace(/\s+/g, " ");
}

/** "Service Revenue" -> "service" | "Accessories" -> "accessory" | "Repair" -> "repair" | null */
function revenueTypeKey_(rt) {
  var t = str_(rt).toLowerCase();
  if (t.indexOf("service") === 0) return "service";
  if (t.indexOf("accessor") === 0) return "accessory";
  if (t.indexOf("repair") === 0) return "repair";
  return null;
}

/** Row qualifies for every dashboard calculation: crm!TranType must be "SALES". */
function isSaleRow_(r) {
  return str_(r.tranType).toUpperCase() === "SALES";
}

/** Parse a date cell that may be a real Date object or a "DD-MM-YYYY HH:mm" string. */
function parseInvoiceDate_(raw) {
  if (raw instanceof Date && !isNaN(raw.getTime())) {
    return { day: raw.getDate(), dayKey: Utilities.formatDate(raw, Session.getScriptTimeZone(), "dd-MM-yyyy") };
  }
  var s = str_(raw);
  if (!s) return { day: 0, dayKey: "" };
  var datePart = s.split(" ")[0];
  var m = datePart.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{2,4})$/);
  if (m) {
    var dd = parseInt(m[1], 10);
    return { day: dd || 0, dayKey: datePart };
  }
  // Fallback: try native Date parsing (handles ISO-ish strings)
  var d = new Date(s);
  if (!isNaN(d.getTime())) {
    return { day: d.getDate(), dayKey: Utilities.formatDate(d, Session.getScriptTimeZone(), "dd-MM-yyyy") };
  }
  return { day: 0, dayKey: datePart };
}

/** Find the column index (0-based) whose header matches any of the given regexes. Returns -1 if none. */
function findColumn_(header, patterns) {
  for (var c = 0; c < header.length; c++) {
    var h = str_(header[c]);
    for (var p = 0; p < patterns.length; p++) {
      if (patterns[p].test(h)) return c;
    }
  }
  return -1;
}

/* ============================================================================
 * CRM ROW PARSING  (columns looked up by header name — order-independent)
 * Required headers (case-insensitive, partial match ok):
 *   InvoiceDate, SalesPersonName, RevenueType, ItemDescription, SubCategory,
 *   SchemeCode, TaxableAmount, GrossProfit, TranType
 * Optional (used for extra KPI breakdowns if present):
 *   InvoiceCode/InvoiceNo/InvoiceNumber, ItemCode, PartType
 * ========================================================================= */

function parseCrmRows(values) {
  if (!values || values.length < 2) return [];
  var header = values[0];

  var col = {
    invoiceCode: findColumn_(header, [/invoice.*(code|no\.?|number|id)/i]),
    invoiceDate: findColumn_(header, [/invoice.*date/i]),
    salesPerson: findColumn_(header, [/sales.*person/i]),
    revenueType: findColumn_(header, [/revenue.*type/i]),
    itemDescription: findColumn_(header, [/item.*desc/i]),
    subCategory: findColumn_(header, [/sub.*categ/i]),
    schemeCode: findColumn_(header, [/scheme.*code/i]),
    taxableAmount: findColumn_(header, [/taxable.*amount/i]),
    grossProfit: findColumn_(header, [/gross.*profit/i]),
    tranType: findColumn_(header, [/tran.*type/i]),
    itemCode: findColumn_(header, [/item.*code/i]),
    partType: findColumn_(header, [/part.*type/i]),
  };

  var missing = [];
  ["invoiceDate", "salesPerson", "revenueType", "itemDescription", "subCategory", "schemeCode", "taxableAmount", "grossProfit", "tranType"].forEach(
    function (k) { if (col[k] === -1) missing.push(k); },
  );
  if (missing.length) {
    throw new Error(
      'Could not find these required columns in "' + CONFIG.CRM_SHEET_NAME + '": ' + missing.join(", ") +
        ". Check the header row spelling.",
    );
  }

  var out = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var invoiceDateRaw = row[col.invoiceDate];
    if (!invoiceDateRaw && col.invoiceCode === -1) continue; // nothing to key this row on
    var parsedDate = parseInvoiceDate_(invoiceDateRaw);
    var invoiceCode =
      col.invoiceCode !== -1 && str_(row[col.invoiceCode])
        ? str_(row[col.invoiceCode])
        : parsedDate.dayKey + "#" + str_(row[col.salesPerson]) + "#" + r; // best-effort fallback key

    out.push({
      invoiceCode: invoiceCode,
      day: parsedDate.day,
      dayKey: parsedDate.dayKey,
      salesPerson: str_(row[col.salesPerson]),
      revenueType: str_(row[col.revenueType]),
      itemDescription: str_(row[col.itemDescription]),
      subCategory: str_(row[col.subCategory]),
      schemeCode: str_(row[col.schemeCode]),
      taxableAmount: num_(row[col.taxableAmount]),
      grossProfit: num_(row[col.grossProfit]),
      tranType: str_(row[col.tranType]),
      itemCode: col.itemCode !== -1 ? str_(row[col.itemCode]) : "",
      partType: col.partType !== -1 ? str_(row[col.partType]) : "",
    });
  }
  return out;
}

/* ============================================================================
 * TARGET SHEET PARSING
 * Grid anchors (matches the layout used by the companion web dashboard for
 * this same spreadsheet). If your target sheet layout differs, adjust the
 * row/column numbers below — search for "ANCHOR" comments.
 * ========================================================================= */

var OPEN_ENDED = 9999;

function parseMoney_(s) {
  return num_(s);
}

function parseRangeLabel_(label) {
  var t = str_(label);
  var m = t.match(/^(\d+)\s*[-–]\s*(\d+)$/);
  if (m) return { min: Number(m[1]), max: Number(m[2]) };
  var single = t.match(/^(\d+)\+$/);
  if (single) return { min: Number(single[1]), max: Infinity };
  var only = t.match(/^(\d+)$/);
  if (only) return { min: Number(only[1]), max: Number(only[1]) };
  return { min: 1, max: Infinity };
}

/** "200/AMC + 750Fixed" -> {rate:200, fixed:750} */
function parseAmcRate_(label) {
  var t = str_(label);
  var rateM = t.match(/(\d+(?:\.\d+)?)\s*\/\s*AMC/i);
  var fixedM = t.match(/\+\s*(\d+(?:\.\d+)?)\s*(?:Fixed)?/i);
  return { rate: rateM ? Number(rateM[1]) : 0, fixed: fixedM ? Number(fixedM[1]) : 0 };
}

function parsePercentLabel_(label) {
  var t = str_(label).replace(/%/g, "");
  if (t.indexOf("<=") === 0) return { min: 0, max: Number(t.slice(2)) };
  var m = t.match(/^(\d+)\s*(?:to|-|–)\s*(\d+)$/);
  if (m) return { min: Number(m[1]), max: Number(m[2]) };
  if (/^\d+$/.test(t)) return { min: Number(t), max: Number(t) };
  return { min: 0, max: 0 };
}

function isBandRow_(label, value) {
  var l = str_(label), v = str_(value);
  return /%|nil/i.test(l + v) || /^[<>=]?\s*\d+\s*%?$/.test(l.trim()) || /\d\s*(?:to|-|–)\s*\d/i.test(l);
}

function parsePctSlabs_(cells) {
  var slabs = [];
  var prevMax = 0;
  for (var i = 0; i < cells.length; i++) {
    var label = cells[i][0], rawVal = cells[i][1];
    if (!label) continue;
    var value = parseMoney_(rawVal);
    var range = parsePercentLabel_(label);
    if (/^<\d/.test(str_(label).trim()) && prevMax > 0) range = { min: prevMax + 1, max: Infinity };
    slabs.push({ label: str_(label), value: value, minPct: range.min, maxPct: Math.min(range.max, OPEN_ENDED) });
    prevMax = Math.max(prevMax, range.max === Infinity ? prevMax : range.max);
  }
  return slabs;
}

function slabForPct_(slabs, pctVal) {
  for (var i = 0; i < slabs.length; i++) {
    if (pctVal >= slabs[i].minPct && pctVal <= slabs[i].maxPct) return slabs[i];
  }
  return null;
}

function incentiveForPct_(slabs, pctVal) {
  var s = slabForPct_(slabs, pctVal);
  return s ? s.value : 0;
}

function findRow_(rows, fromCol, toCol, re) {
  for (var r = 0; r < rows.length; r++) {
    for (var c = fromCol; c <= toCol; c++) {
      if (re.test(str_(rows[r][c]).trim())) return r;
    }
  }
  return -1;
}

function findCell_(rows, fromCol, toCol, re) {
  for (var r = 0; r < rows.length; r++) {
    for (var c = fromCol; c <= toCol; c++) {
      if (re.test(str_(rows[r][c]).trim())) return { row: r, col: c };
    }
  }
  return null;
}

/**
 * Google review block — ANCHOR: target!A9:C10.
 *   Row  9 (A9:C9)  = headers
 *   Row 10 (A10)    = previous month qty
 *   Row 10 (B10)    = current month count
 *   Row 10 (C10)    = difference
 * 0-indexed that is row 9, columns 0 / 1 / 2.
 */
function parseGoogleReview_(rows) {
  function at(r, c) { return str_(rows[r] && rows[r][c]); }
  var previousMonth = parseMoney_(at(9, 0));
  var currentMonth = parseMoney_(at(9, 1));
  var diffRaw = at(9, 2);
  var difference = diffRaw !== "" ? parseMoney_(diffRaw) : currentMonth - previousMonth;
  return { previousMonth: previousMonth, currentMonth: currentMonth, difference: difference };
}

function parseTargetSheet(rows) {
  function at(r, c) { return str_(rows[r] && rows[r][c]); }

  // ---- ANCHOR: store + salesperson targets — target!A1:D6 ----
  //   A1:A6 = revenue type label   (row 1 = header, rows 2..6 = the types)
  //   B     = store target
  //   C, D  = per-salesperson target; C1 / D1 hold the salesperson names.
  // 0-indexed: header row 0, data rows 1..5, cols A=0 B=1 C=2 D=3.
  function slotFor_(label) {
    var l = String(label || "").toLowerCase();
    if (l.indexOf("service") > -1) return "service";
    if (l.indexOf("accessor") > -1) return "accessory";
    if (l.indexOf("repair") > -1) return "repair";
    if (l.indexOf("gross") > -1 || l.indexOf("gp") === 0) return "grossProfit";
    if (l.indexOf("total") > -1) return "total";
    return null;
  }

  var store = { service: 0, accessory: 0, repair: 0, total: 0, grossProfit: 0 };

  // Salesperson names live in the header row, columns C and D onward.
  var people = {};
  var personKeys = [];
  var colIndex = {};
  var headerCells = rows[0] || [];
  for (var c = 2; c < Math.min(headerCells.length, 8); c++) {
    var name = str_(headerCells[c]);
    if (!name) continue;
    // skip generic column captions that aren't people
    if (/^(target|store|amount|value|revenue|type)$/i.test(name)) continue;
    var key = normalizePerson_(name);
    if (!key || personKeys.indexOf(key) > -1) continue;
    personKeys.push(key);
    colIndex[key] = c;
    people[key] = { service: 0, accessory: 0, repair: 0, total: 0, grossProfit: 0 };
  }

  // Data rows 2..6 on the sheet (0-indexed 1..5).
  for (var r = 1; r <= 5 && r < rows.length; r++) {
    var slot = slotFor_(at(r, 0));
    if (!slot) continue;
    store[slot] = parseMoney_(at(r, 1));
    personKeys.forEach(function (key) {
      people[key][slot] = parseMoney_(at(r, colIndex[key]));
    });
  }

  // If the sheet didn't spell out Total / Gross Profit rows, derive Total.
  if (!store.total) store.total = store.service + store.accessory + store.repair;
  personKeys.forEach(function (key) {
    var p = people[key];
    if (!p.total) p.total = p.service + p.accessory + p.repair;
  });

  // ---- ANCHOR: google review (G1:I3) ----
  var googleReview = parseGoogleReview_(rows);

  // ---- ANCHOR: KPI grid — target!B12:L14 ----
  //   Row 12 (B12:L12) = revenue type / category the KPI belongs to (AR/RR/SR)
  //   Row 13 (B13:L13) = KPI name        (Care+, AMC, Display, SRD, ...)
  //   Row 14 (B14:L14) = target count for that KPI
  // 0-indexed: rows 11 / 12 / 13, columns B=1 .. L=11.
  var KPI_FIRST_COL = 1;  // B
  var KPI_LAST_COL = 11;  // L

  // Maps the KPI name written on the sheet to the internal rule code.
  function kpiCodeFor_(label) {
    var l = String(label || "").toUpperCase().replace(/[^A-Z0-9+]/g, "");
    if (!l) return null;
    if (l.indexOf("CARE") === 0) return "CAREPLUS";
    if (l.indexOf("PROTECT") === 0) return "PROTECTPLUS";
    if (l.indexOf("AMC") > -1) return "AMC";
    if (l.indexOf("BUNDLE") > -1) return "BUNDLE";
    if (l.indexOf("DISPLAY") > -1) return "DISPLAY";
    if (l.indexOf("BATTER") > -1) return "BATTERY";
    if (l.indexOf("GLASS") > -1) return "GLASS";
    if (l.indexOf("OTHER") > -1) return "OTHERS";
    if (l.indexOf("SRD") > -1) return "SRD";
    if (l.indexOf("SHC") > -1) return "SHC";
    if (l.indexOf("ARD") > -1) return "ARD";
    return null;
  }

  // Normalises the revenue-type row into the AR / RR / SR bucket label.
  function kpiCategoryFor_(raw, fallback) {
    var l = String(raw || "").toUpperCase();
    if (l.indexOf("AR") > -1 || l.indexOf("ACCESSOR") > -1) return "AR";
    if (l.indexOf("RR") > -1 || l.indexOf("REPAIR") > -1) return "RR";
    if (l.indexOf("SR") > -1 || l.indexOf("SERVICE") > -1) return "SR";
    return fallback || "AR";
  }

  var kpis = [];
  var lastCategory = "";
  for (var kc = KPI_FIRST_COL; kc <= KPI_LAST_COL; kc++) {
    var typeCell = at(11, kc);          // row 12 — revenue type
    var nameCell = at(12, kc);          // row 13 — KPI name
    var targetCell = at(13, kc);        // row 14 — target count
    if (typeCell) lastCategory = typeCell;   // merged cells only fill the first column
    var code = kpiCodeFor_(nameCell);
    if (!code) continue;
    kpis.push({
      category: kpiCategoryFor_(lastCategory, "AR"),
      code: code,
      label: str_(nameCell) || code,
      target: parseMoney_(targetCell),
      value: 0,
      approxAmount: 0,
    });
  }

  // ---- ANCHOR: incentive % slabs (service col A/B, accessory col C/D), anchored on "achievement" header ----
  var pctAnchor = findRow_(rows, 0, 3, /^(achivement|achievement)$/i);
  var svcCells = [], accCells = [];
  if (pctAnchor >= 0) {
    var limit = Math.min(rows.length - 1, pctAnchor + 12);
    for (var rr = pctAnchor + 1; rr <= limit; rr++) {
      var sLabel = at(rr, 0), sVal = at(rr, 1), aLabel = at(rr, 2), aVal = at(rr, 3);
      if (!sLabel && !aLabel) break;
      if (isBandRow_(sLabel, sVal)) svcCells.push([sLabel, sVal]);
      if (isBandRow_(aLabel, aVal)) accCells.push([aLabel, aVal]);
    }
  }
  var serviceSlabs = parsePctSlabs_(svcCells);
  var accessorySlabs = parsePctSlabs_(accCells);

  // ---- ANCHOR: AMC slabs, anchored on "incentive per amc" header (cols F..K) ----
  var amcAnchor = findRow_(rows, 4, 10, /incentive\s*per\s*amc/i);
  var amcSlabs = [];
  if (amcAnchor >= 0) {
    var amcLimit = Math.min(rows.length - 1, amcAnchor + 12);
    for (var ar = amcAnchor + 1; ar <= amcLimit; ar++) {
      var rangeLabel = at(ar, 5), rateLabel = at(ar, 6);
      var accCount = parseMoney_(at(ar, 8)), accRateLabel = at(ar, 9);
      if (!rangeLabel && !rateLabel && !accRateLabel) break;
      var rateInfo = parseAmcRate_(rateLabel);
      var min, max;
      if (rangeLabel && /^\d/.test(rangeLabel)) {
        var range = parseRangeLabel_(rangeLabel);
        min = range.min; max = range.max;
      } else if (amcSlabs.length > 0) {
        min = amcSlabs[amcSlabs.length - 1].max + 1; max = Infinity;
      } else {
        continue;
      }
      var accRateInfo = parseAmcRate_(accRateLabel);
      var slab = {
        label: rangeLabel || (min + "+"),
        rateLabel: rateLabel,
        rate: rateInfo.rate,
        fixed: rateInfo.fixed,
        min: Math.min(min, OPEN_ENDED),
        max: Math.min(Math.max(max, min), OPEN_ENDED),
        accCount: accCount > 0 ? accCount : undefined,
        accRate: accRateInfo.rate > 0 ? accRateInfo.rate : undefined,
        accFixed: accRateInfo.fixed > 0 ? accRateInfo.fixed : undefined,
      };
      amcSlabs.push(slab);
    }
  }

  // ---- ANCHOR: bundle slabs, "slab level" row + "slab incentive" row below it (cols M..P) ----
  var bundleSlabs = [];
  var levelCell = findCell_(rows, 12, 17, /slab\s*level/i);
  if (levelCell) {
    for (var bc = levelCell.col + 1; bc <= levelCell.col + 3; bc++) {
      var level = parseMoney_(at(levelCell.row, bc));
      var per = parseMoney_(at(levelCell.row + 1, bc));
      if (level > 0 && per > 0) bundleSlabs.push({ label: "slab level " + level, min: level, perBundle: per });
    }
  }
  bundleSlabs.sort(function (a, b) { return a.min - b.min; });

  return {
    store: store,
    salespeople: people,
    serviceIncentive: { service: serviceSlabs, accessory: accessorySlabs },
    amcSlabs: amcSlabs,
    bundleSlabs: bundleSlabs,
    kpis: kpis,
    googleReview: googleReview,
  };
}

/* ============================================================================
 * DASHBOARD COMPUTATION  (mirrors metrics.ts computeDashboard / computeKpis)
 * ========================================================================= */

function daysInCurrentMonth_() {
  var now = new Date();
  return new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();
}

function isWorkingDay_(day) {
  return CONFIG.HOLIDAYS.indexOf(day) === -1;
}

function computeDashboard(allRows, target) {
  // Rule: only crm!TranType === "SALES" rows count toward anything.
  var rows = allRows.filter(isSaleRow_);

  var daysInMonth = daysInCurrentMonth_();
  var workingDays = [];
  for (var d = 1; d <= daysInMonth; d++) if (isWorkingDay_(d)) workingDays.push(d);

  var maxDay = 0;
  rows.forEach(function (r) { if (r.day > maxDay) maxDay = r.day; });
  var daysElapsed = Math.max(1, Math.min(daysInMonth, maxDay || 1));
  var daysRemaining = daysInMonth - daysElapsed;
  var activeDays = workingDays.filter(function (dd) { return dd <= daysElapsed; }).length;
  var remainingWorkingAfter = workingDays.filter(function (dd) { return dd > daysElapsed; }).length;
  var remainingWorkingToday = workingDays.filter(function (dd) { return dd >= daysElapsed; }).length;

  var totalWorkingDays = workingDays.length;
  var consumedWorkingDays = activeDays;
  var balanceWorkingDays = Math.max(0, totalWorkingDays - consumedWorkingDays);

  // ---- dynamic salesperson roster ----
  var peopleKeySet = {};
  rows.forEach(function (r) {
    var pk = normalizePerson_(r.salesPerson);
    if (pk) peopleKeySet[pk] = true;
  });
  Object.keys(target.salespeople).forEach(function (k) { if (k) peopleKeySet[k] = true; });
  var peopleKeys = Object.keys(peopleKeySet).sort();
  if (peopleKeys.length === 0) peopleKeys = ["SUBHA", "DINESH"];

  // ---- invoice-level indexes ----
  var invoiceSet = {};
  var invoiceDay = {};
  var invoicePerson = {};
  var invoiceHasServiceOrRepair = {};
  var invoiceAccessoryValue = {};
  var invoiceAccessoryFlag = {};
  var invoiceSchemeAll = {}; // bundle: SchemeCode starts with "DISC"
  var invoiceAmcAll = {}; // SubCategory contains "AMC"
  var invoicePersonSet = {};
  var invoiceLinesByType = { service: 0, accessory: 0, repair: 0 };

  var dayMap = {};
  function ensureDay(day) {
    if (!dayMap[day]) {
      var sample = null;
      for (var i = 0; i < rows.length; i++) { if (rows[i].day === day) { sample = rows[i]; break; } }
      var dateKey = sample ? sample.dayKey : String(day);
      var weekday = "";
      if (dateKey && dateKey.indexOf("-") > -1) {
        var parts = dateKey.split("-");
        if (parts.length === 3) {
          var jsDate = new Date(Number(parts[2]), Number(parts[1]) - 1, Number(parts[0]));
          if (!isNaN(jsDate.getTime())) weekday = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"][jsDate.getDay()];
        }
      }
      dayMap[day] = {
        day: day, dateKey: dateKey, weekday: weekday, holiday: !isWorkingDay_(day), undated: day <= 0,
        uniqueInvoices: 0, serviceRevenue: 0, accessoryRevenue: 0, repairRevenue: 0, totalRevenue: 0,
        grossProfit: 0, byPerson: {},
      };
    }
    return dayMap[day];
  }

  rows.forEach(function (r) {
    var day = ensureDay(r.day);
    invoiceSet[r.invoiceCode] = true;
    invoiceDay[r.invoiceCode] = r.day;
    var pk = normalizePerson_(r.salesPerson) || "UNASSIGNED";
    invoicePerson[r.invoiceCode] = pk;
    if (!invoicePersonSet[r.invoiceCode]) invoicePersonSet[r.invoiceCode] = {};
    invoicePersonSet[r.invoiceCode][pk] = true;

    var rt = revenueTypeKey_(r.revenueType);
    if (rt) invoiceLinesByType[rt] += 1;
    if (rt === "service" || rt === "repair") invoiceHasServiceOrRepair[r.invoiceCode] = true;
    if (rt === "accessory") {
      invoiceAccessoryFlag[r.invoiceCode] = true;
      var prev = invoiceAccessoryValue[r.invoiceCode] || { net: 0, gp: 0, count: 0 };
      invoiceAccessoryValue[r.invoiceCode] = {
        net: prev.net + r.taxableAmount, gp: prev.gp + r.grossProfit, count: prev.count + 1,
      };
    }
    if (r.schemeCode && r.schemeCode.toUpperCase().indexOf("DISC") === 0) invoiceSchemeAll[r.invoiceCode] = true;
    if (r.subCategory && r.subCategory.toUpperCase().indexOf("AMC") > -1) invoiceAmcAll[r.invoiceCode] = true;

    var value = r.taxableAmount;
    if (value >= 0) {
      day.serviceRevenue += rt === "service" ? value : 0;
      day.accessoryRevenue += rt === "accessory" ? value : 0;
      day.repairRevenue += rt === "repair" ? value : 0;
      day.totalRevenue += value;
      day.grossProfit += r.grossProfit;
      var bp = day.byPerson[pk] || {
        serviceRevenue: 0, accessoryRevenue: 0, repairRevenue: 0, totalRevenue: 0, grossProfit: 0,
        // per revenue type: { value, gp, invoices:{code:true} }
        types: {
          service: { value: 0, gp: 0, invoices: {} },
          accessory: { value: 0, gp: 0, invoices: {} },
          repair: { value: 0, gp: 0, invoices: {} },
        },
        invoices: {},
      };
      bp.serviceRevenue += rt === "service" ? value : 0;
      bp.accessoryRevenue += rt === "accessory" ? value : 0;
      bp.repairRevenue += rt === "repair" ? value : 0;
      bp.totalRevenue += value;
      bp.grossProfit += r.grossProfit;
      bp.invoices[r.invoiceCode] = true;
      if (rt && bp.types[rt]) {
        bp.types[rt].value += value;
        bp.types[rt].gp += r.grossProfit;
        bp.types[rt].invoices[r.invoiceCode] = true;
      }
      day.byPerson[pk] = bp;
    }
  });

  var dayUnique = {};
  Object.keys(invoiceSet).forEach(function (code) {
    var dd = invoiceDay[code] || 0;
    if (!dayUnique[dd]) dayUnique[dd] = {};
    dayUnique[dd][code] = true;
  });
  Object.keys(dayUnique).forEach(function (dd) {
    ensureDay(Number(dd)).uniqueInvoices = Object.keys(dayUnique[dd]).length;
  });

  var personInvoiceTotals = {};
  Object.keys(invoiceSet).forEach(function (code) {
    Object.keys(invoicePersonSet[code] || {}).forEach(function (pk) {
      if (!personInvoiceTotals[pk]) personInvoiceTotals[pk] = {};
      personInvoiceTotals[pk][code] = true;
    });
  });

  var days = Object.keys(dayMap).map(function (k) { return dayMap[k]; }).sort(function (a, b) {
    return (a.undated ? 99 : a.day) - (b.undated ? 99 : b.day);
  });

  // Flatten each day's per-person/per-revenue-type detail into rows that the
  // date-wise report can render directly.
  var REV_TYPE_ORDER = [["service", "Service"], ["accessory", "Accessory"], ["repair", "Repair"]];
  days.forEach(function (d) {
    var entries = [];
    Object.keys(d.byPerson).sort().forEach(function (pk) {
      var bp = d.byPerson[pk];
      var typeRows = [];
      REV_TYPE_ORDER.forEach(function (pair) {
        var tk = pair[0], tlabel = pair[1];
        var t = bp.types[tk];
        var invCount = Object.keys(t.invoices).length;
        if (invCount === 0 && t.value === 0) return;
        typeRows.push({
          type: tk, typeLabel: tlabel, invoiceCount: invCount,
          value: Math.round(t.value), gp: Math.round(t.gp),
        });
      });
      entries.push({
        key: pk, name: titleCase_(pk),
        invoiceCount: Object.keys(bp.invoices).length,
        totalRevenue: Math.round(bp.totalRevenue),
        grossProfit: Math.round(bp.grossProfit),
        types: typeRows,
      });
    });
    d.people = entries;
  });

  // ---- store-level revenue rollups (TaxableAmount = achieved value) ----
  function sum(rowsArr, filterFn) {
    var acc = 0;
    rowsArr.forEach(function (r) {
      if (r.taxableAmount >= 0 && (!filterFn || filterFn(r))) acc += r.taxableAmount;
    });
    return acc;
  }
  var sStore = {
    service: sum(rows, function (r) { return revenueTypeKey_(r.revenueType) === "service"; }),
    accessory: sum(rows, function (r) { return revenueTypeKey_(r.revenueType) === "accessory"; }),
    repair: sum(rows, function (r) { return revenueTypeKey_(r.revenueType) === "repair"; }),
    total: sum(rows, null),
    grossProfit: 0,
  };
  rows.forEach(function (r) { if (r.taxableAmount >= 0) sStore.grossProfit += r.grossProfit; });

  function bucket(tgt, ach, remAfter) {
    var runRate = safeDiv_(ach, activeDays);
    return {
      target: tgt, achieved: ach, pct: pct_(ach, tgt), balance: Math.max(0, tgt - ach),
      runRate: Math.round(runRate),
      requiredPerDay: Math.round(safeDiv_(Math.max(0, tgt - ach), Math.max(1, remAfter + 1))),
      projected: Math.round(runRate * daysInMonth),
    };
  }

  var store = {
    daysElapsed: daysElapsed, daysRemaining: daysRemaining, remainingWorkingDays: remainingWorkingToday,
    daysInMonth: daysInMonth,
    totalWorkingDays: totalWorkingDays, consumedWorkingDays: consumedWorkingDays, balanceWorkingDays: balanceWorkingDays,
    totalInvoices: rows.length, uniqueInvoices: Object.keys(invoiceSet).length,
    total: bucket(target.store.total, sStore.total, remainingWorkingAfter),
    service: bucket(target.store.service, sStore.service, remainingWorkingAfter),
    accessory: bucket(target.store.accessory, sStore.accessory, remainingWorkingAfter),
    repair: bucket(target.store.repair, sStore.repair, remainingWorkingAfter),
    grossProfit: bucket(target.store.grossProfit, sStore.grossProfit, remainingWorkingAfter),
  };

  // Daily target = total store target ÷ number of working days in the month.
  // Used to colour the date-wise achieved values green (>= target) or red.
  store.dailyTarget = Math.round(safeDiv_(target.store.total, Math.max(1, totalWorkingDays)));

  // ---- revenue type totals & unique invoice counts ----
  var revenueTypeDefs = [["service", "Service Revenue"], ["accessory", "Accessories Revenue"], ["repair", "Repair Revenue"]];
  var revenueTypes = revenueTypeDefs.map(function (pair) {
    var key = pair[0], label = pair[1];
    var set = {}, revenue = 0, gp = 0;
    rows.forEach(function (r) {
      if (r.taxableAmount < 0) return;
      if (revenueTypeKey_(r.revenueType) !== key) return;
      set[r.invoiceCode] = true; revenue += r.taxableAmount; gp += r.grossProfit;
    });
    return {
      key: key, label: label, invoiceLines: invoiceLinesByType[key], uniqueInvoices: Object.keys(set).length,
      revenue: Math.round(revenue), gp: Math.round(gp),
    };
  });

  var invoiceValuation = {
    totalInvoiceLines: rows.length, totalUniqueInvoices: Object.keys(invoiceSet).length, byType: revenueTypes,
    bundleUniqueInvoices: Object.keys(invoiceSchemeAll).length, amcUniqueInvoices: Object.keys(invoiceAmcAll).length,
  };

  // ---- salesperson revenue & incentive ----
  var salespeople = peopleKeys.map(function (key) {
    var t = target.salespeople[key] || { service: 0, accessory: 0, repair: 0, total: 0, grossProfit: 0 };
    var mine = rows.filter(function (r) { return normalizePerson_(r.salesPerson) === key; });
    var mineInvoices = personInvoiceTotals[key] ? Object.keys(personInvoiceTotals[key]).length : 0;
    var ach = {
      service: sum(mine, function (r) { return revenueTypeKey_(r.revenueType) === "service"; }),
      accessory: sum(mine, function (r) { return revenueTypeKey_(r.revenueType) === "accessory"; }),
      repair: sum(mine, function (r) { return revenueTypeKey_(r.revenueType) === "repair"; }),
    };
    ach.total = ach.service + ach.accessory + ach.repair;

    // Service & accessory are the only revenue types that EARN incentive.
    var svcInc = incentiveForPct_(target.serviceIncentive.service, pct_(ach.service, t.service));
    var accInc = incentiveForPct_(target.serviceIncentive.accessory, pct_(ach.accessory, t.accessory));
    var earnedIncentive = svcInc + accInc;

    // Repair earns nothing itself — it is a GATE. The person only actually
    // receives the earned service+accessory incentive once repair revenue
    // reaches CONFIG.REPAIR_INCENTIVE_GATE_PCT of their repair target.
    var repairPct = pct_(ach.repair, t.repair);
    var gatePct = CONFIG.REPAIR_INCENTIVE_GATE_PCT;
    var gateNeeded = Math.round((gatePct / 100) * t.repair);
    var repairGateMet = t.repair > 0 ? repairPct >= gatePct : false;
    var payableIncentive = repairGateMet ? earnedIncentive : 0;

    function mk(tg, a, inc) {
      return { target: tg, achieved: a, pct: pct_(a, tg), balance: Math.max(0, tg - a), incentive: inc || 0 };
    }

    return {
      key: key, name: titleCase_(key), invoiceCount: mineInvoices,
      total: mk(t.total, ach.total, payableIncentive),
      service: mk(t.service, ach.service, svcInc),
      accessory: mk(t.accessory, ach.accessory, accInc),
      repair: mk(t.repair, ach.repair, 0),
      // incentive gate details, surfaced in the UI
      repairGate: {
        requiredPct: gatePct,
        achievedPct: repairPct,
        requiredValue: gateNeeded,
        shortfall: Math.max(0, gateNeeded - ach.repair),
        met: repairGateMet,
      },
      earnedIncentive: earnedIncentive,
      payableIncentive: payableIncentive,
      totalIncentive: payableIncentive,
    };
  });

  // ---- AMC & Bundle tracking (per person) ----
  var amcByPerson = {}, accAmcByPerson = {}, bundleByPerson = {};
  rows.forEach(function (r) {
    var pk = normalizePerson_(r.salesPerson);
    if (r.subCategory.toUpperCase().indexOf("AMC") > -1 && r.taxableAmount >= 0) {
      amcByPerson[pk] = amcByPerson[pk] || {}; amcByPerson[pk][r.invoiceCode] = true;
      if (!invoiceHasServiceOrRepair[r.invoiceCode]) {
        accAmcByPerson[pk] = accAmcByPerson[pk] || {}; accAmcByPerson[pk][r.invoiceCode] = true;
      }
    }
    if (r.schemeCode.toUpperCase().indexOf("DISC") === 0 && r.taxableAmount >= 0) {
      bundleByPerson[pk] = bundleByPerson[pk] || {}; bundleByPerson[pk][r.invoiceCode] = true;
    }
  });

  var accTiers = target.amcSlabs.filter(function (s) { return (s.accCount || 0) > 0 && (s.accRate || 0) > 0; });
  function accPayout(t, count) { return count * (t.accRate || 0) + (t.accFixed || 0); }

  var amc = peopleKeys.map(function (key) {
    var amcCount = amcByPerson[key] ? Object.keys(amcByPerson[key]).length : 0;
    var accessoryCount = accAmcByPerson[key] ? Object.keys(accAmcByPerson[key]).length : 0;
    var slabs = target.amcSlabs;
    var eligibleSlabs = slabs.filter(function (s) { return amcCount >= s.min; });
    var amcSlab = eligibleSlabs.length ? eligibleSlabs[eligibleSlabs.length - 1] : null;
    var nextSlabs = slabs.filter(function (s) { return s.min > amcCount && s.rate > 0; }).sort(function (a, b) { return a.min - b.min; });
    var amcNextSlab = nextSlabs.length ? nextSlabs[0] : null;
    var amcEligible = amcSlab ? amcCount * amcSlab.rate + amcSlab.fixed : 0;
    var amcRemaining = amcNextSlab ? Math.max(0, amcNextSlab.min - amcCount) : 0;

    var accEligibleTiers = accTiers.filter(function (t) { return accessoryCount >= (t.accCount || 0); })
      .sort(function (a, b) { return accPayout(b, accessoryCount) - accPayout(a, accessoryCount); });
    var accSlab = accEligibleTiers.length ? accEligibleTiers[0] : null;
    var accNextSlabs = accTiers.filter(function (t) { return (t.accCount || 0) > accessoryCount; })
      .sort(function (a, b) { return (a.accCount || 0) - (b.accCount || 0); });
    var accNextSlab = accNextSlabs.length ? accNextSlabs[0] : null;
    var accEligible = accSlab ? accPayout(accSlab, accessoryCount) : 0;
    var accRemaining = accNextSlab ? Math.max(0, (accNextSlab.accCount || 0) - accessoryCount) : 0;

    return {
      key: key, name: titleCase_(key), amcCount: amcCount, amcSlab: amcSlab, amcNextSlab: amcNextSlab,
      amcRemaining: amcRemaining, amcEligible: amcEligible, accessoryCount: accessoryCount,
      accessorySlab: accSlab, accessoryNextSlab: accNextSlab, accessoryRemaining: accRemaining,
      accessoryEligible: accEligible, projectedIncentive: amcEligible + accEligible,
    };
  });

  var bundle = peopleKeys.map(function (key) {
    var count = bundleByPerson[key] ? Object.keys(bundleByPerson[key]).length : 0;
    var slabs = target.bundleSlabs;
    var eligible = slabs.filter(function (s) { return count >= s.min; });
    var slab = eligible.length ? eligible[eligible.length - 1] : null;
    var nextSlabs = slabs.filter(function (s) { return s.min > count; }).sort(function (a, b) { return a.min - b.min; });
    var nextSlab = nextSlabs.length ? nextSlabs[0] : null;
    return {
      key: key, name: titleCase_(key), count: count, slab: slab, nextSlab: nextSlab,
      remaining: nextSlab ? Math.max(0, nextSlab.min - count) : 0,
      projectedIncentive: slab ? count * slab.perBundle : 0,
    };
  });

  var kpi = computeKpis(rows, target, remainingWorkingToday);

  return {
    crmRows: rows, target: target, days: days, activeDays: activeDays, store: store,
    revenueTypes: revenueTypes, invoiceValuation: invoiceValuation, salespeople: salespeople,
    amc: amc, bundle: bundle, kpi: kpi, generatedAt: new Date(),
  };
}

/* ============================================================================
 * KPI RULE ENGINE
 * ========================================================================= */

function computeKpis(crmRows, target, remainingWorkingDays) {
  var kpiByKey = {};
  target.kpis.forEach(function (k) { kpiByKey[k.code] = k; });

  function getKpi(code) {
    var k = kpiByKey[code];
    return { category: k ? k.category : "AR", code: code, label: k ? k.label : code,
      target: k ? k.target : 0, value: k ? k.value : 0, approxAmount: k ? k.approxAmount : 0 };
  }

  var invoicePerson = {};
  crmRows.forEach(function (r) {
    var pk = normalizePerson_(r.salesPerson);
    if (pk && !invoicePerson[r.invoiceCode]) invoicePerson[r.invoiceCode] = pk;
  });

  function mkRow(k, invoiceSetObj) {
    var achievedValue = 0, achievedGp = 0;
    crmRows.forEach(function (r) {
      if (r.taxableAmount >= 0 && invoiceSetObj[r.invoiceCode]) { achievedValue += r.taxableAmount; achievedGp += r.grossProfit; }
    });
    var codes = Object.keys(invoiceSetObj);
    var achieved = codes.length;
    var byPerson = {};
    codes.forEach(function (code) {
      var pk = invoicePerson[code];
      if (pk) byPerson[pk] = (byPerson[pk] || 0) + 1;
    });
    return {
      category: k.category, code: k.code, label: k.label, target: k.target, value: k.value, approxAmount: k.approxAmount,
      achieved: achieved, balance: Math.max(0, k.target - achieved), achievedValue: Math.round(achievedValue),
      achievedGp: Math.round(achievedGp), pct: pct_(achieved, k.target), byPerson: byPerson,
    };
  }

  // ---- AR KPIs ----
  var careplus = {}, protectplus = {}, amcKpi = {}, bundleKpi = {};
  crmRows.forEach(function (r) {
    if (r.taxableAmount < 0) return;
    var pt = r.partType.toUpperCase();
    var d = r.itemDescription.toUpperCase();
    if (pt === "ACPLUS" || d.indexOf("CARE+") > -1 || d.indexOf("CAREPLUS") > -1) careplus[r.invoiceCode] = true;
    if (pt === "PROTECTPLUS" || d.indexOf("PROTECT+") > -1 || d.indexOf("PROTECTPLUS") > -1) protectplus[r.invoiceCode] = true;
    if (r.subCategory.toUpperCase().indexOf("AMC") > -1) amcKpi[r.invoiceCode] = true;
    if (r.schemeCode.toUpperCase().indexOf("DISC") === 0) bundleKpi[r.invoiceCode] = true;
  });

  // ---- RR KPIs — ItemDescription rule: "iphone" AND the part keyword, on repair rows ----
  var rrDisplay = {}, rrBattery = {}, rrGlass = {}, rrOthers = {};
  crmRows.forEach(function (r) {
    if (r.taxableAmount < 0) return;
    if (revenueTypeKey_(r.revenueType) !== "repair") return;
    var d = r.itemDescription.toUpperCase();
    var isIphone = d.indexOf("IPHONE") > -1;
    var isDisplay = isIphone && d.indexOf("DISPLAY") > -1;
    var isBattery = isIphone && d.indexOf("BATTERY") > -1;
    var isGlass = isIphone && d.indexOf("GLASS") > -1;
    if (isDisplay) rrDisplay[r.invoiceCode] = true;
    if (isBattery) rrBattery[r.invoiceCode] = true;
    if (isGlass) rrGlass[r.invoiceCode] = true;
    if (!isDisplay && !isBattery && !isGlass) rrOthers[r.invoiceCode] = true;
  });

  // ---- SR KPIs (ItemCode prefix SRD/SHC/ARD, if that column exists) ----
  var srSrd = {}, srShc = {}, srArd = {};
  crmRows.forEach(function (r) {
    if (r.taxableAmount < 0) return;
    var ic = (r.itemCode || "").toUpperCase();
    if (!ic) return;
    if (ic.indexOf("SRD") === 0) srSrd[r.invoiceCode] = true;
    if (ic.indexOf("SHC") === 0) srShc[r.invoiceCode] = true;
    if (ic.indexOf("ARD") === 0) srArd[r.invoiceCode] = true;
  });

  var ar = [mkRow(getKpi("CAREPLUS"), careplus), mkRow(getKpi("PROTECTPLUS"), protectplus),
    mkRow(getKpi("AMC"), amcKpi), mkRow(getKpi("BUNDLE"), bundleKpi)];
  var rr = [mkRow(getKpi("DISPLAY"), rrDisplay), mkRow(getKpi("BATTERY"), rrBattery),
    mkRow(getKpi("GLASS"), rrGlass), mkRow(getKpi("OTHERS"), rrOthers)];
  var sr = [mkRow(getKpi("SRD"), srSrd), mkRow(getKpi("SHC"), srShc), mkRow(getKpi("ARD"), srArd)];

  var workingDays = Math.max(1, remainingWorkingDays);
  var dailyIdeas = ar.concat(rr, sr).filter(function (k) { return k.target > 0 && k.balance > 0; })
    .map(function (k) {
      return { code: k.code, label: k.label, category: k.category, target: k.target, achieved: k.achieved,
        balance: k.balance, perDay: Math.ceil(k.balance / workingDays) };
    }).sort(function (a, b) { return b.balance - a.balance; });

  var gr = target.googleReview;
  return {
    ar: ar, rr: rr, sr: sr,
    googleReview: { previousMonth: gr.previousMonth, currentMonth: gr.currentMonth,
      difference: gr.difference || Math.max(0, gr.currentMonth - gr.previousMonth) },
    dailyIdeas: dailyIdeas,
  };
}

/* ============================================================================
 * DASHBOARD SHEET WRITER  (colorful, formatted output)
 * ========================================================================= */

var COLORS = {
  headerBg: "#1f2a44", headerFg: "#ffffff",
  sectionBg: "#2563eb", sectionFg: "#ffffff",
  subHeaderBg: "#dbeafe", subHeaderFg: "#1e3a8a",
  good: "#16a34a", warn: "#d97706", bad: "#dc2626",
  stripe: "#f8fafc", border: "#cbd5e1",
  totalBg: "#e2e8f0",
};

function money_(n) { return "₹" + Math.round(n || 0).toLocaleString("en-IN"); }

function writeDashboard(ss, data) {
  var sheet = ss.getSheetByName(CONFIG.DASHBOARD_SHEET_NAME);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(CONFIG.DASHBOARD_SHEET_NAME, 0);
  sheet.setHiddenGridlines(true);

  var row = 1;

  // ---- Title banner ----
  sheet.getRange(row, 1, 1, 10).merge().setValue("📊  SALES PERFORMANCE DASHBOARD")
    .setFontSize(16).setFontWeight("bold").setFontColor(COLORS.headerFg)
    .setBackground(COLORS.headerBg).setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.setRowHeight(row, 34);
  row++;
  sheet.getRange(row, 1, 1, 10).merge()
    .setValue("Generated " + Utilities.formatDate(data.generatedAt, Session.getScriptTimeZone(), "dd MMM yyyy, HH:mm") +
      "  ·  " + data.store.totalInvoices + " SALES rows  ·  " + data.store.uniqueInvoices + " unique invoices")
    .setFontStyle("italic").setFontColor("#475569").setHorizontalAlignment("center");
  row += 2;

  row = writeStoreSection_(sheet, row, data);
  row = writeWorkingDaysSection_(sheet, row, data);
  row = writeSalespersonSection_(sheet, row, data);
  row = writeIncentiveSection_(sheet, row, data);
  row = writeKpiSection_(sheet, row, data);
  row = writeInvoiceValuationSection_(sheet, row, data);
  row = writeDateWiseSection_(sheet, row, data);

  sheet.setFrozenRows(2);
  for (var c = 1; c <= 12; c++) sheet.autoResizeColumn(c);
  sheet.setColumnWidth(1, 190);
  ss.setActiveSheet(sheet);
  sheet.getRange(1, 1).activate();
}

function sectionHeader_(sheet, row, title, span) {
  var rng = sheet.getRange(row, 1, 1, span || 8);
  rng.merge().setValue(title).setBackground(COLORS.sectionBg).setFontColor(COLORS.sectionFg)
    .setFontWeight("bold").setFontSize(12).setVerticalAlignment("middle");
  sheet.setRowHeight(row, 26);
  return row + 1;
}

function tableHeaderRow_(sheet, row, headers) {
  var rng = sheet.getRange(row, 1, 1, headers.length);
  rng.setValues([headers]).setBackground(COLORS.subHeaderBg).setFontColor(COLORS.subHeaderFg)
    .setFontWeight("bold").setHorizontalAlignment("center");
  return row + 1;
}

function pctColor_(pctVal) {
  if (pctVal >= 100) return COLORS.good;
  if (pctVal >= 70) return "#2563eb";
  if (pctVal >= 40) return COLORS.warn;
  return COLORS.bad;
}

function writeStoreSection_(sheet, row, data) {
  row = sectionHeader_(sheet, row, "1 · STORE TARGET vs ACHIEVEMENT", 6);
  row = tableHeaderRow_(sheet, row, ["Revenue Type", "Target", "Achieved", "Achieved %", "Balance", "Req / Day"]);
  var s = data.store;
  var buckets = [["Total Revenue", s.total], ["Service", s.service], ["Accessories", s.accessory],
    ["Repair", s.repair], ["Gross Profit", s.grossProfit]];
  buckets.forEach(function (pair) {
    var label = pair[0], b = pair[1];
    var r = sheet.getRange(row, 1, 1, 6);
    r.setValues([[label, money_(b.target), money_(b.achieved), b.pct + "%", money_(b.balance), money_(b.requiredPerDay)]]);
    sheet.getRange(row, 4).setFontColor(pctColor_(b.pct)).setFontWeight("bold");
    sheet.getRange(row, 1, 1, 6).setBackground(row % 2 === 0 ? COLORS.stripe : "#ffffff");
    row++;
  });
  // Google review — read from target!A9:C10
  var gr = data.kpi.googleReview;
  sheet.getRange(row, 1, 1, 6).setValues([["Google Reviews", "prev month: " + gr.previousMonth,
    "this month: " + gr.currentMonth, (gr.difference >= 0 ? "+" : "") + gr.difference, "", ""]]);
  sheet.getRange(row, 1).setFontWeight("bold");
  sheet.getRange(row, 4).setFontWeight("bold").setFontColor(gr.difference >= 0 ? COLORS.good : COLORS.bad);
  sheet.getRange(row, 1, 1, 6).setBackground("#eef2f9");
  row++;
  return row + 1;
}

function writeWorkingDaysSection_(sheet, row, data) {
  row = sectionHeader_(sheet, row, "2 · WORKING DAYS (CURRENT MONTH)", 4);
  row = tableHeaderRow_(sheet, row, ["Total Working Days", "Consumed", "Balance", "Days Elapsed / In Month"]);
  var s = data.store;
  sheet.getRange(row, 1, 1, 4).setValues([[s.totalWorkingDays, s.consumedWorkingDays, s.balanceWorkingDays,
    s.daysElapsed + " / " + s.daysInMonth]]);
  sheet.getRange(row, 3).setFontColor(s.balanceWorkingDays > 0 ? COLORS.warn : COLORS.good).setFontWeight("bold");
  row++;
  return row + 1;
}

function writeSalespersonSection_(sheet, row, data) {
  row = sectionHeader_(sheet, row, "3 · SALESPERSON PERFORMANCE (per revenue type)", 8);
  row = tableHeaderRow_(sheet, row,
    ["Salesperson", "Type", "Target", "Achieved", "Achieved %", "Balance", "Incentive", "Invoices"]);
  data.salespeople.forEach(function (p, idx) {
    var types = [["Total", p.total], ["Service", p.service], ["Accessory", p.accessory], ["Repair", p.repair]];
    types.forEach(function (pair, ti) {
      var label = pair[0], b = pair[1];
      var incTxt;
      if (label === "Repair") incTxt = "— (gate only)";
      else if (label === "Total") incTxt = money_(p.payableIncentive) + (p.repairGate.met ? "" : " (locked)");
      else incTxt = b.incentive ? money_(b.incentive) : "—";
      var vals = [ti === 0 ? p.name : "", label, money_(b.target), money_(b.achieved), b.pct + "%",
        money_(b.balance), incTxt, ti === 0 ? p.invoiceCount : ""];
      sheet.getRange(row, 1, 1, 8).setValues([vals]);
      sheet.getRange(row, 5).setFontColor(pctColor_(b.pct)).setFontWeight("bold");
      sheet.getRange(row, 1, 1, 8).setBackground(idx % 2 === 0 ? COLORS.stripe : "#ffffff");
      if (ti === 0) sheet.getRange(row, 1).setFontWeight("bold");
      row++;
    });
  });
  // Repair incentive gate summary
  var gateLines = data.salespeople.map(function (p) {
    return p.name + ": repair " + p.repairGate.achievedPct.toFixed(1) + "% of target" +
      (p.repairGate.met ? " — incentive unlocked (" + money_(p.payableIncentive) + ")"
        : " — LOCKED, needs " + money_(p.repairGate.shortfall) + " more");
  }).join("   |   ");
  sheet.getRange(row, 1, 1, 8).merge()
    .setValue("Incentive gate (" + CONFIG.REPAIR_INCENTIVE_GATE_PCT + "% repair target): " + gateLines)
    .setFontSize(9).setFontStyle("italic").setFontColor("#475569").setWrap(true);
  row++;
  return row + 1;
}

function writeIncentiveSection_(sheet, row, data) {
  row = sectionHeader_(sheet, row, "4 · INCENTIVE VALUATION — AMC & BUNDLE", 9);
  row = tableHeaderRow_(sheet, row,
    ["Salesperson", "AMC Count", "AMC Eligible", "AMC To Next Slab", "Acc-AMC Count", "Acc-AMC Eligible",
      "Bundle Count", "Bundle Eligible", "Total Incentive"]);
  var bundleByKey = {};
  data.bundle.forEach(function (b) { bundleByKey[b.key] = b; });
  data.amc.forEach(function (a, idx) {
    var b = bundleByKey[a.key] || { count: 0, projectedIncentive: 0 };
    var totalInc = a.projectedIncentive + b.projectedIncentive;
    sheet.getRange(row, 1, 1, 9).setValues([[a.name, a.amcCount, money_(a.amcEligible),
      a.amcNextSlab ? a.amcRemaining + " more" : "top slab", a.accessoryCount, money_(a.accessoryEligible),
      b.count, money_(b.projectedIncentive), money_(totalInc)]]);
    sheet.getRange(row, 1, 1, 9).setBackground(idx % 2 === 0 ? COLORS.stripe : "#ffffff");
    sheet.getRange(row, 9).setFontWeight("bold").setFontColor(COLORS.good);
    row++;
  });
  return row + 1;
}

function writeKpiSection_(sheet, row, data) {
  // Targets come from target!B12:L14. Google Reviews is NOT listed here —
  // it is reported separately from target!A9:C10.
  row = sectionHeader_(sheet, row, "5 · STORE KPI PARAMETERS (target!B12:L14)", 6);
  row = tableHeaderRow_(sheet, row, ["KPI", "Type", "Target", "Achieved", "Balance", "Achieved Value / GP"]);
  var allKpis = data.kpi.ar.concat(data.kpi.rr, data.kpi.sr);
  allKpis.forEach(function (k, idx) {
    sheet.getRange(row, 1, 1, 6).setValues([[k.label, k.category, k.target || "—", k.achieved, k.target > 0 ? k.balance : "—",
      money_(k.achievedValue) + " / " + money_(k.achievedGp)]]);
    sheet.getRange(row, 5).setFontColor(k.target > 0 && k.balance === 0 ? COLORS.good : COLORS.warn).setFontWeight("bold");
    sheet.getRange(row, 1, 1, 6).setBackground(idx % 2 === 0 ? COLORS.stripe : "#ffffff");
    row++;
  });
  return row + 1;
}

function writeInvoiceValuationSection_(sheet, row, data) {
  row = sectionHeader_(sheet, row, "6 · INVOICE VALUATION", 6);
  row = tableHeaderRow_(sheet, row, ["Revenue Type", "Invoice Lines", "Unique Invoices", "Revenue", "Gross Profit", ""]);
  data.revenueTypes.forEach(function (rt, idx) {
    sheet.getRange(row, 1, 1, 5).setValues([[rt.label, rt.invoiceLines, rt.uniqueInvoices, money_(rt.revenue), money_(rt.gp)]]);
    sheet.getRange(row, 1, 1, 5).setBackground(idx % 2 === 0 ? COLORS.stripe : "#ffffff");
    row++;
  });
  sheet.getRange(row, 1, 1, 5).setValues([["Overall", data.invoiceValuation.totalInvoiceLines,
    data.invoiceValuation.totalUniqueInvoices, money_(data.store.total.achieved), money_(data.store.grossProfit.achieved)]])
    .setBackground(COLORS.totalBg).setFontWeight("bold");
  row++;
  sheet.getRange(row, 1, 1, 4).setValues([["Bundle unique invoices (SchemeCode starts DISC)",
    data.invoiceValuation.bundleUniqueInvoices, "AMC unique invoices (SubCategory has AMC)",
    data.invoiceValuation.amcUniqueInvoices]]);
  row += 2;
  return row;
}

function writeDateWiseSection_(sheet, row, data) {
  row = sectionHeader_(sheet, row, "7 · DATE-WISE SALES REVENUE ANALYSIS (by salesperson)", 8);
  var dailyTarget = data.store.dailyTarget;
  sheet.getRange(row, 1, 1, 4).setValues([["Daily Target (store total ÷ working days)", money_(dailyTarget),
    "Working Days", data.store.totalWorkingDays]]).setFontWeight("bold");
  row += 2;

  var header = ["Date", "Salesperson", "Revenue Type", "Invoice Count", "Value", "GP", "Person Total", "Person GP"];
  row = tableHeaderRow_(sheet, row, header);

  // Alternating tints so each salesperson block is easy to distinguish.
  var personTint = ["#eef3ff", "#fdf1e6", "#e9f7f0", "#f5edfe", "#fdecf4"];
  var personIdx = {};
  data.salespeople.forEach(function (p, i) { personIdx[p.key] = i % personTint.length; });

  data.days.forEach(function (d) {
    // Date summary band
    sheet.getRange(row, 1, 1, 8).setValues([[d.undated ? "No date" : d.dateKey,
      d.weekday + (d.holiday ? " (holiday)" : ""), d.uniqueInvoices + " inv", "", money_(d.totalRevenue),
      money_(d.grossProfit), "daily target", money_(dailyTarget)]]);
    sheet.getRange(row, 1, 1, 8).setBackground(d.holiday ? "#fdf0d5" : "#e3e9f4").setFontWeight("bold");
    sheet.getRange(row, 5).setFontColor(d.totalRevenue >= dailyTarget ? COLORS.good : COLORS.bad);
    row++;

    var people = d.people || [];
    if (!people.length) {
      sheet.getRange(row, 1, 1, 8).setValues([["", "No sales recorded", "", "", "", "", "", ""]]);
      sheet.getRange(row, 1, 1, 8).setFontColor("#94a3b8");
      row++;
      return;
    }

    people.forEach(function (p) {
      var tint = personTint[personIdx[p.key] === undefined ? 0 : personIdx[p.key]];
      var types = p.types.length ? p.types : [{ typeLabel: "—", invoiceCount: 0, value: 0, gp: 0 }];
      types.forEach(function (t, ti) {
        sheet.getRange(row, 1, 1, 8).setValues([["", ti === 0 ? p.name : "", t.typeLabel, t.invoiceCount,
          money_(t.value), money_(t.gp), ti === 0 ? money_(p.totalRevenue) : "",
          ti === 0 ? money_(p.grossProfit) : ""]]);
        sheet.getRange(row, 1, 1, 8).setBackground(tint);
        if (ti === 0) sheet.getRange(row, 2).setFontWeight("bold");
        if (ti === 0) {
          sheet.getRange(row, 7).setFontWeight("bold")
            .setFontColor(p.totalRevenue >= dailyTarget ? COLORS.good : COLORS.bad);
        }
        row++;
      });
    });
  });
  return row + 1;
}

/* ============================================================================
 * WEB APP HTML RENDERER  (Mode B — the standalone "script page")
 * Pure server-side string templating — no client-side calls back into
 * Apps Script are needed, so the page works even for anonymous / link
 * viewers. Reload the URL to recompute from the latest sheet data.
 * ========================================================================= */

function esc_(s) {
  return String(s === null || s === undefined ? "" : s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function pctBadgeHtml_(pctVal) {
  var cls = pctVal >= 100 ? "good" : pctVal >= 70 ? "info" : pctVal >= 40 ? "warn" : "bad";
  return '<span class="badge ' + cls + '">' + pctVal.toFixed(1) + "%</span>";
}

function barHtml_(pctVal) {
  var clamped = Math.max(0, Math.min(100, pctVal));
  var cls = pctVal >= 100 ? "good" : pctVal >= 70 ? "info" : pctVal >= 40 ? "warn" : "bad";
  return '<div class="bar"><div class="bar-fill ' + cls + '" style="width:' + clamped + '%"></div></div>';
}

function renderDashboardHtml_(data) {
  var css = getDashboardCss_();
  var body = "";

  body += renderStoreCards_(data);
  body += renderStoreTable_(data);
  body += renderWorkingDaysCard_(data);
  body += renderSalespersonCards_(data);
  body += renderIncentiveTable_(data);
  body += renderKpiSection_(data);
  body += renderInvoiceValuation_(data);
  body += renderDateWiseTable_(data);

  var generated = Utilities.formatDate(data.generatedAt, Session.getScriptTimeZone(), "dd MMM yyyy, HH:mm");

  // The Refresh button re-requests the Web App URL with ?refresh=1, which
  // bypasses the server-side cache and recomputes from the sheet.
  var refreshScript =
    "<script>" +
    "function doRefresh(btn){" +
    "  btn.classList.add('loading'); btn.disabled=true;" +
    "  btn.querySelector('.txt').textContent='Refreshing...';" +
    "  var base=(window.location.href.split('?')[0]);" +
    "  window.location.replace(base+'?refresh=1&t='+Date.now());" +
    "}" +
    "</script>";

  return (
    "<!DOCTYPE html><html><head><base target=\"_top\"><meta charset=\"utf-8\">" +
    "<title>Sales Performance Dashboard</title><style>" + css + "</style></head><body>" +
    '<header class="topbar"><div class="topbar-inner">' +
    "<div>" +
    "<h1>📊 Sales Performance Dashboard</h1>" +
    '<div class="meta">Updated ' + esc_(generated) + " &middot; " + data.store.totalInvoices +
    " SALES rows &middot; " + data.store.uniqueInvoices + " unique invoices</div>" +
    "</div>" +
    '<button class="refresh-btn" onclick="doRefresh(this)">' +
    '<span class="spin"></span><span class="txt">↻ Refresh data</span></button>' +
    "</div></header>" +
    '<main class="wrap">' + body + "</main>" +
    '<footer class="foot">Press “Refresh data” to pull the latest crm / target sheet values.</footer>' +
    refreshScript +
    "</body></html>"
  );
}

function getDashboardCss_() {
  return (
    "* { box-sizing: border-box; }" +
    // ---- mild dark page background, mild light tiles ----
    "body { margin:0; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Arial, sans-serif;" +
    " background:#33405c; color:#1f2937; }" +
    ".topbar { background:#2b3750; color:#fff; padding:16px 20px; border-bottom:1px solid #46536f;" +
    " position:sticky; top:0; z-index:20; }" +
    ".topbar-inner { max-width:1320px; margin:0 auto; display:flex; align-items:center;" +
    " justify-content:space-between; gap:16px; flex-wrap:wrap; }" +
    ".topbar h1 { margin:0 0 4px; font-size:21px; }" +
    ".topbar .meta { font-size:12px; opacity:.82; }" +
    ".refresh-btn { background:#4f7cff; color:#fff; border:none; border-radius:8px; padding:10px 18px;" +
    " font-size:13px; font-weight:700; cursor:pointer; box-shadow:0 1px 4px rgba(0,0,0,.25);" +
    " display:inline-flex; align-items:center; gap:8px; white-space:nowrap; }" +
    ".refresh-btn:hover { background:#3d6af0; }" +
    ".refresh-btn:disabled { opacity:.65; cursor:progress; }" +
    ".refresh-btn .spin { display:none; width:13px; height:13px; border:2px solid rgba(255,255,255,.4);" +
    " border-top-color:#fff; border-radius:50%; animation:spin .7s linear infinite; }" +
    ".refresh-btn.loading .spin { display:inline-block; }" +
    "@keyframes spin { to { transform:rotate(360deg); } }" +
    ".wrap { max-width:1320px; margin:0 auto; padding:20px 16px 50px; }" +
    // mild light tile surfaces
    ".section { background:#f5f7fb; border-radius:12px; box-shadow:0 2px 8px rgba(0,0,0,.18);" +
    " margin-bottom:20px; overflow:hidden; }" +
    ".section h2 { margin:0; padding:13px 18px; font-size:13px; letter-spacing:.06em; text-transform:uppercase;" +
    " background:#43587f; color:#fff; }" +
    ".section .body { padding:16px 18px; }" +
    ".cards { display:grid; grid-template-columns:repeat(auto-fit,minmax(155px,1fr)); gap:12px; margin-bottom:20px; }" +
    ".card { background:#f5f7fb; border-radius:12px; padding:14px 16px; box-shadow:0 2px 8px rgba(0,0,0,.18); }" +
    ".section .card { box-shadow:none; border:1px solid #dde3ee; background:#fbfcfe; }" +
    ".card .label { font-size:10px; text-transform:uppercase; letter-spacing:.08em; color:#5b6880; font-weight:700; }" +
    ".card .value { font-size:22px; font-weight:700; margin-top:4px; color:#1f2937; }" +
    ".card .sub { font-size:11px; color:#5b6880; margin-top:2px; }" +
    "table { width:100%; border-collapse:collapse; font-size:13px; background:#fbfcfe; }" +
    "th { text-align:right; font-size:10px; text-transform:uppercase; letter-spacing:.06em; color:#4a5a75;" +
    " padding:8px 10px; border-bottom:2px solid #d9e0ec; background:#eef2f9; }" +
    "th:first-child, td:first-child { text-align:left; }" +
    "td { padding:8px 10px; border-bottom:1px solid #e6ebf3; text-align:right; }" +
    "tr:nth-child(even) td { background:#f2f5fa; }" +
    "tr.holiday td { background:#fdf0d5 !important; }" +
    "tr.total-row td { background:#dfe6f2 !important; font-weight:700; }" +
    ".badge { display:inline-block; padding:2px 8px; border-radius:999px; font-weight:700; font-size:11px; }" +
    ".badge.good { background:#d7f5df; color:#15803d; }" +
    ".badge.info { background:#dae7fd; color:#1d4ed8; }" +
    ".badge.warn { background:#fdeecb; color:#b45309; }" +
    ".badge.bad { background:#fcdede; color:#b91c1c; }" +
    ".bar { height:6px; border-radius:999px; background:#dfe5ef; overflow:hidden; min-width:80px; }" +
    ".bar-fill { height:100%; border-radius:999px; }" +
    ".bar-fill.good { background:#16a34a; } .bar-fill.info { background:#2563eb; }" +
    ".bar-fill.warn { background:#d97706; } .bar-fill.bad { background:#dc2626; }" +
    ".person-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(300px,1fr)); gap:14px; }" +
    ".person-card { border:1px solid #dde3ee; border-radius:10px; padding:14px; background:#fbfcfe; }" +
    ".person-card h3 { margin:0 0 8px; font-size:15px; }" +
    ".row-line { display:flex; justify-content:space-between; align-items:center; gap:8px; margin:6px 0 2px; font-size:12px; }" +
    ".scroll-x { overflow-x:auto; }" +
    ".foot { text-align:center; color:#b9c3d6; font-size:11px; padding:6px 0 28px; }" +
    ".pos { color:#15803d; font-weight:700; } .neg { color:#b91c1c; font-weight:700; }" +
    ".muted { color:#6b7688; }" +
    // incentive gate note inside person cards
    ".gate { margin-top:10px; padding:8px 10px; border-radius:8px; font-size:11.5px; line-height:1.45; }" +
    ".gate.met { background:#e6f7ec; border:1px solid #b6e6c6; color:#14663a; }" +
    ".gate.blocked { background:#fdecec; border:1px solid #f3c3c3; color:#8f2020; }" +
    ".incentive-line { display:flex; justify-content:space-between; font-size:12px; margin-top:6px; }" +
    // date-wise per-person grouping
    "table.detail td, table.detail th { font-size:12.5px; }" +
    "tr.date-head td { background:#e3e9f4 !important; font-weight:700; color:#26334d;" +
    " border-top:2px solid #b9c6dd; }" +
    "td.person-cell { font-weight:700; border-left:4px solid transparent; }" +
    ".sp-0 { border-left-color:#4f7cff !important; }" +
    ".sp-1 { border-left-color:#ef8b3c !important; }" +
    ".sp-2 { border-left-color:#22a06b !important; }" +
    ".sp-3 { border-left-color:#a855f7 !important; }" +
    ".sp-4 { border-left-color:#ec4899 !important; }" +
    "tr.person-band-0 td { background:#f4f7ff; } tr.person-band-0:nth-child(even) td { background:#eef3ff; }" +
    "tr.person-band-1 td { background:#fff7f0; } tr.person-band-1:nth-child(even) td { background:#fdf1e6; }" +
    "tr.person-band-2 td { background:#f1fbf6; } tr.person-band-2:nth-child(even) td { background:#e9f7f0; }" +
    "tr.person-band-3 td { background:#faf5ff; } tr.person-band-3:nth-child(even) td { background:#f5edfe; }" +
    "tr.person-band-4 td { background:#fef4f9; } tr.person-band-4:nth-child(even) td { background:#fdecf4; }" +
    "td.val-good { color:#15803d; font-weight:700; }" +
    "td.val-bad { color:#b91c1c; font-weight:700; }" +
    "@media (max-width:640px){ .topbar h1{font-size:18px;} .card .value{font-size:18px;} }"
  );
}

function renderStoreCards_(data) {
  var s = data.store;
  function card(label, value, sub) {
    return '<div class="card"><div class="label">' + esc_(label) + '</div><div class="value">' +
      esc_(value) + '</div><div class="sub">' + esc_(sub || "") + "</div></div>";
  }
  var gr = data.kpi.googleReview;
  // Google review tile — values come straight from target!A9:C10
  var diffCls = gr.difference > 0 ? "pos" : gr.difference < 0 ? "neg" : "muted";
  var grTile =
    '<div class="card"><div class="label">Google Reviews</div>' +
    '<div class="value">' + gr.currentMonth +
    ' <span class="' + diffCls + '" style="font-size:13px">(' +
    (gr.difference >= 0 ? "+" : "") + gr.difference + ")</span></div>" +
    '<div class="sub">this month &middot; previous ' + gr.previousMonth + "</div></div>";

  return (
    '<div class="cards">' +
    card("Month Progress", s.daysElapsed + " / " + s.daysInMonth + " days", s.daysRemaining + " remaining") +
    card("Total Invoices", s.totalInvoices, "invoice lines (SALES only)") +
    card("Unique Invoices", s.uniqueInvoices, "deduped invoice codes") +
    card("Total Revenue", money_(s.total.achieved), s.total.pct + "% of target") +
    card("Gross Profit", money_(s.grossProfit.achieved), s.grossProfit.pct + "% of target") +
    grTile +
    "</div>"
  );
}

function renderStoreTable_(data) {
  var s = data.store;
  var buckets = [["Total Revenue", s.total], ["Service", s.service], ["Accessories", s.accessory],
    ["Repair", s.repair], ["Gross Profit", s.grossProfit]];
  var rows = buckets.map(function (pair) {
    var label = pair[0], b = pair[1];
    return "<tr><td>" + esc_(label) + "</td><td>" + money_(b.target) + "</td><td>" + money_(b.achieved) +
      "</td><td>" + barHtml_(b.pct) + " " + pctBadgeHtml_(b.pct) + "</td><td>" + money_(b.balance) +
      "</td><td>" + money_(b.requiredPerDay) + "</td></tr>";
  }).join("");
  return (
    '<section class="section"><h2>1 &middot; Store Target vs Achievement</h2><div class="body scroll-x">' +
    "<table><thead><tr><th>Revenue Type</th><th>Target</th><th>Achieved</th><th>Achieved %</th>" +
    "<th>Balance</th><th>Req / Day</th></tr></thead><tbody>" + rows + "</tbody></table></div></section>"
  );
}

function renderWorkingDaysCard_(data) {
  var s = data.store;
  return (
    '<section class="section"><h2>2 &middot; Working Days (Current Month)</h2><div class="body">' +
    '<div class="cards" style="margin-bottom:0">' +
    '<div class="card"><div class="label">Total Working Days</div><div class="value">' + s.totalWorkingDays + "</div></div>" +
    '<div class="card"><div class="label">Consumed</div><div class="value">' + s.consumedWorkingDays + "</div></div>" +
    '<div class="card"><div class="label">Balance</div><div class="value" style="color:' +
    (s.balanceWorkingDays > 0 ? "#d97706" : "#16a34a") + '">' + s.balanceWorkingDays + "</div></div>" +
    "</div></div></section>"
  );
}

function renderSalespersonCards_(data) {
  // Incentive-bearing revenue types show target vs achieved, the balance
  // still needed in (brackets), and the incentive that type has earned.
  function incentiveLine_(label, b) {
    var balTxt = b.balance > 0
      ? ' <span class="muted">(' + money_(b.balance) + " to go)</span>"
      : ' <span class="pos">(target met)</span>';
    return (
      '<div class="row-line"><span><b>' + esc_(label) + "</b></span>" +
      "<span>" + money_(b.achieved) + ' <span class="muted">/ ' + money_(b.target) + "</span> " +
      pctBadgeHtml_(b.pct) + "</span></div>" +
      barHtml_(b.pct) +
      '<div class="incentive-line"><span class="muted">Balance to target' + "</span><span>" + balTxt + "</span></div>" +
      '<div class="incentive-line"><span class="muted">Incentive earned</span><span class="' +
      (b.incentive > 0 ? "pos" : "muted") + '">' + money_(b.incentive) + "</span></div>"
    );
  }

  var cards = data.salespeople.map(function (p) {
    var gate = p.repairGate;

    // Repair: no incentive of its own — it only unlocks the others.
    var repairBlock =
      '<div class="row-line" style="margin-top:12px"><span><b>Repair</b> ' +
      '<span class="muted">(no incentive — unlock gate)</span></span>' +
      "<span>" + money_(p.repair.achieved) + ' <span class="muted">/ ' + money_(p.repair.target) +
      "</span> " + pctBadgeHtml_(p.repair.pct) + "</span></div>" +
      barHtml_(p.repair.pct);

    var gateBlock = gate.met
      ? '<div class="gate met">✅ <b>Incentive unlocked.</b> Repair revenue is at ' +
        gate.achievedPct.toFixed(1) + "% of target (needs " + gate.requiredPct + "%). " +
        "Payable incentive: <b>" + money_(p.payableIncentive) + "</b>.</div>"
      : '<div class="gate blocked">🔒 <b>Incentive locked.</b> Repair revenue is at ' +
        gate.achievedPct.toFixed(1) + "% of target — must reach " + gate.requiredPct + "% (" +
        money_(gate.requiredValue) + "). Still needed: <b>" + money_(gate.shortfall) + "</b>.<br>" +
        "Earned but withheld: " + money_(p.earnedIncentive) + ".</div>";

    return (
      '<div class="person-card"><h3>' + esc_(p.name) + "</h3>" +
      '<div class="row-line"><span>Total Revenue</span><span><b>' + money_(p.total.achieved) + "</b>" +
      ' <span class="muted">/ ' + money_(p.total.target) + "</span> " + pctBadgeHtml_(p.total.pct) + "</span></div>" +
      barHtml_(p.total.pct) +
      '<div class="row-line"><span>Invoices</span><span>' + p.invoiceCount + "</span></div>" +
      '<div class="row-line"><span><b>Payable Incentive</b></span><span class="' +
      (p.payableIncentive > 0 ? "pos" : "neg") + '"><b>' + money_(p.payableIncentive) + "</b></span></div>" +
      '<hr style="border:none;border-top:1px solid #e6ebf3;margin:10px 0">' +
      incentiveLine_("Service", p.service) +
      '<div style="height:10px"></div>' +
      incentiveLine_("Accessory", p.accessory) +
      repairBlock +
      gateBlock +
      "</div>"
    );
  }).join("");

  return (
    '<section class="section"><h2>3 &middot; Salesperson Performance</h2><div class="body">' +
    '<div class="person-grid">' + cards + "</div>" +
    '<p class="muted" style="font-size:11.5px;margin:14px 0 0">' +
    "Only <b>Service</b> and <b>Accessory</b> revenue earn incentive. <b>Repair</b> revenue pays nothing on its own — " +
    "it acts as a gate: the salesperson only receives the earned amount once repair revenue reaches <b>" +
    CONFIG.REPAIR_INCENTIVE_GATE_PCT + "%</b> of their repair target.</p>" +
    "</div></section>"
  );
}

function renderIncentiveTable_(data) {
  var bundleByKey = {};
  data.bundle.forEach(function (b) { bundleByKey[b.key] = b; });
  var rows = data.amc.map(function (a) {
    var b = bundleByKey[a.key] || { count: 0, projectedIncentive: 0 };
    var totalInc = a.projectedIncentive + b.projectedIncentive;
    return "<tr><td>" + esc_(a.name) + "</td><td>" + a.amcCount + "</td><td>" + money_(a.amcEligible) +
      "</td><td>" + (a.amcNextSlab ? a.amcRemaining + " more" : "top slab") + "</td><td>" + a.accessoryCount +
      "</td><td>" + money_(a.accessoryEligible) + "</td><td>" + b.count + "</td><td>" +
      money_(b.projectedIncentive) + '</td><td class="pos">' + money_(totalInc) + "</td></tr>";
  }).join("");
  return (
    '<section class="section"><h2>4 &middot; Incentive Valuation — AMC &amp; Bundle</h2><div class="body scroll-x">' +
    "<table><thead><tr><th>Salesperson</th><th>AMC Count</th><th>AMC Eligible</th><th>AMC To Next</th>" +
    "<th>Acc-AMC Count</th><th>Acc-AMC Eligible</th><th>Bundle Count</th><th>Bundle Eligible</th>" +
    "<th>Total Incentive</th></tr></thead><tbody>" + rows + "</tbody></table></div></section>"
  );
}

function renderKpiSection_(data) {
  // Targets come from target!B12:L14 (row 12 = revenue type, row 13 = KPI
  // name, row 14 = target count). Google Reviews is intentionally NOT a KPI
  // row — it is shown as its own tile at the top of the page.
  var allKpis = data.kpi.ar.concat(data.kpi.rr, data.kpi.sr);

  var rows = allKpis.map(function (k) {
    var pctVal = k.target > 0 ? pct_(k.achieved, k.target) : 0;
    return "<tr><td>" + esc_(k.label) + "</td><td>" + esc_(k.category) + "</td><td>" +
      (k.target > 0 ? k.target : "&mdash;") + "</td><td>" + k.achieved + "</td><td>" +
      (k.target > 0 ? pctBadgeHtml_(pctVal) : '<span class="muted">&mdash;</span>') +
      '</td><td class="' + (k.target > 0 && k.balance === 0 ? "pos" : k.balance > 0 ? "neg" : "muted") + '">' +
      (k.target > 0 ? k.balance : "&mdash;") + "</td><td>" +
      money_(k.achievedValue) + "</td><td>" + money_(k.achievedGp) + "</td></tr>";
  }).join("");

  return (
    '<section class="section"><h2>5 &middot; Store KPI Parameters</h2><div class="body scroll-x">' +
    "<table><thead><tr><th>KPI</th><th>Type</th><th>Target</th><th>Achieved</th><th>Achieved %</th>" +
    "<th>Balance</th><th>Achieved Value</th><th>Achieved GP</th></tr></thead><tbody>" + rows +
    "</tbody></table>" +
    '<p class="muted" style="font-size:11.5px;margin:12px 0 0">Targets read from <b>target!B12:L14</b>.</p>' +
    "</div></section>"
  );
}

function renderInvoiceValuation_(data) {
  var rows = data.revenueTypes.map(function (rt) {
    return "<tr><td>" + esc_(rt.label) + "</td><td>" + rt.invoiceLines + "</td><td>" + rt.uniqueInvoices +
      "</td><td>" + money_(rt.revenue) + "</td><td>" + money_(rt.gp) + "</td></tr>";
  }).join("");
  var totalRow = "<tr class=\"total-row\"><td>Overall</td><td>" + data.invoiceValuation.totalInvoiceLines +
    "</td><td>" + data.invoiceValuation.totalUniqueInvoices + "</td><td>" + money_(data.store.total.achieved) +
    "</td><td>" + money_(data.store.grossProfit.achieved) + "</td></tr>";
  return (
    '<section class="section"><h2>6 &middot; Invoice Valuation</h2><div class="body scroll-x">' +
    "<table><thead><tr><th>Revenue Type</th><th>Invoice Lines</th><th>Unique Invoices</th><th>Revenue</th>" +
    "<th>Gross Profit</th></tr></thead><tbody>" + rows + totalRow + "</tbody></table>" +
    '<div class="cards" style="margin-top:14px">' +
    '<div class="card"><div class="label">Bundle Unique Invoices</div><div class="value">' +
    data.invoiceValuation.bundleUniqueInvoices + '</div><div class="sub">SchemeCode starts with DISC</div></div>' +
    '<div class="card"><div class="label">AMC Unique Invoices</div><div class="value">' +
    data.invoiceValuation.amcUniqueInvoices + '</div><div class="sub">SubCategory contains AMC</div></div>' +
    "</div></div></section>"
  );
}

/**
 * Section 7 — a detailed date-wise report broken down by salesperson and
 * revenue type. Each salesperson is given their own colour band / left
 * border so their block is easy to pick out. Achieved values are green when
 * they meet or beat the daily target (store total target ÷ working days),
 * red when they fall short.
 */
function renderDateWiseTable_(data) {
  var dailyTarget = data.store.dailyTarget;

  // Stable colour index per salesperson across the whole table.
  var personIndex = {};
  data.salespeople.forEach(function (p, i) { personIndex[p.key] = i % 5; });
  function idxFor(key) {
    if (personIndex[key] === undefined) personIndex[key] = Object.keys(personIndex).length % 5;
    return personIndex[key];
  }

  function valCell(value, target) {
    var cls = target > 0 ? (value >= target ? "val-good" : "val-bad") : "";
    return '<td class="' + cls + '">' + money_(value) + "</td>";
  }

  var body = "";
  data.days.forEach(function (d) {
    var people = d.people || [];
    // How many detail rows this date spans (for the date rowspan).
    var rowCount = 0;
    people.forEach(function (p) { rowCount += Math.max(1, p.types.length); });
    if (rowCount === 0) rowCount = 1;

    // Date summary band
    var dayCls = d.holiday ? "date-head holiday" : "date-head";
    body +=
      '<tr class="' + dayCls + '"><td colspan="3">' + esc_(d.undated ? "No date" : d.dateKey) +
      (d.weekday ? ' <span class="muted">(' + esc_(d.weekday) + ")</span>" : "") +
      (d.holiday ? " 🔶 holiday" : "") + "</td>" +
      "<td>" + d.uniqueInvoices + " inv</td>" +
      valCell(d.totalRevenue, dailyTarget) +
      "<td>" + money_(d.grossProfit) + "</td>" +
      '<td colspan="2" class="muted">daily target ' + money_(dailyTarget) + "</td></tr>";

    if (people.length === 0) {
      body += '<tr><td colspan="8" class="muted" style="text-align:center">No sales recorded</td></tr>';
      return;
    }

    people.forEach(function (p) {
      var i = idxFor(p.key);
      var types = p.types.length ? p.types : [{ typeLabel: "—", invoiceCount: 0, value: 0, gp: 0 }];
      types.forEach(function (t, ti) {
        body += '<tr class="person-band-' + i + '">';
        // Date column left blank on detail rows (date shown in the band above)
        body += '<td class="muted"></td>';
        if (ti === 0) {
          body += '<td class="person-cell sp-' + i + '" rowspan="' + types.length + '">' + esc_(p.name) + "</td>";
        }
        body += "<td>" + esc_(t.typeLabel) + "</td>";
        body += "<td>" + t.invoiceCount + "</td>";
        body += "<td>" + money_(t.value) + "</td>";
        body += "<td>" + money_(t.gp) + "</td>";
        if (ti === 0) {
          body += '<td rowspan="' + types.length + '"><b>' + money_(p.totalRevenue) + "</b></td>";
          body += '<td rowspan="' + types.length + '">' + money_(p.grossProfit) + "</td>";
        }
        body += "</tr>";
      });
    });
  });

  var legend = data.salespeople.map(function (p) {
    var i = idxFor(p.key);
    return '<span style="display:inline-flex;align-items:center;gap:6px;margin-right:14px;font-size:12px">' +
      '<span class="sp-' + i + '" style="width:14px;height:14px;border-radius:3px;border-left-width:14px;' +
      'border-left-style:solid;display:inline-block"></span>' + esc_(p.name) + "</span>";
  }).join("");

  return (
    '<section class="section"><h2>7 &middot; Date-wise Sales Revenue Analysis (by salesperson)</h2><div class="body">' +
    '<div class="cards" style="margin-bottom:14px">' +
    '<div class="card"><div class="label">Daily Target (store)</div><div class="value">' + money_(dailyTarget) +
    '</div><div class="sub">total target ÷ ' + data.store.totalWorkingDays + " working days</div></div>" +
    '<div class="card"><div class="label">Working Days</div><div class="value">' + data.store.totalWorkingDays +
    '</div><div class="sub">' + data.store.consumedWorkingDays + " consumed / " + data.store.balanceWorkingDays + " left</div></div>" +
    '<div class="card"><div class="label">Required / Day (balance)</div><div class="value">' +
    money_(data.store.total.requiredPerDay) + '</div><div class="sub">to close the remaining gap</div></div>' +
    "</div>" +
    '<div style="margin:0 0 10px">' + legend + "</div>" +
    '<div class="scroll-x"><table class="detail"><thead><tr>' +
    "<th style=\"width:30px\"></th><th>Salesperson</th><th>Revenue Type</th><th>Invoice Count</th>" +
    "<th>Value</th><th>GP</th><th>Person Total</th><th>Person GP</th>" +
    "</tr></thead><tbody>" + body + "</tbody></table></div>" +
    '<p class="muted" style="font-size:11.5px;margin:12px 0 0">' +
    "Date rows show the store total for that day — <span class=\"pos\">green</span> when it meets or beats the daily " +
    "target of " + money_(dailyTarget) + ", <span class=\"neg\">red</span> when it falls short.</p>" +
    "</div></section>"
  );
}
