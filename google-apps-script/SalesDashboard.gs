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
  // ---- Performance page source sheets ----
  // Each entry lists acceptable tab names (first match wins, case-insensitive),
  // so common spellings work without editing anything.
  WALKIN_SHEET_NAMES: ["walkin", "walk-in", "walk in"],
  SERVICE_SHEET_NAMES: ["service", "services"],
  ADVANCE_SHEET_NAMES: ["advance", "advace", "advanced"],
  COMM_SHEET_NAMES: ["comm", "communication", "communications"],
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
  // Background photo for the Web App page, shown at 40% opacity behind the
  // dashboard. Set to "" to use the plain dark background instead.
  BACKGROUND_IMAGE_URL:
    "https://i0.wp.com/mkaion.com/wp-content/uploads/2019/04/business-ideas.jpg?resize=640%2C293&ssl=1",
  // Section 7 starts collapsed showing this many of the most recent dates
  // (current day + the previous two). Expanding reveals every date.
  COLLAPSED_DAYS: 3,
  // Job management lists OPEN jobs only. If the service sheet has a status
  // column, any status containing one of these words is treated as CLOSED and
  // excluded. If there is no status column, a job counts as closed once it has
  // a delivered/closed date, otherwise it stays open.
  JOB_CLOSED_KEYWORDS: ["deliver", "closed", "close", "cancel", "complete", "collected", "returned"],
  // A job is flagged red in the Job Age column once it reaches this many days.
  JOB_AGE_WARN_DAYS: 7,
  JOB_AGE_CRITICAL_DAYS: 15,
  // If this text appears in the GSX or CRM repair status, the job is an
  // unclaimed device and gets a special highlight in the open-jobs table.
  UNCLAIMED_KEYWORD: "unclaim",
  // A job whose latest communication is this many days old or older — or which
  // was never communicated — gets its whole row highlighted as "follow up".
  COMM_STALE_DAYS: 2,
  // How crm!InvoiceDate is written in your sheet: "DMY" = dd-mm-yyyy hh:mm
  // (the default), "MDY" = mm-dd-yyyy hh:mm.
  //
  // This matters because Google Sheets silently converts some of those
  // strings into real dates using the SPREADSHEET's locale. With a US locale
  // "01-09-2026" is read as 9 January, not 1 September — but "14-09-2026"
  // can't be month-first, so it stays text. That mix is what makes a report
  // show a stray "09-01-2026" and only run continuously from the 13th on.
  // The script therefore reads the displayed text and applies this setting,
  // so the interpretation no longer depends on the sheet's locale.
  DATE_ORDER: "DMY",
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

/**
 * Reads the crm sheet, substituting the *displayed* text for the InvoiceDate
 * column only.
 *
 * Why: if the spreadsheet locale is month-first, Sheets silently converts
 * "01-09-2026 12:09" into a Date meaning 9 January, while "14-09-2026" stays
 * text because 14 can't be a month. Reading raw values therefore yields a mix
 * of correct and shifted dates. The displayed string always matches what you
 * see in the cell, so re-parsing it with CONFIG.DATE_ORDER gives one
 * consistent interpretation regardless of locale.
 */
function mergeDisplayDates_(crmSheet) {
  var range = crmSheet.getDataRange();
  var values = range.getValues();
  if (!values.length) return values;

  var dateCol = findColumn_(values[0], [/invoice.*date/i]);
  if (dateCol === -1) return values;

  var display = range.getDisplayValues();
  for (var r = 1; r < values.length; r++) {
    var shown = display[r][dateCol];
    if (shown) values[r][dateCol] = shown;
  }
  return values;
}

/* ============================================================================
 * PERFORMANCE PAGE — shared sheet helpers
 * ========================================================================= */

/** Finds a sheet by any of several acceptable names (case-insensitive). */
function getSheetByNames_(ss, names) {
  var all = ss.getSheets();
  var lookup = {};
  all.forEach(function (sh) { lookup[sh.getName().trim().toLowerCase()] = sh; });
  for (var i = 0; i < names.length; i++) {
    var hit = lookup[String(names[i]).trim().toLowerCase()];
    if (hit) return hit;
  }
  return null;
}

/**
 * Reads a sheet into objects keyed by our internal field names.
 *
 * `spec` maps fieldName -> array of header regexes. Date-ish columns are read
 * from getDisplayValues() so a locale-shifted Date cell is re-parsed from the
 * text actually shown (same defence used for crm!InvoiceDate).
 */
function readSheetRows_(sheet, spec, dateFields) {
  var range = sheet.getDataRange();
  var values = range.getValues();
  if (!values || values.length < 2) return { rows: [], missing: [], header: values[0] || [] };

  var header = values[0];
  var col = {};
  var missing = [];
  Object.keys(spec).forEach(function (field) {
    col[field] = findColumn_(header, spec[field]);
    if (col[field] === -1) missing.push(field);
  });

  // Swap in displayed text for date columns only.
  var display = null;
  (dateFields || []).forEach(function (f) {
    if (col[f] === -1) return;
    if (!display) display = range.getDisplayValues();
    for (var r = 1; r < values.length; r++) {
      var shown = display[r][col[f]];
      if (shown) values[r][col[f]] = shown;
    }
  });

  var out = [];
  for (var r = 1; r < values.length; r++) {
    var src = values[r];
    var blank = true;
    for (var c = 0; c < src.length; c++) { if (str_(src[c])) { blank = false; break; } }
    if (blank) continue;

    var obj = { _row: r + 1 };
    Object.keys(spec).forEach(function (field) {
      obj[field] = col[field] === -1 ? "" : src[col[field]];
    });
    out.push(obj);
  }
  return { rows: out, missing: missing, header: header, col: col };
}

/**
 * Parses any of the date shapes these sheets use — real Date objects,
 * "yyyy-mm-dd hh:mm:ss" (walkin/service/comm) and dd-mm-yyyy hh:mm — into a
 * sortable timestamp plus a normalised dd-MM-yyyy day key.
 */
function parseAnyDate_(raw) {
  var empty = { valid: false, ts: 0, dayKey: "", day: 0, month: 0, year: 0, minutes: 0 };
  if (raw === null || raw === undefined || raw === "") return empty;

  if (raw instanceof Date && !isNaN(raw.getTime())) {
    return {
      valid: true,
      ts: raw.getTime(),
      dayKey: Utilities.formatDate(raw, Session.getScriptTimeZone(), "dd-MM-yyyy"),
      day: raw.getDate(), month: raw.getMonth() + 1, year: raw.getFullYear(),
      minutes: raw.getHours() * 60 + raw.getMinutes(),
    };
  }

  var s = str_(raw);
  if (!s) return empty;

  var parts = s.split(/[\sT]+/);
  var datePart = parts[0];
  var timePart = parts.length > 1 ? parts[1] : "";
  var hh = 0, mi = 0, ss2 = 0;
  var tm = timePart.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?/);
  if (tm) { hh = +tm[1]; mi = +tm[2]; ss2 = tm[3] ? +tm[3] : 0; }

  var dd, mm, yyyy;
  var iso = datePart.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
  if (iso) {
    yyyy = +iso[1]; mm = +iso[2]; dd = +iso[3];
  } else {
    var m = datePart.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{2,4})$/);
    if (!m) {
      var nat = new Date(s);
      if (!isNaN(nat.getTime())) {
        return {
          valid: true, ts: nat.getTime(),
          dayKey: Utilities.formatDate(nat, Session.getScriptTimeZone(), "dd-MM-yyyy"),
          day: nat.getDate(), month: nat.getMonth() + 1, year: nat.getFullYear(),
          minutes: nat.getHours() * 60 + nat.getMinutes(),
        };
      }
      return empty;
    }
    var first = +m[1], second = +m[2];
    yyyy = +m[3];
    if (yyyy < 100) yyyy += 2000;
    if (CONFIG.DATE_ORDER === "MDY") { mm = first; dd = second; } else { dd = first; mm = second; }
    if (mm > 12 && dd <= 12) { var sw = dd; dd = mm; mm = sw; }
  }

  if (!dd || !mm || !yyyy) return empty;
  var js = new Date(yyyy, mm - 1, dd, hh, mi, ss2);
  if (isNaN(js.getTime())) return empty;
  function pad(n) { return (n < 10 ? "0" : "") + n; }
  return {
    valid: true, ts: js.getTime(), dayKey: pad(dd) + "-" + pad(mm) + "-" + yyyy,
    day: dd, month: mm, year: yyyy, minutes: hh * 60 + mi,
  };
}

/** Digits-only mobile key so "+91 98765 43210" and "9876543210" match. */
function mobileKey_(v) {
  var digits = str_(v).replace(/\D/g, "");
  if (digits.length > 10) digits = digits.slice(-10); // drop country code
  return digits;
}

/** Shared pipeline: read crm + target sheets, parse, and compute the dashboard data object. */
function loadDashboardData_() {
  var ss = getSpreadsheet_();
  var crmSheet = ss.getSheetByName(CONFIG.CRM_SHEET_NAME);
  var targetSheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);

  if (!crmSheet) throw new Error('Sheet "' + CONFIG.CRM_SHEET_NAME + '" not found.');
  if (!targetSheet) throw new Error('Sheet "' + CONFIG.TARGET_SHEET_NAME + '" not found.');

  // getDisplayValues() returns what the cell SHOWS, so an InvoiceDate that
  // Google Sheets auto-converted to a real date under a month-first locale is
  // read back as text and re-parsed with CONFIG.DATE_ORDER. Values are merged
  // so only the date column uses display text; everything else keeps its
  // native type (numbers stay numbers).
  var crmValues = mergeDisplayDates_(crmSheet);
  var targetValues = targetSheet.getDataRange().getValues();

  var allRows = parseCrmRows(crmValues);
  var target = parseTargetSheet(targetValues);
  var result = { ss: ss, data: computeDashboard(allRows, target) };

  // Average revenue per unique walk-in — the walkin sheet is optional, so a
  // missing tab simply leaves this metric blank instead of failing the page.
  try {
    var wi = countWalkins_(ss);
    if (wi.unique > 0) {
      result.data.store.uniqueWalkins = wi.unique;
      result.data.store.avgPerWalkin = Math.round(result.data.store.total.achieved / wi.unique);
    }
  } catch (e) { /* walk-in stats are best-effort */ }

  return result;
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
  // "?refresh=1" still works if someone opens the URL directly with it, but
  // the on-page Refresh button no longer navigates — it calls
  // refreshDashboardSections() over google.script.run and swaps the markup
  // in place, which is both faster and safe inside the iframe sandbox.
  var wantsFresh = !!(e && e.parameter && e.parameter.refresh);
  var cache = CacheService.getScriptCache();
  var CACHE_KEY = "dashboard_html_v3";

  var html = null;
  if (!wantsFresh) {
    try { html = cache.get(CACHE_KEY); } catch (err) { html = null; }
  }

  if (!html) {
    var loaded = loadDashboardData_();
    html = buildDashboardPage_(loaded.data);
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
/**
 * Parses crm!InvoiceDate into { day, dayKey, month, year, minutes }.
 *
 * Expected source format: dd-mm-yyyy hh:mm (configurable via
 * CONFIG.DATE_ORDER). dayKey is always normalised to dd-MM-yyyy.
 *
 * Real Date objects are trusted as-is — Apps Script hands back the true
 * instant, so no re-interpretation is needed or wanted there. Strings are
 * split on the configured field order instead of being fed to new Date(),
 * which would otherwise guess month-first and turn 01-09-2026 into 9 Jan.
 */
function parseInvoiceDate_(raw) {
  if (raw instanceof Date && !isNaN(raw.getTime())) {
    return {
      day: raw.getDate(),
      month: raw.getMonth() + 1,
      year: raw.getFullYear(),
      minutes: raw.getHours() * 60 + raw.getMinutes(),
      dayKey: Utilities.formatDate(raw, Session.getScriptTimeZone(), "dd-MM-yyyy"),
    };
  }

  var s = str_(raw);
  if (!s) return { day: 0, month: 0, year: 0, minutes: 0, dayKey: "" };

  // Split "dd-mm-yyyy hh:mm" (or with "/" separators) into date and time.
  var parts = s.split(/[\sT]+/);
  var datePart = parts[0];
  var timePart = parts.length > 1 ? parts[1] : "";

  var minutes = 0;
  var tm = timePart.match(/^(\d{1,2}):(\d{2})/);
  if (tm) minutes = parseInt(tm[1], 10) * 60 + parseInt(tm[2], 10);

  // ISO yyyy-mm-dd is unambiguous, so handle it before the configured order.
  var iso = datePart.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
  if (iso) {
    return mkDateParts_(parseInt(iso[3], 10), parseInt(iso[2], 10), parseInt(iso[1], 10), minutes);
  }

  var m = datePart.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{2,4})$/);
  if (m) {
    var first = parseInt(m[1], 10);
    var second = parseInt(m[2], 10);
    var year = parseInt(m[3], 10);
    if (year < 100) year += 2000;

    var dd, mm;
    if (CONFIG.DATE_ORDER === "MDY") {
      mm = first; dd = second;
    } else {
      dd = first; mm = second;
    }
    // If the "month" field can't be a month, the fields must be the other way
    // round — recover rather than dropping the row.
    if (mm > 12 && dd <= 12) { var swap = dd; dd = mm; mm = swap; }
    return mkDateParts_(dd, mm, year, minutes);
  }

  // Last resort for formats like "9 Sep 2026".
  var d = new Date(s);
  if (!isNaN(d.getTime())) {
    return {
      day: d.getDate(), month: d.getMonth() + 1, year: d.getFullYear(),
      minutes: d.getHours() * 60 + d.getMinutes(),
      dayKey: Utilities.formatDate(d, Session.getScriptTimeZone(), "dd-MM-yyyy"),
    };
  }
  return { day: 0, month: 0, year: 0, minutes: 0, dayKey: datePart };
}

function mkDateParts_(dd, mm, yyyy, minutes) {
  function pad(n) { return (n < 10 ? "0" : "") + n; }
  return {
    day: dd || 0,
    month: mm || 0,
    year: yyyy || 0,
    minutes: minutes || 0,
    dayKey: dd && mm && yyyy ? pad(dd) + "-" + pad(mm) + "-" + yyyy : "",
  };
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
      month: parsedDate.month,
      year: parsedDate.year,
      minutes: parsedDate.minutes,
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
    var approxCell = at(14, kc);        // row 15 — approximate unit value of this KPI
    if (typeCell) lastCategory = typeCell;   // merged cells only fill the first column
    var code = kpiCodeFor_(nameCell);
    if (!code) continue;
    kpis.push({
      category: kpiCategoryFor_(lastCategory, "AR"),
      code: code,
      label: str_(nameCell) || code,
      target: parseMoney_(targetCell),
      value: 0,
      approxAmount: parseMoney_(approxCell),
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

  // "Days elapsed" must come from the newest month present in the data, not
  // the largest day-of-month across all months — otherwise a single row from
  // an earlier month could inflate it.
  var latestStamp = 0;
  rows.forEach(function (r) {
    if (!r.year || !r.month || !r.day) return;
    var stamp = r.year * 10000 + r.month * 100 + r.day;
    if (stamp > latestStamp) latestStamp = stamp;
  });
  var maxDay = latestStamp ? latestStamp % 100 : 0;
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
  // Buckets are keyed by the FULL date (dd-MM-yyyy), not just day-of-month,
  // so the 9th of two different months can never collapse into one row.
  function dayBucketKey_(r) {
    return r.dayKey ? r.dayKey : "";
  }
  function ensureDay(bucketKey) {
    if (!dayMap[bucketKey]) {
      var dateKey = bucketKey;
      var day = 0, month = 0, year = 0;
      var weekday = "";
      if (dateKey && dateKey.indexOf("-") > -1) {
        var parts = dateKey.split("-");
        if (parts.length === 3) {
          day = Number(parts[0]); month = Number(parts[1]); year = Number(parts[2]);
          var jsDate = new Date(year, month - 1, day);
          if (!isNaN(jsDate.getTime())) weekday = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"][jsDate.getDay()];
        }
      }
      var undated = !dateKey;
      dayMap[bucketKey] = {
        day: day, month: month, year: year,
        dateKey: dateKey, weekday: weekday,
        holiday: !undated && !isWorkingDay_(day), undated: undated,
        sortKey: undated ? Infinity : year * 10000 + month * 100 + day,
        uniqueInvoices: 0, serviceRevenue: 0, accessoryRevenue: 0, repairRevenue: 0, totalRevenue: 0,
        grossProfit: 0, byPerson: {},
      };
    }
    return dayMap[bucketKey];
  }

  rows.forEach(function (r) {
    var day = ensureDay(dayBucketKey_(r));
    invoiceSet[r.invoiceCode] = true;
    invoiceDay[r.invoiceCode] = dayBucketKey_(r);
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
    var dd = invoiceDay[code] === undefined ? "" : invoiceDay[code];
    if (!dayUnique[dd]) dayUnique[dd] = {};
    dayUnique[dd][code] = true;
  });
  Object.keys(dayUnique).forEach(function (dd) {
    ensureDay(dd).uniqueInvoices = Object.keys(dayUnique[dd]).length;
  });

  var personInvoiceTotals = {};
  Object.keys(invoiceSet).forEach(function (code) {
    Object.keys(invoicePersonSet[code] || {}).forEach(function (pk) {
      if (!personInvoiceTotals[pk]) personInvoiceTotals[pk] = {};
      personInvoiceTotals[pk][code] = true;
    });
  });

  var days = Object.keys(dayMap).map(function (k) { return dayMap[k]; }).sort(function (a, b) {
    return a.sortKey - b.sortKey;
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

  // Average invoice value = total revenue ÷ unique invoices (i.e. what a
  // typical bill is worth). Uses unique invoice codes rather than line
  // count, so a multi-line bill isn't counted several times.
  store.avgInvoiceValue = Math.round(safeDiv_(sStore.total, Math.max(1, Object.keys(invoiceSet).length)));
  store.avgInvoiceLineValue = Math.round(safeDiv_(sStore.total, Math.max(1, rows.length)));

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

  // ---- AR attachment ----
  // An invoice counts as an AR attachment when it carries accessory revenue
  // ALONGSIDE service or repair revenue. An invoice with only accessory lines
  // is a standalone accessory sale, not an attachment, so it is excluded.
  // Value/GP reported here are the ACCESSORY portion of those invoices —
  // that is the revenue the attachment actually generated.
  var arAttachCodes = {};
  var arAttachValue = 0, arAttachGp = 0, arAttachLines = 0;
  var accessoryOnlyCodes = {};
  var accessoryOnlyValue = 0;

  Object.keys(invoiceAccessoryFlag).forEach(function (code) {
    var acc = invoiceAccessoryValue[code] || { net: 0, gp: 0, count: 0 };
    if (invoiceHasServiceOrRepair[code]) {
      arAttachCodes[code] = true;
      arAttachValue += acc.net;
      arAttachGp += acc.gp;
      arAttachLines += acc.count;
    } else {
      accessoryOnlyCodes[code] = true;
      accessoryOnlyValue += acc.net;
    }
  });

  var totalUnique = Object.keys(invoiceSet).length;
  var arAttachCount = Object.keys(arAttachCodes).length;
  // Invoices that had service/repair and COULD have carried an accessory —
  // the realistic denominator for an attachment rate.
  var attachableCount = Object.keys(invoiceHasServiceOrRepair).length;

  var arAttachment = {
    totalUniqueInvoices: totalUnique,
    attachedInvoices: arAttachCount,
    attachableInvoices: attachableCount,
    missedInvoices: Math.max(0, attachableCount - arAttachCount),
    accessoryOnlyInvoices: Object.keys(accessoryOnlyCodes).length,
    accessoryOnlyValue: Math.round(accessoryOnlyValue),
    attachedLines: arAttachLines,
    value: Math.round(arAttachValue),
    grossProfit: Math.round(arAttachGp),
    // % of all invoices that are AR attachments
    pctOfTotal: pct_(arAttachCount, totalUnique),
    // % of service/repair invoices that also carried an accessory
    attachRate: pct_(arAttachCount, attachableCount),
    avgValue: arAttachCount > 0 ? Math.round(arAttachValue / arAttachCount) : 0,
  };

  var invoiceValuation = {
    totalInvoiceLines: rows.length, totalUniqueInvoices: totalUnique, byType: revenueTypes,
    bundleUniqueInvoices: Object.keys(invoiceSchemeAll).length, amcUniqueInvoices: Object.keys(invoiceAmcAll).length,
    arAttachment: arAttachment,
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
      // Per-day target for this person = their monthly target for that
      // revenue type ÷ the number of working days in the month.
      dailyTargets: {
        service: Math.round(safeDiv_(t.service, Math.max(1, totalWorkingDays))),
        accessory: Math.round(safeDiv_(t.accessory, Math.max(1, totalWorkingDays))),
        repair: Math.round(safeDiv_(t.repair, Math.max(1, totalWorkingDays))),
        total: Math.round(safeDiv_(t.total, Math.max(1, totalWorkingDays))),
      },
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
  var rrDisplay = {}, rrBattery = {}, rrGlass = {};
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
  // "Others" is deliberately excluded from the KPI table — it is a catch-all
  // bucket rather than a tracked target.
  var rr = [mkRow(getKpi("DISPLAY"), rrDisplay), mkRow(getKpi("BATTERY"), rrBattery),
    mkRow(getKpi("GLASS"), rrGlass)];
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
 * PERFORMANCE PAGE — WALK-IN ANALYSIS
 * ========================================================================= */

var WALKIN_SPEC = {
  mobile: [/mobile.*no/i, /^mobile$/i, /phone/i],
  tokenDate: [/token.*date/i, /^date$/i],
  visitPurpose: [/visit.*purpose.*desc/i, /visit.*purpose/i, /purpose/i],
  caseId: [/case.*id/i, /^case$/i],
  invoiceCode: [/invoice.*code/i, /invoice.*no/i, /^invoice$/i],
  remark: [/remark/i, /comment/i, /note/i],
  // Who generated the token. Optional — auto-detected from common spellings.
  createdBy: [/token.*creat.*by/i, /creat.*by/i, /token.*user/i, /user.*id/i, /^user.*name$/i, /staff/i, /employee/i],
};

/**
 * Counts walk-ins from the walkin sheet: total rows and UNIQUE visitors
 * (distinct mobile numbers — the sheet's walk-in identity column, which is
 * located by header name so it works wherever the column sits, e.g. column G).
 */
function countWalkins_(ss) {
  var sheet = getSheetByNames_(ss, CONFIG.WALKIN_SHEET_NAMES);
  if (!sheet) return { rows: 0, unique: 0 };
  var read = readSheetRows_(sheet, { mobile: WALKIN_SPEC.mobile }, []);
  var seen = {};
  read.rows.forEach(function (r) {
    var k = mobileKey_(r.mobile);
    if (k) seen[k] = true;
  });
  return { rows: read.rows.length, unique: Object.keys(seen).length };
}

function computeWalkinAnalysis_(ss) {
  var sheet = getSheetByNames_(ss, CONFIG.WALKIN_SHEET_NAMES);
  if (!sheet) {
    return { available: false, sheetName: CONFIG.WALKIN_SHEET_NAMES[0], days: [], purposes: [],
      creators: [], totals: {}, hasCreator: false };
  }

  var read = readSheetRows_(sheet, WALKIN_SPEC, ["tokenDate"]);
  var rows = read.rows;
  var hasCreator = read.col && read.col.createdBy !== -1;

  var dayMap = {};
  var purposeTotals = {};
  var creatorTotals = {};
  var totals = { walkins: 0, cases: 0, invoices: 0, remarks: 0, deadLeads: 0 };

  rows.forEach(function (r) {
    var d = parseAnyDate_(r.tokenDate);
    var key = d.valid ? d.dayKey : "";
    if (!dayMap[key]) {
      dayMap[key] = {
        dateKey: key, undated: !d.valid,
        sortKey: d.valid ? d.year * 10000 + d.month * 100 + d.day : -1,
        weekday: d.valid ? ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"][new Date(d.year, d.month - 1, d.day).getDay()] : "",
        walkins: 0, cases: 0, invoices: 0, remarks: 0, deadLeads: 0,
        purposes: {}, creators: {},
      };
    }
    var bucket = dayMap[key];

    // Every walk-in row is one visitor; MobileNo is the identity of that visit.
    bucket.walkins += 1;
    totals.walkins += 1;

    var hasCase = !!str_(r.caseId);
    var hasInvoice = !!str_(r.invoiceCode);
    var hasRemark = !!str_(r.remark);

    if (hasCase) { bucket.cases += 1; totals.cases += 1; }
    if (hasInvoice) { bucket.invoices += 1; totals.invoices += 1; }
    if (hasRemark) { bucket.remarks += 1; totals.remarks += 1; }

    // Dead lead: a walk-in with no case, no invoice AND no remark — never converted.
    if (!hasCase && !hasInvoice && !hasRemark) {
      bucket.deadLeads += 1; totals.deadLeads += 1;
    }

    // Visit purpose kept EXACTLY as written in the sheet — the pivot columns.
    var purpose = str_(r.visitPurpose) || "Not specified";
    bucket.purposes[purpose] = (bucket.purposes[purpose] || 0) + 1;
    purposeTotals[purpose] = (purposeTotals[purpose] || 0) + 1;

    if (hasCreator) {
      var by = str_(r.createdBy) || "Unassigned";
      bucket.creators[by] = (bucket.creators[by] || 0) + 1;
      creatorTotals[by] = (creatorTotals[by] || 0) + 1;
    }
  });

  function toSortedList(obj) {
    return Object.keys(obj).map(function (k) { return { name: k, count: obj[k] }; })
      .sort(function (a, b) { return b.count - a.count; });
  }

  var days = Object.keys(dayMap).map(function (k) { return dayMap[k]; })
    .sort(function (a, b) { return b.sortKey - a.sortKey; }); // newest first

  days.forEach(function (d) {
    d.caseRate = pct_(d.cases, d.walkins);
    d.invoiceRate = pct_(d.invoices, d.walkins);
    d.deadRate = pct_(d.deadLeads, d.walkins);
  });

  var creators = toSortedList(creatorTotals).map(function (c) {
    c.share = pct_(c.count, totals.walkins);
    return c;
  });

  // Pivot columns: purposes ordered by volume; beyond 8 they merge into "Others".
  var PURP_COLS_MAX = 8;
  var purposeCols = toSortedList(purposeTotals).map(function (x) { return x.name; });
  var mergedCol = null;
  if (purposeCols.length > PURP_COLS_MAX) {
    mergedCol = "Others";
    var kept = purposeCols.slice(0, PURP_COLS_MAX);
    purposeCols = kept.concat([mergedCol]);
    days.forEach(function (d) {
      var merged = 0;
      Object.keys(d.purposes).forEach(function (name) {
        if (kept.indexOf(name) === -1) {
          merged += d.purposes[name];
          delete d.purposes[name];
        }
      });
      if (merged) d.purposes[mergedCol] = merged;
    });
  }
  // Per-column totals for the footer.
  var purposeColTotals = purposeCols.map(function (name) {
    var sum = 0;
    days.forEach(function (d) { sum += d.purposes[name] || 0; });
    return sum;
  });

  return {
    available: true, sheetName: sheet.getName(), hasCreator: hasCreator,
    days: days,
    purposeCols: purposeCols,
    purposeColTotals: purposeColTotals,
    creators: creators,
    totals: {
      walkins: totals.walkins, cases: totals.cases, invoices: totals.invoices,
      remarks: totals.remarks, deadLeads: totals.deadLeads,
      caseRate: pct_(totals.cases, totals.walkins),
      invoiceRate: pct_(totals.invoices, totals.walkins),
      deadRate: pct_(totals.deadLeads, totals.walkins),
      activeDays: days.filter(function (d) { return !d.undated; }).length,
    },
  };
}

/* ============================================================================
 * PERFORMANCE PAGE — JOB MANAGEMENT (open jobs)
 * ========================================================================= */

var SERVICE_SPEC = {
  caseId: [/case.*id/i, /^case$/i],
  firstName: [/first.*name/i, /customer.*name/i, /^name$/i],
  lastName: [/last.*name/i, /surname/i],
  mobile: [/mobile.*no/i, /^mobile$/i, /phone/i],
  product: [/product.*desc/i, /^product$/i, /item.*desc/i],
  caseDate: [/case.*date/i, /job.*creat.*date/i, /creat.*date/i],
  repairCreatedDate: [/repair.*creat.*date/i, /repair.*date/i],
  jobAge: [/job.*age/i, /^age$/i, /aging/i],
  status: [/job.*status$/i, /case.*status$/i, /^status$/i, /^state$/i],
  // Fallback only: if there is no advance sheet at all, an AvailableBalance
  // column directly on the service sheet is used instead.
  advanceDirect: [/available.*balance/i],
  deliveredDate: [/deliver.*date/i, /closed.*date/i, /collect.*date/i],
  // Two extra repair-status columns: GSX's own status and the CRM repair status.
  gsxStatus: [/gsx.*repair.*status/i, /gsx.*status/i, /^gsx/i],
  crmRepairStatus: [/crm.*job.*status/i, /crm.*repair.*status/i, /^crm_?job/i, /job.*status.*desc/i, /part.*used.*desc/i, /part.*used/i],
};

var ADVANCE_SPEC = {
  mobile: [/mobile.*no/i, /^mobile$/i, /phone/i, /contact/i],
  availableBalance: [/available.*balance/i, /^balance$/i, /advance.*amount/i, /^advance$/i],
  caseId: [/case.*id/i],
};

var COMM_SPEC = {
  caseId: [/case.*id/i, /^case$/i],
  createdDate: [/creat.*date/i, /^date$/i, /comm.*date/i],
  remark: [/remark/i, /message/i, /note/i, /comment/i],
};

function computeJobManagement_(ss) {
  var serviceSheet = getSheetByNames_(ss, CONFIG.SERVICE_SHEET_NAMES);
  if (!serviceSheet) {
    return { available: false, sheetName: CONFIG.SERVICE_SHEET_NAMES[0], jobs: [], totals: {} };
  }

  var read = readSheetRows_(serviceSheet, SERVICE_SPEC,
    ["caseDate", "repairCreatedDate", "deliveredDate"]);
  // The generic job-status regexes would also match CRM_JobStatusDesc — never
  // let the same column drive both fields; the dedicated CRM column wins.
  var hasStatus = read.col && read.col.status !== -1 && read.col.status !== read.col.crmRepairStatus;

  // ---- advance balances, keyed by mobile ----
  var advanceByMobile = {};
  var advanceSheet = getSheetByNames_(ss, CONFIG.ADVANCE_SHEET_NAMES);
  var advanceDirectCol = read.col.advanceDirect === undefined ? -1 : read.col.advanceDirect;
  var advanceByMobileCount = 0;
  if (advanceSheet) {
    var adv = readSheetRows_(advanceSheet, ADVANCE_SPEC, []);
    adv.rows.forEach(function (a) {
      var key = mobileKey_(a.mobile);
      if (!key) return;
      // Same customer can appear more than once — keep the running total.
      advanceByMobile[key] = (advanceByMobile[key] || 0) + num_(a.availableBalance);
    });
  }
  advanceByMobileCount = Object.keys(advanceByMobile).length;

  // ---- latest communication date, keyed by case id ----
  var commByCase = {};
  var commSheet = getSheetByNames_(ss, CONFIG.COMM_SHEET_NAMES);
  if (commSheet) {
    var cm = readSheetRows_(commSheet, COMM_SPEC, ["createdDate"]);
    cm.rows.forEach(function (c) {
      var cid = str_(c.caseId);
      if (!cid) return;
      var d = parseAnyDate_(c.createdDate);
      if (!d.valid) return;
      var prev = commByCase[cid];
      if (!prev || d.ts > prev.ts) {
        commByCase[cid] = { ts: d.ts, dayKey: d.dayKey, remark: str_(c.remark) };
      }
    });
  }

  var today = new Date();
  var todayMid = new Date(today.getFullYear(), today.getMonth(), today.getDate()).getTime();

  var seen = {};
  var jobs = [];
  var closedCount = 0;

  read.rows.forEach(function (r) {
    var caseId = str_(r.caseId);
    if (!caseId) return;
    if (seen[caseId]) return;   // unique case id
    seen[caseId] = true;

    var delivered = parseAnyDate_(r.deliveredDate);
    // Only jobs whose JobStatus is "open" are listed. When there is no status
    // column, anything without a delivered date counts as open.
    var isOpen = hasStatus ? /open/i.test(str_(r.status)) : !delivered.valid;
    if (!isOpen) { closedCount++; return; }       // OPEN jobs only

    var created = parseAnyDate_(r.caseDate);
    var repair = parseAnyDate_(r.repairCreatedDate);

    // Prefer the sheet's own JobAge; fall back to days since the case date.
    var age = num_(r.jobAge);
    if (!age && created.valid) {
      age = Math.max(0, Math.round((todayMid - created.ts) / 86400000));
    }

    var mobile = str_(r.mobile);
    var advance = advanceByMobile[mobileKey_(mobile)] || 0;
    var comm = commByCase[caseId] || null;
    if (!advanceByMobileCount && advanceDirectCol !== -1) {
      advance = num_(r.advanceDirect);        // service!AvailableBalance fallback
    }

    // GSX and CRM repair statuses; "Unclaimed Device" in either triggers the
    // special highlight.
    var gsxStatus = str_(r.gsxStatus);
    var crmRepairStatus = str_(r.crmRepairStatus);
    var isUnclaimed = (gsxStatus + " " + crmRepairStatus).toLowerCase()
      .indexOf(CONFIG.UNCLAIMED_KEYWORD.toLowerCase()) > -1;

    var name = str_(r.firstName);
    if (str_(r.lastName)) name = (name + " " + str_(r.lastName)).trim();

    jobs.push({
      caseId: caseId,
      name: name || "—",
      mobile: mobile || "—",
      product: str_(r.product) || "—",
      status: str_(r.status),
      createdDate: created.valid ? created.dayKey : "—",
      createdTs: created.valid ? created.ts : 0,
      repairDate: repair.valid ? repair.dayKey : "—",
      jobAge: age,
      advance: advance,
      communicated: !!comm,
      commDate: comm ? comm.dayKey : "—",
      commTs: comm ? comm.ts : 0,
      gsxStatus: gsxStatus || "—",
      crmRepairStatus: crmRepairStatus || "—",
      isUnclaimed: isUnclaimed,
      // Communication freshness: days since the latest comm for this case.
      commAgeDays: comm ? Math.floor((Date.now() - comm.ts) / 86400000) : -1,
      // Stale = the latest update is COMM_STALE_DAYS old or more, or no update at all.
      staleComm: comm ? Math.floor((Date.now() - comm.ts) / 86400000) >= CONFIG.COMM_STALE_DAYS : true,
    });
  });

  // Oldest jobs first — they are the ones that need chasing.
  jobs.sort(function (a, b) { return b.jobAge - a.jobAge; });

  var totals = {
    openJobs: jobs.length,
    closedExcluded: closedCount,
    communicated: jobs.filter(function (j) { return j.communicated; }).length,
    notCommunicated: jobs.filter(function (j) { return !j.communicated; }).length,
    advanceTotal: jobs.reduce(function (a, j) { return a + j.advance; }, 0),
    withAdvance: jobs.filter(function (j) { return j.advance > 0; }).length,
    aged: jobs.filter(function (j) { return j.jobAge >= CONFIG.JOB_AGE_CRITICAL_DAYS; }).length,
    unclaimed: jobs.filter(function (j) { return j.isUnclaimed; }).length,
    staleComm: jobs.filter(function (j) { return j.staleComm; }).length,
    avgAge: jobs.length ? Math.round(jobs.reduce(function (a, j) { return a + j.jobAge; }, 0) / jobs.length) : 0,
  };
  totals.commRate = pct_(totals.communicated, totals.openJobs);

  return {
    available: true, sheetName: serviceSheet.getName(),
    hasStatus: hasStatus,
    hasAdvance: !!advanceSheet, hasComm: !!commSheet,
    advanceViaColumn: !advanceSheet && advanceDirectCol !== -1,
    advanceSheetName: advanceSheet ? advanceSheet.getName() : null,
    commSheetName: commSheet ? commSheet.getName() : null,
    jobs: jobs, totals: totals,
  };
}

/** Loads everything the Performance page needs. */
function loadPerformanceData_() {
  var ss = getSpreadsheet_();
  return {
    walkin: computeWalkinAnalysis_(ss),
    jobs: computeJobManagement_(ss),
    generatedAt: new Date(),
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
  // Section 7 is the widest block: Date + Store Total + 3 columns per person.
  var maxCols = Math.max(12, 2 + data.salespeople.length * 3);
  for (var c = 1; c <= maxCols; c++) sheet.autoResizeColumn(c);
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
  var people = data.salespeople;
  var TYPES = [["service", "Service"], ["accessory", "Accessory"], ["repair", "Repair"]];
  var width = 2 + people.length * TYPES.length;

  row = sectionHeader_(sheet, row, "7 · DATE-WISE SALES REVENUE ANALYSIS (by salesperson)", width);

  var dailyTarget = data.store.dailyTarget;
  sheet.getRange(row, 1, 1, 4).setValues([["Store daily target (total ÷ working days)", money_(dailyTarget),
    "Working days", data.store.totalWorkingDays]]).setFontWeight("bold");
  row++;
  // Per-person daily targets, so the /target figures in the grid are explained.
  people.forEach(function (p) {
    sheet.getRange(row, 1, 1, 4).setValues([[p.name + " daily targets",
      "Service " + money_(p.dailyTargets.service),
      "Accessory " + money_(p.dailyTargets.accessory),
      "Repair " + money_(p.dailyTargets.repair)]]);
    sheet.getRange(row, 1).setFontWeight("bold");
    row++;
  });
  row++;

  // ---- two-row banded header: salesperson group over revenue types ----
  var tints = ["#eaf1ff", "#fdefe2", "#e6f6ee", "#f3eafe", "#fce9f2"];
  var hdr1 = ["Date", "Store Total"];
  var hdr2 = ["", ""];
  people.forEach(function (p) {
    TYPES.forEach(function (pair, ti) {
      hdr1.push(ti === 0 ? p.name : "");
      hdr2.push(pair[1]);
    });
  });
  sheet.getRange(row, 1, 1, width).setValues([hdr1])
    .setBackground(COLORS.headerBg).setFontColor(COLORS.headerFg)
    .setFontWeight("bold").setHorizontalAlignment("center");
  people.forEach(function (p, i) {
    var c = 3 + i * TYPES.length;
    sheet.getRange(row, c, 1, TYPES.length).merge()
      .setBackground(COLORS.headerBg).setFontColor(COLORS.headerFg)
      .setFontWeight("bold").setHorizontalAlignment("center").setValue(p.name);
  });
  row++;
  sheet.getRange(row, 1, 1, width).setValues([hdr2])
    .setBackground(COLORS.subHeaderBg).setFontColor(COLORS.subHeaderFg)
    .setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange(row - 1, 1, 2, 1).merge().setValue("Date")
    .setBackground(COLORS.headerBg).setFontColor(COLORS.headerFg)
    .setFontWeight("bold").setVerticalAlignment("middle");
  sheet.getRange(row - 1, 2, 2, 1).merge().setValue("Store Total")
    .setBackground(COLORS.headerBg).setFontColor(COLORS.headerFg)
    .setFontWeight("bold").setVerticalAlignment("middle").setHorizontalAlignment("center");
  row++;

  // ---- one row per date, newest first ----
  var orderedDays = data.days.slice().sort(function (a, b) { return b.sortKey - a.sortKey; });
  orderedDays.forEach(function (d) {
    var byKey = {};
    (d.people || []).forEach(function (pp) { byKey[pp.key] = pp; });

    var vals = [d.undated ? "No date" : (d.dateKey + (d.weekday ? "  " + d.weekday : "") + (d.holiday ? "  🔶" : "")),
      money_(d.totalRevenue)];
    people.forEach(function (p) {
      var pd = byKey[p.key];
      var typeMap = {};
      if (pd) pd.types.forEach(function (t) { typeMap[t.type] = t; });
      TYPES.forEach(function (pair) {
        var tk = pair[0];
        var achieved = typeMap[tk] ? typeMap[tk].value : 0;
        vals.push(money_(achieved) + " / " + money_(p.dailyTargets[tk] || 0));
      });
    });
    sheet.getRange(row, 1, 1, width).setValues([vals]);

    // Date column keeps the store-total rule.
    sheet.getRange(row, 1).setFontWeight("bold");
    sheet.getRange(row, 2).setFontWeight("bold")
      .setFontColor(d.totalRevenue >= dailyTarget ? COLORS.good : COLORS.bad);
    if (d.holiday) sheet.getRange(row, 1, 1, 2).setBackground("#fdf0d5");

    // Each person's cells: tinted group + green/red vs their own daily target.
    people.forEach(function (p, i) {
      var pd = byKey[p.key];
      var typeMap = {};
      if (pd) pd.types.forEach(function (t) { typeMap[t.type] = t; });
      TYPES.forEach(function (pair, ti) {
        var tk = pair[0];
        var achieved = typeMap[tk] ? typeMap[tk].value : 0;
        var tgt = p.dailyTargets[tk] || 0;
        var c = 3 + i * TYPES.length + ti;
        sheet.getRange(row, c).setBackground(tints[i % tints.length]);
        if (tgt > 0) {
          sheet.getRange(row, c).setFontColor(achieved >= tgt ? COLORS.good : COLORS.bad).setFontWeight("bold");
        }
      });
    });
    row++;
  });

  // ---- month-to-date footer ----
  var footer = ["Month to date", money_(data.store.total.achieved)];
  people.forEach(function (p) {
    TYPES.forEach(function (pair) {
      footer.push(money_(p[pair[0]].achieved) + " / " + money_(p[pair[0]].target));
    });
  });
  sheet.getRange(row, 1, 1, width).setValues([footer])
    .setBackground(COLORS.totalBg).setFontWeight("bold");
  row++;

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

/** Builds just the dashboard sections (used for both first paint and refresh). */
function renderSectionsHtml_(data) {
  var body = "";
  body += renderStoreCards_(data);
  body += renderStoreTable_(data);
  body += renderWorkingDaysCard_(data);
  body += renderSalespersonCards_(data);
  body += renderIncentiveTable_(data);
  body += renderKpiSection_(data);
  body += renderInvoiceValuation_(data);
  body += renderDateWiseTable_(data);
  return body;
}

/** The "Updated ... · N rows · N invoices" line under the title. */
function renderMetaText_(data) {
  var generated = Utilities.formatDate(data.generatedAt, Session.getScriptTimeZone(), "dd MMM yyyy, HH:mm");
  return "Updated " + generated + " · " + data.store.totalInvoices +
    " SALES rows · " + data.store.uniqueInvoices + " unique invoices";
}

/**
 * Server endpoint for the Refresh button.
 *
 * Called from the page via google.script.run, which works inside the Apps
 * Script iframe sandbox. It recomputes from the sheet, refreshes the cached
 * page, and hands back only the markup that changed — so refreshing swaps
 * content in place instead of reloading the whole document.
 *
 * NOTE: must stay a public (no trailing underscore) top-level function,
 * otherwise google.script.run cannot see it.
 */
function refreshDashboardSections() {
  var loaded = loadDashboardData_();
  var data = loaded.data;
  var sections = renderSectionsHtml_(data);

  var perfData = loadPerformanceData_();
  var perfHtml = renderPerformanceHtml_(perfData);
  var perfMeta = perfMeta_(perfData);

  // Keep the server-side page cache in step with what we just computed.
  try {
    var full = buildDashboardPage_(data);
    if (full.length < 100000) {
      CacheService.getScriptCache().put("dashboard_html_v3", full, CONFIG.CACHE_SECONDS);
    }
  } catch (err) { /* cache refresh is best-effort */ }

  return JSON.stringify({
    html: sections, meta: renderMetaText_(data),
    perfHtml: perfHtml, perfMeta: perfMeta,
  });
}

/** Assembles the complete HTML document. */
function buildDashboardPage_(data) {
  var css = getDashboardCss_();
  var sections = renderSectionsHtml_(data);

  // ---- second tab: Performance (walk-ins + tokens + open jobs) ----
  var perfData = loadPerformanceData_();
  var perfHtml = renderPerformanceHtml_(perfData);
  var perfMeta = perfMeta_(perfData);

  // Refresh runs through google.script.run because the page is served inside
  // a sandboxed iframe — window.location there points at an internal
  // userCodeAppPanel URL, so navigating it would blank the page.
  var clientScript =
    "<script>" +
    // Section 7 collapse/expand. Rows outside the recent-days window carry
    // .extra-day and are hidden by .collapsed on the table, so toggling is a
    // pure CSS class flip with no server call.
    "function dwToggle(btn){" +
    "  var t=document.getElementById('dwTable'); if(!t) return;" +
    "  var open=t.classList.toggle('collapsed')===false;" +
    "  btn.classList.toggle('open', open);" +
    "  btn.querySelector('.txt').textContent=open?btn.dataset.less:btn.dataset.more;" +
    "}" +
    // Page slider: flip the active tab, show the matching <main>, and swap the
    // subtitle from that container's data-meta attribute.
    "function currentPage(){var a=document.querySelector('.tab-btn.active');return a?a.dataset.page:'business';}" +
    "function switchPage(btn){" +
    "  var tabs=document.querySelectorAll('.tab-btn');" +
    "  for(var i=0;i<tabs.length;i++){tabs[i].classList.toggle('active', tabs[i]===btn);}" +
    "  var page=btn.dataset.page;" +
    "  var biz=document.getElementById('content');" +
    "  var perf=document.getElementById('perfContent');" +
    "  biz.style.display=(page==='business')?'block':'none';" +
    "  perf.style.display=(page==='performance')?'block':'none';" +
    "  var active=(page==='performance')?perf:biz;" +
    "  document.getElementById('meta').textContent=active.getAttribute('data-meta')||'';" +
    "}" +
    "function dashRefresh(btn){" +
    "  if(btn.disabled) return;" +
    "  btn.classList.add('loading'); btn.disabled=true;" +
    "  var t=btn.querySelector('.txt'); var old=t.textContent; t.textContent='Refreshing...';" +
    "  var done=function(){ btn.classList.remove('loading'); btn.disabled=false; t.textContent=old; };" +
    "  if(!(window.google && google.script && google.script.run)){" +
    "    window.top.location.reload(); return;" +
    "  }" +
    "  google.script.run" +
    "    .withSuccessHandler(function(res){" +
    "      try{" +
    "        var prev=document.getElementById('dwTable');" +
    "        var wasOpen=prev?!prev.classList.contains('collapsed'):false;" +
    "        var d=JSON.parse(res);" +
    "        var biz=document.getElementById('content');" +
    "        var perf=document.getElementById('perfContent');" +
    "        biz.innerHTML=d.html;" +
    "        biz.setAttribute('data-meta', d.meta);" +
    "        if(d.perfHtml){ perf.innerHTML=d.perfHtml; perf.setAttribute('data-meta', d.perfMeta); }" +
    "        var active=currentPage()==='performance'?perf:biz;" +
    "        document.getElementById('meta').textContent=active.getAttribute('data-meta')||'';" +
    "        if(wasOpen){" +
    "          var nt=document.getElementById('dwTable');" +
    "          var nb=document.getElementById('dwToggle');" +
    "          if(nt) nt.classList.remove('collapsed');" +
    "          if(nb){ nb.classList.add('open');" +
    "            nb.querySelector('.txt').textContent=nb.dataset.less; }" +
    "        }" +
    "        var n=document.getElementById('note');" +
    "        n.textContent='Updated just now'; n.className='note ok';" +
    "      }catch(e){" +
    "        document.getElementById('note').textContent='Refresh failed: '+e;" +
    "        document.getElementById('note').className='note err';" +
    "      }" +
    "      done();" +
    "    })" +
    "    .withFailureHandler(function(err){" +
    "      var n=document.getElementById('note');" +
    "      n.textContent='Refresh failed: '+(err && err.message ? err.message : err);" +
    "      n.className='note err';" +
    "      done();" +
    "    })" +
    "    .refreshDashboardSections();" +
    "}" +
    "</script>";

  var bizMeta = renderMetaText_(data);

  return (
    "<!DOCTYPE html><html><head><base target=\"_top\"><meta charset=\"utf-8\">" +
    "<title>Sales Performance Dashboard</title><style>" + css + "</style></head><body>" +
    '<header class="topbar"><div class="topbar-inner">' +
    "<div class=\"topbar-left\">" +
    '<div class="brand">📊 Sales Performance Dashboard</div>' +
    // ---- page slider: Business (this page) / Performance (new page) ----
    '<div class="tabs" role="tablist">' +
    '<button class="tab-btn active" data-page="business" onclick="switchPage(this)">💼 Business</button>' +
    '<button class="tab-btn" data-page="performance" onclick="switchPage(this)">📈 Performance</button>' +
    "</div>" +
    '<div class="meta" id="meta">' + esc_(bizMeta) + "</div>" +
    '<div class="note" id="note"></div>' +
    "</div>" +
    '<button class="refresh-btn" onclick="dashRefresh(this)">' +
    '<span class="spin"></span><span class="txt">↻ Refresh data</span></button>' +
    "</div></header>" +
    '<main class="wrap" id="content" data-meta="' + esc_(bizMeta) + '">' + sections + "</main>" +
    '<main class="wrap" id="perfContent" data-meta="' + esc_(perfMeta) + '" style="display:none;">' + perfHtml + "</main>" +
    '<footer class="foot">Business uses the crm / target sheets · Performance uses the walkin / service / advance / comm sheets. Press “Refresh data” to pull the latest values.</footer>' +
    clientScript +
    "</body></html>"
  );
}

function renderDashboardHtml_(data) {
  return buildDashboardPage_(data);
}

/* ============================================================================
 * PERFORMANCE PAGE — HTML RENDERERS
 * ========================================================================= */

function perfMeta_(pd) {
  var bits = [];
  if (pd.walkin.available) bits.push(pd.walkin.totals.walkins + " walk-ins");
  if (pd.jobs.available) bits.push(pd.jobs.totals.openJobs + " open jobs");
  return bits.length
    ? bits.join(" · ") + " · updated " + Utilities.formatDate(pd.generatedAt, Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm")
    : "No performance source sheets found · updated " +
        Utilities.formatDate(pd.generatedAt, Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm");
}

function emptyPerf_(sheetName, msg) {
  return '<div class="empty-perf">📁 No <b>' + esc_(sheetName) + '</b> sheet was found in this spreadsheet.<br>' +
    '<span class="muted">' + esc_(msg) + "</span></div>";
}

function renderWalkinAnalysis_(wl) {
  if (!wl.available) {
    return emptyPerf_(CONFIG.WALKIN_SHEET_NAMES[0],
      "Add a tab named `" + CONFIG.WALKIN_SHEET_NAMES[0] + "` with columns MobileNo, TokenDate, VisitPurposeDesc, CaseId, InvoiceCode and **Remark.");
  }
  var t = wl.totals;
  function card(label, value, sub, cls) {
    return '<div class="card"><div class="label">' + esc_(label) + '</div><div class="value' +
      (cls ? " " + cls : "") + '">' + esc_(value) + '</div><div class="sub">' + esc_(sub || "") + "</div></div>";
  }
  var head =
    '<div class="cards">' +
    card("Total Walk-ins", t.walkins, t.activeDays + " active days") +
    card("Cases Created", t.cases, t.caseRate + "% case conversion") +
    card("Invoices Created", t.invoices, t.invoiceRate + "% invoice conversion") +
    card("Remarks Updated", t.remarks, "rows carrying a remark") +
    card("Dead Leads", t.deadLeads, t.deadRate + "% — no case, no invoice, no remark", t.deadLeads > 0 ? "neg" : "") +
    "</div>";

  function c0(n) { return n ? ('<td class="num">' + n + "</td>") : '<td class="num muted">—</td>'; }

  // Each DATE ROW gets its own soft hue so neighbouring dates are easy to tell
  // apart (pivot-style shading). Hue steps keep consecutive dates distinct.
  var rows = wl.days.map(function (d, i) {
    var hue = (i * 47) % 360;
    var tint = "hsl(" + hue + ",72%,94%)";
    var cells = wl.purposeCols.map(function (name) { return c0(d.purposes[name] || 0); }).join("");
    var dead = d.deadLeads
      ? '<td class="num dead">' + d.deadLeads + "</td>"
      : '<td class="num muted">—</td>';
    var deadPct = d.deadLeads
      ? '<td class="num dead">' + d.deadRate + "%</td>"
      : '<td class="num muted">—</td>';
    var remark = d.remarks
      ? '<td class="num">' + d.remarks + "</td>"
      : '<td class="num muted">—</td>';
    return '<tr style="background:' + tint + '"' + (d.undated ? ' class="undated-row"' : "") + ">" +
      '<td class="daycell"><b>' + esc_(d.undated ? "No date" : d.dateKey) + "</b>" +
        (d.weekday ? ' <span class="muted">' + d.weekday + "</span>" : "") + "</td>" +
      '<td class="num"><b>' + d.walkins + "</b></td>" +
      cells +
      '<td class="num">' + d.cases + "</td>" +
      '<td class="num conv">' + d.caseRate + "%</td>" +
      '<td class="num">' + d.invoices + "</td>" +
      '<td class="num conv">' + d.invoiceRate + "%</td>" +
      remark +
      dead +
      deadPct +
      "</tr>";
  }).join("");

  var purposeFoot = wl.purposeColTotals.map(function (n) {
    return n ? ('<td class="num">' + n + "</td>") : '<td class="num">—</td>';
  }).join("");
  var foot =
    '<tr class="totrow">' +
    "<td><b>Total</b></td>" +
    '<td class="num"><b>' + t.walkins + "</b></td>" +
    purposeFoot +
    '<td class="num">' + t.cases + "</td>" +
    '<td class="num">' + t.caseRate + "%</td>" +
    '<td class="num">' + t.invoices + "</td>" +
    '<td class="num">' + t.invoiceRate + "%</td>" +
    '<td class="num">' + t.remarks + "</td>" +
    '<td class="num dead">' + t.deadLeads + "</td>" +
    '<td class="num dead">' + t.deadRate + "%</td>" +
    "</tr>";

  var purpHeads = wl.purposeCols.map(function (name) { return "<th>" + esc_(name) + "</th>"; }).join("");
  var table =
    '<div class="scroll-x"><table class="ptable walkintable">' +
    "<thead>" +
    '<tr><th rowspan="2">Date</th><th rowspan="2">Walk-ins</th>' +
    '<th colspan="' + wl.purposeCols.length + '" class="grp">Visit purpose</th>' +
    '<th rowspan="2">Cases</th><th rowspan="2">Case conv.</th>' +
    '<th rowspan="2">Invoices</th><th rowspan="2">Invoice conv.</th>' +
    '<th rowspan="2">Remark</th><th rowspan="2">Dead leads</th><th rowspan="2">Dead leads %</th></tr>' +
    "<tr>" + purpHeads + "</tr>" +
    "</thead><tbody>" + rows + foot + "</tbody></table></div>";

  return head + table;
}

function renderTokenCreation_(wl) {
  if (!wl.available) return "";
  if (!wl.hasCreator) {
    return '<div class="subnote">No “created by / token user” column was detected in the ' +
      "<b>" + esc_(wl.sheetName) + "</b> sheet, so token creation by user ID can’t be split yet. " +
      "Add a column such as <b>TokenCreatedBy</b> (or UserId / UserName / Staff) and it will appear automatically here.</div>";
  }
  var t = wl.totals;
  var rows = wl.creators.map(function (c) {
    return "<tr>" +
      "<td><b>" + esc_(c.name) + "</b></td>" +
      '<td class="num">' + c.count + "</td>" +
      '<td class="num">' + c.share + "%</td>" +
      "</tr>";
  }).join("");
  var foot =
    '<tr class="totrow"><td><b>Total tokens generated</b></td>' +
    '<td class="num"><b>' + t.walkins + "</b></td>" +
    '<td class="num"><b>100%</b></td></tr>';
  return (
    '<div class="scroll-x"><table class="ptable tokentable">' +
    "<thead><tr><th>Token created by</th><th>Total tokens created</th><th>% of creation</th></tr></thead>" +
    "<tbody>" + rows + foot + "</tbody></table></div>"
  );
}

function renderJobManagement_(jm) {
  if (!jm.available) {
    return emptyPerf_(CONFIG.SERVICE_SHEET_NAMES[0],
      "Add a tab named `" + CONFIG.SERVICE_SHEET_NAMES[0] + "` with the case / repair / pickup columns listed below. " +
      "If there is a status column, rows not marked delivered/closed are treated as open.");
  }
  var t = jm.totals;
  function card(label, value, sub, cls) {
    return '<div class="card"><div class="label">' + esc_(label) + '</div><div class="value' +
      (cls ? " " + cls : "") + '">' + esc_(value) + '</div><div class="sub">' + esc_(sub || "") + "</div></div>";
  }
  var head =
    '<div class="cards">' +
    card("Open Jobs", t.openJobs, t.closedExcluded + " excluded (status not open)") +
    (t.unclaimed > 0
      ? '<div class="card card-unclaimed"><div class="label">⚠ Unclaimed Devices</div>' +
        '<div class="value neg">' + t.unclaimed + '</div>' +
        '<div class="sub">repair status says “Unclaimed Device” — chase these customers</div></div>'
      : "") +
    card("Communicated", t.communicated, t.commRate + "% of open jobs", "pos") +
    card("Not Communicated", t.notCommunicated, "no comm record found", "neg") +
    card("Needs Follow-up", t.staleComm, "no communication in " + CONFIG.COMM_STALE_DAYS + "+ days — rows highlighted", t.staleComm > 0 ? "neg" : "") +
    card("Avg Job Age", t.avgAge + " d", t.aged + " job(s) " + CONFIG.JOB_AGE_CRITICAL_DAYS + "+ days old", "neg") +
    card("Advance Paid", money_(t.advanceTotal), t.withAdvance + " job(s) carry an advance") +
    "</div>";

  if (!jm.hasAdvance && !jm.advanceViaColumn) {
    head += '<div class="subnote">No advance sheet found (and no service!AvailableBalance column) — the “Advance Paid” column is showing ₹0 for every job.</div>';
  }
  if (!jm.hasComm) {
    head += '<div class="subnote">No comm sheet found — “Communicated” is No for every job.</div>';
  }

  function ageCell(age) {
    var cls = "", label = age + " d";
    if (age >= CONFIG.JOB_AGE_CRITICAL_DAYS) { cls = "age-crit"; label += " ⚠"; }
    else if (age >= CONFIG.JOB_AGE_WARN_DAYS) { cls = "age-warn"; }
    return '<td class="' + cls + ' num">' + label + "</td>";
  }

  // A repair-status cell: the status that reads "Unclaimed Device" gets a
  // strong badge so it stands out from normal statuses.
  function repairStatusCell(txt) {
    var s = str_(txt);
    if (!s || s === "—") return '<td class="muted">—</td>';
    if (s.toLowerCase().indexOf(CONFIG.UNCLAIMED_KEYWORD.toLowerCase()) > -1) {
      return '<td><span class="badge-unclaimed">' + esc_(s) + "</span></td>";
    }
    return "<td>" + esc_(s) + "</td>";
  }

  var rows = jm.jobs.map(function (j) {
    var comm = j.communicated
      ? '<span class="comm-yes">Yes</span>'
      : '<span class="comm-no">No</span>';
    // Comm. updated cell: green when the latest update is within
    // COMM_STALE_DAYS days of today, red when older or not on record at all.
    var commDate;
    if (!j.communicated) {
      commDate = '<td class="comm-neg">Not updated</td>';
    } else if (j.staleComm) {
      commDate = '<td class="comm-neg">' + esc_(j.commDate) + " ⚠</td>";
    } else {
      commDate = '<td class="comm-pos">' + esc_(j.commDate) + " ✓</td>";
    }
    var rowCls = j.isUnclaimed ? "unclaimed-row" : "";
    return "<tr" + (rowCls ? ' class="' + rowCls + '"' : "") + ">" +
      '<td class="sticky-col"><b>' + esc_(j.caseId) + "</b>" +
        (j.isUnclaimed ? ' <span class="badge-unclaimed sm">⚠ Unclaimed</span>' : "") + "</td>" +
      "<td>" + esc_(j.name) + "</td>" +
      "<td>" + esc_(j.mobile) + "</td>" +
      "<td>" + esc_(j.product) + "</td>" +
      repairStatusCell(j.gsxStatus) +
      repairStatusCell(j.crmRepairStatus) +
      "<td>" + esc_(j.createdDate) + "</td>" +
      "<td>" + esc_(j.repairDate) + "</td>" +
      ageCell(j.jobAge) +
      '<td class="num">' + money_(j.advance) + "</td>" +
      "<td>" + comm + "</td>" +
      commDate +
      "</tr>";
  }).join("");

  var table =
    '<div class="scroll-x"><table class="ptable jobtable">' +
    '<thead><tr>' +
    "<th>Case ID</th><th>Customer</th><th>Mobile</th><th>Product</th>" +
    "<th>GSX repair status</th><th>CRM repair status</th>" +
    "<th>Job created</th><th>Repair created</th>" +
    "<th>Job age</th><th>Advance paid</th><th>Communicated</th><th>Comm. updated</th>" +
    "</tr></thead><tbody>" + rows + "</tbody></table></div>";

  return head + table;
}

/** Whole performance tab content (walk-in analysis + token creation + job management). */
function renderPerformanceHtml_(pd) {
  var wl = pd.walkin, jm = pd.jobs;

  var secWalkin =
    '<section class="sec"><div class="sec-head">' +
    '<h2>🚶 Walk-in Analysis</h2>' +
    '<p class="sec-sub">source: ' + esc_(wl.sheetName) + ". Every row is one walk-in — MobileNo identifies the visitor; a row with a CaseId counts as a case, a row with an InvoiceCode as a completed sale.</p>" +
    "</div>" + renderWalkinAnalysis_(wl) + "</section>";

  var secToken =
    '<section class="sec"><div class="sec-head">' +
    '<h2>🎟 Token Creation</h2>' +
    '<p class="sec-sub">how many walk-ins each person’s tokens generated, and their contribution share.</p>' +
    "</div>" + renderTokenCreation_(wl) + "</section>";

  var secJobs =
    '<section class="sec"><div class="sec-head">' +
    '<h2>🔧 Job Management — Open Jobs</h2>' +
    '<p class="sec-sub">source: ' + esc_(jm.sheetName) + '. Only cases whose <b>JobStatus is “open”</b> are listed. Advance is matched on the customer mobile with the advance sheet. Communicated = Yes when a comm record matches the case id — the latest update shows <span style="color:#15803d;font-weight:700">green</span> when within ' + CONFIG.COMM_STALE_DAYS + ' days of today, <span style="color:#b91c1c;font-weight:700">red</span> when older or “Not updated”. “Unclaimed Device” rows are highlighted amber.</p>' +
    "</div>" + renderJobManagement_(jm) + "</section>";

  return secWalkin + secToken + secJobs;
}

function getDashboardCss_() {
  return (
    "* { box-sizing: border-box; }" +
    // ---- mild dark page background, with the background photo at 40% ----
    "body { margin:0; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Arial, sans-serif;" +
    " background:#33405c; color:#1f2937; position:relative; min-height:100vh; }" +
    // The photo sits in a fixed layer behind everything at 40% opacity, so if
    // the image fails to load the mild dark background still shows through.
    (CONFIG.BACKGROUND_IMAGE_URL
      ? "body::before { content:''; position:fixed; inset:0; z-index:-1; opacity:.40;" +
        " background-image:url('" + CONFIG.BACKGROUND_IMAGE_URL + "');" +
        " background-size:cover; background-position:center;" +
        " background-repeat:no-repeat; pointer-events:none; }"
      : "") +
    ".topbar { background:#2b3750; color:#fff; padding:16px 20px; border-bottom:1px solid #46536f;" +
    " position:sticky; top:0; z-index:20; }" +
    ".topbar-inner { max-width:1320px; margin:0 auto; display:flex; align-items:center;" +
    " justify-content:space-between; gap:16px; flex-wrap:wrap; }" +
    ".topbar h1 { margin:0 0 4px; font-size:23px; }" +
    ".topbar .meta { font-size:13.5px; opacity:.85; font-weight:500; }" +
    ".topbar .note { font-size:12.5px; margin-top:3px; min-height:15px; }" +
    ".topbar .note.ok { color:#8ef0b4; }" +
    ".topbar .note.err { color:#ffb4b4; }" +
    ".refresh-btn { background:#4f7cff; color:#fff; border:none; border-radius:8px; padding:10px 18px;" +
    " font-size:14.5px; font-weight:800; cursor:pointer; box-shadow:0 1px 4px rgba(0,0,0,.25);" +
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
    ".section h2 { margin:0; padding:14px 18px; font-size:15.5px; letter-spacing:.06em; text-transform:uppercase;" +
    " background:#43587f; color:#fff; }" +
    ".section .body { padding:16px 18px; }" +
    ".cards { display:grid; grid-template-columns:repeat(auto-fit,minmax(155px,1fr)); gap:12px; margin-bottom:20px; }" +
    ".card { background:#f5f7fb; border-radius:12px; padding:14px 16px; box-shadow:0 2px 8px rgba(0,0,0,.18); }" +
    ".section .card { box-shadow:none; border:1px solid #dde3ee; background:#fbfcfe; }" +
    ".card .label { font-size:12px; text-transform:uppercase; letter-spacing:.08em; color:#55637c; font-weight:800; }" +
    ".card .value { font-size:29px; font-weight:800; margin-top:5px; color:#1f2937;" +
    " font-variant-numeric:tabular-nums; }" +
    ".card .sub { font-size:13.5px; color:#55637c; margin-top:4px; font-weight:500; }" +
    "table { width:100%; border-collapse:collapse; font-size:15.5px; background:#fbfcfe; }" +
    "th { text-align:right; font-size:12.5px; text-transform:uppercase; letter-spacing:.06em; color:#41506b;" +
    " padding:9px 10px; border-bottom:2px solid #c6d0e2; background:#eef2f9; font-weight:800; }" +
    "th:first-child, td:first-child { text-align:left; }" +
    "td { padding:9px 10px; text-align:right; font-size:15.5px; font-weight:500; font-variant-numeric:tabular-nums;" +
    " border-bottom:1px solid rgba(90,110,145,.22); }" +
"tbody tr:last-child td { border-bottom:none; }" +
    "tr:nth-child(even) td { background:#f2f5fa; }" +
    "tr.holiday td { background:#fdf0d5 !important; }" +
    "tr.total-row td { background:#dfe6f2 !important; font-weight:700; }" +
    ".badge { display:inline-block; padding:3px 11px; border-radius:999px; font-weight:800; font-size:14px; }" +
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
    ".person-card h3 { margin:0 0 8px; font-size:17.5px; }" +
    ".row-line { display:flex; justify-content:space-between; align-items:center; gap:8px; margin:7px 0 3px;" +
    " font-size:15.5px; font-weight:600; font-variant-numeric:tabular-nums; }" +
    ".scroll-x { overflow-x:auto; }" +
    ".foot { text-align:center; color:#b9c3d6; font-size:12.5px; padding:6px 0 28px; }" +
    ".pos { color:#15803d; font-weight:700; } .neg { color:#b91c1c; font-weight:700; }" +
    ".muted { color:#6b7688; }" +
    // incentive gate note inside person cards
    ".gate { margin-top:10px; padding:9px 11px; border-radius:8px; font-size:13px; line-height:1.45; }" +
    ".gate.met { background:#e6f7ec; border:1px solid #b6e6c6; color:#14663a; }" +
    ".gate.blocked { background:#fdecec; border:1px solid #f3c3c3; color:#8f2020; }" +
    ".incentive-line { display:flex; justify-content:space-between; font-size:14px; margin-top:6px;" +
    " font-variant-numeric:tabular-nums; padding-bottom:5px;" +
    " border-bottom:1px solid rgba(90,110,145,.16); }" +
    // date-wise per-person grouping
    "table.detail td, table.detail th { font-size:15.5px; }" +
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
    "td.val-good { color:#136c34; font-weight:800; }" +
    "td.val-bad { color:#a81b1b; font-weight:800; }" +
    // ---- section 7 date-by-date matrix ----
    "table.matrix { font-size:15.5px; border-collapse:separate; border-spacing:0; }" +
    "table.matrix th, table.matrix td { white-space:nowrap; padding:8px 10px; }" +
"table.matrix tbody td { border-bottom:1px solid rgba(90,110,145,.28); }" +
    "table.matrix thead th { position:sticky; top:0; z-index:3; }" +
    "table.matrix th.grp { text-align:center; border-left:1px solid #d4dcea; }" +
    "table.matrix .sticky-col { position:sticky; left:0; z-index:2; background:#eef2f9;" +
    " text-align:left; border-right:2px solid #d4dcea; }" +
    "table.matrix tbody tr:nth-child(even) .sticky-col { background:#e8edf6; }" +
    "table.matrix thead .sticky-col { z-index:4; }" +
    "table.matrix td .tgt { display:block; font-size:12.5px; font-weight:600; color:#6b788f; margin-top:2px; }" +
    // tinted column groups, one per salesperson
    "th.grp-0, td.cell-0 { background:#f4f7ff; } tr:nth-child(even) td.cell-0 { background:#eaf1ff; }" +
    "th.grp-1, td.cell-1 { background:#fff7f0; } tr:nth-child(even) td.cell-1 { background:#fdefe2; }" +
    "th.grp-2, td.cell-2 { background:#f1fbf6; } tr:nth-child(even) td.cell-2 { background:#e6f6ee; }" +
    "th.grp-3, td.cell-3 { background:#faf5ff; } tr:nth-child(even) td.cell-3 { background:#f3eafe; }" +
    "th.grp-4, td.cell-4 { background:#fef4f9; } tr:nth-child(even) td.cell-4 { background:#fce9f2; }" +
    "table.matrix tr.holiday .sticky-col { background:#fdf0d5 !important; }" +
    "table.matrix tr.total-row td { background:#dfe6f2 !important; border-top:2px solid #b9c6dd; }" +
    // ---- AR attachment sub-panel inside section 6 ----
    ".subpanel { margin-top:18px; padding:14px 16px; border-radius:10px;" +
    " border:1px solid #d5dded; background:#eef3fb; }" +
    ".subpanel-head { font-size:14px; font-weight:800; text-transform:uppercase; letter-spacing:.06em;" +
    " color:#33456b; margin:0 0 12px; }" +
    ".subpanel .card { background:#ffffff; border:1px solid #dde4f0; }" +
    ".legend-row { display:flex; flex-wrap:wrap; gap:10px; margin:0 0 12px; }" +
    ".legend-card { font-size:13px; padding:8px 12px; border-radius:8px; border:1px solid #dde3ee; line-height:1.5; }" +
    // ---- section 7 collapse / expand ----
    ".toggle-btn { display:inline-flex; align-items:center; gap:9px; margin:0 0 12px;" +
    " background:#43587f; color:#fff; border:none; border-radius:8px; padding:9px 16px;" +
    " font-size:14px; font-weight:800; cursor:pointer; box-shadow:0 1px 3px rgba(0,0,0,.18); }" +
    ".toggle-btn:hover { background:#374a6c; }" +
    ".toggle-btn .chev { display:inline-block; font-size:10px; transition:transform .18s ease; }" +
    ".toggle-btn.open .chev { transform:rotate(180deg); }" +
    ".toggle-btn .count { background:rgba(255,255,255,.22); border-radius:999px;" +
    " padding:2px 10px; font-size:12.5px; font-weight:800; }" +
    ".toggle-btn.open .count { display:none; }" +
    // Collapsed: hide every row outside the recent-days window.
    "table.matrix.collapsed tbody tr.extra-day { display:none; }" +
    // ---- page slider (Business / Performance) ----
    ".topbar-left { display:flex; flex-direction:column; gap:6px; min-width:0; }" +
    ".brand { font-size:17.5px; font-weight:800; letter-spacing:.2px; }" +
    ".tabs { display:inline-flex; background:rgba(255,255,255,.14); border:1px solid rgba(255,255,255,.22);" +
    " border-radius:999px; padding:4px; gap:4px; width:fit-content; }" +
    ".tab-btn { appearance:none; border:none; background:transparent; color:#dfe6f5; font-size:16px;" +
    " font-weight:700; padding:8px 18px; border-radius:999px; cursor:pointer; transition:all .16s ease; }" +
    ".tab-btn:hover { color:#fff; background:rgba(255,255,255,.10); }" +
    ".tab-btn.active { background:#fff; color:#2b3750; box-shadow:0 1px 3px rgba(0,0,0,.35); }" +
    // ---- performance tables ----
    ".sec { background:rgba(255,255,255,.94); border:1px solid #d6deec; border-radius:14px; padding:18px 18px 14px;" +
    " margin-bottom:18px; box-shadow:0 1px 4px rgba(23,32,54,.12); }" +
    ".sec-head h2 { margin:0 0 4px; font-size:20px; color:#22304a; }" +
    ".sec-sub { margin:0 0 12px; font-size:14px; color:#5b6b85; font-weight:500; }" +
    ".scroll-x { overflow-x:auto; }" +
    ".ptable { border-collapse:collapse; width:100%; font-size:15px; }" +
    ".ptable th, .ptable td { text-align:left; padding:9px 12px; border-bottom:1px solid rgba(90,110,145,.22);" +
    " vertical-align:top; white-space:nowrap; font-weight:500; }" +
    ".ptable thead th { background:#eef2f9; color:#33405c; font-size:12.5px; font-weight:800; text-transform:uppercase;" +
    " letter-spacing:.3px; position:sticky; top:0; }" +
    ".ptable tbody tr:hover { background:#f4f7fc; }" +
    ".ptable td.num { text-align:right; font-variant-numeric:tabular-nums; }" +
    ".ptable .daycell { min-width:120px; }" +
    ".jobtable th, .jobtable td { font-size:14.5px; }" +
    ".jobtable .sticky-col { position:sticky; left:0; background:inherit; }" +
    ".chip { display:inline-block; background:#e8eefb; color:#31408a; border-radius:999px;" +
    " padding:3px 9px; font-size:12.5px; margin:1px 2px 1px 0; font-weight:600; }" +
    ".chip b { color:#223; }" +
    // walk-in / comm status colours
    ".comm-yes { color:#15803d; font-weight:700; }" +
    ".comm-no { color:#c2410c; font-weight:700; }" +
    // Unclaimed devices: warm amber row, amber accent bar on the case-id cell,
    // and a solid amber badge — clearly distinct from the green ready rows.
    ".unclaimed-row { background:#fff7e0 !important; }" +
    ".unclaimed-row:hover { background:#fef0c0 !important; }" +
    ".unclaimed-row .sticky-col { box-shadow:inset 4px 0 0 #eab308; }" +
    ".badge-unclaimed { display:inline-block; background:#b45309; color:#fff; font-weight:800;" +
    " font-size:12.5px; padding:3px 10px; border-radius:999px; white-space:nowrap; }" +
    ".badge-unclaimed.sm { font-size:11px; padding:2px 7px; background:#eab308; color:#451a03; }" +
    ".card-unclaimed { background:#fff7e0; border-color:#eab308; }" +
    ".card-unclaimed .value.neg { color:#b45309; }" +
    ".undated-row { background:#fefce8; }" +
    // Comm. updated cell: green tint = updated within the window, red = old / none.
    ".comm-pos { color:#15803d; background:#e6f7ee; font-weight:700; }" +
    ".comm-neg { color:#b91c1c; background:#fdeaea; font-weight:700; }" +
    // Walk-in analysis: grouped purpose header + dead-lead figures.
    ".walkintable thead th.grp { background:#dfe8f6; text-align:center; }" +
    ".walkintable thead th { text-align:center; }" +
    ".walkintable th:first-child, .walkintable td.daycell { text-align:left; }" +
    ".walkintable .totrow td { background:#eef2f9; border-top:2px solid #b9c6dd; }" +
    ".walkintable td.dead { color:#b91c1c; font-weight:700; }" +
    ".walkintable td.conv { color:#31408a; font-weight:600; }" +
    ".tokentable .totrow td { background:#eef2f9; border-top:2px solid #b9c6dd; }" +
    ".age-warn { color:#b45309; font-weight:600; }" +
    ".age-crit { color:#b91c1c; font-weight:700; }" +
    ".subnote { background:#fff8e6; border:1px solid #f0d99a; border-radius:10px; padding:9px 12px;" +
    " font-size:13.5px; color:#7c5c10; margin-bottom:12px; font-weight:500; }" +
    ".empty-perf { background:#fff; border:1px dashed #cbd5e1; border-radius:12px; padding:22px;" +
    " text-align:center; font-size:15.5px; color:#33405c; font-weight:500; }" +
    ".pos { color:#15803d; } .neg { color:#b91c1c; }" +
        "@media (max-width:640px){ .topbar h1{font-size:19px;} .card .value{font-size:21px;} }"
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
    // Avg value tile with TWO details: revenue per unique invoice, and revenue
    // per unique walk-in (from the walkin sheet, when available).
    '<div class="card"><div class="label">Avg Invoice Value</div>' +
    '<div class="value">' + money_(s.avgInvoiceValue) + "</div>" +
    '<div class="sub">revenue ÷ ' + s.uniqueInvoices + " unique invoices</div>" +
    (s.avgPerWalkin
      ? '<div class="value" style="font-size:24px;margin-top:7px">' + money_(s.avgPerWalkin) + "</div>" +
        '<div class="sub">revenue ÷ ' + s.uniqueWalkins + " unique walk-ins</div>"
      : "") +
    "</div>" +
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
      barHtml_(p.repair.pct) +
      // Store-wide revenue target, shown right under the repair target line.
      '<div class="row-line" style="margin-top:6px"><span class="muted">Store Total Revenue Target</span>' +
      '<span class="muted"><b>' + money_(data.store.total.target) + "</b></span></div>";

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
    '<p class="muted" style="font-size:12.5px;margin:14px 0 0">' +
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
    // Approximate achieved value = achieved count x the per-unit approx value
    // listed in target!B15:L15 (shown only when that row has a value).
    var approx = k.approxAmount > 0 ? k.achieved * k.approxAmount : 0;
    return "<tr><td>" + esc_(k.label) + "</td><td>" + esc_(k.category) + "</td><td>" +
      (k.target > 0 ? k.target : "&mdash;") + "</td><td>" + k.achieved + "</td><td>" +
      (k.target > 0 ? pctBadgeHtml_(pctVal) : '<span class="muted">&mdash;</span>') +
      '</td><td class="' + (k.target > 0 && k.balance === 0 ? "pos" : k.balance > 0 ? "neg" : "muted") + '">' +
      (k.target > 0 ? k.balance : "&mdash;") + "</td><td>" +
      (approx > 0 ? money_(approx) : '<span class="muted">&mdash;</span>') +
      "</td><td>" + money_(k.achievedValue) + "</td><td>" + money_(k.achievedGp) + "</td></tr>";
  }).join("");

  return (
    '<section class="section"><h2>5 &middot; Store KPI Parameters</h2><div class="body scroll-x">' +
    "<table><thead><tr><th>KPI</th><th>Type</th><th>Target</th><th>Achieved</th><th>Achieved %</th>" +
    "<th>Balance</th><th>Approx Value</th><th>Achieved Value</th><th>Achieved GP</th></tr></thead><tbody>" + rows +
    "</tbody></table>" +
    '<p class="muted" style="font-size:12.5px;margin:12px 0 0">Targets read from <b>target!B12:L14</b>. ' +
    "<b>Approx Value</b> = achieved count × that KPI’s approximate unit value in <b>target!B15:L15</b>.</p>" +
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
    "</div>" +
    renderArAttachment_(data) +
    "</div></section>"
  );
}

/**
 * AR attachment sub-dashboard (inside section 6).
 *
 * An invoice is an AR attachment when accessory revenue appears on the SAME
 * invoice code as service or repair revenue. Accessory-only invoices are
 * standalone sales and are excluded. The value and GP shown are the
 * accessory portion of those invoices — the revenue the attachment earned.
 */
function renderArAttachment_(data) {
  var a = data.invoiceValuation.arAttachment;

  var rateCls = a.attachRate >= 100 ? "good" : a.attachRate >= 70 ? "info" : a.attachRate >= 40 ? "warn" : "bad";

  return (
    '<div class="subpanel">' +
    '<div class="subpanel-head">AR Attachment' +
    '<span class="muted" style="font-weight:400;text-transform:none;letter-spacing:0"> — accessory sold together with service or repair</span>' +
    "</div>" +
    '<div class="cards" style="margin-bottom:0">' +
    '<div class="card"><div class="label">Total Invoices</div><div class="value">' + a.totalUniqueInvoices +
    '</div><div class="sub">all unique invoice codes</div></div>' +
    '<div class="card"><div class="label">AR Attachment Invoices</div><div class="value pos">' + a.attachedInvoices +
    '</div><div class="sub">' + a.pctOfTotal.toFixed(1) + "% of all invoices</div></div>" +
    '<div class="card"><div class="label">AR Attachment Value</div><div class="value">' + money_(a.value) +
    '</div><div class="sub">accessory portion only</div></div>' +
    '<div class="card"><div class="label">Gross Profit Gained</div><div class="value pos">' + money_(a.grossProfit) +
    '</div><div class="sub">from attached accessories</div></div>' +
    "</div>" +
    '<div class="cards" style="margin:12px 0 0">' +
    '<div class="card"><div class="label">Attach Rate</div><div class="value">' +
    '<span class="badge ' + rateCls + '" style="font-size:18px;padding:3px 12px">' + a.attachRate.toFixed(1) + "%</span>" +
    '</div><div class="sub">' + a.attachedInvoices + " of " + a.attachableInvoices + " service/repair invoices</div></div>" +
    '<div class="card"><div class="label">Missed Opportunities</div><div class="value neg">' + a.missedInvoices +
    '</div><div class="sub">service/repair with no accessory</div></div>' +
    '<div class="card"><div class="label">Avg Attachment Value</div><div class="value">' + money_(a.avgValue) +
    '</div><div class="sub">per attached invoice</div></div>' +
    '<div class="card"><div class="label">Accessory-only Invoices</div><div class="value muted">' +
    a.accessoryOnlyInvoices + '</div><div class="sub">excluded &middot; ' + money_(a.accessoryOnlyValue) + "</div></div>" +
    "</div>" +
    '<p class="muted" style="font-size:12.5px;margin:12px 0 0">' +
    "An invoice counts as an AR attachment only when accessory revenue shares the invoice code with service or repair " +
    "revenue. Invoices carrying accessory lines alone are treated as non-attachment and excluded from the value and GP above.</p>" +
    "</div>"
  );
}

/**
 * Section 7 — a date-by-date matrix.
 *
 * One row per date. Each salesperson owns a group of columns, and within
 * that group every revenue type (Service / Accessory / Repair) gets its own
 * cell showing "achieved / day-target". A cell is green when the achieved
 * value meets or beats that person's daily target for that revenue type,
 * red when it falls short. Each salesperson's column group is tinted so the
 * blocks stay visually distinct.
 *
 * The date column itself keeps the original rule: the STORE total for that
 * day, green when it meets or beats the store daily target, red when not.
 */
function renderDateWiseTable_(data) {
  var dailyTarget = data.store.dailyTarget;
  var people = data.salespeople;
  var TYPES = [["service", "Service"], ["accessory", "Accessory"], ["repair", "Repair"]];

  // Newest date first, so the current day and the two before it sit at the
  // top of the table — exactly what the collapsed view shows.
  var dated = data.days.filter(function (d) { return !d.undated; })
    .slice().sort(function (a, b) { return b.sortKey - a.sortKey; });
  var undated = data.days.filter(function (d) { return d.undated; });

  // Collapsed view keeps the first CONFIG.COLLAPSED_DAYS rows (the most
  // recent dates). Everything else is rendered but hidden, so expanding is
  // instant and needs no extra server round-trip.
  var visibleKeys = {};
  dated.slice(0, CONFIG.COLLAPSED_DAYS).forEach(function (d) { visibleKeys[d.dateKey] = true; });
  var hiddenCount = data.days.length - Math.min(dated.length, CONFIG.COLLAPSED_DAYS);

  // ---- header: two banded rows, one per salesperson group ----
  var head1 = '<tr><th rowspan="2" class="sticky-col">Date</th>' +
    '<th rowspan="2">Store Total</th>';
  var head2 = "<tr>";
  people.forEach(function (p, i) {
    head1 += '<th colspan="' + TYPES.length + '" class="grp grp-' + (i % 5) + '">' + esc_(p.name) + "</th>";
    TYPES.forEach(function (pair) {
      head2 += '<th class="grp grp-' + (i % 5) + '">' + esc_(pair[1]) + "</th>";
    });
  });
  head1 += "</tr>";
  head2 += "</tr>";

  // ---- body: one row per date ----
  function dayRow(d) {
    var byKey = {};
    (d.people || []).forEach(function (pp) { byKey[pp.key] = pp; });

    var dateCls = d.totalRevenue >= dailyTarget ? "val-good" : "val-bad";
    var classes = [];
    if (d.holiday) classes.push("holiday");
    // Rows outside the collapsed window are tagged so the toggle can show them.
    if (!(d.undated ? false : visibleKeys[d.dateKey])) classes.push("extra-day");

    var tr = '<tr class="' + classes.join(" ") + '">';
    tr += '<td class="sticky-col"><b>' + esc_(d.undated ? "No date" : d.dateKey) + "</b>" +
      (d.weekday ? ' <span class="muted">' + esc_(d.weekday) + "</span>" : "") +
      (d.holiday ? ' <span class="muted">🔶</span>' : "") + "</td>";
    tr += '<td class="' + dateCls + '">' + money_(d.totalRevenue) + "</td>";

    people.forEach(function (p, i) {
      var pd = byKey[p.key];
      var typeMap = {};
      if (pd) pd.types.forEach(function (t) { typeMap[t.type] = t; });

      TYPES.forEach(function (pair) {
        var tk = pair[0];
        var achieved = typeMap[tk] ? typeMap[tk].value : 0;
        var tgt = p.dailyTargets[tk] || 0;
        var cls = tgt > 0 ? (achieved >= tgt ? "val-good" : "val-bad") : "muted";
        tr += '<td class="cell-' + (i % 5) + " " + cls + '">' +
          money_(achieved) +
          '<span class="tgt">/ ' + money_(tgt) + "</span></td>";
      });
    });
    tr += "</tr>";
    return tr;
  }

  var body = "";
  dated.forEach(function (d) { body += dayRow(d); });
  undated.forEach(function (d) { body += dayRow(d); });

  // ---- footer: month-to-date totals per person/type (always visible) ----
  var foot = '<tr class="total-row"><td class="sticky-col">Month to date</td>' +
    '<td>' + money_(data.store.total.achieved) + "</td>";
  people.forEach(function (p, i) {
    TYPES.forEach(function (pair) {
      var tk = pair[0];
      foot += '<td class="cell-' + (i % 5) + '">' + money_(p[tk].achieved) +
        '<span class="tgt">/ ' + money_(p[tk].target) + "</span></td>";
    });
  });
  foot += "</tr>";

  // ---- per-person daily target legend ----
  var legend = people.map(function (p, i) {
    return '<div class="legend-card cell-' + (i % 5) + '">' +
      "<b>" + esc_(p.name) + '</b> <span class="muted">daily targets</span><br>' +
      TYPES.map(function (pair) {
        return esc_(pair[1]) + " " + money_(p.dailyTargets[pair[0]]);
      }).join(' <span class="muted">·</span> ') +
      "</div>";
  }).join("");

  var toggle = hiddenCount > 0
    ? '<button class="toggle-btn" id="dwToggle" onclick="dwToggle(this)" ' +
      'data-more="Show all ' + data.days.length + ' dates" ' +
      'data-less="Show only last ' + CONFIG.COLLAPSED_DAYS + ' days">' +
      '<span class="chev">▼</span><span class="txt">Show all ' + data.days.length + ' dates</span>' +
      '<span class="count">+' + hiddenCount + ' hidden</span></button>'
    : "";

  return (
    '<section class="section" id="datewise"><h2>7 &middot; Date-wise Sales Revenue Analysis (by salesperson)</h2><div class="body">' +
    '<div class="cards" style="margin-bottom:14px">' +
    '<div class="card"><div class="label">Store Daily Target</div><div class="value">' + money_(dailyTarget) +
    '</div><div class="sub">total target ÷ ' + data.store.totalWorkingDays + " working days</div></div>" +
    '<div class="card"><div class="label">Working Days</div><div class="value">' + data.store.totalWorkingDays +
    '</div><div class="sub">' + data.store.consumedWorkingDays + " consumed / " + data.store.balanceWorkingDays + " left</div></div>" +
    '<div class="card"><div class="label">Required / Day (balance)</div><div class="value">' +
    money_(data.store.total.requiredPerDay) + '</div><div class="sub">to close the remaining gap</div></div>' +
    "</div>" +
    '<div class="legend-row">' + legend + "</div>" +
    toggle +
    '<div class="scroll-x"><table class="matrix collapsed" id="dwTable"><thead>' + head1 + head2 + "</thead><tbody>" +
    body + foot + "</tbody></table></div>" +
    '<p class="muted" style="font-size:12.5px;margin:12px 0 0">' +
    "Each cell shows <b>achieved / that person's daily target</b> for the revenue type — " +
    '<span class="pos">green</span> when the day meets or beats the target, <span class="neg">red</span> when it falls short. ' +
    "The <b>Store Total</b> column uses the store daily target of " + money_(dailyTarget) + ". " +
    "Collapsed view shows the latest " + CONFIG.COLLAPSED_DAYS + " days; expand for every date.</p>" +
    "</div></section>"
  );
}

/* ============================================================================
 *  ✅  END OF FILE — SalesDashboard.gs
 * ----------------------------------------------------------------------------
 *  PASTE CHECK: this banner must be the LAST thing in your Apps Script editor.
 *  If you cannot see it, your copy was cut short and the script will fail with
 *  "SyntaxError: Unexpected end of input".
 *
 *  Expected length: 2293 lines.
 *  In the Apps Script editor, press Ctrl/Cmd + End to jump to the bottom and
 *  confirm both this banner and the line number.
 *
 *  Reliable way to copy the whole file (a scrolling preview pane often copies
 *  only the part it has rendered):
 *    GitHub → open SalesDashboard.gs → click the "Copy raw file" icon
 *    (top-right of the file view), or "Download raw file".
 * ========================================================================= */
