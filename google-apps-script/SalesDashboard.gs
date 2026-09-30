/**
 * ============================================================================
 *  SALES PERFORMANCE DASHBOARD — Google Apps Script
 * ============================================================================
 *  This script renders the dashboard as a real, standalone web page.
 *  (The old in-sheet "Dashboard" tab writer was removed — the script no
 *  longer creates any tab inside your spreadsheet.)
 *
 *  HOW TO DEPLOY
 *  --------------------------------------------------------------
 *   1. Open your Google Sheet (the one with the "crm" and "target" tabs).
 *   2. Extensions → Apps Script, paste this whole file in, save.
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

  // Store code for the page title — read straight from crm!C2, exactly the
  // cell the previous script used for its location header.
  result.data.store.storeCode =
    (crmValues[1] && crmValues[1].length > 2) ? str_(crmValues[1][2]) : "";

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
  // Pattern-major: the FIRST pattern is the most specific and wins across
  // the whole header row, no matter where that column sits. Each field
  // spec therefore starts with an exact /^CanonicalName$/i pattern.
  for (var p = 0; p < patterns.length; p++) {
    for (var c = 0; c < header.length; c++) {
      if (patterns[p].test(str_(header[c]))) return c;
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

  // Salesperson names live in the header row, columns C..J (target!C1:J1).
  // Empty cells are skipped — the page never invents extra people.
  var people = {};
  var personKeys = [];
  var colIndex = {};
  var headerCells = rows[0] || [];
  for (var c = 2; c < Math.min(headerCells.length, 10); c++) {
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

  // ---- ANCHOR: KPI grid — target!B13:L16 ----
  //   Row 13 (B13:L13) = revenue type / category the KPI belongs to (AR/RR/SR)
  //   Row 14 (B14:L14) = KPI name        (Care+, AMC, Display, SRD, ...)  [fixed]
  //   Row 15 (B15:L15) = target count for that KPI                        [per owner]
  //   Row 16 (B16:L16) = approximate unit value of that KPI
  // 0-indexed: rows 12 / 13 / 14 / 15, columns B=1 .. L=11.
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
    var typeCell = at(12, kc);          // row 13 — revenue type
    var nameCell = at(13, kc);          // row 14 — KPI name
    var targetCell = at(14, kc);        // row 15 — target count
    var approxCell = at(15, kc);        // row 16 — approximate unit value of this KPI
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

  // ---- salesperson roster ----
  // Per owner spec the roster comes EXCLUSIVELY from target!C1:J1 (empty
  // cells are skipped). Names in the CRM that are not on the roster still
  // feed the store/day totals but get no person card or matrix group —
  // this keeps junk values (store codes, C-SAT labels) off the page.
  var peopleKeys = Object.keys(target.salespeople).filter(function (k) { return !!k; }).sort();
  if (peopleKeys.length === 0) {
    // Fallback only when the target roster is completely empty.
    var seen = {};
    rows.forEach(function (r) {
      var pk = normalizePerson_(r.salesPerson);
      if (pk) seen[pk] = true;
    });
    peopleKeys = Object.keys(seen).sort();
  }
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

  // Store-level repair gate gatekeeper: combined store repair ÷ store
  // repair target (0..1). Judged once for the whole store.
  var storeRepairGatePctNow = target.store.repair > 0
    ? sStore.repair / Math.max(1, target.store.repair) : 0;

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
    // Repair (RR) is judged ON THE STORE TARGET, not the salesperson's own
    // ~50% share (owner request): every person's Repair %, balance-to-target
    // and the 80% incentive-unlock gate all compare against the store-wide
    // repair target. Service / Accessory stay on the personal targets.
    var storeRepairTarget = (target.store && target.store.repair) || 0;
    // Gross profit earned by this person (for the date-wise GP column).
    var achGrossProfit = 0;
    mine.forEach(function (r) { if (r.taxableAmount >= 0) achGrossProfit += r.grossProfit; });
    achGrossProfit = Math.round(achGrossProfit);

    // Service & accessory are the only revenue types that EARN incentive.
    var svcInc = incentiveForPct_(target.serviceIncentive.service, pct_(ach.service, t.service));
    var accInc = incentiveForPct_(target.serviceIncentive.accessory, pct_(ach.accessory, t.accessory));
    var earnedIncentive = svcInc + accInc;

    // Repair earns nothing itself — it is a STORE-LEVEL GATE (owner's rule):
    // once the STORE's combined repair revenue reaches
    // CONFIG.REPAIR_INCENTIVE_GATE_PCT of the store repair target, EVERY
    // salesperson becomes eligible and receives the incentive they earned.
    // No per-person repair threshold is applied.
    var storeRepairPct = storeRepairTarget > 0 ? storeRepairGatePctNow : 0;
    var gatePct = CONFIG.REPAIR_INCENTIVE_GATE_PCT;
    var gateNeeded = Math.round((gatePct / 100) * storeRepairTarget);
    var repairGateMet = storeRepairTarget > 0 ? (storeRepairPct * 100) >= gatePct : false;
    var payableIncentive = repairGateMet ? earnedIncentive : 0;

    // Balance still needed to step into the NEXT incentive slab for the two
    // incentive-bearing types. Slab ranges live in target!A18:D25 as % of the
    // personal target, so the money gap is (nextSlab.minPct × target) − won.
    function nextSlab_(slabs, pctVal, tgt, achVal) {
      if (!slabs || !slabs.length || tgt <= 0) return null;
      var sorted = slabs.slice().sort(function (x, y) { return x.minPct - y.minPct; });
      for (var i = 0; i < sorted.length; i++) {
        if (pctVal < sorted[i].minPct) {
          var need = Math.max(0, Math.ceil((sorted[i].minPct / 100) * tgt) - achVal);
          return { label: sorted[i].label, pct: sorted[i].minPct, amount: sorted[i].value, need: need };
        }
      }
      return null;
    }

    function mk(tg, a, inc) {
      return { target: tg, achieved: a, pct: pct_(a, tg), balance: Math.max(0, tg - a), incentive: inc || 0 };
    }

    // Balance still needed per revenue type, used for the balance-based daily targets.
    var balSvc = Math.max(0, t.service - ach.service),
        balAcc = Math.max(0, t.accessory - ach.accessory),
        balRep = Math.max(0, t.repair - ach.repair);

    return {
      key: key, name: titleCase_(key), invoiceCount: mineInvoices,
      total: mk(t.total, ach.total, payableIncentive),
      service: mk(t.service, ach.service, svcInc),
      accessory: mk(t.accessory, ach.accessory, accInc),
      repair: mk(storeRepairTarget, ach.repair, 0),
      // Balance to the NEXT slab (target!A18:D25) so they can chase more incentive.
      nextSvcSlab: nextSlab_(target.serviceIncentive.service, pct_(ach.service, t.service), t.service, ach.service),
      nextAccSlab: nextSlab_(target.serviceIncentive.accessory, pct_(ach.accessory, t.accessory), t.accessory, ach.accessory),
      grossProfit: mk(t.grossProfit, achGrossProfit, 0),
      // Per-day target for this engineer = the BALANCE still needed for that
      // revenue type ÷ the number of working days REMAINING in the month.
      dailyTargets: {
        service: Math.round(safeDiv_(balSvc, Math.max(1, remainingWorkingToday))),
        accessory: Math.round(safeDiv_(balAcc, Math.max(1, remainingWorkingToday))),
        repair: Math.round(safeDiv_(balRep, Math.max(1, remainingWorkingToday))),
        grossProfit: Math.round(safeDiv_(Math.max(0, t.grossProfit - achGrossProfit),
          Math.max(1, remainingWorkingToday))),
        total: Math.round(safeDiv_(balSvc + balAcc + balRep, Math.max(1, remainingWorkingToday))),
      },
      // incentive gate details, surfaced in the UI
      repairGate: {
        requiredPct: gatePct,
        achievedPct: storeRepairPct * 100,
        requiredValue: gateNeeded,
        shortfall: Math.max(0, gateNeeded - sStore.repair),
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
  var allMobiles = {};
  var totals = { walkins: 0, cases: 0, invoices: 0, remarks: 0, deadLeads: 0, unique: 0 };

  rows.forEach(function (r) {
    var d = parseAnyDate_(r.tokenDate);
    var key = d.valid ? d.dayKey : "";
    if (!dayMap[key]) {
      dayMap[key] = {
        dateKey: key, undated: !d.valid,
        sortKey: d.valid ? d.year * 10000 + d.month * 100 + d.day : -1,
        weekday: d.valid ? ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"][new Date(d.year, d.month - 1, d.day).getDay()] : "",
        walkins: 0, unique: 0, cases: 0, invoices: 0, remarks: 0, deadLeads: 0,
        purposes: {}, creators: {}, mobiles: {},
      };
    }
    var bucket = dayMap[key];
    bucket.walkins += 1;
    totals.walkins += 1;

    // Unique walk-ins: distinct mobile numbers seen that day (& overall).
    // Repeated tokens from the same visitor count only once.
    var mk = mobileKey_(r.mobile);
    if (mk) {
      bucket.mobiles = bucket.mobiles || {};
      if (!bucket.mobiles[mk]) { bucket.mobiles[mk] = true; bucket.unique = (bucket.unique || 0) + 1; }
      if (!allMobiles[mk]) { allMobiles[mk] = true; totals.unique = (totals.unique || 0) + 1; }
    }

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
      walkins: totals.walkins, unique: totals.unique || 0, cases: totals.cases, invoices: totals.invoices,
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
  caseId: [/^caseid$/i, /case.*id/i, /^case$/i],
  firstName: [/^firstname$/i, /first.*name/i, /customer.*name/i, /^name$/i],
  lastName: [/^lastname$/i, /last.*name/i, /surname/i],
  mobile: [/^mobileno$/i, /mobile.*no/i, /^mobile$/i, /phone/i],
  product: [/^productdescription$/i, /product.*desc/i, /^product$/i, /item.*desc/i],
  caseDate: [/^casedate$/i, /case.*date/i, /job.*creat.*date/i, /creat.*date/i],
  repairCreatedDate: [/^repaircreateddate$/i, /repair.*creat.*date/i, /repair.*date/i],
  jobAge: [/^jobage$/i, /job.*age/i, /^age$/i, /aging/i],
  status: [/^jobstatus$/i, /job.*status$/i, /case.*status$/i, /^status$/i, /^state$/i],
  // Fallback only: if there is no advance sheet at all, an AvailableBalance
  // column directly on the service sheet is used instead.
  advanceDirect: [/^availablebalance$/i, /available.*balance/i],
  deliveredDate: [/^delivereddate$/i, /deliver.*date/i, /closed.*date/i, /collect.*date/i],
  // Two extra repair-status columns: GSX's own status and the CRM repair status.
  gsxStatus: [/^gsxrepairstatusdescription$/i, /^gsxrepairstatus/i, /gsx.*repair.*status/i, /gsx.*status/i, /^gsx/i],
  crmRepairStatus: [/^crm_jobstatusdesc$/i, /crm.*job.*status/i, /crm.*repair.*status/i, /^crm_?job/i, /job.*status.*desc/i, /part.*used.*desc/i, /part.*used/i],
  // Warranty / coverage classification per case — service!CoverageCodeDescription
  // (GSXWarrantyStatusDesc kept only as a legacy fallback).
  warranty: [/^coveragecodedescription$/i, /coverage.*code.*desc/i, /coverage/i,
    /^gsxwarrantystatusdesc$/i, /gsx.*warranty/i, /warranty.*status/i, /warranty/i],
  // Repair-type classification — service!RepairDocType (also drives the
  // repair analytics on the Performance page).
  repairDocType: [/^repairdoctype$/i, /repair.*doc.*type/i, /repair.*type/i, /doc.*type/i],
  // Ready-for-pickup date — service!Ready_For_Pickup_Date_CRM (average
  // days from CaseDate to pickup-ready on the Performance page).
  readyPickup: [/^ready_for_pickup_date_crm$/i, /ready.*for.*pickup/i, /ready.*pickup/i, /pickup.*date/i],
};

var ADVANCE_SPEC = {
  mobile: [/^mobileno$/i, /mobile.*no/i, /^mobile$/i, /phone/i, /contact/i],
  availableBalance: [/^availablebalance$/i, /available.*balance/i, /^balance$/i, /advance.*amount/i, /^advance$/i],
  caseId: [/^caseid$/i, /case.*id/i],
};

var COMM_SPEC = {
  caseId: [/^caseid$/i, /case.*id/i, /^case$/i],
  createdDate: [/^createddate$/i, /creat.*date/i, /^date$/i, /comm.*date/i],
  remark: [/remark/i, /message/i, /note/i, /comment/i],
};

function computeJobManagement_(ss) {
  var serviceSheet = getSheetByNames_(ss, CONFIG.SERVICE_SHEET_NAMES);
  if (!serviceSheet) {
    return { available: false, sheetName: CONFIG.SERVICE_SHEET_NAMES[0], jobs: [], totals: {} };
  }

  var read = readSheetRows_(serviceSheet, SERVICE_SPEC,
    ["caseDate", "repairCreatedDate", "deliveredDate", "readyPickup"]);
  // The generic job-status regexes would also match CRM_JobStatusDesc — never
  // let the same column drive both fields; the dedicated CRM column wins.
  var hasStatus = read.col && read.col.status !== -1 && read.col.status !== read.col.crmRepairStatus;

  // ---- advance balances, keyed by mobile ----
  var advanceByMobile = {};
  var advanceSheet = getSheetByNames_(ss, CONFIG.ADVANCE_SHEET_NAMES);
  var advanceDirectCol = read.col.advanceDirect === undefined ? -1 : read.col.advanceDirect;
  var advanceByCase = {};
  var advanceByMobileCount = 0;
  if (advanceSheet) {
    var adv = readSheetRows_(advanceSheet, ADVANCE_SPEC, []);
    adv.rows.forEach(function (a) {
      var key = mobileKey_(a.mobile);
      if (key) advanceByMobile[key] = (advanceByMobile[key] || 0) + num_(a.availableBalance);
      var cid = str_(a.caseId);
      // Advance rows can carry a case id — keep a parallel lookup by case.
      if (cid) advanceByCase[cid] = (advanceByCase[cid] || 0) + num_(a.availableBalance);
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

  // ---- UNIQUE case id first -------------------------------------------
  // The sheet can repeat a case on several rows. Before any status filter,
  // collapse to one row per CaseId — the one with the LATEST case date
  // (ties fall to the later row). The jobs table and the repair analytics
  // both work on these unique cases only.
  var byCase = {};
  var order = [];
  read.rows.forEach(function (r) {
    var caseId = str_(r.caseId);
    if (!caseId) return;
    var cd = parseAnyDate_(r.caseDate);
    var score = cd.valid ? cd.ts : 0;
    var prev = byCase[caseId];
    if (!prev || score >= prev.score) byCase[caseId] = { row: r, score: score };
    if (!prev) order.push(caseId);
  });
  var uniqueRows = order.map(function (cid) { return byCase[cid].row; });

  var jobs = [];
  var closedCount = 0;

  uniqueRows.forEach(function (r) {
    var caseId = str_(r.caseId);

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

    // --- conditional rules -------------------------------------------
    var warrantyTxt = str_(r.warranty);
    var repairTypeTxt = str_(r.repairDocType);
    var isSNR = /service\s*non[\s-]*repair/i.test(repairTypeTxt);
    // (a) Service Non-Repair + CRM "ready for pickup" → GSX reads "Completed".
    var gsxCompleted = isSNR && /ready.*for.*pickup/i.test(crmRepairStatus);
    // (b) Non-SNR with NO repair-created date and a case older than 2 days →
    // the whole row is highlighted (nothing has physically started).
    var noRepairLate = !isSNR && !repair.valid && created.valid &&
      (todayMid - created.ts) > 2 * 86400000;
    // (c) Out Of Warranty (No Coverage) with NO advance anywhere for this
    // case or the customer mobile → highlight the line (money at risk).
    var isOOWNoCov = /out\s*of\s*warranty.*no\s*coverage/i.test(warrantyTxt);
    var advanceCaseTotal = advanceByCase[caseId] || 0;
    var anyAdvance = (advance > 0) || (advanceCaseTotal > 0) ||
      (!advanceByMobileCount && advanceDirectCol !== -1 && num_(r.advanceDirect) > 0);
    var oowNoAdvance = isOOWNoCov && !anyAdvance;

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
      warranty: warrantyTxt || "—",
      repairType: repairTypeTxt || "—",
      isSNR: isSNR,
      gsxCompleted: gsxCompleted,
      noRepairLate: noRepairLate,
      oowNoAdvance: oowNoAdvance,
      repairMissing: !repair.valid,
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
  totals.duplicatesMerged = Math.max(0, (function () {
    var n = 0;
    read.rows.forEach(function (r) { if ((str_(r.caseId) || "").trim()) n++; });
    return n - uniqueRows.length;
  })());

  // Unique cases for the repair analytics page (all statuses, not just open).
  var serviceRows = uniqueRows.map(function (r) {
    var cd = parseAnyDate_(r.caseDate);
    var rp = parseAnyDate_(r.readyPickup);
    return {
      caseId: str_(r.caseId),
      caseTs: cd.valid ? cd.ts : 0,
      caseDay: cd.valid ? cd.dayKey : "",
      caseMonth: cd.valid ? cd.dayKey.slice(3) : "",   // "mm-yyyy" from "dd-mm-yyyy"
      readyTs: rp.valid ? rp.ts : 0,
      readyDay: rp.valid ? rp.dayKey : "",
      status: str_(r.status),
      repairType: str_(r.repairDocType) || "",
      warranty: str_(r.warranty) || "",
      product: str_(r.product) || "",
    };
  });

  // Which physical header each displayed column resolved to. Surfaced on
  // the page so a mis-mapped/unspelled column is visible at a glance.
  var sources = {};
  ["caseId", "mobile", "product", "warranty", "repairDocType", "gsxStatus",
   "crmRepairStatus", "caseDate", "repairCreatedDate", "jobAge",
   "advanceDirect", "status", "deliveredDate", "readyPickup"]
    .forEach(function (f) {
      sources[f] = (read.col && read.col[f] !== undefined && read.col[f] !== -1)
        ? str_(read.header[read.col[f]]) : "";
    });

  return {
    available: true, sheetName: serviceSheet.getName(), serviceRows: serviceRows,
    sources: sources,
    hasStatus: hasStatus,
    hasAdvance: !!advanceSheet, hasComm: !!commSheet,
    advanceViaColumn: !advanceSheet && advanceDirectCol !== -1,
    advanceSheetName: advanceSheet ? advanceSheet.getName() : null,
    commSheetName: commSheet ? commSheet.getName() : null,
    jobs: jobs, totals: totals,
  };
}

/**
 * Repair / product / coverage analytics for the PERFORMANCE tab.
 *
 * Sources (per unique CaseId, all statuses):
 *   - repair type:   service!RepairDocType      ("Service Non-Repair" excluded)
 *   - product:       service!ProductDescription → collapsed into families
 *                    (iPhone — incl. model variants & the bare "VIN" export,
 *                    iPad, AirPods, Watch, Mac, Accessory / Non-serialized,
 *                    everything else → Other)
 *   - coverage:      service!CoverageCodeDescription
 *   - date:          service!CaseDate (current month vs the previous 3)
 *   - turnaround:    service!CaseDate → service!Ready_For_Pickup_Date_CRM
 */
function productFamily_(desc) {
  var t = (desc || "").toLowerCase();
  if (/iphone/i.test(t) || /\bvin\b/i.test(t)) return "iPhone";
  if (/ipad/i.test(t)) return "iPad";
  if (/airpod/i.test(t)) return "AirPods";
  if (/watch/i.test(t)) return "Watch";
  if (/mac(book)?/i.test(t)) return "Mac";
  if (/accessor|non[\s-]*serial/.test(t)) return "Accessory (Non-serialized)";
  return "Other";
}

function isSNRType_(t) { return /service\s*non[\s-]*repair/i.test(t || ""); }

function computeRepairAnalytics_(serviceRows) {
  var all = (serviceRows || []).filter(function (r) { return r.caseTs > 0; });
  if (!all.length) return { available: false };
  // Same sheet cleanliness rules as Job Management: keep the latest CaseDate
  // row per CaseId so duplicate exports never double-count.
  var bestByCase = {}, order = [];
  all.forEach(function (r, idx) {
    var k = r.caseId ? r.caseId : ("_row" + idx);
    if (!bestByCase[k]) { bestByCase[k] = r; order.push(k); }
    else if (r.caseTs > bestByCase[k].caseTs) { bestByCase[k] = r; }
  });
  all = order.map(function (k) { return bestByCase[k]; });
  // Service Non-Repair cases are excluded from every analytic on this page.
  var rows = all.filter(function (r) { return !isSNRType_(r.repairType); });

  var now = new Date();
  function pad2(x) { return (x < 10 ? "0" : "") + x; }
  function mkeyOf(y, m) { return y + "-" + pad2(m + 1); }
  function monthAbbr(m) { return ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][m]; }

  var curKey = mkeyOf(now.getFullYear(), now.getMonth());
  var curLabel = monthAbbr(now.getMonth()) + " " + now.getFullYear();
  var prev3 = [];
  for (var k = 3; k >= 1; k--) {
    var dd = new Date(now.getFullYear(), now.getMonth() - k, 1);
    prev3.push({ key: mkeyOf(dd.getFullYear(), dd.getMonth()),
      label: monthAbbr(dd.getMonth()) + " " + dd.getFullYear() });
  }
  var prev3Keys = {};
  prev3.forEach(function (o) { prev3Keys[o.key] = true; });

  // Round integers only on this page, as requested.
  function avgInt(arr) {
    if (!arr.length) return { avg: null, n: 0 };
    var sum = 0; arr.forEach(function (x) { sum += x; });
    return { avg: Math.round(sum / arr.length), n: arr.length };
  }

  // One comparison block per dimension (repair type / product / coverage).
  function dimensionAnalysis(keyFn) {
    var stat = {};
    rows.forEach(function (r) {
      var t = keyFn(r);
      var st = stat[t] || (stat[t] = { cur: 0, byMonth: {}, curDurs: [], prevDurs: [] });
      var mkNorm = r.caseMonth ? r.caseMonth.slice(3) + "-" + r.caseMonth.slice(0, 2) : "";
      var dur = (r.readyTs && r.readyTs >= r.caseTs)
        ? Math.round((r.readyTs - r.caseTs) / 86400000) : -1;
      if (mkNorm === curKey) {
        st.cur++;
        if (dur >= 0) st.curDurs.push(dur);
      } else if (prev3Keys[mkNorm]) {
        st.byMonth[mkNorm] = (st.byMonth[mkNorm] || 0) + 1;
        if (dur >= 0) st.prevDurs.push(dur);
      }
    });
    var names = Object.keys(stat).sort();
    return names.map(function (t) {
      var st = stat[t];
      var prevTotal = 0;
      prev3.forEach(function (o) { prevTotal += st.byMonth[o.key] || 0; });
      var perMonth = Math.round(prevTotal / 3);
      var deltaPct = perMonth > 0 ? Math.round(((st.cur - perMonth) / perMonth) * 100) : null;
      var c = avgInt(st.curDurs), pv = avgInt(st.prevDurs);
      return { name: t, cur: st.cur, perMonth: perMonth, deltaPct: deltaPct,
        curDaysAvg: c.avg, curDaysN: c.n, prevDaysAvg: pv.avg, prevDaysN: pv.n };
    });
  }

  var typeComp = dimensionAnalysis(function (r) {
    return (r.repairType || "").trim() || "Unspecified";
  });
  var productComp = dimensionAnalysis(function (r) { return productFamily_(r.product); });
  var coverComp = dimensionAnalysis(function (r) {
    return (r.warranty || "").trim() || "Unspecified";
  });

  // Month-to-date totals per repair type (no date rows, as requested).
  var mtdTotal = 0;
  var mtdRows = typeComp.filter(function (c) { return c.cur > 0; })
    .sort(function (a, b) { return b.cur - a.cur; });
  mtdRows.forEach(function (c) { mtdTotal += c.cur; });

  // ---- Repair Validation matrix: product families (vertical) × coverage (horizontal) ----
  // One combined table per owner request: every current-month repair counts
  // exactly once, placed at the intersection of its product family and its
  // coverage code. Row/column totals keep the composition checkable.
  var matrix = {}, prodTotals = {}, covTotals = {};
  rows.forEach(function (r) {
    var mkNorm = r.caseMonth ? r.caseMonth.slice(3) + "-" + r.caseMonth.slice(0, 2) : "";
    if (mkNorm !== curKey) return;
    var pf = productFamily_(r.product);
    var cov = (r.warranty || "").trim() || "Unspecified";
    if (!matrix[pf]) matrix[pf] = {};
    matrix[pf][cov] = (matrix[pf][cov] || 0) + 1;
    prodTotals[pf] = (prodTotals[pf] || 0) + 1;
    covTotals[cov] = (covTotals[cov] || 0) + 1;
  });
  function avgFor(comp, name) {
    var hit = null;
    comp.forEach(function (c) { if (c.name === name) hit = c; });
    return hit ? hit.perMonth : 0;
  }
  var valProducts = Object.keys(matrix).sort(function (a, b) {
    return prodTotals[b] - prodTotals[a] || (a < b ? -1 : 1);
  });
  var valCoverages = Object.keys(covTotals).sort(function (a, b) {
    return covTotals[b] - covTotals[a] || (a < b ? -1 : 1);
  });
  var validationMatrix = {
    products: valProducts, coverages: valCoverages, cells: matrix,
    prodTotals: prodTotals, covTotals: covTotals,
    prodAvg: function (name) { return avgFor(productComp, name); },
    covAvg: function (name) { return avgFor(coverComp, name); },
    grand: valProducts.reduce(function (a, pf) { return a + prodTotals[pf]; }, 0),
  };

  return {
    available: true,
    curLabel: curLabel, prev3Labels: prev3.map(function (o) { return o.label; }).join(", "),
    typeComp: typeComp, productComp: productComp, coverComp: coverComp,
    validationMatrix: validationMatrix,
    mtdRows: mtdRows, mtdTotal: mtdTotal,
    typeCount: typeComp.length,
    totalRows: all.length, exclRows: rows.length,
    monthRepairsTotal: typeComp.reduce(function (a, c) { return a + c.cur; }, 0),
    prevRepairAvgTotal: typeComp.reduce(function (a, c) { return a + c.perMonth; }, 0),
  };
}

/** Renders the repair-type & coverage analytics sections (Operation page). */
function renderRepairAnalyticsInner_(ra) {
  if (!ra.available) {
    return '<div class="section"><h2 class="section-title">🔧 Repair Analytics</h2>' +
      '<div class="empty-state">No service data available yet.</div></div>';
  }
  var out = [];

  function deltaPill(c) {
    if (c.deltaPct == null) return '<span class="pill pill-benchmark" title="No repairs logged in the previous 3 months">no 3-month baseline</span>';
    var cls = c.deltaPct > 0 ? "pill-positive" : (c.deltaPct < 0 ? "pill-negative" : "pill-benchmark");
    var sign = c.deltaPct > 0 ? "+" : "";
    return '<span class="pill ' + cls + '" title="This month (to date) vs the 3-month monthly average">' +
      sign + c.deltaPct + "% vs 3-mo avg</span>";
  }
  function daysCell(avg) { return avg == null ? "—" : avg + " d"; }

  function bucketSection(icon, title, labelCol, comp, note) {
    if (!comp.length) return "";
    var o = ['<div class="section">' +
      '<h2 class="section-title">' + icon + " " + esc_(title) + " · <em>" + esc_(ra.curLabel) + "</em></h2>" +
      '<div class="section-sub">' + esc_(note) + ' Counts are whole numbers. "Service Non-Repair" cases are excluded.</div>' +
      '<div class="tile-grid">'];
    comp.forEach(function (c) {
      o.push('<div class="tile"><div class="label">' + esc_(c.name) + "</div>" +
        '<div class="value">' + c.cur + "</div>" +
        '<div class="sub">avg ' + c.perMonth + "/month · prev 3 months</div>" +
        deltaPill(c) + "</div>");
    });
    o.push('</div><div class="table-wrap"><table aria-label="' + esc_(title) + '"><thead><tr>' +
      "<th>" + esc_(labelCol) + "</th>" +
      '<th class="num">This Month</th><th class="num">Avg / Month · Prev 3M</th><th class="num">Change</th>' +
      '<th class="num">Avg Days · Now</th><th class="num">Avg Days · Prev 3M</th>' +
      "</tr></thead><tbody>");
    comp.forEach(function (c) {
      o.push("<tr><td><strong>" + esc_(c.name) + "</strong></td>" +
        '<td class="num">' + c.cur + "</td>" +
        '<td class="num">' + c.perMonth + "</td>" +
        '<td class="num">' + deltaPill(c) + "</td>" +
        '<td class="num">' + daysCell(c.curDaysAvg) + "</td>" +
        '<td class="num">' + daysCell(c.prevDaysAvg) + "</td></tr>");
    });
    o.push("</tbody></table></div></div>");
    return o.join("");
  }

  var typeSec = bucketSection("📈", "Repair Type Performance", "Repair Type", ra.typeComp,
    "Cases logged in " + ra.curLabel + " vs the monthly average of the previous 3 months (" + ra.prev3Labels + ").");

  out.push(typeSec);

  // The "Repair Types — Till Date" composition table was retired on the
  // owner's request; the Repair Type Performance block above already carries
  // the month-to-date counts per type.
  return out.join("");
}

/** Kept for backward-compatible references: the same sections, wrapped. */
function renderRepairAnalytics_(ra) {
  return '<div id="performance" class="tab-page active">' +
    renderRepairAnalyticsInner_(ra) + "</div>";
}

function repairMeta_(ra, generatedAt) {
  var m = ra && ra.available ? ra.monthRepairsTotal : 0;
  var types = ra && ra.typeComp ? ra.typeComp.filter(function (c) { return c.cur > 0; }).length : 0;
  return m + " repairs this month · " + types + " repair types · updated " +
    Utilities.formatDate(generatedAt, Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm");
}

/**
 * Process-page C-SAT block — read from the target sheet, presented cleanly:
 *   target!K1     = the period name shown as the dashboard title
 *   target!L1:U5  = the C-SAT grid (10 cols × 5 rows; row 1 = headers,
 *                   rows 2-5 = one row per salesperson plus a TOTAL row)
 * The sheet's raw grid is NEVER dumped verbatim: junk/empty columns are
 * dropped, the TOTAL row is confirmed against the person rows, the summary
 * numbers (responses, C-SAT, CFI) become tiles, and the table gets
 * first-column styling + a bold total row.
 */
function computeTargetBlock_(ss) {
  var sheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);
  if (!sheet) return { available: false, periodName: "", rows: [], totals: {} };
  var rows = sheet.getDataRange().getDisplayValues();
  function at(r, c) { return (rows[r] && rows[r][c] !== undefined && rows[r][c] !== null) ? String(rows[r][c]).trim() : ""; }

  var K = 10, L = 11;                    // K1 / first col of the table (0-indexed)
  var FIRST_ROW = 0, LAST_ROW = 4;       // sheet rows 1..5
  var periodName = at(FIRST_ROW, K);

  var header = [];
  for (var c = 0; c < 10; c++) header.push(at(FIRST_ROW, L + c));

  // Identify the special columns by header text.
  function isJunkHeader(h) { return h === "" || h === "-" || /^–+$/.test(h); }
  function findCol(re) {
    for (var i = 0; i < header.length; i++) if (re.test(header[i])) return i;
    return -1;
  }
  var totalCol = findCol(/^total$/i);
  var csatCol = findCol(/c[-\s]?sat/i);
  var cfiCol = findCol(/^cfi$/i);

  // Satisfaction categories = every column that is not name / TOTAL / C-SAT / CFI / junk.
  var catCols = [];
  for (var i = 0; i < header.length; i++) {
    if (i === 0 || i === totalCol || i === csatCol || i === cfiCol) continue;
    if (isJunkHeader(header[i])) continue;
    catCols.push(i);
  }
  // Tolerant bucket kind: sheets carry misspellings ("DISSATISIFIED"), so
  // match loose stems. "Neither…" is neutral even though it contains both
  // words; dissatisfaction stems are negative; satisfaction = positive.
  function catKind(h) {
    var t = (h || "").toLowerCase().replace(/[^a-z]/g, "");
    if (t.indexOf("neither") > -1 || t.indexOf("neutral") > -1) return "mid";
    if (/dissatis|disatis|satisn?or/.test(t) && t.indexOf("dis") > -1) return "neg";
    if (t.indexOf("satis") > -1 || t.indexOf("happy") > -1) return "pos";
    return "mid";
  }

  var people = [];
  var totalRow = null;
  for (var r = FIRST_ROW + 1; r <= LAST_ROW && r < (rows.length || 0); r++) {
    var name = at(r, L);
    if (!name) continue;
    var entry = { name: name, cats: [], total: 0, csat: at(r, L + csatCol), cfi: at(r, L + cfiCol) };
    catCols.forEach(function (ci) {
      var n = parseMoney_(at(r, L + ci));
      entry.cats.push(n);
      entry.total += n;
    });
    if (totalCol >= 0) {
      var declared = parseMoney_(at(r, L + totalCol));
      if (declared > 0) entry.total = declared;
    }
    if (/^total$/i.test(name)) totalRow = entry; else people.push(entry);
  }

  if (!people.length && !totalRow && !periodName) {
    return { available: false, periodName: periodName, rows: [], totals: {} };
  }
  // Trust the sheet TOTAL when present; otherwise derive it from the people rows.
  if (!totalRow) {
    totalRow = { name: "TOTAL", cats: catCols.map(function () { return 0; }), total: 0, csat: "", cfi: "" };
    people.forEach(function (p) {
      p.cats.forEach(function (n, i) { totalRow.cats[i] += n; });
    });
    totalRow.total = people.reduce(function (a, p) { return a + p.total; }, 0);
  }
  var labels = catCols.map(function (ci) { return header[ci]; });
  var kinds = catCols.map(function (ci) { return catKind(header[ci]); });

  // How satisfied overall: every "satisfied" (not dissatisfied) bucket ÷ total.
  function satisfiedCount(e) {
    var n = 0;
    e.cats.forEach(function (v, i) { if (kinds[i] === "pos") n += v; });
    return n;
  }
  function dissatisfiedCount(e) {
    var n = 0;
    e.cats.forEach(function (v, i) { if (kinds[i] === "neg") n += v; });
    return n;
  }

  return {
    available: true,
    periodName: periodName,
    nameHeader: header[0] || "Salesperson",
    labels: labels, kinds: kinds,
    people: people, totalRow: totalRow,
    satTotal: satisfiedCount(totalRow), disTotal: dissatisfiedCount(totalRow),
    csat: totalRow.csat, cfi: totalRow.cfi,
    satisfiedCount: satisfiedCount, dissatisfiedCount: dissatisfiedCount,
  };
}

/** Renders the C-SAT block: summary tiles + styled table + repair strip. */
function renderTargetBlock_(tb, ra) {
  if (!tb.available) return "";
  var title = "Customer Satisfaction — " + (tb.periodName || "current period");

  function pctOf(part, whole) { return whole > 0 ? Math.round((part / whole) * 100) : 0; }
  function chip(cls, label, value, sub) {
    return '<div class="csat-chip ' + cls + '"><span class="cc-label">' + esc_(label) + '</span>' +
      '<span class="cc-value">' + esc_(value) + '</span><span class="cc-sub">' + esc_(sub || "") + '</span></div>';
  }

  var chips = [];
  chips.push(chip("", "Feedback responses", tb.totalRow.total,
    tb.people.length + " salesperson" + (tb.people.length === 1 ? "" : "s") + " rated"));
  chips.push(chip("pos", "Satisfied (VS + SS)", pctOf(tb.satTotal, tb.totalRow.total) + "%",
    tb.satTotal + " of " + tb.totalRow.total));
  if (tb.csat) chips.push(chip("info", "C-SAT", tb.csat, "as recorded in sheet"));
  if (tb.cfi) chips.push(chip(tb.cfi === "0.00%" || tb.cfi === "0" ? "pos" : "warn", "CFI", tb.cfi, "as recorded in sheet"));
  chips.push(chip(tb.disTotal > 0 ? "neg" : "pos", "Dissatisfied", pctOf(tb.disTotal, tb.totalRow.total) + "%",
    tb.disTotal + " of " + tb.totalRow.total));

  var thead = "<thead><tr><th>" + esc_(tb.nameHeader) + "</th>";
  tb.labels.forEach(function (h) { thead += "<th>" + esc_(h) + "</th>"; });
  thead += "<th>Total</th><th>% Satisfied</th>" + (tb.csat ? "<th>C-SAT</th>" : "") + (tb.cfi ? "<th>CFI</th>" : "") + "</tr></thead>";

  function dataRow(e, isTotal) {
    var tr = "<tr" + (isTotal ? ' class="totrow"' : "") + "><td><b>" + esc_(e.name) + "</b></td>";
    e.cats.forEach(function (n, i) {
      var cls = tb.kinds[i] === "pos" ? "csat-pos" : tb.kinds[i] === "neg" ? "csat-neg" : "csat-mid";
      tr += '<td class="num ' + cls + '">' + (n || "—") + "</td>";
    });
    tr += '<td class="num"><b>' + e.total + "</b></td>";
    tr += '<td class="num">' + pctOf(tb.satisfiedCount(e), e.total) + "%</td>";
    if (tb.csat) tr += '<td class="num">' + esc_(isTotal ? e.csat : e.csat || "—") + "</td>";
    if (tb.cfi) tr += '<td class="num">' + esc_(isTotal ? e.cfi : e.cfi || "—") + "</td>";
    return tr + "</tr>";
  }
  var body = "<tbody>" + tb.people.map(function (p) { return dataRow(p, false); }).join("") +
    dataRow(tb.totalRow, true) + "</tbody>";

  // Owner layout: table shrunk and left-aligned; the summary chips stack in a
  // column immediately to the table's right.
  var table = '<div class="csat-flex"><div class="scroll-x csat-left"><table class="ptable csat-table">' + thead + body +
    '</table></div><div class="csat-right">' + chips.join("") + "</div></div>";

  // Repairs comparison strip — Service Non-Repair excluded (kept from before).
  var strip = "";
  if (ra && ra.available) {
    var cur = ra.monthRepairsTotal, avg = ra.prevRepairAvgTotal;
    var pill;
    if (avg > 0) {
      var delta = Math.round(((cur - avg) / avg) * 100);
      pill = '<span class="pill ' + (delta >= 0 ? "pill-positive" : "pill-negative") + '">' +
        (delta >= 0 ? "+" : "") + delta + "% vs 3-mo avg</span>";
    } else {
      pill = '<span class="pill pill-benchmark">no 3-month baseline</span>';
    }
    strip = '<div class="repair-strip">' +
      '<div class="rs-item"><span class="rs-label">Total Repairs (excl. Service Non-Repair) — this month</span>' +
        '<span class="rs-value">' + cur + "</span></div>" +
      '<div class="rs-item"><span class="rs-label">3-month monthly average</span>' +
        '<span class="rs-value">' + avg + "</span></div>" +
      '<div class="rs-item">' + pill + "</div></div>";
  }

  return '<section class="sec"><div class="sec-head">' +
    "<h2>🎯 " + esc_(title) + "</h2>" +
    '<p class="sec-sub">Every feedback response lands in a satisfaction bucket. Period name from ' +
    '<span class="src-hint">target!K1</span>, table from <span class="src-hint">target!L1:U5</span> — ' +
    "cleaned and re-presented (junk columns dropped, totals verified).</p></div>" +
    table + strip + "</section>";
}

/** Loads everything the Performance & Process pages need. */
function loadPerformanceData_() {
  var ss = getSpreadsheet_();
  var jm = computeJobManagement_(ss);
  var ra = computeRepairAnalytics_(jm.serviceRows || []);

  // Freshness: the newest dated row of each sheet, surfaced on the index page.
  function latestDateKey_(names, spec, field) {
    var sheet = getSheetByNames_(ss, names);
    if (!sheet) return null;
    var rd = readSheetRows_(sheet, spec, [field]);
    var ts = 0, key = "";
    rd.rows.forEach(function (r) {
      var d = parseAnyDate_(r[field]);
      if (d.valid && d.ts > ts) { ts = d.ts; key = d.dayKey; }
    });
    return key || "no rows";
  }
  var freshness = {
    comm: latestDateKey_(CONFIG.COMM_SHEET_NAMES, { createdDate: COMM_SPEC.createdDate }, "createdDate") || "no comm sheet",
    advance: "no date column",
  };

  return {
    walkin: computeWalkinAnalysis_(ss),
    jobs: jm,
    repair: ra,
    targetBlock: computeTargetBlock_(ss),
    freshness: freshness,
    generatedAt: new Date(),
  };
}



function money_(n) { return "₹" + Math.round(n || 0).toLocaleString("en-IN"); }























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
  // Full Business content (kept for compatibility); pages render only their
  // own slice via renderPageHtml_ below.
  return renderGpHtml_(data) + renderKpiHtml_(data) + renderDateHtml_(data);
}

/** GP page: Store target · Weekly Revenue Analysis · Invoice valuation. */
function renderGpHtml_(data) {
  var body = "";
  body += renderStoreCards_(data);
  body += renderStoreTable_(data);
  body += renderWorkingDaysCard_(data);
  body += renderInvoiceValuation_(data);
  body += renderWeeklyAnalysis_(data);
  return body;
}

/** KPI page: KPI target · Incentive · Salesperson performance · AR attachment. */
function renderKpiHtml_(data) {
  // KPI table sizes to its CONTENT, and the incentive tables compress into
  // the free space on its right instead of leaving a wide empty band.
  var body = '<div class="kpi-pair"><div class="kp-left">' + renderKpiSection_(data) + "</div>" +
    '<div class="kp-right">' + renderIncentiveTable_(data) + "</div></div>";
  body += renderSalespersonCards_(data);
  body += renderArSection_(data);
  return body;
}

/** DATE page: date-wise sales revenue analysis (by salesperson). */
function renderDateHtml_(data) {
  return renderDateWiseTable_(data);
}

/** WALK-IN page: walk-in analysis + token creation side by side. */
function renderWalkinHtml_(pd) {
  var wl = pd.walkin;
  var head =
    '<section class="section" id="sec-walkin"><h2>🚶 Walk-in Analysis &amp; Token Creation</h2><div class="body">' +
    '<div class="sec-desc">source: ' + esc_(wl.sheetName) +
    ". Every row is one walk-in — MobileNo identifies the visitor (Unique = distinct mobiles); a row with a CaseId counts as a case, a row with an InvoiceCode as a completed sale. The token table pins which user created each token.</div>" +
    '<div class="walkin-pair">';

  // Matching heads keep 🚶 Walk-in Analysis and 🎟 Token Creation aligned on
  // ONE line; the wide day-wise table gets the full page width below the pair.
  var left =
    '<div class="wp-left"><div class="wp-head">🚶 Walk-in Analysis</div>' +
    '<p class="sec-desc" style="margin:-2px 0 10px">month-to-date walk-in summary — visitors, cases, invoices, remarks and dead leads.</p>' +
    renderWalkinAnalysis_(wl) + "</div>";
  var right =
    '<div class="wp-right"><div class="wp-head">🎟 Token Creation</div>' +
    '<p class="sec-desc" style="margin:-2px 0 10px">how many walk-in tokens each user created, and their share of the month&rsquo;s tokens.</p>' +
    renderTokenCreation_(wl) + "</div>";

  return head + left + right + "</div>" +
    '<div class="wp-daytable"><div class="wp-head">📅 Day-wise Walk-in &amp; Token Details</div>' +
    renderWalkinDayTable_(wl) + "</div></div></section>";
}

/** OPERATION page: customer satisfaction + repair ops + future sub-modules. */
function renderOperationHtml_(pd) {
  var out = [];
  out.push(renderTargetBlock_(pd.targetBlock, pd.repair));
  out.push(renderRepairAnalyticsInner_(pd.repair));
  // Placeholder sub-modules the owner will fill in later.
  out.push(
    '<section class="sec"><div class="sec-head"><h2>🚫 Decline Repair</h2>' +
    '<p class="sec-sub">Coming soon — jobs the customer declined after diagnosis.</p></div>' +
    '<div class="ph-tile"><div class="ph-icon">⏳</div>' +
    '<div class="ph-text"><b>Placeholder for the Decline-repair module.</b><br>' +
    "This tile is reserved in the new Operation page; the content will be built next.</div></div></section>");
  out.push(
    '<section class="sec"><div class="sec-head"><h2>🔁 Requote &amp; Win-back</h2>' +
    '<p class="sec-sub">Coming soon — re-quoted jobs and customers won back.</p></div>' +
    '<div class="ph-tile"><div class="ph-icon">⏳</div>' +
    '<div class="ph-text"><b>Placeholder for the Requote / win-back module.</b><br>' +
    "This tile is reserved in the new Operation page; the content will be built next.</div></div></section>");
  return out.join("");
}

/** DECK page (Task 24): last tab, not on the index. Summary cards only —
 *  store target vs achievement, KPI achievement, walk-ins, C-SAT, open jobs. */
function renderDeckHtml_(data, pd) {
  var s = data.store;
  var out = [];

  // ---- 1) Store target vs achievement ----
  var storeGroups = [
    ["Service", s.service], ["Accessories", s.accessory], ["Repair", s.repair],
    ["Total", s.total], ["Gross Profit", s.grossProfit],
  ];
  var storeRows = storeGroups.map(function (g) {
    var b = g[1] || { target: 0, achieved: 0 };
    var pct = b.target > 0 ? (100 * b.achieved / b.target) : 0;
    var cls = pct >= 100 ? "pos" : pct >= 60 ? "" : "neg";
    return "<tr><td><b>" + esc_(g[0]) + "</b></td>" +
      '<td class="num">' + money_(b.target) + "</td>" +
      '<td class="num">' + money_(b.achieved) + "</td>" +
      '<td class="num ' + cls + '"><b>' + pct.toFixed(1) + "%</b></td>" +
      '<td class="num">' + money_(Math.max(0, b.target - b.achieved)) + "</td></tr>";
  }).join("");
  out.push(
    '<section class="section" id="deck-store"><h2>🎯 Store Target vs Achievement</h2><div class="body">' +
    '<div class="scroll-x"><table class="ptable decktable">' +
    "<thead><tr><th>Type</th><th class=\"num\">Target</th><th class=\"num\">Achievement</th><th class=\"num\">Achieved %</th><th class=\"num\">Balance</th></tr></thead><tbody>" +
    storeRows + "</tbody></table></div></div></section>");

  // ---- 2) KPI achievement ----
  var kpis = data.kpi.ar.concat(data.kpi.rr, data.kpi.sr);
  var kT = 0, kA = 0, kHit = 0;
  kpis.forEach(function (k) {
    kT += k.target; kA += k.achieved;
    if (k.target > 0 && k.achieved >= k.target) kHit++;
  });
  var kPct = kT > 0 ? (100 * kA / kT) : 0;
  function dc(klabel, kvalue, ksub) {
    return '<div class="card"><div class="label">' + esc_(klabel) + '</div><div class="value">' +
      esc_(String(kvalue)) + '</div><div class="sub">' + esc_(ksub || "") + "</div></div>";
  }
  out.push(
    '<section class="section" id="deck-kpi"><h2>📊 KPI Achievement</h2><div class="body">' +
    '<div class="cards">' +
    dc("KPIs Hit", kHit + " / " + kpis.length, "count at or above target") +
    dc("Overall Target", kT, "sum of all KPI targets") +
    dc("Overall Achieved", kA, "sum of all KPI achievements") +
    dc("Achievement %", kPct.toFixed(1) + "%", "achieved ÷ target") +
    "</div></div></section>");

  // ---- 3) Walk-in details ----
  if (pd.walkin && pd.walkin.available) {
    var t = pd.walkin.totals;
    out.push(
      '<section class="section" id="deck-walkin"><h2>🚶 Walk-in Details</h2><div class="body">' +
      '<div class="cards">' +
      dc("Total Walk-ins", t.walkins, t.activeDays + " active days") +
      dc("Unique Walk-ins", t.unique, "distinct mobile numbers") +
      dc("Cases Created", t.cases, "walk-ins → service case") +
      dc("Invoices Created", t.invoices, "walk-ins → sale") +
      dc("Dead Leads", t.deadLeads, "no case · no invoice · no remark") +
      "</div></div></section>");
  } else {
    out.push('<section class="section" id="deck-walkin"><h2>🚶 Walk-in Details</h2><div class="body">' +
      '<div class="subnote">No walkin sheet found.</div></div></section>');
  }

  // ---- 4) Customer satisfaction ----
  if (pd.targetBlock && pd.targetBlock.available) {
    var tb = pd.targetBlock;
    var satPct = tb.totalRow.total > 0 ? Math.round((100 * tb.satTotal) / tb.totalRow.total) : 0;
    out.push(
      '<section class="section" id="deck-csat"><h2>🎯 Customer Satisfaction — ' + esc_(tb.periodName || "current period") + '</h2><div class="body">' +
      '<div class="cards">' +
      dc("Total Responses", tb.totalRow.total, "people surveyed") +
      dc("Satisfied", tb.satTotal, satPct + "% of responses") +
      dc("Dissatisfied", tb.disTotal, (100 - satPct) + "% of responses") +
      (tb.csat !== "" ? dc("C-SAT", esc_(tb.csat), "from the sheet TOTAL row") : "") +
      (tb.cfi !== "" ? dc("CFI", esc_(tb.cfi), "from the sheet TOTAL row") : "") +
      "</div></div></section>");
  } else {
    out.push('<section class="section" id="deck-csat"><h2>⭐ Customer Satisfaction</h2><div class="body">' +
      '<div class="subnote">C-SAT table not found on the target sheet.</div></div></section>');
  }

  // ---- 5) Open jobs count ----
  if (pd.jobs && pd.jobs.available) {
    var j = pd.jobs.totals;
    out.push(
      '<section class="section" id="deck-jobs"><h2>🔧 Open Jobs</h2><div class="body">' +
      '<div class="cards">' +
      dc("Open Jobs", j.openJobs, j.closedExcluded + " excluded (delivered/closed)") +
      (j.unclaimed > 0 ? dc("⚠ Unclaimed", j.unclaimed, "chase these customers") : "") +
      dc("Communicated", j.communicated, j.commRate + "% of open") +
      dc("Not Communicated", j.notCommunicated, "no comm record yet") +
      "</div></div></section>");
  } else {
    out.push('<section class="section" id="deck-jobs"><h2>🔧 Open Jobs</h2><div class="body">' +
      '<div class="subnote">No service sheet found.</div></div></section>');
  }

  return '<section class="sec"><div class="sec-head"><h2>🗂 Deck — Month at a Glance</h2>' +
    '<p class="sec-sub">Store target vs achievement · KPI achievement · walk-ins · C-SAT · open jobs at a glance.</p></div></section>' +
    out.join("");
}

/** OPEN JOBS page: job management. */
function renderJobsHtml_(pd) {
  return renderJobsSection_(pd);
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

  var perfData = loadPerformanceData_();

  // Keep the server-side page cache in step with what we just computed.
  try {
    var full = buildDashboardPage_(data);
    if (full.length < 100000) {
      CacheService.getScriptCache().put("dashboard_html_v3", full, CONFIG.CACHE_SECONDS);
    }
  } catch (err) { /* cache refresh is best-effort */ }

  return JSON.stringify({
    meta: renderMetaText_(data),
    index: renderIndexHtml_(data, perfData),
    gp: renderGpHtml_(data),
    kpi: renderKpiHtml_(data),
    walkin: renderWalkinHtml_(perfData),
    operation: renderOperationHtml_(perfData),
    jobs: renderJobsHtml_(perfData),
    date: renderDateHtml_(data),
    deck: renderDeckHtml_(data, perfData),
  });
}

/* ============================================================================
 * SHEET-SEARCH CHAT BOT — floating assistant on every page
 *
 * serverQuery(q) answers natural-language questions from the SOURCE sheets:
 *   • "sale date from 15.09.26"          → invoices that day on, revenue +
 *                                           salesperson-wise gross profit
 *   • "part number for iphone 3 display" → matching item codes / part rows
 *   • "9791994639"                       → mobile matched across walk-in,
 *                                           service, advance & CRM sheets
 * ========================================================================= */

/**
 * Public endpoint for the chat widget. No trailing underscore — the Apps
 * Script editor would treat `serverQuery_` as private and google.script.run
 * would not reach it.
 */
function serverQuery(q) {
  var result = { title: "Sheet search", rows: [], note: "" };
  try {
    var query = str_(q).trim();
    if (!query) {
      result.title = "Say something first";
      result.note = "Try: “sales 23.09.26”, “part number for display”, or a mobile number.";
      return JSON.stringify(result);
    }

    var c = chatQuery(query);

    // Task 27: EVERY query also scans every tab in the workbook for relevant
    // matches — the answer is never limited to the one handler's sheet.
    var scan = chatAllSheetsScan_(query);
    result.title = c.title;
    result.rows = c.rows || [];
    if (scan.rows.length) {
      if (result.rows.length) {
        result.rows.push([chatTxt('<div style="margin:4px 0 2px;font-size:.68rem;color:#9fb4e8">— matches across all tabs —</div>')]);
      }
      result.rows = result.rows.concat(scan.rows);
      result.note = (c.note ? c.note + " " : "") +
        "Cross-tab matches found in: " + esc_(scan.touched.join(", ")) + ".";
    } else {
      result.note = c.note || "Scanned all sheet tabs; nothing matched — try a date, part word, or mobile number.";
    }
    return JSON.stringify(result);
  } catch (err) {
    try {
      return JSON.stringify({
        title: "Search failed",
        rows: [],
        note: esc_(str_(err && err.message ? err.message : err)),
      });
    } catch (e2) {
      return '{"title":"Search failed","rows":[],"note":"unknown error"}';
    }
  }
}

/**
 * Small table-cell formatting helpers for chat answers. Returns rows as
 * arrays of cells [[td-html,...] ...], so HTML stays simple and safe.
 */
function chatNum(n) { return '<td class="num">' + Math.round(n || 0).toLocaleString("en-IN") + "</td>"; }
function chatMoneyCell(n) { return '<td class="num">₹' + Math.round(n || 0).toLocaleString("en-IN") + "</td>"; }
// chatTxt is HTML-raw: call sites esc_() sheet data themselves; chatPlain wraps + escapes.
function chatTxt(t) { return "<td>" + t + "</td>"; }
function chatPlain(t) { return "<td>" + esc_(t) + "</td>"; }
function chatBlank() { return '<td class="num muted">—</td>'; }

/** Canonical 10-digit mobile (strip +91 / leading 0). */
function chatDig(q) {
  var d = str_(q).replace(/\D/g, "");
  if (d.length > 10) d = d.slice(-10);
  return d.length === 10 ? d : "";
}

/**
 * Route the free-text question to the right search.
 * Returns { title, rows, note } where rows are table-cell arrays.
 */
function chatQuery(q) {
  var lower = str_(q).toLowerCase();

  // 1) A contiguous 10+ digit run (a real phone number) → customer 360
  // across sheets. Dates like 23.09.2026 never match — the dots break runs.
  var phoneMatch = str_(q).match(/[\d+][\d\s().\-]{8,}\d/);
  var mobile = phoneMatch ? chatDig(phoneMatch[0]) : "";
  if (mobile) return chatMobileQuery(mobile, q);

  // 2) Asks for a DATE / date range + sales → revenue & GP invoice breakdown.
  var wantsSales = /sales|sale\b|revenue|invoice|gross\s*profit|\bgp\b/i.test(lower);
  if (wantsSales) {
    var dates = chatExtractDates(q);
    if (dates.length) return chatSalesQuery(dates, q);
    // "sale date from 15.09.26" — date parsed above; without one, fall through.
    if (/date\s*from\s*[\d]/.test(lower)) {
      return {
        title: "No matching date found",
        rows: [],
        note: "I could not parse the date in that question — try “23.09.26” or “23-09-2026”.",
      };
    }
  }

  // 3) "part number for iphone 3 display" → lookup in CRM item codes +
  // service advance part-ish text. We look in crm ItemCode / PartType /
  // ItemDescription / SubCategory, and walkin/service repair docs too.
  if (/part|item\s*code|model|display\b|battery\b|glass\b|product/i.test(lower)) {
    return chatPartQuery(q);
  }

  // 4) Fall back to a fuzzy name / invoice / case / text search.
  return chatFuzzyQuery(q);
}

/** Extract every date written in the question (dot/slash/hyphen, yy or yyyy). */
function chatExtractDates(q) {
  var out = [];
  var seen = {};
  str_(q).replace(/\b(\d{1,2})[.\-\/](\d{1,2})[.\-\/](\d{2,4})\b/g, function (m, d, mo, y) {
    var q = parseInvoiceDate_(m.replace(/[.\/]/g, "-"));
    if (q.dayKey && !seen[q.dayKey]) {
      seen[q.dayKey] = true;
      out.push(q.dayKey);
    }
    return m;
  });
  out.sort(); // dd-MM-yyyy lexicographic… fine for our equality use; not sorted chronologically
  return out;
}

/** Chat answer for a mobile number — the 360 customer profile. */
function chatMobileQuery(mobile, origQ) {
  var ss = getSpreadsheet_();
  var last4 = mobile.slice(-4);

  // crm — invoice rows billed to this mobile (via walkin InvoiceCode links)
  var crm = chatInvoiceRowsForMobile(ss, mobile);

  // walkin — token history
  var walkin = chatWalkinRowsForMobile(ss, mobile);

  // service — open/closed cases
  var service = chatServiceRowsForMobile(ss, mobile);

  // advance — balances
  var advance = chatAdvanceRowsForMobile(ss, mobile);

  var rows = [];
  var title = "📱 Mobile …" + last4;

  // Customer name
  var name = "";
  if (service.length && (service[0].first || service[0].last)) name = (service[0].first + " " + service[0].last).trim();
  if (!name && crm.length) name = "customer";
  if (name) rows.push([chatTxt("<b>Customer</b>"), chatPlain(name)]);
  rows.push([chatTxt("<b>Mobile</b>"), chatTxt(mobile)]);

  // Cases (service sheet, then walkin case ids not in service)
  var caseIds = {};
  service.forEach(function (s) {
    if (s.caseId) caseIds[s.caseId] = s;
  });
  var caseArr = Object.keys(caseIds);
  if (caseArr.length) {
    rows.push([chatTxt("<b>Case IDs</b>"), chatPlain(caseArr.join(", "))]);
    rows.push([chatTxt("<b>Case detail</b>"), chatPlain(caseArr.map(function (c) {
      var s = caseIds[c];
      return c + " — " + (s.status || "open") + (s.gsx ? ", " + s.gsx : "");
    }).join(" · "))]);
  }

  // Invoices (walkin InvoiceCode joined to crm value)
  if (crm.invoices.length) {
    rows.push([chatTxt("<b>Invoices</b>"), chatPlain(crm.invoices.map(function (i) {
      return i.code + " (₹" + Math.round(i.total).toLocaleString("en-IN") + ")";
    }).join(" · "))]);
  }
  if (crm.total) {
    rows.push([chatTxt("<b>Invoice value</b>"), chatMoneyCell(crm.total)]);
    rows.push([chatTxt("<b>Gross profit</b>"), chatMoneyCell(crm.gp)]);
  }
  if (walkin.length) {
    rows.push([chatTxt("<b>Visits</b>"), chatTxt(String(walkin.length) + " visit" + (walkin.length === 1 ? "" : "s"))]);
  }
  if (advance && advance.total) {
    rows.push([chatTxt("<b>Adv. balance</b>"), chatMoneyCell(advance.total)]);
  }
  if (!crm.invoices.length && !caseArr.length && !service.length && !(advance && advance.count) && !walkin.length) {
    return {
      title: title,
      rows: [[chatTxt('<span class="muted">No records found for this number in any sheet.</span>')]],
      note: "Check the number, or try a date / part-number question.",
    };
  }
  return { title: title, rows: rows, note: "criss-crossed through walk-in, service, advance and CRM sheets." };
}

/** crm rows that reference this mobile (via the invoice-code link from walk-in). */
function chatInvoiceRowsForMobile(ss, mobile) {
  var out = { invoices: [], lines: [], total: 0, gp: 0 };
  var codes = {};
  // First: which invoice codes belong to this mobile? From walk-in sheet.
  var ws = getSheetByNames_(ss, CONFIG.WALKIN_SHEET_NAMES);
  if (ws) {
    var values = chatSheetData(ws);
    var col = {
      mobile: findColumn_(values.header, [/mobile/i, /number/i, /phone/i]),
      invoiceCode: findColumn_(values.header, [/invoice.*(code|no\.?|number|id)/i]),
      caseId: findColumn_(values.header, [/case.*id/i]),
    };
    values.rows.forEach(function (row) {
      var m = chatDig(row[col.mobile]);
      if (m === mobile && col.invoiceCode !== -1 && str_(row[col.invoiceCode]))
        codes[str_(row[col.invoiceCode])] = true;
    });
  }
  // Also: invoices whose TranType is INVOICE and InvoiceCode appears in walkin service sheets — but for a pure mobile answer we only need the codes we just found.
  var crmSheet = getSheetByNames_(ss, [CONFIG.CRM_SHEET_NAME]);
  if (crmSheet) {
    var values = parseCrmRows(chatSheetData(crmSheet).allValues);
    var found = {};
    values.forEach(function (r) {
      if (codes[r.invoiceCode]) {
        if (!found[r.invoiceCode]) {
          found[r.invoiceCode] = { code: r.invoiceCode, total: 0, gp: 0 };
          out.invoices.push(found[r.invoiceCode]);
        }
        found[r.invoiceCode].total += r.taxableAmount;
        found[r.invoiceCode].gp += r.grossProfit;
        out.lines.push(r);
        out.total += r.taxableAmount;
        out.gp += r.grossProfit;
      }
    });
  }
  return out;
}

/** All walk-in rows for this mobile. */
function chatWalkinRowsForMobile(ss, mobile) {
  var ws = getSheetByNames_(ss, CONFIG.WALKIN_SHEET_NAMES);
  if (!ws) return [];
  var values = chatSheetData(ws);
  var cMobile = findColumn_(values.header, [/mobile/i, /number/i, /phone/i]);
  var cDate = findColumn_(values.header, [/token.*date/i, /visit.*date/i, /date/i]);
  var cCase = findColumn_(values.header, [/case.*id/i]);
  var cInvoice = findColumn_(values.header, [/invoice.*(code|no\.?|number|id)/i]);
  var cPurpose = findColumn_(values.header, [/purpose/i]);
  var cRemark = findColumn_(values.header, [/remark/i]);
  var out = [];
  values.rows.forEach(function (r) {
    if (chatDig(r[cMobile]) === mobile) {
      out.push({
        date: cDate !== -1 ? str_(r[cDate]) : "",
        caseId: cCase !== -1 ? str_(r[cCase]) : "",
        invoiceCode: cInvoice !== -1 ? str_(r[cInvoice]) : "",
        purpose: cPurpose !== -1 ? str_(r[cPurpose]) : "",
        remark: cRemark !== -1 ? str_(r[cRemark]) : "",
      });
    }
  });
  return out;
}

/** Service rows for this mobile. */
function chatServiceRowsForMobile(ss, mobile) {
  var sh = getSheetByNames_(ss, CONFIG.SERVICE_SHEET_NAMES);
  if (!sh) return [];
  var values = chatSheetData(sh);
  var cMobile = findColumn_(values.header, [/mobile/i, /number/i, /phone/i]);
  var cFirst = findColumn_(values.header, [/first.*name/i, /name/i]);
  var cLast = findColumn_(values.header, [/last.*name/i]);
  var cCase = findColumn_(values.header, [/case.*id/i]);
  var cStatus = findColumn_(values.header, [/crm.*status/i, /status/i, /job.*status/i]);
  var cGsx = findColumn_(values.header, [/gsx/i, /repair.*status/i]);
  var out = [];
  values.rows.forEach(function (r) {
    if (chatDig(r[cMobile]) === mobile) {
      out.push({
        first: cFirst !== -1 ? str_(r[cFirst]) : "",
        last: cLast !== -1 ? str_(r[cLast]) : "",
        caseId: cCase !== -1 ? str_(r[cCase]) : "",
        status: cStatus !== -1 ? str_(r[cStatus]) : "",
        gsx: cGsx !== -1 ? str_(r[cGsx]) : "",
      });
    }
  });
  return out;
}

/** Advance-balance rows for this mobile. */
function chatAdvanceRowsForMobile(ss, mobile) {
  var sh = getSheetByNames_(ss, CONFIG.ADVANCE_SHEET_NAMES);
  if (!sh) return null;
  var values = chatSheetData(sh);
  var cMobile = findColumn_(values.header, [/mobile/i, /number/i, /phone/i]);
  var cBal = findColumn_(values.header, [/balance/i, /advance/i, /amount/i, /value/i]);
  var total = 0, count = 0;
  values.rows.forEach(function (r) {
    if (chatDig(r[cMobile]) === mobile) { total += num_(r[cBal]); count++; }
  });
  return { total: total, count: count };
}

/**
 * Read a sheet's values once, returning { allValues, header, rows }.
 * allValues is the raw grid (header + rows), header/lrows are split views.
 */
function chatSheetData(sheet) {
  var all = sheet.getDataRange().getDisplayValues();
  return { allValues: all, header: all[0] || [], rows: all.slice(1) };
}

/**
 * "sale date from 15.09.26" (plus optional second date "to 20.09.26").
 * Produces one row per invoice + a salesperson-wise summary with GP.
 */
function chatSalesQuery(dayKeys, origQ) {
  var ss = getSpreadsheet_();
  var crmSheet = getSheetByNames_(ss, [CONFIG.CRM_SHEET_NAME]);
  if (!crmSheet) {
    return { title: "CRM sheet missing", rows: [], note: 'Add a sheet named "' + CONFIG.CRM_SHEET_NAME + '".' };
  }
  var rows = parseCrmRows(chatSheetData(crmSheet).allValues);
  var wanted = {};
  dayKeys.forEach(function (k) { wanted[k] = true; });

  // If they said "from" and there are >=2 dates, treat as a range.
  var lower = str_(origQ).toLowerCase();
  var range = /from/.test(lower) && dayKeys.length >= 2;

  var matched = rows.filter(function (r) {
    if (!r.dayKey) return false;
    if (range) return r.dayKey >= dayKeys[0] && r.dayKey <= dayKeys[dayKeys.length - 1];
    return !!wanted[r.dayKey];
  });

  if (!matched.length) {
    return {
      title: "No invoices found",
      rows: [[chatTxt('<span class="muted">The crm sheet has no rows for ' +
        esc_(dayKeys.join(", ")) + ".</span>")]],
      note: "Try another date — dates are read as <b>dd-MM-yyyy</b> (so 15.09.26 = 15-09-2026).",
    };
  }

  // ---- per-invoice lines ----
  var invoiceMap = matchOrderInvoices_(matched);
  var title = "🧾 Sales on " + esc_(dayKeys.join(", "));
  var rows2 = [];
  invoiceMap.forEach(function (inv) {
    rows2.push([
      chatTxt("<b>" + esc_(inv.date) + "</b>"),
      chatPlain(inv.code),
      chatTxt(esc_(inv.person)),
      chatMoneyCell(inv.total),
      chatMoneyCell(inv.gp),
    ]);
  });
  // ---- salesperson-wise summary ----
  var people = {};
  matched.forEach(function (r) {
    var k = r.salesPerson || "(unassigned)";
    if (!people[k]) people[k] = { name: k, revenue: 0, gp: 0, invoices: {} };
    people[k].revenue += r.taxableAmount;
    people[k].gp += r.grossProfit;
    people[k].invoices[r.invoiceCode] = true;
  });
  var names = Object.keys(people).sort(function (a, b) { return people[b].revenue - people[a].revenue; });
  var totalRevenue = 0, totalGp = 0;
  names.forEach(function (n) {
    totalRevenue += people[n].revenue;
    totalGp += people[n].gp;
    rows2.push([
      chatTxt("<b>by " + esc_(people[n].name) + "</b>"),
      chatTxt(Object.keys(people[n].invoices).length + " inv"),
      chatBlank(),
      chatMoneyCell(people[n].revenue),
      chatMoneyCell(people[n].gp),
    ]);
  });
  rows2.push([
    chatTxt("<b>TOTAL</b>"), chatTxt(invoiceMap.length + " inv"), chatBlank(),
    chatMoneyCell(totalRevenue), chatMoneyCell(totalGp),
  ]);
  return { title: title, rows: rows2, note: "Revenue + gross profit per invoice, then per salesperson." };
}

/** Merge crm lines into one row per invoice, preserving invoice order. */
function matchOrderInvoices_(rows) {
  var order = [], byCode = {};
  rows.forEach(function (r) {
    if (!byCode[r.invoiceCode]) {
      byCode[r.invoiceCode] = { code: r.invoiceCode, date: r.dayKey, person: r.salesPerson, total: 0, gp: 0 };
      order.push(byCode[r.invoiceCode]);
    }
    byCode[r.invoiceCode] && (byCode[r.invoiceCode].total += r.taxableAmount, byCode[r.invoiceCode].gp += r.grossProfit);
  });
  return order;
}

/** "part number for iphone 3 display" → fuzzy match on crm item codes + walkin/service text. */
function chatPartQuery(q) {
  var ss = getSpreadsheet_();
  var stopWords = { whats: 1, what: 1, is: 1, the: 1, for: 1, part: 1, number: 1, find: 1, get: 1, give: 1, tell: 1, me: 1, show: 1, item: 1, code: 1, model: 1, details: 1, please: 1, inventory: 1, list: 1, price: 1 };
  var tokens = str_(q).toLowerCase()
    .replace(/[^a-z0-9\s]/g, " ")
    .split(/\s+/)
    .filter(function (t) { return t && !stopWords[t] && t.length >= 2; });
  if (!tokens.length) {
    return {
      title: "Unclear part name",
      rows: [],
      note: 'Say something like "part number for iPhone display" or "battery for Galaxy S23".',
    };
  }

  var crmSheet = getSheetByNames_(ss, [CONFIG.CRM_SHEET_NAME]);
  if (!crmSheet) return { title: "CRM sheet missing", rows: [], note: "" };
  var rows = parseCrmRows(chatSheetData(crmSheet).allValues);

  // Score every row: how many query tokens appear in a haystack of its fields.
  var scored = rows.map(function (r) {
    var hay = (r.itemCode + " " + r.itemDescription + " " + r.subCategory + " " + r.partType + " " + r.schemeCode + " " + r.revenueType).toLowerCase();
    var hits = tokens.filter(function (t) { return hay.indexOf(t) > -1; }).length;
    return { r: r, hits: hits };
  }).filter(function (s) { return s.hits > 0; });
  scored.sort(function (a, b) { return b.hits - a.hits || (b.r.taxableAmount - a.r.taxableAmount); });

  if (!scored.length) {
    return {
      title: "No matches",
      rows: [[chatTxt('<span class="muted">Nothing found for “' + esc_(tokens.join(" ")) + '” in the CRM sheet.</span>')]],
      note: "Check spelling, or ask for a date or mobile-number search.",
    };
  }

  var title = "🔧 Part: " + esc_(tokens.join(" "));
  var out = [[chatTxt("<b>Item code</b>"), chatTxt("<b>Description</b>"), chatTxt("<b>Part type</b>"), chatTxt("<b>On invoice</b>"), chatTxt("<b>₹</b>")]];
  var seen = {};
  scored.slice(0, 12).forEach(function (s) {
    var key = s.r.itemCode + "|" + s.r.itemDescription;
    if (seen[key]) return;
    seen[key] = true;
    out.push([
      chatTxt("<b>" + esc_(s.r.itemCode || "—") + "</b>"),
      chatPlain(s.r.itemDescription || "—"),
      chatPlain(s.r.partType || s.r.subCategory || "—"),
      chatPlain(s.r.invoiceCode),
      chatMoneyCell(s.r.taxableAmount),
    ]);
  });
  return { title: title, rows: out, note: "Item codes / part types matched against crm!ItemCode, ItemDescription, SubCategory, PartType." };
}

/** Last resort: fuzzy search on invoice code, case id, names, descriptions. */
function chatFuzzyQuery(q) {
  var ss = getSpreadsheet_();
  var needle = str_(q).toLowerCase();
  var crmSheet = getSheetByNames_(ss, [CONFIG.CRM_SHEET_NAME]);
  if (crmSheet) {
    var rows = parseCrmRows(chatSheetData(crmSheet).allValues);
    var matches = rows.filter(function (r) {
      var hay = (r.invoiceCode + " " + r.salesPerson + " " + r.itemDescription + " " + r.itemCode + " " + r.subCategory + " " + r.partType + " " + r.schemeCode).toLowerCase();
      return needle.split(/\s+/).filter(function (t) { return t.length > 2; }).every(function (t) { return hay.indexOf(t) > -1; });
    });
    if (matches.length) {
      return {
        title: "🔎 Matches in crm",
        rows: matches.slice(0, 10).map(function (r) {
          return [chatTxt(r.dayKey), chatPlain(r.invoiceCode), chatPlain(r.salesPerson),
            chatPlain(r.itemDescription || r.itemCode || ""), chatMoneyCell(r.taxableAmount)];
        }),
        note: "",
      };
    }
  }
  return {
    title: "Not sure what you mean",
    rows: [],
    note: 'Try a date: “sales 23.09.26”, a part: “part number for iPhone display”, or a mobile: “919876543210”.',
  };
}

/**
 * Task 27 blanket search: scan EVERY tab in the workbook for rows whose text
 * contains every query token, regardless of which handler answered above.
 * Stop words are stripped so "sale date from 23.09.26" focuses on the date
 * digits, "9791994639" focuses on the digits, etc.
 */
function chatAllSheetsScan_(query) {
  var STOP = { sale:1, sales:1, date:1, day:1, from:1, to:1, since:1, what:1, whats:1, is:1, the:1,
    part:1, number:1, item:1, code:1, for:1, tell:1, me:1, show:1, give:1, details:1, of:1, and:1, or:1,
    on:1, in:1, find:1, get:1, list:1, whole:1, full:1, all:1 };
  var raw = str_(query).toLowerCase();
  var sq = raw.replace(/[^a-z0-9@.\-]/g, " ").trim();
  var tokens = sq.split(/\s+/).filter(function (t) { return t.length >= 2 && !STOP[t]; });
  // Keep at most the 4 most distinctive tokens (longest first) to avoid over-tight matching.
  tokens.sort(function (a, b) { return b.length - a.length; });
  var useTokens = tokens.slice(0, 4);

  var ss = getSpreadsheet_();
  var sheets = ss.getSheets();
  var outRows = [];
  var touched = [];
  var PER_SHEET_CAP = 3;
  var TOTAL_CAP = 12;

  sheets.forEach(function (sheet) {
    if (outRows.length >= TOTAL_CAP) return;
    var name = sheet.getName();
    var values = sheet.getDataRange().getDisplayValues();
    var found = [];
    for (var r = 1; r < values.length && found.length < PER_SHEET_CAP; r++) {
      var cells = values[r].map(function (v) { return str_(v); }).filter(function (v) { return v !== ""; });
      if (!cells.length) continue;
      var hay = cells.join(" ").toLowerCase();
      var hit = useTokens.length
        ? useTokens.every(function (t) { return hay.indexOf(t) > -1; })
        : (sq.length >= 2 && hay.indexOf(sq) > -1);
      if (!hit) continue;
      // Summarise: first 4 non-empty cells, each ≤24 chars.
      var bits = cells.slice(0, 6).map(function (c) {
        return c.length > 24 ? (esc_(c.slice(0, 23)) + "…") : esc_(c);
      });
      found.push([chatTxt("<b>" + esc_(name) + "</b>"), chatTxt(bits.join(" · "))]);
    }
    if (found.length) {
      outRows = outRows.concat(found);
      touched.push(name);
    }
  });

  return { rows: outRows, touched: touched };
}

/** Escape for the chat's HTML table cells (plain text variants live above). */

/* ============================================================================
 * INDEX LANDING PAGE — neon liquid-glass tiles
 *
 * One translucent tile per dashboard. Hovering makes the tile opaque and
 * reveals a handful of key numbers; clicking jumps straight to the pane that
 * holds the content. Groups below follow the owner's page split:
 *   GP / KPI / WALKIN / Operation / OPEN JOBS / DATE.
 * ========================================================================= */

/* ---------------------------------------------------------------------------
 * Small horizontal bar chart used under each index group title — one bar row
 * per sub-title value, combining the group's numbers into a skimmable graph.
 * segments: [[label, value, color], ...]
 * ------------------------------------------------------------------------- */
var PIE_PALETTE = ["#3f5ce0","#63e6b7","#ffd477","#ff9da7","#8fc7ff","#dfb2ff"];
function barBlockHtml_(title, segments) {
  var max = 1;
  (segments || []).forEach(function (sg) { if (sg[1] > max) max = sg[1]; });
  var rows = (segments || []).map(function (sg) {
    var w = Math.max(2, Math.round((sg[1] / max) * 100));
    return '<div class="gbar-row"><span class="gbar-lbl">' + esc_(sg[0]) + "</span>" +
      '<span class="gbar-track"><i class="gbar-fill" style="width:' + w + "%;background:" + sg[2] + '"></i></span>' +
      '<b class="gbar-val">' + esc_(String(sg[1])) + "</b></div>";
  }).join("");
  if (!(segments || []).length) rows = '<div class="gpie-empty">no data yet</div>';
  return '<div class="gpie-wrap"><div class="gpie-title">' + esc_(title) + "</div>" +
    '<div class="gbar-group">' + rows + "</div></div>";
}
function pieSvgHtml_(segments, cx, cy, radius, thick) {
  cx = cx || 60; cy = cy || 60; radius = radius || 44; thick = thick || 16;
  var segs = (segments || []).filter(function (sg) { return sg[1] > 0; });
  var total = 0;
  segs.forEach(function (sg) { total += sg[1]; });
  if (total <= 0) {
    return '<div class="gpie gpie-empty"><span class="gpie-note">no data yet</span></div>';
  }
  var k = { "cx": cx, "cy": cy, "r": radius, "fill": "none", "stroke-width": thick };
  var circ = 2 * Math.PI * radius;
  var acc = -circ / 4; // start at 12 o'clock
  var parts = ['<svg class="gpie" width="' + (cx * 2) + '" height="' + (cy * 2) +
    '" viewBox="0 0 ' + (cx * 2) + " " + (cy * 2) + '" role="img">'];
  parts.push('<circle cx="' + cx + '" cy="' + cy + '" r="' + (radius - thick / 2 - 2) +
    '" fill="rgba(12,29,53,.92)" />');
  segs.forEach(function (sg) {
    var frac = sg[1] / total;
    var len = Math.max(0.01, frac * (circ - segs.length * 2));
    parts.push('<circle cx="' + cx + '" cy="' + cy + '" r="' + radius +
      '" fill="none" stroke="' + sg[2] + '" stroke-width="' + thick +
      '" stroke-dasharray="' + len.toFixed(2) + " " + (circ - len).toFixed(2) +
      '" stroke-dashoffset="' + (-acc).toFixed(2) +
      '" stroke-linecap="butt" transform="rotate(0 ' + cx + " " + cy + ')" />');
    acc += frac * circ;
  });
  parts.push("</svg>");
  return parts.join("");
}

/* ============================================================================
 * INDEX LANDING PAGE — liquid-glass tiles laid out HORIZONTALLY by group,
 * each group's sub-title tiles stacked one below the other, and a horizontal
 * bar chart under every group title combining that group's sub-title numbers.
 * Hover: the glass brightens and the golden glitter video plays over it.
 * ========================================================================= */

function renderIndexHtml_(data, perfData) {
  var s = data.store;
  function p1(x) { return (x || 0).toFixed(1) + "%"; }
  var walkT = perfData && perfData.walkin && perfData.walkin.available ? perfData.walkin.totals : null;
  var jobT = perfData && perfData.jobs && perfData.jobs.available ? perfData.jobs.totals : null;
  var csatT = perfData && perfData.targetBlock && perfData.targetBlock.available ? perfData.targetBlock : null;
  var ra = perfData && perfData.repair && perfData.repair.available ? perfData.repair : null;

  function tile(page, icon, title, blurb, stats) {
    var statHtml = stats && stats.length
      ? '<div class="tile-data">' + stats.map(function (st) {
          return '<div class="tile-stat"><span>' + esc_(st[0]) + "</span><b>" + esc_(String(st[1])) + "</b></div>";
        }).join("") + "</div>"
      : '<div class="tile-data"><div class="tile-soon">⏳ Content coming soon</div>' +
        "<div class=\"tile-soon-sub\">reserved by the owner for this module</div></div>";
    return (
      '<button type="button" class="glass-tile" data-page="' + esc_(page) +
      '" onclick="gotoPage(\'' + esc_(page) + "', this)\" aria-label=\"" + esc_(title) + "\">" +
      '<span class="tile-sparkle" aria-hidden="true"></span>' +
      '<video class="tile-glitter-vid" autoplay muted loop playsinline preload="metadata" ' +
        'poster="https://www.shutterstock.com/shutterstock/videos/1064124616/thumb/8.jpg?ip=x480" aria-hidden="true">' +
        '<source src="https://www.shutterstock.com/shutterstock/videos/1064124616/preview/stock-footage-shining-gold-dust-particles-golden-frame-golden-glitter-sparkling-magic-light-with-bokeh-effect.webm" type="video/webm">' +
        '<source src="https://www.shutterstock.com/shutterstock/videos/1064124616/preview/stock-footage-shining-gold-dust-particles-golden-frame-golden-glitter-sparkling-magic-light-with-bokeh-effect.mp4" type="video/mp4">' +
        '</video>' +
      '<span class="tile-icon">' + icon + "</span>" +
      '<span class="tile-title">' + esc_(title) + "</span>" +
      '<span class="tile-blurb">' + esc_(blurb) + "</span>" +
      statHtml +
      '<span class="tile-go">Open page &rarr;</span>' +
      "</button>"
    );
  }

  function group(title, page, tilesHtml, pieHtml) {
    return '<div class="tile-group"><h3 class="tile-group-head">' +
      '<button type="button" class="tile-group-link" onclick="gotoPage(\'' + esc_(page) +
      "')\" title=\"Open the " + esc_(title) + ' page">' + esc_(title) + " &rarr;</button></h3>" +
      pieHtml + '<div class="tile-col">' + tilesHtml + "</div></div>";
  }

  var C = PIE_PALETTE;

  // ---------- pies: one per group, combining its sub-title numbers ----------
  var gpPie = barBlockHtml_("Revenue won so far, by type", [
    ["Service", Math.round(s.service.achieved), C[0]],
    ["Accessory", Math.round(s.accessory.achieved), C[1]],
    ["Repair", Math.round(s.repair.achieved), C[2]],
  ]);

  var kpis = data.kpi;
  function kpiDone(rows) { var t = 0; rows.forEach(function (k) { t += k.achieved; }); return t; }
  var kpiPie = barBlockHtml_("KPIs hit so far, by category", [
    ["AR", kpiDone(kpis.ar), C[0]],
    ["RR", kpiDone(kpis.rr), C[1]],
    ["SR", kpiDone(kpis.sr), C[2]],
  ]);

  var wlT = walkT || { walkins: 0, unique: 0, cases: 0, invoices: 0, deadLeads: 0 };
  var wlLive = Math.max(0, wlT.walkins - wlT.cases - wlT.deadLeads);
  var walkinPie = barBlockHtml_("This month's walk-ins, split", [
    ["Cases", wlT.cases, C[0]],
    ["Invoices", wlT.invoices, C[1]],
    ["Live leads", wlLive, C[4]],
    ["Dead leads", wlT.deadLeads, C[3]],
  ]);

  var csatNeutral = 0;
  if (csatT) csatNeutral = Math.max(0, csatT.totalRow.total - csatT.satTotal - csatT.disTotal);
  var opPie = barBlockHtml_("Feedback balance", [
    ["Satisfied", csatT ? csatT.satTotal : 0, C[1]],
    ["Neutral", csatT ? csatNeutral : 0, C[4]],
    ["Dissatisfied", csatT ? csatT.disTotal : 0, C[3]],
  ]);

  var jobsList = (perfData && perfData.jobs && perfData.jobs.jobs) || [];
  var commOk = 0, commNo = 0;
  jobsList.forEach(function (jb) { if (jb.communicated) commOk++; else commNo++; });
  var jobsPie = barBlockHtml_("Open jobs, communication cover", [
    ["Communicated", commOk, C[1]],
    ["Not communicated", commNo, C[3]],
  ]);

  var days = data.days || [];
  var dated = days.filter(function (d) { return !d.undated; }).length;
  var datePie = barBlockHtml_("Month coverage", [
    ["Days with sales", dated, C[0]],
    ["Working days used", s.consumedWorkingDays, C[1]],
    ["Working days left", s.balanceWorkingDays, C[4]],
  ]);

  // ---------- tiles (unchanged content, now stacked vertically in-column) ----------
  var kpiRows = kpis.ar.concat(kpis.rr, kpis.sr);
  var kpiMet = kpiRows.filter(function (k) { return k.target > 0 && k.achieved >= k.target; }).length;
  var payable = data.salespeople.reduce(function (a, p) { return a + (p.payableIncentive || 0); }, 0);
  var a = data.invoiceValuation.arAttachment;

  var gpTiles =
    tile("gp", "🎯", "Store Target", "Performance vs the month plan", [
      ["Total", money_(s.total.achieved) + " / " + money_(s.total.target)],
      ["Achieved", p1(s.total.pct)],
      ["Balance", money_(s.total.balance)],
    ]) +
    tile("gp", "🗓", "Weekly Revenue Analysis", "Current month, week by week", [
      ["This week's plan", "target/working-days × week wd"],
      ["Working days", s.totalWorkingDays + " in " + s.daysInMonth + " days"],
      ["Consumed", s.consumedWorkingDays + " · bal " + s.balanceWorkingDays],
    ]) +
    tile("gp", "🧾", "Invoice Valuation", "What each invoice is worth", [
      ["Invoices", s.uniqueInvoices],
      ["Avg value", money_(s.avgInvoiceValue)],
      ["Total revenue", money_(s.total.achieved)],
    ]);

  var kpiTiles =
    tile("kpi", "📊", "KPI Target", "AR / RR / SR scorecard", [
      ["KPIs tracked", kpiRows.length],
      ["Targets met", kpiMet + " / " + kpiRows.length],
      ["Reviews", "+ " + kpis.googleReview.difference + " this month"],
    ]) +
    tile("kpi", "💰", "Sales Person Incentive", "Slab-wise incentive valuation", [
      ["Payable", money_(payable)],
      ["People", data.salespeople.length + " on roster"],
      ["Unlock gate", CONFIG.REPAIR_INCENTIVE_GATE_PCT + "% of store repair"],
    ]) +
    tile("kpi", "👤", "Sales Person Performance", "Per-person target vs achieved", [
      ["Top biller", data.salespeople.length ? data.salespeople[0].name : "—"],
      ["Repair shown", "each person's own win"],
      ["Gate calc", CONFIG.REPAIR_INCENTIVE_GATE_PCT + "% of store repair"],
    ]) +
    tile("kpi", "🔗", "AR Attachment", "Accessory sold with service / repair", [
      ["Attach rate", a.attachRate.toFixed(1) + "%"],
      ["Attached", a.attachedInvoices + " of " + a.attachableInvoices],
      ["AR value", money_(a.value)],
    ]);

  var walkinTiles =
    tile("walkin", "🚶", "Walk-in Analysis", "Visitors, cases & invoices per day", walkT ? [
      ["Walk-ins", wlT.walkins],
      ["Unique", wlT.unique + " visitors"],
      ["Dead leads", wlT.deadLeads],
    ] : [["Status", "walkin sheet not found"], ["", "add the sheet to enable"]]) +
    tile("walkin", "🎟", "Token Creation", "Which user raised each token", walkT && perfData.walkin.creators.length ? [
      ["Top creator", perfData.walkin.creators[0].name],
      ["Tokens", perfData.walkin.creators[0].count + " (" + perfData.walkin.creators[0].share + "%)"],
      ["Creators", perfData.walkin.creators.length + " users"],
    ] : [["Created-by col", "not detected"], ["", "add TokenCreatedBy"]]);

  var operationTiles =
    tile("operation", "😊", "Customer Satisfaction", csatT ? "C-SAT · " + csatT.periodName : "C-SAT feedback board", csatT ? [
      ["Responses", csatT.totalRow.total],
      ["Satisfied", csatT.totalRow.total > 0 ? Math.round((csatT.satTotal / csatT.totalRow.total) * 100) + "%" : "—"],
      ["C-SAT", csatT.csat || "—"],
    ] : [["Status", "target sheet block empty"], ["", "fill L1:U5"]]) +
    tile("operation", "🚫", "Decline Repair", "Estimate declined by customer", null) +
    tile("operation", "🔁", "Requote - Winback", "Re-quoted & won back customers", null);

  var openJobTiles =
    tile("jobs", "🔧", "Job Management", "Open jobs by repair bucket", jobT ? [
      ["Open jobs", jobT.openJobs],
      ["Communicated", jobT.communicated + " (" + jobT.commRate + "%)"],
      ["Avg job age", jobT.avgAge + " d"],
    ] : [["Status", "service sheet not found"], ["", "add the sheet to enable"]]);

  var dateTiles =
    tile("date", "📆", "Date-wise Sales Revenue Analysis", "Day × salesperson × type matrix", [
      ["Today", s.daysElapsed + " of " + s.daysInMonth + " days"],
      ["Daily target", money_(s.dailyTarget)],
      ["Working days left", s.balanceWorkingDays],
    ]);

  return (
    '<section class="index-hero"><h2>Store Performance Dashboard</h2></section>' +
    '<div class="tile-board">' +
    group("GP", "gp", gpTiles, gpPie) +
    group("KPI", "kpi", kpiTiles, kpiPie) +
    group("WALKIN", "walkin", walkinTiles, walkinPie) +
    group("Operation", "operation", operationTiles, opPie) +
    group("OPEN JOBS", "jobs", openJobTiles, jobsPie) +
    group("DATE", "date", dateTiles, datePie) +
    "</div>" +
    freshnessLineHtml_(data, perfData)
  );
}

/** Bottom line on the index: latest dated record seen in each sheet. */
function freshnessLineHtml_(data, perfData) {
  function fmtDateKey(year, month, day) {
    function p(x) { return x < 10 ? "0" + x : x; }
    return p(day) + "-" + p(month) + "-" + year;
  }
  var bits = [];
  // crm — from the date-wise rows of the business data
  var latestCrm = null;
  (data.days || []).forEach(function (d) {
    if (d.undated || !d.dateKey) return;
    if (!latestCrm || d.sortKey > latestCrm.sortKey) latestCrm = d;
  });
  bits.push(["crm (sales)", latestCrm ? latestCrm.dateKey : "no rows"]);
  bits.push(["target (plan)", "refreshed " +
    Utilities.formatDate(data.generatedAt, Session.getScriptTimeZone(), "dd-MM-yyyy" )]);
  // walkin
  var wlD = perfData && perfData.walkin && perfData.walkin.days
    ? perfData.walkin.days.filter(function (d) { return !d.undated; })[0] : null;
  bits.push(["walkin", wlD ? wlD.dateKey : "no rows"]);
  // service (open jobs page) — max CaseDate among parsed rows
  var svcTs = 0, svcKey = "";
  (perfData && perfData.jobs && perfData.jobs.serviceRows || []).forEach(function (r) {
    if (r.caseTs > svcTs) { svcTs = r.caseTs; svcKey = r.createdDate || ""; }
  });
  bits.push(["service", svcTs ? (svcKey || "—") : "no rows"]);
  // advance / comm freshness carried by the perf loader when present
  var fr = perfData && perfData.freshness ? perfData.freshness : null;
  if (fr) {
    bits.push(["advance", fr.advance || "—"]);
    bits.push(["comm", fr.comm || "—"]);
  }
  var li = bits.map(function (b) {
    return '<span class="fresh-chip"><i>' + esc_(b[0]) + "</i><b>" + esc_(b[1]) + "</b></span>";
  }).join("");
  return '<div class="fresh-line"><div class="fresh-line-head">⏱ Latest records per sheet</div>' +
    '<div class="fresh-row">' + li + "</div></div>";
}

/** Assembles the complete HTML document. */
function buildDashboardPage_(data) {
  var css = getDashboardCss_();
  var storeCode = (data.store && data.store.storeCode) ? data.store.storeCode : "";
  var pageTitle = "Store Performance Dashboard" + (storeCode ? " - " + storeCode : "");

  var perfData = loadPerformanceData_();

  var PAGES = [
    // [tab label, data-page id, main id, html, meta-subtitle]
    ["🏠 Index",     "index",     "pageIndex",     renderIndexHtml_(data, perfData), "Landing — pick a tile to jump to its page"],
    ["🎯 GP",        "gp",        "pageGp",        renderGpHtml_(data),       renderMetaText_(data)],
    ["📊 KPI",       "kpi",       "pageKpi",       renderKpiHtml_(data),      renderMetaText_(data)],
    ["🚶 WALK-IN",   "walkin",    "pageWalkin",    renderWalkinHtml_(perfData),  perfMeta_(perfData)],
    ["⚙️ Operation", "operation", "pageOperation", renderOperationHtml_(perfData), repairMeta_(perfData.repair, perfData.generatedAt)],
    ["🧰 Open Jobs", "jobs",      "pageJobs",      renderJobsHtml_(perfData), perfMeta_(perfData)],
    ["📅 DATE",      "date",      "pageDate",      renderDateHtml_(data),     renderMetaText_(data)],
    // DECK lives as the last tab but is intentionally NOT on the index page.
    ["🗂 Deck",      "deck",      "pageDeck",      renderDeckHtml_(data, perfData), "Month at a glance — store, KPI, walk-ins, C-SAT, jobs"],
  ];

  var tabs = PAGES.map(function (pg, idx) {
    return '<button class="tab-btn' + (idx === 0 ? " active" : "") + '" data-page="' + pg[1] +
      '" onclick="switchPage(this)">' + pg[0] + "</button>";
  }).join("");

  var mains = PAGES.map(function (pg, idx) {
    return '<main class="wrap" id="' + pg[2] + '" data-pane="' + pg[1] + '" data-meta="' + esc_(pg[4]) + '"' +
      (idx === 0 ? "" : ' style="display:none;"') + ">" + pg[3] + "</main>";
  }).join("");

  // Refresh runs through google.script.run because the page is served inside
  // a sandboxed iframe — window.location there points at an internal
  // userCodeAppPanel URL, so navigating it would blank the page.
  var clientScript =
    "<script>" +
    // Section collapse/expand for the DATE matrix. Rows outside the
    // recent-days window carry .extra-day and are hidden by .collapsed on the
    // table, so toggling is a pure CSS class flip with no server call.
    "function dwToggle(btn){" +
    "  var t=document.getElementById('dwTable'); if(!t) return;" +
    "  var open=t.classList.toggle('collapsed')===false;" +
    "  btn.classList.toggle('open', open);" +
    "  btn.querySelector('.txt').textContent=open?btn.dataset.less:btn.dataset.more;" +
    "}" +
    // Pane routing: flip the active tab, show the matching <main>, swap the
    // subtitle from its data-meta. Panes are keyed by data-pane.
    "function paneId(page){return {index:'pageIndex',gp:'pageGp',kpi:'pageKpi',walkin:'pageWalkin',operation:'pageOperation',jobs:'pageJobs',date:'pageDate'}[page]||'pageIndex';}" +
    "function paneEl(page){return document.getElementById(paneId(page));}" +
    "function currentPage(){var a=document.querySelector('.tab-btn.active');return a?a.dataset.page:'index';}" +
    "function showPane(page){" +
    "  var mains=document.querySelectorAll('main[data-pane]');" +
    "  for(var i=0;i<mains.length;i++){mains[i].style.display=(mains[i].getAttribute('data-pane')===page)?'block':'none';}" +
    "  var active=paneEl(page);" +
    "  if(active){document.getElementById('meta').textContent=active.getAttribute('data-meta')||'';}" +
    "  window.scrollTo({top:0,behavior:'smooth'});" +
    "}" +
    "function switchPage(btn){" +
    "  var tabs=document.querySelectorAll('.tab-btn');" +
    "  for(var i=0;i<tabs.length;i++){tabs[i].classList.toggle('active', tabs[i]===btn);}" +
    "  showPane(btn.dataset.page);" +
    "}" +
    // Tile click → flip to that page (tab highlight follows automatically).
    "function gotoPage(page, el){" +
    "  var tabs=document.querySelectorAll('.tab-btn');" +
    "  for(var i=0;i<tabs.length;i++){tabs[i].classList.toggle('active', tabs[i].dataset.page===page);}" +
    "  showPane(page);" +
    "}" +
    "function tickClock(){" +
    "  var el=document.getElementById('liveClock'); if(!el) return;" +
    "  var d=new Date(); function p(n){return (n<10?'0':'')+n;}" +
    "  var h=d.getHours(), ap=h>=12?'PM':'AM'; h=h%12||12;" +
    "  el.textContent=p(d.getDate())+'-'+p(d.getMonth()+1)+'-'+d.getFullYear()+' '+" +
    "    p(h)+':'+p(d.getMinutes())+':'+p(d.getSeconds())+' '+ap;" +
    "}" +
    "tickClock(); setInterval(tickClock,1000);" +
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
    "        var map={index:'pageIndex',gp:'pageGp',kpi:'pageKpi',walkin:'pageWalkin',operation:'pageOperation',jobs:'pageJobs',date:'pageDate',deck:'pageDeck'};" +
    "        Object.keys(map).forEach(function(page){" +
    "          var el=document.getElementById(map[page]);" +
    "          if(el && d[page]!==undefined){ el.innerHTML=d[page]; }" +
    "        });" +
    "        var active=paneEl(currentPage());" +
    "        if(active){document.getElementById('meta').textContent=active.getAttribute('data-meta')||'';}" +
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
    "  .withFailureHandler(function(err){" +
    "      var n=document.getElementById('note');" +
    "      n.textContent='Refresh failed: '+(err && err.message ? err.message : err);" +
    "      n.className='note err';" +
    "      done();" +
    "    })" +
    "    .refreshDashboardSections();" +
    "}" +
    // ---- chat bot: toggle / send / answer ----
    "function chatToggle(){" +
    "  var p=document.getElementById('chatPanel'); if(!p) return;" +
    "  p.classList.toggle('open');" +
    "  if(p.classList.contains('open')){var i=document.getElementById('chatInput'); if(i) i.focus();}" +
    "}" +
    "function chatBubble(cls,html){" +
    "  var log=document.getElementById('chatLog'); if(!log) return null;" +
    "  var d=document.createElement('div'); d.className='chat-msg '+cls; d.innerHTML=html;" +
    "  log.appendChild(d); log.scrollTop=log.scrollHeight; return d;" +
    "}" +
    "function chatSend(){" +
    "  var inp=document.getElementById('chatInput'); if(!inp) return;" +
    "  var q=(inp.value||'').trim(); if(!q) return;" +
    "  inp.value='';" +
    "  var safe=q.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');" +
    "  chatBubble('user',safe);" +
    "  var thinking=chatBubble('bot','<span class=\\\"chat-dots\\\"><i></i><i></i><i></i></span>');" +
    "  var finish=function(html){ if(thinking) thinking.innerHTML=html; };" +
    "  if(!(window.google && google.script && google.script.run)){" +
    "    finish('Chat replies need the live web app (this is a static preview).');" +
    "    return;" +
    "  }" +
    "  google.script.run" +
    "    .withSuccessHandler(function(raw){" +
    "      var res; try{res=JSON.parse(raw);}catch(e){res={title:'Reply error',rows:[],note:String(e)};}" +
    "      window.__waLastQuery=q; window.__waLastText=waTextOf_(res);" +
    "      var h='<div class=\\\"chat-title\\\">'+res.title+'</div>';" +
    "      if(res.rows && res.rows.length){" +
    "        h+='<table class=\\\"chat-table\\\"><tbody>';" +
    "        for(var i=0;i<res.rows.length;i++){h+='<tr>'+res.rows[i].join('')+'</tr>';}" +
    "        h+='</tbody></table>';" +
    "      }" +
    "      if(res.note){h+='<div class=\\\"muted\\\">'+res.note+'</div>';}" +
    "      finish(h);" +
    "    })" +
    "    .withFailureHandler(function(err){" +
    "      finish('Search failed: '+(err && err.message ? err.message : err));" +
    "    })" +
    "    .serverQuery(q);" +
    "}" +
    // ---- WhatsApp share of the last bot answer (Task 28) ----
    "function waTextOf_(res){" +
    "  var rows=(res.rows||[]).map(function(r){return r.map(function(c){return c.replace(/<[^>]+>/g,'');}).filter(Boolean).join('  ·  ');});" +
    "  return [res.title].concat(rows).concat(res.note?[res.note]:[]).slice(0,22).join('\\n');" +
    "}" +
    "function waSendLast(){" +
    "  var q=window.__waLastQuery, msg=window.__waLastText;" +
    "  if(!q||!msg){chatBubble('bot','Ask something first, then send the answer to WhatsApp.');return;}" +
    "  var chat=(localStorage.getItem('wa_chatid')||'').trim();" +
    "  if(!chat){" +
    "    chat=(prompt('WhatsApp number (with country code, e.g. 919876543210):')||'').trim();" +
    "    if(chat){localStorage.setItem('wa_chatid',chat);}" +
    "  }" +
    "  if(!chat){chatBubble('bot','No WhatsApp number set — cancelled.');return;}" +
    "  chatBubble('user','WhatsApp→'+chat.replace(/&/g,'&amp;').replace(/</g,'&lt;'));" +
    "  var thinking=chatBubble('bot','<span class=\\\"chat-dots\\\"><i></i><i></i><i></i></span>');" +
    "  var finish=function(html){ if(thinking) thinking.innerHTML=html; };" +
    "  if(!(window.google && google.script && google.script.run)){" +
    "    finish('WhatsApp send needs the live web app (static preview).'); return;" +
    "  }" +
    "  var payload=JSON.stringify({q:q,chat:chat,text:msg});" +
    "  google.script.run" +
    "    .withSuccessHandler(function(raw){" +
    "      var r; try{r=JSON.parse(raw);}catch(e){r={};}" +
    "      if(r.ok){finish('✅ Sent to '+chat+' via OpenWA gateway.');return;}" +
    "      if(r.offline && r.link){finish('Gateway not configured — opening WhatsApp directly…');window.open(r.link,'_blank');return;}" +
    "      finish('WhatsApp send failed: '+(r.error||'unknown'));" +
    "    })" +
    "    .withFailureHandler(function(err){" +
    "      finish('WhatsApp send failed: '+(err && err.message ? err.message : err));" +
    "    })" +
    "    .sendWhatsAppAnswer(payload);" +
    "}" +
    "</script>";

  var bizMeta = renderMetaText_(data);

  return (
    "<!DOCTYPE html><html><head><base target=\"_top\"><meta charset=\"utf-8\">" +
    "<title>" + esc_(pageTitle) + "</title><style>" + css + "</style></head><body>" +
    // Task 27 backdrop for EVERY page: starry deep-blue pinimg GIF at 60%
    // visibility — dark enough that all foreground detail stays readable.
    '<img id="globalBgGif" class="global-bg-gif" ' +
      'src="https://i.pinimg.com/originals/2c/7c/cf/2c7ccfba6490b77f6688ae0d429d3a3a.gif" ' +
      'alt="" aria-hidden="true">' +
    '<header class="topbar"><div class="topbar-inner">' +
    "<div class=\"topbar-left\">" +
    '<div class="brand">' + esc_(pageTitle) + "</div>" +
    '<div class="meta" id="meta">' + esc_(bizMeta) + "</div>" +
    '<div class="note" id="note"></div>' +
    "</div>" +
    // ---- live clock, top-middle: owner-supplied red-LED BRG clock gif as the card background ----
    '<div class="topbar-clock" role="timer">' +
    '<img class="clock-gif" aria-hidden="true" alt="" ' +
    'src="https://www.brgprecision.com/phpinclude/saclocks/alpha/double/images/AC40-01-10-02-R%20animation.gif">' +
    '<span class="clock-dim" aria-hidden="true"></span>' +
    '<span class="clock-text" id="liveClock">—</span>' +
    "</div>" +
    '<button class="refresh-btn" onclick="dashRefresh(this)">' +
    '<span class="spin"></span><span class="txt">↻ Refresh data</span></button>' +
    // ---- tabs: Index / GP / KPI / Walk-in / Operation / Open Jobs / Date ----
    '<div class="tabs-row"><div class="tabs tabs-wide" role="tablist">' + tabs + "</div></div>" +
    "</div></header>" +
    mains +
    // ---- floating sheet-search chat bot, present on every page ----
    '<div class="chat-widget" id="chatWidget">' +
      '<div class="chat-fab" id="chatFab" onclick="chatToggle()" role="button" tabindex="0" ' +
        'aria-label="Open sheet search chat">' +
        '<span class="chat-fab-icon">💬</span><span>ask me anything</span>' +
      '</div>' +
      '<div class="chat-panel" id="chatPanel" role="dialog" aria-label="Sheet search chat">' +
        '<div class="chat-head"><b>🤖 ask me anything</b>' +
          '<span class="chat-sub">ask about sales, parts, customers</span>' +
          '<button class="chat-close" onclick="chatToggle()" aria-label="Close chat">&times;</button>' +
        "</div>" +
        '<div class="chat-log" id="chatLog">' +
          '<div class="chat-msg bot">Ask me anything from the source sheets — e.g. ' +
            "<i>sale date from 23.09.26</i>, <i>part number for display</i>, " +
            "or a mobile number like <i>919876543210</i>." +
            '<br>Tap <b>✆</b> to send an answer over WhatsApp (OpenWA gateway — set ' +
            '<code>WA_GATEWAY_URL/WA_API_KEY/WA_CHATID</code> in Script Properties).</div>' +
        "</div>" +
        '<div class="chat-inputrow">' +
          '<input class="chat-input" id="chatInput" type="text" autocomplete="off" ' +
            'placeholder="Try: sales 23.09.26, part number for…, or a mobile number" ' +
            "onkeydown=\"if(event.key==='Enter')chatSend()\">" +
          '<button class="chat-send" onclick="chatSend()" aria-label="Send">➤</button>' +
          '<button class="chat-send chat-wa" onclick="waSendLast()" ' +
            'aria-label="Send last answer to WhatsApp" title="Send last answer to WhatsApp">✆</button>' +
        "</div>" +
      "</div>" +
    "</div>" +
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
  // Summary rendered as ONE table (was a row of tiles) — one row per metric.
  var head =
    '<div class="scroll-x" style="margin-bottom:14px"><table class="ptable walksum">' +
    "<thead><tr><th>Metric</th>" + '<th class="num">Value</th>' +
    "<th>Reading</th></tr></thead><tbody>" +
    "<tr><td><b>Total Walk-ins</b></td>" + '<td class="num"><b>' + t.walkins + "</b></td>" +
      "<td>" + esc_(t.activeDays + " active days") + "</td></tr>" +
    "<tr><td><b>Unique Walk-ins</b></td>" + '<td class="num"><b>' + t.unique + "</b></td>" +
      "<td>distinct mobile numbers — repeat visitors counted once</td></tr>" +
    "<tr><td><b>Cases Created</b></td>" + '<td class="num"><b>' + t.cases + "</b></td>" +
      "<td>walk-ins that turned into a service case</td></tr>" +
    "<tr><td><b>Invoices Created</b></td>" + '<td class="num"><b>' + t.invoices + "</b></td>" +
      "<td>walk-ins that closed a sale</td></tr>" +
    "<tr><td><b>Remarks Updated</b></td>" + '<td class="num"><b>' + t.remarks + "</b></td>" +
      "<td>rows carrying a remark</td></tr>" +
    '<tr class="deadsum"><td><b>Dead Leads</b></td>' + '<td class="num dead"><b>' + t.deadLeads + "</b></td>" +
      "<td>no case, no invoice, no remark</td></tr>" +
    "</tbody></table></div>";

  return head;
}

/** WALK-IN page: the wide day-wise breakdown table (full page width). */
function renderWalkinDayTable_(wl) {
  if (!wl.available) return "";
  var t = wl.totals;

  function c0(n) { return n ? ('<td class="num">' + n + "</td>") : '<td class="num muted">—</td>'; }

  // Each DATE ROW gets its own soft hue so neighbouring dates are easy to tell
  // apart (pivot-style shading). Hue steps keep consecutive dates distinct.
  // Day rows alternate between just TWO standard navy tints (owner asked to
  // ditch the rotating rainbow: "simple 2 or 3 standard optimum colours").
  var TINTS = ["#1d2354", "#2a3164", "#232a5e"];
  var rows = wl.days.map(function (d, i) {
    var tint = TINTS[i % TINTS.length];
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
    return '<tr style="--tint:' + tint + '"' + (d.undated ? ' class="undated-row"' : "") + ">" +
      '<td class="daycell"><b>' + esc_(d.undated ? "No date" : d.dateKey) + "</b>" +
        (d.weekday ? ' <span class="muted">' + d.weekday + "</span>" : "") + "</td>" +
      '<td class="num"><b>' + d.walkins + "</b></td>" +
      '<td class="num uniq">' + d.unique + "</td>" +
      cells +
      '<td class="num">' + d.cases + "</td>" +
      '<td class="num">' + d.invoices + "</td>" +
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
    '<td class="num uniq"><b>' + t.unique + "</b></td>" +
    purposeFoot +
    '<td class="num">' + t.cases + "</td>" +
    '<td class="num">' + t.invoices + "</td>" +
    '<td class="num">' + t.remarks + "</td>" +
    '<td class="num dead">' + t.deadLeads + "</td>" +
    '<td class="num dead">' + t.deadRate + "%</td>" +
    "</tr>";

  var purpHeads = wl.purposeCols.map(function (name) { return "<th>" + esc_(name) + "</th>"; }).join("");
  var table =
    '<div class="scroll-x"><table class="ptable walkintable">' +
    "<thead>" +
    '<tr><th rowspan="2">Date</th><th rowspan="2">Walk-ins</th><th rowspan="2">Unique</th>' +
    '<th colspan="' + wl.purposeCols.length + '" class="grp">Visit purpose</th>' +
    '<th rowspan="2">Cases</th>' +
    '<th rowspan="2">Invoices</th>' +
    '<th rowspan="2">Remark</th><th rowspan="2">Dead leads</th><th rowspan="2">Dead leads %</th></tr>' +
    "<tr>" + purpHeads + "</tr>" +
    "</thead><tbody>" + rows + foot + "</tbody></table></div>";

  return table;
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

  // ---- club the jobs by CRM repair status (bucket-wise) -------------
  // Buckets are alphabetical; within a bucket the oldest job stays first so
  // the longest-waiting cases in each status sit at the top of their block.
  var grouped = jm.jobs.slice().sort(function (a, b) {
    var ga = (a.crmRepairStatus || "—").toLowerCase();
    var gb = (b.crmRepairStatus || "—").toLowerCase();
    if (ga < gb) return -1;
    if (ga > gb) return 1;
    return b.jobAge - a.jobAge;
  });

  // Column order fixed by the owner (10 columns):
  // Repair bucket · Aging · GSX repair status · Case ID · Name · Number ·
  // Model · Repair type · Communicated · Comm updated date.
  // Rows stay grouped under a highlighted band per repair bucket.
  var NCOLS = 10;
  var prevGroup = null;
  var rowParts = [];
  for (var gi = 0; gi < grouped.length; gi++) {
    var j = grouped[gi];
    var grp = j.crmRepairStatus || "—";
    if (grp !== prevGroup) {
      prevGroup = grp;
      var bucketCount = grouped.filter(function (x) {
        return (x.crmRepairStatus || "—") === grp;
      }).length;
      rowParts.push('<tr class="grphead"><td colspan="' + NCOLS + '">' + esc_(grp) +
        '<span class="grphead-count">' + bucketCount + " open case" + (bucketCount === 1 ? "" : "s") +
        "</span></td></tr>");
    }
    rowParts.push(buildRow_(j));
  }
  var rows = rowParts.join("");

  function buildRow_(j) {
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
    // Rule (a): Service Non-Repair + CRM ready-for-pickup → GSX = Completed.
    var gsxDisplay = j.gsxCompleted ? '<td><span class="badge completed">Completed</span></td>'
      : repairStatusCell(j.gsxStatus);
    // Repair-bucket cell: same value the rows are grouped by, with the
    // "Unclaimed Device" badge kept loud. The <td> is pre-built by
    // repairStatusCell — the grphead band above each block carries the
    // highlight, so the cell itself stays a plain status.
    var bucketCell = j.isUnclaimed
      ? '<td class="bucket-cell"><span class="badge-unclaimed">' + esc_(j.crmRepairStatus || "—") + "</span></td>"
      : '<td class="bucket-cell">' + esc_(j.crmRepairStatus || "—") + "</td>";
    var rowClses = [];
    if (j.isUnclaimed) rowClses.push("unclaimed-row");
    // Rule (b): physical repair never started and the case is 2+ days old.
    if (j.noRepairLate) rowClses.push("norepair-row");
    // Rule (c): Out Of Warranty (No Coverage) with no advance anywhere.
    if (j.oowNoAdvance) rowClses.push("oow-nopay-row");
    return "<tr" + (rowClses.length ? ' class="' + rowClses.join(" ") + '"' : "") + ">" +
      bucketCell +
      ageCell(j.jobAge) +
      gsxDisplay +
      '<td class="sticky-col"><b>' + esc_(j.caseId) + "</b>" +
        (j.isUnclaimed ? ' <span class="badge-unclaimed sm">⚠ Unclaimed</span>' : "") +
        (j.noRepairLate ? ' <span class="badge-norepair sm">⚠ No repair started</span>' : "") +
        (j.oowNoAdvance ? ' <span class="badge-oow sm">⚠ OOW · no advance</span>' : "") + "</td>" +
      "<td>" + esc_(j.name) + "</td>" +
      '<td class="num">' + esc_(j.mobile) + "</td>" +
      "<td>" + esc_(j.product) + "</td>" +
      "<td>" + esc_(j.repairType) + "</td>" +
      "<td>" + comm + "</td>" +
      commDate +
      "</tr>";
  }

  var table =
    '<div class="scroll-x jobwrap"><table class="ptable jobtable">' +
    '<colgroup>' +
    '<col style="width:10.5%"><col style="width:5.5%"><col style="width:12%">' +
    '<col style="width:11%"><col style="width:11%"><col style="width:9.5%">' +
    '<col style="width:14%"><col style="width:11%"><col style="width:7.5%">' +
    '<col style="width:8%">' +
    '</colgroup>' +
    '<thead><tr>' +
    "<th>Repair Bucket</th><th>Aging</th><th>GSX Repair Status</th>" +
    "<th>Case ID</th><th>Name</th><th>Number</th><th>Model</th>" +
    "<th>Repair Type</th><th>Communicated</th><th>Comm. Updated</th>" +
    "</tr></thead><tbody>" + rows + "</tbody></table></div>";

  // Resolved column mapping — printed under the section so the user can
  // verify each displayed column against their sheet headers at a glance.
  function srcLine(pair) {
    var h = jm.sources ? str_(jm.sources[pair[1]]) : "";
    return esc_(pair[0]) + " ← " + (h ? "<b>" + esc_(h) + "</b>" : '<span class="map-missing">not found</span>');
  }
  var mapLine = '<div class="map-line">' + [
    ["Repair Bucket","crmRepairStatus"],["Aging","jobAge"],
    ["GSX Repair Status","gsxStatus"],["Case ID","caseId"],["Number","mobile"],
    ["Model","product"],["Repair Type","repairDocType"],["Open filter","status"]]
    .map(srcLine).join(" · ") + "</div>";

  return head + mapLine + table;
}

/** Job management section (OPEN JOBS page). */
function renderJobsSection_(pd) {
  var jm = pd.jobs;
  return (
    '<section class="sec"><div class="sec-head">' +
    '<h2>🔧 Job Management — Open Jobs</h2>' +
    "</div>" + renderJobManagement_(jm) + "</section>"
  );
}

/** Whole performance tab content (legacy combined view, kept for compat). */
function renderPerformanceHtml_(pd) {
  var wl = pd.walkin;
  var pieces = [];
  var secTarget = renderTargetBlock_(pd.targetBlock, pd.repair);
  if (secTarget) pieces.push(secTarget);
  pieces.push(renderWalkinHtml_(pd), renderJobsSection_(pd));
  return pieces.join("");
}

function getDashboardCss_() {
  // =========================================================================
  // Visual design ported 1:1 from the previous "script page" the user shared:
  // light grey page, dark navy cards, Inter font stack, and these exact
  // palette tokens —
  //   achieved #63e6b7 · target #8fc7ff · balance #ffd477 · percent #dfb2ff
  //   muted #aebdd0 · white #f8fbff · line #3a5575 · danger #ff9da7
  // Only looks changed; no condition or data logic was touched.
  // =========================================================================
  return (
    ":root { --page:#eef2f6; --surface:#17263a; --surface-2:#20344d; --surface-3:#263e5b;" +
    " --line:#3a5575; --muted:#aebdd0; --white:#f8fbff; --achieved:#63e6b7; --target:#8fc7ff;" +
    " --balance:#ffd477; --percent:#dfb2ff; --accent:#45d9f5; --danger:#ff9da7; --completed:#63efbb; }" +
    "* { box-sizing:border-box; }" +
    // +2 points across the whole page: every size below is rem-based, so a
    // slightly larger root font scales all text up evenly.
    "html { font-size:18px; }" +
    "body { margin:0; min-height:100vh;" +
    " background:linear-gradient(180deg,#edf2f7 0%,#e7edf3 100%); color:#17263a;" +
    " font-family:Inter,ui-sans-serif,system-ui,-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif; }" +
    // ---- page frame ----
    ".wrap { max-width:1720px; margin:0 auto; padding:20px 22px 40px; }" +
    // ---- top bar (light, like the old page header) ----
    ".topbar { position:sticky; top:0; z-index:20; padding:14px 22px 12px;" +
    " background:rgba(238,242,246,.92); backdrop-filter:blur(6px); border-bottom:1px solid #d5dde6; }" +
    ".topbar-inner { max-width:1720px; margin:0 auto; display:flex; align-items:center;" +
    " justify-content:space-between; gap:20px; flex-wrap:wrap; }" +
    ".topbar-left { display:flex; flex-direction:column; gap:6px; min-width:0; }" +
    ".brand { color:#16263a; font-size:clamp(1.45rem,2.5vw,2rem); font-weight:850;" +
    " letter-spacing:-.035em; line-height:1.1; }" +
    ".topbar .meta { color:#64748b; font-size:.78rem; font-weight:600; }" +
    ".topbar .note { font-size:.74rem; min-height:15px; font-weight:600; }" +
    ".topbar .note.ok { color:#0d7d56; } .topbar .note.err { color:#c2393f; }" +
    // page slider: a wide pill bar, centered on the page under the title
    ".tabs-row { width:100%; display:flex; justify-content:center; margin-top:10px; }" +
    ".tabs { display:flex; background:#ffffff; border:1px solid #d5dde6; border-radius:999px; padding:5px;" +
    " gap:6px; width:min(760px,100%); box-shadow:0 3px 10px rgba(15,23,42,.06); }" +
    ".tab-btn { appearance:none; border:none; background:transparent; color:#475569; font-size:.92rem;" +
    " font-weight:800; padding:13px 18px; border-radius:999px; cursor:pointer; transition:all .16s ease;" +
    " flex:1; text-align:center; white-space:nowrap; }" +
    ".tab-btn:hover { color:#16263a; background:#eef2f6; }" +
    ".tab-btn.active { background:#172b43; color:#fff; box-shadow:0 3px 8px rgba(15,23,42,.18); }" +
    // ---- refresh button: ported exact ----
    ".refresh-btn { flex-shrink:0; border:0; border-radius:8px; padding:9px 14px; background:#172b43;" +
    " color:#ffffff; cursor:pointer; font-size:.76rem; font-weight:800; white-space:nowrap;" +
    " box-shadow:0 5px 14px rgba(15,23,42,.16); display:inline-flex; align-items:center; gap:8px; }" +
    ".refresh-btn:hover { background:#223b59; }" +
    ".refresh-btn:disabled { opacity:.6; cursor:wait; }" +
    ".refresh-btn .spin { display:none; width:12px; height:12px; border:2px solid rgba(255,255,255,.4);" +
    " border-top-color:#fff; border-radius:50%; animation:spin .7s linear infinite; }" +
    ".refresh-btn.loading .spin { display:inline-block; }" +
    "@keyframes spin { to { transform:rotate(360deg); } }" +
    // ---- section headings: the old .section-heading look (no coloured band) ----
    ".section { margin-bottom:26px; overflow:visible; }" +
    ".section h2 { margin:0 0 9px; padding:0; background:none; color:#182a40;" +
    " font-size:1rem; font-weight:850; letter-spacing:-.015em; text-transform:none; }" +
    ".section .body { padding:0; }" +
    // ---- cards: the old .kpi-card dark navy tile ----
    ".cards { display:grid; grid-template-columns:repeat(auto-fit,minmax(min(100%,225px),1fr)); gap:10px; margin-bottom:16px; }" +
    ".card { background:linear-gradient(145deg,rgba(28,43,82,.98),rgba(22,34,66,.98));" +
    " border:1px solid var(--line); border-radius:12px; box-shadow:0 8px 20px rgba(0,0,0,.18);" +
    " min-width:0; overflow:hidden; padding:13px; }" +
    ".card .label { color:var(--muted); display:block; font-size:.6rem; font-weight:800;" +
    " letter-spacing:.06em; line-height:1.15; text-transform:uppercase; white-space:nowrap;" +
    " overflow:hidden; text-overflow:ellipsis; }" +
    ".card .value { display:block; margin-top:4px; color:var(--white);" +
    " font-size:clamp(1rem,1.6vw,1.3rem); font-weight:850; letter-spacing:-.02em; line-height:1.1;" +
    " font-variant-numeric:tabular-nums; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }" +
    ".card .sub { display:block; margin-top:4px; color:var(--muted); font-size:.72rem; font-weight:500; }" +
    ".card .value.pos { color:var(--achieved); } .card .value.neg { color:var(--danger); }" +
    ".card-unclaimed { border-color:#eab308; background:linear-gradient(145deg,rgba(90,62,8,.95),rgba(46,32,6,.95)); }" +
    ".card-unclaimed .value, .card-unclaimed .value.neg { color:var(--balance); }" +
    // ---- Section 1 store tiles: 1:1 port of the old page's kpi-card pieces ----
    ".rgrid { display:grid; grid-template-columns:repeat(auto-fit,minmax(min(100%,225px),1fr)); gap:10px; }" +
    ".rtile { padding:11px; }" +
    ".rtile.done { background:linear-gradient(145deg,rgba(11,93,71,.98),rgba(9,50,54,.98)); border-color:#57f2bd; }" +
    ".rt-head { align-items:center; display:flex; gap:6px; justify-content:space-between; margin-bottom:8px; min-width:0; }" +
    ".rt-head h3 { color:#fff; font-size:.88rem; font-weight:850; letter-spacing:-.015em; margin:0;" +
    " min-width:0; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }" +
    ".status { border-radius:999px; flex-shrink:0; font-size:.55rem; font-weight:850;" +
    " padding:3px 6px; text-transform:uppercase; }" +
    ".status.on { background:rgba(97,239,187,.17); color:var(--achieved); }" +
    ".status.pending { background:rgba(255,210,122,.15); color:var(--balance); }" +
    ".mg { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:6px; width:100%; }" +
    ".metric { background:rgba(255,255,255,.055); border-radius:7px; min-width:0; padding:7px 6px; }" +
    ".metric span { color:var(--muted); display:block; font-size:.56rem; font-weight:800;" +
    " letter-spacing:.045em; line-height:1.1; text-transform:uppercase; white-space:nowrap; }" +
    ".metric strong { display:block; margin-top:3px; font-size:.82rem; font-weight:850;" +
    " letter-spacing:-.02em; line-height:1.15; font-variant-numeric:tabular-nums; white-space:nowrap; }" +
    ".metric.t strong { color:var(--target); } .metric.a strong { color:var(--achieved); }" +
    ".metric.p strong { color:var(--percent); } .metric.b strong { color:var(--balance); }" +
    ".dreq { background:rgba(69,220,255,.09); border-left:3px solid var(--accent); border-radius:7px;" +
    " margin-top:6px; padding:6px 8px; width:100%; }" +
    ".dreq span { color:var(--muted); display:block; font-size:.54rem; font-weight:800;" +
    " letter-spacing:.045em; line-height:1.1; text-transform:uppercase; }" +
    ".dreq strong { color:#e9fbff; display:block; font-size:.82rem; font-variant-numeric:tabular-nums;" +
    " line-height:1.15; margin-top:2px; white-space:nowrap; }" +
    ".rtile .bar { margin-top:7px; }" +
    ".rtile.done .bar-fill { background:#6dffc7; }" +
    ".sec-desc { color:#64748b; font-size:.72rem; margin:-3px 0 10px; }" +
    // ---- badges / status pills: old .status pills ----
    ".badge { display:inline-block; padding:3px 10px; border-radius:999px; font-weight:850;" +
    " font-size:.66rem; text-transform:uppercase; letter-spacing:.02em; }" +
    ".badge.good { background:rgba(97,239,187,.17); color:var(--achieved); }" +
    ".badge.info { background:rgba(143,199,255,.16); color:var(--target); }" +
    ".badge.warn { background:rgba(255,210,122,.15); color:var(--balance); }" +
    ".badge.bad { background:rgba(255,157,167,.15); color:var(--danger); }" +
    // ---- progress bars: 3D loading-slider look — dark inset track,
    // glossy gradient fill with a glow halo and a travelling light streak ----
    ".bar { height:11px; border-radius:999px; background:#12141f; overflow:hidden; min-width:70px;" +
    " box-shadow:inset 0 2px 5px rgba(0,0,0,.7), inset 0 -1px 2px rgba(255,255,255,.08), 0 1px 0 rgba(255,255,255,.08); }" +
    ".bar-fill { position:relative; height:100%; border-radius:inherit; display:block; overflow:hidden;" +
    " background:var(--barc, linear-gradient(90deg,#2f8dff,#3ce3b4));" +
    " box-shadow:0 0 12px var(--barg, rgba(88,190,255,.55)), inset 0 -2px 3px rgba(0,0,0,.28); }" +
    // top gloss: the "3D" dome highlight
    ".bar-fill::before { content:''; position:absolute; left:1px; right:1px; top:1px; height:44%;" +
    " border-radius:999px 999px 40% 40%; background:linear-gradient(180deg,rgba(255,255,255,.85),rgba(255,255,255,.04)); }" +
    // travelling shine streak
    ".bar-fill::after { content:''; position:absolute; top:0; bottom:0; width:42%; left:-45%;" +
    " background:linear-gradient(90deg,transparent,rgba(255,255,255,.5),transparent);" +
    " transform:skewX(-18deg); animation:barShine 2.6s ease-in-out infinite; }" +
    "@keyframes barShine { to { left:135%; } }" +
    ".bar-fill.good { --barc:linear-gradient(90deg,#0fbf7f,#6dffc7); --barg:rgba(109,255,199,.55); }" +
    ".bar-fill.info { --barc:linear-gradient(90deg,#3b7dff,#8fc7ff); --barg:rgba(113,160,255,.55); }" +
    ".bar-fill.warn { --barc:linear-gradient(90deg,#e0a02c,#ffd477); --barg:rgba(255,212,119,.55); }" +
    ".bar-fill.bad { --barc:linear-gradient(90deg,#e05a68,#ff9da7); --barg:rgba(255,120,135,.6); }" +
    // ---- generic text tones ----
    ".pos { color:var(--achieved); font-weight:700; } .neg { color:var(--danger); font-weight:700; }" +
    ".muted { color:var(--muted); }" +
    "table .muted { color:#94a2d2; }" +
    // ---- tables: old .date-report dark surface ----
    // ==== pricing-chart palette for every table: glow-blue header over deep navy rows ====
    // Deep variant: on the owner's request the Invoice Valuation + Weekly
    // tables get a header a few shades darker — still clearly separate from
    // the navy body cells.
    "table.hdr-deep thead th, table.hdr-deep thead th.grp, table.hdr-deep thead .sticky-col {" +
    " background:linear-gradient(180deg,#25306e 0%,#1c2450 60%,#141b3e 100%) !important;" +
    " color:#dfe6ff; border-bottom:2px solid #5468c9; }" +
    "table.hdr-deep tr.total-row td { background:linear-gradient(90deg,#232b5f,#1a2048) !important; }" +
    "table { width:100%; border-collapse:separate; border-spacing:0; font-size:.8rem;" +
    " background:#161b42; border:1px solid rgba(105,132,255,.34); border-radius:11px; overflow:hidden;" +
    " box-shadow:0 10px 26px rgba(8,12,48,.38), 0 2px 10px rgba(76,111,255,.14); }" +
    "th { text-align:right; background:linear-gradient(180deg,#5b7bff 0%,#4360e0 55%,#3350c2 100%); color:#ffffff;" +
    " font-size:.68rem; font-weight:800; text-shadow:0 1px 2px rgba(12,20,70,.65);" +
    " letter-spacing:.05em; text-transform:uppercase; padding:9px 10px;" +
    " border-bottom:2px solid #8eaaff; box-shadow:inset 0 1px 0 rgba(255,255,255,.25), 0 2px 8px rgba(76,111,255,.30); }" +
    "th:first-child, td:first-child { text-align:left; }" +
    "td { padding:8px 10px; text-align:right; color:#e9edff; font-size:.8rem; font-weight:400;" +
    " font-variant-numeric:tabular-nums; background:#1d2354;" +
    " border-bottom:1px solid rgba(105,132,255,.16); }" +
    "tbody tr:last-child td { border-bottom:none; }" +
    "tbody tr:hover td { background:#262e68; }" +
    "tr.holiday td { background:rgba(255,212,119,.10) !important; }" +
    ".status-chip { display:inline-block; padding:2px 10px; border-radius:999px; font-size:12px; font-weight:700; }" +
    ".status-chip.ok { background:rgba(99,230,183,.28); color:#0f7a52; }" +
    ".status-chip.warn { background:rgba(255,212,119,.30); color:#8a5b00; }" +
    ".status-chip.muted-chip { background:rgba(120,130,150,.16); color:#5b6675; }" +
    ".src-hint { font-family:ui-monospace,Consolas,monospace; font-size:12px; background:rgba(120,130,150,.14);" +
    " padding:1px 6px; border-radius:6px; }" +
    ".walksum { max-width:640px; }" +
    ".walksum td:not(.num) { text-align:left; }" +
    ".walksum tr.deadsum td { background:rgba(255,157,167,.10); }" +
    ".targetblock th, .targetblock td { text-align:left; }" +
    ".repair-strip { display:flex; flex-wrap:wrap; gap:18px; align-items:center; margin-top:12px;" +
    " padding:12px 16px; border:1px solid #d5dde6; border-radius:10px; background:#f6f9fd; }" +
    ".rs-item { display:flex; flex-direction:column; gap:2px; }" +
    ".rs-label { font-size:12px; color:#5b6675; font-weight:600; }" +
    ".rs-value { font-size:20px; font-weight:800; color:#17263a; }" +
    ".csat-flex { display:flex; gap:16px; align-items:flex-start; flex-wrap:wrap; }" +
    ".csat-flex .csat-left { flex:0 0 auto; max-width:100%; }" +
    ".csat-flex table.csat-table { width:auto; min-width:0; }" +
    ".csat-flex .csat-right { flex:0 0 230px; width:230px; min-width:230px; max-width:230px; display:flex; flex-direction:column; gap:10px; }" +
    ".csat-right .csat-chip { flex:0 0 auto; width:100%; }" +
    ".csat-chips { display:flex; flex-wrap:wrap; gap:10px; margin-bottom:12px; }" +
    ".csat-chip { flex:1 1 150px; background:#fff; border:1px solid #d5dde6; border-radius:10px;" +
    " padding:10px 14px; display:flex; flex-direction:column; gap:2px; box-shadow:0 2px 6px rgba(15,23,42,.06); }" +
    // Task 25: navy chip → light text so labels stay readable.
    ".csat-chip .cc-label { font-size:11.5px; font-weight:700; color:#9fb4e8; text-transform:uppercase; letter-spacing:.05em; }" +
    ".csat-chip .cc-value { font-size:1.5rem; font-weight:800; color:#f4f8ff; }" +
    ".csat-chip .cc-sub { font-size:11.5px; color:#93a5d4; }" +
    ".csat-chip.pos { border-left:4px solid #2fb97f; } .csat-chip.pos .cc-value { color:#4fe0a8; }" +
    ".csat-chip.neg { border-left:4px solid #e25567; } .csat-chip.neg .cc-value { color:#ff8fa1; }" +
    ".csat-chip.warn { border-left:4px solid #e0a12d; }" +
    ".csat-chip.info { border-left:4px solid #2f5aa8; }" +
    // C-SAT: columns size to their DATA, long headers wrap instead of
    // stretching the table (owner's "fit the data, not the header").
    "table.csat-table { width:auto; min-width:320px; max-width:100%; }" +
    "table.csat-table th { white-space:normal; overflow-wrap:break-word; vertical-align:bottom;" +
    " line-height:1.25; min-width:62px; max-width:150px; text-align:center; }" +
    "table.csat-table th:first-child { text-align:left; min-width:150px; }" +
    "table.csat-table td { width:1%; white-space:nowrap; }" +
    "table.csat-table td:first-child { text-align:left; white-space:normal; width:auto; }" +
    "table.csat-table td.num { font-variant-numeric:tabular-nums; text-align:center; }" +
    ".csat-table .totrow td { background:linear-gradient(90deg,#3f5ce0,#2f49b5) !important; color:#ffffff; font-weight:800; border-top:1px solid rgba(142,170,255,.6); }" +
    ".csat-pos { color:#52e9ac; font-weight:700; } .csat-neg { color:#ff8f9d; font-weight:700; } .csat-mid { color:#ffd477; font-weight:600; }" +
    "table.vmatrix { width:100%; border-collapse:collapse; background:#161b42;" +
    " border:1px solid rgba(105,132,255,.34); border-radius:11px; overflow:hidden;" +
    " box-shadow:0 10px 26px rgba(8,12,48,.38), 0 2px 10px rgba(76,111,255,.14); }" +
    "table.vmatrix th, table.vmatrix td { padding:8px 12px; border-bottom:1px solid rgba(105,132,255,.16); font-size:.85rem; }" +
    "table.vmatrix thead th { background:linear-gradient(180deg,#5b7bff 0%,#4360e0 55%,#3350c2 100%); color:#ffffff; text-align:center; font-size:.76rem;" +
    " font-weight:800; text-shadow:0 1px 2px rgba(12,20,70,.65);" +
    " text-transform:uppercase; letter-spacing:.04em; position:sticky; top:0; z-index:2; border-bottom:2px solid #8eaaff; }" +
    "table.vmatrix td { background:#1d2354; color:#e9edff; }" +
    "table.vmatrix td.num, table.vmatrix th.num { text-align:right; font-variant-numeric:tabular-nums; }" +
    "table.vmatrix td.muted { color:#8f9ecf; }" +
    "table.vmatrix .sticky-col { position:sticky; left:0; z-index:1; background:#232a63; color:#e9edff; text-align:left; font-weight:700; }" +
    "table.vmatrix thead .sticky-col { z-index:3; background:#3350c2; color:#ffffff; }" +
    "table.vmatrix tbody tr:hover td { background:#262e68; }" +
    "table.vmatrix tbody tr:hover .sticky-col { background:#2b336f; }" +
    ".vmtotal { background:rgba(84,116,255,.16) !important; color:#ffffff; font-weight:800; }" +
    "tr.vmtotrow td { background:linear-gradient(90deg,#3f5ce0,#2f49b5) !important; color:#fff; font-weight:800; border-top:1px solid rgba(142,170,255,.6); }" +
    "tr.vmavgrow td { background:rgba(84,116,255,.12); color:#a9b8ff; font-weight:700; font-style:italic; }" +
    ".vmavg { color:#a9b8ff; }" +
    "tr.total-row td { background:linear-gradient(90deg,#3f5ce0,#2f49b5) !important; color:#ffffff; font-weight:900; }" +
    // ---- salesperson cards ----
    ".person-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(min(100%,400px),1fr)); gap:10px; }" +
    ".person-card { background:rgba(12,29,53,.94); border:1px solid var(--line); border-radius:11px;" +
    " min-width:0; padding:12px; }" +
    ".person-card h3 { color:#ffffff; font-size:.92rem; font-weight:850; margin:0 0 9px; min-width:0; }" +
    ".row-line { display:flex; justify-content:space-between; align-items:center; gap:8px; margin:7px 0 3px;" +
    " color:var(--white); font-size:.82rem; font-weight:600; font-variant-numeric:tabular-nums; }" +
    ".row-line .muted { color:var(--muted); }" +
    ".scroll-x { overflow-x:auto; border-radius:11px; }" +
    // Performance-page tile row: one navy card per bucket, fits any width.
    ".tile-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(min(100%,190px),1fr)); gap:10px; margin-bottom:14px; }" +
    ".tile { background:linear-gradient(145deg,rgba(28,43,82,.98),rgba(22,34,66,.98));" +
    " border:1px solid var(--line); border-radius:12px; box-shadow:0 8px 20px rgba(0,0,0,.18);" +
    " min-width:0; overflow:hidden; padding:13px; }" +
    ".tile .label { color:var(--muted); display:block; font-size:.6rem; font-weight:800;" +
    " letter-spacing:.06em; line-height:1.15; text-transform:uppercase; white-space:nowrap;" +
    " overflow:hidden; text-overflow:ellipsis; }" +
    ".tile .value { display:block; margin-top:4px; color:var(--white); font-size:1.35rem; font-weight:800; }" +
    ".tile .sub { display:block; margin:4px 0 6px; color:var(--muted); font-size:.72rem; font-weight:500; }" +
    // Pivoted matrices keep numbers right-aligned and columns tight when space allows.
    "table.matrix td, table.matrix th { text-align:right; }" +
    "table.matrix td.sticky-col, table.matrix th.sticky-col, table.matrix td:first-child { text-align:left; }" +
    "table.ptable { width:100%; border-collapse:collapse; }" +
    "table.ptable th.num, table.ptable td.num { text-align:right; }" +
    ".foot { text-align:center; color:#64748b; font-size:.74rem; padding:14px 0 26px; font-weight:500; }" +
    // ---- incentive gate notes inside person cards ----
    ".gate { margin-top:10px; padding:9px 11px; border-radius:8px; font-size:.78rem; line-height:1.45; }" +
    ".gate.met { background:rgba(97,239,187,.10); border:1px solid rgba(97,239,187,.35); color:#a5f3d0; }" +
    ".gate.blocked { background:rgba(255,157,167,.10); border:1px solid rgba(255,157,167,.35); color:#ffc7cd; }" +
    ".incentive-line { display:flex; justify-content:space-between; font-size:.78rem; margin-top:6px;" +
    " color:var(--white); font-variant-numeric:tabular-nums; padding-bottom:5px;" +
    " border-bottom:1px solid rgba(58,85,117,.45); }" +
    ".incentive-line .muted { color:var(--muted); }" +
    // ---- date-wise per-person grouping (wide business table) ----
    "table.detail td, table.detail th { font-size:.8rem; }" +
    "tr.date-head td { background:#2b336f !important; color:#ffffff; font-weight:750;" +
    " border-top:1px solid rgba(105,132,255,.35); }" +
    "td.person-cell { font-weight:700; border-left:4px solid transparent; }" +
    ".sp-0 { border-left-color:#8fc7ff !important; }" +
    ".sp-1 { border-left-color:#fbbf24 !important; }" +
    ".sp-2 { border-left-color:#63e6b7 !important; }" +
    ".sp-3 { border-left-color:#dfb2ff !important; }" +
    ".sp-4 { border-left-color:#ff9da7 !important; }" +
    "tr.person-band-0 td { background:rgba(143,199,255,.055); }" +
    "tr.person-band-1 td { background:rgba(251,191,36,.055); }" +
    "tr.person-band-2 td { background:rgba(99,230,183,.055); }" +
    "tr.person-band-3 td { background:rgba(223,178,255,.055); }" +
    "tr.person-band-4 td { background:rgba(255,157,167,.055); }" +
    // Oval achievement pills — exact old-page look: an oval around the
    // achieved number, mint/salmon fill with dark readable text inside.
    // The condition is unchanged (met the daily target → green, else red).
    ".apill { display:inline-flex; align-items:center; justify-content:center; min-width:72px;" +
    " max-width:100%; padding:4px 11px; border-radius:999px; font-size:.78rem; font-weight:800;" +
    " line-height:1; white-space:nowrap; font-variant-numeric:tabular-nums; box-sizing:border-box; }" +
    ".apill.ok { background:#63e6b7; color:#12352d; box-shadow:0 1px 3px rgba(18,53,45,.16); }" +
    ".apill.no { background:#ff9da7; color:#5f2028; box-shadow:0 1px 3px rgba(95,32,40,.16); }" +
    // ---- section 7 date-by-date matrix ----
    "table.matrix { font-size:.8rem; border-collapse:separate; border-spacing:0; }" +
    "table.matrix th, table.matrix td { white-space:nowrap; padding:8px 10px; }" +
    "table.matrix tbody td { border-bottom:1px solid rgba(58,85,117,.45); }" +
    "table.matrix thead th { position:sticky; top:0; z-index:3; }" +
    "table.matrix th.grp { text-align:center; border-left:1px solid var(--line); }" +
    "table.matrix .sticky-col { position:sticky; left:0; z-index:2; background:#1d2354;" +
    " text-align:left; border-right:1px solid rgba(105,132,255,.35); }" +
    "table.matrix thead .sticky-col { z-index:4; background:#3350c2; }" +
    "table.matrix td .tgt { display:block; font-size:.66rem; font-weight:600; color:#94a2d2; margin-top:2px; }" +
    // subtle tinted column groups, one per salesperson
    // Person bands are STRONGLY distinct solids (owner asked for an instant
    // "whose column is this" read on the date-wise matrix).
    // Subha's band: soft periwinkle instead of the loud royal blue (Task 23).
    "td.cell-0 { background:rgba(146,161,214,.12); } th.grp-0 { background:linear-gradient(180deg,#8a9bd4,#6f7fb8); }" +
    "td.cell-1 { background:rgba(196,138,22,.11); } th.grp-1 { background:linear-gradient(180deg,#b8860b,#9c7309); }" +
    "td.cell-2 { background:rgba(15,157,143,.09); } th.grp-2 { background:linear-gradient(180deg,#0f9d8f,#0c8377); }" +
    "td.cell-3 { background:rgba(122,63,212,.10); } th.grp-3 { background:linear-gradient(180deg,#7a3fd4,#6134ad); }" +
    "td.cell-4 { background:rgba(192,74,104,.10); } th.grp-4 { background:linear-gradient(180deg,#c04a68,#a03c55); }" +
    "table.matrix tr.holiday .sticky-col { background:rgba(255,212,119,.16) !important; }" +
    "table.matrix tr.total-row td { background:linear-gradient(90deg,#3f5ce0,#2f49b5) !important; border-top:1px solid rgba(142,170,255,.6); }" +
    // ---- AR attachment sub-panel inside section 6 (light panel like old summary items) ----
    ".subpanel { margin-top:14px; padding:13px 15px; border-radius:10px;" +
    " border:1px solid #d5dde6; background:rgba(255,255,255,.92); box-shadow:0 3px 10px rgba(15,23,42,.05); }" +
    ".subpanel-head { font-size:.68rem; font-weight:850; text-transform:uppercase; letter-spacing:.08em;" +
    " color:#475569; margin:0 0 12px; }" +
    ".subpanel-head .muted { color:#718096; }" +".subpanel .muted { color:#718096; }" +
    ".legend-row { display:flex; flex-wrap:wrap; gap:10px; margin:0 0 12px; }" +
    ".legend-card { font-size:.76rem; padding:8px 12px; border-radius:8px; border:1px solid #d5dde6;" +
    " line-height:1.5; background:rgba(255,255,255,.92); color:#1f2937; text-align:center;" +
    " box-shadow:0 3px 10px rgba(15,23,42,.04); font-variant-numeric:tabular-nums; }" +
    ".legend-card b { color:#111827; }" +
    ".legend-card .muted { color:#475569; }" +
    // ---- section 7 collapse / expand ----
    ".toggle-btn { display:inline-flex; align-items:center; gap:9px; margin:0 0 12px;" +
    " background:#172b43; color:#fff; border:none; border-radius:8px; padding:9px 14px;" +
    " font-size:.76rem; font-weight:800; cursor:pointer; box-shadow:0 5px 14px rgba(15,23,42,.16); }" +
    ".toggle-btn:hover { background:#223b59; }" +
    ".toggle-btn .chev { display:inline-block; font-size:10px; transition:transform .18s ease; }" +
    ".toggle-btn.open .chev { transform:rotate(180deg); }" +
    ".toggle-btn .count { background:rgba(255,255,255,.18); border-radius:999px;" +
    " padding:2px 10px; font-size:.68rem; font-weight:800; }" +
    ".toggle-btn.open .count { display:none; }" +
    "table.matrix.collapsed tbody tr.extra-day { display:none; }" +
    // ---- process page panels: light cards like the old summary strip ----
    ".sec { background:rgba(255,255,255,.92); border:1px solid #d5dde6; border-radius:9px;" +
    " padding:14px 15px 12px; margin-bottom:16px; box-shadow:0 3px 10px rgba(15,23,42,.05); }" +
    ".sec-head h2 { margin:0 0 4px; color:#182a40; font-size:1rem; font-weight:850; letter-spacing:-.015em; }" +
    ".sec-sub { margin:0 0 12px; color:#64748b; font-size:.78rem; font-weight:500; }" +
    ".sec .muted { color:#64748b; }" +

    // dark tables inside the process page (jobs / tokens), matching the old date-report
    ".ptable { border-collapse:separate; border-spacing:0; width:100%; font-size:.8rem;" +
    " background:#161b42; border:1px solid rgba(105,132,255,.34); border-radius:11px; overflow:hidden;" +
    " box-shadow:0 10px 26px rgba(8,12,48,.38), 0 2px 10px rgba(76,111,255,.14); }" +
    ".ptable th, .ptable td { text-align:left; padding:8px 11px; border-bottom:1px solid rgba(105,132,255,.16);" +
    " vertical-align:top; white-space:nowrap; font-weight:400; color:#e9edff; }" +
    ".ptable td { background:#1d2354; }" +
    ".ptable thead th { background:linear-gradient(180deg,#5b7bff 0%,#4360e0 55%,#3350c2 100%); color:#ffffff; font-size:.68rem;" +
    " font-weight:800; text-shadow:0 1px 2px rgba(12,20,70,.65);" +
    " text-transform:uppercase; letter-spacing:.05em; position:sticky; top:0; border-bottom:2px solid #8eaaff;" +
    " box-shadow:inset 0 1px 0 rgba(255,255,255,.25); }" +
    ".ptable tbody tr:hover td { background:#262e68; }" +
    ".ptable td.num { text-align:right; font-variant-numeric:tabular-nums; }" +
    ".ptable .daycell { min-width:110px; }" +
    ".jobtable th, .jobtable td { font-size:.78rem; }" +
    ".jobtable .sticky-col { position:sticky; left:0; background:#1d2354; }" +
    ".jobtable thead .sticky-col { background:#3350c2; }" +
    ".chip { display:inline-block; background:rgba(69,217,245,.12); color:#aee8f7; border-radius:999px;" +
    " padding:3px 9px; font-size:.68rem; margin:1px 2px 1px 0; font-weight:600; }" +
    ".chip b { color:#ffffff; }" +
    // walk-in / comm status colours on the dark tables
    ".comm-yes { color:var(--achieved); font-weight:700; }" +
    ".comm-no { color:#fbbf24; font-weight:700; }" +
    ".comm-pos { color:#12352d; background:#63e6b7 !important; font-weight:700; }" +
    ".comm-neg { color:#5f2028; background:#ff9da7 !important; font-weight:700; }" +
    // unclaimed devices: keeps the amber identity, tuned for the dark table
    ".unclaimed-row td { background:rgba(255,212,119,.10) !important; }" +
    ".unclaimed-row:hover td { background:rgba(255,212,119,.16) !important; }" +
    ".unclaimed-row .sticky-col { box-shadow:inset 4px 0 0 #eab308; }" +
    ".badge-unclaimed { display:inline-block; background:rgba(255,210,122,.18); color:var(--balance); font-weight:850;" +
    " font-size:.68rem; padding:3px 10px; border-radius:999px; white-space:nowrap; text-transform:uppercase; }" +
    ".badge-unclaimed.sm { font-size:.62rem; padding:2px 7px; background:var(--balance); color:#451a03; }" +
    ".undated-row td { background:rgba(223,178,255,.10); }" +
    // walk-in day rows each carry their own --tint; cells fall back to the navy phase
    ".walkintable tbody td { background-color: var(--tint, #1d2354); }" +
    ".walkintable tbody tr:hover td { background-color: var(--tint, #1d2354); box-shadow: inset 0 0 0 999px rgba(84,116,255,.14); }" +
    ".walkintable .totrow td { background:linear-gradient(90deg,#3f5ce0,#2f49b5) !important; border-top:1px solid rgba(142,170,255,.6); color:#ffffff; font-weight:800; }" +
    ".walkintable td.dead { color:#ff9daa; font-weight:700; }" +
    ".walkintable td.uniq { color:#7edbff; font-weight:700; }" +
    ".walkintable thead th { text-align:center; }" +
    ".walkintable thead th.grp { background:linear-gradient(180deg,#5b7bff,#3350c2); text-align:center; }" +
    ".walkintable th:first-child, .walkintable td.daycell { text-align:left; }" +
    ".tokentable .totrow td { background:linear-gradient(90deg,#3f5ce0,#2f49b5) !important; color:#ffffff; font-weight:900; border-top:1px solid rgba(142,170,255,.6); }" +
    ".age-warn { color:var(--balance); font-weight:600; }" +
    ".age-crit { color:var(--danger); font-weight:700; }" +
    ".subnote { background:rgba(255,210,122,.14); border:1px solid rgba(255,210,122,.45); border-radius:8px; padding:9px 12px;" +
    " font-size:.78rem; color:#6b5410; margin-bottom:12px; font-weight:500; }" +
    ".empty-perf { background:rgba(255,255,255,.92); border:1px dashed #cbd5e1; border-radius:9px; padding:22px;" +
    " text-align:center; font-size:.82rem; color:#64748b; font-weight:500; }" +
    ".grphead td { background:linear-gradient(90deg,#546fff,#7c5cff) !important; color:#ffffff; font-weight:850; font-size:.7rem;" +
    " letter-spacing:.07em; text-transform:uppercase; padding:8px 12px; border-bottom:2px solid #9db4ff;" +
    " box-shadow:0 3px 10px rgba(84,111,255,.35); }" +
    ".grphead-count { float:right; color:var(--target); font-weight:800; text-transform:none; letter-spacing:0; }" +
    ".norepair-row td { background:rgba(255,157,167,.12) !important; }" +
    ".norepair-row .sticky-col { box-shadow:inset 4px 0 0 var(--danger); }" +
    ".badge-norepair { display:inline-block; background:rgba(255,157,167,.20); color:var(--danger); font-weight:850;" +
    " font-size:.62rem; padding:2px 7px; border-radius:999px; white-space:nowrap; }" +
    ".oow-nopay-row td { background:rgba(223,178,255,.13) !important; }" +
    ".oow-nopay-row .sticky-col { box-shadow:inset 4px 0 0 var(--percent); }" +
    ".badge-oow { display:inline-block; background:rgba(223,178,255,.20); color:var(--percent); font-weight:850;" +
    " font-size:.62rem; padding:2px 7px; border-radius:999px; white-space:nowrap; }" +
    ".badge.completed { background:rgba(97,239,187,.17); color:var(--achieved); }" +
    ".map-line { font-size:.68rem; color:#64748b; margin:0 0 10px; font-weight:500; line-height:1.6; }" +
    ".map-line b { color:#17263a; font-weight:750; }" +
    ".map-missing { color:#c2393f; font-weight:800; }" +
    // ---- wider tab bar (six pages + index) ----
    ".tabs.tabs-wide { width:min(1180px,100%); flex-wrap:wrap; }" +
    ".tabs.tabs-wide .tab-btn { font-size:.82rem; padding:11px 13px; }" +
    // ---- KPI page: table sizes to CONTENT, incentive fills the right ----
    ".kpi-pair { display:flex; gap:16px; align-items:flex-start; flex-wrap:wrap; }" +
    ".kpi-pair .kp-left { flex:0 0 auto; min-width:0; }" +
    ".kpi-pair .kp-left table { width:auto; min-width:min(100%,760px); }" +
    ".kpi-pair .kp-right { flex:1 1 520px; min-width:340px; }" +
    ".inc-cat { margin-bottom:14px; }" +
    ".inc-cat-title { font-size:.78rem; font-weight:850; color:#33415c; letter-spacing:.04em;" +
    " text-transform:uppercase; margin:0 0 8px; }" +
    ".inc-cat table { width:100%; }" +
    // ---- walk-in page: both halves share the width EQUALLY ----
    ".walkin-pair { display:flex; gap:16px; align-items:stretch; margin-top:2px; flex-wrap:wrap; }" +
    ".walkin-pair .wp-left { flex:1 1 50%; min-width:320px; }" +
    ".walkin-pair .wp-right { flex:1 1 50%; min-width:320px; display:flex; flex-direction:column; }" +
    ".wp-head { font-size:.92rem; font-weight:850; color:#182a40; margin:0 0 4px; letter-spacing:-.015em; }" +
    // ---- placeholder modules (Decline repair / Requote-winback) ----
    ".ph-tile { display:flex; gap:12px; align-items:center; border:2px dashed rgba(105,132,255,.55);" +
    " border-radius:12px; padding:16px; background:rgba(84,116,255,.08); }" +
    ".ph-tile .ph-icon { font-size:1.4rem; }" +
    ".ph-tile .ph-text { font-size:.82rem; color:#475569; line-height:1.5; }" +
    // ---- repair-bucket cell tone on the job table ----
    ".jobtable td.bucket-cell { font-weight:800; color:#bcd0ff; }" +
    // ============ INDEX LANDING — LIQUID GLASS tiles ============
    // Frosted translucent panes laid over a soft gradient backdrop (Apple
    // "liquid glass" look): blurred, edge-lit, gently refracted. Groups run
    // left→right as columns; each group's sub-tiles stack one below one.
    ".index-hero h2 { margin:0 0 6px; font-size:1.3rem; font-weight:850; letter-spacing:-.03em; color:#16263a; }" +
    ".index-sub { margin:0 0 4px; color:#5b6b84; font-size:.86rem; font-weight:500; }" +
    ".tile-board { display:grid; grid-template-columns:repeat(auto-fit,minmax(min(100%,250px),1fr));" +
    " gap:18px; margin:14px 0 8px; padding:18px; border-radius:22px;" +
    " background:linear-gradient(135deg,#c9d6f8 0%,#dfe7fb 42%,#cde4f4 70%,#d9d2f4 100%);" +
    " box-shadow:inset 0 1px 0 rgba(255,255,255,.6), 0 14px 34px rgba(30,45,90,.18); }" +
    ".tile-group { display:flex; flex-direction:column; gap:10px; }" +
    ".tile-group-head { margin:0; font-size:.92rem; font-weight:850; letter-spacing:.02em; color:#22336a; text-transform:none; }" +
    ".tile-group-link { background:none; border:none; padding:0; margin:0; font:inherit; color:#22336a;" +
    " cursor:pointer; letter-spacing:-.02em; }" +
    ".tile-group-link:hover { color:#3f5ce0; }" +
    ".tile-col { display:flex; flex-direction:column; gap:12px; }" +
    // --- the tile: liquid glass — clear, blurred, with a luminous rim ---
    ".glass-tile { position:relative; display:flex; flex-direction:column; gap:5px; text-align:left;" +
    " min-height:158px; padding:14px 14px 36px; border-radius:18px; cursor:pointer; color:#22336a;" +
    " background:linear-gradient(135deg,rgba(255,255,255,.62) 0%,rgba(226,237,255,.34) 55%,rgba(255,255,255,.5) 100%);" +
    " border:1px solid rgba(255,255,255,.75);" +
    " box-shadow:0 10px 24px rgba(40,60,120,.16), inset 0 1px 0 rgba(255,255,255,.95)," +
    " inset 0 -10px 18px rgba(120,150,255,.10), inset 0 0 0 1px rgba(255,255,255,.3);" +
    " backdrop-filter:blur(14px) saturate(1.25); -webkit-backdrop-filter:blur(14px) saturate(1.25);" +
    " overflow:hidden; opacity:.9;" +
    " transition:opacity .16s ease, transform .16s ease, box-shadow .16s ease; }" +
    // refracted top-light rim
    ".glass-tile::before { content:''; position:absolute; inset:0 0 auto 0; height:46%; border-radius:18px 18px 60% 60%;" +
    " background:linear-gradient(180deg,rgba(255,255,255,.7),rgba(255,255,255,0)); pointer-events:none; }" +
    // golden glitter film — hidden until hover
    ".tile-sparkle { position:absolute; inset:0; border-radius:inherit; pointer-events:none; opacity:0;" +
    " background:" +
      "radial-gradient(2px 2px at 18% 30%,rgba(255,221,120,.95),rgba(255,221,120,0) 60%)," +
      "radial-gradient(2px 2px at 38% 70%,rgba(255,196,64,.95),rgba(255,196,64,0) 60%)," +
      "radial-gradient(2.5px 2.5px at 58% 22%,rgba(255,232,150,.95),rgba(255,232,150,0) 60%)," +
      "radial-gradient(2px 2px at 74% 55%,rgba(255,203,80,.95),rgba(255,203,80,0) 60%)," +
      "radial-gradient(2.5px 2.5px at 86% 28%,rgba(255,240,180,.95),rgba(255,240,180,0) 60%)," +
      "radial-gradient(2px 2px at 64% 82%,rgba(255,221,120,.95),rgba(255,221,120,0) 60%)," +
      "radial-gradient(1.8px 1.8px at 26% 58%,rgba(255,232,150,.95),rgba(255,232,150,0) 60%)," +
      "radial-gradient(22% 40% at 50% 0%,rgba(255,230,150,.35),rgba(255,230,150,0) 70%);" +
    " background-repeat:no-repeat; }" +
    ".glass-tile:hover, .glass-tile:focus-visible { opacity:1; transform:translateY(-3px); outline:none;" +
    " box-shadow:0 16px 34px rgba(40,60,120,.28), inset 0 1px 0 rgba(255,255,255,1)," +
    " inset 0 -12px 20px rgba(255,214,120,.18), 0 0 26px rgba(255,205,90,.35); }" +
    // on hover the glitter brightens and twinkles
    ".glass-tile:hover .tile-sparkle, .glass-tile:focus-visible .tile-sparkle { opacity:1;" +
    " animation:tileTwinkle 1.1s ease-in-out infinite alternate; }" +
    "@keyframes tileTwinkle { 0% { opacity:.35; } 100% { opacity:1; } }" +
    ".tile-icon { font-size:2.5rem; line-height:1.1; position:relative; z-index:1; }" +
    ".tile-title { font-size:1rem; font-weight:850; color:#1d2f63; letter-spacing:-.005em; position:relative; z-index:1; }" +
    ".tile-blurb { font-size:.76rem; color:#45558a; line-height:1.45; position:relative; z-index:1; }" +
    // hidden numbers that fade in when the tile brightens on hover
    ".tile-data { margin-top:2px; opacity:0; max-height:0; overflow:hidden; position:relative;" +
    " transition:opacity .18s ease, max-height .22s ease; }" +
    ".glass-tile:hover .tile-data, .glass-tile:focus-visible .tile-data { opacity:1; max-height:180px; }" +
    ".tile-stat { display:flex; justify-content:space-between; gap:8px; font-size:.7rem; color:#3a4a7d;" +
    " padding:3px 0; border-top:1px solid rgba(120,140,220,.3); }" +
    ".tile-stat b { color:#1d2f63; font-weight:800; }" +
    ".tile-soon { margin-top:6px; font-size:.76rem; font-weight:800; color:#8a5b00; }" +
    ".tile-soon-sub { font-size:.68rem; color:#4d5d8a; }" +
    ".tile-go { position:absolute; left:14px; bottom:11px; font-size:.7rem; font-weight:800; color:#3f5ce0; }" +
    ".glass-tile:hover .tile-go { color:#1d2f63; }" +
    // --- pie under each group title ---
    ".gpie-wrap { background:rgba(255,255,255,.55); border:1px solid rgba(255,255,255,.7); border-radius:14px;" +
    " padding:10px; box-shadow:0 6px 16px rgba(40,60,120,.12), inset 0 1px 0 rgba(255,255,255,.8); }" +
    ".gpie-title { font-size:.64rem; font-weight:800; color:#5a6a99; text-transform:uppercase; letter-spacing:.06em; margin-bottom:6px; }" +
    ".gpie-body { display:flex; align-items:center; gap:10px; }" +
    "svg.gpie { width:88px; height:88px; flex-shrink:0; }" +
    ".gpie-leg { display:flex; flex-direction:column; gap:3px; min-width:0; }" +
    ".gpie-leg-item { display:flex; align-items:center; gap:6px; font-size:.66rem; color:#3a4a7d; white-space:nowrap;" +
    " overflow:hidden; text-overflow:ellipsis; }" +
    ".gpie-dot { width:8px; height:8px; border-radius:50%; flex-shrink:0; box-shadow:0 0 4px rgba(0,0,0,.2); }" +
    ".gpie-leg-val { color:#1d2f63; margin-left:auto; padding-left:8px; }" +
    ".gpie-empty { font-size:.66rem; color:#8a96b8; padding:14px 4px; }" +
    // --- latest-records strip at the bottom of the index ---
    ".fresh-line { margin-top:16px; padding:12px 16px; border-radius:14px;" +
    " background:linear-gradient(135deg,rgba(255,255,255,.6),rgba(230,238,255,.45));" +
    " border:1px solid rgba(255,255,255,.7); box-shadow:0 6px 16px rgba(40,60,120,.10); }" +
    ".fresh-line-head { font-size:.66rem; font-weight:850; color:#5a6a99; text-transform:uppercase; letter-spacing:.08em; margin-bottom:6px; }" +
    ".fresh-row { display:flex; flex-wrap:wrap; gap:8px; }" +
    ".fresh-chip { display:inline-flex; align-items:center; gap:6px; background:rgba(255,255,255,.75);" +
    " border:1px solid rgba(150,170,255,.4); border-radius:999px; padding:4px 12px; font-size:.68rem; }" +
    ".fresh-chip i { font-style:normal; color:#5a6a99; font-weight:700; }" +
    ".fresh-chip b { color:#1d2f63; font-weight:800; }" +
    // ============ TASK 21 — dark liquid-glass index + glass material ============
    // Bokeh-hearts video is the always-on background layer of the index pane.
    "#pageIndex { position:relative; background:radial-gradient(1200px 420px at 50% -80px,#1a2350 0%,#0a0f2c 60%,#06091c 100%);" +
    " border-radius:22px; overflow:hidden; }" +
    "#pageIndex > *:not(.bgnd-video) { position:relative; z-index:1; }" +
    ".bgnd-video { position:absolute; z-index:0; inset:0; width:100%; height:100%; object-fit:cover;" +
    " opacity:.8; pointer-events:none; }" +
    "#pageIndex .index-hero h2 { color:#f5f8ff; text-shadow:0 2px 10px rgba(3,6,24,.7); }" +
    "#pageIndex .tile-board { background:rgba(9,13,34,.55); border:1px solid rgba(255,255,255,.14);" +
    " backdrop-filter:blur(10px) saturate(1.2); -webkit-backdrop-filter:blur(10px) saturate(1.2);" +
    " box-shadow:0 16px 40px rgba(2,5,24,.55), inset 0 1px 0 rgba(255,255,255,.14); }" +
    "#pageIndex .tile-group-head, #pageIndex .tile-group-link { color:#e6ecff; }" +
    "#pageIndex .tile-group-link:hover { color:#ffd477; }" +
    // golden-glitter video layer: invisible until the tile is hovered
    ".tile-glitter-vid { position:absolute; z-index:0; inset:0; width:100%; height:100%; object-fit:cover;" +
    " border-radius:inherit; opacity:0; transition:opacity .2s ease; pointer-events:none; }" +
    ".glass-tile:hover .tile-glitter-vid, .glass-tile:focus-visible .tile-glitter-vid { opacity:.62; }" +
    ".glass-tile .tile-data, .glass-tile .tile-go { z-index:1; }" +
    // index-side group charts & freshness strip on the dark backdrop
    "#pageIndex .gpie-wrap { background:linear-gradient(145deg,rgba(28,43,82,.98),rgba(22,34,66,.98)); }" +
    "#pageIndex .fresh-line { background:rgba(12,17,44,.62); border-color:rgba(255,255,255,.18);" +
    " backdrop-filter:blur(10px); }" +
    "#pageIndex .fresh-line-head { color:#aebcff; }" +
    "#pageIndex .fresh-chip { background:linear-gradient(145deg,rgba(28,43,82,.98),rgba(22,34,66,.98)); border-color:rgba(170,190,255,.28); }" +
    "#pageIndex .fresh-chip i { color:#aebcff; } #pageIndex .fresh-chip b { color:#ffffff; }" +
    // ---- horizontal bar chart under each group title ----
    ".gbar-group { display:flex; flex-direction:column; gap:5px; }" +
    ".gbar-row { display:flex; align-items:center; gap:7px; }" +
    ".gbar-lbl { width:74px; flex-shrink:0; font-size:.6rem; font-weight:800; color:#44548c; letter-spacing:.02em;" +
    " white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }" +
    ".gbar-track { flex:1; height:8px; border-radius:999px; background:rgba(90,110,190,.18); overflow:hidden;" +
    " box-shadow:inset 0 1px 2px rgba(20,30,80,.3); }" +
    ".gbar-fill { display:block; height:100%; border-radius:inherit;" +
    " box-shadow:0 0 8px rgba(80,120,255,.45), inset 0 -1px 2px rgba(0,0,0,.25), inset 0 1px 1px rgba(255,255,255,.5); }" +
    ".gbar-val { width:30px; flex-shrink:0; text-align:right; font-size:.64rem; font-weight:850; color:#1d2f63; }" +
    // ---- liquid-glass material for every table & tile (owner's dribbble ref) ----
    "table { background:linear-gradient(145deg,rgba(28,43,82,.98),rgba(22,34,66,.98)); border:1px solid rgba(170,190,255,.32);" +
    " backdrop-filter:blur(12px) saturate(1.2); -webkit-backdrop-filter:blur(12px) saturate(1.2);" +
    " box-shadow:0 12px 30px rgba(8,12,48,.45), inset 0 1px 0 rgba(255,255,255,.10), 0 2px 10px rgba(76,111,255,.14); }" +
    "td { background:rgba(29,35,84,.78); }" +
    "tbody tr:hover td { background:rgba(38,46,104,.85); }" +
    ".card, .person-card, .tile {" +
    " backdrop-filter:blur(12px) saturate(1.2); -webkit-backdrop-filter:blur(12px) saturate(1.2);" +
    " box-shadow:0 12px 28px rgba(8,14,44,.32), inset 0 1px 0 rgba(255,255,255,.14); }" +
    ".card { border:1px solid rgba(170,190,255,.28); }" +
    ".person-card { border:1px solid rgba(170,190,255,.30); }" +
    ".csat-chip { backdrop-filter:blur(10px); }" +
    // ---- AMC & Bundle tables: side by side, splitting the width ----
    ".inc-pair { display:flex; gap:14px; align-items:stretch; flex-wrap:wrap; }" +
    ".inc-pair .inc-cat { flex:1 1 46%; min-width:340px; margin-bottom:0; }" +
    ".inc-pair table { width:100%; }" +
    // ---- repair type performance + validation matrix: side by side ----
    ".duo-pair { display:flex; gap:16px; align-items:stretch; flex-wrap:wrap; }" +
    ".duo-pair > .section { flex:1 1 47%; min-width:460px; margin-bottom:0; border:1px solid #d5dde6;" +
    " border-radius:9px; background:#ffffff; padding:10px 12px 14px; }" +
    // ---- job table: fixed layout compressed to the page width ----
    ".jobwrap { overflow-x:auto; }" +
        ".jobtable { table-layout:fixed; width:100%; min-width:0; }" +
    ".jobtable th, .jobtable td { font-size:.72rem; padding:6px 7px; white-space:nowrap;" +
    " overflow:hidden; text-overflow:ellipsis; font-weight:500; }" +
    ".jobtable th { font-weight:800; }" +

    // ---- next-slab chase line on the salesperson cards ----
    ".next-slab { color:#8fc7ff; font-weight:700; }" +
    // ---- live clock card, top-middle of the page ----
    // Task 29: the AC40 clock gif must be FULLY visible — cover-cropping cut
    // off its edges; use contain inside a fixed-height, non-shrinking card.
    ".topbar-clock { position:relative; overflow:hidden; border-radius:14px; width:300px; height:86px; flex:0 0 auto;" +
    " padding:10px 18px; box-sizing:border-box;" +
    " display:flex; align-items:center; justify-content:center; border:1px solid rgba(140,160,230,.5);" +
    " background:rgba(4,6,18,.85);" +
    " box-shadow:0 8px 22px rgba(6,10,34,.4), inset 0 1px 0 rgba(255,255,255,.25); isolation:isolate; }" +
    ".clock-gif { position:absolute; inset:0; width:100%; height:100%; object-fit:contain; z-index:-2; }" +
    ".clock-dim { position:absolute; inset:0; z-index:-1; background:rgba(6,10,30,.52); }" +
    // Task 25/26: digits glow red-LED to match the AC40 BRG clock texture.
    ".clock-text { font-family:ui-monospace,SFMono-Regular,Consolas,monospace; font-size:1.08rem; font-weight:800;" +
    " letter-spacing:.05em; color:#ff5c66; text-shadow:0 0 6px rgba(255,70,80,.9), 0 0 14px rgba(255,50,60,.6), 0 0 30px rgba(255,45,55,.35);" +
    " white-space:nowrap; font-variant-numeric:tabular-nums; }" +
    // ============ TASK 22 — one black-glitter backdrop for ALL pages ============
    "body { background:#070a1c; color:#dbe4f6; }" +
    "#globalBgVideo { position:fixed; inset:0; z-index:-1; width:100%; height:100%; object-fit:cover; }" +
    ".global-bg-gif { position:fixed; inset:0; z-index:-1; width:100%; height:100%; object-fit:cover;" +
    " opacity:.6; pointer-events:none; }" +
    // index pane no longer needs its own dark shell — the video is global now
    "#pageIndex { background:none; border-radius:0; overflow:visible; }" +
    // sticky topbar + tabs as dark liquid glass on the glitter
    ".topbar { background:rgba(8,12,32,.78); backdrop-filter:blur(14px) saturate(1.25);" +
    " -webkit-backdrop-filter:blur(14px) saturate(1.25); border-bottom:1px solid rgba(170,190,255,.22); }" +
    ".brand { color:#f5f8ff; text-shadow:0 2px 8px rgba(3,6,24,.7); }" +
    ".topbar .meta { color:#aebcff; }" +
    ".topbar .note.ok { color:#6ee7b7; } .topbar .note.err { color:#ffb4c0; }" +
    ".tabs { background:rgba(12,17,42,.72); border-color:rgba(170,190,255,.25);" +
    " backdrop-filter:blur(12px); -webkit-backdrop-filter:blur(12px); box-shadow:0 8px 22px rgba(4,8,26,.45); }" +
    ".tab-btn { color:#c3cfee; }" +
    ".tab-btn.active { background:linear-gradient(180deg,#ffffff,#dfe8ff); color:#1b2a52; box-shadow:0 4px 12px rgba(4,8,26,.5); }" +
    // dark liquid-glass shells for every section on every page
    ".sec { background:rgba(10,15,40,.66); border-color:rgba(170,190,255,.20);" +
    " backdrop-filter:blur(12px) saturate(1.15); -webkit-backdrop-filter:blur(12px) saturate(1.15);" +
    " box-shadow:0 14px 34px rgba(4,8,26,.5), inset 0 1px 0 rgba(255,255,255,.10); }" +
    ".duo-pair > .section { background:rgba(10,15,40,.66); border:1px solid rgba(170,190,255,.20);" +
    " backdrop-filter:blur(12px); -webkit-backdrop-filter:blur(12px); }" +
    ".section h2, .sec h2, .section-title { color:#f2f6ff; }" +
    ".sec-sub, .section-sub, .sec-desc { color:#a9b6d8; }" +
    ".src-hint { background:rgba(120,130,150,.22); color:#cfe0ff; }" +
    ".repair-strip { background:rgba(10,15,40,.66); border-color:rgba(170,190,255,.20); }" +
    ".rs-label { color:#a9b6d8; } .rs-value { color:#ffffff; }" +
    ".ph-tile { border-color:rgba(140,160,230,.45); background:linear-gradient(145deg,rgba(28,43,82,.98),rgba(22,34,66,.98)); }" +
    ".ph-text { color:#c7d2ee; }" +
    ".inc-cat-title { color:#cfe0ff; }" +
    ".wp-head { color:#f2f6ff; }" +
    ".map-line { color:#94a3bd; } .map-line b { color:#e6edff; }" +
    ".legend-card { border-color:rgba(170,190,255,.28); }" +
    ".legend-card b { color:#f2f6ff; } .legend-card .muted { color:#a9b6d8; }" +
    ".subnote { color:#ffe9b3; }" +
    // walk-in analysis + token creation locked side by side, sharing the width
    ".walkin-pair { display:grid; grid-template-columns:minmax(0,1fr) minmax(0,1fr); gap:16px; align-items:stretch; }" +
    ".walkin-pair .wp-left, .walkin-pair .wp-right { min-width:0; }" +
    ".walkin-pair .wp-right { background:rgba(28,43,82,.55); border:1px solid rgba(170,190,255,.18);" +
    " border-radius:12px; padding:12px; backdrop-filter:blur(10px); -webkit-backdrop-filter:blur(10px); }" +
    // job table: entirely left-aligned
    ".jobtable th, .jobtable td { text-align:left; }" +
    // C-SAT chips move BELOW the table, one horizontal row
    ".csat-flex { flex-direction:column; align-items:stretch; }" +
    ".csat-flex .csat-right { flex:0 0 auto; width:100%; min-width:0; max-width:100%;" +
    " flex-direction:row; flex-wrap:wrap; gap:10px; margin-top:12px; }" +
    ".csat-right .csat-chip { flex:1 1 190px; width:auto; max-width:280px; }" +
    // ---- index tiles: lighter glass, brighter luminous fonts ----
    ".glass-tile { background:linear-gradient(145deg,rgba(28,43,82,.98),rgba(22,34,66,.98));" +
    " border:1px solid rgba(255,255,255,.55); color:#f4f8ff; opacity:1; isolation:isolate;" +
    " box-shadow:0 10px 28px rgba(3,7,26,.5), inset 0 1px 0 rgba(255,255,255,.6), inset 0 -10px 18px rgba(255,255,255,.06); }" +
    ".tile-title { color:#ffffff; text-shadow:0 2px 8px rgba(4,8,26,.85); }" +
    ".tile-blurb { color:#e6eeff; text-shadow:0 1px 5px rgba(4,8,26,.8); }" +
    ".tile-stat { color:#e8f0ff; border-top-color:rgba(255,255,255,.40); }" +
    ".tile-stat b { color:#ffe08a; text-shadow:0 1px 5px rgba(4,8,26,.8); }" +
    ".tile-soon { color:#ffd477; text-shadow:0 1px 5px rgba(4,8,26,.8); }" +
    ".tile-soon-sub { color:#e6eeff; }" +
    ".tile-go { color:#ffd477; } .glass-tile:hover .tile-go { color:#ffe08a; }" +
    ".tile-icon { filter:drop-shadow(0 2px 4px rgba(4,8,26,.75)); }" +
    // golden-glitter video hugs the tile outline exactly (same radius, no bleed)
    ".tile-glitter-vid { inset:0; border-radius:18px; }" +
    ".glass-tile:hover .tile-glitter-vid, .glass-tile:focus-visible .tile-glitter-vid { opacity:.8; }" +
    ".glass-tile:hover, .glass-tile:focus-visible {" +
    " box-shadow:0 16px 36px rgba(3,7,26,.6), inset 0 1px 0 rgba(255,255,255,.7), 0 0 28px rgba(255,205,90,.35); }" +
    // ============ TASK 23 — gif backdrop, tile fit, tables & chatbot ============
    // Golden-glitter hover video hugs each tile's outline exactly and always
    // sits BEHIND the text — it can never block what the tile says.
    ".glass-tile { overflow:hidden; }" +
    ".tile-glitter-vid { z-index:0; }" +
    ".glass-tile .tile-icon, .glass-tile .tile-title, .glass-tile .tile-blurb," +
    " .glass-tile .tile-data, .glass-tile .tile-soon, .glass-tile .tile-soon-sub," +
    " .glass-tile .tile-go { position:relative; z-index:2; }" +
    ".glass-tile .tile-sparkle { z-index:1; }" +
    ".glass-tile:hover .tile-glitter-vid, .glass-tile:focus-visible .tile-glitter-vid {" +
    " opacity:.55; filter:brightness(.85) saturate(1.3); }" +
    // KPI table: expand to the entire page width, hugging the left edge.
    ".kpi-pair { flex-direction:column; align-items:stretch; }" +
    ".kpi-pair .kp-left { flex:0 0 auto; width:100%; }" +
    ".kpi-pair .kp-left table { width:100%; min-width:0; }" +
    ".kpi-pair .kp-right { flex:0 0 auto; width:100%; }" +
    // Walk-in / token pair: both sides top-aligned on one line; day-wise
    // table breaks out below and spans the full page width.
    ".wp-daytable { margin-top:16px; width:100%; }" +
    ".wp-daytable .scroll-x { width:100%; }" +
    ".wp-daytable .walkintable { width:100%; }" +
    // C-SAT chips: soften the glare — slightly grey-shaded instead of pure white.
    ".csat-chip { background:linear-gradient(145deg,rgba(28,43,82,.98),rgba(22,34,66,.98)); border-color:rgba(170,190,255,.32); box-shadow:0 2px 8px rgba(10,20,50,.18); }" +
    // Job management: force every cell left (beats .ptable td.num), and bucket
    // bands sweep from sky blue down to deep navy.
    ".jobtable th, .jobtable td { text-align:left !important; }" +
    ".grphead td { background:linear-gradient(90deg,#4fb3f5 0%,#2b8ae0 34%,#175fb0 68%,#0b3672 100%) !important;" +
    " border-bottom-color:#8cc9ff !important; box-shadow:0 3px 10px rgba(11,54,114,.5); }" +
    // ---- floating sheet-search chat bot (every page, bottom-right) ----
    ".chat-widget { position:fixed; right:18px; bottom:18px; z-index:90; display:flex;" +
    " flex-direction:column; align-items:flex-end; gap:10px; font-family:inherit; }" +
    ".chat-fab { display:flex; align-items:center; gap:8px; border:none; cursor:pointer; color:#fff;" +
    " background:linear-gradient(135deg,#4fb3f5,#0b3672); border-radius:999px; padding:11px 18px;" +
    " font-size:.82rem; font-weight:800; letter-spacing:.02em;" +
    " box-shadow:0 10px 26px rgba(4,10,34,.6), inset 0 1px 0 rgba(255,255,255,.4); }" +
    ".chat-fab:hover { transform:translateY(-2px); box-shadow:0 14px 32px rgba(4,10,34,.7); }" +
    ".chat-fab-icon { font-size:1.05rem; }" +
    ".chat-panel { display:none; width:358px; max-width:calc(100vw - 30px); height:430px;" +
    " max-height:calc(100vh - 110px); flex-direction:column; overflow:hidden; border-radius:16px;" +
    " background:rgba(10,15,40,.95); border:1px solid rgba(170,190,255,.3);" +
    " backdrop-filter:blur(14px); -webkit-backdrop-filter:blur(14px);" +
    " box-shadow:0 18px 46px rgba(2,6,26,.75); }" +
    ".chat-panel.open { display:flex; }" +
    ".chat-head { display:flex; align-items:center; gap:8px; padding:10px 12px;" +
    " border-bottom:1px solid rgba(170,190,255,.22); color:#f2f6ff; font-size:.85rem; }" +
    ".chat-head .chat-sub { font-size:.68rem; color:#9fb4e8; font-weight:500; }" +
    ".chat-close { margin-left:auto; background:none; border:none; color:#9fb4e8; font-size:1.15rem;" +
    " cursor:pointer; line-height:1; padding:2px 4px; } .chat-close:hover { color:#fff; }" +
    ".chat-log { flex:1; overflow-y:auto; padding:12px; display:flex; flex-direction:column; gap:8px; }" +
    ".chat-msg { max-width:92%; padding:8px 11px; border-radius:12px; font-size:.78rem; line-height:1.5; }" +
    ".chat-msg.user { align-self:flex-end; color:#fff; border-radius:12px 12px 4px 12px;" +
    " background:linear-gradient(135deg,#2b8ae0,#175fb0); }" +
    ".chat-msg.bot { align-self:flex-start; color:#e8eefc; border-radius:12px 12px 12px 4px;" +
    " background:rgba(255,255,255,.09); border:1px solid rgba(170,190,255,.16); }" +
    ".chat-msg.bot b { color:#ffd477; } .chat-msg.bot .muted { color:#93a5d4; }" +
    ".chat-title { font-weight:800; color:#cfe0ff; margin-bottom:6px; }" +
    ".chat-table { width:100%; border-collapse:collapse; margin:6px 0 2px; font-size:.7rem; }" +
    ".chat-table th { text-align:left; color:#9fb4e8; font-weight:800; padding:3px 5px;" +
    " border-bottom:1px solid rgba(170,190,255,.28); }" +
    ".chat-table td { padding:3px 5px; border-bottom:1px solid rgba(170,190,255,.12); vertical-align:top; }" +
    ".chat-table .num { text-align:right; font-variant-numeric:tabular-nums; }" +
    ".chat-dots i { display:inline-block; width:5px; height:5px; margin-right:3px; border-radius:50%;" +
    " background:#9fb4e8; animation:chatBlink 1s infinite; }" +
    ".chat-dots i:nth-child(2){animation-delay:.2s;} .chat-dots i:nth-child(3){animation-delay:.4s;}" +
    "@keyframes chatBlink { 0%,80%,100%{opacity:.25;} 40%{opacity:1;} }" +
    ".chat-inputrow { display:flex; gap:8px; padding:10px 12px; border-top:1px solid rgba(170,190,255,.22); }" +
    ".chat-input { flex:1; border:1px solid rgba(170,190,255,.3); border-radius:999px; background:rgba(255,255,255,.08);" +
    " color:#fff; padding:8px 14px; font-size:.8rem; outline:none; }" +
    ".chat-input::placeholder { color:#8fa3d4; } .chat-input:focus { border-color:#4fb3f5; }" +
    ".chat-send { border:none; cursor:pointer; border-radius:999px; padding:8px 14px; color:#fff;" +
    " background:linear-gradient(135deg,#4fb3f5,#0b3672); font-size:.85rem; font-weight:800; }" +
    ".chat-send.chat-wa { background:linear-gradient(135deg,#25d366,#128c4b); margin-left:6px; }" +
    "@media (max-width:640px){ .brand{font-size:1.1rem;} .card .value{font-size:1rem;} }"
  );
}

function renderStoreCards_(data) {
  var s = data.store;
  function card(label, value, sub) {
    return '<div class="card"><div class="label">' + esc_(label) + '</div><div class="value">' +
      esc_(value) + '</div><div class="sub">' + esc_(sub || "") + "</div></div>";
  }
  var gr = data.kpi.googleReview;
  // Google review tile — REVERSED emphasis: the gained review count is the
  // big bold number; the current and previous counts sit small beneath it.
  var diffCls = gr.difference > 0 ? "pos" : gr.difference < 0 ? "neg" : "muted";
  var grTile =
    '<div class="card"><div class="label">Google Reviews</div>' +
    '<div class="value ' + diffCls + '">' + (gr.difference >= 0 ? "+" : "") + gr.difference + "</div>" +
    '<div class="sub">' + gr.currentMonth + " this month &middot; previous " + gr.previousMonth + "</div></div>";

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

/**
 * Section 1 — Store Performance tiles.
 *
 * Rendered exactly like the old script page: one dark navy tile per revenue
 * type, each with a status pill (In Progress / Completed), a 2x2 metric grid
 * (Target · Achievement · Achievement % · Balance), a cyan Per-Day
 * Requirement strip and a progress bar. The per-day requirement reuses the
 * existing bucket requiredPerDay value (balance ÷ working days left,
 * today counted) — only the presentation changed, not the conditions.
 */
function renderStoreTable_(data) {
  var s = data.store;
  // Signed money like the old page: a covered target shows the excess with a
  // minus sign (−₹1,40,068) instead of a clamped zero.
  function sMoney(n) { return n < 0 ? "\u2212" + money_(-n) : money_(n); }
  function tile(title, b) {
    var pct = b.target > 0 ? (100 * b.achieved / b.target) : 0;
    var complete = b.target > 0 && pct >= 100;
    var signedBalance = b.target - b.achieved;
    var dayLabel = complete ? "Target Completed / Excess" : "Per Day Requirement";
    var dayValue = complete ? sMoney(signedBalance) : money_(b.requiredPerDay);
    var width = Math.max(0, Math.min(100, pct));
    return '<div class="card rtile' + (complete ? " done" : "") + '">' +
      '<div class="rt-head"><h3>' + esc_(title) + "</h3>" +
      '<span class="status ' + (complete ? "on" : "pending") + '">' +
      (complete ? "Completed" : "In Progress") + "</span></div>" +
      '<div class="mg">' +
      '<div class="metric t"><span>Target</span><strong>' + money_(b.target) + "</strong></div>" +
      '<div class="metric a"><span>Achievement</span><strong>' + money_(b.achieved) + "</strong></div>" +
      '<div class="metric p"><span>Achievement %</span><strong>' + pct.toFixed(1) + "%</strong></div>" +
      '<div class="metric b"><span>Balance</span><strong>' + sMoney(signedBalance) + "</strong></div>" +
      "</div>" +
      '<div class="dreq"><span>' + esc_(dayLabel) + "</span><strong>" + dayValue + "</strong></div>" +
      '<div class="bar"><i class="bar-fill" style="width:' + width.toFixed(1) + '%"></i></div>' +
      "</div>";
  }
  return (
    '<section class="section" id="sec-store"><h2>🎯 Store Target — Performance</h2><div class="body">' +
    '<div class="sec-desc">Target vs Achievement &middot; Run Rate</div>' +
    '<div class="rgrid">' +
    tile("Service Revenue", s.service) +
    tile("Accessory Revenue", s.accessory) +
    tile("Repair Revenue", s.repair) +
    tile("Total Revenue", s.total) +
    tile("Gross Profit", s.grossProfit) +
    "</div></div></section>"
  );
}

function renderWorkingDaysCard_(data) {
  var s = data.store;
  return (
    '<section class="section" id="sec-workdays"><h2>📅 Working Days (Current Month)</h2><div class="body">' +
    '<div class="sec-desc">Definition: days the store is open this month — total, already consumed, and how many selling days are left. Holiday dates come from CONFIG.</div>' +
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
  // still needed in (brackets), the incentive that type has earned, and —
  // from the slab table at target!A18:D25 — the extra revenue that jumps
  // this person into the NEXT (more rewarding) slab.
  function incentiveLine_(label, b, nextSlab) {
    var balTxt = b.balance > 0
      ? ' <span class="muted">(' + money_(b.balance) + " to go)</span>"
      : ' <span class="pos">(target met)</span>';
    var slabLine = nextSlab
      ? '<div class="incentive-line"><span class="muted">Next slab (' + esc_(nextSlab.pct) + '% &rarr; ' +
        esc_(nextSlab.label) + ')</span><span class="next-slab">' + money_(nextSlab.need) +
        ' more &rarr; earns ' + money_(nextSlab.amount) + "</span></div>"
      : '<div class="incentive-line"><span class="muted">Next slab</span>' +
        '<span class="pos">🏆 top slab reached</span></div>';
    return (
      '<div class="row-line"><span><b>' + esc_(label) + "</b></span>" +
      "<span>" + money_(b.achieved) + ' <span class="muted">/ ' + money_(b.target) + "</span> " +
      pctBadgeHtml_(b.pct) + "</span></div>" +
      barHtml_(b.pct) +
      '<div class="incentive-line"><span class="muted">Balance to target' + "</span><span>" + balTxt + "</span></div>" +
      slabLine +
      '<div class="incentive-line"><span class="muted">Incentive earned</span><span class="' +
      (b.incentive > 0 ? "pos" : "muted") + '">' + money_(b.incentive) + "</span></div>"
    );
  }

  if (!data.salespeople.length) {
    return '<section class="section" id="sec-person"><h2>👤 Sales Person Performance</h2><div class="body">' +
      '<div class="empty-state">No salespersons on the roster — add names to <b>target!C1:J1</b> (empty cells are skipped). ' +
      "People named in the CRM but not on the roster count only toward the store totals.</div></div></section>";
  }
  var cards = data.salespeople.map(function (p) {
    var gate = p.repairGate;

    // Repair: the number YOU personally won, shown plainly. No store-target
    // detail is printed on the card (owner's call) — the 80%-of-store-repair
    // eligibility is reported by the shared gate banner beneath the grid.
    var repairBlock =
      '<div class="row-line" style="margin-top:12px"><span><b>Repair</b> ' +
      '<span class="muted">(no incentive — store unlock gate)</span></span>' +
      '<span><b class="rr-val">' + money_(p.repair.achieved) + '</b> ' +
      '<span class="muted">won by you</span></span></div>';

    var gateBlock = gate.met
      ? '<div class="gate met">✅ <b>Incentive unlocked for everyone</b> — the store crossed ' +
        gate.requiredPct + '% of its repair target (store at ' + gate.achievedPct.toFixed(1) + '%). ' +
        'You receive what you earned: <b>' + money_(p.payableIncentive) + '</b>.</div>'
      : '<div class="gate blocked">🔒 <b>Incentive on hold for everyone</b> — the store has reached ' +
        gate.achievedPct.toFixed(1) + '% of its repair target; eligibility opens at ' + gate.requiredPct +
        '%. Your earned amount so far: <b>' + money_(p.earnedIncentive) + '</b>.</div>';

    return (
      '<div class="person-card"><h3>' + esc_(p.name) + "</h3>" +
      '<div class="row-line"><span>Total Revenue</span><span><b>' + money_(p.total.achieved) + "</b>" +
      ' <span class="muted">/ ' + money_(p.total.target) + "</span> " + pctBadgeHtml_(p.total.pct) + "</span></div>" +
      barHtml_(p.total.pct) +
      '<div class="row-line"><span>Invoices</span><span>' + p.invoiceCount + "</span></div>" +
      '<div class="row-line"><span><b>Payable Incentive</b></span><span class="' +
      (p.payableIncentive > 0 ? "pos" : "neg") + '"><b>' + money_(p.payableIncentive) + "</b></span></div>" +
      '<hr style="border:none;border-top:1px solid #e6ebf3;margin:10px 0">' +
      incentiveLine_("Service", p.service, p.nextSvcSlab) +
      '<div style="height:10px"></div>' +
      incentiveLine_("Accessory", p.accessory, p.nextAccSlab) +
      repairBlock +
      gateBlock +
      "</div>"
    );
  }).join("");

  return (
    '<section class="section" id="sec-person"><h2>👤 Sales Person Performance</h2><div class="body">' +
    '<div class="person-grid">' + cards + "</div>" +
    '<p class="muted" style="font-size:12.5px;margin:14px 0 0">' +
    "Only <b>Service</b> and <b>Accessory</b> revenue earn incentive; each also shows the balance that jumps you to the next slab (slab levels from <b>target!A18:D25</b>). " +
    "<b>Repair</b> revenue pays nothing on its own — it is one shared store gate: once the store&rsquo;s combined repair achievement passes <b>" +
    CONFIG.REPAIR_INCENTIVE_GATE_PCT + "%</b> of the store repair target, <b>every salesperson is eligible to receive the incentive they earned</b>.</p>" +
    "</div></section>"
  );
}

function renderIncentiveTable_(data) {
  var bundleByKey = {};
  data.bundle.forEach(function (b) { bundleByKey[b.key] = b; });

  // ---- Category 1: AMC ----
  var amcTot = { cnt: 0, elig: 0, accCnt: 0, accElig: 0, inc: 0 };
  var amcRows = data.amc.map(function (a) {
    amcTot.cnt += a.amcCount; amcTot.elig += a.amcEligible;
    amcTot.accCnt += a.accessoryCount; amcTot.accElig += a.accessoryEligible;
    amcTot.inc += a.projectedIncentive;
    return "<tr><td>" + esc_(a.name) + "</td><td>" + a.amcCount + "</td><td>" + money_(a.amcEligible) +
      "</td><td>" + a.accessoryCount + "</td><td>" + money_(a.accessoryEligible) + "</td><td>" +
      (a.amcNextSlab ? a.amcRemaining + " more" : "top slab") + '</td><td class="pos">' + money_(a.projectedIncentive) +
      "</td></tr>";
  }).join("");
  var amcFoot = '<tr class="total-row"><td>All salespersons</td><td>' + amcTot.cnt + "</td><td>" +
    money_(amcTot.elig) + "</td><td>" + amcTot.accCnt + "</td><td>" + money_(amcTot.accElig) +
    "<td></td>" + '<td class="pos">' + money_(amcTot.inc) + "</td></tr>";

  // ---- Category 2: Bundle (DISC scheme) ----
  var bunTot = { cnt: 0, inc: 0 };
  var bunRows = data.amc.map(function (a) {
    var b = bundleByKey[a.key] || { count: 0, projectedIncentive: 0 };
    bunTot.cnt += b.count; bunTot.inc += b.projectedIncentive;
    return "<tr><td>" + esc_(a.name) + "</td><td>" + b.count + '</td><td class="pos">' +
      money_(b.projectedIncentive) + "</td></tr>";
  }).join("");
  var bunFoot = '<tr class="total-row"><td>All salespersons</td><td>' + bunTot.cnt +
    '</td><td class="pos">' + money_(bunTot.inc) + "</td></tr>";

  return (
    '<section class="section" id="sec-incentive"><h2>💰 Sales Person Incentive — AMC &amp; Bundle</h2><div class="body">' +
    '<div class="inc-pair">' +
    '<div class="inc-cat"><div class="inc-cat-title">Category 1 · AMC</div>' +
    '<div class="scroll-x"><table><thead><tr><th>Salesperson</th><th>AMC Count</th><th>AMC Eligible</th>' +
    "<th>Acc-AMC Count</th><th>Acc-AMC Eligible</th><th>AMC To Next</th><th>AMC Incentive</th></tr></thead><tbody>" +
    amcRows + amcFoot + "</tbody></table></div></div>" +
    '<div class="inc-cat"><div class="inc-cat-title">Category 2 · Bundle (DISC scheme)</div>' +
    '<div class="scroll-x"><table><thead><tr><th>Salesperson</th><th>Bundle Count</th><th>Bundle Incentive</th></tr></thead><tbody>' +
    bunRows + bunFoot + "</tbody></table></div></div></div>" +
    "</div></section>"
  );
}

function renderKpiSection_(data) {
  // Targets come from target!B13:L16 (row 13 = revenue type, row 14 = KPI
  // name, row 15 = target count). Google Reviews is intentionally NOT a KPI
  // row — it is shown as its own tile at the top of the page.
  var allKpis = data.kpi.ar.concat(data.kpi.rr, data.kpi.sr);

  var rows = allKpis.map(function (k) {
    var pctVal = k.target > 0 ? pct_(k.achieved, k.target) : 0;
    // Approx ACHIEVABLE value: the BALANCE count still to be hit multiplied
    // by the per-unit approximate value from target!B16:L16 — how much can
    // still be won against this KPI's balance target this month.
    var approxBal = k.approxAmount > 0 && k.target > 0 ? k.balance * k.approxAmount : 0;
    var approxCell;
    if (k.approxAmount <= 0) {
      approxCell = '<span class="muted">&mdash;</span>';
    } else if (approxBal > 0) {
      approxCell = money_(approxBal) + '<br><span class="muted" style="font-size:12.5px">balance ' +
        k.balance + " × " + money_(k.approxAmount) + "</span>";
    } else {
      approxCell = money_(approxBal) + '<br><span class="pos" style="font-size:12.5px">target covered</span>';
    }
    return "<tr><td>" + esc_(k.label) + "</td><td>" + esc_(k.category) + "</td><td>" +
      (k.target > 0 ? k.target : "&mdash;") + "</td><td>" + k.achieved + "</td><td>" +
      (k.target > 0 ? pctBadgeHtml_(pctVal) : '<span class="muted">&mdash;</span>') +
      '</td><td class="' + (k.target > 0 && k.balance === 0 ? "pos" : k.balance > 0 ? "neg" : "muted") + '">' +
      (k.target > 0 ? k.balance : "&mdash;") + "</td><td>" +
      approxCell +
      "</td><td>" + money_(k.achievedValue) + "</td><td>" + money_(k.achievedGp) + "</td></tr>";
  }).join("");

  return (
    '<section class="section" id="sec-kpi"><h2>📊 KPI Target — Store KPI Parameters</h2><div class="body">' +
    '<div class="scroll-x">' +
    "<table><thead><tr><th>KPI</th><th>Type</th><th>Target</th><th>Achieved</th><th>Achieved %</th>" +
    "<th>Balance</th><th>Approx Achievable Value</th><th>Achieved Value</th><th>Achieved GP</th></tr></thead><tbody>" + rows +
    "</tbody></table></div>" +
    '<p class="muted" style="font-size:12.5px;margin:12px 0 0">Targets read from <b>target!B12:L14</b>. ' +
    "<b>Approx Achievable Value</b> = the balance count still needed × that KPI’s approximate unit value in <b>target!B15:L15</b> — the amount still winnable to close the balance target.</p>" +
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
    '<section class="section" id="sec-invoice"><h2>🧾 Invoice Valuation</h2><div class="body">' +
    '<div class="sec-desc">Definition: what invoices are really worth — line count, unique invoice count, revenue and gross profit per revenue type, plus AR (accessory attachment) rate next to invoices that qualified for one.</div>' +
    '<div class="scroll-x">' +
    "<table class=\"hdr-deep\"><thead><tr><th>Revenue Type</th><th>Invoice Lines</th><th>Unique Invoices</th><th>Revenue</th>" +
    "<th>Gross Profit</th></tr></thead><tbody>" + rows + totalRow + "</tbody></table></div>" +
    '<div class="cards" style="margin-top:14px">' +
    '<div class="card"><div class="label">Bundle Unique Invoices</div><div class="value">' +
    data.invoiceValuation.bundleUniqueInvoices + "</div></div>" +
    '<div class="card"><div class="label">AMC Unique Invoices</div><div class="value">' +
    data.invoiceValuation.amcUniqueInvoices + "</div></div>" +
    "</div>" +
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
function renderArSection_(data) {
  return (
    '<section class="section" id="sec-ar"><h2>🔗 AR Attachment — accessory sold with service / repair</h2><div class="body">' +
    '<div class="sec-desc">Definition: an invoice counts as an AR attachment when accessory revenue shares the invoice code with service or repair revenue; value and GP shown are the accessory portion of those invoices. Accessory-only invoices are excluded.</div>' +
    renderArAttachment_(data) + "</div></section>"
  );
}

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
/** Weekly revenue analysis — weeks of the current month, per revenue type. */
function computeWeeklyAnalysis_(data) {
  var now = new Date();
  var curMonth = now.getMonth() + 1, curYear = now.getFullYear();
  var monthName = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][curMonth - 1] +
    " " + curYear;
  var daysInMonth = daysInCurrentMonth_();
  var TYPES = [
    { key: "service", label: "Service", monthTarget: data.store.service.target },
    { key: "accessory", label: "Accessory", monthTarget: data.store.accessory.target },
    { key: "repair", label: "Repair", monthTarget: data.store.repair.target },
    { key: "grossProfit", label: "Gross Profit", monthTarget: data.store.grossProfit.target },
  ];

  // Week chunks of the month: 1–7, 8–14, 15–21, 22–28, 29–end.
  var weeks = [];
  for (var start = 1; start <= daysInMonth; start += 7) {
    var end = Math.min(daysInMonth, start + 6);
    var w = { start: start, end: end, workingDays: 0, achieved: {}, tgt: {}, idx: weeks.length + 1 };
    for (var d = start; d <= end; d++) if (isWorkingDay_(d)) w.workingDays++;
    TYPES.forEach(function (t) { w.achieved[t.key] = 0; });
    weeks.push(w);
  }
  var totalWorkingDays = 0;
  weeks.forEach(function (w) { totalWorkingDays += w.workingDays; });

  // Weekly target plan (owner's formula): split the month target by the
  // month's ACTUAL working days to get a per-working-day rate, then multiply
  // back by this week's working days.
  //   week target = month target ÷ actual working days × week working days
  weeks.forEach(function (w) {
    TYPES.forEach(function (t) {
      var perDay = t.monthTarget / Math.max(1, totalWorkingDays);
      w.tgt[t.key] = Math.round(perDay * w.workingDays);
    });
  });

  // Achieved per week: only current-month CRM dates count.
  (data.days || []).forEach(function (d) {
    if (d.undated || d.month !== curMonth || d.year !== curYear || !d.day) return;
    var w = weeks[Math.floor((d.day - 1) / 7)];
    if (!w) return;
    w.achieved.service += d.serviceRevenue || 0;
    w.achieved.accessory += d.accessoryRevenue || 0;
    w.achieved.repair += d.repairRevenue || 0;
    w.achieved.grossProfit += d.grossProfit || 0;
  });

  var monthTotals = {};
  TYPES.forEach(function (t) {
    monthTotals[t.key] = 0;
    weeks.forEach(function (w) {
      w.achieved[t.key] = Math.round(w.achieved[t.key]);
      monthTotals[t.key] += w.achieved[t.key];
    });
  });

  var daysElapsed = data.store.daysElapsed;
  weeks.forEach(function (w) {
    w.status = daysElapsed > w.end ? "Completed"
      : (daysElapsed >= w.start ? "In progress" : "Upcoming");
  });

  return { monthName: monthName, weeks: weeks, totalWorkingDays: totalWorkingDays,
    types: TYPES, monthTotals: monthTotals };
}

/** Renders section 7 — weekly achieved vs weekly target per revenue type. */
function renderWeeklyAnalysis_(data) {
  var wa = computeWeeklyAnalysis_(data);
  var TYPES = wa.types;

  var head1 = '<tr><th rowspan="2" class="sticky-col">Week</th><th rowspan="2">Wrk Days</th><th rowspan="2">Status</th>';
  var head2 = "<tr>";
  TYPES.forEach(function (t, i) {
    head1 += '<th colspan="3" class="grp grp-' + (i % 5) + '">' + esc_(t.label) + "</th>";
    head2 += '<th class="grp grp-' + (i % 5) + '">Achieved</th>' +
      '<th class="grp grp-' + (i % 5) + '">Week Target</th>' +
      '<th class="grp grp-' + (i % 5) + '">Balance</th>';
  });
  head1 += "</tr>";
  head2 += "</tr>";

  var body = "";
  wa.weeks.forEach(function (w) {
    var tr = "<tr>" +
      '<td class="sticky-col"><b>Week ' + w.idx + "</b> " +
      '<span class="muted">(' + w.start + "–" + w.end + " " + esc_(wa.monthName).split(" ")[0] + ")</span></td>" +
      '<td style="text-align:center">' + w.workingDays + "</td>" +
      '<td><span class="muted">' + esc_(w.status) + "</span></td>";
    TYPES.forEach(function (t, i) {
      var tk = t.key;
      var ach = w.achieved[tk], tgt = w.tgt[tk];
      var bal = Math.max(0, tgt - ach);
      var pill = tgt > 0
        ? '<span class="apill ' + (ach >= tgt ? "ok" : "no") + '">' + money_(ach) + "</span>"
        : '<span class="muted">' + money_(ach) + "</span>";
      tr += '<td class="cell-' + (i % 5) + '">' + pill + "</td>" +
        '<td class="cell-' + (i % 5) + '"><span class="tgt">' + money_(tgt) + "</span></td>" +
        '<td class="cell-' + (i % 5) + '">' + (bal > 0
          ? '<span class="muted">' + money_(bal) + "</span>"
          : '<span class="apill ok">done</span>') + "</td>";
    });
    body += tr + "</tr>";
  });

  var foot = '<tr class="total-row"><td class="sticky-col">Month</td><td style="text-align:center">' +
    wa.totalWorkingDays + "</td><td></td>";
  TYPES.forEach(function (t, i) {
    var tk = t.key;
    var ach = wa.monthTotals[tk], tgt = t.monthTarget;
    var bal = Math.max(0, tgt - ach);
    foot += '<td class="cell-' + (i % 5) + '"><b>' + money_(ach) + "</b></td>" +
      '<td class="cell-' + (i % 5) + '"><span class="tgt">' + money_(tgt) + "</span></td>" +
      '<td class="cell-' + (i % 5) + '">' + money_(bal) + "</td>";
  });
  foot += "</tr>";

  return (
    '<section class="section" id="sec-weekly"><h2>🗓 Weekly Revenue Analysis — ' +
      esc_(wa.monthName) + '</h2><div class="body">' +
    '<div class="sec-desc">Current month split into ' + wa.weeks.length + ' weeks — achieved vs weekly target, per revenue type</div>' +
    '<div class="scroll-x" style="margin-top:12px"><table class="matrix hdr-deep">' +
      "<thead>" + head1 + head2 + "</thead><tbody>" + body + foot + "</tbody></table></div>" +
    "</div></section>"
  );
}

/** Renders section 8 — the date-wise per-salesperson matrix. */
function renderDateWiseTable_(data) {
  var dailyTarget = data.store.dailyTarget;
  var people = data.salespeople;
  var TYPES = [["service", "Service"], ["accessory", "Accessory"], ["repair", "Repair"], ["grossProfit", "Gross Profit"]];

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
  var head1 = '<tr><th rowspan="2" class="sticky-col">Date</th>';
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

    var classes = [];
    if (d.holiday) classes.push("holiday");
    // Rows outside the collapsed window are tagged so the toggle can show them.
    if (!(d.undated ? false : visibleKeys[d.dateKey])) classes.push("extra-day");

    var tr = '<tr class="' + classes.join(" ") + '">';
    tr += '<td class="sticky-col"><b>' + esc_(d.undated ? "No date" : d.dateKey) + "</b>" +
      (d.weekday ? ' <span class="muted">' + esc_(d.weekday) + "</span>" : "") +
      (d.holiday ? ' <span class="muted">🔶</span>' : "") + "</td>";

    people.forEach(function (p, i) {
      var pd = byKey[p.key];
      var typeMap = {};
      if (pd) pd.types.forEach(function (t) { typeMap[t.type] = t; });

      TYPES.forEach(function (pair) {
        var tk = pair[0];
        // Gross profit uses the flat per-day total (it has no invoice-line
        // type rows); revenue types use their type row. Same condition.
        var achieved = tk === "grossProfit"
          ? (pd ? pd.grossProfit : 0)
          : (typeMap[tk] ? typeMap[tk].value : 0);
        var tgt = p.dailyTargets[tk] || 0;
        // Pill around the achieved number only; the daily target stays small
        // underneath. The met/missed condition is unchanged.
        var figure = tgt > 0
          ? '<span class="apill ' + (achieved >= tgt ? "ok" : "no") + '">' + money_(achieved) + "</span>"
          : '<span class="muted">' + money_(achieved) + "</span>";
        tr += '<td class="cell-' + (i % 5) + '">' + figure +
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
  var foot = '<tr class="total-row"><td class="sticky-col">Month to date</td>';
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
    '<section class="section" id="sec-datewise"><h2>📆 Date-wise Sales Revenue Analysis (by salesperson)</h2><div class="body">' +
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
    "Each cell shows <b>achieved / that engineer's daily target</b> for the revenue type — " +
    "the daily target is the <b>balance needed for the month ÷ working days remaining</b>. " +
    "Achieved values sit inside <b>oval pills</b> that turn <b>green</b> when the day meets the target " +
    "and <b>red</b> when it falls short. " +
    "Collapsed view shows the latest " + CONFIG.COLLAPSED_DAYS + " days; expand for every date.</p>" +
    "</div></section>"
  );
}

/* ============================================================================
 * WHATSAPP (OpenWA gateway) — Task 28
 * ----------------------------------------------------------------------------
 * Three faces:
 *   1) waConfig_()            — reads Script Properties (no secrets in code):
 *        WA_GATEWAY_URL  e.g. https://wa.example.com   (your OpenWA host)
 *        WA_API_KEY      admin/operator key from the OpenWA dashboard
 *        WA_SESSION      session id (default "default")
 *        WA_CHATID       default chat to push replies to (digits or JID)
 *   2) whatsappSendAnswer_(query, chatId) — runs the sheet-search chat bot and
 *      POSTs the reply via OpenWA  POST /api/sessions/{id}/messages/send-text
 *   3) doPost(e)              — inbound webhook target: OpenWA fires
 *      message.received here → we answer over WhatsApp automatically.
 *      If nothing is configured the bot simply won't auto-answer; the page
 *      keeps working and falls back to a "wa.me deep link" share button.
 * ========================================================================= */

function waConfig_() {
  var p = (typeof PropertiesService !== "undefined")
    ? PropertiesService.getScriptProperties() : null;
  function gp(k) { return p ? str_(p.getProperty(k)) : ""; }
  var cfg = {
    base: gp("WA_GATEWAY_URL").replace(/\/+$/, ""),
    key: gp("WA_API_KEY"),
    session: gp("WA_SESSION") || "default",
    chat: gp("WA_CHATID"),
  };
  cfg.live = !!(cfg.base && cfg.key);
  return cfg;
}

/** Normalise any phone/JID/garbage into a WhatsApp chatId ("…@c.us"). */
function waCleanChatId_(v) {
  var s = str_(v).trim();
  if (s.indexOf("@") > -1) return s;              // already a JID
  s = s.replace(/[^\d]/g, "");
  if (s.length === 10) s = "91" + s;              // assume India (10 digits)
  return s ? (s + "@c.us") : "";
}

/** Strip the bot's <td>-wrapped HTML rows into WhatsApp-friendly plain text. */
function waPlainRows_(rows) {
  return (rows || []).map(function (cols) {
    return cols.map(function (c) {
      return c.replace(/<[^>]+>/g, "").replace(/&amp;/g, "&").replace(/&lt;/g, "<")
        .replace(/&gt;/g, ">").replace(/&#39;/g, "'").replace(/&quot;/g, '"').trim();
    }).filter(function (s) { return s !== ""; }).join("  ·  ");
  }).filter(function (s) { return s !== ""; });
}

function waComposeReply_(query) {
  var c = chatQuery(query);
  var scan = chatAllSheetsScan_(query);
  var out = [c.title];
  waPlainRows_(c.rows || []).forEach(function (l) { out.push("  " + l); });
  if (scan.rows.length) {
    out.push("— matches across all tabs —");
    waPlainRows_(scan.rows).forEach(function (l) { out.push("  " + l); });
  }
  if (c.note) out.push("ℹ︎ " + c.note);
  if (scan.touched.length) out.push("📑 matched: " + scan.touched.join(", "));
  return out.filter(function (l) { return str_(l).trim() !== ""; }).join("\n");
}

/** POST a text reply through the OpenWA gateway. Returns {ok, code, body|link}. */
function waSendText_(chatId, text, cfg) {
  cfg = cfg || waConfig_();
  if (!cfg.live) return { ok: false, offline: true };
  var url = cfg.base + "/api/sessions/" + encodeURIComponent(cfg.session) + "/messages/send-text";
  var res = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    headers: { "X-API-Key": cfg.key },
    payload: JSON.stringify({ chatId: chatId, text: text }),
    muteHttpExceptions: true,
  });
  var code = res.getResponseCode();
  return { ok: code >= 200 && code < 300, code: code, body: str_(res.getContentText()).slice(0, 400) };
}

/**
 * google.script.run endpoint: answer a question and send it to WhatsApp.
 * If the gateway is not configured, returns a wa.me deep link the page opens.
 */
function sendWhatsAppAnswer(p) {
  var obj = {};
  try { obj = (typeof p === "string") ? JSON.parse(p) : (p || {}); } catch (e) { obj = {}; }
  var query = str_(obj.q);
  var chatId = waCleanChatId_(obj.chat);
  var cfg = waConfig_();
  if (!query) return JSON.stringify({ ok: false, error: "empty query" });
  var text = waComposeReply_(query);
  if (!chatId) chatId = cfg.chat;
  if (!chatId) return JSON.stringify({ ok: false, error: "no chatId", text: text });
  var sent = waSendText_(chatId, text, cfg);
  if (sent.offline) {
    return JSON.stringify({ ok: false, offline: true,
      link: "https://wa.me/" + chatId.replace(/@.*$/, "") + "?text=" + encodeURIComponent_(text).slice(0, 1800) });
  }
  return JSON.stringify({ ok: sent.ok, code: sent.code, error: sent.ok ? null : sent.body });
}

/** Inbound webhook: OpenWA → Apps Script. Answer over WhatsApp. */
function doPost(e) {
  var out = ContentService.createTextOutput(JSON.stringify({ ok: true }));
  var cfg = waConfig_();
  if (!cfg.live) return out.setMimeType(ContentService.MimeType.JSON);
  var msg = {};
  try { msg = JSON.parse((e && e.postData && e.postData.contents) || "{}"); } catch (ex) { msg = {}; }
  var type = str_(msg.type || msg.event || "");
  if (type !== "message.received" && type !== "message.any") {
    return out.setMimeType(ContentService.MimeType.JSON);
  }
  var m = msg.data || msg.message || msg;
  var from = waCleanChatId_(m.chatId || m.from || m.to || "");
  var body = str_(m.body || m.text || "");
  if (!from || !body || m.fromMe || /@g\.us$/.test(from)) {
    return out.setMimeType(ContentService.MimeType.JSON);
  }
  waSendText_(from, waComposeReply_(body), cfg);
  return out.setMimeType(ContentService.MimeType.JSON);
}
/** @internal encode helper — keeps V8 happy in Apps Script eval too. */
function encodeURIComponent_(s) {
  return encodeURIComponent(str_(s)).replace(/[!'()*]/g, function (c) {
    return "%" + c.charCodeAt(0).toString(16).toUpperCase();
  });
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
