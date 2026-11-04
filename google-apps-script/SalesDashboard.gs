/**
 * SALES PERFORMANCE DASHBOARD — Google Apps Script web app
 * =========================================================
 * One doGet serves a single HTML page that already contains all four pane
 * bodies; the tab bar flips between them with zero server contact.
 *
 * PAGES (tabs)
 *   INDEX  A1 store cards · A2 store table · A6 working days · weekly plan
 *   CRE    A1/A2/A6 + B KPI target · B1 incentive · B2 salesperson · B3 AR
 *          · C/C1 walk-in analysis · D1 customer satisfaction
 *   TECH   D2 repair type · D3 decline repair · D4 requote & win-back
 *          · E1 job buckets (9 aging bands) · communication · open jobs
 *   SM     the full master report, groups A–G
 *
 * THE STANDING RULES THIS FILE IMPLEMENTS
 *   • Everything is prefetched during the load itself: a cold doGet computes
 *     and embeds all four panes in one pass (buildAllPanes_ + cachePanes_).
 *     Once loading completes, no tab tap ever reads sheet data or reloads.
 *   • Tabs and the refresh button are plain <button>s with no href — a tap
 *     can only flip panes; ↻ Refresh re-reads the sheets and swaps the new
 *     values in place (refreshDashboardSections), never a page reload.
 *   • A 15-minute in-page auto-refresh re-primes the same caches silently.
 *   • Caches: chunked whole-page cache (dash_pg_v7_*), per-pane script cache
 *     (dash_pane_v6_* + dash_meta_v6) and a Drive-backed 15-minute store
 *     (SalesDashboard_paneCache_v2.json) that survives cache evictions.
 *     Any pane recompute invalidates the assembled pages and Drive-merges.
 *   • No page-wide export: exactly seven tables carry an ⬇ Export button
 *     (store, ops, repairType, decline, jobs, margin, kpi) → one endpoint,
 *     exportDashboardSectionToSpreadsheet(key).
 *   • C-SAT reads target!E47:F51 (name + C-SAT per CRE, store average
 *     derived) and auto-extends to the full satisfaction grid up to
 *     target!E47:O51 when the sheet carries one; period name from D47.
 *   • CRE page manifest: A1 A2 A6 B B1 B2 B3 C(C1) D1 · TECH: D2 D3 D4 + E1
 *     · INDEX weekly is % only · salesperson performance shows % only.
 *   • Topbar, brand, digital clock, refresh button, data-updated stamp and
 *     text colors/sizes are fixed — never restyle them.
 *   • The first line of the served page is a <!-- BUILD TASK… --> marker so
 *     the live deployment can be verified against this build.
 *
 * TOP-LEVEL FLOW
 *   sheets → loaders/parsers → compute* models → render* pane HTML
 *   → composers → cache layers → doGet / refreshDashboardSections
 *
 * The sheet-search chat assistant (ask me anything) lives in module 13 and
 * answers from the same source sheets. All external sheet reads happen in
 * the loaders only; renderers work purely from computed models.
 */

/* ==========================================================================
 * 01 · CONFIGURATION                                                      *
 * ======================================================================== */

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
  DECLINE_SHEET_NAMES: ["decline", "declined", "declined-repair"],
  QUOTE_SHEET_NAMES: ["quote", "quotes", "requote"],
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

/* ==========================================================================
 * 02 · SPREADSHEET ACCESS, DATE & ROW PARSING                             *
 * ======================================================================== */

/** Resolves the spreadsheet regardless of whether we're in a UI context (Mode A) or an HTTP request (Mode B). */
function getSpreadsheet_() {
  if (CONFIG.SPREADSHEET_ID) return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  return SpreadsheetApp.getActiveSpreadsheet();
}

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
    orderCode: findColumn_(header, [/^ordercode$/i, /^order\s*code$/i, /order.*(code|no|num)/i]),
    partType: findColumn_(header, [/part.*type/i]),
    // optional discount column feeding the GP "Discounted Cases" tile.
    discountAmount: findColumn_(header, [/^discount\s*(amount|amt|value)?$/i, /disc.*(amount|amt|value)/i, /^discount/i]),
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
      orderCode: col.orderCode !== -1 ? str_(row[col.orderCode]) : "",
      partType: col.partType !== -1 ? str_(row[col.partType]) : "",
      discountAmount: col.discountAmount !== -1 ? Math.abs(num_(row[col.discountAmount])) : 0,
    });
  }
  return out;
}

function daysInCurrentMonth_() {
  var now = new Date();
  return new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();
}

function isWorkingDay_(day) {
  return CONFIG.HOLIDAYS.indexOf(day) === -1;
}

/* ==========================================================================
 * 03 · CORE VALUE & TEXT UTILITIES                                        *
 * ======================================================================== */

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

function esc_(s) {
  return String(s === null || s === undefined ? "" : s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function money_(n) { return "₹" + Math.round(n || 0).toLocaleString("en-IN"); }

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

/* ==========================================================================
 * 04 · TARGET SHEET & INCENTIVE PARSING                                   *
 * ======================================================================== */

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

// Chase tiers (ascending) for the "next slab" balance lines on person cards.
var SVC_INC_TIERS33 = [71, 81, 91, 100, 110];

var ACC_INC_TIERS33 = [51, 61, 71, 81, 110];

function serviceIncentiveForPct33_(pctVal) {
  var p = (pctVal || 0) / 100;
  if (p > 1.10) return 12000 * p / 1.10;
  if (p <= 1.09) return 11000 * p / 1.01;
  if (p >= 1.00) return 10000 * p / 1.00;
  if (p >= 0.91) return 8000 * p / 0.99;
  if (p >= 0.81) return 7000 * p / 0.90;
  if (p >= 0.71) return 6000 * p / 0.80;
  return 0;
}

function accessoryIncentiveForPct33_(pctVal) {
  var p = (pctVal || 0) / 100;
  if (p > 1.10) return 8000 * p / 1.10;
  if (p >= 0.81) return 6000 * p / 1.00;
  if (p >= 0.71) return 4000 * p / 0.80;
  if (p >= 0.61) return 3000 * p / 0.70;
  if (p >= 0.51) return 2500 * p / 0.60;
  return 0;
}

function nextIncentiveSlab33_(tiers, formulaFn, pctVal, tgt, achVal) {
  if (!tiers || !tiers.length || tgt <= 0) return null;
  for (var i = 0; i < tiers.length; i++) {
    if (pctVal < tiers[i]) {
      var need = Math.max(0, Math.ceil((tiers[i] / 100) * tgt) - achVal);
      return { label: tiers[i] + "%", pct: tiers[i], amount: formulaFn(tiers[i]), need: need };
    }
  }
  return null;
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

  // ---- Task 40: fixed-address fallback — the owner's slab card lives at
  //      target!A29:F33 (range in A, AMC rate in B, acc-AMC count in D,
  //      acc-AMC rate in E). Used when the loose anchor above found nothing.
  if (!amcSlabs.length) {
    for (var fr = 28; fr <= 32 && fr < rows.length; fr++) {
      var fLab = at(fr, 0), fRate = at(fr, 1), fAccCnt = parseMoney_(at(fr, 3)), fAccRate = at(fr, 4);
      if (!fLab || !/^\d/.test(fLab)) continue;
      var fRange = parseRangeLabel_(fLab);
      var fRateInfo = parseAmcRate_(fRate);
      var fAccInfo = parseAmcRate_(fAccRate);
      if (fRateInfo.rate <= 0 && fRateInfo.fixed <= 0) continue;
      amcSlabs.push({
        label: fLab, rateLabel: fRate, rate: fRateInfo.rate, fixed: fRateInfo.fixed,
        min: Math.min(fRange.min, OPEN_ENDED), max: Math.min(Math.max(fRange.max, fRange.min), OPEN_ENDED),
        accCount: fAccCnt > 0 ? fAccCnt : undefined,
        accRate: fAccInfo.rate > 0 ? fAccInfo.rate : undefined,
        accFixed: fAccInfo.fixed > 0 ? fAccInfo.fixed : undefined,
      });
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

/* ==========================================================================
 * 05 · BUSINESS DATA MODEL                                                *
 * ======================================================================== */

/** Shared pipeline: read crm + target sheets, parse, and compute the dashboard data object. */
function loadDashboardData_() {
  // Deployment hardening: if this helper ever goes missing from the project
  // (stale/partial paste), fall back to the bound spreadsheet directly.
  var ss = (typeof getSpreadsheet_ === "function") ? getSpreadsheet_()
    : SpreadsheetApp.getActiveSpreadsheet();
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
  // D — the deck numbers ride on the same parse, current month only.
  result.data.deckAnalysis = computeDeckAnalysis_(ss, allRows);

  // Store name for the page title — Task 43: taken verbatim from crm!F2;
  // the "iPlanetcare" prefix is no longer glued on.
  result.data.store.storeCode =
    (crmValues[1] && crmValues[1].length > 5) ? str_(crmValues[1][5]) : "";

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
    var svcInc = serviceIncentiveForPct33_(pct_(ach.service, t.service));
    var accInc = accessoryIncentiveForPct33_(pct_(ach.accessory, t.accessory));
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

    // NEXT-tier chase (Task 33 formula tiers, ascending).

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
      nextSvcSlab: nextIncentiveSlab33_(SVC_INC_TIERS33, serviceIncentiveForPct33_, pct_(ach.service, t.service), t.service, ach.service),
      nextAccSlab: nextIncentiveSlab33_(ACC_INC_TIERS33, accessoryIncentiveForPct33_, pct_(ach.accessory, t.accessory), t.accessory, ach.accessory),
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
  var gpExtras = computeGpExtras_(rows);

  return {
    crmRows: rows, target: target, days: days, activeDays: activeDays, store: store,
    revenueTypes: revenueTypes, invoiceValuation: invoiceValuation, salespeople: salespeople,
    amc: amc, bundle: bundle, kpi: kpi, gpExtras: gpExtras, generatedAt: new Date(),
  };
}

/**
 * compact part description for the GP tiles. Owner examples:
 *   "Battery, iPhone 14 Plus, India" → "battery"
 *   "Protect+withAppleCareServicesforiPhone17ProMax" → "protect+"
 *   "MCL TG 2.5D 2ND GEN-APPL-IPH 17PRMAX-BLK" → "temper"
 * Unknown descriptions fall back to their first word (max 12 chars).
 */
function shortPart_(desc) {
  var t = str_(desc).toLowerCase();
  if (!t) return "—";
  if (/protect\s*\+|protect\s*plus/.test(t)) return "protect+";
  if (/temper|temeper|tamper|\btg\b/.test(t)) return "temper";
  if (/batt/.test(t)) return "battery";
  if (/display|screen|lcd/.test(t)) return "display";
  if (/apple\s*care|care\s*\+|\bac\s*\+/.test(t)) return "ac+";
  if (/\bcase\b|\bcover\b|\bhousing\b/.test(t)) return "case";
  if (/adapter|cable|charg/.test(t)) return "cable";
  if (/\bamc\b/.test(t)) return "amc";
  var w = t.replace(/[^a-z0-9+\s]/g, " ").trim().split(/\s+/)[0];
  return w ? w.slice(0, 12) : "—";
}

/**
 * GP-page extras.
 *  • High / Low margin leaders: margin % = Σ crm!GrossProfit ÷ Σ crm!TaxableAmount
 *    computed per unique invoice code (SALES rows only; invoices with a zero
 *    or negative taxable base are skipped — their % is meaningless). The top
 *    5 at each end of the scale are reported, across ALL revenue types.
 *  • Discounted cases: rows with crm!DiscountAmount > 0 — unique invoice
 *    count overall and per revenue type, total discounted value, and the
 *    discounted value as % of the gross profit of those lines.
 */
function computeGpExtras_(rows) {
  var sale = (rows || []).filter(isSaleRow_);

  // ---- margin leaders ----
  var byInv = {};
  sale.forEach(function (r) {
    var k = r.invoiceCode;
    if (!k) return;
    if (!byInv[k]) byInv[k] = { invoice: k, person: r.salesPerson || "—", types: {}, parts: {}, taxable: 0, gp: 0 };
    var b = byInv[k];
    b.taxable += r.taxableAmount;
    b.gp += r.grossProfit;
    // short part label(s) of this invoice's lines (no invoice number
    // is shown on the GP tiles anymore).
    b.parts[shortPart_(r.itemDescription)] = true;
    var tk = revenueTypeKey_(r.revenueType);
    if (tk) b.types[tk] = true;
  });
  var margins = [];
  Object.keys(byInv).forEach(function (k) {
    var b = byInv[k];
    if (!(b.taxable > 0)) return;
    b.margin = (100 * b.gp) / b.taxable;
    b.typeLabel = ["repair", "accessory", "service"].filter(function (t) { return b.types[t]; })
      .map(function (t) { return t === "repair" ? "Repair" : t === "accessory" ? "Accessory" : "Service"; })
      .join(" + ") || "—";
    var pset = Object.keys(b.parts);
    b.partLabel = pset.slice(0, 3).join(" + ") + (pset.length > 3 ? " …" : "");
    margins.push(b);
  });
  var high = margins.slice().sort(function (a, b) { return b.margin - a.margin; }).slice(0, 5);
  var low = margins.slice().sort(function (a, b) { return a.margin - b.margin; }).slice(0, 5);

  // ---- discounted cases ----
  var disc = {
    count: 0, value: 0, gp: 0, pct: 0,
    byType: { repair: 0, accessory: 0, service: 0 },
    byTypeValue: { repair: 0, accessory: 0, service: 0 },
    byTypeTop: { repair: null, accessory: null, service: null },
  };
  var partFreq = { repair: {}, accessory: {}, service: {} };
  var seenInv = {}, seenTypeInv = { repair: {}, accessory: {}, service: {} };
  sale.forEach(function (r) {
    var d = r.discountAmount || 0;
    if (d <= 0) return;
    disc.value += d;
    disc.gp += r.grossProfit;
    var tk = revenueTypeKey_(r.revenueType);
    var cid = r.invoiceCode;
    if (cid) {
      if (!seenInv[cid]) { seenInv[cid] = true; disc.count++; }
      if (tk && !seenTypeInv[tk][cid]) { seenTypeInv[tk][cid] = true; disc.byType[tk]++; }
    }
    if (tk) {
      disc.byTypeValue[tk] += d;
      var sp = shortPart_(r.itemDescription);
      partFreq[tk][sp] = (partFreq[tk][sp] || 0) + 1;
    }
  });
  // Top discounted part per revenue type (short label + how many lines).
  ["repair", "accessory", "service"].forEach(function (tk) {
    var bestK = "", bestN = 0;
    Object.keys(partFreq[tk]).forEach(function (sp) {
      if (partFreq[tk][sp] > bestN) { bestN = partFreq[tk][sp]; bestK = sp; }
    });
    if (bestK) disc.byTypeTop[tk] = { label: bestK, n: bestN };
  });
  disc.value = Math.round(disc.value);
  disc.pct = disc.gp > 0 ? Math.round((1000 * disc.value) / disc.gp) / 10 : 0;

  return { marginHigh: high, marginLow: low, discount: disc };
}

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

/* ==========================================================================
 * 06 · PROCESS DATA MODEL (walk-in, jobs, repair, C-SAT, ops)             *
 * ======================================================================== */

var WALKIN_SPEC = {
  mobile: [/mobile.*no/i, /^mobile$/i, /phone/i],
  tokenDate: [/token.*date/i, /^date$/i],
  visitPurpose: [/visit.*purpose.*desc/i, /visit.*purpose/i, /purpose/i],
  caseId: [/case.*id/i, /^case$/i],
  invoiceCode: [/invoice.*code/i, /invoice.*no/i, /^invoice$/i],
  remark: [/remark/i, /comment/i, /note/i],
  // Who generated the token. Optional — auto-detected from common spellings.
  createdBy: [/token.*creat.*by/i, /creat.*by/i, /token.*user/i, /user.*id/i, /^user.*name$/i, /staff/i, /employee/i],
  // customer source for the source-analysis pivot (walkin!CustomerVisitSource).
  visitSource: [/customer.*visit.*source/i, /visit.*source.*desc/i, /customer.*source/i, /how.*did.*(you|they).*find/i],
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
  var hasSource = read.col && read.col.visitSource !== -1;   // Task 37

  var dayMap = {};
  var purposeTotals = {};
  var creatorTotals = {};
  var sourceTotals = {};
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

    // customer source pivot — walkin!CustomerVisitSource, against the
    // total number of walk-ins. Kept EXACTLY as written in the sheet.
    if (hasSource) {
      var srcv = str_(r.visitSource) || "Not specified";
      sourceTotals[srcv] = (sourceTotals[srcv] || 0) + 1;
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
  // source pivot descending, share of total walk-ins.
  var sources = toSortedList(sourceTotals).map(function (s) {
    s.share = pct_(s.count, totals.walkins);
    return s;
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
    hasSource: hasSource,
    days: days,
    purposeCols: purposeCols,
    purposeColTotals: purposeColTotals,
    creators: creators,
    sources: sources,
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
  // GSX ID and device serial shown on the Open Jobs table.
  gsxCode: [/^gsxcode$/i, /^gsx.*(code|id|no\.?|number)$/i, /gsx.*(job|case).*id/i],
  serialNo1: [/^serialno1$/i, /serial.*no\.?1/i, /serial.*number/i, /^serial/i, /device.*serial/i, /imei/i],
  // replaced part for coverage split (battery / display buckets).
  partDescription: [/^partdescription$/i, /^partdesc$/i, /^part\s*description$/i],
  // Task 40 (Coverage v2): battery/display from service!IssueDescription and
  // the W/NW split from service!ProductWarrantyDesc — owner's final columns.
  issue: [/^issuedescription$/i, /issue.*desc/i, /^issue$/i],
  productWarranty: [/^productwarrantydescription$/i, /^productwarrantydesc$/i,
    /product.*warranty.*desc/i, /product.*warranty/i],
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
    var ready = parseAnyDate_(r.readyPickup);   // Task 38 — popup detail

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
      readyDate: ready.valid ? ready.dayKey : "—",
      jobAge: age,
      advance: advance,
      communicated: !!comm,
      commDate: comm ? comm.dayKey : "—",
      commTs: comm ? comm.ts : 0,
      gsxStatus: gsxStatus || "—",
      crmRepairStatus: crmRepairStatus || "—",
      warranty: warrantyTxt || "—",
      repairType: repairTypeTxt || "—",
      gsxId: str_(r.gsxCode) || "—",
      serialNo: str_(r.serialNo1) || "—",
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
   "advanceDirect", "status", "deliveredDate", "readyPickup",
   "gsxCode", "serialNo1"]
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

/** Task 39: compact repair-type labels for the narrow jobs columns. */
function repairTypeShort_(t) {
  var s = str_(t);
  if (/service\s*non[\s-]*repair/i.test(s)) return "N-repair";
  if (/whole\s*unit.*mail[\s-]*in/i.test(s) || /mail[\s-]*in.*service\s*locat/i.test(s)) return "mail-in";
  return s;
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
  // rows with a blank service!RepairDocType are excluded too.
  var rows = all.filter(function (r) { return str_(r.repairType) !== "" && !isSNRType_(r.repairType); });

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

  // Task 47 (CRE page): completed repairs this month = unique non-SNR cases
  // whose Ready_For_Pickup date falls inside the running month.
  var completedThisMonth = 0;
  rows.forEach(function (r) {
    if (r.readyTs > 0) {
      var rd = new Date(r.readyTs);
      if (rd.getFullYear() === now.getFullYear() && rd.getMonth() === now.getMonth()) completedThisMonth++;
    }
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
    completedThisMonth: completedThisMonth,
    prevRepairAvgTotal: typeComp.reduce(function (a, c) { return a + c.perMonth; }, 0),
  };
}

/**
 * Process-page C-SAT block — read from the target sheet, presented cleanly:
 * the block moved within the same sheet:
 *   target!D47    = the period name shown as the dashboard title
 *   target!E47:O51 = the C-SAT grid (11 cols × 5 rows; row 47 = headers,
 *                   rows 48-51 = one row per salesperson plus a TOTAL row)
 * the sheet may carry only the NARROW C-SAT report —
 *   target!E47:F51 — name in E, C-SAT% in F. The reader auto-detects this
 *   layout: no bucket columns, no TOTAL column, and the store C-SAT is
 *   derived as the average of the per-CRE values when no TOTAL row exists.
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

  // table starts at sheet column E (0-indexed 4), rows 47..51.
  // Period name: cell left of the header (D47), then above the name column
  // (E46), then the legacy K1 position for sheets still on the old layout.
  var L = 4;                             // first column of the table (0-indexed)
  var FIRST_ROW = 46, LAST_ROW = 50;     // sheet rows 47..51
  var periodName = at(FIRST_ROW, L - 1) || at(FIRST_ROW - 1, L) || at(0, 10) || "";

  var header = [];
  for (var c = 0; c < 11; c++) header.push(at(FIRST_ROW, L + c));

  // Identify the special columns by header text.
  function isJunkHeader(h) { return h === "" || h === "-" || /^–+$/.test(h); }
  function findCol(re) {
    for (var i = 0; i < header.length; i++) if (re.test(header[i])) return i;
    return -1;
  }
  var totalCol = findCol(/^total$/i);
  var csatCol = findCol(/c[\-\.\s]*sat/i);
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
  // the narrow C-SAT report (target!E47:F51) has no TOTAL row and
  // no sheet-level C-SAT — derive the store C-SAT / CFI as the average of
  // the per-person values (blank or 0 rows keep out of the average) so the
  // summary chips still show a store number.
  if (csatCol >= 0 && !totalRow.csat) {
    var csSum = 0, csN = 0;
    people.forEach(function (p) {
      var raw = str_(p.csat);
      if (!raw) return;
      var v = parseMoney_(raw);
      if (isFinite(v) && v > 0) { csSum += v; csN++; }
    });
    if (csN) totalRow.csat = (Math.round((csSum / csN) * 100) / 100) + "%";
  }
  if (cfiCol >= 0 && !totalRow.cfi) {
    var cfSum = 0, cfN = 0;
    people.forEach(function (p) {
      var raw2 = str_(p.cfi);
      if (!raw2) return;
      var w = parseMoney_(raw2);
      if (isFinite(w) && w > 0) { cfSum += w; cfN++; }
    });
    if (cfN) totalRow.cfi = (Math.round((cfSum / cfN) * 100) / 100) + "%";
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
    hasTotalCol: totalCol >= 0, hasCsatCol: csatCol >= 0, hasCfiCol: cfiCol >= 0,
    satisfiedCount: satisfiedCount, dissatisfiedCount: dissatisfiedCount,
  };
}

/**
 * TASK 35 — Operation-page extra modules.
 * Decline Repair: unique decline!CaseId where decline!GSXRepairStatusDescription
 *   == "Declined - Rejected by Local Approver".
 * Requote & Win-back (Task 41 rule): a serial number (service!SerialNo1) whose
 *   service case is a parts repair — carry-in / mail-in (service!RepairDocType,
 *   Service Non-Repair excluded) AND out of warranty (service!ProductWarranty
 *   Desc) — must have a "Repair Revenue" quote (quote!RevenueType) that was
 *   first REJECTED and later APPROVED (quote!lastAction) for the same serial.
 *   Count = number of such serials (unique per case ID linkage).
 */
function computeOperationExtras_(ss) {
  var out = {
    decline: { count: 0, matched: 0, sheetName: null },
    winback: { count: 0, matched: 0, sheetName: null },
  };
  // ---- decline ----
  try {
    var ds = getSheetByNames_(ss, CONFIG.DECLINE_SHEET_NAMES);
    if (ds) {
      out.decline.sheetName = ds.getName();
      var rd = readSheetRows_(ds, {
        caseId: [/^caseid$/i, /case.*id/i],
        gsx: [/^gsxrepairstatusdescription$/i, /gsx.*repair.*status/i, /gsx.*status/i, /repair.*status.*desc/i],
      }, []);
      var seen = {};
      rd.rows.forEach(function (r) {
        var g = str_(r.gsx).toLowerCase();
        // Lenient casing/whitespace; the phrase itself must be the reason.
        if (/declined\s*-\s*rejected\s+by\s+local\s+approver/i.test(str_(r.gsx))) {
          out.decline.matched++;
          var cid = str_(r.caseId);
          if (cid && !seen[cid]) { seen[cid] = true; out.decline.count++; }
          else if (!cid) out.decline.count++;
        }
      });
    }
  } catch (e) { /* module is best-effort */ }
  // ---- win-back (Task 41) ----
  try {
    var qs = getSheetByNames_(ss, CONFIG.QUOTE_SHEET_NAMES);
    var svs = getSheetByNames_(ss, CONFIG.SERVICE_SHEET_NAMES);
    if (qs && svs) {
      out.winback.sheetName = qs.getName();
      // Eligible cases: carry-in / mail-in parts repair (Service Non-Repair
      // excluded) AND out of warranty via service!ProductWarrantyDesc
      // (CoverageCodeDescription as fallback). CaseId → serial number.
      var rdS = readSheetRows_(svs, SERVICE_SPEC, ["caseDate"]);
      var eligibleSerial = {};             // caseId → serialNo1 (eligible only)
      rdS.rows.forEach(function (r) {
        var cid = str_(r.caseId);
        if (!cid) return;
        if (isSNRType_(str_(r.repairDocType))) return;
        var t = str_(r.repairDocType);
        if (!/carry|mail/i.test(t)) return;
        var pw = (str_(r.productWarranty) || str_(r.warranty)).toLowerCase();
        if (!/\bout\b|\bow\b|\bno\b|\bnon\b|\bnot\b|expired|void/.test(pw)) return;
        eligibleSerial[cid] = str_(r.serialNo1) || cid;
      });
      // Quote side: Repair-Revenue rows only; for every eligible serial,
      // need one REJECTED row and one APPROVED row (same serial won back).
      var rq = readSheetRows_(qs, {
        caseId: [/^caseid$/i, /case.*id/i],
        revenueType: [/^revenuetype$/i, /revenue.*type/i],
        lastAction: [/^lastaction$/i, /last.*action/i],
      }, []);
      var rejected = {}, approved = {};
      rq.rows.forEach(function (r) {
        if (!/^\s*repair\s*revenue\s*$/i.test(str_(r.revenueType))) return;
        out.winback.matched++;
        var cid = str_(r.caseId);
        var sn = eligibleSerial[cid];
        if (!sn) return;
        var la = str_(r.lastAction).toLowerCase();
        if (/reject/.test(la)) rejected[sn] = true;
        else if (/approved/.test(la)) approved[sn] = true;
      });
      Object.keys(rejected).forEach(function (sn) {
        if (approved[sn]) out.winback.count++;
      });
    } else if (qs) {
      out.winback.sheetName = qs.getName();
    }
  } catch (e) { /* module is best-effort */ }
  return out;
}

/** Loads everything the Performance & Process pages need. */
function loadPerformanceData_() {
  // Deployment hardening: if this helper ever goes missing from the project
  // (stale/partial paste), fall back to the bound spreadsheet directly.
  var ss = (typeof getSpreadsheet_ === "function") ? getSpreadsheet_()
    : SpreadsheetApp.getActiveSpreadsheet();
  var jm = computeJobManagement_(ss);
  var ra = computeRepairAnalytics_(jm.serviceRows || []);
  // stamp EVERY tab with the refresh timestamp — whenever the
  // dashboard is (re)loaded the owning sheets carry the "updated at" time.
  var refreshStamp = stampSheetsRefreshed_(ss);
  var opsExtras = computeOperationExtras_(ss);

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
  // the freshness line also shows the TIME of the last record.
  function latestStamp_(names, spec, field) {
    var sheet = getSheetByNames_(ss, names);
    if (!sheet) return null;
    var rd = readSheetRows_(sheet, spec, [field]);
    var ts = 0, lab = "";
    rd.rows.forEach(function (r) {
      var d = parseAnyDate_(r[field]);
      if (d.valid && d.ts > ts) {
        ts = d.ts;
        var hh = Math.floor(d.minutes / 60), mi = d.minutes % 60;
        lab = d.dayKey + (d.minutes ? " " + (hh < 10 ? "0" + hh : hh) + ":" + (mi < 10 ? "0" + mi : mi) : "");
      }
    });
    return ts ? lab : null;
  }
  var freshness = {
    crm: latestStamp_([CONFIG.CRM_SHEET_NAME], { invoiceDate: [/invoice.*date/i, /^inv.*date/i] }, "invoiceDate") || "no rows",
    walkin: latestStamp_(CONFIG.WALKIN_SHEET_NAMES, { tokenDate: WALKIN_SPEC.tokenDate }, "tokenDate") || "no rows",
    service: latestStamp_(CONFIG.SERVICE_SHEET_NAMES, { caseDate: SERVICE_SPEC.caseDate }, "caseDate") || "no rows",
    comm: latestStamp_(CONFIG.COMM_SHEET_NAMES, { createdDate: COMM_SPEC.createdDate }, "createdDate") || "no comm sheet",
    advance: "no date column",
  };

  return {
    walkin: computeWalkinAnalysis_(ss),
    jobs: jm,
    repair: ra,
    targetBlock: computeTargetBlock_(ss),
    freshness: freshness,
    opsExtras: opsExtras,
    refreshStamp: refreshStamp,
    generatedAt: new Date(),
  };
}

/**
 * TASK 35 — write the dashboard-refresh timestamp into EVERY sheet tab so
 * anyone opening the Google sheet can see when the dashboard last refreshed.
 * The stamp sits one column to the right of the last used column in row 1,
 * as a single-cell label: "⟳ Dashboard refresh · dd-MM-yyyy HH:mm".
 * The same cell is located by its prefix on later refreshes and overwritten,
 * so the sheet never grows extra columns. Returns the stamp string shown on
 * the index page, or "" when the sheet service is unavailable (tests).
 */
function stampSheetsRefreshed_(ss) {
  try {
    var now = new Date();
    function p(n) { return (n < 10 ? "0" : "") + n; }
    var stamp = "⟳ Dashboard refresh · " +
      p(now.getDate()) + "-" + p(now.getMonth() + 1) + "-" + now.getFullYear() +
      " " + p(now.getHours()) + ":" + p(now.getMinutes());
    var PREFIX = "⟳ Dashboard refresh";
    ss.getSheets().forEach(function (sh) {
      try {
        var lastCol = Math.max(1, sh.getLastColumn ? sh.getLastColumn() : 1);
        var target = 1;
        // Look for our previous stamp anywhere in row 1 to overwrite it.
        var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
        var found = 0;
        for (var c = 0; c < hdr.length; c++) {
          if (String(hdr[c]).indexOf(PREFIX) === 0) { found = c + 1; break; }
        }
        target = found || (lastCol + 1);
        sh.getRange(1, target).setValue(stamp);
      } catch (e1) { /* a locked/derived tab is skipped, not fatal */ }
    });
    return stamp;
  } catch (e) {
    return "";
  }
}

/* ==========================================================================
 * 07 · PERIOD ANALYSIS (weekly & date-wise)                               *
 * ======================================================================== */

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
  // one soft colour per revenue type so the groups read apart.
  var WCLASS = { service: "wt-svc", accessory: "wt-acc", repair: "wt-rep", grossProfit: "wt-gp" };
  TYPES.forEach(function (t) { t.wc = WCLASS[t.key]; });

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

/** TASK 36-D: DECK numbers. Everything is CURRENT MONTH only (the owner keeps
 *  previous months manually). crm!TranType == "SALESRETURN" is excluded at the
 *  parse boundary — every valuation below runs on isSaleRow_ rows only. */
function computeDeckAnalysis_(ss, crmRows) {
  var out = { available: true, monthLabel: "" };
  var now = new Date();
  var curY = now.getFullYear(), curM = now.getMonth() + 1;
  var MNAME = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];
  out.monthLabel = MNAME[now.getMonth()] + "-" + curY;
  var curMKey = curY * 100 + curM;                        // 2026*100+9 → 202609
  function mkeyTs(ts) { var d = new Date(ts); return d.getFullYear() * 100 + (d.getMonth() + 1); }

  // ---- CRM rows for the current month, TranType == SALES only ----
  var saleAll = (crmRows || []).filter(isSaleRow_);
  var sale = saleAll.filter(function (r) { return r.year === curY && r.month === curM; });
  // The crm sheet normally holds exactly the running month; if the calendar
  // rolled over before the sheet was swapped, keep the sheet's rows rather
  // than rendering an empty deck.
  if (!sale.length) sale = saleAll;

  function mkType(name) { return { name: name, count: 0, value: 0, gp: 0, invoices: {} }; }
  var rb = { repair: mkType("Repair Revenue"), accessory: mkType("Accessories Revenue"), service: mkType("Service Revenue") };
  var billSet = {};
  sale.forEach(function (r) {
    var k = revenueTypeKey_(r.revenueType);
    if (!k) return;
    var b = rb[k];
    b.count++;                                    // invoice lines — AR+SR+RR sums to Bills
    b.value += r.taxableAmount;
    b.gp += r.grossProfit;
    var inv = str_(r.invoiceCode);
    if (inv) { billSet[inv] = true; b.invoices[inv] = true; }
  });
  function uniqKeys(set) { return Object.keys(set).length; }

  // 1) Sales Revenue Analysis — revenue type, count, value.
  out.salesRev = {
    rows: ["repair", "accessory", "service"].map(function (k) {
      var b = rb[k];
      return { label: b.name, count: b.count, value: Math.round(b.value) };
    }),
  };
  out.salesRev.totalCount = out.salesRev.rows.reduce(function (s, r) { return s + r.count; }, 0);
  out.salesRev.totalValue = out.salesRev.rows.reduce(function (s, r) { return s + r.value; }, 0);

  // 2) Walk-in details — unique walkins (walkin!MobileNo) + unique invoices (crm!InvoiceCode)
  var wk = countWalkins_(ss);
  out.walkins = { unique: wk.unique, uniqBills: uniqKeys(billSet) };

  // ---- service cases: dedupe per CaseId (latest CaseDate wins), current month ----
  var serviceSheet = getSheetByNames_(ss, CONFIG.SERVICE_SHEET_NAMES);
  var caseMonth = [];
  if (serviceSheet) {
    var read = readSheetRows_(serviceSheet, SERVICE_SPEC, ["caseDate"]);
    var best = {}, order = [];
    read.rows.forEach(function (r, idx) {
      var id = str_(r.caseId) || ("_row" + idx);
      var pd = parseAnyDate_(r.caseDate);
      var ts = pd.valid ? pd.ts : 0;
      if (!best[id]) { best[id] = { row: r, ts: ts }; order.push(id); }
      else if (ts > best[id].ts) best[id] = { row: r, ts: ts };
    });
    var allCases = [];
    order.forEach(function (id) {
      var b = best[id];
      allCases.push(b.row);
      if (b.ts && mkeyTs(b.ts) === curMKey) caseMonth.push(b.row);
    });
    // at month change the service export still holds last month's
    // cases — mirror the CRM fallback above so the two case-driven tables
    // (Repair Revenue & Coverage) never sit at zero on a stale sheet.
    if (!caseMonth.length) caseMonth = allCases;
  }

  function deckFamily_(prod) {
    var t = str_(prod).toLowerCase();
    if (/iphone|\bvin\b/.test(t)) return "iPhone";
    if (/\bmac|imac|macbook/.test(t)) return "Mac";
    if (/ipad|airpods?/.test(t)) return "iPad / Airpods";
    return "Others";
  }
  var FAMS = ["iPhone", "Mac", "iPad / Airpods", "Others"];

  // 3) Repair Revenue Analysis — count = unique service!SerialNo1 per product
  //    family; value = crm!TaxableAmount matched on service!CaseId ↔ crm!OrderCode.
  var covCases = caseMonth;                                     // coverage keeps every case type
  var repairCases = caseMonth.filter(function (r) { return !isSNRType_(str_(r.repairDocType)); });
  var famSerial = {}, famValue = {};
  FAMS.forEach(function (f) { famSerial[f] = {}; famValue[f] = 0; });
  var caseFamily = {};                                          // service!CaseId → family bucket
  repairCases.forEach(function (r) {
    var f = deckFamily_(r.product);
    var sn = str_(r.serialNo1);
    var cid = str_(r.caseId);
    famSerial[f][sn || ("#" + cid)] = true;                     // unique count, deduped
  });
  // the value join must run over the SAME population the count is
  // taken from (the month-window repair set; caseMonth already carries the
  // stale-sheet fallback). Linking to a wider set printed values beside a
  // zero count — the owner's "0 counts but value shows" bug.
  var linkCases = repairCases;
  linkCases.forEach(function (r) {
    var cid = str_(r.caseId);
    if (cid && !caseFamily[cid]) caseFamily[cid] = deckFamily_(r.product);
  });
  sale.forEach(function (r) {
    var oc = str_(r.orderCode);
    if (oc && caseFamily[oc]) famValue[caseFamily[oc]] += r.taxableAmount;
  });
  out.repairRev = {
    rows: FAMS.map(function (f) {
      return { label: f, count: uniqKeys(famSerial[f]), value: Math.round(famValue[f]) };
    }),
  };
  out.repairRev.totalCount = out.repairRev.rows.reduce(function (s, r) { return s + r.count; }, 0);
  out.repairRev.totalValue = out.repairRev.rows.reduce(function (s, r) { return s + r.value; }, 0);

  // 4) Accessories Revenue Analysis — crm!ItemDescription buckets.
  //    Owner's wildcard spellings kept verbatim (tg / temper(ed) / temepered /
  //    tamper(ed) glass; adapter|cable; AppleCare+/AC+; protect+; amc).
  var ACC_DEFS = [
    ["Tempered Glass", /(\btg\b|temper|temeper|tamper)/i],
    ["Case", /case/i],
    ["Adapter & Cable", /adapter|cable/i],
    ["AC+", /apple\s*care\s*\+|care\s*\+|\bac\s*\+/i],
    ["Protect Plus", /protect\s*\+|protect\s*plus/i],
    ["AMC Charge", /\bamc\b/i],
  ];
  var acc = ACC_DEFS.map(function (d) { return { label: d[0], count: 0, value: 0 }; });
  acc.push({ label: "Others", count: 0, value: 0 });
  sale.forEach(function (r) {
    if (revenueTypeKey_(r.revenueType) !== "accessory") return;
    var d = str_(r.itemDescription);
    var bucket = acc[acc.length - 1];
    for (var i = 0; i < ACC_DEFS.length; i++) {
      if (ACC_DEFS[i][1].test(d)) { bucket = acc[i]; break; }
    }
    bucket.count++;
    bucket.value += r.taxableAmount;
  });
  out.accessoriesRev = {
    rows: acc,
    totalCount: acc.reduce(function (s, r) { return s + r.count; }, 0),
    totalValue: acc.reduce(function (s, r) { return s + r.value; }, 0),
  };

  // 5) Service Revenue Analysis — pivot on crm!ItemCode, Service rows only.
  var svcPivot = {};
  sale.forEach(function (r) {
    if (revenueTypeKey_(r.revenueType) !== "service") return;
    var code = str_(r.itemCode) || "(blank item code)";
    if (!svcPivot[code]) svcPivot[code] = { label: code, count: 0, value: 0 };
    svcPivot[code].count++;
    svcPivot[code].value += r.taxableAmount;
  });
  out.serviceRev = {
    rows: Object.keys(svcPivot).sort().map(function (k) {
      var b = svcPivot[k];
      return { label: b.label, count: b.count, value: Math.round(b.value) };
    }),
  };
  out.serviceRev.totalCount = out.serviceRev.rows.reduce(function (s, r) { return s + r.count; }, 0);
  out.serviceRev.totalValue = out.serviceRev.rows.reduce(function (s, r) { return s + r.value; }, 0);

  // 6) Gross Profit Analysis — Value: invoice value, GP value, GP%.
  out.gpValue = {
    rows: ["accessory", "repair", "service"].map(function (k) {
      var b = rb[k];
      return { label: b.name, value: Math.round(b.value), gp: Math.round(b.gp),
        pct: b.value > 0 ? Math.round((1000 * b.gp) / b.value) / 10 : 0 };
    }),
  };
  var tv = out.gpValue.rows.reduce(function (s, r) { return s + r.value; }, 0);
  var tg = out.gpValue.rows.reduce(function (s, r) { return s + r.gp; }, 0);
  out.gpValue.total = { value: tv, gp: tg, pct: tv > 0 ? Math.round((1000 * tg) / tv) / 10 : 0 };

  // 7) Walk-in Analysis — unique walkins from walkin!MobileNo; unique bills from
  //    crm!InvoiceCode; Bills = total invoice lines; AR/RR/SR = unique invoice
  //    count per crm!RevenueType; % = unique bills ÷ walk-ins; Enquiry =
  //    walk-ins − unique bills.
  var walkins = out.walkins.unique;
  var uniqBills = out.walkins.uniqBills;
  out.walkinAnalysis = [
    { label: "No. Of Walkin", value: walkins },
    { label: "Unique Bill", value: uniqBills },
    { label: "Bills", value: sale.length },
    { label: "AR", value: uniqKeys(rb.accessory.invoices) },
    { label: "RR", value: uniqKeys(rb.repair.invoices) },
    { label: "SR", value: uniqKeys(rb.service.invoices) },
    { label: "Walkin to Bills %", value: walkins > 0 ? Math.round((100 * uniqBills) / walkins) + "%" : "—" },
    { label: "Enquiry", value: walkins - uniqBills },
  ];

  // 8) Coverage Analysis | LOB (Task 40 spec, Task-41 W/NW fix) — family from
  //    service!ProductDescription (iPhone / Mac / Watch / iPad / AirPods /
  //    Others). Battery & display are split off the iPhone row via text in
  //    service!IssueDescription (was PartDescription). The W/NW split reads
  //    service!ProductWarrantyDesc first and falls back to
  //    service!CoverageCodeDescription when that cell is blank:
  //    "in / product / consumer / limited warranty", applecare or coverage
  //    wording (and no out/no/not/non/expired/void marker) → W; anything
  //    else, including blank, → NW.
  function covBucket_(prod, issue) {
    var t = str_(prod).toLowerCase();
    var p = str_(issue).toLowerCase();
    var iphone = /iphone|\bvin\b/.test(t);
    if (iphone && /batt/.test(p)) return "iPhone Battery";
    if (iphone && /display|screen|lcd/.test(p)) return "iPhone Display";
    if (iphone) return "iPhone";
    if (/\bmac|imac|macbook/.test(t)) return "Mac";
    if (/watch/.test(t)) return "Watch";
    if (/ipad/.test(t)) return "iPad";
    if (/airpods?/.test(t)) return "AirPods";
    return "Others";
  }
  var COV_ORDER = ["iPhone", "iPhone Battery", "iPhone Display", "Mac", "Watch", "iPad", "AirPods", "Others"];
  var cov = {};
  COV_ORDER.forEach(function (n) { cov[n] = { label: n, count: 0, w: 0, nw: 0 }; });
  covCases.forEach(function (r) {
    var c = cov[covBucket_(r.product, r.issue)];
    c.count++;
    // product-warranty text first, legacy coverage-code text as the
    // fallback when the product column is empty — mixed exports under-counted
    // the W side before.
    var wtxt = (str_(r.productWarranty) || str_(r.warranty)).toLowerCase();
    var outMark = /\bout\b|\bow\b|\bno\b|\bnon\b|\bnot\b|expired|void/.test(wtxt);
    if (!outMark && (/warrant|coverage|applecare|\biw\b/.test(wtxt))) c.w++;
    else c.nw++;
  });
  out.coverage = { rows: COV_ORDER.map(function (n) { return cov[n]; }) };
  out.coverage.total = {
    count: out.coverage.rows.reduce(function (s, r) { return s + r.count; }, 0),
    w: out.coverage.rows.reduce(function (s, r) { return s + r.w; }, 0),
    nw: out.coverage.rows.reduce(function (s, r) { return s + r.nw; }, 0),
  };

  return out;
}

/* ==========================================================================
 * 08 · RENDER PRIMITIVES                                                  *
 * ======================================================================== */

function pctBadgeHtml_(pctVal) {
  var cls = pctVal >= 100 ? "good" : pctVal >= 70 ? "info" : pctVal >= 40 ? "warn" : "bad";
  return '<span class="badge ' + cls + '">' + pctVal.toFixed(1) + "%</span>";
}

function barHtml_(pctVal) {
  var clamped = Math.max(0, Math.min(100, pctVal));
  var cls = pctVal >= 100 ? "good" : pctVal >= 70 ? "info" : pctVal >= 40 ? "warn" : "bad";
  return '<div class="bar"><div class="bar-fill ' + cls + '" style="width:' + clamped + '%"></div></div>';
}

/** Task 47 — group banner separating the master's A-G groups on the SM page. */
function roleBanner_(letter, name) {
  return '<div class="group-banner"><span class="gb-letter">' + esc_(letter) + "</span>" +
    '<span class="gb-name">' + esc_(name) + "</span></div>";
}

function emptyPerf_(sheetName, msg) {
  return '<div class="empty-perf">📁 No <b>' + esc_(sheetName) + '</b> sheet was found in this spreadsheet.<br>' +
    '<span class="muted">' + esc_(msg) + "</span></div>";
}

/** The "Updated ... · N rows · N invoices" line under the title. */
function renderMetaText_(data) {
  var generated = Utilities.formatDate(data.generatedAt, Session.getScriptTimeZone(), "dd MMM yyyy, HH:mm");
  return "Updated " + generated + " · " + data.store.totalInvoices +
    " SALES rows · " + data.store.uniqueInvoices + " unique invoices";
}

/* ==========================================================================
 * 09 · RENDER — STORE & SALES BLOCKS                                      *
 * ======================================================================== */

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
    '<section class="section" id="sec-store"><h2>A1. 🎯 Store Target — Performance ' + exportSectionBtn_("store") + '</h2><div class="body">' +
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
    '<section class="section" id="sec-workdays"><h2>A2. 📅 Working Days (Current Month)</h2><div class="body">' +
    '<div class="sec-desc">Definition: days the store is open this month — total, already consumed, and how many selling days are left. Holiday dates come from CONFIG.</div>' +
    '<div class="cards" style="margin-bottom:0">' +
    '<div class="card"><div class="label">Total Working Days</div><div class="value">' + s.totalWorkingDays + "</div></div>" +
    '<div class="card"><div class="label">Consumed</div><div class="value">' + s.consumedWorkingDays + "</div></div>" +
    '<div class="card"><div class="label">Balance</div><div class="value" style="color:' +
    (s.balanceWorkingDays > 0 ? "#d97706" : "#16a34a") + '">' + s.balanceWorkingDays + "</div></div>" +
    "</div></div></section>"
  );
}

/** Task 49 — INDEX weekly summary: percentage view only.
 *  Completed weeks → achieved % of that week's target per revenue type;
 *  in-progress week → current % so far; upcoming weeks → the % of the MONTH
 *  target each week must still contribute (remaining balance split evenly). */
function renderWeeklySummary_(data) {
  var wa = computeWeeklyAnalysis_(data);
  function r1(x) { return Math.round(x * 10) / 10; }
  var upcoming = wa.weeks.filter(function (w) { return w.status === "Upcoming"; });
  var nUp = upcoming.length;
  var upPct = {};
  wa.types.forEach(function (t) {
    if (!nUp) { upPct[t.key] = null; return; }
    var bal = Math.max(0, t.monthTarget - (wa.monthTotals[t.key] || 0));
    upPct[t.key] = r1(((bal / nUp) / Math.max(1, t.monthTarget)) * 100);
  });
  function typeCell(w, t) {
    if (w.status === "Upcoming") {
      return '<td class="num muted">' + (upPct[t.key] == null ? "—" : upPct[t.key] + "%") + "</td>";
    }
    if (!w.tgt[t.key]) return '<td class="num muted">—</td>';
    var p = r1((w.achieved[t.key] / w.tgt[t.key]) * 100);
    var cls = p >= 100 ? "pos" : (w.status === "Completed" ? "neg" : "");
    return '<td class="num' + (cls ? " " + cls : "") + '">' + p + "%</td>";
  }
  var rows = wa.weeks.map(function (w) {
    var tag = w.status === "Completed" ? '<span class="pill pill-positive">completed</span>'
      : w.status === "In progress" ? '<span class="pill pill-benchmark">in progress</span>'
      : '<span class="pill">upcoming</span>';
    return "<tr><td><b>Week " + w.idx + '</b> <span class="muted">(' + w.start + "–" + w.end + ")</span></td>" +
      "<td>" + tag + "</td>" + wa.types.map(function (t) { return typeCell(w, t); }).join("") + "</tr>";
  }).join("");
  return '<section class="section" id="idx-weekly"><h2>A6. 🗓 Weekly Revenue Analysis — ' + esc_(wa.monthName) + "</h2>" +
    '<div class="body"><div class="sec-desc">completed weeks: achieved % of each week&rsquo;s target · in-progress week: current % so far · ' +
    "upcoming weeks: % of the month target each week must still contribute (balance split evenly) to close the month at 100%.</div>" +
    '<div class="scroll-x"><table class="ptable"><thead><tr><th>Week</th><th>Status</th>' +
    wa.types.map(function (t) { return '<th class="num">' + esc_(t.label) + "</th>"; }).join("") +
    "</tr></thead><tbody>" + rows + "</tbody></table></div></div></section>";
}

/** Renders section 7 — weekly achieved vs weekly target per revenue type. */
function renderWeeklyAnalysis_(data) {
  var wa = computeWeeklyAnalysis_(data);
  var TYPES = wa.types;

  var head1 = '<tr><th rowspan="2" class="sticky-col">Week</th><th rowspan="2">Wrk Days</th><th rowspan="2">Status</th>';
  var head2 = "<tr>";
  TYPES.forEach(function (t, i) {
    var wc = " " + t.wc;
    head1 += '<th colspan="3" class="grp grp-' + (i % 5) + wc + '">' + esc_(t.label) + "</th>";
    head2 += '<th class="grp grp-' + (i % 5) + wc + '">Achieved</th>' +
      '<th class="grp grp-' + (i % 5) + wc + '">Week Target</th>' +
      '<th class="grp grp-' + (i % 5) + wc + '">Balance</th>';
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
      var wc = " " + t.wc;
      var ach = w.achieved[tk], tgt = w.tgt[tk];
      var bal = Math.max(0, tgt - ach);
      var pill = tgt > 0
        ? '<span class="apill ' + (ach >= tgt ? "ok" : "no") + '">' + money_(ach) + "</span>"
        : '<span class="muted">' + money_(ach) + "</span>";
      tr += '<td class="cell-' + (i % 5) + wc + '">' + pill + "</td>" +
        '<td class="cell-' + (i % 5) + wc + '"><span class="tgt">' + money_(tgt) + "</span></td>" +
        '<td class="cell-' + (i % 5) + wc + '">' + (bal > 0
          ? '<span class="muted">' + money_(bal) + "</span>"
          : '<span class="apill ok">done</span>') + "</td>";
    });
    body += tr + "</tr>";
  });

  var foot = '<tr class="total-row"><td class="sticky-col">Month</td><td style="text-align:center">' +
    wa.totalWorkingDays + "</td><td></td>";
  TYPES.forEach(function (t, i) {
    var tk = t.key;
    var wc = " " + t.wc;
    var ach = wa.monthTotals[tk], tgt = t.monthTarget;
    var bal = Math.max(0, tgt - ach);
    foot += '<td class="cell-' + (i % 5) + wc + '"><b>' + money_(ach) + "</b></td>" +
      '<td class="cell-' + (i % 5) + wc + '"><span class="tgt">' + money_(tgt) + "</span></td>" +
      '<td class="cell-' + (i % 5) + wc + '">' + money_(bal) + "</td>";
  });
  foot += "</tr>";

  return (
    '<section class="section" id="sec-weekly"><h2>A6. 🗓 Weekly Revenue Analysis — ' +
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
  // every salesperson block alternates between a light and a dark
  // shade (sp-a / sp-b) so neighbouring people are visually separated.
  people.forEach(function (p, i) {
    var sp = " sp-" + (i % 2 ? "b" : "a");
    head1 += '<th colspan="' + TYPES.length + '" class="grp grp-' + (i % 5) + sp + '">' + esc_(p.name) + "</th>";
    TYPES.forEach(function (pair) {
      head2 += '<th class="grp grp-' + (i % 5) + sp + '">' + esc_(pair[1]) + "</th>";
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
        tr += '<td class="cell-' + (i % 5) + (i % 2 ? " sp-b" : " sp-a") + '">' + figure +
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
      foot += '<td class="cell-' + (i % 5) + (i % 2 ? " sp-b" : " sp-a") + '">' + money_(p[tk].achieved) +
        '<span class="tgt">/ ' + money_(p[tk].target) + "</span></td>";
    });
  });
  foot += "</tr>";

  // ---- per-person daily target legend ----
  var legend = people.map(function (p, i) {
    return '<div class="legend-card cell-' + (i % 5) + (i % 2 ? " sp-b" : " sp-a") + '">' +
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
    '<section class="section" id="sec-datewise"><h2>F1. 📆 Date-wise Sales Revenue Analysis (by salesperson)</h2><div class="body">' +
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
    '<section class="section" id="sec-kpi"><h2>B. 📊 KPI Target — Store KPI Parameters ' + exportSectionBtn_("kpi") + '</h2><div class="body">' +
    '<div class="scroll-x">' +
    "<table><thead><tr><th>KPI</th><th>Type</th><th>Target</th><th>Achieved</th><th>Achieved %</th>" +
    "<th>Balance</th><th>Approx Achievable Value</th><th>Achieved Value</th><th>Achieved GP</th></tr></thead><tbody>" + rows +
    "</tbody></table></div>" +
    '<p class="muted" style="font-size:12.5px;margin:12px 0 0">Targets read from <b>target!B12:L14</b>. ' +
    "<b>Approx Achievable Value</b> = the balance count still needed × that KPI’s approximate unit value in <b>target!B15:L15</b> — the amount still winnable to close the balance target.</p>" +
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
    // the Acc-AMC cell must answer "how many more to close the next
    // slab" (slab card at target!A29:F33), not just restate the raw count.
    return "<tr><td>" + esc_(a.name) + "</td><td>" + a.amcCount + "</td><td>" + money_(a.amcEligible) +
      "</td><td>" + (a.accessoryNextSlab ? a.accessoryRemaining + " more (slab " + a.accessoryNextSlab.accCount + ")" : "top slab") +
      "</td><td>" + money_(a.accessoryEligible) + "</td><td>" +
      (a.amcNextSlab ? a.amcRemaining + " more" : "top slab") + '</td><td class="pos">' + money_(a.projectedIncentive) +
      "</td></tr>";
  }).join("");
  var amcFoot = '<tr class="total-row"><td>All salespersons</td><td>' + amcTot.cnt + "</td><td>" +
    money_(amcTot.elig) + "</td><td></td><td>" + money_(amcTot.accElig) +
    "<td></td>" + '<td class="pos">' + money_(amcTot.inc) + "</td></tr>";

  // ---- Category 2: Bundle (DISC scheme) ----
  var bunTot = { cnt: 0, inc: 0 };
  // bundle rows also carry the gap to the next slab (minimum
  // incentive tier) — the owner wants the count needed, not just the count.
  var bunRows = data.amc.map(function (a) {
    var b = bundleByKey[a.key] || { count: 0, projectedIncentive: 0 };
    bunTot.cnt += b.count; bunTot.inc += b.projectedIncentive;
    return "<tr><td>" + esc_(a.name) + "</td><td>" + b.count + "</td><td>" +
      (b.nextSlab ? b.remaining + " more (slab " + b.nextSlab.min + ")" : "top slab") +
      '</td><td class="pos">' + money_(b.projectedIncentive) + "</td></tr>";
  }).join("");
  var bunFoot = '<tr class="total-row"><td>All salespersons</td><td>' + bunTot.cnt +
    '</td><td></td><td class="pos">' + money_(bunTot.inc) + "</td></tr>";

  return (
    '<section class="section" id="sec-incentive"><h2>B1. 💰 Sales Person Incentive — AMC &amp; Bundle</h2><div class="body">' +
    '<div class="inc-pair">' +
    '<div class="inc-cat"><div class="inc-cat-title">Category 1 · AMC</div>' +
    '<div class="scroll-x"><table><thead><tr><th>Salesperson</th><th>AMC Count</th><th>AMC Eligible</th>' +
    "<th>Need To Close Next Slab</th><th>Acc-AMC Eligible</th><th>AMC To Next</th><th>AMC Incentive</th></tr></thead><tbody>" +
    amcRows + amcFoot + "</tbody></table></div></div>" +
    '<div class="inc-cat"><div class="inc-cat-title">Category 2 · Bundle (DISC scheme)</div>' +
    '<div class="scroll-x"><table><thead><tr><th>Salesperson</th><th>Bundle Count</th><th>Bundle To Next</th><th>Bundle Incentive</th></tr></thead><tbody>' +
    bunRows + bunFoot + "</tbody></table></div></div></div>" +
    "</div></section>"
  );
}

function renderSalespersonCards_(data) {
  // Incentive-bearing revenue types show target vs achieved, the balance
  // still needed in (brackets), the incentive that type has earned, and the
  // next chase tier of the Task-33 IFS formula (70s tier &rarr; 110% top).
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
    return '<section class="section" id="sec-person"><h2>B2. 👤 Sales Person Performance</h2><div class="body">' +
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
    '<section class="section" id="sec-person"><h2>B2. 👤 Sales Person Performance</h2><div class="body">' +
    '<div class="person-grid">' + cards + "</div>" +
    '<p class="muted" style="font-size:12.5px;margin:14px 0 0">' +
    "Only <b>Service</b> and <b>Accessory</b> revenue earn incentive — payout is the pro-rata IFS formula. Service tiers 71/81/91/100&#8209;&#62;110% (&uarr;to &#8377;12,000); Accessory tiers 51/61/71/81&#8209;&#62;110% (&uarr;to &#8377;8,000). " +
    "<b>Repair</b> revenue pays nothing on its own — it is one shared store gate: once the store&rsquo;s combined repair achievement passes <b>" +
    CONFIG.REPAIR_INCENTIVE_GATE_PCT + "%</b> of the store repair target, <b>every salesperson is eligible to receive the incentive they earned</b>.</p>" +
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
    '<section class="section" id="sec-ar"><h2>B3. 🔗 AR Attachment — accessory sold with service / repair</h2><div class="body">' +
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

function renderInvoiceValuation_(data) {
  var rows = data.revenueTypes.map(function (rt) {
    return "<tr><td>" + esc_(rt.label) + "</td><td>" + rt.invoiceLines + "</td><td>" + rt.uniqueInvoices +
      "</td><td>" + money_(rt.revenue) + "</td><td>" + money_(rt.gp) + "</td></tr>";
  }).join("");
  var totalRow = "<tr class=\"total-row\"><td>Overall</td><td>" + data.invoiceValuation.totalInvoiceLines +
    "</td><td>" + data.invoiceValuation.totalUniqueInvoices + "</td><td>" + money_(data.store.total.achieved) +
    "</td><td>" + money_(data.store.grossProfit.achieved) + "</td></tr>";
  return (
    '<section class="section" id="sec-invoice"><h2>A3. 🧾 Invoice Valuation</h2><div class="body">' +
    '<div class="sec-desc">Definition: what invoices are really worth — line count, unique invoice count, revenue and gross profit per revenue type, plus AR (accessory attachment) rate next to invoices that qualified for one.</div>' +
    '<div class="iv-split">' +
    '<div class="iv-left"><div class="scroll-x">' +
    "<table class=\"hdr-deep\"><thead><tr><th>Revenue Type</th><th>Invoice Lines</th><th>Unique Invoices</th><th>Revenue</th>" +
    "<th>Gross Profit</th></tr></thead><tbody>" + rows + totalRow + "</tbody></table></div></div>" +
    '<div class="iv-right">' +
    '<div class="card"><div class="label">Bundle Unique Invoices</div><div class="value">' +
    data.invoiceValuation.bundleUniqueInvoices + "</div></div>" +
    '<div class="card"><div class="label">AMC Unique Invoices</div><div class="value">' +
    data.invoiceValuation.amcUniqueInvoices + "</div></div>" +
    "</div>" +
    "</div>" +
    "</div></section>"
  );
}

/** GP page: Store target · Weekly Revenue Analysis · Invoice valuation. */
function renderGpHtml_(data) {
  var body = "";
  body += renderStoreCards_(data);
  body += renderStoreTable_(data);
  body += renderWorkingDaysCard_(data);
  body += renderInvoiceValuation_(data);
  body += renderGpExtras_(data.gpExtras);
  body += renderWeeklyAnalysis_(data);
  return body;
}

/** Task 42 — the two GP tiles: high/low margin + discounted cases. */
function renderGpExtras_(x) {
  x = x || {};
  var discount = x.discount || { count: 0, value: 0, gp: 0, pct: 0, byType: {}, byTypeValue: {} };
  var hi = x.marginHigh || [], lo = x.marginLow || [];
  function money0(v) { return money_(v || 0); }

  function marginTable(rows) {
    var body = rows.map(function (m) {
      // invoice number back under the short part (owner wants both);
      // the cell carries a tooltip since the column is width-capped.
      return '<td title="' + esc_(m.partLabel + " · " + m.invoice) + '"><b>' + esc_(m.partLabel) +
        '</b><span class="cell-sub">' + esc_(m.invoice) + "</span></td>" +
        "<td>" + esc_(m.typeLabel) + "</td>" +
        '<td class="num">' + money0(m.taxable) + "</td>" +
        '<td class="num"><span class="' + (m.gp >= 0 ? "pos" : "neg") + '">' + money0(m.gp) + "</span></td>" +
        '<td class="num"><span class="' + (m.margin >= 0 ? "pos" : "neg") + '">' + m.margin.toFixed(1) + "%</span></td></tr>";
    }).join("");
    if (!rows.length) body = '<tr><td colspan="5"><span class="muted">no invoice valuation rows</span></td></tr>';
    return '<table class="ptable"><thead><tr><th>Part</th><th>Revenue Type</th>' +
      '<th class="num">Taxable</th><th class="num">Gross Profit</th><th class="num">Margin %</th></tr></thead>' +
      "<tbody>" + body + "</tbody></table>";
  }

  var TYPE_LBL = [["repair", "Repair Revenue"], ["accessory", "Accessory Revenue"], ["service", "Service Revenue"]];
  var dRows = TYPE_LBL.map(function (t) {
    var top = discount.byTypeTop && discount.byTypeTop[t[0]];
    return "<tr><td>" + t[1] + '</td><td class="num">' + (discount.byType[t[0]] || 0) +
      '</td><td class="num">' + money0(discount.byTypeValue[t[0]]) +
      "</td><td>" + (top ? esc_(top.label) + (top.n > 1 ? " ×" + top.n : "") : '<span class="muted">—</span>') + "</td></tr>";
  }).join("");
  var dTotal = '<tr class="totrow"><td><b>Total</b></td><td class="num"><b>' + discount.count +
    '</b></td><td class="num"><b>' + money0(discount.value) + "</b></td><td></td></tr>";

  var marginPanel =
    '<section class="section" id="sec-gpmargin"><h2>A4. 📈 High / Low Margin ' + exportSectionBtn_("margin") + '</h2><div class="body">' +
    '<div class="sec-desc">margin % = crm!GrossProfit ÷ crm!TaxableAmount per unique invoice (SALES only; ' +
    "zero-base invoices skipped). Top 5 at each end, across all revenue types.</div>" +
    '<div class="hl-pair">' +
    '<div class="hl-half"><div class="subpanel-head">Top 5 — High Margin</div><div class="scroll-x">' + marginTable(hi) + "</div></div>" +
    '<div class="hl-half"><div class="subpanel-head">Top 5 — Low Margin</div><div class="scroll-x">' + marginTable(lo) + "</div></div>" +
    "</div></div></section>";

  var discPanel =
    '<section class="section" id="sec-discount"><h2>A5. 🏷 Discounted Cases</h2><div class="body">' +
    '<div class="sec-desc">crm!DiscountAmount &gt; 0 rows (SALES only); cases counted once per unique invoice code.</div>' +
    '<div class="cards" style="margin-bottom:12px">' +
    '<div class="card"><div class="label">Discounted Invoices</div><div class="value">' + discount.count +
    '</div><div class="sub">unique invoice codes</div></div>' +
    '<div class="card"><div class="label">Total Discount Value</div><div class="value neg">' + money0(discount.value) +
    '</div><div class="sub">Σ crm!DiscountAmount</div></div>' +
    '<div class="card"><div class="label">Discount ÷ Gross Profit</div><div class="value">' + discount.pct +
    '%</div><div class="sub">vs GP of the discounted lines</div></div>' +
    "</div>" +
    '<div class="scroll-x"><table class="ptable"><thead><tr><th>Revenue Type</th><th class="num">Cases</th>' +
    '<th class="num">Discount Value</th><th>Top Part</th></tr></thead><tbody>' +
    dRows + dTotal + "</tbody></table></div>" +
    "</div></section>";

  // stacked vertically (each panel gets the full page width) with
  // High / Low Margin ABOVE Discounted Cases — no more side-by-side pair.
  return '<div class="gp-stack">' + marginPanel + discPanel + "</div>";
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

/* ==========================================================================
 * 10 · RENDER — PROCESS BLOCKS                                            *
 * ======================================================================== */

/** WALK-IN page: walk-in analysis + token creation side by side. */
function renderWalkinHtml_(pd) {
  var wl = pd.walkin;
  var head =
    '<section class="section" id="sec-walkin"><h2>C. 🚶 Walk-in Analysis &amp; Token Creation</h2><div class="body">' +
    '<div class="sec-desc">source: ' + esc_(wl.sheetName) +
    ". Every row is one walk-in — MobileNo identifies the visitor (Unique = distinct mobiles); a row with a CaseId counts as a case, a row with an InvoiceCode as a completed sale. The token table pins which user created each token.</div>" +
    '<div class="walkin-pair">';

  // Matching heads keep 🚶 Walk-in Analysis and 🎟 Token Creation aligned on
  // ONE line; the wide day-wise table gets the full page width below the pair.
  var left =
    '<div class="wp-left"><div class="wp-head">C1. 🚶 Walk-in Analysis</div>' +
    '<p class="sec-desc" style="margin:-2px 0 10px">month-to-date walk-in summary — visitors, cases, invoices, remarks and dead leads.</p>' +
    renderWalkinAnalysis_(wl) + "</div>";
  var mid =
    '<div class="wp-mid"><div class="wp-head">C2. 🧭 Customer Source Analysis</div>' +
    '<p class="sec-desc" style="margin:-2px 0 10px">pivot on walkin!CustomerVisitSource against the total number of walk-ins.</p>' +
    renderCustomerSource_(wl) + "</div>";
  var right =
    '<div class="wp-right"><div class="wp-head">C3. 🎟 Token Creation</div>' +
    '<p class="sec-desc" style="margin:-2px 0 10px">how many walk-in tokens each user created, and their share of the month&rsquo;s tokens.</p>' +
    renderTokenCreation_(wl) + "</div>";

  return head + left + mid + right + "</div>" +
    '<div class="wp-daytable"><div class="wp-head">C4. 📅 Day-wise Walk-in &amp; Token Details</div>' +
    renderWalkinDayTable_(wl) + "</div></div></section>";
}

function renderWalkinAnalysis_(wl) {
  if (!wl.available) {
    return emptyPerf_(CONFIG.WALKIN_SHEET_NAMES[0],
      "Add a tab named `" + CONFIG.WALKIN_SHEET_NAMES[0] + "` with columns MobileNo, TokenDate, VisitPurposeDesc, CaseId, InvoiceCode and **Remark.");
  }
  var t = wl.totals;
  // Summary rendered as ONE table (was a row of tiles) — one row per metric.
  // the free-text "Reading" column becomes the metric's % share of
  // the total walk-in tokens (one decimal).
  function contrib(n) {
    return '<td class="num">' + (t.walkins > 0 ? Math.round((1000 * n) / t.walkins) / 10 : 0) + "%</td>";
  }
  var head =
    '<div class="scroll-x" style="margin-bottom:14px"><table class="ptable walksum">' +
    "<thead><tr><th>Metric</th>" + '<th class="num">Value</th>' +
    '<th class="num">% Contribution</th></tr></thead><tbody>' +
    "<tr><td><b>Total Walk-ins</b></td>" + '<td class="num"><b>' + t.walkins + "</b></td>" +
      contrib(t.walkins) + "</tr>" +
    "<tr><td><b>Unique Walk-ins</b></td>" + '<td class="num"><b>' + t.unique + "</b></td>" +
      contrib(t.unique) + "</tr>" +
    "<tr><td><b>Cases Created</b></td>" + '<td class="num"><b>' + t.cases + "</b></td>" +
      contrib(t.cases) + "</tr>" +
    "<tr><td><b>Invoices Created</b></td>" + '<td class="num"><b>' + t.invoices + "</b></td>" +
      contrib(t.invoices) + "</tr>" +
    "<tr><td><b>Remarks Updated</b></td>" + '<td class="num"><b>' + t.remarks + "</b></td>" +
      contrib(t.remarks) + "</tr>" +
    '<tr class="deadsum"><td><b>Dead Leads</b></td>' + '<td class="num dead"><b>' + t.deadLeads + "</b></td>" +
      contrib(t.deadLeads) + "</tr>" +
    '</tbody></table></div>' +
    '<p class="muted" style="font-size:11.5px;margin:-8px 0 10px 2px">% Contribution = metric value ÷ total walk-in tokens of the month.</p>';

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
  var rows = wl.days.map(function (d, i) {
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
    return "<tr" + (d.undated ? ' class="undated-row"' : "") + ">" +
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

/** Task 37: customer source analysis — pivot walkin!CustomerVisitSource. */
function renderCustomerSource_(wl) {
  if (!wl.available) return "";
  if (!wl.hasSource) {
    return '<div class="subnote">No customer-source column was detected in the ' +
      "<b>" + esc_(wl.sheetName) + "</b> sheet yet. Add a column named <b>CustomerVisitSource</b> " +
      "(e.g. Walk-by / Google / Reference) and the pivot appears here automatically.</div>";
  }
  var t = wl.totals;
  var rows = (wl.sources || []).map(function (s) {
    return "<tr>" +
      "<td><b>" + esc_(s.name) + "</b></td>" +
      '<td class="num">' + s.count + "</td>" +
      '<td class="num">' + s.share + "%</td>" +
      "</tr>";
  }).join("");
  var foot =
    '<tr class="totrow"><td><b>Total walk-ins</b></td>' +
    '<td class="num"><b>' + t.walkins + "</b></td>" +
    '<td class="num"><b>100%</b></td></tr>';
  return (
    '<div class="scroll-x"><table class="ptable srcstable">' +
    "<thead><tr><th>Customer Source</th><th>Walk-ins</th><th>% of walk-ins</th></tr></thead>" +
    "<tbody>" + rows + foot + "</tbody></table></div>"
  );
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

  // one distinguishable accent per owner-named repair bucket.
  function bucketAccent_(grp) {
    var g = str_(grp).toLowerCase();
    if (g.indexOf("unclaimed") > -1) return "bk-unclaimed";
    if ((g.indexOf("ready") > -1 && (g.indexOf("pick") > -1 || g.indexOf("collect") > -1))) return "bk-ready";
    if (/repair\s*in\s*(process|progress)/.test(g) || /work\s*in\s*progress/.test(g) ||
        /\bin[- ](process|progress)\b/.test(g)) return "bk-process";
    if (g.indexOf("estimat") > -1) return "bk-estimate";
    return "";
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

  // Column order fixed by the owner (12 columns):
  // Repair bucket · Aging · GSX repair status · Repair type · CRM status ·
  // Case ID · Name · Number · GSX ID · Device S.No · Model · Comm updated.
  // Rows stay grouped under a colour-coded band per repair bucket.
  var NCOLS = 11;
  // column header line repeated under every bucket band — CRM Status
  // (service!CRM_JobStatusDesc) sits between GSX Repair Status and Repair Type.
  var BUCKET_COLS =
    "<th>Aging</th><th>GSX Repair Status</th><th>CRM Status</th><th>Repair Type</th>" +
    "<th>Case ID</th><th>Name</th><th>Number</th>" +
    "<th>GSX ID</th><th>S.No</th><th>Model</th><th>Comm. Updated</th>";
  var prevGroup = null;
  var grpAccent = "";
  var rowParts = [];
  for (var gi = 0; gi < grouped.length; gi++) {
    var j = grouped[gi];
    var grp = j.crmRepairStatus || "—";
    if (grp !== prevGroup) {
      prevGroup = grp;
      var bucketCount = grouped.filter(function (x) {
        return (x.crmRepairStatus || "—") === grp;
      }).length;
      grpAccent = bucketAccent_(grp);
      // Task 45 owner's layout: Repair Bucket → bucket name → column headers,
      // then the rows of that bucket.
      rowParts.push('<tr class="bucket-sub-row"><th colspan="' + NCOLS + '"><span class="grphead-sub">Repair Bucket</span></th></tr>');
      rowParts.push('<tr class="grphead' + (grpAccent ? " " + grpAccent : "") + '"><td colspan="' + NCOLS + '">' + esc_(grp) +
        '<span class="grphead-count">' + bucketCount + " open case" + (bucketCount === 1 ? "" : "s") +
        "</span></td></tr>");
      rowParts.push('<tr class="bucket-cols">' + BUCKET_COLS + "</tr>");
    }
    rowParts.push(buildRow_(j, grpAccent));
  }
  var rows = rowParts.join("");

  function buildRow_(j, bk) {
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
    // first column is titled "CRM Status" now — value unchanged:
    // service!CRM_JobStatusDesc (the same field the bucket bands group by).
    var bucketCell = j.isUnclaimed
      ? '<td class="bucket-cell' + (bk ? " " + bk : "") + '"><span class="badge-unclaimed">' + esc_(j.crmRepairStatus || "—") + "</span></td>"
      : '<td class="bucket-cell' + (bk ? " " + bk : "") + '">' + esc_(j.crmRepairStatus || "—") + "</td>";
    var rowClses = [];
    if (bk) rowClses.push(bk);
    if (j.isUnclaimed) rowClses.push("unclaimed-row");
    // Rule (b): physical repair never started and the case is 2+ days old.
    if (j.noRepairLate) rowClses.push("norepair-row");
    // Rule (c): Out Of Warranty (No Coverage) with no advance anywhere.
    if (j.oowNoAdvance) rowClses.push("oow-nopay-row");
    return "<tr" + (rowClses.length ? ' class="' + rowClses.join(" ") + '"' : "") + ">" +
      ageCell(j.jobAge) +
      gsxDisplay +
      // CRM Status (service!CRM_JobStatusDesc) between GSX Repair
      // Status and Repair Type — the row order matches BUCKET_COLS.
      bucketCell +
      "<td>" + esc_(repairTypeShort_(j.repairType)) + "</td>" +
      '<td class="sticky-col"><button type="button" class="caselink" data-case="' + esc_(j.caseId) +
        '" onclick="openCaseDetails(this.dataset.case)"><b>' + esc_(j.caseId) + "</b></button>" +
        "</td>" +
      "<td>" + esc_(j.name) + "</td>" +
      '<td class="num">' + esc_(j.mobile) + "</td>" +
      // GSX ID + device serial right after Number, before Model.
      "<td>" + esc_(j.gsxId || "—") + "</td>" +
      "<td>" + esc_(j.serialNo || "—") + "</td>" +
      "<td>" + esc_(j.product) + "</td>" +
      commDate +
      "</tr>";
  }

  var table =
    '<div class="scroll-x jobwrap"><table class="ptable jobtable">' +
    '<colgroup>' +
    '<col style="width:10%"><col style="width:5%"><col style="width:10%">' +
    '<col style="width:8%"><col style="width:9%"><col style="width:9%">' +
    '<col style="width:9%"><col style="width:9%"><col style="width:9%">' +
    '<col style="width:10%"><col style="width:12%">' +
    '</colgroup>' +
    // every bucket block carries its own "Repair Bucket" subtitle,
    // the bucket band and a fresh column header row — no global thead.
    "<tbody>" + rows + "</tbody></table></div>";

  // Resolved column mapping — printed under the section so the user can
  // verify each displayed column against their sheet headers at a glance.
  function srcLine(pair) {
    var h = jm.sources ? str_(jm.sources[pair[1]]) : "";
    return esc_(pair[0]) + " ← " + (h ? "<b>" + esc_(h) + "</b>" : '<span class="map-missing">not found</span>');
  }
  var mapLine = '<div class="map-line">' + [
    ["CRM Status","crmRepairStatus"],["Aging","jobAge"],
    ["GSX Repair Status","gsxStatus"],["Repair Type","repairDocType"],["Case ID","caseId"],["Number","mobile"],
    ["GSX ID","gsxCode"],["Device S.No","serialNo1"],
    ["Model","product"],["Open filter","status"]]
    .map(srcLine).join(" · ") + "</div>";

  // case-id popup details. A hidden data island carries one JSON
  // row per open job; openCaseDetails() (client script) renders the modal.
  var cdMap = {};
  jm.jobs.forEach(function (j) {
    cdMap[j.caseId] = {
      caseId: j.caseId, createdDate: j.createdDate, repairDate: j.repairDate,
      bucket: j.crmRepairStatus, name: j.name, product: j.product,
      commDate: j.commDate, readyDate: j.readyDate || "—", gsxStatus: j.gsxStatus,
    };
  });
  var caseBits =
    '<div id="caseDetailData" hidden>' + JSON.stringify(cdMap).replace(/</g, "\u003c") + "</div>" +
    '<div class="casemodal-overlay" id="caseModal" style="display:none" ' +
    'onclick="if(event.target===this)closeCaseDetails()">' +
    '<div class="casemodal"><div class="cm-top"><b id="cmTitle">Case details</b>' +
    '<button type="button" onclick="closeCaseDetails()" aria-label="Close">&times;</button></div>' +
    '<table id="cmTable"></table></div></div>';

  return head + mapLine + table + caseBits;
}

/** Renders the repair-type & coverage analytics sections (Operation page). */
function renderRepairAnalyticsInner_(ra) {
  if (!ra.available) {
    return '<div class="section"><h2 class="section-title">🔧 Repair Analytics</h2>' +
      '<div class="empty-state">No service data available yet.</div></div>';
  }
  var out = [];

  // no 3-month average pills anywhere — current month only.
  function daysCell(avg) { return avg == null ? "—" : avg + " d"; }

  function bucketSection(icon, title, labelCol, comp, note) {
    if (!comp.length) return "";
    var o = ['<div class="section">' +
      '<h2 class="section-title">' + icon + " " + esc_(title) + " · <em>" + esc_(ra.curLabel) + "</em></h2>" +
      '<div class="section-sub">' + esc_(note) + ' Counts are whole numbers. Blank and "Service Non-Repair" repair types are excluded.</div>' +
      '<div class="tile-grid">'];
    comp.forEach(function (c) {
      o.push('<div class="tile"><div class="label">' + esc_(c.name) + "</div>" +
        '<div class="value">' + c.cur + "</div></div>");
    });
    o.push('</div><div class="table-wrap"><table aria-label="' + esc_(title) + '"><thead><tr>' +
      "<th>" + esc_(labelCol) + "</th>" +
      '<th class="num">This Month</th><th class="num">Avg Days · Now</th>' +
      "</tr></thead><tbody>");
    comp.forEach(function (c) {
      o.push("<tr><td><strong>" + esc_(c.name) + "</strong></td>" +
        '<td class="num">' + c.cur + "</td>" +
        '<td class="num">' + daysCell(c.curDaysAvg) + "</td></tr>");
    });
    o.push("</tbody></table></div></div>");
    return o.join("");
  }

  var typeSec = bucketSection("📈", "D2. Repair Type Performance", "Repair Type", ra.typeComp,
    "Case counts in " + ra.curLabel + " (month to date).");
  typeSec = insertExportBtn_(typeSec, "repairType");

  out.push(typeSec);

  // The "Repair Types — Till Date" composition table was retired on the
  // owner's request; the Repair Type Performance block above already carries
  // the month-to-date counts per type.
  return out.join("");
}

/** Task 49: decline / win-back card pair — shared by the OPERATION
 * group (SM page) and the TECH page. */
function renderOpsPairHtml_(ex) {
  return (
    '<div class="ops-pair"><section class="sec"><div class="sec-head"><h2>D3. 🚫 Decline Repair ' + exportSectionBtn_("decline") + '</h2>' +
    '<p class="sec-sub">Unique CaseIDs in <b>decline</b> where GSXRepairStatusDescription is ' +
    '"Declined - Rejected by Local Approver".</p></div>' +
    '<div class="cards">' +
    '<div class="card"><div class="label">Declined repairs</div><div class="value">' +
      (ex.decline.count || 0) + '</div>' +
      '<div class="sub">' + (ex.decline.sheetName ? "unique CaseIDs · from " + esc_(ex.decline.sheetName) : "no decline sheet found") + '</div></div>' +
    "</div></section>" +
  // Requote & Win-back: eligible cases are carry-in / mail-in
  // parts repairs, OUT of warranty (service!ProductWarrantyDesc); a win-back
  // is a serial whose Repair-Revenue quote was rejected and later approved.
    '<section class="sec"><div class="sec-head"><h2>D4. 🔁 Requote &amp; Win-back</h2>' +
    '<p class="sec-sub">Serials (service!SerialNo1) of out-of-warranty carry-in / mail-in cases whose ' +
    '<b>quote</b> rows (RevenueType "Repair Revenue") show lastAction rejected and later approved for the same serial.</p></div>' +
    '<div class="cards">' +
    '<div class="card"><div class="label">Win-backs</div><div class="value">' +
      (ex.winback.count || 0) + '</div>' +
      '<div class="sub">' + (ex.winback.sheetName ? "rejected → approved · unique serials · from " + esc_(ex.winback.sheetName) : "no quote sheet found") + '</div></div>' +
    "</div></section></div>");
}

/** OPERATION page: customer satisfaction + repair ops + sub-modules. */
function renderOperationHtml_(pd) {
  var out = [];
  out.push(renderTargetBlock_(pd.targetBlock, pd.repair));
  out.push(renderRepairAnalyticsInner_(pd.repair));
  var ex = pd.opsExtras || { decline: {}, winback: {} };
  // Decline Repair: unique decline!CaseId whose GSX status is
  // "Declined - Rejected by Local Approver".
  out.push(renderOpsPairHtml_(ex));

  return out.join("");
}

/** DECK page (Task 36): full rework — the owner's 8 tables, current month
 *  only, in the given order. All values come from computeDeckAnalysis_. */
function renderDeckHtml_(data, pd) {
  var d = (data && data.deckAnalysis) || {};
  var month = d.monthLabel || "";
  var stamp = month ? '<span class="fresh-chip fresh-stamp"><i>current month</i><b>' + esc_(month) + "</b></span>" : "";

  function th(cells) {
    return "<thead><tr>" + cells.map(function (c, i) {
      return "<th" + (i > 0 ? ' class="num"' : "") + ">" + esc_(c) + "</th>";
    }).join("") + "</tr></thead>";
  }
  function tr(cells, total) {
    return "<tr" + (total ? ' class="totrow"' : "") + ">" + cells.map(function (c, i) {
      return "<td" + (i === 0 ? ">" : ' class="num">') + c + "</td>";
    }).join("") + "</tr>";
  }
  function money0(v) { return money_(v || 0); }
  function table(id, icon, title, sub, head, body, foot) {
    return '<section class="section" id="' + id + '"><h2>' + icon + " " + esc_(title) + "</h2>" +
      '<div class="body">' +
      '<div class="scroll-x"><table class="ptable decktable">' + th(head) + "<tbody>" +
      body + (foot || "") + "</tbody></table></div></div></section>";
  }

  var sections = [];

  // ---- 1) Sales Revenue Analysis (revenue details) ----
  var sr = d.salesRev || { rows: [] };
  sections.push(table("deck-sales", "📊", "G1. Sales Revenue Analysis",
    "Revenue details — count & value per crm!RevenueType · crm!TaxableAmount. " +
    "crm!TranType = SALESRETURN is excluded from every valuation on this page.",
    ["Revenue Type", "Count", "Value"],
    sr.rows.map(function (r) { return tr([esc_(r.label), r.count, money0(r.value)]); }).join(""),
    tr(["<b>Total</b>", "<b>" + (sr.totalCount || 0) + "</b>", "<b>" + money0(sr.totalValue) + "</b>"], true)));

  // ---- 2) Walk-in details ----
  var w = d.walkins || {};
  sections.push(table("deck-walkin", "🚶", "G2. Walk-in Details",
    "Total unique walk-in count from walkin!MobileNo · unique invoice count from crm!InvoiceCode.",
    ["Metric", "Count"],
    tr(["No. Of Walkin (unique)", w.unique || 0]) +
    tr(["Unique Invoice", w.uniqBills || 0])));

  // ---- 3) Repair Revenue Analysis ----
  var rr = d.repairRev || { rows: [] };
  sections.push(table("deck-repair", "🔧", "G3. Repair Revenue Analysis",
    "Product families by service!ProductDescription · unique count from service!SerialNo1 · " +
    "value = crm!TaxableAmount where service!CaseId = crm!OrderCode. Service Non-Repair excluded.",
    ["Product", "Count", "Value"],
    rr.rows.map(function (r) { return tr([esc_(r.label), r.count, money0(r.value)]); }).join(""),
    tr(["<b>Total</b>", "<b>" + (rr.totalCount || 0) + "</b>", "<b>" + money0(rr.totalValue) + "</b>"], true)));

  // ---- 4) Accessories Revenue Analysis ----
  var ar = d.accessoriesRev || { rows: [] };
  sections.push(table("deck-accessories", "📱", "G4. Accessories Revenue Analysis",
    "Buckets from crm!ItemDescription — tg / temper(ed) glass / tamper → Tempered Glass; " +
    "Case; Adapter & Cable; AppleCare+/AC+; Protect Plus; AMC; rest → Others. Value = crm!TaxableAmount.",
    ["Product", "Count", "Value"],
    ar.rows.map(function (r) { return tr([esc_(r.label), r.count, money0(r.value)]); }).join(""),
    tr(["<b>Total</b>", "<b>" + (ar.totalCount || 0) + "</b>", "<b>" + money0(ar.totalValue) + "</b>"], true)));

  // ---- 5) Service Revenue Analysis ----
  var sv = d.serviceRev || { rows: [] };
  sections.push(table("deck-service", "🧾", "G5. Service Revenue Analysis",
    "Pivot on crm!ItemCode — Service Revenue only; crm!TranType = SALESRETURN excluded.",
    ["Item Code", "Count", "Value"],
    sv.rows.map(function (r) { return tr([esc_(r.label), r.count, money0(r.value)]); }).join("") ||
      tr(['<span class="muted">no Service Revenue rows this month</span>', "0", money0(0)]),
    tr(["<b>Total</b>", "<b>" + (sv.totalCount || 0) + "</b>", "<b>" + money0(sv.totalValue) + "</b>"], true)));

  // ---- 6) Gross Profit Analysis — Value ----
  var gp = d.gpValue || { rows: [] };
  sections.push(table("deck-gp", "💰", "G6. Gross Profit Analysis — Value",
    "Invoice value = Σ crm!TaxableAmount · GP value = Σ crm!GrossProfit · GP% = GP ÷ invoice value.",
    ["Revenue Type", "Invoice Value", "GP Value", "GP %"],
    gp.rows.map(function (r) {
      return tr([esc_(r.label), money0(r.value), money0(r.gp), r.pct.toFixed(1) + "%"]);
    }).join(""),
    (gp.total ? tr(["<b>Total</b>", "<b>" + money0(gp.total.value) + "</b>",
      "<b>" + money0(gp.total.gp) + "</b>", "<b>" + gp.total.pct.toFixed(1) + "%</b>"], true) : "")));

  // ---- 7) Walk-in Analysis (rank: most important) ----
  var wa = d.walkinAnalysis || [];
  sections.push(table("deck-walkin-analysis", "📈", "G7. Walk-in Analysis",
    "No. Of Walkin = unique walkin!MobileNo · Unique Bill = unique crm!InvoiceCode · Bills = total invoices · " +
    "AR/RR/SR = unique invoice count per crm!RevenueType · Walkin to Bills % = unique bills ÷ walk-ins · Enquiry = walk-ins − unique bills.",
    ["Metric", "Count"],
    wa.map(function (r) { return tr([esc_(r.label), r.value]); }).join("")));

  // ---- 8) Coverage Analysis | LOB ----
  var cov = d.coverage || { rows: [] };
  sections.push(table("deck-coverage", "🛡", "G8. Coverage Analysis | LOB",
    "Product + battery/display from service!ProductDescription & service!PartDescription (must match) · " +
    "W / NW from service!CoverageCodeDescription.",
    ["Product", "Cases", "W", "NW"],
    cov.rows.map(function (r) { return tr([esc_(r.label), r.count, r.w, r.nw]); }).join(""),
    (cov.total ? tr(["<b>Total</b>", "<b>" + cov.total.count + "</b>",
      "<b>" + cov.total.w + "</b>", "<b>" + cov.total.nw + "</b>"], true) : "")));

  // same owner order, laid out as four side-by-side pairs.
  var grid = "";
  for (var pi = 0; pi < sections.length; pi += 2) {
    grid += '<div class="deck-pair">' + sections[pi] + (sections[pi + 1] || "") + "</div>";
  }
  return '<section class="sec"><div class="sec-head"><h2>G. 🗂 Deck — ' + esc_(month || "Current Month") + " " + stamp + '</h2>' +
    '<p class="sec-sub">Owner order — revenue · walk-in · repair · accessories · service · GP · walk-in analysis · coverage. ' +
    "Current month only; previous months are kept manually by the owner. SALESRETURN is excluded everywhere.</p></div></section>" +
    grid;
}

/** OPEN JOBS page: job management. */
function renderJobsHtml_(pd) {
  return renderJobsSection_(pd);
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
  if (tb.labels.length) {
    chips.push(chip("", "Feedback responses", tb.totalRow.total,
      tb.people.length + " salesperson" + (tb.people.length === 1 ? "" : "s") + " rated"));
    chips.push(chip("pos", "Satisfied (VS + SS)", pctOf(tb.satTotal, tb.totalRow.total) + "%",
      tb.satTotal + " of " + tb.totalRow.total));
    if (tb.csat) chips.push(chip("info", "C-SAT", tb.csat, "as recorded in sheet"));
    if (tb.cfi) chips.push(chip(tb.cfi === "0.00%" || tb.cfi === "0" ? "pos" : "warn", "CFI", tb.cfi, "as recorded in sheet"));
    chips.push(chip(tb.disTotal > 0 ? "neg" : "pos", "Dissatisfied", pctOf(tb.disTotal, tb.totalRow.total) + "%",
      tb.disTotal + " of " + tb.totalRow.total));
  } else {
    // narrow target!E47:F51 report — no satisfaction buckets; show
    // the store C-SAT (sheet total, else average of the per-CRE values) and
    // how many CREs were rated.
    if (tb.csat) chips.push(chip("info", "Store C-SAT", tb.csat, "from target!F48:F51"));
    chips.push(chip("", "CREs rated", tb.people.length, "rows from target!E48:F51"));
    if (tb.cfi) chips.push(chip("info", "CFI", tb.cfi, "as recorded in sheet"));
  }

  var thead = "<thead><tr><th>" + esc_(tb.nameHeader) + "</th>";
  tb.labels.forEach(function (h) { thead += "<th>" + esc_(h) + "</th>"; });
  if (tb.labels.length || tb.hasTotalCol) thead += "<th>Total</th>";
  if (tb.labels.length) thead += "<th>% Satisfied</th>";
  if (tb.hasCsatCol) thead += "<th>C-SAT</th>";
  if (tb.hasCfiCol) thead += "<th>CFI</th>";
  thead += "</tr></thead>";

  function dataRow(e, isTotal) {
    var tr = "<tr" + (isTotal ? ' class="totrow"' : "") + "><td><b>" + esc_(e.name) + "</b></td>";
    e.cats.forEach(function (n, i) {
      var cls = tb.kinds[i] === "pos" ? "csat-pos" : tb.kinds[i] === "neg" ? "csat-neg" : "csat-mid";
      tr += '<td class="num ' + cls + '">' + (n || "—") + "</td>";
    });
    if (tb.labels.length || tb.hasTotalCol) tr += '<td class="num"><b>' + e.total + "</b></td>";
    if (tb.labels.length) tr += '<td class="num">' + pctOf(tb.satisfiedCount(e), e.total) + "%</td>";
    if (tb.hasCsatCol) tr += '<td class="num">' + esc_(e.csat || "—") + "</td>";
    if (tb.hasCfiCol) tr += '<td class="num">' + esc_(e.cfi || "—") + "</td>";
    return tr + "</tr>";
  }
  var body = "<tbody>" + tb.people.map(function (p) { return dataRow(p, false); }).join("") +
    dataRow(tb.totalRow, true) + "</tbody>";

  // Owner layout: table shrunk and left-aligned; the summary chips stack in a
  // column immediately to the table's right.
  // chips sit BELOW the table in one horizontal band (owner's layout).
  var table = '<div class="scroll-x csat-left"><table class="ptable csat-table">' + thead + body +
    '</table></div><div class="csat-brow">' + chips.join("") + "</div>";

  // the repairs 3-month comparison strip was removed (owner: no
  // last-3-month average comparisons anywhere on the dashboards).

  return '<section class="sec"><div class="sec-head">' +
    "<h2>🎯 D1. " + esc_(title) + "</h2>" +
    '<p class="sec-sub">Every feedback response lands in a satisfaction bucket. Period name from ' +
    '<span class="src-hint">target!D47</span>, table from <span class="src-hint">target!E47:F51</span> ' +
    '(auto-extends up to <span class="src-hint">:O51</span> when the full satisfaction grid is filled) — ' +
    "cleaned and re-presented (junk columns dropped, totals verified).</p></div>" +
    table + "</section>";
}

/** Job management section (OPEN JOBS page). */
function renderJobsSection_(pd) {
  var jm = pd.jobs;
  return (
    '<section class="sec"><div class="sec-head">' +
    '<h2>E1. 🔧 Job Management — Open Jobs ' + exportSectionBtn_("jobs") + '</h2>' +
    "</div>" + renderJobManagement_(jm) + "</section>"
  );
}

/* ==========================================================================
 * 11 · PAGE COMPOSERS                                                     *
 * ======================================================================== */

function renderIndexHtml_(data, perfData) {
  var s = data.store;
  var inv = data.invoiceValuation || {};
  var wlT = perfData && perfData.walkin && perfData.walkin.available ? perfData.walkin.totals : null;
  var jbT = perfData && perfData.jobs && perfData.jobs.available ? perfData.jobs.totals : null;
  var ex = perfData && perfData.opsExtras ? perfData.opsExtras : { decline: {}, winback: {} };
  function m0(v) { return money_(v || 0); }

  function statCard(label, value, sub, cls) {
    return '<div class="card"><div class="label">' + esc_(label) + "</div>" +
      '<div class="value' + (cls ? " " + cls : "") + '">' + value + "</div>" +
      (sub ? '<div class="sub">' + sub + "</div>" : "") + "</div>";
  }
  function idxSection(id, h2, desc, body) {
    return '<section class="section" id="' + id + '"><h2>' + h2 + "</h2><div class=\"body\">" +
      (desc ? '<div class="sec-desc">' + desc + "</div>" : "") + body + "</div></section>";
  }

  // ---- 1) Store performance -------------------------------------------------
  function storeCard(label, b) {
    return statCard(label, "<b>" + m0(b.achieved) + "</b>",
      "target " + m0(b.target) + " · <b>" + b.pct + "%</b> achieved");
  }
  var storeSec = idxSection("idx-store", "🎯 Store Performance",
    "month-to-date achievement against the store revenue targets (crm vs target!store).",
    '<div class="cards">' +
      storeCard("Total Revenue", s.total) + storeCard("Service Revenue", s.service) +
      storeCard("Accessories Revenue", s.accessory) + storeCard("Repair Revenue", s.repair) +
    "</div>");

  // ---- 2) Total walk-ins  +  3) invoices in each revenue type ---------------
  var types = inv.byType || [];
  var typeRows = types.map(function (t) {
    return "<tr><td><b>" + esc_(t.label) + "</b></td>" +
      '<td class="num"><b>' + t.uniqueInvoices + "</b></td>" +
      '<td class="num">' + t.invoiceLines + "</td>" +
      '<td class="num">' + m0(t.revenue) + "</td>" +
      '<td class="num">' + m0(t.gp) + "</td></tr>";
  }).join("");
  var totLines = 0, totRev = 0, totGp = 0;
  types.forEach(function (t) { totLines += t.invoiceLines; totRev += t.revenue; totGp += t.gp; });
  typeRows += '<tr class="totrow"><td><b>Total</b></td>' +
    '<td class="num"><b>' + s.uniqueInvoices + "</b></td>" +
    '<td class="num"><b>' + totLines + "</b></td>" +
    '<td class="num"><b>' + m0(s.total.achieved) + "</b></td>" +
    '<td class="num"><b>' + m0(s.grossProfit.achieved) + "</b></td></tr>";
  var walkinSec = idxSection("idx-walkin", "🚶 Total Walk-ins · 🧾 Invoices — each Revenue Type",
    "walk-in tokens from the walkin sheet; invoice counts are unique crm!InvoiceCode per crm!RevenueType (TranType = SALESRETURN excluded).",
    '<div class="cards" style="margin-bottom:12px">' +
      statCard("Total Walk-ins", "<b>" + (wlT ? wlT.walkins : 0) + "</b>",
        wlT ? (wlT.unique + " unique visitors · " + wlT.cases + " jobs · " + wlT.invoices + " invoices") : "walkin sheet not found") +
    "</div>" +
    '<div class="scroll-x"><table class="ptable"><thead><tr><th>Revenue Type</th>' +
    '<th class="num">Total Invoices</th><th class="num">Invoice Lines</th>' +
    '<th class="num">Revenue</th><th class="num">Gross Profit</th></tr></thead><tbody>' +
    typeRows + "</tbody></table></div>");

  // ---- 4) Weekly Revenue Analysis — INDEX shows the % summary, not the
  // detailed table (Task 49): completed/in-progress achieved %, upcoming weeks
  // the balance % each must contribute to close the month.
  var weeklySec = renderWeeklySummary_(data);

  // ---- 5) KPI Target — overall achieved % ------------------------------------
  var kAch = 0, kTgt = 0;
  var catChips = [["AR", data.kpi.ar], ["RR", data.kpi.rr], ["SR", data.kpi.sr]].map(function (c) {
    var a = 0, t = 0;
    c[1].forEach(function (k) { a += k.achieved || 0; t += k.target || 0; });
    kAch += a; kTgt += t;
    return statCard(c[0] + " KPIs", "<b>" + pct_(a, t) + "%</b>", a + " of " + t + " achieved");
  }).join("");
  var kpiSec = idxSection("idx-kpi", "📊 KPI Target — Store KPI Parameters",
    "overall achievement across every KPI parameter (AR + RR + SR · Σ achieved ÷ Σ target).",
    '<div class="cards">' +
      statCard("Overall KPI Achievement", "<b>" + pct_(kAch, kTgt) + "%</b>", kAch + " of " + kTgt + " parameters achieved", "headline") +
      catChips +
    "</div>");

  // ---- 6) AR Attachment — total count ----------------------------------------
  var aa = inv.arAttachment || {};
  var arSec = idxSection("idx-ar", "🔗 AR Attachment",
    "invoices that carried an accessory together with a service / repair.",
    '<div class="cards">' +
      statCard("AR Attachments", "<b>" + (aa.attachedInvoices || 0) + "</b>",
        "of " + (aa.attachableInvoices || 0) + " service / repair invoices · " + (aa.attachRate || 0) + "% attach rate") +
    "</div>");

  // ---- 7) Sales Person Performance — % of achievement only --------------------
  var spRows = (data.salespeople || []).map(function (p) {
    function pc(seg) { return '<td class="num">' + seg.pct + "%</td>"; }
    return "<tr><td><b>" + esc_(p.name) + "</b></td>" +
      pc(p.service) + pc(p.accessory) + pc(p.repair) + pc(p.grossProfit) +
      '<td class="num"><b>' + p.total.pct + "%</b></td></tr>";
  }).join("");
  var personSec = idxSection("idx-person", "👤 Sales Person Performance",
    "% of achievement in each segment — percentages only, no values (repair % is judged on the store repair target).",
    '<div class="scroll-x"><table class="ptable"><thead><tr><th>Salesperson</th>' +
    '<th class="num">Service %</th><th class="num">Accessory %</th><th class="num">Repair %</th>' +
    '<th class="num">Gross Profit %</th><th class="num">Total %</th></tr></thead><tbody>' +
    spRows + "</tbody></table></div>");

  // ---- 8/9/10) decline repairs · win-backs · open jobs ------------------------
  var opsSec = idxSection("idx-ops", "📉 Repairs & Jobs — counts",
    "headline counts from the service / decline / quote / job sources (details on the SM and TECH pages).",
    '<div class="cards">' +
      statCard("Declined Repairs", "<b>" + (ex.decline.count || 0) + "</b>", "unique decline!CaseId · local approver rejected", "neg") +
      statCard("Requote Win-backs", "<b>" + (ex.winback.count || 0) + "</b>", "rejected → approved · unique serials", "pos") +
      statCard("Open Jobs", "<b>" + (jbT ? jbT.openJobs : 0) + "</b>",
        jbT ? ("communication updated " + jbT.communicated + " · pending " + jbT.notCommunicated) : "service sheet not found") +
    "</div>");
  opsSec = insertExportBtn_(opsSec, "ops");

  // ---- 11) Working Days (Current Month) — the master A2 module, moved to the
  // top of the page directly below 🎯 Store Performance (Task 49).
  var wdSec = renderWorkingDaysCard_(data).replace('id="sec-workdays"', 'id="idx-workdays"');

  return (
    '<section class="index-hero"><h2>Store Performance Dashboard</h2></section>' +
    storeSec + wdSec + walkinSec + weeklySec + kpiSec + arSec + personSec + opsSec +
    freshnessLineHtml_(data, perfData)
  );
}

/**
 * CRE page manifest (owner's list, replaces the Task-47 minimal
 * view): A1 Store Target, A2 Working Days, A6 Weekly, B / B1 / B2 / B3 KPI
 * family, C with C1 walk-in analysis only, and D1 customer satisfaction —
 * all rendered by the same master modules the SM page uses.
 */
function renderCreHtml_(data, pd) {
  var wl = pd.walkin || { available: false };
  var c1 = !wl.available
    ? emptyPerf_(CONFIG.WALKIN_SHEET_NAMES[0], "Add a walk-in sheet to see the C1 walk-in analysis.")
    : '<section class="section" id="sec-cre-walkin"><h2>C. 🚶 Walk-in Analysis &amp; Token Creation</h2><div class="body">' +
      '<div class="wp-left"><div class="wp-head">C1. 🚶 Walk-in Analysis</div>' +
      '<p class="sec-desc" style="margin:-2px 0 10px">month-to-date walk-in summary — each metric&rsquo;s % of the month&rsquo;s walk-in tokens.</p>' +
      renderWalkinAnalysis_(wl) + "</div></div></section>";
  return (
    // A1 · A2 · A6
    renderStoreCards_(data) + renderStoreTable_(data) + renderWorkingDaysCard_(data) +
      renderWeeklyAnalysis_(data) +
    // B · B1 · B2 · B3
    '<div class="kpi-pair"><div class="kp-left">' + renderKpiSection_(data) + "</div>" +
      '<div class="kp-right">' + renderIncentiveTable_(data) + "</div></div>" +
      renderSalespersonCards_(data) + renderArSection_(data) +
    // C · C1
    c1 +
    // D1
    renderTargetBlock_(pd.targetBlock, pd.repair)
  );
}

/**
 * TECH page manifest (owner's list): D2 Repair Type Performance,
 * D3 Decline Repair, D4 Requote & Win-back, then the E1 job-management zone —
 * the Task-47 bucket summary (counts · aging bands · communication) followed
 * by the full master open-jobs table.
 */
function renderTechHtml_(pd) {
  var jm = pd.jobs || { available: false };
  var d2 = renderRepairAnalyticsInner_(pd.repair || { available: false });
  var d34 = renderOpsPairHtml_(pd.opsExtras || { decline: {}, winback: {} });
  if (!jm.available || !jm.jobs) {
    return d2 + d34 +
      emptyPerf_("service", "Add a service sheet to see the bucket, aging and communication summary.");
  }
  var t = jm.totals || {};
  // aging split-up 0-2 · 3-5 · 6-10 · 11-15 · 16-30 ·
  // 31-45 · 46-60 · 61-100 · 100+ days.
  var BANDS = [[0, 2, "0–2 d"], [3, 5, "3–5 d"], [6, 10, "6–10 d"], [11, 15, "11–15 d"],
    [16, 30, "16–30 d"], [31, 45, "31–45 d"], [46, 60, "46–60 d"], [61, 100, "61–100 d"], [101, 1e12, "100+ d"]];
  var order = [], map = {};
  jm.jobs.forEach(function (j) {
    var b = (j.crmRepairStatus || "").trim() || "—";
    if (!map[b]) { map[b] = { bands: BANDS.map(function () { return 0; }), noAge: 0, total: 0 }; order.push(b); }
    var rec = map[b];
    rec.total++;
    var age = (typeof j.jobAge === "number" && isFinite(j.jobAge) && j.jobAge >= 0) ? j.jobAge : null;
    if (age == null) rec.noAge++;
    else BANDS.forEach(function (bd, i) { if (age >= bd[0] && age <= bd[1]) rec.bands[i]++; });
  });

  // Alphabetical bucket order — mirrors the band order of the detailed SM table.
  order.sort();
  function cell(n) { return n ? ('<td class="num">' + n + "</td>") : '<td class="num muted">—</td>'; }
  var rows = order.map(function (b) {
    var r = map[b];
    return "<tr><td><b>" + esc_(b) + "</b></td>" +
      BANDS.map(function (_, i2) { return cell(r.bands[i2]); }).join("") + cell(r.noAge) +
      '<td class="num"><b>' + r.total + "</b></td></tr>";
  }).join("");
  var g = BANDS.map(function () { return 0; }), gNoAge = 0, gTotal = 0;
  order.forEach(function (b) {
    var r = map[b];
    BANDS.forEach(function (_, i2) { g[i2] += r.bands[i2]; });
    gNoAge += r.noAge; gTotal += r.total;
  });
  var tot = '<tr class="totrow"><td><b>All open jobs</b></td>' +
    g.map(function (n) { return '<td class="num"><b>' + n + "</b></td>"; }).join("") +
    '<td class="num"><b>' + (gNoAge || 0) + "</b></td>" +
    '<td class="num"><b>' + gTotal + "</b></td></tr>";

  var commCards =
    '<div class="cards">' +
    '<div class="card"><div class="label">Open jobs</div><div class="value">' + (t.openJobs || gTotal) +
    '</div><div class="sub">across ' + order.length + " repair bucket" + (order.length === 1 ? "" : "s") + '</div></div>' +
    '<div class="card"><div class="label">Communication updated</div><div class="value pos">' + (t.communicated || 0) +
    '</div><div class="sub">' + (t.commRate || 0) + '% of open jobs</div></div>' +
    '<div class="card"><div class="label">Not updated</div><div class="value neg">' + (t.notCommunicated || 0) +
    '</div><div class="sub">no comm entry for the case</div></div>' +
    "</div>";

  return d2 + d34 +
    '<section class="section" id="sec-tech-jobs"><h2>E1. 🔧 Job Management — Open Jobs · Bucket Summary</h2><div class="body">' +
    '<div class="sec-desc">open jobs per repair bucket with aging inside each bucket (Jobs without a readable age are ' +
    'listed in a separate column); then communication updated vs not updated for all open jobs.</div>' +
    commCards +
    '<div class="scroll-x"><table class="ptable"><thead><tr><th>Repair Bucket</th>' +
    BANDS.map(function (bd) { return '<th class="num">' + bd[2] + "</th>"; }).join("") +
    '<th class="num">Age —</th><th class="num">Total</th></tr></thead><tbody>' +
    rows + tot + "</tbody></table></div></div></section>" +
    renderJobsHtml_(pd);
}

/**
 * SM page: the full & detailed master report. Every section of the
 * legacy pages in the owner's master order A → G, separated by group banners.
 */
function renderSmHtml_(data, pd) {
  return (
    roleBanner_("A", "GP — store target · working days · valuation · margin · discount · weekly") +
      renderGpHtml_(data) +
    roleBanner_("B", "KPI — target parameters · incentive · salesperson performance · AR attachment") +
      renderKpiHtml_(data) +
    roleBanner_("C", "WALK-IN — analysis · customer source · token creation · day-wise") +
      renderWalkinHtml_(pd) +
    roleBanner_("D", "OPERATION — customer satisfaction · repair types · decline · win-back") +
      renderOperationHtml_(pd) +
    roleBanner_("E", "OPEN JOBS — job management") +
      renderJobsHtml_(pd) +
    roleBanner_("F", "DATE — date-wise sales revenue by salesperson") +
      renderDateHtml_(data) +
    roleBanner_("G", "DECK — current-month summary tables") +
      renderDeckHtml_(data, pd)
  );
}

/** Bottom line on the index: latest dated record seen in each sheet. */
function freshnessLineHtml_(data, perfData) {
  var bits = [];
  // every chip reads the sheet's own LAST record date AND time.
  var fr = perfData && perfData.freshness ? perfData.freshness : {};
  bits.push(["crm (sales)", fr.crm || "no rows"]);
  bits.push(["target (plan)", "refreshed " +
    Utilities.formatDate(data.generatedAt, Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm" )]);
  bits.push(["walkin", fr.walkin || "no rows"]);
  bits.push(["service", fr.service || "no rows"]);
  bits.push(["advance", fr.advance || "—"]);
  bits.push(["comm", fr.comm || "—"]);
  var li = bits.map(function (b) {
    return '<span class="fresh-chip"><i>' + esc_(b[0]) + "</i><b>" + esc_(b[1]) + "</b></span>";
  }).join("");
  // every tab is stamped with the dashboard-refresh time — show it.
  var stamp = (perfData && perfData.refreshStamp)
    ? '<span class="fresh-chip fresh-stamp"><i>all sheets</i><b>' +
      esc_(perfData.refreshStamp.replace("⟳ Dashboard refresh · ", "")) + "</b></span>"
    : "";
  return '<div class="fresh-line"><div class="fresh-line-head">⏱ Latest records per sheet</div>' +
    '<div class="fresh-row">' + stamp + li + "</div></div>";
}

/**
 * Single source of truth for the four page panes. The cold load, the
 * refresh endpoint and the 15-minute background top-up all compute and
 * cache pane markup through this one builder, so the paths can never
 * drift apart.
 */
function buildAllPanes_(data, perfData) {
  return {
    index: renderIndexHtml_(data, perfData),
    cre: renderCreHtml_(data, perfData),
    tech: renderTechHtml_(perfData),
    sm: renderSmHtml_(data, perfData)
  };
}

/* ==========================================================================
 * 12 · PER-TABLE EXPORT                                                   *
 * ======================================================================== */

/**
 * per-table export. The page-wide "Export to Sheets" is gone; a
 * small "⬇ Export" button sits on each owner-listed table only:
 *   A1 Store Target — Performance · Repairs & Jobs — counts ·
 *   D2 Repair Type Performance · D3 Decline Repair ·
 *   E1 Job Management — Open Jobs · A4 High / Low Margin ·
 *   B KPI Target — Store KPI Parameters
 * A click creates a NEW spreadsheet containing just that table, as shown.
 */
function exportSectionBtn_(key) {
  return '<button class="mini-export" data-export-key="' + key + '" onclick="dashExportSection(\'' + key + '\',this)" ' +
    'title="Export this table to a spreadsheet">⬇ Export</button>';
}

function insertExportBtn_(html, key) {
  return html.replace(/<h2[^>]*>/, function (m) { return m + exportSectionBtn_(key); });
}

/** Table blocks for one export key (headers + rows, exactly as shown). */
function exportBlocksFor_(key) {
  // Buttons send camelCase keys (e.g. 'repairType'); normalize before the
  // comparisons below so case can never silently drop a table again.
  key = str_(key).trim().toLowerCase();
  key = String(key || "").toLowerCase();
  var loaded = loadDashboardData_();
  var data = loaded.data;
  var pd = loadPerformanceData_();

  function storePerfRows_() {
    var st = data.store;
    function tile(title, b) {
      var pct = b.target > 0 ? (100 * b.achieved / b.target) : 0;
      var complete = b.target > 0 && pct >= 100;
      var signedBalance = b.target - b.achieved;
      return [title, complete ? "Completed" : "In Progress",
        b.target, b.achieved, (Math.round(pct * 10) / 10) + "%",
        b.target > 0 ? signedBalance : 0,
        complete ? signedBalance : b.requiredPerDay];
    }
    return [tile("Service Revenue", st.service), tile("Accessory Revenue", st.accessory),
      tile("Repair Revenue", st.repair), tile("Total Revenue", st.total),
      tile("Gross Profit", st.grossProfit)];
  }
  function opsCountRows_() {
    var ex = pd.opsExtras || {};
    var t = (pd.jobs && pd.jobs.totals) || {};
    return [
      ["Declined Repairs — unique decline!CaseId · local approver rejected",
        ex.decline ? ex.decline.count || 0 : 0],
      ["Requote Win-backs — serials rejected → approved",
        ex.winback ? ex.winback.count || 0 : 0],
      ["Open Jobs", t.openJobs || 0],
      ["Communication updated", t.communicated || 0],
      ["Update pending", t.notCommunicated || 0]];
  }
  function repairTypeRows_() {
    var ra = pd.repair || {};
    return (ra.typeComp || []).map(function (c) {
      return [c.name, c.cur, c.curDaysAvg == null ? "—" : c.curDaysAvg + " d"];
    });
  }
  function declineRows_() {
    var ex = pd.opsExtras || {}, d = ex.decline || {};
    return [
      ["Declined repairs — unique CaseIDs · 'Declined - Rejected by Local Approver'",
        d.count || 0, d.sheetName || ""],
      ["Decline rows matched (all statuses)", d.matched || 0, d.sheetName || ""]];
  }
  var JOB_HEADERS = ["Aging", "GSX Repair Status", "CRM Status", "Repair Type",
    "Case ID", "Name", "Number", "GSX ID", "S.No", "Model", "Comm. Updated"];
  function jobRows_() {
    var now = Date.now(), rows = [];
    var jobs = (pd.jobs && pd.jobs.available && pd.jobs.jobs) ? pd.jobs.jobs : [];
    jobs.forEach(function (j) {
      var comm = "—";
      if (j.commTs > 0) {
        var dfc = Math.max(0, Math.floor((now - j.commTs) / 86400000));
        comm = dfc === 0 ? "today" : dfc + "d";
      }
      rows.push([j.jobAge == null ? "" : j.jobAge,
        j.gsxStatus || "—", j.crmRepairStatus || "—", j.repairType || "—",
        j.caseId || "—", j.name || "—", j.mobile || "—", j.gsxId || "—",
        j.serialNo || "—", j.product || "—", comm]);
    });
    return rows;
  }
  var BANDS = [[0, 2, "0–2 d"], [3, 5, "3–5 d"], [6, 10, "6–10 d"], [11, 15, "11–15 d"],
    [16, 30, "16–30 d"], [31, 45, "31–45 d"], [46, 60, "46–60 d"], [61, 100, "61–100 d"], [101, 1e12, "100+ d"]];
  function bucketHeaders_() {
    return ["Repair Bucket"].concat(BANDS.map(function (bd) { return bd[2]; })).concat(["Age —", "Total"]);
  }
  function bucketRows_() {
    var jobs = (pd.jobs && pd.jobs.available && pd.jobs.jobs) ? pd.jobs.jobs : [];
    var bMap = {}, bOrder = [];
    jobs.forEach(function (j) {
      var b = (j.crmRepairStatus || "—");
      if (!bMap[b]) { bMap[b] = { bands: BANDS.map(function () { return 0; }), noAge: 0, total: 0 }; bOrder.push(b); }
      var rec = bMap[b]; rec.total++;
      var age = (typeof j.jobAge === "number" && isFinite(j.jobAge) && j.jobAge >= 0) ? j.jobAge : null;
      if (age == null) rec.noAge++;
      else BANDS.forEach(function (bd, i) { if (age >= bd[0] && age <= bd[1]) rec.bands[i]++; });
    });
    bOrder.sort();
    return bOrder.map(function (b) {
      var r = bMap[b];
      return [b].concat(r.bands).concat([r.noAge, r.total]);
    });
  }
  function marginRows_(mk) {
    var x = data.gpExtras || {};
    return (x[mk] || []).map(function (m) {
      return [m.partLabel || "—", m.invoice || "—", m.typeLabel || "—",
        Math.round(m.taxable || 0), Math.round(m.gp || 0),
        (Math.round((m.margin || 0) * 10) / 10) + "%"];
    });
  }
  var MARGIN_HEADERS = ["Part", "Invoice", "Revenue Type", "Taxable", "Gross Profit", "Margin %"];
  function kpiTargetRows_() {
    var allKpis = data.kpi.ar.concat(data.kpi.rr, data.kpi.sr);
    return allKpis.map(function (k) {
      var pctVal = k.target > 0 ? pct_(k.achieved, k.target) : 0;
      var approxBal = k.approxAmount > 0 && k.target > 0 ? k.balance * k.approxAmount : "—";
      return [k.label || k.name, k.category || "",
        k.target > 0 ? k.target : "—", k.achieved,
        k.target > 0 ? pctVal + "%" : "—",
        k.target > 0 ? k.balance : "—",
        approxBal === "—" ? "—" : Math.round(approxBal),
        Math.round(k.achievedValue || 0), Math.round(k.achievedGp || 0)];
    });
  }

  if (key === "store") return { name: "Store Target — Performance", blocks: [
    { title: "A1. Store Target — Performance",
      headers: ["Revenue Type", "Status", "Target", "Achievement", "Achievement %", "Balance (Target − Achievement)", "Per Day Requirement / Excess"],
      rows: storePerfRows_() }] };
  if (key === "ops") return { name: "Repairs & Jobs — counts", blocks: [
    { title: "Repairs & Jobs — counts", headers: ["Metric", "Count"], rows: opsCountRows_() }] };
  if (key === "repairtype") return { name: "Repair Type Performance", blocks: [
    { title: "D2. Repair Type Performance · " + ((pd.repair && pd.repair.curLabel) || "current month"),
      headers: ["Repair Type", "This Month", "Avg Days · Now"], rows: repairTypeRows_() }] };
  if (key === "decline") return { name: "Decline Repair", blocks: [
    { title: "D3. Decline Repair", headers: ["Metric", "Count", "Source sheet"], rows: declineRows_() }] };
  if (key === "jobs") return { name: "Job Management — Open Jobs", blocks: [
    { title: "Bucket Summary — 9 aging bands", headers: bucketHeaders_(), rows: bucketRows_() },
    { title: "E1. Job Management — Open Jobs", headers: JOB_HEADERS, rows: jobRows_() }] };
  if (key === "margin") return { name: "High Low Margin", blocks: [
    { title: "A4. High Margin — top 5", headers: MARGIN_HEADERS, rows: marginRows_("marginHigh") },
    { title: "A4. Low Margin — top 5", headers: MARGIN_HEADERS, rows: marginRows_("marginLow") }] };
  if (key === "kpi") return { name: "KPI Target — Store KPI Parameters", blocks: [
    { title: "B. KPI Target — Store KPI Parameters",
      headers: ["KPI", "Category", "Target", "Achieved", "%", "Balance", "Approx Balance Value", "Achieved Value", "Achieved GP"],
      rows: kpiTargetRows_() }] };
  return null;
}

/** Public endpoint called by the per-table "⬇ Export" buttons. */
function exportDashboardSectionToSpreadsheet(key) {
  var spec = exportBlocksFor_(key);
  if (!spec) {
    if (key === null || key === undefined || str_(key).trim() === "") {
      throw new Error("No export table was specified. This function takes a " +
        "table name — open the dashboard page and click the ⬇ Export button " +
        "on the table you want; running it from the editor (or a sheet " +
        "shortcut) passes nothing.");
    }
    throw new Error("Unknown export table: " + key);
  }
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy HH:mm");
  var book = SpreadsheetApp.create(spec.name + " — " + stamp);
  var sh = book.getSheets()[0];
  sh.setName(spec.name.replace(/[\[\]:*?/\\]/g, " ").slice(0, 31));
  var r = 1;
  spec.blocks.forEach(function (b) {
    sh.getRange(r, 1).setValue(b.title).setFontWeight("bold");
    sh.getRange(r + 1, 1, 1, b.headers.length).setValues([b.headers]).setFontWeight("bold");
    if (b.rows.length) sh.getRange(r + 2, 1, b.rows.length, b.headers.length).setValues(b.rows);
    r += 2 + b.rows.length + 1;
  });
  return JSON.stringify({ name: book.getName(), url: book.getUrl() });
}

/* ==========================================================================
 * 13 · SHEET-SEARCH CHAT ASSISTANT                                        *
 * ======================================================================== */

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
      result.note = "Try a customer mobile number or a Case ID.";
      return JSON.stringify(result);
    }

    var c = chatQuery(query);
    result.title = c.title;
    result.rows = c.rows || [];
    result.note = c.note || "";
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

// chatTxt is HTML-raw: call sites esc_() sheet data themselves; chatPlain wraps + escapes.
function chatTxt(t) { return "<td>" + t + "</td>"; }

function chatPlain(t) { return "<td>" + esc_(t) + "</td>"; }

/** Canonical 10-digit mobile (strip +91 / leading 0). */
function chatDig(q) {
  var d = str_(q).replace(/\D/g, "");
  if (d.length > 10) d = d.slice(-10);
  return d.length === 10 ? d : "";
}

function chatMoneyCell(n) { return '<td class="num">₹' + Math.round(n || 0).toLocaleString("en-IN") + "</td>"; }

/**
 * Route the free-text question to the right search.
 * Returns { title, rows, note } where rows are table-cell arrays.
 */
/**
 * Task 30 hard limit: the bot answers ONLY two kinds of lookups —
 *   1. a customer mobile number  → service!MobileNo → case ID, CRM repair
 *      status (CRM_JobStatusDesc), customer name, product, plus the rest of
 *      the 360 (walk-in visits, advance balance, CRM invoices).
 *   2. a Case ID                 → service!CaseId  → the full case detail.
 * Anything else gets a gentle guidance reply instead of a fuzzy search.
 */
function chatQuery(q) {
  // 1) A contiguous 10+ digit run (a real phone number) — dates never match.
  var phoneMatch = str_(q).match(/[\d+][\d\s().\-]{8,}\d/);
  var mobile = phoneMatch ? chatDig(phoneMatch[0]) : "";
  if (mobile) return chatMobileQuery(mobile, q);

  // 2) A case-id shaped token (optionally letters + 3+ digits): C1001,
  // CS-200145, or a bare 6+ digit job number.
  var tok = null;
  var tokens = (str_(q).toLowerCase().replace(/[^a-z0-9\- ]/g, " ").split(/\s+/));
  for (var i = 0; i < tokens.length; i++) {
    var t = tokens[i];
    if (t === "case" || t === "caseid" || t === "id" || t === "job" || t === "status") continue;
    if (/^[a-z]{0,4}\d{3,}[a-z0-9\-]*$/.test(t)) { tok = t; break; }
  }
  if (tok) {
    var hit = chatCaseQuery(tok);
    if (hit) return hit;
  }

  // 3) Task 35 — universal search: every other query is scanned across ALL
  //    sheets of this spreadsheet (never external sources). Exact token
  //    matches only — no fuzzy/similar suggestions are ever offered.
  var uni = chatUniversalQuery_(q);
  if (uni) return uni;
  return {
    title: "No matches for “" + str_(q).trim() + "”",
    rows: [[chatTxt('<span class="muted">Nothing in any sheet contains all those words. ' +
      "Numbers, names, case IDs, products, statuses and dates are searched exactly as typed.</span>")]],
    note: "Searched every tab of this spreadsheet.",
    wide: false,
  };
}

/**
 * TASK 35 — universal cross-sheet search. A data row matches when EVERY
 * query token (case-insensitive, 1-char tokens ignored) appears somewhere in
 * that row. Results are grouped by sheet (tab order), identical rows are
 * de-duplicated, and output is capped so the panel stays readable.
 */
function chatUniversalQuery_(q) {
  var query = str_(q).trim();
  if (query.length < 2) return null;
  var tokens = query.toLowerCase().split(/\s+/).filter(function (t) { return t.length >= 2; });
  if (!tokens.length) return null;
  // Deployment hardening: if this helper ever goes missing from the project
  // (stale/partial paste), fall back to the bound spreadsheet directly.
  var ss = (typeof getSpreadsheet_ === "function") ? getSpreadsheet_()
    : SpreadsheetApp.getActiveSpreadsheet();
  var rows = [];
  var totalHits = 0, hitSheets = 0;
  var MAX_PER_SHEET = 4, MAX_ROWS = 18;
  ss.getSheets().forEach(function (sh) {
    if (rows.length >= MAX_ROWS) return;
    var values;
    try { values = chatSheetData(sh); } catch (e) { return; }
    var header = values.header;
    var seenRows = {};
    var sheetRows = [];
    for (var i = 0; i < values.rows.length && sheetRows.length < MAX_PER_SHEET; i++) {
      var r = values.rows[i];
      var cells = r.map(function (v) { return str_(v).toLowerCase(); });
      var joined = " " + cells.join(" │ ") + " ";
      var all = tokens.every(function (t) {
        var tn = t.replace(/[.\/\-]/g, " ");
        return joined.indexOf(t) > -1 ||
          (tn !== t && (" " + joined.replace(/[.\/\-]/g, " ") + " ").indexOf(" " + tn + " ") > -1);
      });
      if (!all) continue;
      var sig = r.join("|").toLowerCase();
      if (seenRows[sig]) continue;
      seenRows[sig] = true;
      var pairs = [];
      for (var c = 0; c < r.length && pairs.length < 8; c++) {
        var v = str_(r[c]);
        if (!v) continue;
        var h = (header[c] !== undefined && str_(header[c])) ? str_(header[c]) : "col " + (c + 1);
        pairs.push(h.length > 26 ? h.slice(0, 24) + "…" : h);
        pairs.push(v.length > 60 ? v.slice(0, 58) + "…" : v);
      }
      if (pairs.length) sheetRows.push({ rowNo: i + 2, pairs: pairs });
    }
    if (sheetRows.length) {
      hitSheets++;
      rows.push([chatTxt("<b>📄 " + esc_(sh.getName()) + "</b>"),
        chatTxt('<span class="muted">' + sheetRows.length + " match" + (sheetRows.length === 1 ? "" : "es") + "</span>")]);
      sheetRows.forEach(function (m) {
        if (rows.length >= MAX_ROWS + 2) return;
        rows.push([chatTxt('<span class="muted">row ' + m.rowNo + "</span>"),
          chatPlain(chatPairsInline_(m.pairs))]);
        totalHits++;
      });
    }
  });
  if (!totalHits) return null;
  return {
    title: "🔍 “" + query + "” → " + totalHits + " row" + (totalHits === 1 ? "" : "s") +
      " in " + hitSheets + " sheet" + (hitSheets === 1 ? "" : "s"),
    rows: rows,
    note: "Exact matches only — nothing similar/related is included.",
    wide: true,
  };
}

/** header/value pairs → "H: V · H: V" single-line text. */
function chatPairsInline_(pairs) {
  var out = [];
  for (var i = 0; i + 1 < pairs.length; i += 2) out.push(pairs[i] + ": " + pairs[i + 1]);
  return out.join(" · ");
}

/** Chat answer for a mobile number — the 360 customer profile. */
function chatMobileQuery(mobile, origQ) {
  // Deployment hardening: if this helper ever goes missing from the project
  // (stale/partial paste), fall back to the bound spreadsheet directly.
  var ss = (typeof getSpreadsheet_ === "function") ? getSpreadsheet_()
    : SpreadsheetApp.getActiveSpreadsheet();
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
  // one diagnostic block per case — Case ID, customer, product,
  // CRM case status, GSX case status, CRM bucket, case date and the most
  // recent communication-update date.
  caseArr.forEach(function (c) {
    var s = caseIds[c];
    rows.push([chatTxt("<b>&#128736;&#65039; " + c + "</b>")]);
    rows.push([chatTxt("Case ID"), chatTxt(c)]);
    if (name) rows.push([chatTxt("Customer"), chatPlain(name)]);
    rows.push([chatTxt("Product"), chatPlain(s.product || "—")]);
    rows.push([chatTxt("CRM case status"), chatPlain(s.caseStatus || "—")]);
    rows.push([chatTxt("GSX case status"), chatPlain(s.gsx || "—")]);
    rows.push([chatTxt("CRM bucket"), chatPlain(s.status || "—")]);
    rows.push([chatTxt("Case date"), chatPlain(s.caseDate || "—")]);
    var commD = chatCommDateForCase_(ss, c);
    rows.push([chatTxt("Comm updated"), chatPlain(commD || "no updates logged")]);
  });

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
      note: "Check the number, or try a Case ID instead.",
      wide: true,
    };
  }
  return { title: title, rows: rows, note: "criss-crossed through walk-in, service, advance and CRM sheets.", wide: true };
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
  var cStatus = findColumn_(values.header, [/^crm_?jobstatusdesc$/i, /crm.*status/i, /job.*status.*desc/i]);
  var cGsx = findColumn_(values.header, [/gsx/i, /repair.*status/i]);
  var cProd = findColumn_(values.header, [/product.*desc/i, /^product/i, /model/i]);
  // case status (service!Status), the CRM bucket and the case date
  // ride along so the answer can list the full per-case diagnostic block.
  var cCaseStatus = findColumn_(values.header, [/^status$/i, /^jobstatus$/i, /case.*status/i]);
  var cCaseDate = findColumn_(values.header, [/^casedate$/i, /case.*date/i]);
  var out = [];
  values.rows.forEach(function (r) {
    if (chatDig(r[cMobile]) === mobile) {
      out.push({
        first: cFirst !== -1 ? str_(r[cFirst]) : "",
        last: cLast !== -1 ? str_(r[cLast]) : "",
        caseId: cCase !== -1 ? str_(r[cCase]) : "",
        status: cStatus !== -1 ? str_(r[cStatus]) : "",
        caseStatus: cCaseStatus !== -1 ? str_(r[cCaseStatus]) : "",
        caseDate: cCaseDate !== -1 ? str_(r[cCaseDate]) : "",
        gsx: cGsx !== -1 ? str_(r[cGsx]) : "",
        product: cProd !== -1 ? str_(r[cProd]) : "",
      });
    }
  });
  return out;
}

/**
 * Latest communication-created date for a Case ID (Task 33). Reads the
 * chat-display values of the communication sheet and keeps the newest
 * parseable stamp; the day key comes back as dd-MM-yyyy for display.
 */
function chatCommDateForCase_(ss, caseId) {
  var sh = getSheetByNames_(ss, CONFIG.COMM_SHEET_NAMES);
  if (!sh) return "";
  var values = chatSheetData(sh);
  var cCase = findColumn_(values.header, COMM_SPEC.caseId);
  var cDate = findColumn_(values.header, COMM_SPEC.createdDate);
  if (cCase === -1 || cDate === -1) return "";
  var want = str_(caseId).toLowerCase();
  var best = null;
  values.rows.forEach(function (r) {
    if (str_(r[cCase]).toLowerCase() === want) {
      var p = parseAnyDate_(r[cDate]);
      if (p.valid && (!best || p.ts > best.ts)) best = p;
    }
  });
  return best ? best.dayKey : "";
}

/**
 * Direct Case-ID lookup against service!CaseId (Task 30). Returns null when
 * the ID isn't on the sheet so the router can fall back to the guidance text.
 */
function chatCaseQuery(caseId) {
  // Deployment hardening: if this helper ever goes missing from the project
  // (stale/partial paste), fall back to the bound spreadsheet directly.
  var ss = (typeof getSpreadsheet_ === "function") ? getSpreadsheet_()
    : SpreadsheetApp.getActiveSpreadsheet();
  var sh = getSheetByNames_(ss, CONFIG.SERVICE_SHEET_NAMES);
  if (!sh) return null;
  var values = chatSheetData(sh);
  var cCase = findColumn_(values.header, [/^caseid$/i, /^case.*id/i]);
  if (cCase === -1) return null;
  function col(patterns) { return findColumn_(values.header, patterns); }
  var cFirst = col([/first.*name/i, /name/i]);
  var cLast = col([/last.*name/i]);
  var cMobile = col([/mobile/i, /number/i, /phone/i]);
  var cProd = col([/product.*desc/i, /^product/i, /model/i]);
  var cStatus = col([/^jobstatus$/i, /^status$/i, /case.*status/i]);
  var cGsx = col([/gsx/i, /repair.*status/i]);
  var cCrm = col([/^crm_?jobstatusdesc$/i, /crm.*status/i, /crm.*job/i]);
  var cDoc = col([/^repairdoctype$/i, /repair.*doc.*type/i, /doc.*type/i]);
  var cAge = col([/^jobage$/i, /job.*age/i, /^age$/i]);
  var cDate = col([/^casedate$/i, /case.*date/i]);
  var rows = [];
  var want = str_(caseId).toLowerCase();
  values.rows.forEach(function (r) {
    if (str_(r[cCase]).toLowerCase() === want) rows.push(r);
  });
  if (!rows.length) return null;
  var outRows = [];
  outRows.push([chatTxt("<b>Case ID</b>"), chatTxt(str_(rows[0][cCase]).toUpperCase())]);
  var name = ((cFirst !== -1 ? str_(rows[0][cFirst]) : "") + " " + (cLast !== -1 ? str_(rows[0][cLast]) : "")).trim();
  if (name) outRows.push([chatTxt("<b>Customer</b>"), chatPlain(name)]);
  if (cMobile !== -1 && rows[0][cMobile]) outRows.push([chatTxt("<b>Mobile</b>"), chatTxt(str_(rows[0][cMobile]))]);
  if (cProd !== -1) outRows.push([chatTxt("<b>Product</b>"), chatPlain(str_(rows[0][cProd]) || "—")]);
  if (cCrm !== -1) outRows.push([chatTxt("<b>CRM repair status</b>"), chatPlain(str_(rows[0][cCrm]) || "—")]);
  if (cGsx !== -1) outRows.push([chatTxt("<b>GSX repair status</b>"), chatPlain(str_(rows[0][cGsx]) || "—")]);
  if (cDoc !== -1 && rows[0][cDoc]) outRows.push([chatTxt("<b>Repair type</b>"), chatPlain(str_(rows[0][cDoc]))]);
  if (cStatus !== -1 && rows[0][cStatus]) outRows.push([chatTxt("<b>Case status</b>"), chatPlain(str_(rows[0][cStatus]))]);
  if (cAge !== -1 && rows[0][cAge] !== "") outRows.push([chatTxt("<b>Job age</b>"), chatTxt(String(rows[0][cAge]) + " days")]);
  if (cDate !== -1 && rows[0][cDate]) outRows.push([chatTxt("<b>Case opened</b>"), chatPlain(str_(rows[0][cDate]))]);
  var commD = chatCommDateForCase_(ss, caseId);
  outRows.push([chatTxt("<b>Comm updated</b>"), chatPlain(commD || "no updates logged")]);
  if (rows.length > 1) outRows.push([chatTxt('<span class="muted">' + rows.length + " rows carry this Case ID on the service sheet.</span>")]);
  return { title: "🧰 Case " + str_(caseId).toUpperCase(), rows: outRows, note: "From service!CaseId.", wide: true };
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

/* ==========================================================================
 * 14 · CACHING LAYER (script cache + Drive persistence)                   *
 * ======================================================================== */

/**
 * best-effort pane cache: stores each rendered dashboard pane
 * (skipping any over the 100 KB CacheService cap) plus the small frame
 * record (store code + meta line) so doGet can serve repeat loads without a
 * single sheet read. TTL matches the 15-minute background refresh cadence.
 */
// CacheService caps each value at 100 KB — the SM pane alone can be larger
// than that with real data. Every pane is therefore stored as 90 KB chunks:
// key_n = chunk count, key_0 .. key_{n-1} = the pieces. Without this the
// whole-pane cache missed EVERY load, which is why each tab tap recomputed.
var PANE_CACHE_CHUNK = 90000;

function paneCachePut_(cache, key, html) {
  try {
    if (!html) return;
    var n = Math.ceil(html.length / PANE_CACHE_CHUNK);
    if (n > 40) return; // pathological size — serve without caching
    for (var i = 0; i < n; i++) {
      cache.put(key + "_" + i, html.substr(i * PANE_CACHE_CHUNK, PANE_CACHE_CHUNK), 900);
    }
    cache.put(key + "_n", String(n), 900);
  } catch (err) { /* best-effort */ }
}

function paneCacheGet_(cache, key) {
  try {
    var nS = cache.get(key + "_n");
    if (!nS) return null;
    var n = parseInt(nS, 10) || 0, out = "";
    for (var i = 0; i < n; i++) {
      var part = cache.get(key + "_" + i);
      if (part == null) return null;
      out += part;
    }
    return out;
  } catch (err) { return null; }
}

// CacheService alone proved unreliable live (entries can be evicted
// at any moment; the whole SM pane is large). The panes + frame record are
// therefore ALSO persisted to a small Drive file: a "cold" recompute then
// happens only when the data is genuinely due (the 15-minute cadence), never
// because a cache entry vanished. All Drive work is best-effort: any failure
// (e.g. scope not granted) silently falls back to CacheService-only.
var PANE_DRIVE_FILE = "SalesDashboard_paneCache_v2.json";

var PANE_DRIVE_TTL_MS = 900000; // 15 minutes, matches the auto-refresh cadence

function paneDriveGetAll_() {
  try {
    var files = DriveApp.getFilesByName(PANE_DRIVE_FILE);
    if (!files.hasNext()) return null;
    var d = JSON.parse(files.next().getBlob().getDataAsString() || "{}");
    if (!d || !d.t || (Date.now() - d.t) > PANE_DRIVE_TTL_MS) return null;
    return d;
  } catch (err) { return null; }
}

function paneDriveMerge_(single, meta) {
  try {
    var files = DriveApp.getFilesByName(PANE_DRIVE_FILE);
    var d = null;
    if (files.hasNext()) {
      try { d = JSON.parse(files.next().getBlob().getDataAsString() || "{}"); } catch (e1) { d = null; }
    }
    if (!d || typeof d !== "object") d = {};
    d.t = Date.now();
    d.panes = d.panes || {};
    if (single) for (var k in single) if (single[k]) d.panes[k] = single[k];
    if (meta) d.meta = meta;
    var body = JSON.stringify(d);
    var write = files ? null : null;
    files = DriveApp.getFilesByName(PANE_DRIVE_FILE);
    if (files.hasNext()) files.next().setContent(body);
    else DriveApp.createFile(PANE_DRIVE_FILE, body, MimeType.PLAIN_TEXT);
  } catch (err) { /* best-effort */ }
}

/** Read all warm panes + frame meta from the persistent store (or null). */
function paneDriveState_() {
  var d = paneDriveGetAll_();
  if (!d || !d.panes) return null;
  var all = ["index", "cre", "tech", "sm"].every(function (pg) { return !!d.panes[pg]; });
  return all ? d : null;
}

function cachePanes_(data, panes) {
  try {
    var cache = CacheService.getScriptCache();
    ["index", "cre", "tech", "sm"].forEach(function (pg) {
      if (panes[pg]) paneCachePut_(cache, "dash_pane_v6_" + pg, panes[pg]);
    });
    cache.put("dash_meta_v6", JSON.stringify({
      storeCode: (data && data.store && data.store.storeCode) ? data.store.storeCode : "",
      metaText: renderMetaText_(data)
    }), 900);
    // Fresh panes ⇒ previously assembled pages are stale; drop them.
    try {
      cache.removeAll(["dash_pg_v7_index", "dash_pg_v7_cre", "dash_pg_v7_tech", "dash_pg_v7_sm"]);
    } catch (e2) { /* non-fatal */ }
    // also persist to the Drive-backed store (merge, timestamped).
    paneDriveMerge_(panes, {
      storeCode: (data && data.store && data.store.storeCode) ? data.store.storeCode : "",
      metaText: renderMetaText_(data)
    });
  } catch (err) { /* caching is best-effort only */ }
}

/* ==========================================================================
 * 15 · PAGE ASSEMBLY & STYLES                                             *
 * ======================================================================== */

/** Assembles the complete HTML document. */
function buildDashboardPage_(data, initialPage, opts) {
  opts = opts || {};
  var css = getDashboardCss_();
  var storeCode = (opts.storeCode !== undefined) ? opts.storeCode :
    ((data && data.store && data.store.storeCode) ? data.store.storeCode : "");
  // Task 35/43: brand line 1 = the crm!F2 store name verbatim; generic title
  // moves to a second line but stays the page header.
  var pageTitle = (storeCode ? storeCode : "iPlanetcare") + " — Store Performance Dashboard";

  var perfData = opts.paneJson ? null : loadPerformanceData_();

  // ?page=… deep links — resolve which tab shows first, build the
  // anchor-fallback URL (blank outside the live Apps Script env), and stamp
  // the digital clock with the server time so it is never blank.
  var ACTIVE = "index";
  PAGES_ACTIVE: {
    var rq = String(initialPage || "index").toLowerCase();
    if (["index", "cre", "tech", "sm"].indexOf(rq) !== -1) ACTIVE = rq;
  }
  var clockStamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MM-yyyy hh:mm:ss a");

  // the eight legacy pages are replaced by four ROLE pages.
  // INDEX - overall store performance at a glance; CRE - minimum basic
  // details; TECH - job management basics only; SM - the full master report
  // (groups A-G). Fine-grained per-role gating per target!X1:Y50 follows in
  // the next iteration.
  var cached = opts.paneJson || null;
  var PAGES = [
    // [tab label, data-page id, main id, html, meta-subtitle]
    ["INDEX", "index", "pageIndex", cached ? cached.index : renderIndexHtml_(data, perfData), "Overall store performance — every dashboard at a glance"],
    ["CRE",   "cre",   "pageCre",   cached ? cached.cre : renderCreHtml_(data, perfData),     "CRE — walk-in tokens · jobs & invoices created · C-SAT summary"],
    ["TECH",  "tech",  "pageTech",  cached ? cached.tech : renderTechHtml_(perfData),         "TECH — job buckets · bucket aging · communication status"],
    ["SM",    "sm",    "pageSm",    cached ? cached.sm : renderSmHtml_(data, perfData),       "SM — full & detailed master report (groups A-G)"],
  ];
  // Full renders also warm the pane caches (Task 51b) so repeat page loads
  // and tab clicks skip all server compute.
  if (!cached && data) {
    cachePanes_(data, { index: PAGES[0][3], cre: PAGES[1][3], tech: PAGES[2][3], sm: PAGES[3][3] });
  }

  // tabs are plain <button>s — a tap can ONLY flip the panes that
  // were prefetched into the DOM at load. There is no href at all, so a tap
  // can never hit the server, never re-read sheet data, never reload.
  var tabs = PAGES.map(function (pg) {
    var act = (pg[1] === ACTIVE) ? " active" : "";
    return '<button type="button" class="tab-btn' + act + '" data-page="' + pg[1] +
      '" onclick="switchPage(this);return false;">' + pg[0] + "</button>";
  }).join("");

  var mains = PAGES.map(function (pg) {
    return '<main class="wrap" id="' + pg[2] + '" data-pane="' + pg[1] + '" data-meta="' + esc_(pg[4]) + '"' +
      (pg[1] === ACTIVE ? "" : ' style="display:none;"') + ">" + pg[3] + "</main>";
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
    "function paneId(page){return {index:'pageIndex',cre:'pageCre',tech:'pageTech',sm:'pageSm'}[page]||'pageIndex';}" +
    "function paneEl(page){return document.getElementById(paneId(page));}" +
    "function currentPage(){var a=document.querySelector('.tab-btn.active');return a?a.dataset.page:'index';}" +
    "function showPane(page){" +
    "  var mains=document.querySelectorAll('main[data-pane]');" +
    "  for(var i=0;i<mains.length;i++){mains[i].style.display=(mains[i].getAttribute('data-pane')===page)?'block':'none';}" +
    "  var active=paneEl(page);" +
    "  if(active){document.getElementById('meta').textContent=active.getAttribute('data-meta')||'';}" +
    "  try{window.scrollTo(0,0);}catch(e){}" +
    "}" +
    // Keep the address bar truthful without any navigation (best effort —
    // sandboxed hosts may forbid it, and failure there is harmless).
    "function syncUrl(page){try{if(window.history&&history.replaceState){history.replaceState(null,'','?page='+page);}}catch(e){}}" +
    "function switchPage(btn){" +
    "  var tabs=document.querySelectorAll('.tab-btn');" +
    "  for(var i=0;i<tabs.length;i++){tabs[i].classList.toggle('active', tabs[i]===btn);}" +
    "  showPane(btn.dataset.page);" +
    "  syncUrl(btn.dataset.page);" +
    "}" +
    // Tile click → flip to that page (tab highlight follows automatically).
    "function gotoPage(page, el){" +
    "  var tabs=document.querySelectorAll('.tab-btn');" +
    "  for(var i=0;i<tabs.length;i++){tabs[i].classList.toggle('active', tabs[i].dataset.page===page);}" +
    "  showPane(page);" +
    "  syncUrl(page);" +
    "}" +
    "function tickClock(){" +
    "  var el=document.getElementById('liveClock'); if(!el) return;" +
    "  var d=new Date(); function p(n){return (n<10?'0':'')+n;}" +
    "  var h=d.getHours(), ap=h>=12?'PM':'AM'; h=h%12||12;" +
    "  el.textContent=p(d.getDate())+'-'+p(d.getMonth()+1)+'-'+d.getFullYear()+' '+" +
    "    p(h)+':'+p(d.getMinutes())+':'+p(d.getSeconds())+' '+ap;" +
    "}" +
    // Bind clicks via addEventListener too — pure belt-and-braces, since
    // buttons carry no href and thus cannot navigate anyway (Task 57).
    "function bindTabs(){" +
    "  var tabs=document.querySelectorAll('.tab-btn');" +
    "  for(var i=0;i<tabs.length;i++){(function(t){t.addEventListener('click',function(ev){ev.preventDefault();switchPage(t);});})(tabs[i]);}" +
    "}" +
    "try{bindTabs();}catch(e){}" +
    // Back/forward buttons only flip the already-loaded panes — no reload.
    "try{window.addEventListener('popstate',function(){var m=((window.location&&location.search)||'').match(/[?&]page=(index|cre|tech|sm)/);if(m)gotoPage(m[1],null);});}catch(e){}" +
    "try{tickClock(); setInterval(tickClock,1000);}catch(e){}" +
    "" +
    // silent background updater: right after the page paints it
    // fills every pane (so tab switches become pure display flips), and from
    // then on it re-checks for updates every 15 minutes without a reload.
    "function applySections(res,onlyLoading){" +
    "  try{" +
    "    var d=JSON.parse(res);" +
    "    var map={index:'pageIndex',cre:'pageCre',tech:'pageTech',sm:'pageSm'};" +
    "    Object.keys(map).forEach(function(pg){var el=document.getElementById(map[pg]);" +
    "      if(el&&d[pg]!==undefined&&(!onlyLoading||el.querySelector('.pane-loading'))){el.innerHTML=d[pg];}});" +
    "    var active=paneEl(currentPage());" +
    "    if(active){document.getElementById('meta').textContent=active.getAttribute('data-meta')||'';}" +
    "    var n=document.getElementById('note');" +
    "    if(n){n.textContent=(d.meta||'Up to date'); n.className='note ok';}" +
    "  }catch(e){}" +
    "}" +
    "function autoRefresh(prefillOnly){" +
    "  if(!(window.google&&google.script&&google.script.run)) return;" +
    "  google.script.run" +
    "    .withSuccessHandler(function(res){applySections(res,prefillOnly);})" +
    "    .withFailureHandler(function(){})" +
    "    .refreshDashboardSections();" +
    "}" +
    "if(window.google&&google.script&&google.script.run){" +
    "  setInterval(function(){autoRefresh(false);},15*60*1000);" +
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
    "        var map={index:'pageIndex',cre:'pageCre',tech:'pageTech',sm:'pageSm'};" +
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
    "        n.textContent=(d.meta||'Updated just now'); n.className='note ok';" +
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
    "" +
    "// Task 54: per-table export buttons call the single-table endpoint." +
    "function dashExportSection(key,btn){" +
    "  if(btn&&btn.getAttribute){var dk=btn.getAttribute('data-export-key'); if(dk&&typeof key!=='string') key=dk;}" +
    "  if(!key){var n0=document.getElementById('note'); if(n0){n0.textContent='Nothing to export — use the ⬇ Export button on a table.'; n0.className='note err';} return;}" +
    "  if(btn.disabled) return;" +
    "  var old=btn.textContent; var n=document.getElementById('note');" +
    "  if(!(window.google && google.script && google.script.run)){" +
    "    if(n) n.textContent='Export works in the live add-on only.'; return; }" +
    "  btn.disabled=true; btn.textContent='…'; if(n) n.textContent='Exporting table…';" +
    "  google.script.run" +
    "    .withSuccessHandler(function(res){" +
    "      try{ var d=JSON.parse(res);" +
    "        if(n) n.innerHTML='Exported · <a target=\"_top\" href=\"'+d.url+'\" style=\"color:#63e6b7;font-weight:800\">open '+d.name+'</a>';" +
    "      }catch(e){ if(n) n.textContent='Export finished.'; }" +
    "      btn.disabled=false; btn.textContent=old; })" +
    "    .withFailureHandler(function(err){" +
    "      if(n) n.textContent='Export failed: '+(err && err.message ? err.message : err);" +
    "      btn.disabled=false; btn.textContent=old; })" +
    "    .exportDashboardSectionToSpreadsheet(key);" +
    "}" +
    "" +
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
    "      var h='<div class=\\\"chat-title\\\">'+res.title+'</div>';" +
    "      if(res.rows && res.rows.length){" +
    "        h+='<table class=\\\"chat-table\\\"><tbody>';" +
    "        for(var i=0;i<res.rows.length;i++){h+='<tr>'+res.rows[i].join('')+'</tr>';}" +
    "        h+='</tbody></table>';" +
    "      }" +
    "      if(res.note){h+='<div class=\\\"muted\\\">'+res.note+'</div>';}" +
    "      var pl=document.getElementById('chatPanel'); if(pl){pl.classList.toggle('half', !!res.wide);}" +
    "      finish(h);" +
    "    })" +
    "    .withFailureHandler(function(err){" +
    "      finish('Search failed: '+(err && err.message ? err.message : err));" +
    "    })" +
    "    .serverQuery(q);" +
    "}" +
    // case detail popover. Values arrive in the #caseDetailData
    // island on the jobs page; escape everything before innerHTML.
    "function escText_(s){return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}" +
    "var CASE_DETAILS_CACHE=null;" +
    "function caseDetailsMap_(){if(CASE_DETAILS_CACHE===null){" +
    "  var el=document.getElementById('caseDetailData');" +
    "  try{CASE_DETAILS_CACHE=el?JSON.parse(el.textContent||'{}'):{};}catch(e){CASE_DETAILS_CACHE={};}" +
    "}return CASE_DETAILS_CACHE;}" +
    "function openCaseDetails(cid){" +
    "  var d=(caseDetailsMap_()||{})[cid]; if(!d)return;" +
    "  var rows=[['Case ID',d.caseId],['Job Created Date',d.createdDate]," +
    "    ['Repair Created Date',d.repairDate],['Repair Bucket',d.bucket]," +
    "    ['Customer Name',d.name],['Product',d.product]," +
    "    ['Communication Updated',d.commDate],['Ready For Pickup',d.readyDate]," +
    "    ['GSX Repair Status',d.gsxStatus]];" +
    "  var h='';for(var i=0;i<rows.length;i++){" +
    "    h+='<tr><td>'+escText_(rows[i][0])+'</td><td>'+escText_(rows[i][1])+'</td></tr>';}" +
    "  document.getElementById('cmTable').innerHTML=h;" +
    "  document.getElementById('cmTitle').textContent='CASE '+d.caseId;" +
    "  document.getElementById('caseModal').style.display='flex';" +
    "}" +
    "function closeCaseDetails(){var m=document.getElementById('caseModal');if(m)m.style.display='none';}" +
    "document.addEventListener('keydown',function(e){if(e.key==='Escape')closeCaseDetails();});" +
    "</script>";

  var bizMeta = (opts.metaText !== undefined) ? opts.metaText : renderMetaText_(data);

  return (
    "<!-- BUILD TASK61 · export keys pinned on buttons + caches rebased · prefetched panes · tabs never navigate -->" +
    "<!DOCTYPE html><html><head><base target=\"_top\"><meta charset=\"utf-8\">" +
    "<title>" + esc_(pageTitle) + "</title><style>" + css + "</style></head><body>" +
    // plain dark-ash backdrop for EVERY page — no image at all.
    '<header class="topbar"><div class="topbar-inner">' +
    "<div class=\"topbar-left\">" +
    // brand shows the crm!F2 store name as-is (fallback when blank).
    '<div class="brand">' + esc_(storeCode ? storeCode : "iPlanetcare") + "</div>" +
    '<div class="brand-sub">Store Performance Dashboard</div>' +
    '<div class="meta" id="meta">' + esc_(bizMeta) + "</div>" +
    '<div class="note" id="note"></div>' +
    "</div>" +
    // ---- live clock, top-middle: Task 30 — gif removed; blue LED digits only ----
    '<div class="topbar-clock" role="timer">' +
    '<span class="clock-text" id="liveClock">' + esc_(clockStamp) + "</span>" +
    "</div>" +
    // refresh is a plain button — it can ONLY run dashRefresh(),
    // which re-reads the source data server-side and swaps the new values in
    // place. No ?refresh=1 navigation, no full page load.
    '<button type="button" class="refresh-btn" onclick="dashRefresh(this);return false;">' +
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
          '<span class="chat-sub">ask by mobile number or Case ID</span>' +
          '<button class="chat-close" onclick="chatToggle()" aria-label="Close chat">&times;</button>' +
        "</div>" +
        '<div class="chat-log" id="chatLog">' +
          '<div class="chat-msg bot"><b>Ask me anything</b> — I search <b>every sheet</b> of this ' +
            "spreadsheet: customer mobiles, Case IDs, names, products, statuses, dates, " +
            "amounts… Exact matches only, straight from the sheets.</div>" +
        "</div>" +
        '<div class="chat-inputrow">' +
          '<input class="chat-input" id="chatInput" type="text" autocomplete="off" ' +
            'placeholder="Search every sheet — mobile, Case ID, name, product…" ' +
            "onkeydown=\"if(event.key==='Enter')chatSend()\">" +
          '<button class="chat-send" onclick="chatSend()" aria-label="Send">➤</button>' +
        "</div>" +
      "</div>" +
    "</div>" +
    clientScript +
    "</body></html>"
  );
}

function getDashboardCss_() {
  // Presentation-only replacement. All business rules, calculations,
  // renderers, markup generators, endpoints and client-side handlers remain unchanged.
  return [
    ":root{--page:#eef2f7;--ink:#17263a;--muted:#64748b;--line:#dbe3ec;--surface:#ffffff;--navy:#172b43;--navy2:#203a58;--achieved:#16a879;--target:#3488d8;--balance:#c58a18;--percent:#8a57c2;--accent:#2689a6;--danger:#d94c59;--completed:#16a879;--radius:14px;--shadow:0 5px 18px rgba(20,40,65,.07);--f-th:.66rem;--f-td:.72rem;--f-label:.62rem;--f-val:1.12rem}",
    "*{box-sizing:border-box}",
    "html{font-size:16px;scroll-behavior:smooth}",
    "body{margin:0;min-height:100vh;background:linear-gradient(180deg,#f4f7fb 0%,#edf2f7 100%) fixed;color:var(--ink);font-family:Inter,ui-sans-serif,system-ui,-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;font-size:14px;line-height:1.5}",
    "button,input,select,textarea{font:inherit}",
    ".wrap{width:min(100%,1800px);margin:0 auto;padding:22px clamp(14px,2vw,30px) 48px;overflow:visible}",
    ".topbar{position:sticky;top:0;z-index:50;padding:14px clamp(14px,2vw,30px) 12px;background:rgba(255,255,255,.96);border-bottom:1px solid #dce4ed;box-shadow:0 3px 14px rgba(20,40,65,.05)}",
    ".topbar-inner{position:relative;width:min(100%,1800px);margin:0 auto;display:flex;align-items:center;justify-content:space-between;gap:18px;flex-wrap:wrap}",
    ".topbar-left{display:flex;flex-direction:column;gap:3px;min-width:220px}",
    ".brand{font-size:clamp(1.35rem,2.1vw,1.8rem);font-weight:850;letter-spacing:-.045em;line-height:1.12;color:#172b43}",
    ".brand-sub{color:#64748b;font-size:.75rem;font-weight:700;letter-spacing:.04em}",
    ".topbar .meta{font-size:.72rem;color:#64748b;font-weight:600}",
    ".topbar .note{min-height:14px;font-size:.7rem;font-weight:650}",
    ".topbar .note.ok{color:#087c58}.topbar .note.err{color:#c53b48}",
    ".topbar-clock{position:absolute;left:50%;top:50%;transform:translate(-50%,-50%);font-size:.78rem;font-weight:750;color:#415875;font-variant-numeric:tabular-nums;white-space:nowrap}",
    ".tabs-row{width:100%;display:flex;justify-content:center;margin-top:12px}",
    ".tabs{display:flex;gap:4px;width:min(900px,100%);padding:5px;border:1px solid #dfe6ef;border-radius:12px;background:#f3f6fa;box-shadow:inset 0 1px 2px rgba(20,40,65,.035)}",
    ".tab-btn{appearance:none;flex:1;min-width:0;border:0;border-radius:9px;padding:10px 12px;background:transparent;color:#53657a;font-size:.73rem;font-weight:800;letter-spacing:.025em;cursor:pointer;white-space:nowrap;transition:background .15s,color .15s,box-shadow .15s}",
    ".tab-btn:hover{background:#e7edf5;color:#172b43}.tab-btn.active{background:#172b43;color:#fff;box-shadow:0 3px 8px rgba(23,43,67,.18)}",
    ".refresh-btn{flex-shrink:0;display:inline-flex;align-items:center;gap:8px;border:1px solid #172b43;border-radius:9px;padding:9px 14px;background:#172b43;color:#fff;font-size:.73rem;font-weight:800;cursor:pointer;box-shadow:0 3px 9px rgba(23,43,67,.14)}",
    ".refresh-btn:hover{background:#254360}.refresh-btn:disabled{opacity:.6;cursor:wait}.refresh-btn .spin{display:none;width:12px;height:12px;border:2px solid rgba(255,255,255,.4);border-top-color:#fff;border-radius:50%;animation:spin .7s linear infinite}.refresh-btn.loading .spin{display:inline-block}@keyframes spin{to{transform:rotate(360deg)}}",
    ".section{margin:0 0 26px;overflow:visible}.section h2{margin:0 0 12px;padding:0 0 9px;background:none!important;color:#1c3048!important;border-bottom:1px solid #dce5ef;border-radius:0!important;box-shadow:none!important;font-size:.96rem;font-weight:850;letter-spacing:-.015em}.section .body{padding:0}",
    ".cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(min(100%,205px),1fr));gap:12px;margin:0 0 18px}",
    ".card{position:relative;min-width:0;overflow:hidden;padding:15px;background:#fff!important;color:#20334b!important;border:1px solid #e0e7ef!important;border-radius:13px!important;box-shadow:var(--shadow)!important}",
    ".card:before{content:'';position:absolute;left:0;top:0;bottom:0;width:3px;background:#b8c9dc;opacity:.8}",
    ".card.rtile.done:before,.card.done:before{background:#18a879}",
    ".card .label{display:block;color:#718198!important;font-size:.61rem!important;font-weight:800!important;letter-spacing:.055em!important;line-height:1.25!important;text-transform:uppercase!important;white-space:normal}",
    ".card .value{display:block;margin-top:6px;color:#172b43!important;font-size:1.3rem!important;font-weight:850!important;letter-spacing:-.035em;line-height:1.15;font-variant-numeric:tabular-nums;white-space:nowrap}",
    ".card .sub,.card .muted{color:#738197!important;font-size:.66rem}",
    ".metric{display:flex;align-items:center;justify-content:space-between;gap:10px;margin-top:9px;padding:8px 10px;border:1px solid #edf1f5!important;border-radius:8px;background:#f7f9fc!important}",
    ".metric span{color:#718198!important;font-size:.61rem!important;font-weight:750!important}.metric strong{font-size:.78rem!important;font-weight:850!important;font-variant-numeric:tabular-nums;color:#283e58}",
    ".metric.t strong,.tgt,.grphead-count{color:#3488d8!important}.metric.a strong,.value.pos,.pos{color:#13996e!important}.metric.p strong,.pct{color:#8656bd!important}.metric.b strong{color:#b17b13!important}",
    ".dreq{display:flex;align-items:center;justify-content:space-between;gap:10px;margin-top:9px;padding:9px 10px;background:#f5f8fc!important;border:1px solid #e5ebf2!important;border-left:3px solid #83a9d2!important;border-radius:8px}",
    ".dreq span{font-size:.62rem!important;font-weight:750!important;color:#718198!important}.dreq strong{font-size:.82rem!important;font-weight:850!important;color:#203a58!important;font-variant-numeric:tabular-nums}",
    ".trend{display:flex;gap:6px;flex-wrap:wrap;margin-top:8px}.trend span{flex:1;min-width:90px;padding:6px 8px;border:1px solid #e8edf3;border-radius:7px;background:#f7f9fc;font-size:.58rem;font-weight:700;color:#738197}.trend b{color:#263c56;font-variant-numeric:tabular-nums;white-space:nowrap}",
    ".subpanel{min-width:0;margin:0 0 14px;padding:14px;background:#fff!important;border:1px solid #e0e7ef!important;border-radius:13px;box-shadow:var(--shadow)!important;color:#263a52!important}",
    ".subpanel-head{display:flex;align-items:center;justify-content:space-between;gap:10px;margin-bottom:10px;color:#263a52!important;font-size:.82rem!important;font-weight:850!important}",
    ".rt-head h3{margin:0;color:#263a52!important;font-size:.85rem!important;font-weight:850!important}",
    ".legend-card{padding:10px 12px;background:#f7f9fc!important;border:1px solid #e5ebf2!important;border-radius:10px;color:#43546a!important;box-shadow:none!important}",
    ".legend,.small,.note{font-size:.68rem;font-weight:550;color:#68788f!important}.muted{color:#718198!important}",
    ".status,.badge{display:inline-flex;align-items:center;gap:4px;padding:4px 8px;border-radius:999px;background:#edf3f9!important;color:#47617d!important;font-size:.58rem!important;font-weight:800!important;letter-spacing:.025em;text-transform:uppercase;border:1px solid #e2eaf2!important}",
    ".status.on,.badge.good{background:#e8f7f1!important;color:#087b58!important;border-color:#cdeee0!important}.status.pending,.badge.warn{background:#fff6df!important;color:#956b12!important;border-color:#f7e8bc!important}.badge.bad{background:#fff0f0!important;color:#b83e48!important;border-color:#f4d5d7!important}",
    ".comm-yes{color:#0f9a70!important}.neg{color:#cf4653!important}.bal-neg,.bal-warn,.age-warn{color:#ac7611!important}",
    ".scroll-x{width:100%;overflow-x:auto;overscroll-behavior-x:contain;border-radius:10px}",
    "table{max-width:100%;border-collapse:separate;border-spacing:0;font-variant-numeric:tabular-nums}",
    "table.ptable,table.matrix,table.jobtable,table.csat-table{width:100%;min-width:760px;table-layout:auto;background:#fff!important;border:1px solid #e0e7ef!important;border-radius:11px;overflow:hidden;box-shadow:0 3px 12px rgba(20,40,65,.045)!important}",
    "table th{position:sticky;top:0;padding:10px 11px;background:#eef3f8!important;color:#405671!important;border-bottom:1px solid #dce5ef!important;text-align:left;font-size:var(--f-th)!important;font-weight:850!important;letter-spacing:.045em!important;text-transform:uppercase}",
    "table td{padding:9px 11px;background:#fff!important;color:#34485f!important;border-bottom:1px solid #edf1f5!important;font-size:var(--f-td)!important;font-weight:600!important}",
    "tbody tr:nth-child(even) td{background:#f8fafc!important}tbody tr:hover td{background:#f0f5fa!important}tbody tr:last-child td{border-bottom:0}",
    "table.hdr-deep tr.total-row td{background:#eaf0f7!important;color:#1c3048!important;font-weight:850!important}.grphead td{background:#eaf0f7!important;color:#405671!important;font-weight:850!important}",
    ".card .small,.card .note,.subpanel .small,.subpanel .note,table .small,table .note{color:#718198!important}.card a,.subpanel a{color:#286ea9}",
    ".empty-note,.sec-desc2{color:#718198!important}.sec-desc,p.defline,.grphead-sub{color:#718198}.grphead-sub{font-size:.65rem}",
    ".chat-widget{z-index:100}.chat-fab{border:1px solid #dbe4ed!important;background:#172b43!important;color:#fff!important;box-shadow:0 7px 22px rgba(23,43,67,.2)!important;border-radius:999px!important}.chat-panel{border:1px solid #dbe4ed!important;background:#fff!important;color:#253951!important;box-shadow:0 14px 42px rgba(23,43,67,.2)!important;border-radius:14px!important}.chat-head{background:#172b43!important;color:#fff!important}.chat-sub{color:#d2deeb!important}.chat-msg{border-radius:10px}.chat-msg.bot{background:#f1f5f9!important;color:#263b54!important}.chat-inputrow{background:#fff!important;border-top:1px solid #e5ebf2}.chat-input{border:1px solid #dce5ef!important;border-radius:8px!important;color:#20364f!important}.chat-send{background:#172b43!important;color:#fff!important;border-radius:8px!important}",
    ".export-btn,.section button:not(.tab-btn):not(.refresh-btn){border-radius:8px}",
    ".scroll-x::-webkit-scrollbar{height:9px}.scroll-x::-webkit-scrollbar-thumb{background:#c5d1df;border-radius:99px}.scroll-x::-webkit-scrollbar-track{background:#f0f3f7}",
    ".wrap img,.wrap svg,.wrap video{max-width:100%}",
    ".sec-desc,p.defline{display:none}",
    "@media(min-width:1500px){.cards{grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:14px}.wrap{padding-left:30px;padding-right:30px}}",
    "@media(max-width:900px){.topbar-clock{position:static;transform:none;margin-left:auto}.topbar-inner{gap:10px}.tabs{overflow-x:auto;justify-content:flex-start}.tab-btn{flex:0 0 auto;padding:10px 13px}.wrap{padding:16px 14px 36px}}",
    "@media(max-width:600px){.topbar{padding:12px}.topbar-left{min-width:0;width:calc(100% - 120px)}.topbar-clock{order:3;width:100%;margin:0;text-align:left}.refresh-btn{font-size:.68rem;padding:8px 10px}.tabs-row{margin-top:9px}.tabs{width:100%;border-radius:10px}.tab-btn{font-size:.67rem;padding:9px 11px}.cards{grid-template-columns:repeat(2,minmax(0,1fr));gap:9px}.card{padding:12px}.card .value{font-size:1.1rem!important}.section{margin-bottom:20px}.subpanel{padding:11px}.wrap{padding:13px 10px 30px}}",
    "@media(prefers-reduced-motion:reduce){*,*:before,*:after{scroll-behavior:auto!important;animation-duration:.01ms!important;transition-duration:.01ms!important}}"
  ].join("");
}

/* ==========================================================================
 * 16 · PUBLIC ENDPOINTS                                                   *
 * ======================================================================== */

/**
 * Web App entry point. Apps Script calls this automatically
 * whenever someone opens the deployed Web App URL (see the setup steps in
 * the header comment above — Deploy → New deployment → Web app).
 *
 * Fast path: the fully-rendered HTML is cached for CONFIG.CACHE_SECONDS so
 * repeat visits are near-instant. Opening the URL with "?refresh=1" (what
 * the on-page Refresh button does) skips the cache and recomputes.
 */
function doGet(e) {
  var wantsFresh = !!(e && e.parameter && e.parameter.refresh);
  // deep-linkable tabs: "?page=cre" serves the CRE pane already
  // visible, so navigation keeps working even where inline JS is stripped.
  var page = (e && e.parameter && e.parameter.page) ? String(e.parameter.page).toLowerCase() : "index";
  if (["index", "cre", "tech", "sm"].indexOf(page) === -1) page = "index";
  var cache = CacheService.getScriptCache();

  // ---- fast page loads, three tiers --------------------------------------
  // (a) whole-page cache, chunked per page — zero compute, all panes embedded.
  // (b) pane caches (script + Drive) + frame record warm → cheap rebuild.
  // (c) cold → the load itself PREFETCHES everything: all four panes are
  //     computed, embedded and cached in one pass (Task 56), so once loading
  //     completes switching between pages never recomputes.
  var html = null;
  if (!wantsFresh) {
    try { html = paneCacheGet_(cache, "dash_pg_v7_" + page); } catch (e0) { html = null; }
    if (!html) {
      try {
        var metaJ = cache.get("dash_meta_v6");
        if (metaJ) {
          var panes = {}, cold = false;
          ["index", "cre", "tech", "sm"].forEach(function (pg) {
            var h = paneCacheGet_(cache, "dash_pane_v6_" + pg);
            if (h == null) cold = true; else panes[pg] = h;
          });
          if (!cold) {
            var meta = JSON.parse(metaJ);
            html = buildDashboardPage_(null, page, {
              storeCode: meta.storeCode, metaText: meta.metaText, paneJson: panes
            });
          }
        }
      } catch (err) { html = null; }
    }
    if (!html) {
      // (b2) Task 55 — persistent Drive-backed pane store: survives cache
      // evictions, so warm loads stay warm between taps.
      try {
        var d = paneDriveState_();
        if (d) {
          html = buildDashboardPage_(null, page, {
            storeCode: d.meta ? d.meta.storeCode : "", metaText: d.meta ? d.meta.metaText : "",
            paneJson: d.panes
          });
        }
      } catch (err2) { html = null; }
    }
  }

  if (!html) {
    var loaded = loadDashboardData_();
    var perfData = loadPerformanceData_();
    var data = loaded.data;
    // everything is prefetched DURING the load itself —
    // the cold load computes and embeds all four panes in one pass and caches
    // them, so once loading completes no tab tap ever touches the sheets again.
    var panes2 = buildAllPanes_(data, perfData);
    cachePanes_(data, panes2);
    html = buildDashboardPage_(data, page, { paneJson: panes2 });
  }
  try { paneCachePut_(cache, "dash_pg_v7_" + page, html); } catch (err3) { /* skip */ }

  return HtmlService.createHtmlOutput(html)
    .setTitle("Sales Performance Dashboard")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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
  var panes = buildAllPanes_(data, perfData);
  // refill the per-pane caches, so the next page load (or tab
  // navigation via anchors) is served with zero sheet reads.
  cachePanes_(data, panes);

  return JSON.stringify({
    meta: renderMetaText_(data),
    index: panes.index,
    cre: panes.cre,
    tech: panes.tech,
    sm: panes.sm
  });
}

/* ==========================================================================
 *  ✅  END OF FILE — SalesDashboard.gs
 * ------------------------------------------------------------------------ *
 *  PASTE CHECK: this banner must be the LAST thing in your Apps Script editor.
 *  If you cannot see it, your paste was cut short and the script will fail
 *  with the unexpected-end-of-input syntax error — copy the raw
 *  file again (GitHub → SalesDashboard.gs → Copy raw file) and paste it whole.
 * ======================================================================== */
