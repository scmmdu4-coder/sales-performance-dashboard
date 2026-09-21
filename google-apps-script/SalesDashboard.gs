/**
 * ============================================================================
 *  SALES PERFORMANCE DASHBOARD — Google Apps Script
 * ============================================================================
 *  HOW TO USE
 *  ----------
 *  1. Open your Google Sheet (the one with the "crm" and "target" tabs).
 *  2. Extensions → Apps Script.
 *  3. Delete anything in the editor and paste this entire file in.
 *  4. Click the disk icon to save (name the project anything you like).
 *  5. In the function dropdown at the top, choose "buildDashboard" and
 *     click ▶ Run. The first run will ask you to authorize the script —
 *     accept it (it only touches this spreadsheet).
 *  6. Reload the spreadsheet — you'll see a new "📊 Sales Dashboard" menu.
 *     Use it any time to rebuild the dashboard after the crm/target sheets
 *     change.
 *
 *  WHAT IT BUILDS (tab "Dashboard")
 *  ---------------------------------
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
  // Day-of-month numbers treated as holidays / non-working days this month.
  // Edit this list to match your store's actual holiday calendar.
  HOLIDAYS: [1, 15, 28],
};

/* ============================================================================
 * MENU
 * ========================================================================= */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📊 Sales Dashboard")
    .addItem("Build / Refresh Dashboard", "buildDashboard")
    .addToUi();
}

/* ============================================================================
 * ENTRY POINT
 * ========================================================================= */

function buildDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var crmSheet = ss.getSheetByName(CONFIG.CRM_SHEET_NAME);
  var targetSheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);

  if (!crmSheet) throw new Error('Sheet "' + CONFIG.CRM_SHEET_NAME + '" not found.');
  if (!targetSheet) throw new Error('Sheet "' + CONFIG.TARGET_SHEET_NAME + '" not found.');

  var crmValues = crmSheet.getDataRange().getValues();
  var targetValues = targetSheet.getDataRange().getValues();

  var allRows = parseCrmRows(crmValues);
  var target = parseTargetSheet(targetValues);
  var data = computeDashboard(allRows, target);

  writeDashboard(ss, data);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Dashboard rebuilt: " + data.store.totalInvoices + " SALES rows, " +
      data.store.uniqueInvoices + " unique invoices.",
    "Sales Dashboard",
    6,
  );
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

function parseGoogleReview_(rows) {
  function numbersIn(row) {
    var out = [];
    if (!row) return out;
    for (var c = 5; c <= 12; c++) {
      var raw = str_(row[c]).replace(/[₹\s]/g, "");
      if (!/^-?\d+(?:,\d+)*(?:\.\d+)?%?$/.test(raw)) continue;
      out.push(parseMoney_(raw));
    }
    return out;
  }
  var lastRow = Math.min(5, rows.length - 1);
  for (var r = 0; r <= lastRow; r++) {
    var nums = numbersIn(rows[r]);
    if (nums.length < 2) continue;
    var previousMonth = nums[0], currentMonth = nums[1];
    var difference = nums.length >= 3 && nums[2] !== 0 ? nums[2] : Math.max(0, currentMonth - previousMonth);
    return { previousMonth: previousMonth, currentMonth: currentMonth, difference: difference };
  }
  return { previousMonth: 0, currentMonth: 0, difference: 0 };
}

function parseTargetSheet(rows) {
  function at(r, c) { return str_(rows[r] && rows[r][c]); }

  // ---- ANCHOR: store target (A1:B5 -> 0-indexed rows 1..5, col B=1) ----
  var store = {
    service: parseMoney_(at(1, 1)),
    accessory: parseMoney_(at(2, 1)),
    repair: parseMoney_(at(3, 1)),
    total: parseMoney_(at(4, 1)),
    grossProfit: parseMoney_(at(5, 1)),
  };

  // ---- ANCHOR: salesperson targets (D1:E6 header + rows) ----
  var people = {};
  var personKeys = [];
  var headerCells = rows[0] || [];
  for (var c = 3; c < Math.min(headerCells.length, 5); c++) {
    var name = headerCells[c];
    if (name) {
      var key = normalizePerson_(name);
      if (personKeys.indexOf(key) === -1) personKeys.push(key);
    }
  }
  personKeys.forEach(function (key) {
    people[key] = { service: 0, accessory: 0, repair: 0, total: 0, grossProfit: 0 };
  });
  var colIndex = {};
  personKeys.forEach(function (key, i) { colIndex[key] = 3 + i; });
  for (var r = 1; r <= 5; r++) {
    var label = at(r, 2).toLowerCase();
    var slot = null;
    if (label.indexOf("service") === 0) slot = "service";
    else if (label.indexOf("accessor") === 0) slot = "accessory";
    else if (label.indexOf("repair") === 0) slot = "repair";
    else if (label.indexOf("total") === 0) slot = "total";
    else if (label.indexOf("gross") === 0) slot = "grossProfit";
    if (!slot) continue;
    personKeys.forEach(function (key) {
      people[key][slot] = parseMoney_(at(r, colIndex[key]));
    });
  }

  // ---- ANCHOR: google review (G1:I3) ----
  var googleReview = parseGoogleReview_(rows);

  // ---- ANCHOR: KPI grid (row 9 target / row 10 value / row 11 approx, 0-indexed rows 8/9/10) ----
  var kpiLabels = {
    CAREPLUS: "Care+", PROTECTPLUS: "Protect+", AMC: "AMC", BUNDLE: "Bundle",
    DISPLAY: "Display", BATTERY: "Battery", GLASS: "Glass", OTHERS: "Others",
    SRD: "SRD", SHC: "SHC", ARD: "ARD",
  };
  var KPI_LAYOUT = [
    ["CAREPLUS", 1], ["PROTECTPLUS", 2], ["AMC", 3], ["BUNDLE", 4],
    ["DISPLAY", 6], ["BATTERY", 7], ["GLASS", 8], ["OTHERS", 9],
    ["SRD", 10], ["SHC", 11], ["ARD", 12],
  ];
  var kpis = [];
  KPI_LAYOUT.forEach(function (pair) {
    var code = pair[0], colI = pair[1];
    var targetVal = parseMoney_(at(8, colI));
    if (targetVal <= 0) return;
    kpis.push({
      category: colI <= 4 ? "AR" : colI <= 9 ? "RR" : "SR",
      code: code,
      label: kpiLabels[code] || code,
      target: targetVal,
      value: parseMoney_(at(9, colI)),
      approxAmount: parseMoney_(at(10, colI)),
    });
  });

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
      var bp = day.byPerson[pk] || { serviceRevenue: 0, accessoryRevenue: 0, repairRevenue: 0, totalRevenue: 0, grossProfit: 0 };
      bp.serviceRevenue += rt === "service" ? value : 0;
      bp.accessoryRevenue += rt === "accessory" ? value : 0;
      bp.repairRevenue += rt === "repair" ? value : 0;
      bp.totalRevenue += value;
      bp.grossProfit += r.grossProfit;
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
    var svcInc = incentiveForPct_(target.serviceIncentive.service, pct_(ach.service, t.service));
    var accInc = incentiveForPct_(target.serviceIncentive.accessory, pct_(ach.accessory, t.accessory));
    function mk(tg, a, inc) { return { target: tg, achieved: a, pct: pct_(a, tg), balance: Math.max(0, tg - a), incentive: inc }; }
    return {
      key: key, name: titleCase_(key), invoiceCount: mineInvoices,
      total: mk(t.total, ach.total, svcInc + accInc), service: mk(t.service, ach.service, svcInc),
      accessory: mk(t.accessory, ach.accessory, accInc), repair: mk(t.repair, ach.repair, 0),
      totalIncentive: svcInc + accInc,
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
      var vals = [ti === 0 ? p.name : "", label, money_(b.target), money_(b.achieved), b.pct + "%",
        money_(b.balance), b.incentive ? money_(b.incentive) : "—", ti === 0 ? p.invoiceCount : ""];
      sheet.getRange(row, 1, 1, 8).setValues([vals]);
      sheet.getRange(row, 5).setFontColor(pctColor_(b.pct)).setFontWeight("bold");
      sheet.getRange(row, 1, 1, 8).setBackground(idx % 2 === 0 ? COLORS.stripe : "#ffffff");
      if (ti === 0) sheet.getRange(row, 1).setFontWeight("bold");
      row++;
    });
  });
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
  row = sectionHeader_(sheet, row, "5 · STORE KPI PARAMETERS (target sheet grid)", 6);
  row = tableHeaderRow_(sheet, row, ["KPI", "Category", "Target", "Achieved", "Balance", "Achieved Value / GP"]);
  var allKpis = data.kpi.ar.concat(data.kpi.rr, data.kpi.sr);
  allKpis.forEach(function (k, idx) {
    sheet.getRange(row, 1, 1, 6).setValues([[k.label, k.category, k.target, k.achieved, k.balance,
      money_(k.achievedValue) + " / " + money_(k.achievedGp)]]);
    sheet.getRange(row, 5).setFontColor(k.balance === 0 ? COLORS.good : COLORS.warn).setFontWeight("bold");
    sheet.getRange(row, 1, 1, 6).setBackground(idx % 2 === 0 ? COLORS.stripe : "#ffffff");
    row++;
  });
  row++;
  sheet.getRange(row, 1, 1, 4).setValues([["Google Review", "prev " + data.kpi.googleReview.previousMonth,
    "current " + data.kpi.googleReview.currentMonth,
    (data.kpi.googleReview.difference >= 0 ? "+" : "") + data.kpi.googleReview.difference]]);
  sheet.getRange(row, 1, 1, 4).setFontWeight("bold");
  row += 2;
  return row;
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
  row = sectionHeader_(sheet, row, "7 · DATE-WISE SALES REVENUE ANALYSIS", 8);
  var perTypeDailyTarget = function (key) {
    var bucket = data.store[key];
    return data.store.balanceWorkingDays > 0 ? bucket.balance / data.store.balanceWorkingDays : 0;
  };
  sheet.getRange(row, 1, 1, 4).setValues([["Daily Target — Service", money_(perTypeDailyTarget("service")),
    "Daily Target — Accessory", money_(perTypeDailyTarget("accessory"))]]).setFontWeight("bold");
  row++;
  sheet.getRange(row, 1, 1, 4).setValues([["Daily Target — Repair", money_(perTypeDailyTarget("repair")),
    "Daily Target — Total", money_(data.store.total.requiredPerDay)]]).setFontWeight("bold");
  row += 2;

  var header = ["Date", "Weekday", "Unique Invoices", "Service Rev", "Accessory Rev", "Repair Rev", "Total Rev", "Gross Profit"];
  row = tableHeaderRow_(sheet, row, header);
  data.days.forEach(function (d, idx) {
    sheet.getRange(row, 1, 1, 8).setValues([[d.undated ? "No date" : d.dateKey, d.weekday, d.uniqueInvoices,
      money_(d.serviceRevenue), money_(d.accessoryRevenue), money_(d.repairRevenue), money_(d.totalRevenue), money_(d.grossProfit)]]);
    sheet.getRange(row, 1, 1, 8).setBackground(d.holiday ? "#fef3c7" : (idx % 2 === 0 ? COLORS.stripe : "#ffffff"));
    row++;
  });
  return row + 1;
}
