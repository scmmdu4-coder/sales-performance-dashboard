// Data layer: raw Google Sheets CSV parsing + target-sheet interpretation.
// The "crm" sheet (gid=0) holds item-level invoice rows; the "target" sheet
// holds store / salesperson targets, KPI grid, Google review counters and
// incentive slabs.

export const SHEET_ID = "1Y7U0hfbj5zyYAQaxTSka2AqSYu7OSeI0ySMjCTosPfo";
export const CRM_GID = "0";

/** Export URLs for the two sheets we consume (public / link-shared). */
export function buildSheetsFetchUrls() {
  return {
    crmUrl: `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:csv&headers=1&gid=${CRM_GID}`,
    targetUrl: `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:csv&headers=1&sheet=target`,
  };
}

// ---------------------------------------------------------------------------
// Minimal RFC-4180 CSV parser (quoted fields, embedded commas/newlines)
// ---------------------------------------------------------------------------

export function parseCsv(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = "";
  let inQuotes = false;
  let i = 0;
  while (i < text.length) {
    const ch = text[i];
    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i += 2;
          continue;
        }
        inQuotes = false;
        i++;
        continue;
      }
      field += ch;
      i++;
      continue;
    }
    if (ch === '"') {
      inQuotes = true;
      i++;
      continue;
    }
    if (ch === ",") {
      row.push(field);
      field = "";
      i++;
      continue;
    }
    if (ch === "\r") {
      i++;
      continue;
    }
    if (ch === "\n") {
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
      i++;
      continue;
    }
    field += ch;
    i++;
  }
  row.push(field);
  rows.push(row);
  // Drop fully-empty trailing rows
  return rows.filter((r) => r.some((c) => c.trim() !== ""));
}

export function csvToObjects(rows: string[][]): Record<string, string>[] {
  if (rows.length === 0) return [];
  const header = rows[0].map((h) => h.trim());
  return rows.slice(1).map((r) => {
    const obj: Record<string, string> = {};
    header.forEach((h, idx) => {
      obj[h] = (r[idx] ?? "").trim();
    });
    return obj;
  });
}

// ---------------------------------------------------------------------------
// CRM row model (subset of columns we need)
// ---------------------------------------------------------------------------

export interface CrmRow {
  invoiceCode: string;
  invoiceDate: string; // raw "DD-MM-YYYY HH:mm"
  dayKey: string; // "DD-MM-YYYY"
  day: number;
  salesPerson: string; // raw
  employeeCode: string;
  partType: string;
  subCategory: string;
  revenueType: string;
  itemCode: string;
  itemDescription: string;
  schemeCode: string;
  quantity: number;
  netAmount: number;
  taxableAmount: number;
  grossProfit: number;
}

function num(s: string): number {
  const v = Number.parseFloat((s ?? "").replace(/,/g, "").trim());
  return Number.isFinite(v) ? v : 0;
}

export function parseCrmRows(rows: string[][]): CrmRow[] {
  const objects = csvToObjects(rows);
  const out: CrmRow[] = [];
  for (const o of objects) {
    const invoiceCode = (o["InvoiceCode"] ?? "").trim();
    if (!invoiceCode) continue;
    const invoiceDate = (o["InvoiceDate"] ?? "").trim();
    const dayKey = invoiceDate.split(" ")[0] ?? invoiceDate;
    const day = Number.parseInt(dayKey.split("-")[0] ?? "0", 10) || 0;
    out.push({
      invoiceCode,
      invoiceDate,
      dayKey,
      day,
      salesPerson: (o["SalesPersonName"] ?? "").trim(),
      employeeCode: (o["EmployeeCode"] ?? "").trim(),
      partType: (o["PartType"] ?? "").trim(),
      subCategory: (o["SubCategory"] ?? "").trim(),
      revenueType: (o["RevenueType"] ?? "").trim(),
      itemCode: (o["ItemCode"] ?? "").trim(),
      itemDescription: (o["ItemDescription"] ?? "").trim(),
      schemeCode: (o["SchemeCode"] ?? "").trim(),
      quantity: num(o["Quantity"] ?? "0"),
      netAmount: num(o["NetAmount"] ?? "0"),
      taxableAmount: num(o["TaxableAmount"] ?? "0"),
      grossProfit: num(o["GrossProfit"] ?? "0"),
    });
  }
  return out;
}

// ---------------------------------------------------------------------------
// Canonical names & helpers
// ---------------------------------------------------------------------------

/** Uppercase key used to merge employee spellings across sheets. */
export function normalizePerson(name: string): string {
  return (name ?? "").trim().toUpperCase().replace(/\s+/g, " ");
}

/** Friendly display name for a canonical person key. */
export function displayPerson(key: string): string {
  const k = key.toUpperCase();
  if (k === "SUBHA") return "SUBHA";
  if (k === "DINESH") return "DINESH";
  return key;
}

export type RevenueTypeKey = "service" | "accessory" | "repair" | "total" | "grossProfit";

export function revenueTypeKey(rt: string): Exclude<RevenueTypeKey, "total" | "grossProfit"> | null {
  const t = (rt ?? "").trim().toLowerCase();
  if (t.startsWith("service")) return "service";
  if (t.startsWith("accessor")) return "accessory";
  if (t.startsWith("repair")) return "repair";
  return null;
}

// ---------------------------------------------------------------------------
// Target sheet model (grid-anchored, A1 notation)
// ---------------------------------------------------------------------------

export interface SalespersonTarget {
  service: number;
  accessory: number;
  repair: number;
  total: number;
  grossProfit: number;
}

/**
 * Upper bound used for open-ended slab ranges. Convex rejects non-finite
 * numbers in function return values, so "and above" is stored as a sentinel
 * that is far beyond any realisable achievement percentage.
 */
export const OPEN_ENDED = 9999;

export interface IncentiveSlab {
  label: string; // "70% to 80%"
  value: number;
  /** Lower edge of the achievement % this slab pays for. */
  minPct: number;
  /** Upper edge of the achievement % this slab pays for. */
  maxPct: number;
}

export interface StoreTargets {
  service: number;
  accessory: number;
  repair: number;
  total: number;
  grossProfit: number;
}

export interface ServiceIncentiveSlabs {
  service: IncentiveSlab[];
  accessory: IncentiveSlab[];
}

export interface AmcSlab {
  label: string; // "4-6"
  rateLabel: string; // "200/AMC"
  rate: number; // 200
  fixed: number; // 0 (or 750 etc. for accessory column)
  min: number; // 4
  max: number; // 6 (OPEN_ENDED for the open-ended slab)
  incentive: number; // expected payout if landed in this slab
  /** Accessory-AMC variant (columns J:K) */
  accCount?: number;
  accRate?: number;
  accFixed?: number;
}

export interface BundleSlab {
  label: string; // "slab level 3"
  min: number;
  perBundle: number;
}

export interface KpiTarget {
  category: "AR" | "RR" | "SR";
  code: string; // CAREPLUS / PROTECTPLUS / AMC / BUNDLE / DISPLAY / BATTERY / GLASS / OTHERS / SRD / SHC / ARD
  label: string;
  target: number;
  value: number;
  approxAmount: number;
}

export interface GoogleReview {
  previousMonth: number;
  currentMonth: number;
  difference: number;
}

export interface TargetSheet {
  store: StoreTargets;
  salespeople: Record<string, SalespersonTarget>; // canonical key -> target
  serviceIncentive: ServiceIncentiveSlabs;
  amcSlabs: AmcSlab[]; // service AMC slabs (G18:H23) + accessory fixed (J:K)
  bundleSlabs: BundleSlab[];
  kpis: KpiTarget[];
  googleReview: GoogleReview;
}

/** Parse "104,500₹" / "₹ 500" / "125000" into a number. */
function parseMoney(s: string | undefined): number {
  if (!s) return 0;
  const v = Number.parseFloat(s.replace(/[₹,\s]/g, ""));
  return Number.isFinite(v) ? v : 0;
}

function parseRangeLabel(label: string): { min: number; max: number } {
  const t = (label ?? "").trim();
  // e.g. "4-6", "7-11", "12-15"
  const m = t.match(/^(\d+)\s*[-–]\s*(\d+)$/);
  if (m) return { min: Number(m[1]), max: Number(m[2]) };
  const single = t.match(/^(\d+)\+$/);
  if (single) return { min: Number(single[1]), max: Number.POSITIVE_INFINITY };
  const only = t.match(/^(\d+)$/);
  if (only) return { min: Number(only[1]), max: Number(only[1]) };
  return { min: 1, max: Number.POSITIVE_INFINITY };
}

/** "200/AMC + 750Fixed" -> { rate: 200, fixed: 750 } ; "200/AMC" -> { rate: 200, fixed: 0 } */
function parseAmcRate(label: string): { rate: number; fixed: number } {
  const t = (label ?? "").trim();
  const rateM = t.match(/(\d+(?:\.\d+)?)\s*\/\s*AMC/i);
  const fixedM = t.match(/\+\s*(\d+(?:\.\d+)?)\s*(?:Fixed)?/i);
  return {
    rate: rateM ? Number(rateM[1]) : 0,
    fixed: fixedM ? Number(fixedM[1]) : 0,
  };
}

function parsePercentLabel(label: string): { min: number; max: number } {
  const t = (label ?? "").replace(/%/g, "").trim();
  if (t.startsWith("<=")) return { min: 0, max: Number(t.slice(2)) };
  const m = t.match(/^(\d+)\s*(?:to|-|–)\s*(\d+)$/);
  if (m) return { min: Number(m[1]), max: Number(m[2]) };
  if (/^\d+$/.test(t)) return { min: Number(t), max: Number(t) };
  return { min: 0, max: 0 };
}

/**
 * Parse the % incentive slab rows. The trailing "<110%" row follows an
 * ascending ladder (…, "100% to 109%") so it is re-interpreted as
 * "above the previous slab" rather than "below 110".
 */
function parsePctSlabs(cells: Array<[string, string]>): IncentiveSlab[] {
  const slabs: IncentiveSlab[] = [];
  let prevMax = 0;
  for (const [label, rawVal] of cells) {
    if (!label) continue;
    const value = parseMoney(rawVal);
    let range = parsePercentLabel(label);
    if (/^<\d/.test(label.trim()) && prevMax > 0) {
      range = { min: prevMax + 1, max: Number.POSITIVE_INFINITY };
    }
    slabs.push({
      label,
      value,
      minPct: range.min,
      maxPct: Math.min(range.max, OPEN_ENDED),
    });
    prevMax = Math.max(prevMax, range.max === Number.POSITIVE_INFINITY ? prevMax : range.max);
  }
  return slabs;
}

function slabForPct(slabs: IncentiveSlab[], pct: number): IncentiveSlab | null {
  for (const s of slabs) {
    if (pct >= s.minPct && pct <= s.maxPct) return s;
  }
  return null;
}

export function incentiveForPct(slabs: IncentiveSlab[], pct: number): number {
  return slabForPct(slabs, pct)?.value ?? 0;
}

/**
 * The Google-review counters sit in G1:I3:
 *   G1 "google review" · G2 "previous month" · G3 previous count
 *   H3 current-month count · I3 difference
 * They are located by scanning the top block for the first row holding two or
 * more numbers, so the real sheet values are always used (no hardcoded
 * placeholders) even if the columns shift.
 */
function parseGoogleReview(rows: string[][]): GoogleReview {
  const numbersIn = (row: string[] | undefined): number[] => {
    const out: number[] = [];
    for (let c = 5; c <= 12; c++) {
      const raw = (row?.[c] ?? "").replace(/[₹\s]/g, "");
      if (!/^-?\d+(?:,\d+)*(?:\.\d+)?%?$/.test(raw)) continue;
      out.push(parseMoney(raw));
    }
    return out;
  };
  const lastRow = Math.min(5, rows.length - 1);
  for (let r = 0; r <= lastRow; r++) {
    const nums = numbersIn(rows[r]);
    if (nums.length < 2) continue;
    const previousMonth = nums[0];
    const currentMonth = nums[1];
    const difference = nums.length >= 3 && nums[2] !== 0
      ? nums[2]
      : Math.max(0, currentMonth - previousMonth);
    return { previousMonth, currentMonth, difference };
  }
  return { previousMonth: 0, currentMonth: 0, difference: 0 };
}

/** First row in [fromCol,toCol] matching `re`, or -1. */
function findRow(rows: string[][], fromCol: number, toCol: number, re: RegExp): number {
  for (let r = 0; r < rows.length; r++) {
    for (let c = fromCol; c <= toCol; c++) {
      if (re.test((rows[r]?.[c] ?? "").trim())) return r;
    }
  }
  return -1;
}

/** First matching cell (row + column) in a column window. */
function findCell(
  rows: string[][],
  fromCol: number,
  toCol: number,
  re: RegExp,
): { row: number; col: number } | null {
  for (let r = 0; r < rows.length; r++) {
    for (let c = fromCol; c <= toCol; c++) {
      if (re.test((rows[r]?.[c] ?? "").trim())) return { row: r, col: c };
    }
  }
  return null;
}

/**
 * Grid-anchored target sheet parser (A1 style anchors from the source doc):
 *  - store target  A1:B5 (rows: Service, Accessories, Repair, Total, Gross Profit)
 *  - salesperson targets D1:E6 (names in D1:E1, labels in C, values in D:E)
 *  - google review G1:I3 (G3 prev, H3 current, I3 diff)
 *  - KPI grid A9:M13 (row 9 targets, row 10 values, row 11 approx)
 *  - AMC slabs F:H + J:K (anchored on the "Incentive per AMC" header)
 *  - Bundle slabs M:P (anchored on the "slab level" row)
 */
export function parseTargetSheet(rows: string[][]): TargetSheet {
  const at = (r: number, c: number): string => (rows[r]?.[c] ?? "").trim(); // 0-indexed

  // --- store target (A1:B5 -> rows 0..4, cols 0..1) ---
  const store: StoreTargets = {
    service: parseMoney(at(1, 1)),
    accessory: parseMoney(at(2, 1)),
    repair: parseMoney(at(3, 1)),
    total: parseMoney(at(4, 1)),
    grossProfit: parseMoney(at(5, 1)),
  };

  // --- salesperson targets (D1:E6 header + rows) ---
  // Row 0: "store target","","","subha","dinesh"
  // Rows 1..5: "<type>","<store val>","<type>","subha val","dinesh val"
  const people: Record<string, SalespersonTarget> = {};
  const personKeys: string[] = [];
  const headerCells = rows[0] ?? [];
  for (let c = 3; c < Math.min(headerCells.length, 5); c++) {
    const name = headerCells[c];
    if (name && !personKeys.includes(normalizePerson(name))) {
      personKeys.push(normalizePerson(name));
    }
  }
  for (const key of personKeys) {
    people[key] = { service: 0, accessory: 0, repair: 0, total: 0, grossProfit: 0 };
  }
  const colIndex: Record<string, number> = {};
  personKeys.forEach((key, i) => {
    colIndex[key] = 3 + i;
  });
  for (let r = 1; r <= 5; r++) {
    const label = at(r, 2).toLowerCase();
    let slot: keyof SalespersonTarget | null = null;
    if (label.startsWith("service")) slot = "service";
    else if (label.startsWith("accessor")) slot = "accessory";
    else if (label.startsWith("repair")) slot = "repair";
    else if (label.startsWith("total")) slot = "total";
    else if (label.startsWith("gross")) slot = "grossProfit";
    if (!slot) continue;
    for (const key of personKeys) {
      people[key][slot] = parseMoney(at(r, colIndex[key]));
    }
  }

  // --- google review (G1:I3) ---
  const googleReview: GoogleReview = parseGoogleReview(rows);

  // --- KPI grid (A8:M12) ---
  // Row 7 (A8): "KPI","","protect plus",...,"RR","","DISPLAY",...
  // Row 8 (A9): targets
  // Row 9 (A10): per-unit values
  // Row 10 (A11): approx amounts
  const kpiCategory = (c: number): "AR" | "RR" | "SR" => (c <= 4 ? "AR" : c <= 10 ? "RR" : "SR");
  void kpiCategory;
  const kpiLabels: Record<string, string> = {
    CAREPLUS: "Care+",
    PROTECTPLUS: "Protect+",
    AMC: "AMC",
    BUNDLE: "Bundle",
    DISPLAY: "Display",
    BATTERY: "Battery",
    GLASS: "Glass",
    OTHERS: "Others",
    SRD: "SRD",
    SHC: "SHC",
    ARD: "ARD",
  };
  // --- KPI grid (A8:M12) ---
  // The sheet's category labels are visually offset from their columns, so the
  // 11 KPIs are mapped positionally (verified against the live sheet):
  //   col1 Care+ 5/20900 · col2 Protect+ 10/8500 · col3 AMC 20/4000 ·
  //   col4 Bundle 20/3000 · col6 Display 5/25000 · col7 Battery 40/8600 ·
  //   col8 Glass 3/18500 · col9 Others 10/9000 · col10 SRD 50/2500 ·
  //   col11 SHC 30/1000 · col12 ARD 30/2500
  const KPI_LAYOUT: Array<[string, number]> = [
    ["CAREPLUS", 1],
    ["PROTECTPLUS", 2],
    ["AMC", 3],
    ["BUNDLE", 4],
    ["DISPLAY", 6],
    ["BATTERY", 7],
    ["GLASS", 8],
    ["OTHERS", 9],
    ["SRD", 10],
    ["SHC", 11],
    ["ARD", 12],
  ];
  const kpis: KpiTarget[] = [];
  for (const [code, col] of KPI_LAYOUT) {
    const targetVal = parseMoney(at(8, col));
    if (targetVal <= 0) continue;
    kpis.push({
      category: col <= 4 ? "AR" : col <= 9 ? "RR" : "SR",
      code,
      label: kpiLabels[code] ?? code,
      target: targetVal,
      value: parseMoney(at(9, col)),
      approxAmount: parseMoney(at(10, col)),
    });
  }

  // --- incentive slabs (service % A/B, accessory % C/D) ---
  // Anchored on the "achivement" header row so rows inserted in the source
  // sheet cannot silently shift the mapping.
  const pctAnchor = findRow(rows, 0, 3, /^(achivement|achievement)$/i);
  const svcCells: Array<[string, string]> = [];
  const accCells: Array<[string, string]> = [];
  if (pctAnchor >= 0) {
    const limit = Math.min(rows.length - 1, pctAnchor + 12);
    for (let r = pctAnchor + 1; r <= limit; r++) {
      const sLabel = at(r, 0);
      const sVal = at(r, 1);
      const aLabel = at(r, 2);
      const aVal = at(r, 3);
      if (!sLabel && !aLabel) break;
      if (/%|nil/i.test(sLabel + sVal)) svcCells.push([sLabel, sVal]);
      if (/%|nil/i.test(aLabel + aVal)) accCells.push([aLabel, aVal]);
    }
  }
  const serviceSlabs = parsePctSlabs(svcCells);
  const accessorySlabs = parsePctSlabs(accCells);

  // --- AMC slabs: F range · G rate per AMC · I accessory-AMC count ·
  // J accessory rate (+fixed) · K accessory payout ---
  const amcAnchor = findRow(rows, 4, 10, /incentive\s*per\s*amc/i);
  const amcSlabs: AmcSlab[] = [];
  if (amcAnchor >= 0) {
    const limit = Math.min(rows.length - 1, amcAnchor + 12);
    for (let r = amcAnchor + 1; r <= limit; r++) {
      const rangeLabel = at(r, 5);
      const rateLabel = at(r, 6);
      const accCount = parseMoney(at(r, 8));
      const accRateLabel = at(r, 9);
      if (!rangeLabel && !rateLabel && !accRateLabel) break;
      const { rate, fixed } = parseAmcRate(rateLabel);
      let min: number;
      let max: number;
      if (rangeLabel && /^\d/.test(rangeLabel)) {
        const range = parseRangeLabel(rangeLabel);
        min = range.min;
        max = range.max;
      } else if (amcSlabs.length > 0) {
        // Rate-only trailing row (e.g. "500/AMC") continues the ladder.
        min = amcSlabs[amcSlabs.length - 1].max + 1;
        max = Number.POSITIVE_INFINITY;
      } else {
        continue;
      }
      const accRateInfo = parseAmcRate(accRateLabel);
      const slab: AmcSlab = {
        label: rangeLabel || `${min}+`,
        rateLabel,
        rate,
        fixed,
        min: Math.min(min, OPEN_ENDED),
        max: Math.min(Math.max(max, min), OPEN_ENDED),
        incentive: Math.max(rate * min + fixed, 0),
      };
      // Only attach the accessory-AMC variant when the sheet really has one —
      // Convex return validators reject explicit undefined values.
      if (accCount > 0) slab.accCount = accCount;
      if (accRateInfo.rate > 0) slab.accRate = accRateInfo.rate;
      if (accRateInfo.fixed > 0) slab.accFixed = accRateInfo.fixed;
      amcSlabs.push(slab);
    }
  }

  // --- Bundle slabs: "slab level" row + "slab incentive" row below it ---
  const bundleSlabs: BundleSlab[] = [];
  const levelCell = findCell(rows, 12, 17, /slab\s*level/i);
  if (levelCell) {
    for (let c = levelCell.col + 1; c <= levelCell.col + 3; c++) {
      const level = parseMoney(at(levelCell.row, c));
      const per = parseMoney(at(levelCell.row + 1, c));
      if (level > 0 && per > 0)
        bundleSlabs.push({ label: `slab level ${level}`, min: level, perBundle: per });
    }
  }
  bundleSlabs.sort((a, b) => a.min - b.min);

  return {
    store,
    salespeople: people,
    serviceIncentive: { service: serviceSlabs, accessory: accessorySlabs },
    amcSlabs,
    bundleSlabs,
    kpis,
    googleReview,
  };
}
