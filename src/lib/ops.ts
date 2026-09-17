// Data layer for the operational dashboards (PROCESS + PRODUCTIVITY).
//
// Sources, all read from the same public workbook as the business page:
//   walkin · walk-in tokens (customer flow, case + invoice creation)
//   aod    · service order detail (case status, technician, TAT)
//   comm   · customer communication log (per-case update history)
//   crm    · invoice lines, used only for invoice revenue attribution
//
// Column lookups are header-driven (not positional) so the dashboards keep
// working if the sheet columns move.

import { SHEET_ID } from "./sheets";

/**
 * Column projections. The full AOD export is ~6.5 MB, which is too heavy for a
 * serverless isolate, so each sheet is fetched with only the columns these
 * dashboards actually read (~1.3 MB in total).
 */
const OPS_COLUMNS = {
  walkin: "N,G,Q,R,U",
  aod: "I,J,K,L,N,W,X,Y,AF,AK,AL,BB,BH,BL,BO,BQ,AY,CE",
  comm: "G,I",
};

export function buildOpsFetchUrls() {
  const url = (sheet: keyof typeof OPS_COLUMNS) =>
    `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:csv&headers=1&sheet=${sheet}` +
    `&tq=${encodeURIComponent(`select ${OPS_COLUMNS[sheet]}`)}`;
  return {
    walkinUrl: url("walkin"),
    aodUrl: url("aod"),
    commUrl: url("comm"),
  };
}

/**
 * Allocation-light RFC-4180 CSV reader for the large operational exports.
 * Fields are materialised with slice() instead of character concatenation, so
 * a 1.3 MB sheet stays comfortably inside a small isolate's memory budget.
 */
export function parseCsvLite(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  const len = text.length;
  let i = 0;

  while (i < len) {
    let field: string;
    if (text.charCodeAt(i) === 34) {
      i += 1;
      let buffer = "";
      let segmentStart = i;
      while (i < len) {
        const code = text.charCodeAt(i);
        if (code === 34) {
          if (text.charCodeAt(i + 1) === 34) {
            buffer += text.slice(segmentStart, i + 1);
            i += 2;
            segmentStart = i;
            continue;
          }
          buffer += text.slice(segmentStart, i);
          i += 1;
          break;
        }
        i += 1;
      }
      field = buffer;
      while (i < len) {
        const code = text.charCodeAt(i);
        if (code === 44 || code === 10) break;
        i += 1;
      }
    } else {
      const start = i;
      while (i < len) {
        const code = text.charCodeAt(i);
        if (code === 44 || code === 10) break;
        i += 1;
      }
      field = text.slice(start, i);
    }

    row.push(field.charCodeAt(field.length - 1) === 13 ? field.slice(0, -1) : field);

    if (i < len) {
      if (text.charCodeAt(i) === 44) {
        i += 1;
        continue;
      }
      i += 1;
      rows.push(row);
      row = [];
    }
  }
  if (row.length) rows.push(row);
  return rows;
}

/* ------------------------------------------------------------------ */
/* Small helpers                                                       */
/* ------------------------------------------------------------------ */

/** The sheets store local (IST) wall-clock strings without an offset. */
const IST_OFFSET_MS = 330 * 60 * 1000;

const MONTH_NAMES = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December",
];

export const str = (v: unknown): string => String(v ?? "").trim();

const num = (v: unknown): number => {
  const n = Number.parseFloat(str(v).replace(/[,\s]/g, ""));
  return Number.isFinite(n) ? n : 0;
};

/** Read a cell by a header index that may not exist in the sheet. */
const at = (row: string[], index: number | null): string =>
  index === null ? "" : str(row[index]);

function columnIndex(header: string[]): (name: string) => number | null {
  const map = new Map<string, number>();
  header.forEach((h, i) => {
    const key = str(h).toLowerCase();
    if (key && !map.has(key)) map.set(key, i);
  });
  return (name) => {
    const i = map.get(name.toLowerCase());
    return i === undefined ? null : i;
  };
}

/**
 * Parse both timestamp shapes present in the workbook into epoch ms.
 *   "2026-09-12 12:52:00" (walkin / aod)  ·  "12-09-2026" (comm)
 * Everything is treated as a single fixed zone, so all durations stay exact.
 */
export function parseSheetTs(raw: string): number | null {
  const t = str(raw);
  if (!t) return null;
  const iso = t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (iso) {
    return Date.UTC(
      Number(iso[1]),
      Number(iso[2]) - 1,
      Number(iso[3]),
      Number(iso[4] ?? 0),
      Number(iso[5] ?? 0),
      Number(iso[6] ?? 0),
    );
  }
  const dmy = t.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (dmy) {
    return Date.UTC(
      Number(dmy[3]),
      Number(dmy[2]) - 1,
      Number(dmy[1]),
      Number(dmy[4] ?? 0),
      Number(dmy[5] ?? 0),
      Number(dmy[6] ?? 0),
    );
  }
  const fallback = Date.parse(t);
  return Number.isFinite(fallback) ? fallback : null;
}

const dayKeyOf = (ts: number): string => new Date(ts).toISOString().slice(0, 10);
const monthKeyOf = (ts: number): string => new Date(ts).toISOString().slice(0, 7);

function monthLabel(key: string): string {
  const [y, m] = key.split("-");
  const name = MONTH_NAMES[Number(m) - 1] ?? key;
  return `${name} ${y}`;
}

function shortMonthLabel(key: string): string {
  const [y, m] = key.split("-");
  return `${(MONTH_NAMES[Number(m) - 1] ?? key).slice(0, 3)} ${y}`;
}

function shiftMonth(key: string, delta: number): string {
  const [y, m] = key.split("-").map(Number);
  const d = new Date(Date.UTC(y, (m ?? 1) - 1 + delta, 1));
  return d.toISOString().slice(0, 7);
}

const round1 = (n: number): number => Math.round(n * 10) / 10;

/** "3d 4h 12m" · "4h 12m" · "12m" · "—" */
export function formatMinutes(minutes: number | null): string {
  if (minutes === null || !Number.isFinite(minutes)) return "—";
  const total = Math.max(0, Math.round(minutes));
  const days = Math.floor(total / 1440);
  const hours = Math.floor((total % 1440) / 60);
  const mins = total % 60;
  if (days > 0) return `${days}d ${hours}h ${mins}m`;
  if (hours > 0) return `${hours}h ${mins}m`;
  return `${mins}m`;
}

/** "Dinesh Kumar" from "dinesh.kumar@consolidated.one". */
function personLabel(raw: string): string {
  const t = str(raw);
  if (!t) return "UNASSIGNED";
  const local = t.includes("@") ? t.split("@")[0] : t;
  return local
    .split(/[._\-]+/)
    .filter(Boolean)
    .map((w) => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase())
    .join(" ");
}

const digitsOnly = (v: string): string => str(v).replace(/\D/g, "");

function average(values: number[]): number {
  if (!values.length) return 0;
  return values.reduce((a, b) => a + b, 0) / values.length;
}

/* ------------------------------------------------------------------ */
/* Source parsers                                                      */
/* ------------------------------------------------------------------ */

export interface WalkinRow {
  mobile: string;
  createdBy: string;
  caseId: string;
  invoiceCode: string;
  ts: number | null;
}

export function parseWalkinRows(rows: string[][]): WalkinRow[] {
  if (rows.length < 2) return [];
  const col = columnIndex(rows[0]);
  const cMobile = col("MobileNo");
  const cDate = col("TokenDate");
  const cCreated = col("CreatedBy");
  const cCase = col("CaseId");
  const cInvoice = col("InvoiceCode");
  const out: WalkinRow[] = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const mobile = at(row, cMobile);
    if (!mobile) continue;
    out.push({
      mobile,
      createdBy: at(row, cCreated) || "UNASSIGNED",
      caseId: at(row, cCase),
      invoiceCode: at(row, cInvoice),
      ts: parseSheetTs(at(row, cDate)),
    });
  }
  return out;
}

export interface AodRow {
  caseId: string;
  mobile: string;
  customer: string;
  product: string;
  coverage: string;
  repairType: string;
  bucket: string;
  jobStatus: string;
  technician: string;
  challanStatus: string;
  jobAge: number | null;
  caseTs: number | null;
  assignTs: number | null;
  completeTs: number | null;
  /** Which sheet field supplied the completion date. */
  completeSource: string;
}

/** CIN1 → Carry-In · SNR → N-Repair · anything else passed through. */
function normaliseServiceType(raw: string): string {
  const t = str(raw).toUpperCase();
  if (t === "CIN1") return "Carry-In";
  if (t === "SNR") return "N-Repair";
  return str(raw);
}

export function parseAodRows(rows: string[][]): AodRow[] {
  if (rows.length < 2) return [];
  const col = columnIndex(rows[0]);
  const c = {
    caseId: col("CaseId"),
    mobile: col("MobileNo"),
    first: col("FirstName"),
    last: col("LastName"),
    product: col("productDescription"),
    coverage: col("GSXWarrantyStatusDesc"),
    serviceType: col("ServiceType"),
    repairDoc: col("RepairDocType"),
    jobStatus: col("JobStatus"),
    bucket: col("CRM_JobStatusDesc"),
    technician: col("TechCode"),
    challan: col("DeliveryChallanStatus"),
    jobAge: col("JobAge"),
    caseDate: col("CaseDate"),
    assignDate: col("AssignDate"),
    rfp: col("GSX_RFPUDateTime"),
    ready: col("Ready_For_Pickup_Date_CRM"),
    handover: col("HandOverDate"),
  };

  const out: AodRow[] = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const caseId = at(row, c.caseId);
    if (!caseId) continue;
    const customer = [at(row, c.first), at(row, c.last)].filter(Boolean).join(" ");
    const repairTypeLabel = at(row, c.repairDoc) || normaliseServiceType(at(row, c.serviceType));
    const assignTs = parseSheetTs(at(row, c.assignDate));

    // Completion ladder: RFPU → CRM ready-for-pickup → hand-over.
    let completeTs: number | null = null;
    let completeSource = "";
    const sources: Array<[string, string]> = [
      ["RFPU", at(row, c.rfp)],
      ["Ready For Pickup", at(row, c.ready)],
      ["Hand-over", at(row, c.handover)],
    ];
    for (const [label, raw] of sources) {
      const ts = parseSheetTs(raw);
      if (ts !== null) {
        completeTs = ts;
        completeSource = label;
        break;
      }
    }

    const age = at(row, c.jobAge);
    out.push({
      caseId,
      mobile: at(row, c.mobile),
      customer,
      product: at(row, c.product),
      coverage: at(row, c.coverage),
      repairType: repairTypeLabel || "Unclassified",
      bucket: at(row, c.bucket),
      jobStatus: at(row, c.jobStatus),
      technician: personLabel(at(row, c.technician)),
      challanStatus: at(row, c.challan),
      jobAge: age ? num(age) : null,
      caseTs: parseSheetTs(at(row, c.caseDate)),
      assignTs,
      completeTs,
      completeSource,
    });
  }
  return out;
}

export interface CommRow {
  caseId: string;
  ts: number | null;
}

export function parseCommRows(rows: string[][]): CommRow[] {
  if (rows.length < 2) return [];
  const col = columnIndex(rows[0]);
  const cCase = col("CaseId");
  const cDate = col("CreatedDate");
  const out: CommRow[] = [];
  for (let i = 1; i < rows.length; i++) {
    const caseId = at(rows[i], cCase);
    if (!caseId) continue;
    out.push({ caseId, ts: parseSheetTs(at(rows[i], cDate)) });
  }
  return out;
}

/* ------------------------------------------------------------------ */
/* Result shapes                                                       */
/* ------------------------------------------------------------------ */

export const AGEING_SLABS = ["0-2", "3-5", "6-10", "11-15", "16-20", "20+"];

export interface RepairTypeSummary {
  repairType: string;
  current: number;
  previousThree: number[];
  previousAverage: number;
  difference: number;
  percentVsAverage: number;
}

export interface RepairPerformance {
  available: boolean;
  currentMonth: string;
  previousMonths: string[];
  totalCurrent: number;
  previousThreeAverage: number;
  difference: number;
  percentVsAverage: number;
  typeSummary: RepairTypeSummary[];
  detail: Array<{
    repairType: string;
    coverage: string;
    product: string;
    current: number;
    previousAverage: number;
    difference: number;
    percentVsAverage: number;
  }>;
}

export interface OpenJob {
  caseId: string;
  customer: string;
  mobile: string;
  product: string;
  coverage: string;
  repairType: string;
  caseDate: string;
  caseAge: string;
  bucket: string;
  unclaimed: boolean;
  lastCommunication: string;
  communicationAging: string;
}

export interface OpenJobAgeing {
  ages: string[];
  rows: Array<{ bucket: string; cells: number[]; total: number }>;
  totals: number[];
  grandTotal: number;
}

export interface ProcessDashboard {
  available: boolean;
  month: string;
  monthKey: string;
  summary: {
    total: number;
    unique: number;
    jobs: number;
    invoices: number;
    communications: number;
    openJobs: number;
    unclaimed: number;
  };
  daily: Array<{
    date: string;
    dateKey: string;
    total: number;
    cases: number;
    invoices: number;
    communications: number;
  }>;
  repair: RepairPerformance;
  openJobs: OpenJob[];
  ageing: OpenJobAgeing;
}

export interface ProductivityDashboard {
  available: boolean;
  month: string;
  monthKey: string;
  summary: {
    walkins: number;
    uniqueWalkins: number;
    jobs: number;
    invoices: number;
    revenue: number;
    averageWalkinsPerDay: number;
    averageJobsPerDay: number;
    averageInvoicesPerDay: number;
    walkinToJob: number;
    walkinToInvoice: number;
    handlingMinutes: number;
    handlingSample: number;
    repairTatMinutes: number;
    nRepairTatMinutes: number;
    totalRepairs: number;
    carryInRepairs: number;
    nRepairRepairs: number;
    tatSample: number;
  };
  daily: Array<{
    date: string;
    dateKey: string;
    walkins: number;
    jobs: number;
    invoices: number;
    jobConversion: number;
    invoiceConversion: number;
  }>;
  people: Array<{
    name: string;
    walkins: number;
    uniqueWalkins: number;
    jobs: number;
    invoices: number;
    walkinToJob: number;
    walkinToInvoice: number;
    invoiceRevenue: number;
  }>;
  technicians: Array<{
    name: string;
    total: number;
    carryIn: number;
    nRepair: number;
    carryInTat: string;
    nRepairTat: string;
    averageTat: string;
  }>;
  tatCases: Array<{
    caseId: string;
    technician: string;
    repairType: string;
    customer: string;
    mobile: string;
    assignDate: string;
    completeDate: string;
    completeSource: string;
    tat: string;
    tatMinutes: number;
  }>;
  tatCaseTotal: number;
}

export interface OpsResult {
  process: ProcessDashboard;
  productivity: ProductivityDashboard;
}

export interface OpsInput {
  walkin: WalkinRow[];
  aod: AodRow[];
  comm: CommRow[];
  /** invoice code → billed value, from the current-month CRM lines. */
  invoiceRevenue: Map<string, number>;
  /** "now" expressed as IST wall-clock epoch ms. */
  nowTs?: number;
}

/* ------------------------------------------------------------------ */
/* Shared derivations                                                  */
/* ------------------------------------------------------------------ */

/** One row per case id — the last row in sheet order wins. */
function latestPerCase(aod: AodRow[]): AodRow[] {
  const map = new Map<string, AodRow>();
  for (const row of aod) map.set(row.caseId, row);
  return Array.from(map.values());
}

const isClosed = (r: AodRow): boolean => r.jobStatus.toLowerCase() === "closed";
const isOpen = (r: AodRow): boolean => r.jobStatus.toLowerCase() === "open";

function latestCommunication(comm: CommRow[]): Map<string, number> {
  const map = new Map<string, number>();
  for (const row of comm) {
    if (row.ts === null) continue;
    const current = map.get(row.caseId);
    if (current === undefined || row.ts > current) map.set(row.caseId, row.ts);
  }
  return map;
}

function ageingSlabIndex(age: number): number {
  if (age <= 2) return 0;
  if (age <= 5) return 1;
  if (age <= 10) return 2;
  if (age <= 15) return 3;
  if (age <= 20) return 4;
  return 5;
}

/* ------------------------------------------------------------------ */
/* Repair type performance (current month vs previous 3-month average) */
/* ------------------------------------------------------------------ */

function computeRepairPerformance(aod: AodRow[], monthKey: string): RepairPerformance {
  const previousMonths = [
    shiftMonth(monthKey, -3),
    shiftMonth(monthKey, -2),
    shiftMonth(monthKey, -1),
  ];
  const completed = latestPerCase(aod).filter((r) => isClosed(r) && r.completeTs !== null);
  const current = completed.filter((r) => monthKeyOf(r.completeTs as number) === monthKey);
  const previousPool = completed.filter((r) =>
    previousMonths.includes(monthKeyOf(r.completeTs as number)),
  );

  const countFor = (pool: AodRow[], type: string, month?: string): number =>
    pool.filter(
      (r) =>
        r.repairType === type &&
        (month === undefined || monthKeyOf(r.completeTs as number) === month),
    ).length;

  const types = Array.from(
    new Set([...current.map((r) => r.repairType), ...previousPool.map((r) => r.repairType)]),
  );

  const typeSummary: RepairTypeSummary[] = types
    .map((repairType) => {
      const currentCount = countFor(current, repairType);
      const previousThree = previousMonths.map((m) => countFor(previousPool, repairType, m));
      const previousAverage = round1(previousThree.reduce((a, b) => a + b, 0) / previousMonths.length);
      const difference = round1(currentCount - previousAverage);
      return {
        repairType,
        current: currentCount,
        previousThree,
        previousAverage,
        difference,
        percentVsAverage:
          previousAverage > 0 ? round1((currentCount / previousAverage) * 100) : currentCount > 0 ? 100 : 0,
      };
    })
    .sort((a, b) => b.current - a.current || b.previousAverage - a.previousAverage);

  // Detail: repair type × coverage × product.
  const groupKey = (r: AodRow) => `${r.repairType}|${r.coverage || "—"}|${r.product || "—"}`;
  const detailMap = new Map<string, { repairType: string; coverage: string; product: string; current: number }>();
  for (const r of current) {
    const key = groupKey(r);
    const entry = detailMap.get(key) ?? {
      repairType: r.repairType,
      coverage: r.coverage || "—",
      product: r.product || "—",
      current: 0,
    };
    entry.current += 1;
    detailMap.set(key, entry);
  }
  const previousByGroup = new Map<string, number>();
  for (const r of previousPool) {
    const key = groupKey(r);
    previousByGroup.set(key, (previousByGroup.get(key) ?? 0) + 1);
  }
  for (const [key, entry] of detailMap) {
    if (previousByGroup.has(key)) continue;
    const [repairType, coverage, product] = key.split("|");
    detailMap.set(key, { repairType, coverage, product, current: 0 });
  }

  const detail = Array.from(detailMap.entries())
    .map(([key, entry]) => {
      const previousAverage = round1((previousByGroup.get(key) ?? 0) / previousMonths.length);
      return {
        ...entry,
        previousAverage,
        difference: round1(entry.current - previousAverage),
        percentVsAverage:
          previousAverage > 0
            ? round1((entry.current / previousAverage) * 100)
            : entry.current > 0
              ? 100
              : 0,
      };
    })
    .sort((a, b) => b.current - a.current || b.previousAverage - a.previousAverage)
    .slice(0, 40);

  const totalCurrent = current.length;
  const previousThreeAverage = round1(previousPool.length / previousMonths.length);

  return {
    available: totalCurrent > 0 || previousPool.length > 0,
    currentMonth: monthLabel(monthKey),
    previousMonths: previousMonths.map(shortMonthLabel),
    totalCurrent,
    previousThreeAverage,
    difference: round1(totalCurrent - previousThreeAverage),
    percentVsAverage:
      previousThreeAverage > 0
        ? round1((totalCurrent / previousThreeAverage) * 100)
        : totalCurrent > 0
          ? 100
          : 0,
    typeSummary,
    detail,
  };
}

/* ------------------------------------------------------------------ */
/* Process dashboard                                                   */
/* ------------------------------------------------------------------ */

function computeProcessDashboard(input: OpsInput, monthKey: string, nowTs: number): ProcessDashboard {
  const { walkin, aod, comm } = input;
  const rows = walkin.filter((r) => r.ts !== null && monthKeyOf(r.ts) === monthKey);

  const dailyMap = new Map<
    string,
    { total: number; cases: Set<string>; invoices: Set<string> }
  >();
  const mobiles = new Set<string>();
  const cases = new Set<string>();
  const invoices = new Set<string>();

  for (const row of rows) {
    const key = dayKeyOf(row.ts as number);
    const day = dailyMap.get(key) ?? { total: 0, cases: new Set<string>(), invoices: new Set<string>() };
    day.total += 1;
    mobiles.add(digitsOnly(row.mobile) || row.mobile);
    if (row.caseId) {
      day.cases.add(row.caseId);
      cases.add(row.caseId);
    }
    if (row.invoiceCode) {
      day.invoices.add(row.invoiceCode);
      invoices.add(row.invoiceCode);
    }
    dailyMap.set(key, day);
  }

  // Communication updates logged for the month (comm!CreatedDate).
  const commByDay = new Map<string, number>();
  for (const entry of comm) {
    if (entry.ts === null || monthKeyOf(entry.ts) !== monthKey) continue;
    const key = dayKeyOf(entry.ts);
    commByDay.set(key, (commByDay.get(key) ?? 0) + 1);
  }
  let communications = 0;
  for (const [key, count] of commByDay) {
    communications += count;
    if (!dailyMap.has(key)) {
      dailyMap.set(key, { total: 0, cases: new Set<string>(), invoices: new Set<string>() });
    }
  }

  const daily = Array.from(dailyMap.entries())
    .sort((a, b) => a[0].localeCompare(b[0]))
    .map(([dateKey, day]) => ({
      date: dateKey.split("-").reverse().join("-"),
      dateKey,
      total: day.total,
      cases: day.cases.size,
      invoices: day.invoices.size,
      communications: commByDay.get(dateKey) ?? 0,
    }));

  // ----- open jobs -----
  const lastComm = latestCommunication(comm);
  const openCases = latestPerCase(aod).filter(isOpen);
  const openJobs: OpenJob[] = openCases
    .map((r) => {
      const commTs = lastComm.get(r.caseId);
      return {
        caseId: r.caseId,
        customer: r.customer || "—",
        mobile: r.mobile || "—",
        product: r.product || "—",
        coverage: r.coverage || "—",
        repairType: r.repairType || "—",
        caseDate: r.caseTs === null ? "—" : dayKeyOf(r.caseTs).split("-").reverse().join("-"),
        caseAge: r.jobAge === null ? "—" : `${Math.round(r.jobAge)} days`,
        bucket: r.bucket || "Unclassified",
        unclaimed: r.bucket.trim().toLowerCase() === "unclaimed device" && !r.challanStatus,
        lastCommunication:
          commTs === undefined ? "—" : dayKeyOf(commTs).split("-").reverse().join("-"),
        communicationAging:
          commTs === undefined ? "—" : `${Math.max(0, Math.floor((nowTs - commTs) / 86400000))} days`,
      };
    })
    .sort((a, b) => a.caseDate.split("-").reverse().join("").localeCompare(b.caseDate.split("-").reverse().join("")));

  // ----- ageing pivot -----
  const bucketMap = new Map<string, { cells: number[]; total: number }>();
  for (const r of openCases) {
    if (r.jobAge === null) continue;
    const bucket = r.bucket || "Unclassified";
    const entry = bucketMap.get(bucket) ?? { cells: [0, 0, 0, 0, 0, 0], total: 0 };
    entry.cells[ageingSlabIndex(r.jobAge)] += 1;
    entry.total += 1;
    bucketMap.set(bucket, entry);
  }
  const totals = [0, 0, 0, 0, 0, 0];
  const ageingRows = Array.from(bucketMap.entries())
    .map(([bucket, entry]) => {
      entry.cells.forEach((v, i) => {
        totals[i] += v;
      });
      return { bucket, cells: entry.cells, total: entry.total };
    })
    .sort((a, b) => b.total - a.total);
  const grandTotal = ageingRows.reduce((acc, r) => acc + r.total, 0);

  return {
    available: true,
    month: monthLabel(monthKey),
    monthKey,
    summary: {
      total: rows.length,
      unique: mobiles.size,
      jobs: cases.size,
      invoices: invoices.size,
      communications,
      openJobs: openJobs.length,
      unclaimed: openJobs.filter((j) => j.unclaimed).length,
    },
    daily,
    repair: computeRepairPerformance(aod, monthKey),
    openJobs,
    ageing: { ages: AGEING_SLABS, rows: ageingRows, totals, grandTotal },
  };
}

/* ------------------------------------------------------------------ */
/* Productivity dashboard                                              */
/* ------------------------------------------------------------------ */

function computeProductivityDashboard(
  input: OpsInput,
  monthKey: string,
): ProductivityDashboard {
  const { walkin, aod, invoiceRevenue } = input;
  const rows = walkin.filter((r) => r.ts !== null && monthKeyOf(r.ts) === monthKey);

  // mobile → candidate cases, so a walk-in can be matched to its job.
  const byMobile = new Map<string, AodRow[]>();
  for (const r of latestPerCase(aod)) {
    const key = digitsOnly(r.mobile);
    if (!key) continue;
    const list = byMobile.get(key) ?? [];
    list.push(r);
    byMobile.set(key, list);
  }

  interface PersonAcc {
    name: string;
    walkins: number;
    mobiles: Set<string>;
    jobs: Set<string>;
    invoices: Set<string>;
    revenue: number;
    revenueInvoices: Set<string>;
  }
  const peopleMap = new Map<string, PersonAcc>();
  const dailyMap = new Map<string, { walkins: number; jobs: Set<string>; invoices: Set<string> }>();
  const allJobs = new Set<string>();
  const allInvoices = new Set<string>();
  const handlingSamples: number[] = [];

  for (const row of rows) {
    const ts = row.ts as number;
    const key = dayKeyOf(ts);
    const day = dailyMap.get(key) ?? { walkins: 0, jobs: new Set<string>(), invoices: new Set<string>() };
    day.walkins += 1;
    dailyMap.set(key, day);

    const person =
      peopleMap.get(row.createdBy) ??
      ({
        name: personLabel(row.createdBy),
        walkins: 0,
        mobiles: new Set<string>(),
        jobs: new Set<string>(),
        invoices: new Set<string>(),
        revenue: 0,
        revenueInvoices: new Set<string>(),
      } satisfies PersonAcc);
    person.walkins += 1;
    person.mobiles.add(digitsOnly(row.mobile) || row.mobile);

    // Match the walk-in to the nearest service order for the same mobile.
    const candidates = byMobile.get(digitsOnly(row.mobile)) ?? [];
    let matched: AodRow | null = null;
    let bestScore = Number.POSITIVE_INFINITY;
    for (const candidate of candidates) {
      if (candidate.caseTs === null) continue;
      const diff = candidate.caseTs - ts;
      const score = diff >= 0 ? diff : Math.abs(diff) + 864000000;
      if (score < bestScore) {
        bestScore = score;
        matched = candidate;
      }
      if (diff >= 0 && diff < 604800000) break;
    }
    if (matched) {
      person.jobs.add(matched.caseId);
      allJobs.add(matched.caseId);
      day.jobs.add(matched.caseId);
      if (matched.caseTs !== null && handlingSamples.length < 5000) {
        const diff = (matched.caseTs - ts) / 60000;
        if (diff >= 0) handlingSamples.push(diff);
      }
    }

    if (row.invoiceCode && invoiceRevenue.has(row.invoiceCode)) {
      person.invoices.add(row.invoiceCode);
      allInvoices.add(row.invoiceCode);
      day.invoices.add(row.invoiceCode);
      if (!person.revenueInvoices.has(row.invoiceCode)) {
        person.revenueInvoices.add(row.invoiceCode);
        person.revenue += invoiceRevenue.get(row.invoiceCode) ?? 0;
      }
    }

    peopleMap.set(row.createdBy, person);
  }

  const activeDays = Math.max(1, dailyMap.size);

  // ----- technician productivity + case TAT -----
  const completed = latestPerCase(aod).filter(
    (r) => isClosed(r) && r.completeTs !== null && monthKeyOf(r.completeTs as number) === monthKey,
  );

  interface TechAcc {
    name: string;
    total: number;
    carryIn: number;
    nRepair: number;
    carryInTat: number[];
    nRepairTat: number[];
    allTat: number[];
  }
  const techMap = new Map<string, TechAcc>();
  const tatCases: ProductivityDashboard["tatCases"] = [];

  for (const r of completed) {
    const completeTs = r.completeTs as number;
    const minutes = r.assignTs !== null ? Math.max(0, (completeTs - r.assignTs) / 60000) : null;
    const isCarryIn = r.repairType.toLowerCase().includes("carry");
    const tech =
      techMap.get(r.technician) ??
      ({
        name: r.technician,
        total: 0,
        carryIn: 0,
        nRepair: 0,
        carryInTat: [],
        nRepairTat: [],
        allTat: [],
      } satisfies TechAcc);
    tech.total += 1;
    if (isCarryIn) tech.carryIn += 1;
    else tech.nRepair += 1;
    if (minutes !== null) {
      tech.allTat.push(minutes);
      if (isCarryIn) tech.carryInTat.push(minutes);
      else tech.nRepairTat.push(minutes);
      tatCases.push({
        caseId: r.caseId,
        technician: r.technician,
        repairType: r.repairType,
        customer: r.customer || "—",
        mobile: r.mobile || "—",
        assignDate: r.assignTs === null ? "—" : dayKeyOf(r.assignTs).split("-").reverse().join("-"),
        completeDate: dayKeyOf(completeTs).split("-").reverse().join("-"),
        completeSource: r.completeSource,
        tat: formatMinutes(minutes),
        tatMinutes: Math.round(minutes),
      });
    }
    techMap.set(r.technician, tech);
  }

  const technicians = Array.from(techMap.values())
    .map((t) => ({
      name: t.name,
      total: t.total,
      carryIn: t.carryIn,
      nRepair: t.nRepair,
      carryInTat: t.carryInTat.length ? formatMinutes(average(t.carryInTat)) : "—",
      nRepairTat: t.nRepairTat.length ? formatMinutes(average(t.nRepairTat)) : "—",
      averageTat: t.allTat.length ? formatMinutes(average(t.allTat)) : "—",
    }))
    .sort((a, b) => b.total - a.total);

  const carryInTat = completed
    .filter((r) => r.repairType.toLowerCase().includes("carry") && r.assignTs !== null)
    .map((r) => ((r.completeTs as number) - (r.assignTs as number)) / 60000);
  const nRepairTat = completed
    .filter((r) => !r.repairType.toLowerCase().includes("carry") && r.assignTs !== null)
    .map((r) => ((r.completeTs as number) - (r.assignTs as number)) / 60000);
  const people = Array.from(peopleMap.values())
    .map((p) => {
      const uniqueWalkins = p.mobiles.size;
      return {
        name: p.name,
        walkins: p.walkins,
        uniqueWalkins,
        jobs: p.jobs.size,
        invoices: p.invoices.size,
        walkinToJob: uniqueWalkins ? round1((p.jobs.size / uniqueWalkins) * 100) : 0,
        walkinToInvoice: uniqueWalkins ? round1((p.invoices.size / uniqueWalkins) * 100) : 0,
        invoiceRevenue: Math.round(p.revenue),
      };
    })
    .sort((a, b) => b.invoiceRevenue - a.invoiceRevenue || b.walkins - a.walkins);

  const daily = Array.from(dailyMap.entries())
    .sort((a, b) => a[0].localeCompare(b[0]))
    .map(([dateKey, day]) => ({
      date: dateKey.split("-").reverse().join("-"),
      dateKey,
      walkins: day.walkins,
      jobs: day.jobs.size,
      invoices: day.invoices.size,
      jobConversion: day.walkins ? round1((day.jobs.size / day.walkins) * 100) : 0,
      invoiceConversion: day.walkins ? round1((day.invoices.size / day.walkins) * 100) : 0,
    }));

  const uniqueWalkins = new Set(rows.map((r) => digitsOnly(r.mobile) || r.mobile)).size;
  const revenue = Array.from(allInvoices).reduce(
    (acc, code) => acc + (invoiceRevenue.get(code) ?? 0),
    0,
  );

  tatCases.sort((a, b) => b.tatMinutes - a.tatMinutes);

  return {
    available: true,
    month: monthLabel(monthKey),
    monthKey,
    summary: {
      walkins: rows.length,
      uniqueWalkins,
      jobs: allJobs.size,
      invoices: allInvoices.size,
      revenue: Math.round(revenue),
      averageWalkinsPerDay: round1(rows.length / activeDays),
      averageJobsPerDay: round1(allJobs.size / activeDays),
      averageInvoicesPerDay: round1(allInvoices.size / activeDays),
      walkinToJob: uniqueWalkins ? round1((allJobs.size / uniqueWalkins) * 100) : 0,
      walkinToInvoice: uniqueWalkins ? round1((allInvoices.size / uniqueWalkins) * 100) : 0,
      handlingMinutes: Math.round(average(handlingSamples)),
      handlingSample: handlingSamples.length,
      repairTatMinutes: Math.round(average(carryInTat)),
      nRepairTatMinutes: Math.round(average(nRepairTat)),
      totalRepairs: completed.length,
      carryInRepairs: completed.filter((r) => r.repairType.toLowerCase().includes("carry")).length,
      nRepairRepairs: completed.filter((r) => !r.repairType.toLowerCase().includes("carry")).length,
      tatSample: carryInTat.length + nRepairTat.length,
    },
    daily,
    people,
    technicians,
    tatCases: tatCases.slice(0, 120),
    tatCaseTotal: tatCases.length,
  };
}

/* ------------------------------------------------------------------ */
/* Entry point                                                         */
/* ------------------------------------------------------------------ */

export function computeOps(input: OpsInput): OpsResult {
  const nowTs = input.nowTs ?? Date.now() + IST_OFFSET_MS;

  // The workbook is updated during the month, so the operational month is the
  // current one when walk-ins exist for it, otherwise the latest month present.
  const currentMonth = monthKeyOf(nowTs);
  const months = Array.from(
    new Set(
      input.walkin
        .filter((r) => r.ts !== null)
        .map((r) => monthKeyOf(r.ts as number)),
    ),
  ).sort();
  const monthKey = months.includes(currentMonth)
    ? currentMonth
    : months.length
      ? months[months.length - 1]
      : currentMonth;

  return {
    process: computeProcessDashboard(input, monthKey, nowTs),
    productivity: computeProductivityDashboard(input, monthKey),
  };
}
