import {
  incentiveForPct,
  normalizePerson,
  revenueTypeKey,
  type AmcSlab,
  type BundleSlab,
  type CrmRow,
  type IncentiveSlab,
  type KpiTarget,
  type TargetSheet,
} from "./sheets";

export const WORKING_DAYS_PER_MONTH = 30;
export const HOLIDAYS = new Set<number>([1, 15, 28]);

export const PEOPLE_KEYS = ["SUBHA", "DINESH"] as const;
export type PersonKey = (typeof PEOPLE_KEYS)[number];

export interface DayRow {
  day: number;
  dateKey: string;
  weekday: string;
  holiday: boolean;
  /** Rows in the CRM sheet with no InvoiceDate (cannot be placed on a day) */
  undated: boolean;
  uniqueInvoices: number;
  serviceRevenue: number;
  accessoryRevenue: number;
  repairRevenue: number;
  totalRevenue: number;
  grossProfit: number;
  /** Per-person achievement for the day */
  byPerson: Record<
    string,
    { serviceRevenue: number; accessoryRevenue: number; repairRevenue: number; totalRevenue: number; grossProfit: number; invoices: number }
  >;
}

export interface StoreMetrics {
  daysElapsed: number;
  daysRemaining: number;
  /** Working (non-holiday) days left in the month, including the current day */
  remainingWorkingDays: number;
  monthProgressLabel: string;
  totalInvoices: number;
  uniqueInvoices: number;
  total: { target: number; achieved: number; pct: number; balance: number; runRate: number; requiredPerDay: number; projected: number };
  service: { target: number; achieved: number; pct: number; balance: number; runRate: number; requiredPerDay: number; projected: number };
  accessory: { target: number; achieved: number; pct: number; balance: number; runRate: number; requiredPerDay: number; projected: number };
  repair: { target: number; achieved: number; pct: number; balance: number; runRate: number; requiredPerDay: number; projected: number };
  grossProfit: { target: number; achieved: number; pct: number; balance: number; runRate: number; requiredPerDay: number; projected: number };
}

export interface SalespersonRow {
  key: string;
  name: string;
  invoiceCount: number;
  total: { target: number; achieved: number; pct: number; balance: number; incentive: number };
  service: { target: number; achieved: number; pct: number; balance: number; incentive: number };
  accessory: { target: number; achieved: number; pct: number; balance: number; incentive: number };
  repair: { target: number; achieved: number; pct: number; balance: number; incentive: number };
  totalIncentive: number;
}

export interface AmcRow {
  key: string;
  name: string;
  amcCount: number;
  /** Slab the count currently falls in (its rate may be Nil) */
  amcSlab: AmcSlab | null;
  /** Next incentivised slab to push for */
  amcNextSlab: AmcSlab | null;
  /** Count still required to reach the next slab */
  amcRemaining: number;
  /** Incentive earned on the current slab */
  amcEligible: number;
  accessoryCount: number;
  accessorySlab: AmcSlab | null;
  accessoryNextSlab: AmcSlab | null;
  accessoryRemaining: number;
  accessoryEligible: number;
  projectedIncentive: number;
}

export interface BundleRow {
  key: string;
  name: string;
  count: number;
  /** Slab the count currently falls in */
  slab: BundleSlab | null;
  nextSlab: BundleSlab | null;
  /** Count still required to reach the next slab */
  remaining: number;
  projectedIncentive: number;
}

export interface KpiRow extends KpiTarget {
  achieved: number;
  balance: number;
  achievedValue: number;
  achievedGp: number;
  pct: number;
  /** Unique-invoice achievement for the criterion, per salesperson */
  byPerson: Record<string, number>;
  /** Extra breakdown rows, e.g. SRD/SHC/ARD code chips or product mix */
  sub?: Array<{ label: string; count: number; value: number; gp: number }>;
}

/** "Daily target achievement idea" — how many units/day are needed to close a KPI gap. */
export interface KpiIdea {
  code: string;
  label: string;
  category: KpiTarget["category"];
  target: number;
  achieved: number;
  balance: number;
  perDay: number;
  hint: string;
}

export interface RevenueTypeStat {
  key: string;
  label: string;
  revenue: number;
  gp: number;
  uniqueInvoices: number;
}

export interface KpiSnapshot {
  ar: KpiRow[];
  rr: KpiRow[];
  sr: KpiRow[];
  googleReview: { previousMonth: number; currentMonth: number; difference: number };
  threePP: { achieved: number; target: number };
  /** Units/day required per criterion to close the month's gap */
  dailyIdeas: KpiIdea[];
}

export interface BundleOpportunity {
  invoiceCode: string;
  dayKey: string;
  salesPerson: string;
  accessoryNet: number;
  accessoryGp: number;
  accessoryCount: number;
}

export interface DashboardData {
  crmRows: CrmRow[];
  target: TargetSheet;
  days: DayRow[];
  activeDays: number;
  store: StoreMetrics;
  /** Total + unique invoice split per revenue type */
  revenueTypes: RevenueTypeStat[];
  salespeople: SalespersonRow[];
  amc: AmcRow[];
  bundle: BundleRow[];
  bundleOpportunities: BundleOpportunity[];
  kpi: KpiSnapshot;
  generatedAt: number;
}

const WD = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];

function isWorkingDay(day: number): boolean {
  return !HOLIDAYS.has(day);
}

export function countWorkingDays(fromDay: number, toDay: number): number {
  let n = 0;
  for (let d = fromDay; d <= toDay; d++) if (isWorkingDay(d)) n++;
  return n;
}

function pct(part: number, whole: number): number {
  if (whole <= 0) return 0;
  return Math.round((part / whole) * 1000) / 10;
}

function safeDiv(a: number, b: number): number {
  return b > 0 ? a / b : 0;
}

// ---------------------------------------------------------------------------
// Aggregation
// ---------------------------------------------------------------------------

export function computeDashboard(crmRows: CrmRow[], target: TargetSheet): DashboardData {
  // ----- month / day scaffolding -----
  const workingDays: number[] = [];
  for (let d = 1; d <= WORKING_DAYS_PER_MONTH; d++) if (isWorkingDay(d)) workingDays.push(d);
  const maxDay = crmRows.reduce((m, r) => Math.max(m, r.day), 0);
  const daysElapsed = Math.max(1, Math.min(WORKING_DAYS_PER_MONTH, maxDay || 1));
  const daysRemaining = WORKING_DAYS_PER_MONTH - daysElapsed;
  const activeDays = workingDays.filter((d) => d <= daysElapsed).length;

  // working days remaining AFTER the last data day (used for forward requirements)
  const remainingWorkingAfter = workingDays.filter((d) => d > daysElapsed).length;
  // working days remaining including today
  const remainingWorkingToday = workingDays.filter((d) => d >= daysElapsed).length;

  // ----- day rows -----
  const dayMap = new Map<number, DayRow>();
  const ensureDay = (day: number): DayRow => {
    let row = dayMap.get(day);
    if (!row) {
      const sample = crmRows.find((r) => r.day === day);
      const dateKey = sample?.dayKey ?? String(day).padStart(2, "0");
      const [dd, mm, yy] = dateKey.split("-");
      const weekday = dd && mm && yy ? WD[new Date(Number(yy), Number(mm) - 1, Number(dd)).getDay()] : "";
      row = {
        day,
        dateKey,
        weekday,
        holiday: !isWorkingDay(day),
        undated: day <= 0,
        uniqueInvoices: 0,
        serviceRevenue: 0,
        accessoryRevenue: 0,
        repairRevenue: 0,
        totalRevenue: 0,
        grossProfit: 0,
        byPerson: {},
      };
      dayMap.set(day, row);
    }
    return row;
  };

  // ----- invoice-level indexes -----
  const invoiceSet = new Set<string>();
  const invoiceDay = new Map<string, number>();
  const invoicePerson = new Map<string, string>();
  const invoiceHasServiceOrRepair = new Set<string>();
  const invoiceAccessoryValue = new Map<string, { net: number; gp: number; count: number }>();
  const invoiceAccessoryFlag = new Set<string>();
  const invoiceScheme = new Set<string>();
  const invoicePersonSet = new Map<string, Set<string>>();

  for (const r of crmRows) {
    const day = ensureDay(r.day);
    invoiceSet.add(r.invoiceCode);
    invoiceDay.set(r.invoiceCode, r.day);
    const pk = normalizePerson(r.salesPerson) || "UNASSIGNED";
    invoicePerson.set(r.invoiceCode, pk);
    if (!invoicePersonSet.has(r.invoiceCode)) invoicePersonSet.set(r.invoiceCode, new Set());
    invoicePersonSet.get(r.invoiceCode)!.add(pk);

    const rt = revenueTypeKey(r.revenueType);
    if (rt === "service" || rt === "repair") invoiceHasServiceOrRepair.add(r.invoiceCode);
    if (rt === "accessory") {
      invoiceAccessoryFlag.add(r.invoiceCode);
      const prev = invoiceAccessoryValue.get(r.invoiceCode) ?? { net: 0, gp: 0, count: 0 };
      invoiceAccessoryValue.set(r.invoiceCode, {
        net: prev.net + r.netAmount,
        gp: prev.gp + r.grossProfit,
        count: prev.count + 1,
      });
    }
    if (r.schemeCode && r.schemeCode.toUpperCase().startsWith("DISC")) {
      invoiceScheme.add(r.invoiceCode);
    }

    // row-level day aggregation (positive amounts only)
    if (r.netAmount >= 0) {
      day.serviceRevenue += rt === "service" ? r.netAmount : 0;
      day.accessoryRevenue += rt === "accessory" ? r.netAmount : 0;
      day.repairRevenue += rt === "repair" ? r.netAmount : 0;
      day.totalRevenue += r.netAmount;
      day.grossProfit += r.grossProfit;
      const bp = day.byPerson[pk] ?? {
        serviceRevenue: 0,
        accessoryRevenue: 0,
        repairRevenue: 0,
        totalRevenue: 0,
        grossProfit: 0,
        invoices: 0,
      };
      bp.serviceRevenue += rt === "service" ? r.netAmount : 0;
      bp.accessoryRevenue += rt === "accessory" ? r.netAmount : 0;
      bp.repairRevenue += rt === "repair" ? r.netAmount : 0;
      bp.totalRevenue += r.netAmount;
      bp.grossProfit += r.grossProfit;
      day.byPerson[pk] = bp;
    }
  }

  // unique invoices per day & per person-day
  const dayUnique = new Map<number, Set<string>>();
  for (const code of invoiceSet) {
    const d = invoiceDay.get(code) ?? 0;
    if (!dayUnique.has(d)) dayUnique.set(d, new Set());
    dayUnique.get(d)!.add(code);
  }
  const personDayInvoices = new Map<string, Set<string>>();
  for (const code of invoiceSet) {
    const pk = invoicePerson.get(code) ?? "UNASSIGNED";
    const d = invoiceDay.get(code) ?? 0;
    const k = `${pk}|${d}`;
    if (!personDayInvoices.has(k)) personDayInvoices.set(k, new Set());
    personDayInvoices.get(k)!.add(code);
  }
  for (const [d, set] of dayUnique) {
    ensureDay(d).uniqueInvoices = set.size;
  }
  const personInvoiceTotals = new Map<string, Set<string>>();
  for (const code of invoiceSet) {
    for (const pk of invoicePersonSet.get(code) ?? []) {
      if (!personInvoiceTotals.has(pk)) personInvoiceTotals.set(pk, new Set());
      personInvoiceTotals.get(pk)!.add(code);
    }
  }

  // Undated rows (blank InvoiceDate) are listed last instead of leading as "day 0".
  const days = Array.from(dayMap.values()).sort(
    (a, b) => (a.undated ? 99 : a.day) - (b.undated ? 99 : b.day),
  );

  // ----- store-level revenue rollups -----
  const sum = (f: (r: CrmRow) => number, filter?: (r: CrmRow) => boolean) =>
    crmRows.reduce((acc, r) => (r.netAmount >= 0 && (!filter || filter(r)) ? acc + f(r) : acc), 0);

  const sStore = {
    service: sum((r) => r.netAmount, (r) => revenueTypeKey(r.revenueType) === "service"),
    accessory: sum((r) => r.netAmount, (r) => revenueTypeKey(r.revenueType) === "accessory"),
    repair: sum((r) => r.netAmount, (r) => revenueTypeKey(r.revenueType) === "repair"),
    total: sum((r) => r.netAmount),
    grossProfit: sum((r) => r.grossProfit),
  };

  const bucket = (
    tgt: number,
    ach: number,
    elapsedWorking: number,
    remAfter: number,
  ) => {
    const runRate = safeDiv(ach, elapsedWorking);
    return {
      target: tgt,
      achieved: ach,
      pct: pct(ach, tgt),
      balance: Math.max(0, tgt - ach),
      runRate: Math.round(runRate),
      requiredPerDay: Math.round(safeDiv(Math.max(0, tgt - ach), Math.max(1, remAfter + 1))),
      projected: Math.round(runRate * WORKING_DAYS_PER_MONTH),
    };
  };

  const store: StoreMetrics = {
    daysElapsed,
    daysRemaining,
    remainingWorkingDays: remainingWorkingToday,
    monthProgressLabel: `${daysElapsed} of ${WORKING_DAYS_PER_MONTH} days`,
    totalInvoices: crmRows.length,
    uniqueInvoices: invoiceSet.size,
    total: bucket(target.store.total, sStore.total, activeDays, remainingWorkingAfter),
    service: bucket(target.store.service, sStore.service, activeDays, remainingWorkingAfter),
    accessory: bucket(target.store.accessory, sStore.accessory, activeDays, remainingWorkingAfter),
    repair: bucket(target.store.repair, sStore.repair, activeDays, remainingWorkingAfter),
    grossProfit: bucket(target.store.grossProfit, sStore.grossProfit, activeDays, remainingWorkingAfter),
  };

  // ----- revenue type totals & unique invoice counts -----
  const revenueTypes: RevenueTypeStat[] = (
    [
      ["service", "Service Revenue"],
      ["accessory", "Accessories Revenue"],
      ["repair", "Repair Revenue"],
    ] as const
  ).map(([key, label]) => {
    const set = new Set<string>();
    let revenue = 0;
    let gp = 0;
    for (const r of crmRows) {
      if (r.netAmount < 0) continue;
      if (revenueTypeKey(r.revenueType) !== key) continue;
      set.add(r.invoiceCode);
      revenue += r.netAmount;
      gp += r.grossProfit;
    }
    return {
      key,
      label,
      revenue: Math.round(revenue),
      gp: Math.round(gp),
      uniqueInvoices: set.size,
    };
  });

  // ----- salesperson revenue & incentive -----
  const salespeople: SalespersonRow[] = PEOPLE_KEYS.map((key) => {
    const t = target.salespeople[key] ?? {
      service: 0,
      accessory: 0,
      repair: 0,
      total: 0,
      grossProfit: 0,
    };
    const mine = crmRows.filter((r) => normalizePerson(r.salesPerson) === key);
    const mineInvoices = personInvoiceTotals.get(key) ?? new Set<string>();
    const sumMine = (f: (r: CrmRow) => number, filter?: (r: CrmRow) => boolean) =>
      mine.reduce((acc, r) => (r.netAmount >= 0 && (!filter || filter(r)) ? acc + f(r) : acc), 0);
    const ach = {
      service: sumMine((r) => r.netAmount, (r) => revenueTypeKey(r.revenueType) === "service"),
      accessory: sumMine((r) => r.netAmount, (r) => revenueTypeKey(r.revenueType) === "accessory"),
      repair: sumMine((r) => r.netAmount, (r) => revenueTypeKey(r.revenueType) === "repair"),
      total: 0,
    };
    ach.total = ach.service + ach.accessory + ach.repair;

    const svc = incentiveForPct(target.serviceIncentive.service, pct(ach.service, t.service));
    const acc = incentiveForPct(target.serviceIncentive.accessory, pct(ach.accessory, t.accessory));

    const mk = (tg: number, a: number, inc: number) => ({
      target: tg,
      achieved: a,
      pct: pct(a, tg),
      balance: Math.max(0, tg - a),
      incentive: inc,
    });

    return {
      key,
      name: key.charAt(0) + key.slice(1).toLowerCase(),
      invoiceCount: mineInvoices.size,
      total: mk(t.total, ach.total, svc + acc),
      service: mk(t.service, ach.service, svc),
      accessory: mk(t.accessory, ach.accessory, acc),
      repair: mk(t.repair, ach.repair, 0),
      totalIncentive: svc + acc,
    } as SalespersonRow;
  });

  // ----- AMC & Bundle tracking (per person) -----
  // AMC KPI rule: crm!SubCategory == "AMC Charge", unique invoice codes.
  // Accessory-AMC variant: AMC invoice with no service/repair line on it.
  const amcByPerson = new Map<string, Set<string>>();
  const accAmcByPerson = new Map<string, Set<string>>();
  const bundleByPerson = new Map<string, Set<string>>();
  for (const r of crmRows) {
    const pk = normalizePerson(r.salesPerson);
    const rt = revenueTypeKey(r.revenueType);
    if (r.subCategory.toUpperCase() === "AMC CHARGE" && r.netAmount >= 0) {
      if (!amcByPerson.has(pk)) amcByPerson.set(pk, new Set());
      amcByPerson.get(pk)!.add(r.invoiceCode);
      if (!invoiceHasServiceOrRepair.has(r.invoiceCode)) {
        if (!accAmcByPerson.has(pk)) accAmcByPerson.set(pk, new Set());
        accAmcByPerson.get(pk)!.add(r.invoiceCode);
      }
    }
    if (r.schemeCode.toUpperCase().startsWith("DISC") && r.netAmount >= 0) {
      if (!bundleByPerson.has(pk)) bundleByPerson.set(pk, new Set());
      bundleByPerson.get(pk)!.add(r.invoiceCode);
    }
  }

  // Accessory-AMC tiers are keyed on their accessory count threshold; the
  // best-paying tier the count qualifies for wins.
  const accTiers = target.amcSlabs.filter(
    (s) => (s.accCount ?? 0) > 0 && (s.accRate ?? 0) > 0,
  );
  const accPayout = (t: AmcSlab, count: number) =>
    count * (t.accRate ?? 0) + (t.accFixed ?? 0);

  const amc: AmcRow[] = PEOPLE_KEYS.map((key) => {
    const amcCount = amcByPerson.get(key)?.size ?? 0;
    const accessoryCount = accAmcByPerson.get(key)?.size ?? 0;
    const slabs = target.amcSlabs;
    const amcSlab = slabs.filter((s) => amcCount >= s.min).slice(-1)[0] ?? null;
    const amcNextSlab =
      slabs.filter((s) => s.min > amcCount && s.rate > 0).sort((a, b) => a.min - b.min)[0] ??
      null;
    const amcEligible = amcSlab ? amcCount * amcSlab.rate + amcSlab.fixed : 0;
    const amcRemaining = amcNextSlab ? Math.max(0, amcNextSlab.min - amcCount) : 0;

    const accSlab =
      accTiers
        .filter((t) => accessoryCount >= (t.accCount ?? 0))
        .sort((a, b) => accPayout(b, accessoryCount) - accPayout(a, accessoryCount))[0] ?? null;
    const accNextSlab =
      accTiers
        .filter((t) => (t.accCount ?? 0) > accessoryCount)
        .sort((a, b) => (a.accCount ?? 0) - (b.accCount ?? 0))[0] ?? null;
    const accEligible = accSlab ? accPayout(accSlab, accessoryCount) : 0;
    const accRemaining = accNextSlab
      ? Math.max(0, (accNextSlab.accCount ?? 0) - accessoryCount)
      : 0;

    return {
      key,
      name: key.charAt(0) + key.slice(1).toLowerCase(),
      amcCount,
      amcSlab,
      amcNextSlab,
      amcRemaining,
      amcEligible,
      accessoryCount,
      accessorySlab: accSlab,
      accessoryNextSlab: accNextSlab,
      accessoryRemaining: accRemaining,
      accessoryEligible: accEligible,
      projectedIncentive: amcEligible + accEligible,
    };
  });

  const bundle: BundleRow[] = PEOPLE_KEYS.map((key) => {
    const count = bundleByPerson.get(key)?.size ?? 0;
    const slabs = target.bundleSlabs;
    const slab = slabs.filter((s) => count >= s.min).slice(-1)[0] ?? null;
    const nextSlab =
      slabs.filter((s) => s.min > count).sort((a, b) => a.min - b.min)[0] ?? null;
    return {
      key,
      name: key.charAt(0) + key.slice(1).toLowerCase(),
      count,
      slab,
      nextSlab,
      remaining: nextSlab ? Math.max(0, nextSlab.min - count) : 0,
      projectedIncentive: slab ? count * slab.perBundle : 0,
    };
  });

  // bundle opportunities: accessory-only invoices
  const bundleOpportunities: BundleOpportunity[] = [];
  for (const code of invoiceAccessoryFlag) {
    if (invoiceHasServiceOrRepair.has(code)) continue;
    const v = invoiceAccessoryValue.get(code);
    if (!v) continue;
    bundleOpportunities.push({
      invoiceCode: code,
      dayKey: crmRows.find((r) => r.invoiceCode === code)?.dayKey ?? "",
      salesPerson: invoicePerson.get(code) ?? "UNASSIGNED",
      accessoryNet: v.net,
      accessoryGp: v.gp,
      accessoryCount: v.count,
    });
  }
  bundleOpportunities.sort((a, b) => a.dayKey.localeCompare(b.dayKey));

  // ----- KPI computations -----
  const kpi = computeKpis(crmRows, target, invoiceSet, remainingWorkingToday);

  return {
    crmRows,
    target,
    days,
    activeDays,
    store,
    revenueTypes,
    salespeople,
    amc,
    bundle,
    bundleOpportunities,
    kpi,
    generatedAt: Date.now(),
  };
}

// ---------------------------------------------------------------------------
// KPI rule engine
// ---------------------------------------------------------------------------

const itemCodeUp = (r: CrmRow) => r.itemCode.toUpperCase();
const descUp = (r: CrmRow) => r.itemDescription.toUpperCase();

/** Short, actionable line per KPI shown in the daily-ideas panel. */
const KPI_HINTS: Record<string, string> = {
  CAREPLUS: "Pitch Care+ at every service check-in",
  PROTECTPLUS: "Bundle Protect+ with device sales",
  AMC: "Convert repair walk-ins into AMC plans",
  BUNDLE: "Attach DISC bundle schemes to accessory bills",
  DISPLAY: "Quote display replacement on cracked iPhones",
  BATTERY: "Run a battery-health check on every iPhone service",
  GLASS: "Upsell glass protection on iPhone repairs",
  OTHERS: "Log repair revenue lines that use non-ACSH part codes",
  SRD: "Close standard diagnosis jobs (SRD codes)",
  SHC: "Push standard hardware-care jobs (SHC codes)",
  ARD: "Move complex repairs onto ARD codes",
};

export function computeKpis(
  crmRows: CrmRow[],
  target: TargetSheet,
  invoiceSet: Set<string>,
  remainingWorkingDays: number,
): KpiSnapshot {
  void invoiceSet;
  const kpiByKey = new Map<string, KpiTarget>();
  for (const k of target.kpis) kpiByKey.set(k.code, k);

  const getKpi = (code: string): KpiTarget => ({
    category: kpiByKey.get(code)?.category ?? "AR",
    code,
    label: kpiByKey.get(code)?.label ?? code,
    target: kpiByKey.get(code)?.target ?? 0,
    value: kpiByKey.get(code)?.value ?? 0,
    approxAmount: kpiByKey.get(code)?.approxAmount ?? 0,
  });

  // Salesperson attribution for KPI counts (first salesperson seen per invoice).
  const invoicePerson = new Map<string, string>();
  for (const r of crmRows) {
    const pk = normalizePerson(r.salesPerson);
    if (pk && !invoicePerson.has(r.invoiceCode)) invoicePerson.set(r.invoiceCode, pk);
  }

  const mkRow = (k: KpiTarget, invoices: Set<string>): KpiRow => {
    let achievedValue = 0;
    let achievedGp = 0;
    for (const r of crmRows) {
      if (r.netAmount >= 0 && invoices.has(r.invoiceCode)) {
        achievedValue += r.netAmount;
        achievedGp += r.grossProfit;
      }
    }
    const achieved = invoices.size;
    const byPerson: Record<string, number> = {};
    for (const code of invoices) {
      const pk = invoicePerson.get(code);
      if (pk) byPerson[pk] = (byPerson[pk] ?? 0) + 1;
    }
    return {
      ...k,
      achieved,
      balance: Math.max(0, k.target - achieved),
      achievedValue: Math.round(achievedValue),
      achievedGp: Math.round(achievedGp),
      pct: pct(achieved, k.target),
      byPerson,
    };
  };

  // ---- AR KPIs (unique invoice counts) ----
  const careplus = new Set<string>();
  const protectplus = new Set<string>();
  const amcKpi = new Set<string>();
  const bundleKpi = new Set<string>();
  const threePP = new Set<string>();

  for (const r of crmRows) {
    if (r.netAmount < 0) continue;
    const pt = r.partType.toUpperCase();
    const rt = revenueTypeKey(r.revenueType);
    if (pt === "ACPLUS") careplus.add(r.invoiceCode);
    if (pt === "PROTECTPLUS") protectplus.add(r.invoiceCode);
    if (r.subCategory.toUpperCase() === "AMC CHARGE") amcKpi.add(r.invoiceCode);
    if (r.schemeCode.toUpperCase().startsWith("DISC")) bundleKpi.add(r.invoiceCode);
    if (rt === "accessory") threePP.add(r.invoiceCode);
  }
  // 3PP: accessory invoices that ALSO have service or repair revenue on the same invoice
  const invoiceHasSvcRep = new Set<string>();
  for (const r of crmRows) {
    const rt = revenueTypeKey(r.revenueType);
    if ((rt === "service" || rt === "repair") && r.netAmount >= 0) invoiceHasSvcRep.add(r.invoiceCode);
  }
  const threePPFinal = new Set(Array.from(threePP).filter((c) => invoiceHasSvcRep.has(c)));

  // ---- RR KPIs ----
  const rrDisplay = new Set<string>();
  const rrBattery = new Set<string>();
  const rrGlass = new Set<string>();
  const rrOthers = new Set<string>();

  for (const r of crmRows) {
    if (r.netAmount < 0) continue;
    const rt = revenueTypeKey(r.revenueType);
    const d = descUp(r);
    const ic = itemCodeUp(r);
    const isAcsh = ic.startsWith("ACSH");
    const isIphone = d.includes("IPHONE");
    if (rt === "repair") {
      if (isIphone && d.includes("DISPLAY") && !isAcsh) rrDisplay.add(r.invoiceCode);
      if (isIphone && d.includes("BATTERY")) rrBattery.add(r.invoiceCode);
      if (isIphone && d.includes("GLASS") && !isAcsh) rrGlass.add(r.invoiceCode);
      if (!isAcsh) rrOthers.add(r.invoiceCode);
    }
  }

  // ---- SR KPIs ----
  const srSrd = new Set<string>();
  const srShc = new Set<string>();
  const srArd = new Set<string>();
  for (const r of crmRows) {
    if (r.netAmount < 0) continue;
    const ic = itemCodeUp(r);
    if (ic.startsWith("SRD")) srSrd.add(r.invoiceCode);
    if (ic.startsWith("SHC")) srShc.add(r.invoiceCode);
    if (ic.startsWith("ARD")) srArd.add(r.invoiceCode);
  }

  // SR product mix (iph/air/mac/wat/pad/con inside item code), per SR code
  // so the dashboard can render an SHC/SRD/ARD × product matrix.
  const productMix = new Map<string, Map<string, { count: Set<string>; value: number; gp: number }>>();
  const PRODUCT_CODES: Array<[string, string]> = [
    ["IPH", "iPhone"],
    ["AIR", "AirPods"],
    ["MAC", "Mac"],
    ["WAT", "Watch"],
    ["PAD", "iPad"],
    ["CON", "Consumables"],
  ];
  for (const r of crmRows) {
    if (r.netAmount < 0) continue;
    const ic = itemCodeUp(r);
    const srCode = ic.startsWith("SRD") ? "SRD" : ic.startsWith("SHC") ? "SHC" : ic.startsWith("ARD") ? "ARD" : null;
    if (!srCode) continue;
    // Product family is the token sitting between the separators of the item
    // code (e.g. "SRD-IPH-002" -> IPH -> iPhone).
    const tokens = ic.split(/[^A-Z0-9]+/);
    for (const [code, label] of PRODUCT_CODES) {
      if (!tokens.includes(code)) continue;
      if (!productMix.has(srCode)) productMix.set(srCode, new Map());
      const byProduct = productMix.get(srCode)!;
      if (!byProduct.has(label)) byProduct.set(label, { count: new Set(), value: 0, gp: 0 });
      const entry = byProduct.get(label)!;
      entry.count.add(r.invoiceCode);
      entry.value += r.netAmount;
      entry.gp += r.grossProfit;
    }
  }

  const ar = [
    mkRow(getKpi("CAREPLUS"), careplus),
    mkRow(getKpi("PROTECTPLUS"), protectplus),
    mkRow(getKpi("AMC"), amcKpi),
    mkRow(getKpi("BUNDLE"), bundleKpi),
  ];
  const rr = [
    mkRow(getKpi("DISPLAY"), rrDisplay),
    mkRow(getKpi("BATTERY"), rrBattery),
    mkRow(getKpi("GLASS"), rrGlass),
    mkRow(getKpi("OTHERS"), rrOthers),
  ];
  const srBase = [
    mkRow(getKpi("SRD"), srSrd),
    mkRow(getKpi("SHC"), srShc),
    mkRow(getKpi("ARD"), srArd),
  ];
  const sr = srBase.map((row) => ({
    ...row,
    sub: Array.from((productMix.get(row.code) ?? new Map()).entries()).map(([label, v]) => ({
      label,
      count: v.count.size,
      value: Math.round(v.value),
      gp: Math.round(v.gp),
    })),
  }));

  const gr = target.googleReview;

  // "Daily target achievement ideas": units/day needed to close each KPI gap.
  const workingDays = Math.max(1, remainingWorkingDays);
  const dailyIdeas: KpiIdea[] = [...ar, ...rr, ...sr]
    .filter((k) => k.target > 0 && k.balance > 0)
    .map((k) => ({
      code: k.code,
      label: k.label,
      category: k.category,
      target: k.target,
      achieved: k.achieved,
      balance: k.balance,
      perDay: Math.ceil(k.balance / workingDays),
      hint: KPI_HINTS[k.code] ?? "",
    }))
    .sort((a, b) => b.balance - a.balance);

  return {
    ar,
    rr,
    sr,
    googleReview: {
      previousMonth: gr.previousMonth,
      currentMonth: gr.currentMonth,
      difference: gr.difference || Math.max(0, gr.currentMonth - gr.previousMonth),
    },
    threePP: {
      achieved: threePPFinal.size,
      target: kpiByKey.get("CAREPLUS")?.target ?? 0,
    },
    dailyIdeas,
  };
}
