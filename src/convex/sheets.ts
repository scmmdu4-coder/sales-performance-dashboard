import { action, internalMutation, internalQuery } from "./_generated/server";
import { v } from "convex/values";
import { internal } from "./_generated/api";
import type { Doc, Id } from "./_generated/dataModel";
import {
  buildSheetsFetchUrls,
  parseCsv,
  parseCrmRows,
  parseTargetSheet,
} from "../lib/sheets";
import {
  buildOpsFetchUrls,
  computeOps,
  parseAodRows,
  parseCommRows,
  parseCsvLite,
  parseWalkinRows,
  type ProcessDashboard,
  type ProductivityDashboard,
} from "../lib/ops";

/**
 * Server-side fetch of the two public Google Sheets (crm + target).
 * Returns parsed rows + parsed target config so the client computes metrics
 * with one payload. Runs as a Convex action because it performs network I/O.
 */
export const fetchDashboardData = action({
  args: {},
  handler: async () => {
    const { crmUrl, targetUrl } = buildSheetsFetchUrls();
    const [crmRes, targetRes] = await Promise.all([
      fetch(crmUrl, { redirect: "follow", cache: "no-store" }),
      fetch(targetUrl, { redirect: "follow", cache: "no-store" }),
    ]);
    if (!crmRes.ok) {
      throw new Error(`CRM sheet fetch failed: ${crmRes.status} ${crmRes.statusText}`);
    }
    if (!targetRes.ok) {
      throw new Error(`Target sheet fetch failed: ${targetRes.status} ${targetRes.statusText}`);
    }
    const [crmCsv, targetCsv] = await Promise.all([crmRes.text(), targetRes.text()]);
    const crmRows = parseCrmRows(parseCsv(crmCsv));
    const target = parseTargetSheet(parseCsv(targetCsv));
    return {
      crmRows,
      target,
      fetchedAt: Date.now(),
      crmRowCount: crmRows.length,
    };
  },
  returns: v.object({
    crmRows: v.array(
      v.object({
        invoiceCode: v.string(),
        invoiceDate: v.string(),
        dayKey: v.string(),
        day: v.number(),
        salesPerson: v.string(),
        employeeCode: v.string(),
        partType: v.string(),
        subCategory: v.string(),
        revenueType: v.string(),
        itemCode: v.string(),
        itemDescription: v.string(),
        schemeCode: v.string(),
        quantity: v.number(),
        netAmount: v.number(),
        taxableAmount: v.number(),
        grossProfit: v.number(),
      }),
    ),
    target: v.object({
      store: v.object({
        service: v.number(),
        accessory: v.number(),
        repair: v.number(),
        total: v.number(),
        grossProfit: v.number(),
      }),
      salespeople: v.record(
        v.string(),
        v.object({
          service: v.number(),
          accessory: v.number(),
          repair: v.number(),
          total: v.number(),
          grossProfit: v.number(),
        }),
      ),
      serviceIncentive: v.object({
        service: v.array(
          v.object({
            label: v.string(),
            value: v.number(),
            minPct: v.number(),
            maxPct: v.number(),
          }),
        ),
        accessory: v.array(
          v.object({
            label: v.string(),
            value: v.number(),
            minPct: v.number(),
            maxPct: v.number(),
          }),
        ),
      }),
      amcSlabs: v.array(
        v.object({
          label: v.string(),
          rateLabel: v.string(),
          rate: v.number(),
          fixed: v.number(),
          min: v.number(),
          max: v.number(),
          incentive: v.number(),
          accCount: v.optional(v.number()),
          accRate: v.optional(v.number()),
          accFixed: v.optional(v.number()),
        }),
      ),
      bundleSlabs: v.array(
        v.object({ label: v.string(), min: v.number(), perBundle: v.number() }),
      ),
      kpis: v.array(
        v.object({
          category: v.string(),
          code: v.string(),
          label: v.string(),
          target: v.number(),
          value: v.number(),
          approxAmount: v.number(),
        }),
      ),
      googleReview: v.object({
        previousMonth: v.number(),
        currentMonth: v.number(),
        difference: v.number(),
      }),
    }),
    fetchedAt: v.number(),
    crmRowCount: v.number(),
  }),
});

/* ------------------------------------------------------------------ */
/* Operational sheets (process + productivity)                         */
/* ------------------------------------------------------------------ */

const repairPerformanceValidator = v.object({
  available: v.boolean(),
  currentMonth: v.string(),
  previousMonths: v.array(v.string()),
  totalCurrent: v.number(),
  previousThreeAverage: v.number(),
  difference: v.number(),
  percentVsAverage: v.number(),
  typeSummary: v.array(
    v.object({
      repairType: v.string(),
      current: v.number(),
      previousThree: v.array(v.number()),
      previousAverage: v.number(),
      difference: v.number(),
      percentVsAverage: v.number(),
    }),
  ),
  detail: v.array(
    v.object({
      repairType: v.string(),
      coverage: v.string(),
      product: v.string(),
      current: v.number(),
      previousAverage: v.number(),
      difference: v.number(),
      percentVsAverage: v.number(),
    }),
  ),
});

const processValidator = v.object({
  available: v.boolean(),
  month: v.string(),
  monthKey: v.string(),
  summary: v.object({
    total: v.number(),
    unique: v.number(),
    jobs: v.number(),
    invoices: v.number(),
    communications: v.number(),
    openJobs: v.number(),
    unclaimed: v.number(),
  }),
  daily: v.array(
    v.object({
      date: v.string(),
      dateKey: v.string(),
      total: v.number(),
      cases: v.number(),
      invoices: v.number(),
      communications: v.number(),
    }),
  ),
  repair: repairPerformanceValidator,
  openJobs: v.array(
    v.object({
      caseId: v.string(),
      customer: v.string(),
      mobile: v.string(),
      product: v.string(),
      coverage: v.string(),
      repairType: v.string(),
      caseDate: v.string(),
      caseAge: v.string(),
      bucket: v.string(),
      unclaimed: v.boolean(),
      lastCommunication: v.string(),
      communicationAging: v.string(),
    }),
  ),
  ageing: v.object({
    ages: v.array(v.string()),
    rows: v.array(
      v.object({
        bucket: v.string(),
        cells: v.array(v.number()),
        total: v.number(),
      }),
    ),
    totals: v.array(v.number()),
    grandTotal: v.number(),
  }),
});

const productivityValidator = v.object({
  available: v.boolean(),
  month: v.string(),
  monthKey: v.string(),
  summary: v.object({
    walkins: v.number(),
    uniqueWalkins: v.number(),
    jobs: v.number(),
    invoices: v.number(),
    revenue: v.number(),
    averageWalkinsPerDay: v.number(),
    averageJobsPerDay: v.number(),
    averageInvoicesPerDay: v.number(),
    walkinToJob: v.number(),
    walkinToInvoice: v.number(),
    handlingMinutes: v.number(),
    handlingSample: v.number(),
    repairTatMinutes: v.number(),
    nRepairTatMinutes: v.number(),
    totalRepairs: v.number(),
    carryInRepairs: v.number(),
    nRepairRepairs: v.number(),
    tatSample: v.number(),
  }),
  daily: v.array(
    v.object({
      date: v.string(),
      dateKey: v.string(),
      walkins: v.number(),
      jobs: v.number(),
      invoices: v.number(),
      jobConversion: v.number(),
      invoiceConversion: v.number(),
    }),
  ),
  people: v.array(
    v.object({
      name: v.string(),
      walkins: v.number(),
      uniqueWalkins: v.number(),
      jobs: v.number(),
      invoices: v.number(),
      walkinToJob: v.number(),
      walkinToInvoice: v.number(),
      invoiceRevenue: v.number(),
    }),
  ),
  technicians: v.array(
    v.object({
      name: v.string(),
      total: v.number(),
      carryIn: v.number(),
      nRepair: v.number(),
      carryInTat: v.string(),
      nRepairTat: v.string(),
      averageTat: v.string(),
    }),
  ),
  tatCases: v.array(
    v.object({
      caseId: v.string(),
      technician: v.string(),
      repairType: v.string(),
      customer: v.string(),
      mobile: v.string(),
      assignDate: v.string(),
      completeDate: v.string(),
      completeSource: v.string(),
      tat: v.string(),
      tatMinutes: v.number(),
    }),
  ),
  tatCaseTotal: v.number(),
});

const OPS_CACHE_KEY = "ops-aggregates-v1";
const OPS_CACHE_TTL_MS = 15 * 60 * 1000; // 15 minutes

/** Read the cached ops payload (no TTL check — the caller decides freshness). */
export const readOpsCache = internalQuery({
  args: { key: v.string() },
  handler: async (ctx, args): Promise<{ payload: unknown; createdAt: number } | null> => {
    const row = await ctx.db
      .query("kvCache")
      .withIndex("by_key", (q) => q.eq("key", args.key))
      .unique();
    return row ? { payload: row.payload, createdAt: row.createdAt } : null;
  },
});

/** Replace the cached ops payload, keeping at most one row per key. */
export const writeOpsCache = internalMutation({
  args: { key: v.string(), payload: v.any() },
  handler: async (ctx, args): Promise<void> => {
    const existing = await ctx.db
      .query("kvCache")
      .withIndex("by_key", (q) => q.eq("key", args.key))
      .unique();
    if (existing) {
      await ctx.db.patch(existing._id, { payload: args.payload, createdAt: Date.now() });
      return;
    }
    await ctx.db.insert("kvCache", {
      key: args.key,
      payload: args.payload,
      createdAt: Date.now(),
    });
  },
});

/**
 * Fetches the walkin / aod / comm sheets (plus the CRM lines for invoice
 * revenue) and returns the already-aggregated process + productivity views.
 * Only aggregates cross the wire — the raw AOD sheet is several megabytes.
 *
 * The aggregates are shared through the kvCache table for a short TTL, so a
 * team of viewers hits the expensive Google Sheets export at most once per
 * TTL window instead of on every page load. `force` skips the cache read but
 * still refreshes the stored copy for everyone else.
 */
interface OpsPayload {
  process: ProcessDashboard;
  productivity: ProductivityDashboard;
  fetchedAt: number;
}

type OpsActionResult = OpsPayload & { cache: "hit" | "miss" };

export const fetchOpsData = action({
  args: { force: v.optional(v.boolean()) },
  handler: async (ctx, args): Promise<OpsActionResult> => {
    const now = Date.now();
    if (!args.force) {
      const cached = await ctx.runQuery(internal.sheets.readOpsCache, { key: OPS_CACHE_KEY });
      if (cached && now - cached.createdAt < OPS_CACHE_TTL_MS) {
        const payload = cached.payload as {
          process: ProcessDashboard;
          productivity: ProductivityDashboard;
        };
        return {
          process: payload.process,
          productivity: payload.productivity,
          fetchedAt: cached.createdAt,
          cache: "hit",
        };
      }
    }

    const fresh = await fetchOpsFromSheets();
    await ctx.runMutation(internal.sheets.writeOpsCache, {
      key: OPS_CACHE_KEY,
      payload: { process: fresh.process, productivity: fresh.productivity },
    });
    return { ...fresh, cache: "miss" };
  },
  returns: v.object({
    process: processValidator,
    productivity: productivityValidator,
    fetchedAt: v.number(),
    cache: v.union(v.literal("hit"), v.literal("miss")),
  }),
});

async function fetchOpsFromSheets(): Promise<OpsPayload> {
  const { walkinUrl, aodUrl, commUrl } = buildOpsFetchUrls();
  const { crmUrl } = buildSheetsFetchUrls();
  const [walkinRes, aodRes, commRes, crmRes] = await Promise.all([
    fetch(walkinUrl, { redirect: "follow", cache: "no-store" }),
    fetch(aodUrl, { redirect: "follow", cache: "no-store" }),
    fetch(commUrl, { redirect: "follow", cache: "no-store" }),
    fetch(crmUrl, { redirect: "follow", cache: "no-store" }),
  ]);
  const failed = [
    ["walkin", walkinRes],
    ["aod", aodRes],
    ["comm", commRes],
    ["crm", crmRes],
  ].find(([, res]) => !(res as Response).ok);
  if (failed) {
    const [name, res] = failed as [string, Response];
    throw new Error(`${name} sheet fetch failed: ${res.status} ${res.statusText}`);
  }

  const [walkinCsv, aodCsv, commCsv, crmCsv] = await Promise.all([
    walkinRes.text(),
    aodRes.text(),
    commRes.text(),
    crmRes.text(),
  ]);

  const invoiceRevenue = new Map<string, number>();
  // Latest month present in the CRM lines — the operational views report on
  // the same month as the business dashboard.
  let preferredMonth: string | undefined;
  for (const row of parseCrmRows(parseCsv(crmCsv))) {
    invoiceRevenue.set(
      row.invoiceCode,
      (invoiceRevenue.get(row.invoiceCode) ?? 0) + row.netAmount,
    );
    const parts = row.dayKey.split("-");
    if (parts.length === 3 && /^\d{4}$/.test(parts[2])) {
      const month = `${parts[2]}-${parts[1]}`;
      if (preferredMonth === undefined || month > preferredMonth) preferredMonth = month;
    }
  }

  const result = computeOps({
    walkin: parseWalkinRows(parseCsvLite(walkinCsv)),
    aod: parseAodRows(parseCsvLite(aodCsv)),
    comm: parseCommRows(parseCsvLite(commCsv)),
    invoiceRevenue,
    preferredMonth,
  });

  return {
    process: result.process,
    productivity: result.productivity,
    fetchedAt: Date.now(),
  };
}
