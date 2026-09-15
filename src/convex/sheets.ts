import { action } from "./_generated/server";
import { v } from "convex/values";
import {
  buildSheetsFetchUrls,
  parseCsv,
  parseCrmRows,
  parseTargetSheet,
} from "../lib/sheets";

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
        service: v.array(v.object({ label: v.string(), value: v.number() })),
        accessory: v.array(v.object({ label: v.string(), value: v.number() })),
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
