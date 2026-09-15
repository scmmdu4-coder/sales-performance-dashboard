// Temporary verification script — compares computed metrics against expected values.
import {
  buildSheetsFetchUrls,
  parseCsv,
  parseCrmRows,
  parseTargetSheet,
} from "../src/lib/sheets";
import { computeDashboard } from "../src/lib/metrics";

const { crmUrl, targetUrl } = buildSheetsFetchUrls();
const [crmCsv, targetCsv] = await Promise.all([
  (await fetch(crmUrl, { redirect: "follow" })).text(),
  (await fetch(targetUrl, { redirect: "follow" })).text(),
]);
const crmRows = parseCrmRows(parseCsv(crmCsv));
const target = parseTargetSheet(parseCsv(targetCsv));
const data = computeDashboard(crmRows, target);

console.log("rows:", crmRows.length, "| invoices:", data.store.uniqueInvoices, "| lines:", data.store.totalInvoices);
console.log("days:", data.store.daysElapsed, "elapsed,", data.store.daysRemaining, "remaining");
console.log("store.total:", data.store.total);
console.log("google review:", data.kpi.googleReview);
console.log("kpis AR:", data.kpi.ar.map((k) => `${k.code}=${k.achieved}/${k.target}`).join(" "));
console.log("kpis RR:", data.kpi.rr.map((k) => `${k.code}=${k.achieved}/${k.target}`).join(" "));
console.log("kpis SR:", data.kpi.sr.map((k) => `${k.code}=${k.achieved}/${k.target}`).join(" "));
console.log("3PP:", data.kpi.threePP.achieved);
console.log("salespeople:", data.salespeople.map((p) => `${p.key} svc=${p.service.achieved}/${p.service.target} acc=${p.accessory.achieved}/${p.accessory.target} inc=${p.totalIncentive}`).join(" | "));
console.log("amc:", data.amc.map((a) => `${a.key} amc=${a.amcCount} acc=${a.accessoryCount} inc=${a.projectedIncentive}`).join(" | "));
console.log("bundle:", data.bundle.map((b) => `${b.key} n=${b.count} inc=${b.projectedIncentive}`).join(" | "));
