# Sales Dashboard — Google Apps Script version

`SalesDashboard.gs` is a standalone Google Apps Script that reproduces the
same metrics logic as `src/lib/metrics.ts` (TranType = "SALES" filter,
TaxableAmount as achieved value, dynamic salesperson roster, AMC/bundle
slabs, KPI grid, invoice valuation, date-wise analysis) directly inside your
Google Sheet — no web app / Convex backend required.

## Setup

1. Open the Google Sheet that has your `crm` and `target` tabs.
2. `Extensions → Apps Script`.
3. Delete any placeholder code and paste in the full contents of
   `SalesDashboard.gs`.
4. Save the project.
5. Run the `buildDashboard` function once from the editor (▶ button) and
   accept the authorization prompt — it only touches this spreadsheet.
6. Reload the spreadsheet. A **📊 Sales Dashboard** menu appears; use
   "Build / Refresh Dashboard" any time the `crm`/`target` data changes.

## What it builds

A new **Dashboard** tab with:

1. Store target vs achievement vs balance, per revenue type
2. Working days: total / consumed / balance for the current month
3. Salesperson performance per revenue type (every name found in `crm`)
4. Incentive valuation — AMC & Bundle slabs, per salesperson
5. Store KPI parameters from the `target` KPI grid
6. Invoice valuation — total & unique invoices, per revenue type
7. Date-wise sales revenue analysis with per-revenue-type daily target
   (balance ÷ balance working days)

## Config

Edit the `CONFIG` object at the top of the script if your tab names differ
from `crm` / `target`, or to update the `HOLIDAYS` list for the current
month's non-working days.

## Notes

- Column lookup is by header name (case-insensitive, partial match), so
  column order in `crm` doesn't matter as long as the header text is
  recognizable (`InvoiceDate`, `SalesPersonName`, `RevenueType`,
  `ItemDescription`, `SubCategory`, `SchemeCode`, `TaxableAmount`,
  `GrossProfit`, `TranType`).
- The `target` sheet parser expects the same grid layout as the existing
  web dashboard (see comments marked `ANCHOR` in the script) — adjust the
  row/column numbers there if your target sheet layout differs.
