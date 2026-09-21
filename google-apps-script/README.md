# Sales Dashboard — Google Apps Script version

`SalesDashboard.gs` is a standalone Google Apps Script that reproduces the
same metrics logic as `src/lib/metrics.ts` (TranType = "SALES" filter,
TaxableAmount as achieved value, dynamic salesperson roster, AMC/bundle
slabs, KPI grid, invoice valuation, date-wise analysis).

It supports **two output modes** from the same script — pick either, or use
both:

## Mode A — a tab inside the spreadsheet

1. Open the Google Sheet that has your `crm` and `target` tabs.
2. `Extensions → Apps Script`.
3. Delete any placeholder code and paste in the full contents of
   `SalesDashboard.gs`.
4. Save the project.
5. Run the `buildDashboard` function once from the editor (▶ button) and
   accept the authorization prompt — it only touches this spreadsheet.
6. Reload the spreadsheet. A **📊 Sales Dashboard** menu appears; use
   "Build / Refresh Dashboard (sheet tab)" any time the `crm`/`target` data
   changes.

This writes a formatted **Dashboard** tab with cells/colors inside the
spreadsheet itself.

## Mode B — a real, standalone web page ("script page")

This is the mode to use if you want a dedicated URL you can open in any
browser, share with people who shouldn't need spreadsheet edit access, or
put in a bookmark / TV display, instead of a tab inside the sheet.

1. Do the paste-and-save step from Mode A (steps 1–4 above).
2. In the Apps Script editor, click **Deploy** (top-right) → **New
   deployment**.
3. Click the gear icon next to "Select type" → choose **Web app**.
4. Fill in:
   - Description: `Sales Dashboard`
   - Execute as: **Me** (your account)
   - Who has access: **Anyone with the link** (or "Anyone within your
     organization" if you're on Google Workspace and want to restrict it)
5. Click **Deploy**, then authorize when prompted.
6. Copy the **Web app URL** you're given. Open it in any browser — it
   renders a live, colorful HTML dashboard page, recomputed straight from
   the current `crm`/`target` sheet contents on every page load.
7. To refresh after editing sheet data, just reload the URL — no rebuild
   step needed.
8. If you edit the script later, go to **Deploy → Manage deployments** →
   pencil icon → **New version** → **Deploy**, so the URL picks up the
   change (Web app URLs are pinned to a specific deployed version).

The `onOpen()` menu also has a **"How do I get the Web App page?"** item
that shows these steps again as a quick reminder inside the sheet.

## What it computes (same for both modes)

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
- Mode B's `doGet(e)` is what Apps Script calls automatically when someone
  opens the deployed Web App URL — you never call it yourself. If you
  previously saw a "Script function not found: doGet" error, that was
  because the editor's function selector dropdown was set to `doGet`
  instead of `buildDashboard`; that only matters for Mode A, not Mode B.
