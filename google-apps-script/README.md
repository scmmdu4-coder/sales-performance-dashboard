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

## Where each number is read from (`target` sheet)

| Data | Range | Layout |
|---|---|---|
| Store + salesperson targets | `target!A1:D6` | Col A = revenue type, col B = store target, cols C/D = per-person targets (names in C1/D1) |
| Google reviews | `target!A9:C10` | A10 = previous month qty, B10 = current month count, C10 = difference |
| KPI targets | `target!B12:L14` | Row 12 = revenue type (AR/RR/SR), row 13 = KPI name, row 14 = target count |

KPI columns are matched by the **name** written in row 13 (Care+, Protect+,
AMC, Bundle, Display, Battery, Glass, Others, SRD, SHC, ARD), so you can
reorder or add columns without editing the script.

## Incentive rules

- **Service** and **Accessory** revenue are the only types that *earn*
  incentive, from the % achievement slabs on the target sheet.
- **Repair** revenue earns nothing on its own. It is a **gate**: a
  salesperson only actually receives their earned service + accessory
  incentive once repair revenue reaches **80 %** of their repair target.
  Below that, the amount is shown as "earned but withheld".
- Change the threshold with `CONFIG.REPAIR_INCENTIVE_GATE_PCT`.

## Daily target colouring

In section 7, the daily target is the **store total target ÷ number of
working days in the month**. Each day's achieved value is shown in
**green** when it meets or beats that figure and **red** when it falls
short. Each salesperson also gets their own colour band and left border so
their block is easy to pick out.

## Refresh button

The page header has a **↻ Refresh data** button. The rendered page is
cached server-side for 90 seconds so ordinary loads are fast; pressing
Refresh reloads with `?refresh=1`, which bypasses the cache and recomputes
straight from the sheet. Tune with `CONFIG.CACHE_SECONDS` (set it to `0` to
disable caching entirely).

## Config

Edit the `CONFIG` object at the top of the script if your tab names differ
from `crm` / `target`, to update the `HOLIDAYS` list for the current
month's non-working days, to change `REPAIR_INCENTIVE_GATE_PCT`, or to
adjust `CACHE_SECONDS`.

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

## Performance page (second tab)

The dashboard's top bar has a **Business / Performance** slider. Business is
the original page (unchanged). Performance is a second page that reads four
extra tabs. All four are looked up by tab name (case-insensitive, a few
spellings accepted) — see `CONFIG.WALKIN_SHEET_NAMES`, `SERVICE_SHEET_NAMES`,
`ADVANCE_SHEET_NAMES`, `COMM_SHEET_NAMES`. If a tab is missing, that section
simply shows guidance instead of breaking.

### Walk-in Analysis + Token Creation — `walkin` tab
| Column | Used for |
| --- | --- |
| `MobileNo` | every row = **one walk-in** (the visitor's mobile is the count) |
| `TokenDate` | the visit date (`yyyy-mm-dd hh:mm:ss`) — drives the date-wise table, newest first |
| `VisitPurposeDesc` | visit purpose, bucketed into fixed columns: **Collection** (contains "collect"), **General/Sales Enquiry** (contains "enqu/sales/general"), **Repair Submission** (contains "repair"); anything else lands in **Other** |
| `CaseId` | a row with a case id counts as a **case created** (+ case conversion %) |
| `InvoiceCode` | a row with an invoice code counts as an **invoice created** (+ invoice conversion %) |
| `**Remark` / `Remark` | a row with a remark text counts toward "remarks updated" |
| `TokenCreatedBy` *(optional)* | who generated the token → the **Token Creation** panel: total tokens generated vs each user ID's count and % share. Auto-detects `TokenCreatedBy` / `UserId` / `UserName` / `Staff` / `CreatedBy`. If absent, the panel explains how to add it. |

**Dead leads** — a walk-in with no CaseId, no InvoiceCode **and** no Remark
(never converted). Shown per day and as a total with dead-leads %, plus a
summary card.

### Job Management (open jobs only) — `service` tab
One row per case. **Duplicate case ids are de-duplicated; every listed job is "open".**
| Column | Used for |
| --- | --- |
| `CaseId` | unique job id (also the key for the communication lookup) |
| `FirstName` (+ `LastName`) | customer name |
| `MobileNo` | customer mobile — matched to the advance amount |
| `ProductDescription` | product |
| `CaseDate` | job created date (fallback for job age) |
| `RepairCreatedDate` | repair created date |
| `Ready_For_Pickup_Date_CRM` | ready-for-pickup date (row highlighted green when set) |
| `JobAge` | job age in days — if blank, computed as days since `CaseDate` |
| `Status` *(optional)* | any value containing deliver/closed/cancel/complete/collected is treated as **closed** and excluded. Otherwise the job is open. |
| `DeliveredDate` *(optional, used when there is no `Status`)* | a date here marks the job closed |


### Job Management — GSX / CRM repair status & unclaimed devices
- `GSXRepairStatusDescription` — shown as **GSX repair status**.
- `CRM_JobStatusDesc` — shown as **CRM repair status**.
- If either status contains **"Unclaimed Device"** (matched case-insensitively
  via `CONFIG.UNCLAIMED_KEYWORD`, default `"unclaim"`), the job stays in the
  open-jobs list but the whole row is highlighted in amber, the matching status
  cell becomes a solid badge, a "⚠ Unclaimed" tag appears next to the case id,
  and a count card appears at the top of the section.

### Job management — `advance` and `comm` tabs
- `advance`: `MobileNo` + `AvailableBalance` (typos like the tab name `advace` are accepted). Balances for the same mobile are **summed**.
- `comm`: `CaseId` + `CreatedDate` (+ optional `Remark`/`Message`). The job is **Communicated = Yes** when any matching case is found, and the **latest** `CreatedDate` is shown (otherwise **Not updated**). If the latest update is `COMM_STALE_DAYS` (default 2) old or more — or there is no update at all — the whole job row is highlighted red and a **Needs Follow-up** card counts it. Priority: unclaimed (amber) > stale communication (red) > ready for pickup (green).

### Tunables (CONFIG)
- `JOB_CLOSED_KEYWORDS` — words that mark a job closed.
- `JOB_AGE_WARN_DAYS` (7) / `JOB_AGE_CRITICAL_DAYS` (15) — job-age colour thresholds.
- `COMM_STALE_DAYS` (2) — how old the latest communication may be before a job row is highlighted as needing follow-up.

## Business page extras (new)

- **Avg Invoice Value tile** shows two details: *avg value against total
  invoice* (revenue ÷ unique invoices) and *avg invoice value against unique
  walk-in* (revenue ÷ distinct mobiles in the `walkin` sheet — the column is
  found by its `MobileNo` header, so `walkin!G:G` works wherever it sits).
- **Salesperson Performance** now lists the **store total revenue target**
  under each person's individual Repair revenue target.

## Performance page — current behaviour

- **Walk-in Analysis** is a true pivot: one row per date (each date tinted a
  different colour) and one column per **actual `VisitPurposeDesc` value** as
  written in the sheet — plus walk-ins, cases (case conv %), invoices (invoice
  conv %), remarks, dead leads and dead-leads % with a totals row.
- **Token Creation** is a table (not a chart): token created by · total tokens
  created · % of creation, with a total row.
- **Job Management** lists only cases whose `service!JobStatus` is **open**
  (non-open statuses are excluded and counted). The Ready-for-pickup column
  was removed. **Comm. updated** is the latest `comm!CreatedDate` for the case
  id: green when within `COMM_STALE_DAYS` (2) of today, red when older or when
  there is no record ("Not updated"). Advance paid comes from
  `advance!AvailableBalance` matched on `advance!MobileNo` ↔
  `service!MobileNo` (balances summed when a customer repeats).
- If there is **no `advance` tab** (or it matches nothing), the script falls
  back to a per-row `AvailableBalance` column read **directly from the
  `service` sheet** (any header matching `available balance`, e.g.
  `service!AvailableBalance`), and the table notes say which source is in
  use. Commit `e8378b1`.

## Latest changes (Tasks 12)

- **Tab rename** — "📈 Performance" is now **"🛠 Process"** (internal page id
  is unchanged; this is label-only).
- **Google Reviews tile (Section 3) is reversed** — the big number is now the
  **gained count** (e.g. `+19`, green, or red when negative), with the small
  caption `1494 this month · previous 1475`.
- **Avg Invoice Value tile (Section 3) shows three values**: avg invoice value
  ÷ total invoices (invoice lines), total invoice value ÷ unique walk-ins,
  and avg gross profit ÷ total invoices.
- **Date-wise execution table (Business Section 7)** uses full-cell
  *sensex-style* backgrounds: solid green (`#16a34a`) cells for target-achieved
  days and solid red (`#e23b3b`) cells for days below target, with the target
  value shown in translucent white inside the cell.
- **Salesperson daily targets (Section 4)** are now
  `balance ÷ working days remaining` per revenue type per engineer, i.e.
  `max(0, target − achieved) ÷ remaining working days` (today included when
  it is a working day) — instead of the earlier even `target ÷ total working
  days`. Clamped at 0 once the monthly target is already covered. The store
  monthly target daily split (Section 2) is unchanged.
- **Section 5 Approx Value cell adds a guidance line**: in addition to the
  approximate value of the achieved count it now shows
  `₹X to target` (balance count × the unit value from `target!B15:L15`), or
  "target covered" when the balance is 0. When a KPI has no unit value the
  cell shows an em-dash as before.

## Node test harnesses

`google-apps-script/tests/node/perftest.cjs` and `edge2.cjs` stub the Apps
Script services, feed mock sheets through the real `SalesDashboard.gs`, and
assert the computed values and rendered HTML (daily-target balance math,
reversed review tile, 3-value avg tile, approx-to-target line, sensex CSS,
advance fallback, comm 2-day window, KPI/pivot/token tables, …). Run either
file with plain `node` — they exit non-zero on failure.
