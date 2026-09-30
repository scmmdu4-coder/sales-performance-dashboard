# Sales Dashboard — Google Apps Script version

`SalesDashboard.gs` is a standalone Google Apps Script that reproduces the
same metrics logic as `src/lib/metrics.ts` (TranType = "SALES" filter,
TaxableAmount as achieved value, dynamic salesperson roster, AMC/bundle
slabs, KPI grid, invoice valuation, date-wise analysis).

## The web app — a real, standalone page ("script page")

> **Sheet-writer removed:** this script no longer creates a "Dashboard" tab
> inside your spreadsheet on deploy. The ONLY output is the standalone web
> page below.

This is the mode to use if you want a dedicated URL you can open in any
browser, share with people who shouldn't need spreadsheet edit access, or
put in a bookmark / TV display, instead of a tab inside the sheet.

1. Open your Google Sheet → `Extensions → Apps Script`, delete any
   placeholder code, paste the full `SalesDashboard.gs`, and save.
2. Click **Deploy** (top-right) → **New deployment**.
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


## What it computes

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
- The `target` sheet layout is anchored (see comments marked `ANCHOR` in
  the script): salesperson roster = **C1:J1** (only those names get cards
  and matrix columns — C-SAT labels or store codes elsewhere never show up
  as people); KPI grid = **B13:L16** (row 13 AR/RR/SR type, row 14 names,
  row 15 target counts, row 16 approx unit values); C-SAT block = period in
  **K1**, table in **L1:U5**.
- `doGet(e)` is what Apps Script calls automatically when someone opens the
  deployed Web App URL — you never call it yourself.

## Performance page (second tab)

The dashboard's top bar has **Business / Process / Performance** tabs.
Business is the original page. Process reads four extra tabs. All four are looked up by tab name (case-insensitive, a few
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

## Design system (latest)

The whole dashboard now uses the **exact look of the previous "script page"**
the user shared (its HTML/CSS was provided): light grey page gradient
(`#edf2f7 → #e7edf3`), dark navy cards
(`linear-gradient(145deg, rgba(23,49,84,.98), rgba(11,25,47,.98))` over
`rgba(12,29,53,.96)` tables with `#203852` headers), the Inter font stack,
tabular numbers, and the old palette tokens:

- achieved `#63e6b7` · target `#8fc7ff` · balance `#ffd477` · percent `#dfb2ff`
- muted `#aebdd0` · white `#f8fbff` · border `#3a5575` · danger `#ff9da7`
- refresh button `#172b43`; section headings on the light page `#182a40`.

Translating rules applied (no conditions or data touched):

- Date-wise "sensex" cells keep the same condition (day ≥ daily target) but
  now use the old page's pill palette as full-cell fills: `#63e6b7`/`#12352d`
  green and `#ff9da7`/`#5f2028` red.
- Salesperson column tints are subtle translucent versions of the accent
  colors (blue/amber/mint/purple/salmon) instead of pastel card colors.
- The walk-in analysis table stays on a **light** surface because its
  date rows use pastel `hsl(…,94%)` tints; every other table is dark navy.
- The background photo layer was removed so the page gradient matches the
  old page exactly.

### Design tweaks after the screenshots

- **Page title** is now `Store Performance Dashboard - <store code>`, with the
  store code read from **`crm!C2`** (the cell the old script used for its
  location header). The browser `<title>` uses the same string.
- **Section 1 · Store Performance** is a set of tiles (one per revenue type:
  Service / Accessory / Repair / Total / Gross Profit) exactly like the old
  page: status pill (In Progress / Completed with the green "done" tint), a
  2×2 metric grid (Target · Achievement · Achievement % · Balance, with the
  signed balance shown like `−₹1,40,068` when a target is exceeded), the
  cyan Per-Day Requirement strip (reuses the existing per-bucket
  `requiredPerDay`; Completed tiles show the excess) and the gradient
  progress bar.
- **Section 7 date-wise** achievement numbers now sit inside **oval pills**
  (mint `#63e6b7` with `#12352d` text when the day meets its target, salmon
  `#ff9da7` with `#5f2028` text when it misses), matching the old page's
  pill style. The met/missed condition and the small `/ daily target` line
  are unchanged.

## Latest changes (post screenshots)

- **Avg Invoice Value tile removed** from the top card row.
- **Fonts +2pt page-wide** — the whole stylesheet is rem-based, so the root
  font-size went from 16px to 18px, scaling everything evenly.
- **Tab slider**: the Business / Process / Performance chooser is now a wide
  centered pill bar under the title (`.tabs-row`), with bigger buttons.
- **Section 5 "Approx Value" → "Approx Achievable Value"**: balance count ×
  the row-15 unit value (what can still be won), with the small line showing
  `balance N × ₹unit`; "target covered" when the balance is 0.
- The **"SchemeCode starts with DISC" / "SubCategory contains AMC"** helper
  notes are hidden from Section 6 (the cards remain).
- Section 7 **legend cards (per-person daily targets)** are darker and
  centre-aligned like the person headers in the date-wise table itself.
- Section 7 now has a **Gross Profit column per salesperson** (achieved pill
  + per-day target = balance ÷ remaining working days, same rule as the
  revenue types). Person objects expose `grossProfit` bucket +
  `dailyTargets.grossProfit`.
- **Job Management**: rows are de-duplicated on **unique CaseId** *before*
  the open-status filter (latest CaseDate wins, ties → later row; the merged
  count is noted in the section subtitle). New columns **Warranty Type**
  (`service!GSXWarrantyStatusDesc`) and **Repair Type**
  (`service!RepairDocType`) per case.
- **New third tab 📈 Performance** — repair-type analytics over unique case
  ids (`service!RepairDocType` + `service!CaseDate`):
  - one **tile per repair type** with this month's count and the prev-3-
    months monthly average + delta;
  - a **comparison table** — this month vs avg/month over the 3 previous
    months, Δ %, and **average days from CaseDate →
    Ready_For_Pickup_Date_CRM** for both windows;
  - a **date × repair-type pivot** for the current month (newest first,
    totals row).

### Job Management — mapping guarantees

- Column resolution is now **pattern-major with exact-name tiers**: every
  service/advance/comm field starts with an exact `/^CanonicalName$/i`
  pattern (`CaseId`, `MobileNo`, `JobStatus`, `GSXRepairStatusDescription`,
  `CRM_JobStatusDesc`, `GSXWarrantyStatusDesc`, `RepairDocType`,
  `Ready_For_Pickup_Date_CRM`, `AvailableBalance`, `CreatedDate`, …), so the
  canonical headers always win even if the sheet carries extra similarly
  named columns or reorders them. Looser fallbacks remain for custom sheets.
- The Job Management section prints a small **mapping line**
  (`Case ID ← CaseId · Mobile ← MobileNo · …`) resolved from the live sheet
  headers — any column that could not be found shows as red `not found`
  instead of silently rendering dashes.
- Comm. updated rule restated: green when the latest `comm!CreatedDate` for
  the case is within 48 h (COMM_STALE_DAYS = 2) of now, red otherwise or
  when missing. Only `JobStatus` containing "open" is listed, on unique
  `service!CaseId` rows.

## October refresh — what changed & new sections

**Job Management (Process page)**

- Rows are **clubbed by CRM repair status** — bucket headers show
  `N open cases`, sorted alphabetically, oldest first inside each bucket.
- **Service Non-Repair + "ready for pickup"** ⇒ the GSX cell displays
  `Completed` (badge) even if the raw GSX status says otherwise.
- Non-SNR cases with **no repair-created date** and a **job-created date
  older than 2 days** — the whole row is highlighted rose.
- **Out Of Warranty (No Coverage)** warranty (`service!CoverageCodeDescription`)
  with **no advance** matched by case id OR customer mobile — highlighted violet.
- Warranty Type column now reads **service!CoverageCodeDescription**
  (was `GSXWarrantyStatusDesc`), with fallback to the old names.
- Advance is matched BOTH by `advance!MobileNo` **and** `advance!CaseId`
  (`advance!AvailableBalance` is summed per case).

**Performance page (repair analytics)** — "Service Non-Repair" excluded everywhere.

- Round **whole numbers only**; every tile shows `avg N/month · prev 3 months`
  and a delta pill vs the 3-month monthly average.
- Two new dashboards: **🧰 Product Performance** (families: every iPhone
  variant incl. the bare "VIN" export → iPhone; iPad/AirPods/Watch/Mac;
  Accessory–Non-serialized; else Other) and **🛡 Coverage Performance**
  (grouped by `service!CoverageCodeDescription`). Both keep the
  3-month-average comparison table with rounded average days-to-ready.
- "Repair Types by Date" replaced by **🗓 Repair Types — Till Date**: one
  composition row per type for the running month (no per-date rows).

**Process page**

- New **🎯 period summary** at the top: title from `target!F1`, grid rendered
  verbatim from `target!G1:P7`, plus a repairs comparison strip
  (this month vs 3-month average, Service Non-Repair excluded).
- Walk-in Analysis summary is now **one table** (was a five-tile row).

**Business page**

- Section 7 GP fix: gross profit now comes from `crm!GrossProfit` per
  `crm!InvoiceDate` and `crm!SalesPersonName` (cells were rendering ₹0).
- New **§7 · Weekly Revenue Analysis** above the date-wise matrix
  (date-wise became §8): weeks of the current month, per revenue type
  achieved vs weekly target (= month target × week working days ÷ total
  working days) plus balance and a month totals row.
- **Store Total column removed** from the date-wise matrix; tiles across
  all pages now fit any screen width via responsive grids, and pivoted
  numbers are right-aligned.

## October refresh v2 — layout corrections (owner-specified)

- **Salesperson roster = target!C1:J1 exactly.** Only those names get
  performance cards / date-wise columns; empty cells are skipped; CRM-only
  names (store codes, C-SAT labels such as "Very Satisfied") fall off the
  page but still count toward store totals.
- **Store KPI re-anchored:** KPI names fixed at **target!B14:L14**, target
  numbers at **B15:L15**, approx unit values at **B16:L16**, AR/RR/SR type
  row at **B13:L13** — TYPE / TARGET / ACHIEVED / % / BALANCE all resolve
  again (no more collapsed dashes).
- **C-SAT block:** period title from **target!K1**, grid from **target!L1:U5**
  — never merged with salesperson data. Rendered as summary chips
  (responses, satisfied %, C-SAT, CFI, dissatisfied %) plus a cleaned table
  (junk columns dropped, TOTAL row verified against the people rows).
- **Repair Validation = one combined table:** product families vertical
  (iPhone incl. "VIN" exports, iPad, AirPods, Watch, Mac,
  Accessory–Non-serialized, Other) × coverage codes horizontal
  (CoverageCodeDescription), counts per intersection, row/column totals and
  the previous-3-month monthly average strip. Service Non-Repair excluded.
- Every section carries a one-line **Definition** of what it shows and where
  the numbers come from.

---

## WhatsApp Business integration (OpenWA) — Task 28

The dashboard now speaks WhatsApp both ways via the open-source
[rmyndharis/OpenWA](https://github.com/rmyndharis/OpenWA) API gateway
(plugins: [rmyndharis/OpenWA-plugins](https://github.com/rmyndharis/OpenWA-plugins)).

### What you get

| Surface | Behavior |
| --- | --- |
| Chat widget **✆ button** | Sends the last bot answer to any WhatsApp number via the gateway — or, when the gateway is not configured yet, falls back to opening a `wa.me` deep link with the answer prefilled. The number is remembered in `localStorage` (`wa_chatid`). |
| `sendWhatsAppAnswer({q, chat})` | Server function: runs the sheet-search chat bot on `q`, composes a WhatsApp-friendly multiline reply, and POSTs it via OpenWA. |
| `doPost` webhook | Point OpenWA's message webhook at the deployed web app and every inbound DM (groups & `fromMe` filtered out) is automatically answered with the same sheet-bot search. Full WhatsApp chatbot, zero extra code. |
| Phone handling | Any free-form number is normalised to a WhatsApp JID: 10 digits ⇒ `91…` prefix, digits ⇒ `@c.us`; JIDs pass through untouched. |

### Configuration (Script Properties — no secrets in code)

Project Settings → Script Properties:

| Property | Example | Purpose |
| --- | --- | --- |
| `WA_GATEWAY_URL` | `https://wa.example.com` | Your OpenWA host (`docker compose up -d`, see its README; API on port 2785). |
| `WA_API_KEY` | `…` | Admin/operator API key from the OpenWA dashboard. |
| `WA_SESSION` | `sales` | Session id to send with (default `default`). |
| `WA_CHATID` | `919791994639@c.us` or `919791994639` | Default chat for page sends with no number given. |

Webhook: in the OpenWA dashboard (or via
`POST /api/sessions/{id}/webhooks`), register the web-app `/exec` URL with
events `["message.received"]`. Apps Script accepts anonymous `doPost` when
deployed with "Execute as: Me / access: Anyone".

### Plugins worth enabling on the OpenWA side

- `http-action` — expose the spreadsheet bot to WhatsApp users **without**
  this Apps Script (runs inside OpenWA's plugin worker).
- `after-hours` — polite auto-reply when the store is closed.
- `gsheets-logger` — mirror the WhatsApp conversation into a dedicated tab
  (keeps the owner's source sheet untouched).

All three need only `messages:send` and live config.

### Test run

- **Unit/perm harness**: `node google-apps-script/tests/node/perftest.cjs`
  → `== TASK 28 batch ==` 10 checks (offline `wa.me` fallback, JID
  normalisation, gateway POST shape, default chat, webhook answer incl.
  cross-tab matches, `fromMe`/group/other-event filtering, silent when
  unconfigured, chat-widget wiring).
- **Live smoke** (spins up a mock OpenWA on :29785 and pushes real HTTP):
  `node google-apps-script/tests/node/wa-live-test.cjs`
  → verifies send-text exchanges, default chat, webhook auto-reply, and that
  a wrong API key surfaces `401` cleanly.

```text
[1] session=sales chatId=919988776655@c.us  "…" (5 lines)
[2] session=sales chatId=919791994639@c.us  "🧾 Sales on 23-09-2026" (5 lines)
[3] session=sales chatId=919988776655@c.us  "🔧 Part: display" (7 lines)
LIVE SMOKE RUN: PASS
```

### macOS walkthrough (Apple MacBook — M1/M2/M3 or Intel)

Same integration as above, spelled out for a Mac:

**1. Install what you need (Terminal, ~5 min)**
```bash
# package manager (if you don't have brew yet)
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"

# Docker Desktop — the easiest way to run OpenWA
brew install --cask docker
open /Applications/Docker.app      # wait for the whale icon to say "Docker Desktop is running"

# cloudflared — gives the gateway a public HTTPS URL for Apps Script
brew install cloudflared
```

**2. Start the OpenWA gateway**
```bash
git clone https://github.com/rmyndharis/OpenWA.git
cd OpenWA
docker compose -f docker-compose.dev.yml up -d
# dashboard + API now at http://localhost:2785 (Swagger at /api/docs)

# grab the admin key
docker exec openwa-api cat /app/data/.api-key
```
OpenWA's published image is multi-arch (`linux/arm64` included), so Apple Silicon pulls the native build — no Rosetta needed.

**3. Link your WhatsApp Business number**
1. Open `http://localhost:2785` → sign in with the admin key.
2. Sessions → New → name `sales` → Start.
3. Open the session's **QR** screen → in the WhatsApp Business app →
   *Settings → Linked Devices → Link a Device* → scan.
4. Status turns **CONNECTED**.

**4. Give the gateway a public HTTPS URL (so Google Apps Script can call it)**
```bash
# quick test URL (changes each run — fine for a trial):
cloudflared tunnel --url http://localhost:2785
# keep this Terminal window open; it prints https://random-words.trycloudflare.com
```
For a permanent address, create a named tunnel:
```bash
cloudflared tunnel login            # browser opens; pick your domain
cloudflared tunnel create openwa
# then bind a hostname (e.g. wa.yourdomain.com) per the dashboard prompts
cloudflared tunnel --url http://localhost:2785    # replaced by the named-tunnel 'run' command once configured
```

**5. Configure Apps Script (no code change)**
Project Settings → Script Properties:

| Property | Example |
| --- | --- |
| `WA_GATEWAY_URL` | `https://random-words.trycloudflare.com` |
| `WA_API_KEY` | admin key from step 2 |
| `WA_SESSION` | `sales` |
| `WA_CHATID` | `919876543210` |

Then **Deploy → New deployment → Web app** (Execute as: Me · Access: Anyone) and copy the `/exec` URL.

**6. Test**
- Outbound: ask the sheet bot a question, tap the green **✆**, enter a number → message arrives.
- Inbound: in the OpenWA dashboard add a webhook on your `/exec` URL with event `message.received` → WhatsApp a question to the linked number → it auto-replies from the sheets.

**7. Keep it running overnight**
```bash
caffeinate -t 43200     # keep the Mac awake 12 h while you leave it unattended
```
Docker Desktop will restart the container on reboot if you enable "Start Docker Desktop when you log in" (Docker Desktop → Settings → General).
