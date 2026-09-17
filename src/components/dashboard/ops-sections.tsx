import { cn } from "@/lib/utils";
import { useState, type ReactNode } from "react";
import {
  Panel,
  PanelHead,
  Section,
  Stat,
  money,
  moneyCompact,
} from "@/components/dashboard/sections";
import type { OpsResult } from "@/lib/ops";

/* ------------------------------------------------------------------ */
/* Shared                                                              */
/* ------------------------------------------------------------------ */

const num = (n: number) => Math.round(n).toLocaleString("en-IN");
const one = (n: number) => n.toFixed(1);

function Table({
  head,
  rows,
  footer,
  right = [],
  minWidth = 720,
}: {
  head: ReactNode[];
  rows: ReactNode[][];
  footer?: ReactNode[];
  /** column indexes that hold numbers (right aligned) */
  right?: number[];
  minWidth?: number;
}) {
  const isRight = (i: number) => right.includes(i);
  return (
    <div className="overflow-x-auto">
      <table className="w-full text-sm" style={{ minWidth }}>
        <thead>
          <tr className="border-b border-border/70 text-[10px] font-semibold uppercase tracking-[0.12em] text-muted-foreground">
            {head.map((h, i) => (
              <th
                key={i}
                className={cn(
                  "whitespace-nowrap px-3 py-2",
                  isRight(i) ? "text-right" : "text-left",
                )}
              >
                {h}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((r, i) => (
            <tr
              key={i}
              className="border-b border-border/40 transition-colors last:border-0 hover:bg-accent/30"
            >
              {r.map((c, j) => (
                <td
                  key={j}
                  className={cn(
                    "px-3 py-2 align-top",
                    isRight(j) ? "text-right tabular-nums" : "text-left",
                  )}
                >
                  {c}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
        {footer ? (
          <tfoot>
            <tr className="border-t border-border/70 bg-accent/20 font-semibold">
              {footer.map((f, j) => (
                <td
                  key={j}
                  className={cn("px-3 py-2", isRight(j) ? "text-right tabular-nums" : "text-left")}
                >
                  {f}
                </td>
              ))}
            </tr>
          </tfoot>
        ) : null}
      </table>
    </div>
  );
}

function Delta({ value, suffix = "" }: { value: number; suffix?: string }) {
  const positive = value >= 0;
  return (
    <span
      className={cn(
        "font-mono text-[11px] font-semibold tabular-nums",
        positive ? "text-positive" : "text-negative",
      )}
    >
      {positive ? "+" : ""}
      {one(value)}
      {suffix}
    </span>
  );
}

const AGE_TONES = [
  "text-positive",
  "text-primary",
  "text-foreground",
  "text-warning",
  "text-warning",
  "text-negative font-bold",
];

const BUCKET_TONE: Record<string, string> = {
  "unclaimed device": "bg-warning/10",
  diagnosis: "bg-primary/5",
  "ready for pickup": "bg-positive/5",
  "repair in process": "bg-accent/20",
};

/* ------------------------------------------------------------------ */
/* PROCESS VIEW                                                        */
/* ------------------------------------------------------------------ */

export function ProcessView({
  process,
  googleReview,
  uniqueInvoices,
}: {
  process: OpsResult["process"];
  googleReview: { previousMonth: number; currentMonth: number; difference: number };
  uniqueInvoices: number;
}) {
  const [showAllDates, setShowAllDates] = useState(false);
  const [showUnclaimed, setShowUnclaimed] = useState(false);

  const s = process.summary;
  const daily = showAllDates ? process.daily : process.daily.slice(-3);
  const openJobs = showUnclaimed
    ? process.openJobs
    : process.openJobs.filter((j) => !j.unclaimed);
  const reviewRate = uniqueInvoices > 0 ? (googleReview.currentMonth / uniqueInvoices) * 100 : 0;

  const dailyHead = [
    "Date",
    "Total Walk-ins",
    "Cases Created",
    "Invoices Created",
    "Communication Updated",
  ];
  const dailyRight = [1, 2, 3, 4];
  const dailyRows: ReactNode[][] = daily.map((d) => [
    d.date,
    num(d.total),
    num(d.cases),
    num(d.invoices),
    num(d.communications),
  ]);
  if (!dailyRows.length) {
    dailyRows.push([
      <span key="none" className="text-muted-foreground">
        No walk-in records for {process.month}
      </span>,
      "—",
      "—",
      "—",
      "—",
    ]);
  }

  const repair = process.repair;
  const ageing = process.ageing;

  return (
    <div className="space-y-8">
      <Section
        title={`Walk-in Process — ${process.month}`}
        subtitle="Walk-in flow, case creation, invoice creation and open service orders"
        right={
          <span className="font-mono text-[11px] text-muted-foreground">
            walkin · aod · comm
          </span>
        }
      >
        <div className="grid grid-cols-2 gap-3 sm:grid-cols-3 lg:grid-cols-7">
          <Stat label="Total Walk-ins" value={num(s.total)} sub="tokens this month" />
          <Stat
            label="Unique Tokens"
            value={num(s.unique)}
            sub="unique mobile numbers"
            tone="primary"
          />
          <Stat label="Cases Created" value={num(s.jobs)} sub="unique case ids" tone="positive" />
          <Stat label="Invoices Created" value={num(s.invoices)} sub="walk-in invoices" />
          <Stat
            label="Comm Updates"
            value={num(s.communications)}
            sub="customer updates logged"
          />
          <Stat
            label="Open Jobs"
            value={num(s.openJobs)}
            sub="aod · JobStatus = Open"
            tone="warning"
          />
          <Stat
            label="Unclaimed"
            value={num(s.unclaimed)}
            sub="no delivery challan"
            tone="negative"
          />
        </div>
      </Section>

      <Section
        title="Date-wise Walk-in Process"
        subtitle="One line per date — walk-ins, cases and invoices created"
        right={
          process.daily.length > 3 ? (
            <button
              type="button"
              onClick={() => setShowAllDates((v) => !v)}
              className="rounded border border-border/70 px-2.5 py-1 font-mono text-[11px] text-muted-foreground transition-colors hover:bg-accent hover:text-foreground"
            >
              {showAllDates ? "Latest 3 dates" : `All ${process.daily.length} dates`}
            </button>
          ) : null
        }
      >
        <Panel>
          <Table head={dailyHead} rows={dailyRows} right={dailyRight} minWidth={660} />
        </Panel>
      </Section>

      <Section
        title="Google Reviews"
        subtitle="Reviews earned this month against the store's unique invoice base"
        right={<span className="font-mono text-[11px] text-muted-foreground">target!G1:I3</span>}
      >
        <div className="grid grid-cols-2 gap-3 lg:grid-cols-4">
          <Stat label="Last Month" value={num(googleReview.previousMonth)} sub="reviews" />
          <Stat
            label="Current Month"
            value={num(googleReview.currentMonth)}
            sub="reviews"
            tone="primary"
          />
          <Stat
            label="Difference"
            value={`${googleReview.difference >= 0 ? "+" : ""}${num(googleReview.difference)}`}
            sub="current − last month"
            tone={googleReview.difference >= 0 ? "positive" : "negative"}
          />
          <Stat
            label="GR / Invoice"
            value={`${one(reviewRate)}%`}
            sub={`${num(googleReview.currentMonth)} reviews ÷ ${num(uniqueInvoices)} invoices`}
          />
        </div>
      </Section>

      <Section
        title="Repair Type Performance"
        subtitle={`Completed service orders (AOD · RFPU / ready-for-pickup) · ${repair.currentMonth} vs previous 3-month average`}
        right={
          <span className="font-mono text-[11px] text-muted-foreground">
            {repair.previousMonths.join(" · ")}
          </span>
        }
      >
        {!repair.available ? (
          <Panel className="p-4 text-sm text-muted-foreground">
            No completed repair orders found for {repair.currentMonth}.
          </Panel>
        ) : (
          <div className="space-y-3">
            <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-4">
              <Panel className="p-3">
                <p className="text-[10px] font-semibold uppercase tracking-[0.14em] text-muted-foreground">
                  Total completed
                </p>
                <p className="mt-1 font-display text-2xl font-bold tabular-nums">
                  {num(repair.totalCurrent)}
                </p>
                <p className="mt-0.5 font-mono text-[11px] text-muted-foreground">
                  3M avg {one(repair.previousThreeAverage)} ·{" "}
                  <Delta value={repair.difference} /> · {one(repair.percentVsAverage)}%
                </p>
              </Panel>
              {repair.typeSummary.slice(0, 3).map((t) => (
                <Panel key={t.repairType} className="p-3">
                  <p className="truncate text-[10px] font-semibold uppercase tracking-[0.14em] text-muted-foreground">
                    {t.repairType}
                  </p>
                  <p className="mt-1 font-display text-2xl font-bold tabular-nums">
                    {num(t.current)}
                  </p>
                  <p className="mt-0.5 font-mono text-[11px] text-muted-foreground">
                    3M avg {one(t.previousAverage)} · <Delta value={t.difference} /> ·{" "}
                    {one(t.percentVsAverage)}%
                  </p>
                  <div className="mt-2 flex gap-1 font-mono text-[10px] text-muted-foreground">
                    {t.previousThree.map((v, i) => (
                      <span key={i} className="rounded border border-border/60 px-1.5 py-0.5">
                        {repair.previousMonths[i]?.slice(0, 3)} {v}
                      </span>
                    ))}
                  </div>
                </Panel>
              ))}
            </div>
            <Panel>
              <PanelHead
                title="Repair type × coverage × product"
                hint="Top rows for the current month against the 3-month average"
              />
              <Table
                head={[
                  "#",
                  "Repair Type",
                  "Coverage",
                  "Product",
                  "Current",
                  "3M Avg",
                  "Diff",
                  "Vs Avg",
                ]}
                right={[0, 4, 5, 6, 7]}
                minWidth={860}
                rows={repair.detail.map((d, i) => [
                  num(i + 1),
                  d.repairType,
                  <span key="c" className="text-muted-foreground">
                    {d.coverage}
                  </span>,
                  d.product,
                  <span key="v" className="font-semibold">
                    {num(d.current)}
                  </span>,
                  one(d.previousAverage),
                  <Delta key="d" value={d.difference} />,
                  `${one(d.percentVsAverage)}%`,
                ])}
              />
            </Panel>
          </div>
        )}
      </Section>

      <Section
        title="Open Jobs — Device Bucket & Ageing"
        subtitle="Unique open Case ID count · status bucket × ageing slab"
        right={
          <span className="font-mono text-[11px] text-muted-foreground">
            aod · JobAge (days)
          </span>
        }
      >
        <Panel>
          <Table
            minWidth={760}
            head={["CRM_JobStatusDesc", ...ageing.ages, "Total"]}
            right={ageing.ages.map((_, i) => i + 1).concat(ageing.ages.length + 1)}
            rows={ageing.rows.map((r) => [
              <span key="b" className="font-medium">
                {r.bucket}
              </span>,
              ...r.cells.map((v, i) => (
                <span key={i} className={v ? AGE_TONES[i] : "text-muted-foreground/50"}>
                  {v ? num(v) : "·"}
                </span>
              )),
              <span key="t" className="font-semibold">
                {num(r.total)}
              </span>,
            ])}
            footer={[
              "Grand Total",
              ...ageing.totals.map((v, i) => (
                <span key={i} className={v ? AGE_TONES[i] : "text-muted-foreground/50"}>
                  {v ? num(v) : "·"}
                </span>
              )),
              num(ageing.grandTotal),
            ]}
          />
        </Panel>
      </Section>

      <Section
        title="Open Jobs"
        subtitle={`${num(process.openJobs.length)} unique open cases · latest AOD state with communication history`}
        right={
          process.summary.unclaimed > 0 ? (
            <button
              type="button"
              onClick={() => setShowUnclaimed((v) => !v)}
              className="rounded border border-border/70 px-2.5 py-1 font-mono text-[11px] text-muted-foreground transition-colors hover:bg-accent hover:text-foreground"
            >
              {showUnclaimed
                ? `Hide ${process.summary.unclaimed} unclaimed`
                : `Show ${process.summary.unclaimed} unclaimed`}
            </button>
          ) : null
        }
      >
        <Panel>
          <Table
            minWidth={1180}
            head={[
              "Case ID",
              "Customer",
              "Mobile",
              "Product",
              "Coverage",
              "Repair Type",
              "Case Date",
              "Case Ageing",
              "Bucket",
              "Last Comm",
              "Comm Ageing",
            ]}
            rows={openJobs.map((j) => [
              <span key="id" className="font-mono text-[12px]">
                {j.caseId}
              </span>,
              j.customer,
              j.mobile,
              j.product,
              <span key="cov" className="text-muted-foreground">
                {j.coverage}
              </span>,
              j.repairType,
              j.caseDate,
              <span key="age" className="text-warning">
                {j.caseAge}
              </span>,
              <span
                key="bucket"
                className={cn(
                  "rounded px-1.5 py-0.5",
                  BUCKET_TONE[j.bucket.trim().toLowerCase()] ?? "",
                )}
              >
                {j.bucket}
              </span>,
              j.lastCommunication,
              j.communicationAging,
            ])}
          />
        </Panel>
      </Section>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* PRODUCTIVITY VIEW                                                   */
/* ------------------------------------------------------------------ */

export function ProductivityView({ productivity }: { productivity: OpsResult["productivity"] }) {
  const s = productivity.summary;
  const maxWalkins = Math.max(1, ...productivity.daily.map((d) => d.walkins));
  const maxJobs = Math.max(1, ...productivity.daily.map((d) => d.jobs));
  const maxInvoices = Math.max(1, ...productivity.daily.map((d) => d.invoices));

  const bar = (value: number, max: number, tone: string) => (
    <div className="flex items-center gap-2">
      <div className="h-1.5 w-full min-w-16 overflow-hidden rounded-full bg-border/40">
        <div className={cn("h-full rounded-full", tone)} style={{ width: `${(value / max) * 100}%` }} />
      </div>
      <span className="w-7 text-right font-mono text-[11px] tabular-nums">{value}</span>
    </div>
  );

  const funnel = [
    { label: "Walk-ins", value: s.walkins, tone: "bg-chart-1" },
    { label: "Jobs Created", value: s.jobs, tone: "bg-positive" },
    { label: "Invoices", value: s.invoices, tone: "bg-chart-2" },
  ];
  const funnelMax = Math.max(1, s.walkins);

  return (
    <div className="space-y-8">
      <Section
        title={`Productivity & Conversion — ${productivity.month}`}
        subtitle="Customer flow, conversion efficiency and technician turnaround"
        right={
          <span className="font-mono text-[11px] text-muted-foreground">
            {productivity.daily.length} active days
          </span>
        }
      >
        <div className="grid grid-cols-2 gap-3 sm:grid-cols-3 lg:grid-cols-4">
          <Stat
            label="Walk-ins / Day"
            value={one(s.averageWalkinsPerDay)}
            sub={`${num(s.walkins)} walk-ins · ${num(s.uniqueWalkins)} unique`}
          />
          <Stat label="Jobs / Day" value={one(s.averageJobsPerDay)} sub={`${num(s.jobs)} cases`} tone="positive" />
          <Stat
            label="Invoices / Day"
            value={one(s.averageInvoicesPerDay)}
            sub={`${num(s.invoices)} invoices · ${moneyCompact(s.revenue)}`}
            tone="primary"
          />
          <Stat label="Job Conversion" value={`${one(s.walkinToJob)}%`} sub="unique walk-in → case" />
          <Stat
            label="Invoice Conversion"
            value={`${one(s.walkinToInvoice)}%`}
            sub="unique walk-in → invoice"
          />
          <Stat
            label="Customer Handling"
            value={`${num(s.handlingMinutes)} min`}
            sub={`walk-in → case creation · ${num(s.handlingSample)} samples`}
            tone="warning"
          />
          <Stat
            label="Repair TAT"
            value={s.carryInRepairs > 0 ? `${num(s.repairTatMinutes)} min` : "—"}
            sub={`assign → ready for pickup · ${num(s.carryInRepairs)} carry-in cases`}
            tone="negative"
          />
          <Stat
            label="N-Repair TAT"
            value={s.nRepairRepairs > 0 ? `${num(s.nRepairTatMinutes)} min` : "—"}
            sub={`${num(s.nRepairRepairs)} n-repair · ${num(s.totalRepairs)} completed`}
          />
        </div>
      </Section>

      <div className="grid gap-4 lg:grid-cols-2">
        <Section title="Customer Flow Funnel" subtitle="Walk-in → case → invoice">
          <Panel className="space-y-3 p-4">
            {funnel.map((f) => (
              <div key={f.label} className="flex items-center gap-3">
                <span className="w-28 shrink-0 text-xs text-muted-foreground">{f.label}</span>
                <div className="h-2.5 w-full overflow-hidden rounded-full bg-border/40">
                  <div
                    className={cn("h-full rounded-full", f.tone)}
                    style={{ width: `${(f.value / funnelMax) * 100}%` }}
                  />
                </div>
                <span className="w-12 shrink-0 text-right font-mono text-xs tabular-nums">
                  {num(f.value)}
                </span>
              </div>
            ))}
            <div className="flex flex-wrap gap-4 border-t border-border/60 pt-3 font-mono text-[11px] text-muted-foreground">
              <span>
                job conversion{" "}
                <span className="text-foreground">{one(s.walkinToJob)}%</span>
              </span>
              <span>
                invoice conversion{" "}
                <span className="text-foreground">{one(s.walkinToInvoice)}%</span>
              </span>
              <span>
                billed value <span className="text-positive">{money(s.revenue)}</span>
              </span>
            </div>
          </Panel>
        </Section>

        <Section title="Technician Productivity" subtitle="Repairs handled by type and average assign → pickup TAT">
          <Panel>
            <PanelHead
              title="Repair + N-repair"
              hint="Completed cases in the current month"
              right={
                <span className="font-mono text-[11px] text-muted-foreground">
                  {num(s.carryInRepairs)} repair · {num(s.nRepairRepairs)} n-repair
                </span>
              }
            />
            <Table
              minWidth={620}
              head={[
                "Technician",
                "Total",
                "Repair",
                "N-Repair",
                "Repair TAT",
                "N-Repair TAT",
                "Avg TAT",
              ]}
              right={[1, 2, 3, 4, 5, 6]}
              rows={productivity.technicians.map((t) => [
                t.name,
                num(t.total),
                num(t.carryIn),
                num(t.nRepair),
                t.carryInTat,
                t.nRepairTat,
                <span key="a" className="font-semibold">
                  {t.averageTat}
                </span>,
              ])}
            />
          </Panel>
        </Section>
      </div>

      <Section
        title="Daily Productivity"
        subtitle="Daily walk-ins against cases and invoices created"
        right={
          <span className="font-mono text-[11px] text-muted-foreground">
            average per active day
          </span>
        }
      >
        <Panel>
          <Table
            minWidth={900}
            head={[
              "Date",
              "Walk-ins",
              "Jobs Created",
              "Invoices Created",
              "Job Conversion",
              "Invoice Conversion",
            ]}
            right={[4, 5]}
            rows={productivity.daily.map((d) => [
              d.date,
              bar(d.walkins, maxWalkins, "bg-chart-1"),
              bar(d.jobs, maxJobs, "bg-positive"),
              bar(d.invoices, maxInvoices, "bg-chart-2"),
              `${one(d.jobConversion)}%`,
              `${one(d.invoiceConversion)}%`,
            ])}
          />
        </Panel>
      </Section>

      <Section
        title="Salesperson Productivity Ranking"
        subtitle="Walk-in owner → mobile-matched job → invoice → billed revenue"
        right={<span className="font-mono text-[11px] text-muted-foreground">sorted by revenue</span>}
      >
        <Panel>
          <Table
            minWidth={980}
            head={[
              "#",
              "Person",
              "Walk-ins",
              "Unique Walk-ins",
              "Jobs",
              "Job Conversion",
              "Invoices",
              "Invoice Conversion",
              "Invoice Revenue",
            ]}
            right={[0, 2, 3, 4, 5, 6, 7, 8]}
            rows={productivity.people.map((p, i) => [
              num(i + 1),
              <span key="n" className="font-medium">
                {p.name}
              </span>,
              num(p.walkins),
              num(p.uniqueWalkins),
              num(p.jobs),
              `${one(p.walkinToJob)}%`,
              num(p.invoices),
              `${one(p.walkinToInvoice)}%`,
              <span key="r" className="font-semibold text-positive">
                {money(p.invoiceRevenue)}
              </span>,
            ])}
          />
        </Panel>
      </Section>

      <Section
        title="Repair TAT — Case by Case"
        subtitle="Assigned to technician → ready for pickup"
        right={
          <span className="font-mono text-[11px] text-muted-foreground">
            worst first · showing {num(productivity.tatCases.length)} of {num(productivity.tatCaseTotal)}
          </span>
        }
      >
        <Panel>
          <Table
            minWidth={1080}
            head={[
              "#",
              "Case ID",
              "Technician",
              "Repair Type",
              "Customer",
              "Mobile",
              "Assigned",
              "Completed",
              "Source",
              "TAT",
            ]}
            right={[0, 9]}
            rows={productivity.tatCases.map((c, i) => [
              num(i + 1),
              <span key="id" className="font-mono text-[12px]">
                {c.caseId}
              </span>,
              c.technician,
              c.repairType,
              c.customer,
              c.mobile,
              c.assignDate,
              c.completeDate,
              <span key="src" className="text-muted-foreground">
                {c.completeSource}
              </span>,
              <span key="tat" className="font-semibold text-warning">
                {c.tat}
              </span>,
            ])}
          />
        </Panel>
        <p className="mt-3 rounded border border-border/60 bg-accent/20 px-3 py-2 font-mono text-[11px] leading-relaxed text-muted-foreground">
          method · walk-in → job matches the walk-in mobile to the AOD customer
          mobile and takes the closest case created after the visit · TAT uses
          AssignDate → GSX_RFPUDateTime, falling back to the CRM ready-for-pickup
          date when RFPU is empty
        </p>
      </Section>
    </div>
  );
}
