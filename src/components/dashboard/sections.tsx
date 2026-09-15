import { cn } from "@/lib/utils";
import type { ReactNode } from "react";
import { Badge } from "@/components/ui/badge";
import type {
  DashboardData,
  KpiRow,
  SalespersonRow,
  StoreMetrics,
} from "@/lib/metrics";

/* ------------------------------------------------------------------ */
/* Shell                                                               */
/* ------------------------------------------------------------------ */

export function DashShell({ children }: { children: ReactNode }) {
  return (
    <div className="min-h-screen bg-grid bg-background">
      <div className="bg-glow-primary min-h-screen">{children}</div>
    </div>
  );
}

export function Section({
  id,
  title,
  subtitle,
  right,
  children,
}: {
  id?: string;
  title: string;
  subtitle?: string;
  right?: ReactNode;
  children: ReactNode;
}) {
  return (
    <section id={id} className="scroll-mt-24">
      <div className="mb-3 flex flex-wrap items-end justify-between gap-2">
        <div>
          <h2 className="font-display text-sm font-bold uppercase tracking-[0.18em] text-foreground">
            {title}
          </h2>
          {subtitle ? (
            <p className="mt-0.5 text-xs text-muted-foreground">{subtitle}</p>
          ) : null}
        </div>
        {right}
      </div>
      {children}
    </section>
  );
}

export function Panel({
  className,
  children,
}: {
  className?: string;
  children: ReactNode;
}) {
  return (
    <div
      className={cn(
        "rounded-lg border border-border/80 bg-card/60 shadow-[0_1px_0_0_oklch(0.955_0.008_230/0.04)_inset] backdrop-blur-sm",
        className,
      )}
    >
      {children}
    </div>
  );
}

export function PanelHead({
  title,
  hint,
  right,
}: {
  title: string;
  hint?: string;
  right?: ReactNode;
}) {
  return (
    <div className="flex items-center justify-between gap-3 border-b border-border/70 px-4 py-2.5">
      <div className="min-w-0">
        <p className="truncate text-xs font-semibold uppercase tracking-[0.14em] text-muted-foreground">
          {title}
        </p>
        {hint ? (
          <p className="mt-0.5 truncate text-[11px] text-muted-foreground/70">
            {hint}
          </p>
        ) : null}
      </div>
      {right}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Primitives                                                          */
/* ------------------------------------------------------------------ */

export function Stat({
  label,
  value,
  sub,
  tone = "default",
  className,
}: {
  label: string;
  value: ReactNode;
  sub?: ReactNode;
  tone?: "default" | "primary" | "positive" | "warning" | "negative";
  className?: string;
}) {
  const toneCls =
    tone === "primary"
      ? "text-primary"
      : tone === "positive"
        ? "text-positive"
        : tone === "warning"
          ? "text-warning"
          : tone === "negative"
            ? "text-negative"
            : "text-foreground";
  return (
    <div
      className={cn(
        "rounded-md border border-border/60 bg-background/40 px-3 py-2.5",
        className,
      )}
    >
      <p className="text-[10px] font-semibold uppercase tracking-[0.14em] text-muted-foreground">
        {label}
      </p>
      <p className={cn("mt-1 font-display text-xl font-bold tabular-nums", toneCls)}>
        {value}
      </p>
      {sub ? <p className="mt-0.5 text-[11px] text-muted-foreground">{sub}</p> : null}
    </div>
  );
}

export function Progress({
  pct,
  tone,
  className,
}: {
  pct: number;
  tone?: "auto" | "primary";
  className?: string;
}) {
  const clamped = Math.max(0, Math.min(100, pct));
  const auto =
    pct >= 100
      ? "bg-positive"
      : pct >= 70
        ? "bg-primary"
        : pct >= 40
          ? "bg-warning"
          : "bg-negative";
  return (
    <div
      className={cn(
        "h-1.5 w-full overflow-hidden rounded-full bg-border/40",
        className,
      )}
    >
      <div
        className={cn("h-full rounded-full transition-all", tone === "primary" ? "bg-primary" : auto)}
        style={{ width: `${clamped}%` }}
      />
    </div>
  );
}

export function PctBadge({ value }: { value: number }) {
  const tone =
    value >= 100
      ? "border-positive/40 bg-positive/10 text-positive"
      : value >= 70
        ? "border-primary/40 bg-primary/10 text-primary"
        : value >= 40
          ? "border-warning/40 bg-warning/10 text-warning"
          : "border-negative/40 bg-negative/10 text-negative";
  return (
    <span
      className={cn(
        "inline-flex items-center rounded border px-1.5 py-0.5 font-mono text-[11px] font-semibold tabular-nums",
        tone,
      )}
    >
      {value.toFixed(1)}%
    </span>
  );
}

export const money = (n: number) =>
  `₹${Math.round(n).toLocaleString("en-IN")}`;

export function MiniTable({
  head,
  rows,
  footer,
  className,
}: {
  head: ReactNode[];
  rows: ReactNode[][];
  footer?: ReactNode[];
  className?: string;
}) {
  return (
    <div className={cn("overflow-x-auto", className)}>
      <table className="w-full min-w-[560px] text-left text-sm">
        <thead>
          <tr className="border-b border-border/70">
            {head.map((h, i) => (
              <th
                key={i}
                className={cn(
                  "px-3 py-2 text-[10px] font-semibold uppercase tracking-[0.12em] text-muted-foreground",
                  i === 0 ? "text-left" : "text-right",
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
                    "px-3 py-2 tabular-nums",
                    j === 0 ? "text-left font-medium" : "text-right",
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
                  className={cn(
                    "px-3 py-2 tabular-nums",
                    j === 0 ? "text-left" : "text-right",
                  )}
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

/* ------------------------------------------------------------------ */
/* Store performance (Section 1)                                       */
/* ------------------------------------------------------------------ */

interface Bucket {
  target: number;
  achieved: number;
  pct: number;
  balance: number;
  runRate: number;
  requiredPerDay: number;
  projected: number;
}

function RevRow({ label, b }: { label: string; b: Bucket }) {
  return (
    <div className="grid grid-cols-2 items-center gap-3 px-4 py-2.5 sm:grid-cols-[150px_1fr_auto]">
      <div className="text-sm font-medium">{label}</div>
      <div className="col-span-2 space-y-1 sm:col-span-1">
        <Progress pct={b.pct} />
        <div className="flex flex-wrap gap-x-3 gap-y-0.5 font-mono text-[11px] text-muted-foreground">
          <span>
            ach <span className="text-foreground">{money(b.achieved)}</span>
          </span>
          <span>
            target <span className="text-foreground">{money(b.target)}</span>
          </span>
          <span>
            balance <span className="text-foreground">{money(b.balance)}</span>
          </span>
          <span>
            req/day{" "}
            <span className="text-foreground">{money(b.requiredPerDay)}</span>
          </span>
        </div>
      </div>
      <div className="flex items-center justify-end gap-2">
        <span className="hidden font-mono text-[11px] text-muted-foreground sm:inline">
          run {money(b.runRate)}/d
        </span>
        <PctBadge value={b.pct} />
      </div>
    </div>
  );
}

export function StorePerformanceSection({
  data,
}: {
  data: DashboardData;
}) {
  const s: StoreMetrics = data.store;
  const buckets: Array<[string, Bucket]> = [
    ["Total Revenue", s.total],
    ["Service", s.service],
    ["Accessories", s.accessory],
    ["Repair", s.repair],
    ["Gross Profit", s.grossProfit],
  ];

  return (
    <Panel>
      <PanelHead
        title="Store Performance"
        hint={`${s.monthProgressLabel} · ${s.daysRemaining} days remaining · ${s.totalInvoices} invoice lines / ${s.uniqueInvoices} unique invoices`}
        right={
          <span className="font-mono text-[11px] text-muted-foreground">
            projected {money(s.total.projected)}
          </span>
        }
      />
      <div className="divide-y divide-border/40">
        {buckets.map(([label, b]) => (
          <RevRow key={label} label={label} b={b} />
        ))}
      </div>
      <div className="grid grid-cols-2 gap-2 border-t border-border/70 p-3 sm:grid-cols-5">
        <Stat label="Total Invoices" value={s.totalInvoices} />
        <Stat label="Unique Invoices" value={s.uniqueInvoices} tone="primary" />
        <Stat label="Revenue Rate" value={`${s.total.pct.toFixed(1)}%`} tone={s.total.pct >= 100 ? "positive" : "default"} />
        <Stat label="Run Rate / Day" value={money(s.total.runRate)} />
        <Stat label="Req / Day" value={money(s.total.requiredPerDay)} sub="to close balance" />
      </div>
    </Panel>
  );
}

/* ------------------------------------------------------------------ */
/* Salesperson performance (Section 2)                                 */
/* ------------------------------------------------------------------ */

function PersonCard({ row }: { row: SalespersonRow }) {
  const rows: Array<[string, SalespersonRow["service"]]> = [
    ["Service", row.service],
    ["Accessory", row.accessory],
    ["Repair", row.repair],
  ];
  return (
    <Panel className="overflow-hidden">
      <PanelHead
        title={row.name}
        hint={`${row.invoiceCount} invoices`}
        right={<PctBadge value={row.total.pct} />}
      />
      <div className="p-3">
        <div className="mb-3 grid grid-cols-2 gap-2">
          <Stat
            label="Total Revenue"
            value={money(row.total.achieved)}
            sub={`of ${money(row.total.target)}`}
          />
          <Stat
            label="Total Incentive"
            value={money(row.totalIncentive)}
            sub="service + accessory"
            tone="positive"
          />
        </div>
        <div className="space-y-2">
          {rows.map(([label, b]) => (
            <div key={label}>
              <div className="flex items-center justify-between text-xs">
                <span className="font-medium">{label}</span>
                <span className="font-mono text-muted-foreground">
                  {money(b.achieved)} / {money(b.target)} ·{" "}
                  <span
                    className={cn(
                      b.pct >= 100
                        ? "text-positive"
                        : b.pct >= 70
                          ? "text-primary"
                          : b.pct >= 40
                            ? "text-warning"
                            : "text-negative",
                    )}
                  >
                    {b.pct.toFixed(1)}%
                  </span>
                </span>
              </div>
              <Progress pct={b.pct} className="mt-1" />
              <div className="mt-0.5 flex justify-between font-mono text-[10px] text-muted-foreground">
                <span>balance {money(b.balance)}</span>
                <span>
                  eligible{" "}
                  <span className="text-positive">
                    {b.incentive > 0 ? money(b.incentive) : "—"}
                  </span>
                </span>
              </div>
            </div>
          ))}
        </div>
      </div>
    </Panel>
  );
}

export function SalespersonSection({ data }: { data: DashboardData }) {
  return (
    <div className="grid gap-4 lg:grid-cols-2">
      {data.salespeople.map((row) => (
        <PersonCard key={row.key} row={row} />
      ))}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* AMC & Bundle incentives (Section 3)                                 */
/* ------------------------------------------------------------------ */

export function AmcBundleSection({ data }: { data: DashboardData }) {
  const { amc, bundle, target } = data;
  const slabLabels = (slabs: { label: string; value?: number; perBundle?: number }[]) =>
    slabs
      .map((s) => `${s.label}${s.value != null ? `: ₹${s.value}` : s.perBundle != null ? `: ₹${s.perBundle}` : ""}`)
      .join("  ·  ");

  return (
    <div className="space-y-4">
      <div className="grid gap-4 lg:grid-cols-2">
        {amc.map((row) => (
          <Panel key={row.key}>
            <PanelHead
              title={`${row.name} · AMC`}
              hint={`achieved ${row.amcCount} AMCs`}
              right={
                <span className="font-mono text-[11px] text-positive">
                  {money(row.projectedIncentive)}
                </span>
              }
            />
            <div className="grid grid-cols-2 gap-2 p-3">
              <Stat
                label="AMC Count"
                value={row.amcCount}
                sub={row.amcNextSlab ? `slab ${row.amcNextSlab.label} @ ₹${row.amcNextSlab.rate}/AMC` : "no slab yet"}
              />
              <Stat
                label="To Next Slab"
                value={row.amcRemaining}
                sub={row.amcNextSlab ? `${row.amcRemaining} more to close ${row.amcNextSlab.label}` : "—"}
                tone={row.amcRemaining === 0 ? "positive" : "warning"}
              />
              <Stat
                label="Accessory AMC"
                value={row.accessoryCount}
                sub={row.accessoryNextSlab ? `fixed ₹${row.accessoryNextSlab.accFixed ?? 0}` : "below slab"}
              />
              <Stat
                label="Accessory Left"
                value={row.accessoryRemaining}
                sub="to close current tier"
              />
            </div>
          </Panel>
        ))}
      </div>
      <div className="grid gap-4 lg:grid-cols-2">
        {bundle.map((row) => (
          <Panel key={row.key}>
            <PanelHead
              title={`${row.name} · Bundle`}
              hint={`${row.count} DISC invoices`}
              right={
                <span className="font-mono text-[11px] text-positive">
                  {money(row.projectedIncentive)}
                </span>
              }
            />
            <div className="grid grid-cols-2 gap-2 p-3">
              <Stat
                label="Bundle Count"
                value={row.count}
                sub={row.nextSlab ? `next slab at ${row.nextSlab.min}` : "top slab"}
              />
              <Stat
                label="To Next Slab"
                value={row.remaining}
                sub={
                  row.nextSlab
                    ? `${row.remaining} more → ₹${row.nextSlab.perBundle}/bundle`
                    : "—"
                }
                tone={row.remaining === 0 ? "positive" : "warning"}
              />
            </div>
          </Panel>
        ))}
      </div>
      <Panel className="p-3">
        <p className="text-[10px] font-semibold uppercase tracking-[0.14em] text-muted-foreground">
          Plan reference
        </p>
        <p className="mt-1 font-mono text-[11px] leading-relaxed text-muted-foreground">
          {target.amcSlabs.map((s) => `${s.label}: ${s.rateLabel}`).join("  ·  ")}
        </p>
        <p className="mt-1.5 font-mono text-[11px] leading-relaxed text-muted-foreground">
          {slabLabels(target.bundleSlabs)}
        </p>
      </Panel>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* KPI grid (Section 4)                                                */
/* ------------------------------------------------------------------ */

function KpiCard({ row }: { row: KpiRow }) {
  return (
    <div className="rounded-md border border-border/60 bg-background/40 p-3">
      <div className="flex items-start justify-between gap-2">
        <div>
          <p className="text-sm font-semibold">{row.label}</p>
          <p className="font-mono text-[10px] text-muted-foreground">
            value ₹{row.value.toLocaleString("en-IN")} · approx {money(row.approxAmount)}
          </p>
        </div>
        <PctBadge value={row.pct} />
      </div>
      <Progress pct={row.pct} className="mt-2" />
      <div className="mt-2 grid grid-cols-4 gap-1.5 font-mono text-[10px]">
        <div>
          <p className="text-muted-foreground">TGT</p>
          <p className="text-foreground">{row.target}</p>
        </div>
        <div>
          <p className="text-muted-foreground">ACH</p>
          <p className="text-foreground">{row.achieved}</p>
        </div>
        <div>
          <p className="text-muted-foreground">BAL</p>
          <p className={row.balance === 0 ? "text-positive" : "text-warning"}>
            {row.balance}
          </p>
        </div>
        <div>
          <p className="text-muted-foreground">VALUE</p>
          <p className="text-foreground">{money(row.achievedValue)}</p>
        </div>
      </div>
      <p className="mt-1 font-mono text-[10px] text-muted-foreground">
        GP {money(row.achievedGp)}
      </p>
      {row.sub && row.sub.length > 0 ? (
        <div className="mt-2 flex flex-wrap gap-1 border-t border-border/40 pt-2">
          {row.sub.map((s) => (
            <Badge
              key={s.label}
              variant="outline"
              className="gap-1 font-mono text-[10px] font-normal"
            >
              {s.label} <span className="text-primary">{s.count}</span>
              <span className="text-muted-foreground">{money(s.value)}</span>
            </Badge>
          ))}
        </div>
      ) : null}
    </div>
  );
}

export function KpiSection({ data }: { data: DashboardData }) {
  const { ar, rr, sr, googleReview, threePP } = data.kpi;
  return (
    <div className="space-y-4">
      <div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-4">
        {ar.map((row) => (
          <KpiCard key={row.code} row={row} />
        ))}
      </div>
      <div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-4">
        {rr.map((row) => (
          <KpiCard key={row.code} row={row} />
        ))}
      </div>
      <div className="grid gap-3 sm:grid-cols-3">
        {sr.map((row) => (
          <KpiCard key={row.code} row={row} />
        ))}
      </div>
      <div className="grid gap-3 sm:grid-cols-2">
        <Panel className="p-4">
          <p className="text-[10px] font-semibold uppercase tracking-[0.14em] text-muted-foreground">
            Google Review KPI
          </p>
          <div className="mt-2 grid grid-cols-3 gap-2">
            <Stat label="Previous" value={googleReview.previousMonth} />
            <Stat label="Current" value={googleReview.currentMonth} tone="primary" />
            <Stat
              label="Difference"
              value={`+${googleReview.difference}`}
              tone={googleReview.difference > 0 ? "positive" : "negative"}
            />
          </div>
        </Panel>
        <Panel className="p-4">
          <p className="text-[10px] font-semibold uppercase tracking-[0.14em] text-muted-foreground">
            3PP Attachment (accessory + service/repair invoice)
          </p>
          <div className="mt-2 grid grid-cols-2 gap-2">
            <Stat label="Qualifying Invoices" value={threePP.achieved} tone="primary" />
            <Stat label="Attachment Rate" value={`${((threePP.achieved / Math.max(1, data.store.uniqueInvoices)) * 100).toFixed(1)}%`} />
          </div>
        </Panel>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Invoice dashboard (Section 5)                                       */
/* ------------------------------------------------------------------ */

const PERSON_COLORS: Record<string, { bg: string; chip: string; dot: string }> = {
  SUBHA: {
    bg: "bg-chart-2/10",
    chip: "border-chart-2/40 text-chart-2",
    dot: "bg-chart-2",
  },
  DINESH: {
    bg: "bg-chart-1/10",
    chip: "border-chart-1/40 text-chart-1",
    dot: "bg-chart-1",
  },
};

export function personColor(key: string) {
  return PERSON_COLORS[key] ?? { bg: "bg-accent/30", chip: "border-border text-foreground", dot: "bg-muted-foreground" };
}

export function InvoiceDashboardSection({ data }: { data: DashboardData }) {
  const { days } = data;

  const totals = days.reduce(
    (acc, d) => ({
      inv: acc.inv + d.uniqueInvoices,
      svc: acc.svc + d.serviceRevenue,
      acc: acc.acc + d.accessoryRevenue,
      rep: acc.rep + d.repairRevenue,
      tot: acc.tot + d.totalRevenue,
      gp: acc.gp + d.grossProfit,
    }),
    { inv: 0, svc: 0, acc: 0, rep: 0, tot: 0, gp: 0 },
  );

  return (
    <Panel>
      <PanelHead
        title="Invoice Dashboard — Date-wise Revenue & GP"
        hint="Per-day achievement with salesperson color split and store per-day requirement"
        right={
          <span className="font-mono text-[11px] text-muted-foreground">
            {totals.inv} invoices · {money(totals.tot)} revenue
          </span>
        }
      />
      <div className="overflow-x-auto">
        <table className="w-full min-w-[980px] text-sm">
          <thead>
            <tr className="border-b border-border/70 text-[10px] uppercase tracking-[0.12em] text-muted-foreground">
              <th className="px-3 py-2 text-left font-semibold">Date</th>
              <th className="px-3 py-2 text-right font-semibold">Invoices</th>
              <th className="px-3 py-2 text-right font-semibold">Service</th>
              <th className="px-3 py-2 text-right font-semibold">Accessories</th>
              <th className="px-3 py-2 text-right font-semibold">Repair</th>
              <th className="px-3 py-2 text-right font-semibold">Total Revenue</th>
              <th className="px-3 py-2 text-right font-semibold">GP</th>
              <th className="px-3 py-2 text-right font-semibold">Store Req/Day</th>
              <th className="px-3 py-2 text-left font-semibold">Salesperson Split</th>
            </tr>
          </thead>
          <tbody>
            {days.map((d) => (
              <tr
                key={d.day}
                className="border-b border-border/40 transition-colors last:border-0 hover:bg-accent/20"
              >
                <td className="px-3 py-2">
                  <div className="flex items-center gap-2">
                    <span className="font-mono text-xs font-semibold">{d.dateKey}</span>
                    {d.holiday ? (
                      <Badge variant="outline" className="border-warning/40 text-warning">
                        holiday
                      </Badge>
                    ) : null}
                  </div>
                  <span className="text-[10px] text-muted-foreground">{d.weekday}</span>
                </td>
                <td className="px-3 py-2 text-right font-mono tabular-nums">
                  {d.uniqueInvoices}
                </td>
                <td className="px-3 py-2 text-right font-mono tabular-nums">
                  {money(d.serviceRevenue)}
                </td>
                <td className="px-3 py-2 text-right font-mono tabular-nums">
                  {money(d.accessoryRevenue)}
                </td>
                <td className="px-3 py-2 text-right font-mono tabular-nums">
                  {money(d.repairRevenue)}
                </td>
                <td className="px-3 py-2 text-right font-mono font-semibold tabular-nums">
                  {money(d.totalRevenue)}
                </td>
                <td className="px-3 py-2 text-right font-mono tabular-nums text-positive">
                  {money(d.grossProfit)}
                </td>
                <td className="px-3 py-2 text-right font-mono tabular-nums text-muted-foreground">
                  {money(data.store.total.requiredPerDay)}
                </td>
                <td className="px-3 py-2">
                  <div className="flex flex-wrap gap-1.5">
                    {Object.entries(d.byPerson)
                      .filter(([, v]) => v.totalRevenue > 0 || v.invoices > 0)
                      .map(([k, v]) => (
                        <span
                          key={k}
                          className={`inline-flex items-center gap-1 rounded border px-1.5 py-0.5 font-mono text-[10px] ${personColor(k).chip} ${personColor(k).bg}`}
                        >
                          <span className={`size-1.5 rounded-full ${personColor(k).dot}`} />
                          {k} · {money(v.totalRevenue)}
                        </span>
                      ))}
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
          <tfoot>
            <tr className="border-t border-border/70 bg-accent/20 font-semibold">
              <td className="px-3 py-2">Total</td>
              <td className="px-3 py-2 text-right font-mono">{totals.inv}</td>
              <td className="px-3 py-2 text-right font-mono">{money(totals.svc)}</td>
              <td className="px-3 py-2 text-right font-mono">{money(totals.acc)}</td>
              <td className="px-3 py-2 text-right font-mono">{money(totals.rep)}</td>
              <td className="px-3 py-2 text-right font-mono">{money(totals.tot)}</td>
              <td className="px-3 py-2 text-right font-mono">{money(totals.gp)}</td>
              <td className="px-3 py-2" />
              <td className="px-3 py-2" />
            </tr>
          </tfoot>
        </table>
      </div>
    </Panel>
  );
}
