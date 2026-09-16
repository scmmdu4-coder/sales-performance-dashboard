import { cn } from "@/lib/utils";
import { Fragment, type ReactNode } from "react";
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
          <p className="mt-0.5 text-[11px] leading-snug text-muted-foreground/70">
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

/** Short money for dense cells: ₹1.2k / ₹1.4L */
export const moneyCompact = (n: number) =>
  n >= 100000
    ? `₹${(n / 100000).toFixed(1)}L`
    : n >= 1000
      ? `₹${(n / 1000).toFixed(1)}k`
      : `₹${Math.round(n)}`;

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
/* Per-salesperson colour language                                     */
/* ------------------------------------------------------------------ */

interface PersonColor {
  bg: string;
  chip: string;
  dot: string;
  text: string;
}

const PERSON_COLORS: Record<string, PersonColor> = {
  SUBHA: {
    bg: "bg-chart-2/10",
    chip: "border-chart-2/40 text-chart-2",
    dot: "bg-chart-2",
    text: "text-chart-2",
  },
  DINESH: {
    bg: "bg-chart-1/10",
    chip: "border-chart-1/40 text-chart-1",
    dot: "bg-chart-1",
    text: "text-chart-1",
  },
};

export function personColor(key: string): PersonColor {
  return (
    PERSON_COLORS[key] ?? {
      bg: "bg-accent/30",
      chip: "border-border text-foreground",
      dot: "bg-muted-foreground",
      text: "text-foreground",
    }
  );
}

export function PersonTag({ personKey, label }: { personKey: string; label?: ReactNode }) {
  const c = personColor(personKey);
  return (
    <span className="inline-flex items-center gap-1.5">
      <span className={cn("size-1.5 rounded-full", c.dot)} />
      <span className={c.text}>{label ?? personKey}</span>
    </span>
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
            req/day <span className="text-foreground">{money(b.requiredPerDay)}</span>
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

export function StorePerformanceSection({ data }: { data: DashboardData }) {
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
        <Stat
          label="Revenue Rate"
          value={`${s.total.pct.toFixed(1)}%`}
          tone={s.total.pct >= 100 ? "positive" : "default"}
        />
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
  const c = personColor(row.key);
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
        <p className={cn("mt-3 border-t border-border/40 pt-2 font-mono text-[10px]", c.text)}>
          target · service {money(row.service.target)} · accessory {money(row.accessory.target)} ·
          repair {money(row.repair.target)}
        </p>
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
  const amcTotal = amc.reduce((a, r) => a + r.projectedIncentive, 0);
  const bundleTotal = bundle.reduce((a, r) => a + r.projectedIncentive, 0);

  return (
    <div className="space-y-4">
      <Panel>
        <PanelHead
          title="AMC Incentive — achieved count vs next slab"
          hint="AMC = unique invoices with crm!SubCategory “AMC Charge” · accessory AMC = those AMC invoices without a service/repair line"
          right={<span className="font-mono text-[11px] text-positive">{money(amcTotal)}</span>}
        />
        <MiniTable
          head={[
            "Salesperson",
            "AMC Ach",
            "Current Slab",
            "Eligible",
            "Next Slab",
            "To Next",
            "Acc. AMC",
            "Acc. Tier",
            "Acc. To Next",
          ]}
          rows={amc.map((row) => [
            <PersonTag personKey={row.key} label={row.name} />,
            <span className="font-semibold">{row.amcCount}</span>,
            row.amcSlab ? (
              <span className="text-[11px]">
                {row.amcSlab.label}
                {row.amcSlab.rate > 0 ? ` · ₹${row.amcSlab.rate}/AMC` : " · Nil"}
              </span>
            ) : (
              <span className="text-muted-foreground">none</span>
            ),
            <span className={row.amcEligible > 0 ? "text-positive" : "text-muted-foreground"}>
              {row.amcEligible > 0 ? money(row.amcEligible) : "Nil"}
            </span>,
            row.amcNextSlab ? (
              <span className="text-[11px]">
                {row.amcNextSlab.label} · ₹{row.amcNextSlab.rate}/AMC
              </span>
            ) : (
              <span className="text-muted-foreground">top slab</span>
            ),
            <span className={row.amcRemaining === 0 ? "text-positive" : "text-warning"}>
              {row.amcNextSlab ? `${row.amcRemaining} more` : "—"}
            </span>,
            row.accessoryCount,
            row.accessorySlab ? (
              <span className="text-[11px]">
                {row.accessorySlab.accCount}+ · ₹{row.accessorySlab.accRate}
                {row.accessorySlab.accFixed ? ` + ₹${row.accessorySlab.accFixed}` : ""}
              </span>
            ) : (
              <span className="text-muted-foreground">—</span>
            ),
            <span
              className={
                row.accessoryNextSlab ? "text-warning" : "text-muted-foreground"
              }
            >
              {row.accessoryNextSlab ? `${row.accessoryRemaining} more` : "—"}
            </span>,
          ])}
        />
      </Panel>

      <Panel>
        <PanelHead
          title="Bundle Incentive — achieved count vs next slab"
          hint="Bundle = unique invoices whose crm!SchemeCode starts with “DISC”"
          right={<span className="font-mono text-[11px] text-positive">{money(bundleTotal)}</span>}
        />
        <MiniTable
          head={[
            "Salesperson",
            "Bundles",
            "Current Slab",
            "Rate / Bundle",
            "Eligible",
            "Next Slab",
            "To Next",
          ]}
          rows={bundle.map((row) => [
            <PersonTag personKey={row.key} label={row.name} />,
            <span className="font-semibold">{row.count}</span>,
            row.slab ? (
              <span className="text-[11px]">level {row.slab.min}+</span>
            ) : (
              <span className="text-muted-foreground">not eligible</span>
            ),
            row.slab ? `₹${row.slab.perBundle}` : "—",
            <span className={row.projectedIncentive > 0 ? "text-positive" : "text-muted-foreground"}>
              {money(row.projectedIncentive)}
            </span>,
            row.nextSlab ? (
              <span className="text-[11px]">
                {row.nextSlab.min} bundles · ₹{row.nextSlab.perBundle}
              </span>
            ) : (
              <span className="text-muted-foreground">top slab</span>
            ),
            <span className={row.remaining === 0 ? "text-positive" : "text-warning"}>
              {row.nextSlab ? `${row.remaining} more` : "—"}
            </span>,
          ])}
        />
      </Panel>

      <Panel className="p-3">
        <p className="text-[10px] font-semibold uppercase tracking-[0.14em] text-muted-foreground">
          Plan reference
        </p>
        <p className="mt-1 font-mono text-[11px] leading-relaxed text-muted-foreground">
          {target.amcSlabs
            .map((s) => `${s.label}: ${s.rateLabel || "—"}${s.accRate ? ` · acc ${s.accCount}+ @ ₹${s.accRate}/AMC` : ""}`)
            .join("  ·  ")}
        </p>
        <p className="mt-1.5 font-mono text-[11px] leading-relaxed text-muted-foreground">
          {target.bundleSlabs.map((s) => `${s.label} → ₹${s.perBundle}/bundle`).join("  ·  ")}
        </p>
      </Panel>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* KPI grid (Section 4)                                                */
/* ------------------------------------------------------------------ */

function PersonChips({ byPerson }: { byPerson: Record<string, number> }) {
  const entries = Object.entries(byPerson).filter(([, n]) => n > 0);
  if (entries.length === 0) return null;
  return (
    <div className="mt-2 flex flex-wrap gap-1">
      {entries
        .sort((a, b) => b[1] - a[1])
        .map(([key, n]) => (
          <span
            key={key}
            className={cn(
              "inline-flex items-center gap-1 rounded border px-1.5 py-0.5 font-mono text-[10px] font-normal",
              personColor(key).chip,
              personColor(key).bg,
            )}
          >
            <span className={cn("size-1.5 rounded-full", personColor(key).dot)} />
            {key} {n}
          </span>
        ))}
    </div>
  );
}

function KpiCard({ row }: { row: KpiRow }) {
  return (
    <div className="rounded-md border border-border/60 bg-background/40 p-3">
      <div className="flex items-start justify-between gap-2">
        <div>
          <p className="text-sm font-semibold">{row.label}</p>
          <p className="font-mono text-[10px] text-muted-foreground">
            {row.category} · value ₹{row.value.toLocaleString("en-IN")} · approx{" "}
            {money(row.approxAmount)}
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
      <PersonChips byPerson={row.byPerson} />
    </div>
  );
}

function GoogleReviewCard({
  googleReview,
}: {
  googleReview: DashboardData["kpi"]["googleReview"];
}) {
  const { previousMonth, currentMonth, difference } = googleReview;
  return (
    <div className="rounded-md border border-border/60 bg-background/40 p-3">
      <div className="flex items-start justify-between gap-2">
        <div>
          <p className="text-sm font-semibold">Google Review</p>
          <p className="font-mono text-[10px] text-muted-foreground">
            AR category · target!G1:I3
          </p>
        </div>
        <span
          className={cn(
            "inline-flex items-center rounded border px-1.5 py-0.5 font-mono text-[11px] font-semibold tabular-nums",
            difference >= 0
              ? "border-positive/40 bg-positive/10 text-positive"
              : "border-negative/40 bg-negative/10 text-negative",
          )}
        >
          {difference >= 0 ? "+" : ""}
          {difference}
        </span>
      </div>
      <div className="mt-2 grid grid-cols-2 gap-1.5 font-mono text-[10px]">
        <div>
          <p className="text-muted-foreground">PREV MONTH</p>
          <p className="text-foreground">{previousMonth}</p>
        </div>
        <div>
          <p className="text-muted-foreground">CURRENT MONTH</p>
          <p className="text-primary">{currentMonth}</p>
        </div>
      </div>
      <p className="mt-2 font-mono text-[10px] text-muted-foreground">
        difference {difference >= 0 ? "+" : ""}
        {difference} reviews vs previous month
      </p>
    </div>
  );
}

const SR_PRODUCT_ROWS = ["iPhone", "Mac", "Watch", "iPad", "AirPods", "Consumables"];
const SR_MATRIX_CODES = ["SHC", "SRD", "ARD"] as const;

/** Small-font product mix under the SR KPI, e.g. SRD-IPH-002 → iPhone. */
function SrProductMatrix({ sr }: { sr: KpiRow[] }) {
  const byCode = new Map(sr.map((r) => [r.code, r]));
  const cell = (code: string, label: string) =>
    (byCode.get(code)?.sub ?? []).find((s) => s.label === label) ?? null;

  return (
    <Panel>
      <PanelHead
        title="SR KPI — product mix (small font)"
        hint="Item codes starting with SRD / SHC / ARD, split by the product token inside the code (iph · air · mac · wat · pad · con)"
        right={
          <span className="font-mono text-[11px] text-muted-foreground">
            prefix rule · crm!ItemCode
          </span>
        }
      />
      <div className="overflow-x-auto">
        <table className="w-full min-w-[520px] text-sm">
          <thead>
            <tr className="border-b border-border/70 text-[10px] uppercase tracking-[0.12em] text-muted-foreground">
              <th className="px-3 py-2 text-left font-semibold">Product</th>
              {SR_MATRIX_CODES.map((code) => (
                <th key={code} className="px-3 py-2 text-right font-semibold">
                  {code}
                  <span className="block text-[9px] font-normal normal-case tracking-normal text-muted-foreground/70">
                    achieved {byCode.get(code)?.achieved ?? 0} / target{" "}
                    {byCode.get(code)?.target ?? 0}
                  </span>
                </th>
              ))}
              <th className="px-3 py-2 text-right font-semibold">Total</th>
            </tr>
          </thead>
          <tbody>
            {SR_PRODUCT_ROWS.map((label) => {
              const cells = SR_MATRIX_CODES.map((code) => cell(code, label));
              const rowTotal = cells.reduce((a, c) => a + (c?.count ?? 0), 0);
              const rowValue = cells.reduce((a, c) => a + (c?.value ?? 0), 0);
              return (
                <tr
                  key={label}
                  className="border-b border-border/40 transition-colors last:border-0 hover:bg-accent/20"
                >
                  <td className="px-3 py-1.5 text-left text-xs">{label}</td>
                  {cells.map((c, i) => (
                    <td key={i} className="px-3 py-1.5 text-right align-top">
                      <span className="font-mono text-sm font-semibold tabular-nums">
                        {c?.count ?? 0}
                      </span>
                      <span className="block font-mono text-[9px] text-muted-foreground/70">
                        {moneyCompact(c?.value ?? 0)}
                      </span>
                    </td>
                  ))}
                  <td className="px-3 py-1.5 text-right align-top">
                    <span className="font-mono text-sm font-semibold tabular-nums">
                      {rowTotal}
                    </span>
                    <span className="block font-mono text-[9px] text-muted-foreground/70">
                      {moneyCompact(rowValue)}
                    </span>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </Panel>
  );
}

function SalespersonKpiTable({ data }: { data: DashboardData }) {
  const people = data.salespeople.map((p) => p.key);
  const rows = [...data.kpi.ar, ...data.kpi.rr, ...data.kpi.sr];
  const workingDays = Math.max(1, data.store.remainingWorkingDays);

  return (
    <Panel>
      <PanelHead
        title="Salesperson KPI — target vs achievement"
        hint="targets from target!A9:M13 · achievement is unique invoices per salesperson"
        right={
          <span className="font-mono text-[11px] text-muted-foreground">
            {data.store.remainingWorkingDays} working days left
          </span>
        }
      />
      <div className="overflow-x-auto">
        <table className="w-full min-w-[560px] text-sm">
          <thead>
            <tr className="border-b border-border/70 text-[10px] uppercase tracking-[0.12em] text-muted-foreground">
              <th className="px-3 py-2 text-left font-semibold">KPI</th>
              <th className="px-3 py-2 text-left font-semibold">Cat</th>
              {people.map((p) => (
                <th
                  key={p}
                  className={cn("px-3 py-2 text-right font-semibold", personColor(p).text)}
                >
                  {p}
                </th>
              ))}
              <th className="px-3 py-2 text-right font-semibold">Target</th>
              <th className="px-3 py-2 text-right font-semibold">Balance</th>
              <th className="px-3 py-2 text-right font-semibold">Req / day</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((row) => (
              <tr
                key={row.code}
                className="border-b border-border/40 transition-colors last:border-0 hover:bg-accent/20"
              >
                <td className="px-3 py-2 text-left text-xs font-medium">{row.label}</td>
                <td className="px-3 py-2 text-left font-mono text-[10px] text-muted-foreground">
                  {row.category}
                </td>
                {people.map((p) => (
                  <td
                    key={p}
                    className={cn(
                      "px-3 py-2 text-right font-mono tabular-nums",
                      personColor(p).text,
                    )}
                  >
                    {row.byPerson[p] ?? 0}
                  </td>
                ))}
                <td className="px-3 py-2 text-right font-mono tabular-nums">{row.target}</td>
                <td
                  className={cn(
                    "px-3 py-2 text-right font-mono tabular-nums",
                    row.balance === 0 ? "text-positive" : "text-warning",
                  )}
                >
                  {row.balance}
                </td>
                <td className="px-3 py-2 text-right font-mono tabular-nums text-muted-foreground">
                  {row.target > 0 ? Math.ceil(row.balance / workingDays) : "—"}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </Panel>
  );
}

function DailyIdeasPanel({ data }: { data: DashboardData }) {
  const ideas = data.kpi.dailyIdeas;
  return (
    <Panel>
      <PanelHead
        title="Daily target achievement ideas"
        hint={`balance ÷ ${data.store.remainingWorkingDays} working days left (target!A9:M13)`}
      />
      {ideas.length === 0 ? (
        <p className="px-4 py-3 text-xs text-muted-foreground">
          Every KPI target is already closed for the month.
        </p>
      ) : (
        <div className="divide-y divide-border/40">
          {ideas.map((idea) => (
            <div
              key={idea.code}
              className="flex items-start justify-between gap-3 px-4 py-2.5"
            >
              <div className="min-w-0">
                <p className="text-xs font-semibold">
                  {idea.label}{" "}
                  <span className="font-mono text-[10px] text-muted-foreground">
                    {idea.category}
                  </span>
                </p>
                <p className="mt-0.5 text-[11px] text-muted-foreground">{idea.hint}</p>
              </div>
              <div className="shrink-0 text-right">
                <p className="font-mono text-sm font-semibold text-warning">
                  {idea.perDay}/day
                </p>
                <p className="font-mono text-[10px] text-muted-foreground">
                  {idea.achieved}/{idea.target} · {idea.balance} left
                </p>
              </div>
            </div>
          ))}
        </div>
      )}
    </Panel>
  );
}

export function KpiSection({ data }: { data: DashboardData }) {
  const { ar, rr, sr, googleReview, threePP } = data.kpi;
  return (
    <div className="space-y-4">
      <div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-5">
        {ar.map((row) => (
          <KpiCard key={row.code} row={row} />
        ))}
        <GoogleReviewCard googleReview={googleReview} />
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
      <SrProductMatrix sr={sr} />
      <div className="grid gap-4 xl:grid-cols-2">
        <SalespersonKpiTable data={data} />
        <DailyIdeasPanel data={data} />
      </div>
      <Panel className="p-4">
        <p className="text-[10px] font-semibold uppercase tracking-[0.14em] text-muted-foreground">
          3PP Attachment (accessory + service/repair on the same invoice)
        </p>
        <div className="mt-2 grid grid-cols-2 gap-2 sm:grid-cols-3">
          <Stat label="Qualifying Invoices" value={threePP.achieved} tone="primary" />
          <Stat
            label="Attachment Rate"
            value={`${((threePP.achieved / Math.max(1, data.store.uniqueInvoices)) * 100).toFixed(1)}%`}
            sub="of unique invoices"
          />
          <Stat
            label="Accessory Invoices"
            value={
              data.revenueTypes.find((r) => r.key === "accessory")?.uniqueInvoices ?? 0
            }
            sub="unique accessory bills"
          />
        </div>
      </Panel>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Invoice dashboard (Section 5)                                       */
/* ------------------------------------------------------------------ */

export function InvoiceDashboardSection({ data }: { data: DashboardData }) {
  const { days, store, revenueTypes } = data;
  const people = data.salespeople.map((p) => p.key);
  const targetPerDay = store.total.requiredPerDay;

  const totals = days.reduce(
    (acc, d) => ({
      inv: acc.inv + d.uniqueInvoices,
      tot: acc.tot + d.totalRevenue,
      gp: acc.gp + d.grossProfit,
    }),
    { inv: 0, tot: 0, gp: 0 },
  );

  const personTotals = people.map((p) =>
    days.reduce(
      (acc, d) => {
        const v = d.byPerson[p];
        return {
          service: acc.service + (v?.serviceRevenue ?? 0),
          achieved: acc.achieved + (v?.totalRevenue ?? 0),
        };
      },
      { service: 0, achieved: 0 },
    ),
  );

  return (
    <div className="space-y-4">
      <div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-5">
        <Stat
          label="Overall Invoices"
          value={store.totalInvoices}
          sub="invoice lines"
        />
        <Stat
          label="Unique Invoices"
          value={store.uniqueInvoices}
          sub="deduped invoice codes"
          tone="primary"
        />
        {revenueTypes.map((rt) => (
          <Stat
            key={rt.key}
            label={rt.label}
            value={money(rt.revenue)}
            sub={`${rt.uniqueInvoices} unique · GP ${money(rt.gp)}`}
          />
        ))}
      </div>

      <Panel>
        <PanelHead
          title="Date-wise Revenue & GP Report"
          hint="One line per date · salesperson cells colour-coded · target/day taken from the store per-day requirement"
          right={
            <span className="font-mono text-[11px] text-muted-foreground">
              {totals.inv} invoices · {money(totals.tot)} revenue
            </span>
          }
        />
        <div className="flex flex-wrap items-center gap-3 border-b border-border/50 px-4 py-2">
          {people.map((p) => (
            <span
              key={p}
              className={cn(
                "inline-flex items-center gap-1.5 rounded border px-1.5 py-0.5 font-mono text-[10px]",
                personColor(p).chip,
                personColor(p).bg,
              )}
            >
              <span className={cn("size-1.5 rounded-full", personColor(p).dot)} />
              {p}
            </span>
          ))}
          <span className="font-mono text-[10px] text-muted-foreground">
            target / day {money(targetPerDay)}
          </span>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full min-w-[1040px] text-sm">
            <thead>
              <tr className="border-b border-border/40 text-[10px] uppercase tracking-[0.12em] text-muted-foreground">
                <th rowSpan={2} className="px-3 py-2 text-left align-bottom font-semibold">
                  Date
                </th>
                <th rowSpan={2} className="px-3 py-2 text-right align-bottom font-semibold">
                  Invoices
                </th>
                {people.map((p) => (
                  <th
                    key={p}
                    colSpan={2}
                    className={cn(
                      "border-l border-border/40 px-3 py-2 text-center font-semibold",
                      personColor(p).text,
                    )}
                  >
                    {p}
                  </th>
                ))}
                <th
                  rowSpan={2}
                  className="border-l border-border/40 px-3 py-2 text-right align-bottom font-semibold"
                >
                  Target / Day
                </th>
                <th rowSpan={2} className="px-3 py-2 text-right align-bottom font-semibold">
                  Achieved
                </th>
                <th rowSpan={2} className="px-3 py-2 text-right align-bottom font-semibold">
                  GP
                </th>
                <th rowSpan={2} className="px-3 py-2 text-right align-bottom font-semibold">
                  vs Target
                </th>
              </tr>
              <tr className="border-b border-border/70 text-[10px] uppercase tracking-[0.12em] text-muted-foreground">
                {people.map((p) => (
                  <Fragment key={p}>
                    <th className="border-l border-border/40 px-3 py-1.5 text-right font-semibold">
                      Service Rev
                    </th>
                    <th className="px-3 py-1.5 text-right font-semibold">Achieved</th>
                  </Fragment>
                ))}
              </tr>
            </thead>
            <tbody>
              {days.map((d) => {
                const deltaPct =
                  targetPerDay > 0
                    ? ((d.totalRevenue - targetPerDay) / targetPerDay) * 100
                    : 0;
                return (
                  <tr
                    key={d.day}
                    className="border-b border-border/40 transition-colors last:border-0 hover:bg-accent/20"
                  >
                    <td className="px-3 py-2">
                      <div className="flex items-center gap-2">
                        <span className="font-mono text-xs font-semibold">
                          {d.undated ? "No date" : d.dateKey}
                        </span>
                        {d.holiday ? (
                          <Badge variant="outline" className="border-warning/40 text-warning">
                            holiday
                          </Badge>
                        ) : null}
                        {d.undated ? (
                          <Badge variant="outline" className="border-muted-foreground/40 text-muted-foreground">
                            undated
                          </Badge>
                        ) : null}
                      </div>
                      <span className="text-[10px] text-muted-foreground">
                        {d.undated ? "InvoiceDate blank in crm" : d.weekday}
                      </span>
                    </td>
                    <td className="px-3 py-2 text-right font-mono tabular-nums">
                      {d.uniqueInvoices}
                    </td>
                    {people.map((p) => {
                      const v = d.byPerson[p];
                      const c = personColor(p);
                      return (
                        <Fragment key={p}>
                          <td
                            className={cn(
                              "border-l border-border/40 px-3 py-2 text-right font-mono tabular-nums",
                              c.bg,
                              v && v.serviceRevenue > 0 ? c.text : "text-muted-foreground",
                            )}
                          >
                            {v && v.serviceRevenue > 0 ? money(v.serviceRevenue) : "—"}
                          </td>
                          <td
                            className={cn(
                              "px-3 py-2 text-right font-mono font-semibold tabular-nums",
                              c.bg,
                              v && v.totalRevenue > 0 ? "text-foreground" : "text-muted-foreground",
                            )}
                          >
                            {v && v.totalRevenue > 0 ? money(v.totalRevenue) : "—"}
                          </td>
                        </Fragment>
                      );
                    })}
                    <td className="border-l border-border/40 px-3 py-2 text-right font-mono tabular-nums text-muted-foreground">
                      {money(targetPerDay)}
                    </td>
                    <td className="px-3 py-2 text-right font-mono font-semibold tabular-nums">
                      {money(d.totalRevenue)}
                    </td>
                    <td className="px-3 py-2 text-right font-mono tabular-nums text-positive">
                      {money(d.grossProfit)}
                    </td>
                    <td className="px-3 py-2 text-right font-mono text-[11px] tabular-nums">
                      {d.undated ? (
                        <span className="text-muted-foreground">—</span>
                      ) : (
                        <span className={deltaPct >= 0 ? "text-positive" : "text-negative"}>
                          {deltaPct >= 0 ? "+" : ""}
                          {deltaPct.toFixed(0)}%
                        </span>
                      )}
                    </td>
                  </tr>
                );
              })}
            </tbody>
            <tfoot>
              <tr className="border-t border-border/70 bg-accent/20 font-semibold">
                <td className="px-3 py-2">Total</td>
                <td className="px-3 py-2 text-right font-mono">{totals.inv}</td>
                {people.map((p, i) => (
                  <Fragment key={p}>
                    <td
                      className={cn(
                        "border-l border-border/40 px-3 py-2 text-right font-mono",
                        personColor(p).text,
                      )}
                    >
                      {money(personTotals[i].service)}
                    </td>
                    <td className="px-3 py-2 text-right font-mono">
                      {money(personTotals[i].achieved)}
                    </td>
                  </Fragment>
                ))}
                <td className="border-l border-border/40 px-3 py-2 text-right font-mono text-muted-foreground">
                  {money(targetPerDay)}
                </td>
                <td className="px-3 py-2 text-right font-mono">{money(totals.tot)}</td>
                <td className="px-3 py-2 text-right font-mono">{money(totals.gp)}</td>
                <td className="px-3 py-2" />
              </tr>
            </tfoot>
          </table>
        </div>
      </Panel>
    </div>
  );
}
