import { useAuth } from "@/hooks/use-auth";
import { Button } from "@/components/ui/button";
import {
  DashShell,
  InvoiceDashboardSection,
  KpiSection,
  Panel,
  AmcBundleSection,
  SalespersonSection,
  Section,
  StorePerformanceSection,
  money,
} from "@/components/dashboard/sections";
import {
  ProcessView,
  ProductivityView,
} from "@/components/dashboard/ops-sections";
import { computeDashboard, type DashboardData } from "@/lib/metrics";
import type { TargetSheet } from "@/lib/sheets";
import type { OpsResult } from "@/lib/ops";
import { cn } from "@/lib/utils";
import { api } from "@/convex/_generated/api";
import { useAction } from "convex/react";
import { motion } from "framer-motion";
import {
  AlertTriangle,
  BadgeCheck,
  Gauge,
  LogOut,
  RefreshCw,
} from "lucide-react";
import { useCallback, useEffect, useState } from "react";
import { useNavigate } from "react-router";

/* ------------------------------------------------------------------ */
/* Dashboard views                                                     */
/* ------------------------------------------------------------------ */

const MODES = [
  { id: "business", label: "Business", hint: "revenue · KPI · incentive" },
  { id: "process", label: "Process", hint: "walk-in · repair · open jobs" },
  { id: "productivity", label: "Productivity", hint: "flow · conversion · TAT" },
] as const;

type DashboardMode = (typeof MODES)[number]["id"];

const BUSINESS_NAV = [
  { href: "#store", label: "Store" },
  { href: "#salesperson", label: "Salesperson" },
  { href: "#amc-bundle", label: "AMC & Bundle" },
  { href: "#kpi", label: "KPI" },
  { href: "#invoice", label: "Invoice" },
];

function ModeSwitch({
  mode,
  onChange,
}: {
  mode: DashboardMode;
  onChange: (mode: DashboardMode) => void;
}) {
  return (
    <div className="inline-flex items-center gap-1 rounded-full border border-border/70 bg-card/60 p-1 backdrop-blur-sm">
      {MODES.map((m) => (
        <button
          key={m.id}
          type="button"
          onClick={() => onChange(m.id)}
          aria-pressed={mode === m.id}
          className={cn(
            "rounded-full px-3.5 py-1.5 text-[11px] font-semibold uppercase tracking-[0.14em] transition-colors sm:px-4",
            mode === m.id
              ? "bg-primary/90 text-primary-foreground shadow"
              : "text-muted-foreground hover:bg-accent hover:text-foreground",
          )}
        >
          {m.label}
        </button>
      ))}
    </div>
  );
}

function Header({
  name,
  mode,
  syncing,
  onSync,
  onSignOut,
}: {
  name?: string;
  mode: DashboardMode;
  syncing: boolean;
  onSync: () => void;
  onSignOut: () => void;
}) {
  return (
    <header className="sticky top-0 z-30 border-b border-border/70 bg-background/80 backdrop-blur-md">
      <div className="mx-auto flex max-w-7xl items-center justify-between gap-4 px-4 py-3 sm:px-6">
        <div className="flex items-center gap-3">
          <div className="flex size-9 items-center justify-center rounded-md border border-primary/40 bg-primary/10">
            <Gauge className="size-4.5 text-primary" />
          </div>
          <div>
            <p className="font-display text-sm font-bold uppercase tracking-[0.2em] text-foreground">
              MDU4<span className="text-primary">/</span>Performance
            </p>
            <p className="font-mono text-[10px] uppercase tracking-[0.18em] text-muted-foreground">
              iPlanet Care · Simmakal
            </p>
          </div>
        </div>
        <nav className="hidden items-center gap-1 lg:flex">
          {mode === "business" ? (
            BUSINESS_NAV.map((n) => (
              <a
                key={n.href}
                href={n.href}
                className="rounded px-2.5 py-1.5 text-xs font-medium text-muted-foreground transition-colors hover:bg-accent hover:text-foreground"
              >
                {n.label}
              </a>
            ))
          ) : (
            <span className="font-mono text-[10px] uppercase tracking-[0.18em] text-muted-foreground">
              {MODES.find((m) => m.id === mode)?.hint}
            </span>
          )}
        </nav>
        <div className="flex items-center gap-2">
          {name ? (
            <span className="hidden font-mono text-xs text-muted-foreground sm:inline">
              {name}
            </span>
          ) : null}
          <Button
            variant="ghost"
            size="icon"
            className="size-8 text-muted-foreground hover:text-foreground"
            onClick={onSync}
            disabled={syncing}
            title="Re-read the MDU4 sheets"
          >
            <RefreshCw className={cn("size-4", syncing && "animate-spin")} />
          </Button>
          <Button
            variant="ghost"
            size="icon"
            className="size-8 text-muted-foreground hover:text-foreground"
            onClick={onSignOut}
          >
            <LogOut className="size-4" />
          </Button>
        </div>
      </div>
      <div className="scan-line h-px w-full opacity-60" />
    </header>
  );
}

function LoadingState({ message }: { message: string }) {
  return (
    <DashShell>
      <div className="flex min-h-screen flex-col items-center justify-center gap-3">
        <RefreshCw className="size-6 animate-spin text-primary" />
        <p className="font-mono text-xs uppercase tracking-[0.2em] text-muted-foreground">
          {message}
        </p>
      </div>
    </DashShell>
  );
}

function ErrorState({
  message,
  onRetry,
}: {
  message: string;
  onRetry: () => void;
}) {
  return (
    <DashShell>
      <div className="flex min-h-screen flex-col items-center justify-center gap-4 px-6">
        <AlertTriangle className="size-8 text-warning" />
        <p className="text-center text-sm text-muted-foreground">{message}</p>
        <p className="max-w-md text-center font-mono text-[11px] text-muted-foreground/70">
          The dashboard reads the shared MDU4 sheets directly. Make sure the
          spreadsheet is shared as “Anyone with the link”.
        </p>
        <Button
          variant="outline"
          size="sm"
          className="gap-2"
          onClick={onRetry}
        >
          <RefreshCw className="size-3.5" /> Retry sync
        </Button>
      </div>
    </DashShell>
  );
}

/** Loading / failure block for the operational sheets. */
function OpsState({
  loading,
  error,
  onRetry,
}: {
  loading: boolean;
  error: string | null;
  onRetry: () => void;
}) {
  return (
    <Panel className="flex flex-col items-center gap-3 p-10 text-center">
      {loading ? (
        <>
          <RefreshCw className="size-5 animate-spin text-primary" />
          <p className="font-mono text-xs uppercase tracking-[0.2em] text-muted-foreground">
            Loading walk-in, service order and communication sheets…
          </p>
          <p className="max-w-md text-[11px] text-muted-foreground/70">
            The service order register is large, so the first load of this view
            takes a few seconds.
          </p>
        </>
      ) : (
        <>
          <AlertTriangle className="size-5 text-warning" />
          <p className="text-sm text-muted-foreground">
            {error ?? "The operational sheets could not be loaded."}
          </p>
          <Button variant="outline" size="sm" className="gap-2" onClick={onRetry}>
            <RefreshCw className="size-3.5" /> Retry
          </Button>
        </>
      )}
    </Panel>
  );
}

/* ------------------------------------------------------------------ */
/* Business view                                                       */
/* ------------------------------------------------------------------ */

function BusinessView({ data }: { data: DashboardData }) {
  const { store, target } = data;
  return (
    <>
      {/* Month summary strip */}
      <motion.div
        initial={{ opacity: 0, y: 8 }}
        animate={{ opacity: 1, y: 0 }}
        transition={{ duration: 0.4 }}
        className="grid grid-cols-2 gap-3 sm:grid-cols-3 lg:grid-cols-6"
      >
        <Panel className="p-3">
          <p className="text-[10px] font-semibold uppercase tracking-[0.14em] text-muted-foreground">
            Month Progress
          </p>
          <p className="mt-1 font-display text-xl font-bold">
            {store.daysElapsed}
            <span className="text-sm text-muted-foreground"> / 30 days</span>
          </p>
          <p className="font-mono text-[10px] text-muted-foreground">
            {store.daysRemaining} remaining
          </p>
        </Panel>
        <Panel className="p-3">
          <p className="text-[10px] font-semibold uppercase tracking-[0.14em] text-muted-foreground">
            Total Invoices
          </p>
          <p className="mt-1 font-display text-xl font-bold tabular-nums">
            {store.totalInvoices}
          </p>
          <p className="font-mono text-[10px] text-muted-foreground">invoice lines</p>
        </Panel>
        <Panel className="p-3">
          <p className="text-[10px] font-semibold uppercase tracking-[0.14em] text-muted-foreground">
            Unique Invoices
          </p>
          <p className="mt-1 font-display text-xl font-bold tabular-nums text-primary">
            {store.uniqueInvoices}
          </p>
          <p className="font-mono text-[10px] text-muted-foreground">deduped codes</p>
        </Panel>
        <Panel className="p-3">
          <p className="text-[10px] font-semibold uppercase tracking-[0.14em] text-muted-foreground">
            Total Revenue
          </p>
          <p className="mt-1 font-display text-xl font-bold tabular-nums">
            {money(store.total.achieved)}
          </p>
          <p className="font-mono text-[10px] text-muted-foreground">
            {store.total.pct.toFixed(1)}% of target
          </p>
        </Panel>
        <Panel className="p-3">
          <p className="text-[10px] font-semibold uppercase tracking-[0.14em] text-muted-foreground">
            Gross Profit
          </p>
          <p className="mt-1 font-display text-xl font-bold tabular-nums text-positive">
            {money(store.grossProfit.achieved)}
          </p>
          <p className="font-mono text-[10px] text-muted-foreground">
            {store.grossProfit.pct.toFixed(1)}% of target
          </p>
        </Panel>
        <Panel className="p-3">
          <p className="text-[10px] font-semibold uppercase tracking-[0.14em] text-muted-foreground">
            Google Reviews
          </p>
          <p className="mt-1 font-display text-xl font-bold tabular-nums">
            {target.googleReview.currentMonth}
          </p>
          <p className="font-mono text-[10px] text-muted-foreground">
            prev {target.googleReview.previousMonth} ·{" "}
            <span
              className={
                target.googleReview.difference >= 0 ? "text-positive" : "text-negative"
              }
            >
              {target.googleReview.difference >= 0 ? "+" : ""}
              {target.googleReview.difference}
            </span>
          </p>
        </Panel>
      </motion.div>

      <Section
        id="store"
        title="Store Performance"
        subtitle="Revenue-wise target, achievement, %, balance and run rate"
        right={
          <span className="font-mono text-[11px] text-muted-foreground">
            targets from target!A1:B5
          </span>
        }
      >
        <StorePerformanceSection data={data} />
      </Section>

      <Section
        id="salesperson"
        title="Salesperson Performance"
        subtitle="Revenue-wise target, achievement, %, balance and eligible incentive per revenue type"
        right={
          <span className="font-mono text-[11px] text-muted-foreground">
            targets from target!D1:E6
          </span>
        }
      >
        <SalespersonSection data={data} />
      </Section>

      <Section
        id="amc-bundle"
        title="AMC & Bundle Incentive Plan"
        subtitle="Every salesperson’s achieved count and how many more are needed to close the next slab"
        right={
          <span className="font-mono text-[11px] text-muted-foreground">
            slabs from target!F17:K23 · M18:P21
          </span>
        }
      >
        <AmcBundleSection data={data} />
      </Section>

      <Section
        id="kpi"
        title="Store KPI — Target vs Achievement"
        subtitle="Target, achievement, balance, achieved value & achieved GP per KPI criterion, with the salesperson split and daily ideas"
        right={
          <span className="font-mono text-[11px] text-muted-foreground">
            grid from target!A9:M13
          </span>
        }
      >
        <KpiSection data={data} />
      </Section>

      <Section
        id="invoice"
        title="Invoice Dashboard"
        subtitle="Day-wise invoices, revenue-type totals, salesperson split & per-day requirement"
        right={
          <span className="font-mono text-[11px] text-muted-foreground">
            req/day = balance ÷ remaining working days
          </span>
        }
      >
        <InvoiceDashboardSection data={data} />
      </Section>
    </>
  );
}

/* ------------------------------------------------------------------ */
/* Page                                                                */
/* ------------------------------------------------------------------ */

type SheetResult = {
  crmRows: import("@/lib/sheets").CrmRow[];
  target: TargetSheet;
  fetchedAt: number;
  crmRowCount: number;
};

type OpsSheetResult = OpsResult & { fetchedAt: number };

export default function Dashboard() {
  const { user, signOut } = useAuth();
  const navigate = useNavigate();
  const fetchSheets = useAction(api.sheets.fetchDashboardData);
  const fetchOps = useAction(api.sheets.fetchOpsData);

  const [sheetData, setSheetData] = useState<SheetResult | null>(null);
  const [loadError, setLoadError] = useState<string | null>(null);
  const [reloadKey, setReloadKey] = useState(0);

  const [mode, setMode] = useState<DashboardMode>("business");
  const [opsData, setOpsData] = useState<OpsSheetResult | null>(null);
  const [opsError, setOpsError] = useState<string | null>(null);
  const [opsLoading, setOpsLoading] = useState(false);
  const [syncing, setSyncing] = useState(false);

  const load = useCallback(async () => {
    setLoadError(null);
    try {
      const res = await fetchSheets({});
      setSheetData(res as SheetResult);
    } catch (err) {
      console.error("Sheet fetch failed", err);
      setLoadError(
        err instanceof Error ? err.message : "Failed to fetch the MDU4 sheets.",
      );
    }
  }, [fetchSheets]);

  const loadOps = useCallback(async () => {
    setOpsLoading(true);
    setOpsError(null);
    try {
      const res = await fetchOps({});
      setOpsData(res as OpsSheetResult);
    } catch (err) {
      console.error("Operational sheet fetch failed", err);
      setOpsError(
        err instanceof Error
          ? err.message
          : "Failed to fetch the operational sheets.",
      );
    } finally {
      setOpsLoading(false);
    }
  }, [fetchOps]);

  useEffect(() => {
    void load();
  }, [load, reloadKey]);

  // The operational sheets are only fetched when one of their views is opened.
  useEffect(() => {
    if (mode === "business") return;
    if (opsData || opsLoading) return;
    void loadOps();
  }, [mode, opsData, opsLoading, loadOps]);

  // Header sync: always refresh the business sheets, and the operational ones
  // when they have already been loaded for this session.
  const syncAll = useCallback(async () => {
    setSyncing(true);
    try {
      await load();
      if (opsData || opsError) await loadOps();
    } finally {
      setSyncing(false);
    }
  }, [load, loadOps, opsData, opsError]);

  const handleSignOut = async () => {
    await signOut();
    navigate("/");
  };

  if (!sheetData && !loadError) {
    return <LoadingState message="Loading MDU4 telemetry…" />;
  }
  if (!sheetData || loadError) {
    return (
      <ErrorState
        message={loadError ?? "Could not compute metrics from the sheet data."}
        onRetry={() => setReloadKey((k) => k + 1)}
      />
    );
  }

  let data: DashboardData;
  try {
    data = computeDashboard(sheetData.crmRows, sheetData.target);
  } catch (err) {
    console.error("computeDashboard failed", err);
    return (
      <ErrorState
        message="Could not compute metrics from the sheet data."
        onRetry={() => setReloadKey((k) => k + 1)}
      />
    );
  }

  const opsReady = opsData !== null && opsError === null;

  return (
    <DashShell>
      <Header
        name={user?.name ?? undefined}
        mode={mode}
        syncing={syncing}
        onSync={() => {
          void syncAll();
        }}
        onSignOut={handleSignOut}
      />

      <main className="mx-auto max-w-7xl space-y-8 px-4 py-6 sm:px-6 sm:py-8">
        <div className="flex flex-wrap items-center justify-between gap-3">
          <ModeSwitch mode={mode} onChange={setMode} />
          <span className="font-mono text-[10px] uppercase tracking-[0.18em] text-muted-foreground">
            {MODES.find((m) => m.id === mode)?.hint}
          </span>
        </div>

        {mode === "business" ? (
          <BusinessView data={data} />
        ) : !opsReady ? (
          <OpsState
            loading={opsLoading}
            error={opsError}
            onRetry={() => {
              void loadOps();
            }}
          />
        ) : mode === "process" ? (
          <ProcessView
            process={opsData!.process}
            googleReview={data.kpi.googleReview}
            uniqueInvoices={data.store.uniqueInvoices}
          />
        ) : (
          <ProductivityView productivity={opsData!.productivity} />
        )}

        <div className="flex flex-wrap items-center justify-between gap-2 border-t border-border/50 pt-4 pb-2 font-mono text-[10px] text-muted-foreground">
          <span>
            {mode === "business"
              ? `source · google sheets crm (gid 0) + target · ${sheetData.crmRowCount} rows`
              : "source · google sheets walkin + aod + comm · server-side aggregates"}
          </span>
          <span className="flex items-center gap-1.5">
            <BadgeCheck className="size-3 text-positive" />
            synced{" "}
            {new Date(
              mode === "business"
                ? sheetData.fetchedAt
                : (opsData?.fetchedAt ?? sheetData.fetchedAt),
            ).toLocaleTimeString()}
          </span>
        </div>
      </main>
    </DashShell>
  );
}
