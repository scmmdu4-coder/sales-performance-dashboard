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
import { computeDashboard } from "@/lib/metrics";
import type { TargetSheet } from "@/lib/sheets";
import { api } from "@/convex/_generated/api";
import { useAction } from "convex/react";
import { motion } from "framer-motion";
import {
  Activity,
  AlertTriangle,
  BadgeCheck,
  Gauge,
  LogOut,
  RefreshCw,
  Users,
} from "lucide-react";
import { useCallback, useEffect, useState } from "react";
import { useNavigate } from "react-router";

function Header({ name, onSignOut }: { name?: string; onSignOut: () => void }) {
  const nav = [
    { href: "#store", label: "Store" },
    { href: "#salesperson", label: "Salesperson" },
    { href: "#amc-bundle", label: "AMC & Bundle" },
    { href: "#kpi", label: "KPI" },
    { href: "#invoice", label: "Invoice" },
  ];
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
          {nav.map((n) => (
            <a
              key={n.href}
              href={n.href}
              className="rounded px-2.5 py-1.5 text-xs font-medium text-muted-foreground transition-colors hover:bg-accent hover:text-foreground"
            >
              {n.label}
            </a>
          ))}
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

type SheetResult = {
  crmRows: import("@/lib/sheets").CrmRow[];
  target: TargetSheet;
  fetchedAt: number;
  crmRowCount: number;
};

export default function Dashboard() {
  const { user, signOut } = useAuth();
  const navigate = useNavigate();
  const fetchSheets = useAction(api.sheets.fetchDashboardData);
  const [sheetData, setSheetData] = useState<SheetResult | null>(null);
  const [loadError, setLoadError] = useState<string | null>(null);
  const [reloadKey, setReloadKey] = useState(0);

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

  useEffect(() => {
    void load();
  }, [load, reloadKey]);

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

  let data: ReturnType<typeof computeDashboard>;
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

  const { store, target } = data;

  return (
    <DashShell>
      <Header name={user?.name ?? undefined} onSignOut={handleSignOut} />

      <main className="mx-auto max-w-7xl space-y-8 px-4 py-6 sm:px-6 sm:py-8">
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
              <span className="text-sm text-positive">
                {" "}
                +{target.googleReview.difference}
              </span>
            </p>
            <p className="font-mono text-[10px] text-muted-foreground">
              prev {target.googleReview.previousMonth}
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
          subtitle="Achieved count vs remaining count to close the next slab"
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
          subtitle="Target, achievement, balance, achieved value & achieved GP per KPI criterion"
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

        <div className="flex flex-wrap items-center justify-between gap-2 border-t border-border/50 pt-4 pb-2 font-mono text-[10px] text-muted-foreground">
          <span>
            source · google sheets crm (gid 0) + target · {sheetData.crmRowCount} rows
          </span>
          <span className="flex items-center gap-1.5">
            <BadgeCheck className="size-3 text-positive" />
            synced {new Date(sheetData.fetchedAt).toLocaleTimeString()}
          </span>
        </div>
      </main>
    </DashShell>
  );
}
