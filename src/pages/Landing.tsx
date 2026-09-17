import { motion } from "framer-motion";
import {
  ArrowRight,
  Gauge,
  HeartPulse,
  LineChart,
  Lock,
  ShieldCheck,
  Timer,
  Target,
  TrendingUp,
  Users,
  Zap,
} from "lucide-react";
import { Link } from "react-router";

const features = [
  {
    icon: Target,
    title: "Store targets, live",
    desc: "Revenue-wise target, achievement, balance and run rate for service, accessories and repair — straight from the target sheet.",
  },
  {
    icon: Users,
    title: "Salesperson scorecards",
    desc: "Per-person achievement, eligible incentive per revenue type and total incentive, updated with every invoice.",
  },
  {
    icon: Zap,
    title: "AMC & bundle slabs",
    desc: "Count achieved vs remaining to close the next slab, for AMC and bundle schemes alike.",
  },
  {
    icon: ShieldCheck,
    title: "KPI criteria grid",
    desc: "Care+, Protect+, AMC, bundle, display, battery, glass, others, SRD/SHC/ARD — with achieved value and GP.",
  },
  {
    icon: LineChart,
    title: "Day-wise invoice dashboard",
    desc: "Every day's invoices, revenue-type split and per-person achievement with store per-day requirement.",
  },
  {
    icon: TrendingUp,
    title: "Run-rate projections",
    desc: "Balance over remaining working days so managers always know the pace needed to land the month.",
  },
  {
    icon: HeartPulse,
    title: "Walk-in process",
    desc: "Daily walk-ins, case creation, invoice creation and communication updates — with open-jobs ageing and repair-type trends.",
  },
  {
    icon: Timer,
    title: "Technician TAT",
    desc: "Conversion funnel, technician productivity and case-by-case repair turnaround time, all in one productivity view.",
  },
];

export default function Landing() {
  return (
    <div className="bg-grid min-h-screen bg-background">
      <div className="bg-glow-primary min-h-screen">
        {/* Nav */}
        <header className="mx-auto flex max-w-6xl items-center justify-between px-6 py-5">
          <div className="flex items-center gap-2.5">
            <div className="flex size-9 items-center justify-center rounded-md border border-primary/40 bg-primary/10">
              <Gauge className="size-4.5 text-primary" />
            </div>
            <span className="font-display text-sm font-bold uppercase tracking-[0.2em]">
              MDU4<span className="text-primary">/</span>Performance
            </span>
          </div>
          <Link
            to="/auth"
            className="rounded-md border border-border px-4 py-2 text-sm font-medium text-foreground transition-colors hover:border-primary/50 hover:text-primary"
          >
            Sign in
          </Link>
        </header>

        {/* Hero */}
        <section className="mx-auto max-w-6xl px-6 pt-14 pb-16 text-center sm:pt-20">
          <motion.p
            initial={{ opacity: 0, y: 8 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.5 }}
            className="mx-auto inline-flex items-center gap-2 rounded-full border border-primary/30 bg-primary/10 px-3 py-1 font-mono text-[11px] uppercase tracking-[0.18em] text-primary"
          >
            <span className="size-1.5 animate-pulse rounded-full bg-positive" />
            iPlanet Care · Simmakal · September
          </motion.p>
          <motion.h1
            initial={{ opacity: 0, y: 12 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.55, delay: 0.05 }}
            className="font-display mx-auto mt-5 max-w-3xl text-4xl font-extrabold leading-tight tracking-tight sm:text-6xl"
          >
            The MDU4 performance{" "}
            <span className="text-primary text-glow">dashboard</span> for
            sales managers
          </motion.h1>
          <motion.p
            initial={{ opacity: 0, y: 12 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.55, delay: 0.12 }}
            className="mx-auto mt-5 max-w-2xl text-base text-muted-foreground sm:text-lg"
          >
            One console for store targets, salesperson incentives, KPI
            criteria and the date-wise revenue &amp; GP report — computed live
            from the MDU4 sheets.
          </motion.p>
          <motion.div
            initial={{ opacity: 0, y: 12 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.55, delay: 0.2 }}
            className="mt-8 flex flex-wrap items-center justify-center gap-3"
          >
            <Link
              to="/dashboard"
              className="inline-flex items-center gap-2 rounded-md bg-primary px-5 py-2.5 text-sm font-semibold text-primary-foreground transition-transform hover:scale-[1.02]"
            >
              Open dashboard <ArrowRight className="size-4" />
            </Link>
            <a
              href="#features"
              className="inline-flex items-center gap-2 rounded-md border border-border px-5 py-2.5 text-sm font-medium text-foreground transition-colors hover:border-primary/50"
            >
              Explore sections
            </a>
          </motion.div>

          {/* Hero telemetry strip */}
          <motion.div
            initial={{ opacity: 0, y: 16 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.6, delay: 0.28 }}
            className="mx-auto mt-12 grid max-w-4xl grid-cols-2 gap-3 sm:grid-cols-4"
          >
            {[
              ["Month progress", "11 / 30 days"],
              ["Revenue target", "₹9,60,000"],
              ["KPI criteria", "11 tracked"],
              ["Invoice lines", "live sync"],
            ].map(([label, value]) => (
              <div
                key={label}
                className="rounded-lg border border-border/70 bg-card/60 px-4 py-3 text-left backdrop-blur-sm"
              >
                <p className="font-mono text-[10px] uppercase tracking-[0.14em] text-muted-foreground">
                  {label}
                </p>
                <p className="font-display mt-1 text-lg font-bold">{value}</p>
              </div>
            ))}
          </motion.div>
        </section>

        {/* Features */}
        <section id="features" className="mx-auto max-w-6xl px-6 py-14">
          <p className="font-mono text-[11px] uppercase tracking-[0.2em] text-primary">
            What's inside
          </p>
          <h2 className="font-display mt-2 text-2xl font-bold sm:text-3xl">
            Every number a sales manager needs, in one screen
          </h2>
          <div className="mt-8 grid gap-4 sm:grid-cols-2 lg:grid-cols-3">
            {features.map((f, i) => (
              <motion.div
                key={f.title}
                initial={{ opacity: 0, y: 14 }}
                whileInView={{ opacity: 1, y: 0 }}
                viewport={{ once: true, margin: "-60px" }}
                transition={{ duration: 0.45, delay: i * 0.05 }}
                className="group rounded-lg border border-border/70 bg-card/60 p-5 transition-colors hover:border-primary/40"
              >
                <div className="flex size-9 items-center justify-center rounded-md border border-primary/30 bg-primary/10 text-primary">
                  <f.icon className="size-4.5" />
                </div>
                <h3 className="font-display mt-3 text-base font-semibold">
                  {f.title}
                </h3>
                <p className="mt-1.5 text-sm leading-relaxed text-muted-foreground">
                  {f.desc}
                </p>
              </motion.div>
            ))}
          </div>
        </section>

        {/* CTA */}
        <section className="mx-auto max-w-6xl px-6 pb-20">
          <div className="relative overflow-hidden rounded-xl border border-primary/25 bg-primary/5 p-8 text-center sm:p-12">
            <div className="scan-line absolute inset-x-0 top-0 h-px opacity-70" />
            <h2 className="font-display text-2xl font-bold sm:text-3xl">
              Ready to run the month at a glance?
            </h2>
            <p className="mx-auto mt-3 max-w-xl text-sm text-muted-foreground">
              Sign in to open the MDU4 performance console — Business,
              Process and Productivity views, all backed by the live sheets.
            </p>
            <Link
              to="/auth"
              className="mt-6 inline-flex items-center gap-2 rounded-md bg-primary px-6 py-3 text-sm font-semibold text-primary-foreground transition-transform hover:scale-[1.02]"
            >
              <Lock className="size-4" /> Sign in to continue
            </Link>
          </div>
        </section>

        <footer className="border-t border-border/50">
          <div className="mx-auto flex max-w-6xl flex-wrap items-center justify-between gap-2 px-6 py-5 font-mono text-[10px] uppercase tracking-[0.16em] text-muted-foreground">
            <span>MDU4 Performance Console</span>
            <span>iPlanet Care · Simmakal · Madurai</span>
          </div>
        </footer>
      </div>
    </div>
  );
}
