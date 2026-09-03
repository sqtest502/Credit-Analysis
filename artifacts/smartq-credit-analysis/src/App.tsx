import { useRef, useState, type ChangeEvent, type DragEvent, type ReactNode } from 'react';
import {
  AlertTriangle,
  ArrowDownToLine,
  ArrowRight,
  Building2,
  Check,
  CheckCircle2,
  Database,
  FileSpreadsheet,
  Files,
  History,
  Info,
  Layers3,
  LoaderCircle,
  RefreshCw,
  Search,
  ShieldCheck,
  Table2,
  UploadCloud,
} from 'lucide-react';
import { Route, Switch, Router as WouterRouter } from 'wouter';
import NotFound from '@/pages/not-found';
import {
  AnalysisError,
  analyzeWorkbook,
  buildWorkbook,
  formatDate,
  formatNumber,
  type AnalysisResult,
  type DailyMetric,
  type FoodcourtMetric,
  type UserDailyRow,
  type WeekdayMetric,
} from '@/lib/analysis';

type ProcessState = 'empty' | 'parsing' | 'success' | 'invalid';
type ViewKey = 'overview' | 'daily' | 'foodcourts' | 'users' | 'audit' | 'weekday';

const navItems: { key: ViewKey; label: string; icon: typeof Layers3; target: string }[] = [
  { key: 'overview', label: 'Overview', icon: Layers3, target: 'overview-section' },
  { key: 'daily', label: 'Daily metrics', icon: Table2, target: 'daily-section' },
  { key: 'foodcourts', label: 'Foodcourt validation', icon: Building2, target: 'foodcourt-section' },
  { key: 'users', label: 'User metrics', icon: Files, target: 'users-section' },
  { key: 'audit', label: 'Audit queue', icon: AlertTriangle, target: 'audit-section' },
  { key: 'weekday', label: 'Weekday summary', icon: History, target: 'weekday-section' },
];

const numeric = (value: number) => formatNumber(value);

function AppShell({ children }: { children: ReactNode }) {
  return (
    <div className="min-h-[100dvh] bg-[hsl(var(--background))] text-[hsl(var(--foreground))]">
      <aside className="fixed inset-y-0 left-0 z-20 hidden w-[248px] flex-col bg-[hsl(var(--sidebar))] px-5 py-6 text-[hsl(var(--sidebar-foreground))] md:flex">
        <div className="flex items-center gap-3 px-2">
          <div className="grid size-10 place-items-center rounded-xl bg-[hsl(var(--accent))] text-[hsl(var(--accent-foreground))] shadow-[0_8px_22px_rgba(224,165,54,.16)]">
            <ShieldCheck size={21} strokeWidth={2.5} />
          </div>
          <div>
            <p className="font-display text-[17px] font-bold tracking-tight">SmartQ</p>
            <p className="font-mono text-[10px] uppercase tracking-[.18em] text-[hsl(var(--sidebar-foreground)/.58)]">Credit analysis</p>
          </div>
        </div>
        <div className="mt-12 px-2">
          <p className="font-mono text-[10px] uppercase tracking-[.2em] text-[hsl(var(--sidebar-foreground)/.42)]">Workbench</p>
          <nav className="mt-3 space-y-1" aria-label="Analysis sections">
            {navItems.map(({ key, label, icon: Icon, target }) => (
              <a
                href={`#${target}`}
                key={key}
                data-testid={`link-${key}`}
                className="group flex items-center gap-3 rounded-lg px-3 py-2.5 text-[13px] text-[hsl(var(--sidebar-foreground)/.64)] transition-colors hover:bg-[hsl(var(--sidebar-foreground)/.08)] hover:text-[hsl(var(--sidebar-foreground))] focus-visible:text-[hsl(var(--sidebar-foreground))]"
              >
                <Icon size={16} strokeWidth={1.8} className="transition-transform group-hover:translate-x-0.5" />
                <span>{label}</span>
                {key === 'audit' && <span className="ml-auto size-1.5 rounded-full bg-[hsl(var(--accent))]" />}
              </a>
            ))}
          </nav>
        </div>
        <div className="mt-auto rounded-xl border border-[hsl(var(--sidebar-foreground)/.12)] bg-[hsl(var(--sidebar-foreground)/.05)] p-4">
          <div className="flex items-center gap-2 text-[hsl(var(--accent))]">
            <ShieldCheck size={15} />
            <span className="font-mono text-[10px] uppercase tracking-[.12em]">Local by design</span>
          </div>
          <p className="mt-2 text-[12px] leading-5 text-[hsl(var(--sidebar-foreground)/.58)]">
            Files are read in this browser and never uploaded to a server.
          </p>
        </div>
      </aside>
      <div className="md:pl-[248px]">{children}</div>
    </div>
  );
}

function MobileHeader() {
  return (
    <div className="flex items-center justify-between border-b border-[hsl(var(--border))] bg-[hsl(var(--card)/.92)] px-5 py-4 md:hidden">
      <div className="flex items-center gap-2.5">
        <div className="grid size-8 place-items-center rounded-lg bg-[hsl(var(--accent))] text-[hsl(var(--accent-foreground))]"><ShieldCheck size={17} /></div>
        <span className="font-display font-bold">SmartQ</span>
      </div>
      <span className="font-mono text-[10px] uppercase tracking-[.13em] text-[hsl(var(--muted-foreground))]">Credit analysis</span>
    </div>
  );
}

function TopBar({ filename, onReset, onDownload, canDownload }: { filename: string | null; onReset: () => void; onDownload: () => void; canDownload: boolean }) {
  return (
    <header className="sticky top-0 z-10 flex min-h-[73px] items-center justify-between gap-4 border-b border-[hsl(var(--border))] bg-[hsl(var(--background)/.9)] px-5 backdrop-blur-xl md:px-10">
      <div className="min-w-0">
        <p className="font-mono text-[10px] uppercase tracking-[.2em] text-[hsl(var(--muted-foreground))]">Operations / weekly review</p>
        <p data-testid="text-source-file" className="mt-1 truncate text-[13px] font-medium text-[hsl(var(--foreground)/.72)]">{filename ? filename : 'No report loaded'}</p>
      </div>
      <div className="flex shrink-0 items-center gap-2">
        {filename && <button type="button" onClick={onReset} data-testid="button-new-report" className="hidden items-center gap-2 rounded-lg border border-[hsl(var(--border))] bg-[hsl(var(--card))] px-3 py-2 text-[12px] font-semibold text-[hsl(var(--foreground)/.75)] transition-all hover:-translate-y-px hover:border-[hsl(var(--primary)/.4)] sm:flex"><RefreshCw size={14} /> New report</button>}
        <button type="button" onClick={onDownload} disabled={!canDownload} data-testid="button-download-header" className="flex items-center gap-2 rounded-lg bg-[hsl(var(--primary))] px-3 py-2 text-[12px] font-semibold text-[hsl(var(--primary-foreground))] shadow-sm transition-all hover:-translate-y-px hover:bg-[hsl(var(--primary)/.9)] disabled:cursor-not-allowed disabled:opacity-35"><ArrowDownToLine size={14} /> <span className="hidden sm:inline">Download workbook</span><span className="sm:hidden">Export</span></button>
      </div>
    </header>
  );
}

function EmptyState({ onPick }: { onPick: () => void }) {
  return (
    <section data-testid="status-empty" className="animate-rise-in mx-auto max-w-[1060px] px-5 pb-20 pt-12 md:px-10 md:pt-20">
      <div className="max-w-2xl">
        <div className="mb-5 inline-flex items-center gap-2 rounded-full border border-[hsl(var(--primary)/.18)] bg-[hsl(var(--primary)/.06)] px-3 py-1.5 font-mono text-[10px] uppercase tracking-[.15em] text-[hsl(var(--primary))]"><span className="size-1.5 rounded-full bg-[hsl(var(--primary))]" /> Browser-only workbench</div>
        <h1 className="font-display text-[clamp(2.55rem,6vw,5.25rem)] font-bold leading-[.98] tracking-[-.055em] text-[hsl(var(--foreground))]">
          Find the signal<br /><span className="text-[hsl(var(--primary))]">before it becomes noise.</span>
        </h1>
        <p className="mt-6 max-w-xl text-[16px] leading-7 text-[hsl(var(--muted-foreground))]">Drop in your weekly credit extract. SmartQ cleans the messy edges, maps the day-by-day picture, and puts audit exceptions in plain sight.</p>
      </div>
      <button type="button" onClick={onPick} data-testid="button-upload-empty" className="group relative mt-12 flex w-full max-w-3xl flex-col items-center justify-center overflow-hidden rounded-2xl border border-dashed border-[hsl(var(--primary)/.35)] bg-[hsl(var(--card))] px-6 py-14 text-center shadow-[var(--shadow-soft)] transition-all hover:-translate-y-1 hover:border-[hsl(var(--accent))] hover:bg-[hsl(var(--card)/.8)]">
        <div className="absolute left-0 top-0 h-1 w-1/3 bg-[hsl(var(--accent))] transition-all group-hover:w-2/3" />
        <div className="grid size-14 place-items-center rounded-2xl bg-[hsl(var(--primary)/.1)] text-[hsl(var(--primary))] transition-transform group-hover:-translate-y-1"><UploadCloud size={26} strokeWidth={1.8} /></div>
        <p className="mt-5 font-display text-[19px] font-semibold">Choose a weekly report</p>
        <p className="mt-2 text-[13px] text-[hsl(var(--muted-foreground))]">or drag and drop it here</p>
        <span className="mt-5 rounded-md bg-[hsl(var(--secondary))] px-2.5 py-1 font-mono text-[10px] uppercase tracking-[.12em] text-[hsl(var(--muted-foreground))]">.xlsx · .xls · .csv</span>
      </button>
      <div className="mt-10 grid max-w-3xl gap-3 sm:grid-cols-3">
        <MiniContract icon={Database} label="Input" value="OrderLog worksheet" />
        <MiniContract icon={Table2} label="Required fields" value="4 columns" />
        <MiniContract icon={ShieldCheck} label="Data handling" value="Local only" />
      </div>
    </section>
  );
}

function MiniContract({ icon: Icon, label, value }: { icon: typeof Database; label: string; value: string }) {
  return <div className="rounded-xl border border-[hsl(var(--border))] bg-[hsl(var(--card)/.55)] px-4 py-3"><Icon size={15} className="text-[hsl(var(--primary))]" /><p className="mt-3 font-mono text-[10px] uppercase tracking-[.14em] text-[hsl(var(--muted-foreground))]">{label}</p><p className="mt-1 text-[13px] font-semibold">{value}</p></div>;
}

function ParsingState({ filename }: { filename: string }) {
  return (
    <section data-testid="status-parsing" className="mx-auto max-w-[1060px] px-5 pb-20 pt-12 md:px-10 md:pt-20">
      <div className="max-w-xl animate-rise-in">
        <div className="flex size-14 items-center justify-center rounded-2xl bg-[hsl(var(--primary)/.1)] text-[hsl(var(--primary))]"><LoaderCircle size={27} className="animate-spin" /></div>
        <p className="mt-7 font-mono text-[10px] uppercase tracking-[.2em] text-[hsl(var(--primary))]">Reading report</p>
        <h1 className="mt-3 font-display text-4xl font-bold tracking-[-.04em]">Making the numbers<br />reviewable.</h1>
        <p className="mt-4 text-[15px] text-[hsl(var(--muted-foreground))]"><span className="font-semibold text-[hsl(var(--foreground))]">{filename}</span> is being checked locally.</p>
      </div>
      <div className="mt-12 max-w-3xl space-y-3">
        {['Reading first worksheet', 'Cleaning dates and credit values', 'Building audit views'].map((label, index) => <div key={label} className="flex items-center gap-3 rounded-xl border border-[hsl(var(--border))] bg-[hsl(var(--card)/.6)] px-4 py-4 animate-rise-in" style={{ animationDelay: `${index * 90}ms` }}><div className="size-2 rounded-full bg-[hsl(var(--accent))] animate-pulse-line" /><span className="text-[13px] text-[hsl(var(--muted-foreground))]">{label}</span><div className="ml-auto h-1 w-20 overflow-hidden rounded-full bg-[hsl(var(--secondary))]"><div className="h-full w-2/3 rounded-full bg-[hsl(var(--primary))] animate-pulse-line" /></div></div>)}
      </div>
    </section>
  );
}

function InvalidState({ message, missingColumns, onRetry }: { message: string; missingColumns: string[]; onRetry: () => void }) {
  return (
    <section data-testid="status-invalid-file" className="mx-auto max-w-[1060px] px-5 pb-20 pt-12 md:px-10 md:pt-20">
      <div className="max-w-2xl rounded-2xl border border-[hsl(var(--destructive)/.25)] bg-[hsl(var(--card))] p-7 shadow-[var(--shadow-soft)] md:p-10">
        <div className="grid size-12 place-items-center rounded-xl bg-[hsl(var(--destructive)/.1)] text-[hsl(var(--destructive))]"><AlertTriangle size={23} /></div>
        <p className="mt-6 font-mono text-[10px] uppercase tracking-[.2em] text-[hsl(var(--destructive))]">Invalid file</p>
        <h1 className="mt-2 font-display text-3xl font-bold tracking-[-.04em]">This report needs a closer look.</h1>
        <p className="mt-3 text-[14px] leading-6 text-[hsl(var(--muted-foreground))]">{message}</p>
        {missingColumns.length > 0 && <div className="mt-6 rounded-xl bg-[hsl(var(--destructive)/.06)] p-4"><p className="font-mono text-[10px] uppercase tracking-[.15em] text-[hsl(var(--destructive))]">Missing required columns</p><div className="mt-3 flex flex-wrap gap-2">{missingColumns.map((column) => <span key={column} className="rounded-md border border-[hsl(var(--destructive)/.2)] px-2.5 py-1 text-[12px] font-semibold">{column}</span>)}</div></div>}
        <div className="mt-8 flex flex-wrap gap-3"><button type="button" onClick={onRetry} data-testid="button-retry-upload" className="flex items-center gap-2 rounded-lg bg-[hsl(var(--primary))] px-4 py-2.5 text-[13px] font-semibold text-[hsl(var(--primary-foreground))] transition-transform hover:-translate-y-px"><RefreshCw size={15} /> Try another file</button><div className="flex items-center gap-2 text-[12px] text-[hsl(var(--muted-foreground))]"><Info size={14} /> Accepted: .xlsx, .xls, or .csv</div></div>
      </div>
    </section>
  );
}

function KpiCard({ label, value, note, tone = 'default', testId }: { label: string; value: string; note: string; tone?: 'default' | 'warning'; testId: string }) {
  return <div data-testid={testId} className={`rounded-xl border bg-[hsl(var(--card))] p-5 shadow-[0_7px_20px_rgba(52,61,46,.035)] ${tone === 'warning' ? 'border-[hsl(var(--accent)/.48)]' : 'border-[hsl(var(--border))]'}`}><div className="flex items-start justify-between gap-3"><p className="font-mono text-[10px] uppercase tracking-[.14em] text-[hsl(var(--muted-foreground))]">{label}</p>{tone === 'warning' && <AlertTriangle size={15} className="text-[hsl(var(--accent-foreground))] fill-[hsl(var(--accent))]" />}</div><p className="mt-4 font-display text-3xl font-bold tracking-[-.04em]">{value}</p><p className="mt-1 text-[11px] text-[hsl(var(--muted-foreground))]">{note}</p></div>;
}

function SectionHeading({ eyebrow, title, count, id }: { eyebrow: string; title: string; count?: string; id: string }) {
  return <div id={id} className="mb-5 flex scroll-mt-28 items-end justify-between gap-4"><div><p className="font-mono text-[10px] uppercase tracking-[.18em] text-[hsl(var(--primary))]">{eyebrow}</p><h2 className="mt-1 font-display text-2xl font-bold tracking-[-.035em]">{title}</h2></div>{count && <span className="rounded-full bg-[hsl(var(--secondary))] px-2.5 py-1 font-mono text-[10px] text-[hsl(var(--muted-foreground))]">{count}</span>}</div>;
}

function DataTable({ headers, rows, emptyText, testId }: { headers: string[]; rows: (string | number)[][]; emptyText: string; testId: string }) {
  return <div data-testid={testId} className="overflow-hidden rounded-xl border border-[hsl(var(--border))] bg-[hsl(var(--card))]"><div className="overflow-x-auto"><table className="w-full min-w-[690px] border-collapse text-left"><thead><tr className="border-b border-[hsl(var(--border))] bg-[hsl(var(--secondary)/.45)]">{headers.map((header) => <th key={header} className="whitespace-nowrap px-4 py-3 font-mono text-[9px] font-medium uppercase tracking-[.12em] text-[hsl(var(--muted-foreground))]">{header}</th>)}</tr></thead><tbody>{rows.length ? rows.map((row, rowIndex) => <tr key={`${testId}-${rowIndex}`} data-testid={`row-${testId}-${rowIndex}`} className="border-b border-[hsl(var(--border)/.7)] last:border-0 transition-colors hover:bg-[hsl(var(--secondary)/.26)]">{row.map((cell, cellIndex) => <td key={`${rowIndex}-${cellIndex}`} className={`whitespace-nowrap px-4 py-3 text-[12px] ${cellIndex === 0 ? 'font-medium' : 'text-[hsl(var(--muted-foreground))]'}`}>{cell}</td>)}</tr>) : <tr><td colSpan={headers.length} className="px-4 py-10 text-center text-[13px] text-[hsl(var(--muted-foreground))]">{emptyText}</td></tr>}</tbody></table></div></div>;
}

function AuditTable({ rows, emptyText, testId }: { rows: UserDailyRow[]; emptyText: string; testId: string }) {
  return <DataTable testId={testId} emptyText={emptyText} headers={['Date', 'Foodcourt', 'User', 'User type', 'Daily credits', 'Status']} rows={rows.map((row) => [formatDate(row.Date), row.Foodcourt, row.User || 'Unnamed user', row['User Type'] || '—', numeric(row['Daily User Credits']), 'REVIEW'])} />;
}

function Overview({ result, onJump }: { result: AnalysisResult; onJump: (target: string) => void }) {
  const exceptionCount = result.over200.length + result.vendorCreditUsers.length;
  const dailyMax = Math.max(...result.dailyMetrics.map((row) => row['Total Credits']), 0);
  return (
    <>
      <section id="overview-section" className="scroll-mt-28 animate-rise-in">
        <div className="flex flex-col justify-between gap-5 lg:flex-row lg:items-end">
          <div><p className="font-mono text-[10px] uppercase tracking-[.2em] text-[hsl(var(--primary))]">Weekly signal</p><h1 className="mt-2 font-display text-[clamp(2rem,4vw,3.4rem)] font-bold leading-[1] tracking-[-.055em]">A clear read on<br /><span className="text-[hsl(var(--primary))]">credit behaviour.</span></h1><p className="mt-4 max-w-xl text-[14px] leading-6 text-[hsl(var(--muted-foreground))]">Cleaned from {numeric(result.sourceRows)} rows across {numeric(result.sourceColumns)} columns. Review the exceptions first, then use the workbook as your audit trail.</p></div>
          <div className="flex shrink-0 items-center gap-2 rounded-lg border border-[hsl(var(--border))] bg-[hsl(var(--card)/.7)] px-3 py-2 text-[11px] text-[hsl(var(--muted-foreground))]"><CheckCircle2 size={15} className="text-[hsl(var(--primary))]" /> Analysis complete</div>
        </div>
        {result.notices.length > 0 && <div data-testid="status-format-notice" className="mt-7 space-y-2 rounded-xl border border-[hsl(var(--accent)/.45)] bg-[hsl(var(--accent)/.1)] px-4 py-3.5 text-[12px] leading-5 text-[hsl(var(--foreground)/.8)]">{result.notices.map((notice, index) => <div key={notice} className="flex items-start gap-3"><Info size={16} className="mt-0.5 shrink-0 text-[hsl(var(--accent-foreground))]" /><span data-testid={`text-format-notice-${index}`}>{notice}</span></div>)}</div>}
        {result.invalidDates > 0 && <div data-testid="status-invalid-dates" className="mt-7 flex items-start gap-3 rounded-xl border border-[hsl(var(--accent)/.45)] bg-[hsl(var(--accent)/.1)] px-4 py-3.5 text-[12px] leading-5 text-[hsl(var(--foreground)/.8)]"><Info size={16} className="mt-0.5 shrink-0 text-[hsl(var(--accent-foreground))]" /><span><strong>{numeric(result.invalidDates)} rows</strong> have an invalid or missing Date and were excluded from date-based metrics.</span></div>}
        <div className="mt-8 grid gap-3 sm:grid-cols-2 xl:grid-cols-6">
          <KpiCard label="Total users" value={numeric(result.totalUsers)} note="Distinct users in file" testId="metric-total-users" />
          <KpiCard label="Credit users" value={numeric(result.creditUsers)} note="Users with credits > 0" testId="metric-credit-users" />
          <KpiCard label="Total credits" value={numeric(result.totalCredits)} note="Across all source rows" testId="metric-total-credits" />
          <KpiCard label="Foodcourts" value={numeric(result.foodcourtCount)} note="Locations detected" testId="metric-foodcourts" />
          <KpiCard label="Users >200" value={numeric(new Set(result.over200.map((row) => row.User)).size)} note="Daily threshold exceptions" tone={result.over200.length ? 'warning' : 'default'} testId="metric-over-200" />
          <KpiCard label="Vendor credit users" value={numeric(new Set(result.vendorCreditUsers.map((row) => row.User)).size)} note="VendorNoCredit with usage" tone={result.vendorCreditUsers.length ? 'warning' : 'default'} testId="metric-vendor-users" />
        </div>
        <div className={`mt-5 flex flex-col gap-4 rounded-xl border p-5 sm:flex-row sm:items-center sm:justify-between ${exceptionCount ? 'border-[hsl(var(--accent)/.52)] bg-[hsl(var(--accent)/.12)]' : 'border-[hsl(var(--primary)/.2)] bg-[hsl(var(--primary)/.07)]'}`} data-testid="status-audit-summary">
          <div className="flex items-start gap-3"><div className={`mt-0.5 grid size-8 shrink-0 place-items-center rounded-lg ${exceptionCount ? 'bg-[hsl(var(--accent))] text-[hsl(var(--accent-foreground))]' : 'bg-[hsl(var(--primary))] text-[hsl(var(--primary-foreground))]'}`}>{exceptionCount ? <AlertTriangle size={16} /> : <Check size={17} />}</div><div><p className="text-[13px] font-semibold">{exceptionCount ? `${numeric(exceptionCount)} audit ${exceptionCount === 1 ? 'exception needs' : 'exceptions need'} review` : 'No audit exceptions found'}</p><p className="mt-1 text-[12px] text-[hsl(var(--muted-foreground))]">{exceptionCount ? 'The queue below is sorted by date and daily credit impact.' : 'Thresholds checked: over 200 credits per day and VendorNoCredit usage.'}</p></div></div>{exceptionCount > 0 && <button type="button" onClick={() => onJump('audit-section')} data-testid="button-review-exceptions" className="flex items-center gap-2 self-start rounded-lg bg-[hsl(var(--foreground))] px-3 py-2 text-[12px] font-semibold text-[hsl(var(--card))] transition-transform hover:-translate-y-px sm:self-auto">Review queue <ArrowRight size={14} /></button>}</div>
      </section>
      <section className="mt-14">
        <SectionHeading id="daily-section" eyebrow="At a glance" title="Daily rhythm" count={`${result.dailyMetrics.length} active days`} />
        <div className="grid gap-5 xl:grid-cols-[1.45fr_.85fr]">
          <DataTable testId="table-daily-metrics" emptyText="No valid date data was found." headers={['Date', 'Day', 'Users', 'Credit users', 'Credits', 'Avg / user', 'Orders']} rows={result.dailyMetrics.map((row: DailyMetric) => [formatDate(row.Date), row['Week Day'].slice(0, 3), row['Total Users'], row['Credit Users'], numeric(row['Total Credits']), numeric(row['Avg Credit/User']), row['Total Orders']])} />
          <div className="rounded-xl border border-[hsl(var(--border))] bg-[hsl(var(--card))] p-5"><div className="flex items-center justify-between"><div><p className="font-mono text-[10px] uppercase tracking-[.15em] text-[hsl(var(--muted-foreground))]">Volume map</p><p className="mt-1 text-[13px] font-semibold">Credits by day</p></div><span className="font-mono text-[10px] text-[hsl(var(--muted-foreground))]">max {numeric(dailyMax)}</span></div><div className="mt-7 space-y-4">{result.dailyMetrics.length ? result.dailyMetrics.map((row) => <div key={row.Date.toISOString()} data-testid={`bar-day-${row['Week Day']}`}><div className="mb-1.5 flex justify-between text-[11px]"><span className="text-[hsl(var(--muted-foreground))]">{row['Week Day'].slice(0, 3)} <span className="font-mono text-[10px] opacity-60">{formatDate(row.Date).slice(0, 6)}</span></span><span className="font-mono font-medium">{numeric(row['Total Credits'])}</span></div><div className="h-2 overflow-hidden rounded-full bg-[hsl(var(--secondary))]"><div className="h-full rounded-full bg-[hsl(var(--primary))] transition-all duration-700" style={{ width: `${dailyMax ? Math.max(4, (row['Total Credits'] / dailyMax) * 100) : 0}%` }} /></div></div>) : <p className="py-8 text-center text-[12px] text-[hsl(var(--muted-foreground))]">Nothing to chart yet.</p>}</div></div>
        </div>
      </section>
    </>
  );
}

function UsersSection({ result }: { result: AnalysisResult }) {
  return <section id="users-section" className="mt-14 scroll-mt-28"><SectionHeading id="users-heading" eyebrow="Who is using credit" title="User metrics" count={`${result.userMetrics.length} user / court records`} /><DataTable testId="table-user-metrics" emptyText="No user metrics are available." headers={['User', 'Foodcourt', 'User type', 'Days used', 'Total credits', 'Max / day', '>200 days', 'Vendor days', 'Avg / day']} rows={result.userMetrics.map((row) => [row.User || 'Unnamed user', row.Foodcourt, row['User Type'] || '—', row['Days Used'], numeric(row['Total Credits']), numeric(row['Max Credits/Day']), row['>200 Days'], row['Vendor Credit Days'], numeric(row['Avg Credits/Day'])])} /></section>;
}

function FoodcourtSection({ result }: { result: AnalysisResult }) {
  const [selectedFoodcourt, setSelectedFoodcourt] = useState('all');
  const visibleMetrics = selectedFoodcourt === 'all'
    ? result.foodcourtMetrics
    : result.foodcourtMetrics.filter((row) => row.Foodcourt === selectedFoodcourt);
  const reviewCount = visibleMetrics.filter((row) => row['Validation Status'] === 'REVIEW').length;

  return (
    <section id="foodcourt-section" className="mt-14 scroll-mt-28">
      <div className="flex flex-col justify-between gap-4 sm:flex-row sm:items-end">
        <SectionHeading id="foodcourt-heading" eyebrow="Validate by location" title="Foodcourt validation" count={`${result.foodcourtCount} locations`} />
        <label className="flex items-center gap-2 text-[11px] text-[hsl(var(--muted-foreground))]">
          <span className="font-mono text-[10px] uppercase tracking-[.12em]">View</span>
          <select value={selectedFoodcourt} onChange={(event) => setSelectedFoodcourt(event.target.value)} data-testid="select-foodcourt" className="rounded-lg border border-[hsl(var(--border))] bg-[hsl(var(--card))] px-3 py-2 text-[12px] outline-none focus:border-[hsl(var(--primary)/.5)]">
            <option value="all">All foodcourts</option>
            {result.foodcourtMetrics.map((row) => <option key={row.Foodcourt} value={row.Foodcourt}>{row.Foodcourt}</option>)}
          </select>
        </label>
      </div>
      <p className="mb-5 max-w-2xl text-[13px] leading-6 text-[hsl(var(--muted-foreground))]">Each location is validated independently, so a Bengaluru, Hyderabad, or Chennai issue can be isolated without searching through the full user population.</p>
      <div className={`mb-4 rounded-xl border px-4 py-3 text-[12px] ${reviewCount ? 'border-[hsl(var(--accent)/.52)] bg-[hsl(var(--accent)/.12)]' : 'border-[hsl(var(--primary)/.2)] bg-[hsl(var(--primary)/.07)]'}`} data-testid="status-foodcourt-validation">
        {reviewCount ? <span><strong>{numeric(reviewCount)} {reviewCount === 1 ? 'foodcourt needs' : 'foodcourts need'} review</strong> in this view. Open the audit queue to see the exact users and dates.</span> : <span><strong>All foodcourts are clear</strong> in this view. No over-200 or VendorNoCredit usage flags were found.</span>}
      </div>
      <DataTable
        testId="table-foodcourt-metrics"
        emptyText="No foodcourt data is available."
        headers={['Foodcourt', 'Active days', 'Users', 'Credit users', 'Credits', 'Avg / user', 'Users >200', 'Vendor users', 'Review rows', 'Orders', 'Status']}
        rows={visibleMetrics.map((row: FoodcourtMetric) => [row.Foodcourt, row['Active Days'], row['Total Users'], row['Credit Users'], numeric(row['Total Credits']), numeric(row['Avg Credit/User']), row['Users >200'], row['Vendor Credit Users'], row['Review Rows'], row['Total Orders'], row['Validation Status']])}
      />
    </section>
  );
}

function AuditSection({ result }: { result: AnalysisResult }) {
  const [search, setSearch] = useState('');
  const [foodcourt, setFoodcourt] = useState('all');
  const matches = (row: UserDailyRow) =>
    (foodcourt === 'all' || row.Foodcourt === foodcourt) &&
    `${row.Foodcourt} ${row.User} ${row['User Type']} ${row['Week Day']}`.toLowerCase().includes(search.toLowerCase());
  const over200 = result.over200.filter(matches);
  const vendor = result.vendorCreditUsers.filter(matches);
  return <section id="audit-section" className="mt-14 scroll-mt-28"><div className="flex flex-col justify-between gap-4 sm:flex-row sm:items-end"><SectionHeading id="audit-heading" eyebrow="Exceptions first" title="Audit queue" count={`${result.over200.length + result.vendorCreditUsers.length} flags`} /><div className="flex flex-col gap-2 sm:flex-row"><select value={foodcourt} onChange={(event) => setFoodcourt(event.target.value)} data-testid="select-audit-foodcourt" aria-label="Filter audit queue by foodcourt" className="rounded-lg border border-[hsl(var(--border))] bg-[hsl(var(--card))] px-3 py-2 text-[12px] outline-none focus:border-[hsl(var(--primary)/.5)]"><option value="all">All foodcourts</option>{result.foodcourtMetrics.map((row) => <option key={row.Foodcourt} value={row.Foodcourt}>{row.Foodcourt}</option>)}</select><div className="relative"><Search size={14} className="pointer-events-none absolute left-3 top-2.5 text-[hsl(var(--muted-foreground))]" /><input value={search} onChange={(event) => setSearch(event.target.value)} data-testid="input-audit-search" aria-label="Filter audit queue" placeholder="Filter users…" className="w-full rounded-lg border border-[hsl(var(--border))] bg-[hsl(var(--card))] py-2 pl-9 pr-3 text-[12px] outline-none transition-colors placeholder:text-[hsl(var(--muted-foreground)/.7)] focus:border-[hsl(var(--primary)/.5)] sm:w-48" /></div></div></div><div className="grid gap-5 xl:grid-cols-2"><div><div className="mb-3 flex items-center gap-2"><div className="size-2 rounded-full bg-[hsl(var(--accent))]" /><h3 className="text-[13px] font-semibold">Over 200 credits in a day</h3><span className="font-mono text-[10px] text-[hsl(var(--muted-foreground))]">{over200.length}</span></div><AuditTable testId="table-over-200" rows={over200} emptyText={search || foodcourt !== 'all' ? 'No matching threshold exceptions.' : 'No users exceeded 200 credits.'} /></div><div><div className="mb-3 flex items-center gap-2"><div className="size-2 rounded-full bg-[hsl(var(--destructive))]" /><h3 className="text-[13px] font-semibold">VendorNoCredit with usage</h3><span className="font-mono text-[10px] text-[hsl(var(--muted-foreground))]">{vendor.length}</span></div><AuditTable testId="table-vendor-credit" rows={vendor} emptyText={search || foodcourt !== 'all' ? 'No matching vendor exceptions.' : 'No VendorNoCredit users used credits.'} /></div></div></section>;
}

function WeekdaySection({ result }: { result: AnalysisResult }) {
  return <section id="weekday-section" className="mt-14 scroll-mt-28"><SectionHeading id="weekday-heading" eyebrow="Pattern by weekday" title="Weekday summary" count="Mon — Sun" /><DataTable testId="table-weekday-summary" emptyText="No weekday data is available." headers={['Weekday', 'Unique users', 'Credit users', 'Total credits', 'Avg / credit user', 'Users >200', 'Vendor users']} rows={result.weekdaySummary.map((row: WeekdayMetric) => [row['Week Day'], row['Unique Users'], row['Credit Users'], numeric(row['Total Credits']), numeric(row['Avg Credits/Credit User']), row['Users >200'], row['Vendor Credit Users']])} /></section>;
}

function SuccessView({ result, onJump }: { result: AnalysisResult; onJump: (target: string) => void }) {
  return <main data-testid="status-success" className="mx-auto max-w-[1300px] px-5 pb-24 pt-9 md:px-10 md:pt-12"><Overview result={result} onJump={onJump} /><FoodcourtSection result={result} /><UsersSection result={result} /><AuditSection result={result} /><WeekdaySection result={result} /><div className="mt-14 flex flex-col justify-between gap-5 rounded-2xl bg-[hsl(var(--sidebar))] p-6 text-[hsl(var(--sidebar-foreground))] sm:flex-row sm:items-center md:p-8"><div><div className="flex items-center gap-2 text-[hsl(var(--accent))]"><ArrowDownToLine size={17} /><p className="font-mono text-[10px] uppercase tracking-[.18em]">Take it with you</p></div><h2 className="mt-3 font-display text-2xl font-bold tracking-[-.03em]">Your audit trail is ready.</h2><p className="mt-2 max-w-lg text-[13px] leading-5 text-[hsl(var(--sidebar-foreground)/.58)]">The workbook includes overall metrics plus separate foodcourt validation and audit sheets.</p></div><button type="button" onClick={() => onJump('download')} data-testid="button-download-footer" className="flex shrink-0 items-center justify-center gap-2 rounded-lg bg-[hsl(var(--accent))] px-4 py-3 text-[13px] font-bold text-[hsl(var(--accent-foreground))] transition-transform hover:-translate-y-px">Download workbook <ArrowDownToLine size={15} /></button></div><div id="download" className="pt-6 text-center font-mono text-[10px] uppercase tracking-[.13em] text-[hsl(var(--muted-foreground))]">Processed locally · Nothing leaves this browser</div></main>;
}

function Home() {
  const [status, setStatus] = useState<ProcessState>('empty');
  const [result, setResult] = useState<AnalysisResult | null>(null);
  const [filename, setFilename] = useState<string | null>(null);
  const [errorMessage, setErrorMessage] = useState('');
  const [missingColumns, setMissingColumns] = useState<string[]>([]);
  const [dragging, setDragging] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  const reset = () => {
    setStatus('empty');
    setResult(null);
    setFilename(null);
    setErrorMessage('');
    setMissingColumns([]);
    setDragging(false);
  };

  const handleFile = async (file?: File) => {
    if (!file) return;
    const normalizedName = file.name.trim().toLowerCase().split(/[?#]/, 1)[0];
    const excelMimeTypes = new Set([
      'application/vnd.ms-excel',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'application/octet-stream',
    ]);
    const hasCsvExtension = /\.csv$/.test(normalizedName);
    const hasCsvMime = file.type.toLowerCase() === 'text/csv';
    const hasExcelExtension = /\.(xlsx|xls)$/.test(normalizedName);
    const hasExcelMime = excelMimeTypes.has(file.type.toLowerCase());

    // Some upload sources omit the filename extension and browsers often report
    // Excel files as application/octet-stream. Let those files through so the
    // workbook parser can make the final determination.
    if (!hasExcelExtension && !hasExcelMime && !hasCsvExtension && !hasCsvMime && file.type) {
      setFilename(file.name);
      setStatus('invalid');
      setErrorMessage('SmartQ accepts Excel workbooks and CSV files with .xlsx, .xls, or .csv formats.');
      setMissingColumns([]);
      return;
    }
    setFilename(file.name);
    setStatus('parsing');
    setErrorMessage('');
    try {
      const isCsv = hasCsvExtension || hasCsvMime;
      const analysis = await analyzeWorkbook(
        isCsv ? await file.text() : await file.arrayBuffer(),
        isCsv ? 'csv' : 'excel',
        file.name,
      );
      setResult(analysis);
      setStatus('success');
    } catch (error) {
      const analysisError = error instanceof AnalysisError ? error : null;
      setErrorMessage(analysisError?.message ?? 'The workbook could not be read. It may be damaged or protected.');
      setMissingColumns(analysisError?.missingColumns ?? []);
      setStatus('invalid');
    }
  };

  const onFileChange = (event: ChangeEvent<HTMLInputElement>) => {
    void handleFile(event.target.files?.[0]);
    event.target.value = '';
  };

  const onDrop = (event: DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setDragging(false);
    void handleFile(event.dataTransfer.files?.[0]);
  };

  const download = () => {
    if (!result) return;
    const workbook = buildWorkbook(result);
    const blob = new Blob([workbook], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = `Weekly_Credit_Analysis_${new Date().toISOString().replace(/\D/g, '').slice(0, 15)}.xlsx`;
    anchor.click();
    URL.revokeObjectURL(url);
  };

  const jump = (target: string) => document.getElementById(target)?.scrollIntoView({ behavior: 'smooth', block: 'start' });

  return (
    <AppShell>
      <MobileHeader />
      <TopBar filename={filename} onReset={reset} onDownload={download} canDownload={status === 'success'} />
      <input ref={inputRef} onChange={onFileChange} type="file" accept=".xlsx,.xls,.csv,text/csv,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" className="hidden" data-testid="input-file-upload" />
      {status === 'empty' && <div onDragOver={(event) => { event.preventDefault(); setDragging(true); }} onDragLeave={() => setDragging(false)} onDrop={onDrop} className={dragging ? 'bg-[hsl(var(--primary)/.04)]' : ''}><EmptyState onPick={() => inputRef.current?.click()} /></div>}
      {status === 'parsing' && filename && <ParsingState filename={filename} />}
      {status === 'invalid' && <InvalidState message={errorMessage} missingColumns={missingColumns} onRetry={() => { reset(); setTimeout(() => inputRef.current?.click(), 0); }} />}
      {status === 'success' && result && <SuccessView result={result} onJump={jump} />}
    </AppShell>
  );
}

function Router() {
  return <Switch><Route path="/" component={Home} /><Route component={NotFound} /></Switch>;
}

function App() {
  return <WouterRouter base={import.meta.env.BASE_URL.replace(/\/$/, '')}><Router /></WouterRouter>;
}

export default App;