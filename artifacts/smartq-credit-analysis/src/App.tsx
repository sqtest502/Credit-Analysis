import { useEffect, useRef, useState, type ChangeEvent, type DragEvent, type ReactNode } from 'react';
import {
  AlertTriangle,
  ArrowDownToLine,
  ArrowRight,
  Building2,
  Check,
  CheckCircle2,
  ChevronLeft,
  ChevronRight,
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
  ArrowUpDown,
  Table2,
  UploadCloud,
} from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Badge } from '@/components/ui/badge';
import { Card } from '@/components/ui/card';
import { Route, Switch, Router as WouterRouter } from 'wouter';
import {
  Bar,
  BarChart,
  CartesianGrid,
  Line,
  LineChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from 'recharts';
import NotFound from '@/pages/not-found';
import {
  AnalysisError,
  WEEKDAYS,
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

const analyzeInWorker = (input: ArrayBuffer | string, format: 'excel' | 'csv', sourceName: string) => new Promise<AnalysisResult>((resolve, reject) => {
  const worker = new Worker(new URL('./lib/analysis.worker.ts', import.meta.url), { type: 'module' });
  worker.onmessage = (event: MessageEvent<{ type: 'success'; result: AnalysisResult } | { type: 'error'; message: string; missingColumns: string[] }>) => {
    worker.terminate();
    if (event.data.type === 'success') resolve(event.data.result);
    else reject(new AnalysisError(event.data.message, event.data.missingColumns));
  };
  worker.onerror = (event) => {
    worker.terminate();
    void analyzeWorkbook(input instanceof ArrayBuffer ? input.slice(0) : input, format, sourceName)
      .then(resolve)
      .catch((error) => reject(error instanceof Error ? error : new Error(event.message || 'The workbook could not be analyzed.')));
  };
  const request = { input, format, sourceName };
  if (input instanceof ArrayBuffer) worker.postMessage(request, [input]);
  else worker.postMessage(request);
});

const dateKeyFor = (date: Date) => `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;

function dailyMetricsForFoodcourt(result: AnalysisResult, foodcourt: string): DailyMetric[] {
  if (foodcourt === 'all') return result.dailyMetrics;
  const groups = new Map<string, UserDailyRow[]>();
  result.userDaily.filter((row) => row.Foodcourt === foodcourt).forEach((row) => {
    const key = row.dateKey;
    const group = groups.get(key);
    if (group) group.push(row);
    else groups.set(key, [row]);
  });
  const ordersByDate = new Map<string, number>();
  result.rawData.forEach((row) => {
    if (row.Foodcourt !== foodcourt || !row.Date) return;
    const dateKey = dateKeyFor(row.Date);
    ordersByDate.set(dateKey, (ordersByDate.get(dateKey) ?? 0) + 1);
  });
  return [...groups.entries()].sort(([left], [right]) => left.localeCompare(right)).map(([dateKey, rows]) => {
    const creditUsers = new Set(rows.filter((row) => row['Daily User Credits'] > 0).map((row) => row.User));
    const totalCredits = rows.reduce((sum, row) => sum + row['Daily User Credits'], 0);
    return {
      Date: rows[0].Date,
      'Week Day': rows[0]['Week Day'],
      'Total Users': new Set(rows.map((row) => row.User)).size,
      'Credit Users': creditUsers.size,
      'Total Credits': totalCredits,
      'Avg Credit/User': creditUsers.size ? totalCredits / creditUsers.size : 0,
      'Users >200': new Set(rows.filter((row) => row['Over 200']).map((row) => row.User)).size,
      'Vendor Credit Users': new Set(rows.filter((row) => row['Vendor Used Credit']).map((row) => row.User)).size,
      'Total Orders': ordersByDate.get(dateKey) ?? 0,
    };
  });
}

function AppShell({ children }: { children: ReactNode }) {
  return (
    <div className="min-h-[100dvh] bg-[hsl(var(--background))] text-[hsl(var(--foreground))]">
      <aside className="fixed inset-y-0 left-0 z-20 hidden w-[276px] flex-col border-r border-white/10 bg-[linear-gradient(180deg,_#101827_0%,_#121b2d_100%)] px-6 py-7 text-[hsl(var(--sidebar-foreground))] shadow-[0_18px_45px_rgba(15,23,42,0.28)] md:flex">
        <div className="px-1">
          <img src="/smartq-logo.avif" alt="SmartQ, a Compass Group company" className="h-auto w-full max-w-[210px] object-contain" />
        </div>
        <div className="mt-12 px-2">
          <p className="font-mono text-[10px] uppercase tracking-[.2em] text-slate-400">Workbench</p>
          <nav className="mt-3 space-y-1.5" aria-label="Analysis sections">
            {navItems.map(({ key, label, icon: Icon, target }) => (
              <a
                href={`#${target}`}
                key={key}
                data-testid={`link-${key}`}
                className="group flex items-center gap-3 rounded-xl border border-transparent bg-white/0 px-3 py-2.5 text-[13px] font-medium text-slate-200 transition-all duration-200 hover:border-white/10 hover:bg-white/5 hover:text-white"
              >
                <Icon size={16} strokeWidth={1.8} className="transition-transform group-hover:translate-x-0.5 text-sky-300" />
                <span>{label}</span>
                {key === 'audit' && <span className="ml-auto size-1.5 rounded-full bg-[hsl(var(--accent))]" />}
              </a>
            ))}
          </nav>
        </div>
        <div className="mt-auto rounded-2xl border border-white/10 bg-white/5 p-4">
          <div className="flex items-center gap-2 text-[hsl(var(--accent))]">
            <ShieldCheck size={15} />
            <span className="font-mono text-[10px] uppercase tracking-[.12em]">Support-ready by design</span>
          </div>
          <p className="mt-2 text-[12px] leading-5 text-slate-300">
            Diagnose credit issues faster, keep customer data private, and hand off a clear audit trail when support needs it.
          </p>
        </div>
      </aside>
      <div className="md:pl-[276px]">{children}</div>
    </div>
  );
}

function MobileHeader() {
  return (
    <div className="flex items-center justify-between border-b border-[hsl(var(--border))] bg-[hsl(var(--card)/.92)] px-5 py-4 md:hidden">
      <div className="flex items-center">
        <img src="/smartq-logo.avif" alt="SmartQ, a Compass Group company" className="h-auto w-[142px] object-contain" />
      </div>
      <span className="font-mono text-[10px] uppercase tracking-[.13em] text-[hsl(var(--muted-foreground))]">Credit analysis</span>
    </div>
  );
}

function TopBar({ filename, onReset, onDownload, canDownload }: { filename: string | null; onReset: () => void; onDownload: () => void; canDownload: boolean }) {
  return (
    <header className="sticky top-0 z-10 flex min-h-[78px] items-center justify-between gap-4 border-b border-[hsl(var(--border))] bg-[hsl(var(--background)/.94)] px-5 backdrop-blur-xl md:px-12">
      <div className="min-w-0">
        <p className="font-mono text-[10px] uppercase tracking-[.2em] text-[hsl(var(--muted-foreground))]">Operations / weekly review</p>
        <div className="mt-1 flex items-center gap-2">
          <p data-testid="text-source-file" className="truncate text-[13px] font-medium text-[hsl(var(--foreground)/.72)]">{filename ? filename : 'No report loaded'}</p>
          {filename && <Badge variant="secondary" className="hidden text-[10px] sm:inline-flex">Live</Badge>}
        </div>
      </div>
      <div className="flex shrink-0 items-center gap-2">
        {filename && (
          <Button type="button" variant="outline" size="sm" onClick={onReset} data-testid="button-new-report" className="hidden sm:inline-flex">
            <RefreshCw size={14} /> New report
          </Button>
        )}
        <Button type="button" size="sm" onClick={onDownload} disabled={!canDownload} data-testid="button-download-header" className="inline-flex gap-2">
          <ArrowDownToLine size={14} />
          <span className="hidden sm:inline">Download workbook</span>
          <span className="sm:hidden">Export</span>
        </Button>
      </div>
    </header>
  );
}

function EmptyState({ onPick, dragging }: { onPick: () => void; dragging: boolean }) {
  return (
    <section data-testid="status-empty" className="launch-stage mx-auto max-w-[1060px] px-5 pb-20 pt-12 md:px-10 md:pt-20">
      <div className="launch-copy max-w-2xl">
        <Badge variant="secondary" className="mb-5 border-[hsl(var(--primary)/.15)] bg-[hsl(var(--primary)/.06)] text-[hsl(var(--primary))]">
          <span className="mr-2 size-1.5 rounded-full bg-[hsl(var(--primary))]" /> Browser-only workbench
        </Badge>
        <h1 className="launch-title font-display text-[clamp(2.55rem,6vw,5.25rem)] font-bold leading-[.98] tracking-[-.055em] text-[hsl(var(--foreground))]">
          Find the signal<br /><span className="launch-gradient bg-gradient-to-r from-[hsl(var(--primary))] to-[hsl(var(--accent))] bg-clip-text text-transparent">before it becomes noise.</span>
        </h1>
        <p className="launch-description mt-6 max-w-xl text-[16px] leading-7 text-[hsl(var(--muted-foreground))]">Drop in your weekly credit extract. SmartQ cleans the messy edges, maps the day-by-day picture, and puts audit exceptions in plain sight.</p>
      </div>
      <Card className={`launch-upload group relative mt-12 w-full max-w-3xl overflow-hidden border p-0 transition-all duration-300 hover:-translate-y-1 ${dragging ? 'border-[hsl(var(--accent))] bg-[hsl(var(--primary)/.06)] shadow-[var(--shadow-lifted)]' : 'border-[hsl(var(--border))] bg-[linear-gradient(180deg,_rgba(255,255,255,0.9),_rgba(248,250,252,0.92))] shadow-[var(--shadow-soft)]'}`}>
        <button type="button" onClick={onPick} data-testid="button-upload-empty" aria-label="Choose a weekly credit report" className="launch-upload-action group relative flex w-full flex-col items-center justify-center overflow-hidden rounded-xl px-6 py-14 text-center">
          <div className="absolute inset-x-0 top-0 h-1 bg-[linear-gradient(90deg,_hsl(var(--primary))_0%,_hsl(var(--accent))_100%)]" />
          <div className={`launch-upload-icon grid size-14 place-items-center rounded-2xl text-[hsl(var(--primary))] transition-transform group-hover:-translate-y-1 ${dragging ? 'bg-[hsl(var(--accent)/.28)]' : 'bg-[linear-gradient(135deg,_hsl(var(--primary)/.12),_hsl(var(--accent)/.14))]'}`}><UploadCloud size={26} strokeWidth={1.8} /></div>
          <p className="mt-5 font-display text-[19px] font-semibold">Choose a weekly report</p>
          <p className="mt-2 text-[13px] text-[hsl(var(--muted-foreground))]">{dragging ? 'Release to start the analysis' : 'or drag and drop it here'}</p>
          <div className="mt-5 flex items-center gap-2">
            <Badge variant="secondary" className="text-[10px] uppercase tracking-[.12em]">.xlsx</Badge>
            <Badge variant="secondary" className="text-[10px] uppercase tracking-[.12em]">.xls</Badge>
            <Badge variant="secondary" className="text-[10px] uppercase tracking-[.12em]">.csv</Badge>
          </div>
        </button>
      </Card>
      <div className="launch-features mt-10 grid max-w-3xl gap-3 sm:grid-cols-3">
        <MiniContract icon={Database} label="Input" value="OrderLog worksheet" />
        <MiniContract icon={Table2} label="Required fields" value="4 columns" />
        <MiniContract icon={ShieldCheck} label="Data handling" value="Local only" />
      </div>
      <div className="launch-steps mt-10 flex max-w-3xl flex-wrap items-center gap-x-5 gap-y-3 border-t border-[hsl(var(--border))] pt-5 text-[11px] text-[hsl(var(--muted-foreground))]">
        {['Upload', 'Review signals', 'Export workbook'].map((step, index) => <div key={step} className="flex items-center gap-2"><span className="grid size-5 place-items-center rounded-full bg-[hsl(var(--primary)/.1)] font-mono text-[10px] font-medium text-[hsl(var(--primary))]">{index + 1}</span><span>{step}</span>{index < 2 && <ArrowRight size={13} className="ml-2 opacity-40" />}</div>)}
      </div>
    </section>
  );
}

function MiniContract({ icon: Icon, label, value }: { icon: typeof Database; label: string; value: string }) {
  return <div className="launch-feature rounded-xl border border-[hsl(var(--border))] bg-[hsl(var(--card)/.55)] px-4 py-3"><Icon size={15} className="text-[hsl(var(--primary))]" /><p className="mt-3 font-mono text-[10px] uppercase tracking-[.14em] text-[hsl(var(--muted-foreground))]">{label}</p><p className="mt-1 text-[13px] font-semibold">{value}</p></div>;
}

function ParsingState({ filename }: { filename: string }) {
  const [activeStep, setActiveStep] = useState(0);
  const steps = ['Reading first worksheet', 'Cleaning dates and credit values', 'Building audit views'];

  useEffect(() => {
    const timer = window.setInterval(() => setActiveStep((step) => (step + 1) % steps.length), 900);
    return () => window.clearInterval(timer);
  }, [steps.length]);

  return (
    <section data-testid="status-parsing" className="mx-auto max-w-[1060px] px-5 pb-20 pt-12 md:px-10 md:pt-20">
      <div className="max-w-xl animate-rise-in">
        <div className="relative flex size-14 items-center justify-center rounded-2xl bg-[hsl(var(--primary)/.1)] text-[hsl(var(--primary))]"><span className="absolute inset-0 rounded-2xl border-2 border-[hsl(var(--primary)/.2)] animate-analysis-ring" /><LoaderCircle size={27} className="animate-spin" /></div>
        <p className="mt-7 font-mono text-[10px] uppercase tracking-[.2em] text-[hsl(var(--primary))]">Reading report</p>
        <h1 className="mt-3 font-display text-4xl font-bold tracking-[-.04em]">Making the numbers<br />reviewable.</h1>
        <p className="mt-4 text-[15px] text-[hsl(var(--muted-foreground))]"><span className="font-semibold text-[hsl(var(--foreground))]">{filename}</span> is being checked locally.</p>
        <div className="mt-3 flex items-center gap-2 text-[11px] text-[hsl(var(--muted-foreground))]"><span className="size-1.5 rounded-full bg-[hsl(var(--primary))] animate-pulse-line" />No data is uploaded or sent to a server</div>
      </div>
      <div className="mt-12 max-w-3xl space-y-3">
        {steps.map((label, index) => { const isActive = index === activeStep; const isComplete = index < activeStep; return <div key={label} className={`flex items-center gap-3 rounded-xl border px-4 py-4 transition-all duration-500 ${isActive ? 'border-[hsl(var(--primary)/.5)] bg-[hsl(var(--primary)/.09)] shadow-[var(--shadow-soft)]' : 'border-[hsl(var(--border))] bg-[hsl(var(--card)/.6)]'}`} style={{ animationDelay: `${index * 140}ms` }}><div className={`grid size-5 place-items-center rounded-full text-[10px] font-bold transition-all duration-500 ${isActive ? 'bg-[hsl(var(--primary))] text-[hsl(var(--primary-foreground))] scale-110' : isComplete ? 'bg-[hsl(var(--primary)/.18)] text-[hsl(var(--primary))]' : 'bg-[hsl(var(--secondary))] text-[hsl(var(--muted-foreground))]'}`}>{isComplete ? '✓' : index + 1}</div><span className={`text-[13px] transition-colors ${isActive ? 'font-semibold text-[hsl(var(--foreground))]' : 'text-[hsl(var(--muted-foreground))]'}`}>{label}</span><div className="ml-auto h-1.5 w-24 overflow-hidden rounded-full bg-[hsl(var(--secondary))]"><div className={`h-full rounded-full transition-all duration-700 ${isActive ? 'w-full bg-[hsl(var(--primary))]' : isComplete ? 'w-full bg-[hsl(var(--primary)/.45)]' : 'w-0'}`} /></div></div>; })}
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
  return (
    <Card data-testid={testId} className={`surface-lift overflow-hidden border p-5 ${tone === 'warning' ? 'border-[hsl(var(--accent)/.45)] bg-[linear-gradient(180deg,_rgba(251,146,60,0.06),_rgba(255,255,255,0.98))]' : 'border-[hsl(var(--border))] bg-[linear-gradient(180deg,_rgba(255,255,255,0.96),_rgba(248,250,252,0.98))]'}`}>
      <div className="flex items-start justify-between gap-3">
        <p className="font-mono text-[11px] uppercase tracking-[.14em] text-[hsl(var(--muted-foreground))]">{label}</p>
        {tone === 'warning' && <Badge variant="destructive" className="text-[9px] uppercase tracking-[.08em]">Review</Badge>}
      </div>
      <p className="mt-4 break-words font-display text-[clamp(1.25rem,2vw,2.15rem)] font-bold leading-none tracking-[-.04em] text-[hsl(var(--foreground))]">{value}</p>
      <p className="mt-2 text-[12px] text-[hsl(var(--muted-foreground))]">{note}</p>
    </Card>
  );
}

function SectionHeading({ eyebrow, title, count, id }: { eyebrow: string; title: string; count?: string; id: string }) {
  return (
    <div id={id} className="mb-5 flex scroll-mt-28 items-end justify-between gap-4">
      <div>
        <p className="font-mono text-[11px] uppercase tracking-[.18em] text-[hsl(var(--primary))]">{eyebrow}</p>
        <h2 className="mt-1 font-display text-[1.7rem] font-bold tracking-[-.035em]">{title}</h2>
      </div>
      {count && <Badge variant="secondary" className="rounded-full px-3 py-1.5 font-mono text-[11px] font-medium text-[hsl(var(--foreground)/.74)]">{count}</Badge>}
    </div>
  );
}

function DataTable({ headers, rows, emptyText, testId }: { headers: string[]; rows: (string | number)[][]; emptyText: string; testId: string }) {
  return <div data-testid={testId} className="overflow-hidden rounded-xl border border-[hsl(var(--border))] bg-[hsl(var(--card))] shadow-[0_7px_20px_rgba(52,61,46,.035)]"><div className="data-table-scroll overflow-x-auto"><table className="w-full min-w-[690px] border-collapse text-left"><thead><tr className="border-b border-[hsl(var(--border))] bg-[hsl(var(--secondary)/.72)]">{headers.map((header) => <th key={header} className="whitespace-nowrap px-4 py-3.5 font-mono text-[10px] font-semibold uppercase tracking-[.12em] text-[hsl(var(--foreground)/.72)]">{header}</th>)}</tr></thead><tbody>{rows.length ? rows.map((row, rowIndex) => <tr key={`${testId}-${rowIndex}`} data-testid={`row-${testId}-${rowIndex}`} className="border-b border-[hsl(var(--border)/.7)] last:border-0 transition-colors hover:bg-[hsl(var(--primary)/.07)]">{row.map((cell, cellIndex) => <td key={`${rowIndex}-${cellIndex}`} className={`whitespace-nowrap px-4 py-3.5 text-[13px] ${cellIndex === 0 ? 'font-medium text-[hsl(var(--foreground)/.9)]' : 'text-[hsl(var(--muted-foreground))]'}`}>{cell}</td>)}</tr>) : <tr><td colSpan={headers.length} className="px-4 py-10 text-center text-[14px] text-[hsl(var(--muted-foreground))]">{emptyText}</td></tr>}</tbody></table></div></div>;
}

function AuditTable({ rows, emptyText, testId }: { rows: UserDailyRow[]; emptyText: string; testId: string }) {
  return <DataTable testId={testId} emptyText={emptyText} headers={['Date', 'Foodcourt', 'User', 'User type', 'Daily credits', 'Status']} rows={rows.map((row) => [formatDate(row.Date), row.Foodcourt, row.User || 'Unnamed user', row['User Type'] || '—', numeric(row['Daily User Credits']), 'REVIEW'])} />;
}

function SignalCharts({ result, selectedFoodcourt }: { result: AnalysisResult; selectedFoodcourt: string }) {
  const [hiddenFoodcourts, setHiddenFoodcourts] = useState<string[]>([]);
  const foodcourts = selectedFoodcourt === 'all' ? result.foodcourtMetrics.map((row) => row.Foodcourt) : [selectedFoodcourt];
  const colors = ['#0284c7', '#f97316', '#4f46e5', '#0d9488', '#d97706', '#db2777', '#475569', '#7c3aed', '#65a30d', '#0891b2', '#ea580c', '#9333ea', '#16a34a', '#e11d48'];
  const dailyGroups = new Map<string, Map<string, number>>();
  result.userDaily.forEach((row) => {
    if (selectedFoodcourt !== 'all' && row.Foodcourt !== selectedFoodcourt) return;
    const group = dailyGroups.get(row.dateKey) ?? new Map<string, number>();
    group.set(row.Foodcourt, (group.get(row.Foodcourt) ?? 0) + row['Daily User Credits']);
    dailyGroups.set(row.dateKey, group);
  });
  const dailyData = [...dailyGroups.entries()].sort(([left], [right]) => left.localeCompare(right)).map(([dateKey, values]) => {
    const date = new Date(`${dateKey}T00:00:00`);
    return { label: formatDate(date), ...Object.fromEntries(values) };
  });
  const weekdayData = result.weekdaySummary.map((row) => ({
    label: row['Week Day'].slice(0, 3),
    credits: row['Total Credits'],
  }));
  const tooltipStyle = {
    backgroundColor: 'hsl(var(--card))',
    border: '1px solid hsl(var(--border))',
    borderRadius: '10px',
    boxShadow: 'var(--shadow-soft)',
    fontSize: '11px',
  };

  return (
    <section className="mt-8 grid gap-5 xl:grid-cols-[1.4fr_.9fr]" aria-label="Credit pattern charts">
      <div className="surface-lift rounded-xl border border-[hsl(var(--border))] bg-[hsl(var(--card))] p-5 md:p-6">
        <div>
          <p className="font-mono text-[10px] uppercase tracking-[.15em] text-[hsl(var(--primary))]">Trend line</p>
          <h2 className="mt-1 font-display text-lg font-bold tracking-[-.03em]">Credit activity over time</h2>
          <p className="mt-1 text-[12px] text-[hsl(var(--muted-foreground))]">Daily credits by foodcourt.</p>
        </div>
        <div className="mt-4 h-[230px] w-full">
          {dailyData.length ? <ResponsiveContainer width="100%" height="100%"><LineChart data={dailyData} margin={{ top: 8, right: 8, left: -14, bottom: 0 }}>
            <CartesianGrid stroke="hsl(var(--border))" strokeDasharray="3 5" vertical={false} />
            <XAxis dataKey="label" axisLine={false} tickLine={false} tick={{ fill: 'hsl(var(--muted-foreground))', fontSize: 9 }} minTickGap={24} />
            <YAxis axisLine={false} tickLine={false} tick={{ fill: 'hsl(var(--muted-foreground))', fontSize: 9 }} width={42} tickFormatter={(value) => numeric(Number(value))} />
            <Tooltip contentStyle={tooltipStyle} formatter={(value, name) => [numeric(Number(value)), name]} />
            {foodcourts.filter((foodcourt) => !hiddenFoodcourts.includes(foodcourt)).map((foodcourt) => {
              const index = result.foodcourtMetrics.findIndex((row) => row.Foodcourt === foodcourt);
              return <Line key={foodcourt} type="monotone" dataKey={foodcourt} name={foodcourt} stroke={colors[(index < 0 ? 0 : index) % colors.length]} strokeWidth={2} dot={false} activeDot={{ r: 4 }} connectNulls />;
            })}
          </LineChart></ResponsiveContainer> : <div className="grid h-full place-items-center text-[12px] text-[hsl(var(--muted-foreground))]">No valid date data to chart.</div>}
        </div>
        <div className="mt-3 flex flex-wrap gap-1.5">
          {foodcourts.map((foodcourt) => {
            const index = result.foodcourtMetrics.findIndex((row) => row.Foodcourt === foodcourt);
            const hidden = hiddenFoodcourts.includes(foodcourt);
            return <button key={foodcourt} type="button" onClick={() => setHiddenFoodcourts((current) => hidden ? current.filter((item) => item !== foodcourt) : [...current, foodcourt])} aria-pressed={!hidden}
              className={`inline-flex items-center gap-1.5 rounded-full border px-2 py-1 text-[9px] transition-opacity ${hidden ? 'opacity-40' : 'border-[hsl(var(--border))] bg-[hsl(var(--background))]'}`}>
              <span className="size-1.5 rounded-full" style={{ backgroundColor: colors[(index < 0 ? 0 : index) % colors.length] }} />{foodcourt}
            </button>;
          })}
        </div>
      </div>
      <div className="surface-lift rounded-xl border border-[hsl(var(--border))] bg-[hsl(var(--card))] p-5 md:p-6">
        <div><p className="font-mono text-[10px] uppercase tracking-[.15em] text-[hsl(var(--primary))]">Pattern check</p><h2 className="mt-1 font-display text-lg font-bold tracking-[-.03em]">Weekday comparison</h2><p className="mt-1 text-[12px] text-[hsl(var(--muted-foreground))]">Where credit demand concentrates across the report.</p></div>
        <div className="mt-5 h-[230px] w-full">
          {weekdayData.length ? <ResponsiveContainer width="100%" height="100%"><BarChart data={weekdayData} margin={{ top: 8, right: 4, left: -20, bottom: 0 }}>
            <CartesianGrid stroke="hsl(var(--border))" strokeDasharray="3 5" vertical={false} />
            <XAxis dataKey="label" axisLine={false} tickLine={false} tick={{ fill: 'hsl(var(--muted-foreground))', fontSize: 10 }} dy={8} />
            <YAxis axisLine={false} tickLine={false} tick={{ fill: 'hsl(var(--muted-foreground))', fontSize: 10 }} width={34} />
            <Tooltip contentStyle={tooltipStyle} cursor={{ fill: 'hsl(var(--secondary) / .5)' }} />
            <Bar dataKey="credits" name="Credits" fill="hsl(var(--primary))" radius={[5, 5, 0, 0]} maxBarSize={28} />
          </BarChart></ResponsiveContainer> : <div className="grid h-full place-items-center text-[12px] text-[hsl(var(--muted-foreground))]">No weekday data to chart.</div>}
        </div>
      </div>
    </section>
  );
}

function CreditActivity({ result, selectedFoodcourt, onFoodcourtChange, onInspectDate }: { result: AnalysisResult; selectedFoodcourt: string; onFoodcourtChange: (foodcourt: string) => void; onInspectDate: (dateKey: string) => void }) {
  const [hiddenFoodcourts, setHiddenFoodcourts] = useState<string[]>([]);
  const [view, setView] = useState<'stacked' | 'proportional' | 'heatmap'>('stacked');
  const foodcourts = selectedFoodcourt === 'all'
    ? result.foodcourtMetrics.map((row) => row.Foodcourt)
    : [selectedFoodcourt];
  const colors = ['#0284c7', '#f97316', '#4f46e5', '#0d9488', '#d97706', '#db2777', '#475569', '#7c3aed', '#65a30d', '#0891b2', '#ea580c', '#9333ea', '#16a34a', '#e11d48'];
  const colorFor = (foodcourt: string) => {
    const index = result.foodcourtMetrics.findIndex((row) => row.Foodcourt === foodcourt);
    return colors[(index < 0 ? 0 : index) % colors.length];
  };
  const dailyGroups = new Map<string, { date: Date; totalCredits: number; orders: number; creditsByFoodcourt: Map<string, number> }>();
  result.userDaily.forEach((row) => {
    if (selectedFoodcourt !== 'all' && row.Foodcourt !== selectedFoodcourt) return;
    const group = dailyGroups.get(row.dateKey) ?? {
      date: row.Date,
      totalCredits: 0,
      orders: 0,
      creditsByFoodcourt: new Map<string, number>(),
    };
    group.totalCredits += row['Daily User Credits'];
    group.creditsByFoodcourt.set(row.Foodcourt, (group.creditsByFoodcourt.get(row.Foodcourt) ?? 0) + row['Daily User Credits']);
    dailyGroups.set(row.dateKey, group);
  });
  result.rawData.forEach((row) => {
    if (!row.Date || (selectedFoodcourt !== 'all' && row.Foodcourt !== selectedFoodcourt)) return;
    const group = dailyGroups.get(dateKeyFor(row.Date));
    if (group) group.orders += 1;
  });
  const flaggedDates = new Set([...result.over200, ...result.vendorCreditUsers]
    .filter((row) => selectedFoodcourt === 'all' || row.Foodcourt === selectedFoodcourt)
    .map((row) => row.dateKey));
  const topDays = [...dailyGroups.entries()]
    .sort(([, left], [, right]) => right.totalCredits - left.totalCredits)
    .slice(0, 7)
    .map(([dateKey, group]) => ({
      dateKey,
      date: group.date,
      totalCredits: group.totalCredits,
      orders: group.orders,
      creditsByFoodcourt: [...group.creditsByFoodcourt.entries()].filter(([name]) => !hiddenFoodcourts.includes(name)),
      isFlagged: flaggedDates.has(dateKey),
    }));
  const totalForScope = [...dailyGroups.values()].reduce((sum, row) => sum + row.totalCredits, 0);
  const topDayShare = totalForScope ? topDays.reduce((sum, row) => sum + row.totalCredits, 0) / totalForScope * 100 : 0;
  const sortedTotals = [...dailyGroups.values()].map((row) => row.totalCredits).sort((left, right) => left - right);
  const peak = Math.max(...sortedTotals, 0);
  const middle = Math.floor(sortedTotals.length / 2);
  const median = sortedTotals.length
    ? sortedTotals.length % 2 ? sortedTotals[middle] : (sortedTotals[middle - 1] + sortedTotals[middle]) / 2
    : 0;
  const averageOrders = topDays.length ? topDays.reduce((sum, row) => sum + row.orders, 0) / topDays.length : 0;
  const creditsByFoodcourt = new Map<string, number>();
  dailyGroups.forEach((group) => group.creditsByFoodcourt.forEach((credits, foodcourt) => {
    creditsByFoodcourt.set(foodcourt, (creditsByFoodcourt.get(foodcourt) ?? 0) + credits);
  }));
  const topFoodcourt = [...creditsByFoodcourt.entries()].sort(([, left], [, right]) => right - left)[0];
  const topFoodcourtShare = topFoodcourt && totalForScope ? topFoodcourt[1] / totalForScope * 100 : 0;
  const maxCreditsByFoodcourt = new Map<string, number>();
  dailyGroups.forEach((group) => group.creditsByFoodcourt.forEach((credits, foodcourt) => {
    maxCreditsByFoodcourt.set(foodcourt, Math.max(maxCreditsByFoodcourt.get(foodcourt) ?? 0, credits));
  }));

  return (
    <section className="mt-8 overflow-hidden rounded-xl border border-[hsl(var(--border))] bg-[hsl(var(--card))] shadow-[var(--shadow-soft)]" aria-label="Daily credit highlights">
      <div className="p-5 md:p-6">
        <div className="flex flex-col justify-between gap-4 lg:flex-row lg:items-start">
          <div>
            <div className="flex flex-wrap items-center gap-2">
              <Badge variant="secondary" className="text-[9px] uppercase tracking-[.1em]">Peak volume analysis</Badge>
              <Badge variant="secondary" className="text-[9px]">Reconciled daily data</Badge>
            </div>
            <h2 className="mt-2 font-display text-2xl font-bold tracking-[-.04em]">Highest-credit days</h2>
            <p className="mt-1 text-[12px] text-[hsl(var(--muted-foreground))]">
              Top 7 days by credits across {selectedFoodcourt === 'all' ? `${result.foodcourtCount} foodcourts` : selectedFoodcourt}.
            </p>
            <p className="mt-2 text-[11px]">
              <span className="font-medium text-amber-700">Amber: audit flags</span>
              <span className="mx-2 text-[hsl(var(--border))]">·</span>
              <span className="font-medium text-emerald-700">Green: no critical audit breaches</span>
            </p>
          </div>
          <div className="flex flex-wrap items-center gap-2">
            <div className="flex rounded-lg border border-[hsl(var(--border))] bg-[hsl(var(--background))] p-1" aria-label="Peak day chart display">
              {([
                ['stacked', 'Stacked'],
                ['proportional', 'Proportional'],
                ['heatmap', 'Location heatmap'],
              ] as const).map(([value, label]) => (
                <button key={value} type="button" onClick={() => setView(value)} aria-pressed={view === value}
                  className={`rounded-md px-2.5 py-1.5 text-[10px] font-medium transition-colors ${view === value ? 'bg-[hsl(var(--primary))] text-white shadow-sm' : 'text-[hsl(var(--muted-foreground))] hover:text-[hsl(var(--foreground))]'}`}>
                  {label}
                </button>
              ))}
            </div>
            <label className="sr-only" htmlFor="peak-day-foodcourt">Peak day foodcourt scope</label>
            <select id="peak-day-foodcourt" value={selectedFoodcourt} onChange={(event) => { setHiddenFoodcourts([]); onFoodcourtChange(event.target.value); }}
              className="rounded-lg border border-[hsl(var(--border))] bg-[hsl(var(--background))] px-3 py-2 text-[11px] font-medium outline-none focus:border-[hsl(var(--primary)/.6)]">
              <option value="all">All foodcourts ({result.foodcourtCount})</option>
              {result.foodcourtMetrics.map((row) => <option key={row.Foodcourt} value={row.Foodcourt}>{row.Foodcourt}</option>)}
            </select>
          </div>
        </div>

        <div className="mt-5 flex flex-wrap gap-2" aria-label="Filter chart by foodcourt">
          {foodcourts.map((foodcourt) => {
            const metric = result.foodcourtMetrics.find((row) => row.Foodcourt === foodcourt);
            const share = metric && result.totalCredits ? metric['Total Credits'] / result.totalCredits * 100 : 0;
            const hidden = hiddenFoodcourts.includes(foodcourt);
            return (
              <button key={foodcourt} type="button" onClick={() => setHiddenFoodcourts((current) => hidden ? current.filter((item) => item !== foodcourt) : [...current, foodcourt])}
                aria-pressed={!hidden} aria-label={`${hidden ? 'Show' : 'Hide'} ${foodcourt}`}
                className={`inline-flex items-center gap-1.5 rounded-full border px-2.5 py-1.5 text-[9px] font-medium transition-opacity ${hidden ? 'opacity-40' : 'border-[hsl(var(--border))] bg-[hsl(var(--background))]'}`}>
                <span className="size-2 rounded-full" style={{ backgroundColor: colorFor(foodcourt) }} />
                <span>{foodcourt}</span>
                {metric && <span className="font-mono text-[hsl(var(--muted-foreground))]">{share.toFixed(1)}%</span>}
              </button>
            );
          })}
        </div>

        <div className="mt-5 space-y-2.5">
          {topDays.length ? topDays.map((row) => {
            const positiveTotal = row.creditsByFoodcourt.reduce((sum, [, credits]) => sum + Math.max(0, credits), 0);
            const scaleTotal = view === 'proportional' ? positiveTotal : peak;
            return (
              <div key={row.dateKey} className="grid gap-2 rounded-lg border border-transparent px-2 py-2 transition-colors hover:border-[hsl(var(--border))] hover:bg-[hsl(var(--background)/.6)] sm:grid-cols-[82px_minmax(0,1fr)_112px] sm:items-center">
                <div className="flex items-center gap-2">
                  <div className="grid size-9 shrink-0 place-items-center rounded-lg bg-[hsl(var(--secondary))] text-center">
                    <span className="text-[9px] leading-3 text-[hsl(var(--muted-foreground))]">{row.date.toLocaleDateString('en-US', { weekday: 'short' }).toUpperCase()}</span>
                    <span className="font-mono text-[12px] font-bold leading-3">{String(row.date.getDate()).padStart(2, '0')}</span>
                  </div>
                  <div className="min-w-0">
                    <p className="text-[10px] font-semibold">{formatDate(row.date)}</p>
                    <p className="text-[9px] text-[hsl(var(--muted-foreground))]">{numeric(row.orders)} orders</p>
                  </div>
                </div>

                {view === 'heatmap' ? (
                  <div className="flex min-w-0 gap-1 overflow-x-auto py-1" aria-label={`Foodcourt credits for ${formatDate(row.date)}`}>
                    {foodcourts.filter((name) => !hiddenFoodcourts.includes(name)).map((name) => {
                      const credits = row.creditsByFoodcourt.find(([court]) => court === name)?.[1] ?? 0;
                      const maximum = maxCreditsByFoodcourt.get(name) ?? 0;
                      const intensity = maximum ? Math.max(0.1, credits / maximum) : 0;
                      return <div key={name} title={`${name}: ${numeric(credits)} credits`} className="grid h-8 min-w-10 flex-1 place-items-center rounded px-1 text-[8px] font-semibold"
                        style={{ backgroundColor: credits > 0 ? `color-mix(in srgb, ${colorFor(name)} ${Math.round(intensity * 75)}%, white)` : 'hsl(var(--secondary))', color: intensity > 0.55 ? 'white' : 'hsl(var(--foreground))' }}>
                        {credits > 0 ? numeric(credits) : '—'}
                      </div>;
                    })}
                  </div>
                ) : (
                  <div className={`flex h-7 min-w-0 overflow-hidden rounded-md bg-[hsl(var(--secondary)/.7)] ${row.isFlagged ? 'ring-2 ring-amber-500/70 ring-offset-1' : ''}`}
                    aria-label={`${formatDate(row.date)} foodcourt credit distribution`}>
                    {row.creditsByFoodcourt.filter(([, credits]) => credits > 0).map(([foodcourt, credits]) => {
                      const width = scaleTotal ? Math.max(credits / scaleTotal * 100, 0) : 0;
                      if (!width) return null;
                      return <div key={foodcourt} title={`${foodcourt}: ${numeric(credits)} credits`}
                        className="flex h-full min-w-0 items-center justify-center overflow-hidden px-1 text-[8px] font-semibold text-white"
                        style={{ width: `${width}%`, backgroundColor: colorFor(foodcourt) }}>
                        {width >= 12 ? numeric(credits) : ''}
                      </div>;
                    })}
                    {!row.creditsByFoodcourt.some(([, credits]) => credits > 0) && <span className="px-3 text-[9px] text-[hsl(var(--muted-foreground))]">No positive credit volume</span>}
                  </div>
                )}

                <div className="flex items-center justify-between gap-2 sm:justify-end">
                  <div className="text-right">
                    <p className="font-mono text-[11px] font-bold text-[hsl(var(--primary))]">{numeric(row.totalCredits)}</p>
                    <p className="text-[8px] text-[hsl(var(--muted-foreground))]">{peak ? `${(row.totalCredits / peak * 100).toFixed(1)}% of peak` : '—'}</p>
                  </div>
                  <button type="button" onClick={() => onInspectDate(row.dateKey)} aria-label={`Inspect audit log for ${formatDate(row.date)}`}
                    className="rounded-md px-2 py-1 text-[9px] font-semibold text-[hsl(var(--primary))] hover:bg-[hsl(var(--primary)/.08)]">Inspect</button>
                </div>
              </div>
            );
          }) : <p className="py-10 text-center text-[12px] text-[hsl(var(--muted-foreground))]">No valid date data available.</p>}
        </div>

        <div className="mt-5 grid gap-3 rounded-xl border border-[hsl(var(--primary)/.12)] bg-[hsl(var(--primary)/.04)] p-4 sm:grid-cols-[1fr_auto] sm:items-center">
          <div>
            <p className="text-[11px] font-semibold">Peak-day summary</p>
            <p className="mt-1 text-[10px] text-[hsl(var(--muted-foreground))]">
              Top {topDays.length} days make up {topDayShare.toFixed(1)}% of tracked credits
              {' · '}Median daily volume: {numeric(median)}
              {' · '}Avg. orders on peak days: {numeric(averageOrders)}
              {topFoodcourt && ` · Leading location: ${topFoodcourt[0]} (${topFoodcourtShare.toFixed(1)}%)`}
            </p>
          </div>
          <button type="button" onClick={() => onInspectDate(topDays[0]?.dateKey ?? '')} disabled={!topDays.length}
            className="inline-flex items-center justify-center gap-1.5 rounded-lg bg-[hsl(var(--primary))] px-3 py-2 text-[10px] font-semibold text-white transition-opacity hover:opacity-90 disabled:cursor-not-allowed disabled:opacity-50">
            <Search size={12} /> Deep dive audit
          </button>
        </div>
        <div className="mt-3 flex justify-between font-mono text-[8px] text-[hsl(var(--muted-foreground))]">
          <span>{view === 'proportional' ? 'Each bar normalized to 100%' : `0 ${view === 'heatmap' ? 'credits' : 'credits'}`}</span>
          <span>{view === 'proportional' ? 'Share of day total' : `Peak ${numeric(peak)} credits`}</span>
        </div>
      </div>
    </section>
  );
}

function DataQuality({ result }: { result: AnalysisResult }) {
  const blankUsers = result.rawData.filter((row) => !row.User).length;
  const zeroCreditRows = result.rawData.filter((row) => row.Credits === 0).length;
  const unassignedFoodcourts = result.rawData.filter((row) => row.Foodcourt === 'Unassigned').length;
  const checks = [
    { label: 'Rows read', value: result.sourceRows, note: 'Source rows processed', tone: 'good' },
    { label: 'Invalid dates', value: result.invalidDates, note: result.invalidDates ? 'Excluded from date metrics' : 'All dates recognized', tone: result.invalidDates ? 'warn' : 'good' },
    { label: 'Blank users', value: blankUsers, note: blankUsers ? 'Review source values' : 'Every row has a user', tone: blankUsers ? 'warn' : 'good' },
    { label: 'Zero-credit rows', value: zeroCreditRows, note: 'Included in source totals', tone: 'neutral' },
    { label: 'Unassigned location', value: unassignedFoodcourts, note: result.foodcourtColumn ? 'Rows without a foodcourt' : 'No foodcourt column found', tone: unassignedFoodcourts ? 'warn' : 'good' },
  ];

  return <section className="mt-8 rounded-xl border border-[hsl(var(--border))] bg-[hsl(var(--card))] p-5 shadow-[0_7px_20px_rgba(31,45,67,.04)] md:p-6" aria-label="Data quality summary">
    <div className="flex flex-col justify-between gap-2 sm:flex-row sm:items-start"><div><p className="font-mono text-[11px] uppercase tracking-[.15em] text-[hsl(var(--primary))]">Trust check</p><h2 className="mt-1 font-display text-[1.45rem] font-bold tracking-[-.03em]">Data quality</h2></div><span className="text-[12px] text-[hsl(var(--muted-foreground))]">Calculated locally from the uploaded file</span></div>
    <div className="mt-5 grid gap-3 sm:grid-cols-2 xl:grid-cols-5">{checks.map((check) => <div key={check.label} className="rounded-lg border border-[hsl(var(--border))] bg-[hsl(var(--background)/.58)] p-4"><div className="flex items-center justify-between gap-2"><p className="font-mono text-[10px] font-medium uppercase tracking-[.1em] text-[hsl(var(--muted-foreground))]">{check.label}</p><span className={`size-2 rounded-full ${check.tone === 'warn' ? 'bg-[hsl(var(--accent))]' : check.tone === 'good' ? 'bg-[hsl(var(--primary))]' : 'bg-[hsl(var(--muted-foreground)/.45)]'}`} /></div><p className="mt-3 font-display text-2xl font-bold">{numeric(check.value)}</p><p className="mt-1 text-[11px] leading-4 text-[hsl(var(--muted-foreground))]">{check.note}</p></div>)}</div>
  </section>;
}

function Overview({ result, onJump, onInspectDate, selectedFoodcourt, onFoodcourtChange }: { result: AnalysisResult; onJump: (target: string) => void; onInspectDate: (dateKey: string) => void; selectedFoodcourt: string; onFoodcourtChange: (foodcourt: string) => void }) {
  const exceptionCount = result.over200.length + result.vendorCreditUsers.length;
  const dailyMetrics = dailyMetricsForFoodcourt(result, selectedFoodcourt);
  const dailyMax = Math.max(...dailyMetrics.map((row) => row['Total Credits']), 0);
  const creditsByFoodcourt = new Map<string, number>();
  result.rawData.forEach((row) => creditsByFoodcourt.set(row.Foodcourt, (creditsByFoodcourt.get(row.Foodcourt) ?? 0) + row.Credits));
  const topFoodcourt = [...creditsByFoodcourt.entries()].sort(([, left], [, right]) => right - left)[0];
  const topFoodcourtShare = topFoodcourt && result.totalCredits
    ? `${(Math.round(topFoodcourt[1] / result.totalCredits * 1000) / 10).toFixed(1)}%`
    : '0%';
  const peakDay = [...result.dailyMetrics].sort((left, right) => right['Total Credits'] - left['Total Credits'])[0];
  const flaggedDayCount = new Set([...result.over200, ...result.vendorCreditUsers].map((row) => row.dateKey)).size;
  return (
    <>
      <section id="overview-section" className="scroll-mt-28 animate-rise-in">
        <div className="flex flex-col justify-between gap-5 lg:flex-row lg:items-end">
          <div><p className="font-mono text-[10px] uppercase tracking-[.2em] text-[hsl(var(--primary))]">Weekly signal</p><h1 className="mt-2 font-display text-[clamp(2rem,4vw,3.4rem)] font-bold leading-[1] tracking-[-.055em]">A clear read on<br /><span className="text-[hsl(var(--primary))]">credit behaviour.</span></h1><p className="mt-4 max-w-xl text-[14px] leading-6 text-[hsl(var(--muted-foreground))]">Cleaned from {numeric(result.sourceRows)} rows across {numeric(result.sourceColumns)} columns. Review the exceptions first, then use the workbook as your audit trail.</p></div>
          <div className="flex shrink-0 items-center gap-2 rounded-lg border border-[hsl(var(--border))] bg-[hsl(var(--card)/.7)] px-3 py-2 text-[11px] text-[hsl(var(--muted-foreground))]"><CheckCircle2 size={15} className="text-[hsl(var(--primary))]" /> Analysis complete</div>
        </div>
        {result.notices.length > 0 && <div data-testid="status-format-notice" className="mt-7 space-y-2 rounded-xl border border-[hsl(var(--accent)/.45)] bg-[hsl(var(--accent)/.1)] px-4 py-3.5 text-[12px] leading-5 text-[hsl(var(--foreground)/.8)]">{result.notices.map((notice, index) => <div key={notice} className="flex items-start gap-3"><Info size={16} className="mt-0.5 shrink-0 text-[hsl(var(--accent-foreground))]" /><span data-testid={`text-format-notice-${index}`}>{notice}</span></div>)}</div>}
        {result.invalidDates > 0 && <div data-testid="status-invalid-dates" className="mt-7 flex items-start gap-3 rounded-xl border border-[hsl(var(--accent)/.45)] bg-[hsl(var(--accent)/.1)] px-4 py-3.5 text-[12px] leading-5 text-[hsl(var(--foreground)/.8)]"><Info size={16} className="mt-0.5 shrink-0 text-[hsl(var(--accent-foreground))]" /><span><strong>{numeric(result.invalidDates)} rows</strong> have an invalid or missing Date and were excluded from date-based metrics.</span></div>}
        <div className="mt-8 grid gap-3 sm:grid-cols-2 xl:grid-cols-4">
          <KpiCard label="Total credits tracked" value={numeric(result.totalCredits)} note={`${numeric(result.sourceRows)} source rows processed`} testId="metric-total-credits" />
          <KpiCard label="Top foodcourt" value={topFoodcourt?.[0] ?? '—'} note={`${topFoodcourtShare} of credits`} testId="metric-foodcourts" />
          <KpiCard label="Peak single-day total" value={numeric(peakDay?.['Total Credits'] ?? 0)} note={peakDay ? formatDate(peakDay.Date) : 'No valid date data'} testId="metric-peak-day" />
          <KpiCard label="Audit flags" value={`${numeric(flaggedDayCount)} days`} note="Days with exceptions" tone={flaggedDayCount ? 'warning' : 'default'} testId="metric-audit-days" />
        </div>
        <div className={`mt-5 flex flex-col gap-4 rounded-xl border p-5 sm:flex-row sm:items-center sm:justify-between ${exceptionCount ? 'border-[hsl(var(--accent)/.52)] bg-[hsl(var(--accent)/.12)]' : 'border-[hsl(var(--primary)/.2)] bg-[hsl(var(--primary)/.07)]'}`} data-testid="status-audit-summary">
          <div className="flex items-start gap-3"><div className={`mt-0.5 grid size-8 shrink-0 place-items-center rounded-lg ${exceptionCount ? 'bg-[hsl(var(--accent))] text-[hsl(var(--accent-foreground))]' : 'bg-[hsl(var(--primary))] text-[hsl(var(--primary-foreground))]'}`}>{exceptionCount ? <AlertTriangle size={16} /> : <Check size={17} />}</div><div><p className="text-[13px] font-semibold">{exceptionCount ? `${numeric(exceptionCount)} audit ${exceptionCount === 1 ? 'exception needs' : 'exceptions need'} review` : 'No audit exceptions found'}</p><p className="mt-1 text-[12px] text-[hsl(var(--muted-foreground))]">{exceptionCount ? 'The queue below is sorted by date and daily credit impact.' : 'Thresholds checked: over 200 credits per day and VendorNoCredit usage.'}</p></div></div>{exceptionCount > 0 && <button type="button" onClick={() => onJump('audit-section')} data-testid="button-review-exceptions" className="flex items-center gap-2 self-start rounded-lg bg-[hsl(var(--foreground))] px-3 py-2 text-[12px] font-semibold text-[hsl(var(--card))] transition-transform hover:-translate-y-px sm:self-auto">Review queue <ArrowRight size={14} /></button>}</div>
      </section>
      <DataQuality result={result} />
      <SignalCharts result={result} selectedFoodcourt={selectedFoodcourt} />
      <CreditActivity result={result} selectedFoodcourt={selectedFoodcourt} onFoodcourtChange={onFoodcourtChange} onInspectDate={onInspectDate} />
      <section className="mt-14">
        <div className="flex flex-col justify-between gap-4 sm:flex-row sm:items-end"><SectionHeading id="daily-section" eyebrow="At a glance" title="Daily rhythm" count={`${dailyMetrics.length} active days`} /><label className="flex items-center gap-2 text-[11px] text-[hsl(var(--muted-foreground))]"><span className="font-mono text-[10px] uppercase tracking-[.12em]">Foodcourt</span><select value={selectedFoodcourt} onChange={(event) => onFoodcourtChange(event.target.value)} data-testid="select-daily-foodcourt" aria-label="Filter daily rhythm by foodcourt" className="rounded-lg border border-[hsl(var(--border))] bg-[hsl(var(--card))] px-3 py-2 text-[12px] font-medium text-[hsl(var(--foreground))] outline-none focus:border-[hsl(var(--primary)/.7)] focus:ring-2 focus:ring-[hsl(var(--primary)/.15)]"><option value="all">All foodcourts</option>{result.foodcourtMetrics.map((row) => <option key={row.Foodcourt} value={row.Foodcourt}>{row.Foodcourt}</option>)}</select></label></div>
        <div className="grid gap-5 xl:grid-cols-[1.45fr_.85fr]">
          <DataTable testId="table-daily-metrics" emptyText="No valid date data was found for this foodcourt." headers={['Date', 'Day', 'Users', 'Credit users', 'Credits', 'Avg / user', 'Orders']} rows={dailyMetrics.map((row: DailyMetric) => [formatDate(row.Date), row['Week Day'].slice(0, 3), row['Total Users'], row['Credit Users'], numeric(row['Total Credits']), numeric(row['Avg Credit/User']), row['Total Orders']])} />
          <div className="rounded-xl border border-[hsl(var(--border))] bg-[hsl(var(--card))] p-5"><div className="flex items-center justify-between"><div><p className="font-mono text-[10px] uppercase tracking-[.15em] text-[hsl(var(--muted-foreground))]">Volume map</p><p className="mt-1 text-[13px] font-semibold">Credits by day</p></div><span className="font-mono text-[10px] text-[hsl(var(--muted-foreground))]">max {numeric(dailyMax)}</span></div><div className="mt-7 space-y-4">{dailyMetrics.length ? dailyMetrics.map((row) => <div key={row.Date.toISOString()} data-testid={`bar-day-${row['Week Day']}`}><div className="mb-1.5 flex justify-between text-[11px]"><span className="text-[hsl(var(--muted-foreground))]">{row['Week Day'].slice(0, 3)} <span className="font-mono text-[10px] opacity-60">{formatDate(row.Date).slice(0, 6)}</span></span><span className="font-mono font-medium">{numeric(row['Total Credits'])}</span></div><div className="h-2 overflow-hidden rounded-full bg-[hsl(var(--secondary))]"><div className="h-full rounded-full bg-[hsl(var(--primary))] transition-all duration-700" style={{ width: `${dailyMax ? Math.max(4, (row['Total Credits'] / dailyMax) * 100) : 0}%` }} /></div></div>) : <p className="py-8 text-center text-[12px] text-[hsl(var(--muted-foreground))]">Nothing to chart yet.</p>}</div></div>
        </div>
      </section>
    </>
  );
}

function UsersSection({ result }: { result: AnalysisResult }) {
  const [query, setQuery] = useState('');
  const [page, setPage] = useState(1);
  const [sortBy, setSortBy] = useState<'total' | 'max' | 'flags' | 'days'>('total');
  const pageSize = 10;
  const normalizedQuery = query.trim().toLowerCase();
  const filteredRows = [...result.userMetrics]
    .filter((row) => `${row.User} ${row.Foodcourt} ${row['User Type']}`.toLowerCase().includes(normalizedQuery))
    .sort((left, right) => {
      if (sortBy === 'max') return right['Max Credits/Day'] - left['Max Credits/Day'];
      if (sortBy === 'flags') return (right['>200 Days'] + right['Vendor Credit Days']) - (left['>200 Days'] + left['Vendor Credit Days']);
      if (sortBy === 'days') return right['Days Used'] - left['Days Used'];
      return right['Total Credits'] - left['Total Credits'];
    });
  const pageCount = Math.max(1, Math.ceil(filteredRows.length / pageSize));
  const currentPage = Math.min(page, pageCount);
  const visibleRows = filteredRows.slice((currentPage - 1) * pageSize, currentPage * pageSize);
  const firstVisible = filteredRows.length ? (currentPage - 1) * pageSize + 1 : 0;
  const lastVisible = Math.min(currentPage * pageSize, filteredRows.length);

  const updateQuery = (value: string) => {
    setQuery(value);
    setPage(1);
  };

  return <section id="users-section" className="mt-14 scroll-mt-28">
    <div className="flex flex-col justify-between gap-4 lg:flex-row lg:items-end">
      <SectionHeading id="users-heading" eyebrow="Who is using credit" title="User metrics" count={`${result.userMetrics.length} user / court records`} />
      <div className="flex w-full flex-col gap-2 sm:flex-row lg:mb-5 lg:w-auto">
        <label className="flex items-center gap-2 rounded-lg border border-[hsl(var(--border))] bg-[hsl(var(--card))] px-3 text-[11px] text-[hsl(var(--muted-foreground))]">
          <ArrowUpDown size={14} />
          <span className="sr-only">Sort user metrics</span>
          <select value={sortBy} onChange={(event) => { setSortBy(event.target.value as typeof sortBy); setPage(1); }} aria-label="Sort user metrics" className="bg-transparent py-2.5 text-[12px] font-medium text-[hsl(var(--foreground))] outline-none"><option value="total">Highest total credits</option><option value="max">Highest daily maximum</option><option value="flags">Most audit flags</option><option value="days">Most active days</option></select>
        </label>
        <div className="relative w-full sm:w-72">
        <label htmlFor="user-metrics-search" className="sr-only">Search user metrics</label>
        <Search size={15} className="pointer-events-none absolute left-3 top-2.5 text-[hsl(var(--muted-foreground))]" />
        <input id="user-metrics-search" value={query} onChange={(event) => updateQuery(event.target.value)} data-testid="input-user-search" aria-label="Search user metrics" placeholder="Search user, foodcourt, or type" className="w-full rounded-lg border border-[hsl(var(--border))] bg-[hsl(var(--card))] py-2.5 pl-9 pr-3 text-[12px] outline-none transition-colors placeholder:text-[hsl(var(--muted-foreground)/.85)] focus:border-[hsl(var(--primary)/.75)] focus:ring-2 focus:ring-[hsl(var(--primary)/.15)]" />
        </div>
      </div>
    </div>
    <div className="mb-4 flex flex-wrap items-center justify-between gap-3 text-[11px] text-[hsl(var(--muted-foreground))]">
      <span>{filteredRows.length ? `Showing ${firstVisible}–${lastVisible} of ${filteredRows.length}` : 'No matching users'}</span>
      {query && <button type="button" onClick={() => updateQuery('')} className="font-semibold text-[hsl(var(--primary))] hover:underline">Clear search</button>}
    </div>
    <DataTable testId="table-user-metrics" emptyText={query ? 'No users match this search.' : 'No user metrics are available.'} headers={['User', 'Foodcourt', 'User type', 'Days used', 'Total credits', 'Max / day', '>200 days', 'Vendor days', 'Avg / day']} rows={visibleRows.map((row) => [row.User || 'Unnamed user', row.Foodcourt, row['User Type'] || '—', row['Days Used'], numeric(row['Total Credits']), numeric(row['Max Credits/Day']), row['>200 Days'], row['Vendor Credit Days'], numeric(row['Avg Credits/Day'])])} />
    <div className="mt-4 flex items-center justify-between gap-4">
      <span className="font-mono text-[10px] uppercase tracking-[.12em] text-[hsl(var(--muted-foreground))]">Page {currentPage} of {pageCount}</span>
      <div className="flex items-center gap-1.5">
        <button type="button" onClick={() => setPage((value) => Math.max(1, value - 1))} disabled={currentPage === 1} aria-label="Previous user metrics page" className="grid size-8 place-items-center rounded-lg border border-[hsl(var(--border))] bg-[hsl(var(--card))] text-[hsl(var(--foreground))] transition-colors hover:border-[hsl(var(--primary)/.5)] disabled:cursor-not-allowed disabled:opacity-35"><ChevronLeft size={15} /></button>
        <button type="button" onClick={() => setPage((value) => Math.min(pageCount, value + 1))} disabled={currentPage === pageCount} aria-label="Next user metrics page" className="grid size-8 place-items-center rounded-lg border border-[hsl(var(--border))] bg-[hsl(var(--card))] text-[hsl(var(--foreground))] transition-colors hover:border-[hsl(var(--primary)/.5)] disabled:cursor-not-allowed disabled:opacity-35"><ChevronRight size={15} /></button>
      </div>
    </div>
  </section>;
}

function FoodcourtSection({ result, selectedFoodcourt, onFoodcourtChange }: { result: AnalysisResult; selectedFoodcourt: string; onFoodcourtChange: (foodcourt: string) => void }) {
  const [comparisonMetric, setComparisonMetric] = useState<'credits' | 'users' | 'orders' | 'average'>('credits');
  const visibleMetrics = selectedFoodcourt === 'all'
    ? result.foodcourtMetrics
    : result.foodcourtMetrics.filter((row) => row.Foodcourt === selectedFoodcourt);
  const reviewCount = visibleMetrics.filter((row) => row['Validation Status'] === 'REVIEW').length;
  const comparison = {
    credits: { label: 'Total credits', shortLabel: 'Credits', getValue: (row: FoodcourtMetric) => row['Total Credits'] },
    users: { label: 'Active users', shortLabel: 'Users', getValue: (row: FoodcourtMetric) => row['Total Users'] },
    orders: { label: 'Total orders', shortLabel: 'Orders', getValue: (row: FoodcourtMetric) => row['Total Orders'] },
    average: { label: 'Average credits / user', shortLabel: 'Avg / user', getValue: (row: FoodcourtMetric) => row['Avg Credit/User'] },
  }[comparisonMetric];

  return (
    <section id="foodcourt-section" className="mt-14 scroll-mt-28">
      <div className="flex flex-col justify-between gap-4 sm:flex-row sm:items-end">
        <SectionHeading id="foodcourt-heading" eyebrow="Validate by location" title="Foodcourt validation" count={`${result.foodcourtCount} locations`} />
        <label className="flex items-center gap-2 text-[11px] text-[hsl(var(--muted-foreground))]">
          <span className="font-mono text-[10px] uppercase tracking-[.12em]">View</span>
          <select value={selectedFoodcourt} onChange={(event) => onFoodcourtChange(event.target.value)} data-testid="select-foodcourt" className="rounded-lg border border-[hsl(var(--border))] bg-[hsl(var(--card))] px-3 py-2 text-[12px] outline-none focus:border-[hsl(var(--primary)/.5)]">
            <option value="all">All foodcourts</option>
            {result.foodcourtMetrics.map((row) => <option key={row.Foodcourt} value={row.Foodcourt}>{row.Foodcourt}</option>)}
          </select>
        </label>
      </div>
      <p className="mb-5 max-w-2xl text-[13px] leading-6 text-[hsl(var(--muted-foreground))]">Each location is validated independently, so a Bengaluru, Hyderabad, or Chennai issue can be isolated without searching through the full user population.</p>
      <div className={`mb-4 rounded-xl border px-4 py-3 text-[12px] ${reviewCount ? 'border-[hsl(var(--accent)/.52)] bg-[hsl(var(--accent)/.12)]' : 'border-[hsl(var(--primary)/.2)] bg-[hsl(var(--primary)/.07)]'}`} data-testid="status-foodcourt-validation">
        {reviewCount ? <span><strong>{numeric(reviewCount)} {reviewCount === 1 ? 'foodcourt needs' : 'foodcourts need'} review</strong> in this view. Open the audit queue to see the exact users and dates.</span> : <span><strong>All foodcourts are clear</strong> in this view. No over-200 or VendorNoCredit usage flags were found.</span>}
      </div>
      <div className="surface-lift mb-5 rounded-xl border border-[hsl(var(--border))] bg-[hsl(var(--card))] p-5 md:p-6">
        <div className="flex items-start justify-between gap-4">
          <div><p className="font-mono text-[11px] uppercase tracking-[.15em] text-[hsl(var(--primary))]">Location signal</p><h3 className="mt-1 font-display text-[1.35rem] font-bold tracking-[-.03em]">Foodcourt comparison</h3><p className="mt-1 text-[12px] text-[hsl(var(--muted-foreground))]">Compare locations by the measure that matters for this review.</p></div>
          <div className="flex shrink-0 items-center gap-2"><label htmlFor="foodcourt-comparison-metric" className="sr-only">Compare foodcourts by</label><select id="foodcourt-comparison-metric" value={comparisonMetric} onChange={(event) => setComparisonMetric(event.target.value as typeof comparisonMetric)} data-testid="select-foodcourt-comparison" className="rounded-lg border border-[hsl(var(--border))] bg-[hsl(var(--card))] px-3 py-2 text-[12px] font-medium text-[hsl(var(--foreground))] outline-none focus:border-[hsl(var(--primary)/.7)] focus:ring-2 focus:ring-[hsl(var(--primary)/.15)]"><option value="credits">Total credits</option><option value="users">Active users</option><option value="orders">Total orders</option><option value="average">Avg credits / user</option></select><span className="hidden rounded-full bg-[hsl(var(--secondary))] px-3 py-1.5 font-mono text-[10px] text-[hsl(var(--foreground)/.7)] sm:block">{visibleMetrics.length} shown</span></div>
        </div>
        <div className="mt-5 h-[250px] w-full">
          {visibleMetrics.length ? <ResponsiveContainer width="100%" height="100%"><BarChart data={visibleMetrics.map((row) => ({ name: row.Foodcourt, value: comparison.getValue(row) }))} layout="vertical" margin={{ top: 4, right: 18, left: 8, bottom: 4 }}>
            <CartesianGrid stroke="hsl(var(--border))" strokeDasharray="3 5" horizontal={false} />
            <XAxis type="number" axisLine={false} tickLine={false} tick={{ fill: 'hsl(var(--muted-foreground))', fontSize: 10 }} />
            <YAxis type="category" dataKey="name" axisLine={false} tickLine={false} width={92} tick={{ fill: 'hsl(var(--foreground) / .72)', fontSize: 11 }} />
            <Tooltip contentStyle={{ backgroundColor: 'hsl(var(--card))', border: '1px solid hsl(var(--border))', borderRadius: '10px', boxShadow: 'var(--shadow-soft)', fontSize: '11px' }} cursor={{ fill: 'hsl(var(--secondary) / .5)' }} />
            <Bar dataKey="value" name={comparison.label} fill="hsl(var(--primary))" radius={[0, 5, 5, 0]} maxBarSize={24} />
          </BarChart></ResponsiveContainer> : <div className="grid h-full place-items-center text-[13px] text-[hsl(var(--muted-foreground))]">No foodcourt data to chart.</div>}
        </div>
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

function AuditSection({ result, focusDate, onClearFocus }: { result: AnalysisResult; focusDate: string | null; onClearFocus: () => void }) {
  const [search, setSearch] = useState('');
  const [foodcourt, setFoodcourt] = useState('all');
  const matches = (row: UserDailyRow) =>
    (!focusDate || row.dateKey === focusDate) &&
    (foodcourt === 'all' || row.Foodcourt === foodcourt) &&
    `${row.Foodcourt} ${row.User} ${row['User Type']} ${row['Week Day']}`.toLowerCase().includes(search.toLowerCase());
  const over200 = result.over200.filter(matches);
  const vendor = result.vendorCreditUsers.filter(matches);
  return <section id="audit-section" className="mt-14 scroll-mt-28"><div className="flex flex-col justify-between gap-4 sm:flex-row sm:items-end p-4"><SectionHeading id="audit-heading" eyebrow="Exceptions first" title="Audit queue" count={`${result.over200.length + result.vendorCreditUsers.length} flags`} /><div className="flex flex-col gap-2 sm:flex-row"><select value={foodcourt} onChange={(event) => setFoodcourt(event.target.value)} data-testid="select-audit-foodcourt" aria-label="Filter audit queue by foodcourt" className="rounded-lg border border-[hsl(var(--border))] bg-[hsl(var(--card))] px-3 py-2 text-[12px] outline-none focus:border-[hsl(var(--primary)/.5)]"><option value="all">All foodcourts</option>{result.foodcourtMetrics.map((row) => <option key={row.Foodcourt} value={row.Foodcourt}>{row.Foodcourt}</option>)}</select><div className="relative"><Search size={14} className="pointer-events-none absolute left-3 top-2.5 text-[hsl(var(--muted-foreground))]" /><input value={search} onChange={(event) => setSearch(event.target.value)} data-testid="input-audit-search" aria-label="Filter audit queue" placeholder="Filter users…" className="w-full rounded-lg border border-[hsl(var(--border))] bg-[hsl(var(--card))] py-2 pl-9 pr-3 text-[12px] outline-none transition-colors placeholder:text-[hsl(var(--muted-foreground)/.7)] focus:border-[hsl(var(--primary)/.5)] sm:w-48" /></div></div></div>{focusDate && <div className="mb-4 flex items-center justify-between gap-3 rounded-lg border border-[hsl(var(--primary)/.2)] bg-[hsl(var(--primary)/.06)] px-4 py-3 text-[12px]"><span>Showing audit exceptions for <strong>{formatDate(new Date(`${focusDate}T00:00:00`))}</strong>.</span><button type="button" onClick={onClearFocus} className="font-semibold text-[hsl(var(--primary))] hover:underline">Show all dates</button></div>}<div className="mb-5 grid gap-3 sm:grid-cols-2"><div className="rounded-lg border border-[hsl(var(--accent)/.5)] bg-[hsl(var(--accent)/.1)] px-4 py-3"><p className="font-mono text-[10px] uppercase tracking-[.12em] text-[hsl(var(--accent-foreground))]">Threshold review</p><p className="mt-1 text-[13px] font-semibold">{numeric(over200.length)} flagged rows</p><p className="mt-1 text-[11px] text-[hsl(var(--muted-foreground))]">Daily credits above 200.</p></div><div className="rounded-lg border border-[hsl(var(--destructive)/.35)] bg-[hsl(var(--destructive)/.07)] px-4 py-3"><p className="font-mono text-[10px] uppercase tracking-[.12em] text-[hsl(var(--destructive))]">Policy review</p><p className="mt-1 text-[13px] font-semibold">{numeric(vendor.length)} flagged rows</p><p className="mt-1 text-[11px] text-[hsl(var(--muted-foreground))]">VendorNoCredit users with usage.</p></div></div><div className="grid gap-5 xl:grid-cols-2"><div><div className="mb-3 flex items-center gap-2"><div className="size-2 rounded-full bg-[hsl(var(--accent))]" /><h3 className="text-[13px] font-semibold">Over 200 credits in a day</h3><span className="font-mono text-[10px] text-[hsl(var(--muted-foreground))]">{over200.length}</span></div><AuditTable testId="table-over-200" rows={over200} emptyText={search || foodcourt !== 'all' || focusDate ? 'No matching threshold exceptions.' : 'No users exceeded 200 credits.'} /></div><div><div className="mb-3 flex items-center gap-2"><div className="size-2 rounded-full bg-[hsl(var(--destructive))]" /><h3 className="text-[13px] font-semibold">VendorNoCredit with usage</h3><span className="font-mono text-[10px] text-[hsl(var(--muted-foreground))]">{vendor.length}</span></div><AuditTable testId="table-vendor-credit" rows={vendor} emptyText={search || foodcourt !== 'all' || focusDate ? 'No matching vendor exceptions.' : 'No VendorNoCredit users used credits.'} /></div></div></section>;
}

function WeekdaySection({ result }: { result: AnalysisResult }) {
  const weekStartFor = (date: Date) => {
    const start = new Date(date.getFullYear(), date.getMonth(), date.getDate());
    const daysFromMonday = (start.getDay() + 6) % 7;
    start.setDate(start.getDate() - daysFromMonday);
    return start;
  };
  const weekKeyFor = (date: Date) => weekStartFor(date).toISOString().slice(0, 10);
  const weeklyGroups = new Map<string, UserDailyRow[]>();
  result.userDaily.forEach((row) => {
    const key = weekKeyFor(row.Date);
    const group = weeklyGroups.get(key);
    if (group) group.push(row);
    else weeklyGroups.set(key, [row]);
  });
  const weeklySummaries = [...weeklyGroups.entries()].sort(([left], [right]) => left.localeCompare(right)).map(([weekKey, rows]) => {
    const weekStart = weekStartFor(rows[0].Date);
    const weekEnd = new Date(weekStart);
    weekEnd.setDate(weekEnd.getDate() + 6);
    const rowsByDay = WEEKDAYS.map((day) => rows.filter((row) => row['Week Day'] === day));
    return {
      weekKey,
      label: `${formatDate(weekStart)} – ${formatDate(weekEnd)}`,
      rows: rowsByDay.flatMap((dayRows, index) => {
        if (!dayRows.length) return [];
        const creditUsers = new Set(dayRows.filter((row) => row['Daily User Credits'] > 0).map((row) => row.User));
        const totalCredits = dayRows.reduce((sum, row) => sum + row['Daily User Credits'], 0);
        return [{
          'Week Day': WEEKDAYS[index],
          'Unique Users': new Set(dayRows.map((row) => row.User)).size,
          'Credit Users': creditUsers.size,
          'Total Credits': totalCredits,
          'Avg Credits/Credit User': creditUsers.size ? totalCredits / creditUsers.size : 0,
          'Users >200': new Set(dayRows.filter((row) => row['Over 200']).map((row) => row.User)).size,
          'Vendor Credit Users': new Set(dayRows.filter((row) => row['Vendor Used Credit']).map((row) => row.User)).size,
        }];
      }),
    };
  });

  return <section id="weekday-section" className="mt-14 scroll-mt-28">
    <SectionHeading id="weekday-heading" eyebrow="Pattern by weekday" title="Weekday summary" count={`${weeklySummaries.length} ${weeklySummaries.length === 1 ? 'week' : 'weeks'}`} />
    <div className="space-y-5">
      {weeklySummaries.length ? weeklySummaries.map((week) => <article key={week.weekKey} className="surface-lift rounded-xl border border-[hsl(var(--border))] bg-[hsl(var(--card))] p-5 md:p-6">
        <div className="mb-4 flex flex-col justify-between gap-2 sm:flex-row sm:items-center"><div><p className="font-mono text-[11px] uppercase tracking-[.14em] text-[hsl(var(--primary))]">Reporting week</p><h3 className="mt-1 font-display text-xl font-bold tracking-[-.03em]">{week.label}</h3></div><span className="rounded-full bg-[hsl(var(--secondary))] px-3 py-1.5 font-mono text-[10px] text-[hsl(var(--foreground)/.72)]">{week.rows.length} active days</span></div>
        <DataTable testId={`table-weekday-summary-${week.weekKey}`} emptyText="No weekday data is available for this week." headers={['Weekday', 'Unique users', 'Credit users', 'Total credits', 'Avg / credit user', 'Users >200', 'Vendor users']} rows={week.rows.map((row) => [row['Week Day'], row['Unique Users'], row['Credit Users'], numeric(row['Total Credits']), numeric(row['Avg Credits/Credit User']), row['Users >200'], row['Vendor Credit Users']])} />
      </article>) : <DataTable testId="table-weekday-summary" emptyText="No weekday data is available." headers={['Weekday', 'Unique users', 'Credit users', 'Total credits', 'Avg / credit user', 'Users >200', 'Vendor users']} rows={result.weekdaySummary.map((row: WeekdayMetric) => [row['Week Day'], row['Unique Users'], row['Credit Users'], numeric(row['Total Credits']), numeric(row['Avg Credits/Credit User']), row['Users >200'], row['Vendor Credit Users']])} />}
    </div>
  </section>;
}

function SuccessView({ result, onJump, onDownload }: { result: AnalysisResult; onJump: (target: string) => void; onDownload: () => void }) {
  const [selectedFoodcourt, setSelectedFoodcourt] = useState('all');
  const [auditFocusDate, setAuditFocusDate] = useState<string | null>(null);
  const inspectDate = (dateKey: string) => {
    setAuditFocusDate(dateKey);
    window.requestAnimationFrame(() => onJump('audit-section'));
  };
  return <>
    <nav aria-label="Analysis sections" className="sticky top-[78px] z-[9] border-b border-[hsl(var(--border))] bg-[hsl(var(--background)/.96)] backdrop-blur-xl md:hidden">
      <div className="flex gap-1 overflow-x-auto px-3 py-2">
        {navItems.map(({ key, label, icon: Icon, target }) => (
          <a
            key={key}
            href={`#${target}`}
            data-testid={`mobile-link-${key}`}
            className="flex shrink-0 items-center gap-2 rounded-lg px-3 py-2 text-[12px] font-medium text-[hsl(var(--muted-foreground))] transition-colors hover:bg-[hsl(var(--primary)/.08)] hover:text-[hsl(var(--foreground))] focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-[hsl(var(--primary))]"
          >
            <Icon size={15} />
            <span>{label}</span>
          </a>
        ))}
      </div>
    </nav>
    <main data-testid="status-success" className="mx-auto max-w-[1500px] px-5 pb-24 pt-10 md:px-12 md:pt-14"><Overview result={result} onJump={onJump} onInspectDate={inspectDate} selectedFoodcourt={selectedFoodcourt} onFoodcourtChange={setSelectedFoodcourt} /><FoodcourtSection result={result} selectedFoodcourt={selectedFoodcourt} onFoodcourtChange={setSelectedFoodcourt} /><UsersSection result={result} /><AuditSection result={result} focusDate={auditFocusDate} onClearFocus={() => setAuditFocusDate(null)} /><WeekdaySection result={result} /><div className="mt-14 flex flex-col justify-between gap-5 rounded-2xl bg-[hsl(var(--sidebar))] p-6 text-[hsl(var(--sidebar-foreground))] sm:flex-row sm:items-center md:p-8"><div><div className="flex items-center gap-2 text-[hsl(var(--accent))]"><ArrowDownToLine size={17} /><p className="font-mono text-[10px] uppercase tracking-[.18em]">Take it with you</p></div><h2 className="mt-3 font-display text-2xl font-bold tracking-[-.03em]">Your audit trail is ready.</h2><p className="mt-2 max-w-lg text-[13px] leading-5 text-[hsl(var(--sidebar-foreground)/.58)]">The workbook includes overall metrics plus separate foodcourt validation and audit sheets.</p></div><button type="button" onClick={onDownload} data-testid="button-download-footer" className="flex shrink-0 items-center justify-center gap-2 rounded-lg bg-[hsl(var(--accent))] px-4 py-3 text-[13px] font-bold text-[hsl(var(--accent-foreground))] transition-transform hover:-translate-y-px">Download workbook <ArrowDownToLine size={15} /></button></div><div id="download" className="pt-6 text-center font-mono text-[10px] uppercase tracking-[.13em] text-[hsl(var(--muted-foreground))]">Processed locally · Nothing leaves this browser</div></main>
  </>;
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
      const analysis = await analyzeInWorker(
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
      {status === 'empty' && <div onDragOver={(event) => { event.preventDefault(); setDragging(true); }} onDragLeave={() => setDragging(false)} onDrop={onDrop} className={dragging ? 'bg-[hsl(var(--primary)/.04)]' : ''}><EmptyState onPick={() => inputRef.current?.click()} dragging={dragging} /></div>}
      {status === 'parsing' && filename && <ParsingState filename={filename} />}
      {status === 'invalid' && <InvalidState message={errorMessage} missingColumns={missingColumns} onRetry={() => { reset(); setTimeout(() => inputRef.current?.click(), 0); }} />}
      {status === 'success' && result && <SuccessView result={result} onJump={jump} onDownload={download} />}
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