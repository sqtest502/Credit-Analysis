import * as XLSX from 'xlsx';

export const REQUIRED_COLUMNS = ['User', 'User Type', 'Date', 'Credits'] as const;
export const WEEKDAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];

export type CleanRow = Record<string, unknown> & {
  User: string;
  'User Type': string;
  Date: Date | null;
  Credits: number;
  'Week Day': string;
};

export type UserDailyRow = {
  Date: Date;
  dateKey: string;
  'Week Day': string;
  User: string;
  'User Type': string;
  'Daily User Credits': number;
  'Over 200': boolean;
  'Vendor No Credit': boolean;
  'Vendor Used Credit': boolean;
  'Audit Status': 'OK' | 'REVIEW';
};

export type DailyMetric = {
  Date: Date;
  'Week Day': string;
  'Total Users': number;
  'Credit Users': number;
  'Total Credits': number;
  'Avg Credit/User': number;
  'Users >200': number;
  'Vendor Credit Users': number;
  'Total Orders': number;
};

export type UserMetric = {
  User: string;
  'User Type': string;
  'Days Used': number;
  'Total Credits': number;
  'Max Credits/Day': number;
  '>200 Days': number;
  'Vendor Credit Days': number;
  'Avg Credits/Day': number;
};

export type WeekdayMetric = {
  'Week Day': string;
  'Unique Users': number;
  'Credit Users': number;
  'Total Credits': number;
  'Avg Credits/Credit User': number;
  'Users >200': number;
  'Vendor Credit Users': number;
};

export type AnalysisResult = {
  sourceRows: number;
  sourceColumns: number;
  rawData: CleanRow[];
  userDaily: UserDailyRow[];
  dailyMetrics: DailyMetric[];
  userMetrics: UserMetric[];
  over200: UserDailyRow[];
  vendorCreditUsers: UserDailyRow[];
  weekdaySummary: WeekdayMetric[];
  invalidDates: number;
  totalUsers: number;
  creditUsers: number;
  totalCredits: number;
};

export class AnalysisError extends Error {
  missingColumns: string[];

  constructor(message: string, missingColumns: string[] = []) {
    super(message);
    this.name = 'AnalysisError';
    this.missingColumns = missingColumns;
  }
}

const asCleanHeader = (value: unknown) => String(value ?? '').trim().replace(/\s+/g, ' ');

const round2 = (value: number) => Math.round((value + Number.EPSILON) * 100) / 100;

const pad = (value: number) => String(value).padStart(2, '0');

const dateKeyFor = (date: Date) =>
  `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;

const fromDateParts = (year: number, month: number, day: number) => {
  const date = new Date(year, month - 1, day);
  return Number.isNaN(date.getTime()) || date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day
    ? null
    : date;
};

const cleanDate = (value: unknown): Date | null => {
  if (value instanceof Date) {
    return Number.isNaN(value.getTime()) ? null : new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }
  if (typeof value === 'number' && Number.isFinite(value)) {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) return fromDateParts(parsed.y, parsed.m, parsed.d);
    return null;
  }
  const text = String(value ?? '').trim();
  if (!text) return null;
  const explicit = text.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/);
  if (explicit) {
    const [, first, second, year] = explicit;
    const firstNumber = Number(first);
    const secondNumber = Number(second);
    // Match pandas' dayfirst=False default where possible, then support unambiguous day-first exports.
    return firstNumber > 12
      ? fromDateParts(Number(year), secondNumber, firstNumber)
      : fromDateParts(Number(year), firstNumber, secondNumber);
  }
  const parsed = new Date(text);
  if (Number.isNaN(parsed.getTime())) return null;
  return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
};

const cleanCredits = (value: unknown) => {
  const text = String(value ?? '').replace(/,/g, '').replace(/₹/g, '').trim();
  const number = Number(text);
  return Number.isFinite(number) ? number : 0;
};

const weekday = (date: Date) => date.toLocaleDateString('en-US', { weekday: 'long' });

const unique = (values: string[]) => new Set(values).size;

const toRows = (sheet: XLSX.WorkSheet) =>
  XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: null, raw: true });

export const analyzeWorkbook = async (
  input: ArrayBuffer | string,
  format: 'excel' | 'csv' = 'excel',
): Promise<AnalysisResult> => {
  const workbook = XLSX.read(input, {
    type: format === 'csv' ? 'string' : 'array',
    cellDates: true,
    raw: true,
  });
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  if (!firstSheet) throw new AnalysisError('The workbook does not contain a readable first sheet.');

  const sourceRows = toRows(firstSheet);
  const headerMatrix = XLSX.utils.sheet_to_json<unknown[]>(firstSheet, { header: 1, defval: null, raw: true });
  const rawHeaders = headerMatrix[0] ?? Object.keys(sourceRows[0] ?? {});
  const headers = rawHeaders.map(asCleanHeader);
  const missing = REQUIRED_COLUMNS.filter((column) => !headers.includes(column));
  if (missing.length) {
    throw new AnalysisError(`Required columns are missing: ${missing.join(', ')}`, missing);
  }

  const rawData: CleanRow[] = sourceRows.map((source) => {
    const row: Record<string, unknown> = {};
    Object.entries(source).forEach(([key, value]) => { row[asCleanHeader(key)] = value; });
    const date = cleanDate(row.Date);
    row.User = String(row.User ?? '').trim();
    row['User Type'] = String(row['User Type'] ?? '').trim();
    row.Date = date;
    row.Credits = cleanCredits(row.Credits);
    row['Week Day'] = date ? weekday(date) : '';
    return row as CleanRow;
  });

  const invalidDates = rawData.filter((row) => !row.Date).length;
  const validRows = rawData.filter((row): row is CleanRow & { Date: Date } => Boolean(row.Date));
  const grouped = new Map<string, UserDailyRow>();
  validRows.forEach((row) => {
    const date = row.Date;
    const dateKey = dateKeyFor(date);
    const key = `${dateKey}\u0000${row.User}\u0000${row['User Type']}`;
    const existing = grouped.get(key);
    if (existing) existing['Daily User Credits'] += row.Credits;
    else grouped.set(key, {
      Date: date,
      dateKey,
      'Week Day': row['Week Day'],
      User: row.User,
      'User Type': row['User Type'],
      'Daily User Credits': row.Credits,
      'Over 200': false,
      'Vendor No Credit': false,
      'Vendor Used Credit': false,
      'Audit Status': 'OK',
    });
  });

  const userDaily = [...grouped.values()].map((row) => {
    const over = row['Daily User Credits'] > 200;
    const vendor = row['User Type'].trim().toLowerCase() === 'vendornocredit';
    const vendorUsed = vendor && row['Daily User Credits'] > 0;
    return { ...row, 'Over 200': over, 'Vendor No Credit': vendor, 'Vendor Used Credit': vendorUsed, 'Audit Status': over || vendorUsed ? 'REVIEW' : 'OK' as 'OK' | 'REVIEW' };
  });

  const dates = [...new Set(userDaily.map((row) => row.dateKey))].sort();
  const dailyMetrics: DailyMetric[] = dates.map((dateKey) => {
    const group = userDaily.filter((row) => row.dateKey === dateKey);
    const date = group[0].Date;
    const creditUsers = unique(group.filter((row) => row['Daily User Credits'] > 0).map((row) => row.User));
    const totalCredits = group.reduce((sum, row) => sum + row['Daily User Credits'], 0);
    return {
      Date: date,
      'Week Day': group[0]['Week Day'],
      'Total Users': unique(group.map((row) => row.User)),
      'Credit Users': creditUsers,
      'Total Credits': totalCredits,
      'Avg Credit/User': creditUsers ? round2(totalCredits / creditUsers) : 0,
      'Users >200': unique(group.filter((row) => row['Over 200']).map((row) => row.User)),
      'Vendor Credit Users': unique(group.filter((row) => row['Vendor Used Credit']).map((row) => row.User)),
      'Total Orders': validRows.filter((row) => dateKeyFor(row.Date) === dateKey).length,
    };
  });

  const userKeys = [...new Set(userDaily.map((row) => `${row.User}\u0000${row['User Type']}`))];
  const userMetrics: UserMetric[] = userKeys.map((key) => {
    const [user, userType] = key.split('\u0000');
    const group = userDaily.filter((row) => row.User === user && row['User Type'] === userType);
    const total = group.reduce((sum, row) => sum + row['Daily User Credits'], 0);
    return {
      User: user,
      'User Type': userType,
      'Days Used': new Set(group.map((row) => row.dateKey)).size,
      'Total Credits': total,
      'Max Credits/Day': Math.max(...group.map((row) => row['Daily User Credits'])),
      '>200 Days': group.filter((row) => row['Over 200']).length,
      'Vendor Credit Days': group.filter((row) => row['Vendor Used Credit']).length,
      'Avg Credits/Day': round2(total / group.length),
    };
  }).sort((a, b) => b['Total Credits'] - a['Total Credits'] || a.User.localeCompare(b.User));

  const over200 = userDaily.filter((row) => row['Over 200']).sort((a, b) => a.dateKey.localeCompare(b.dateKey) || b['Daily User Credits'] - a['Daily User Credits']);
  const vendorCreditUsers = userDaily.filter((row) => row['Vendor Used Credit']).sort((a, b) => a.dateKey.localeCompare(b.dateKey) || b['Daily User Credits'] - a['Daily User Credits']);

  const weekdaySummary = WEEKDAYS.flatMap((day) => {
    const group = userDaily.filter((row) => row['Week Day'] === day);
    if (!group.length) return [];
    const creditRows = group.filter((row) => row['Daily User Credits'] > 0);
    const credits = group.reduce((sum, row) => sum + row['Daily User Credits'], 0);
    const creditUserCount = unique(creditRows.map((row) => row.User));
    return [{
      'Week Day': day,
      'Unique Users': unique(group.map((row) => row.User)),
      'Credit Users': creditUserCount,
      'Total Credits': credits,
      'Avg Credits/Credit User': creditUserCount ? round2(credits / creditUserCount) : 0,
      'Users >200': unique(group.filter((row) => row['Over 200']).map((row) => row.User)),
      'Vendor Credit Users': unique(group.filter((row) => row['Vendor Used Credit']).map((row) => row.User)),
    }];
  });

  return {
    sourceRows: rawData.length,
    sourceColumns: headers.length,
    rawData,
    userDaily,
    dailyMetrics,
    userMetrics,
    over200,
    vendorCreditUsers,
    weekdaySummary,
    invalidDates,
    totalUsers: unique(rawData.map((row) => row.User)),
    creditUsers: unique(rawData.filter((row) => row.Credits > 0).map((row) => row.User)),
    totalCredits: rawData.reduce((sum, row) => sum + row.Credits, 0),
  };
};

const appendSheet = (workbook: XLSX.WorkBook, name: string, rows: unknown[], headers?: string[]) => {
  const sheet = headers?.length
    ? XLSX.utils.json_to_sheet(rows, { header: headers })
    : XLSX.utils.json_to_sheet(rows);
  const range = XLSX.utils.decode_range(sheet['!ref'] ?? 'A1');
  for (let column = range.s.c; column <= range.e.c; column += 1) {
    let max = 10;
    for (let row = range.s.r; row <= Math.min(range.e.r, 200); row += 1) {
      const cell = sheet[XLSX.utils.encode_cell({ r: row, c: column })];
      max = Math.max(max, String(cell?.v ?? '').length + 2);
    }
    if (!sheet['!cols']) sheet['!cols'] = [];
    sheet['!cols'][column] = { wch: Math.min(max, 35) };
  }
  sheet['!freeze'] = { xSplit: 0, ySplit: 1 };
  if (sheet['!ref']) sheet['!autofilter'] = { ref: sheet['!ref'] };
  XLSX.utils.book_append_sheet(workbook, sheet, name);
};

export const buildWorkbook = (result: AnalysisResult) => {
  const workbook = XLSX.utils.book_new();
  appendSheet(workbook, 'Daily_Metrics', result.dailyMetrics);
  appendSheet(workbook, 'User_Metrics', result.userMetrics);
  appendSheet(workbook, 'Users_Over_200', result.over200);
  appendSheet(workbook, 'Vendor_Credit_Users', result.vendorCreditUsers);
  appendSheet(workbook, 'Weekday_Summary', result.weekdaySummary);
  appendSheet(workbook, 'User_Daily_Audit', result.userDaily);
  WEEKDAYS.forEach((day) => appendSheet(workbook, day, result.rawData.filter((row) => row['Week Day'] === day)));
  appendSheet(workbook, 'All_Data', result.rawData);
  return XLSX.write(workbook, { bookType: 'xlsx', type: 'array', compression: true }) as ArrayBuffer;
};

export const formatDate = (date: Date | null) =>
  date ? date.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }) : '—';

export const formatNumber = (number: number) =>
  new Intl.NumberFormat('en-US', { maximumFractionDigits: 2 }).format(number);