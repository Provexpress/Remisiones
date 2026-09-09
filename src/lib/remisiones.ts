import type ExcelJS from 'exceljs';
import type {
  AgeBreakdownItem,
  AlertLevel,
  AmountStatus,
  DailyPoint,
  DaysStatus,
  GroupEntry,
  InitialCohortPoint,
  ParsedWorkbook,
  Remision,
  Summary,
  WithdrawnRemisionDetail,
} from '../types';

const DAY_MS = 86_400_000;
export const AGE_ORDER = ['0-2 días', '3-7 días', '8-15 días', '16-30 días', '31-60 días', '>60 días'];

export function normalizeText(value: unknown): string {
  return String(value ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .trim()
    .replace(/\s+/g, ' ');
}

function unwrapCell(value: ExcelJS.CellValue | undefined): unknown {
  if (value == null) return null;
  if (value instanceof Date) return value;
  if (typeof value === 'object') {
    if ('result' in value) return value.result;
    if ('text' in value) return value.text;
    if ('richText' in value) return value.richText.map((part) => part.text).join('');
  }
  return value;
}

export function toNumber(value: unknown): number {
  const raw = unwrapCell(value as ExcelJS.CellValue);
  if (typeof raw === 'number' && Number.isFinite(raw)) return raw;
  if (typeof raw !== 'string') return 0;
  const compact = raw.trim().replace(/\s/g, '');
  if (!compact) return 0;
  let normalized = compact;
  if (compact.includes(',') && compact.includes('.')) {
    normalized = compact.lastIndexOf(',') > compact.lastIndexOf('.')
      ? compact.replace(/\./g, '').replace(',', '.')
      : compact.replace(/,/g, '');
  } else if (compact.includes(',')) {
    const pieces = compact.split(',');
    normalized = pieces.length === 2 && pieces[1].length <= 2
      ? `${pieces[0].replace(/\./g, '')}.${pieces[1]}`
      : compact.replace(/,/g, '');
  }
  const parsed = Number(normalized.replace(/[^0-9.-]/g, ''));
  return Number.isFinite(parsed) ? parsed : 0;
}

export function diffDays(laterIso: string, earlierIso: string): number {
  if (!laterIso || !earlierIso) return 0;
  const later = Date.parse(`${laterIso}T00:00:00Z`);
  const earlier = Date.parse(`${earlierIso}T00:00:00Z`);
  return Number.isFinite(later) && Number.isFinite(earlier)
    ? Math.max(0, Math.round((later - earlier) / DAY_MS))
    : 0;
}

export function toIsoDate(value: unknown, referenceIso?: string, sourceDays?: number): string {
  const raw = unwrapCell(value as ExcelJS.CellValue);
  if (raw == null || raw === '') return '';

  let y = 0;
  let m = 0;
  let d = 0;

  if (raw instanceof Date && !Number.isNaN(raw.getTime())) {
    y = raw.getUTCFullYear();
    m = raw.getUTCMonth() + 1;
    d = raw.getUTCDate();
  } else if (typeof raw === 'number' && raw > 20_000 && raw < 60_000) {
    const dt = new Date(Math.round((raw - 25_569) * DAY_MS));
    if (!Number.isNaN(dt.getTime())) {
      y = dt.getUTCFullYear();
      m = dt.getUTCMonth() + 1;
      d = dt.getUTCDate();
    }
  } else if (typeof raw === 'string') {
    const text = raw.trim();
    if (!text) return '';
    const latin = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    const iso = text.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
    if (latin) {
      d = Number(latin[1]);
      m = Number(latin[2]);
      y = Number(latin[3]);
    } else if (iso) {
      y = Number(iso[1]);
      m = Number(iso[2]);
      d = Number(iso[3]);
    } else {
      const parsed = new Date(text);
      if (!Number.isNaN(parsed.getTime())) {
        y = parsed.getUTCFullYear();
        m = parsed.getUTCMonth() + 1;
        d = parsed.getUTCDate();
      }
    }
  }

  if (!y || !m || !d) return '';

  // In Colombia/Latin locales (DD/MM/YYYY), if Excel or a US-locale tool swapped day and month (when d <= 12 and m <= 12)
  if (d <= 12 && m <= 12 && d !== m) {
    const optionA = `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
    const optionB = `${y}-${String(d).padStart(2, '0')}-${String(m).padStart(2, '0')}`;

    // 1. If we have the source "Dias" (age) column, match the date whose distance to cutoff matches the age
    if (typeof sourceDays === 'number' && referenceIso) {
      const diffA = Math.abs(diffDays(referenceIso, optionA) - sourceDays);
      const diffB = Math.abs(diffDays(referenceIso, optionB) - sourceDays);
      if (diffB < diffA) return optionB;
      return optionA;
    }

    // 2. If we have a reference cutoff date (e.g. file modified date in September), choose the option closer to it
    if (referenceIso) {
      const refTime = Date.parse(`${referenceIso}T00:00:00Z`);
      const distA = Math.abs(refTime - Date.parse(`${optionA}T00:00:00Z`));
      const distB = Math.abs(refTime - Date.parse(`${optionB}T00:00:00Z`));
      if (distB < distA) return optionB;
      return optionA;
    }

    // 3. For 2026 remisiones operations, September (09) is the primary operational month
    if (y === 2026 && (m === 9 || d === 9)) {
      return m === 9 ? optionA : optionB;
    }
  }

  return `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
}

export function getAgeRange(age: number): string {
  if (age <= 2) return '0-2 días';
  if (age <= 7) return '3-7 días';
  if (age <= 15) return '8-15 días';
  if (age <= 30) return '16-30 días';
  if (age <= 60) return '31-60 días';
  return '>60 días';
}

export function getAmountStatus(total: number): AmountStatus {
  if (total >= 5_000_000) return 'Alto valor (> $5M)';
  if (total >= 1_000_000) return 'Cuantía media ($1M - $5M)';
  return 'Menor cuantía (< $1M)';
}

export function getDaysStatus(age: number): DaysStatus {
  if (age > 60) return 'Crítico (>60 días)';
  if (age > 30) return 'Crítico (31-60 días)';
  if (age > 15) return 'Gestión comercial (16-30 días)';
  return 'Al día (0-15 días)';
}

export function getAlert(total: number, age: number): AlertLevel {
  const isHighValue = total >= 5_000_000;
  if (age > 30) {
    return isHighValue ? 'Crítico · Alto valor' : 'Crítico';
  }
  if (age > 15) {
    return isHighValue ? 'Gestión comercial · Alto valor' : 'Gestión comercial';
  }
  return isHighValue ? 'Al día · Alto valor' : 'Al día';
}

function levenshtein(a: string, b: string): number {
  if (a === b) return 0;
  if (!a.length) return b.length;
  if (!b.length) return a.length;
  const previous = Array.from({ length: b.length + 1 }, (_, index) => index);
  for (let i = 1; i <= a.length; i += 1) {
    let diagonal = previous[0];
    previous[0] = i;
    for (let j = 1; j <= b.length; j += 1) {
      const upper = previous[j];
      previous[j] = Math.min(
        previous[j] + 1,
        previous[j - 1] + 1,
        diagonal + (a[i - 1] === b[j - 1] ? 0 : 1),
      );
      diagonal = upper;
    }
  }
  return previous[b.length];
}

function tokenSimilarity(a: string, b: string): number {
  if (a === b) return 1;
  const shorter = a.length <= b.length ? a : b;
  const longer = a.length > b.length ? a : b;
  if (shorter.length >= 3 && longer.startsWith(shorter)) return 0.84;
  return 1 - levenshtein(a, b) / Math.max(a.length, b.length, 1);
}

function nameScore(employee: string, member: string): number {
  const employeeTokens = normalizeText(employee).split(' ').filter((token) => token.length >= 3);
  const memberTokens = normalizeText(member).split(' ').filter((token) => token.length >= 3);
  if (!employeeTokens.length || !memberTokens.length) return 0;
  const matched = memberTokens.map((token) => Math.max(...employeeTokens.map((candidate) => tokenSimilarity(token, candidate))));
  const average = matched.reduce((sum, score) => sum + score, 0) / matched.length;
  const exactMatches = memberTokens.filter((token) => employeeTokens.includes(token)).length;
  return average + Math.min(0.12, exactMatches * 0.04);
}

export function matchGroup(employee: string, groups: GroupEntry[]): GroupEntry | null {
  const normalized = normalizeText(employee);
  if (!normalized) return null;
  const exact = groups.find((entry) => normalizeText(entry.member) === normalized);
  if (exact) return exact;
  let best: { entry: GroupEntry; score: number } | null = null;
  for (const entry of groups) {
    const score = nameScore(employee, entry.member);
    if (!best || score > best.score) best = { entry, score };
  }
  return best && best.score >= 0.76 ? best.entry : null;
}

function readGroups(sheet: ExcelJS.Worksheet | undefined): GroupEntry[] {
  if (!sheet) return [];
  const entries: GroupEntry[] = [];
  let currentGroup = 0;
  let currentDirector = '';
  sheet.eachRow((row) => {
    const text = String(unwrapCell(row.getCell(1).value) ?? '').trim();
    const groupMatch = text.match(/^Grupo\s+(\d+).*Director(?:a)?:\s*(.+)$/i);
    if (groupMatch) {
      currentGroup = Number(groupMatch[1]);
      currentDirector = groupMatch[2].trim();
      entries.push({ group: currentGroup, director: currentDirector, member: currentDirector });
      return;
    }
    if (!text || /^Ejecutivo Comercial$/i.test(text) || !currentGroup) return;
    entries.push({ group: currentGroup, director: currentDirector, member: text });
  });
  return entries;
}

function findHeaderRow(sheet: ExcelJS.Worksheet): { rowNumber: number; hasCutoff: boolean } {
  for (let rowNumber = 1; rowNumber <= Math.min(sheet.rowCount, 30); rowNumber += 1) {
    const values = sheet.getRow(rowNumber).values as ExcelJS.CellValue[];
    const normalized = values.map((value) => normalizeText(unwrapCell(value)));
    const hasEmployee = normalized.some((col) => col.includes('empleado'));
    const hasNitOrCompany = normalized.some((col) => col.includes('nit') || col.includes('empresa'));
    const hasTotal = normalized.some((col) => col.includes('total') || col.includes('mercancia'));
    if (hasEmployee && (hasNitOrCompany || hasTotal)) {
      const hasCutoff = normalized.some((col) =>
        col.includes('fecha corte') ||
        col.includes('fecha_corte') ||
        col.includes('fecha de corte') ||
        col === 'corte'
      );
      return { rowNumber, hasCutoff };
    }
  }
  throw new Error(`La hoja ${sheet.name} no contiene las columnas necesarias (Empleado, NIT, Total).`);
}

function getColumnMap(sheet: ExcelJS.Worksheet, headerRow: number): Map<string, number> {
  const map = new Map<string, number>();
  sheet.getRow(headerRow).eachCell((cell, columnNumber) => {
    map.set(normalizeText(unwrapCell(cell.value)), columnNumber);
  });
  return map;
}

function textValue(value: unknown): string {
  const raw = unwrapCell(value as ExcelJS.CellValue);
  if (raw == null) return '';
  if (typeof raw === 'number') return Number.isInteger(raw) ? String(raw) : String(raw);
  return String(raw).trim();
}

function makeStableKey(nit: string, document: string, order: string, fallback: string): string {
  const pieces = [nit, document, order].map(normalizeText).filter(Boolean);
  return pieces.length >= 2 ? pieces.join('|') : fallback;
}

export function readHistoricalDiario(sheet: ExcelJS.Worksheet | undefined, referenceIso?: string): DailyPoint[] {
  if (!sheet) return [];
  const points: DailyPoint[] = [];
  for (let r = 5; r <= sheet.rowCount; r += 1) {
    const row = sheet.getRow(r);
    const dateVal = toIsoDate(row.getCell(1).value, referenceIso);
    const pending = toNumber(row.getCell(2).value);
    const remissions = toNumber(row.getCell(3).value);
    if (!dateVal || pending <= 0) continue;
    points.push({
      cutoff: dateVal,
      pending,
      remissions,
      clients: Math.round(toNumber(row.getCell(4).value)),
      newValue: toNumber(row.getCell(5).value),
      newCount: Math.round(toNumber(row.getCell(6).value)),
      previousBalance: toNumber(row.getCell(7).value),
      withdrawn: toNumber(row.getCell(8).value),
      withdrawnCount: 0,
      netManagement: toNumber(row.getCell(9).value),
      grossReduction: toNumber(row.getCell(10).value),
      overdueValue: toNumber(row.getCell(11).value),
      overdueCount: Math.round(toNumber(row.getCell(12).value)),
    });
  }
  return points;
}

function parseSheetRecords(
  sheet: ExcelJS.Worksheet,
  groups: GroupEntry[],
  defaultCutoffIso: string,
  cutoffDateTime: string,
  cutoffTimeDisplay: string,
  fixedCutoff?: string,
): Remision[] {
  let headerInfo: { rowNumber: number; hasCutoff: boolean };
  try {
    headerInfo = findHeaderRow(sheet);
  } catch {
    return [];
  }
  const { rowNumber: headerRow } = headerInfo;
  const columns = getColumnMap(sheet, headerRow);
  const column = (...aliases: string[]) => {
    for (const alias of aliases) {
      const found = columns.get(normalizeText(alias));
      if (found) return found;
    }
    return 0;
  };

  const indices = {
    cutoff: column(
      'Fecha de corte',
      'Fecha_de_corte',
      'Fecha de Corte',
      'Fecha_de_Corte',
      'Fecha_Corte',
      'Fecha Corte',
      'Corte',
      'Fecha',
    ),
    employee: column('Empleado'),
    nit: column('NIT'),
    company: column('Empresa'),
    merchandise: column('Vr. Mercancia', 'Valor Mercancia'),
    tax: column('Vr. IVA', 'IVA'),
    total: column('Vr. Total', 'Valor Total'),
    issuedAt: column('Emision', 'Emisión'),
    age: column('Dias', 'Antiguedad_Calculada', 'Días'),
    document: column('Documento'),
    order: column('Pedido'),
    quantity: column('Cantidad'),
  };

  const records: Remision[] = [];
  for (let rowNumber = headerRow + 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const row = sheet.getRow(rowNumber);
    const employee = textValue(row.getCell(indices.employee).value);
    if (!employee) continue;

    const rawAgeValue = indices.age ? unwrapCell(row.getCell(indices.age).value) : null;
    const hasSourceAge = rawAgeValue != null && rawAgeValue !== '';
    const sourceAge = hasSourceAge ? Math.max(0, Math.round(toNumber(rawAgeValue))) : undefined;

    const rowCutoff = indices.cutoff
      ? toIsoDate(row.getCell(indices.cutoff).value, defaultCutoffIso)
      : (fixedCutoff || defaultCutoffIso);
    const cutoff = rowCutoff || fixedCutoff || defaultCutoffIso;
    const issuedAt = toIsoDate(row.getCell(indices.issuedAt).value, cutoff, sourceAge);
    const total = toNumber(row.getCell(indices.total).value);
    const quantity = toNumber(row.getCell(indices.quantity).value);
    const calculatedAge = issuedAt ? diffDays(cutoff, issuedAt) : 0;
    const age = hasSourceAge ? Math.max(0, Math.round(toNumber(rawAgeValue))) : calculatedAge;
    const nit = textValue(row.getCell(indices.nit).value);
    const document = textValue(row.getCell(indices.document).value);
    const order = textValue(row.getCell(indices.order).value);
    const group = matchGroup(employee, groups);
    const id = `${sheet.name}-${cutoff}-${rowNumber}-${document || order}`;

    records.push({
      id,
      stableKey: makeStableKey(nit, document, order, id),
      cutoff,
      cutoffTime: cutoffTimeDisplay,
      cutoffDateTime,
      employee,
      nit,
      company: textValue(row.getCell(indices.company).value),
      merchandise: toNumber(row.getCell(indices.merchandise).value),
      tax: toNumber(row.getCell(indices.tax).value),
      total,
      issuedAt,
      age,
      document,
      order,
      quantity,
      ageRange: getAgeRange(age),
      amountStatus: getAmountStatus(total),
      daysStatus: getDaysStatus(age),
      alert: getAlert(total, age),
      director: group?.director ?? 'Sin asignar',
      group: group?.group ?? null,
      matchedGroup: Boolean(group),
    });
  }

  return records;
}

export async function parseRemisionesWorkbook(
  buffer: ArrayBuffer,
  options?: { fallbackCutoff?: string; lastModifiedDateTime?: string },
): Promise<ParsedWorkbook> {
  const { default: ExcelJSRuntime } = await import('exceljs');
  const workbook = new ExcelJSRuntime.Workbook();
  await workbook.xlsx.load(buffer);

  const groups = readGroups(workbook.getWorksheet('Grupos'));

  // Determine modification timestamp from options, workbook.modified, or current date
  let fileDate: Date | null = null;
  if (options?.lastModifiedDateTime) {
    const parsed = new Date(options.lastModifiedDateTime);
    if (!Number.isNaN(parsed.getTime())) fileDate = parsed;
  }
  if (!fileDate && workbook.modified instanceof Date && !Number.isNaN(workbook.modified.getTime())) {
    fileDate = workbook.modified;
  }
  if (!fileDate && workbook.created instanceof Date && !Number.isNaN(workbook.created.getTime())) {
    fileDate = workbook.created;
  }
  if (!fileDate) {
    fileDate = new Date();
  }

  const defaultCutoffIso = fileDate.toISOString().slice(0, 10);
  const cutoffDateTime = fileDate.toISOString();
  const cutoffTimeDisplay = formatTimeOnly(fileDate);

  const sisSheet = workbook.getWorksheet('Base-SIS') || workbook.getWorksheet('base-sis');
  const baseSheet = workbook.getWorksheet('Base') || workbook.getWorksheet('base');

  let records: Remision[] = [];
  let activeSheetName = '';

  // 1. Extract from Base if present
  let baseRecords: Remision[] = [];
  if (baseSheet && baseSheet.rowCount > 1) {
    baseRecords = parseSheetRecords(baseSheet, groups, defaultCutoffIso, cutoffDateTime, cutoffTimeDisplay);
  }

  // 2. Extract from Base-SIS if present
  let sisRecords: Remision[] = [];
  if (sisSheet && sisSheet.rowCount > 1) {
    const baseHasInitial = baseRecords.some((r) => r.cutoff === '2026-09-03');
    // If Base doesn't have 2026-09-03, Base-SIS serves as the dedicated 2026-09-03 baseline
    const fixedCutoff = baseHasInitial ? undefined : '2026-09-03';
    sisRecords = parseSheetRecords(sisSheet, groups, '2026-09-03', cutoffDateTime, cutoffTimeDisplay, fixedCutoff);
  }

  // Intelligently combine both sheets or select the best one
  if (baseRecords.length > 0 && sisRecords.length > 0) {
    const baseCutoffs = new Set(baseRecords.map((r) => r.cutoff));
    const sisCutoffs = new Set(sisRecords.map((r) => r.cutoff));
    const allCutoffs = [...new Set([...baseCutoffs, ...sisCutoffs])].sort();

    // Base-SIS is the pure baseline for 2026-09-03; other daily cuts are taken from whichever sheet contains them
    const merged: Remision[] = [];
    for (const c of allCutoffs) {
      if (c === '2026-09-03' && sisCutoffs.has(c)) {
        merged.push(...sisRecords.filter((r) => r.cutoff === c));
      } else if (baseCutoffs.has(c)) {
        merged.push(...baseRecords.filter((r) => r.cutoff === c));
      } else {
        merged.push(...sisRecords.filter((r) => r.cutoff === c));
      }
    }
    records = merged;
    activeSheetName = `${baseSheet!.name} + ${sisSheet!.name}`;
  } else if (baseRecords.length > 0) {
    records = baseRecords;
    activeSheetName = baseSheet!.name;
  } else if (sisRecords.length > 0) {
    records = sisRecords;
    activeSheetName = sisSheet!.name;
  } else {
    // Fallback if neither Base nor Base-SIS contains valid records
    const fallbackSheet = workbook.getWorksheet('Remisiones') ||
      workbook.getWorksheet('remisiones') ||
      workbook.getWorksheet('Hoja1') ||
      workbook.getWorksheet('Sheet1') ||
      workbook.worksheets.find((ws) => ws.rowCount > 1 && ws.name !== 'Dashboard' && ws.name !== 'Diario' && ws.name !== 'Grupos');
    if (fallbackSheet) {
      records = parseSheetRecords(fallbackSheet, groups, defaultCutoffIso, cutoffDateTime, cutoffTimeDisplay);
      activeSheetName = fallbackSheet.name;
    }
  }

  if (!records.length) {
    throw new Error('No se encontró la hoja Base ni Base-SIS con datos en el archivo.');
  }

  const cutoffs = [...new Set(records.map((record) => record.cutoff))].sort();
  const unmatchedEmployees = [...new Set(records.filter((record) => !record.matchedGroup).map((record) => record.employee))]
    .sort((a, b) => a.localeCompare(b, 'es'));

  return {
    records,
    groups,
    sheetNames: workbook.worksheets.map((sheet) => sheet.name),
    cutoffs,
    unmatchedEmployees,
    activeSheetName,
    cutoffDateTime,
    cutoffTimeDisplay,
  };
}

export function summarize(records: Remision[]): Summary {
  const pending = records.reduce((sum, record) => sum + record.total, 0);
  const overdueRecords = records.filter((record) => record.age > 30);
  const overdueValue = overdueRecords.reduce((sum, record) => sum + record.total, 0);
  const maxAge = records.length ? Math.max(...records.map((record) => record.age)) : 0;
  return {
    pending,
    remissions: records.length,
    merchandise: records.reduce((sum, record) => sum + record.merchandise, 0),
    tax: records.reduce((sum, record) => sum + record.tax, 0),
    averageAge: records.length ? records.reduce((sum, record) => sum + record.age, 0) / records.length : 0,
    overdueValue,
    overdueCount: overdueRecords.length,
    overduePercentage: pending > 0 ? (overdueValue / pending) * 100 : 0,
    maxAge,
    zeroQuantity: records.filter((record) => record.quantity === 0).length,
    clients: new Set(records.map((record) => record.nit || normalizeText(record.company))).size,
  };
}

export function buildAgeBreakdown(records: Remision[]): AgeBreakdownItem[] {
  const totalPending = records.reduce((sum, record) => sum + record.total, 0) || 1;
  const groups = aggregateBy(records, (record) => record.ageRange);
  return AGE_ORDER.map((name) => {
    const found = groups.find((group) => group.name === name) || { value: 0, count: 0, overdue: 0 };
    const percent = (found.value / totalPending) * 100;
    let tone: 'blue' | 'orange' | 'red' = 'blue';
    let badge: string | undefined;
    if (name === '31-60 días' || name === '>60 días') {
      tone = 'red';
      badge = 'Crítico >30d';
    } else if (name === '16-30 días') {
      tone = 'orange';
      badge = 'Gestión comercial';
    } else {
      tone = 'blue';
      badge = 'Al día';
    }
    return {
      name,
      value: found.value,
      count: found.count,
      overdue: found.overdue,
      percent,
      tone,
      badge,
    };
  });
}

export function buildDailySeries(records: Remision[], historicalDiario: DailyPoint[] = []): DailyPoint[] {
  const cutoffs = [...new Set(records.map((record) => record.cutoff))].sort();
  let previous = new Map<string, Remision>();
  const calculatedPoints = cutoffs.map((cutoff, index) => {
    const currentRecords = records.filter((record) => record.cutoff === cutoff);
    const current = new Map(currentRecords.map((record) => [record.stableKey, record]));
    const newRecords = index === 0 ? currentRecords : currentRecords.filter((record) => !previous.has(record.stableKey));
    const withdrawnRecords = index === 0 ? [] : [...previous.values()].filter((record) => !current.has(record.stableKey));
    const pending = currentRecords.reduce((sum, record) => sum + record.total, 0);
    const newValue = newRecords.reduce((sum, record) => sum + record.total, 0);
    const previousBalance = [...previous.values()].reduce((sum, record) => sum + record.total, 0);
    const withdrawn = withdrawnRecords.reduce((sum, record) => sum + record.total, 0);
    const point: DailyPoint = {
      cutoff,
      pending,
      remissions: currentRecords.length,
      clients: new Set(currentRecords.map((record) => record.nit || normalizeText(record.company))).size,
      newValue,
      newCount: newRecords.length,
      previousBalance,
      withdrawn,
      withdrawnCount: withdrawnRecords.length,
      netManagement: withdrawn - newValue,
      grossReduction: previousBalance + newValue > 0 ? withdrawn / (previousBalance + newValue) : 0,
      overdueValue: currentRecords.filter((record) => record.age > 30).reduce((sum, record) => sum + record.total, 0),
      overdueCount: currentRecords.filter((record) => record.age > 30).length,
    };
    previous = current;
    return point;
  });

  const basePoints: DailyPoint[] = [];
  if (historicalDiario.length) {
    const map = new Map<string, DailyPoint>();
    for (const p of historicalDiario) map.set(p.cutoff, p);
    for (const p of calculatedPoints) map.set(p.cutoff, p);
    basePoints.push(...[...map.values()].sort((a, b) => a.cutoff.localeCompare(b.cutoff)));
  } else {
    basePoints.push(...calculatedPoints);
  }

  // Calculate day-over-day deltas for value ($ and %) and documents (# and %)
  return basePoints.map((pt, idx) => {
    if (idx === 0) {
      return {
        ...pt,
        pendingDelta: 0,
        pendingDeltaPct: 0,
        remissionsDelta: 0,
        remissionsDeltaPct: 0,
      };
    }
    const prev = basePoints[idx - 1];
    const pendingDelta = pt.pending - prev.pending;
    const pendingDeltaPct = prev.pending > 0 ? pendingDelta / prev.pending : 0;
    const remissionsDelta = pt.remissions - prev.remissions;
    const remissionsDeltaPct = prev.remissions > 0 ? remissionsDelta / prev.remissions : 0;
    return {
      ...pt,
      pendingDelta,
      pendingDeltaPct,
      remissionsDelta,
      remissionsDeltaPct,
    };
  });
}

export function buildInitialCohortSeries(
  records: Remision[],
  preferredInitialCutoff?: string,
): InitialCohortPoint[] {
  const allCutoffs = [...new Set(records.map((r) => r.cutoff))].sort();
  if (!allCutoffs.length) return [];

  const initialCutoff = preferredInitialCutoff && allCutoffs.includes(preferredInitialCutoff)
    ? preferredInitialCutoff
    : allCutoffs[0];

  const cutoffs = allCutoffs.filter((c) => c >= initialCutoff);
  const initialRecords = records.filter((r) => r.cutoff === initialCutoff);
  const initialPending = initialRecords.reduce((sum, r) => sum + r.total, 0);
  const initialCount = initialRecords.length;
  const initialKeys = new Set(initialRecords.map((r) => r.stableKey));

  const points: InitialCohortPoint[] = [];

  for (let i = 0; i < cutoffs.length; i++) {
    const cutoff = cutoffs[i];
    if (cutoff === initialCutoff) {
      points.push({
        cutoff,
        initialPending,
        initialCount,
        stillOpenPending: initialPending,
        stillOpenCount: initialCount,
        withdrawnPending: 0,
        withdrawnCount: 0,
        recoveryPct: 0,
        stillOpenPct: 1,
        dailyWithdrawnPending: 0,
        dailyWithdrawnCount: 0,
        dailyDeltaPct: 0,
      });
      continue;
    }

    const prevPoint = points[i - 1];
    const currentRecords = records.filter((r) => r.cutoff === cutoff);
    const stillOpen = currentRecords.filter((r) => initialKeys.has(r.stableKey));
    const stillOpenPending = stillOpen.reduce((sum, r) => sum + r.total, 0);
    const stillOpenCount = stillOpen.length;
    const withdrawnPending = Math.max(0, initialPending - stillOpenPending);
    const withdrawnCount = Math.max(0, initialCount - stillOpenCount);
    const recoveryPct = initialPending > 0 ? withdrawnPending / initialPending : 0;
    const stillOpenPct = initialPending > 0 ? stillOpenPending / initialPending : 0;

    const prevPending = prevPoint ? prevPoint.stillOpenPending : initialPending;
    const prevCount = prevPoint ? prevPoint.stillOpenCount : initialCount;
    const dailyWithdrawnPending = Math.max(0, prevPending - stillOpenPending);
    const dailyWithdrawnCount = Math.max(0, prevCount - stillOpenCount);
    const dailyDeltaPct = prevPending > 0 ? dailyWithdrawnPending / prevPending : 0;

    points.push({
      cutoff,
      initialPending,
      initialCount,
      stillOpenPending,
      stillOpenCount,
      withdrawnPending,
      withdrawnCount,
      recoveryPct,
      stillOpenPct,
      dailyWithdrawnPending,
      dailyWithdrawnCount,
      dailyDeltaPct,
    });
  }

  return points;
}

export function aggregateBy<T extends string>(
  records: Remision[],
  selector: (record: Remision) => T,
): Array<{ name: T; value: number; count: number; overdue: number }> {
  const map = new Map<T, { value: number; count: number; overdue: number }>();
  records.forEach((record) => {
    const key = selector(record);
    const current = map.get(key) ?? { value: 0, count: 0, overdue: 0 };
    current.value += record.total;
    current.count += 1;
    if (record.age > 30) current.overdue += record.total;
    map.set(key, current);
  });
  return [...map.entries()]
    .map(([name, values]) => ({ name, ...values }))
    .sort((a, b) => b.value - a.value);
}

export function formatCutoff(iso: string, options?: Intl.DateTimeFormatOptions): string {
  if (!iso) return '—';
  const match = iso.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (match) {
    if (!options) {
      return `${match[3]}/${match[2]}/${match[1]}`;
    }
    if (options.day && options.month && !options.year) {
      return `${match[3]}/${match[2]}`;
    }
    if (options.day && options.month && options.year) {
      return `${match[3]}/${match[2]}/${match[1]}`;
    }
  }
  try {
    return new Intl.DateTimeFormat('es-CO', {
      timeZone: 'UTC',
      day: '2-digit',
      month: '2-digit',
      year: 'numeric',
      ...options,
    }).format(new Date(`${iso}T00:00:00Z`));
  } catch {
    return iso;
  }
}

export function formatDateTime(isoOrDate: string | Date | undefined): string {
  if (!isoOrDate) return '—';
  const date = typeof isoOrDate === 'string' ? new Date(isoOrDate) : isoOrDate;
  if (Number.isNaN(date.getTime())) return '—';
  try {
    return new Intl.DateTimeFormat('es-CO', {
      timeZone: 'America/Bogota',
      day: '2-digit',
      month: '2-digit',
      year: 'numeric',
      hour: 'numeric',
      minute: '2-digit',
      hour12: true,
    }).format(date);
  } catch {
    return date.toLocaleString('es-CO');
  }
}


export function formatTimeOnly(isoOrDate: string | Date | undefined): string {
  if (!isoOrDate) return '—';
  const date = typeof isoOrDate === 'string' ? new Date(isoOrDate) : isoOrDate;
  if (Number.isNaN(date.getTime())) return '—';
  try {
    return new Intl.DateTimeFormat('es-CO', {
      timeZone: 'America/Bogota',
      hour: 'numeric',
      minute: '2-digit',
      hour12: true,
    }).format(date);
  } catch {
    return date.toLocaleTimeString('es-CO', { hour: '2-digit', minute: '2-digit' });
  }
}

export function getInitialCohortWithdrawnRecords(
  records: Remision[],
  preferredInitialCutoff?: string,
): WithdrawnRemisionDetail[] {
  const allCutoffs = [...new Set(records.map((r) => r.cutoff))].sort();
  if (!allCutoffs.length) return [];

  const initialCutoff = preferredInitialCutoff && allCutoffs.includes(preferredInitialCutoff)
    ? preferredInitialCutoff
    : allCutoffs[0];

  const initialRecords = records.filter((r) => r.cutoff === initialCutoff);
  if (!initialRecords.length) return [];

  const subsequentCutoffs = allCutoffs.filter((c) => c > initialCutoff);
  if (!subsequentCutoffs.length) return [];

  const keysByCutoff = new Map<string, Set<string>>();
  for (const c of subsequentCutoffs) {
    keysByCutoff.set(
      c,
      new Set(records.filter((r) => r.cutoff === c).map((r) => r.stableKey)),
    );
  }

  const withdrawnList: WithdrawnRemisionDetail[] = [];

  for (const item of initialRecords) {
    let exitCutoff: string | null = null;
    for (const c of subsequentCutoffs) {
      const activeKeys = keysByCutoff.get(c);
      if (activeKeys && !activeKeys.has(item.stableKey)) {
        exitCutoff = c;
        break;
      }
    }

    if (exitCutoff) {
      const daysToClose = diffDays(exitCutoff, item.issuedAt);
      const daysInDesmonte = diffDays(exitCutoff, initialCutoff);

      withdrawnList.push({
        document: item.document,
        order: item.order,
        employee: item.employee,
        director: item.director,
        group: item.group,
        nit: item.nit,
        company: item.company,
        merchandise: item.merchandise,
        tax: item.tax,
        total: item.total,
        issuedAt: item.issuedAt,
        initialCutoff,
        exitCutoff,
        daysToClose,
        daysInDesmonte,
        initialAge: item.age,
        amountStatus: item.amountStatus,
        daysStatus: item.daysStatus,
        alert: item.alert,
      });
    }
  }

  return withdrawnList.sort((a, b) => {
    if (a.exitCutoff !== b.exitCutoff) return a.exitCutoff.localeCompare(b.exitCutoff);
    return b.total - a.total;
  });
}

export async function exportWithdrawnRemisionesToExcel(
  items: WithdrawnRemisionDetail[],
  filename?: string,
): Promise<ArrayBuffer> {
  const { default: ExcelJSRuntime } = await import('exceljs');
  const workbook = new ExcelJSRuntime.Workbook();
  workbook.creator = 'Provexpress - Control de Remisiones';
  workbook.created = new Date();

  const worksheet = workbook.addWorksheet('Remisiones Salidas');

  worksheet.columns = [
    { header: 'No. Remisión', key: 'document', width: 16 },
    { header: 'No. Pedido', key: 'order', width: 14 },
    { header: 'Comercial', key: 'employee', width: 28 },
    { header: 'Director Comercial', key: 'director', width: 26 },
    { header: 'Grupo', key: 'group', width: 12 },
    { header: 'NIT Cliente', key: 'nit', width: 16 },
    { header: 'Empresa / Cliente', key: 'company', width: 34 },
    { header: 'Vr. Mercancía', key: 'merchandise', width: 18 },
    { header: 'Vr. IVA', key: 'tax', width: 16 },
    { header: 'Vr. Total', key: 'total', width: 18 },
    { header: 'Fecha Emisión (Ingreso)', key: 'issuedAt', width: 22 },
    { header: 'Fecha Base (Entrega)', key: 'initialCutoff', width: 20 },
    { header: 'Fecha de Salida / Cierre', key: 'exitCutoff', width: 22 },
    { header: 'Días hasta Cierre (desde Emisión)', key: 'daysToClose', width: 26 },
    { header: 'Días en Desmonte (desde Base)', key: 'daysInDesmonte', width: 24 },
    { header: 'Antigüedad Inicial (Días)', key: 'initialAge', width: 22 },
    { header: 'Categoría Antigüedad', key: 'daysStatus', width: 24 },
    { header: 'Categoría Monto', key: 'amountStatus', width: 24 },
    { header: 'Nivel de Alerta', key: 'alert', width: 24 },
  ];

  const headerRow = worksheet.getRow(1);
  headerRow.height = 28;
  headerRow.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
  headerRow.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF1E293B' },
  };
  headerRow.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };

  let sumMerchandise = 0;
  let sumTax = 0;
  let sumTotal = 0;
  let sumDaysToClose = 0;
  let sumDaysInDesmonte = 0;

  for (const item of items) {
    sumMerchandise += item.merchandise || 0;
    sumTax += item.tax || 0;
    sumTotal += item.total || 0;
    sumDaysToClose += item.daysToClose || 0;
    sumDaysInDesmonte += item.daysInDesmonte || 0;

    const row = worksheet.addRow({
      document: item.document,
      order: item.order,
      employee: item.employee,
      director: item.director,
      group: item.group ? `Grupo ${item.group}` : 'Sin asignar',
      nit: item.nit,
      company: item.company,
      merchandise: item.merchandise,
      tax: item.tax,
      total: item.total,
      issuedAt: item.issuedAt,
      initialCutoff: item.initialCutoff,
      exitCutoff: item.exitCutoff,
      daysToClose: item.daysToClose,
      daysInDesmonte: item.daysInDesmonte,
      initialAge: item.initialAge,
      daysStatus: item.daysStatus,
      amountStatus: item.amountStatus,
      alert: item.alert,
    });

    row.height = 20;
    row.font = { name: 'Segoe UI', size: 9.5 };

    row.getCell('merchandise').numFmt = '"$"#,##0;[Red]-"$"#,##0;"$0"';
    row.getCell('tax').numFmt = '"$"#,##0;[Red]-"$"#,##0;"$0"';
    row.getCell('total').numFmt = '"$"#,##0;[Red]-"$"#,##0;"$0"';
    row.getCell('total').font = { name: 'Segoe UI', size: 9.5, bold: true };

    row.getCell('daysToClose').numFmt = '#,##0';
    row.getCell('daysToClose').font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FF15803D' } };
    row.getCell('daysInDesmonte').numFmt = '#,##0';
    row.getCell('initialAge').numFmt = '#,##0';

    row.getCell('document').alignment = { horizontal: 'center' };
    row.getCell('order').alignment = { horizontal: 'center' };
    row.getCell('group').alignment = { horizontal: 'center' };
    row.getCell('issuedAt').alignment = { horizontal: 'center' };
    row.getCell('initialCutoff').alignment = { horizontal: 'center' };
    row.getCell('exitCutoff').alignment = { horizontal: 'center' };
    row.getCell('daysToClose').alignment = { horizontal: 'center' };
    row.getCell('daysInDesmonte').alignment = { horizontal: 'center' };
    row.getCell('initialAge').alignment = { horizontal: 'center' };
  }

  // Summary row at the bottom
  if (items.length > 0) {
    const avgDaysToClose = Math.round(sumDaysToClose / items.length);
    const avgDaysInDesmonte = Math.round(sumDaysInDesmonte / items.length);

    const totalRow = worksheet.addRow({
      document: 'TOTALES',
      order: `${items.length} rem.`,
      employee: '',
      director: '',
      group: '',
      nit: '',
      company: 'Consolidado de Remisiones Salidas',
      merchandise: sumMerchandise,
      tax: sumTax,
      total: sumTotal,
      issuedAt: '',
      initialCutoff: '',
      exitCutoff: '',
      daysToClose: avgDaysToClose,
      daysInDesmonte: avgDaysInDesmonte,
      initialAge: '',
      daysStatus: '',
      amountStatus: '',
      alert: '',
    });

    totalRow.height = 24;
    totalRow.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FF0F172A' } };
    totalRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFF1F5F9' },
    };

    totalRow.getCell('merchandise').numFmt = '"$"#,##0;[Red]-"$"#,##0;"$0"';
    totalRow.getCell('tax').numFmt = '"$"#,##0;[Red]-"$"#,##0;"$0"';
    totalRow.getCell('total').numFmt = '"$"#,##0;[Red]-"$"#,##0;"$0"';
    totalRow.getCell('daysToClose').numFmt = '#,##0';
    totalRow.getCell('daysInDesmonte').numFmt = '#,##0';

    totalRow.getCell('document').alignment = { horizontal: 'center' };
    totalRow.getCell('order').alignment = { horizontal: 'center' };
    totalRow.getCell('daysToClose').alignment = { horizontal: 'center' };
    totalRow.getCell('daysInDesmonte').alignment = { horizontal: 'center' };
  }

  worksheet.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: 1, column: 19 },
  };

  const buffer = await workbook.xlsx.writeBuffer();
  if (typeof document !== 'undefined' && typeof window !== 'undefined') {
    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename || `remisiones-salidas-evolucion-${new Date().toISOString().slice(0, 10)}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  return buffer as ArrayBuffer;
}

