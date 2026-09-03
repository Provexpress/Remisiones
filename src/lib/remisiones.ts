import type ExcelJS from 'exceljs';
import type {
  AgeBreakdownItem,
  AlertLevel,
  DailyPoint,
  GroupEntry,
  ParsedWorkbook,
  Remision,
  Summary,
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

export function toIsoDate(value: unknown): string {
  const raw = unwrapCell(value as ExcelJS.CellValue);
  let date: Date | null = null;
  if (raw instanceof Date && !Number.isNaN(raw.getTime())) {
    date = new Date(Date.UTC(raw.getUTCFullYear(), raw.getUTCMonth(), raw.getUTCDate()));
  } else if (typeof raw === 'number' && raw > 20_000) {
    date = new Date(Math.round((raw - 25_569) * DAY_MS));
  } else if (typeof raw === 'string') {
    const text = raw.trim();
    const iso = text.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
    const latin = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (iso) date = new Date(Date.UTC(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3])));
    else if (latin) date = new Date(Date.UTC(Number(latin[3]), Number(latin[2]) - 1, Number(latin[1])));
    else {
      const parsed = new Date(text);
      if (!Number.isNaN(parsed.getTime())) date = parsed;
    }
  }
  return date && !Number.isNaN(date.getTime()) ? date.toISOString().slice(0, 10) : '';
}

export function diffDays(laterIso: string, earlierIso: string): number {
  if (!laterIso || !earlierIso) return 0;
  const later = Date.parse(`${laterIso}T00:00:00Z`);
  const earlier = Date.parse(`${earlierIso}T00:00:00Z`);
  return Number.isFinite(later) && Number.isFinite(earlier)
    ? Math.max(0, Math.round((later - earlier) / DAY_MS))
    : 0;
}

export function getAgeRange(age: number): string {
  if (age <= 2) return '0-2 días';
  if (age <= 7) return '3-7 días';
  if (age <= 15) return '8-15 días';
  if (age <= 30) return '16-30 días';
  if (age <= 60) return '31-60 días';
  return '>60 días';
}

export function getAlert(total: number, quantity: number, age: number): AlertLevel {
  if (total <= 1) return 'Revisar valor';
  if (quantity === 0) return 'Cantidad en cero';
  if (age > 30) return 'Vencida >30 días';
  if (age > 15) return 'Prioritaria';
  return 'Normal';
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
      const hasCutoff = normalized.some((col) => col.includes('fecha corte') || col.includes('fecha_corte'));
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

export function readHistoricalDiario(sheet: ExcelJS.Worksheet | undefined): DailyPoint[] {
  if (!sheet) return [];
  const points: DailyPoint[] = [];
  for (let r = 5; r <= sheet.rowCount; r += 1) {
    const row = sheet.getRow(r);
    const dateVal = toIsoDate(row.getCell(1).value);
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

export async function parseRemisionesWorkbook(
  buffer: ArrayBuffer,
  options?: { fallbackCutoff?: string; lastModifiedDateTime?: string },
): Promise<ParsedWorkbook> {
  const { default: ExcelJSRuntime } = await import('exceljs');
  const workbook = new ExcelJSRuntime.Workbook();
  await workbook.xlsx.load(buffer);

  // Prioritize Base-SIS sheet if present and contains data, else Base
  const sisSheet = workbook.getWorksheet('Base-SIS') || workbook.getWorksheet('base-sis');
  const baseSheet = workbook.getWorksheet('Base') || workbook.getWorksheet('base');
  const targetSheet = (sisSheet && sisSheet.rowCount > 1) ? sisSheet : baseSheet;
  if (!targetSheet) {
    throw new Error('No se encontró la hoja Base-SIS ni Base en el archivo.');
  }

  const groups = readGroups(workbook.getWorksheet('Grupos'));
  const { rowNumber: headerRow } = findHeaderRow(targetSheet);
  const columns = getColumnMap(targetSheet, headerRow);
  const column = (...aliases: string[]) => {
    for (const alias of aliases) {
      const found = columns.get(normalizeText(alias));
      if (found) return found;
    }
    return 0;
  };

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

  const indices = {
    cutoff: column('Fecha_Corte', 'Fecha Corte'),
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
  for (let rowNumber = headerRow + 1; rowNumber <= targetSheet.rowCount; rowNumber += 1) {
    const row = targetSheet.getRow(rowNumber);
    const employee = textValue(row.getCell(indices.employee).value);
    if (!employee) continue;

    const rowCutoff = indices.cutoff ? toIsoDate(row.getCell(indices.cutoff).value) : '';
    const cutoff = rowCutoff || defaultCutoffIso;
    const issuedAt = toIsoDate(row.getCell(indices.issuedAt).value);
    const total = toNumber(row.getCell(indices.total).value);
    const quantity = toNumber(row.getCell(indices.quantity).value);
    const rawAgeValue = indices.age ? unwrapCell(row.getCell(indices.age).value) : null;
    const hasSourceAge = rawAgeValue != null && rawAgeValue !== '';
    const calculatedAge = issuedAt ? diffDays(cutoff, issuedAt) : 0;
    const age = hasSourceAge ? Math.max(0, Math.round(toNumber(rawAgeValue))) : calculatedAge;
    const nit = textValue(row.getCell(indices.nit).value);
    const document = textValue(row.getCell(indices.document).value);
    const order = textValue(row.getCell(indices.order).value);
    const group = matchGroup(employee, groups);
    const id = `${cutoff}-${rowNumber}-${document || order}`;

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
      alert: getAlert(total, quantity, age),
      director: group?.director ?? 'Sin asignar',
      group: group?.group ?? null,
      matchedGroup: Boolean(group),
    });
  }

  if (!records.length) {
    throw new Error(`La hoja ${targetSheet.name} no contiene registros de remisiones válidos.`);
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
    activeSheetName: targetSheet.name,
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
      badge = 'Vencida >30d';
    } else if (name === '16-30 días') {
      tone = 'orange';
      badge = 'Por vencer';
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

  if (!historicalDiario.length) return calculatedPoints;
  const map = new Map<string, DailyPoint>();
  for (const p of historicalDiario) map.set(p.cutoff, p);
  for (const p of calculatedPoints) map.set(p.cutoff, p);
  return [...map.values()].sort((a, b) => a.cutoff.localeCompare(b.cutoff));
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

export function formatCutoff(iso: string, options: Intl.DateTimeFormatOptions = {}): string {
  if (!iso) return '—';
  try {
    return new Intl.DateTimeFormat('es-CO', {
      timeZone: 'UTC',
      day: '2-digit',
      month: 'short',
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
      month: 'short',
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
