/**
 * Utilidades para segmentación, filtrado cronológico por Mes y Semana
 * en tableros de Trazabilidad Operativa, Desmonte y Monitor en Vivo.
 */

export type PeriodMode = 'current-week' | 'previous-week' | 'current-month' | 'all' | 'custom';

export interface PeriodFilterState {
  mode: PeriodMode;
  monthKey: string; // 'all' | 'YYYY-MM'
  weekKey: string;  // 'all' | 'YYYY-MM-DD_YYYY-MM-DD'
}

export interface MonthOption {
  key: string;       // '2026-11'
  label: string;     // 'Noviembre 2026'
  shortLabel: string;// 'Nov 2026'
  count: number;     // Cantidad de cortes en el mes
  cutoffs: string[];
}

export interface WeekOption {
  key: string;       // '2026-11-02_2026-11-08'
  label: string;     // 'Semana 02/11 al 08/11'
  monthKey: string;  // '2026-11'
  startIso: string;  // '2026-11-02' (Lunes)
  endIso: string;    // '2026-11-08' (Domingo)
  count: number;     // Cantidad de cortes en la semana
  cutoffs: string[];
}

export interface PeriodStructure {
  months: MonthOption[];
  weeks: WeekOption[];
  currentWeekKey: string;
  previousWeekKey: string;
  currentMonthKey: string;
  currentMonthLabel: string;
  currentWeekLabel: string;
  latestCutoff: string;
}

const MONTH_NAMES = [
  'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
  'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre',
];

const MONTH_NAMES_SHORT = [
  'Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun',
  'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic',
];

function pad(n: number): string {
  return String(n).padStart(2, '0');
}

export function parseIsoDate(iso: string): Date {
  const [y, m, d] = iso.split('-').map(Number);
  return new Date(y, (m || 1) - 1, d || 1);
}

export function formatIsoDate(d: Date): string {
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}

export function formatDateShort(d: Date): string {
  return `${pad(d.getDate())}/${pad(d.getMonth() + 1)}`;
}

/**
 * Obtiene el lunes de la semana a la que pertenece una fecha dada.
 */
export function getMondayOfWeek(date: Date): Date {
  const d = new Date(date);
  const day = d.getDay();
  const diff = day === 0 ? -6 : 1 - day; // Si es domingo (0), restar 6 días
  d.setDate(d.getDate() + diff);
  d.setHours(0, 0, 0, 0);
  return d;
}

/**
 * Obtiene la clave y etiqueta de semana (Lunes a Domingo) para un corte ISO.
 */
export function getCutoffWeekInfo(cutoff: string): {
  weekKey: string;
  label: string;
  startIso: string;
  endIso: string;
  monthKey: string;
} {
  const date = parseIsoDate(cutoff);
  const monday = getMondayOfWeek(date);
  const sunday = new Date(monday);
  sunday.setDate(monday.getDate() + 6);

  const startIso = formatIsoDate(monday);
  const endIso = formatIsoDate(sunday);
  const weekKey = `${startIso}_${endIso}`;
  const label = `Semana ${formatDateShort(monday)} al ${formatDateShort(sunday)}`;
  const monthKey = cutoff.slice(0, 7);

  return { weekKey, label, startIso, endIso, monthKey };
}

/**
 * Analiza la lista de cortes y construye la estructura completa de meses y semanas disponibles.
 */
export function extractPeriodStructure(cutoffs: string[]): PeriodStructure {
  const uniqueCutoffs = [...new Set(cutoffs)].filter(Boolean).sort();
  const latestCutoff = uniqueCutoffs.at(-1) || '';

  // 1. Agrupar por mes
  const monthsMap = new Map<string, string[]>();
  for (const c of uniqueCutoffs) {
    const mKey = c.slice(0, 7);
    const list = monthsMap.get(mKey) || [];
    list.push(c);
    monthsMap.set(mKey, list);
  }

  const months: MonthOption[] = [];
  for (const [key, cList] of monthsMap.entries()) {
    const [yStr, mStr] = key.split('-');
    const y = Number(yStr);
    const m = Number(mStr);
    const mIdx = m >= 1 && m <= 12 ? m - 1 : 0;
    months.push({
      key,
      label: `${MONTH_NAMES[mIdx]} ${y}`,
      shortLabel: `${MONTH_NAMES_SHORT[mIdx]} ${y}`,
      count: cList.length,
      cutoffs: cList,
    });
  }
  // Ordenar meses de más reciente a más antiguo
  months.sort((a, b) => b.key.localeCompare(a.key));

  // 2. Agrupar por semana
  const weeksMap = new Map<string, { label: string; monthKey: string; startIso: string; endIso: string; cutoffs: string[] }>();
  for (const c of uniqueCutoffs) {
    const info = getCutoffWeekInfo(c);
    const existing = weeksMap.get(info.weekKey);
    if (existing) {
      existing.cutoffs.push(c);
    } else {
      weeksMap.set(info.weekKey, {
        label: info.label,
        monthKey: info.monthKey,
        startIso: info.startIso,
        endIso: info.endIso,
        cutoffs: [c],
      });
    }
  }

  const weeks: WeekOption[] = [];
  for (const [key, wData] of weeksMap.entries()) {
    weeks.push({
      key,
      label: wData.label,
      monthKey: wData.monthKey,
      startIso: wData.startIso,
      endIso: wData.endIso,
      count: wData.cutoffs.length,
      cutoffs: wData.cutoffs,
    });
  }
  // Ordenar semanas de más reciente a más antigua por startIso
  weeks.sort((a, b) => b.startIso.localeCompare(a.startIso));

  const currentWeekKey = weeks[0]?.key || '';
  const previousWeekKey = weeks[1]?.key || '';
  const currentMonthKey = months[0]?.key || '';
  const currentMonthLabel = months[0]?.label || 'Mes Actual';
  const currentWeekLabel = weeks[0]?.label || 'Esta Semana';

  return {
    months,
    weeks,
    currentWeekKey,
    previousWeekKey,
    currentMonthKey,
    currentMonthLabel,
    currentWeekLabel,
    latestCutoff,
  };
}

/**
 * Determina el conjunto de cortes que coinciden con el estado del filtro de periodo.
 */
export function filterCutoffsByPeriod(
  cutoffs: string[],
  filter: PeriodFilterState,
  structure: PeriodStructure,
): string[] {
  const set = new Set(cutoffs);

  if (filter.mode === 'all') {
    return [...cutoffs];
  }

  if (filter.mode === 'current-week') {
    const currentWeek = structure.weeks.find((w) => w.key === structure.currentWeekKey);
    if (!currentWeek) return [...cutoffs];
    return cutoffs.filter((c) => currentWeek.cutoffs.includes(c));
  }

  if (filter.mode === 'previous-week') {
    const prevWeek = structure.weeks.find((w) => w.key === structure.previousWeekKey);
    if (!prevWeek) return [...cutoffs];
    return cutoffs.filter((c) => prevWeek.cutoffs.includes(c));
  }

  if (filter.mode === 'current-month') {
    return cutoffs.filter((c) => c.startsWith(structure.currentMonthKey));
  }

  // Modo 'custom' con dropdowns
  return cutoffs.filter((c) => {
    if (filter.monthKey && filter.monthKey !== 'all') {
      if (!c.startsWith(filter.monthKey)) return false;
    }
    if (filter.weekKey && filter.weekKey !== 'all') {
      const wInfo = getCutoffWeekInfo(c);
      if (wInfo.weekKey !== filter.weekKey) return false;
    }
    return set.has(c);
  });
}

/**
 * Devuelve un texto descriptivo del periodo actualmente seleccionado.
 */
export function getPeriodDescription(
  filter: PeriodFilterState,
  structure: PeriodStructure,
  visibleDaysCount: number,
): string {
  const countLabel = `${visibleDaysCount} ${visibleDaysCount === 1 ? 'corte' : 'cortes'}`;

  if (filter.mode === 'current-week') {
    return `Esta Semana (${structure.currentWeekLabel}) · ${countLabel}`;
  }
  if (filter.mode === 'previous-week') {
    const prev = structure.weeks.find((w) => w.key === structure.previousWeekKey);
    return `Semana Anterior (${prev?.label || 'Previa'}) · ${countLabel}`;
  }
  if (filter.mode === 'current-month') {
    return `Mes Actual (${structure.currentMonthLabel}) · ${countLabel}`;
  }
  if (filter.mode === 'all') {
    return `Histórico Completo (Desde Septiembre) · ${countLabel}`;
  }

  // Custom
  const parts: string[] = [];
  if (filter.monthKey && filter.monthKey !== 'all') {
    const m = structure.months.find((item) => item.key === filter.monthKey);
    if (m) parts.push(m.label);
  }
  if (filter.weekKey && filter.weekKey !== 'all') {
    const w = structure.weeks.find((item) => item.key === filter.weekKey);
    if (w) parts.push(w.label);
  }

  const periodText = parts.length > 0 ? parts.join(' · ') : 'Periodo Personalizado';
  return `${periodText} · ${countLabel}`;
}

export interface PeriodWithdrawnRecord {
  document: string;
  order: string;
  company: string;
  nit: string;
  employee: string;
  director: string;
  group: number | null;
  total: number;
  merchandise?: number;
  tax?: number;
  issuedAt: string;
  exitCutoff: string;
  age?: number;
  status?: string;
  alert?: string;
}

/**
 * Extrae todas las remisiones que salieron (facturadas / cobradas) en un conjunto de cortes objetivo.
 * Se compara cada corte objetivo contra su corte inmediatamente precedente.
 */
export function extractWithdrawnRecordsForCutoffs(
  allRecords: Array<{
    cutoff: string;
    stableKey: string;
    document: string;
    order: string;
    company: string;
    nit: string;
    employee: string;
    director: string;
    group: number | null;
    total: number;
    merchandise?: number;
    tax?: number;
    issuedAt: string;
    age?: number;
    status?: string;
    alert?: string;
  }>,
  allCutoffs: string[],
  targetCutoffs: string[],
): PeriodWithdrawnRecord[] {
  const targetSet = new Set(targetCutoffs);
  const sortedAllCutoffs = [...new Set(allCutoffs)].sort();

  const byCutoff = new Map<string, Map<string, typeof allRecords[0]>>();
  for (const r of allRecords) {
    let m = byCutoff.get(r.cutoff);
    if (!m) {
      m = new Map();
      byCutoff.set(r.cutoff, m);
    }
    m.set(r.stableKey, r);
  }

  const results: PeriodWithdrawnRecord[] = [];
  const seen = new Set<string>();

  for (let i = 1; i < sortedAllCutoffs.length; i++) {
    const prevCutoff = sortedAllCutoffs[i - 1];
    const currCutoff = sortedAllCutoffs[i];

    if (!targetSet.has(currCutoff)) continue;

    const prevMap = byCutoff.get(prevCutoff);
    const currMap = byCutoff.get(currCutoff);
    if (!prevMap || !currMap) continue;

    for (const [key, prevRecord] of prevMap.entries()) {
      if (!currMap.has(key)) {
        const uniqueKey = `${currCutoff}_${key}`;
        if (!seen.has(uniqueKey)) {
          seen.add(uniqueKey);
          results.push({
            document: prevRecord.document,
            order: prevRecord.order,
            company: prevRecord.company,
            nit: prevRecord.nit,
            employee: prevRecord.employee,
            director: prevRecord.director,
            group: prevRecord.group,
            total: prevRecord.total,
            merchandise: prevRecord.merchandise,
            tax: prevRecord.tax,
            issuedAt: prevRecord.issuedAt,
            exitCutoff: currCutoff,
            age: prevRecord.age,
            status: prevRecord.status,
            alert: prevRecord.alert,
          });
        }
      }
    }
  }

  return results.sort((a, b) => {
    if (a.exitCutoff !== b.exitCutoff) return b.exitCutoff.localeCompare(a.exitCutoff);
    return b.total - a.total;
  });
}

