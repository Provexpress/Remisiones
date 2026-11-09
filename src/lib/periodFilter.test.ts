import { describe, it, expect } from 'vitest';
import {
  extractPeriodStructure,
  filterCutoffsByPeriod,
  getPeriodDescription,
  getCutoffWeekInfo,
  extractWithdrawnRecordsForCutoffs,
  type PeriodFilterState,
} from './periodFilter';

describe('Period Filter Utilities (Mes & Semana)', () => {
  const sampleCutoffs = [
    '2026-09-03',
    '2026-09-04',
    '2026-09-07',
    '2026-09-25',
    '2026-10-02',
    '2026-10-15',
    '2026-10-30',
    '2026-11-03',
    '2026-11-04',
    '2026-11-05',
    '2026-11-06',
    '2026-11-09',
  ];

  it('determina correctamente la semana de lunes a domingo para una fecha', () => {
    // 2026-11-09 es lunes
    const infoMon = getCutoffWeekInfo('2026-11-09');
    expect(infoMon.startIso).toBe('2026-11-09');
    expect(infoMon.endIso).toBe('2026-11-15');
    expect(infoMon.label).toBe('Semana 09/11 al 15/11');

    // 2026-11-06 es viernes
    const infoFri = getCutoffWeekInfo('2026-11-06');
    expect(infoFri.startIso).toBe('2026-11-02');
    expect(infoFri.endIso).toBe('2026-11-08');
    expect(infoFri.label).toBe('Semana 02/11 al 08/11');
  });

  it('construye la estructura de periodos con meses y semanas ordenadas de más reciente a más antigua', () => {
    const structure = extractPeriodStructure(sampleCutoffs);

    expect(structure.months.map((m) => m.key)).toEqual(['2026-11', '2026-10', '2026-09']);
    expect(structure.months[0].label).toBe('Noviembre 2026');
    expect(structure.currentMonthKey).toBe('2026-11');
    expect(structure.latestCutoff).toBe('2026-11-09');

    // Semana más reciente: 09/11 al 15/11
    expect(structure.currentWeekKey).toBe('2026-11-09_2026-11-15');
    // Semana anterior: 02/11 al 08/11
    expect(structure.previousWeekKey).toBe('2026-11-02_2026-11-08');
  });

  it('filtra por Esta Semana de manera predeterminada', () => {
    const structure = extractPeriodStructure(sampleCutoffs);
    const filter: PeriodFilterState = {
      mode: 'current-week',
      monthKey: 'all',
      weekKey: 'all',
    };

    const result = filterCutoffsByPeriod(sampleCutoffs, filter, structure);
    expect(result).toEqual(['2026-11-09']);
  });

  it('filtra por Semana Anterior', () => {
    const structure = extractPeriodStructure(sampleCutoffs);
    const filter: PeriodFilterState = {
      mode: 'previous-week',
      monthKey: 'all',
      weekKey: 'all',
    };

    const result = filterCutoffsByPeriod(sampleCutoffs, filter, structure);
    expect(result).toEqual(['2026-11-03', '2026-11-04', '2026-11-05', '2026-11-06']);
  });

  it('filtra por Mes Actual', () => {
    const structure = extractPeriodStructure(sampleCutoffs);
    const filter: PeriodFilterState = {
      mode: 'current-month',
      monthKey: 'all',
      weekKey: 'all',
    };

    const result = filterCutoffsByPeriod(sampleCutoffs, filter, structure);
    expect(result).toEqual([
      '2026-11-03',
      '2026-11-04',
      '2026-11-05',
      '2026-11-06',
      '2026-11-09',
    ]);
  });

  it('filtra por Todo el Histórico', () => {
    const structure = extractPeriodStructure(sampleCutoffs);
    const filter: PeriodFilterState = {
      mode: 'all',
      monthKey: 'all',
      weekKey: 'all',
    };

    const result = filterCutoffsByPeriod(sampleCutoffs, filter, structure);
    expect(result).toHaveLength(sampleCutoffs.length);
  });

  it('filtra de forma personalizada por un mes específico (Octubre)', () => {
    const structure = extractPeriodStructure(sampleCutoffs);
    const filter: PeriodFilterState = {
      mode: 'custom',
      monthKey: '2026-10',
      weekKey: 'all',
    };

    const result = filterCutoffsByPeriod(sampleCutoffs, filter, structure);
    expect(result).toEqual(['2026-10-02', '2026-10-15', '2026-10-30']);
  });

  it('genera descripciones claras y amigables para el usuario', () => {
    const structure = extractPeriodStructure(sampleCutoffs);
    const desc = getPeriodDescription(
      { mode: 'current-week', monthKey: 'all', weekKey: 'all' },
      structure,
      1,
    );
    expect(desc).toContain('Esta Semana');
    expect(desc).toContain('1 corte');
  });

  it('extrae correctamente las remisiones que salieron en cortes objetivo', () => {
    const records = [
      { cutoff: '2026-11-03', stableKey: 'R1', document: '101', order: 'O1', company: 'C1', nit: 'N1', employee: 'E1', director: 'D1', group: 1, total: 1000, issuedAt: '2026-11-01' },
      { cutoff: '2026-11-03', stableKey: 'R2', document: '102', order: 'O2', company: 'C2', nit: 'N2', employee: 'E2', director: 'D1', group: 1, total: 2000, issuedAt: '2026-11-01' },
      // En 2026-11-04 salió R1 (ya no está) y se mantiene R2
      { cutoff: '2026-11-04', stableKey: 'R2', document: '102', order: 'O2', company: 'C2', nit: 'N2', employee: 'E2', director: 'D1', group: 1, total: 2000, issuedAt: '2026-11-01' },
      // En 2026-11-05 salió R2
      { cutoff: '2026-11-05', stableKey: 'R3', document: '103', order: 'O3', company: 'C3', nit: 'N3', employee: 'E3', director: 'D1', group: 1, total: 3000, issuedAt: '2026-11-04' },
    ];

    const cutoffs = ['2026-11-03', '2026-11-04', '2026-11-05'];
    const withdrawnNov4 = extractWithdrawnRecordsForCutoffs(records, cutoffs, ['2026-11-04']);
    expect(withdrawnNov4).toHaveLength(1);
    expect(withdrawnNov4[0].document).toBe('101');
    expect(withdrawnNov4[0].exitCutoff).toBe('2026-11-04');

    const withdrawnAll = extractWithdrawnRecordsForCutoffs(records, cutoffs, ['2026-11-04', '2026-11-05']);
    expect(withdrawnAll).toHaveLength(2);
    expect(withdrawnAll.map((r) => r.document)).toEqual(['102', '101']); // ordenado por total desc o fecha
  });
});
