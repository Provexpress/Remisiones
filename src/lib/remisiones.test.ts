import { readFile } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import {
  buildDailySeries,
  getAgeRange,
  matchGroup,
  normalizeText,
  parseRemisionesWorkbook,
  summarize,
  toNumber,
} from './remisiones';
import type { GroupEntry } from '../types';

describe('normalización de datos', () => {
  it('normaliza acentos y separadores', () => {
    expect(normalizeText('  María ÁNGELICA—Caballero ')).toBe('maria angelica caballero');
  });

  it('convierte números con coma decimal', () => {
    expect(toNumber('979184,36')).toBeCloseTo(979184.36);
    expect(toNumber('1.234.567,89')).toBeCloseTo(1234567.89);
  });

  it('clasifica rangos de antigüedad', () => {
    expect(getAgeRange(2)).toBe('0-2 días');
    expect(getAgeRange(31)).toBe('31-60 días');
    expect(getAgeRange(92)).toBe('>60 días');
  });

  it('cruza nombres abreviados con su director', () => {
    const groups: GroupEntry[] = [
      { group: 2, director: 'Angélica Caballero', member: 'Dayana Chala' },
      { group: 3, director: 'Óscar Beltrán', member: 'Tatiana Parra' },
    ];
    expect(matchGroup('Dayana Marcela Chala', groups)?.director).toBe('Angélica Caballero');
    expect(matchGroup('Angie Tatiana Parra', groups)?.director).toBe('Óscar Beltrán');
  });
});

describe('parser del libro', () => {
  it('procesa una Base y cruza el grupo sin depender del archivo corporativo', async () => {
    const workbook = new ExcelJS.Workbook();
    const base = workbook.addWorksheet('Base');
    base.addRow(['Base histórica']);
    base.addRow([]);
    base.addRow(['Fecha_Corte', 'Empleado', 'NIT', 'Empresa', 'Vr. Mercancia', 'Vr. IVA', 'Vr. Total', 'Emision', 'Dias', 'Documento', 'Pedido', 'Cantidad']);
    base.addRow([new Date('2026-09-03T00:00:00Z'), 'Dayana Marcela Chala', '9001', 'Cliente ejemplo', 1000, 190, 1190, new Date('2026-08-01T00:00:00Z'), 33, 'R1', 'P1', 2]);
    const groups = workbook.addWorksheet('Grupos');
    groups.addRow(['Grupo 2 — Directora: Angélica Caballero']);
    groups.addRow(['Ejecutivo Comercial']);
    groups.addRow(['Dayana Chala']);
    const output = await workbook.xlsx.writeBuffer();
    const parsed = await parseRemisionesWorkbook(output as ArrayBuffer);
    expect(parsed.records).toHaveLength(1);
    expect(parsed.records[0].director).toBe('Angélica Caballero');
    expect(parsed.records[0].age).toBe(33);
    expect(parsed.records[0].total).toBe(1190);
  });
});

const corporateWorkbookPath = fileURLToPath(new URL('../../Remisiones.xlsx', import.meta.url));
const describeCorporateWorkbook = existsSync(corporateWorkbookPath) ? describe : describe.skip;

describeCorporateWorkbook('libro real de remisiones', () => {
  it('lee Base y reconcilia el dashboard del corte', async () => {
    const file = await readFile(corporateWorkbookPath);
    const arrayBuffer = file.buffer.slice(file.byteOffset, file.byteOffset + file.byteLength) as ArrayBuffer;
    const parsed = await parseRemisionesWorkbook(arrayBuffer);
    const summary = summarize(parsed.records);
    const daily = buildDailySeries(parsed.records);

    expect(parsed.cutoffs).toEqual(['2026-09-03']);
    expect(parsed.records).toHaveLength(674);
    expect(summary.pending).toBeCloseTo(3_599_705_468.61, 1);
    expect(summary.merchandise).toBeCloseTo(3_047_575_498.03, 1);
    expect(summary.tax).toBeCloseTo(552_129_970.58, 1);
    expect(summary.overdueCount).toBe(105);
    expect(daily).toHaveLength(1);
    expect(parsed.records.find((record) => record.employee.includes('Dayana'))?.director).toBe('Angélica Caballero');
    expect(parsed.records.find((record) => record.employee.includes('Tatiana'))?.director).toBe('Óscar Beltrán');
    expect(parsed.unmatchedEmployees.length).toBeGreaterThan(0);
  }, 20_000);
});

