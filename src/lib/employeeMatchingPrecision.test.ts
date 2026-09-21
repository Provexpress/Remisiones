import { existsSync } from 'node:fs';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, it, expect } from 'vitest';
import {
  nameMatches,
  resolveCommercialOrDirector,
  getCommercialInfo,
  ESTRUCTURA_COMERCIAL_2026,
} from './commercialDirectory';
import { buildCommercialEmailSummary } from './commercialEmailTemplate';
import { parseRemisionesWorkbook } from './remisiones';

describe('Employee Matching Precision & Surname Collision Protection', () => {
  it('prevents name collisions between individuals sharing first names', () => {
    // 1. María Angélica Álvarez vs María Angélica Caballero
    expect(nameMatches('Maria Angelica Alvar', 'Angélica Álvarez')).toBe(true);
    expect(nameMatches('Maria Angelica Alvar', 'Maria Angelica Alvarez Morales')).toBe(true);
    expect(nameMatches('Maria Angelica Cabal', 'Angélica Caballero')).toBe(true);
    expect(nameMatches('Maria Angelica Cabal', 'Maria Angelica Caballero')).toBe(true);

    // CRITICAL: They must NEVER match each other despite sharing "Maria Angelica"
    expect(nameMatches('Maria Angelica Alvar', 'Angélica Caballero')).toBe(false);
    expect(nameMatches('Maria Angelica Cabal', 'Angélica Álvarez')).toBe(false);
    expect(nameMatches('Maria Angelica Alvar', 'Maria Angelica Cabal')).toBe(false);

    // 2. Juan David Martínez vs Juan David Novoa Camacho
    expect(nameMatches('Juan David Martínez', 'Juan Martínez')).toBe(true);
    expect(nameMatches('Juan David Novoa Cam', 'Juan Martínez')).toBe(false);
    expect(nameMatches('Juan David Martínez', 'Juan David Novoa Cam')).toBe(false);

    // 3. Jhonatan Steven Acevedo vs Jhonatan Camilo Hernández
    expect(nameMatches('Jhonatan Steven Acev', 'Jhonatan Acevedo')).toBe(true);
    expect(nameMatches('Jhonatan Camilo Hern', 'Camilo Hernández')).toBe(true);
    expect(nameMatches('Jhonatan Steven Acev', 'Camilo Hernández')).toBe(false);
    expect(nameMatches('Jhonatan Camilo Hern', 'Jhonatan Acevedo')).toBe(false);
    expect(nameMatches('Jhonatan Steven Acev', 'Jhonatan Camilo Hern')).toBe(false);

    // 4. Fernando Alberto Quiñonez vs Wilson Fernando Sánchez
    expect(nameMatches('Fernando Alberto Qui', 'Fernando Quiñonez')).toBe(true);
    expect(nameMatches('Wilson Fernando Sanc', 'Wilson Sánchez')).toBe(true);
    expect(nameMatches('Fernando Alberto Qui', 'Wilson Sánchez')).toBe(false);
    expect(nameMatches('Wilson Fernando Sanc', 'Fernando Quiñonez')).toBe(false);
    expect(nameMatches('Fernando Alberto Qui', 'Wilson Fernando Sanc')).toBe(false);

    // 5. Adriana Cucaita vs Adriana Cecilia Ramírez
    expect(nameMatches('Adriana Cucaita Boni', 'Adriana Cucaita')).toBe(true);
    expect(nameMatches('Adriana Cecilia Rami', 'Adriana Cucaita')).toBe(false);

    // 6. Ángela Torres vs María Angélica Álvarez / Caballero
    expect(nameMatches('Angela Rocio Torres', 'Ángela Torres')).toBe(true);
    expect(nameMatches('Angela Rocio Torres', 'Angélica Álvarez')).toBe(false);
    expect(nameMatches('Angela Rocio Torres', 'Angélica Caballero')).toBe(false);
  });

  it('resolves unique corporate emails for all commercial advisors and directors', () => {
    const alvarez = resolveCommercialOrDirector('Maria Angelica Alvar');
    expect(alvarez.type).toBe('ejecutivo');
    expect(alvarez.email).toBe('angelica.alvarez@provexpress.com.co');
    expect(alvarez.grupo).toBe(3);

    const caballero = resolveCommercialOrDirector('Maria Angelica Cabal');
    expect(caballero.type).toBe('director_directa');
    expect(caballero.email).toBe('angelica.caballero@provexpress.com.co');
    expect(caballero.grupo).toBe(2);

    const jmartinez = resolveCommercialOrDirector('Juan David Martínez');
    expect(jmartinez.type).toBe('ejecutivo');
    expect(jmartinez.email).toBe('juan.martinez@provexpress.com.co');

    const novoa = resolveCommercialOrDirector('Juan David Novoa Cam');
    expect(novoa.type).toBe('unassigned');
    expect(novoa.email).toBe('');
  });

  const corporateWorkbookPath = fileURLToPath(new URL('../../Remisiones.xlsx', import.meta.url));
  const hasRealWorkbook = existsSync(corporateWorkbookPath);
  const describeWorkbook = hasRealWorkbook ? describe : describe.skip;

  describeWorkbook('Workbook Verification (Remisiones.xlsx)', () => {
    it('strictly isolates Maria Angelica Alvarez and Maria Angelica Caballero in summary and dispatch', async () => {
      const file = await readFile(corporateWorkbookPath);
      const arrayBuffer = file.buffer.slice(file.byteOffset, file.byteOffset + file.byteLength) as ArrayBuffer;
      const parsed = await parseRemisionesWorkbook(arrayBuffer);

      const cutoff = '2026-09-21';
      const currentRecords = parsed.records.filter((r) => r.cutoff === cutoff);
      if (currentRecords.length === 0) return;

      // Executive Summary for Maria Angelica Alvar
      const summaryAlvarez = buildCommercialEmailSummary('Maria Angelica Alvar', parsed.records, cutoff);
      expect(summaryAlvarez.commercialEmail).toBe('angelica.alvarez@provexpress.com.co');
      expect(summaryAlvarez.totalCount).toBe(14);
      expect(Math.round(summaryAlvarez.totalValue)).toBe(5365691);

      // Verify that NONE of Caballero's records are in Alvarez's remisiones
      const hasCaballeroRecords = summaryAlvarez.allRemisiones.some(
        (r) => r.employee === 'Maria Angelica Cabal' || r.company.toLowerCase().includes('cooperativa nacional') || r.company.toLowerCase().includes('price res')
      );
      expect(hasCaballeroRecords).toBe(false);

      // Verify Caballero's direct records
      const caballeroRecords = currentRecords.filter((r) => r.employee === 'Maria Angelica Cabal');
      expect(caballeroRecords.length).toBe(5);
      expect(Math.round(caballeroRecords.reduce((s, r) => s + r.total, 0))).toBe(164626089);
    }, 30000);
  });
});
