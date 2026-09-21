import { describe, expect, it } from 'vitest';
import { buildGerenciaEmailSummary, generateGerenciaEmailHtml } from './gerenciaEmailTemplate';
import { generateGerenciaExcelBase64 } from './gerenciaExcelGenerator';
import type { Remision } from '../types';

describe('Sistema de Notificación para Dirección y Gerencia Comercial', () => {
  const mockRemisiones: Remision[] = [
    {
      id: '1',
      cutoff: '2026-09-21',
      employee: 'Mario Reyes',
      nit: '900123456',
      company: 'Distribuidora Ejemplo SAS',
      merchandise: 10_000_000,
      tax: 1_900_000,
      total: 11_900_000,
      issuedAt: '2026-08-20',
      age: 32,
      ageRange: 'Más de 30 días',
      daysStatus: 'Gestión comercial (>30 días)',
      document: 'REM-1001',
      order: 'PED-501',
      quantity: 10,
      amountStatus: 'Alto valor (> $5M)',
      alert: 'Gestión comercial · Alto valor',
      sourceRow: 2,
      sourceSheet: 'Base-SIS',
      stableKey: '900123456|REM-1001|PED-501',
      director: 'Rafael Novoa',
      matchedGroup: 1,
    },
    {
      id: '2',
      cutoff: '2026-09-21',
      employee: 'Dayana Marcela Chala',
      nit: '800654321',
      company: 'Comercializadora del Norte SAS',
      merchandise: 20_000_000,
      tax: 3_800_000,
      total: 23_800_000,
      issuedAt: '2026-09-10',
      age: 11,
      ageRange: '8-15 días',
      daysStatus: 'Gestión comercial (8-15 días)',
      document: 'REM-2001',
      order: 'PED-601',
      quantity: 5,
      amountStatus: 'Alto valor (> $5M)',
      alert: 'Gestión comercial · Alto valor',
      sourceRow: 3,
      sourceSheet: 'Base-SIS',
      stableKey: '800654321|REM-2001|PED-601',
      director: 'Angélica Caballero',
      matchedGroup: 2,
    },
    {
      id: '3',
      cutoff: '2026-09-21',
      employee: 'Gina García',
      nit: '860111222',
      company: 'Industrias Centro SAS',
      merchandise: 4_000_000,
      tax: 760_000,
      total: 4_760_000,
      issuedAt: '2026-09-15',
      age: 6,
      ageRange: '0-7 días',
      daysStatus: 'Al día (0-15 días)',
      document: 'REM-3001',
      order: 'PED-701',
      quantity: 2,
      amountStatus: 'Normal (< $5M)',
      alert: 'Al día',
      sourceRow: 4,
      sourceSheet: 'Base-SIS',
      stableKey: '860111222|REM-3001|PED-701',
      director: 'Óscar Beltrán',
      matchedGroup: 3,
    },
  ];

  it('consolida métricas de los 4 grupos y de toda la compañía', () => {
    const summary = buildGerenciaEmailSummary(mockRemisiones, '2026-09-21', {
      name: 'Rafael Novoa',
      cargo: 'Director Comercial',
      email: 'rafael.novoa@provexpress.com.co',
      genero: 'M',
    });

    expect(summary.recipientName).toBe('Rafael Novoa');
    expect(summary.recipientCargo).toBe('Director Comercial');
    expect(summary.totalCount).toBe(3);
    expect(summary.totalValue).toBe(40_460_000);
    expect(summary.groupsSummary).toHaveLength(4);

    // Grupo 1: Novoa
    const g1 = summary.groupsSummary.find((g) => g.grupo === 1);
    expect(g1?.directorName).toBe('Rafael Novoa');
    expect(g1?.totalCount).toBe(1);
    expect(g1?.totalValue).toBe(11_900_000);

    // Grupo 2: Caballero
    const g2 = summary.groupsSummary.find((g) => g.grupo === 2);
    expect(g2?.directorName).toBe('Angélica Caballero');
    expect(g2?.totalCount).toBe(1);
    expect(g2?.totalValue).toBe(23_800_000);

    // Top Oportunidades globales
    expect(summary.topMayorValor?.document).toBe('REM-2001');
    expect(summary.topMayorAntiguedad?.document).toBe('REM-1001');

    // Top Ejecutivos
    expect(summary.topExecutives[0].name).toBe('Dayana Chala');
    expect(summary.topExecutives[0].total).toBe(23_800_000);
  });

  it('genera el HTML para gerencia respetando estándares corporativos y SIN términos prohibidos', () => {
    const summary = buildGerenciaEmailSummary(mockRemisiones, '2026-09-21');
    const html = generateGerenciaEmailHtml(summary);

    expect(html).toContain('Provexpress');
    expect(html).toContain('Consolidado General de Gestión Comercial');
    expect(html).toContain('Resultado Consolidado por Directores de Grupo');
    expect(html).toContain('Rafael Novoa');
    expect(html).toContain('Angélica Caballero');
    expect(html).toContain('Top 5 Asesores');
    expect(html).toContain('Mayor valor por facturar de la compañía');
    expect(html).toContain('Mayor antigüedad pendiente de la compañía');
    expect(html).toContain('Remisiones_Consolidado_General_Gerencia_2026-09-21.xlsx');

    // VERIFICACIÓN ESTRICTA
    expect(html.toLowerCase()).not.toContain('cartera');
  });

  it('genera el libro Excel maestro con 3 hojas', async () => {
    const summary = buildGerenciaEmailSummary(mockRemisiones, '2026-09-21');
    const result = await generateGerenciaExcelBase64(summary);

    expect(result.filename).toBe('Remisiones_Consolidado_General_Gerencia_2026-09-21.xlsx');
    expect(result.base64).toBeTruthy();
    expect(typeof result.base64).toBe('string');
  });
});
