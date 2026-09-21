import { describe, it, expect } from 'vitest';
import type { Remision } from '../types';
import {
  buildDirectorEmailSummary,
  generateDirectorEmailHtml,
} from './directorEmailTemplate';
import { generateDirectorExcelBase64 } from './directorExcelGenerator';

describe('Director Email Notification System', () => {
  const mockRemisiones: Remision[] = [
    {
      cutoff: '2026-09-16',
      employee: 'Dayana Marcela Chala',
      document: '116729',
      nit: '800123456',
      company: 'Makro Supermayorista',
      total: 25000000,
      merchandise: 21008403,
      tax: 3991597,
      age: 25,
      emission: '2026-08-22',
      order: 'PED-101',
    },
    {
      cutoff: '2026-09-16',
      employee: 'Dayana Marcela Chala',
      document: '109459',
      nit: '900987654',
      company: 'Fabrica de Especias',
      total: 12000000,
      merchandise: 10084033,
      tax: 1915967,
      age: 145,
      emission: '2026-04-24',
      order: 'PED-102',
    },
    {
      cutoff: '2026-09-16',
      employee: 'Angela Rocio Torres',
      document: '115800',
      nit: '860001234',
      company: 'Supertiendas Olimpica',
      total: 8000000,
      merchandise: 6722689,
      tax: 1277311,
      age: 18,
      emission: '2026-08-29',
      order: 'PED-103',
    },
    {
      cutoff: '2026-09-16',
      employee: 'Maria Alejandra Velásquez',
      document: '114300',
      nit: '890900123',
      company: 'Exito Colombia',
      total: 15000000,
      merchandise: 12605042,
      tax: 2394958,
      age: 41,
      emission: '2026-08-06',
      order: 'PED-104',
    },
  ];

  it('calcula métricas consolidadas correctamente para Grupo 2 (Angélica Caballero)', () => {
    const summary = buildDirectorEmailSummary(2, mockRemisiones, '2026-09-16');

    expect(summary.directorGroup).toBe(2);
    expect(summary.directorName).toBe('Angélica Caballero');
    expect(summary.totalCount).toBe(4);
    expect(summary.totalValue).toBe(60000000);
    expect(summary.activeExecutivesCount).toBe(3); // Dayana, Ángela, Alejandra
    expect(summary.topMayorValor?.document).toBe('116729');
    expect(summary.topMayorValor?.total).toBe(25000000);
    expect(summary.topMayorAntiguedad?.document).toBe('109459');
    expect(summary.topMayorAntiguedad?.age).toBe(145);
  });

  it('genera plantilla HTML ejecutiva completa y sin la palabra prohibida', () => {
    const summary = buildDirectorEmailSummary(2, mockRemisiones, '2026-09-16');
    const html = generateDirectorEmailHtml(summary, { forWebPreview: true });

    expect(summary.genero).toBe('F');
    expect(html).toContain('Angélica Caballero');
    expect(html).toContain('Dirección Grupo 2');
    expect(html).toContain('Makro Supermayorista');
    expect(html).toContain('Fabrica de Especias');
    expect(html).toContain('Archivo adjunto');
    expect(html).toContain('¡Acompañando a nuestros equipos alcanzamos las metas!');
    expect(html).not.toContain('Gerencia Administrativa y Financiera');

    // Angélica Caballero es mujer -> muñeca (avatar femenino)
    expect(html).toContain('alt="Directora Comercial"');

    // Rafael Novoa es hombre -> muñeco (avatar masculino)
    const summaryNovoa = buildDirectorEmailSummary(1, mockRemisiones, '2026-09-16');
    expect(summaryNovoa.genero).toBe('M');
    const htmlNovoa = generateDirectorEmailHtml(summaryNovoa);
    expect(htmlNovoa).toContain('cid:avatar_man');

    // Regla estricta: Jamás usar el término prohibido
    expect(html.toLowerCase()).not.toContain('cartera');
  });

  it('genera libro Excel consolidado con pestañas de resumen y detalle', async () => {
    const summary = buildDirectorEmailSummary(2, mockRemisiones, '2026-09-16');
    const excel = await generateDirectorExcelBase64(summary);

    expect(excel.filename).toContain('Remisiones_Consolidado_Grupo_2');
    expect(excel.filename.endsWith('.xlsx')).toBe(true);
    expect(excel.base64.length).toBeGreaterThan(100);
  });
});
