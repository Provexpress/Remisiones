import { describe, expect, it } from 'vitest';
import { getCommercialInfo, getExecutiveEmailByName } from './commercialDirectory';
import { buildCommercialEmailSummary, formatCOP, generateCommercialEmailHtml } from './commercialEmailTemplate';
import type { Remision } from '../types';

describe('directorio comercial y emparejamiento', () => {
  it('encuentra el correo corporativo a partir de nombres y variantes de Excel', () => {
    expect(getExecutiveEmailByName('Mario Reyes Gutierre')).toBe('mario.reyes@provexpress.com.co');
    expect(getExecutiveEmailByName('Dayana Marcela Chala')).toBe('dayana.chala@provexpress.com.co');
    expect(getExecutiveEmailByName('Wilson Fernando Sanc')).toBe('wilson.sanchez@provexpress.com.co');
    expect(getExecutiveEmailByName('Rosmira Rojas Puente')).toBe('rosmira.rojas@provexpress.com.co');
    expect(getExecutiveEmailByName('Tatiana Parra')).toBe('tatiana.parra@provexpress.com.co');
    expect(getExecutiveEmailByName('Miller Romero')).toBe(''); // Es director, no ejecutivo
  });

  it('obtiene la información completa del comercial y su director', () => {
    const infoMario = getCommercialInfo('Mario Reyes');
    expect(infoMario.email).toBe('mario.reyes@provexpress.com.co');
    expect(infoMario.grupo).toBe(1);
    expect(infoMario.directorNombre).toBe('Rafael Novoa');

    const infoDayana = getCommercialInfo('dayana.chala@provexpress.com.co');
    expect(infoDayana.nombre).toBe('Dayana Chala');
    expect(infoDayana.grupo).toBe(2);
    expect(infoDayana.directorNombre).toBe('Angélica Caballero');
  });
});

describe('generador de plantilla HTML corporativa de notificación', () => {
  const mockRemisiones: Remision[] = [
    {
      id: '1',
      cutoff: '2026-09-16',
      employee: 'Mario Reyes',
      nit: '900123456',
      company: 'Distribuidora Ejemplo SAS',
      merchandise: 10_000_000,
      tax: 1_900_000,
      total: 11_900_000,
      issuedAt: '2026-08-20',
      age: 27,
      ageRange: '16-30 días',
      daysStatus: 'Gestión comercial (16-30 días)',
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
      cutoff: '2026-09-16',
      employee: 'Mario Reyes',
      nit: '800654321',
      company: 'Comercializadora del Norte',
      merchandise: 5_000_000,
      tax: 950_000,
      total: 5_950_000,
      issuedAt: '2026-09-10',
      age: 6,
      ageRange: '0-7 días',
      daysStatus: 'Al día (0-15 días)',
      document: 'REM-1002',
      order: 'PED-502',
      quantity: 5,
      amountStatus: 'Alto valor (> $5M)',
      alert: 'Al día · Alto valor',
      sourceRow: 3,
      sourceSheet: 'Base-SIS',
      stableKey: '800654321|REM-1002|PED-502',
      director: 'Rafael Novoa',
      matchedGroup: 1,
    },
  ];

  it('construye el resumen de métricas del comercial para el corte', () => {
    const summary = buildCommercialEmailSummary('Mario Reyes', mockRemisiones, '2026-09-16');
    expect(summary.commercialName).toBe('Mario Reyes');
    expect(summary.commercialEmail).toBe('mario.reyes@provexpress.com.co');
    expect(summary.directorName).toBe('Rafael Novoa');
    expect(summary.totalCount).toBe(2);
    expect(summary.totalValue).toBe(17_850_000);
    expect(summary.avgAge).toBe(17);
    expect(summary.destacadas).toHaveLength(2);
    expect(summary.topOpportunity?.document).toBe('REM-1001');
  });

  it('formatea montos COP con separadores colombianos', () => {
    expect(formatCOP(58_320_000)).toContain('58.320.000');
    expect(formatCOP(0)).toBe('$ 0');
  });

  it('genera el HTML con las secciones corporativas exactas y SIN términos prohibidos', () => {
    const summary = buildCommercialEmailSummary('Mario Reyes', mockRemisiones, '2026-09-16');
    const html = generateCommercialEmailHtml(summary);

    // Encabezado corporativo y saludo
    expect(html).toContain('PROVEXPRESS');
    expect(html).toContain('Hola, <strong style="color: #0F172A; font-size: 17px;">Mario Reyes</strong>');
    expect(html).toContain('Oportunidades listas para convertirse en <span style="color: #15803D;">ventas facturadas</span>');
    expect(html).toContain('Tu gestión hace la diferencia.');

    // Tarjetas KPI
    expect(html).toContain('Listas para gestionar');
    expect(html).toContain('Valor por facturar');
    expect(html).toContain('Antigüedad promedio');
    expect(html).toContain('Ventas por reconocer');

    // Tabla de remisiones
    expect(html).toContain('Tus remisiones destacadas');
    expect(html).toContain('REM-1001');
    expect(html).toContain('Distribuidora Ejemplo SAS');
    expect(html).toContain('Más de 15 días');

    // Hero banner y pie
    expect(html).toContain('Top oportunidad del día');
    expect(html).toContain('Cada factura cuenta');
    expect(html).toContain('Gerencia Administrativa y Financiera');

    // VERIFICACIÓN ESTRICTA: cero ocurrencias del término prohibido
    expect(html.toLowerCase()).not.toContain('cartera');
  });
});
