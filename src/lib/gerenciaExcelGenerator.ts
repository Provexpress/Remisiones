import type { Remision } from '../types';
import type { GerenciaEmailSummary } from './gerenciaEmailTemplate';
import { ESTRUCTURA_COMERCIAL_2026, LISTA_DIRECTORES, normalizeName } from './commercialDirectory';

/**
 * Construye el libro de trabajo Excel maestro para Dirección y Gerencia Comercial.
 * Incluye 3 hojas de trabajo:
 * - Hoja 1: Resumen por Director / Grupo
 * - Hoja 2: Resumen por Ejecutivo Comercial
 * - Hoja 3: Detalle General de Remisiones de toda la compañía
 */
async function buildGerenciaWorkbook(summary: GerenciaEmailSummary) {
  const { default: ExcelJSRuntime } = await import('exceljs');
  const wb = new ExcelJSRuntime.Workbook();
  wb.creator = 'Provexpress SAS - Sistema de Gestión de Remisiones';
  wb.created = new Date();

  // ═══════════════════════════════════════════════════════════════════════════
  // HOJA 1: RESUMEN POR DIRECTOR / GRUPO
  // ═══════════════════════════════════════════════════════════════════════════
  const wsDirectors = wb.addWorksheet('Resumen por Director', {
    properties: { defaultRowHeight: 20 },
    views: [{ state: 'frozen', ySplit: 5 }],
  });

  // Título institucional
  wsDirectors.mergeCells('A1:K1');
  const t1 = wsDirectors.getCell('A1');
  t1.value = 'PROVEXPRESS SAS · CONSOLIDADO GERENCIAL DE REMISIONES · ALTA DIRECCIÓN';
  t1.font = { name: 'Calibri', size: 13, bold: true, color: { argb: 'FFFFFFFF' } };
  t1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
  t1.alignment = { horizontal: 'center', vertical: 'middle' };
  wsDirectors.getRow(1).height = 28;

  // Metadatos
  wsDirectors.getCell('A2').value = 'Destinatario:';
  wsDirectors.getCell('A2').font = { bold: true, size: 10, color: { argb: 'FF475569' } };
  wsDirectors.getCell('B2').value = `${summary.recipientName} (${summary.recipientCargo})`;
  wsDirectors.getCell('B2').font = { bold: true, size: 11, color: { argb: 'FF0F172A' } };

  wsDirectors.getCell('E2').value = 'Fecha de Corte:';
  wsDirectors.getCell('E2').font = { bold: true, size: 10, color: { argb: 'FF475569' } };
  wsDirectors.getCell('F2').value = summary.cutoffDate || 'Consolidado';
  wsDirectors.getCell('F2').font = { bold: true, size: 11, color: { argb: 'FF0F172A' } };

  wsDirectors.getCell('I2').value = 'Remisiones Empresa:';
  wsDirectors.getCell('I2').font = { bold: true, size: 10, color: { argb: 'FF475569' } };
  wsDirectors.getCell('J2').value = summary.totalCount;
  wsDirectors.getCell('J2').font = { bold: true, size: 11, color: { argb: 'FF1E3A8A' } };

  wsDirectors.getCell('A3').value = 'Grupos Comerciales:';
  wsDirectors.getCell('A3').font = { bold: true, size: 10, color: { argb: 'FF475569' } };
  wsDirectors.getCell('B3').value = `${summary.activeDirectorsCount} de ${summary.totalDirectorsCount} con gestión activa`;
  wsDirectors.getCell('B3').font = { bold: true, size: 10, color: { argb: 'FF0F172A' } };

  wsDirectors.getCell('E3').value = 'Antigüedad Promedio:';
  wsDirectors.getCell('E3').font = { bold: true, size: 10, color: { argb: 'FF475569' } };
  wsDirectors.getCell('F3').value = `${summary.avgAge} días`;
  wsDirectors.getCell('F3').font = { bold: true, size: 10, color: { argb: 'FFDC2626' } };

  wsDirectors.getCell('I3').value = 'Total por Facturar:';
  wsDirectors.getCell('I3').font = { bold: true, size: 10, color: { argb: 'FF475569' } };
  wsDirectors.getCell('J3').value = summary.totalValue;
  wsDirectors.getCell('J3').numFmt = '"$"#,##0';
  wsDirectors.getCell('J3').font = { bold: true, size: 11, color: { argb: 'FF15803D' } };

  wsDirectors.getRow(4).height = 8;

  const dirHeaders = [
    'N°',
    'Grupo',
    'Director Comercial',
    'Correo Corporativo',
    'Asesores Activos',
    'Total Asesores',
    'Remisiones Abiertas',
    'Valor por Facturar ($ COP)',
    '% Participación',
    'Antigüedad Promedio',
    'Estado',
  ];

  const headerRow1 = wsDirectors.getRow(5);
  headerRow1.height = 24;
  dirHeaders.forEach((text, i) => {
    const cell = headerRow1.getCell(i + 1);
    cell.value = text;
    cell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F172A' } };
    cell.alignment = {
      horizontal: i === 0 || i === 1 || i === 4 || i === 5 || i === 6 || i === 8 || i === 9 || i === 10 ? 'center' : i === 7 ? 'right' : 'left',
      vertical: 'middle',
    };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FF94A3B8' } },
      bottom: { style: 'medium', color: { argb: 'FF0F172A' } },
      left: { style: 'thin', color: { argb: 'FF94A3B8' } },
      right: { style: 'thin', color: { argb: 'FF94A3B8' } },
    };
  });

  summary.groupsSummary.forEach((g, idx) => {
    const row = wsDirectors.getRow(6 + idx);
    row.height = 22;
    const isEven = idx % 2 === 0;
    const bgArgb = isEven ? 'FFFFFFFF' : 'FFF8FAFC';

    row.getCell(1).value = idx + 1;
    row.getCell(2).value = `Grupo ${g.grupo}`;
    row.getCell(2).font = { bold: true, color: { argb: 'FF1E3A8A' } };
    row.getCell(3).value = g.directorName;
    row.getCell(3).font = { bold: true, color: { argb: 'FF0F172A' } };
    row.getCell(4).value = g.directorEmail;
    row.getCell(4).font = { color: { argb: 'FF475569' } };
    row.getCell(5).value = g.activeExecutivesCount;
    row.getCell(5).font = { bold: true };
    row.getCell(6).value = g.totalExecutivesCount;
    row.getCell(7).value = g.totalCount;
    row.getCell(7).font = { bold: true };
    row.getCell(8).value = g.totalValue;
    row.getCell(8).numFmt = '"$"#,##0';
    row.getCell(8).font = { bold: true, color: { argb: 'FF15803D' } };
    row.getCell(9).value = `${g.pctCompanyValue}%`;
    row.getCell(9).font = { bold: true, color: { argb: 'FF1E3A8A' } };
    row.getCell(10).value = `${g.avgAge} días`;
    row.getCell(10).font = { bold: true, color: g.avgAge > 25 ? { argb: 'FFDC2626' } : { argb: 'FF0F172A' } };
    row.getCell(11).value = g.status;
    row.getCell(11).font = { bold: true, color: g.status === 'Al día' ? { argb: 'FF15803D' } : g.status === 'Atención prioritaria' ? { argb: 'FFDC2626' } : { argb: 'FF1E3A8A' } };

    for (let col = 1; col <= 11; col++) {
      const cell = row.getCell(col);
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgArgb } };
      cell.alignment = {
        horizontal: col === 1 || col === 2 || col === 5 || col === 6 || col === 7 || col === 9 || col === 10 || col === 11 ? 'center' : col === 8 ? 'right' : 'left',
        vertical: 'middle',
      };
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        right: { style: 'thin', color: { argb: 'FFE2E8F0' } },
      };
    }
  });

  // Fila Total de Directores
  const totalDirRowIdx = 6 + summary.groupsSummary.length;
  const totalDirRow = wsDirectors.getRow(totalDirRowIdx);
  totalDirRow.height = 24;

  totalDirRow.getCell(1).value = 'TOTAL';
  totalDirRow.getCell(2).value = 'EMPRESA';
  totalDirRow.getCell(3).value = 'CONSOLIDADO 4 GRUPOS';
  totalDirRow.getCell(3).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  totalDirRow.getCell(5).value = summary.activeExecutivesCount;
  totalDirRow.getCell(6).value = summary.totalExecutivesCount;
  totalDirRow.getCell(7).value = summary.totalCount;
  totalDirRow.getCell(8).value = summary.totalValue;
  totalDirRow.getCell(8).numFmt = '"$"#,##0';
  totalDirRow.getCell(9).value = '100%';
  totalDirRow.getCell(10).value = `${summary.avgAge} días`;
  totalDirRow.getCell(11).value = 'GENERAL';

  for (let col = 1; col <= 11; col++) {
    const cell = totalDirRow.getCell(col);
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F172A' } };
    cell.font = { bold: true, size: 10, color: { argb: 'FFFFFFFF' } };
    cell.alignment = {
      horizontal: col === 1 || col === 2 || col === 5 || col === 6 || col === 7 || col === 9 || col === 10 || col === 11 ? 'center' : col === 8 ? 'right' : 'left',
      vertical: 'middle',
    };
    cell.border = {
      top: { style: 'medium', color: { argb: 'FF0F172A' } },
      bottom: { style: 'medium', color: { argb: 'FF0F172A' } },
    };
  }

  wsDirectors.columns = [
    { width: 6 },
    { width: 14 },
    { width: 26 },
    { width: 34 },
    { width: 16 },
    { width: 16 },
    { width: 18 },
    { width: 22 },
    { width: 16 },
    { width: 20 },
    { width: 20 },
  ];

  // ═══════════════════════════════════════════════════════════════════════════
  // HOJA 2: RESUMEN POR EJECUTIVO COMERCIAL (TODOS LOS EJECUTIVOS DE LA EMPRESA)
  // ═══════════════════════════════════════════════════════════════════════════
  const wsExecs = wb.addWorksheet('Resumen por Ejecutivo', {
    properties: { defaultRowHeight: 20 },
    views: [{ state: 'frozen', ySplit: 4 }],
  });

  wsExecs.mergeCells('A1:J1');
  const t2 = wsExecs.getCell('A1');
  t2.value = 'PROVEXPRESS SAS · CONSOLIDADO DE REMISIONES POR ASESOR COMERCIAL';
  t2.font = { name: 'Calibri', size: 13, bold: true, color: { argb: 'FFFFFFFF' } };
  t2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
  t2.alignment = { horizontal: 'center', vertical: 'middle' };
  wsExecs.getRow(1).height = 28;

  wsExecs.getCell('A2').value = `Total Asesores con pendientes: ${summary.activeExecutivesCount} de ${summary.totalExecutivesCount} activos en la compañía`;
  wsExecs.getCell('A2').font = { bold: true, size: 10, color: { argb: 'FF475569' } };
  wsExecs.getRow(3).height = 8;

  const execHeaders = [
    'N°',
    'Grupo',
    'Director Responsable',
    'Asesor Comercial',
    'Correo Corporativo',
    'Remisiones Abiertas',
    'Valor por Facturar ($ COP)',
    '% Participación',
    'Antigüedad Promedio',
    'Estado',
  ];

  const headerRow2 = wsExecs.getRow(4);
  headerRow2.height = 24;
  execHeaders.forEach((text, i) => {
    const cell = headerRow2.getCell(i + 1);
    cell.value = text;
    cell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
    cell.alignment = {
      horizontal: i === 0 || i === 1 || i === 5 || i === 7 || i === 8 || i === 9 ? 'center' : i === 6 ? 'right' : 'left',
      vertical: 'middle',
    };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FF94A3B8' } },
      bottom: { style: 'medium', color: { argb: 'FF0F172A' } },
      left: { style: 'thin', color: { argb: 'FF94A3B8' } },
      right: { style: 'thin', color: { argb: 'FF94A3B8' } },
    };
  });

  // Recopilar todos los ejecutivos con métricas
  const allExecsFullList: {
    grupo: number;
    directorName: string;
    name: string;
    email: string;
    count: number;
    total: number;
    avgAge: number;
    pct: number;
  }[] = [];

  const allExecEntries = Object.entries(ESTRUCTURA_COMERCIAL_2026.ejecutivos);
  for (const [email, data] of allExecEntries) {
    const dirInfo = LISTA_DIRECTORES.find((d) => d.grupo === data.grupo);
    const dirName = dirInfo?.nombre || `Director Grupo ${data.grupo}`;
    const userRem = summary.allCompanyRemisiones.filter((r) => normalizeName(r.employee) === normalizeName(data.nombre) || normalizeName(r.employee) === normalizeName(data.archivo));

    const count = userRem.length;
    const total = userRem.reduce((s, r) => s + (r.total || 0), 0);
    const avgAge = count > 0 ? Math.round(userRem.reduce((s, r) => s + (r.age || 0), 0) / count) : 0;
    const pct = summary.totalValue > 0 ? Number(((total / summary.totalValue) * 100).toFixed(1)) : 0;

    allExecsFullList.push({
      grupo: data.grupo,
      directorName: dirName,
      name: data.nombre,
      email,
      count,
      total,
      avgAge,
      pct,
    });
  }

  // Ordenar por total por facturar desc
  allExecsFullList.sort((a, b) => b.total - a.total);

  allExecsFullList.forEach((e, idx) => {
    const row = wsExecs.getRow(5 + idx);
    row.height = 20;
    const isEven = idx % 2 === 0;
    const bgArgb = isEven ? 'FFFFFFFF' : 'FFF8FAFC';

    row.getCell(1).value = idx + 1;
    row.getCell(2).value = `Grupo ${e.grupo}`;
    row.getCell(3).value = e.directorName;
    row.getCell(4).value = e.name;
    row.getCell(4).font = { bold: true, color: { argb: 'FF0F172A' } };
    row.getCell(5).value = e.email;
    row.getCell(5).font = { color: { argb: 'FF64748B' } };
    row.getCell(6).value = e.count;
    row.getCell(6).font = { bold: true };
    row.getCell(7).value = e.total;
    row.getCell(7).numFmt = '"$"#,##0';
    row.getCell(7).font = { bold: true, color: e.count > 0 ? { argb: 'FF15803D' } : { argb: 'FF94A3B8' } };
    row.getCell(8).value = `${e.pct}%`;
    row.getCell(8).font = { bold: true, color: { argb: 'FF1E3A8A' } };
    row.getCell(9).value = e.count > 0 ? `${e.avgAge} días` : '—';
    row.getCell(9).font = { bold: true, color: e.avgAge > 30 ? { argb: 'FFDC2626' } : { argb: 'FF0F172A' } };
    row.getCell(10).value = e.count === 0 ? 'Al día ✓' : e.avgAge > 30 ? 'Urgente (>30d)' : 'En gestión';
    row.getCell(10).font = { bold: true, color: e.count === 0 ? { argb: 'FF15803D' } : e.avgAge > 30 ? { argb: 'FFDC2626' } : { argb: 'FF1E3A8A' } };

    for (let col = 1; col <= 10; col++) {
      const cell = row.getCell(col);
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgArgb } };
      cell.alignment = {
        horizontal: col === 1 || col === 2 || col === 6 || col === 8 || col === 9 || col === 10 ? 'center' : col === 7 ? 'right' : 'left',
        vertical: 'middle',
      };
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        right: { style: 'thin', color: { argb: 'FFE2E8F0' } },
      };
    }
  });

  // Fila Total de Ejecutivos
  const totalExecRowIdx = 5 + allExecsFullList.length;
  const totalExecRow = wsExecs.getRow(totalExecRowIdx);
  totalExecRow.height = 24;

  totalExecRow.getCell(1).value = 'TOTAL';
  totalExecRow.getCell(2).value = 'TODOS';
  totalExecRow.getCell(3).value = '4 GRUPOS';
  totalExecRow.getCell(4).value = 'CONSOLIDADO EMPRESA';
  totalExecRow.getCell(4).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  totalExecRow.getCell(6).value = summary.totalCount;
  totalExecRow.getCell(7).value = summary.totalValue;
  totalExecRow.getCell(7).numFmt = '"$"#,##0';
  totalExecRow.getCell(8).value = '100%';
  totalExecRow.getCell(9).value = `${summary.avgAge} días`;
  totalExecRow.getCell(10).value = 'GENERAL';

  for (let col = 1; col <= 10; col++) {
    const cell = totalExecRow.getCell(col);
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F172A' } };
    cell.font = { bold: true, size: 10, color: { argb: 'FFFFFFFF' } };
    cell.alignment = {
      horizontal: col === 1 || col === 2 || col === 6 || col === 8 || col === 9 || col === 10 ? 'center' : col === 7 ? 'right' : 'left',
      vertical: 'middle',
    };
    cell.border = {
      top: { style: 'medium', color: { argb: 'FF0F172A' } },
      bottom: { style: 'medium', color: { argb: 'FF0F172A' } },
    };
  }

  wsExecs.columns = [
    { width: 6 },
    { width: 14 },
    { width: 24 },
    { width: 28 },
    { width: 34 },
    { width: 18 },
    { width: 22 },
    { width: 16 },
    { width: 20 },
    { width: 18 },
  ];

  // ═══════════════════════════════════════════════════════════════════════════
  // HOJA 3: DETALLE GENERAL REMISIONES (TODAS LAS REMISIONES ABIERTAS EMPRESA)
  // ═══════════════════════════════════════════════════════════════════════════
  const wsDetail = wb.addWorksheet('Detalle General Remisiones', {
    properties: { defaultRowHeight: 20 },
    views: [{ state: 'frozen', ySplit: 5 }],
  });

  wsDetail.mergeCells('A1:N1');
  const t3 = wsDetail.getCell('A1');
  t3.value = 'PROVEXPRESS SAS · DETALLE COMPLETO DE REMISIONES ABIERTAS · TODA LA COMPAÑÍA';
  t3.font = { name: 'Calibri', size: 13, bold: true, color: { argb: 'FFFFFFFF' } };
  t3.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
  t3.alignment = { horizontal: 'center', vertical: 'middle' };
  wsDetail.getRow(1).height = 28;

  wsDetail.getCell('A2').value = `Total Remisiones: ${summary.totalCount} | Valor Total: ${summary.totalValue}`;
  wsDetail.getCell('A2').font = { bold: true, size: 10, color: { argb: 'FF475569' } };
  wsDetail.getRow(4).height = 8;

  const detailHeaders = [
    'N°',
    'Grupo',
    'Director',
    'Asesor Comercial',
    'Remisión',
    'Fecha Emisión',
    'Días Abierta',
    'Rango Antigüedad',
    'NIT',
    'Cliente / Razón Social',
    'Vr. Mercancía',
    'Vr. IVA',
    'Vr. Total',
    'Pedido',
  ];

  const headerRow3 = wsDetail.getRow(5);
  headerRow3.height = 24;
  detailHeaders.forEach((text, i) => {
    const cell = headerRow3.getCell(i + 1);
    cell.value = text;
    cell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F172A' } };
    cell.alignment = {
      horizontal: i === 0 || i === 1 || i === 4 || i === 5 || i === 6 || i === 7 || i === 8 || i === 13 ? 'center' : i === 10 || i === 11 || i === 12 ? 'right' : 'left',
      vertical: 'middle',
    };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FF94A3B8' } },
      bottom: { style: 'medium', color: { argb: 'FF0F172A' } },
      left: { style: 'thin', color: { argb: 'FF94A3B8' } },
      right: { style: 'thin', color: { argb: 'FF94A3B8' } },
    };
  });

  // Ordenar remisiones: mayor antigüedad desc, luego mayor total desc
  const sortedCompanyRemisiones = [...summary.allCompanyRemisiones].sort((a, b) => {
    const ageA = a.age || 0;
    const ageB = b.age || 0;
    if (ageB !== ageA) return ageB - ageA;
    return (b.total || 0) - (a.total || 0);
  });

  sortedCompanyRemisiones.forEach((r, idx) => {
    const row = wsDetail.getRow(6 + idx);
    row.height = 20;
    const isEven = idx % 2 === 0;
    const bgArgb = isEven ? 'FFFFFFFF' : 'FFF8FAFC';

    const docStr = r.document?.startsWith('REM-') ? r.document : `REM-${r.document || 'S/N'}`;
    const age = r.age || 0;
    const rangeStr = age > 15 ? 'Más de 15 días' : age >= 8 ? '8 a 15 días' : '0 a 7 días';

    row.getCell(1).value = idx + 1;
    row.getCell(2).value = `Grupo ${r.group || 1}`;
    row.getCell(3).value = r.director || 'Dirección Comercial';
    row.getCell(4).value = r.employee || 'Comercial';
    row.getCell(4).font = { bold: true, color: { argb: 'FF0F172A' } };
    row.getCell(5).value = docStr;
    row.getCell(5).font = { bold: true, color: { argb: 'FF1E3A8A' } };
    row.getCell(6).value = r.issuedAt || summary.cutoffDate;
    row.getCell(7).value = age;
    row.getCell(7).font = { bold: true, color: age > 15 ? { argb: 'FFDC2626' } : { argb: 'FF0F172A' } };
    row.getCell(8).value = rangeStr;
    row.getCell(9).value = r.nit || '';
    row.getCell(10).value = r.company || 'Cliente';
    row.getCell(11).value = r.merchandise || r.total || 0;
    row.getCell(11).numFmt = '"$"#,##0';
    row.getCell(12).value = r.tax || 0;
    row.getCell(12).numFmt = '"$"#,##0';
    row.getCell(13).value = r.total || 0;
    row.getCell(13).numFmt = '"$"#,##0';
    row.getCell(13).font = { bold: true, color: { argb: 'FF15803D' } };
    row.getCell(14).value = r.order || '';

    for (let col = 1; col <= 14; col++) {
      const cell = row.getCell(col);
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgArgb } };
      cell.alignment = {
        horizontal: col === 1 || col === 2 || col === 5 || col === 6 || col === 7 || col === 8 || col === 9 || col === 14 ? 'center' : col === 11 || col === 12 || col === 13 ? 'right' : 'left',
        vertical: 'middle',
      };
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        right: { style: 'thin', color: { argb: 'FFE2E8F0' } },
      };
    }
  });

  // Fila total del detalle
  const totalDetailRowIdx = 6 + sortedCompanyRemisiones.length;
  const totalDetailRow = wsDetail.getRow(totalDetailRowIdx);
  totalDetailRow.height = 24;

  totalDetailRow.getCell(1).value = 'TOTAL';
  totalDetailRow.getCell(4).value = 'TOTAL CONSOLIDADO EMPRESA';
  totalDetailRow.getCell(4).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  totalDetailRow.getCell(5).value = `${summary.totalCount} docs`;
  totalDetailRow.getCell(7).value = `${summary.avgAge}d prom.`;
  totalDetailRow.getCell(13).value = summary.totalValue;
  totalDetailRow.getCell(13).numFmt = '"$"#,##0';

  for (let col = 1; col <= 14; col++) {
    const cell = totalDetailRow.getCell(col);
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
    cell.font = { bold: true, size: 10, color: { argb: 'FFFFFFFF' } };
    cell.alignment = {
      horizontal: col === 1 || col === 2 || col === 5 || col === 6 || col === 7 || col === 8 || col === 9 || col === 14 ? 'center' : col === 11 || col === 12 || col === 13 ? 'right' : 'left',
      vertical: 'middle',
    };
    cell.border = {
      top: { style: 'medium', color: { argb: 'FF0F172A' } },
      bottom: { style: 'medium', color: { argb: 'FF0F172A' } },
    };
  }

  wsDetail.columns = [
    { width: 6 },  // N°
    { width: 12 }, // Grupo
    { width: 22 }, // Director
    { width: 26 }, // Asesor
    { width: 16 }, // Remisión
    { width: 14 }, // Emisión
    { width: 12 }, // Días
    { width: 16 }, // Rango
    { width: 16 }, // NIT
    { width: 34 }, // Cliente
    { width: 16 }, // Mercancía
    { width: 14 }, // IVA
    { width: 18 }, // Total
    { width: 14 }, // Pedido
  ];

  return wb;
}

/**
 * Genera el archivo Excel maestro gerencial en formato Base64 para adjuntarlo vía Microsoft Graph API.
 */
export async function generateGerenciaExcelBase64(summary: GerenciaEmailSummary): Promise<{
  filename: string;
  base64: string;
}> {
  const wb = await buildGerenciaWorkbook(summary);
  const buffer = await wb.xlsx.writeBuffer();

  const binary = Array.from(new Uint8Array(buffer))
    .map((b) => String.fromCharCode(b))
    .join('');
  const base64 = btoa(binary);

  const filename = `Remisiones_Consolidado_General_Gerencia_${summary.cutoffDate || 'Corte'}.xlsx`;

  return { filename, base64 };
}

/**
 * Descarga directamente el archivo Excel maestro gerencial en el navegador.
 */
export async function downloadGerenciaExcel(summary: GerenciaEmailSummary): Promise<void> {
  const wb = await buildGerenciaWorkbook(summary);
  const buffer = await wb.xlsx.writeBuffer();

  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `Remisiones_Consolidado_General_Gerencia_${summary.cutoffDate || 'Corte'}.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
