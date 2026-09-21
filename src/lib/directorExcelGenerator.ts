import type { Remision } from '../types';
import type { DirectorEmailSummary } from './directorEmailTemplate';

/**
 * Construye el libro de trabajo Excel consolidado para el Director de Grupo.
 * Incluye:
 * - Hoja 1: Resumen Ejecutivo por Asesor Comercial
 * - Hoja 2: Detalle Completo de Remisiones del Grupo
 */
async function buildDirectorWorkbook(summary: DirectorEmailSummary) {
  const { default: ExcelJSRuntime } = await import('exceljs');
  const wb = new ExcelJSRuntime.Workbook();
  wb.creator = 'Provexpress SAS - Sistema de Gestión de Remisiones';
  wb.created = new Date();

  // ═══════════════════════════════════════════════════════════════════════════
  // HOJA 1: RESUMEN EJECUTIVO POR ASESOR COMERCIAL
  // ═══════════════════════════════════════════════════════════════════════════
  const wsSummary = wb.addWorksheet('Resumen por Ejecutivo', {
    properties: { defaultRowHeight: 20 },
    views: [{ state: 'frozen', ySplit: 5 }],
  });

  // Título institucional
  wsSummary.mergeCells('A1:I1');
  const t1 = wsSummary.getCell('A1');
  t1.value = `PROVEXPRESS SAS · CONSOLIDADO DE REMISIONES · DIRECCIÓN GRUPO ${summary.directorGroup}`;
  t1.font = { name: 'Calibri', size: 13, bold: true, color: { argb: 'FFFFFFFF' } };
  t1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
  t1.alignment = { horizontal: 'center', vertical: 'middle' };
  wsSummary.getRow(1).height = 28;

  // Metadatos
  wsSummary.getCell('A2').value = 'Director de Grupo:';
  wsSummary.getCell('A2').font = { bold: true, size: 10, color: { argb: 'FF475569' } };
  wsSummary.getCell('B2').value = summary.directorName;
  wsSummary.getCell('B2').font = { bold: true, size: 11, color: { argb: 'FF0F172A' } };

  wsSummary.getCell('D2').value = 'Fecha de Corte:';
  wsSummary.getCell('D2').font = { bold: true, size: 10, color: { argb: 'FF475569' } };
  wsSummary.getCell('E2').value = summary.cutoffDate;
  wsSummary.getCell('E2').font = { bold: true, size: 11, color: { argb: 'FF0F172A' } };

  wsSummary.getCell('G2').value = 'Total Remisiones:';
  wsSummary.getCell('G2').font = { bold: true, size: 10, color: { argb: 'FF475569' } };
  wsSummary.getCell('H2').value = summary.totalCount;
  wsSummary.getCell('H2').font = { bold: true, size: 11, color: { argb: 'FF1E3A8A' } };

  wsSummary.getCell('A3').value = 'Total Ejecutivos:';
  wsSummary.getCell('A3').font = { bold: true, size: 10, color: { argb: 'FF475569' } };
  wsSummary.getCell('B3').value = `${summary.activeExecutivesCount} con pendientes de ${summary.totalExecutivesCount} activos`;
  wsSummary.getCell('B3').font = { bold: true, size: 10, color: { argb: 'FF0F172A' } };

  wsSummary.getCell('D3').value = 'Antigüedad Promedio:';
  wsSummary.getCell('D3').font = { bold: true, size: 10, color: { argb: 'FF475569' } };
  wsSummary.getCell('E3').value = `${summary.avgAge} días`;
  wsSummary.getCell('E3').font = { bold: true, size: 10, color: { argb: 'FFDC2626' } };

  wsSummary.getCell('G3').value = 'Valor Total Grupo:';
  wsSummary.getCell('G3').font = { bold: true, size: 10, color: { argb: 'FF475569' } };
  wsSummary.getCell('H3').value = summary.totalValue;
  wsSummary.getCell('H3').numFmt = '"$"#,##0';
  wsSummary.getCell('H3').font = { bold: true, size: 11, color: { argb: 'FF15803D' } };

  // Fila vacía separadora
  wsSummary.getRow(4).height = 8;

  // Cabecera de la tabla de resumen
  const summaryHeaders = [
    'N°',
    'Asesor Comercial',
    'Correo Corporativo',
    'Remisiones Abiertas',
    'Valor Pendiente ($ COP)',
    '% Participación',
    'Antigüedad Promedio',
    'Remisión Crítica (Más Días)',
    'Estado',
  ];

  const headerRow1 = wsSummary.getRow(5);
  headerRow1.height = 24;
  summaryHeaders.forEach((text, i) => {
    const cell = headerRow1.getCell(i + 1);
    cell.value = text;
    cell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
    cell.alignment = {
      horizontal: i === 0 || i === 3 || i === 5 || i === 6 || i === 8 ? 'center' : i === 4 ? 'right' : 'left',
      vertical: 'middle',
    };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FF94A3B8' } },
      bottom: { style: 'medium', color: { argb: 'FF0F172A' } },
      left: { style: 'thin', color: { argb: 'FF94A3B8' } },
      right: { style: 'thin', color: { argb: 'FF94A3B8' } },
    };
  });

  // Filas de ejecutivos
  let currentSummaryRow = 6;
  summary.executives.forEach((exec, idx) => {
    const row = wsSummary.getRow(currentSummaryRow);
    row.height = 20;

    const isEven = idx % 2 === 0;
    const rowBg = isEven ? 'FFFFFFFF' : 'FFF8FAFC';
    const pct = summary.totalValue > 0 ? (exec.total / summary.totalValue) : 0;

    let statusText = 'En gestión';
    if (exec.count === 0) statusText = 'Al día ✓';
    else if (exec.maxAge > 30) statusText = 'Urgente (>30d)';

    const criticalDoc = exec.criticalRemision
      ? `${exec.criticalRemision.document?.startsWith('REM-') ? exec.criticalRemision.document : `REM-${exec.criticalRemision.document || 'S/N'}`} (${exec.criticalRemision.age}d)`
      : '—';

    row.getCell(1).value = idx + 1;
    row.getCell(2).value = exec.name;
    row.getCell(3).value = exec.email;
    row.getCell(4).value = exec.count;
    row.getCell(5).value = exec.total;
    row.getCell(5).numFmt = '"$"#,##0';
    row.getCell(6).value = pct;
    row.getCell(6).numFmt = '0.0%';
    row.getCell(7).value = exec.count > 0 ? `${exec.avgAge} días` : '—';
    row.getCell(8).value = criticalDoc;
    row.getCell(9).value = statusText;

    for (let c = 1; c <= 9; c++) {
      const cell = row.getCell(c);
      cell.font = { name: 'Calibri', size: 10, color: { argb: 'FF0F172A' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg } };
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        right: { style: 'thin', color: { argb: 'FFE2E8F0' } },
      };
      if (c === 1 || c === 4 || c === 6 || c === 7 || c === 9) {
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      } else if (c === 5) {
        cell.alignment = { horizontal: 'right', vertical: 'middle' };
        if (exec.count > 0) cell.font = { bold: true, color: { argb: 'FF15803D' } };
      } else {
        cell.alignment = { horizontal: 'left', vertical: 'middle' };
      }
    }

    currentSummaryRow++;
  });

  // Fila de totales del resumen
  const totalSummaryRow = wsSummary.getRow(currentSummaryRow);
  totalSummaryRow.height = 24;
  totalSummaryRow.getCell(1).value = '';
  totalSummaryRow.getCell(2).value = 'TOTAL GRUPO';
  totalSummaryRow.getCell(2).font = { bold: true, size: 10, color: { argb: 'FFFFFFFF' } };
  totalSummaryRow.getCell(3).value = '';
  totalSummaryRow.getCell(4).value = summary.totalCount;
  totalSummaryRow.getCell(4).font = { bold: true, size: 10, color: { argb: 'FFFFFFFF' } };
  totalSummaryRow.getCell(4).alignment = { horizontal: 'center', vertical: 'middle' };
  totalSummaryRow.getCell(5).value = summary.totalValue;
  totalSummaryRow.getCell(5).numFmt = '"$"#,##0';
  totalSummaryRow.getCell(5).font = { bold: true, size: 11, color: { argb: 'FFFFFFFF' } };
  totalSummaryRow.getCell(5).alignment = { horizontal: 'right', vertical: 'middle' };
  totalSummaryRow.getCell(6).value = 1.0;
  totalSummaryRow.getCell(6).numFmt = '0.0%';
  totalSummaryRow.getCell(6).font = { bold: true, size: 10, color: { argb: 'FFFFFFFF' } };
  totalSummaryRow.getCell(6).alignment = { horizontal: 'center', vertical: 'middle' };
  totalSummaryRow.getCell(7).value = `${summary.avgAge} días`;
  totalSummaryRow.getCell(7).font = { bold: true, size: 10, color: { argb: 'FFFFFFFF' } };
  totalSummaryRow.getCell(7).alignment = { horizontal: 'center', vertical: 'middle' };
  totalSummaryRow.getCell(8).value = '';
  totalSummaryRow.getCell(9).value = '';

  for (let c = 1; c <= 9; c++) {
    const cell = totalSummaryRow.getCell(c);
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
    cell.border = {
      top: { style: 'medium', color: { argb: 'FF0F172A' } },
      bottom: { style: 'medium', color: { argb: 'FF0F172A' } },
    };
  }

  // Anchos de columna hoja 1
  wsSummary.columns = [
    { width: 6 },  // N°
    { width: 26 }, // Asesor
    { width: 32 }, // Email
    { width: 20 }, // Remisiones
    { width: 22 }, // Valor Total
    { width: 16 }, // % Part
    { width: 20 }, // Días Prom
    { width: 24 }, // Remisión Crítica
    { width: 18 }, // Estado
  ];

  // ═══════════════════════════════════════════════════════════════════════════
  // HOJA 2: DETALLE COMPLETO DE REMISIONES DEL GRUPO
  // ═══════════════════════════════════════════════════════════════════════════
  const wsDetail = wb.addWorksheet('Detalle de Remisiones', {
    properties: { defaultRowHeight: 20 },
    views: [{ state: 'frozen', ySplit: 3 }],
  });

  // Título detalle
  wsDetail.mergeCells('A1:L1');
  const t2 = wsDetail.getCell('A1');
  t2.value = `PROVEXPRESS SAS · DETALLE DE REMISIONES ABIERTAS · GRUPO ${summary.directorGroup}`;
  t2.font = { name: 'Calibri', size: 13, bold: true, color: { argb: 'FFFFFFFF' } };
  t2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
  t2.alignment = { horizontal: 'center', vertical: 'middle' };
  wsDetail.getRow(1).height = 28;

  wsDetail.getRow(2).height = 6;

  const detailHeaders = [
    'N°',
    'Asesor Comercial',
    'Remisión',
    'NIT Cliente',
    'Nombre / Razón Social del Cliente',
    'Vr. Mercancía',
    'Vr. IVA',
    'Vr. Total ($ COP)',
    'Fecha Emisión',
    'Días Abierta',
    'Rango Antigüedad',
    'N° Pedido',
  ];

  const headerRow2 = wsDetail.getRow(3);
  headerRow2.height = 24;
  detailHeaders.forEach((text, i) => {
    const cell = headerRow2.getCell(i + 1);
    cell.value = text;
    cell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
    cell.alignment = {
      horizontal: i === 0 || i === 2 || i === 8 || i === 9 || i === 10 || i === 11 ? 'center' : i >= 5 && i <= 7 ? 'right' : 'left',
      vertical: 'middle',
    };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FF94A3B8' } },
      bottom: { style: 'medium', color: { argb: 'FF0F172A' } },
      left: { style: 'thin', color: { argb: 'FF94A3B8' } },
      right: { style: 'thin', color: { argb: 'FF94A3B8' } },
    };
  });

  // Ordenar remisiones del grupo por antigüedad desc, luego total desc
  const sortedDetail = [...summary.allGroupRemisiones].sort((a, b) => {
    if ((b.age || 0) !== (a.age || 0)) return (b.age || 0) - (a.age || 0);
    return (b.total || 0) - (a.total || 0);
  });

  let currentDetailRow = 4;
  sortedDetail.forEach((r, idx) => {
    const row = wsDetail.getRow(currentDetailRow);
    row.height = 20;

    const isEven = idx % 2 === 0;
    const rowBg = isEven ? 'FFFFFFFF' : 'FFF8FAFC';

    const age = r.age || 0;
    let ageRange = '0 a 7 días';
    if (age > 30) ageRange = 'Más de 30 días';
    else if (age > 15) ageRange = '16 a 30 días';
    else if (age > 7) ageRange = '8 a 15 días';

    const docStr = r.document?.startsWith('REM-') ? r.document : `REM-${r.document || 'S/N'}`;

    row.getCell(1).value = idx + 1;
    row.getCell(2).value = r.employee || 'Sin Asignar';
    row.getCell(3).value = docStr;
    row.getCell(4).value = r.nit || '—';
    row.getCell(5).value = r.company || '—';
    row.getCell(6).value = r.merchandise || 0;
    row.getCell(6).numFmt = '"$"#,##0';
    row.getCell(7).value = r.tax || 0;
    row.getCell(7).numFmt = '"$"#,##0';
    row.getCell(8).value = r.total || 0;
    row.getCell(8).numFmt = '"$"#,##0';
    row.getCell(9).value = r.issuedAt || '—';
    row.getCell(10).value = age;
    row.getCell(11).value = ageRange;
    row.getCell(12).value = r.order || '—';

    for (let c = 1; c <= 12; c++) {
      const cell = row.getCell(c);
      cell.font = { name: 'Calibri', size: 10, color: { argb: 'FF0F172A' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg } };
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        right: { style: 'thin', color: { argb: 'FFE2E8F0' } },
      };

      if (c === 1 || c === 2 || c === 3 || c === 9 || c === 10 || c === 11 || c === 12) {
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      } else if (c >= 6 && c <= 8) {
        cell.alignment = { horizontal: 'right', vertical: 'middle' };
        if (c === 8) cell.font = { bold: true, color: { argb: 'FF15803D' } };
      } else {
        cell.alignment = { horizontal: 'left', vertical: 'middle' };
      }
    }

    currentDetailRow++;
  });

  // Fila total del detalle
  const totalDetailRow = wsDetail.getRow(currentDetailRow);
  totalDetailRow.height = 24;
  totalDetailRow.getCell(1).value = '';
  totalDetailRow.getCell(2).value = 'TOTALES DEL GRUPO';
  totalDetailRow.getCell(2).font = { bold: true, size: 10, color: { argb: 'FFFFFFFF' } };
  totalDetailRow.getCell(8).value = summary.totalValue;
  totalDetailRow.getCell(8).numFmt = '"$"#,##0';
  totalDetailRow.getCell(8).font = { bold: true, size: 11, color: { argb: 'FFFFFFFF' } };
  totalDetailRow.getCell(8).alignment = { horizontal: 'right', vertical: 'middle' };
  totalDetailRow.getCell(10).value = `${summary.avgAge} días prom.`;
  totalDetailRow.getCell(10).font = { bold: true, size: 10, color: { argb: 'FFFFFFFF' } };
  totalDetailRow.getCell(10).alignment = { horizontal: 'center', vertical: 'middle' };

  for (let c = 1; c <= 12; c++) {
    const cell = totalDetailRow.getCell(c);
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
    cell.border = {
      top: { style: 'medium', color: { argb: 'FF0F172A' } },
      bottom: { style: 'medium', color: { argb: 'FF0F172A' } },
    };
  }

  // Anchos detalle
  wsDetail.columns = [
    { width: 6 },  // N°
    { width: 26 }, // Asesor
    { width: 16 }, // Remisión
    { width: 16 }, // NIT
    { width: 34 }, // Cliente
    { width: 16 }, // Mercancía
    { width: 14 }, // IVA
    { width: 18 }, // Total
    { width: 14 }, // Emisión
    { width: 14 }, // Días
    { width: 18 }, // Rango
    { width: 16 }, // Pedido
  ];

  return wb;
}

/**
 * Genera el archivo Excel en formato Base64 para adjuntarlo vía Microsoft Graph API.
 */
export async function generateDirectorExcelBase64(summary: DirectorEmailSummary): Promise<{
  filename: string;
  base64: string;
}> {
  const wb = await buildDirectorWorkbook(summary);
  const buffer = await wb.xlsx.writeBuffer();

  const binary = Array.from(new Uint8Array(buffer))
    .map((b) => String.fromCharCode(b))
    .join('');
  const base64 = btoa(binary);

  const cleanName = summary.directorName.replace(/\s+/g, '_');
  const filename = `Remisiones_Consolidado_Grupo_${summary.directorGroup}_${cleanName}_${summary.cutoffDate || 'Corte'}.xlsx`;

  return { filename, base64 };
}

/**
 * Descarga directamente el archivo Excel en el navegador.
 */
export async function downloadDirectorExcel(summary: DirectorEmailSummary): Promise<void> {
  const wb = await buildDirectorWorkbook(summary);
  const buffer = await wb.xlsx.writeBuffer();

  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  const cleanName = summary.directorName.replace(/\s+/g, '_');
  a.download = `Remisiones_Consolidado_Grupo_${summary.directorGroup}_${cleanName}_${summary.cutoffDate || 'Corte'}.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
