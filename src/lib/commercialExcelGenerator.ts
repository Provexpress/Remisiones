import type { Remision } from '../types';

/**
 * Formatea un número como pesos colombianos sin decimales.
 */
function formatCOPNumber(amount: number): string {
  if (amount == null || Number.isNaN(amount)) return '$ 0';
  return new Intl.NumberFormat('es-CO', {
    style: 'currency',
    currency: 'COP',
    maximumFractionDigits: 0,
  }).format(amount).replace('COP', '').replace(/\u00A0/g, ' ').trim();
}

/**
 * Construye el libro de trabajo Excel para las remisiones abiertas de un comercial.
 */
async function buildWorkbook(
  commercialName: string,
  cutoffDate: string,
  remisiones: Remision[],
) {
  const { default: ExcelJSRuntime } = await import('exceljs');
  const wb = new ExcelJSRuntime.Workbook();
  wb.creator = 'Provexpress SAS - Sistema de Gestión de Remisiones';
  wb.created = new Date();

  const ws = wb.addWorksheet('Remisiones Abiertas', {
    properties: { defaultRowHeight: 20 },
    views: [{ state: 'frozen', ySplit: 6 }],
  });

  // 1. Encabezado institucional
  ws.mergeCells('A1:L1');
  const titleCell = ws.getCell('A1');
  titleCell.value = 'PROVEXPRESS SAS · GESTIÓN COMERCIAL DE REMISIONES ABIERTAS';
  titleCell.font = { name: 'Calibri', size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
  titleCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF1E3A8A' }, // Azul institucional oscuro
  };
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getRow(1).height = 30;

  // 2. Metadatos del reporte
  const totalValue = remisiones.reduce((sum, r) => sum + (r.total || 0), 0);
  const avgAge = remisiones.length > 0
    ? Math.round(remisiones.reduce((sum, r) => sum + (r.age || 0), 0) / remisiones.length)
    : 0;

  ws.getCell('A2').value = 'Asesor Comercial:';
  ws.getCell('A2').font = { bold: true, size: 10, color: { argb: 'FF475569' } };
  ws.getCell('B2').value = commercialName;
  ws.getCell('B2').font = { bold: true, size: 11, color: { argb: 'FF0F172A' } };

  ws.getCell('D2').value = 'Fecha de Corte:';
  ws.getCell('D2').font = { bold: true, size: 10, color: { argb: 'FF475569' } };
  ws.getCell('E2').value = cutoffDate;
  ws.getCell('E2').font = { bold: true, size: 11, color: { argb: 'FF0F172A' } };

  ws.getCell('G2').value = 'Total Remisiones:';
  ws.getCell('G2').font = { bold: true, size: 10, color: { argb: 'FF475569' } };
  ws.getCell('H2').value = remisiones.length;
  ws.getCell('H2').font = { bold: true, size: 11, color: { argb: 'FF1E3A8A' } };

  ws.getCell('J2').value = 'Valor Total:';
  ws.getCell('J2').font = { bold: true, size: 10, color: { argb: 'FF475569' } };
  ws.getCell('K2').value = totalValue;
  ws.getCell('K2').numFmt = '"$"#,##0';
  ws.getCell('K2').font = { bold: true, size: 11, color: { argb: 'FF15803D' } };

  ws.getCell('A3').value = 'Antigüedad Promedio:';
  ws.getCell('A3').font = { bold: true, size: 10, color: { argb: 'FF475569' } };
  ws.getCell('B3').value = `${avgAge} días`;
  ws.getCell('B3').font = { bold: true, size: 10, color: { argb: 'FF0F172A' } };

  ws.getCell('D3').value = 'Generado:';
  ws.getCell('D3').font = { bold: true, size: 10, color: { argb: 'FF475569' } };
  ws.getCell('E3').value = new Date().toLocaleString('es-CO');
  ws.getCell('E3').font = { size: 10, color: { argb: 'FF64748B' } };

  // Fila vacía en 4
  ws.getRow(4).height = 10;

  // 3. Encabezados de columnas de la tabla (Fila 6)
  const headers = [
    { key: 'index', title: 'N°', width: 6, align: 'center' },
    { key: 'document', title: 'Remisión', width: 15, align: 'center' },
    { key: 'issuedAt', title: 'Fecha Emisión', width: 14, align: 'center' },
    { key: 'age', title: 'Días Abierta', width: 13, align: 'right' },
    { key: 'ageRange', title: 'Rango Antigüedad', width: 18, align: 'center' },
    { key: 'nit', title: 'NIT', width: 15, align: 'left' },
    { key: 'company', title: 'Cliente / Razón Social', width: 36, align: 'left' },
    { key: 'merchandise', title: 'Vr. Mercancía', width: 16, align: 'right' },
    { key: 'tax', title: 'Vr. IVA', width: 14, align: 'right' },
    { key: 'total', title: 'Vr. Total', width: 16, align: 'right' },
    { key: 'order', title: 'Pedido', width: 14, align: 'left' },
    { key: 'quantity', title: 'Cantidad', width: 11, align: 'right' },
  ];

  const headerRow = ws.getRow(6);
  headerRow.height = 24;

  headers.forEach((h, idx) => {
    const cell = headerRow.getCell(idx + 1);
    cell.value = h.title;
    cell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF0F172A' }, // Slate 900
    };
    cell.alignment = { horizontal: h.align as any, vertical: 'middle' };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
      bottom: { style: 'medium', color: { argb: 'FF94A3B8' } },
      left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
      right: { style: 'thin', color: { argb: 'FFE2E8F0' } },
    };
    ws.getColumn(idx + 1).width = h.width;
  });

  // 4. Agregar filas con los datos de las remisiones
  // Ordenar primero por mayor antigüedad y luego por mayor valor
  const sorted = [...remisiones].sort((a, b) => {
    if (b.age !== a.age) return b.age - a.age;
    return b.total - a.total;
  });

  let currentRowIdx = 7;
  sorted.forEach((r, idx) => {
    const row = ws.getRow(currentRowIdx);
    row.height = 20;

    const isEven = idx % 2 === 0;
    const bgArgb = isEven ? 'FFFFFFFF' : 'FFF8FAFC';

    const docStr = r.document?.startsWith('REM-') ? r.document : `REM-${r.document || 'S/N'}`;

    row.getCell(1).value = idx + 1;
    row.getCell(2).value = docStr;
    row.getCell(3).value = r.issuedAt ? r.issuedAt.slice(0, 10) : '';
    row.getCell(4).value = r.age || 0;
    row.getCell(5).value = r.age > 15 ? 'Más de 15 días' : (r.age >= 8 ? '8 a 15 días' : '0 a 7 días');
    row.getCell(6).value = r.nit || '';
    row.getCell(7).value = r.company || '';
    row.getCell(8).value = r.merchandise || 0;
    row.getCell(9).value = r.tax || 0;
    row.getCell(10).value = r.total || 0;
    row.getCell(11).value = r.order || '';
    row.getCell(12).value = r.quantity || 0;

    // Formato numérico
    row.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
    row.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' };
    row.getCell(2).font = { bold: true, color: { argb: 'FF1E3A8A' } };
    row.getCell(3).alignment = { horizontal: 'center', vertical: 'middle' };
    row.getCell(4).alignment = { horizontal: 'right', vertical: 'middle' };
    row.getCell(4).font = { bold: true, color: r.age > 15 ? { argb: 'FFDC2626' } : { argb: 'FF0F172A' } };
    row.getCell(5).alignment = { horizontal: 'center', vertical: 'middle' };
    row.getCell(6).alignment = { horizontal: 'left', vertical: 'middle' };
    row.getCell(7).alignment = { horizontal: 'left', vertical: 'middle' };

    row.getCell(8).numFmt = '"$"#,##0';
    row.getCell(8).alignment = { horizontal: 'right', vertical: 'middle' };
    row.getCell(9).numFmt = '"$"#,##0';
    row.getCell(9).alignment = { horizontal: 'right', vertical: 'middle' };
    row.getCell(10).numFmt = '"$"#,##0';
    row.getCell(10).alignment = { horizontal: 'right', vertical: 'middle' };
    row.getCell(10).font = { bold: true, color: { argb: 'FF15803D' } };

    row.getCell(11).alignment = { horizontal: 'left', vertical: 'middle' };
    row.getCell(12).numFmt = '#,##0';
    row.getCell(12).alignment = { horizontal: 'right', vertical: 'middle' };

    // Bordes y fondo
    for (let c = 1; c <= 12; c++) {
      const cell = row.getCell(c);
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: bgArgb },
      };
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFF1F5F9' } },
        bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        left: { style: 'thin', color: { argb: 'FFF1F5F9' } },
        right: { style: 'thin', color: { argb: 'FFF1F5F9' } },
      };
    }

    currentRowIdx++;
  });

  // 5. Fila de Totales
  const totalRow = ws.getRow(currentRowIdx);
  totalRow.height = 24;
  totalRow.getCell(1).value = '';
  totalRow.getCell(2).value = 'TOTALES';
  totalRow.getCell(2).font = { bold: true, size: 11, color: { argb: 'FF0F172A' } };
  totalRow.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' };

  totalRow.getCell(4).value = `${avgAge} días prom.`;
  totalRow.getCell(4).font = { bold: true, size: 10, color: { argb: 'FF0F172A' } };
  totalRow.getCell(4).alignment = { horizontal: 'right', vertical: 'middle' };

  const startRow = 7;
  const endRow = currentRowIdx - 1;

  totalRow.getCell(8).value = { formula: `SUM(H${startRow}:H${endRow})`, result: remisiones.reduce((s, r) => s + (r.merchandise || 0), 0) };
  totalRow.getCell(8).numFmt = '"$"#,##0';
  totalRow.getCell(8).font = { bold: true, size: 11, color: { argb: 'FF0F172A' } };

  totalRow.getCell(9).value = { formula: `SUM(I${startRow}:I${endRow})`, result: remisiones.reduce((s, r) => s + (r.tax || 0), 0) };
  totalRow.getCell(9).numFmt = '"$"#,##0';
  totalRow.getCell(9).font = { bold: true, size: 11, color: { argb: 'FF0F172A' } };

  totalRow.getCell(10).value = { formula: `SUM(J${startRow}:J${endRow})`, result: totalValue };
  totalRow.getCell(10).numFmt = '"$"#,##0';
  totalRow.getCell(10).font = { bold: true, size: 11, color: { argb: 'FF15803D' } };

  totalRow.getCell(12).value = { formula: `SUM(L${startRow}:L${endRow})`, result: remisiones.reduce((s, r) => s + (r.quantity || 0), 0) };
  totalRow.getCell(12).numFmt = '#,##0';
  totalRow.getCell(12).font = { bold: true, size: 10, color: { argb: 'FF0F172A' } };

  for (let c = 1; c <= 12; c++) {
    const cell = totalRow.getCell(c);
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFF1F5F9' }, // Fondo gris sutil
    };
    cell.border = {
      top: { style: 'medium', color: { argb: 'FF94A3B8' } },
      bottom: { style: 'double', color: { argb: 'FF0F172A' } },
    };
  }

  return wb;
}

/**
 * Genera el ArrayBuffer del archivo Excel para un comercial.
 */
export async function generateCommercialExcelBuffer(
  commercialName: string,
  cutoffDate: string,
  remisiones: Remision[],
): Promise<ArrayBuffer> {
  const wb = await buildWorkbook(commercialName, cutoffDate, remisiones);
  return await wb.xlsx.writeBuffer();
}

/**
 * Genera el archivo Excel en formato base64 listo para adjuntar a correos de Microsoft Graph.
 */
export async function generateCommercialExcelBase64(
  commercialName: string,
  cutoffDate: string,
  remisiones: Remision[],
): Promise<string> {
  const buffer = await generateCommercialExcelBuffer(commercialName, cutoffDate, remisiones);
  const bytes = new Uint8Array(buffer);
  let binary = '';
  const len = bytes.byteLength;
  for (let i = 0; i < len; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  if (typeof btoa === 'function') {
    return btoa(binary);
  }
  const globalBuf = (globalThis as any).Buffer;
  if (globalBuf) {
    return globalBuf.from(buffer).toString('base64');
  }
  return '';
}

/**
 * Descarga directamente en el navegador el archivo Excel de remisiones abiertas.
 */
export async function downloadCommercialExcel(
  commercialName: string,
  cutoffDate: string,
  remisiones: Remision[],
): Promise<void> {
  const buffer = await generateCommercialExcelBuffer(commercialName, cutoffDate, remisiones);
  const cleanName = commercialName.replace(/\s+/g, '_');
  const filename = `Remisiones_Abiertas_${cleanName}_${cutoffDate || 'actual'}.xlsx`;

  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
