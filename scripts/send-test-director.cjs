/**
 * Script standalone para despachar el correo de prueba para Directores de Grupo a especialista.preventa@provexpress.com.co
 * Utiliza Microsoft Graph API autenticado con las credenciales de Azure Entra ID
 * Ubicadas en C:\Proyectos\m365-report-mailer\.env
 *
 * Muestra el consolidado de ejecutivos comerciales para Angélica Caballero (Directora Grupo 2)
 * Incluye archivo Excel consolidado adjunto con Hoja 1: Resumen y Hoja 2: Detalle.
 */

const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

const TARGET_TEST_EMAIL = 'especialista.preventa@provexpress.com.co';

// 1. Cargar dependencias de m365-report-mailer
const m365Dir = 'C:/Proyectos/m365-report-mailer';
const { ClientSecretCredential } = require(path.join(m365Dir, 'node_modules/@azure/identity'));
const axios = require(path.join(m365Dir, 'node_modules/axios'));

// 2. Cargar variables de entorno de m365-report-mailer
function loadEnv(filePath) {
  const env = {};
  if (fs.existsSync(filePath)) {
    const lines = fs.readFileSync(filePath, 'utf8').split('\n');
    for (const line of lines) {
      const parts = line.split('=');
      if (parts.length >= 2) {
        const key = parts[0].trim();
        let val = parts.slice(1).join('=').trim();
        if ((val.startsWith('"') && val.endsWith('"')) || (val.startsWith("'") && val.endsWith("'"))) {
          val = val.slice(1, -1);
        }
        env[key] = val.trim();
      }
    }
  }
  return env;
}

const env = loadEnv(path.join(m365Dir, '.env'));
const tenantId = env.TENANT_ID;
const clientId = env.CLIENT_ID;
const clientSecret = env.CLIENT_SECRET;
const senderEmail = env.SENDER_EMAIL || env.MAIL_SENDER_USER || 'juannovoa@provexpress.com.co';

if (!tenantId || !clientId || !clientSecret) {
  console.error('❌ Error: Faltan credenciales de Microsoft Entra ID en .env');
  process.exit(1);
}

function formatCOP(amount) {
  if (amount == null || isNaN(amount)) return '$ 0';
  return new Intl.NumberFormat('es-CO', {
    style: 'currency',
    currency: 'COP',
    maximumFractionDigits: 0,
  }).format(amount).replace('COP', '').replace(/\u00A0/g, ' ').trim();
}

function formatNumber(num) {
  if (num == null || isNaN(num)) return '0';
  return new Intl.NumberFormat('es-CO').format(Math.round(num));
}

function normalizeName(val) {
  return String(val || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

const GRUPO_2_EJECUTIVOS = [
  { nombre: 'Dayana Chala', email: 'dayana.chala@provexpress.com.co' },
  { nombre: 'Ángela Torres', email: 'angela.torres@provexpress.com.co' },
  { nombre: 'Alejandra Velásquez', email: 'alejandra.velasquez@provexpress.com.co' },
  { nombre: 'Daniel Galindo', email: 'daniel.galindo@provexpress.com.co' },
  { nombre: 'César Céspedes', email: 'cesar.cespedes@provexpress.com.co' },
  { nombre: 'Yurany Andrea Vargas', email: 'andrea.vargas@provexpress.com.co' },
  { nombre: 'Johanna Jaime', email: 'johanna.jaime@provexpress.com.co' },
  { nombre: 'Jasbleidy Mójica', email: 'johana.mojica@provexpress.com.co' },
  { nombre: 'Adriana Cucaita', email: 'adriana.cucaita@provexpress.com.co' },
  { nombre: 'Yovanny Herrera', email: 'yovanny.herrera@provexpress.com.co' },
  { nombre: 'Fernando Quiñonez', email: 'fernando.quinonez@provexpress.com.co' },
];

async function main() {
  console.log('================================================================');
  console.log('🚀 PREPARANDO CORREO DE PRUEBA PARA DIRECTOR DE GRUPO');
  console.log('📩 Destinatario:', TARGET_TEST_EMAIL);
  console.log('📤 Remitente institucional:', senderEmail);
  console.log('👔 Director simulado: Angélica Caballero (Directora Grupo 2)');
  console.log('================================================================');

  // 1. Cargar remisiones del archivo real
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(path.join(__dirname, '../Remisiones.xlsx'));
  const sis = wb.getWorksheet('Base-SIS');

  const dataSourceCutoff = '2026-09-16';
  const TODAY_CUTOFF = '2026-09-21';
  const TODAY_FORMATTED = '21 de septiembre de 2026';
  const cutoffRecords = [];

  for (let r = 2; r <= sis.rowCount; r++) {
    const row = sis.getRow(r);
    const dVal = row.getCell(1).value;
    const iso = dVal instanceof Date ? dVal.toISOString().slice(0, 10) : String(dVal || '').trim();
    if (iso === dataSourceCutoff) {
      cutoffRecords.push({
        emp: String(row.getCell(2).value || '').trim(),
        nit: String(row.getCell(3).value || '').trim(),
        company: String(row.getCell(4).value || '').trim(),
        merchandise: Number(row.getCell(5).value || 0),
        tax: Number(row.getCell(6).value || 0),
        total: Number(row.getCell(7).value || 0),
        emission: String(row.getCell(8).value || '').trim(),
        age: Number(row.getCell(9).value || 0),
        doc: String(row.getCell(10).value || '').trim(),
        order: String(row.getCell(11).value || '').trim(),
      });
    }
  }

  // 2. Mapear ejecutivos del Grupo 2
  const execSummaries = [];
  const allGroupRemisiones = [];

  for (const exec of GRUPO_2_EJECUTIVOS) {
    const targetNorm = normalizeName(exec.nombre);
    const targetTokens = targetNorm.split(' ').filter(Boolean);

    const execRecords = cutoffRecords.filter((r) => {
      const empNorm = normalizeName(r.emp);
      if (!empNorm) return false;
      if (empNorm.includes(targetNorm) || targetNorm.includes(empNorm)) return true;
      const empTokens = empNorm.split(' ').filter(Boolean);
      const common = targetTokens.filter((t) => empTokens.includes(t));
      return common.length >= 2;
    });

    const count = execRecords.length;
    const total = execRecords.reduce((s, r) => s + r.total, 0);
    const avgAge = count > 0 ? Math.round(execRecords.reduce((s, r) => s + r.age, 0) / count) : 0;
    const maxAge = count > 0 ? Math.max(...execRecords.map((r) => r.age)) : 0;

    const sortedByAge = [...execRecords].sort((a, b) => b.age - a.age);
    const criticalDoc = sortedByAge[0] || null;

    execSummaries.push({
      name: exec.nombre,
      email: exec.email,
      count,
      total,
      avgAge,
      maxAge,
      criticalDoc,
      records: execRecords,
    });

    allGroupRemisiones.push(...execRecords);
  }

  // Ordenar ejecutivos: primero con remisiones (valor desc), luego los en 0
  execSummaries.sort((a, b) => {
    if (a.count === 0 && b.count > 0) return 1;
    if (b.count === 0 && a.count > 0) return -1;
    return b.total - a.total;
  });

  const totalCount = allGroupRemisiones.length;
  const totalValue = allGroupRemisiones.reduce((s, r) => s + r.total, 0);
  const avgAge = totalCount > 0 ? Math.round(allGroupRemisiones.reduce((s, r) => s + r.age, 0) / totalCount) : 0;
  const activeCount = execSummaries.filter((e) => e.count > 0).length;
  const totalExecs = execSummaries.length;

  // Top Oportunidades del Grupo
  const sortedByValue = [...allGroupRemisiones].sort((a, b) => b.total - a.total);
  const topMayorValor = sortedByValue[0] || null;

  const sortedByAgeAll = [...allGroupRemisiones].sort((a, b) => b.age - a.age);
  const topMayorAntiguedad = sortedByAgeAll[0] || null;

  console.log(`✓ Remisiones consolidadas Grupo 2 (Angélica Caballero): ${totalCount}`);
  console.log(`   - Ejecutivos con pendientes: ${activeCount} de ${totalExecs}`);
  console.log(`   - Total Valor Grupo: ${formatCOP(totalValue)}`);
  console.log(`   - Antigüedad Promedio Grupo: ${avgAge} días`);
  console.log(`   - Top Mayor Valor: ${topMayorValor?.doc} - ${topMayorValor?.company} (${formatCOP(topMayorValor?.total)}) [${topMayorValor?.emp}]`);
  console.log(`   - Top Mayor Antigüedad: ${topMayorAntiguedad?.doc} - ${topMayorAntiguedad?.company} (${topMayorAntiguedad?.age} días) [${topMayorAntiguedad?.emp}]`);

  // 3. Generar Excel consolidado
  const excelWb = new ExcelJS.Workbook();
  excelWb.creator = 'Provexpress SAS - Sistema de Remisiones';

  // HOJA 1: Resumen
  const wsSum = excelWb.addWorksheet('Resumen por Ejecutivo', {
    views: [{ state: 'frozen', ySplit: 5 }],
  });

  wsSum.mergeCells('A1:I1');
  const t1 = wsSum.getCell('A1');
  t1.value = 'PROVEXPRESS SAS · CONSOLIDADO DE REMISIONES · DIRECCIÓN GRUPO 2';
  t1.font = { name: 'Calibri', size: 13, bold: true, color: { argb: 'FFFFFFFF' } };
  t1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
  t1.alignment = { horizontal: 'center', vertical: 'middle' };
  wsSum.getRow(1).height = 28;

  wsSum.getCell('A2').value = 'Directora de Grupo:';
  wsSum.getCell('A2').font = { bold: true, size: 10, color: { argb: 'FF475569' } };
  wsSum.getCell('B2').value = 'Angélica Caballero';
  wsSum.getCell('B2').font = { bold: true, size: 11, color: { argb: 'FF0F172A' } };

  wsSum.getCell('D2').value = 'Fecha de Corte:';
  wsSum.getCell('D2').font = { bold: true, size: 10, color: { argb: 'FF475569' } };
  wsSum.getCell('E2').value = TODAY_FORMATTED;
  wsSum.getCell('E2').font = { bold: true, size: 11, color: { argb: 'FF0F172A' } };

  wsSum.getCell('G2').value = 'Total Remisiones:';
  wsSum.getCell('G2').font = { bold: true, size: 10, color: { argb: 'FF475569' } };
  wsSum.getCell('H2').value = totalCount;
  wsSum.getCell('H2').font = { bold: true, size: 11, color: { argb: 'FF1E3A8A' } };

  wsSum.getCell('A3').value = 'Total Ejecutivos:';
  wsSum.getCell('A3').font = { bold: true, size: 10, color: { argb: 'FF475569' } };
  wsSum.getCell('B3').value = `${activeCount} con pendientes de ${totalExecs} activos`;
  wsSum.getCell('B3').font = { bold: true, size: 10, color: { argb: 'FF0F172A' } };

  wsSum.getCell('D3').value = 'Antigüedad Promedio:';
  wsSum.getCell('D3').font = { bold: true, size: 10, color: { argb: 'FF475569' } };
  wsSum.getCell('E3').value = `${avgAge} días`;
  wsSum.getCell('E3').font = { bold: true, size: 10, color: { argb: 'FFDC2626' } };

  wsSum.getCell('G3').value = 'Valor Total Grupo:';
  wsSum.getCell('G3').font = { bold: true, size: 10, color: { argb: 'FF475569' } };
  wsSum.getCell('H3').value = totalValue;
  wsSum.getCell('H3').numFmt = '"$"#,##0';
  wsSum.getCell('H3').font = { bold: true, size: 11, color: { argb: 'FF15803D' } };

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

  const headerRow1 = wsSum.getRow(5);
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
  });

  let sumRowIdx = 6;
  execSummaries.forEach((exec, idx) => {
    const row = wsSum.getRow(sumRowIdx);
    row.height = 20;
    const isEven = idx % 2 === 0;
    const rowBg = isEven ? 'FFFFFFFF' : 'FFF8FAFC';
    const pct = totalValue > 0 ? exec.total / totalValue : 0;

    let status = 'En gestión';
    if (exec.count === 0) status = 'Al día ✓';
    else if (exec.maxAge > 30) status = 'Urgente (>30d)';

    const critStr = exec.criticalDoc ? `${exec.criticalDoc.doc} (${exec.criticalDoc.age}d)` : '—';

    row.getCell(1).value = idx + 1;
    row.getCell(2).value = exec.name;
    row.getCell(3).value = exec.email;
    row.getCell(4).value = exec.count;
    row.getCell(5).value = exec.total;
    row.getCell(5).numFmt = '"$"#,##0';
    row.getCell(6).value = pct;
    row.getCell(6).numFmt = '0.0%';
    row.getCell(7).value = exec.count > 0 ? `${exec.avgAge} días` : '—';
    row.getCell(8).value = critStr;
    row.getCell(9).value = status;

    for (let c = 1; c <= 9; c++) {
      const cell = row.getCell(c);
      cell.font = { name: 'Calibri', size: 10 };
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
    sumRowIdx++;
  });

  const totRow = wsSum.getRow(sumRowIdx);
  totRow.height = 24;
  totRow.getCell(2).value = 'TOTAL GRUPO';
  totRow.getCell(2).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  totRow.getCell(4).value = totalCount;
  totRow.getCell(4).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  totRow.getCell(4).alignment = { horizontal: 'center', vertical: 'middle' };
  totRow.getCell(5).value = totalValue;
  totRow.getCell(5).numFmt = '"$"#,##0';
  totRow.getCell(5).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  totRow.getCell(5).alignment = { horizontal: 'right', vertical: 'middle' };
  totRow.getCell(6).value = 1.0;
  totRow.getCell(6).numFmt = '0.0%';
  totRow.getCell(6).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  totRow.getCell(6).alignment = { horizontal: 'center', vertical: 'middle' };
  totRow.getCell(7).value = `${avgAge} días`;
  totRow.getCell(7).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  totRow.getCell(7).alignment = { horizontal: 'center', vertical: 'middle' };

  for (let c = 1; c <= 9; c++) {
    const cell = totRow.getCell(c);
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
    cell.border = { top: { style: 'medium', color: { argb: 'FF0F172A' } }, bottom: { style: 'medium', color: { argb: 'FF0F172A' } } };
  }

  wsSum.columns = [
    { width: 6 }, { width: 26 }, { width: 32 }, { width: 20 },
    { width: 22 }, { width: 16 }, { width: 20 }, { width: 24 }, { width: 18 }
  ];

  // HOJA 2: Detalle
  const wsDet = excelWb.addWorksheet('Detalle de Remisiones', {
    views: [{ state: 'frozen', ySplit: 3 }],
  });

  wsDet.mergeCells('A1:L1');
  const t2 = wsDet.getCell('A1');
  t2.value = 'PROVEXPRESS SAS · DETALLE DE REMISIONES ABIERTAS · GRUPO 2';
  t2.font = { name: 'Calibri', size: 13, bold: true, color: { argb: 'FFFFFFFF' } };
  t2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
  t2.alignment = { horizontal: 'center', vertical: 'middle' };
  wsDet.getRow(1).height = 28;

  const detHeaders = [
    'N°', 'Asesor Comercial', 'Remisión', 'NIT Cliente',
    'Nombre / Razón Social del Cliente', 'Vr. Mercancía', 'Vr. IVA',
    'Vr. Total ($ COP)', 'Fecha Emisión', 'Días Abierta', 'Rango Antigüedad', 'N° Pedido'
  ];

  const headerRow2 = wsDet.getRow(3);
  headerRow2.height = 24;
  detHeaders.forEach((text, i) => {
    const cell = headerRow2.getCell(i + 1);
    cell.value = text;
    cell.font = { name: 'Calibri', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
    cell.alignment = {
      horizontal: i === 0 || i === 2 || i === 8 || i === 9 || i === 10 || i === 11 ? 'center' : i >= 5 && i <= 7 ? 'right' : 'left',
      vertical: 'middle',
    };
  });

  const sortedDetail = [...allGroupRemisiones].sort((a, b) => (b.age !== a.age ? b.age - a.age : b.total - a.total));

  let detRowIdx = 4;
  sortedDetail.forEach((r, idx) => {
    const row = wsDet.getRow(detRowIdx);
    row.height = 20;
    const isEven = idx % 2 === 0;
    const rowBg = isEven ? 'FFFFFFFF' : 'FFF8FAFC';

    let ageRange = '0 a 7 días';
    if (r.age > 30) ageRange = 'Más de 30 días';
    else if (r.age > 15) ageRange = '16 a 30 días';
    else if (r.age > 7) ageRange = '8 a 15 días';

    row.getCell(1).value = idx + 1;
    row.getCell(2).value = r.emp || 'Sin Asignar';
    row.getCell(3).value = r.doc?.startsWith('REM-') ? r.doc : `REM-${r.doc || 'S/N'}`;
    row.getCell(4).value = r.nit || '—';
    row.getCell(5).value = r.company || '—';
    row.getCell(6).value = r.merchandise || 0;
    row.getCell(6).numFmt = '"$"#,##0';
    row.getCell(7).value = r.tax || 0;
    row.getCell(7).numFmt = '"$"#,##0';
    row.getCell(8).value = r.total || 0;
    row.getCell(8).numFmt = '"$"#,##0';
    row.getCell(9).value = r.emission || '—';
    row.getCell(10).value = r.age;
    row.getCell(11).value = ageRange;
    row.getCell(12).value = r.order || '—';

    for (let c = 1; c <= 12; c++) {
      const cell = row.getCell(c);
      cell.font = { name: 'Calibri', size: 10 };
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
    detRowIdx++;
  });

  wsDet.columns = [
    { width: 6 }, { width: 26 }, { width: 16 }, { width: 16 },
    { width: 34 }, { width: 16 }, { width: 14 }, { width: 18 },
    { width: 14 }, { width: 14 }, { width: 18 }, { width: 16 }
  ];

  const excelBuffer = await excelWb.xlsx.writeBuffer();
  const excelBase64 = excelBuffer.toString('base64');
  const excelFilename = `Remisiones_Consolidado_Grupo_2_Angelica_Caballero_${TODAY_CUTOFF}.xlsx`;
  console.log(`✓ Archivo Excel consolidado generado: ${excelFilename} (${excelBuffer.length} bytes)`);

  // 4. Generar HTML del correo
  const rowsHtml = execSummaries
    .map((exec, index) => {
      const isEven = index % 2 === 0;
      const bgColor = isEven ? '#FFFFFF' : '#F8FAFC';

      let ageBadgeColor = '#15803D';
      let ageBadgeBg = '#DCFCE7';
      if (exec.avgAge > 30) {
        ageBadgeColor = '#B91C1C';
        ageBadgeBg = '#FEE2E2';
      } else if (exec.avgAge > 15) {
        ageBadgeColor = '#B45309';
        ageBadgeBg = '#FEF3C7';
      }

      let statusBadge = '';
      if (exec.count === 0) {
        statusBadge = '<span style="display: inline-block; padding: 2px 8px; border-radius: 999px; font-size: 10px; font-weight: 700; color: #15803D; background-color: #DCFCE7;">Al día ✓</span>';
      } else if (exec.maxAge > 30) {
        statusBadge = '<span style="display: inline-block; padding: 2px 8px; border-radius: 999px; font-size: 10px; font-weight: 700; color: #B91C1C; background-color: #FEE2E2;">Urgente (>30d)</span>';
      } else {
        statusBadge = '<span style="display: inline-block; padding: 2px 8px; border-radius: 999px; font-size: 10px; font-weight: 700; color: #1E3A8A; background-color: #DBEAFE;">En gestión</span>';
      }

      const critInfo = exec.criticalDoc
        ? `${exec.criticalDoc.doc?.startsWith('REM-') ? exec.criticalDoc.doc : `REM-${exec.criticalDoc.doc}`} (${exec.criticalDoc.age}d)`
        : '—';

      return `
      <tr style="background-color: ${bgColor}; border-bottom: 1px solid #E2E8F0;">
        <td style="padding: 10px 12px; font-size: 12px; font-weight: 700; color: #1E293B; font-family: 'Segoe UI', Arial, sans-serif;">
          ${exec.name}
          <div style="font-size: 10px; font-weight: 400; color: #64748B;">${exec.email}</div>
        </td>
        <td style="padding: 10px 12px; text-align: center; font-size: 12px; font-weight: 700; color: #0F172A; font-family: 'Segoe UI', Arial, sans-serif;">
          ${exec.count > 0 ? formatNumber(exec.count) : '<span style="color: #94A3B8;">0</span>'}
        </td>
        <td style="padding: 10px 12px; text-align: right; font-size: 12.5px; font-weight: 800; color: ${exec.count > 0 ? '#15803D' : '#94A3B8'}; font-family: 'Segoe UI', Arial, sans-serif;">
          ${formatCOP(exec.total)}
        </td>
        <td style="padding: 10px 12px; text-align: center;">
          ${exec.count > 0 ? `
          <span style="display: inline-block; padding: 2px 8px; border-radius: 6px; font-size: 11px; font-weight: 700; color: ${ageBadgeColor}; background-color: ${ageBadgeBg}; font-family: 'Segoe UI', Arial, sans-serif;">
            ${exec.avgAge} días
          </span>
          ` : '<span style="font-size: 11px; color: #94A3B8;">—</span>'}
        </td>
        <td style="padding: 10px 12px; font-size: 11px; color: #475569; font-family: 'Segoe UI', Arial, sans-serif;">
          ${critInfo}
        </td>
        <td style="padding: 10px 12px; text-align: center;">
          ${statusBadge}
        </td>
      </tr>
      `;
    })
    .join('');

  const htmlContent = `
<!DOCTYPE html>
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <title>Resumen de Remisiones · Dirección Grupo 2 · Provexpress</title>
  <style type="text/css">
    body, table, td, a { -webkit-text-size-adjust: 100%; -ms-text-size-adjust: 100%; }
    table, td { mso-table-lspace: 0pt; mso-table-rspace: 0pt; }
    img { -ms-interpolation-mode: bicubic; border: 0; outline: none; text-decoration: none; }
    body { margin: 0; padding: 0; width: 100% !important; background-color: #F1F4F8; font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, Helvetica, Arial, sans-serif; }
    @media only screen and (max-width: 920px) {
      .main-card { width: 100% !important; border-radius: 0 !important; }
      .stack-kpi { display: inline-block !important; width: 48% !important; margin-bottom: 8px !important; }
    }
  </style>
</head>
<body style="margin: 0; padding: 24px 0; background-color: #F1F4F8;">

  <!-- CONTENEDOR PRINCIPAL BLANCO (960px) IDÉNTICO A COMERCIALES -->
  <table role="presentation" class="main-card" width="960" align="center" border="0" cellpadding="0" cellspacing="0" style="width: 960px; margin: 0 auto; background-color: #FFFFFF; border-radius: 16px; overflow: hidden; border: 1.5px solid #E2E8F0; box-shadow: 0 4px 20px rgba(15, 23, 42, 0.05);">
    <tr>
      <td style="padding: 24px 28px;">

        <!-- ═════════════════════════════════════════════════════════════════════
             1. ENCABEZADO: LOGO + SALUDO + MUÑECO (AVATAR) + CARD DIRECCIÓN
             ═════════════════════════════════════════════════════════════════════ -->
        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="margin-bottom: 20px;">
          <tr>
            <!-- LOGO PROVEXPRESS OFICIAL -->
            <td width="160" valign="middle" align="left" style="padding-right: 18px;">
              <img src="cid:logo_provexpress" width="144" height="110" style="display: block; border: 0; width: 144px; height: 110px; max-width: 144px;" alt="Provexpress" />
            </td>

            <!-- TEXTO DE SALUDO Y TITULAR DE LIDERAZGO -->
            <td valign="middle" style="padding-right: 12px;">
              <div style="font-size: 15px; color: #475569; font-weight: 600; font-family: 'Segoe UI', Arial, sans-serif;">
                Hola, <strong style="color: #0F172A; font-size: 16px;">Angélica Caballero</strong> 👋
              </div>
              <h1 style="margin: 3px 0 6px 0; font-size: 24px; font-weight: 900; color: #0F172A; line-height: 1.15; letter-spacing: -0.02em; font-family: 'Segoe UI', Arial, sans-serif;">
                Oportunidades de tu equipo listas para<br>
                convertirse en <span style="color: #16A34A;">ventas facturadas</span>
              </h1>
              <p style="margin: 0; font-size: 12px; color: #475569; line-height: 1.35; font-family: 'Segoe UI', Arial, sans-serif; max-width: 390px;">
                Excelente labor de liderazgo con tu equipo comercial. Te presentamos el consolidado de remisiones abiertas de tus ejecutivos para acompañar, priorizar y acelerar el cierre comercial del grupo.
              </p>
              <p style="margin: 5px 0 0 0; font-size: 12px; color: #0F172A; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">
                El seguimiento oportuno transforma el esfuerzo de tu equipo en <span style="color: #16A34A;">resultados reales</span>.
              </p>
            </td>

            <!-- ILUSTRACIÓN DIRECTORA COMERCIAL: LA MUÑECA (THUMBS UP) -->
            <td width="190" valign="bottom" align="center" style="padding-right: 10px;">
              <img src="cid:avatar_woman" width="180" height="154" style="display: block; border: 0; width: 180px; height: auto; margin: 0 auto;" alt="Directora Comercial" />
            </td>

            <!-- CARD DIRECCIÓN DE GRUPO -->
            <td width="130" valign="middle" align="right">
              <table role="presentation" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 14px; padding: 14px 10px; width: 126px; text-align: center;">
                <tr>
                  <td align="center">
                    <div style="background-color: #1E3A8A; color: #FFFFFF; width: 28px; height: 28px; border-radius: 50%; font-size: 14px; font-weight: 900; line-height: 28px; margin: 0 auto 8px auto;">
                      👔
                    </div>
                    <div style="font-size: 11.5px; font-weight: 700; color: #1E293B; line-height: 1.3; font-family: 'Segoe UI', Arial, sans-serif;">
                      Tu liderazgo<br>
                      hace la<br>
                      <strong style="color: #1E3A8A; font-size: 12.5px;">diferencia.</strong>
                    </div>
                    <div style="margin-top: 6px; padding-top: 6px; border-top: 1px dashed #E2E8F0; font-size: 9.5px; font-weight: 800; color: #2563EB; text-transform: uppercase;">
                      Grupo 2
                    </div>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>

        <!-- ═════════════════════════════════════════════════════════════════════
             2. FILA DE 4 TARJETAS KPI CON EL ESTILO EXACTO DE COMERCIALES
             ═════════════════════════════════════════════════════════════════════ -->
        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="margin-bottom: 20px;">
          <tr>
            <!-- KPI 1: Ejecutivos con pendientes -->
            <td width="24%" valign="top" class="stack-kpi" style="padding-right: 8px;">
              <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 12px; padding: 12px 10px; height: 95px;">
                <tr>
                  <td width="32" valign="top">
                    <div style="background-color: #312E81; width: 30px; height: 30px; border-radius: 50%; text-align: center; line-height: 30px; font-size: 15px; color: #FFFFFF;">
                      👥
                    </div>
                  </td>
                  <td valign="top" style="padding-left: 6px;">
                    <div style="font-size: 10.5px; color: #475569; font-weight: 600; line-height: 1.2; font-family: 'Segoe UI', Arial, sans-serif;">
                      Ejecutivos con<br>remisiones activas
                    </div>
                  </td>
                </tr>
                <tr>
                  <td colspan="2" valign="bottom" style="padding-top: 6px;">
                    <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td valign="bottom" style="font-size: 22px; font-weight: 900; color: #312E81; font-family: 'Segoe UI', Arial, sans-serif; line-height: 1;">
                          ${activeCount} / ${totalExecs}
                        </td>
                        <td align="right" valign="bottom" style="font-size: 16px; color: #94A3B8; opacity: 0.65;">
                          📋
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>

            <!-- KPI 2: Total Remisiones del Grupo -->
            <td width="24%" valign="top" class="stack-kpi" style="padding-right: 8px;">
              <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 12px; padding: 12px 10px; height: 95px;">
                <tr>
                  <td width="32" valign="top">
                    <div style="background-color: #1E3A8A; width: 30px; height: 30px; border-radius: 50%; text-align: center; line-height: 30px; font-size: 15px; color: #FFFFFF;">
                      📦
                    </div>
                  </td>
                  <td valign="top" style="padding-left: 6px;">
                    <div style="font-size: 10.5px; color: #475569; font-weight: 600; line-height: 1.2; font-family: 'Segoe UI', Arial, sans-serif;">
                      Remisiones listas<br>en el equipo
                    </div>
                  </td>
                </tr>
                <tr>
                  <td colspan="2" valign="bottom" style="padding-top: 6px;">
                    <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td valign="bottom" style="font-size: 22px; font-weight: 900; color: #1E3A8A; font-family: 'Segoe UI', Arial, sans-serif; line-height: 1;">
                          ${formatNumber(totalCount)}
                        </td>
                        <td align="right" valign="bottom" style="font-size: 16px; color: #94A3B8; opacity: 0.65;">
                          📝
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>

            <!-- KPI 3: Valor por facturar del grupo -->
            <td width="28%" valign="top" class="stack-kpi" style="padding-right: 8px;">
              <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 12px; padding: 12px 10px; height: 95px;">
                <tr>
                  <td width="32" valign="top">
                    <div style="background-color: #15803D; width: 30px; height: 30px; border-radius: 50%; text-align: center; line-height: 30px; font-size: 15px; font-weight: 900; color: #FFFFFF;">
                      $
                    </div>
                  </td>
                  <td valign="top" style="padding-left: 6px;">
                    <div style="font-size: 10.5px; color: #475569; font-weight: 600; line-height: 1.2; font-family: 'Segoe UI', Arial, sans-serif;">
                      Valor por facturar<br>del equipo
                    </div>
                  </td>
                </tr>
                <tr>
                  <td colspan="2" valign="bottom" style="padding-top: 6px;">
                    <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td valign="bottom" style="font-size: 17px; font-weight: 900; color: #15803D; font-family: 'Segoe UI', Arial, sans-serif; line-height: 1; white-space: nowrap;">
                          ${formatCOP(totalValue)}
                        </td>
                        <td align="right" valign="bottom" style="font-size: 16px; color: #86EFAC; opacity: 0.75;">
                          🪙
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>

            <!-- KPI 4: Antigüedad promedio del grupo -->
            <td width="24%" valign="top" class="stack-kpi">
              <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 12px; padding: 12px 10px; height: 95px;">
                <tr>
                  <td width="32" valign="top">
                    <div style="background-color: #D97706; width: 30px; height: 30px; border-radius: 50%; text-align: center; line-height: 30px; font-size: 15px; color: #FFFFFF;">
                      🕒
                    </div>
                  </td>
                  <td valign="top" style="padding-left: 6px;">
                    <div style="font-size: 10.5px; color: #475569; font-weight: 600; line-height: 1.2; font-family: 'Segoe UI', Arial, sans-serif;">
                      Antigüedad<br>promedio equipo
                    </div>
                  </td>
                </tr>
                <tr>
                  <td colspan="2" valign="bottom" style="padding-top: 6px;">
                    <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td valign="bottom" style="font-size: 22px; font-weight: 900; color: #D97706; font-family: 'Segoe UI', Arial, sans-serif; line-height: 1;">
                          ${avgAge} <span style="font-size: 12px; font-weight: 700;">días</span>
                        </td>
                        <td align="right" valign="bottom" style="font-size: 16px; color: #FDBA74; opacity: 0.75;">
                          📅
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>

        <!-- ═════════════════════════════════════════════════════════════════════
             3. TABLA DE DESEMPEÑO POR ASESOR COMERCIAL
             ═════════════════════════════════════════════════════════════════════ -->
        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="margin-bottom: 8px;">
          <tr>
            <td valign="middle">
              <span style="font-size: 14.5px; font-weight: 850; color: #1E293B; font-family: 'Segoe UI', Arial, sans-serif;">
                📋 Detalle y Desempeño por Asesor Comercial
              </span>
              <span style="font-size: 13px; color: #D97706; margin: 0 4px;">☆</span>
              <span style="font-size: 11.5px; color: #64748B; font-family: 'Segoe UI', Arial, sans-serif;">
                Enfócate en acompañar los casos de mayor valor y mayor antigüedad.
              </span>
            </td>
            <td align="right" valign="middle">
              <span style="font-size: 11px; font-weight: 700; color: #64748B; font-family: 'Segoe UI', Arial, sans-serif;">
                Corte: 16 de septiembre de 2026
              </span>
            </td>
          </tr>
        </table>

        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="border: 1px solid #E2E8F0; border-radius: 12px; overflow: hidden; border-collapse: separate; border-spacing: 0; margin-bottom: 18px;">
          <thead>
            <tr style="background-color: #1E293B; color: #FFFFFF;">
              <th style="padding: 11px 14px; text-align: left; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">Ejecutivo Comercial</th>
              <th style="padding: 11px 12px; text-align: center; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">Remisiones</th>
              <th style="padding: 11px 14px; text-align: right; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">Valor Total</th>
              <th style="padding: 11px 12px; text-align: center; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">Días Prom.</th>
              <th style="padding: 11px 14px; text-align: left; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">Remisión Crítica</th>
              <th style="padding: 11px 12px; text-align: center; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">Estado</th>
            </tr>
          </thead>
          <tbody>
            ${rowsHtml}
          </tbody>
        </table>

        <!-- ═════════════════════════════════════════════════════════════════════
             4. HERO BANNER: TOP OPORTUNIDADES DEL EQUIPO (FULL WIDTH)
             ═════════════════════════════════════════════════════════════════════ -->
        ${(topMayorValor || topMayorAntiguedad) ? `
        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 12px; padding: 14px 18px; margin-bottom: 14px;">
          <!-- FILA 1: MAYOR VALOR DEL EQUIPO -->
          ${topMayorValor ? `
          <tr>
            <td width="42" valign="middle" align="center" style="padding-right: 12px;">
              <span style="font-size: 26px;">🏆</span>
            </td>
            <td valign="middle" style="padding-right: 14px;">
              <div style="font-size: 10px; font-weight: 700; color: #15803D; text-transform: uppercase; letter-spacing: 0.04em; font-family: 'Segoe UI', Arial, sans-serif;">
                Mayor valor por facturar del equipo <span style="color: #D97706;">☆</span>
              </div>
              <div style="font-size: 15px; font-weight: 900; color: #1E3A8A; font-family: 'Segoe UI', Arial, sans-serif; margin: 1px 0;">
                ${topMayorValor.doc?.startsWith('REM-') ? topMayorValor.doc : `REM-${topMayorValor.doc}`}
              </div>
              <div style="font-size: 11px; color: #475569; font-family: 'Segoe UI', Arial, sans-serif; max-width: 250px;">
                ${topMayorValor.company} · <em>${topMayorValor.emp}</em>
              </div>
            </td>
            <td width="130" valign="middle" style="padding-right: 14px;">
              <div style="font-size: 15px; font-weight: 900; color: #15803D; font-family: 'Segoe UI', Arial, sans-serif;">
                ${formatCOP(topMayorValor.total)}
              </div>
              <div style="font-size: 10px; color: #64748B; font-family: 'Segoe UI', Arial, sans-serif;">
                Mayor valor
              </div>
            </td>
            <td width="95" valign="middle" style="padding-right: 14px;">
              <div style="font-size: 14.5px; font-weight: 900; color: #0F172A; font-family: 'Segoe UI', Arial, sans-serif;">
                ${topMayorValor.age} días
              </div>
              <div style="font-size: 10px; color: #64748B; font-family: 'Segoe UI', Arial, sans-serif;">
                Antigüedad
              </div>
            </td>
            <td valign="middle" align="right">
              <table role="presentation" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td valign="middle" style="padding-right: 4px; font-size: 16px; color: #15803D;">⤷</td>
                  <td valign="middle" style="font-size: 11.5px; font-weight: 850; font-style: italic; color: #15803D; font-family: 'Segoe UI', Arial, sans-serif; line-height: 1.2;">
                    ¡Mayor impacto<br>en la meta!
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          ` : ''}

          ${(topMayorAntiguedad && (!topMayorValor || topMayorAntiguedad.doc !== topMayorValor.doc)) ? `
          <!-- SEPARADOR -->
          <tr>
            <td colspan="5" style="padding: 10px 0;">
              <div style="border-top: 1px dashed #E2E8F0; height: 1px; line-height: 1px; font-size: 0;">&nbsp;</div>
            </td>
          </tr>

          <!-- FILA 2: MAYOR ANTIGÜEDAD DEL EQUIPO -->
          <tr>
            <td width="42" valign="middle" align="center" style="padding-right: 12px;">
              <span style="font-size: 26px;">⏳</span>
            </td>
            <td valign="middle" style="padding-right: 14px;">
              <div style="font-size: 10px; font-weight: 700; color: #DC2626; text-transform: uppercase; letter-spacing: 0.04em; font-family: 'Segoe UI', Arial, sans-serif;">
                Mayor antigüedad pendiente del equipo <span style="color: #DC2626;">⏱</span>
              </div>
              <div style="font-size: 15px; font-weight: 900; color: #1E3A8A; font-family: 'Segoe UI', Arial, sans-serif; margin: 1px 0;">
                ${topMayorAntiguedad.doc?.startsWith('REM-') ? topMayorAntiguedad.doc : `REM-${topMayorAntiguedad.doc}`}
              </div>
              <div style="font-size: 11px; color: #475569; font-family: 'Segoe UI', Arial, sans-serif; max-width: 250px;">
                ${topMayorAntiguedad.company} · <em>${topMayorAntiguedad.emp}</em>
              </div>
            </td>
            <td width="130" valign="middle" style="padding-right: 14px;">
              <div style="font-size: 15px; font-weight: 900; color: #1E3A8A; font-family: 'Segoe UI', Arial, sans-serif;">
                ${formatCOP(topMayorAntiguedad.total)}
              </div>
              <div style="font-size: 10px; color: #64748B; font-family: 'Segoe UI', Arial, sans-serif;">
                Valor
              </div>
            </td>
            <td width="95" valign="middle" style="padding-right: 14px;">
              <div style="font-size: 14.5px; font-weight: 900; color: #DC2626; font-family: 'Segoe UI', Arial, sans-serif;">
                ${topMayorAntiguedad.age} días
              </div>
              <div style="font-size: 10px; color: #64748B; font-family: 'Segoe UI', Arial, sans-serif;">
                Mayor antigüedad
              </div>
            </td>
            <td valign="middle" align="right">
              <table role="presentation" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td valign="middle" style="padding-right: 4px; font-size: 16px; color: #DC2626;">⤷</td>
                  <td valign="middle" style="font-size: 11.5px; font-weight: 850; font-style: italic; color: #DC2626; font-family: 'Segoe UI', Arial, sans-serif; line-height: 1.2;">
                    ¡Prioridad de gestión<br>inmediata!
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          ` : ''}
        </table>
        ` : ''}

        <!-- ═════════════════════════════════════════════════════════════════════
             5. BANNER DE ARCHIVO ADJUNTO EXCEL CONSOLIDADO
             ═════════════════════════════════════════════════════════════════════ -->
        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #F8FAFC; border: 1.5px solid #E2E8F0; border-radius: 12px; padding: 12px 18px; margin-bottom: 14px;">
          <tr>
            <td width="38" valign="middle" align="center" style="padding-right: 12px;">
              <div style="background-color: #107C41; color: #FFFFFF; width: 32px; height: 32px; border-radius: 8px; text-align: center; line-height: 32px; font-size: 16px; font-weight: 900;">
                📊
              </div>
            </td>
            <td valign="middle">
              <div style="font-size: 12px; font-weight: 800; color: #0F172A; font-family: 'Segoe UI', Arial, sans-serif;">
                📎 Archivo adjunto: ${excelFilename}
              </div>
              <div style="font-size: 11px; color: #475569; font-family: 'Segoe UI', Arial, sans-serif; margin-top: 2px;">
                Incluye <strong>Hoja 1: Resumen por Ejecutivo</strong> (totales, promedio de días y % de participación) y <strong>Hoja 2: Detalle Completo</strong> con las <strong>${formatNumber(totalCount)}</strong> remisiones abiertas (${formatCOP(totalValue)}) de tu equipo comercial.
              </div>
            </td>
          </tr>
        </table>

        <!-- ═════════════════════════════════════════════════════════════════════
             6. PIE DE PÁGINA: CORAZÓN AZUL + FIRMA + SOBRE POSTAL (FULL WIDTH)
             ═════════════════════════════════════════════════════════════════════ -->
        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #F8FAFC; border: 1.5px solid #E2E8F0; border-radius: 14px; padding: 16px 22px;">
          <tr>
            <!-- Corazón azul grande -->
            <td width="38" valign="middle" align="center" style="padding-right: 12px;">
              <div style="font-size: 28px; line-height: 1; color: #1E3A8A;">
                💙
              </div>
            </td>

            <!-- Mensaje de agradecimiento y liderazgo -->
            <td valign="middle">
              <div style="font-size: 12px; color: #475569; font-family: 'Segoe UI', Arial, sans-serif;">
                Gracias por tu liderazgo y gestión continua.
              </div>
              <div style="font-size: 13.5px; font-weight: 850; color: #1E3A8A; font-family: 'Segoe UI', Arial, sans-serif; margin-top: 2px;">
                ¡Acompañando a nuestros equipos alcanzamos las metas!
              </div>
            </td>

            <!-- Sobre postal azul -->
            <td width="55" align="right" valign="middle">
              <svg width="48" height="38" viewBox="0 0 52 42" fill="none" xmlns="http://www.w3.org/2000/svg">
                <rect x="8" y="2" width="36" height="26" rx="3" fill="#FFFFFF" stroke="#60A5FA" stroke-width="1.5"/>
                <line x1="13" y1="8" x2="31" y2="8" stroke="#93C5FD" stroke-width="2" stroke-linecap="round"/>
                <line x1="13" y1="13" x2="27" y2="13" stroke="#93C5FD" stroke-width="2" stroke-linecap="round"/>
                <rect x="2" y="10" width="48" height="30" rx="4" fill="#1E3A8A"/>
                <path d="M2 12L26 28L50 12" stroke="#3B82F6" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
                <path d="M3 15L26 29L49 15" stroke="#FFFFFF" stroke-width="1.5"/>
                <circle cx="41" cy="31" r="7" fill="#16A34A" stroke="#FFFFFF" stroke-width="1.5"/>
                <path d="M38 31L40 33L44 29" stroke="#FFFFFF" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/>
              </svg>
            </td>
          </tr>
        </table>

      </td>
    </tr>
  </table>

</body>
</html>
  `.trim();

  // 5. Cargar Logo y Muñeca (Avatar Femenino) para adjuntos inline CID
  const logoBytes = fs.readFileSync(path.join(__dirname, '../public/logo_provexpress_stacked.png')).toString('base64');
  const avatarBytes = fs.readFileSync(path.join(__dirname, '../public/avatar_woman.png')).toString('base64');

  // 6. Autenticar con Microsoft Graph
  console.log('\n🔑 Autenticando con Microsoft Entra ID (Graph)...');
  const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
  const tokenResponse = await credential.getToken('https://graph.microsoft.com/.default');
  console.log('✓ Token de Microsoft Graph obtenido exitosamente.');

  const mailPayload = {
    message: {
      subject: `[PRUEBA] Reporte Consolidado de Remisiones · Dirección Grupo 2 · Angélica Caballero · Corte ${TODAY_FORMATTED}`,
      body: {
        contentType: 'HTML',
        content: htmlContent,
      },
      toRecipients: [
        {
          emailAddress: {
            address: TARGET_TEST_EMAIL,
            name: 'Especialista Preventa',
          },
        },
      ],
      attachments: [
        {
          '@odata.type': '#microsoft.graph.fileAttachment',
          name: 'logo_provexpress.png',
          contentType: 'image/png',
          contentBytes: logoBytes,
          isInline: true,
          contentId: 'logo_provexpress',
        },
        {
          '@odata.type': '#microsoft.graph.fileAttachment',
          name: 'avatar_woman.png',
          contentType: 'image/png',
          contentBytes: avatarBytes,
          isInline: true,
          contentId: 'avatar_woman',
        },
        {
          '@odata.type': '#microsoft.graph.fileAttachment',
          name: excelFilename,
          contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          contentBytes: excelBase64,
          isInline: false,
        },
      ],
    },
    saveToSentItems: true,
  };

  console.log(`\n📨 Enviando correo de prueba a ${TARGET_TEST_EMAIL}...`);
  const graphUrl = `https://graph.microsoft.com/v1.0/users/${senderEmail}/sendMail`;

  const response = await axios.post(graphUrl, mailPayload, {
    headers: {
      Authorization: `Bearer ${tokenResponse.token}`,
      'Content-Type': 'application/json',
    },
  });

  console.log('\n✅ ¡CORREO DE DIRECTOR ENVIADO CON ÉXITO!');
  console.log(`   - Código de respuesta HTTP: ${response.status} ${response.statusText}`);
  console.log(`   - Destino: ${TARGET_TEST_EMAIL}`);
  console.log(`   - Archivo adjunto: ${excelFilename}`);
  console.log('================================================================');
}

main().catch((err) => {
  console.error('\n❌ ERROR enviando correo:', err.response?.data || err.message);
  process.exit(1);
});
