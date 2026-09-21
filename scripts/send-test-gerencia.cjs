/**
 * Script standalone para despachar el correo de prueba para Dirección y Gerencia Comercial ("Los Jefes Jefes")
 * a especialista.preventa@provexpress.com.co
 * Utiliza Microsoft Graph API autenticado con credenciales corporativas en C:\Proyectos\m365-report-mailer\.env
 *
 * Muestra el consolidado de todos los grupos comerciales (Directores: Rafael Novoa, Angélica Caballero, Óscar Beltrán, Miller Romero)
 * Incluye archivo Excel maestro con 3 hojas: (1) Resumen por Director, (2) Resumen por Ejecutivo, (3) Detalle General.
 * Fecha de hoy: 21/09/2026
 */

const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

const TARGET_TEST_EMAIL = 'especialista.preventa@provexpress.com.co';
const TODAY_CUTOFF = '2026-09-21';
const TODAY_FORMATTED = '21 de septiembre de 2026';

// 1. Cargar dependencias de m365-report-mailer
const m365Dir = 'C:/Proyectos/m365-report-mailer';
const { ClientSecretCredential } = require(path.join(m365Dir, 'node_modules/@azure/identity'));
const axios = require(path.join(m365Dir, 'node_modules/axios'));

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

function levenshtein(a, b) {
  const previous = Array.from({ length: b.length + 1 }, (_, index) => index);
  for (let i = 1; i <= a.length; i += 1) {
    let diagonal = previous[0];
    previous[0] = i;
    for (let j = 1; j <= b.length; j += 1) {
      const upper = previous[j];
      previous[j] = Math.min(
        previous[j] + 1,
        previous[j - 1] + 1,
        diagonal + (a[i - 1] === b[j - 1] ? 0 : 1)
      );
      diagonal = upper;
    }
  }
  return previous[b.length];
}

function tokenMatches(token, candidateList) {
  if (candidateList.includes(token)) return true;
  return candidateList.some((cand) => {
    if (cand === token) return true;
    const shorter = cand.length <= token.length ? cand : token;
    const longer = cand.length > token.length ? cand : token;
    if (shorter.length >= 3 && longer.startsWith(shorter)) return true;
    if (token.length >= 5 && cand.length >= 5 && levenshtein(token, cand) <= 1) return true;
    return false;
  });
}

function nameMatches(fullName, targetName) {
  const normFull = normalizeName(fullName);
  const normTarget = normalizeName(targetName);
  if (!normFull || !normTarget) return false;
  if (normFull === normTarget || normFull.includes(normTarget) || normTarget.includes(normFull)) return true;

  const fullTokens = normFull.split(' ').filter((t) => t.length >= 3);
  const targetTokens = normTarget.split(' ').filter((t) => t.length >= 3);

  const matched = targetTokens.filter((t) => tokenMatches(t, fullTokens));
  return matched.length >= 2 || (targetTokens.length === 1 && matched.length === 1);
}

const ESTRUCTURA_COMERCIAL_2026 = {
  directores: {
    1: { nombre: 'Rafael Novoa', email: 'rafael.novoa@provexpress.com.co', genero: 'M' },
    2: { nombre: 'Angélica Caballero', email: 'angelica.caballero@provexpress.com.co', genero: 'F' },
    3: { nombre: 'Óscar Beltrán', email: 'oscar.beltran@provexpress.com.co', genero: 'M' },
    4: { nombre: 'Miller Romero', email: 'miller.romero@provexpress.com.co', genero: 'M' },
  },
  ejecutivos: {
    // Grupo 1: Novoa (9)
    'rosmira.rojas@provexpress.com.co': { grupo: 1, nombre: 'Rosmira Rojas', archivo: 'Rosmira Rojas.xlsx' },
    'mario.reyes@provexpress.com.co': { grupo: 1, nombre: 'Mario Reyes', archivo: 'Mario Reyes.xlsx' },
    'wilson.sanchez@provexpress.com.co': { grupo: 1, nombre: 'Wilson Sánchez', archivo: 'Wilson Fernando Sánchez.xlsx' },
    'maria.cruz@provexpress.com.co': { grupo: 1, nombre: 'María Eugenia Cruz', archivo: 'Maria Eugenia Cruz.xlsx' },
    'javier.cortes@provexpress.com.co': { grupo: 1, nombre: 'Javier Cortés', archivo: 'Javier Cortés.xlsx' },
    'rosa.mendoza@provexpress.com.co': { grupo: 1, nombre: 'Rosa Mendoza', archivo: 'Rosa María Mendoza.xlsx' },
    'mariela.ramirez@provexpress.com.co': { grupo: 1, nombre: 'Mariela Ramírez', archivo: 'Mariela Ramírez.xlsx' },
    'jenny.gonzalez@provexpress.com.co': { grupo: 1, nombre: 'Jenny Gónzalez', archivo: 'Jenny Gónzalez.xlsx' },
    'julieth.galindo@provexpress.com.co': { grupo: 1, nombre: 'Julieth Galindo', archivo: 'Julieth Galindo.xlsx' },

    // Grupo 2: Caballero (11)
    'angela.torres@provexpress.com.co': { grupo: 2, nombre: 'Ángela Torres', archivo: 'Ángela Torres.xlsx' },
    'andrea.vargas@provexpress.com.co': { grupo: 2, nombre: 'Yurany Andrea Vargas', archivo: 'Yurany Andrea Vargas.xlsx' },
    'alejandra.velasquez@provexpress.com.co': { grupo: 2, nombre: 'Alejandra Velásquez', archivo: 'Alejandra Velásquez.xlsx' },
    'fernando.quinonez@provexpress.com.co': { grupo: 2, nombre: 'Fernando Quiñonez', archivo: 'Fernando Quiñonez.xlsx' },
    'johana.mojica@provexpress.com.co': { grupo: 2, nombre: 'Jasbleidy Mójica', archivo: 'Jasbleidy Mójica.xlsx' },
    'johanna.jaime@provexpress.com.co': { grupo: 2, nombre: 'Johanna Jaime', archivo: 'Johanna Jaime.xlsx' },
    'dayana.chala@provexpress.com.co': { grupo: 2, nombre: 'Dayana Chala', archivo: 'Dayana Chala.xlsx' },
    'yovanny.herrera@provexpress.com.co': { grupo: 2, nombre: 'Yovanny Herrera', archivo: 'Yovanny Herrera.xlsx' },
    'cesar.cespedes@provexpress.com.co': { grupo: 2, nombre: 'César Céspedes', archivo: 'César Cespedes.xlsx' },
    'daniel.galindo@provexpress.com.co': { grupo: 2, nombre: 'Daniel Galindo', archivo: 'Daniel Galindo.xlsx' },
    'adriana.cucaita@provexpress.com.co': { grupo: 2, nombre: 'Adriana Cucaita', archivo: 'Adriana Cucaita.xlsx' },

    // Grupo 3: Beltrán (10)
    'paola.garcia@provexpress.com.co': { grupo: 3, nombre: 'Gina García', archivo: 'Gina García.xlsx' },
    'karen.carrillo@provexpress.com.co': { grupo: 3, nombre: 'Karent Carrillo', archivo: 'Karent Carrillo.xlsx' },
    'lington.linares@provexpress.com.co': { grupo: 3, nombre: 'Lington Linares', archivo: 'Lington Linares.xlsx' },
    'angelica.alvarez@provexpress.com.co': { grupo: 3, nombre: 'Angélica Álvarez', archivo: 'Angélica Álvarez.xlsx' },
    'andres.pena@provexpress.com.co': { grupo: 3, nombre: 'Andrés Peña', archivo: 'Andrés Peña.xlsx' },
    'tatiana.parra@provexpress.com.co': { grupo: 3, nombre: 'Tatiana Parra', archivo: 'Tatiana Parra.xlsx' },
    'claudia.triana@provexpress.com.co': { grupo: 3, nombre: 'Claudia Triana', archivo: 'Claudia Triana.xlsx' },
    'dilma.cuesta@provexpress.com.co': { grupo: 3, nombre: 'Dilma Cuesta', archivo: 'Dilma Cuesta.xlsx' },
    'juan.martinez@provexpress.com.co': { grupo: 3, nombre: 'Juan Martínez', archivo: 'Juan Martínez.xlsx' },
    'deisy.mogollon@provexpress.com.co': { grupo: 3, nombre: 'Deisy Mogollón', archivo: 'Deisy Mogollón.xlsx' },

    // Grupo 4: Romero (8)
    'astrid.jimenez@provexpress.com.co': { grupo: 4, nombre: 'Astrid Jiménez', archivo: 'Astrid Jiménez.xlsx' },
    'maria.briceno@provexpress.com.co': { grupo: 4, nombre: 'María Paola Briceño', archivo: 'María Paola Briceño.xlsx' },
    'dafne.ruiz@provexpress.com.co': { grupo: 4, nombre: 'Dafne Ruiz', archivo: 'Dafne Lizeth Ruiz.xlsx' },
    'jessica.valencia@provexpress.com.co': { grupo: 4, nombre: 'Jessica Valencia', archivo: 'Jessica Valencia.xlsx' },
    'jhonatan.acevedo@provexpress.com.co': { grupo: 4, nombre: 'Jhonatan Acevedo', archivo: 'Jhonatan Acevedo.xlsx' },
    'camilo.hernandez@provexpress.com.co': { grupo: 4, nombre: 'Camilo Hernández', archivo: 'Jhonatan Camilo Hernández.xlsx' },
    'yeison.urrego@provexpress.com.co': { grupo: 4, nombre: 'Yeison Urrego', archivo: 'Yeison Urrego.xlsx' },
    'diana.castro@provexpress.com.co': { grupo: 4, nombre: 'Diana Castro', archivo: 'Diana Catalina Castro.xlsx' },
  },
};

const DIRECTORES = [
  { grupo: 1, nombre: 'Rafael Novoa', email: 'rafael.novoa@provexpress.com.co' },
  { grupo: 2, nombre: 'Angélica Caballero', email: 'angelica.caballero@provexpress.com.co' },
  { grupo: 3, nombre: 'Óscar Beltrán', email: 'oscar.beltran@provexpress.com.co' },
  { grupo: 4, nombre: 'Miller Romero', email: 'miller.romero@provexpress.com.co' },
];

async function main() {
  console.log('================================================================');
  console.log('🚀 PREPARANDO INFORME GERENCIAL PARA DIRECCIÓN Y GERENCIA COMERCIAL');
  console.log(`📩 Destinatario de prueba: ${TARGET_TEST_EMAIL}`);
  console.log(`📤 Remitente institucional: ${senderEmail}`);
  console.log(`📅 Fecha de corte activa: ${TODAY_FORMATTED} (${TODAY_CUTOFF})`);
  console.log('👑 Destinatarios corporativos: Rafael Novoa (Director Comercial) & Juan Novoa (Gerente)');
  console.log('================================================================');

  // 1. Cargar remisiones del archivo real
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(path.join(__dirname, '../Remisiones.xlsx'));
  const sis = wb.getWorksheet('Base-SIS');

  const rawCutoff = '2026-09-16'; // datos más recientes del archivo
  const rawRows = [];

  for (let r = 2; r <= sis.rowCount; r++) {
    const row = sis.getRow(r);
    const dVal = row.getCell(1).value;
    const iso = dVal instanceof Date ? dVal.toISOString().slice(0, 10) : String(dVal || '').trim();
    if (iso === rawCutoff) {
      rawRows.push({
        emp: String(row.getCell(2).value || '').trim(),
        nit: String(row.getCell(3).value || '').trim(),
        company: String(row.getCell(4).value || '').trim(),
        merchandise: Number(row.getCell(5).value || 0),
        tax: Number(row.getCell(6).value || 0),
        total: Number(row.getCell(7).value || 0),
        age: Number(row.getCell(9).value || 0),
        doc: String(row.getCell(10).value || '').trim(),
        order: String(row.getCell(11).value || '').trim(),
      });
    }
  }

  // Clasificar cada fila
  const allRemisiones = [];
  const execEntries = Object.entries(ESTRUCTURA_COMERCIAL_2026.ejecutivos);
  const dirEntries = Object.entries(ESTRUCTURA_COMERCIAL_2026.directores);

  for (const r of rawRows) {
    let matched = null;

    // 1. Probar contra ejecutivos comerciales
    for (const [email, exec] of execEntries) {
      if (nameMatches(r.emp, exec.nombre) || nameMatches(r.emp, exec.archivo)) {
        matched = {
          ...r,
          email,
          employee: exec.nombre,
          grupo: exec.grupo,
          dirName: ESTRUCTURA_COMERCIAL_2026.directores[exec.grupo].nombre,
          isDirector: false,
          isUnassigned: false,
        };
        break;
      }
    }

    // 2. Probar si es remisión directa a nombre del Director
    if (!matched) {
      for (const [gStr, dir] of dirEntries) {
        if (nameMatches(r.emp, dir.nombre)) {
          matched = {
            ...r,
            email: dir.email,
            employee: `${dir.nombre} (Gestión Directa)`,
            grupo: Number(gStr),
            dirName: dir.nombre,
            isDirector: true,
            isUnassigned: false,
          };
          break;
        }
      }
    }

    // 3. Cuentas Especiales / Otras Áreas
    if (!matched) {
      matched = {
        ...r,
        email: 'corporativo@provexpress.com.co',
        employee: r.emp,
        grupo: 0,
        dirName: 'Otras Áreas / Especiales',
        isDirector: false,
        isUnassigned: true,
      };
    }

    allRemisiones.push(matched);
  }

  const totalCount = allRemisiones.length;
  const totalValue = allRemisiones.reduce((s, x) => s + x.total, 0);
  const avgAge = totalCount > 0 ? Math.round(allRemisiones.reduce((s, x) => s + x.age, 0) / totalCount) : 0;

  console.log(`✓ Total Remisiones Empresa: ${totalCount}`);
  console.log(`✓ Total Valor Empresa: ${formatCOP(totalValue)}`);
  console.log(`✓ Antigüedad Promedio Global: ${avgAge} días`);

  // Consolidar por grupo
  const groupsSummary = DIRECTORES.map((d) => {
    const gRem = allRemisiones.filter((r) => r.grupo === d.grupo);
    const count = gRem.length;
    const val = gRem.reduce((s, r) => s + r.total, 0);
    const age = count > 0 ? Math.round(gRem.reduce((s, r) => s + r.age, 0) / count) : 0;
    const pct = totalValue > 0 ? Number(((val / totalValue) * 100).toFixed(1)) : 0;

    const byAge = [...gRem].sort((a, b) => b.age - a.age);
    const crit = byAge[0] || null;

    // Asesores activos únicos (excluyendo gestión directa del director)
    const uniqueExecs = new Set(gRem.filter((r) => !r.isDirector).map((r) => r.employee)).size;
    const totalExecs = d.grupo === 1 ? 9 : d.grupo === 2 ? 11 : d.grupo === 3 ? 10 : 8;

    let status = 'Al día';
    let statusColor = '#15803D';
    let statusBg = '#DCFCE7';
    if (age > 25 || (crit && crit.age > 60)) {
      status = 'Atención prioritaria';
      statusColor = '#B91C1C';
      statusBg = '#FEE2E2';
    } else if (age > 15) {
      status = 'Gestión activa';
      statusColor = '#1E3A8A';
      statusBg = '#DBEAFE';
    }

    return {
      grupo: d.grupo,
      nombre: d.nombre,
      email: d.email,
      count,
      total: val,
      avgAge: age,
      pct,
      crit,
      activeExecs: uniqueExecs,
      totalExecs,
      status,
      statusColor,
      statusBg,
    };
  });

  // Cuentas Especiales / Otras Áreas
  const unassignedRem = allRemisiones.filter((r) => r.isUnassigned || r.grupo === 0);
  const unassignedTotal = unassignedRem.reduce((s, r) => s + r.total, 0);
  const unassignedAvgAge = unassignedRem.length > 0 ? Math.round(unassignedRem.reduce((s, r) => s + r.age, 0) / unassignedRem.length) : 0;
  const unassignedPct = totalValue > 0 ? Number(((unassignedTotal / totalValue) * 100).toFixed(1)) : 0;
  const unassignedCrit = [...unassignedRem].sort((a, b) => b.age - a.age)[0] || null;
  const unassignedSummary = {
    label: 'Cuentas Especiales / Otras Áreas',
    count: unassignedRem.length,
    total: unassignedTotal,
    avgAge: unassignedAvgAge,
    pct: unassignedPct,
    crit: unassignedCrit,
  };

  const totalActiveExecs = groupsSummary.reduce((s, g) => s + g.activeExecs, 0);
  const totalAllExecs = groupsSummary.reduce((s, g) => s + g.totalExecs, 0);

  // Top Oportunidades Globales
  const byValAll = [...allRemisiones].sort((a, b) => b.total - a.total);
  const topMayorValor = byValAll[0] || null;

  const byAgeAll = [...allRemisiones].sort((a, b) => b.age - a.age);
  const topMayorAntiguedad = byAgeAll[0] || null;

  // Agrupar asesores comerciales para Top 5 (solo asesores comerciales)
  const execRanking = [];
  for (const [email, exec] of execEntries) {
    const execRem = allRemisiones.filter((r) => !r.isDirector && r.grupo === exec.grupo && r.employee === exec.nombre);
    const count = execRem.length;
    const total = execRem.reduce((s, r) => s + r.total, 0);
    const age = count > 0 ? Math.round(execRem.reduce((s, r) => s + r.age, 0) / count) : 0;
    const pct = totalValue > 0 ? Number(((total / totalValue) * 100).toFixed(1)) : 0;
    execRanking.push({
      email,
      name: exec.nombre,
      grupo: exec.grupo,
      dirName: ESTRUCTURA_COMERCIAL_2026.directores[exec.grupo].nombre,
      count,
      total,
      avgAge: age,
      pct,
    });
  }

  // Ordenar por total por facturar desc
  execRanking.sort((a, b) => b.total - a.total);
  const top5Execs = execRanking.slice(0, 5);

  // 2. Construir Libro Excel Maestro con 3 Hojas
  const excelWb = new ExcelJS.Workbook();
  excelWb.creator = 'Provexpress SAS - Sistema de Gestión de Remisiones';
  excelWb.created = new Date();

  // ── HOJA 1: RESUMEN POR DIRECTOR ──
  const wsDir = excelWb.addWorksheet('Resumen por Director', {
    views: [{ state: 'frozen', ySplit: 5 }],
  });
  wsDir.mergeCells('A1:K1');
  const h1 = wsDir.getCell('A1');
  h1.value = 'PROVEXPRESS SAS · CONSOLIDADO GERENCIAL DE REMISIONES · ALTA DIRECCIÓN';
  h1.font = { name: 'Calibri', size: 13, bold: true, color: { argb: 'FFFFFFFF' } };
  h1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
  h1.alignment = { horizontal: 'center', vertical: 'middle' };
  wsDir.getRow(1).height = 28;

  wsDir.getCell('A2').value = 'Destinatario:';
  wsDir.getCell('A2').font = { bold: true, size: 10, color: { argb: 'FF475569' } };
  wsDir.getCell('B2').value = 'Rafael Novoa / Juan Novoa (Dirección Comercial & Gerencia)';
  wsDir.getCell('B2').font = { bold: true, size: 11, color: { argb: 'FF0F172A' } };

  wsDir.getCell('E2').value = 'Fecha de Corte:';
  wsDir.getCell('E2').font = { bold: true, size: 10, color: { argb: 'FF475569' } };
  wsDir.getCell('F2').value = TODAY_FORMATTED;
  wsDir.getCell('F2').font = { bold: true, size: 11, color: { argb: 'FF0F172A' } };

  wsDir.getCell('I2').value = 'Remisiones Empresa:';
  wsDir.getCell('I2').font = { bold: true, size: 10, color: { argb: 'FF475569' } };
  wsDir.getCell('J2').value = totalCount;
  wsDir.getCell('J2').font = { bold: true, size: 11, color: { argb: 'FF1E3A8A' } };

  wsDir.getCell('A3').value = 'Antigüedad Promedio:';
  wsDir.getCell('A3').font = { bold: true, size: 10, color: { argb: 'FF475569' } };
  wsDir.getCell('B3').value = `${avgAge} días`;
  wsDir.getCell('B3').font = { bold: true, size: 11, color: { argb: 'FFDC2626' } };

  wsDir.getCell('I3').value = 'Total por Facturar:';
  wsDir.getCell('I3').font = { bold: true, size: 10, color: { argb: 'FF475569' } };
  wsDir.getCell('J3').value = totalValue;
  wsDir.getCell('J3').numFmt = '"$"#,##0';
  wsDir.getCell('J3').font = { bold: true, size: 11, color: { argb: 'FF15803D' } };

  wsDir.getRow(4).height = 8;

  const dirCols = ['N°', 'Grupo', 'Director Comercial', 'Correo Corporativo', 'Asesores Activos', 'Total Asesores', 'Remisiones Abiertas', 'Valor por Facturar ($ COP)', '% Participación', 'Antigüedad Promedio', 'Estado'];
  const hRow1 = wsDir.getRow(5);
  hRow1.height = 24;
  dirCols.forEach((text, i) => {
    const c = hRow1.getCell(i + 1);
    c.value = text;
    c.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F172A' } };
    c.alignment = { horizontal: i === 7 ? 'right' : i === 0 || i === 1 || i === 4 || i === 5 || i === 6 || i === 8 || i === 9 || i === 10 ? 'center' : 'left', vertical: 'middle' };
  });

  groupsSummary.forEach((g, idx) => {
    const r = wsDir.getRow(6 + idx);
    r.height = 22;
    r.getCell(1).value = idx + 1;
    r.getCell(2).value = `Grupo ${g.grupo}`;
    r.getCell(3).value = g.nombre;
    r.getCell(3).font = { bold: true };
    r.getCell(4).value = g.email;
    r.getCell(5).value = g.activeExecs;
    r.getCell(6).value = g.totalExecs;
    r.getCell(7).value = g.count;
    r.getCell(7).font = { bold: true };
    r.getCell(8).value = g.total;
    r.getCell(8).numFmt = '"$"#,##0';
    r.getCell(8).font = { bold: true, color: { argb: 'FF15803D' } };
    r.getCell(9).value = `${g.pct}%`;
    r.getCell(9).font = { bold: true, color: { argb: 'FF1E3A8A' } };
    r.getCell(10).value = `${g.avgAge} días`;
    r.getCell(10).font = { bold: true, color: g.avgAge > 25 ? { argb: 'FFDC2626' } : { argb: 'FF0F172A' } };
    r.getCell(11).value = g.status;
    r.getCell(11).font = { bold: true };

    for (let col = 1; col <= 11; col++) {
      r.getCell(col).alignment = { horizontal: col === 8 ? 'right' : col === 1 || col === 2 || col === 5 || col === 6 || col === 7 || col === 9 || col === 10 || col === 11 ? 'center' : 'left', vertical: 'middle' };
    }
  });

  let currentDirRowIdx = 6 + groupsSummary.length;

  // Fila de Cuentas Especiales / Otras Áreas
  if (unassignedSummary && unassignedSummary.count > 0) {
    const uRow = wsDir.getRow(currentDirRowIdx);
    uRow.height = 22;
    uRow.getCell(1).value = groupsSummary.length + 1;
    uRow.getCell(2).value = 'Especiales';
    uRow.getCell(2).font = { bold: true, color: { argb: 'FF64748B' } };
    uRow.getCell(3).value = unassignedSummary.label;
    uRow.getCell(3).font = { bold: true, color: { argb: 'FF475569' } };
    uRow.getCell(4).value = 'corporativo@provexpress.com.co';
    uRow.getCell(4).font = { color: { argb: 'FF94A3B8' } };
    uRow.getCell(5).value = '—';
    uRow.getCell(6).value = '—';
    uRow.getCell(7).value = unassignedSummary.count;
    uRow.getCell(7).font = { bold: true };
    uRow.getCell(8).value = unassignedSummary.total;
    uRow.getCell(8).numFmt = '"$"#,##0';
    uRow.getCell(8).font = { bold: true, color: { argb: 'FF15803D' } };
    uRow.getCell(9).value = `${unassignedSummary.pct}%`;
    uRow.getCell(9).font = { bold: true, color: { argb: 'FF64748B' } };
    uRow.getCell(10).value = `${unassignedSummary.avgAge} días`;
    uRow.getCell(10).font = { bold: true, color: { argb: 'FF475569' } };
    uRow.getCell(11).value = 'En gestión';
    uRow.getCell(11).font = { bold: true, color: { argb: 'FF475569' } };

    for (let col = 1; col <= 11; col++) {
      uRow.getCell(col).alignment = { horizontal: col === 8 ? 'right' : col === 1 || col === 2 || col === 5 || col === 6 || col === 7 || col === 9 || col === 10 || col === 11 ? 'center' : 'left', vertical: 'middle' };
      uRow.getCell(col).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };
    }
    currentDirRowIdx++;
  }

  const totDirR = wsDir.getRow(currentDirRowIdx);
  totDirR.height = 24;
  totDirR.getCell(1).value = 'TOTAL';
  totDirR.getCell(2).value = 'EMPRESA';
  totDirR.getCell(3).value = 'CONSOLIDADO COMPAÑÍA';
  totDirR.getCell(5).value = totalActiveExecs;
  totDirR.getCell(6).value = totalAllExecs;
  totDirR.getCell(7).value = totalCount;
  totDirR.getCell(8).value = totalValue;
  totDirR.getCell(8).numFmt = '"$"#,##0';
  totDirR.getCell(9).value = '100%';
  totDirR.getCell(10).value = `${avgAge} días`;
  totDirR.getCell(11).value = 'GENERAL';
  for (let c = 1; c <= 11; c++) {
    const cell = totDirR.getCell(c);
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F172A' } };
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.alignment = { horizontal: c === 8 ? 'right' : c === 1 || c === 2 || c === 5 || c === 6 || c === 7 || c === 9 || c === 10 || c === 11 ? 'center' : 'left', vertical: 'middle' };
  }
  wsDir.columns = [{ width: 6 }, { width: 14 }, { width: 26 }, { width: 34 }, { width: 16 }, { width: 16 }, { width: 18 }, { width: 22 }, { width: 16 }, { width: 20 }, { width: 20 }];

  // ── HOJA 2: RESUMEN POR EJECUTIVO ──
  const wsExec = excelWb.addWorksheet('Resumen por Ejecutivo', {
    views: [{ state: 'frozen', ySplit: 4 }],
  });
  wsExec.mergeCells('A1:J1');
  const h2 = wsExec.getCell('A1');
  h2.value = 'PROVEXPRESS SAS · CONSOLIDADO DE REMISIONES POR ASESOR COMERCIAL';
  h2.font = { name: 'Calibri', size: 13, bold: true, color: { argb: 'FFFFFFFF' } };
  h2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
  h2.alignment = { horizontal: 'center', vertical: 'middle' };
  wsExec.getRow(1).height = 28;

  wsExec.getCell('A2').value = `Total Asesores con pendientes: ${totalActiveExecs} de ${totalAllExecs} en la compañía | Fecha de corte: ${TODAY_FORMATTED}`;
  wsExec.getCell('A2').font = { bold: true, size: 10, color: { argb: 'FF475569' } };
  wsExec.getRow(3).height = 8;

  const execCols = ['N°', 'Grupo', 'Director Responsable', 'Asesor Comercial', 'Correo Corporativo', 'Remisiones Abiertas', 'Valor por Facturar ($ COP)', '% Participación', 'Antigüedad Promedio', 'Estado'];
  const hRow2 = wsExec.getRow(4);
  hRow2.height = 24;
  execCols.forEach((text, i) => {
    const c = hRow2.getCell(i + 1);
    c.value = text;
    c.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
    c.alignment = { horizontal: i === 6 ? 'right' : i === 0 || i === 1 || i === 5 || i === 7 || i === 8 || i === 9 ? 'center' : 'left', vertical: 'middle' };
  });

  // Lista completa de ejecutivos + gestión directa directores + cuentas especiales
  const fullExecListForSheet = [];

  // 1. Ejecutivos comerciales (38)
  for (const [email, exec] of execEntries) {
    const userRem = allRemisiones.filter((r) => !r.isDirector && r.grupo === exec.grupo && r.employee === exec.nombre);
    const count = userRem.length;
    const total = userRem.reduce((s, r) => s + r.total, 0);
    const uAvgAge = count > 0 ? Math.round(userRem.reduce((s, r) => s + r.age, 0) / count) : 0;
    const uPct = totalValue > 0 ? Number(((total / totalValue) * 100).toFixed(1)) : 0;
    fullExecListForSheet.push({
      grupo: `Grupo ${exec.grupo}`,
      dirName: ESTRUCTURA_COMERCIAL_2026.directores[exec.grupo].nombre,
      name: exec.nombre,
      email,
      count,
      total,
      avgAge: uAvgAge,
      pct: uPct,
    });
  }

  // 2. Gestión directa directores
  for (const [gStr, dir] of dirEntries) {
    const dirRem = allRemisiones.filter((r) => r.isDirector && r.grupo === Number(gStr));
    if (dirRem.length > 0) {
      const count = dirRem.length;
      const total = dirRem.reduce((s, r) => s + r.total, 0);
      const uAvgAge = count > 0 ? Math.round(dirRem.reduce((s, r) => s + r.age, 0) / count) : 0;
      const uPct = totalValue > 0 ? Number(((total / totalValue) * 100).toFixed(1)) : 0;
      fullExecListForSheet.push({
        grupo: `Grupo ${gStr}`,
        dirName: dir.nombre,
        name: `${dir.nombre} (Gestión Directa)`,
        email: dir.email,
        count,
        total,
        avgAge: uAvgAge,
        pct: uPct,
      });
    }
  }

  // 3. Cuentas Especiales / Otras Áreas
  if (unassignedRem.length > 0) {
    const unMap = new Map();
    unassignedRem.forEach((r) => {
      const c = unMap.get(r.emp) || { count: 0, total: 0, ageSum: 0 };
      c.count++;
      c.total += r.total;
      c.ageSum += r.age;
      unMap.set(r.emp, c);
    });
    for (const [emp, d] of unMap.entries()) {
      const uAvgAge = Math.round(d.ageSum / d.count);
      const uPct = totalValue > 0 ? Number(((d.total / totalValue) * 100).toFixed(1)) : 0;
      fullExecListForSheet.push({
        grupo: 'Especial',
        dirName: 'Otras Áreas',
        name: emp,
        email: 'corporativo@provexpress.com.co',
        count: d.count,
        total: d.total,
        avgAge: uAvgAge,
        pct: uPct,
      });
    }
  }

  // Ordenar por total desc
  fullExecListForSheet.sort((a, b) => b.total - a.total);

  fullExecListForSheet.forEach((e, idx) => {
    const r = wsExec.getRow(5 + idx);
    r.height = 20;
    r.getCell(1).value = idx + 1;
    r.getCell(2).value = e.grupo;
    r.getCell(3).value = e.dirName;
    r.getCell(4).value = e.name;
    r.getCell(4).font = { bold: true };
    r.getCell(5).value = e.email;
    r.getCell(5).font = { color: { argb: 'FF64748B' } };
    r.getCell(6).value = e.count;
    r.getCell(6).font = { bold: true };
    r.getCell(7).value = e.total;
    r.getCell(7).numFmt = '"$"#,##0';
    r.getCell(7).font = { bold: true, color: e.count > 0 ? { argb: 'FF15803D' } : { argb: 'FF94A3B8' } };
    r.getCell(8).value = `${e.pct}%`;
    r.getCell(8).font = { bold: true, color: { argb: 'FF1E3A8A' } };
    r.getCell(9).value = e.count > 0 ? `${e.avgAge} días` : '—';
    r.getCell(9).font = { bold: true, color: e.avgAge > 30 ? { argb: 'FFDC2626' } : { argb: 'FF0F172A' } };
    r.getCell(10).value = e.count === 0 ? 'Al día ✓' : e.avgAge > 30 ? 'Urgente (>30d)' : 'En gestión';
    r.getCell(10).font = { bold: true, color: e.count === 0 ? { argb: 'FF15803D' } : e.avgAge > 30 ? { argb: 'FFDC2626' } : { argb: 'FF1E3A8A' } };

    for (let col = 1; col <= 10; col++) {
      r.getCell(col).alignment = { horizontal: col === 7 ? 'right' : col === 1 || col === 2 || col === 6 || col === 8 || col === 9 || col === 10 ? 'center' : 'left', vertical: 'middle' };
    }
  });

  const totExecR = wsExec.getRow(5 + fullExecListForSheet.length);
  totExecR.height = 24;
  totExecR.getCell(1).value = 'TOTAL';
  totExecR.getCell(2).value = 'TODOS';
  totExecR.getCell(4).value = 'CONSOLIDADO COMPAÑÍA';
  totExecR.getCell(6).value = totalCount;
  totExecR.getCell(7).value = totalValue;
  totExecR.getCell(7).numFmt = '"$"#,##0';
  totExecR.getCell(8).value = '100%';
  totExecR.getCell(9).value = `${avgAge} días`;
  totExecR.getCell(10).value = 'GENERAL';
  for (let c = 1; c <= 10; c++) {
    const cell = totExecR.getCell(c);
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F172A' } };
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.alignment = { horizontal: c === 7 ? 'right' : c === 1 || c === 2 || c === 6 || c === 8 || c === 9 || c === 10 ? 'center' : 'left', vertical: 'middle' };
  }
  wsExec.columns = [{ width: 6 }, { width: 14 }, { width: 24 }, { width: 28 }, { width: 34 }, { width: 18 }, { width: 22 }, { width: 16 }, { width: 20 }, { width: 18 }];

  // ── HOJA 3: DETALLE GENERAL REMISIONES ──
  const wsDet = excelWb.addWorksheet('Detalle General Remisiones', {
    views: [{ state: 'frozen', ySplit: 5 }],
  });
  wsDet.mergeCells('A1:N1');
  const h3 = wsDet.getCell('A1');
  h3.value = 'PROVEXPRESS SAS · DETALLE COMPLETO DE REMISIONES ABIERTAS · TODA LA COMPAÑÍA';
  h3.font = { name: 'Calibri', size: 13, bold: true, color: { argb: 'FFFFFFFF' } };
  h3.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
  h3.alignment = { horizontal: 'center', vertical: 'middle' };
  wsDet.getRow(1).height = 28;

  wsDet.getCell('A2').value = `Total Remisiones: ${totalCount} | Valor Total: ${totalValue} | Fecha de corte: ${TODAY_FORMATTED}`;
  wsDet.getCell('A2').font = { bold: true, size: 10, color: { argb: 'FF475569' } };
  wsDet.getRow(4).height = 8;

  const detCols = ['N°', 'Grupo', 'Director', 'Asesor Comercial', 'Remisión', 'Fecha Emisión', 'Días Abierta', 'Rango Antigüedad', 'NIT', 'Cliente / Razón Social', 'Vr. Mercancía', 'Vr. IVA', 'Vr. Total', 'Pedido'];
  const hRow3 = wsDet.getRow(5);
  hRow3.height = 24;
  detCols.forEach((text, i) => {
    const c = hRow3.getCell(i + 1);
    c.value = text;
    c.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F172A' } };
    c.alignment = { horizontal: i === 10 || i === 11 || i === 12 ? 'right' : i === 0 || i === 1 || i === 4 || i === 5 || i === 6 || i === 7 || i === 8 || i === 13 ? 'center' : 'left', vertical: 'middle' };
  });

  const sortedRem = [...allRemisiones].sort((a, b) => b.age - a.age);
  sortedRem.forEach((r, idx) => {
    const row = wsDet.getRow(6 + idx);
    row.height = 20;
    const isEven = idx % 2 === 0;
    const bg = isEven ? 'FFFFFFFF' : 'FFF8FAFC';

    row.getCell(1).value = idx + 1;
    row.getCell(2).value = r.grupo > 0 ? `Grupo ${r.grupo}` : 'Especial';
    row.getCell(3).value = r.dirName;
    row.getCell(4).value = r.employee;
    row.getCell(4).font = { bold: true };
    row.getCell(5).value = r.doc?.startsWith('REM-') ? r.doc : `REM-${r.doc || 'S/N'}`;
    row.getCell(5).font = { bold: true, color: { argb: 'FF1E3A8A' } };
    row.getCell(6).value = TODAY_CUTOFF;
    row.getCell(7).value = r.age;
    row.getCell(7).font = { bold: true, color: r.age > 15 ? { argb: 'FFDC2626' } : { argb: 'FF0F172A' } };
    row.getCell(8).value = r.age > 15 ? 'Más de 15 días' : (r.age >= 8 ? '8 a 15 días' : '0 a 7 días');
    row.getCell(9).value = r.nit;
    row.getCell(10).value = r.company;
    row.getCell(11).value = r.merchandise;
    row.getCell(11).numFmt = '"$"#,##0';
    row.getCell(12).value = r.tax;
    row.getCell(12).numFmt = '"$"#,##0';
    row.getCell(13).value = r.total;
    row.getCell(13).numFmt = '"$"#,##0';
    row.getCell(13).font = { bold: true, color: { argb: 'FF15803D' } };
    row.getCell(14).value = r.order;

    for (let col = 1; col <= 14; col++) {
      row.getCell(col).alignment = { horizontal: col === 11 || col === 12 || col === 13 ? 'right' : col === 1 || col === 2 || col === 5 || col === 6 || col === 7 || col === 8 || col === 9 || col === 14 ? 'center' : 'left', vertical: 'middle' };
    }
  });

  const totDetR = wsDet.getRow(6 + sortedRem.length);
  totDetR.height = 24;
  totDetR.getCell(1).value = 'TOTAL';
  totDetR.getCell(4).value = 'TOTAL CONSOLIDADO COMPAÑÍA';
  totDetR.getCell(5).value = `${totalCount} docs`;
  totDetR.getCell(7).value = `${avgAge}d prom.`;
  totDetR.getCell(13).value = totalValue;
  totDetR.getCell(13).numFmt = '"$"#,##0';
  for (let c = 1; c <= 14; c++) {
    const cell = totDetR.getCell(c);
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.alignment = { horizontal: c === 11 || c === 12 || c === 13 ? 'right' : c === 1 || c === 2 || c === 5 || c === 6 || c === 7 || c === 8 || c === 9 || c === 14 ? 'center' : 'left', vertical: 'middle' };
  }
  wsDet.columns = [{ width: 6 }, { width: 12 }, { width: 22 }, { width: 26 }, { width: 16 }, { width: 14 }, { width: 12 }, { width: 16 }, { width: 16 }, { width: 34 }, { width: 16 }, { width: 14 }, { width: 18 }, { width: 14 }];

  const excelBuffer = await excelWb.xlsx.writeBuffer();
  const excelBase64 = Buffer.from(excelBuffer).toString('base64');
  const excelFilename = `Remisiones_Consolidado_General_Gerencia_${TODAY_CUTOFF}.xlsx`;
  console.log(`✓ Archivo Excel Maestro Gerencial generado: ${excelFilename} (${excelBuffer.byteLength} bytes con 3 hojas)`);

  // 3. Cargar imágenes inline
  const logoPath = path.join(__dirname, '../public/logo_provexpress_stacked.png');
  const avatarPath = path.join(__dirname, '../public/avatar_man.png');
  const logoBase64 = fs.readFileSync(logoPath).toString('base64');
  const avatarBase64 = fs.readFileSync(avatarPath).toString('base64');

  // Filas de directores para el correo HTML
  const rowsDirectorsHtml = groupsSummary
    .map((g, idx) => {
      const isEven = idx % 2 === 0;
      const bg = isEven ? '#FFFFFF' : '#F8FAFC';

      let ageColor = '#15803D';
      let ageBg = '#DCFCE7';
      if (g.avgAge > 25) {
        ageColor = '#B91C1C';
        ageBg = '#FEE2E2';
      } else if (g.avgAge > 15) {
        ageColor = '#B45309';
        ageBg = '#FEF3C7';
      }

      const critText = g.crit ? `${g.crit.doc} (${g.crit.age}d)` : '—';

      return `
      <tr style="background-color: ${bg}; border-bottom: 1px solid #E2E8F0;">
        <td style="padding: 11px 14px; font-size: 13px; font-weight: 800; color: #1E293B; font-family: 'Segoe UI', Arial, sans-serif;">
          <span style="display: inline-block; width: 22px; height: 22px; line-height: 22px; text-align: center; border-radius: 6px; background-color: #1E3A8A; color: #FFFFFF; font-size: 11px; margin-right: 6px;">
            G${g.grupo}
          </span>
          ${g.nombre}
          <div style="font-size: 10.5px; font-weight: 500; color: #64748B; margin-left: 28px;">
            ${g.email}
          </div>
        </td>
        <td style="padding: 11px 10px; text-align: center; font-size: 12px; font-weight: 700; color: #334155; font-family: 'Segoe UI', Arial, sans-serif;">
          <span style="color: #0F172A; font-weight: 900;">${g.activeExecs}</span> <span style="color: #94A3B8;">/ ${g.totalExecs}</span>
        </td>
        <td style="padding: 11px 10px; text-align: center; font-size: 13px; font-weight: 850; color: #0F172A; font-family: 'Segoe UI', Arial, sans-serif;">
          ${formatNumber(g.count)}
        </td>
        <td style="padding: 11px 14px; text-align: right; font-size: 13.5px; font-weight: 900; color: #15803D; font-family: 'Segoe UI', Arial, sans-serif;">
          ${formatCOP(g.total)}
        </td>
        <td style="padding: 11px 10px; text-align: center; font-size: 12px; font-weight: 800; color: #1E3A8A; font-family: 'Segoe UI', Arial, sans-serif;">
          ${g.pct}%
        </td>
        <td style="padding: 11px 10px; text-align: center;">
          <span style="display: inline-block; padding: 3px 8px; border-radius: 6px; font-size: 11px; font-weight: 700; color: ${ageColor}; background-color: ${ageBg}; font-family: 'Segoe UI', Arial, sans-serif;">
            ${g.avgAge} días
          </span>
        </td>
        <td style="padding: 11px 12px; font-size: 11.5px; color: #475569; font-family: 'Segoe UI', Arial, sans-serif;">
          ${critText}
        </td>
        <td style="padding: 11px 12px; text-align: center;">
          <span style="display: inline-block; padding: 2px 8px; border-radius: 999px; font-size: 10px; font-weight: 700; color: ${g.statusColor}; background-color: ${g.statusBg};">
            ${g.status}
          </span>
        </td>
      </tr>
      `;
    })
    .join('');

  // Filas top 5 ejecutivos
  const rowsTopExecsHtml = top5Execs
    .map((e, idx) => {
      const isEven = idx % 2 === 0;
      const bg = isEven ? '#FFFFFF' : '#F8FAFC';
      return `
      <tr style="background-color: ${bg}; border-bottom: 1px solid #E2E8F0;">
        <td style="padding: 9px 12px; text-align: center; font-size: 12px; font-weight: 800; color: #1E3A8A; font-family: 'Segoe UI', Arial, sans-serif;">
          #${idx + 1}
        </td>
        <td style="padding: 9px 12px; font-size: 12.5px; font-weight: 750; color: #0F172A; font-family: 'Segoe UI', Arial, sans-serif;">
          ${e.name}
          <div style="font-size: 10px; color: #64748B;">Grupo ${e.grupo} · ${e.dirName}</div>
        </td>
        <td style="padding: 9px 10px; text-align: center; font-size: 12px; font-weight: 800; color: #0F172A; font-family: 'Segoe UI', Arial, sans-serif;">
          ${e.count}
        </td>
        <td style="padding: 9px 12px; text-align: right; font-size: 12.5px; font-weight: 900; color: #15803D; font-family: 'Segoe UI', Arial, sans-serif;">
          ${formatCOP(e.total)}
        </td>
        <td style="padding: 9px 10px; text-align: center; font-size: 11.5px; font-weight: 700; color: #D97706; font-family: 'Segoe UI', Arial, sans-serif;">
          ${e.avgAge} días
        </td>
        <td style="padding: 9px 10px; text-align: center; font-size: 11.5px; font-weight: 800; color: #1E3A8A; font-family: 'Segoe UI', Arial, sans-serif;">
          ${e.pct}%
        </td>
      </tr>
      `;
    })
    .join('');

  const htmlContent = `<!DOCTYPE html>
<html lang="es" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Consolidado General de Gestión Comercial · Dirección & Gerencia · Provexpress SAS</title>
  <!--[if mso]>
  <noscript>
    <xml>
      <o:OfficeDocumentSettings>
        <o:PixelsPerInch>96</o:PixelsPerInch>
      </o:OfficeDocumentSettings>
    </xml>
  </noscript>
  <![endif]-->
  <style>
    body, table, td, a { -webkit-text-size-adjust: 100%; -ms-text-size-adjust: 100%; }
    table, td { mso-table-lspace: 0pt; mso-table-rspace: 0pt; }
    img { -ms-interpolation-mode: bicubic; border: 0; outline: none; text-decoration: none; }
    body { margin: 0; padding: 0; width: 100% !important; background-color: #F1F4F8; font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, Helvetica, Arial, sans-serif; }
    @media only screen and (max-width: 920px) {
      .main-card { width: 100% !important; border-radius: 0 !important; }
      .stack-col { display: block !important; width: 100% !important; padding-right: 0 !important; }
      .kpi-cell { display: inline-block !important; width: 48% !important; margin-bottom: 8px !important; }
    }
  </style>
</head>
<body style="margin: 0; padding: 24px 0; background-color: #F1F4F8;">

  <!-- CONTENEDOR PRINCIPAL BLANCO (960px) -->
  <table role="presentation" class="main-card" width="960" align="center" border="0" cellpadding="0" cellspacing="0" style="width: 960px; margin: 0 auto; background-color: #FFFFFF; border-radius: 16px; overflow: hidden; border: 1.5px solid #E2E8F0; box-shadow: 0 4px 20px rgba(15, 23, 42, 0.05);">
    <tr>
      <td style="padding: 24px 28px;">

        <!-- 1. ENCABEZADO INSTITUCIONAL -->
        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="margin-bottom: 20px;">
          <tr>
            <td width="160" valign="middle" align="left" style="padding-right: 18px;">
              <img src="cid:logo_provexpress" width="144" height="110" style="display: block; border: 0; width: 144px; height: 110px; max-width: 144px;" alt="Provexpress" />
            </td>

            <td valign="middle" style="padding-right: 12px;">
              <div style="font-size: 14.5px; color: #475569; font-weight: 600; font-family: 'Segoe UI', Arial, sans-serif;">
                Hola, <strong style="color: #0F172A; font-size: 16px;">Rafael Novoa & Juan Novoa</strong>
                <span style="display: inline-block; margin-left: 8px; font-size: 11px; background-color: #EFF6FF; color: #1E3A8A; padding: 2px 8px; border-radius: 999px; font-weight: 700;">Dirección Comercial & Gerencia General</span>
              </div>
              <h1 style="margin: 3px 0 6px 0; font-size: 23px; font-weight: 900; color: #0F172A; line-height: 1.15; letter-spacing: -0.02em; font-family: 'Segoe UI', Arial, sans-serif;">
                Consolidado General de Gestión Comercial<br>
                <span style="color: #16A34A;">Dirección & Gerencia</span> · Corte ${TODAY_FORMATTED}
              </h1>
              <p style="margin: 0; font-size: 12px; color: #475569; line-height: 1.35; font-family: 'Segoe UI', Arial, sans-serif; max-width: 410px;">
                Visión integral del estado de remisiones abiertas y oportunidades de facturación en los <strong>4 grupos comerciales</strong> y sus <strong>${totalActiveExecs} ejecutivos con gestión</strong>.
              </p>
              <p style="margin: 5px 0 0 0; font-size: 12px; color: #0F172A; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">
                Articulación y seguimiento estratégico para convertir entregas en <span style="color: #16A34A;">ventas facturadas</span>.
              </p>
            </td>

            <td width="190" valign="bottom" align="center" style="padding-right: 10px;">
              <img src="cid:avatar_man" width="180" height="154" style="display: block; border: 0; width: 180px; height: auto; margin: 0 auto;" alt="Dirección Comercial" />
            </td>

            <td width="125" valign="middle" align="right">
              <table role="presentation" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 14px; padding: 14px 8px; width: 120px; text-align: center;">
                <tr>
                  <td align="center">
                    <div style="background-color: #1E3A8A; color: #FFFFFF; width: 26px; height: 26px; border-radius: 50%; font-size: 14px; font-weight: 900; line-height: 26px; margin: 0 auto 8px auto;">
                      ★
                    </div>
                    <div style="font-size: 11.5px; font-weight: 700; color: #1E293B; line-height: 1.3; font-family: 'Segoe UI', Arial, sans-serif;">
                      Control Total<br>
                      Visión<br>
                      <strong style="color: #1E3A8A; font-size: 12.5px;">Corporativa</strong>
                    </div>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>

        <!-- 2. 4 TARJETAS KPI GLOBALES DE LA COMPAÑÍA -->
        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="margin-bottom: 20px;">
          <tr>
            <td width="24%" valign="top" class="kpi-cell" style="padding-right: 8px;">
              <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 12px; padding: 12px 10px; height: 95px;">
                <tr>
                  <td width="32" valign="top">
                    <div style="background-color: #1E3A8A; width: 30px; height: 30px; border-radius: 50%; text-align: center; line-height: 30px; font-size: 15px; color: #FFFFFF;">
                      👔
                    </div>
                  </td>
                  <td valign="top" style="padding-left: 6px;">
                    <div style="font-size: 10.5px; color: #475569; font-weight: 600; line-height: 1.2; font-family: 'Segoe UI', Arial, sans-serif;">
                      Grupos y Ejecutivos<br>con pendientes
                    </div>
                  </td>
                </tr>
                <tr>
                  <td colspan="2" valign="bottom" style="padding-top: 6px;">
                    <div style="font-size: 16px; font-weight: 900; color: #1E3A8A; font-family: 'Segoe UI', Arial, sans-serif;">
                      4 grupos · ${totalActiveExecs} asesores
                    </div>
                  </td>
                </tr>
              </table>
            </td>

            <td width="24%" valign="top" class="kpi-cell" style="padding-right: 8px;">
              <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 12px; padding: 12px 10px; height: 95px;">
                <tr>
                  <td width="32" valign="top">
                    <div style="background-color: #312E81; width: 30px; height: 30px; border-radius: 50%; text-align: center; line-height: 30px; font-size: 15px; color: #FFFFFF;">
                      📦
                    </div>
                  </td>
                  <td valign="top" style="padding-left: 6px;">
                    <div style="font-size: 10.5px; color: #475569; font-weight: 600; line-height: 1.2; font-family: 'Segoe UI', Arial, sans-serif;">
                      Remisiones abiertas<br>en la compañía
                    </div>
                  </td>
                </tr>
                <tr>
                  <td colspan="2" valign="bottom" style="padding-top: 6px;">
                    <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td valign="bottom" style="font-size: 24px; font-weight: 900; color: #312E81; font-family: 'Segoe UI', Arial, sans-serif; line-height: 1;">
                          ${formatNumber(totalCount)}
                        </td>
                        <td align="right" valign="bottom" style="font-size: 17px; color: #94A3B8; opacity: 0.65;">
                          📝
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>

            <td width="28%" valign="top" class="kpi-cell" style="padding-right: 8px;">
              <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 12px; padding: 12px 10px; height: 95px;">
                <tr>
                  <td width="32" valign="top">
                    <div style="background-color: #15803D; width: 30px; height: 30px; border-radius: 50%; text-align: center; line-height: 30px; font-size: 15px; font-weight: 900; color: #FFFFFF;">
                      $
                    </div>
                  </td>
                  <td valign="top" style="padding-left: 6px;">
                    <div style="font-size: 10.5px; color: #475569; font-weight: 600; line-height: 1.2; font-family: 'Segoe UI', Arial, sans-serif;">
                      Total por facturar<br>compañía
                    </div>
                  </td>
                </tr>
                <tr>
                  <td colspan="2" valign="bottom" style="padding-top: 6px;">
                    <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td valign="bottom" style="font-size: 18px; font-weight: 900; color: #15803D; font-family: 'Segoe UI', Arial, sans-serif; line-height: 1;">
                          ${formatCOP(totalValue)}
                        </td>
                        <td align="right" valign="bottom" style="font-size: 17px; color: #86EFAC; opacity: 0.75;">
                          🪙
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>

            <td width="24%" valign="top" class="kpi-cell">
              <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 12px; padding: 12px 10px; height: 95px;">
                <tr>
                  <td width="32" valign="top">
                    <div style="background-color: #D97706; width: 30px; height: 30px; border-radius: 50%; text-align: center; line-height: 30px; font-size: 15px; color: #FFFFFF;">
                      🕒
                    </div>
                  </td>
                  <td valign="top" style="padding-left: 6px;">
                    <div style="font-size: 10.5px; color: #475569; font-weight: 600; line-height: 1.2; font-family: 'Segoe UI', Arial, sans-serif;">
                      Antigüedad promedio<br>global
                    </div>
                  </td>
                </tr>
                <tr>
                  <td colspan="2" valign="bottom" style="padding-top: 6px;">
                    <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td valign="bottom" style="font-size: 22px; font-weight: 900; color: #D97706; font-family: 'Segoe UI', Arial, sans-serif; line-height: 1;">
                          ${avgAge} días
                        </td>
                        <td align="right" valign="bottom" style="font-size: 17px; color: #FDBA74; opacity: 0.75;">
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

        <!-- 3. TABLA 1: EL RESULTADO DE LOS DIRECTIVOS -->
        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="margin-bottom: 8px;">
          <tr>
            <td>
              <span style="font-size: 15.5px; font-weight: 850; color: #0F172A; font-family: 'Segoe UI', Arial, sans-serif;">
                Resultado Consolidado por Directores de Grupo
              </span>
              <span style="color: #1E3A8A; font-size: 14px; margin-left: 4px;">👔</span>
              <span style="font-size: 11.5px; color: #475569; margin-left: 12px; font-family: 'Segoe UI', Arial, sans-serif;">
                Desempeño y volumen pendiente en cada una de las 4 direcciones comerciales.
              </span>
            </td>
          </tr>
        </table>

        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="border: 1.5px solid #E2E8F0; border-radius: 12px; overflow: hidden; border-collapse: separate; border-spacing: 0; margin-bottom: 22px;">
          <thead>
            <tr style="background-color: #0F172A; color: #FFFFFF;">
              <th style="padding: 10px 14px; text-align: left; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">Director & Grupo</th>
              <th style="padding: 10px 10px; text-align: center; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">Asesores Activos</th>
              <th style="padding: 10px 10px; text-align: center; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">Remisiones</th>
              <th style="padding: 10px 14px; text-align: right; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">Valor por Facturar</th>
              <th style="padding: 10px 10px; text-align: center; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">% Empresa</th>
              <th style="padding: 10px 10px; text-align: center; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">Promedio</th>
              <th style="padding: 10px 12px; text-align: left; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">Remisión Crítica</th>
              <th style="padding: 10px 12px; text-align: center; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">Estado</th>
            </tr>
          </thead>
          <tbody style="background-color: #FFFFFF;">
            ${rowsDirectorsHtml}
            ${unassignedSummary && unassignedSummary.count > 0 ? `
            <tr style="background-color: #FAFAFA; border-bottom: 1.5px dashed #CBD5E1;">
              <td style="padding: 10px 14px; font-size: 12.5px; font-weight: 750; color: #475569; font-family: 'Segoe UI', Arial, sans-serif;">
                <span style="display: inline-block; width: 22px; height: 22px; line-height: 22px; text-align: center; border-radius: 6px; background-color: #64748B; color: #FFFFFF; font-size: 10px; margin-right: 6px;">
                  CE
                </span>
                ${unassignedSummary.label}
                <div style="font-size: 10px; font-weight: 500; color: #94A3B8; margin-left: 28px;">
                  Cuentas directas y áreas de apoyo
                </div>
              </td>
              <td style="padding: 10px 10px; text-align: center; font-size: 12px; font-weight: 600; color: #94A3B8; font-family: 'Segoe UI', Arial, sans-serif;">
                —
              </td>
              <td style="padding: 10px 10px; text-align: center; font-size: 12.5px; font-weight: 850; color: #334155; font-family: 'Segoe UI', Arial, sans-serif;">
                ${formatNumber(unassignedSummary.count)}
              </td>
              <td style="padding: 10px 14px; text-align: right; font-size: 13px; font-weight: 850; color: #15803D; font-family: 'Segoe UI', Arial, sans-serif;">
                ${formatCOP(unassignedSummary.total)}
              </td>
              <td style="padding: 10px 10px; text-align: center; font-size: 12px; font-weight: 800; color: #64748B; font-family: 'Segoe UI', Arial, sans-serif;">
                ${unassignedSummary.pct}%
              </td>
              <td style="padding: 10px 10px; text-align: center;">
                <span style="display: inline-block; padding: 2px 7px; border-radius: 6px; font-size: 11px; font-weight: 700; color: #475569; background-color: #F1F5F9; font-family: 'Segoe UI', Arial, sans-serif;">
                  ${unassignedSummary.avgAge} días
                </span>
              </td>
              <td style="padding: 10px 12px; font-size: 11px; color: #64748B; font-family: 'Segoe UI', Arial, sans-serif;">
                ${unassignedSummary.crit ? `${unassignedSummary.crit.doc} (${unassignedSummary.crit.age}d)` : '—'}
              </td>
              <td style="padding: 10px 12px; text-align: center;">
                <span style="display: inline-block; padding: 2px 8px; border-radius: 999px; font-size: 10px; font-weight: 700; color: #475569; background-color: #F1F5F9;">
                  En gestión
                </span>
              </td>
            </tr>
            ` : ''}
            <tr style="background-color: #F1F5F9; border-top: 2px solid #CBD5E1;">
              <td style="padding: 12px 14px; font-size: 13px; font-weight: 900; color: #0F172A; font-family: 'Segoe UI', Arial, sans-serif;">
                TOTAL GENERAL COMPAÑÍA
              </td>
              <td style="padding: 12px 10px; text-align: center; font-size: 12.5px; font-weight: 900; color: #0F172A; font-family: 'Segoe UI', Arial, sans-serif;">
                ${totalActiveExecs} / ${totalAllExecs}
              </td>
              <td style="padding: 12px 10px; text-align: center; font-size: 13.5px; font-weight: 900; color: #0F172A; font-family: 'Segoe UI', Arial, sans-serif;">
                ${formatNumber(totalCount)}
              </td>
              <td style="padding: 12px 14px; text-align: right; font-size: 14px; font-weight: 950; color: #15803D; font-family: 'Segoe UI', Arial, sans-serif;">
                ${formatCOP(totalValue)}
              </td>
              <td style="padding: 12px 10px; text-align: center; font-size: 12px; font-weight: 900; color: #1E3A8A; font-family: 'Segoe UI', Arial, sans-serif;">
                100%
              </td>
              <td style="padding: 12px 10px; text-align: center; font-size: 12px; font-weight: 900; color: #D97706; font-family: 'Segoe UI', Arial, sans-serif;">
                ${avgAge} días
              </td>
              <td colspan="2" style="padding: 12px 12px; text-align: center; font-size: 11px; font-weight: 700; color: #475569; font-family: 'Segoe UI', Arial, sans-serif;">
                4 grupos comerciales bajo seguimiento
              </td>
            </tr>
          </tbody>
        </table>

        <!-- 4. TABLA 2: TOP 5 EJECUTIVOS CON MAYOR VOLUMEN -->
        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="margin-bottom: 8px;">
          <tr>
            <td>
              <span style="font-size: 14.5px; font-weight: 850; color: #0F172A; font-family: 'Segoe UI', Arial, sans-serif;">
                Top 5 Asesores con Mayor Volumen por Facturar
              </span>
              <span style="color: #15803D; font-size: 14px; margin-left: 4px;">📈</span>
              <span style="font-size: 11.5px; color: #475569; margin-left: 12px; font-family: 'Segoe UI', Arial, sans-serif;">
                Concentración de valor para acompañamiento prioritario de la gerencia comercial.
              </span>
            </td>
          </tr>
        </table>

        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="border: 1px solid #E2E8F0; border-radius: 12px; overflow: hidden; border-collapse: separate; border-spacing: 0; margin-bottom: 22px;">
          <thead>
            <tr style="background-color: #1E293B; color: #FFFFFF;">
              <th style="padding: 9px 12px; text-align: center; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif; width: 45px;">Pos.</th>
              <th style="padding: 9px 12px; text-align: left; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">Asesor Comercial</th>
              <th style="padding: 9px 10px; text-align: center; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif; width: 90px;">Remisiones</th>
              <th style="padding: 9px 12px; text-align: right; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif; width: 140px;">Valor por Facturar</th>
              <th style="padding: 9px 10px; text-align: center; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif; width: 90px;">Promedio</th>
              <th style="padding: 9px 10px; text-align: center; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif; width: 85px;">% Empresa</th>
            </tr>
          </thead>
          <tbody style="background-color: #FFFFFF;">
            ${rowsTopExecsHtml}
          </tbody>
        </table>

        <!-- 5. BANNERS ANCHO COMPLETO: TOP VALOR Y TOP ANTIGÜEDAD -->
        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 14px; padding: 14px 18px; margin-bottom: 14px;">
          ${topMayorValor ? `
          <tr>
            <td width="42" valign="middle" align="center" style="padding-right: 12px;">
              <span style="font-size: 26px;">🏆</span>
            </td>
            <td valign="middle" style="padding-right: 14px;">
              <div style="font-size: 10px; font-weight: 700; color: #15803D; text-transform: uppercase; letter-spacing: 0.04em; font-family: 'Segoe UI', Arial, sans-serif;">
                Mayor valor por facturar de la compañía <span style="color: #15803D;">★</span>
              </div>
              <div style="font-size: 15px; font-weight: 900; color: #1E3A8A; font-family: 'Segoe UI', Arial, sans-serif; margin: 1px 0;">
                ${topMayorValor.doc?.startsWith('REM-') ? topMayorValor.doc : `REM-${topMayorValor.doc || 'S/N'}`}
              </div>
              <div style="font-size: 11px; color: #475569; font-family: 'Segoe UI', Arial, sans-serif;">
                ${topMayorValor.company || 'Cliente'}
              </div>
              <div style="font-size: 10.5px; color: #1E3A8A; font-weight: 700; margin-top: 2px;">
                Asesor: ${topMayorValor.employee} (${topMayorValor.dirName})
              </div>
            </td>
            <td width="140" valign="middle" style="padding-right: 14px;">
              <div style="font-size: 15.5px; font-weight: 900; color: #15803D; font-family: 'Segoe UI', Arial, sans-serif;">
                ${formatCOP(topMayorValor.total)}
              </div>
              <div style="font-size: 10px; color: #64748B; font-family: 'Segoe UI', Arial, sans-serif;">
                Monto más representativo
              </div>
            </td>
            <td width="95" valign="middle" style="padding-right: 14px;">
              <div style="font-size: 14px; font-weight: 850; color: #0F172A; font-family: 'Segoe UI', Arial, sans-serif;">
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
                    ¡Mayor impacto<br>global!
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          ` : ''}

          ${(topMayorAntiguedad && (!topMayorValor || topMayorAntiguedad.doc !== topMayorValor.doc)) ? `
          <tr>
            <td colspan="5" style="padding: 10px 0;">
              <div style="border-top: 1px dashed #E2E8F0; height: 1px; line-height: 1px; font-size: 0;">&nbsp;</div>
            </td>
          </tr>
          <tr>
            <td width="42" valign="middle" align="center" style="padding-right: 12px;">
              <span style="font-size: 26px;">⏳</span>
            </td>
            <td valign="middle" style="padding-right: 14px;">
              <div style="font-size: 10px; font-weight: 700; color: #DC2626; text-transform: uppercase; letter-spacing: 0.04em; font-family: 'Segoe UI', Arial, sans-serif;">
                Mayor antigüedad pendiente de la compañía <span style="color: #DC2626;">⏱</span>
              </div>
              <div style="font-size: 15px; font-weight: 900; color: #1E3A8A; font-family: 'Segoe UI', Arial, sans-serif; margin: 1px 0;">
                ${topMayorAntiguedad.doc?.startsWith('REM-') ? topMayorAntiguedad.doc : `REM-${topMayorAntiguedad.doc || 'S/N'}`}
              </div>
              <div style="font-size: 11px; color: #475569; font-family: 'Segoe UI', Arial, sans-serif;">
                ${topMayorAntiguedad.company || 'Cliente'}
              </div>
              <div style="font-size: 10.5px; color: #1E3A8A; font-weight: 700; margin-top: 2px;">
                Asesor: ${topMayorAntiguedad.employee} (${topMayorAntiguedad.dirName})
              </div>
            </td>
            <td width="140" valign="middle" style="padding-right: 14px;">
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
                    ¡Prioridad crítica<br>por tiempo!
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          ` : ''}
        </table>

        <!-- 6. BANNER ARCHIVO ADJUNTO EXCEL MASTER GERENCIAL -->
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
                Hemos adjuntado el libro Excel maestro con <strong>3 hojas de trabajo</strong>: (1) Resumen por Director, (2) Resumen por Ejecutivo Comercial y (3) Detalle General de las <strong>${formatNumber(totalCount)}</strong> remisiones abiertas (${formatCOP(totalValue)}) de toda la empresa.
              </div>
            </td>
          </tr>
        </table>

        <!-- 7. PIE DE PÁGINA -->
        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #F8FAFC; border: 1.5px solid #E2E8F0; border-radius: 14px; padding: 16px 22px; margin-top: 18px;">
          <tr>
            <td width="38" valign="middle" align="center" style="padding-right: 12px;">
              <div style="font-size: 28px; line-height: 1; color: #1E3A8A;">
                💙
              </div>
            </td>
            <td valign="middle">
              <div style="font-size: 12px; color: #475569; font-family: 'Segoe UI', Arial, sans-serif;">
                Liderazgo, articulación comercial y compromiso corporativo.
              </div>
              <div style="font-size: 13.5px; font-weight: 850; color: #1E3A8A; font-family: 'Segoe UI', Arial, sans-serif; margin-top: 2px;">
                ¡Sigamos impulsando los resultados de Provexpress juntos!
              </div>
            </td>
            <td width="55" align="right" valign="middle">
              <svg width="48" height="38" viewBox="0 0 52 42" fill="none" xmlns="http://www.w3.org/2000/svg">
                <rect x="8" y="2" width="36" height="26" rx="3" fill="#FFFFFF" stroke="#60A5FA" stroke-width="1.5"/>
                <line x1="13" y1="8" x2="31" y2="8" stroke="#93C5FD" stroke-width="2" stroke-linecap="round"/>
                <line x1="13" y1="13" x2="27" y2="13" stroke="#93C5FD" stroke-width="2" stroke-linecap="round"/>
                <path d="M4 14H48C49.1046 14 50 14.8954 50 16V36C50 37.1046 49.1046 38 48 38H4C2.89543 38 2 37.1046 2 36V16C2 14.8954 2.89543 14 4 14Z" fill="#2563EB"/>
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
`;

  // 4. Autenticar y enviar vía Microsoft Graph
  console.log('\n🔑 Autenticando con Microsoft Entra ID (Graph)...');
  const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
  const tokenResponse = await credential.getToken('https://graph.microsoft.com/.default');
  const token = tokenResponse.token;
  console.log('✓ Token de Microsoft Graph obtenido exitosamente.');

  console.log(`\n📨 Enviando informe gerencial de prueba a ${TARGET_TEST_EMAIL}...`);
  const graphSendUrl = `https://graph.microsoft.com/v1.0/users/${senderEmail}/sendMail`;

  const payload = {
    message: {
      subject: `[PRUEBA] Consolidado General de Gestión Comercial · Dirección & Gerencia · Corte ${TODAY_CUTOFF}`,
      body: {
        contentType: 'HTML',
        content: htmlContent,
      },
      toRecipients: [
        {
          emailAddress: {
            address: TARGET_TEST_EMAIL,
          },
        },
      ],
      attachments: [
        {
          '@odata.type': '#microsoft.graph.fileAttachment',
          name: 'logo_provexpress.png',
          contentType: 'image/png',
          contentBytes: logoBase64,
          isInline: true,
          contentId: 'logo_provexpress',
        },
        {
          '@odata.type': '#microsoft.graph.fileAttachment',
          name: 'avatar_man.png',
          contentType: 'image/png',
          contentBytes: avatarBase64,
          isInline: true,
          contentId: 'avatar_man',
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
    saveToSentItems: false,
  };

  const response = await axios.post(graphSendUrl, payload, {
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
  });

  console.log(`\n✅ ¡CORREO GERENCIAL ENVIADO CON ÉXITO!`);
  console.log(`   - Código de respuesta HTTP: ${response.status} ${response.statusText || 'Accepted'}`);
  console.log(`   - Destino: ${TARGET_TEST_EMAIL}`);
  console.log(`   - Adjunto maestro: ${excelFilename}`);
  console.log('================================================================');
}

main().catch((err) => {
  console.error('❌ Error enviando el correo gerencial de prueba:', err.response?.data || err.message);
  process.exit(1);
});
