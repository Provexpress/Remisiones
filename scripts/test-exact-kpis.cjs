const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

function normalizeText(value) {
  if (value == null) return '';
  return String(value)
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

function tokenSimilarity(a, b) {
  if (a === b) return 1;
  const shorter = a.length <= b.length ? a : b;
  const longer = a.length > b.length ? a : b;
  if (shorter.length >= 3 && longer.startsWith(shorter)) return 0.84;
  return 1 - levenshtein(a, b) / Math.max(a.length, b.length, 1);
}

function nameScore(employee, member) {
  const employeeTokens = normalizeText(employee).split(' ').filter((token) => token.length >= 3);
  const memberTokens = normalizeText(member).split(' ').filter((token) => token.length >= 3);
  if (!employeeTokens.length || !memberTokens.length) return 0;
  const matched = memberTokens.map((token) => Math.max(...employeeTokens.map((candidate) => tokenSimilarity(token, candidate))));
  const average = matched.reduce((sum, score) => sum + score, 0) / matched.length;
  const exactMatches = memberTokens.filter((token) => employeeTokens.includes(token)).length;
  return average + Math.min(0.12, exactMatches * 0.04);
}

function matchGroup(employee, groups) {
  const normalized = normalizeText(employee);
  if (!normalized) return null;
  const exact = groups.find((entry) => normalizeText(entry.member) === normalized);
  if (exact) return exact;
  let best = null;
  for (const entry of groups) {
    const score = nameScore(employee, entry.member);
    if (!best || score > best.score) best = { entry, score };
  }
  return best && best.score >= 0.76 ? best.entry : null;
}

async function run() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(path.join(__dirname, '../Remisiones.xlsx'));

  const sheet = wb.getWorksheet('Grupos');
  const groups = [];
  let currentGroup = 0;
  let currentDirector = '';
  sheet.eachRow((row) => {
    const text = String(row.getCell(1).value || '').trim();
    const groupMatch = text.match(/^Grupo\s+(\d+).*Director(?:a)?:\s*(.+)$/i);
    if (groupMatch) {
      currentGroup = Number(groupMatch[1]);
      currentDirector = groupMatch[2].trim();
      // Incluir también al director como miembro del grupo para captar remisiones a su nombre
      groups.push({ group: currentGroup, director: currentDirector, member: currentDirector, isDirector: true });
      return;
    }
    if (!text || /^Ejecutivo Comercial$/i.test(text) || !currentGroup) return;
    groups.push({ group: currentGroup, director: currentDirector, member: text, isDirector: false });
  });

  const sis = wb.getWorksheet('Base-SIS');
  const allRows = [];

  for (let r = 2; r <= sis.rowCount; r++) {
    const row = sis.getRow(r);
    const dVal = row.getCell(1).value;
    const iso = dVal instanceof Date ? dVal.toISOString().slice(0, 10) : String(dVal || '').trim();
    if (iso === '2026-09-16') {
      const emp = String(row.getCell(2).value || '').trim();
      const nit = String(row.getCell(3).value || '').trim();
      const company = String(row.getCell(4).value || '').trim();
      const merchandise = Number(row.getCell(5).value || 0);
      const tax = Number(row.getCell(6).value || 0);
      const total = Number(row.getCell(7).value || 0);
      const emission = String(row.getCell(8).value || '').trim();
      const age = Number(row.getCell(9).value || 0);
      const doc = String(row.getCell(10).value || '').trim();
      const order = String(row.getCell(11).value || '').trim();

      const m = matchGroup(emp, groups);
      allRows.push({
        emp,
        nit,
        company,
        merchandise,
        tax,
        total,
        emission,
        age,
        doc,
        order,
        group: m ? m.group : null,
        director: m ? m.director : 'Sin asignar',
        matchedMember: m ? m.member : null,
        isDirectorMember: m ? m.isDirector : false,
      });
    }
  }

  console.log(`Total remisiones leídas: ${allRows.length}`);
  const grandTotal = allRows.reduce((s, r) => s + r.total, 0);
  console.log(`Total Valor Empresa: $${Math.round(grandTotal).toLocaleString('es-CO')}`);

  const DIRECTORES_INFO = [
    { grupo: 1, nombre: 'Rafael Novoa', email: 'rafael.novoa@provexpress.com.co', totalAsesores: 9 },
    { grupo: 2, nombre: 'Angélica Caballero', email: 'angelica.caballero@provexpress.com.co', totalAsesores: 11 },
    { grupo: 3, nombre: 'Óscar Beltrán', email: 'oscar.beltran@provexpress.com.co', totalAsesores: 10 },
    { grupo: 4, nombre: 'Miller Romero', email: 'miller.romero@provexpress.com.co', totalAsesores: 8 },
  ];

  for (const dir of DIRECTORES_INFO) {
    const gRows = allRows.filter((r) => r.group === dir.grupo);
    const count = gRows.length;
    const totalVal = gRows.reduce((s, r) => s + r.total, 0);
    const avgAge = count > 0 ? Math.round(gRows.reduce((s, r) => s + r.age, 0) / count) : 0;
    const pct = Number(((totalVal / grandTotal) * 100).toFixed(1));

    // Ejecutivos activos (solo ejecutivos comerciales, sin duplicar al director)
    const execMembers = new Set(gRows.filter((r) => !r.isDirectorMember).map((r) => r.matchedMember));
    const activeExecsCount = execMembers.size;

    const sortedByVal = [...gRows].sort((a, b) => b.total - a.total);
    const topVal = sortedByVal[0];

    const sortedByAge = [...gRows].sort((a, b) => b.age - a.age);
    const topAge = sortedByAge[0];

    console.log(`\n======================================================`);
    console.log(`GRUPO ${dir.grupo}: ${dir.nombre}`);
    console.log(`- Remisiones Abiertas: ${count}`);
    console.log(`- Valor Pendiente: $${Math.round(totalVal).toLocaleString('es-CO')}`);
    console.log(`- % Participación: ${pct}%`);
    console.log(`- Antigüedad Promedio: ${avgAge} días`);
    console.log(`- Asesores Activos: ${activeExecsCount} de ${dir.totalAsesores}`);
    console.log(`- Top Mayor Valor: ${topVal.doc} - ${topVal.company} ($${Math.round(topVal.total).toLocaleString('es-CO')}) [${topVal.emp}]`);
    console.log(`- Top Mayor Antigüedad: ${topAge.doc} - ${topAge.company} (${topAge.age} días) [${topAge.emp}]`);
  }

  const unassigned = allRows.filter((r) => !r.group);
  const unCount = unassigned.length;
  const unVal = unassigned.reduce((s, r) => s + r.total, 0);
  const unAvgAge = unCount > 0 ? Math.round(unassigned.reduce((s, r) => s + r.age, 0) / unCount) : 0;
  const unPct = Number(((unVal / grandTotal) * 100).toFixed(1));
  console.log(`\n======================================================`);
  console.log(`CUENTAS ESPECIALES / SIN ASIGNAR`);
  console.log(`- Remisiones: ${unCount}`);
  console.log(`- Valor: $${Math.round(unVal).toLocaleString('es-CO')}`);
  console.log(`- % Participación: ${unPct}%`);
  console.log(`- Antigüedad Promedio: ${unAvgAge} días`);
}

run();
