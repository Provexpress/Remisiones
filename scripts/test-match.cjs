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

async function testMatch() {
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
      // Nota: No incluir al director como miembro a menos que se quiera emparejar
      return;
    }
    if (!text || /^Ejecutivo Comercial$/i.test(text) || !currentGroup) return;
    groups.push({ group: currentGroup, director: currentDirector, member: text });
  });

  const sis = wb.getWorksheet('Base-SIS');
  const res = {
    1: { count: 0, total: 0, execs: new Map() },
    2: { count: 0, total: 0, execs: new Map() },
    3: { count: 0, total: 0, execs: new Map() },
    4: { count: 0, total: 0, execs: new Map() },
    unmatched: { count: 0, total: 0, list: [] }
  };

  for (let r = 2; r <= sis.rowCount; r++) {
    const row = sis.getRow(r);
    const dVal = row.getCell(1).value;
    const iso = dVal instanceof Date ? dVal.toISOString().slice(0, 10) : String(dVal || '').trim();
    if (iso === '2026-09-16') {
      const emp = String(row.getCell(2).value || '').trim();
      const tot = Number(row.getCell(7).value || 0);
      const m = matchGroup(emp, groups);
      if (m) {
        res[m.group].count++;
        res[m.group].total += tot;
        const curr = res[m.group].execs.get(m.member) || { count: 0, total: 0, rawEmps: new Set() };
        curr.count++;
        curr.total += tot;
        curr.rawEmps.add(emp);
        res[m.group].execs.set(m.member, curr);
      } else {
        res.unmatched.count++;
        res.unmatched.total += tot;
        res.unmatched.list.push({ emp, tot });
      }
    }
  }

  for (let g = 1; g <= 4; g++) {
    console.log(`\n=== OFICIAL GRUPO ${g} ===`);
    console.log(`Remisiones: ${res[g].count}`);
    console.log(`Total: $${Math.round(res[g].total).toLocaleString('es-CO')}`);
    console.log(`Asesores activos (${res[g].execs.size}):`);
    for (const [member, d] of res[g].execs.entries()) {
      console.log(`  - ${member}: ${d.count} remisiones, $${Math.round(d.total).toLocaleString('es-CO')} [de ${Array.from(d.rawEmps).join(', ')}]`);
    }
  }

  console.log(`\n=== NO ASIGNADOS ===`);
  console.log(`Total: ${res.unmatched.count} remisiones, $${Math.round(res.unmatched.total).toLocaleString('es-CO')}`);
  const unMap = new Map();
  res.unmatched.list.forEach(x => unMap.set(x.emp, (unMap.get(x.emp) || { count: 0, total: 0 })));
  res.unmatched.list.forEach(x => {
    const c = unMap.get(x.emp);
    c.count++;
    c.total += x.tot;
  });
  console.log('Emps no asignados:');
  for (const [emp, d] of unMap.entries()) {
    console.log(`  - "${emp}": ${d.count} remisiones, $${Math.round(d.total).toLocaleString('es-CO')}`);
  }
}

testMatch();
