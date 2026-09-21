const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

function normalizeText(val) {
  return String(val || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

async function analyze() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(path.join(__dirname, '../Remisiones.xlsx'));
  
  const gSheet = wb.getWorksheet('Grupos');
  const groups = [];
  let curGroup = 0, curDir = '';
  gSheet.eachRow((r) => {
    const text = String(r.getCell(1).value || '').trim();
    const m = text.match(/^Grupo\s+(\d+).*Director(?:a)?:\s*(.+)$/i);
    if (m) {
      curGroup = Number(m[1]);
      curDir = m[2].trim();
      return;
    }
    if (!text || /^Ejecutivo Comercial$/i.test(text) || !curGroup) return;
    groups.push({ grupo: curGroup, director: curDir, member: text });
  });

  const sis = wb.getWorksheet('Base-SIS');
  const empMap = new Map();
  let totalRows = 0, totalVal = 0;

  for (let r = 2; r <= sis.rowCount; r++) {
    const row = sis.getRow(r);
    const dVal = row.getCell(1).value;
    const iso = dVal instanceof Date ? dVal.toISOString().slice(0, 10) : String(dVal || '').trim();
    if (iso === '2026-09-16') {
      totalRows++;
      const emp = String(row.getCell(2).value || '').trim();
      const tot = Number(row.getCell(7).value || 0);
      totalVal += tot;

      const curr = empMap.get(emp) || { count: 0, total: 0 };
      curr.count++;
      curr.total += tot;
      empMap.set(emp, curr);
    }
  }

  console.log('Total 2026-09-16:', totalRows, 'remisiones, total COP:', totalVal);
  console.log('Distinct raw employee strings:', empMap.size);

  const groupStats = {
    1: { count: 0, total: 0, matchedEmps: new Set(), execs: new Set() },
    2: { count: 0, total: 0, matchedEmps: new Set(), execs: new Set() },
    3: { count: 0, total: 0, matchedEmps: new Set(), execs: new Set() },
    4: { count: 0, total: 0, matchedEmps: new Set(), execs: new Set() },
    unmatched: { count: 0, total: 0, emps: [] }
  };

  for (const [emp, data] of empMap.entries()) {
    const norm = normalizeText(emp);
    const tokens = norm.split(' ').filter(Boolean);
    let matchedGroup = null;

    // Check direct matching against group members
    for (const g of groups) {
      const gNorm = normalizeText(g.member);
      const gTokens = gNorm.split(' ').filter(Boolean);
      const common = tokens.filter(t => gTokens.includes(t));
      if (common.length >= 2 || (tokens.length === 1 && gTokens.includes(tokens[0]))) {
        matchedGroup = g;
        break;
      }
    }

    if (matchedGroup) {
      groupStats[matchedGroup.grupo].count += data.count;
      groupStats[matchedGroup.grupo].total += data.total;
      groupStats[matchedGroup.grupo].matchedEmps.add(emp);
      groupStats[matchedGroup.grupo].execs.add(matchedGroup.member);
    } else {
      groupStats.unmatched.count += data.count;
      groupStats.unmatched.total += data.total;
      groupStats.unmatched.emps.push({ emp, count: data.count, total: data.total });
    }
  }

  for (let g = 1; g <= 4; g++) {
    console.log(`\n=================== GRUPO ${g} ===================`);
    console.log(`Remisiones: ${groupStats[g].count}`);
    console.log(`Total COP: ${groupStats[g].total}`);
    console.log(`Ejecutivos oficiales activos (${groupStats[g].execs.size}):`, Array.from(groupStats[g].execs));
    console.log(`Variaciones de nombres en Excel (${groupStats[g].matchedEmps.size}):`, Array.from(groupStats[g].matchedEmps));
  }

  console.log('\n=================== NO ASIGNADOS A NINGÚN GRUPO ===================');
  console.log(`Remisiones: ${groupStats.unmatched.count}, Total COP: ${groupStats.unmatched.total}`);
  console.log('Listado no asignados:');
  groupStats.unmatched.emps.forEach(e => console.log(`  - "${e.emp}": ${e.count} remisiones, $${e.total}`));
}

analyze();
