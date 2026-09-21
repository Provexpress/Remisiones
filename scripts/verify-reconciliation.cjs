const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

function normalizeName(value) {
  return String(value || '')
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

async function testReconciliation() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(path.join(__dirname, '../Remisiones.xlsx'));
  const sis = wb.getWorksheet('Base-SIS');

  const rows = [];
  for (let r = 2; r <= sis.rowCount; r++) {
    const row = sis.getRow(r);
    const dVal = row.getCell(1).value;
    const iso = dVal instanceof Date ? dVal.toISOString().slice(0, 10) : String(dVal || '').trim();
    if (iso === '2026-09-16') {
      rows.push({
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

  console.log(`Total remisiones leídas: ${rows.length}`);
  const grandTotal = rows.reduce((s, r) => s + r.total, 0);

  // Clasificar cada fila
  const assigned = [];
  const unassigned = [];

  const execEntries = Object.entries(ESTRUCTURA_COMERCIAL_2026.ejecutivos);
  const dirEntries = Object.entries(ESTRUCTURA_COMERCIAL_2026.directores);

  for (const r of rows) {
    let matched = null;

    // 1. Probar contra ejecutivos
    for (const [email, exec] of execEntries) {
      if (nameMatches(r.emp, exec.nombre) || nameMatches(r.emp, exec.archivo)) {
        matched = {
          ...r,
          email,
          nombreAsesor: exec.nombre,
          grupo: exec.grupo,
          director: ESTRUCTURA_COMERCIAL_2026.directores[exec.grupo].nombre,
          isDirector: false,
        };
        break;
      }
    }

    // 2. Si no es ejecutivo, probar si es remisión directa a nombre del Director
    if (!matched) {
      for (const [gStr, dir] of dirEntries) {
        if (nameMatches(r.emp, dir.nombre)) {
          matched = {
            ...r,
            email: dir.email,
            nombreAsesor: `${dir.nombre} (Gestión Directa)`,
            grupo: Number(gStr),
            director: dir.nombre,
            isDirector: true,
          };
          break;
        }
      }
    }

    if (matched) {
      assigned.push(matched);
    } else {
      unassigned.push(r);
    }
  }

  console.log(`Asignadas a los 4 grupos: ${assigned.length}`);
  console.log(`Sin asignar (Otras áreas): ${unassigned.length}`);

  for (let g = 1; g <= 4; g++) {
    const gRows = assigned.filter((r) => r.grupo === g);
    const count = gRows.length;
    const total = gRows.reduce((s, r) => s + r.total, 0);
    const avgAge = Math.round(gRows.reduce((s, r) => s + r.age, 0) / count);
    const pct = Number(((total / grandTotal) * 100).toFixed(1));

    // Ejecutivos activos
    const uniqueExecs = new Set(gRows.filter((r) => !r.isDirector).map((r) => r.nombreAsesor));
    const totalExecs = g === 1 ? 9 : g === 2 ? 11 : g === 3 ? 10 : 8;

    console.log(`\n=================================================`);
    console.log(`GRUPO ${g} (${ESTRUCTURA_COMERCIAL_2026.directores[g].nombre}):`);
    console.log(`  Remisiones: ${count}`);
    console.log(`  Total COP: $${Math.round(total).toLocaleString('es-CO')}`);
    console.log(`  % Participación: ${pct}%`);
    console.log(`  Antigüedad: ${avgAge} días`);
    console.log(`  Asesores Activos: ${uniqueExecs.size} de ${totalExecs}`);
    console.log(`  Asesores con remisiones:`, Array.from(uniqueExecs));
    const dirRows = gRows.filter((r) => r.isDirector);
    if (dirRows.length > 0) {
      console.log(`  + Remisiones Gestión Directa Director: ${dirRows.length} ($${Math.round(dirRows.reduce((s, r) => s + r.total, 0)).toLocaleString('es-CO')})`);
    }
  }

  console.log(`\n=================================================`);
  console.log(`SIN ASIGNAR: ${unassigned.length} remisiones ($${Math.round(unassigned.reduce((s, r) => s + r.total, 0)).toLocaleString('es-CO')})`);
  const unMap = new Map();
  unassigned.forEach((r) => {
    const c = unMap.get(r.emp) || { count: 0, total: 0 };
    c.count++;
    c.total += r.total;
    unMap.set(r.emp, c);
  });
  for (const [emp, d] of unMap.entries()) {
    console.log(`  - "${emp}": ${d.count} remisiones, $${Math.round(d.total).toLocaleString('es-CO')}`);
  }
}

testReconciliation();
