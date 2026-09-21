import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { describe, it } from 'vitest';
import { parseRemisionesWorkbook } from './remisiones';
import { ESTRUCTURA_COMERCIAL_2026, LISTA_DIRECTORES } from './commercialDirectory';

describe('check groups on 2026-09-16', () => {
  it('analyzes parsed records for 2026-09-16', async () => {
    const corporateWorkbookPath = fileURLToPath(new URL('../../Remisiones.xlsx', import.meta.url));
    const file = await readFile(corporateWorkbookPath);
    const arrayBuffer = file.buffer.slice(file.byteOffset, file.byteOffset + file.byteLength) as ArrayBuffer;
    const parsed = await parseRemisionesWorkbook(arrayBuffer);
    const records16 = parsed.records.filter((r) => r.cutoff === '2026-09-16');

    console.log('\n======================================================');
    console.log(`TOTAL RECORDS 2026-09-16: ${records16.length}`);
    const totalCompanyValue = records16.reduce((s, r) => s + r.total, 0);
    console.log(`TOTAL COMPANY VALUE: $${Math.round(totalCompanyValue).toLocaleString('es-CO')}`);

    // Group distribution
    for (let g = 1; g <= 4; g++) {
      const gRecords = records16.filter((r) => r.group === g);
      const gTotal = gRecords.reduce((s, r) => s + r.total, 0);
      const uniqueEmployees = [...new Set(gRecords.map((r) => r.employee))];
      console.log(`\n--- GRUPO ${g} ---`);
      console.log(`Remisiones: ${gRecords.length}`);
      console.log(`Valor Total: $${Math.round(gTotal).toLocaleString('es-CO')}`);
      console.log(`Empleados en el grupo (${uniqueEmployees.length}):`, uniqueEmployees);
    }

    const unassigned = records16.filter((r) => !r.group);
    const unassignedTotal = unassigned.reduce((s, r) => s + r.total, 0);
    const unassignedEmps = [...new Set(unassigned.map((r) => r.employee))];
    console.log(`\n--- SIN ASIGNAR (${unassigned.length} remisiones, $${Math.round(unassignedTotal).toLocaleString('es-CO')}) ---`);
    console.log('Empleados sin grupo:', unassignedEmps);
    console.log('======================================================\n');
  }, 30000);
});
