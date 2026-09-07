import { describe, it, expect } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { parseRemisionesWorkbook } from './remisiones';
import {
  CORPORATE_DIRECTORY,
  resolveUserAccess,
  matchExcelEmployee,
  getEmployeesForDirector,
} from './permissions';

describe('Sistema de permisos y control de acceso corporativo', () => {
  const buffer = fs.readFileSync(path.resolve(process.cwd(), 'Remisiones.xlsx'));

  it('valida que el directorio contenga 4 directores y 38 ejecutivos', () => {
    const directors = CORPORATE_DIRECTORY.filter((u) => u.role === 'director');
    const executives = CORPORATE_DIRECTORY.filter((u) => u.role === 'executive');
    expect(directors.length).toBe(4);
    expect(executives.length).toBe(38);
    expect(CORPORATE_DIRECTORY.length).toBe(42);
  });

  it('resuelve correctamente a los 4 directores bloqueando su dirección', async () => {
    const wb = await parseRemisionesWorkbook(buffer.buffer);
    const excelDirectors = [...new Set(wb.records.map((r) => r.director))];
    const excelEmployees = [...new Set(wb.records.map((r) => r.employee))];

    // 1. Rafael Novoa
    const novoa = resolveUserAccess('rafael.novoa@provexpress.com.co', excelDirectors, excelEmployees);
    expect(novoa.role).toBe('director');
    expect(novoa.group).toBe(1);
    expect(novoa.lockedDirector).toBe('Rafael Novoa');
    expect(novoa.isRestricted).toBe(true);
    expect(novoa.allowedEmployees.length).toBeGreaterThan(0);

    // 2. Angélica Caballero
    const caballero = resolveUserAccess('angelica.caballero@provexpress.com.co', excelDirectors, excelEmployees);
    expect(caballero.role).toBe('director');
    expect(caballero.group).toBe(2);
    expect(caballero.lockedDirector).toBe('Angélica Caballero');
    expect(caballero.isRestricted).toBe(true);
    expect(caballero.allowedEmployees.length).toBeGreaterThan(0);

    // 3. Óscar Beltrán
    const beltran = resolveUserAccess('oscar.beltran@provexpress.com.co', excelDirectors, excelEmployees);
    expect(beltran.role).toBe('director');
    expect(beltran.group).toBe(3);
    expect(beltran.lockedDirector).toBe('Óscar Beltrán');
    expect(beltran.isRestricted).toBe(true);

    // 4. Miller Romero
    const romero = resolveUserAccess('miller.romero@provexpress.com.co', excelDirectors, excelEmployees);
    expect(romero.role).toBe('director');
    expect(romero.group).toBe(4);
    expect(romero.lockedDirector).toBe('Miller Romero');
    expect(romero.isRestricted).toBe(true);
  });

  it('resuelve correctamente ejecutivos bloqueando a su propio nombre y dirección', async () => {
    const wb = await parseRemisionesWorkbook(buffer.buffer);
    const excelDirectors = [...new Set(wb.records.map((r) => r.director))];
    const excelEmployees = [...new Set(wb.records.map((r) => r.employee))];

    // Rosmira Rojas (Grupo Novoa)
    const rosmira = resolveUserAccess('rosmira.rojas@provexpress.com.co', excelDirectors, excelEmployees);
    expect(rosmira.role).toBe('executive');
    expect(rosmira.lockedDirector).toBe('Rafael Novoa');
    expect(rosmira.lockedEmployee).toBe('Rosmira Rojas Puente');
    expect(rosmira.isRestricted).toBe(true);
    expect(rosmira.category).toBe('Master');

    // Yurany Andrea Vargas (Grupo Caballero)
    const andrea = resolveUserAccess('andrea.vargas@provexpress.com.co', excelDirectors, excelEmployees);
    expect(andrea.role).toBe('executive');
    expect(andrea.lockedDirector).toBe('Angélica Caballero');
    expect(andrea.lockedEmployee).toBe('Yurany Andrea Vargas');
    expect(andrea.isRestricted).toBe(true);

    // Camilo Hernández (Grupo Romero)
    const camilo = resolveUserAccess('camilo.hernandez@provexpress.com.co', excelDirectors, excelEmployees);
    expect(camilo.role).toBe('executive');
    expect(camilo.lockedDirector).toBe('Miller Romero');
    expect(camilo.lockedEmployee).toBe('Jhonatan Camilo Hern');
    expect(camilo.isRestricted).toBe(true);
    expect(camilo.category).toBe('Enterprise');
  });

  it('asigna rol de administrador con acceso total para cuentas no restringidas o modo local', () => {
    const admin = resolveUserAccess('especialista.preventa@provexpress.com');
    expect(admin.role).toBe('admin');
    expect(admin.isRestricted).toBe(false);
    expect(admin.lockedDirector).toBeUndefined();
    expect(admin.lockedEmployee).toBeUndefined();

    const local = resolveUserAccess(null);
    expect(local.role).toBe('admin');
    expect(local.isRestricted).toBe(false);
  });
});
