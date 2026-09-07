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

const corporateWorkbookPath = path.resolve(process.cwd(), 'Remisiones.xlsx');
const hasRealWorkbook = fs.existsSync(corporateWorkbookPath);

// Lista representativa de nombres de directores y empleados en Excel (para ejecución en CI / GitHub Actions)
const excelDirectorsMock = ['Rafael Novoa', 'Angélica Caballero', 'Óscar Beltrán', 'Miller Romero'];
const excelEmployeesMock = [
  'Rosmira Rojas Puente',
  'Mario Reyes Gutierre',
  'Wilson Fernando Sanc',
  'Maria Eugenia Cruz H',
  'Javier Antonio Corte',
  'Rosa Maria Mendoza M',
  'Mariela Ramírez Cast',
  'Jenny Alexandra Gonz',
  'Julieth Milena Galin',
  'Yurany Andrea Vargas',
  'Maria Alejandra Velá',
  'Fernando Alberto Qui',
  'Jasbleidy Johana Moj',
  'Johanna Jaime Murcia',
  'Dayana Marcela Chala',
  'Jair Yovanny Herrea',
  'Cesar Augusto Cesped',
  'Daniel Galindo Giron',
  'Adriana Cucaita Boni',
  'Gina Paola Garcia Qu',
  'Karent Carrillo Mari',
  'Lington Linares Lina',
  'Maria Angelica Alvar',
  'Freddy Andres Peña S',
  'Angie Tatiana Parra',
  'Claudia Patricia Tri',
  'Dilma Constanza Cues',
  'Juan David Martínez',
  'Deisy Mogollon',
  'Leidy Astrid Jimenez',
  'Maria Paola Briceño',
  'Dafne Lizeth Ruiz Be',
  'Jessica Lorena Valen',
  'Jhonatan Steven Acev',
  'Jhonatan Camilo Hern',
  'Yeison Alonso Urrego',
  'Diana Catalina Castr',
];

describe('Sistema de permisos y control de acceso corporativo (RBAC)', () => {
  it('valida que el directorio contenga 4 miembros de gerencia, 3 directores de grupo y 38 ejecutivos', () => {
    const gerencia = CORPORATE_DIRECTORY.filter((u) => u.role === 'admin');
    const directors = CORPORATE_DIRECTORY.filter((u) => u.role === 'director');
    const executives = CORPORATE_DIRECTORY.filter((u) => u.role === 'executive');
    expect(gerencia.length).toBe(4);
    expect(directors.length).toBe(3);
    expect(executives.length).toBe(38);
    expect(CORPORATE_DIRECTORY.length).toBe(45);
  });

  it('resuelve correctamente a los miembros de Gerencia con Acceso Total', () => {
    // 1. Rafael Novoa (Gerencia / Dirección Comercial)
    const novoa = resolveUserAccess('rafael.novoa@provexpress.com.co', excelDirectorsMock, excelEmployeesMock);
    expect(novoa.role).toBe('admin');
    expect(novoa.isRestricted).toBe(false);
    expect(novoa.lockedDirector).toBeUndefined();
    expect(novoa.label).toContain('Gerencia');

    // 2. Juan Novoa (Gerencia General)
    const juan = resolveUserAccess('juannovoa@provexpress.com.co', excelDirectorsMock, excelEmployeesMock);
    expect(juan.role).toBe('admin');
    expect(juan.isRestricted).toBe(false);

    // 3. Cuentas Estratégicas (Gerencia)
    const estrategica = resolveUserAccess('c.estrategica@provexpress.com.co', excelDirectorsMock, excelEmployeesMock);
    expect(estrategica.role).toBe('admin');
    expect(estrategica.isRestricted).toBe(false);

    // 4. Preventa Software (Gerencia)
    const preventa = resolveUserAccess('preventa.software@provexpress.com.co', excelDirectorsMock, excelEmployeesMock);
    expect(preventa.role).toBe('admin');
    expect(preventa.isRestricted).toBe(false);
  });

  it('resuelve correctamente a los directores de grupo bloqueando su dirección', () => {
    // 1. Angélica Caballero
    const caballero = resolveUserAccess('angelica.caballero@provexpress.com.co', excelDirectorsMock, excelEmployeesMock);
    expect(caballero.role).toBe('director');
    expect(caballero.group).toBe(2);
    expect(caballero.lockedDirector).toBe('Angélica Caballero');
    expect(caballero.isRestricted).toBe(true);
    expect(caballero.allowedEmployees.length).toBeGreaterThan(0);

    // 2. Óscar Beltrán
    const beltran = resolveUserAccess('oscar.beltran@provexpress.com.co', excelDirectorsMock, excelEmployeesMock);
    expect(beltran.role).toBe('director');
    expect(beltran.group).toBe(3);
    expect(beltran.lockedDirector).toBe('Óscar Beltrán');
    expect(beltran.isRestricted).toBe(true);

    // 3. Miller Romero
    const romero = resolveUserAccess('miller.romero@provexpress.com.co', excelDirectorsMock, excelEmployeesMock);
    expect(romero.role).toBe('director');
    expect(romero.group).toBe(4);
    expect(romero.lockedDirector).toBe('Miller Romero');
    expect(romero.isRestricted).toBe(true);
  });

  it('resuelve correctamente ejecutivos bloqueando a su propio nombre y dirección', () => {
    // Rosmira Rojas (Grupo Novoa)
    const rosmira = resolveUserAccess('rosmira.rojas@provexpress.com.co', excelDirectorsMock, excelEmployeesMock);
    expect(rosmira.role).toBe('executive');
    expect(rosmira.lockedDirector).toBe('Rafael Novoa');
    expect(rosmira.lockedEmployee).toBe('Rosmira Rojas Puente');
    expect(rosmira.isRestricted).toBe(true);
    expect(rosmira.category).toBe('Master');

    // Yurany Andrea Vargas (Grupo Caballero)
    const andrea = resolveUserAccess('andrea.vargas@provexpress.com.co', excelDirectorsMock, excelEmployeesMock);
    expect(andrea.role).toBe('executive');
    expect(andrea.lockedDirector).toBe('Angélica Caballero');
    expect(andrea.lockedEmployee).toBe('Yurany Andrea Vargas');
    expect(andrea.isRestricted).toBe(true);

    // Camilo Hernández (Grupo Romero)
    const camilo = resolveUserAccess('camilo.hernandez@provexpress.com.co', excelDirectorsMock, excelEmployeesMock);
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

const describeCorporateWorkbook = hasRealWorkbook ? describe : describe.skip;

describeCorporateWorkbook('validación de permisos contra libro Excel real', () => {
  it('resuelve directores y ejecutivos contra el archivo real Remisiones.xlsx', async () => {
    const buffer = fs.readFileSync(corporateWorkbookPath);
    const arrayBuffer = buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength) as ArrayBuffer;
    const wb = await parseRemisionesWorkbook(arrayBuffer);
    const realDirectors = [...new Set(wb.records.map((r) => r.director))];
    const realEmployees = [...new Set(wb.records.map((r) => r.employee))];

    const rosmira = resolveUserAccess('rosmira.rojas@provexpress.com.co', realDirectors, realEmployees);
    expect(rosmira.lockedEmployee).toBe('Rosmira Rojas Puente');

    const caballero = resolveUserAccess('angelica.caballero@provexpress.com.co', realDirectors, realEmployees);
    expect(caballero.lockedDirector).toBe('Angélica Caballero');
  });
});
