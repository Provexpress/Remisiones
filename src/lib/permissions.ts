import { normalizeText, matchGroup } from './remisiones';

export type UserRole = 'admin' | 'director' | 'executive';

export interface DirectoryUser {
  email: string;
  name: string;
  role: UserRole;
  group: number;
  groupName: string;
  directorName: string;
  category?: 'Master' | 'Junior' | 'Enterprise';
  monthlyQuota?: number;
  aliases?: string[];
}

export interface UserAccess {
  email: string;
  name: string;
  role: UserRole;
  group?: number;
  groupName?: string;
  directorName?: string;
  category?: string;
  monthlyQuota?: number;
  /**
   * If restricted, only records belonging to this director are allowed.
   */
  lockedDirector?: string;
  /**
   * If restricted to an executive, only records matching their Excel employee name are allowed.
   */
  lockedEmployee?: string;
  /**
   * Names of employees the user is permitted to see. If empty, all employees under lockedDirector or all in company.
   */
  allowedEmployees: string[];
  isRestricted: boolean;
  label: string;
}

export const CORPORATE_DIRECTORY: DirectoryUser[] = [
  // ══ DIRECTORES ═════════════════════════════════════════════════════════════
  {
    email: 'rafael.novoa@provexpress.com.co',
    name: 'Rafael Novoa',
    role: 'director',
    group: 1,
    groupName: 'Grupo Novoa',
    directorName: 'Rafael Novoa',
    aliases: ['Rafael Francisco Nov', 'Rafael Novoa'],
  },
  {
    email: 'angelica.caballero@provexpress.com.co',
    name: 'Angélica Caballero',
    role: 'director',
    group: 2,
    groupName: 'Grupo Caballero',
    directorName: 'Angélica Caballero',
    aliases: ['Maria Angelica Cabal', 'Angelica Caballero'],
  },
  {
    email: 'oscar.beltran@provexpress.com.co',
    name: 'Óscar Beltrán',
    role: 'director',
    group: 3,
    groupName: 'Grupo Beltrán',
    directorName: 'Óscar Beltrán',
    aliases: ['Oscar Alejandro Belt', 'Oscar Beltran'],
  },
  {
    email: 'miller.romero@provexpress.com.co',
    name: 'Miller Romero',
    role: 'director',
    group: 4,
    groupName: 'Grupo Romero',
    directorName: 'Miller Romero',
    aliases: ['Miller Leonardo Rome', 'Miller Romero'],
  },

  // ══ GRUPO 1: NOVOA (9 Ejecutivos) ══════════════════════════════════════════
  {
    email: 'rosmira.rojas@provexpress.com.co',
    name: 'Rosmira Rojas',
    role: 'executive',
    group: 1,
    groupName: 'Grupo Novoa',
    directorName: 'Rafael Novoa',
    category: 'Master',
    monthlyQuota: 28_000_000,
    aliases: ['Rosmira Rojas Puente'],
  },
  {
    email: 'mario.reyes@provexpress.com.co',
    name: 'Mario Reyes',
    role: 'executive',
    group: 1,
    groupName: 'Grupo Novoa',
    directorName: 'Rafael Novoa',
    category: 'Master',
    monthlyQuota: 28_000_000,
    aliases: ['Mario Reyes Gutierre'],
  },
  {
    email: 'wilson.sanchez@provexpress.com.co',
    name: 'Wilson Sánchez',
    role: 'executive',
    group: 1,
    groupName: 'Grupo Novoa',
    directorName: 'Rafael Novoa',
    category: 'Master',
    monthlyQuota: 28_000_000,
    aliases: ['Wilson Fernando Sanc'],
  },
  {
    email: 'maria.cruz@provexpress.com.co',
    name: 'María Eugenia Cruz',
    role: 'executive',
    group: 1,
    groupName: 'Grupo Novoa',
    directorName: 'Rafael Novoa',
    category: 'Master',
    monthlyQuota: 28_000_000,
    aliases: ['Maria Eugenia Cruz H'],
  },
  {
    email: 'javier.cortes@provexpress.com.co',
    name: 'Javier Cortés',
    role: 'executive',
    group: 1,
    groupName: 'Grupo Novoa',
    directorName: 'Rafael Novoa',
    category: 'Master',
    monthlyQuota: 28_000_000,
    aliases: ['Javier Antonio Corte'],
  },
  {
    email: 'rosa.mendoza@provexpress.com.co',
    name: 'Rosa Mendoza',
    role: 'executive',
    group: 1,
    groupName: 'Grupo Novoa',
    directorName: 'Rafael Novoa',
    category: 'Master',
    monthlyQuota: 28_000_000,
    aliases: ['Rosa Maria Mendoza M'],
  },
  {
    email: 'mariela.ramirez@provexpress.com.co',
    name: 'Mariela Ramírez',
    role: 'executive',
    group: 1,
    groupName: 'Grupo Novoa',
    directorName: 'Rafael Novoa',
    category: 'Junior',
    monthlyQuota: 18_000_000,
    aliases: ['Mariela Ramírez Cast'],
  },
  {
    email: 'jenny.gonzalez@provexpress.com.co',
    name: 'Jenny Gónzalez',
    role: 'executive',
    group: 1,
    groupName: 'Grupo Novoa',
    directorName: 'Rafael Novoa',
    category: 'Junior',
    monthlyQuota: 18_000_000,
    aliases: ['Jenny Alexandra Gonz'],
  },
  {
    email: 'julieth.galindo@provexpress.com.co',
    name: 'Julieth Galindo',
    role: 'executive',
    group: 1,
    groupName: 'Grupo Novoa',
    directorName: 'Rafael Novoa',
    category: 'Junior',
    monthlyQuota: 18_000_000,
    aliases: ['Julieth Milena Galin'],
  },

  // ══ GRUPO 2: CABALLERO (11 Ejecutivos) ═════════════════════════════════════
  {
    email: 'angela.torres@provexpress.com.co',
    name: 'Ángela Torres',
    role: 'executive',
    group: 2,
    groupName: 'Grupo Caballero',
    directorName: 'Angélica Caballero',
    category: 'Junior',
    monthlyQuota: 18_000_000,
    aliases: ['Angela Rocio Torres'],
  },
  {
    email: 'andrea.vargas@provexpress.com.co',
    name: 'Yurany Andrea Vargas',
    role: 'executive',
    group: 2,
    groupName: 'Grupo Caballero',
    directorName: 'Angélica Caballero',
    category: 'Master',
    monthlyQuota: 28_000_000,
    aliases: ['Yurany Andrea Vargas'],
  },
  {
    email: 'alejandra.velasquez@provexpress.com.co',
    name: 'Alejandra Velásquez',
    role: 'executive',
    group: 2,
    groupName: 'Grupo Caballero',
    directorName: 'Angélica Caballero',
    category: 'Master',
    monthlyQuota: 28_000_000,
    aliases: ['Maria Alejandra Velá'],
  },
  {
    email: 'fernando.quinonez@provexpress.com.co',
    name: 'Fernando Quiñonez',
    role: 'executive',
    group: 2,
    groupName: 'Grupo Caballero',
    directorName: 'Angélica Caballero',
    category: 'Master',
    monthlyQuota: 28_000_000,
    aliases: ['Fernando Alberto Qui'],
  },
  {
    email: 'johana.mojica@provexpress.com.co',
    name: 'Jasbleidy Mójica',
    role: 'executive',
    group: 2,
    groupName: 'Grupo Caballero',
    directorName: 'Angélica Caballero',
    category: 'Master',
    monthlyQuota: 28_000_000,
    aliases: ['Jasbleidy Johana Moj'],
  },
  {
    email: 'johanna.jaime@provexpress.com.co',
    name: 'Johanna Jaime',
    role: 'executive',
    group: 2,
    groupName: 'Grupo Caballero',
    directorName: 'Angélica Caballero',
    category: 'Master',
    monthlyQuota: 28_000_000,
    aliases: ['Johanna Jaime Murcia'],
  },
  {
    email: 'dayana.chala@provexpress.com.co',
    name: 'Dayana Chala',
    role: 'executive',
    group: 2,
    groupName: 'Grupo Caballero',
    directorName: 'Angélica Caballero',
    category: 'Junior',
    monthlyQuota: 18_000_000,
    aliases: ['Dayana Marcela Chala'],
  },
  {
    email: 'yovanny.herrera@provexpress.com.co',
    name: 'Yovanny Herrera',
    role: 'executive',
    group: 2,
    groupName: 'Grupo Caballero',
    directorName: 'Angélica Caballero',
    category: 'Junior',
    monthlyQuota: 18_000_000,
    aliases: ['Jair Yovanny Herrea'],
  },
  {
    email: 'cesar.cespedes@provexpress.com.co',
    name: 'César Céspedes',
    role: 'executive',
    group: 2,
    groupName: 'Grupo Caballero',
    directorName: 'Angélica Caballero',
    category: 'Master',
    monthlyQuota: 28_000_000,
    aliases: ['Cesar Augusto Cesped'],
  },
  {
    email: 'daniel.galindo@provexpress.com.co',
    name: 'Daniel Galindo',
    role: 'executive',
    group: 2,
    groupName: 'Grupo Caballero',
    directorName: 'Angélica Caballero',
    category: 'Master',
    monthlyQuota: 28_000_000,
    aliases: ['Daniel Galindo Giron'],
  },
  {
    email: 'adriana.cucaita@provexpress.com.co',
    name: 'Adriana Cucaita',
    role: 'executive',
    group: 2,
    groupName: 'Grupo Caballero',
    directorName: 'Angélica Caballero',
    category: 'Junior',
    monthlyQuota: 18_000_000,
    aliases: ['Adriana Cucaita Boni'],
  },

  // ══ GRUPO 3: BELTRÁN (10 Ejecutivos) ═══════════════════════════════════════
  {
    email: 'paola.garcia@provexpress.com.co',
    name: 'Gina García',
    role: 'executive',
    group: 3,
    groupName: 'Grupo Beltrán',
    directorName: 'Óscar Beltrán',
    category: 'Master',
    monthlyQuota: 28_000_000,
    aliases: ['Gina Paola Garcia Qu', 'Gina Garcia'],
  },
  {
    email: 'karen.carrillo@provexpress.com.co',
    name: 'Karent Carrillo',
    role: 'executive',
    group: 3,
    groupName: 'Grupo Beltrán',
    directorName: 'Óscar Beltrán',
    category: 'Master',
    monthlyQuota: 28_000_000,
    aliases: ['Karent Carrillo Mari'],
  },
  {
    email: 'lington.linares@provexpress.com.co',
    name: 'Lington Linares',
    role: 'executive',
    group: 3,
    groupName: 'Grupo Beltrán',
    directorName: 'Óscar Beltrán',
    category: 'Master',
    monthlyQuota: 28_000_000,
    aliases: ['Lington Linares Lina'],
  },
  {
    email: 'angelica.alvarez@provexpress.com.co',
    name: 'Angélica Álvarez',
    role: 'executive',
    group: 3,
    groupName: 'Grupo Beltrán',
    directorName: 'Óscar Beltrán',
    category: 'Junior',
    monthlyQuota: 18_000_000,
    aliases: ['Maria Angelica Alvar'],
  },
  {
    email: 'andres.pena@provexpress.com.co',
    name: 'Andrés Peña',
    role: 'executive',
    group: 3,
    groupName: 'Grupo Beltrán',
    directorName: 'Óscar Beltrán',
    category: 'Master',
    monthlyQuota: 28_000_000,
    aliases: ['Freddy Andres Peña S'],
  },
  {
    email: 'tatiana.parra@provexpress.com.co',
    name: 'Tatiana Parra',
    role: 'executive',
    group: 3,
    groupName: 'Grupo Beltrán',
    directorName: 'Óscar Beltrán',
    category: 'Junior',
    monthlyQuota: 18_000_000,
    aliases: ['Angie Tatiana Parra'],
  },
  {
    email: 'claudia.triana@provexpress.com.co',
    name: 'Claudia Triana',
    role: 'executive',
    group: 3,
    groupName: 'Grupo Beltrán',
    directorName: 'Óscar Beltrán',
    category: 'Master',
    monthlyQuota: 28_000_000,
    aliases: ['Claudia Patricia Tri'],
  },
  {
    email: 'dilma.cuesta@provexpress.com.co',
    name: 'Dilma Cuesta',
    role: 'executive',
    group: 3,
    groupName: 'Grupo Beltrán',
    directorName: 'Óscar Beltrán',
    category: 'Junior',
    monthlyQuota: 18_000_000,
    aliases: ['Dilma Constanza Cues'],
  },
  {
    email: 'juan.martinez@provexpress.com.co',
    name: 'Juan Martínez',
    role: 'executive',
    group: 3,
    groupName: 'Grupo Beltrán',
    directorName: 'Óscar Beltrán',
    category: 'Junior',
    monthlyQuota: 14_000_000,
    aliases: ['Juan David Martínez'],
  },
  {
    email: 'deisy.mogollon@provexpress.com.co',
    name: 'Deisy Mogollón',
    role: 'executive',
    group: 3,
    groupName: 'Grupo Beltrán',
    directorName: 'Óscar Beltrán',
    category: 'Junior',
    monthlyQuota: 18_000_000,
    aliases: ['Deisy Mogollon'],
  },

  // ══ GRUPO 4: ROMERO (8 Ejecutivos) ═════════════════════════════════════════
  {
    email: 'astrid.jimenez@provexpress.com.co',
    name: 'Astrid Jiménez',
    role: 'executive',
    group: 4,
    groupName: 'Grupo Romero',
    directorName: 'Miller Romero',
    category: 'Enterprise',
    monthlyQuota: 48_000_000,
    aliases: ['Leidy Astrid Jimenez'],
  },
  {
    email: 'maria.briceno@provexpress.com.co',
    name: 'María Paola Briceño',
    role: 'executive',
    group: 4,
    groupName: 'Grupo Romero',
    directorName: 'Miller Romero',
    category: 'Enterprise',
    monthlyQuota: 48_000_000,
    aliases: ['Maria Paola Briceño'],
  },
  {
    email: 'dafne.ruiz@provexpress.com.co',
    name: 'Dafne Ruiz',
    role: 'executive',
    group: 4,
    groupName: 'Grupo Romero',
    directorName: 'Miller Romero',
    category: 'Enterprise',
    monthlyQuota: 48_000_000,
    aliases: ['Dafne Lizeth Ruiz Be'],
  },
  {
    email: 'jessica.valencia@provexpress.com.co',
    name: 'Jessica Valencia',
    role: 'executive',
    group: 4,
    groupName: 'Grupo Romero',
    directorName: 'Miller Romero',
    category: 'Enterprise',
    monthlyQuota: 48_000_000,
    aliases: ['Jessica Lorena Valen'],
  },
  {
    email: 'jhonatan.acevedo@provexpress.com.co',
    name: 'Jhonatan Acevedo',
    role: 'executive',
    group: 4,
    groupName: 'Grupo Romero',
    directorName: 'Miller Romero',
    category: 'Enterprise',
    monthlyQuota: 48_000_000,
    aliases: ['Jhonatan Steven Acev'],
  },
  {
    email: 'camilo.hernandez@provexpress.com.co',
    name: 'Camilo Hernández',
    role: 'executive',
    group: 4,
    groupName: 'Grupo Romero',
    directorName: 'Miller Romero',
    category: 'Enterprise',
    monthlyQuota: 48_000_000,
    aliases: ['Jhonatan Camilo Hern'],
  },
  {
    email: 'yeison.urrego@provexpress.com.co',
    name: 'Yeison Urrego',
    role: 'executive',
    group: 4,
    groupName: 'Grupo Romero',
    directorName: 'Miller Romero',
    category: 'Enterprise',
    monthlyQuota: 48_000_000,
    aliases: ['Yeison Alonso Urrego'],
  },
  {
    email: 'diana.castro@provexpress.com.co',
    name: 'Diana Castro',
    role: 'executive',
    group: 4,
    groupName: 'Grupo Romero',
    directorName: 'Miller Romero',
    category: 'Enterprise',
    monthlyQuota: 48_000_000,
    aliases: ['Diana Catalina Castr'],
  },
];

/**
 * Finds a matching employee name in Excel for a directory user using aliases or fuzzy matching.
 */
export function matchExcelEmployee(user: DirectoryUser, excelEmployees: string[]): string | undefined {
  if (user.aliases) {
    for (const alias of user.aliases) {
      const exact = excelEmployees.find((e) => normalizeText(e) === normalizeText(alias));
      if (exact) return exact;
    }
  }
  // Try matching against name
  const match = excelEmployees.find((e) => {
    const res = matchGroup(e, [{ group: user.group, director: user.directorName, member: user.name }]);
    return !!res;
  });
  return match;
}

/**
 * Finds all employees belonging to a specific direction/group.
 */
export function getEmployeesForDirector(directorName: string, excelEmployees: string[]): string[] {
  const normDirector = normalizeText(directorName);
  const groupUsers = CORPORATE_DIRECTORY.filter(
    (u) => u.role === 'executive' && normalizeText(u.directorName) === normDirector,
  );
  const set = new Set<string>();
  for (const u of groupUsers) {
    const matched = matchExcelEmployee(u, excelEmployees);
    if (matched) set.add(matched);
  }
  return [...set];
}

/**
 * Resolves permissions and access control for a given user email against actual Excel records.
 */
export function resolveUserAccess(
  email: string | null | undefined,
  excelDirectors: string[] = [],
  excelEmployees: string[] = [],
): UserAccess {
  if (!email) {
    return {
      email: '',
      name: 'Vista Local',
      role: 'admin',
      allowedEmployees: [],
      isRestricted: false,
      label: 'Acceso Total (Vista Local)',
    };
  }

  const cleanEmail = email.trim().toLowerCase();
  const found = CORPORATE_DIRECTORY.find((u) => u.email.toLowerCase() === cleanEmail);

  if (!found) {
    // Admin or corporate user without strict restrictions
    return {
      email: cleanEmail,
      name: email.split('@')[0],
      role: 'admin',
      allowedEmployees: [],
      isRestricted: false,
      label: 'Dirección General / Administrador',
    };
  }

  // 1. Director role: can only view their own group
  if (found.role === 'director') {
    // Match director name in Excel
    const normDir = normalizeText(found.directorName);
    const matchedDirector = excelDirectors.find((d) => normalizeText(d) === normDir) || found.directorName;
    const groupEmployees = getEmployeesForDirector(found.directorName, excelEmployees);

    return {
      email: found.email,
      name: found.name,
      role: 'director',
      group: found.group,
      groupName: found.groupName,
      directorName: matchedDirector,
      lockedDirector: matchedDirector,
      allowedEmployees: groupEmployees,
      isRestricted: true,
      label: `Director · ${found.groupName}`,
    };
  }

  // 2. Executive role: can only view their own assigned remisiones
  if (found.role === 'executive') {
    const normDir = normalizeText(found.directorName);
    const matchedDirector = excelDirectors.find((d) => normalizeText(d) === normDir) || found.directorName;
    const matchedEmployee = matchExcelEmployee(found, excelEmployees) || found.name;

    return {
      email: found.email,
      name: found.name,
      role: 'executive',
      group: found.group,
      groupName: found.groupName,
      directorName: matchedDirector,
      lockedDirector: matchedDirector,
      lockedEmployee: matchedEmployee,
      allowedEmployees: [matchedEmployee],
      category: found.category,
      monthlyQuota: found.monthlyQuota,
      isRestricted: true,
      label: `Ejecutivo ${found.category || ''} · ${found.groupName}`,
    };
  }

  return {
    email: cleanEmail,
    name: found.name,
    role: 'admin',
    allowedEmployees: [],
    isRestricted: false,
    label: 'Administrador General',
  };
}
