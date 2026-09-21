/**
 * Directorio comercial corporativo de Provexpress SAS.
 * Adaptado e integrado desde la estructura oficial del proyecto ForeCast 2026.
 */

export interface CommercialEntry {
  email: string;
  nombre: string;
  grupo: number;
  directorNombre: string;
  directorEmail: string;
  archivo?: string;
  aliases?: string[];
}

export const ESTRUCTURA_COMERCIAL_2026: {
  gerencia: string[];
  directores: Record<string, { grupo: number; nombre: string; carpeta: string }>;
  ejecutivos: Record<string, { grupo: number; nombre: string; archivo: string }>;
} = {
  gerencia: [
    'juannovoa@provexpress.com.co',
    'oscar.beltran@provexpress.com.co',
    'rafael.novoa@provexpress.com.co',
    'rafaelnovoa@provexpress.com.co',
    'c.estrategica@provexpress.com.co',
    'nini.beltran@provexpress.com.co',
    'maribel.virguez@provexpress.com.co',
    'especialista.preventa@provexpress.com.co',
    'preventa.software@provexpress.com.co',
    'oscar.perez@provexpress.com.co',
  ],
  directores: {
    'rafael.novoa@provexpress.com.co': {
      grupo: 1,
      nombre: 'Rafael Novoa',
      carpeta: 'Grupo Rafael Novoa',
    },
    'rafaelnovoa@provexpress.com.co': {
      grupo: 1,
      nombre: 'Rafael Novoa',
      carpeta: 'Grupo Rafael Novoa',
    },
    'angelica.caballero@provexpress.com.co': {
      grupo: 2,
      nombre: 'Angélica Caballero',
      carpeta: 'Grupo Maria Angelica caballero',
    },
    'oscar.beltran@provexpress.com.co': {
      grupo: 3,
      nombre: 'Óscar Beltrán',
      carpeta: 'Grupo Oscar Beltran',
    },
    'miller.romero@provexpress.com.co': {
      grupo: 4,
      nombre: 'Miller Romero',
      carpeta: 'Grupo Miller Romero',
    },
  },
  ejecutivos: {
    // Grupo 1: Novoa
    'rosmira.rojas@provexpress.com.co': { grupo: 1, nombre: 'Rosmira Rojas', archivo: 'Rosmira Rojas.xlsx' },
    'mario.reyes@provexpress.com.co': { grupo: 1, nombre: 'Mario Reyes', archivo: 'Mario Reyes.xlsx' },
    'wilson.sanchez@provexpress.com.co': { grupo: 1, nombre: 'Wilson Sánchez', archivo: 'Wilson Fernando Sánchez.xlsx' },
    'maria.cruz@provexpress.com.co': { grupo: 1, nombre: 'María Eugenia Cruz', archivo: 'Maria Eugenia Cruz.xlsx' },
    'javier.cortes@provexpress.com.co': { grupo: 1, nombre: 'Javier Cortés', archivo: 'Javier Cortés.xlsx' },
    'rosa.mendoza@provexpress.com.co': { grupo: 1, nombre: 'Rosa Mendoza', archivo: 'Rosa María Mendoza.xlsx' },
    'mariela.ramirez@provexpress.com.co': { grupo: 1, nombre: 'Mariela Ramírez', archivo: 'Mariela Ramírez.xlsx' },
    'jenny.gonzalez@provexpress.com.co': { grupo: 1, nombre: 'Jenny Gónzalez', archivo: 'Jenny Gónzalez.xlsx' },
    'julieth.galindo@provexpress.com.co': { grupo: 1, nombre: 'Julieth Galindo', archivo: 'Julieth Galindo.xlsx' },

    // Grupo 2: Caballero
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

    // Grupo 3: Beltrán
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

    // Grupo 4: Romero
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

export function normalizeEmail(value: unknown): string {
  return String(value || '').toLowerCase().trim();
}

export function normalizeName(value: unknown): string {
  return String(value || '')
    .replace(/\.(xlsx|xls)$/i, '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

const DIRECTORES_BY_GRUPO: Record<number, { nombre: string; email: string }> = {
  1: { nombre: 'Rafael Novoa', email: 'rafael.novoa@provexpress.com.co' },
  2: { nombre: 'Angélica Caballero', email: 'angelica.caballero@provexpress.com.co' },
  3: { nombre: 'Óscar Beltrán', email: 'oscar.beltran@provexpress.com.co' },
  4: { nombre: 'Miller Romero', email: 'miller.romero@provexpress.com.co' },
};

/**
 * Busca el correo de un ejecutivo a partir de su nombre o alias registrado en Excel.
 */
export function getExecutiveEmailByName(name: string): string {
  const target = normalizeName(name);
  if (!target) return '';

  const entries = Object.entries(ESTRUCTURA_COMERCIAL_2026.ejecutivos);

  // 1. Coincidencia exacta de nombre o archivo
  for (const [email, data] of entries) {
    if (
      normalizeName(data.nombre) === target ||
      normalizeName(data.archivo) === target
    ) {
      return email;
    }
  }

  // 2. Coincidencia parcial (subconjunto de palabras, ej. 'Mario Reyes Gutierre' -> 'Mario Reyes')
  const targetTokens = target.split(' ').filter(Boolean);
  for (const [email, data] of entries) {
    const dataTokens = normalizeName(data.nombre).split(' ').filter(Boolean);
    const commonTokens = targetTokens.filter((t) => dataTokens.includes(t));
    if (commonTokens.length >= 2) {
      return email;
    }
  }

  // 3. Coincidencia por archivo tokens
  for (const [email, data] of entries) {
    const fileTokens = normalizeName(data.archivo).split(' ').filter(Boolean);
    const commonTokens = targetTokens.filter((t) => fileTokens.includes(t));
    if (commonTokens.length >= 2) {
      return email;
    }
  }

  return '';
}

/**
 * Obtiene la información completa del comercial para el envío de notificaciones.
 */
export function getCommercialInfo(nameOrEmail: string): CommercialEntry {
  const normalized = normalizeEmail(nameOrEmail);

  // Si ya es un email conocido
  if (normalized.includes('@') && ESTRUCTURA_COMERCIAL_2026.ejecutivos[normalized]) {
    const data = ESTRUCTURA_COMERCIAL_2026.ejecutivos[normalized];
    const dir = DIRECTORES_BY_GRUPO[data.grupo] || { nombre: 'Dirección Comercial', email: '' };
    return {
      email: normalized,
      nombre: data.nombre,
      grupo: data.grupo,
      directorNombre: dir.nombre,
      directorEmail: dir.email,
      archivo: data.archivo,
    };
  }

  // Si es un nombre
  const email = getExecutiveEmailByName(nameOrEmail);
  if (email && ESTRUCTURA_COMERCIAL_2026.ejecutivos[email]) {
    const data = ESTRUCTURA_COMERCIAL_2026.ejecutivos[email];
    const dir = DIRECTORES_BY_GRUPO[data.grupo] || { nombre: 'Dirección Comercial', email: '' };
    return {
      email,
      nombre: data.nombre,
      grupo: data.grupo,
      directorNombre: dir.nombre,
      directorEmail: dir.email,
      archivo: data.archivo,
    };
  }

  // Fallback seguro
  return {
    email: normalized.includes('@') ? normalized : '',
    nombre: nameOrEmail,
    grupo: 0,
    directorNombre: 'Dirección Comercial',
    directorEmail: '',
  };
}

export interface DirectorInfo {
  grupo: number;
  nombre: string;
  email: string;
  carpeta: string;
}

export const LISTA_DIRECTORES: DirectorInfo[] = [
  { grupo: 1, nombre: 'Rafael Novoa', email: 'rafael.novoa@provexpress.com.co', carpeta: 'Grupo Rafael Novoa' },
  { grupo: 2, nombre: 'Angélica Caballero', email: 'angelica.caballero@provexpress.com.co', carpeta: 'Grupo Maria Angelica caballero' },
  { grupo: 3, nombre: 'Óscar Beltrán', email: 'oscar.beltran@provexpress.com.co', carpeta: 'Grupo Oscar Beltran' },
  { grupo: 4, nombre: 'Miller Romero', email: 'miller.romero@provexpress.com.co', carpeta: 'Grupo Miller Romero' },
];

export function getDirectorInfo(emailOrGroup: string | number): DirectorInfo | null {
  if (typeof emailOrGroup === 'number') {
    return LISTA_DIRECTORES.find((d) => d.grupo === emailOrGroup) || null;
  }
  const norm = normalizeEmail(emailOrGroup);
  return (
    LISTA_DIRECTORES.find((d) => normalizeEmail(d.email) === norm || normalizeName(d.nombre) === normalizeName(emailOrGroup)) ||
    null
  );
}

