import type { Remision, AmountStatus, DaysStatus, AlertLevel, ParsedWorkbook, GroupEntry } from '../types';
import { supabase } from './supabaseClient';
import {
  diffDays,
  getAgeRange,
  getAmountStatus,
  getDaysStatus,
  getAlert,
  normalizeText,
} from './remisiones';
import {
  resolveCommercialOrDirector,
  getDirectorInfo,
  normalizeName,
  ESTRUCTURA_COMERCIAL_2026,
} from './commercialDirectory';

export interface LiveProduct {
  codigo: string;
  descripcion: string;
  cantidad: number;
  pedido: string;
}

export interface LiveMovement {
  tipo: 'ENTRADA' | 'SALIDA';
  remision: Remision;
  detectadoAt: string;
}

export interface LiveSyncResult {
  cutoff: string;
  cutoffTime: string;
  records: Remision[];
  entrantes: Remision[];
  salientes: Remision[];
  totalValor: number;
  totalCount: number;
  saldoAnterior: number;
  saldoAnteriorCount?: number;
  corteAnteriorFecha?: string;
  gestionNeta: number;
  fromSupabase: boolean;
}

function getErpBaseUrl(): string {
  if (typeof window !== 'undefined') {
    if (window.location.protocol === 'https:') {
      return '/api/erp-proxy';
    }
    return '/erp-api';
  }
  return import.meta.env.VITE_REMISIONES_API_BASE || 'http://152.200.146.226:50010';
}

function formatTodayIso(): string {
  const d = new Date();
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

function formatCurrentTime(): string {
  return new Date().toLocaleTimeString('es-CO', {
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12: true,
  });
}

function makeStableKey(nit: string, doc: string, order: string, id: string): string {
  const cleanNit = normalizeText(nit);
  const cleanDoc = normalizeText(doc);
  const cleanOrder = normalizeText(order);
  if (cleanNit && cleanDoc) return `${cleanNit}|${cleanDoc}|${cleanOrder}`;
  if (cleanDoc) return `${cleanDoc}|${cleanOrder}`;
  return id;
}

/**
 * Obtiene el token JWT del backend del ERP
 */
async function authenticateErp(): Promise<string> {
  const baseUrl = getErpBaseUrl();
  const username = String(import.meta.env.VITE_REMISIONES_API_USER || 'powerbi').trim();
  let password = String(import.meta.env.VITE_REMISIONES_API_PASS || '3xpress#2025').trim();
  // Salvaguarda: si dotenv cortó la clave en el signo '#' por tratarlo como comentario
  if (password === '3xpress' || !password.includes('#')) {
    password = '3xpress#2025';
  }

  const res = await fetch(`${baseUrl}/api/getKey`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      username,
      password,
    }),
  });

  if (!res.ok) {
    throw new Error(`Error de autenticación ERP (${res.status}): ${res.statusText}`);
  }

  const data = await res.json();
  const token = data.token || data.access;
  if (!token) {
    throw new Error('La respuesta de autenticación del ERP no contiene token válido.');
  }
  return token;
}

/**
 * Consulta en tiempo real las remisiones vivas cruzando Valores con Pedidos y Productos
 */
export async function fetchLiveRemisionesFromErp(): Promise<{
  records: Remision[];
  cutoff: string;
  cutoffTime: string;
}> {
  const token = await authenticateErp();
  const baseUrl = getErpBaseUrl();
  const headers = {
    Authorization: `Bearer ${token}`,
    'Content-Type': 'application/json',
  };

  // Consultar ambas APIs en paralelo
  const [resValores, resProductos] = await Promise.all([
    fetch(`${baseUrl}/consultas/api/consultaDetalleEntregasMercanciaDashboardPBI`, { headers }),
    fetch(`${baseUrl}/consultas/api/consultaEntregasMercanciaDashboardPBI`, { headers }),
  ]);

  if (!resValores.ok) {
    throw new Error(`Error en API Valores (${resValores.status}): ${resValores.statusText}`);
  }
  if (!resProductos.ok) {
    throw new Error(`Error en API Productos (${resProductos.status}): ${resProductos.statusText}`);
  }

  const dataValores = await resValores.json();
  const dataProductos = await resProductos.json();

  const remisionesRaw: any[] = dataValores.response || [];
  const itemsRaw: any[] = dataProductos.response || [];

  // Mapear pedidos y productos por número de remisión
  const ordersMap = new Map<number, string>();
  const productsMap = new Map<number, LiveProduct[]>();

  for (const it of itemsRaw) {
    const num = it.Numero;
    if (num == null) continue;

    if (!ordersMap.has(num) && it.Pedido) {
      ordersMap.set(num, String(it.Pedido).trim());
    }

    if (!productsMap.has(num)) {
      productsMap.set(num, []);
    }
    productsMap.get(num)!.push({
      codigo: String(it.Codigo || '').trim(),
      descripcion: String(it.Producto || '').trim(),
      cantidad: Number(it.Cantidad || 1),
      pedido: String(it.Pedido || '').trim(),
    });
  }

  const cutoff = formatTodayIso();
  const cutoffTime = formatCurrentTime();
  const cutoffDateTime = new Date().toISOString();

  const records: Remision[] = [];

  for (let idx = 0; idx < remisionesRaw.length; idx += 1) {
    const raw = remisionesRaw[idx];
    const num = raw.Numero;
    const docStr = String(num != null ? num : '');
    const prefijo = String(raw.Prefijo || '').trim();
    const document = prefijo ? `${prefijo}${docStr}` : docStr;

    const nit = String(raw.Identificacion || '').trim();
    const order = ordersMap.get(num) || '';
    const employee = String(raw.Nombre_Empleado || raw.Empleado || '').trim();
    const company = String(raw.Empresa || '').trim();

    const merchandise = Number(raw.Valor_Mercancia || 0);
    const tax = Number(raw.Valor_Iva || 0);
    const total = Number(raw.Valor_Total || merchandise + tax);

    // Fecha de emisión
    let issuedAt = '';
    if (raw.Fecha_Emision) {
      issuedAt = String(raw.Fecha_Emision).slice(0, 10);
    }
    const age = issuedAt ? diffDays(cutoff, issuedAt) : 0;

    // Resolver director y grupo
    const role = resolveCommercialOrDirector(employee);
    const grupoPersonalStr = String(raw.Grupo_Personal || '');
    let director = role.directorNombre || 'Sin asignar';
    let groupNum: number | null = role.grupo > 0 ? role.grupo : null;

    if (!groupNum && grupoPersonalStr) {
      const match = grupoPersonalStr.match(/\d+/);
      if (match) groupNum = parseInt(match[0], 10);
    }

    const id = `ERP-LIVE-${cutoff}-${idx + 1}-${document || order}`;
    const stableKey = makeStableKey(nit, document, order, id);

    records.push({
      id,
      stableKey,
      cutoff,
      cutoffTime,
      cutoffDateTime,
      employee,
      nit,
      company,
      merchandise,
      tax,
      total,
      issuedAt,
      age,
      document,
      order,
      quantity: 1,
      ageRange: getAgeRange(age),
      amountStatus: getAmountStatus(total),
      daysStatus: getDaysStatus(age),
      alert: getAlert(total, age),
      director,
      group: groupNum,
      matchedGroup: groupNum != null,
    });
  }

  return { records, cutoff, cutoffTime };
}

/**
 * Sincroniza la API viva contra Supabase:
 * 1. Consulta la API viva de Remisiones.
 * 2. Compara contra el último corte registrado en Supabase.
 * 3. Identifica Salientes (Facturadas) y Entrantes (Nuevas).
 * 4. Guarda la foto de hoy y los movimientos en Supabase.
 */
export async function syncLiveRemisiones(): Promise<LiveSyncResult> {
  const { records, cutoff, cutoffTime } = await fetchLiveRemisionesFromErp();

  // 1. Obtener el corte anterior más reciente desde Supabase (fecha < hoy)
  let prevRecords: Remision[] = [];
  let latestCorteData: any = null;
  try {
    const { data: latestCorte, error } = await supabase
      .from('cortes_historicos')
      .select('*')
      .lt('fecha', cutoff)
      .order('fecha', { ascending: false })
      .limit(1)
      .maybeSingle();

    if (!error && latestCorte) {
      latestCorteData = latestCorte;
      if (latestCorte.records_json) {
        prevRecords = latestCorte.records_json as Remision[];
      }
    }
  } catch (err) {
    console.warn('No se pudo consultar el corte anterior en Supabase:', err);
  }

  // 2. Comparar conjuntos por documento y NIT
  const getDocumentKey = (r: any): string => {
    const doc = String(r.document || r.Numero || '').trim().toLowerCase();
    const nit = String(r.nit || r.Identificacion || '').trim().toLowerCase();
    if (doc) return `${nit}|${doc}`;
    return String(r.stableKey || r.id || '');
  };

  const currentKeyMap = new Map<string, Remision>();
  for (const r of records) {
    currentKeyMap.set(getDocumentKey(r), r);
  }

  const prevKeyMap = new Map<string, Remision>();
  for (const r of prevRecords) {
    prevKeyMap.set(getDocumentKey(r), r);
  }

  // Salientes: estaban en el corte previo pero ya no en el vivo (fueron facturadas!)
  const salientes: Remision[] = [];
  for (const [key, prevRem] of prevKeyMap.entries()) {
    if (!currentKeyMap.has(key)) {
      salientes.push(prevRem);
    }
  }

  // Entrantes: están en el vivo pero no estaban en el previo (nuevas!)
  const entrantes: Remision[] = [];
  for (const [key, currRem] of currentKeyMap.entries()) {
    if (!prevKeyMap.has(key)) {
      entrantes.push(currRem);
    }
  }

  const totalValor = records.reduce((s, r) => s + r.total, 0);
  const totalCount = records.length;
  const saldoAnterior = prevRecords.reduce((s, r) => s + r.total, 0);
  const salientesValor = salientes.reduce((s, r) => s + r.total, 0);
  const entrantesValor = entrantes.reduce((s, r) => s + r.total, 0);
  const gestionNeta = salientesValor - entrantesValor;

  // 3. Guardar el nuevo corte en Supabase
  try {
    const totalClientes = new Set(records.map((r) => r.nit).filter(Boolean)).size;

    await supabase.from('cortes_historicos').upsert({
      id: cutoff,
      fecha: cutoff,
      total_valor: Math.round(totalValor),
      total_remisiones: totalCount,
      total_clientes: totalClientes,
      saldo_anterior: Math.round(saldoAnterior),
      entrantes_valor: Math.round(entrantesValor),
      entrantes_count: entrantes.length,
      salientes_valor: Math.round(salientesValor),
      salientes_count: salientes.length,
      gestion_neta: Math.round(gestionNeta),
      records_json: records,
    });

    // Registrar movimientos en la tabla remisiones_movimientos
    const movimientosRows = [
      ...salientes.map((s) => ({
        fecha_corte: cutoff,
        tipo: 'SALIDA',
        numero: parseInt(s.document.replace(/\D/g, ''), 10) || 0,
        prefijo: '',
        pedido: s.order || '',
        nit: s.nit,
        empresa: s.company,
        empleado: s.employee,
        grupo_personal: s.director,
        valor_total: s.total,
        fecha_emision: s.issuedAt ? `${s.issuedAt}T00:00:00` : null,
      })),
      ...entrantes.map((e) => ({
        fecha_corte: cutoff,
        tipo: 'ENTRADA',
        numero: parseInt(e.document.replace(/\D/g, ''), 10) || 0,
        prefijo: '',
        pedido: e.order || '',
        nit: e.nit,
        empresa: e.company,
        empleado: e.employee,
        grupo_personal: e.director,
        valor_total: e.total,
        fecha_emision: e.issuedAt ? `${e.issuedAt}T00:00:00` : null,
      })),
    ];

    if (movimientosRows.length > 0) {
      await supabase.from('remisiones_movimientos').insert(movimientosRows);
    }
  } catch (err) {
    console.warn('Error guardando corte en Supabase:', err);
  }

  return {
    cutoff,
    cutoffTime,
    records,
    entrantes,
    salientes,
    totalValor,
    totalCount,
    saldoAnterior,
    saldoAnteriorCount: prevRecords.length,
    corteAnteriorFecha: latestCorteData?.fecha || '2026-09-25',
    gestionNeta,
    fromSupabase: true,
  };
}

/**
 * Carga todo el histórico de cortes desde Supabase para armar las gráficas de Evolución
 */
export async function loadHistoricalCortesFromSupabase(): Promise<{
  cortes: Array<{
    id: string;
    fecha: string;
    total_valor: number;
    total_remisiones: number;
    total_clientes: number;
    saldo_anterior: number;
    entrantes_valor: number;
    entrantes_count: number;
    salientes_valor: number;
    salientes_count: number;
    gestion_neta: number;
    records_json: Remision[];
  }>;
}> {
  const { data, error } = await supabase
    .from('cortes_historicos')
    .select('*')
    .order('fecha', { ascending: true });

  if (error || !data) {
    throw new Error(`Error cargando cortes de Supabase: ${error?.message || 'Sin datos'}`);
  }

  return { cortes: data as any };
}

/**
 * Reconstruye el objeto ParsedWorkbook completo a partir de todos los cortes
 * guardados en la tabla cortes_historicos de Supabase.
 * Permite que la aplicación arranque instantáneamente sin descargar el Excel de SharePoint.
 */
export async function buildWorkbookFromSupabase(): Promise<ParsedWorkbook> {
  const { data, error } = await supabase
    .from('cortes_historicos')
    .select('*')
    .order('fecha', { ascending: true });

  if (error || !data || data.length === 0) {
    throw new Error(`No hay cortes en Supabase: ${error?.message || 'Sin registros'}`);
  }

  const dirNames: Record<number, string> = {
    1: 'Rafael Novoa',
    2: 'Angélica Caballero',
    3: 'Óscar Beltrán',
    4: 'Miller Romero',
  };

  const groups: GroupEntry[] = Object.values(ESTRUCTURA_COMERCIAL_2026.ejecutivos).map((e) => ({
    group: e.grupo,
    director: dirNames[e.grupo] || `Grupo ${e.grupo}`,
    member: e.nombre,
  }));

  const allRecords: Remision[] = [];
  const cutoffs: string[] = [];

  for (const corte of data) {
    const cutoffDate = String(corte.fecha).trim();
    if (cutoffDate && !cutoffs.includes(cutoffDate)) {
      cutoffs.push(cutoffDate);
    }
    const rawRecords = (corte.records_json || []) as any[];

    for (let idx = 0; idx < rawRecords.length; idx++) {
      const r = rawRecords[idx];
      const cutoff = r.cutoff || cutoffDate;
      const employee = String(r.employee || '').trim();
      const nit = String(r.nit || '').trim();
      const company = String(r.company || '').trim();
      const document = String(r.document || '').trim();
      const order = String(r.order || '').trim();
      const issuedAt = r.issuedAt ? String(r.issuedAt).slice(0, 10) : '';
      const total = Number(r.total || 0);
      const merchandise = typeof r.merchandise === 'number' ? r.merchandise : total;
      const tax = typeof r.tax === 'number' ? r.tax : Math.max(0, total - merchandise);
      const age = typeof r.age === 'number' ? r.age : (issuedAt ? Math.max(0, diffDays(cutoff, issuedAt)) : 0);
      const quantity = typeof r.quantity === 'number' ? r.quantity : 1;

      const role = resolveCommercialOrDirector(employee);
      const director = (r.director && r.director !== 'Sin asignar') ? r.director : (role.directorNombre || 'Sin asignar');
      const groupNum = typeof r.group === 'number' ? r.group : (role.grupo > 0 ? role.grupo : null);

      const id = r.id || `SB-${cutoff}-${idx + 1}-${document || order}`;
      const stableKey = r.stableKey || makeStableKey(nit, document, order, id);

      allRecords.push({
        id,
        stableKey,
        cutoff,
        cutoffTime: r.cutoffTime,
        cutoffDateTime: r.cutoffDateTime,
        employee,
        nit,
        company,
        merchandise,
        tax,
        total,
        issuedAt,
        age,
        document,
        order,
        quantity,
        ageRange: r.ageRange || getAgeRange(age),
        amountStatus: r.amountStatus || getAmountStatus(total),
        daysStatus: r.daysStatus || getDaysStatus(age),
        alert: r.alert || getAlert(total, age),
        director,
        group: groupNum,
        matchedGroup: groupNum != null,
      });
    }
  }

  const uniqueCutoffs = [...new Set(cutoffs)].sort();
  const latestCutoff = uniqueCutoffs.at(-1) || '';
  const unmatchedEmployees = [...new Set(allRecords.filter((r) => !r.matchedGroup).map((r) => r.employee))]
    .sort((a, b) => a.localeCompare(b, 'es'));

  return {
    records: allRecords,
    groups,
    sheetNames: ['Base', 'Grupos', 'Diario', 'Supabase'],
    cutoffs: uniqueCutoffs,
    unmatchedEmployees,
    activeSheetName: 'Base (Supabase Cloud)',
    cutoffDateTime: `${latestCutoff}T12:00:00Z`,
    cutoffTimeDisplay: 'Tiempo Real',
  };
}
