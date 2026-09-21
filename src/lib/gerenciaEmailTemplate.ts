import type { Remision } from '../types';
import {
  ESTRUCTURA_COMERCIAL_2026,
  LISTA_DIRECTORES,
  normalizeName,
  resolveCommercialOrDirector,
} from './commercialDirectory';
import { LOGO_PROVEXPRESS_DATA_URI, AVATAR_MAN_DATA_URI, AVATAR_WOMAN_DATA_URI } from './commercialEmailAssets';
import { formatCOP, formatNumber } from './commercialEmailTemplate';

export interface GerenciaGroupSummary {
  grupo: number;
  directorName: string;
  directorEmail: string;
  activeExecutivesCount: number;
  totalExecutivesCount: number;
  totalCount: number;
  totalValue: number;
  pctCompanyValue: number;
  avgAge: number;
  criticalRemision: (Remision & { executiveName?: string }) | null;
  status: 'Al día' | 'Gestión activa' | 'Atención prioritaria';
  statusColor: string;
  statusBg: string;
}

export interface GerenciaUnassignedSummary {
  label: string;
  totalCount: number;
  totalValue: number;
  pctCompanyValue: number;
  avgAge: number;
  criticalRemision: (Remision & { executiveName?: string; directorName?: string }) | null;
}

export interface GerenciaTopExecutiveSummary {
  name: string;
  email: string;
  grupo: number;
  directorName: string;
  count: number;
  total: number;
  avgAge: number;
  pctCompanyValue: number;
}

export interface GerenciaEmailSummary {
  recipientName: string;
  recipientCargo: string;
  recipientEmail: string;
  cutoffDate: string;
  totalCount: number;
  totalValue: number;
  avgAge: number;
  activeDirectorsCount: number;
  totalDirectorsCount: number;
  activeExecutivesCount: number;
  totalExecutivesCount: number;
  topMayorValor: (Remision & { executiveName?: string; directorName?: string }) | null;
  topMayorAntiguedad: (Remision & { executiveName?: string; directorName?: string }) | null;
  groupsSummary: GerenciaGroupSummary[];
  unassignedSummary?: GerenciaUnassignedSummary;
  topExecutives: GerenciaTopExecutiveSummary[];
  allCompanyRemisiones: (Remision & { executiveName?: string; directorName?: string; isDirectorDirect?: boolean })[];
  genero: 'M' | 'F';
}

export interface GenerateGerenciaEmailOptions {
  forWebPreview?: boolean;
}

/**
 * Construye el resumen consolidado a nivel gerencial de toda la compañía
 */
export function buildGerenciaEmailSummary(
  records: Remision[],
  cutoffDate: string,
  recipientOptions?: {
    name?: string;
    cargo?: string;
    email?: string;
    genero?: 'M' | 'F';
  },
): GerenciaEmailSummary {
  const recipientName = recipientOptions?.name || 'Dirección y Gerencia Comercial';
  const recipientCargo = recipientOptions?.cargo || 'Dirección Comercial & Gerencia General';
  const recipientEmail = recipientOptions?.email || 'rafael.novoa@provexpress.com.co';
  const genero = recipientOptions?.genero || 'M';

  // Filtrar remisiones de este corte si se especifica
  const cutoffRecords = records.filter((r) => !cutoffDate || r.cutoff === cutoffDate);

  // Clasificar cada registro con la lógica unificada de resolución
  const classifiedRecords: (Remision & {
    executiveName: string;
    directorName: string;
    resolvedGroup: number;
    isDirectorDirect: boolean;
  })[] = [];

  for (const r of cutoffRecords) {
    const role = resolveCommercialOrDirector(r.employee);
    if (role.type === 'ejecutivo') {
      classifiedRecords.push({
        ...r,
        group: role.grupo,
        director: role.directorNombre,
        executiveName: role.nombre,
        directorName: role.directorNombre,
        resolvedGroup: role.grupo,
        isDirectorDirect: false,
      });
    } else if (role.type === 'director_directa') {
      classifiedRecords.push({
        ...r,
        group: role.grupo,
        director: role.directorNombre,
        executiveName: role.nombre,
        directorName: role.directorNombre,
        resolvedGroup: role.grupo,
        isDirectorDirect: true,
      });
    } else {
      classifiedRecords.push({
        ...r,
        group: 0,
        director: 'Otras Áreas / Cuentas Especiales',
        executiveName: r.employee || 'Cuentas Especiales',
        directorName: 'Otras Áreas / Especiales',
        resolvedGroup: 0,
        isDirectorDirect: false,
      });
    }
  }

  const allCompanyRemisiones = classifiedRecords;
  const totalCount = allCompanyRemisiones.length;
  const totalValue = allCompanyRemisiones.reduce((s, r) => s + (r.total || 0), 0);
  const avgAge = totalCount > 0 ? Math.round(allCompanyRemisiones.reduce((s, r) => s + (r.age || 0), 0) / totalCount) : 0;

  // Mapear ejecutivos y grupos
  const allExecEntries = Object.entries(ESTRUCTURA_COMERCIAL_2026.ejecutivos);
  const totalExecutivesCount = allExecEntries.length;

  // Resumen de ejecutivos para Top 5
  const execSummariesList: GerenciaTopExecutiveSummary[] = [];
  for (const [execEmail, execData] of allExecEntries) {
    const dirInfo = LISTA_DIRECTORES.find((d) => d.grupo === execData.grupo);
    const directorName = dirInfo?.nombre || `Director Grupo ${execData.grupo}`;
    const execRemisiones = classifiedRecords.filter(
      (r) => !r.isDirectorDirect && r.resolvedGroup === execData.grupo && r.executiveName === execData.nombre
    );

    if (execRemisiones.length > 0) {
      const execTotal = execRemisiones.reduce((s, r) => s + (r.total || 0), 0);
      const execAvgAge = Math.round(execRemisiones.reduce((s, r) => s + (r.age || 0), 0) / execRemisiones.length);
      execSummariesList.push({
        name: execData.nombre,
        email: execEmail,
        grupo: execData.grupo,
        directorName,
        count: execRemisiones.length,
        total: execTotal,
        avgAge: execAvgAge,
        pctCompanyValue: totalValue > 0 ? Number(((execTotal / totalValue) * 100).toFixed(1)) : 0,
      });
    }
  }

  const activeExecutivesCount = execSummariesList.length;
  execSummariesList.sort((a, b) => b.total - a.total);
  const topExecutives = execSummariesList.slice(0, 5);

  // Top oportunidades globales
  const byValue = [...allCompanyRemisiones].sort((a, b) => (b.total || 0) - (a.total || 0));
  const topMayorValor = byValue[0] || null;

  const byAge = [...allCompanyRemisiones].sort((a, b) => (b.age || 0) - (a.age || 0));
  const topMayorAntiguedad = byAge[0] || null;

  // Consolidar por cada uno de los 4 directores / grupos
  const groupsSummary: GerenciaGroupSummary[] = LISTA_DIRECTORES.map((dir) => {
    const groupGroupExecutives = allExecEntries.filter(([, data]) => data.grupo === dir.grupo);
    const groupRemisiones = classifiedRecords.filter((r) => r.resolvedGroup === dir.grupo);

    const gCount = groupRemisiones.length;
    const gTotal = groupRemisiones.reduce((s, r) => s + (r.total || 0), 0);
    const gAvgAge = gCount > 0 ? Math.round(groupRemisiones.reduce((s, r) => s + (r.age || 0), 0) / gCount) : 0;
    const gPct = totalValue > 0 ? Number(((gTotal / totalValue) * 100).toFixed(1)) : 0;

    // Asesores activos únicos (excluyendo gestión directa de directores)
    const activeExecsSet = new Set(
      groupRemisiones.filter((r) => !r.isDirectorDirect).map((r) => r.executiveName)
    );
    const gActiveExecs = activeExecsSet.size;

    const sortedGroupAge = [...groupRemisiones].sort((a, b) => (b.age || 0) - (a.age || 0));
    const criticalRemision = sortedGroupAge[0] || null;

    let status: 'Al día' | 'Gestión activa' | 'Atención prioritaria' = 'Gestión activa';
    let statusColor = '#1E3A8A';
    let statusBg = '#DBEAFE';

    if (gCount === 0 || gAvgAge <= 15) {
      status = 'Al día';
      statusColor = '#15803D';
      statusBg = '#DCFCE7';
    } else if (gAvgAge > 25 || (criticalRemision && (criticalRemision.age || 0) > 60)) {
      status = 'Atención prioritaria';
      statusColor = '#B91C1C';
      statusBg = '#FEE2E2';
    }

    return {
      grupo: dir.grupo,
      directorName: dir.nombre,
      directorEmail: dir.email,
      activeExecutivesCount: gActiveExecs,
      totalExecutivesCount: groupGroupExecutives.length,
      totalCount: gCount,
      totalValue: gTotal,
      pctCompanyValue: gPct,
      avgAge: gAvgAge,
      criticalRemision,
      status,
      statusColor,
      statusBg,
    };
  });

  const activeDirectorsCount = groupsSummary.filter((g) => g.totalCount > 0).length;

  // Cuentas Especiales / Otras Áreas (sin director comercial asignado)
  const unassignedRemisiones = classifiedRecords.filter((r) => r.resolvedGroup === 0);
  let unassignedSummary: GerenciaUnassignedSummary | undefined;
  if (unassignedRemisiones.length > 0) {
    const uTotal = unassignedRemisiones.reduce((s, r) => s + (r.total || 0), 0);
    const uAvgAge = Math.round(unassignedRemisiones.reduce((s, r) => s + (r.age || 0), 0) / unassignedRemisiones.length);
    const sortedU = [...unassignedRemisiones].sort((a, b) => (b.age || 0) - (a.age || 0));
    unassignedSummary = {
      label: 'Cuentas Especiales / Otras Áreas',
      totalCount: unassignedRemisiones.length,
      totalValue: uTotal,
      pctCompanyValue: totalValue > 0 ? Number(((uTotal / totalValue) * 100).toFixed(1)) : 0,
      avgAge: uAvgAge,
      criticalRemision: sortedU[0] || null,
    };
  }

  return {
    recipientName,
    recipientCargo,
    recipientEmail,
    cutoffDate,
    totalCount,
    totalValue,
    avgAge,
    activeDirectorsCount,
    totalDirectorsCount: LISTA_DIRECTORES.length,
    activeExecutivesCount,
    totalExecutivesCount,
    topMayorValor,
    topMayorAntiguedad,
    groupsSummary,
    unassignedSummary,
    topExecutives,
    allCompanyRemisiones,
    genero,
  };
}

/**
 * Genera la plantilla HTML corporativa de alta dirección para Outlook Desktop, Web y Móvil.
 */
export function generateGerenciaEmailHtml(
  summary: GerenciaEmailSummary,
  options: GenerateGerenciaEmailOptions = {},
): string {
  const { forWebPreview = false } = options;
  const isFemale = summary.genero === 'F';
  const logoSrc = forWebPreview ? LOGO_PROVEXPRESS_DATA_URI : 'cid:logo_provexpress';
  const avatarSrc = forWebPreview
    ? (isFemale ? AVATAR_WOMAN_DATA_URI : AVATAR_MAN_DATA_URI)
    : (isFemale ? 'cid:avatar_woman' : 'cid:avatar_man');
  const avatarAlt = isFemale ? 'Líder Comercial' : 'Líder Comercial';

  const {
    recipientName,
    recipientCargo,
    cutoffDate,
    totalCount,
    totalValue,
    avgAge,
    activeDirectorsCount,
    totalDirectorsCount,
    activeExecutivesCount,
    totalExecutivesCount,
    topMayorValor,
    topMayorAntiguedad,
    groupsSummary,
    unassignedSummary,
    topExecutives,
  } = summary;

  // Formato de fecha
  let formattedCutoff = cutoffDate;
  if (cutoffDate) {
    try {
      const [year, month, day] = cutoffDate.split('-');
      if (year && month && day) {
        const d = new Date(Number(year), Number(month) - 1, Number(day));
        formattedCutoff = d.toLocaleDateString('es-CO', {
          day: 'numeric',
          month: 'long',
          year: 'numeric',
        });
      }
    } catch {
      formattedCutoff = cutoffDate;
    }
  }

  // Filas de la tabla de directores (El resultado de los directivos)
  const rowsDirectorsHtml = groupsSummary
    .map((g, idx) => {
      const isEven = idx % 2 === 0;
      const bg = isEven ? '#FFFFFF' : '#F8FAFC';

      let ageColor = '#15803D';
      let ageBg = '#DCFCE7';
      if (g.avgAge > 25) {
        ageColor = '#B91C1C';
        ageBg = '#FEE2E2';
      } else if (g.avgAge > 15) {
        ageColor = '#B45309';
        ageBg = '#FEF3C7';
      }

      const critText = g.criticalRemision
        ? `${g.criticalRemision.document?.startsWith('REM-') ? g.criticalRemision.document : `REM-${g.criticalRemision.document || 'S/N'}`} (${g.criticalRemision.age}d)`
        : '—';

      return `
      <tr style="background-color: ${bg}; border-bottom: 1px solid #E2E8F0;">
        <td style="padding: 11px 14px; font-size: 13px; font-weight: 800; color: #1E293B; font-family: 'Segoe UI', Arial, sans-serif;">
          <span style="display: inline-block; width: 22px; height: 22px; line-height: 22px; text-align: center; border-radius: 6px; background-color: #1E3A8A; color: #FFFFFF; font-size: 11px; margin-right: 6px;">
            G${g.grupo}
          </span>
          ${g.directorName}
          <div style="font-size: 10.5px; font-weight: 500; color: #64748B; margin-left: 28px;">
            ${g.directorEmail}
          </div>
        </td>
        <td style="padding: 11px 10px; text-align: center; font-size: 12px; font-weight: 700; color: #334155; font-family: 'Segoe UI', Arial, sans-serif;">
          <span style="color: #0F172A; font-weight: 900;">${g.activeExecutivesCount}</span> <span style="color: #94A3B8;">/ ${g.totalExecutivesCount}</span>
        </td>
        <td style="padding: 11px 10px; text-align: center; font-size: 13px; font-weight: 850; color: #0F172A; font-family: 'Segoe UI', Arial, sans-serif;">
          ${formatNumber(g.totalCount)}
        </td>
        <td style="padding: 11px 14px; text-align: right; font-size: 13.5px; font-weight: 900; color: #15803D; font-family: 'Segoe UI', Arial, sans-serif;">
          ${formatCOP(g.totalValue)}
        </td>
        <td style="padding: 11px 10px; text-align: center; font-size: 12px; font-weight: 800; color: #1E3A8A; font-family: 'Segoe UI', Arial, sans-serif;">
          ${g.pctCompanyValue}%
        </td>
        <td style="padding: 11px 10px; text-align: center;">
          <span style="display: inline-block; padding: 3px 8px; border-radius: 6px; font-size: 11px; font-weight: 700; color: ${ageColor}; background-color: ${ageBg}; font-family: 'Segoe UI', Arial, sans-serif;">
            ${g.avgAge} días
          </span>
        </td>
        <td style="padding: 11px 12px; font-size: 11.5px; color: #475569; font-family: 'Segoe UI', Arial, sans-serif;">
          ${critText}
        </td>
        <td style="padding: 11px 12px; text-align: center;">
          <span style="display: inline-block; padding: 2px 8px; border-radius: 999px; font-size: 10px; font-weight: 700; color: ${g.statusColor}; background-color: ${g.statusBg};">
            ${g.status}
          </span>
        </td>
      </tr>
      `;
    })
    .join('');

  // Filas de Top 5 Ejecutivos Comerciales
  const rowsTopExecsHtml = topExecutives
    .map((e, idx) => {
      const isEven = idx % 2 === 0;
      const bg = isEven ? '#FFFFFF' : '#F8FAFC';
      return `
      <tr style="background-color: ${bg}; border-bottom: 1px solid #E2E8F0;">
        <td style="padding: 9px 12px; text-align: center; font-size: 12px; font-weight: 800; color: #1E3A8A; font-family: 'Segoe UI', Arial, sans-serif;">
          #${idx + 1}
        </td>
        <td style="padding: 9px 12px; font-size: 12.5px; font-weight: 750; color: #0F172A; font-family: 'Segoe UI', Arial, sans-serif;">
          ${e.name}
          <div style="font-size: 10px; color: #64748B;">Grupo ${e.grupo} · ${e.directorName}</div>
        </td>
        <td style="padding: 9px 10px; text-align: center; font-size: 12px; font-weight: 800; color: #0F172A; font-family: 'Segoe UI', Arial, sans-serif;">
          ${e.count}
        </td>
        <td style="padding: 9px 12px; text-align: right; font-size: 12.5px; font-weight: 900; color: #15803D; font-family: 'Segoe UI', Arial, sans-serif;">
          ${formatCOP(e.total)}
        </td>
        <td style="padding: 9px 10px; text-align: center; font-size: 11.5px; font-weight: 700; color: #D97706; font-family: 'Segoe UI', Arial, sans-serif;">
          ${e.avgAge} días
        </td>
        <td style="padding: 9px 10px; text-align: center; font-size: 11.5px; font-weight: 800; color: #1E3A8A; font-family: 'Segoe UI', Arial, sans-serif;">
          ${e.pctCompanyValue}%
        </td>
      </tr>
      `;
    })
    .join('');

  return `<!DOCTYPE html>
<html lang="es" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Informe Gerencial de Remisiones Abiertas · Provexpress SAS</title>
  <!--[if mso]>
  <noscript>
    <xml>
      <o:OfficeDocumentSettings>
        <o:PixelsPerInch>96</o:PixelsPerInch>
      </o:OfficeDocumentSettings>
    </xml>
  </noscript>
  <![endif]-->
  <style>
    body, table, td, a { -webkit-text-size-adjust: 100%; -ms-text-size-adjust: 100%; }
    table, td { mso-table-lspace: 0pt; mso-table-rspace: 0pt; }
    img { -ms-interpolation-mode: bicubic; border: 0; outline: none; text-decoration: none; }
    body { margin: 0; padding: 0; width: 100% !important; background-color: #F1F4F8; font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, Helvetica, Arial, sans-serif; }
    @media only screen and (max-width: 920px) {
      .main-card { width: 100% !important; border-radius: 0 !important; }
      .stack-col { display: block !important; width: 100% !important; padding-right: 0 !important; }
      .kpi-cell { display: inline-block !important; width: 48% !important; margin-bottom: 8px !important; }
    }
  </style>
</head>
<body style="margin: 0; padding: 24px 0; background-color: #F1F4F8;">

  <!-- CONTENEDOR PRINCIPAL BLANCO (960px) -->
  <table role="presentation" class="main-card" width="960" align="center" border="0" cellpadding="0" cellspacing="0" style="width: 960px; margin: 0 auto; background-color: #FFFFFF; border-radius: 16px; overflow: hidden; border: 1.5px solid #E2E8F0; box-shadow: 0 4px 20px rgba(15, 23, 42, 0.05);">
    <tr>
      <td style="padding: 24px 28px;">

        <!-- ═════════════════════════════════════════════════════════════════════
             1. ENCABEZADO: LOGO + SALUDO + AVATAR + CARD CONTROL TOTAL
             ═════════════════════════════════════════════════════════════════════ -->
        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="margin-bottom: 20px;">
          <tr>
            <!-- LOGO PROVEXPRESS OFICIAL -->
            <td width="160" valign="middle" align="left" style="padding-right: 18px;">
              <img src="${logoSrc}" width="144" height="110" style="display: block; border: 0; width: 144px; height: 110px; max-width: 144px;" alt="Provexpress" />
            </td>

            <!-- TEXTO DE SALUDO Y TITULAR -->
            <td valign="middle" style="padding-right: 12px;">
              <div style="font-size: 14.5px; color: #475569; font-weight: 600; font-family: 'Segoe UI', Arial, sans-serif;">
                Hola, <strong style="color: #0F172A; font-size: 16px;">${recipientName}</strong>
                <span style="display: inline-block; margin-left: 8px; font-size: 11px; background-color: #EFF6FF; color: #1E3A8A; padding: 2px 8px; border-radius: 999px; font-weight: 700;">${recipientCargo}</span>
              </div>
              <h1 style="margin: 3px 0 6px 0; font-size: 23px; font-weight: 900; color: #0F172A; line-height: 1.15; letter-spacing: -0.02em; font-family: 'Segoe UI', Arial, sans-serif;">
                Consolidado General de Gestión Comercial<br>
                <span style="color: #16A34A;">Dirección & Gerencia</span> · Corte ${formattedCutoff}
              </h1>
              <p style="margin: 0; font-size: 12px; color: #475569; line-height: 1.35; font-family: 'Segoe UI', Arial, sans-serif; max-width: 410px;">
                Visión integral del estado de remisiones abiertas y oportunidades de facturación en todos los <strong>4 grupos comerciales</strong> y sus <strong>${totalExecutivesCount} ejecutivos</strong>.
              </p>
              <p style="margin: 5px 0 0 0; font-size: 12px; color: #0F172A; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">
                Coordinación estratégica para transformar entregas en <span style="color: #16A34A;">ventas facturadas</span>.
              </p>
            </td>

            <!-- ILUSTRACIÓN LÍDER COMERCIAL: AVATAR THUMBS UP -->
            <td width="190" valign="bottom" align="center" style="padding-right: 10px;">
              <img src="${avatarSrc}" width="180" height="154" style="display: block; border: 0; width: 180px; height: auto; margin: 0 auto;" alt="${avatarAlt}" />
            </td>

            <!-- CARD "Control Total · Visión Corporativa" -->
            <td width="125" valign="middle" align="right">
              <table role="presentation" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 14px; padding: 14px 8px; width: 120px; text-align: center;">
                <tr>
                  <td align="center">
                    <div style="background-color: #1E3A8A; color: #FFFFFF; width: 26px; height: 26px; border-radius: 50%; font-size: 14px; font-weight: 900; line-height: 26px; margin: 0 auto 8px auto;">
                      ★
                    </div>
                    <div style="font-size: 11.5px; font-weight: 700; color: #1E293B; line-height: 1.3; font-family: 'Segoe UI', Arial, sans-serif;">
                      Control Total<br>
                      Visión<br>
                      <strong style="color: #1E3A8A; font-size: 12.5px;">Corporativa</strong>
                    </div>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>

        <!-- ═════════════════════════════════════════════════════════════════════
             2. 4 TARJETAS KPI GLOBALES DE LA COMPAÑÍA
             ═════════════════════════════════════════════════════════════════════ -->
        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="margin-bottom: 20px;">
          <tr>
            <!-- KPI 1: Grupos y Ejecutivos Activos -->
            <td width="24%" valign="top" class="kpi-cell" style="padding-right: 8px;">
              <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 12px; padding: 12px 10px; height: 95px;">
                <tr>
                  <td width="32" valign="top">
                    <div style="background-color: #1E3A8A; width: 30px; height: 30px; border-radius: 50%; text-align: center; line-height: 30px; font-size: 15px; color: #FFFFFF;">
                      👔
                    </div>
                  </td>
                  <td valign="top" style="padding-left: 6px;">
                    <div style="font-size: 10.5px; color: #475569; font-weight: 600; line-height: 1.2; font-family: 'Segoe UI', Arial, sans-serif;">
                      Grupos y Ejecutivos<br>con pendientes
                    </div>
                  </td>
                </tr>
                <tr>
                  <td colspan="2" valign="bottom" style="padding-top: 6px;">
                    <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td valign="bottom" style="font-size: 17px; font-weight: 900; color: #1E3A8A; font-family: 'Segoe UI', Arial, sans-serif; line-height: 1;">
                          ${activeDirectorsCount} grupos · ${activeExecutivesCount} asesores
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>

            <!-- KPI 2: Total Remisiones Abiertas Empresa -->
            <td width="24%" valign="top" class="kpi-cell" style="padding-right: 8px;">
              <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 12px; padding: 12px 10px; height: 95px;">
                <tr>
                  <td width="32" valign="top">
                    <div style="background-color: #312E81; width: 30px; height: 30px; border-radius: 50%; text-align: center; line-height: 30px; font-size: 15px; color: #FFFFFF;">
                      📦
                    </div>
                  </td>
                  <td valign="top" style="padding-left: 6px;">
                    <div style="font-size: 10.5px; color: #475569; font-weight: 600; line-height: 1.2; font-family: 'Segoe UI', Arial, sans-serif;">
                      Remisiones abiertas<br>en la compañía
                    </div>
                  </td>
                </tr>
                <tr>
                  <td colspan="2" valign="bottom" style="padding-top: 6px;">
                    <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td valign="bottom" style="font-size: 24px; font-weight: 900; color: #312E81; font-family: 'Segoe UI', Arial, sans-serif; line-height: 1;">
                          ${formatNumber(totalCount)}
                        </td>
                        <td align="right" valign="bottom" style="font-size: 17px; color: #94A3B8; opacity: 0.65;">
                          📝
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>

            <!-- KPI 3: Total Valor Empresa por Facturar -->
            <td width="28%" valign="top" class="kpi-cell" style="padding-right: 8px;">
              <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 12px; padding: 12px 10px; height: 95px;">
                <tr>
                  <td width="32" valign="top">
                    <div style="background-color: #15803D; width: 30px; height: 30px; border-radius: 50%; text-align: center; line-height: 30px; font-size: 15px; font-weight: 900; color: #FFFFFF;">
                      $
                    </div>
                  </td>
                  <td valign="top" style="padding-left: 6px;">
                    <div style="font-size: 10.5px; color: #475569; font-weight: 600; line-height: 1.2; font-family: 'Segoe UI', Arial, sans-serif;">
                      Total por facturar<br>compañía
                    </div>
                  </td>
                </tr>
                <tr>
                  <td colspan="2" valign="bottom" style="padding-top: 6px;">
                    <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td valign="bottom" style="font-size: 18px; font-weight: 900; color: #15803D; font-family: 'Segoe UI', Arial, sans-serif; line-height: 1;">
                          ${formatCOP(totalValue)}
                        </td>
                        <td align="right" valign="bottom" style="font-size: 17px; color: #86EFAC; opacity: 0.75;">
                          🪙
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>

            <!-- KPI 4: Antigüedad Promedio Global -->
            <td width="24%" valign="top" class="kpi-cell">
              <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 12px; padding: 12px 10px; height: 95px;">
                <tr>
                  <td width="32" valign="top">
                    <div style="background-color: #D97706; width: 30px; height: 30px; border-radius: 50%; text-align: center; line-height: 30px; font-size: 15px; color: #FFFFFF;">
                      🕒
                    </div>
                  </td>
                  <td valign="top" style="padding-left: 6px;">
                    <div style="font-size: 10.5px; color: #475569; font-weight: 600; line-height: 1.2; font-family: 'Segoe UI', Arial, sans-serif;">
                      Antigüedad promedio<br>global
                    </div>
                  </td>
                </tr>
                <tr>
                  <td colspan="2" valign="bottom" style="padding-top: 6px;">
                    <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td valign="bottom" style="font-size: 22px; font-weight: 900; color: #D97706; font-family: 'Segoe UI', Arial, sans-serif; line-height: 1;">
                          ${avgAge} días
                        </td>
                        <td align="right" valign="bottom" style="font-size: 17px; color: #FDBA74; opacity: 0.75;">
                          📅
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>

        <!-- ═════════════════════════════════════════════════════════════════════
             3. TABLA 1: EL RESULTADO DE LOS DIRECTIVOS (CONSOLIDADO POR GRUPOS)
             ═════════════════════════════════════════════════════════════════════ -->
        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="margin-bottom: 8px;">
          <tr>
            <td>
              <span style="font-size: 15.5px; font-weight: 850; color: #0F172A; font-family: 'Segoe UI', Arial, sans-serif;">
                Resultado Consolidado por Directores de Grupo
              </span>
              <span style="color: #1E3A8A; font-size: 14px; margin-left: 4px;">👔</span>
              <span style="font-size: 11.5px; color: #475569; margin-left: 12px; font-family: 'Segoe UI', Arial, sans-serif;">
                Desempeño y volumen pendiente en cada una de las 4 direcciones comerciales.
              </span>
            </td>
          </tr>
        </table>

        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="border: 1.5px solid #E2E8F0; border-radius: 12px; overflow: hidden; border-collapse: separate; border-spacing: 0; margin-bottom: 22px;">
          <thead>
            <tr style="background-color: #0F172A; color: #FFFFFF;">
              <th style="padding: 10px 14px; text-align: left; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">Director & Grupo</th>
              <th style="padding: 10px 10px; text-align: center; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">Asesores Activos</th>
              <th style="padding: 10px 10px; text-align: center; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">Remisiones</th>
              <th style="padding: 10px 14px; text-align: right; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">Valor por Facturar</th>
              <th style="padding: 10px 10px; text-align: center; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">% Empresa</th>
              <th style="padding: 10px 10px; text-align: center; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">Promedio</th>
              <th style="padding: 10px 12px; text-align: left; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">Remisión Crítica</th>
              <th style="padding: 10px 12px; text-align: center; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">Estado</th>
            </tr>
          </thead>
          <tbody style="background-color: #FFFFFF;">
            ${rowsDirectorsHtml}
            ${unassignedSummary && unassignedSummary.totalCount > 0 ? `
            <tr style="background-color: #FAFAFA; border-bottom: 1.5px dashed #CBD5E1;">
              <td style="padding: 10px 14px; font-size: 12.5px; font-weight: 750; color: #475569; font-family: 'Segoe UI', Arial, sans-serif;">
                <span style="display: inline-block; width: 22px; height: 22px; line-height: 22px; text-align: center; border-radius: 6px; background-color: #64748B; color: #FFFFFF; font-size: 10px; margin-right: 6px;">
                  CE
                </span>
                ${unassignedSummary.label}
                <div style="font-size: 10px; font-weight: 500; color: #94A3B8; margin-left: 28px;">
                  Cuentas directas y áreas de apoyo
                </div>
              </td>
              <td style="padding: 10px 10px; text-align: center; font-size: 12px; font-weight: 600; color: #94A3B8; font-family: 'Segoe UI', Arial, sans-serif;">
                —
              </td>
              <td style="padding: 10px 10px; text-align: center; font-size: 12.5px; font-weight: 800; color: #334155; font-family: 'Segoe UI', Arial, sans-serif;">
                ${formatNumber(unassignedSummary.totalCount)}
              </td>
              <td style="padding: 10px 14px; text-align: right; font-size: 13px; font-weight: 850; color: #15803D; font-family: 'Segoe UI', Arial, sans-serif;">
                ${formatCOP(unassignedSummary.totalValue)}
              </td>
              <td style="padding: 10px 10px; text-align: center; font-size: 12px; font-weight: 800; color: #64748B; font-family: 'Segoe UI', Arial, sans-serif;">
                ${unassignedSummary.pctCompanyValue}%
              </td>
              <td style="padding: 10px 10px; text-align: center;">
                <span style="display: inline-block; padding: 2px 7px; border-radius: 6px; font-size: 11px; font-weight: 700; color: #475569; background-color: #F1F5F9; font-family: 'Segoe UI', Arial, sans-serif;">
                  ${unassignedSummary.avgAge} días
                </span>
              </td>
              <td style="padding: 10px 12px; font-size: 11px; color: #64748B; font-family: 'Segoe UI', Arial, sans-serif;">
                ${unassignedSummary.criticalRemision ? `${unassignedSummary.criticalRemision.document?.startsWith('REM-') ? unassignedSummary.criticalRemision.document : `REM-${unassignedSummary.criticalRemision.document || 'S/N'}`} (${unassignedSummary.criticalRemision.age}d)` : '—'}
              </td>
              <td style="padding: 10px 12px; text-align: center;">
                <span style="display: inline-block; padding: 2px 8px; border-radius: 999px; font-size: 10px; font-weight: 700; color: #475569; background-color: #F1F5F9;">
                  En gestión
                </span>
              </td>
            </tr>
            ` : ''}
            <!-- FILA DE TOTALES GENERALES DE LA COMPAÑÍA -->
            <tr style="background-color: #F1F5F9; border-top: 2px solid #CBD5E1;">
              <td style="padding: 12px 14px; font-size: 13px; font-weight: 900; color: #0F172A; font-family: 'Segoe UI', Arial, sans-serif;">
                TOTAL GENERAL COMPAÑÍA
              </td>
              <td style="padding: 12px 10px; text-align: center; font-size: 12.5px; font-weight: 900; color: #0F172A; font-family: 'Segoe UI', Arial, sans-serif;">
                ${activeExecutivesCount} / ${totalExecutivesCount}
              </td>
              <td style="padding: 12px 10px; text-align: center; font-size: 13.5px; font-weight: 900; color: #0F172A; font-family: 'Segoe UI', Arial, sans-serif;">
                ${formatNumber(totalCount)}
              </td>
              <td style="padding: 12px 14px; text-align: right; font-size: 14px; font-weight: 950; color: #15803D; font-family: 'Segoe UI', Arial, sans-serif;">
                ${formatCOP(totalValue)}
              </td>
              <td style="padding: 12px 10px; text-align: center; font-size: 12px; font-weight: 900; color: #1E3A8A; font-family: 'Segoe UI', Arial, sans-serif;">
                100%
              </td>
              <td style="padding: 12px 10px; text-align: center; font-size: 12px; font-weight: 900; color: #D97706; font-family: 'Segoe UI', Arial, sans-serif;">
                ${avgAge} días
              </td>
              <td colspan="2" style="padding: 12px 12px; text-align: center; font-size: 11px; font-weight: 700; color: #475569; font-family: 'Segoe UI', Arial, sans-serif;">
                ${activeDirectorsCount} de ${totalDirectorsCount} grupos comerciales con gestión
              </td>
            </tr>
          </tbody>
        </table>

        <!-- ═════════════════════════════════════════════════════════════════════
             4. TABLA 2: TOP 5 EJECUTIVOS CON MAYOR VALOR PENDIENTE
             ═════════════════════════════════════════════════════════════════════ -->
        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="margin-bottom: 8px;">
          <tr>
            <td>
              <span style="font-size: 14.5px; font-weight: 850; color: #0F172A; font-family: 'Segoe UI', Arial, sans-serif;">
                Top 5 Asesores con Mayor Volumen por Facturar
              </span>
              <span style="color: #15803D; font-size: 14px; margin-left: 4px;">📈</span>
              <span style="font-size: 11.5px; color: #475569; margin-left: 12px; font-family: 'Segoe UI', Arial, sans-serif;">
                Concentración de valor para acompañamiento prioritario de la gerencia comercial.
              </span>
            </td>
          </tr>
        </table>

        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="border: 1px solid #E2E8F0; border-radius: 12px; overflow: hidden; border-collapse: separate; border-spacing: 0; margin-bottom: 22px;">
          <thead>
            <tr style="background-color: #1E293B; color: #FFFFFF;">
              <th style="padding: 9px 12px; text-align: center; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif; width: 45px;">Pos.</th>
              <th style="padding: 9px 12px; text-align: left; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">Asesor Comercial</th>
              <th style="padding: 9px 10px; text-align: center; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif; width: 90px;">Remisiones</th>
              <th style="padding: 9px 12px; text-align: right; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif; width: 140px;">Valor por Facturar</th>
              <th style="padding: 9px 10px; text-align: center; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif; width: 90px;">Promedio</th>
              <th style="padding: 9px 10px; text-align: center; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif; width: 85px;">% Empresa</th>
            </tr>
          </thead>
          <tbody style="background-color: #FFFFFF;">
            ${rowsTopExecsHtml}
          </tbody>
        </table>

        <!-- ═════════════════════════════════════════════════════════════════════
             5. BANNERS ANCHO COMPLETO (100%): MAYOR VALOR Y MAYOR ANTIGÜEDAD
             ═════════════════════════════════════════════════════════════════════ -->
        ${(topMayorValor || topMayorAntiguedad) ? `
        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 14px; padding: 14px 18px; margin-bottom: 14px;">
          ${topMayorValor ? `
          <!-- FILA 1: MAYOR VALOR COMPAÑÍA -->
          <tr>
            <td width="42" valign="middle" align="center" style="padding-right: 12px;">
              <span style="font-size: 26px;">🏆</span>
            </td>
            <td valign="middle" style="padding-right: 14px;">
              <div style="font-size: 10px; font-weight: 700; color: #15803D; text-transform: uppercase; letter-spacing: 0.04em; font-family: 'Segoe UI', Arial, sans-serif;">
                Mayor valor por facturar de la compañía <span style="color: #15803D;">★</span>
              </div>
              <div style="font-size: 15px; font-weight: 900; color: #1E3A8A; font-family: 'Segoe UI', Arial, sans-serif; margin: 1px 0;">
                ${topMayorValor.document?.startsWith('REM-') ? topMayorValor.document : `REM-${topMayorValor.document || 'S/N'}`}
              </div>
              <div style="font-size: 11px; color: #475569; font-family: 'Segoe UI', Arial, sans-serif;">
                ${topMayorValor.company || 'Cliente'}
              </div>
              <div style="font-size: 10.5px; color: #1E3A8A; font-weight: 700; margin-top: 2px;">
                Asesor: ${topMayorValor.executiveName || topMayorValor.employee || 'Comercial'} (${topMayorValor.directorName || 'Dirección'})
              </div>
            </td>
            <td width="140" valign="middle" style="padding-right: 14px;">
              <div style="font-size: 15.5px; font-weight: 900; color: #15803D; font-family: 'Segoe UI', Arial, sans-serif;">
                ${formatCOP(topMayorValor.total)}
              </div>
              <div style="font-size: 10px; color: #64748B; font-family: 'Segoe UI', Arial, sans-serif;">
                Monto más representativo
              </div>
            </td>
            <td width="95" valign="middle" style="padding-right: 14px;">
              <div style="font-size: 14px; font-weight: 850; color: #0F172A; font-family: 'Segoe UI', Arial, sans-serif;">
                ${topMayorValor.age} días
              </div>
              <div style="font-size: 10px; color: #64748B; font-family: 'Segoe UI', Arial, sans-serif;">
                Antigüedad
              </div>
            </td>
            <td valign="middle" align="right">
              <table role="presentation" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td valign="middle" style="padding-right: 4px; font-size: 16px; color: #15803D;">⤷</td>
                  <td valign="middle" style="font-size: 11.5px; font-weight: 850; font-style: italic; color: #15803D; font-family: 'Segoe UI', Arial, sans-serif; line-height: 1.2;">
                    ¡Mayor impacto<br>global!
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          ` : ''}

          ${(topMayorAntiguedad && (!topMayorValor || topMayorAntiguedad.document !== topMayorValor.document)) ? `
          <!-- SEPARADOR ENTRE MAYOR VALOR Y MAYOR ANTIGÜEDAD -->
          <tr>
            <td colspan="5" style="padding: 10px 0;">
              <div style="border-top: 1px dashed #E2E8F0; height: 1px; line-height: 1px; font-size: 0;">&nbsp;</div>
            </td>
          </tr>

          <!-- FILA 2: MAYOR ANTIGÜEDAD COMPAÑÍA -->
          <tr>
            <td width="42" valign="middle" align="center" style="padding-right: 12px;">
              <span style="font-size: 26px;">⏳</span>
            </td>
            <td valign="middle" style="padding-right: 14px;">
              <div style="font-size: 10px; font-weight: 700; color: #DC2626; text-transform: uppercase; letter-spacing: 0.04em; font-family: 'Segoe UI', Arial, sans-serif;">
                Mayor antigüedad pendiente de la compañía <span style="color: #DC2626;">⏱</span>
              </div>
              <div style="font-size: 15px; font-weight: 900; color: #1E3A8A; font-family: 'Segoe UI', Arial, sans-serif; margin: 1px 0;">
                ${topMayorAntiguedad.document?.startsWith('REM-') ? topMayorAntiguedad.document : `REM-${topMayorAntiguedad.document || 'S/N'}`}
              </div>
              <div style="font-size: 11px; color: #475569; font-family: 'Segoe UI', Arial, sans-serif;">
                ${topMayorAntiguedad.company || 'Cliente'}
              </div>
              <div style="font-size: 10.5px; color: #1E3A8A; font-weight: 700; margin-top: 2px;">
                Asesor: ${topMayorAntiguedad.executiveName || topMayorAntiguedad.employee || 'Comercial'} (${topMayorAntiguedad.directorName || 'Dirección'})
              </div>
            </td>
            <td width="140" valign="middle" style="padding-right: 14px;">
              <div style="font-size: 15px; font-weight: 900; color: #1E3A8A; font-family: 'Segoe UI', Arial, sans-serif;">
                ${formatCOP(topMayorAntiguedad.total)}
              </div>
              <div style="font-size: 10px; color: #64748B; font-family: 'Segoe UI', Arial, sans-serif;">
                Valor
              </div>
            </td>
            <td width="95" valign="middle" style="padding-right: 14px;">
              <div style="font-size: 14.5px; font-weight: 900; color: #DC2626; font-family: 'Segoe UI', Arial, sans-serif;">
                ${topMayorAntiguedad.age} días
              </div>
              <div style="font-size: 10px; color: #64748B; font-family: 'Segoe UI', Arial, sans-serif;">
                Mayor antigüedad
              </div>
            </td>
            <td valign="middle" align="right">
              <table role="presentation" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td valign="middle" style="padding-right: 4px; font-size: 16px; color: #DC2626;">⤷</td>
                  <td valign="middle" style="font-size: 11.5px; font-weight: 850; font-style: italic; color: #DC2626; font-family: 'Segoe UI', Arial, sans-serif; line-height: 1.2;">
                    ¡Prioridad crítica<br>por tiempo!
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          ` : ''}
        </table>
        ` : ''}

        <!-- ═════════════════════════════════════════════════════════════════════
             6. BANNER DE ARCHIVO ADJUNTO EXCEL MASTER GERENCIAL (3 PESTAÑAS)
             ═════════════════════════════════════════════════════════════════════ -->
        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #F8FAFC; border: 1.5px solid #E2E8F0; border-radius: 12px; padding: 12px 18px; margin-bottom: 14px;">
          <tr>
            <td width="38" valign="middle" align="center" style="padding-right: 12px;">
              <div style="background-color: #107C41; color: #FFFFFF; width: 32px; height: 32px; border-radius: 8px; text-align: center; line-height: 32px; font-size: 16px; font-weight: 900;">
                📊
              </div>
            </td>
            <td valign="middle">
              <div style="font-size: 12px; font-weight: 800; color: #0F172A; font-family: 'Segoe UI', Arial, sans-serif;">
                📎 Archivo adjunto: Remisiones_Consolidado_General_Gerencia_${cutoffDate}.xlsx
              </div>
              <div style="font-size: 11px; color: #475569; font-family: 'Segoe UI', Arial, sans-serif; margin-top: 2px;">
                Hemos adjuntado el libro Excel maestro con <strong>3 hojas de trabajo</strong>: (1) Resumen por Director, (2) Resumen por Ejecutivo Comercial y (3) Detalle General de las <strong>${formatNumber(totalCount)}</strong> remisiones abiertas (${formatCOP(totalValue)}) de toda la empresa.
              </div>
            </td>
          </tr>
        </table>

        <!-- ═════════════════════════════════════════════════════════════════════
             7. PIE DE PÁGINA: CORAZÓN AZUL + MENSAJE LIDERAZGO + SOBRE POSTAL
             ═════════════════════════════════════════════════════════════════════ -->
        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #F8FAFC; border: 1.5px solid #E2E8F0; border-radius: 14px; padding: 16px 22px; margin-top: 18px;">
          <tr>
            <!-- Corazón azul grande -->
            <td width="38" valign="middle" align="center" style="padding-right: 12px;">
              <div style="font-size: 28px; line-height: 1; color: #1E3A8A;">
                💙
              </div>
            </td>

            <!-- Mensaje de liderazgo y gratitud -->
            <td valign="middle">
              <div style="font-size: 12px; color: #475569; font-family: 'Segoe UI', Arial, sans-serif;">
                Liderazgo, articulación comercial y compromiso corporativo.
              </div>
              <div style="font-size: 13.5px; font-weight: 850; color: #1E3A8A; font-family: 'Segoe UI', Arial, sans-serif; margin-top: 2px;">
                ¡Sigamos impulsando los resultados de Provexpress juntos!
              </div>
            </td>

            <!-- Sobre postal azul con carta y checkmark verde -->
            <td width="55" align="right" valign="middle">
              <svg width="48" height="38" viewBox="0 0 52 42" fill="none" xmlns="http://www.w3.org/2000/svg">
                <rect x="8" y="2" width="36" height="26" rx="3" fill="#FFFFFF" stroke="#60A5FA" stroke-width="1.5"/>
                <line x1="13" y1="8" x2="31" y2="8" stroke="#93C5FD" stroke-width="2" stroke-linecap="round"/>
                <line x1="13" y1="13" x2="27" y2="13" stroke="#93C5FD" stroke-width="2" stroke-linecap="round"/>
                <path d="M4 14H48C49.1046 14 50 14.8954 50 16V36C50 37.1046 49.1046 38 48 38H4C2.89543 38 2 37.1046 2 36V16C2 14.8954 2.89543 14 4 14Z" fill="#2563EB"/>
                <path d="M3 15L26 29L49 15" stroke="#FFFFFF" stroke-width="1.5"/>
                <circle cx="41" cy="31" r="7" fill="#16A34A" stroke="#FFFFFF" stroke-width="1.5"/>
                <path d="M38 31L40 33L44 29" stroke="#FFFFFF" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/>
              </svg>
            </td>
          </tr>
        </table>

      </td>
    </tr>
  </table>

</body>
</html>`;
}
