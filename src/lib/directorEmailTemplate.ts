import type { Remision } from '../types';
import {
  ESTRUCTURA_COMERCIAL_2026,
  getDirectorInfo,
  normalizeName,
  type DirectorInfo,
} from './commercialDirectory';
import { LOGO_PROVEXPRESS_DATA_URI, AVATAR_MAN_DATA_URI, AVATAR_WOMAN_DATA_URI } from './commercialEmailAssets';
import { formatCOP, formatNumber } from './commercialEmailTemplate';

export interface DirectorExecutiveSummary {
  name: string;
  email: string;
  count: number;
  total: number;
  avgAge: number;
  maxAge: number;
  criticalRemision: Remision | null;
  onTrackCount: number;
  remisiones: Remision[];
}

export interface DirectorEmailSummary {
  directorName: string;
  directorEmail: string;
  directorGroup: number;
  cutoffDate: string;
  totalCount: number;
  totalValue: number;
  avgAge: number;
  activeExecutivesCount: number;
  totalExecutivesCount: number;
  topMayorValor: (Remision & { executiveName?: string }) | null;
  topMayorAntiguedad: (Remision & { executiveName?: string }) | null;
  executives: DirectorExecutiveSummary[];
  allGroupRemisiones: Remision[];
  genero: 'M' | 'F';
}

export interface GenerateDirectorEmailOptions {
  forWebPreview?: boolean;
}

export function buildDirectorEmailSummary(
  directorEmailOrGroup: string | number,
  records: Remision[],
  cutoffDate: string,
): DirectorEmailSummary {
  const dirInfo = getDirectorInfo(directorEmailOrGroup) || {
    grupo: typeof directorEmailOrGroup === 'number' ? directorEmailOrGroup : 1,
    nombre: 'Director Comercial',
    email: typeof directorEmailOrGroup === 'string' ? directorEmailOrGroup : '',
    carpeta: '',
    genero: 'M' as const,
  };

  const groupNumber = dirInfo.grupo;

  // Filtrar ejecutivos asignados a este grupo
  const groupExecutivesEntries = Object.entries(ESTRUCTURA_COMERCIAL_2026.ejecutivos).filter(
    ([, data]) => data.grupo === groupNumber,
  );

  // Filtrar remisiones de este corte
  const cutoffRecords = records.filter((r) => !cutoffDate || r.cutoff === cutoffDate);

  const executivesSummary: DirectorExecutiveSummary[] = [];
  const allGroupRemisiones: Remision[] = [];

  for (const [execEmail, execData] of groupExecutivesEntries) {
    const execTarget = normalizeName(execData.nombre);
    const execFileTarget = normalizeName(execData.archivo);

    // Buscar remisiones que correspondan a este ejecutivo
    const execRemisiones = cutoffRecords.filter((r) => {
      const emp = normalizeName(r.employee);
      if (!emp) return false;
      if (emp === execTarget || emp === execFileTarget) return true;
      const empTokens = emp.split(' ').filter(Boolean);
      const targetTokens = execTarget.split(' ').filter(Boolean);
      const common = targetTokens.filter((t) => empTokens.includes(t));
      return common.length >= 2;
    });

    const count = execRemisiones.length;
    const total = execRemisiones.reduce((s, r) => s + (r.total || 0), 0);
    const avgAge = count > 0 ? Math.round(execRemisiones.reduce((s, r) => s + (r.age || 0), 0) / count) : 0;
    const maxAge = count > 0 ? Math.max(...execRemisiones.map((r) => r.age || 0)) : 0;

    // Remisión más crítica (mayor días)
    const sortedByAge = [...execRemisiones].sort((a, b) => (b.age || 0) - (a.age || 0));
    const criticalRemision = sortedByAge.length > 0 ? sortedByAge[0] : null;
    const onTrackCount = execRemisiones.filter((r) => (r.age || 0) <= 15).length;

    executivesSummary.push({
      name: execData.nombre,
      email: execEmail,
      count,
      total,
      avgAge,
      maxAge,
      criticalRemision,
      onTrackCount,
      remisiones: execRemisiones,
    });

    allGroupRemisiones.push(...execRemisiones);
  }

  // Ordenar ejecutivos: primero los que tienen remisiones abiertas (mayor valor desc), luego los que están en 0
  executivesSummary.sort((a, b) => {
    if (a.count === 0 && b.count > 0) return 1;
    if (b.count === 0 && a.count > 0) return -1;
    return b.total - a.total;
  });

  const totalCount = allGroupRemisiones.length;
  const totalValue = allGroupRemisiones.reduce((s, r) => s + (r.total || 0), 0);
  const avgAge = totalCount > 0 ? Math.round(allGroupRemisiones.reduce((s, r) => s + (r.age || 0), 0) / totalCount) : 0;
  const activeExecutivesCount = executivesSummary.filter((e) => e.count > 0).length;
  const totalExecutivesCount = executivesSummary.length;

  // Top Mayor Valor del grupo
  const sortedByValue = [...allGroupRemisiones].sort((a, b) => (b.total || 0) - (a.total || 0));
  const topMayorValor = sortedByValue.length > 0 ? sortedByValue[0] : null;

  // Top Mayor Antigüedad del grupo
  const sortedByAgeAll = [...allGroupRemisiones].sort((a, b) => (b.age || 0) - (a.age || 0));
  const topMayorAntiguedad = sortedByAgeAll.length > 0 ? sortedByAgeAll[0] : null;

  return {
    directorName: dirInfo.nombre,
    directorEmail: dirInfo.email,
    directorGroup: groupNumber,
    cutoffDate,
    totalCount,
    totalValue,
    avgAge,
    activeExecutivesCount,
    totalExecutivesCount,
    topMayorValor,
    topMayorAntiguedad,
    executives: executivesSummary,
    allGroupRemisiones,
    genero: dirInfo.genero || 'M',
  };
}

export function generateDirectorEmailHtml(
  summary: DirectorEmailSummary,
  options: GenerateDirectorEmailOptions = {},
): string {
  const { forWebPreview = false } = options;
  const isFemale = summary.genero === 'F';
  const logoSrc = forWebPreview ? LOGO_PROVEXPRESS_DATA_URI : 'cid:logo_provexpress';
  const avatarSrc = forWebPreview
    ? (isFemale ? AVATAR_WOMAN_DATA_URI : AVATAR_MAN_DATA_URI)
    : (isFemale ? 'cid:avatar_woman' : 'cid:avatar_man');
  const avatarAlt = isFemale ? 'Directora Comercial' : 'Director Comercial';

  const {
    directorName,
    directorGroup,
    cutoffDate,
    totalCount,
    totalValue,
    avgAge,
    activeExecutivesCount,
    totalExecutivesCount,
    topMayorValor,
    topMayorAntiguedad,
    executives,
  } = summary;

  // Formateo de fecha de corte
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

  // Generación de filas de la tabla de ejecutivos
  const rowsExecutivesHtml = executives
    .map((exec, index) => {
      const isEven = index % 2 === 0;
      const bgColor = isEven ? '#FFFFFF' : '#F8FAFC';

      let ageBadgeColor = '#15803D'; // verde
      let ageBadgeBg = '#DCFCE7';
      if (exec.avgAge > 30) {
        ageBadgeColor = '#B91C1C'; // rojo
        ageBadgeBg = '#FEE2E2';
      } else if (exec.avgAge > 15) {
        ageBadgeColor = '#B45309'; // ámbar
        ageBadgeBg = '#FEF3C7';
      }

      let statusBadge = '';
      if (exec.count === 0) {
        statusBadge = `<span style="display: inline-block; padding: 2px 8px; border-radius: 999px; font-size: 10px; font-weight: 700; color: #15803D; background-color: #DCFCE7;">Al día ✓</span>`;
      } else if (exec.maxAge > 30) {
        statusBadge = `<span style="display: inline-block; padding: 2px 8px; border-radius: 999px; font-size: 10px; font-weight: 700; color: #B91C1C; background-color: #FEE2E2;">Urgente (>30d)</span>`;
      } else {
        statusBadge = `<span style="display: inline-block; padding: 2px 8px; border-radius: 999px; font-size: 10px; font-weight: 700; color: #1E3A8A; background-color: #DBEAFE;">En gestión</span>`;
      }

      const criticalDocInfo = exec.criticalRemision
        ? `${exec.criticalRemision.document?.startsWith('REM-') ? exec.criticalRemision.document : `REM-${exec.criticalRemision.document || 'S/N'}`} (${exec.criticalRemision.age}d)`
        : '—';

      return `
      <tr style="background-color: ${bgColor}; border-bottom: 1px solid #E2E8F0;">
        <td style="padding: 10px 14px; font-size: 12.5px; font-weight: 700; color: #1E293B; font-family: 'Segoe UI', Arial, sans-serif;">
          ${exec.name}
          <div style="font-size: 10px; font-weight: 400; color: #64748B;">${exec.email}</div>
        </td>
        <td style="padding: 10px 12px; text-align: center; font-size: 12.5px; font-weight: 800; color: #0F172A; font-family: 'Segoe UI', Arial, sans-serif;">
          ${exec.count > 0 ? formatNumber(exec.count) : '<span style="color: #94A3B8;">0</span>'}
        </td>
        <td style="padding: 10px 14px; text-align: right; font-size: 13px; font-weight: 850; color: ${exec.count > 0 ? '#15803D' : '#94A3B8'}; font-family: 'Segoe UI', Arial, sans-serif;">
          ${formatCOP(exec.total)}
        </td>
        <td style="padding: 10px 12px; text-align: center;">
          ${exec.count > 0 ? `
          <span style="display: inline-block; padding: 3px 9px; border-radius: 6px; font-size: 11px; font-weight: 700; color: ${ageBadgeColor}; background-color: ${ageBadgeBg}; font-family: 'Segoe UI', Arial, sans-serif;">
            ${exec.avgAge} días
          </span>
          ` : '<span style="font-size: 11px; color: #94A3B8;">—</span>'}
        </td>
        <td style="padding: 10px 14px; font-size: 11.5px; color: #475569; font-family: 'Segoe UI', Arial, sans-serif;">
          ${criticalDocInfo}
        </td>
        <td style="padding: 10px 12px; text-align: center;">
          ${statusBadge}
        </td>
      </tr>
      `;
    })
    .join('');

  return `
<!DOCTYPE html>
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <title>Resumen de Remisiones · Dirección Grupo ${directorGroup} · Provexpress</title>
  <!--[if mso]>
  <style type="text/css">
    body, table, td, div, p, a { font-family: Arial, sans-serif !important; }
  </style>
  <![endif]-->
  <style type="text/css">
    body, table, td, a { -webkit-text-size-adjust: 100%; -ms-text-size-adjust: 100%; }
    table, td { mso-table-lspace: 0pt; mso-table-rspace: 0pt; }
    img { -ms-interpolation-mode: bicubic; border: 0; outline: none; text-decoration: none; }
    body { margin: 0; padding: 0; width: 100% !important; background-color: #F1F4F8; font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, Helvetica, Arial, sans-serif; }
    @media only screen and (max-width: 920px) {
      .main-card { width: 100% !important; border-radius: 0 !important; }
      .stack-kpi { display: inline-block !important; width: 48% !important; margin-bottom: 8px !important; }
    }
  </style>
</head>
<body style="margin: 0; padding: 24px 0; background-color: #F1F4F8;">

  <!-- CONTENEDOR PRINCIPAL BLANCO (960px) IDÉNTICO A COMERCIALES -->
  <table role="presentation" class="main-card" width="960" align="center" border="0" cellpadding="0" cellspacing="0" style="width: 960px; margin: 0 auto; background-color: #FFFFFF; border-radius: 16px; overflow: hidden; border: 1.5px solid #E2E8F0; box-shadow: 0 4px 20px rgba(15, 23, 42, 0.05);">
    <tr>
      <td style="padding: 24px 28px;">

        <!-- ═════════════════════════════════════════════════════════════════════
             1. ENCABEZADO: LOGO + SALUDO + MUÑECO (AVATAR) + CARD DIRECCIÓN
             ═════════════════════════════════════════════════════════════════════ -->
        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="margin-bottom: 20px;">
          <tr>
            <!-- LOGO PROVEXPRESS OFICIAL -->
            <td width="160" valign="middle" align="left" style="padding-right: 18px;">
              <img src="${logoSrc}" width="144" height="110" style="display: block; border: 0; width: 144px; height: 110px; max-width: 144px;" alt="Provexpress" />
            </td>

            <!-- TEXTO DE SALUDO Y TITULAR DE LIDERAZGO -->
            <td valign="middle" style="padding-right: 12px;">
              <div style="font-size: 15px; color: #475569; font-weight: 600; font-family: 'Segoe UI', Arial, sans-serif;">
                Hola, <strong style="color: #0F172A; font-size: 16px;">${directorName}</strong> 👋
              </div>
              <h1 style="margin: 3px 0 6px 0; font-size: 24px; font-weight: 900; color: #0F172A; line-height: 1.15; letter-spacing: -0.02em; font-family: 'Segoe UI', Arial, sans-serif;">
                Oportunidades de tu equipo listas para<br>
                convertirse en <span style="color: #16A34A;">ventas facturadas</span>
              </h1>
              <p style="margin: 0; font-size: 12px; color: #475569; line-height: 1.35; font-family: 'Segoe UI', Arial, sans-serif; max-width: 390px;">
                Excelente labor de liderazgo con tu equipo comercial. Te presentamos el consolidado de remisiones abiertas de tus ejecutivos para acompañar, priorizar y acelerar el cierre comercial del grupo.
              </p>
              <p style="margin: 5px 0 0 0; font-size: 12px; color: #0F172A; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">
                El seguimiento oportuno transforma el esfuerzo de tu equipo en <span style="color: #16A34A;">resultados reales</span>.
              </p>
            </td>

            <!-- ILUSTRACIÓN ASESOR COMERCIAL: EL MUÑECO O LA MUÑECA (THUMBS UP) -->
            <td width="190" valign="bottom" align="center" style="padding-right: 10px;">
              <img src="${avatarSrc}" width="180" height="154" style="display: block; border: 0; width: 180px; height: auto; margin: 0 auto;" alt="${avatarAlt}" />
            </td>

            <!-- CARD DIRECCIÓN DE GRUPO -->
            <td width="130" valign="middle" align="right">
              <table role="presentation" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 14px; padding: 14px 10px; width: 126px; text-align: center;">
                <tr>
                  <td align="center">
                    <div style="background-color: #1E3A8A; color: #FFFFFF; width: 28px; height: 28px; border-radius: 50%; font-size: 14px; font-weight: 900; line-height: 28px; margin: 0 auto 8px auto;">
                      👔
                    </div>
                    <div style="font-size: 11.5px; font-weight: 700; color: #1E293B; line-height: 1.3; font-family: 'Segoe UI', Arial, sans-serif;">
                      Tu liderazgo<br>
                      hace la<br>
                      <strong style="color: #1E3A8A; font-size: 12.5px;">diferencia.</strong>
                    </div>
                    <div style="margin-top: 6px; padding-top: 6px; border-top: 1px dashed #E2E8F0; font-size: 9.5px; font-weight: 800; color: #2563EB; text-transform: uppercase;">
                      Grupo ${directorGroup}
                    </div>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>

        <!-- ═════════════════════════════════════════════════════════════════════
             2. FILA DE 4 TARJETAS KPI CON EL ESTILO EXACTO DE COMERCIALES
             ═════════════════════════════════════════════════════════════════════ -->
        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="margin-bottom: 20px;">
          <tr>
            <!-- KPI 1: Ejecutivos con pendientes -->
            <td width="24%" valign="top" class="stack-kpi" style="padding-right: 8px;">
              <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 12px; padding: 12px 10px; height: 95px;">
                <tr>
                  <td width="32" valign="top">
                    <div style="background-color: #312E81; width: 30px; height: 30px; border-radius: 50%; text-align: center; line-height: 30px; font-size: 15px; color: #FFFFFF;">
                      👥
                    </div>
                  </td>
                  <td valign="top" style="padding-left: 6px;">
                    <div style="font-size: 10.5px; color: #475569; font-weight: 600; line-height: 1.2; font-family: 'Segoe UI', Arial, sans-serif;">
                      Ejecutivos con<br>remisiones activas
                    </div>
                  </td>
                </tr>
                <tr>
                  <td colspan="2" valign="bottom" style="padding-top: 6px;">
                    <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td valign="bottom" style="font-size: 22px; font-weight: 900; color: #312E81; font-family: 'Segoe UI', Arial, sans-serif; line-height: 1;">
                          ${activeExecutivesCount} / ${totalExecutivesCount}
                        </td>
                        <td align="right" valign="bottom" style="font-size: 16px; color: #94A3B8; opacity: 0.65;">
                          📋
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>

            <!-- KPI 2: Total Remisiones del Grupo -->
            <td width="24%" valign="top" class="stack-kpi" style="padding-right: 8px;">
              <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 12px; padding: 12px 10px; height: 95px;">
                <tr>
                  <td width="32" valign="top">
                    <div style="background-color: #1E3A8A; width: 30px; height: 30px; border-radius: 50%; text-align: center; line-height: 30px; font-size: 15px; color: #FFFFFF;">
                      📦
                    </div>
                  </td>
                  <td valign="top" style="padding-left: 6px;">
                    <div style="font-size: 10.5px; color: #475569; font-weight: 600; line-height: 1.2; font-family: 'Segoe UI', Arial, sans-serif;">
                      Remisiones listas<br>en el equipo
                    </div>
                  </td>
                </tr>
                <tr>
                  <td colspan="2" valign="bottom" style="padding-top: 6px;">
                    <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td valign="bottom" style="font-size: 22px; font-weight: 900; color: #1E3A8A; font-family: 'Segoe UI', Arial, sans-serif; line-height: 1;">
                          ${formatNumber(totalCount)}
                        </td>
                        <td align="right" valign="bottom" style="font-size: 16px; color: #94A3B8; opacity: 0.65;">
                          📝
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>

            <!-- KPI 3: Valor por facturar del grupo -->
            <td width="28%" valign="top" class="stack-kpi" style="padding-right: 8px;">
              <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 12px; padding: 12px 10px; height: 95px;">
                <tr>
                  <td width="32" valign="top">
                    <div style="background-color: #15803D; width: 30px; height: 30px; border-radius: 50%; text-align: center; line-height: 30px; font-size: 15px; font-weight: 900; color: #FFFFFF;">
                      $
                    </div>
                  </td>
                  <td valign="top" style="padding-left: 6px;">
                    <div style="font-size: 10.5px; color: #475569; font-weight: 600; line-height: 1.2; font-family: 'Segoe UI', Arial, sans-serif;">
                      Valor por facturar<br>del equipo
                    </div>
                  </td>
                </tr>
                <tr>
                  <td colspan="2" valign="bottom" style="padding-top: 6px;">
                    <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td valign="bottom" style="font-size: 17px; font-weight: 900; color: #15803D; font-family: 'Segoe UI', Arial, sans-serif; line-height: 1; white-space: nowrap;">
                          ${formatCOP(totalValue)}
                        </td>
                        <td align="right" valign="bottom" style="font-size: 16px; color: #86EFAC; opacity: 0.75;">
                          🪙
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>

            <!-- KPI 4: Antigüedad promedio del grupo -->
            <td width="24%" valign="top" class="stack-kpi">
              <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 12px; padding: 12px 10px; height: 95px;">
                <tr>
                  <td width="32" valign="top">
                    <div style="background-color: #D97706; width: 30px; height: 30px; border-radius: 50%; text-align: center; line-height: 30px; font-size: 15px; color: #FFFFFF;">
                      🕒
                    </div>
                  </td>
                  <td valign="top" style="padding-left: 6px;">
                    <div style="font-size: 10.5px; color: #475569; font-weight: 600; line-height: 1.2; font-family: 'Segoe UI', Arial, sans-serif;">
                      Antigüedad<br>promedio equipo
                    </div>
                  </td>
                </tr>
                <tr>
                  <td colspan="2" valign="bottom" style="padding-top: 6px;">
                    <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td valign="bottom" style="font-size: 22px; font-weight: 900; color: #D97706; font-family: 'Segoe UI', Arial, sans-serif; line-height: 1;">
                          ${avgAge} <span style="font-size: 12px; font-weight: 700;">días</span>
                        </td>
                        <td align="right" valign="bottom" style="font-size: 16px; color: #FDBA74; opacity: 0.75;">
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
             3. TABLA DE DESEMPEÑO POR ASESOR COMERCIAL
             ═════════════════════════════════════════════════════════════════════ -->
        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="margin-bottom: 8px;">
          <tr>
            <td valign="middle">
              <span style="font-size: 14.5px; font-weight: 850; color: #1E293B; font-family: 'Segoe UI', Arial, sans-serif;">
                📋 Detalle y Desempeño por Asesor Comercial
              </span>
              <span style="font-size: 13px; color: #D97706; margin: 0 4px;">☆</span>
              <span style="font-size: 11.5px; color: #64748B; font-family: 'Segoe UI', Arial, sans-serif;">
                Enfócate en acompañar los casos de mayor valor y mayor antigüedad.
              </span>
            </td>
            <td align="right" valign="middle">
              <span style="font-size: 11px; font-weight: 700; color: #64748B; font-family: 'Segoe UI', Arial, sans-serif;">
                Corte: ${formattedCutoff}
              </span>
            </td>
          </tr>
        </table>

        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="border: 1px solid #E2E8F0; border-radius: 12px; overflow: hidden; border-collapse: separate; border-spacing: 0; margin-bottom: 18px;">
          <thead>
            <tr style="background-color: #1E293B; color: #FFFFFF;">
              <th style="padding: 11px 14px; text-align: left; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">Ejecutivo Comercial</th>
              <th style="padding: 11px 12px; text-align: center; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">Remisiones</th>
              <th style="padding: 11px 14px; text-align: right; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">Valor Total</th>
              <th style="padding: 11px 12px; text-align: center; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">Días Prom.</th>
              <th style="padding: 11px 14px; text-align: left; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">Remisión Crítica</th>
              <th style="padding: 11px 12px; text-align: center; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">Estado</th>
            </tr>
          </thead>
          <tbody>
            ${rowsExecutivesHtml}
          </tbody>
        </table>

        <!-- ═════════════════════════════════════════════════════════════════════
             4. HERO BANNER: TOP OPORTUNIDADES DEL EQUIPO (FULL WIDTH)
             ═════════════════════════════════════════════════════════════════════ -->
        ${(topMayorValor || topMayorAntiguedad) ? `
        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 12px; padding: 14px 18px; margin-bottom: 14px;">
          <!-- FILA 1: MAYOR VALOR DEL EQUIPO -->
          ${topMayorValor ? `
          <tr>
            <td width="42" valign="middle" align="center" style="padding-right: 12px;">
              <span style="font-size: 26px;">🏆</span>
            </td>
            <td valign="middle" style="padding-right: 14px;">
              <div style="font-size: 10px; font-weight: 700; color: #15803D; text-transform: uppercase; letter-spacing: 0.04em; font-family: 'Segoe UI', Arial, sans-serif;">
                Mayor valor por facturar del equipo <span style="color: #D97706;">☆</span>
              </div>
              <div style="font-size: 15px; font-weight: 900; color: #1E3A8A; font-family: 'Segoe UI', Arial, sans-serif; margin: 1px 0;">
                ${topMayorValor.document?.startsWith('REM-') ? topMayorValor.document : `REM-${topMayorValor.document || 'S/N'}`}
              </div>
              <div style="font-size: 11px; color: #475569; font-family: 'Segoe UI', Arial, sans-serif; max-width: 250px;">
                ${topMayorValor.company || 'Cliente'} ${topMayorValor.employee ? `· <em>${topMayorValor.employee}</em>` : ''}
              </div>
            </td>
            <td width="130" valign="middle" style="padding-right: 14px;">
              <div style="font-size: 15px; font-weight: 900; color: #15803D; font-family: 'Segoe UI', Arial, sans-serif;">
                ${formatCOP(topMayorValor.total)}
              </div>
              <div style="font-size: 10px; color: #64748B; font-family: 'Segoe UI', Arial, sans-serif;">
                Mayor valor
              </div>
            </td>
            <td width="95" valign="middle" style="padding-right: 14px;">
              <div style="font-size: 14.5px; font-weight: 900; color: #0F172A; font-family: 'Segoe UI', Arial, sans-serif;">
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
                    ¡Mayor impacto<br>en la meta!
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          ` : ''}

          ${(topMayorAntiguedad && (!topMayorValor || topMayorAntiguedad.document !== topMayorValor.document)) ? `
          <!-- SEPARADOR -->
          <tr>
            <td colspan="5" style="padding: 10px 0;">
              <div style="border-top: 1px dashed #E2E8F0; height: 1px; line-height: 1px; font-size: 0;">&nbsp;</div>
            </td>
          </tr>

          <!-- FILA 2: MAYOR ANTIGÜEDAD DEL EQUIPO -->
          <tr>
            <td width="42" valign="middle" align="center" style="padding-right: 12px;">
              <span style="font-size: 26px;">⏳</span>
            </td>
            <td valign="middle" style="padding-right: 14px;">
              <div style="font-size: 10px; font-weight: 700; color: #DC2626; text-transform: uppercase; letter-spacing: 0.04em; font-family: 'Segoe UI', Arial, sans-serif;">
                Mayor antigüedad pendiente del equipo <span style="color: #DC2626;">⏱</span>
              </div>
              <div style="font-size: 15px; font-weight: 900; color: #1E3A8A; font-family: 'Segoe UI', Arial, sans-serif; margin: 1px 0;">
                ${topMayorAntiguedad.document?.startsWith('REM-') ? topMayorAntiguedad.document : `REM-${topMayorAntiguedad.document || 'S/N'}`}
              </div>
              <div style="font-size: 11px; color: #475569; font-family: 'Segoe UI', Arial, sans-serif; max-width: 250px;">
                ${topMayorAntiguedad.company || 'Cliente'} ${topMayorAntiguedad.employee ? `· <em>${topMayorAntiguedad.employee}</em>` : ''}
              </div>
            </td>
            <td width="130" valign="middle" style="padding-right: 14px;">
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
                    ¡Prioridad de gestión<br>inmediata!
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          ` : ''}
        </table>
        ` : ''}

        <!-- ═════════════════════════════════════════════════════════════════════
             5. BANNER DE ARCHIVO ADJUNTO EXCEL CONSOLIDADO
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
                📎 Archivo adjunto: Remisiones_Consolidado_Grupo_${directorGroup}_${cutoffDate || 'Corte'}.xlsx
              </div>
              <div style="font-size: 11px; color: #475569; font-family: 'Segoe UI', Arial, sans-serif; margin-top: 2px;">
                Incluye <strong>Hoja 1: Resumen por Ejecutivo</strong> (totales, promedio de días y % de participación) y <strong>Hoja 2: Detalle Completo</strong> con las <strong>${formatNumber(totalCount)}</strong> remisiones abiertas (${formatCOP(totalValue)}) de tu equipo comercial.
              </div>
            </td>
          </tr>
        </table>

        <!-- ═════════════════════════════════════════════════════════════════════
             6. PIE DE PÁGINA: CORAZÓN AZUL + FIRMA + SOBRE POSTAL (FULL WIDTH)
             ═════════════════════════════════════════════════════════════════════ -->
        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #F8FAFC; border: 1.5px solid #E2E8F0; border-radius: 14px; padding: 16px 22px;">
          <tr>
            <!-- Corazón azul grande -->
            <td width="38" valign="middle" align="center" style="padding-right: 12px;">
              <div style="font-size: 28px; line-height: 1; color: #1E3A8A;">
                💙
              </div>
            </td>

            <!-- Mensaje de agradecimiento y liderazgo -->
            <td valign="middle">
              <div style="font-size: 12px; color: #475569; font-family: 'Segoe UI', Arial, sans-serif;">
                Gracias por tu liderazgo y gestión continua.
              </div>
              <div style="font-size: 13.5px; font-weight: 850; color: #1E3A8A; font-family: 'Segoe UI', Arial, sans-serif; margin-top: 2px;">
                ¡Acompañando a nuestros equipos alcanzamos las metas!
              </div>
            </td>

            <!-- Sobre postal azul -->
            <td width="55" align="right" valign="middle">
              <svg width="48" height="38" viewBox="0 0 52 42" fill="none" xmlns="http://www.w3.org/2000/svg">
                <rect x="8" y="2" width="36" height="26" rx="3" fill="#FFFFFF" stroke="#60A5FA" stroke-width="1.5"/>
                <line x1="13" y1="8" x2="31" y2="8" stroke="#93C5FD" stroke-width="2" stroke-linecap="round"/>
                <line x1="13" y1="13" x2="27" y2="13" stroke="#93C5FD" stroke-width="2" stroke-linecap="round"/>
                <rect x="2" y="10" width="48" height="30" rx="4" fill="#1E3A8A"/>
                <path d="M2 12L26 28L50 12" stroke="#3B82F6" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
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
</html>
  `.trim();
}
