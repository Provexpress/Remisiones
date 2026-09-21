import type { Remision } from '../types';
import {
  ESTRUCTURA_COMERCIAL_2026,
  getDirectorInfo,
  normalizeName,
  type DirectorInfo,
} from './commercialDirectory';
import { LOGO_PROVEXPRESS_DATA_URI } from './commercialEmailAssets';
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
  };
}

export function generateDirectorEmailHtml(
  summary: DirectorEmailSummary,
  options: GenerateDirectorEmailOptions = {},
): string {
  const { forWebPreview = false } = options;
  const logoSrc = forWebPreview ? LOGO_PROVEXPRESS_DATA_URI : 'cid:logo_provexpress';

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
        <td style="padding: 10px 12px; font-size: 12px; font-weight: 700; color: #1E293B; font-family: 'Segoe UI', Arial, sans-serif;">
          ${exec.name}
          <div style="font-size: 10px; font-weight: 400; color: #64748B;">${exec.email}</div>
        </td>
        <td style="padding: 10px 12px; text-align: center; font-size: 12px; font-weight: 700; color: #0F172A; font-family: 'Segoe UI', Arial, sans-serif;">
          ${exec.count > 0 ? formatNumber(exec.count) : '<span style="color: #94A3B8;">0</span>'}
        </td>
        <td style="padding: 10px 12px; text-align: right; font-size: 12.5px; font-weight: 800; color: ${exec.count > 0 ? '#15803D' : '#94A3B8'}; font-family: 'Segoe UI', Arial, sans-serif;">
          ${formatCOP(exec.total)}
        </td>
        <td style="padding: 10px 12px; text-align: center;">
          ${exec.count > 0 ? `
          <span style="display: inline-block; padding: 2px 8px; border-radius: 6px; font-size: 11px; font-weight: 700; color: ${ageBadgeColor}; background-color: ${ageBadgeBg}; font-family: 'Segoe UI', Arial, sans-serif;">
            ${exec.avgAge} días
          </span>
          ` : '<span style="font-size: 11px; color: #94A3B8;">—</span>'}
        </td>
        <td style="padding: 10px 12px; font-size: 11px; color: #475569; font-family: 'Segoe UI', Arial, sans-serif;">
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
    body {
      margin: 0;
      padding: 0;
      background-color: #F1F5F9;
      font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, Helvetica, Arial, sans-serif;
      -webkit-font-smoothing: antialiased;
      -webkit-text-size-adjust: 100%;
      -ms-text-size-adjust: 100%;
      color: #0F172A;
    }
    table {
      border-collapse: collapse;
      mso-table-lspace: 0pt;
      mso-table-rspace: 0pt;
    }
    img {
      border: 0;
      line-height: 100%;
      outline: none;
      text-decoration: none;
    }
    @media only screen and (max-width: 680px) {
      .container {
        width: 100% !important;
        max-width: 100% !important;
      }
      .stack-kpi {
        display: block !important;
        width: 100% !important;
        margin-bottom: 8px !important;
      }
    }
  </style>
</head>
<body style="margin: 0; padding: 16px 0; background-color: #F1F5F9;">

  <center>
    <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #F1F5F9;">
      <tr>
        <td align="center" style="padding: 12px;">

          <!-- TARJETA CONTENEDOR PRINCIPAL (MAX 680PX) -->
          <table role="presentation" class="container" width="680" border="0" cellpadding="0" cellspacing="0" style="width: 680px; max-width: 680px; background-color: #FFFFFF; border-radius: 18px; overflow: hidden; box-shadow: 0 4px 20px rgba(0, 0, 0, 0.06); border: 1px solid #E2E8F0; padding: 24px 28px;">
            <tr>
              <td>

                <!-- ═════════════════════════════════════════════════════════════
                     1. ENCABEZADO EJECUTIVO PARA DIRECTOR DE GRUPO
                     ═════════════════════════════════════════════════════════════ -->
                <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="margin-bottom: 20px;">
                  <tr>
                    <!-- LOGO PROVEXPRESS -->
                    <td width="200" valign="middle" align="left">
                      <img src="${logoSrc}" alt="Provexpress" width="180" style="display: block; width: 180px; max-width: 180px; height: auto;">
                    </td>

                    <!-- INSIGNIA DE DIRECCIÓN -->
                    <td valign="middle" align="right">
                      <table role="presentation" border="0" cellpadding="0" cellspacing="0" style="background-color: #EFF6FF; border: 1px solid #BFDBFE; border-radius: 10px; padding: 6px 14px;">
                        <tr>
                          <td style="font-size: 11px; font-weight: 800; color: #1E3A8A; font-family: 'Segoe UI', Arial, sans-serif; text-transform: uppercase; letter-spacing: 0.05em;">
                            👔 Dirección Grupo ${directorGroup}
                          </td>
                        </tr>
                        <tr>
                          <td style="font-size: 10px; color: #64748B; font-family: 'Segoe UI', Arial, sans-serif; text-align: right;">
                            Corte: ${formattedCutoff}
                          </td>
                        </tr>
                      </table>
                    </td>
                  </tr>
                </table>

                <!-- BANNER DE BIENVENIDA Y RESUMEN DIRECTIVO -->
                <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background: linear-gradient(135deg, #1E3A8A 0%, #1E40AF 100%); background-color: #1E3A8A; border-radius: 14px; padding: 18px 22px; margin-bottom: 20px; color: #FFFFFF;">
                  <tr>
                    <td>
                      <div style="font-size: 11px; font-weight: 700; color: #93C5FD; text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 4px; font-family: 'Segoe UI', Arial, sans-serif;">
                        Reporte Consolidado de Gestión Comercial
                      </div>
                      <div style="font-size: 20px; font-weight: 900; color: #FFFFFF; font-family: 'Segoe UI', Arial, sans-serif; margin-bottom: 6px;">
                        Hola, ${directorName} 👋
                      </div>
                      <div style="font-size: 12.5px; color: #E0E7FF; line-height: 1.4; font-family: 'Segoe UI', Arial, sans-serif;">
                        Te compartimos el estado de remisiones abiertas de los ejecutivos de tu equipo comercial para acompañar, priorizar y acelerar el reconocimiento y facturación oportuna de ventas.
                      </div>
                    </td>
                  </tr>
                </table>

                <!-- ═════════════════════════════════════════════════════════════
                     2. TARJETAS KPI CONSOLIDADAS DEL GRUPO (4 CARDS)
                     ═════════════════════════════════════════════════════════════ -->
                <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="margin-bottom: 22px;">
                  <tr>
                    <!-- KPI 1: Ejecutivos con pendientes -->
                    <td width="24%" valign="top" class="stack-kpi" style="padding-right: 8px;">
                      <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #F8FAFC; border: 1.5px solid #E2E8F0; border-radius: 12px; padding: 12px 10px;">
                        <tr>
                          <td>
                            <div style="font-size: 10px; font-weight: 800; color: #64748B; text-transform: uppercase; letter-spacing: 0.04em;">👥 Ejecutivos</div>
                            <div style="font-size: 19px; font-weight: 900; color: #1E3A8A; margin-top: 4px;">${activeExecutivesCount} / ${totalExecutivesCount}</div>
                            <div style="font-size: 10px; color: #64748B; margin-top: 2px;">Con remisiones activas</div>
                          </td>
                        </tr>
                      </table>
                    </td>

                    <!-- KPI 2: Total Remisiones -->
                    <td width="24%" valign="top" class="stack-kpi" style="padding-right: 8px;">
                      <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #F8FAFC; border: 1.5px solid #E2E8F0; border-radius: 12px; padding: 12px 10px;">
                        <tr>
                          <td>
                            <div style="font-size: 10px; font-weight: 800; color: #64748B; text-transform: uppercase; letter-spacing: 0.04em;">📦 Remisiones</div>
                            <div style="font-size: 19px; font-weight: 900; color: #0F172A; margin-top: 4px;">${formatNumber(totalCount)}</div>
                            <div style="font-size: 10px; color: #64748B; margin-top: 2px;">Por facturar en grupo</div>
                          </td>
                        </tr>
                      </table>
                    </td>

                    <!-- KPI 3: Valor por facturar -->
                    <td width="28%" valign="top" class="stack-kpi" style="padding-right: 8px;">
                      <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #F8FAFC; border: 1.5px solid #E2E8F0; border-radius: 12px; padding: 12px 10px;">
                        <tr>
                          <td>
                            <div style="font-size: 10px; font-weight: 800; color: #15803D; text-transform: uppercase; letter-spacing: 0.04em;">💵 Valor Grupo</div>
                            <div style="font-size: 17px; font-weight: 900; color: #15803D; margin-top: 4px; white-space: nowrap;">${formatCOP(totalValue)}</div>
                            <div style="font-size: 10px; color: #64748B; margin-top: 2px;">Impacto en ventas</div>
                          </td>
                        </tr>
                      </table>
                    </td>

                    <!-- KPI 4: Antigüedad Promedio -->
                    <td width="24%" valign="top" class="stack-kpi">
                      <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #F8FAFC; border: 1.5px solid #E2E8F0; border-radius: 12px; padding: 12px 10px;">
                        <tr>
                          <td>
                            <div style="font-size: 10px; font-weight: 800; color: #DC2626; text-transform: uppercase; letter-spacing: 0.04em;">⏱ Antigüedad</div>
                            <div style="font-size: 19px; font-weight: 900; color: #DC2626; margin-top: 4px;">${avgAge} <span style="font-size: 11px; font-weight: 700;">días</span></div>
                            <div style="font-size: 10px; color: #64748B; margin-top: 2px;">Promedio del equipo</div>
                          </td>
                        </tr>
                      </table>
                    </td>
                  </tr>
                </table>

                <!-- ═════════════════════════════════════════════════════════════
                     3. TABLA CONSOLIDADA POR EJECUTIVO COMERCIAL
                     ═════════════════════════════════════════════════════════════ -->
                <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="margin-bottom: 8px;">
                  <tr>
                    <td>
                      <div style="font-size: 14px; font-weight: 850; color: #1E293B; font-family: 'Segoe UI', Arial, sans-serif;">
                        📋 Detalle y Desempeño por Asesor Comercial
                      </div>
                      <div style="font-size: 11px; color: #64748B; font-family: 'Segoe UI', Arial, sans-serif; margin-top: 2px;">
                        Ordenado por mayor valor pendiente por facturar para enfocar esfuerzos de cierre.
                      </div>
                    </td>
                  </tr>
                </table>

                <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="border: 1px solid #E2E8F0; border-radius: 12px; overflow: hidden; border-collapse: separate; border-spacing: 0; margin-bottom: 18px;">
                  <thead>
                    <tr style="background-color: #1E293B; color: #FFFFFF;">
                      <th style="padding: 10px 12px; text-align: left; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">Ejecutivo Comercial</th>
                      <th style="padding: 10px 12px; text-align: center; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">Remisiones</th>
                      <th style="padding: 10px 12px; text-align: right; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">Valor Total</th>
                      <th style="padding: 10px 12px; text-align: center; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">Días Prom.</th>
                      <th style="padding: 10px 12px; text-align: left; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">Remisión Crítica</th>
                      <th style="padding: 10px 12px; text-align: center; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">Estado</th>
                    </tr>
                  </thead>
                  <tbody>
                    ${rowsExecutivesHtml}
                  </tbody>
                </table>

                <!-- ═════════════════════════════════════════════════════════════
                     4. TOP OPORTUNIDADES CLAVE DEL GRUPO (ANCHO COMPLETO)
                     ═════════════════════════════════════════════════════════════ -->
                ${(topMayorValor || topMayorAntiguedad) ? `
                <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 12px; padding: 14px 18px; margin-bottom: 14px;">
                  <!-- FILA 1: MAYOR VALOR DEL GRUPO -->
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
                      <div style="font-size: 11px; color: #475569; font-family: 'Segoe UI', Arial, sans-serif; max-width: 220px;">
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
                  <!-- SEPARADOR ENTRE MAYOR VALOR Y MAYOR ANTIGÜEDAD -->
                  <tr>
                    <td colspan="5" style="padding: 10px 0;">
                      <div style="border-top: 1px dashed #E2E8F0; height: 1px; line-height: 1px; font-size: 0;">&nbsp;</div>
                    </td>
                  </tr>

                  <!-- FILA 2: MAYOR ANTIGÜEDAD DEL GRUPO -->
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
                      <div style="font-size: 11px; color: #475569; font-family: 'Segoe UI', Arial, sans-serif; max-width: 220px;">
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

                <!-- ═════════════════════════════════════════════════════════════
                     5. BANNER DE ARCHIVO ADJUNTO EXCEL CONSOLIDADO
                     ═════════════════════════════════════════════════════════════ -->
                <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #F8FAFC; border: 1.5px solid #E2E8F0; border-radius: 12px; padding: 12px 18px; margin-bottom: 14px;">
                  <tr>
                    <td width="38" valign="middle" align="center" style="padding-right: 12px;">
                      <div style="background-color: #1E3A8A; color: #FFFFFF; width: 32px; height: 32px; border-radius: 8px; text-align: center; line-height: 32px; font-size: 16px; font-weight: 900;">
                        📊
                      </div>
                    </td>
                    <td valign="middle">
                      <div style="font-size: 12px; font-weight: 800; color: #0F172A; font-family: 'Segoe UI', Arial, sans-serif;">
                        📎 Archivo adjunto: Remisiones_Consolidado_Grupo_${directorGroup}_${cutoffDate || 'Corte'}.xlsx
                      </div>
                      <div style="font-size: 11px; color: #475569; font-family: 'Segoe UI', Arial, sans-serif; margin-top: 2px;">
                        Incluye <strong>Hoja 1: Resumen por Ejecutivo</strong> (totales, promedio de días y participación) y <strong>Hoja 2: Detalle Completo</strong> con las <strong>${formatNumber(totalCount)}</strong> remisiones abiertas (${formatCOP(totalValue)}) de tu equipo comercial.
                      </div>
                    </td>
                  </tr>
                </table>

                <!-- ═════════════════════════════════════════════════════════════
                     6. PIE DE PÁGINA: CORAZÓN AZUL + FIRMA INSTITUCIONAL
                     ═════════════════════════════════════════════════════════════ -->
                <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #F8FAFC; border: 1.5px solid #E2E8F0; border-radius: 14px; padding: 16px 22px;">
                  <tr>
                    <td width="38" valign="middle" align="center" style="padding-right: 12px;">
                      <div style="font-size: 28px; line-height: 1; color: #1E3A8A;">
                        💙
                      </div>
                    </td>
                    <td valign="middle">
                      <div style="font-size: 12px; color: #475569; font-family: 'Segoe UI', Arial, sans-serif;">
                        Gracias por tu liderazgo y gestión continua.
                      </div>
                      <div style="font-size: 13.5px; font-weight: 850; color: #1E3A8A; font-family: 'Segoe UI', Arial, sans-serif; margin-top: 2px;">
                        ¡Acompañando a nuestros equipos alcanzamos las metas!
                      </div>
                    </td>
                    <td width="20" align="center" valign="middle" style="color: #CBD5E1; font-size: 22px; font-weight: 300;">
                      |
                    </td>
                    <td valign="middle" style="padding-left: 10px;">
                      <div style="font-size: 11.5px; color: #475569; font-family: 'Segoe UI', Arial, sans-serif;">
                        Cordialmente,
                      </div>
                      <div style="font-size: 11.5px; font-weight: 600; color: #334155; font-family: 'Segoe UI', Arial, sans-serif; margin: 1px 0;">
                        Gerencia Administrativa y Financiera
                      </div>
                      <div style="font-size: 12px; font-weight: 850; color: #2563EB; font-family: 'Segoe UI', Arial, sans-serif;">
                        PROVEXPRESS SAS
                      </div>
                    </td>
                  </tr>
                </table>

              </td>
            </tr>
          </table>

        </td>
      </tr>
    </table>
  </center>

</body>
</html>
  `.trim();
}
