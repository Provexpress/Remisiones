import type { Remision } from '../types';
import { getCommercialInfo } from './commercialDirectory';

export interface CommercialEmailData {
  commercialName: string;
  commercialEmail: string;
  directorName: string;
  cutoffDate: string; // ej. 2026-09-16
  remisiones: Remision[];
}

export function formatCOP(amount: number): string {
  if (amount == null || Number.isNaN(amount)) return '$ 0';
  return new Intl.NumberFormat('es-CO', {
    style: 'currency',
    currency: 'COP',
    maximumFractionDigits: 0,
  }).format(amount).replace('COP', '').replace(/\u00A0/g, ' ').trim();
}

export function formatNumber(num: number): string {
  if (num == null || Number.isNaN(num)) return '0';
  return new Intl.NumberFormat('es-CO').format(Math.round(num));
}

function formatDateReadable(iso: string): string {
  if (!iso) return '';
  const [y, m, d] = iso.split('-').map(Number);
  if (!y || !m || !d) return iso;
  const date = new Date(y, m - 1, d);
  return date.toLocaleDateString('es-CO', { day: 'numeric', month: 'long', year: 'numeric' });
}

export function buildCommercialEmailSummary(
  commercialName: string,
  records: Remision[],
  cutoffDate: string,
): {
  commercialName: string;
  commercialEmail: string;
  directorName: string;
  totalCount: number;
  totalValue: number;
  avgAge: number;
  destacadas: Remision[];
  topOpportunity: Remision | null;
  progressPct: number;
  cutoffDate: string;
} {
  const info = getCommercialInfo(commercialName);
  const userRemisiones = records.filter((r) => r.cutoff === cutoffDate);

  const totalCount = userRemisiones.length;
  const totalValue = userRemisiones.reduce((sum, r) => sum + (r.total || 0), 0);
  const avgAge = totalCount > 0
    ? Math.round(userRemisiones.reduce((sum, r) => sum + (r.age || 0), 0) / totalCount)
    : 0;

  // Ordenadas por mayor antigüedad y luego por mayor valor
  const sorted = [...userRemisiones].sort((a, b) => {
    if (b.age !== a.age) return b.age - a.age;
    return b.total - a.total;
  });

  const destacadas = sorted.slice(0, 5);
  const topOpportunity = sorted[0] || null;

  // Porcentaje de remisiones con gestión reciente (< 15 días) vs total
  const onTrackCount = userRemisiones.filter((r) => r.age <= 15).length;
  const progressPct = totalCount > 0 ? Math.round((onTrackCount / totalCount) * 100) : 100;

  return {
    commercialName: info.nombre || commercialName,
    commercialEmail: info.email,
    directorName: info.directorNombre,
    totalCount,
    totalValue,
    avgAge,
    destacadas,
    topOpportunity,
    progressPct,
    cutoffDate,
  };
}

/**
 * Genera la plantilla HTML corporativa de notificación diaria por correo.
 * Compatible con Outlook Desktop (Word engine), Outlook Web, Gmail y dispositivos móviles.
 */
export function generateCommercialEmailHtml(summary: ReturnType<typeof buildCommercialEmailSummary>): string {
  const {
    commercialName,
    directorName,
    totalCount,
    totalValue,
    avgAge,
    destacadas,
    topOpportunity,
    progressPct,
    cutoffDate,
  } = summary;

  const formattedCutoff = formatDateReadable(cutoffDate);

  const rowsDestacadasHtml = destacadas.length === 0
    ? `<tr><td colspan="5" style="padding: 20px; text-align: center; color: #64748B; font-size: 13px;">¡Felicitaciones! No tienes remisiones pendientes por facturar en este corte.</td></tr>`
    : destacadas.map((r) => {
      let badgeBg = '#DCFCE7';
      let badgeText = '#15803D';
      let badgeLabel = '0 a 7 días';
      if (r.age > 15) {
        badgeBg = '#FEE2E2';
        badgeText = '#B91C1C';
        badgeLabel = 'Más de 15 días';
      } else if (r.age > 7) {
        badgeBg = '#FEF3C7';
        badgeText = '#B45309';
        badgeLabel = '8 a 15 días';
      }

      return `
        <tr style="border-bottom: 1px solid #E2E8F0;">
          <td style="padding: 12px 10px; font-family: 'Segoe UI', Arial, sans-serif;">
            <div style="font-weight: 800; color: #1E293B; font-size: 13px;">
              📄 ${r.document || 'S/N'}
            </div>
          </td>
          <td style="padding: 12px 10px; font-family: 'Segoe UI', Arial, sans-serif;">
            <div style="font-weight: 600; color: #334155; font-size: 12px; max-width: 200px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="${r.company}">
              ${r.company}
            </div>
            <div style="font-size: 10px; color: #64748B;">NIT: ${r.nit || '—'}</div>
          </td>
          <td style="padding: 12px 10px; font-family: 'Segoe UI', Arial, sans-serif; font-weight: 800; color: #2563EB; font-size: 13px; text-align: right;">
            ${formatCOP(r.total)}
          </td>
          <td style="padding: 12px 10px; font-family: 'Segoe UI', Arial, sans-serif; font-weight: 700; color: #0F172A; font-size: 13px; text-align: center;">
            ${r.age}
          </td>
          <td style="padding: 12px 10px; font-family: 'Segoe UI', Arial, sans-serif; text-align: center;">
            <span style="background-color: ${badgeBg}; color: ${badgeText}; padding: 4px 10px; border-radius: 12px; font-size: 10.5px; font-weight: 700; display: inline-block;">
              ● ${badgeLabel}
            </span>
          </td>
        </tr>
      `;
    }).join('');

  return `<!DOCTYPE html>
<html lang="es" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Oportunidades de Facturación · Provexpress SAS</title>
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
    body { margin: 0; padding: 0; width: 100% !important; background-color: #F1F5F9; font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, Helvetica, Arial, sans-serif; }
    @media only screen and (max-width: 680px) {
      .main-table { width: 100% !important; }
      .stack-col { display: block !important; width: 100% !important; padding-right: 0 !important; }
      .kpi-table { width: 100% !important; margin-bottom: 8px !important; }
    }
  </style>
</head>
<body style="margin: 0; padding: 24px 8px; background-color: #F1F5F9;">

  <!-- CONTENEDOR PRINCIPAL (MAX 840px) -->
  <table role="presentation" class="main-table" width="840" align="center" border="0" cellpadding="0" cellspacing="0" style="width: 840px; margin: 0 auto; background-color: #FFFFFF; border-radius: 16px; overflow: hidden; border: 1px solid #E2E8F0; box-shadow: 0 4px 20px rgba(15, 23, 42, 0.06);">
    
    <!-- 1. ENCABEZADO CORPORATIVO CON SALUDO -->
    <tr>
      <td style="padding: 28px 36px 20px 36px; border-bottom: 1px solid #E2E8F0; background: linear-gradient(180deg, #FFFFFF 0%, #F8FAFC 100%);">
        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <!-- Logo y saludo -->
            <td valign="top" style="padding-right: 20px;">
              <table role="presentation" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td valign="middle" style="padding-bottom: 12px;">
                    <!-- Logo tipográfico corporativo -->
                    <div style="display: inline-block;">
                      <span style="font-size: 22px; font-weight: 900; color: #0F172A; letter-spacing: -0.02em; font-family: 'Segoe UI', Arial, sans-serif;">
                        PROVEXPRESS <span style="color: #2563EB; font-size: 13px; font-weight: 800;">SAS</span>
                      </span>
                      <div style="font-size: 9px; font-weight: 700; color: #64748B; text-transform: uppercase; letter-spacing: 0.08em; margin-top: -2px;">
                        Cumplimos siempre
                      </div>
                    </div>
                  </td>
                </tr>
                <tr>
                  <td>
                    <div style="font-size: 14px; color: #475569; font-weight: 600;">
                      Hola, <strong style="color: #0F172A; font-size: 17px;">${commercialName}</strong>
                    </div>
                    <h1 style="margin: 6px 0 8px 0; font-size: 24px; font-weight: 850; color: #0F172A; line-height: 1.25; letter-spacing: -0.02em;">
                      Oportunidades listas para convertirse en <span style="color: #15803D;">ventas facturadas</span>
                    </h1>
                    <p style="margin: 0; font-size: 13.5px; color: #475569; line-height: 1.45; max-width: 530px;">
                      Excelente trabajo en la generación de negocio. Tienes remisiones entregadas que están listas para avanzar hacia la facturación.
                    </p>
                    <p style="margin: 6px 0 0 0; font-size: 13px; color: #15803D; font-weight: 700;">
                      Cada factura emitida transforma tu esfuerzo comercial en resultados reales.
                    </p>
                  </td>
                </tr>
              </table>
            </td>

            <!-- Badge Motivacional lateral -->
            <td width="200" valign="top" align="right" class="stack-col">
              <table role="presentation" border="0" cellpadding="0" cellspacing="0" style="background-color: #F0FDF4; border: 1.5px solid #86EFAC; border-radius: 12px; padding: 14px 16px; text-align: center; width: 190px;">
                <tr>
                  <td align="center">
                    <div style="background-color: #15803D; color: #FFFFFF; width: 28px; height: 28px; line-height: 28px; border-radius: 50%; font-size: 16px; font-weight: 900; margin: 0 auto 8px auto;">
                      ✓
                    </div>
                    <div style="font-size: 13px; font-weight: 800; color: #166534; line-height: 1.25;">
                      Tu gestión hace la diferencia.
                    </div>
                    <div style="font-size: 10.5px; color: #475569; margin-top: 6px; border-top: 1px dashed #BBF7D0; padding-top: 6px;">
                      Corte: <strong>${formattedCutoff}</strong>
                    </div>
                    <div style="font-size: 10px; color: #64748B;">
                      Dir: ${directorName}
                    </div>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
      </td>
    </tr>

    <!-- 2. TARJETAS KPI RESUMEN (4 COLUMNAS) -->
    <tr>
      <td style="padding: 24px 36px 16px 36px; background-color: #FAFAFA;">
        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <!-- KPI 1: Cantidad -->
            <td width="24%" valign="top" style="padding-right: 8px;">
              <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 12px; padding: 14px 12px; text-align: left;">
                <tr>
                  <td style="font-size: 10px; font-weight: 800; color: #64748B; text-transform: uppercase; letter-spacing: 0.04em;">
                    📦 Listas para gestionar
                  </td>
                </tr>
                <tr>
                  <td style="padding-top: 6px; font-size: 26px; font-weight: 900; color: #1E293B;">
                    ${formatNumber(totalCount)} <span style="font-size: 13px; font-weight: 600; color: #64748B;">rem.</span>
                  </td>
                </tr>
              </table>
            </td>

            <!-- KPI 2: Valor -->
            <td width="26%" valign="top" style="padding-right: 8px;">
              <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1px solid #BBF7D0; border-radius: 12px; padding: 14px 12px; text-align: left;">
                <tr>
                  <td style="font-size: 10px; font-weight: 800; color: #15803D; text-transform: uppercase; letter-spacing: 0.04em;">
                    💵 Valor por facturar
                  </td>
                </tr>
                <tr>
                  <td style="padding-top: 6px; font-size: 20px; font-weight: 900; color: #15803D;">
                    ${formatCOP(totalValue)}
                  </td>
                </tr>
              </table>
            </td>

            <!-- KPI 3: Antigüedad promedio -->
            <td width="24%" valign="top" style="padding-right: 8px;">
              <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1px solid #FED7AA; border-radius: 12px; padding: 14px 12px; text-align: left;">
                <tr>
                  <td style="font-size: 10px; font-weight: 800; color: #C2410C; text-transform: uppercase; letter-spacing: 0.04em;">
                    ⏱ Antigüedad promedio
                  </td>
                </tr>
                <tr>
                  <td style="padding-top: 6px; font-size: 24px; font-weight: 900; color: #C2410C;">
                    ${avgAge} <span style="font-size: 13px; font-weight: 600; color: #9A3412;">días</span>
                  </td>
                </tr>
              </table>
            </td>

            <!-- KPI 4: Potencial -->
            <td width="26%" valign="top">
              <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1px solid #BFDBFE; border-radius: 12px; padding: 14px 12px; text-align: left;">
                <tr>
                  <td style="font-size: 10px; font-weight: 800; color: #1D4ED8; text-transform: uppercase; letter-spacing: 0.04em;">
                    📈 Ventas por reconocer
                  </td>
                </tr>
                <tr>
                  <td style="padding-top: 6px; font-size: 20px; font-weight: 900; color: #1D4ED8;">
                    ${formatCOP(totalValue)}
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
      </td>
    </tr>

    <!-- 3. CUERPO: TABLA DE REMISIONES DESTACADAS + LATERAL INSPIRACIONAL -->
    <tr>
      <td style="padding: 16px 36px 24px 36px;">
        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <!-- Columna Izquierda: Tabla de Remisiones Destacadas -->
            <td width="66%" valign="top" style="padding-right: 18px;" class="stack-col">
              <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 12px; overflow: hidden;">
                <!-- Header de sección -->
                <tr>
                  <td style="padding: 14px 16px; background-color: #F8FAFC; border-bottom: 1px solid #E2E8F0;">
                    <div style="font-size: 14px; font-weight: 800; color: #0F172A;">
                      ⭐ Tus remisiones destacadas
                    </div>
                    <div style="font-size: 11px; color: #64748B; margin-top: 2px;">
                      Enfócate primero en las de mayor antigüedad para acelerar su facturación.
                    </div>
                  </td>
                </tr>
                <!-- Tabla -->
                <tr>
                  <td style="padding: 0;">
                    <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="font-size: 12px; border-collapse: collapse;">
                      <thead>
                        <tr style="background-color: #0F172A; color: #FFFFFF; font-size: 11px; text-transform: uppercase; letter-spacing: 0.04em;">
                          <th style="padding: 10px 10px; text-align: left; font-weight: 700;">Remisión</th>
                          <th style="padding: 10px 10px; text-align: left; font-weight: 700;">Cliente</th>
                          <th style="padding: 10px 10px; text-align: right; font-weight: 700;">Valor</th>
                          <th style="padding: 10px 10px; text-align: center; font-weight: 700;">Días</th>
                          <th style="padding: 10px 10px; text-align: center; font-weight: 700;">Estado</th>
                        </tr>
                      </thead>
                      <tbody>
                        ${rowsDestacadasHtml}
                      </tbody>
                    </table>
                  </td>
                </tr>
              </table>
            </td>

            <!-- Columna Derecha: Tarjetas "Cada factura cuenta" y "Meta del día" -->
            <td width="34%" valign="top" class="stack-col">
              <!-- Tarjeta 1: Cada factura cuenta -->
              <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 12px; padding: 16px; margin-bottom: 14px;">
                <tr>
                  <td>
                    <div style="font-size: 13px; font-weight: 800; color: #15803D; margin-bottom: 8px;">
                      ★ Cada factura cuenta
                    </div>
                    <div style="font-size: 11.5px; color: #475569; font-weight: 600; margin-bottom: 8px;">
                      Facturar a tiempo nos ayuda a:
                    </div>
                    <ul style="margin: 0; padding-left: 16px; font-size: 11px; color: #475569; line-height: 1.6;">
                      <li>Reflejar el valor de tu gestión comercial.</li>
                      <li>Cumplir nuestras metas juntos.</li>
                      <li>Mejorar flujo de caja y fortalecer la operación.</li>
                    </ul>
                    <div style="font-size: 11.5px; font-weight: 800; color: #15803D; margin-top: 12px;">
                      ¡Sigamos logrando grandes cosas juntos! ♡
                    </div>
                  </td>
                </tr>
              </table>

              <!-- Tarjeta 2: Meta del día -->
              <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #EFF6FF; border: 1px solid #BFDBFE; border-radius: 12px; padding: 16px;">
                <tr>
                  <td>
                    <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td valign="middle">
                          <div style="font-size: 12.5px; font-weight: 800; color: #1E40AF;">
                            🎯 Meta del día
                          </div>
                        </td>
                        <td align="right" valign="middle">
                          <span style="font-size: 18px;">🏆</span>
                        </td>
                      </tr>
                    </table>
                    <div style="font-size: 11px; color: #334155; line-height: 1.4; margin-top: 6px;">
                      Gestionar las remisiones con mayor antigüedad para acelerar el reconocimiento de ventas.
                    </div>
                    <!-- Barra de progreso -->
                    <div style="background-color: #DBEAFE; border-radius: 8px; height: 10px; width: 100%; margin: 12px 0 6px 0; overflow: hidden;">
                      <div style="background-color: #2563EB; height: 10px; width: ${Math.min(100, Math.max(15, progressPct))}%; border-radius: 8px;"></div>
                    </div>
                    <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td style="font-size: 10.5px; font-weight: 700; color: #1E40AF;">
                          Vamos por más, ¡tú puedes! 💪
                        </td>
                        <td align="right" style="font-size: 11px; font-weight: 800; color: #1E40AF;">
                          ${progressPct}%
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
      </td>
    </tr>

    <!-- 4. TOP OPORTUNIDAD DEL DÍA (HERO BANNER) -->
    ${topOpportunity ? `
    <tr>
      <td style="padding: 0 36px 24px 36px;">
        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background: linear-gradient(135deg, #FFFBEB 0%, #FEF3C7 100%); border: 1.5px solid #FCD34D; border-radius: 12px; padding: 18px 22px;">
          <tr>
            <td width="48" valign="middle" align="center" style="padding-right: 14px;">
              <span style="font-size: 32px;">🏆</span>
            </td>
            <td valign="middle">
              <div style="font-size: 11px; font-weight: 800; color: #B45309; text-transform: uppercase; letter-spacing: 0.05em;">
                Top oportunidad del día ☆
              </div>
              <div style="font-size: 17px; font-weight: 900; color: #78350F; margin-top: 2px;">
                Remisión ${topOpportunity.document || 'S/N'} · ${topOpportunity.company}
              </div>
              <div style="font-size: 12px; color: #92400E; margin-top: 2px;">
                Valor: <strong style="color: #0F172A; font-size: 13px;">${formatCOP(topOpportunity.total)}</strong> &nbsp;·&nbsp;
                Antigüedad: <strong style="color: #DC2626; font-size: 13px;">${topOpportunity.age} días</strong>
              </div>
            </td>
            <td align="right" valign="middle" class="stack-col" style="padding-left: 16px;">
              <div style="background-color: #FFFFFF; border: 1px solid #FDE68A; border-radius: 8px; padding: 8px 14px; display: inline-block; font-size: 12px; font-weight: 800; color: #B45309; white-space: nowrap;">
                ⤤ ¡Gran impacto si la gestionas hoy!
              </div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    ` : ''}

    <!-- 5. PIE DE PÁGINA Y FIRMA CORPORATIVA -->
    <tr>
      <td style="padding: 22px 36px; background-color: #F8FAFC; border-top: 1px solid #E2E8F0;">
        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td valign="middle">
              <div style="font-size: 12.5px; color: #475569; font-weight: 600;">
                💙 Gracias por tu compromiso y dedicación.
              </div>
              <div style="font-size: 13px; font-weight: 800; color: #0F172A; margin-top: 2px;">
                ¡Sigamos construyendo resultados juntos!
              </div>
            </td>
            <td align="right" valign="middle" style="padding-left: 16px;">
              <div style="font-size: 11px; color: #64748B;">
                Cordialmente,
              </div>
              <div style="font-size: 12px; font-weight: 800; color: #0F172A; margin-top: 1px;">
                Gerencia Administrativa y Financiera
              </div>
              <div style="font-size: 11px; font-weight: 800; color: #2563EB;">
                PROVEXPRESS SAS
              </div>
            </td>
          </tr>
        </table>
      </td>
    </tr>

  </table>

  <!-- NOTA DE PIE EXTERIOR -->
  <table role="presentation" width="840" align="center" border="0" cellpadding="0" cellspacing="0" style="width: 840px; margin: 12px auto 0 auto;">
    <tr>
      <td align="center" style="font-size: 11px; color: #94A3B8; font-family: 'Segoe UI', Arial, sans-serif;">
        Este es un mensaje institucional automático de apoyo comercial emitido por Provexpress SAS.
      </td>
    </tr>
  </table>

</body>
</html>`;
}
