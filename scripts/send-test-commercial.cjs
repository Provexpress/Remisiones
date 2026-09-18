const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

// 1. Cargar dependencias de m365-report-mailer
const m365Dir = 'C:/Proyectos/m365-report-mailer';
const { ClientSecretCredential } = require(path.join(m365Dir, 'node_modules/@azure/identity'));
const axios = require(path.join(m365Dir, 'node_modules/axios'));

// 2. Cargar variables de entorno de m365-report-mailer
function loadEnv(filePath) {
  const env = {};
  if (fs.existsSync(filePath)) {
    const lines = fs.readFileSync(filePath, 'utf8').split('\n');
    for (const line of lines) {
      const parts = line.split('=');
      if (parts.length >= 2) {
        const key = parts[0].trim();
        let val = parts.slice(1).join('=').trim();
        if ((val.startsWith('"') && val.endsWith('"')) || (val.startsWith("'") && val.endsWith("'"))) {
          val = val.slice(1, -1);
        }
        env[key] = val.trim();
      }
    }
  }
  return env;
}

const env = loadEnv(path.join(m365Dir, '.env'));
const tenantId = env.TENANT_ID;
const clientId = env.CLIENT_ID;
const clientSecret = env.CLIENT_SECRET;
const senderEmail = env.SENDER_EMAIL || 'juannovoa@provexpress.com.co';
const TARGET_TEST_EMAIL = 'especialista.preventa@provexpress.com.co';

function formatCOP(amount) {
  if (amount == null || isNaN(amount)) return '$ 0';
  return new Intl.NumberFormat('es-CO', {
    style: 'currency',
    currency: 'COP',
    maximumFractionDigits: 0,
  }).format(amount).replace('COP', '').replace(/\u00A0/g, ' ').trim();
}

function formatNumber(num) {
  if (num == null || isNaN(num)) return '0';
  return new Intl.NumberFormat('es-CO').format(Math.round(num));
}

async function main() {
  console.log('================================================================');
  console.log('🚀 PREPARANDO CORREO DE PRUEBA DE NOTIFICACIÓN COMERCIAL');
  console.log(`📩 Destinatario: ${TARGET_TEST_EMAIL}`);
  console.log(`📤 Remitente institucional: ${senderEmail}`);
  console.log('================================================================');

  // Leer libro de remisiones
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(path.join(__dirname, '../Remisiones.xlsx'));
  const sis = wb.getWorksheet('Base-SIS');

  const targetEmployee = 'Dayana Marcela Chala';
  const cutoffDate = '2026-09-16';
  const items = [];

  for (let r = 2; r <= sis.rowCount; r++) {
    const row = sis.getRow(r);
    const dVal = row.getCell(1).value;
    const iso = dVal instanceof Date ? dVal.toISOString().slice(0, 10) : String(dVal || '').trim();
    if (iso === cutoffDate) {
      const emp = String(row.getCell(2).value || '').trim();
      if (emp.toLowerCase().includes('dayana')) {
        const total = Number(row.getCell(7).value || 0);
        const age = Number(row.getCell(9).value || 0);
        const doc = String(row.getCell(10).value || '').trim();
        const nit = String(row.getCell(3).value || '').trim();
        const company = String(row.getCell(4).value || '').trim();
        items.push({ doc, nit, company, total, age });
      }
    }
  }

  console.log(`✓ Remisiones encontradas para ${targetEmployee} en fecha ${cutoffDate}: ${items.length}`);

  // Ordenar por antigüedad desc y total desc
  items.sort((a, b) => (b.age !== a.age ? b.age - a.age : b.total - a.total));

  const totalCount = items.length;
  const totalValue = items.reduce((s, x) => s + x.total, 0);
  const avgAge = totalCount > 0 ? Math.round(items.reduce((s, x) => s + x.age, 0) / totalCount) : 0;
  const destacadas = items.slice(0, 5);
  const topOpportunity = items[0];
  const onTrackCount = items.filter(x => x.age <= 15).length;
  const progressPct = totalCount > 0 ? Math.round((onTrackCount / totalCount) * 100) : 100;

  console.log(`   - Total Documentos: ${totalCount}`);
  console.log(`   - Total Valor: ${formatCOP(totalValue)}`);
  console.log(`   - Antigüedad Promedio: ${avgAge} días`);
  console.log(`   - Top Oportunidad: ${topOpportunity.doc} - ${topOpportunity.company} (${formatCOP(topOpportunity.total)})`);

  // Filas de remisiones destacadas (3 filas con temas de color exactos)
  const rowColors = [
    { circleBg: '#4F46E5', valColor: '#3B28CC' },
    { circleBg: '#D97706', valColor: '#D97706' },
    { circleBg: '#15803D', valColor: '#15803D' },
  ];

  const rowsHtml = items.slice(0, 3).map((r, idx) => {
    const colTheme = rowColors[idx % rowColors.length];
    let badgeBg = '#FEF3C7';
    let badgeDot = '#D97706';
    let badgeText = '#B45309';
    let badgeLabel = '8 a 15 días';

    if (r.age > 15) {
      badgeBg = '#FEE2E2';
      badgeDot = '#DC2626';
      badgeText = '#DC2626';
      badgeLabel = 'Más de 15 días';
    } else if (r.age <= 7) {
      badgeBg = '#DCFCE7';
      badgeDot = '#16A34A';
      badgeText = '#15803D';
      badgeLabel = '0 a 7 días';
    }

    return `
      <tr style="border-bottom: 1px solid #F1F5F9;">
        <td style="padding: 12px 14px; font-family: 'Segoe UI', Arial, sans-serif;">
          <table role="presentation" border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td style="vertical-align: middle; padding-right: 8px;">
                <div style="background-color: ${colTheme.circleBg}; width: 26px; height: 26px; border-radius: 6px; text-align: center; line-height: 26px; color: #FFFFFF; font-size: 13px;">
                  📄
                </div>
              </td>
              <td style="vertical-align: middle;">
                <span style="font-weight: 800; color: #0F172A; font-size: 13px;">
                  ${r.doc?.startsWith('REM-') ? r.doc : `REM-${r.doc || 'S/N'}`}
                </span>
              </td>
            </tr>
          </table>
        </td>
        <td style="padding: 12px 14px; font-family: 'Segoe UI', Arial, sans-serif;">
          <span style="font-size: 13px; color: #334155; font-weight: 500;">
            ${r.company || 'Cliente'}
          </span>
        </td>
        <td style="padding: 12px 14px; font-family: 'Segoe UI', Arial, sans-serif; text-align: right;">
          <span style="font-weight: 850; color: ${colTheme.valColor}; font-size: 14px;">
            ${formatCOP(r.total)}
          </span>
        </td>
        <td style="padding: 12px 14px; font-family: 'Segoe UI', Arial, sans-serif; text-align: center;">
          <span style="font-weight: 700; color: #1E293B; font-size: 14px;">
            ${r.age}
          </span>
        </td>
        <td style="padding: 12px 14px; font-family: 'Segoe UI', Arial, sans-serif; text-align: center;">
          <span style="background-color: ${badgeBg}; color: ${badgeText}; padding: 4px 12px; border-radius: 999px; font-size: 11px; font-weight: 700; display: inline-block; white-space: nowrap;">
            <span style="color: ${badgeDot}; font-size: 9px; vertical-align: middle; margin-right: 4px;">●</span>${badgeLabel}
          </span>
        </td>
      </tr>
    `;
  }).join('');

  const htmlContent = `<!DOCTYPE html>
<html lang="es" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Oportunidades listas para convertirse en ventas facturadas · Provexpress SAS</title>
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
    body { margin: 0; padding: 0; width: 100% !important; background-color: #F0F3F7; font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, Helvetica, Arial, sans-serif; }
    @media only screen and (max-width: 860px) {
      .main-card { width: 100% !important; border-radius: 0 !important; }
      .stack-col { display: block !important; width: 100% !important; padding-right: 0 !important; }
      .kpi-cell { display: block !important; width: 100% !important; margin-bottom: 8px !important; }
    }
  </style>
</head>
<body style="margin: 0; padding: 20px 0; background-color: #F0F3F7;">

  <table role="presentation" class="main-card" width="880" align="center" border="0" cellpadding="0" cellspacing="0" style="width: 880px; margin: 0 auto; background-color: #FFFFFF; border-radius: 20px; overflow: hidden; border: 1px solid #E2E8F0; box-shadow: 0 8px 30px rgba(15, 23, 42, 0.07);">
    <tr>
      <td style="padding: 28px 32px 32px 32px;">

        <!-- 1. ENCABEZADO -->
        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="margin-bottom: 24px;">
          <tr>
            <!-- Logo Provexpress SAS -->
            <td width="130" valign="top" style="padding-right: 16px;">
              <div style="font-family: 'Segoe UI', Arial, sans-serif; text-align: left;">
                <svg width="44" height="44" viewBox="0 0 44 44" fill="none" xmlns="http://www.w3.org/2000/svg">
                  <path d="M8 8H26C31.5228 8 36 12.4772 36 18C36 23.5228 31.5228 28 26 28H18V38H8V8Z" fill="#1E3A8A"/>
                  <path d="M18 16H25C26.1046 16 27 16.8954 27 18C27 19.1046 26.1046 20 25 20H18V16Z" fill="#FFFFFF"/>
                  <path d="M8 8H28C33 8 36 12 36 17C36 22 32 26 27 26H16" stroke="#1E3A8A" stroke-width="4" stroke-linecap="round"/>
                </svg>
                <div style="font-size: 13px; font-weight: 900; color: #1E3A8A; letter-spacing: 0.02em; margin-top: 4px; line-height: 1.1;">
                  PROVEXPRESS
                </div>
                <div style="font-size: 9px; font-weight: 800; color: #1E3A8A; letter-spacing: 0.1em; text-align: right; width: 88px;">
                  SAS
                </div>
              </div>
            </td>

            <!-- Saludo y titular -->
            <td valign="top" style="padding-right: 12px;">
              <div style="font-size: 15px; color: #475569; font-weight: 600; font-family: 'Segoe UI', Arial, sans-serif;">
                Hola, <strong style="color: #0F172A; font-size: 16px;">Dayana Chala</strong>
              </div>
              <h1 style="margin: 4px 0 6px 0; font-size: 25px; font-weight: 900; color: #0F172A; line-height: 1.2; letter-spacing: -0.02em; font-family: 'Segoe UI', Arial, sans-serif;">
                Oportunidades listas para convertirse<br>
                en <span style="color: #16A34A;">ventas facturadas</span>
              </h1>
              <p style="margin: 0; font-size: 13px; color: #475569; line-height: 1.45; font-family: 'Segoe UI', Arial, sans-serif; max-width: 440px;">
                Excelente trabajo en la generación de negocio. Tienes remisiones entregadas que están listas para avanzar hacia la facturación.
              </p>
              <p style="margin: 6px 0 0 0; font-size: 12.5px; color: #0F172A; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">
                Cada factura emitida transforma tu esfuerzo comercial en <span style="color: #16A34A;">resultados reales</span>.
              </p>
            </td>

            <!-- Ilustración comercial con pulgar arriba -->
            <td width="130" valign="bottom" align="center" style="padding-right: 12px;">
              <svg width="120" height="120" viewBox="0 0 120 120" fill="none" xmlns="http://www.w3.org/2000/svg">
                <circle cx="60" cy="60" r="54" fill="#F1F5F9"/>
                <path d="M30 115C30 96 42 82 60 82C78 82 90 96 90 115" fill="#1E3A8A"/>
                <path d="M52 82L60 92L68 82" stroke="#FFFFFF" stroke-width="2.5" stroke-linecap="round"/>
                <path d="M53 70H67V82C67 82 65 85 60 85C55 85 53 82 53 82V70Z" fill="#FBD38D"/>
                <ellipse cx="60" cy="55" rx="16" ry="19" fill="#FBD38D"/>
                <path d="M44 52C44 40 50 36 60 36C70 36 76 40 76 52C74 50 71 47 67 47C62 47 58 49 53 47C49 47 46 50 44 52Z" fill="#1E293B"/>
                <circle cx="55" cy="54" r="1.5" fill="#1E293B"/>
                <circle cx="65" cy="54" r="1.5" fill="#1E293B"/>
                <path d="M57 62C58 64 62 64 63 62" stroke="#1E293B" stroke-width="1.5" stroke-linecap="round"/>
                <path d="M38 100L42 78C42 78 44 75 48 76C52 77 50 82 50 82L46 100" fill="#1E3A8A"/>
                <path d="M43 72C43 70 45 67 47 67C49 67 50 69 50 72V76H43V72Z" fill="#FBD38D"/>
                <rect x="42" y="74" width="9" height="10" rx="3" fill="#FBD38D"/>
              </svg>
            </td>

            <!-- Card "Tu gestión hace la diferencia" -->
            <td width="135" valign="top" align="right">
              <table role="presentation" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 14px; padding: 12px 10px; width: 125px; text-align: center; box-shadow: 0 2px 8px rgba(0,0,0,0.02);">
                <tr>
                  <td align="center">
                    <div style="background-color: #16A34A; color: #FFFFFF; width: 26px; height: 26px; border-radius: 50%; font-size: 14px; font-weight: 900; line-height: 26px; margin: 0 auto 8px auto;">
                      ✓
                    </div>
                    <div style="font-size: 12.5px; font-weight: 700; color: #1E293B; line-height: 1.3; font-family: 'Segoe UI', Arial, sans-serif;">
                      Tu gestión<br>
                      hace la<br>
                      <strong style="color: #16A34A; font-size: 13.5px;">diferencia.</strong>
                    </div>
                    <div style="margin-top: 8px; font-size: 14px; line-height: 1;">
                      🌿
                    </div>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>

        <!-- 2. TARJETAS KPI (4 EN FILA) -->
        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="margin-bottom: 22px;">
          <tr>
            <td width="24%" valign="top" class="kpi-cell" style="padding-right: 8px;">
              <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 14px; padding: 14px 12px; height: 100px;">
                <tr>
                  <td width="36" valign="top">
                    <div style="background-color: #312E81; width: 34px; height: 34px; border-radius: 50%; text-align: center; line-height: 34px; font-size: 16px; color: #FFFFFF;">
                      📦
                    </div>
                  </td>
                  <td valign="top" style="padding-left: 8px;">
                    <div style="font-size: 11px; color: #475569; font-weight: 600; line-height: 1.25; font-family: 'Segoe UI', Arial, sans-serif;">
                      Remisiones listas<br>para gestionar
                    </div>
                  </td>
                </tr>
                <tr>
                  <td colspan="2" valign="bottom" style="padding-top: 8px;">
                    <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td valign="bottom" style="font-size: 26px; font-weight: 900; color: #312E81; font-family: 'Segoe UI', Arial, sans-serif; line-height: 1;">
                          ${formatNumber(totalCount)}
                        </td>
                        <td align="right" valign="bottom" style="font-size: 18px; color: #94A3B8; opacity: 0.65;">
                          📝
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>

            <td width="26%" valign="top" class="kpi-cell" style="padding-right: 8px;">
              <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 14px; padding: 14px 12px; height: 100px;">
                <tr>
                  <td width="36" valign="top">
                    <div style="background-color: #15803D; width: 34px; height: 34px; border-radius: 50%; text-align: center; line-height: 34px; font-size: 16px; font-weight: 900; color: #FFFFFF;">
                      $
                    </div>
                  </td>
                  <td valign="top" style="padding-left: 8px;">
                    <div style="font-size: 11px; color: #475569; font-weight: 600; line-height: 1.25; font-family: 'Segoe UI', Arial, sans-serif;">
                      Valor por facturar
                    </div>
                  </td>
                </tr>
                <tr>
                  <td colspan="2" valign="bottom" style="padding-top: 8px;">
                    <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td valign="bottom" style="font-size: 19px; font-weight: 900; color: #15803D; font-family: 'Segoe UI', Arial, sans-serif; line-height: 1;">
                          ${formatCOP(totalValue)}
                        </td>
                        <td align="right" valign="bottom" style="font-size: 18px; color: #86EFAC; opacity: 0.75;">
                          🪙
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>

            <td width="24%" valign="top" class="kpi-cell" style="padding-right: 8px;">
              <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 14px; padding: 14px 12px; height: 100px;">
                <tr>
                  <td width="36" valign="top">
                    <div style="background-color: #D97706; width: 34px; height: 34px; border-radius: 50%; text-align: center; line-height: 34px; font-size: 16px; color: #FFFFFF;">
                      🕒
                    </div>
                  </td>
                  <td valign="top" style="padding-left: 8px;">
                    <div style="font-size: 11px; color: #475569; font-weight: 600; line-height: 1.25; font-family: 'Segoe UI', Arial, sans-serif;">
                      Antigüedad<br>promedio
                    </div>
                  </td>
                </tr>
                <tr>
                  <td colspan="2" valign="bottom" style="padding-top: 8px;">
                    <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td valign="bottom" style="font-size: 22px; font-weight: 900; color: #D97706; font-family: 'Segoe UI', Arial, sans-serif; line-height: 1;">
                          ${avgAge} días
                        </td>
                        <td align="right" valign="bottom" style="font-size: 18px; color: #FDBA74; opacity: 0.75;">
                          📅
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>

            <td width="26%" valign="top" class="kpi-cell">
              <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 14px; padding: 14px 12px; height: 100px;">
                <tr>
                  <td width="36" valign="top">
                    <div style="background-color: #1E40AF; width: 34px; height: 34px; border-radius: 50%; text-align: center; line-height: 34px; font-size: 16px; color: #FFFFFF;">
                      📊
                    </div>
                  </td>
                  <td valign="top" style="padding-left: 8px;">
                    <div style="font-size: 11px; color: #475569; font-weight: 600; line-height: 1.25; font-family: 'Segoe UI', Arial, sans-serif;">
                      Potencial de ventas<br>por reconocer
                    </div>
                  </td>
                </tr>
                <tr>
                  <td colspan="2" valign="bottom" style="padding-top: 8px;">
                    <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td valign="bottom" style="font-size: 19px; font-weight: 900; color: #1E40AF; font-family: 'Segoe UI', Arial, sans-serif; line-height: 1;">
                          ${formatCOP(totalValue)}
                        </td>
                        <td align="right" valign="bottom" style="font-size: 18px; color: #93C5FD; opacity: 0.75;">
                          🏆
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>

        <!-- 3. TABLA Y LATERAL -->
        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="margin-bottom: 20px;">
          <tr>
            <!-- Columna Izquierda: Tabla -->
            <td width="66%" valign="top" style="padding-right: 18px;" class="stack-col">
              <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="margin-bottom: 8px;">
                <tr>
                  <td>
                    <span style="font-size: 16px; font-weight: 850; color: #0F172A; font-family: 'Segoe UI', Arial, sans-serif;">
                      Tus remisiones destacadas
                    </span>
                    <span style="color: #D97706; font-size: 15px; margin-left: 4px;">☆</span>
                    <span style="font-size: 12px; color: #475569; margin-left: 14px; font-family: 'Segoe UI', Arial, sans-serif;">
                      Enfócate primero en las de mayor antigüedad.
                    </span>
                  </td>
                </tr>
              </table>

              <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="border: 1px solid #E2E8F0; border-radius: 12px; overflow: hidden; border-collapse: separate; border-spacing: 0;">
                <thead>
                  <tr style="background-color: #1E293B; color: #FFFFFF;">
                    <th style="padding: 10px 14px; text-align: left; font-size: 11.5px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">Remisión</th>
                    <th style="padding: 10px 14px; text-align: left; font-size: 11.5px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">Cliente</th>
                    <th style="padding: 10px 14px; text-align: right; font-size: 11.5px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">Valor</th>
                    <th style="padding: 10px 14px; text-align: center; font-size: 11.5px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">Días abierta</th>
                    <th style="padding: 10px 14px; text-align: center; font-size: 11.5px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">Antigüedad</th>
                  </tr>
                </thead>
                <tbody style="background-color: #FFFFFF;">
                  ${rowsHtml}
                </tbody>
              </table>
            </td>

            <!-- Columna Derecha: Tarjetas motivacionales -->
            <td width="34%" valign="top" class="stack-col">
              <!-- Cada factura cuenta -->
              <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 14px; padding: 16px; margin-bottom: 12px;">
                <tr>
                  <td>
                    <div style="font-size: 15px; font-weight: 850; color: #166534; font-family: 'Segoe UI', Arial, sans-serif; margin-bottom: 8px;">
                      <span style="font-size: 16px;">★</span> Cada factura cuenta
                    </div>
                    <div style="font-size: 12px; color: #475569; font-family: 'Segoe UI', Arial, sans-serif; margin-bottom: 10px;">
                      Facturar a tiempo nos ayuda a:
                    </div>
                    <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="font-size: 11.5px; color: #334155; font-family: 'Segoe UI', Arial, sans-serif;">
                      <tr>
                        <td width="20" valign="top" style="padding-bottom: 8px; color: #16A34A; font-weight: 900; font-size: 13px;">✓</td>
                        <td valign="top" style="padding-bottom: 8px; line-height: 1.35;">Reflejar el valor de tu gestión comercial.</td>
                      </tr>
                      <tr>
                        <td width="20" valign="top" style="padding-bottom: 8px; color: #16A34A; font-weight: 900; font-size: 13px;">✓</td>
                        <td valign="top" style="padding-bottom: 8px; line-height: 1.35;">Cumplir nuestras metas juntos.</td>
                      </tr>
                      <tr>
                        <td width="20" valign="top" style="padding-bottom: 8px; color: #16A34A; font-weight: 900; font-size: 13px;">✓</td>
                        <td valign="top" style="padding-bottom: 8px; line-height: 1.35;">Mejorar flujo de caja y fortalecer nuestra operación.</td>
                      </tr>
                    </table>
                    <div style="font-size: 13px; font-weight: 800; color: #166534; font-family: 'Segoe UI', Arial, sans-serif; margin-top: 6px;">
                      ¡Sigamos logrando<br>grandes cosas juntos! <span style="font-size: 14px;">♡</span>
                    </div>
                  </td>
                </tr>
              </table>

              <!-- Meta del día -->
              <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 14px; padding: 14px 16px;">
                <tr>
                  <td>
                    <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="margin-bottom: 6px;">
                      <tr>
                        <td valign="middle">
                          <span style="font-size: 15px; margin-right: 4px;">🎯</span>
                          <span style="font-size: 13.5px; font-weight: 850; color: #1E3A8A; font-family: 'Segoe UI', Arial, sans-serif;">
                            Meta del día
                          </span>
                        </td>
                        <td align="right" valign="middle">
                          <span style="font-size: 20px;">🏆</span>
                        </td>
                      </tr>
                    </table>
                    <div style="font-size: 11px; color: #475569; line-height: 1.35; font-family: 'Segoe UI', Arial, sans-serif; margin-bottom: 10px;">
                      Gestionar las remisiones con mayor antigüedad para acelerar el reconocimiento de ventas.
                    </div>
                    <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="margin-bottom: 8px;">
                      <tr>
                        <td valign="middle" style="padding-right: 10px;">
                          <div style="background-color: #E5E7EB; border-radius: 999px; height: 10px; width: 100%; overflow: hidden;">
                            <div style="background-color: #16A34A; height: 10px; width: ${progressPct}%; border-radius: 999px;"></div>
                          </div>
                        </td>
                        <td width="36" align="right" valign="middle" style="font-size: 13px; font-weight: 850; color: #1E293B; font-family: 'Segoe UI', Arial, sans-serif;">
                          ${progressPct}%
                        </td>
                      </tr>
                    </table>
                    <div style="font-size: 11.5px; color: #334155; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">
                      Vamos por más, ¡tú puedes! 💪
                    </div>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>

        <!-- 4. TOP OPORTUNIDAD DEL DÍA -->
        ${topOpportunity ? `
        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 14px; padding: 16px 20px; margin-bottom: 20px;">
          <tr>
            <td width="55" valign="middle" align="center" style="padding-right: 14px;">
              <span style="font-size: 38px;">🏆</span>
            </td>
            <td valign="middle" style="padding-right: 16px;">
              <div style="font-size: 11px; font-weight: 700; color: #475569; font-family: 'Segoe UI', Arial, sans-serif;">
                Top oportunidad del día <span style="color: #D97706;">☆</span>
              </div>
              <div style="font-size: 17px; font-weight: 900; color: #1E3A8A; font-family: 'Segoe UI', Arial, sans-serif; margin: 1px 0;">
                ${topOpportunity.doc?.startsWith('REM-') ? topOpportunity.doc : `REM-${topOpportunity.doc || 'S/N'}`}
              </div>
              <div style="font-size: 12px; color: #475569; font-family: 'Segoe UI', Arial, sans-serif;">
                ${topOpportunity.company || 'Cliente'}
              </div>
            </td>
            <td width="130" valign="middle" style="padding-right: 16px;">
              <div style="font-size: 17px; font-weight: 900; color: #1E3A8A; font-family: 'Segoe UI', Arial, sans-serif;">
                ${formatCOP(topOpportunity.total)}
              </div>
              <div style="font-size: 11px; color: #64748B; font-family: 'Segoe UI', Arial, sans-serif;">
                Valor
              </div>
            </td>
            <td width="100" valign="middle" style="padding-right: 16px;">
              <div style="font-size: 17px; font-weight: 900; color: #DC2626; font-family: 'Segoe UI', Arial, sans-serif;">
                ${topOpportunity.age} días
              </div>
              <div style="font-size: 11px; color: #64748B; font-family: 'Segoe UI', Arial, sans-serif;">
                Antigüedad
              </div>
            </td>
            <td valign="middle" align="right">
              <table role="presentation" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td valign="middle" style="padding-right: 6px; font-size: 22px; color: #1E3A8A;">
                    ⤷
                  </td>
                  <td valign="middle" style="font-size: 14px; font-weight: 850; font-style: italic; color: #1E3A8A; font-family: 'Segoe UI', Arial, sans-serif; line-height: 1.2;">
                    ¡Gran impacto<br>si la gestionas hoy!
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
        ` : ''}

        <!-- 5. PIE DE PÁGINA -->
        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #F8FAFC; border: 1.5px solid #E2E8F0; border-radius: 14px; padding: 18px 24px;">
          <tr>
            <td width="42" valign="middle" align="center" style="padding-right: 14px;">
              <div style="font-size: 32px; line-height: 1; color: #1E3A8A;">
                💙
              </div>
            </td>
            <td valign="middle">
              <div style="font-size: 12px; color: #475569; font-family: 'Segoe UI', Arial, sans-serif;">
                Gracias por tu compromiso y dedicación.
              </div>
              <div style="font-size: 14px; font-weight: 850; color: #1E3A8A; font-family: 'Segoe UI', Arial, sans-serif; margin-top: 2px;">
                ¡Sigamos construyendo resultados juntos!
              </div>
            </td>
            <td width="20" align="center" valign="middle" style="color: #CBD5E1; font-size: 22px; font-weight: 300;">
              |
            </td>
            <td valign="middle" style="padding-left: 10px;">
              <div style="font-size: 11.5px; color: #475569; font-family: 'Segoe UI', Arial, sans-serif;">
                Cordialmente, <span style="font-size: 12px;">♡</span>
              </div>
              <div style="font-size: 11.5px; font-weight: 600; color: #334155; font-family: 'Segoe UI', Arial, sans-serif; margin: 1px 0;">
                Gerencia Administrativa y Financiera
              </div>
              <div style="font-size: 12px; font-weight: 850; color: #2563EB; font-family: 'Segoe UI', Arial, sans-serif;">
                PROVEXPRESS SAS
              </div>
            </td>
            <td width="60" align="right" valign="middle">
              <svg width="52" height="42" viewBox="0 0 52 42" fill="none" xmlns="http://www.w3.org/2000/svg">
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

  // 4. Autenticar y enviar vía Microsoft Graph
  console.log('\n🔑 Autenticando con Microsoft Entra ID (Graph)...');
  const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
  const tokenResponse = await credential.getToken('https://graph.microsoft.com/.default');
  const token = tokenResponse.token;
  console.log('✓ Token de Microsoft Graph obtenido exitosamente.');

  console.log(`\n📨 Enviando correo a ${TARGET_TEST_EMAIL}...`);
  const graphSendUrl = `https://graph.microsoft.com/v1.0/users/${senderEmail}/sendMail`;

  const payload = {
    message: {
      subject: `[PRUEBA] Oportunidades de Facturación · Dayana Chala · Corte 16/09/2026`,
      body: {
        contentType: 'HTML',
        content: htmlContent,
      },
      toRecipients: [
        {
          emailAddress: {
            address: TARGET_TEST_EMAIL,
          },
        },
      ],
    },
    saveToSentItems: false,
  };

  const response = await axios.post(graphSendUrl, payload, {
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
  });

  console.log(`\n✅ ¡CORREO ENVIADO CON ÉXITO!`);
  console.log(`   - Código de respuesta HTTP: ${response.status} ${response.statusText || 'Accepted'}`);
  console.log(`   - Destino: ${TARGET_TEST_EMAIL}`);
  console.log('================================================================');
}

main().catch(err => {
  console.error('❌ Error enviando el correo de prueba:', err.response?.data || err.message);
  process.exit(1);
});
