/**
 * Script standalone para despachar el correo de prueba a especialista.preventa@provexpress.com.co
 * Utiliza Microsoft Graph API autenticado con las credenciales de Azure Entra ID
 * Ubicadas en C:\Proyectos\m365-report-mailer\.env
 *
 * Diseño fiel al mockup de referencia:
 * - Logo Provexpress SAS oficial de ForeCast / Provex One (adjunto CID inline)
 * - Ilustración asesor comercial con pulgar arriba / el muñeco (adjunto CID inline)
 * - Tarjeta lateral "Tu gestión hace la diferencia"
 * - Grilla de 2 columnas (Izquierda: 4 KPIs + Tabla destacadas + Hero banner; Derecha: Cada factura cuenta + Meta del día)
 * - Pie institucional: corazón azul, firma oficial Gerencia y sobre postal ilustrado
 */

const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

const TARGET_TEST_EMAIL = 'especialista.preventa@provexpress.com.co';

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
const senderEmail = env.SENDER_EMAIL || env.MAIL_SENDER_USER || 'juannovoa@provexpress.com.co';

if (!tenantId || !clientId || !clientSecret) {
  console.error('❌ Error: Faltan credenciales de Microsoft Entra ID en .env');
  process.exit(1);
}

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

  // 2. Cargar datos reales de Remisiones.xlsx para el corte del 16 de septiembre de 2026
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
  const destacadas = items.slice(0, 3);
  const topOpportunity = items[0];
  const onTrackCount = items.filter(x => x.age <= 15).length;
  const progressPct = totalCount > 0 ? Math.round((onTrackCount / totalCount) * 100) : 100;

  console.log(`   - Total Documentos: ${totalCount}`);
  console.log(`   - Total Valor: ${formatCOP(totalValue)}`);
  console.log(`   - Antigüedad Promedio: ${avgAge} días`);
  console.log(`   - Top Oportunidad: ${topOpportunity.doc} - ${topOpportunity.company} (${formatCOP(topOpportunity.total)})`);

  // Cargar imágenes para adjuntar como CID inline
  const logoPath = path.join(__dirname, '../public/logo_provexpress_stacked.png');
  const avatarPath = path.join(__dirname, '../public/avatar_man.png');
  const logoBase64 = fs.readFileSync(logoPath).toString('base64');
  const avatarBase64 = fs.readFileSync(avatarPath).toString('base64');

  // Filas de remisiones destacadas (3 filas con temas de color exactos)
  const rowColors = [
    { circleBg: '#4F46E5', valColor: '#1E3A8A' },
    { circleBg: '#D97706', valColor: '#D97706' },
    { circleBg: '#15803D', valColor: '#15803D' },
  ];

  const rowsHtml = destacadas.map((r, idx) => {
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
        <td style="padding: 10px 12px; font-family: 'Segoe UI', Arial, sans-serif;">
          <table role="presentation" border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td style="vertical-align: middle; padding-right: 8px;">
                <div style="background-color: ${colTheme.circleBg}; width: 24px; height: 24px; border-radius: 6px; text-align: center; line-height: 24px; color: #FFFFFF; font-size: 12px;">
                  📄
                </div>
              </td>
              <td style="vertical-align: middle;">
                <span style="font-weight: 800; color: #0F172A; font-size: 12.5px;">
                  ${r.doc?.startsWith('REM-') ? r.doc : `REM-${r.doc || 'S/N'}`}
                </span>
              </td>
            </tr>
          </table>
        </td>
        <td style="padding: 10px 12px; font-family: 'Segoe UI', Arial, sans-serif;">
          <span style="font-size: 12.5px; color: #334155; font-weight: 500;">
            ${r.company || 'Cliente'}
          </span>
        </td>
        <td style="padding: 10px 12px; font-family: 'Segoe UI', Arial, sans-serif; text-align: right;">
          <span style="font-weight: 850; color: ${colTheme.valColor}; font-size: 13.5px;">
            ${formatCOP(r.total)}
          </span>
        </td>
        <td style="padding: 10px 12px; font-family: 'Segoe UI', Arial, sans-serif; text-align: center;">
          <span style="font-weight: 700; color: #1E293B; font-size: 13px;">
            ${r.age}
          </span>
        </td>
        <td style="padding: 10px 12px; font-family: 'Segoe UI', Arial, sans-serif; text-align: center;">
          <span style="background-color: ${badgeBg}; color: ${badgeText}; padding: 3px 10px; border-radius: 999px; font-size: 10.5px; font-weight: 700; display: inline-block; white-space: nowrap;">
            <span style="color: ${badgeDot}; font-size: 8px; vertical-align: middle; margin-right: 4px;">●</span>${badgeLabel}
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
             1. ENCABEZADO: LOGO + SALUDO + MUÑECO (AVATAR) + CARD GESTIÓN
             ═════════════════════════════════════════════════════════════════════ -->
        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="margin-bottom: 20px;">
          <tr>
            <!-- LOGO PROVEXPRESS SAS (OFICIAL FORECAST / PROVEX ONE) -->
            <td width="210" valign="middle" style="padding-right: 18px;">
              <img src="cid:logo_provexpress" width="195" height="31" style="display: block; border: 0; max-width: 195px; height: auto;" alt="Provexpress" />
            </td>

            <!-- TEXTO DE SALUDO Y TITULAR -->
            <td valign="middle" style="padding-right: 12px;">
              <div style="font-size: 15px; color: #475569; font-weight: 600; font-family: 'Segoe UI', Arial, sans-serif;">
                Hola, <strong style="color: #0F172A; font-size: 16px;">Dayana Chala</strong>
              </div>
              <h1 style="margin: 3px 0 6px 0; font-size: 24px; font-weight: 900; color: #0F172A; line-height: 1.15; letter-spacing: -0.02em; font-family: 'Segoe UI', Arial, sans-serif;">
                Oportunidades listas para convertirse<br>
                en <span style="color: #16A34A;">ventas facturadas</span>
              </h1>
              <p style="margin: 0; font-size: 12px; color: #475569; line-height: 1.35; font-family: 'Segoe UI', Arial, sans-serif; max-width: 380px;">
                Excelente trabajo en la generación de negocio. Tienes remisiones entregadas que están listas para avanzar hacia la facturación.
              </p>
              <p style="margin: 5px 0 0 0; font-size: 12px; color: #0F172A; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">
                Cada factura emitida transforma tu esfuerzo comercial en <span style="color: #16A34A;">resultados reales</span>.
              </p>
            </td>

            <!-- ILUSTRACIÓN ASESOR COMERCIAL: EL MUÑECO (THUMBS UP) -->
            <td width="190" valign="bottom" align="center" style="padding-right: 10px;">
              <img src="cid:avatar_man" width="180" height="154" style="display: block; border: 0; width: 180px; height: auto; margin: 0 auto;" alt="Asesor Comercial" />
            </td>

            <!-- CARD "Tu gestión hace la diferencia." -->
            <td width="125" valign="middle" align="right">
              <table role="presentation" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 14px; padding: 12px 8px; width: 120px; text-align: center;">
                <tr>
                  <td align="center">
                    <div style="background-color: #16A34A; color: #FFFFFF; width: 24px; height: 24px; border-radius: 50%; font-size: 13px; font-weight: 900; line-height: 24px; margin: 0 auto 6px auto;">
                      ✓
                    </div>
                    <div style="font-size: 12px; font-weight: 700; color: #1E293B; line-height: 1.25; font-family: 'Segoe UI', Arial, sans-serif;">
                      Tu gestión<br>
                      hace la<br>
                      <strong style="color: #16A34A; font-size: 13px;">diferencia.</strong>
                    </div>
                    <div style="margin-top: 6px; font-size: 15px; line-height: 1;">
                      🌿
                    </div>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>

        <!-- ═════════════════════════════════════════════════════════════════════
             2. CUERPO EN 2 COLUMNAS:
                - IZQUIERDA (71%): 4 KPIS + TABLA DESTACADAS + HERO TOP OPORTUNIDAD
                - DERECHA (29%):   CADA FACTURA CUENTA + META DEL DÍA
             ═════════════════════════════════════════════════════════════════════ -->
        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <!-- ── COLUMNA IZQUIERDA (71%) ── -->
            <td width="71%" valign="top" style="padding-right: 18px;" class="stack-col">

              <!-- FILA DE 4 TARJETAS KPI -->
              <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="margin-bottom: 16px;">
                <tr>
                  <!-- KPI 1: Remisiones listas para gestionar -->
                  <td width="24%" valign="top" class="kpi-cell" style="padding-right: 6px;">
                    <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 12px; padding: 12px 10px; height: 95px;">
                      <tr>
                        <td width="32" valign="top">
                          <div style="background-color: #312E81; width: 30px; height: 30px; border-radius: 50%; text-align: center; line-height: 30px; font-size: 15px; color: #FFFFFF;">
                            📦
                          </div>
                        </td>
                        <td valign="top" style="padding-left: 6px;">
                          <div style="font-size: 10.5px; color: #475569; font-weight: 600; line-height: 1.2; font-family: 'Segoe UI', Arial, sans-serif;">
                            Remisiones listas<br>para gestionar
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

                  <!-- KPI 2: Valor por facturar -->
                  <td width="26%" valign="top" class="kpi-cell" style="padding-right: 6px;">
                    <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 12px; padding: 12px 10px; height: 95px;">
                      <tr>
                        <td width="32" valign="top">
                          <div style="background-color: #15803D; width: 30px; height: 30px; border-radius: 50%; text-align: center; line-height: 30px; font-size: 15px; font-weight: 900; color: #FFFFFF;">
                            $
                          </div>
                        </td>
                        <td valign="top" style="padding-left: 6px;">
                          <div style="font-size: 10.5px; color: #475569; font-weight: 600; line-height: 1.2; font-family: 'Segoe UI', Arial, sans-serif;">
                            Valor por facturar
                          </div>
                        </td>
                      </tr>
                      <tr>
                        <td colspan="2" valign="bottom" style="padding-top: 6px;">
                          <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0">
                            <tr>
                              <td valign="bottom" style="font-size: 17px; font-weight: 900; color: #15803D; font-family: 'Segoe UI', Arial, sans-serif; line-height: 1;">
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

                  <!-- KPI 3: Antigüedad promedio -->
                  <td width="23%" valign="top" class="kpi-cell" style="padding-right: 6px;">
                    <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 12px; padding: 12px 10px; height: 95px;">
                      <tr>
                        <td width="32" valign="top">
                          <div style="background-color: #D97706; width: 30px; height: 30px; border-radius: 50%; text-align: center; line-height: 30px; font-size: 15px; color: #FFFFFF;">
                            🕒
                          </div>
                        </td>
                        <td valign="top" style="padding-left: 6px;">
                          <div style="font-size: 10.5px; color: #475569; font-weight: 600; line-height: 1.2; font-family: 'Segoe UI', Arial, sans-serif;">
                            Antigüedad<br>promedio
                          </div>
                        </td>
                      </tr>
                      <tr>
                        <td colspan="2" valign="bottom" style="padding-top: 6px;">
                          <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0">
                            <tr>
                              <td valign="bottom" style="font-size: 20px; font-weight: 900; color: #D97706; font-family: 'Segoe UI', Arial, sans-serif; line-height: 1;">
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

                  <!-- KPI 4: Potencial de ventas por reconocer -->
                  <td width="27%" valign="top" class="kpi-cell">
                    <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 12px; padding: 12px 10px; height: 95px;">
                      <tr>
                        <td width="32" valign="top">
                          <div style="background-color: #1E40AF; width: 30px; height: 30px; border-radius: 50%; text-align: center; line-height: 30px; font-size: 15px; color: #FFFFFF;">
                            📊
                          </div>
                        </td>
                        <td valign="top" style="padding-left: 6px;">
                          <div style="font-size: 10.5px; color: #475569; font-weight: 600; line-height: 1.2; font-family: 'Segoe UI', Arial, sans-serif;">
                            Potencial de ventas<br>por reconocer
                          </div>
                        </td>
                      </tr>
                      <tr>
                        <td colspan="2" valign="bottom" style="padding-top: 6px;">
                          <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0">
                            <tr>
                              <td valign="bottom" style="font-size: 17px; font-weight: 900; color: #1E40AF; font-family: 'Segoe UI', Arial, sans-serif; line-height: 1;">
                                ${formatCOP(totalValue)}
                              </td>
                              <td align="right" valign="bottom" style="font-size: 17px; color: #93C5FD; opacity: 0.75;">
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

              <!-- TÍTULO DE LA TABLA -->
              <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="margin-bottom: 8px;">
                <tr>
                  <td>
                    <span style="font-size: 15px; font-weight: 850; color: #0F172A; font-family: 'Segoe UI', Arial, sans-serif;">
                      Tus remisiones destacadas
                    </span>
                    <span style="color: #D97706; font-size: 14px; margin-left: 4px;">☆</span>
                    <span style="font-size: 11.5px; color: #475569; margin-left: 12px; font-family: 'Segoe UI', Arial, sans-serif;">
                      Enfócate primero en las de mayor antigüedad.
                    </span>
                  </td>
                </tr>
              </table>

              <!-- TABLA DESTACADAS (3 FILAS) -->
              <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="border: 1px solid #E2E8F0; border-radius: 12px; overflow: hidden; border-collapse: separate; border-spacing: 0; margin-bottom: 14px;">
                <thead>
                  <tr style="background-color: #1E293B; color: #FFFFFF;">
                    <th style="padding: 10px 12px; text-align: left; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">Remisión</th>
                    <th style="padding: 10px 12px; text-align: left; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">Cliente</th>
                    <th style="padding: 10px 12px; text-align: right; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">Valor</th>
                    <th style="padding: 10px 12px; text-align: center; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">Días abierta</th>
                    <th style="padding: 10px 12px; text-align: center; font-size: 11px; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">Antigüedad</th>
                  </tr>
                </thead>
                <tbody style="background-color: #FFFFFF;">
                  ${rowsHtml}
                </tbody>
              </table>

              <!-- HERO BANNER: TOP OPORTUNIDAD DEL DÍA (DENTRO DE COLUMNA IZQUIERDA) -->
              ${topOpportunity ? `
              <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 12px; padding: 14px 16px;">
                <tr>
                  <!-- Trofeo dorado -->
                  <td width="48" valign="middle" align="center" style="padding-right: 12px;">
                    <span style="font-size: 34px;">🏆</span>
                  </td>

                  <!-- Info Remisión y Cliente -->
                  <td valign="middle" style="padding-right: 12px;">
                    <div style="font-size: 10.5px; font-weight: 700; color: #475569; font-family: 'Segoe UI', Arial, sans-serif;">
                      Top oportunidad del día <span style="color: #D97706;">☆</span>
                    </div>
                    <div style="font-size: 16px; font-weight: 900; color: #1E3A8A; font-family: 'Segoe UI', Arial, sans-serif; margin: 1px 0;">
                      ${topOpportunity.doc?.startsWith('REM-') ? topOpportunity.doc : `REM-${topOpportunity.doc || 'S/N'}`}
                    </div>
                    <div style="font-size: 11.5px; color: #475569; font-family: 'Segoe UI', Arial, sans-serif;">
                      ${topOpportunity.company || 'Cliente'}
                    </div>
                  </td>

                  <!-- Valor -->
                  <td width="115" valign="middle" style="padding-right: 12px;">
                    <div style="font-size: 16px; font-weight: 900; color: #1E3A8A; font-family: 'Segoe UI', Arial, sans-serif;">
                      ${formatCOP(topOpportunity.total)}
                    </div>
                    <div style="font-size: 10.5px; color: #64748B; font-family: 'Segoe UI', Arial, sans-serif;">
                      Valor
                    </div>
                  </td>

                  <!-- Antigüedad -->
                  <td width="85" valign="middle" style="padding-right: 12px;">
                    <div style="font-size: 16px; font-weight: 900; color: #DC2626; font-family: 'Segoe UI', Arial, sans-serif;">
                      ${topOpportunity.age} días
                    </div>
                    <div style="font-size: 10.5px; color: #64748B; font-family: 'Segoe UI', Arial, sans-serif;">
                      Antigüedad
                    </div>
                  </td>

                  <!-- Flecha y Call to Action -->
                  <td valign="middle" align="right">
                    <table role="presentation" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td valign="middle" style="padding-right: 6px; font-size: 20px; color: #1E3A8A;">
                          ⤷
                        </td>
                        <td valign="middle" style="font-size: 13px; font-weight: 850; font-style: italic; color: #1E3A8A; font-family: 'Segoe UI', Arial, sans-serif; line-height: 1.2;">
                          ¡Gran impacto<br>si la gestionas hoy!
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
              ` : ''}

            </td>

            <!-- ── COLUMNA DERECHA: SIDEBAR (29%) ── -->
            <td width="29%" valign="top" class="stack-col">

              <!-- CARD 1: CADA FACTURA CUENTA -->
              <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 14px; padding: 14px; margin-bottom: 12px;">
                <tr>
                  <td>
                    <div style="font-size: 14.5px; font-weight: 850; color: #166534; font-family: 'Segoe UI', Arial, sans-serif; margin-bottom: 6px;">
                      <span style="font-size: 15px;">★</span> Cada factura cuenta
                    </div>
                    <div style="font-size: 11.5px; color: #475569; font-family: 'Segoe UI', Arial, sans-serif; margin-bottom: 8px;">
                      Facturar a tiempo nos ayuda a:
                    </div>

                    <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="font-size: 11px; color: #334155; font-family: 'Segoe UI', Arial, sans-serif;">
                      <tr>
                        <td width="18" valign="top" style="padding-bottom: 6px; color: #16A34A; font-weight: 900; font-size: 12px;">✓</td>
                        <td valign="top" style="padding-bottom: 6px; line-height: 1.3;">Reflejar el valor de tu gestión comercial.</td>
                      </tr>
                      <tr>
                        <td width="18" valign="top" style="padding-bottom: 6px; color: #16A34A; font-weight: 900; font-size: 12px;">✓</td>
                        <td valign="top" style="padding-bottom: 6px; line-height: 1.3;">Cumplir nuestras metas juntos.</td>
                      </tr>
                      <tr>
                        <td width="18" valign="top" style="padding-bottom: 6px; color: #16A34A; font-weight: 900; font-size: 12px;">✓</td>
                        <td valign="top" style="padding-bottom: 6px; line-height: 1.3;">Mejorar flujo de caja y fortalecer nuestra operación.</td>
                      </tr>
                    </table>

                    <div style="font-size: 12px; font-weight: 800; color: #166534; font-family: 'Segoe UI', Arial, sans-serif; margin-top: 6px;">
                      ¡Sigamos logrando<br>grandes cosas juntos! <span style="font-size: 13px;">♡</span>
                    </div>
                  </td>
                </tr>
              </table>

              <!-- CARD 2: META DEL DÍA -->
              <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 14px; padding: 14px;">
                <tr>
                  <td>
                    <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="margin-bottom: 6px;">
                      <tr>
                        <td valign="middle">
                          <span style="font-size: 14px; margin-right: 4px;">🎯</span>
                          <span style="font-size: 13px; font-weight: 850; color: #1E3A8A; font-family: 'Segoe UI', Arial, sans-serif;">
                            Meta del día
                          </span>
                        </td>
                        <td align="right" valign="middle">
                          <span style="font-size: 18px;">🏆</span>
                        </td>
                      </tr>
                    </table>

                    <div style="font-size: 10.5px; color: #475569; line-height: 1.3; font-family: 'Segoe UI', Arial, sans-serif; margin-bottom: 8px;">
                      Gestionar las remisiones con mayor antigüedad para acelerar el reconocimiento de ventas.
                    </div>

                    <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="margin-bottom: 6px;">
                      <tr>
                        <td valign="middle" style="padding-right: 8px;">
                          <div style="background-color: #E5E7EB; border-radius: 999px; height: 9px; width: 100%; overflow: hidden;">
                            <div style="background-color: #16A34A; height: 9px; width: ${progressPct}%; border-radius: 999px;"></div>
                          </div>
                        </td>
                        <td width="34" align="right" valign="middle" style="font-size: 12.5px; font-weight: 850; color: #1E293B; font-family: 'Segoe UI', Arial, sans-serif;">
                          ${progressPct}%
                        </td>
                      </tr>
                    </table>

                    <div style="font-size: 11px; color: #334155; font-weight: 700; font-family: 'Segoe UI', Arial, sans-serif;">
                      Vamos por más, ¡tú puedes! 💪
                    </div>
                  </td>
                </tr>
              </table>

            </td>
          </tr>
        </table>

        <!-- ═════════════════════════════════════════════════════════════════════
             3. PIE DE PÁGINA: CORAZÓN AZUL + FIRMA + SOBRE POSTAL (FULL WIDTH)
             ═════════════════════════════════════════════════════════════════════ -->
        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #F8FAFC; border: 1.5px solid #E2E8F0; border-radius: 14px; padding: 16px 22px; margin-top: 18px;">
          <tr>
            <!-- Corazón azul grande -->
            <td width="38" valign="middle" align="center" style="padding-right: 12px;">
              <div style="font-size: 28px; line-height: 1; color: #1E3A8A;">
                💙
              </div>
            </td>

            <!-- Mensaje de agradecimiento -->
            <td valign="middle">
              <div style="font-size: 12px; color: #475569; font-family: 'Segoe UI', Arial, sans-serif;">
                Gracias por tu compromiso y dedicación.
              </div>
              <div style="font-size: 13.5px; font-weight: 850; color: #1E3A8A; font-family: 'Segoe UI', Arial, sans-serif; margin-top: 2px;">
                ¡Sigamos construyendo resultados juntos!
              </div>
            </td>

            <!-- Divisor vertical -->
            <td width="20" align="center" valign="middle" style="color: #CBD5E1; font-size: 22px; font-weight: 300;">
              |
            </td>

            <!-- Firma oficial -->
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
</html>
`;

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
      attachments: [
        {
          '@odata.type': '#microsoft.graph.fileAttachment',
          name: 'logo_provexpress.png',
          contentType: 'image/png',
          contentBytes: logoBase64,
          isInline: true,
          contentId: 'logo_provexpress',
        },
        {
          '@odata.type': '#microsoft.graph.fileAttachment',
          name: 'avatar_man.png',
          contentType: 'image/png',
          contentBytes: avatarBase64,
          isInline: true,
          contentId: 'avatar_man',
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
