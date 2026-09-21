/**
 * Script unificado para despachar las 3 pruebas solicitadas por el usuario con la fecha de hoy (2026-09-21):
 * 1. Ejecutivo Comercial Hombre: Mario Reyes (Grupo 1 - Rafael Novoa) · avatar_man.png · 3 KPIs · Excel individual
 * 2. Ejecutiva Comercial Mujer: Dayana Chala (Grupo 2 - Angélica Caballero) · avatar_woman.png · 3 KPIs · Excel individual
 * 3. Dirección y Gerencia Comercial: Rafael Novoa / Juan Novoa · Consolidado General · 4 Directores · Excel Maestro 3 hojas
 *
 * Destinatario de las 3 pruebas: especialista.preventa@provexpress.com.co
 * Remitente: juannovoa@provexpress.com.co (Microsoft Graph API)
 */

const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

const TARGET_TEST_EMAIL = 'especialista.preventa@provexpress.com.co';
const TODAY_CUTOFF = '2026-09-21';
const TODAY_FORMATTED = '21 de septiembre de 2026';
const DATA_SOURCE_CUTOFF = '2026-09-16'; // corte fuente de datos más reciente

const m365Dir = 'C:/Proyectos/m365-report-mailer';
const { ClientSecretCredential } = require(path.join(m365Dir, 'node_modules/@azure/identity'));
const axios = require(path.join(m365Dir, 'node_modules/axios'));

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
  console.error('❌ Error: Faltan credenciales de Microsoft Entra ID');
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

function normalizeName(val) {
  return String(val || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

const logoPath = path.join(__dirname, '../public/logo_provexpress_stacked.png');
const avatarManPath = path.join(__dirname, '../public/avatar_man.png');
const avatarWomanPath = path.join(__dirname, '../public/avatar_woman.png');
const logoBase64 = fs.readFileSync(logoPath).toString('base64');
const avatarManBase64 = fs.readFileSync(avatarManPath).toString('base64');
const avatarWomanBase64 = fs.readFileSync(avatarWomanPath).toString('base64');

async function sendMail(token, subject, htmlContent, attachments) {
  const graphSendUrl = `https://graph.microsoft.com/v1.0/users/${senderEmail}/sendMail`;
  const payload = {
    message: {
      subject,
      body: { contentType: 'HTML', content: htmlContent },
      toRecipients: [{ emailAddress: { address: TARGET_TEST_EMAIL } }],
      attachments,
    },
    saveToSentItems: false,
  };
  const res = await axios.post(graphSendUrl, payload, {
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
  });
  return res.status;
}

// Generador de Excel individual para comercial
async function generateCommercialExcel(name, items) {
  const excelWb = new ExcelJS.Workbook();
  excelWb.creator = 'Provexpress SAS - Sistema de Remisiones';
  const wsOpen = excelWb.addWorksheet('Remisiones Abiertas', { views: [{ state: 'frozen', ySplit: 6 }] });

  wsOpen.mergeCells('A1:L1');
  const tCell = wsOpen.getCell('A1');
  tCell.value = 'PROVEXPRESS SAS · GESTIÓN COMERCIAL DE REMISIONES ABIERTAS';
  tCell.font = { name: 'Calibri', size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
  tCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
  tCell.alignment = { horizontal: 'center', vertical: 'middle' };
  wsOpen.getRow(1).height = 30;

  wsOpen.getCell('A2').value = 'Asesor Comercial:';
  wsOpen.getCell('B2').value = name;
  wsOpen.getCell('E2').value = 'Fecha de Corte:';
  wsOpen.getCell('F2').value = TODAY_FORMATTED;
  wsOpen.getCell('A3').value = 'Total Documentos:';
  wsOpen.getCell('B3').value = items.length;
  wsOpen.getCell('E3').value = 'Total por Facturar:';
  wsOpen.getCell('F3').value = items.reduce((s, x) => s + x.total, 0);
  wsOpen.getCell('F3').numFmt = '"$"#,##0';

  const headers = [
    { title: 'N°', width: 6, align: 'center' },
    { title: 'Remisión', width: 15, align: 'center' },
    { title: 'Fecha Emisión', width: 14, align: 'center' },
    { title: 'Días Abierta', width: 13, align: 'right' },
    { title: 'Rango Antigüedad', width: 18, align: 'center' },
    { title: 'NIT', width: 15, align: 'left' },
    { title: 'Cliente / Razón Social', width: 36, align: 'left' },
    { title: 'Vr. Mercancía', width: 16, align: 'right' },
    { title: 'Vr. IVA', width: 14, align: 'right' },
    { title: 'Vr. Total', width: 16, align: 'right' },
    { title: 'Pedido', width: 14, align: 'left' },
    { title: 'Cantidad', width: 11, align: 'right' },
  ];

  const hRow = wsOpen.getRow(6);
  hRow.height = 24;
  headers.forEach((h, i) => {
    const c = hRow.getCell(i + 1);
    c.value = h.title;
    c.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F172A' } };
    c.alignment = { vertical: 'middle', horizontal: h.align };
    wsOpen.getColumn(i + 1).width = h.width;
  });

  items.forEach((item, idx) => {
    const r = wsOpen.getRow(7 + idx);
    r.height = 20;
    r.getCell(1).value = idx + 1;
    r.getCell(2).value = item.doc?.startsWith('REM-') ? item.doc : `REM-${item.doc || 'S/N'}`;
    r.getCell(2).font = { bold: true, color: { argb: 'FF1E3A8A' } };
    r.getCell(3).value = TODAY_CUTOFF;
    r.getCell(4).value = item.age;
    r.getCell(4).font = { bold: true, color: item.age > 15 ? { argb: 'FFDC2626' } : { argb: 'FF0F172A' } };
    r.getCell(5).value = item.age > 15 ? 'Más de 15 días' : (item.age >= 8 ? '8 a 15 días' : '0 a 7 días');
    r.getCell(6).value = item.nit;
    r.getCell(7).value = item.company;
    r.getCell(8).value = item.total;
    r.getCell(8).numFmt = '"$"#,##0';
    r.getCell(9).value = 0;
    r.getCell(9).numFmt = '"$"#,##0';
    r.getCell(10).value = item.total;
    r.getCell(10).numFmt = '"$"#,##0';
    r.getCell(10).font = { bold: true, color: { argb: 'FF15803D' } };
    r.getCell(11).value = item.order || '';
    r.getCell(12).value = 1;

    for (let col = 1; col <= 12; col++) {
      r.getCell(col).border = { top: { style: 'thin', color: { argb: 'FFE2E8F0' } }, bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } } };
    }
  });

  const buf = await excelWb.xlsx.writeBuffer();
  return Buffer.from(buf).toString('base64');
}

// Plantilla HTML corporativa de comercial con 3 KPIs
function generateCommercialHtml(name, items, gender = 'M', dirName = 'Rafael Novoa') {
  const isFemale = gender === 'F';
  const avatarCid = isFemale ? 'cid:avatar_woman' : 'cid:avatar_man';
  const avatarAlt = isFemale ? 'Asesora Comercial' : 'Asesor Comercial';

  const byAge = [...items].sort((a, b) => b.age - a.age);
  const destacadas = byAge.slice(0, 3);
  const topMayorAntiguedad = byAge[0] || null;

  const byVal = [...items].sort((a, b) => b.total - a.total);
  const topMayorValor = byVal[0] || null;

  const totalCount = items.length;
  const totalValue = items.reduce((s, x) => s + x.total, 0);
  const avgAge = totalCount > 0 ? Math.round(items.reduce((s, x) => s + x.age, 0) / totalCount) : 0;
  const onTrack = items.filter((x) => x.age <= 15).length;
  const progressPct = totalCount > 0 ? Math.round((onTrack / totalCount) * 100) : 100;

  const rowsDestacadas = destacadas.map((r) => {
    let badgeBg = '#DCFCE7';
    let badgeText = '#15803D';
    let badgeLabel = '0 a 7 días';
    if (r.age > 15) {
      badgeBg = '#FEE2E2';
      badgeText = '#DC2626';
      badgeLabel = 'Más de 15 días';
    } else if (r.age >= 8) {
      badgeBg = '#FEF3C7';
      badgeText = '#B45309';
      badgeLabel = '8 a 15 días';
    }

    return `
      <tr style="border-bottom: 1px solid #F1F5F9;">
        <td style="padding: 10px 12px; font-family: 'Segoe UI', Arial, sans-serif;">
          <strong style="color: #0F172A; font-size: 12.5px;">${r.doc?.startsWith('REM-') ? r.doc : `REM-${r.doc || 'S/N'}`}</strong>
        </td>
        <td style="padding: 10px 12px; font-size: 12.5px; color: #334155; font-family: 'Segoe UI', Arial, sans-serif;">
          ${r.company || 'Cliente'}
        </td>
        <td style="padding: 10px 12px; text-align: right; font-weight: 850; color: #15803D; font-size: 13.5px; font-family: 'Segoe UI', Arial, sans-serif;">
          ${formatCOP(r.total)}
        </td>
        <td style="padding: 10px 12px; text-align: center; font-weight: 700; color: #1E293B; font-size: 13px; font-family: 'Segoe UI', Arial, sans-serif;">
          ${r.age}
        </td>
        <td style="padding: 10px 12px; text-align: center;">
          <span style="background-color: ${badgeBg}; color: ${badgeText}; padding: 3px 10px; border-radius: 999px; font-size: 10.5px; font-weight: 700; display: inline-block;">
            ${badgeLabel}
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
  <title>Oportunidades listas para convertirse en ventas facturadas · Provexpress SAS</title>
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

  <table role="presentation" class="main-card" width="960" align="center" border="0" cellpadding="0" cellspacing="0" style="width: 960px; margin: 0 auto; background-color: #FFFFFF; border-radius: 16px; overflow: hidden; border: 1.5px solid #E2E8F0; box-shadow: 0 4px 20px rgba(15, 23, 42, 0.05);">
    <tr>
      <td style="padding: 24px 28px;">

        <!-- 1. ENCABEZADO -->
        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="margin-bottom: 20px;">
          <tr>
            <td width="160" valign="middle" align="left" style="padding-right: 18px;">
              <img src="cid:logo_provexpress" width="144" height="110" style="display: block; border: 0; width: 144px; height: 110px; max-width: 144px;" alt="Provexpress" />
            </td>

            <td valign="middle" style="padding-right: 12px;">
              <div style="font-size: 15px; color: #475569; font-weight: 600; font-family: 'Segoe UI', Arial, sans-serif;">
                Hola, <strong style="color: #0F172A; font-size: 16px;">${name}</strong>
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

            <td width="190" valign="bottom" align="center" style="padding-right: 10px;">
              <img src="${avatarCid}" width="180" height="154" style="display: block; border: 0; width: 180px; height: auto; margin: 0 auto;" alt="${avatarAlt}" />
            </td>

            <td width="125" valign="middle" align="right">
              <table role="presentation" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 14px; padding: 14px 8px; width: 120px; text-align: center;">
                <tr>
                  <td align="center">
                    <div style="background-color: #16A34A; color: #FFFFFF; width: 26px; height: 26px; border-radius: 50%; font-size: 14px; font-weight: 900; line-height: 26px; margin: 0 auto 8px auto;">
                      ✓
                    </div>
                    <div style="font-size: 12px; font-weight: 700; color: #1E293B; line-height: 1.3; font-family: 'Segoe UI', Arial, sans-serif;">
                      Tu gestión<br>hace la<br><strong style="color: #16A34A; font-size: 13px;">diferencia.</strong>
                    </div>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>

        <!-- 2. CUERPO EN 2 COLUMNAS (3 KPIS + DESTACADAS // CADA FACTURA CUENTA) -->
        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td width="71%" valign="top" style="padding-right: 18px;" class="stack-col">

              <!-- FILA DE 3 TARJETAS KPI -->
              <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="margin-bottom: 16px;">
                <tr>
                  <td width="31%" valign="top" class="kpi-cell" style="padding-right: 8px;">
                    <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 12px; padding: 12px 10px; height: 95px;">
                      <tr>
                        <td width="32" valign="top">
                          <div style="background-color: #312E81; width: 30px; height: 30px; border-radius: 50%; text-align: center; line-height: 30px; font-size: 15px; color: #FFFFFF;">📦</div>
                        </td>
                        <td valign="top" style="padding-left: 6px;">
                          <div style="font-size: 10.5px; color: #475569; font-weight: 600; line-height: 1.2; font-family: 'Segoe UI', Arial, sans-serif;">Remisiones listas<br>para gestionar</div>
                        </td>
                      </tr>
                      <tr>
                        <td colspan="2" valign="bottom" style="padding-top: 6px;">
                          <div style="font-size: 24px; font-weight: 900; color: #312E81; font-family: 'Segoe UI', Arial, sans-serif; line-height: 1;">${formatNumber(totalCount)}</div>
                        </td>
                      </tr>
                    </table>
                  </td>

                  <td width="38%" valign="top" class="kpi-cell" style="padding-right: 8px;">
                    <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 12px; padding: 12px 10px; height: 95px;">
                      <tr>
                        <td width="32" valign="top">
                          <div style="background-color: #15803D; width: 30px; height: 30px; border-radius: 50%; text-align: center; line-height: 30px; font-size: 15px; font-weight: 900; color: #FFFFFF;">$</div>
                        </td>
                        <td valign="top" style="padding-left: 6px;">
                          <div style="font-size: 10.5px; color: #475569; font-weight: 600; line-height: 1.2; font-family: 'Segoe UI', Arial, sans-serif;">Valor por facturar</div>
                        </td>
                      </tr>
                      <tr>
                        <td colspan="2" valign="bottom" style="padding-top: 6px;">
                          <div style="font-size: 18px; font-weight: 900; color: #15803D; font-family: 'Segoe UI', Arial, sans-serif; line-height: 1;">${formatCOP(totalValue)}</div>
                        </td>
                      </tr>
                    </table>
                  </td>

                  <td width="31%" valign="top" class="kpi-cell">
                    <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 12px; padding: 12px 10px; height: 95px;">
                      <tr>
                        <td width="32" valign="top">
                          <div style="background-color: #D97706; width: 30px; height: 30px; border-radius: 50%; text-align: center; line-height: 30px; font-size: 15px; color: #FFFFFF;">🕒</div>
                        </td>
                        <td valign="top" style="padding-left: 6px;">
                          <div style="font-size: 10.5px; color: #475569; font-weight: 600; line-height: 1.2; font-family: 'Segoe UI', Arial, sans-serif;">Antigüedad<br>promedio</div>
                        </td>
                      </tr>
                      <tr>
                        <td colspan="2" valign="bottom" style="padding-top: 6px;">
                          <div style="font-size: 22px; font-weight: 900; color: #D97706; font-family: 'Segoe UI', Arial, sans-serif; line-height: 1;">${avgAge} días</div>
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>

              <!-- TABLA DESTACADAS -->
              <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="border: 1px solid #E2E8F0; border-radius: 12px; overflow: hidden; border-collapse: separate; border-spacing: 0;">
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
                  ${rowsDestacadas}
                </tbody>
              </table>

            </td>

            <!-- SIDEBAR DERECHO -->
            <td width="29%" valign="top" class="stack-col">
              <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 14px; padding: 14px; margin-bottom: 12px;">
                <tr>
                  <td>
                    <div style="font-size: 14.5px; font-weight: 850; color: #166534; font-family: 'Segoe UI', Arial, sans-serif; margin-bottom: 6px;">
                      <span style="font-size: 15px;">★</span> Cada factura cuenta
                    </div>
                    <div style="font-size: 11.5px; color: #475569; font-family: 'Segoe UI', Arial, sans-serif; margin-bottom: 8px;">
                      Facturar a tiempo nos ayuda a:
                    </div>
                    <div style="font-size: 11px; color: #334155; line-height: 1.4;">
                      ✓ Reflejar el valor de tu gestión comercial.<br>
                      ✓ Cumplir nuestras metas juntos.<br>
                      ✓ Garantizar una operación ágil y sólida.
                    </div>
                  </td>
                </tr>
              </table>

              <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #F8FAFC; border: 1.5px solid #E2E8F0; border-radius: 14px; padding: 14px;">
                <tr>
                  <td>
                    <div style="font-size: 13.5px; font-weight: 850; color: #1E293B; font-family: 'Segoe UI', Arial, sans-serif; margin-bottom: 6px;">
                      🎯 Meta del día
                    </div>
                    <div style="font-size: 11px; color: #64748B; margin-bottom: 8px;">
                      ${onTrack} de ${totalCount} remisiones al día (&le; 15 días)
                    </div>
                    <div style="background-color: #E2E8F0; border-radius: 999px; height: 8px; overflow: hidden; margin-bottom: 8px;">
                      <div style="background-color: #16A34A; height: 8px; width: ${progressPct}%;"></div>
                    </div>
                    <div style="font-size: 11px; font-weight: 700; color: #15803D;">
                      ¡Tu esfuerzo impulsa a Provexpress!
                    </div>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>

        <!-- 3. TOP OPORTUNIDADES (ANCHO COMPLETO 100%) -->
        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #FFFFFF; border: 1.5px solid #E2E8F0; border-radius: 14px; padding: 14px 18px; margin-top: 14px; margin-bottom: 14px;">
          ${topMayorValor ? `
          <tr>
            <td width="42" valign="middle" align="center" style="padding-right: 12px;">
              <span style="font-size: 26px;">🏆</span>
            </td>
            <td valign="middle" style="padding-right: 14px;">
              <div style="font-size: 10px; font-weight: 700; color: #15803D; text-transform: uppercase; letter-spacing: 0.04em;">Mayor valor por facturar ★</div>
              <div style="font-size: 15px; font-weight: 900; color: #1E3A8A;">${topMayorValor.doc?.startsWith('REM-') ? topMayorValor.doc : `REM-${topMayorValor.doc || 'S/N'}`}</div>
              <div style="font-size: 11px; color: #475569;">${topMayorValor.company}</div>
            </td>
            <td width="130" valign="middle" style="padding-right: 14px;">
              <div style="font-size: 15px; font-weight: 900; color: #15803D;">${formatCOP(topMayorValor.total)}</div>
              <div style="font-size: 10px; color: #64748B;">Valor</div>
            </td>
            <td width="95" valign="middle" style="padding-right: 14px;">
              <div style="font-size: 14.5px; font-weight: 900; color: #0F172A;">${topMayorValor.age} días</div>
              <div style="font-size: 10px; color: #64748B;">Antigüedad</div>
            </td>
            <td valign="middle" align="right">
              <span style="font-size: 11.5px; font-weight: 850; font-style: italic; color: #15803D;">¡Mayor impacto en facturación!</span>
            </td>
          </tr>
          ` : ''}

          ${(topMayorAntiguedad && (!topMayorValor || topMayorAntiguedad.doc !== topMayorValor.doc)) ? `
          <tr><td colspan="5" style="padding: 10px 0;"><div style="border-top: 1px dashed #E2E8F0; height: 1px;"></div></td></tr>
          <tr>
            <td width="42" valign="middle" align="center" style="padding-right: 12px;">
              <span style="font-size: 26px;">⏳</span>
            </td>
            <td valign="middle" style="padding-right: 14px;">
              <div style="font-size: 10px; font-weight: 700; color: #DC2626; text-transform: uppercase; letter-spacing: 0.04em;">Mayor antigüedad pendiente ⏱</div>
              <div style="font-size: 15px; font-weight: 900; color: #1E3A8A;">${topMayorAntiguedad.doc?.startsWith('REM-') ? topMayorAntiguedad.doc : `REM-${topMayorAntiguedad.doc || 'S/N'}`}</div>
              <div style="font-size: 11px; color: #475569;">${topMayorAntiguedad.company}</div>
            </td>
            <td width="130" valign="middle" style="padding-right: 14px;">
              <div style="font-size: 15px; font-weight: 900; color: #1E3A8A;">${formatCOP(topMayorAntiguedad.total)}</div>
              <div style="font-size: 10px; color: #64748B;">Valor</div>
            </td>
            <td width="95" valign="middle" style="padding-right: 14px;">
              <div style="font-size: 14.5px; font-weight: 900; color: #DC2626;">${topMayorAntiguedad.age} días</div>
              <div style="font-size: 10px; color: #64748B;">Mayor antigüedad</div>
            </td>
            <td valign="middle" align="right">
              <span style="font-size: 11.5px; font-weight: 850; font-style: italic; color: #DC2626;">¡Prioridad urgente por tiempo!</span>
            </td>
          </tr>
          ` : ''}
        </table>

        <!-- 4. BANNER ARCHIVO ADJUNTO (ANCHO COMPLETO 100%) -->
        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #F8FAFC; border: 1.5px solid #E2E8F0; border-radius: 12px; padding: 12px 18px; margin-bottom: 14px;">
          <tr>
            <td width="38" valign="middle" align="center" style="padding-right: 12px;">
              <div style="background-color: #107C41; color: #FFFFFF; width: 32px; height: 32px; border-radius: 8px; text-align: center; line-height: 32px; font-size: 16px; font-weight: 900;">📊</div>
            </td>
            <td valign="middle">
              <div style="font-size: 12px; font-weight: 800; color: #0F172A;">
                📎 Archivo adjunto: Remisiones_Abiertas_${name.replace(/\s+/g, '_')}_${TODAY_CUTOFF}.xlsx
              </div>
              <div style="font-size: 11px; color: #475569; margin-top: 2px;">
                Hemos adjuntado a este correo tu archivo Excel con el detalle completo de tus <strong>${formatNumber(totalCount)}</strong> remisiones abiertas (${formatCOP(totalValue)}) para tu gestión y descarga.
              </div>
            </td>
          </tr>
        </table>

        <!-- 5. PIE DE PÁGINA -->
        <table role="presentation" width="100%" border="0" cellpadding="0" cellspacing="0" style="background-color: #F8FAFC; border: 1.5px solid #E2E8F0; border-radius: 14px; padding: 16px 22px; margin-top: 18px;">
          <tr>
            <td width="38" valign="middle" align="center" style="padding-right: 12px;">
              <div style="font-size: 28px; line-height: 1; color: #1E3A8A;">💙</div>
            </td>
            <td valign="middle">
              <div style="font-size: 12px; color: #475569;">Gracias por tu compromiso y dedicación.</div>
              <div style="font-size: 13.5px; font-weight: 850; color: #1E3A8A; margin-top: 2px;">¡Sigamos construyendo resultados juntos!</div>
            </td>
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

async function main() {
  console.log('================================================================');
  console.log('🚀 ENVIANDO LAS 3 PRUEBAS SOLICITADAS CON LA FECHA DE HOY (21/09/2026)');
  console.log(`📩 Destinatario común: ${TARGET_TEST_EMAIL}`);
  console.log(`📤 Remitente institucional: ${senderEmail}`);
  console.log('================================================================');

  // Autenticar
  const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
  const tokenResp = await credential.getToken('https://graph.microsoft.com/.default');
  const token = tokenResp.token;
  console.log('✓ Token de autenticación Microsoft Graph obtenido exitosamente.');

  // Cargar libro Excel real
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(path.join(__dirname, '../Remisiones.xlsx'));
  const sis = wb.getWorksheet('Base-SIS');

  const marioItems = [];
  const dayanaItems = [];

  for (let r = 2; r <= sis.rowCount; r++) {
    const row = sis.getRow(r);
    const dVal = row.getCell(1).value;
    const iso = dVal instanceof Date ? dVal.toISOString().slice(0, 10) : String(dVal || '').trim();
    if (iso === DATA_SOURCE_CUTOFF) {
      const emp = String(row.getCell(2).value || '').trim();
      const nit = String(row.getCell(3).value || '').trim();
      const company = String(row.getCell(4).value || '').trim();
      const total = Number(row.getCell(7).value || 0);
      const age = Number(row.getCell(9).value || 0);
      const doc = String(row.getCell(10).value || '').trim();
      const order = String(row.getCell(11).value || '').trim();

      if (normalizeName(emp).includes('mario reyes')) {
        marioItems.push({ doc, nit, company, total, age, order });
      }
      if (normalizeName(emp).includes('dayana')) {
        dayanaItems.push({ doc, nit, company, total, age, order });
      }
    }
  }

  // ───────────────────────────────────────────────────────────────────────────
  // 1. EJECUTIVO COMERCIAL HOMBRE: Mario Reyes (Grupo 1 - Rafael Novoa)
  // ───────────────────────────────────────────────────────────────────────────
  console.log('\n----------------------------------------------------------------');
  console.log(`📨 1/3: Enviando prueba EJECUTIVO HOMBRE (Mario Reyes · Grupo 1 · ${marioItems.length} remisiones)...`);
  const marioExcelBase64 = await generateCommercialExcel('Mario Reyes', marioItems);
  const marioHtml = generateCommercialHtml('Mario Reyes', marioItems, 'M', 'Rafael Novoa');
  const marioFilename = `Remisiones_Abiertas_Mario_Reyes_${TODAY_CUTOFF}.xlsx`;

  await sendMail(
    token,
    `[PRUEBA] Oportunidades de Facturación · Mario Reyes (Grupo 1) · Corte ${TODAY_FORMATTED}`,
    marioHtml,
    [
      { '@odata.type': '#microsoft.graph.fileAttachment', name: 'logo_provexpress.png', contentType: 'image/png', contentBytes: logoBase64, isInline: true, contentId: 'logo_provexpress' },
      { '@odata.type': '#microsoft.graph.fileAttachment', name: 'avatar_man.png', contentType: 'image/png', contentBytes: avatarManBase64, isInline: true, contentId: 'avatar_man' },
      { '@odata.type': '#microsoft.graph.fileAttachment', name: marioFilename, contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', contentBytes: marioExcelBase64, isInline: false },
    ]
  );
  console.log('   ✅ Correo de Mario Reyes (Hombre · Grupo 1) enviado exitosamente.');

  // Pausa preventiva de 1.5s
  await new Promise((resolve) => setTimeout(resolve, 1500));

  // ───────────────────────────────────────────────────────────────────────────
  // 2. EJECUTIVA COMERCIAL MUJER: Dayana Chala (Grupo 2 - Angélica Caballero)
  // ───────────────────────────────────────────────────────────────────────────
  console.log('\n----------------------------------------------------------------');
  console.log(`📨 2/3: Enviando prueba EJECUTIVA MUJER (Dayana Chala · Grupo 2 · ${dayanaItems.length} remisiones)...`);
  const dayanaExcelBase64 = await generateCommercialExcel('Dayana Chala', dayanaItems);
  const dayanaHtml = generateCommercialHtml('Dayana Chala', dayanaItems, 'F', 'Angélica Caballero');
  const dayanaFilename = `Remisiones_Abiertas_Dayana_Chala_${TODAY_CUTOFF}.xlsx`;

  await sendMail(
    token,
    `[PRUEBA] Oportunidades de Facturación · Dayana Chala (Grupo 2) · Corte ${TODAY_FORMATTED}`,
    dayanaHtml,
    [
      { '@odata.type': '#microsoft.graph.fileAttachment', name: 'logo_provexpress.png', contentType: 'image/png', contentBytes: logoBase64, isInline: true, contentId: 'logo_provexpress' },
      { '@odata.type': '#microsoft.graph.fileAttachment', name: 'avatar_woman.png', contentType: 'image/png', contentBytes: avatarWomanBase64, isInline: true, contentId: 'avatar_woman' },
      { '@odata.type': '#microsoft.graph.fileAttachment', name: dayanaFilename, contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', contentBytes: dayanaExcelBase64, isInline: false },
    ]
  );
  console.log('   ✅ Correo de Dayana Chala (Mujer · Grupo 2 · Nueva muñeca) enviado exitosamente.');

  // Pausa preventiva de 1.5s
  await new Promise((resolve) => setTimeout(resolve, 1500));

  const { execSync } = require('child_process');

  // ───────────────────────────────────────────────────────────────────────────
  // 3. DIRECTORA DE GRUPO COMERCIAL: Angélica Caballero (Grupo 2)
  // ───────────────────────────────────────────────────────────────────────────
  console.log('\n----------------------------------------------------------------');
  console.log('📨 3/4: Enviando prueba DIRECTORA DE GRUPO (Angélica Caballero · Grupo 2)...');
  const directorOutput = execSync(`node "${path.join(__dirname, 'send-test-director.cjs')}"`, { encoding: 'utf8' });
  console.log(directorOutput);

  // Pausa preventiva de 1.5s
  await new Promise((resolve) => setTimeout(resolve, 1500));

  // ───────────────────────────────────────────────────────────────────────────
  // 4. DIRECCIÓN Y GERENCIA COMERCIAL ("Los Jefes Jefes")
  // ───────────────────────────────────────────────────────────────────────────
  console.log('\n----------------------------------------------------------------');
  console.log('📨 4/4: Enviando prueba DIRECCIÓN Y GERENCIA COMERCIAL (Consolidado General 4 Grupos)...');
  const gerenciaOutput = execSync(`node "${path.join(__dirname, 'send-test-gerencia.cjs')}"`, { encoding: 'utf8' });
  console.log(gerenciaOutput);

  console.log('================================================================');
  console.log('🎉 ¡TODAS LAS PRUEBAS FUERON DESPACHADAS EXITOSAMENTE!');
  console.log(`   1. Ejecutivo Hombre: Mario Reyes (Grupo 1) con fecha ${TODAY_CUTOFF}`);
  console.log(`   2. Ejecutiva Mujer: Dayana Chala (Grupo 2) con fecha ${TODAY_CUTOFF}`);
  console.log(`   3. Directora de Grupo: Angélica Caballero (Grupo 2) con fecha ${TODAY_CUTOFF}`);
  console.log(`   4. Informe Gerencial General: Rafael & Juan Novoa con fecha ${TODAY_CUTOFF}`);
  console.log(`   Destino de todas: ${TARGET_TEST_EMAIL}`);
  console.log('================================================================');
}

main().catch((err) => {
  console.error('❌ Error enviando las pruebas:', err.response?.data || err.message);
  process.exit(1);
});
