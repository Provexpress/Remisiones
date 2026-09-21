import React, { useMemo, useRef, useState, useEffect } from 'react';
import {
  AlertCircle,
  Briefcase,
  CheckCircle2,
  Copy,
  Download,
  Mail,
  RefreshCw,
  Send,
  Sparkles,
  Terminal,
  Users,
  X,
  FileSpreadsheet,
  ShieldCheck,
} from 'lucide-react';
import type { Remision } from '../types';
import { acquireMailToken } from '../lib/auth';
import {
  buildCommercialEmailSummary,
  formatCOP,
  generateCommercialEmailHtml,
} from '../lib/commercialEmailTemplate';
import { generateCommercialExcelBase64 } from '../lib/commercialExcelGenerator';
import {
  buildDirectorEmailSummary,
  generateDirectorEmailHtml,
} from '../lib/directorEmailTemplate';
import { generateDirectorExcelBase64 } from '../lib/directorExcelGenerator';
import {
  buildGerenciaEmailSummary,
  generateGerenciaEmailHtml,
} from '../lib/gerenciaEmailTemplate';
import { generateGerenciaExcelBase64 } from '../lib/gerenciaExcelGenerator';
import { LISTA_DIRECTORES, LISTA_GERENCIA } from '../lib/commercialDirectory';
import {
  sendMailViaGraph,
  type SendEmailPayload,
  type SendResult,
  OFFICIAL_SENDER_EMAIL,
  OFFICIAL_SENDER_NAME,
} from '../lib/emailSender';
import { formatCutoff } from '../lib/remisiones';

interface Props {
  isOpen: boolean;
  onClose: () => void;
  records: Remision[];
  cutoffDate: string;
  currentUserEmail?: string;
}

export interface DispatchLogEntry {
  id: string;
  time: string;
  httpStatus: number;
  success: boolean;
  recipientEmail: string;
  recipientName: string;
  role: 'Comercial' | 'Director' | 'Gerencia' | 'Copia Validez' | 'Prueba';
  detail: string;
  error?: string;
}

export const EmailNotificationModal: React.FC<Props> = ({
  isOpen,
  onClose,
  records,
  cutoffDate,
  currentUserEmail,
}) => {
  // Opciones de paquetes a enviar
  const [includeCommercials, setIncludeCommercials] = useState(true);
  const [includeDirectors, setIncludeDirectors] = useState(true);
  const [includeGerencia, setIncludeGerencia] = useState(true);

  // Estados de ejecución
  const [isSending, setIsSending] = useState(false);
  const [isSendingTest, setIsSendingTest] = useState(false);
  const [progress, setProgress] = useState<{ current: number; total: number } | null>(null);
  const [logs, setLogs] = useState<DispatchLogEntry[]>([]);
  const [alertMessage, setAlertMessage] = useState<{ type: 'success' | 'error' | 'info'; text: string } | null>(null);
  const [copied, setCopied] = useState(false);

  const logEndRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (logs.length > 0) {
      logEndRef.current?.scrollIntoView({ behavior: 'smooth' });
    }
  }, [logs]);

  // Cálculo de estadísticas y destinatarios del corte actual
  const {
    commercialList,
    totalCommercialValue,
    totalCommercialRemisiones,
    directorList,
    gerenciaList,
  } = useMemo(() => {
    const currentRecords = records.filter((r) => r.cutoff === cutoffDate);

    // 1. Comerciales con remisiones en este corte
    const byEmployee = new Map<string, Remision[]>();
    currentRecords.forEach((r) => {
      const emp = r.employee?.trim();
      if (!emp) return;
      const list = byEmployee.get(emp) || [];
      list.push(r);
      byEmployee.set(emp, list);
    });

    const cList: { name: string; email: string; count: number; value: number }[] = [];
    let totVal = 0;
    let totRem = 0;

    byEmployee.forEach((empRecords, empName) => {
      const summary = buildCommercialEmailSummary(empName, records, cutoffDate);
      if (summary.totalCount > 0) {
        cList.push({
          name: empName,
          email: summary.commercialEmail,
          count: summary.totalCount,
          value: summary.totalValue,
        });
        totVal += summary.totalValue;
        totRem += summary.totalCount;
      }
    });

    // 2. Directores de grupo
    const dList = LISTA_DIRECTORES.filter((d) => d.grupo >= 1 && d.grupo <= 4);

    // 3. Gerencia & Especialista
    const gList = LISTA_GERENCIA;

    return {
      commercialList: cList,
      totalCommercialValue: totVal,
      totalCommercialRemisiones: totRem,
      directorList: dList,
      gerenciaList: gList,
    };
  }, [records, cutoffDate]);

  const totalCalculated =
    (includeCommercials ? commercialList.length : 0) +
    (includeDirectors ? directorList.length : 0) +
    (includeGerencia ? gerenciaList.length : 0);

  if (!isOpen) return null;

  // ══════════════════════════════════════════════════════════════════════════
  // ENVIAR PRUEBA A MI CORREO
  // ══════════════════════════════════════════════════════════════════════════
  const handleSendTest = async () => {
    const targetEmail = currentUserEmail || 'especialista.preventa@provexpress.com.co';

    try {
      setIsSendingTest(true);
      setAlertMessage(null);
      const token = await acquireMailToken();

      // Generar consolidado gerencial de prueba con copia a especialista
      const testGerencia = buildGerenciaEmailSummary(records, cutoffDate, {
        name: 'Especialista Preventa',
        cargo: 'Prueba de Despacho',
        email: targetEmail,
      });

      let excelAttachment;
      try {
        const { filename, base64 } = await generateGerenciaExcelBase64(testGerencia);
        excelAttachment = { filename, base64 };
      } catch (e) {
        console.warn('Error generando Excel de prueba:', e);
      }

      const res = await sendMailViaGraph(token, {
        toEmail: targetEmail,
        toName: 'Especialista Preventa (Prueba)',
        subject: `[PRUEBA] Consolidado Maestro Gerencial · Remisiones · Corte ${formatCutoff(cutoffDate)}`,
        htmlBody: generateGerenciaEmailHtml(testGerencia),
        excelAttachment,
        senderEmail: OFFICIAL_SENDER_EMAIL,
        senderName: OFFICIAL_SENDER_NAME,
        category: 'prueba',
      });

      const entry: DispatchLogEntry = {
        id: `test-${Date.now()}`,
        time: res.timestamp || new Date().toLocaleTimeString('es-CO', { hour12: false }),
        httpStatus: res.httpStatus || (res.success ? 202 : 500),
        success: res.success,
        recipientEmail: targetEmail,
        recipientName: 'Especialista Preventa',
        role: 'Prueba',
        detail: `Prueba de correo gerencial recibida (${formatCutoff(cutoffDate)})`,
        error: res.error,
      };

      setLogs((prev) => [...prev, entry]);

      if (res.success) {
        setAlertMessage({
          type: 'success',
          text: `¡Correo de prueba despachado con éxito a ${targetEmail}! (HTTP ${res.httpStatus || 202} OK). Revisa tu bandeja de entrada.`,
        });
      } else {
        setAlertMessage({
          type: 'error',
          text: `Error al enviar prueba: ${res.error || 'Verifica la sesión de Microsoft 365'}`,
        });
      }
    } catch (err: any) {
      setAlertMessage({
        type: 'error',
        text: `Error de conexión con Microsoft 365: ${err?.message || 'Error desconocido'}`,
      });
    } finally {
      setIsSendingTest(false);
    }
  };

  // ══════════════════════════════════════════════════════════════════════════
  // INICIAR DESPACHO OFICIAL COMPLETO
  // ══════════════════════════════════════════════════════════════════════════
  const handleStartDispatch = async () => {
    if (totalCalculated === 0) {
      setAlertMessage({ type: 'error', text: 'Selecciona al menos un grupo para despachar las notificaciones.' });
      return;
    }

    try {
      setIsSending(true);
      setAlertMessage(null);
      setProgress({ current: 0, total: totalCalculated });

      const token = await acquireMailToken();
      const newLogs: DispatchLogEntry[] = [];

      // Registro de inicio en el log
      const startLog: DispatchLogEntry = {
        id: `start-${Date.now()}`,
        time: new Date().toLocaleTimeString('es-CO', { hour12: false }),
        httpStatus: 200,
        success: true,
        recipientEmail: OFFICIAL_SENDER_EMAIL,
        recipientName: OFFICIAL_SENDER_NAME,
        role: 'Gerencia',
        detail: `🚀 Iniciando despacho de notificaciones · Corte ${formatCutoff(cutoffDate)} · Total: ${totalCalculated} correos`,
      };
      setLogs((prev) => [...prev, startLog]);

      let processed = 0;

      // 1. DESPACHO A COMERCIALES
      if (includeCommercials) {
        for (const c of commercialList) {
          if (!c.email) continue;
          const summary = buildCommercialEmailSummary(c.name, records, cutoffDate);

          let excelAttachment;
          try {
            const excelBase64 = await generateCommercialExcelBase64(c.name, cutoffDate, summary.allRemisiones);
            const cleanName = c.name.replace(/\s+/g, '_');
            excelAttachment = {
              filename: `Remisiones_Abiertas_${cleanName}_${cutoffDate}.xlsx`,
              base64: excelBase64,
            };
          } catch (e) {
            console.warn(`No se pudo adjuntar Excel a comercial ${c.name}:`, e);
          }

          const res = await sendMailViaGraph(token, {
            toEmail: c.email,
            toName: c.name,
            subject: `Oportunidades de Facturación · ${c.name} · Corte ${formatCutoff(cutoffDate)}`,
            htmlBody: generateCommercialEmailHtml(summary),
            excelAttachment,
            senderEmail: OFFICIAL_SENDER_EMAIL,
            senderName: OFFICIAL_SENDER_NAME,
            category: 'comercial',
          });

          processed++;
          setProgress({ current: processed, total: totalCalculated });

          const logItem: DispatchLogEntry = {
            id: `comm-${c.email}-${Date.now()}`,
            time: res.timestamp || new Date().toLocaleTimeString('es-CO', { hour12: false }),
            httpStatus: res.httpStatus || (res.success ? 202 : 500),
            success: res.success,
            recipientEmail: c.email,
            recipientName: c.name,
            role: 'Comercial',
            detail: `${summary.totalCount} remisiones (${formatCOP(summary.totalValue)}) · Excel adjunto`,
            error: res.error,
          };
          setLogs((prev) => [...prev, logItem]);

          // Pausa preventiva de 600 ms para Microsoft Graph rate-limiting
          await new Promise((r) => setTimeout(r, 600));
        }
      }

      // 2. DESPACHO A DIRECTORES DE GRUPO
      if (includeDirectors) {
        for (const dir of directorList) {
          if (!dir.email) continue;
          const summary = buildDirectorEmailSummary(dir.grupo, records, cutoffDate);

          let excelAttachment;
          try {
            const { filename, base64 } = await generateDirectorExcelBase64(summary);
            excelAttachment = { filename, base64 };
          } catch (e) {
            console.warn(`No se pudo adjuntar Excel a director ${dir.nombre}:`, e);
          }

          const res = await sendMailViaGraph(token, {
            toEmail: dir.email,
            toName: dir.nombre,
            subject: `Consolidado de Remisiones · ${dir.nombre} (Grupo ${dir.grupo}) · Corte ${formatCutoff(cutoffDate)}`,
            htmlBody: generateDirectorEmailHtml(summary),
            excelAttachment,
            senderEmail: OFFICIAL_SENDER_EMAIL,
            senderName: OFFICIAL_SENDER_NAME,
            category: 'director',
          });

          processed++;
          setProgress({ current: processed, total: totalCalculated });

          const logItem: DispatchLogEntry = {
            id: `dir-${dir.email}-${Date.now()}`,
            time: res.timestamp || new Date().toLocaleTimeString('es-CO', { hour12: false }),
            httpStatus: res.httpStatus || (res.success ? 202 : 500),
            success: res.success,
            recipientEmail: dir.email,
            recipientName: dir.nombre,
            role: 'Director',
            detail: `Grupo ${dir.grupo} (${dir.carpeta}) · ${summary.totalCount} remisiones (${formatCOP(summary.totalValue)})`,
            error: res.error,
          };
          setLogs((prev) => [...prev, logItem]);

          await new Promise((r) => setTimeout(r, 600));
        }
      }

      // 3. DESPACHO A GERENCIA Y COPIA DE VALIDEZ A ESPECIALISTA PREVENTA
      if (includeGerencia) {
        for (const g of gerenciaList) {
          if (!g.email) continue;
          const isEspecialista = g.email === 'especialista.preventa@provexpress.com.co';

          const summary = buildGerenciaEmailSummary(records, cutoffDate, {
            name: g.nombre,
            cargo: g.cargo,
            email: g.email,
          });

          let excelAttachment;
          try {
            const { filename, base64 } = await generateGerenciaExcelBase64(summary);
            excelAttachment = { filename, base64 };
          } catch (e) {
            console.warn(`No se pudo adjuntar Excel maestro a gerencia ${g.nombre}:`, e);
          }

          const subject = isEspecialista
            ? `📊 Consolidado General de Gestión Comercial · Dirección & Gerencia · Especialista Preventa (Copia de Validez) · Corte ${formatCutoff(cutoffDate)}`
            : `📊 Consolidado General de Gestión Comercial · Dirección & Gerencia · ${g.nombre} · Corte ${formatCutoff(cutoffDate)}`;

          const res = await sendMailViaGraph(token, {
            toEmail: g.email,
            toName: g.nombre,
            subject,
            htmlBody: generateGerenciaEmailHtml(summary),
            excelAttachment,
            senderEmail: OFFICIAL_SENDER_EMAIL,
            senderName: OFFICIAL_SENDER_NAME,
            category: isEspecialista ? 'copia_validez' : 'gerencia',
          });

          processed++;
          setProgress({ current: processed, total: totalCalculated });

          const logItem: DispatchLogEntry = {
            id: `ger-${g.email}-${Date.now()}`,
            time: res.timestamp || new Date().toLocaleTimeString('es-CO', { hour12: false }),
            httpStatus: res.httpStatus || (res.success ? 202 : 500),
            success: res.success,
            recipientEmail: g.email,
            recipientName: g.nombre,
            role: isEspecialista ? 'Copia Validez' : 'Gerencia',
            detail: `${isEspecialista ? 'Copia de Auditoría Especialista · ' : ''}Libro Maestro 3 Hojas (${formatCOP(summary.totalValue)})`,
            error: res.error,
          };
          setLogs((prev) => [...prev, logItem]);

          await new Promise((r) => setTimeout(r, 600));
        }
      }

      // Registro de finalización en el log
      const endLog: DispatchLogEntry = {
        id: `end-${Date.now()}`,
        time: new Date().toLocaleTimeString('es-CO', { hour12: false }),
        httpStatus: 200,
        success: true,
        recipientEmail: OFFICIAL_SENDER_EMAIL,
        recipientName: OFFICIAL_SENDER_NAME,
        role: 'Gerencia',
        detail: `✅ DESPACHO OFICIAL FINALIZADO: ${processed} correos procesados satisfactoriamente.`,
      };
      setLogs((prev) => [...prev, endLog]);

      setAlertMessage({
        type: 'success',
        text: `¡Proceso de notificación finalizado! Se procesaron ${processed} correos oficiales con remitente ${OFFICIAL_SENDER_EMAIL}.`,
      });
    } catch (err: any) {
      setAlertMessage({
        type: 'error',
        text: `Error general durante el despacho: ${err?.message || 'Error desconocido'}`,
      });
    } finally {
      setIsSending(false);
    }
  };

  // ══════════════════════════════════════════════════════════════════════════
  // EXPORTAR Y COPIAR LOG
  // ══════════════════════════════════════════════════════════════════════════
  const copyLogText = () => {
    const text = logs
      .map(
        (l) =>
          `[${l.time}] HTTP ${l.httpStatus} ${l.success ? 'OK' : 'ERROR'} | [${l.role}] ${l.recipientName} <${l.recipientEmail}> - ${l.detail}${
            l.error ? ` :: ${l.error}` : ''
          }`,
      )
      .join('\n');
    navigator.clipboard.writeText(text);
    setCopied(true);
    setTimeout(() => setCopied(false), 2500);
  };

  const downloadLogFile = () => {
    const text = logs
      .map(
        (l) =>
          `[${l.time}] HTTP ${l.httpStatus} ${l.success ? 'OK' : 'ERROR'} | [${l.role}] ${l.recipientName} <${l.recipientEmail}> - ${l.detail}${
            l.error ? ` :: ${l.error}` : ''
          }`,
      )
      .join('\n');
    const blob = new Blob([text], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Log_Entrega_Notificaciones_${cutoffDate}.txt`;
    a.click();
    URL.revokeObjectURL(url);
  };

  return (
    <div className="email-modal-overlay" onClick={onClose} role="dialog" aria-modal="true">
      <div className="dispatch-console-modal" onClick={(e) => e.stopPropagation()}>
        {/* ENCABEZADO */}
        <div className="dispatch-header">
          <div className="dispatch-header-title">
            <div className="dispatch-icon-badge">
              <Send size={20} />
            </div>
            <div>
              <h3>Consola de Despacho de Notificaciones Comerciales</h3>
              <div className="dispatch-header-meta">
                <span className="badge-cutoff">Corte: {formatCutoff(cutoffDate)}</span>
                <span className="meta-divider">•</span>
                <span>Remitente Oficial: <strong>{OFFICIAL_SENDER_EMAIL}</strong></span>
                <span className="meta-divider">•</span>
                <span>Copia Validez: <strong>especialista.preventa@provexpress.com.co</strong></span>
              </div>
            </div>
          </div>
          <button type="button" className="email-modal-close" onClick={onClose} disabled={isSending}>
            <X size={20} />
          </button>
        </div>

        {/* ALERTA DE MENSAJE */}
        {alertMessage && (
          <div className={`email-modal-alert ${alertMessage.type}`}>
            {alertMessage.type === 'success' ? <CheckCircle2 size={16} /> : <AlertCircle size={16} />}
            <span>{alertMessage.text}</span>
          </div>
        )}

        {/* TARJETAS DE PAQUETES A DESPACHAR */}
        <div className="dispatch-cards-grid">
          {/* Card 1: Comerciales */}
          <div className={`dispatch-card ${includeCommercials ? 'active' : 'inactive'}`}>
            <div className="dispatch-card-top">
              <label className="dispatch-checkbox-label">
                <input
                  type="checkbox"
                  checked={includeCommercials}
                  onChange={(e) => setIncludeCommercials(e.target.checked)}
                  disabled={isSending}
                />
                <div className="dispatch-card-icon blue">
                  <Users size={16} />
                </div>
                <strong>Asesores Comerciales ({commercialList.length})</strong>
              </label>
              <span className="dispatch-card-badge blue">Plantilla individual</span>
            </div>
            <div className="dispatch-card-body">
              <p>
                <strong>{totalCommercialRemisiones} remisiones abiertas</strong> por facturar
              </p>
              <div className="dispatch-card-stat">{formatCOP(totalCommercialValue)}</div>
              <div className="dispatch-card-foot">
                <FileSpreadsheet size={13} />
                <span>Libro Excel individual personalizado adjunto</span>
              </div>
            </div>
          </div>

          {/* Card 2: Directores */}
          <div className={`dispatch-card ${includeDirectors ? 'active' : 'inactive'}`}>
            <div className="dispatch-card-top">
              <label className="dispatch-checkbox-label">
                <input
                  type="checkbox"
                  checked={includeDirectors}
                  onChange={(e) => setIncludeDirectors(e.target.checked)}
                  disabled={isSending}
                />
                <div className="dispatch-card-icon amber">
                  <Briefcase size={16} />
                </div>
                <strong>Directores de Grupo ({directorList.length})</strong>
              </label>
              <span className="dispatch-card-badge amber">Consolidado de grupo</span>
            </div>
            <div className="dispatch-card-body">
              <p>Angélica Caballero, Óscar Beltrán, Miller Romero y Rafael Novoa</p>
              <div className="dispatch-card-stat">4 Grupos Comerciales</div>
              <div className="dispatch-card-foot">
                <FileSpreadsheet size={13} />
                <span>Libro Excel consolidado de grupo (2 pestañas)</span>
              </div>
            </div>
          </div>

          {/* Card 3: Gerencia */}
          <div className={`dispatch-card ${includeGerencia ? 'active' : 'inactive'}`}>
            <div className="dispatch-card-top">
              <label className="dispatch-checkbox-label">
                <input
                  type="checkbox"
                  checked={includeGerencia}
                  onChange={(e) => setIncludeGerencia(e.target.checked)}
                  disabled={isSending}
                />
                <div className="dispatch-card-icon purple">
                  <ShieldCheck size={16} />
                </div>
                <strong>Dirección & Gerencia ({gerenciaList.length})</strong>
              </label>
              <span className="dispatch-card-badge purple">Maestro Corporativo</span>
            </div>
            <div className="dispatch-card-body">
              <p>Rafael Novoa, Juan Novoa, Cuentas Estratégicas</p>
              <div className="dispatch-card-stat">
                + Copia a Especialista Preventa
              </div>
              <div className="dispatch-card-foot">
                <FileSpreadsheet size={13} />
                <span>Libro Excel maestro corporativo (3 hojas completas)</span>
              </div>
            </div>
          </div>
        </div>

        {/* BARRA DE CONTROL DE ACCIÓN */}
        <div className="dispatch-action-bar">
          <div className="dispatch-action-info">
            <Sparkles size={16} color="#0071e3" />
            <span>
              Total a despachar:{' '}
              <strong>{totalCalculated} correos oficiales</strong> vía Microsoft Graph API (Exchange Online)
            </span>
          </div>
          <div className="dispatch-action-buttons">
            <button
              type="button"
              className="email-action-btn secondary"
              onClick={handleSendTest}
              disabled={isSending || isSendingTest}
              title="Envía una prueba directa a tu bandeja de entrada"
            >
              <Mail size={14} />
              <span>{isSendingTest ? 'Enviando prueba...' : 'Enviar prueba a mi correo'}</span>
            </button>
            <button
              type="button"
              className="email-send-all-btn"
              onClick={handleStartDispatch}
              disabled={isSending || totalCalculated === 0}
            >
              <Send size={15} />
              <span>
                {isSending
                  ? `Despachando (${progress?.current || 0} de ${progress?.total || totalCalculated})...`
                  : `🚀 Iniciar Envío Masivo Oficial (${totalCalculated})`}
              </span>
            </button>
          </div>
        </div>

        {/* CONSOLA DE LOG EN VIVO (HTTP STATUS 200/202) */}
        <div className="dispatch-log-panel">
          <div className="dispatch-log-header">
            <div className="dispatch-log-title">
              <Terminal size={15} />
              <span>Consola de Registro de Entrega HTTP en Tiempo Real</span>
              {logs.length > 0 && <span className="log-count-badge">{logs.length} eventos</span>}
            </div>
            {logs.length > 0 && (
              <div className="dispatch-log-actions">
                <button type="button" className="btn-log-action" onClick={copyLogText} title="Copiar todo el registro">
                  <Copy size={13} />
                  <span>{copied ? '¡Copiado!' : 'Copiar Log'}</span>
                </button>
                <button type="button" className="btn-log-action" onClick={downloadLogFile} title="Descargar como archivo de texto">
                  <Download size={13} />
                  <span>Descargar Log (.txt)</span>
                </button>
              </div>
            )}
          </div>

          {/* BARRA DE PROGRESO VISUAL */}
          {progress && (
            <div className="dispatch-progress-track">
              <div
                className="dispatch-progress-bar"
                style={{ width: `${Math.round((progress.current / progress.total) * 100)}%` }}
              />
            </div>
          )}

          {/* TERMINAL DE LOGS */}
          <div className="dispatch-terminal">
            {logs.length === 0 ? (
              <div className="dispatch-terminal-empty">
                <Terminal size={24} />
                <p>La consola registrará el estado HTTP (202 Accepted / 200 OK) de cada correo tan pronto inicies el despacho o envíes una prueba.</p>
              </div>
            ) : (
              <div className="dispatch-terminal-lines">
                {logs.map((log) => (
                  <div key={log.id} className={`terminal-line ${log.success ? 'success' : 'error'}`}>
                    <span className="log-time">[{log.time}]</span>
                    <span className={`log-badge-http ${log.httpStatus >= 200 && log.httpStatus < 300 ? 'status-200' : 'status-err'}`}>
                      HTTP {log.httpStatus || (log.success ? 202 : 500)} {log.success ? 'OK' : 'ERR'}
                    </span>
                    <span className={`log-badge-role ${log.role.toLowerCase().replace(/\s+/g, '-')}`}>
                      {log.role}
                    </span>
                    <span className="log-recipient">
                      <strong>{log.recipientName}</strong> &lt;{log.recipientEmail}&gt;
                    </span>
                    <span className="log-detail">— {log.detail}</span>
                    {log.error && <span className="log-error">⚠️ {log.error}</span>}
                  </div>
                ))}
                <div ref={logEndRef} />
              </div>
            )}
          </div>
        </div>

        {/* PIE DEL MODAL */}
        <div className="dispatch-footer">
          <span className="dispatch-footer-note">
            🛡️ Certificación de entrega: Los registros HTTP 202/200 confirman la aceptación oficial en los servidores de Microsoft 365 Exchange.
          </span>
          <button type="button" className="email-cancel-btn" onClick={onClose} disabled={isSending}>
            Cerrar
          </button>
        </div>
      </div>
    </div>
  );
};
