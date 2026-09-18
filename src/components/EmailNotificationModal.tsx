import React, { useMemo, useState } from 'react';
import {
  CheckCircle2,
  Copy,
  Download,
  ExternalLink,
  Eye,
  Mail,
  RefreshCw,
  Send,
  UserCheck,
  X,
  AlertCircle,
  Clock,
  Sparkles,
} from 'lucide-react';
import type { Remision } from '../types';
import { acquireMailToken } from '../lib/auth';
import {
  buildCommercialEmailSummary,
  formatCOP,
  formatNumber,
  generateCommercialEmailHtml,
} from '../lib/commercialEmailTemplate';
import {
  downloadCommercialExcel,
  generateCommercialExcelBase64,
} from '../lib/commercialExcelGenerator';
import { sendBatchEmails, sendMailViaGraph, type SendEmailPayload, type SendResult } from '../lib/emailSender';

interface Props {
  isOpen: boolean;
  onClose: () => void;
  records: Remision[];
  cutoffDate: string;
  currentUserEmail?: string;
}

interface CommercialTarget {
  name: string;
  email: string;
  directorName: string;
  count: number;
  value: number;
  avgAge: number;
  selected: boolean;
  status: 'idle' | 'sending' | 'success' | 'error';
  errorMessage?: string;
}

export const EmailNotificationModal: React.FC<Props> = ({
  isOpen,
  onClose,
  records,
  cutoffDate,
  currentUserEmail,
}) => {
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedCommercialName, setSelectedCommercialName] = useState<string>('');
  const [commercialTargets, setCommercialTargets] = useState<CommercialTarget[]>([]);
  const [isSendingBatch, setIsSendingBatch] = useState(false);
  const [isSendingTest, setIsSendingTest] = useState(false);
  const [batchProgress, setBatchProgress] = useState<{ current: number; total: number } | null>(null);
  const [alertMessage, setAlertMessage] = useState<{ type: 'success' | 'error' | 'info'; text: string } | null>(null);
  const [copied, setCopied] = useState(false);
  const [viewMode, setViewMode] = useState<'preview' | 'html'>('preview');

  // Inicializar lista de comerciales al cambiar el corte o los registros
  useMemo(() => {
    const currentRecords = records.filter((r) => r.cutoff === cutoffDate);
    const byEmployee = new Map<string, Remision[]>();

    currentRecords.forEach((r) => {
      const emp = r.employee?.trim();
      if (!emp) return;
      const list = byEmployee.get(emp) || [];
      list.push(r);
      byEmployee.set(emp, list);
    });

    const targets: CommercialTarget[] = [];
    byEmployee.forEach((empRecords, empName) => {
      const summary = buildCommercialEmailSummary(empName, records, cutoffDate);
      targets.push({
        name: empName,
        email: summary.commercialEmail,
        directorName: summary.directorName,
        count: summary.totalCount,
        value: summary.totalValue,
        avgAge: summary.avgAge,
        selected: Boolean(summary.commercialEmail), // Preseleccionar si tiene correo
        status: 'idle',
      });
    });

    // Ordenar por mayor valor pendiente
    targets.sort((a, b) => b.value - a.value);
    setCommercialTargets(targets);

    if (targets.length > 0 && (!selectedCommercialName || !targets.some((t) => t.name === selectedCommercialName))) {
      setSelectedCommercialName(targets[0].name);
    }
  }, [records, cutoffDate]);

  const activeCommercialSummary = useMemo(() => {
    if (!selectedCommercialName) return null;
    return buildCommercialEmailSummary(selectedCommercialName, records, cutoffDate);
  }, [selectedCommercialName, records, cutoffDate]);

  const activeHtml = useMemo(() => {
    if (!activeCommercialSummary) return '';
    return generateCommercialEmailHtml(activeCommercialSummary, { forWebPreview: true });
  }, [activeCommercialSummary]);

  if (!isOpen) return null;

  const filteredTargets = commercialTargets.filter((t) =>
    t.name.toLowerCase().includes(searchTerm.toLowerCase()) ||
    t.email.toLowerCase().includes(searchTerm.toLowerCase()) ||
    t.directorName.toLowerCase().includes(searchTerm.toLowerCase()),
  );

  const selectedCount = commercialTargets.filter((t) => t.selected).length;
  const totalValueSelected = commercialTargets
    .filter((t) => t.selected)
    .reduce((sum, t) => sum + t.value, 0);

  const toggleSelectAll = () => {
    const allSelected = filteredTargets.every((t) => t.selected);
    setCommercialTargets((prev) =>
      prev.map((t) => {
        if (filteredTargets.some((ft) => ft.name === t.name)) {
          return { ...t, selected: !allSelected };
        }
        return t;
      }),
    );
  };

  const toggleTarget = (name: string) => {
    setCommercialTargets((prev) =>
      prev.map((t) => (t.name === name ? { ...t, selected: !t.selected } : t)),
    );
  };

  const handleCopyHtml = async () => {
    if (!activeHtml) return;
    try {
      await navigator.clipboard.writeText(activeHtml);
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    } catch {
      // Fallback
    }
  };

  const handleDownloadExcel = async () => {
    if (!activeCommercialSummary) return;
    try {
      await downloadCommercialExcel(
        activeCommercialSummary.commercialName,
        activeCommercialSummary.cutoffDate,
        activeCommercialSummary.allRemisiones,
      );
      setAlertMessage({
        type: 'success',
        text: `¡Archivo Excel de ${activeCommercialSummary.commercialName} descargado exitosamente!`,
      });
    } catch (err: any) {
      setAlertMessage({
        type: 'error',
        text: `Error al descargar Excel: ${err?.message || 'Error desconocido'}`,
      });
    }
  };

  const handleSendTest = async () => {
    if (!activeCommercialSummary) return;
    const testRecipient = currentUserEmail || activeCommercialSummary.commercialEmail;
    if (!testRecipient) {
      setAlertMessage({ type: 'error', text: 'No se detectó un correo de destino para la prueba.' });
      return;
    }

    try {
      setIsSendingTest(true);
      setAlertMessage(null);
      const token = await acquireMailToken();
      const subject = `[PRUEBA] Oportunidades de Facturación · ${activeCommercialSummary.commercialName} · ${cutoffDate}`;

      let excelAttachment;
      try {
        const excelBase64 = await generateCommercialExcelBase64(
          activeCommercialSummary.commercialName,
          activeCommercialSummary.cutoffDate,
          activeCommercialSummary.allRemisiones,
        );
        const cleanName = activeCommercialSummary.commercialName.replace(/\s+/g, '_');
        excelAttachment = {
          filename: `Remisiones_Abiertas_${cleanName}_${cutoffDate}.xlsx`,
          base64: excelBase64,
        };
      } catch (e) {
        console.warn('No se pudo generar adjunto Excel:', e);
      }

      const res = await sendMailViaGraph(token, {
        toEmail: testRecipient,
        toName: activeCommercialSummary.commercialName,
        subject,
        htmlBody: generateCommercialEmailHtml(activeCommercialSummary),
        excelAttachment,
      });

      if (res.success) {
        setAlertMessage({
          type: 'success',
          text: `¡Correo de prueba enviado con éxito a ${testRecipient}! Incluye el archivo Excel adjunto.`,
        });
      } else {
        setAlertMessage({
          type: 'error',
          text: `Error al enviar prueba: ${res.error || 'Verifica permisos de Microsoft 365'}`,
        });
      }
    } catch (err: any) {
      setAlertMessage({
        type: 'error',
        text: `No se pudo conectar a Microsoft 365: ${err?.message || 'Inicia sesión con tu cuenta corporativa'}`,
      });
    } finally {
      setIsSendingTest(false);
    }
  };

  const handleSendBatch = async () => {
    const toSend = commercialTargets.filter((t) => t.selected && t.email);
    if (toSend.length === 0) {
      setAlertMessage({ type: 'error', text: 'Selecciona al menos un comercial con correo válido.' });
      return;
    }

    try {
      setIsSendingBatch(true);
      setAlertMessage(null);
      setBatchProgress({ current: 0, total: toSend.length });

      const token = await acquireMailToken();

      const items: SendEmailPayload[] = await Promise.all(
        toSend.map(async (t) => {
          const sum = buildCommercialEmailSummary(t.name, records, cutoffDate);
          let excelAttachment;
          try {
            const excelBase64 = await generateCommercialExcelBase64(
              sum.commercialName,
              sum.cutoffDate,
              sum.allRemisiones,
            );
            const cleanName = sum.commercialName.replace(/\s+/g, '_');
            excelAttachment = {
              filename: `Remisiones_Abiertas_${cleanName}_${cutoffDate}.xlsx`,
              base64: excelBase64,
            };
          } catch (e) {
            console.warn('No se pudo generar adjunto Excel para', t.name, e);
          }
          return {
            toEmail: t.email,
            toName: t.name,
            subject: `Oportunidades de Facturación · ${t.name} · Corte ${cutoffDate}`,
            htmlBody: generateCommercialEmailHtml(sum),
            excelAttachment,
          };
        }),
      );

      // Actualizar estado a sending
      setCommercialTargets((prev) =>
        prev.map((t) => (t.selected && t.email ? { ...t, status: 'sending', errorMessage: undefined } : t)),
      );

      await sendBatchEmails(
        token,
        items,
        (current, total, lastResult) => {
          setBatchProgress({ current, total });
          setCommercialTargets((prev) =>
            prev.map((t) => {
              if (t.email === lastResult.toEmail) {
                return {
                  ...t,
                  status: lastResult.success ? 'success' : 'error',
                  errorMessage: lastResult.error,
                };
              }
              return t;
            }),
          );
        },
        500, // Pausa de 500ms
      );

      setAlertMessage({
        type: 'success',
        text: `Envío completado: se procesaron las notificaciones para ${toSend.length} comerciales.`,
      });
    } catch (err: any) {
      setAlertMessage({
        type: 'error',
        text: `Error durante el envío masivo: ${err?.message || 'Verifica tu conexión y permisos'}`,
      });
    } finally {
      setIsSendingBatch(false);
      setBatchProgress(null);
    }
  };

  return (
    <div className="email-modal-overlay" onClick={onClose}>
      <div className="email-modal-card" onClick={(e) => e.stopPropagation()}>
        {/* ENCABEZADO MODAL */}
        <div className="email-modal-header">
          <div className="email-modal-header-left">
            <div className="email-modal-icon-badge">
              <Mail size={20} />
            </div>
            <div>
              <div className="email-modal-title">
                Notificación Diaria a Ejecutivos Comerciales
              </div>
              <div className="email-modal-subtitle">
                Corte activo: <strong>{cutoffDate}</strong> · Formato corporativo Provexpress SAS
              </div>
            </div>
          </div>
          <button className="email-modal-close-btn" onClick={onClose} title="Cerrar modal">
            <X size={20} />
          </button>
        </div>

        {/* ALERTA O MENSAJE DE ESTADO */}
        {alertMessage && (
          <div className={`email-modal-alert ${alertMessage.type}`}>
            {alertMessage.type === 'success' ? <CheckCircle2 size={16} /> : <AlertCircle size={16} />}
            <span>{alertMessage.text}</span>
            <button onClick={() => setAlertMessage(null)} className="alert-dismiss-btn">✕</button>
          </div>
        )}

        {/* CUERPO DEL MODAL (2 COLUMNAS) */}
        <div className="email-modal-body">
          {/* PANEL IZQUIERDO: LISTA DE COMERCIALES */}
          <div className="email-modal-sidebar">
            <div className="email-sidebar-header">
              <div className="email-sidebar-title-row">
                <span className="email-sidebar-title">
                  Comerciales con saldo ({commercialTargets.length})
                </span>
                <button
                  type="button"
                  className="email-select-all-btn"
                  onClick={toggleSelectAll}
                >
                  {filteredTargets.every((t) => t.selected) ? 'Deseleccionar' : 'Seleccionar todo'}
                </button>
              </div>

              <input
                type="text"
                className="email-search-input"
                placeholder="Buscar comercial o director..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
              />

              <div className="email-selected-summary">
                Seleccionados: <strong>{selectedCount} de {commercialTargets.length}</strong> · {formatCOP(totalValueSelected)}
              </div>
            </div>

            {/* LISTADO */}
            <div className="email-targets-list">
              {filteredTargets.length === 0 ? (
                <div className="email-empty-search">No se encontraron comerciales</div>
              ) : (
                filteredTargets.map((target) => {
                  const isSelectedForPreview = selectedCommercialName === target.name;
                  return (
                    <div
                      key={target.name}
                      className={`email-target-item ${isSelectedForPreview ? 'active-preview' : ''}`}
                      onClick={() => setSelectedCommercialName(target.name)}
                    >
                      <input
                        type="checkbox"
                        checked={target.selected}
                        disabled={isSendingBatch}
                        onChange={(e) => {
                          e.stopPropagation();
                          toggleTarget(target.name);
                        }}
                        className="email-target-checkbox"
                      />
                      <div className="email-target-info">
                        <div className="email-target-top">
                          <strong className="email-target-name">{target.name}</strong>
                          {target.status === 'success' && (
                            <span className="email-status-pill success"><CheckCircle2 size={11} /> Enviado</span>
                          )}
                          {target.status === 'sending' && (
                            <span className="email-status-pill sending"><RefreshCw size={11} className="spin" /> Enviando</span>
                          )}
                          {target.status === 'error' && (
                            <span className="email-status-pill error" title={target.errorMessage}><AlertCircle size={11} /> Falló</span>
                          )}
                        </div>
                        <div className="email-target-meta">
                          <span>{target.email || 'Sin correo'}</span>
                          <span>·</span>
                          <span className="email-target-director">{target.directorName}</span>
                        </div>
                        <div className="email-target-numbers">
                          <span className="email-target-count">{target.count} rem.</span>
                          <span className="email-target-val">{formatCOP(target.value)}</span>
                          <span className="email-target-age">⏱ {target.avgAge}d</span>
                        </div>
                      </div>
                    </div>
                  );
                })
              )}
            </div>
          </div>

          {/* PANEL DERECHO: PREVISUALIZACIÓN EN VIVO */}
          <div className="email-modal-preview-pane">
            <div className="email-preview-toolbar">
              <div className="email-preview-user-badge">
                <Eye size={15} />
                <span>Vista previa: <strong>{selectedCommercialName || 'Ninguno'}</strong></span>
                {activeCommercialSummary?.commercialEmail && (
                  <small>({activeCommercialSummary.commercialEmail})</small>
                )}
              </div>

              <div className="email-preview-actions">
                <div className="email-view-mode-tabs">
                  <button
                    className={`email-mode-tab ${viewMode === 'preview' ? 'active' : ''}`}
                    onClick={() => setViewMode('preview')}
                  >
                    Visual
                  </button>
                  <button
                    className={`email-mode-tab ${viewMode === 'html' ? 'active' : ''}`}
                    onClick={() => setViewMode('html')}
                  >
                    HTML
                  </button>
                </div>

                <button
                  type="button"
                  className="email-action-btn secondary"
                  onClick={handleCopyHtml}
                  title="Copiar código HTML del correo"
                >
                  <Copy size={13} />
                  <span>{copied ? '¡Copiado!' : 'Copiar HTML'}</span>
                </button>

                <button
                  type="button"
                  className="email-action-btn secondary"
                  onClick={handleDownloadExcel}
                  title="Descargar archivo Excel con todas las remisiones abiertas del comercial"
                >
                  <Download size={13} />
                  <span>Descargar Excel ({activeCommercialSummary?.totalCount || 0})</span>
                </button>

                <button
                  type="button"
                  className="email-action-btn primary"
                  onClick={handleSendTest}
                  disabled={isSendingTest || isSendingBatch}
                  title="Enviar correo de muestra a tu bandeja de entrada"
                >
                  <Send size={13} />
                  <span>{isSendingTest ? 'Enviando...' : 'Enviar prueba a mi correo'}</span>
                </button>
              </div>
            </div>

            {/* CONTENIDO PREVIEW */}
            <div className="email-preview-render-area">
              {viewMode === 'preview' ? (
                <iframe
                  title="Vista Previa Correo"
                  srcDoc={activeHtml}
                  className="email-preview-iframe"
                  sandbox="allow-same-origin"
                />
              ) : (
                <pre className="email-preview-code">
                  <code>{activeHtml}</code>
                </pre>
              )}
            </div>
          </div>
        </div>

        {/* PIE DEL MODAL CON ACCIÓN MASIVA */}
        <div className="email-modal-footer">
          <div className="email-footer-left">
            {batchProgress && (
              <div className="email-batch-progress">
                <RefreshCw size={14} className="spin" />
                <span>Enviando notificaciones: <strong>{batchProgress.current} de {batchProgress.total}</strong>...</span>
                <div className="email-progress-bar-track">
                  <div
                    className="email-progress-bar-fill"
                    style={{ width: `${Math.round((batchProgress.current / batchProgress.total) * 100)}%` }}
                  />
                </div>
              </div>
            )}
            {!batchProgress && (
              <div className="email-footer-info">
                <Sparkles size={14} color="#15803D" />
                <span>Se enviará un correo personalizado a cada comercial con su logo y remisiones.</span>
              </div>
            )}
          </div>

          <div className="email-footer-right">
            <button
              type="button"
              className="email-cancel-btn"
              onClick={onClose}
              disabled={isSendingBatch}
            >
              Cerrar
            </button>
            <button
              type="button"
              className="email-send-all-btn"
              onClick={handleSendBatch}
              disabled={isSendingBatch || selectedCount === 0}
            >
              <Send size={15} />
              <span>
                {isSendingBatch
                  ? 'Enviando...'
                  : `Enviar a los ${selectedCount} comerciales seleccionados`}
              </span>
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
