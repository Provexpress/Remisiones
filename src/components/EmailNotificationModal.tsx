import React, { useMemo, useState } from 'react';
import {
  CheckCircle2,
  Copy,
  Download,
  Eye,
  Mail,
  RefreshCw,
  Send,
  X,
  AlertCircle,
  Sparkles,
  Users,
  Briefcase,
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
import {
  buildDirectorEmailSummary,
  generateDirectorEmailHtml,
  type DirectorEmailSummary,
} from '../lib/directorEmailTemplate';
import {
  downloadDirectorExcel,
  generateDirectorExcelBase64,
} from '../lib/directorExcelGenerator';
import { LISTA_DIRECTORES } from '../lib/commercialDirectory';
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

interface DirectorTarget {
  grupo: number;
  name: string;
  email: string;
  count: number;
  value: number;
  avgAge: number;
  activeExecutivesCount: number;
  totalExecutivesCount: number;
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
  const [recipientType, setRecipientType] = useState<'commercials' | 'directors'>('commercials');
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedCommercialName, setSelectedCommercialName] = useState<string>('');
  const [selectedDirectorGroup, setSelectedDirectorGroup] = useState<number>(2);
  const [commercialTargets, setCommercialTargets] = useState<CommercialTarget[]>([]);
  const [directorTargets, setDirectorTargets] = useState<DirectorTarget[]>([]);
  const [isSendingBatch, setIsSendingBatch] = useState(false);
  const [isSendingTest, setIsSendingTest] = useState(false);
  const [batchProgress, setBatchProgress] = useState<{ current: number; total: number } | null>(null);
  const [alertMessage, setAlertMessage] = useState<{ type: 'success' | 'error' | 'info'; text: string } | null>(null);
  const [copied, setCopied] = useState(false);
  const [viewMode, setViewMode] = useState<'preview' | 'html'>('preview');

  // Inicializar listas al cambiar registros o corte
  useMemo(() => {
    const currentRecords = records.filter((r) => r.cutoff === cutoffDate);

    // 1. Comerciales
    const byEmployee = new Map<string, Remision[]>();
    currentRecords.forEach((r) => {
      const emp = r.employee?.trim();
      if (!emp) return;
      const list = byEmployee.get(emp) || [];
      list.push(r);
      byEmployee.set(emp, list);
    });

    const cTargets: CommercialTarget[] = [];
    byEmployee.forEach((empRecords, empName) => {
      const summary = buildCommercialEmailSummary(empName, records, cutoffDate);
      cTargets.push({
        name: empName,
        email: summary.commercialEmail,
        directorName: summary.directorName,
        count: summary.totalCount,
        value: summary.totalValue,
        avgAge: summary.avgAge,
        selected: Boolean(summary.commercialEmail),
        status: 'idle',
      });
    });

    cTargets.sort((a, b) => b.value - a.value);
    setCommercialTargets(cTargets);

    if (cTargets.length > 0 && (!selectedCommercialName || !cTargets.some((t) => t.name === selectedCommercialName))) {
      setSelectedCommercialName(cTargets[0].name);
    }

    // 2. Directores
    const dTargets: DirectorTarget[] = LISTA_DIRECTORES.map((dir) => {
      const sum = buildDirectorEmailSummary(dir.grupo, records, cutoffDate);
      return {
        grupo: dir.grupo,
        name: dir.nombre,
        email: dir.email,
        count: sum.totalCount,
        value: sum.totalValue,
        avgAge: sum.avgAge,
        activeExecutivesCount: sum.activeExecutivesCount,
        totalExecutivesCount: sum.totalExecutivesCount,
        selected: Boolean(dir.email),
        status: 'idle',
      };
    });

    setDirectorTargets(dTargets);
  }, [records, cutoffDate]);

  // Resumen del comercial seleccionado
  const activeCommercialSummary = useMemo(() => {
    if (recipientType !== 'commercials' || !selectedCommercialName) return null;
    return buildCommercialEmailSummary(selectedCommercialName, records, cutoffDate);
  }, [recipientType, selectedCommercialName, records, cutoffDate]);

  // Resumen del director seleccionado
  const activeDirectorSummary = useMemo(() => {
    if (recipientType !== 'directors' || !selectedDirectorGroup) return null;
    return buildDirectorEmailSummary(selectedDirectorGroup, records, cutoffDate);
  }, [recipientType, selectedDirectorGroup, records, cutoffDate]);

  // HTML activo para visualización
  const activeHtml = useMemo(() => {
    if (recipientType === 'directors') {
      if (!activeDirectorSummary) return '';
      return generateDirectorEmailHtml(activeDirectorSummary, { forWebPreview: true });
    }
    if (!activeCommercialSummary) return '';
    return generateCommercialEmailHtml(activeCommercialSummary, { forWebPreview: true });
  }, [recipientType, activeDirectorSummary, activeCommercialSummary]);

  if (!isOpen) return null;

  // Filtrado de comerciales
  const filteredCommercials = commercialTargets.filter((t) =>
    t.name.toLowerCase().includes(searchTerm.toLowerCase()) ||
    t.email.toLowerCase().includes(searchTerm.toLowerCase()) ||
    t.directorName.toLowerCase().includes(searchTerm.toLowerCase()),
  );

  // Filtrado de directores
  const filteredDirectors = directorTargets.filter((t) =>
    t.name.toLowerCase().includes(searchTerm.toLowerCase()) ||
    t.email.toLowerCase().includes(searchTerm.toLowerCase()) ||
    `grupo ${t.grupo}`.includes(searchTerm.toLowerCase()),
  );

  // Conteos de selección
  const selectedCount = recipientType === 'directors'
    ? directorTargets.filter((t) => t.selected).length
    : commercialTargets.filter((t) => t.selected).length;

  const totalValueSelected = recipientType === 'directors'
    ? directorTargets.filter((t) => t.selected).reduce((sum, t) => sum + t.value, 0)
    : commercialTargets.filter((t) => t.selected).reduce((sum, t) => sum + t.value, 0);

  const toggleSelectAll = () => {
    if (recipientType === 'directors') {
      const allSelected = filteredDirectors.every((t) => t.selected);
      setDirectorTargets((prev) =>
        prev.map((t) => (filteredDirectors.some((fd) => fd.grupo === t.grupo) ? { ...t, selected: !allSelected } : t)),
      );
    } else {
      const allSelected = filteredCommercials.every((t) => t.selected);
      setCommercialTargets((prev) =>
        prev.map((t) => (filteredCommercials.some((ft) => ft.name === t.name) ? { ...t, selected: !allSelected } : t)),
      );
    }
  };

  const toggleTarget = (key: string | number) => {
    if (recipientType === 'directors') {
      setDirectorTargets((prev) =>
        prev.map((t) => (t.grupo === key ? { ...t, selected: !t.selected } : t)),
      );
    } else {
      setCommercialTargets((prev) =>
        prev.map((t) => (t.name === key ? { ...t, selected: !t.selected } : t)),
      );
    }
  };

  const handleCopyHtml = async () => {
    if (!activeHtml) return;
    try {
      await navigator.clipboard.writeText(activeHtml);
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    } catch {
      // Fallback silencioso
    }
  };

  const handleDownloadExcel = async () => {
    try {
      if (recipientType === 'directors') {
        if (!activeDirectorSummary) return;
        await downloadDirectorExcel(activeDirectorSummary);
        setAlertMessage({
          type: 'success',
          text: `¡Excel consolidado de ${activeDirectorSummary.directorName} descargado exitosamente!`,
        });
      } else {
        if (!activeCommercialSummary) return;
        await downloadCommercialExcel(
          activeCommercialSummary.commercialName,
          activeCommercialSummary.cutoffDate,
          activeCommercialSummary.allRemisiones,
        );
        setAlertMessage({
          type: 'success',
          text: `¡Archivo Excel de ${activeCommercialSummary.commercialName} descargado exitosamente!`,
        });
      }
    } catch (err: any) {
      setAlertMessage({
        type: 'error',
        text: `Error al descargar Excel: ${err?.message || 'Error desconocido'}`,
      });
    }
  };

  const handleSendTest = async () => {
    const testRecipient = currentUserEmail || (recipientType === 'directors' ? activeDirectorSummary?.directorEmail : activeCommercialSummary?.commercialEmail);
    if (!testRecipient) {
      setAlertMessage({ type: 'error', text: 'No se detectó un correo de destino para la prueba.' });
      return;
    }

    try {
      setIsSendingTest(true);
      setAlertMessage(null);
      const token = await acquireMailToken();

      if (recipientType === 'directors') {
        if (!activeDirectorSummary) return;
        const subject = `[PRUEBA] Consolidado de Remisiones · ${activeDirectorSummary.directorName} (Grupo ${activeDirectorSummary.directorGroup}) · ${cutoffDate}`;

        let excelAttachment;
        try {
          const { filename, base64 } = await generateDirectorExcelBase64(activeDirectorSummary);
          excelAttachment = { filename, base64 };
        } catch (e) {
          console.warn('No se pudo generar adjunto Excel de director:', e);
        }

        const res = await sendMailViaGraph(token, {
          toEmail: testRecipient,
          toName: activeDirectorSummary.directorName,
          subject,
          htmlBody: generateDirectorEmailHtml(activeDirectorSummary),
          excelAttachment,
        });

        if (res.success) {
          setAlertMessage({
            type: 'success',
            text: `¡Correo de prueba de Director enviado con éxito a ${testRecipient}! Incluye el archivo Excel consolidado adjunto.`,
          });
        } else {
          setAlertMessage({
            type: 'error',
            text: `Error al enviar prueba: ${res.error || 'Verifica permisos de Microsoft 365'}`,
          });
        }
      } else {
        if (!activeCommercialSummary) return;
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
          console.warn('No se pudo generar adjunto Excel comercial:', e);
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
    if (recipientType === 'directors') {
      const toSend = directorTargets.filter((t) => t.selected && t.email);
      if (toSend.length === 0) {
        setAlertMessage({ type: 'error', text: 'Selecciona al menos un Director con correo válido.' });
        return;
      }

      try {
        setIsSendingBatch(true);
        setAlertMessage(null);
        setBatchProgress({ current: 0, total: toSend.length });

        const token = await acquireMailToken();

        const items: SendEmailPayload[] = await Promise.all(
          toSend.map(async (t) => {
            const sum = buildDirectorEmailSummary(t.grupo, records, cutoffDate);
            let excelAttachment;
            try {
              const { filename, base64 } = await generateDirectorExcelBase64(sum);
              excelAttachment = { filename, base64 };
            } catch (e) {
              console.warn('No se pudo generar adjunto Excel para director', t.name, e);
            }
            return {
              toEmail: t.email,
              toName: t.name,
              subject: `📊 Reporte Consolidado de Remisiones · Dirección Grupo ${t.grupo} · ${t.name} · ${cutoffDate}`,
              htmlBody: generateDirectorEmailHtml(sum),
              excelAttachment,
            };
          }),
        );

        setDirectorTargets((prev) =>
          prev.map((t) => (t.selected && t.email ? { ...t, status: 'sending', errorMessage: undefined } : t)),
        );

        await sendBatchEmails(
          token,
          items,
          (sent, total, lastResult) => {
            setBatchProgress({ current: sent, total });
            setDirectorTargets((prev) =>
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
          500,
        );

        setAlertMessage({
          type: 'success',
          text: `Envío completado: se procesaron los reportes para ${toSend.length} Directores de Grupo.`,
        });
      } catch (err: any) {
        setAlertMessage({
          type: 'error',
          text: `Error durante el envío a directores: ${err?.message || 'Verifica tu conexión y permisos'}`,
        });
      } finally {
        setIsSendingBatch(false);
        setBatchProgress(null);
      }
    } else {
      // Envío a comerciales
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

        setCommercialTargets((prev) =>
          prev.map((t) => (t.selected && t.email ? { ...t, status: 'sending', errorMessage: undefined } : t)),
        );

        await sendBatchEmails(
          token,
          items,
          (sent, total, lastResult) => {
            setBatchProgress({ current: sent, total });
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
          500,
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
                Notificaciones Corporativas de Remisiones por Correo
              </div>
              <div className="email-modal-subtitle">
                Corte activo: <strong>{cutoffDate}</strong> · Formato oficial Provexpress SAS
              </div>
            </div>
          </div>
          <button className="email-modal-close-btn" onClick={onClose} title="Cerrar modal">
            <X size={20} />
          </button>
        </div>

        {/* SELECTOR DE PESTAÑAS (COMERCIALES / DIRECTORES) */}
        <div className="email-modal-tabs">
          <button
            type="button"
            className={`email-tab-btn ${recipientType === 'commercials' ? 'active' : ''}`}
            onClick={() => setRecipientType('commercials')}
          >
            <Users size={14} />
            <span>👤 Ejecutivos Comerciales ({commercialTargets.length})</span>
          </button>
          <button
            type="button"
            className={`email-tab-btn ${recipientType === 'directors' ? 'active' : ''}`}
            onClick={() => setRecipientType('directors')}
          >
            <Briefcase size={14} />
            <span>👔 Directores de Grupo ({directorTargets.length})</span>
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
          {/* PANEL IZQUIERDO: LISTA DE DESTINATARIOS */}
          <div className="email-modal-sidebar">
            <div className="email-sidebar-header">
              <div className="email-sidebar-title-row">
                <span className="email-sidebar-title">
                  {recipientType === 'directors'
                    ? `Directores de Grupo (${directorTargets.length})`
                    : `Comerciales con saldo (${commercialTargets.length})`}
                </span>
                <button
                  type="button"
                  className="email-select-all-btn"
                  onClick={toggleSelectAll}
                >
                  {(recipientType === 'directors' ? filteredDirectors : filteredCommercials).every((t) => t.selected)
                    ? 'Deseleccionar'
                    : 'Seleccionar todo'}
                </button>
              </div>

              <input
                type="text"
                className="email-search-input"
                placeholder={recipientType === 'directors' ? 'Buscar director o grupo...' : 'Buscar comercial o director...'}
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
              />

              <div className="email-selected-summary">
                Seleccionados: <strong>{selectedCount} de {recipientType === 'directors' ? directorTargets.length : commercialTargets.length}</strong> · {formatCOP(totalValueSelected)}
              </div>
            </div>

            {/* LISTADO SEGÚN PESTAÑA ACTIVA */}
            <div className="email-targets-list">
              {recipientType === 'directors' ? (
                filteredDirectors.length === 0 ? (
                  <div className="email-empty-search">No se encontraron directores</div>
                ) : (
                  filteredDirectors.map((target) => {
                    const isSelectedForPreview = selectedDirectorGroup === target.grupo;
                    return (
                      <div
                        key={target.grupo}
                        className={`email-target-item ${isSelectedForPreview ? 'active-preview' : ''}`}
                        onClick={() => setSelectedDirectorGroup(target.grupo)}
                      >
                        <input
                          type="checkbox"
                          checked={target.selected}
                          disabled={isSendingBatch}
                          onChange={(e) => {
                            e.stopPropagation();
                            toggleTarget(target.grupo);
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
                            <span style={{ fontWeight: 700, color: '#1E3A8A' }}>Grupo {target.grupo}</span>
                            <span>·</span>
                            <span>{target.email}</span>
                          </div>
                          <div className="email-target-numbers">
                            <span className="email-target-count">{target.count} remisiones</span>
                            <span className="email-target-val">{formatCOP(target.value)}</span>
                            <span className="email-target-age">👥 {target.activeExecutivesCount}/{target.totalExecutivesCount} ejec.</span>
                          </div>
                        </div>
                      </div>
                    );
                  })
                )
              ) : (
                filteredCommercials.length === 0 ? (
                  <div className="email-empty-search">No se encontraron comerciales</div>
                ) : (
                  filteredCommercials.map((target) => {
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
                )
              )}
            </div>
          </div>

          {/* PANEL DERECHO: PREVISUALIZACIÓN EN VIVO */}
          <div className="email-modal-preview-pane">
            <div className="email-preview-toolbar">
              <div className="email-preview-user-badge">
                <Eye size={15} />
                <span>
                  Vista previa:{' '}
                  <strong>
                    {recipientType === 'directors'
                      ? `${activeDirectorSummary?.directorName || 'Ninguno'} (Grupo ${activeDirectorSummary?.directorGroup})`
                      : selectedCommercialName || 'Ninguno'}
                  </strong>
                </span>
                <small>
                  (
                  {recipientType === 'directors'
                    ? activeDirectorSummary?.directorEmail
                    : activeCommercialSummary?.commercialEmail}
                  )
                </small>
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
                  title="Descargar archivo Excel adjunto"
                >
                  <Download size={13} />
                  <span>
                    {recipientType === 'directors'
                      ? `Descargar Excel Consolidado (${activeDirectorSummary?.totalCount || 0})`
                      : `Descargar Excel (${activeCommercialSummary?.totalCount || 0})`}
                  </span>
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
                <span>Enviando reportes: <strong>{batchProgress.current} de {batchProgress.total}</strong>...</span>
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
                <span>
                  {recipientType === 'directors'
                    ? 'Se enviará el reporte consolidado con archivo Excel de 2 pestañas a cada Director seleccionado.'
                    : 'Se enviará un correo personalizado a cada comercial con su logo y remisiones.'}
                </span>
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
                  : recipientType === 'directors'
                  ? `Enviar a los ${selectedCount} Directores seleccionados`
                  : `Enviar a los ${selectedCount} comerciales seleccionados`}
              </span>
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
