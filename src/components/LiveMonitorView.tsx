import React, { useState, useMemo } from 'react';
import {
  Zap,
  RefreshCw,
  ArrowDownLeft,
  ArrowUpRight,
  Package,
  Search,
  X,
  Building2,
  Download,
  CheckCircle2,
} from 'lucide-react';
import type { Remision } from '../types';
import type { LiveProduct } from '../lib/remisionesApi';
import { resolveCommercialOrDirector } from '../lib/commercialDirectory';

const DIR_BY_GRUPO: Record<number, string> = {
  1: 'Rafael Novoa',
  2: 'Angélica Caballero',
  3: 'Óscar Beltrán',
  4: 'Miller Romero',
};

export function getRecordDirector(r: { director?: string; employee?: string; group?: number | null }): string {
  if (r.director && r.director !== 'Sin asignar' && String(r.director).trim() !== '') {
    return r.director;
  }
  const role = resolveCommercialOrDirector(r.employee || '');
  if (role.directorNombre && role.directorNombre !== 'Otras Áreas / Especiales' && role.directorNombre !== 'Sin asignar') {
    return role.directorNombre;
  }
  if (r.group && DIR_BY_GRUPO[r.group]) {
    return DIR_BY_GRUPO[r.group];
  }
  return 'Sin asignar';
}

interface LiveMonitorViewProps {
  records: Remision[];
  entrantes: Remision[];
  salientes: Remision[];
  lastSyncTime: string;
  isLoading: boolean;
  onSync: () => Promise<void>;
  productsMap?: Map<number, LiveProduct[]>;
  onSelectDirector?: (director: string) => void;
  saldoAnterior?: number;
  saldoAnteriorCount?: number;
  corteAnteriorFecha?: string;
}

const currencyFmt = new Intl.NumberFormat('es-CO', {
  style: 'currency',
  currency: 'COP',
  maximumFractionDigits: 0,
});

export const LiveMonitorView: React.FC<LiveMonitorViewProps> = ({
  records,
  entrantes,
  salientes,
  lastSyncTime,
  isLoading,
  onSync,
  productsMap,
}) => {
  const [activeTab, setActiveTab] = useState<'salientes' | 'entrantes' | 'vivas'>('salientes');
  const [selectedDirector, setSelectedDirector] = useState<string>('Todos');
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedProductRemision, setSelectedProductRemision] = useState<Remision | null>(null);
  const [isExportingExcel, setIsExportingExcel] = useState(false);

  // Totales
  const totalSalientesValor = useMemo(() => salientes.reduce((s, r) => s + (r.total || 0), 0), [salientes]);
  const totalEntrantesValor = useMemo(() => entrantes.reduce((s, r) => s + (r.total || 0), 0), [entrantes]);
  const totalVivasValor = useMemo(() => records.reduce((s, r) => s + (r.total || 0), 0), [records]);

  // Lista de directores para filtrar
  const directorsList = useMemo(() => {
    const set = new Set<string>();
    [...salientes, ...entrantes, ...records].forEach((r) => {
      const d = getRecordDirector(r);
      if (d && d !== 'Sin asignar') set.add(d);
    });
    return ['Todos', ...Array.from(set).sort()];
  }, [salientes, entrantes, records]);

  // Filtrado de la lista activa
  const currentList = useMemo(() => {
    let baseList = activeTab === 'salientes' ? salientes : activeTab === 'entrantes' ? entrantes : records;

    if (selectedDirector !== 'Todos') {
      baseList = baseList.filter((r) => getRecordDirector(r) === selectedDirector);
    }

    if (!searchTerm.trim()) return baseList;
    const q = searchTerm.toLowerCase().trim();
    return baseList.filter(
      (r) =>
        r.company?.toLowerCase().includes(q) ||
        r.nit?.toLowerCase().includes(q) ||
        r.document?.toLowerCase().includes(q) ||
        r.order?.toLowerCase().includes(q) ||
        r.employee?.toLowerCase().includes(q) ||
        getRecordDirector(r).toLowerCase().includes(q),
    );
  }, [activeTab, salientes, entrantes, records, selectedDirector, searchTerm]);

  const activeProducts = useMemo(() => {
    if (!selectedProductRemision || !productsMap) return [];
    const num = parseInt(selectedProductRemision.document.replace(/\D/g, ''), 10);
    return productsMap.get(num) || [];
  }, [selectedProductRemision, productsMap]);

  // Exportar listado a Excel
  const handleExportExcel = async () => {
    try {
      setIsExportingExcel(true);
      const { default: ExcelJSRuntime } = await import('exceljs');
      const workbook = new ExcelJSRuntime.Workbook();
      workbook.creator = 'Provexpress';
      workbook.created = new Date();

      const sheetName =
        activeTab === 'salientes'
          ? 'Lo que Salió (Facturadas)'
          : activeTab === 'entrantes'
          ? 'Lo que Entró (Nuevas)'
          : 'Total en ERP';

      const headerColor =
        activeTab === 'salientes' ? 'FF166534' : activeTab === 'entrantes' ? 'FF1D4ED8' : 'FF0F172A';

      const worksheet = workbook.addWorksheet(sheetName);
      worksheet.columns = [
        { header: '#', key: 'rank', width: 6 },
        { header: 'No. Remisión', key: 'document', width: 16 },
        { header: 'No. Pedido', key: 'order', width: 14 },
        { header: 'Cliente / Empresa', key: 'company', width: 34 },
        { header: 'NIT', key: 'nit', width: 16 },
        { header: 'Comercial Asesor', key: 'employee', width: 28 },
        { header: 'Director Comercial', key: 'director', width: 26 },
        { header: 'Vr. Total', key: 'total', width: 18 },
        { header: 'Fecha Emisión', key: 'issuedAt', width: 16 },
        { header: 'Días', key: 'age', width: 12 },
        { header: 'Movimiento', key: 'tipo', width: 22 },
      ];

      const headerRow = worksheet.getRow(1);
      headerRow.height = 28;
      headerRow.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
      headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: headerColor } };
      headerRow.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };

      currentList.forEach((r, idx) => {
        const row = worksheet.addRow({
          rank: idx + 1,
          document: r.document,
          order: r.order || '',
          company: r.company,
          nit: r.nit,
          employee: r.employee,
          director: getRecordDirector(r),
          total: r.total || 0,
          issuedAt: r.issuedAt,
          age: r.age,
          tipo:
            activeTab === 'salientes'
              ? 'SALIÓ (FACTURADA)'
              : activeTab === 'entrantes'
              ? 'ENTRÓ (NUEVA)'
              : 'EN ERP',
        });
        row.height = 20;
        row.font = { name: 'Segoe UI', size: 9.5 };
        row.getCell('total').numFmt = '"$"#,##0;[Red]-"$"#,##0;"$0"';
        row.getCell('total').font = { name: 'Segoe UI', size: 9.5, bold: true };
        row.getCell('rank').alignment = { horizontal: 'center' };
        row.getCell('document').alignment = { horizontal: 'center' };
        row.getCell('order').alignment = { horizontal: 'center' };
        row.getCell('issuedAt').alignment = { horizontal: 'center' };
        row.getCell('age').alignment = { horizontal: 'center' };
        row.getCell('tipo').alignment = { horizontal: 'center' };
      });

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `remisiones-${activeTab}-${new Date().toISOString().slice(0, 10)}.xlsx`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch (err) {
      console.error('Error exportando Excel:', err);
    } finally {
      setIsExportingExcel(false);
    }
  };

  return (
    <div className="live-monitor-container" style={{ padding: '0 0 32px 0' }}>
      {/* 1. Header en vivo */}
      <div
        style={{
          display: 'flex',
          justifyContent: 'space-between',
          alignItems: 'center',
          flexWrap: 'wrap',
          gap: '16px',
          marginBottom: '16px',
          padding: '18px 24px',
          background: 'linear-gradient(135deg, #0F172A 0%, #1E293B 100%)',
          borderRadius: '16px',
          color: '#FFFFFF',
          boxShadow: '0 4px 20px -2px rgba(15, 23, 42, 0.25)',
        }}
      >
        <div>
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '4px' }}>
            <span
              style={{
                width: '8px',
                height: '8px',
                borderRadius: '50%',
                backgroundColor: '#22C55E',
                boxShadow: '0 0 8px #22C55E',
              }}
            />
            <span style={{ fontSize: '12px', fontWeight: 700, color: '#4ADE80', textTransform: 'uppercase', letterSpacing: '0.04em' }}>
              En Vivo ERP
            </span>
            <span style={{ fontSize: '13px', color: '#94A3B8', marginLeft: '6px' }}>
              Última consulta: <strong style={{ color: '#F8FAFC' }}>{lastSyncTime || 'Hoy'}</strong>
            </span>
          </div>
          <h2 style={{ fontSize: '20px', fontWeight: 800, margin: 0, color: '#FFFFFF' }}>
            Movimientos en Tiempo Real · ¿Qué Salió y Qué Entró?
          </h2>
        </div>

        <button
          onClick={onSync}
          disabled={isLoading}
          style={{
            display: 'inline-flex',
            alignItems: 'center',
            gap: '8px',
            padding: '10px 20px',
            backgroundColor: '#2563EB',
            color: '#FFFFFF',
            border: 'none',
            borderRadius: '10px',
            fontSize: '13px',
            fontWeight: 700,
            cursor: isLoading ? 'not-allowed' : 'pointer',
            boxShadow: '0 4px 12px rgba(37, 99, 235, 0.35)',
            transition: 'all 0.2s ease',
            opacity: isLoading ? 0.7 : 1,
          }}
        >
          <RefreshCw size={16} className={isLoading ? 'spin-animation' : ''} />
          {isLoading ? 'Consultando ERP...' : 'Consultar ERP en Vivo Ahora'}
        </button>
      </div>

      {/* 2. Tarjetas Operativas Directas (Solo Qué Salió, Qué Entró y Qué Hay) */}
      <div
        style={{
          display: 'grid',
          gridTemplateColumns: 'repeat(auto-fit, minmax(260px, 1fr))',
          gap: '16px',
          marginBottom: '20px',
        }}
      >
        {/* Tarjeta 1: LO QUE SALIÓ */}
        <div
          onClick={() => setActiveTab('salientes')}
          style={{
            backgroundColor: activeTab === 'salientes' ? '#F0FDF4' : '#FFFFFF',
            border: activeTab === 'salientes' ? '2px solid #22C55E' : '1px solid #E2E8F0',
            borderRadius: '14px',
            padding: '18px 20px',
            cursor: 'pointer',
            transition: 'all 0.15s ease',
            boxShadow: activeTab === 'salientes' ? '0 4px 12px rgba(34, 197, 94, 0.15)' : 'none',
          }}
        >
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px' }}>
            <span style={{ fontSize: '13px', fontWeight: 800, color: '#15803D', textTransform: 'uppercase' }}>
              🟢 Lo que Salió (Facturadas)
            </span>
            <div
              style={{
                width: '34px',
                height: '34px',
                borderRadius: '8px',
                backgroundColor: '#DCFCE7',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                color: '#16A34A',
              }}
            >
              <ArrowUpRight size={20} />
            </div>
          </div>
          <div style={{ fontSize: '26px', fontWeight: 800, color: '#166534', marginBottom: '4px' }}>
            {currencyFmt.format(totalSalientesValor)}
          </div>
          <div style={{ fontSize: '13px', fontWeight: 700, color: '#15803D' }}>
            {salientes.length} remisiones que salieron
          </div>
          <div style={{ fontSize: '11px', color: '#64748B', marginTop: '4px' }}>
            Ya no están pendientes (fueron facturadas/cobradas)
          </div>
        </div>

        {/* Tarjeta 2: LO QUE ENTRÓ */}
        <div
          onClick={() => setActiveTab('entrantes')}
          style={{
            backgroundColor: activeTab === 'entrantes' ? '#EFF6FF' : '#FFFFFF',
            border: activeTab === 'entrantes' ? '2px solid #3B82F6' : '1px solid #E2E8F0',
            borderRadius: '14px',
            padding: '18px 20px',
            cursor: 'pointer',
            transition: 'all 0.15s ease',
            boxShadow: activeTab === 'entrantes' ? '0 4px 12px rgba(59, 130, 246, 0.15)' : 'none',
          }}
        >
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px' }}>
            <span style={{ fontSize: '13px', fontWeight: 800, color: '#1D4ED8', textTransform: 'uppercase' }}>
              🔵 Lo que Entró (Nuevas)
            </span>
            <div
              style={{
                width: '34px',
                height: '34px',
                borderRadius: '8px',
                backgroundColor: '#DBEAFE',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                color: '#2563EB',
              }}
            >
              <ArrowDownLeft size={20} />
            </div>
          </div>
          <div style={{ fontSize: '26px', fontWeight: 800, color: '#1E40AF', marginBottom: '4px' }}>
            {currencyFmt.format(totalEntrantesValor)}
          </div>
          <div style={{ fontSize: '13px', fontWeight: 700, color: '#1D4ED8' }}>
            {entrantes.length} remisiones nuevas
          </div>
          <div style={{ fontSize: '11px', color: '#64748B', marginTop: '4px' }}>
            Nuevos despachos de bodega recién emitidos
          </div>
        </div>

        {/* Tarjeta 3: TODAS LAS ACTIVAS */}
        <div
          onClick={() => setActiveTab('vivas')}
          style={{
            backgroundColor: activeTab === 'vivas' ? '#F8FAFC' : '#FFFFFF',
            border: activeTab === 'vivas' ? '2px solid #0F172A' : '1px solid #E2E8F0',
            borderRadius: '14px',
            padding: '18px 20px',
            cursor: 'pointer',
            transition: 'all 0.15s ease',
            boxShadow: activeTab === 'vivas' ? '0 4px 12px rgba(15, 23, 42, 0.08)' : 'none',
          }}
        >
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px' }}>
            <span style={{ fontSize: '13px', fontWeight: 800, color: '#0F172A', textTransform: 'uppercase' }}>
              📋 Todas las Activas en ERP
            </span>
            <div
              style={{
                width: '34px',
                height: '34px',
                borderRadius: '8px',
                backgroundColor: '#F1F5F9',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                color: '#334155',
              }}
            >
              <Package size={20} />
            </div>
          </div>
          <div style={{ fontSize: '26px', fontWeight: 800, color: '#0F172A', marginBottom: '4px' }}>
            {currencyFmt.format(totalVivasValor)}
          </div>
          <div style={{ fontSize: '13px', fontWeight: 700, color: '#475569' }}>
            {records.length} remisiones activas hoy
          </div>
          <div style={{ fontSize: '11px', color: '#64748B', marginTop: '4px' }}>
            Total de remisiones vivas en este segundo
          </div>
        </div>
      </div>

      {/* 3. Barra de Controles: Selector de Pestaña + Filtro Director + Buscador + Excel */}
      <div
        style={{
          display: 'flex',
          justifyContent: 'space-between',
          alignItems: 'center',
          flexWrap: 'wrap',
          gap: '12px',
          marginBottom: '16px',
          background: '#FFFFFF',
          padding: '12px 16px',
          borderRadius: '12px',
          border: '1px solid #E2E8F0',
        }}
      >
        {/* Selector de Pestañas */}
        <div style={{ display: 'flex', gap: '8px' }}>
          <button
            onClick={() => setActiveTab('salientes')}
            style={{
              padding: '8px 16px',
              fontSize: '13px',
              fontWeight: 700,
              borderRadius: '8px',
              border: 'none',
              cursor: 'pointer',
              backgroundColor: activeTab === 'salientes' ? '#16A34A' : '#F1F5F9',
              color: activeTab === 'salientes' ? '#FFFFFF' : '#475569',
              transition: 'all 0.15s ease',
            }}
          >
            🟢 Lo que Salió ({salientes.length})
          </button>
          <button
            onClick={() => setActiveTab('entrantes')}
            style={{
              padding: '8px 16px',
              fontSize: '13px',
              fontWeight: 700,
              borderRadius: '8px',
              border: 'none',
              cursor: 'pointer',
              backgroundColor: activeTab === 'entrantes' ? '#2563EB' : '#F1F5F9',
              color: activeTab === 'entrantes' ? '#FFFFFF' : '#475569',
              transition: 'all 0.15s ease',
            }}
          >
            🔵 Lo que Entró ({entrantes.length})
          </button>
          <button
            onClick={() => setActiveTab('vivas')}
            style={{
              padding: '8px 16px',
              fontSize: '13px',
              fontWeight: 700,
              borderRadius: '8px',
              border: 'none',
              cursor: 'pointer',
              backgroundColor: activeTab === 'vivas' ? '#0F172A' : '#F1F5F9',
              color: activeTab === 'vivas' ? '#FFFFFF' : '#475569',
              transition: 'all 0.15s ease',
            }}
          >
            📋 Total en ERP ({records.length})
          </button>
        </div>

        {/* Filtros: Director Comercial + Buscador + Excel */}
        <div style={{ display: 'flex', gap: '10px', alignItems: 'center', flexWrap: 'wrap' }}>
          {/* Filtro Director */}
          <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
            <Building2 size={16} color="#64748B" />
            <select
              value={selectedDirector}
              onChange={(e) => setSelectedDirector(e.target.value)}
              style={{
                padding: '7px 12px',
                fontSize: '12.5px',
                fontWeight: 600,
                border: '1px solid #CBD5E1',
                borderRadius: '8px',
                backgroundColor: '#FFFFFF',
                color: '#1E293B',
                outline: 'none',
              }}
            >
              {directorsList.map((d) => (
                <option key={d} value={d}>
                  {d === 'Todos' ? 'Todos los Directores' : d}
                </option>
              ))}
            </select>
          </div>

          {/* Buscador */}
          <div style={{ position: 'relative', width: '240px' }}>
            <Search
              size={15}
              style={{ position: 'absolute', left: '10px', top: '50%', transform: 'translateY(-50%)', color: '#94A3B8' }}
            />
            <input
              type="text"
              placeholder="Buscar cliente, NIT, remisión..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              style={{
                width: '100%',
                padding: '7px 10px 7px 32px',
                fontSize: '12.5px',
                border: '1px solid #CBD5E1',
                borderRadius: '8px',
                outline: 'none',
              }}
            />
          </div>

          {/* Botón Excel */}
          <button
            onClick={handleExportExcel}
            disabled={isExportingExcel}
            style={{
              display: 'inline-flex',
              alignItems: 'center',
              gap: '6px',
              padding: '7px 14px',
              background: '#FFFFFF',
              color: '#0F172A',
              border: '1px solid #CBD5E1',
              borderRadius: '8px',
              fontSize: '12.5px',
              fontWeight: 700,
              cursor: isExportingExcel ? 'not-allowed' : 'pointer',
            }}
            title="Descargar listado a Excel"
          >
            <Download size={14} color="#16A34A" />
            Excel ({currentList.length})
          </button>
        </div>
      </div>

      {/* 4. Tabla Directa de Remisiones */}
      <div
        style={{
          backgroundColor: '#FFFFFF',
          borderRadius: '14px',
          border: '1px solid #E2E8F0',
          overflow: 'hidden',
          boxShadow: '0 2px 8px rgba(0, 0, 0, 0.04)',
        }}
      >
        <div
          style={{
            padding: '12px 18px',
            background: activeTab === 'salientes' ? '#F0FDF4' : activeTab === 'entrantes' ? '#EFF6FF' : '#F8FAFC',
            borderBottom: '1px solid #E2E8F0',
            display: 'flex',
            justifyContent: 'space-between',
            alignItems: 'center',
            flexWrap: 'wrap',
            gap: '8px',
          }}
        >
          <span style={{ fontSize: '13px', color: '#1E293B', fontWeight: 700 }}>
            {activeTab === 'salientes'
              ? `🟢 Listado de lo que Salió (Facturadas / Cobradas): ${currentList.length} remisiones`
              : activeTab === 'entrantes'
              ? `🔵 Listado de lo que Entró (Nuevas Remisiones de Bodega): ${currentList.length} remisiones`
              : `📋 Total de Remisiones Activas en ERP: ${currentList.length} remisiones`}
          </span>
          <span style={{ fontSize: '13px', color: '#64748B' }}>
            Suma total:{' '}
            <strong
              style={{
                color: activeTab === 'salientes' ? '#166534' : activeTab === 'entrantes' ? '#1E40AF' : '#0F172A',
                fontSize: '14px',
              }}
            >
              {currencyFmt.format(currentList.reduce((s, r) => s + (r.total || 0), 0))}
            </strong>
          </span>
        </div>

        <div style={{ overflowX: 'auto' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left', fontSize: '13px' }}>
            <thead>
              <tr style={{ backgroundColor: '#F8FAFC', borderBottom: '1px solid #E2E8F0' }}>
                <th style={{ padding: '12px 16px', fontWeight: 700, color: '#475569' }}>No. Remisión</th>
                <th style={{ padding: '12px 16px', fontWeight: 700, color: '#475569' }}>Pedido</th>
                <th style={{ padding: '12px 16px', fontWeight: 700, color: '#475569' }}>Cliente / Razón Social</th>
                <th style={{ padding: '12px 16px', fontWeight: 700, color: '#475569' }}>Comercial</th>
                <th style={{ padding: '12px 16px', fontWeight: 700, color: '#475569' }}>Director</th>
                <th style={{ padding: '12px 16px', fontWeight: 700, color: '#475569' }}>Fecha Emisión</th>
                <th style={{ padding: '12px 16px', fontWeight: 700, color: '#475569', textAlign: 'right' }}>
                  Valor ($)
                </th>
                <th style={{ padding: '12px 16px', fontWeight: 700, color: '#475569', textAlign: 'center' }}>
                  Productos
                </th>
              </tr>
            </thead>
            <tbody>
              {currentList.length === 0 ? (
                <tr>
                  <td colSpan={8} style={{ padding: '40px', textAlign: 'center', color: '#94A3B8' }}>
                    No hay remisiones para mostrar con los filtros aplicados.
                  </td>
                </tr>
              ) : (
                currentList.map((r, i) => {
                  const num = parseInt(r.document.replace(/\D/g, ''), 10);
                  const prods = productsMap?.get(num) || [];

                  return (
                    <tr
                      key={r.id || `${r.document}-${i}`}
                      style={{
                        borderBottom: '1px solid #F1F5F9',
                        backgroundColor: i % 2 === 0 ? '#FFFFFF' : '#FAFAFA',
                        transition: 'background-color 0.15s',
                      }}
                    >
                      <td style={{ padding: '12px 16px', fontWeight: 700, color: '#0F172A' }}>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                          <span>{r.document || '—'}</span>
                          {activeTab === 'salientes' && (
                            <span
                              style={{
                                fontSize: '10px',
                                background: '#DCFCE7',
                                color: '#166534',
                                padding: '2px 6px',
                                borderRadius: '4px',
                                fontWeight: 700,
                              }}
                            >
                              Salió
                            </span>
                          )}
                          {activeTab === 'entrantes' && (
                            <span
                              style={{
                                fontSize: '10px',
                                background: '#DBEAFE',
                                color: '#1D4ED8',
                                padding: '2px 6px',
                                borderRadius: '4px',
                                fontWeight: 700,
                              }}
                            >
                              Entró
                            </span>
                          )}
                        </div>
                      </td>
                      <td style={{ padding: '12px 16px', color: '#64748B' }}>
                        <span
                          style={{
                            padding: '2px 8px',
                            background: '#F1F5F9',
                            borderRadius: '6px',
                            fontSize: '12px',
                            fontWeight: 600,
                          }}
                        >
                          {r.order || '—'}
                        </span>
                      </td>
                      <td style={{ padding: '12px 16px' }}>
                        <div style={{ fontWeight: 600, color: '#1E293B' }}>{r.company}</div>
                        <div style={{ fontSize: '11px', color: '#94A3B8' }}>NIT: {r.nit}</div>
                      </td>
                      <td style={{ padding: '12px 16px', color: '#334155', fontWeight: 500 }}>
                        {r.employee}
                      </td>
                      <td style={{ padding: '12px 16px', color: '#475569' }}>
                        {getRecordDirector(r)}
                      </td>
                      <td style={{ padding: '12px 16px', color: '#64748B', whiteSpace: 'nowrap' }}>
                        {r.issuedAt} ({r.age} d)
                      </td>
                      <td
                        style={{
                          padding: '12px 16px',
                          textAlign: 'right',
                          fontWeight: 700,
                          color:
                            activeTab === 'salientes'
                              ? '#15803D'
                              : activeTab === 'entrantes'
                              ? '#1D4ED8'
                              : '#0F172A',
                          whiteSpace: 'nowrap',
                        }}
                      >
                        {currencyFmt.format(r.total)}
                      </td>
                      <td style={{ padding: '12px 16px', textAlign: 'center' }}>
                        {prods.length > 0 ? (
                          <button
                            onClick={() => setSelectedProductRemision(r)}
                            style={{
                              display: 'inline-flex',
                              alignItems: 'center',
                              gap: '4px',
                              padding: '4px 10px',
                              backgroundColor: '#EEF2FF',
                              color: '#4F46E5',
                              border: '1px solid #C7D2FE',
                              borderRadius: '6px',
                              fontSize: '11.5px',
                              fontWeight: 600,
                              cursor: 'pointer',
                            }}
                          >
                            <Package size={13} />
                            {prods.length} {prods.length === 1 ? 'ítem' : 'ítems'}
                          </button>
                        ) : (
                          <span style={{ fontSize: '11px', color: '#CBD5E1' }}>—</span>
                        )}
                      </td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* 5. Modal de detalle de productos de la remisión */}
      {selectedProductRemision && (
        <div
          style={{
            position: 'fixed',
            top: 0,
            left: 0,
            right: 0,
            bottom: 0,
            backgroundColor: 'rgba(15, 23, 42, 0.65)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            zIndex: 9999,
            padding: '16px',
          }}
        >
          <div
            style={{
              backgroundColor: '#FFFFFF',
              borderRadius: '16px',
              maxWidth: '650px',
              width: '100%',
              maxHeight: '85vh',
              display: 'flex',
              flexDirection: 'column',
              boxShadow: '0 20px 25px -5px rgba(0, 0, 0, 0.25)',
              overflow: 'hidden',
            }}
          >
            <div
              style={{
                padding: '20px 24px',
                borderBottom: '1px solid #E2E8F0',
                display: 'flex',
                justifyContent: 'space-between',
                alignItems: 'center',
                backgroundColor: '#F8FAFC',
              }}
            >
              <div>
                <h3 style={{ margin: 0, fontSize: '17px', fontWeight: 800, color: '#0F172A' }}>
                  Remisión #{selectedProductRemision.document} · Detalle de Productos
                </h3>
                <div style={{ fontSize: '12px', color: '#64748B', marginTop: '2px' }}>
                  {selectedProductRemision.company} (Pedido: {selectedProductRemision.order || 'Sin pedido'})
                </div>
              </div>
              <button
                onClick={() => setSelectedProductRemision(null)}
                style={{
                  background: 'none',
                  border: 'none',
                  cursor: 'pointer',
                  color: '#64748B',
                  padding: '4px',
                }}
              >
                <X size={20} />
              </button>
            </div>

            <div style={{ padding: '20px 24px', overflowY: 'auto' }}>
              <div
                style={{
                  display: 'flex',
                  justifyContent: 'space-between',
                  padding: '12px 16px',
                  backgroundColor: '#F1F5F9',
                  borderRadius: '10px',
                  marginBottom: '16px',
                }}
              >
                <div>
                  <div style={{ fontSize: '11px', color: '#64748B' }}>Asesor Comercial</div>
                  <div style={{ fontSize: '13px', fontWeight: 700, color: '#0F172A' }}>
                    {selectedProductRemision.employee}
                  </div>
                </div>
                <div style={{ textAlign: 'right' }}>
                  <div style={{ fontSize: '11px', color: '#64748B' }}>Valor Total Remisión</div>
                  <div style={{ fontSize: '15px', fontWeight: 800, color: '#15803D' }}>
                    {currencyFmt.format(selectedProductRemision.total)}
                  </div>
                </div>
              </div>

              <h4 style={{ fontSize: '13px', fontWeight: 700, color: '#334155', marginBottom: '10px' }}>
                Mercancía física ({activeProducts.length} ítems):
              </h4>

              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12.5px' }}>
                <thead>
                  <tr style={{ borderBottom: '1px solid #E2E8F0', color: '#64748B' }}>
                    <th style={{ padding: '8px', textAlign: 'left' }}>Código</th>
                    <th style={{ padding: '8px', textAlign: 'left' }}>Producto / Descripción</th>
                    <th style={{ padding: '8px', textAlign: 'right' }}>Cant.</th>
                  </tr>
                </thead>
                <tbody>
                  {activeProducts.map((p, idx) => (
                    <tr key={idx} style={{ borderBottom: '1px solid #F1F5F9' }}>
                      <td style={{ padding: '10px 8px', color: '#64748B', fontFamily: 'monospace' }}>
                        {p.codigo || '—'}
                      </td>
                      <td style={{ padding: '10px 8px', fontWeight: 600, color: '#1E293B' }}>
                        {p.descripcion}
                      </td>
                      <td style={{ padding: '10px 8px', textAlign: 'right', fontWeight: 700, color: '#0F172A' }}>
                        {p.cantidad}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <div
              style={{
                padding: '16px 24px',
                borderTop: '1px solid #E2E8F0',
                display: 'flex',
                justifyContent: 'flex-end',
                backgroundColor: '#F8FAFC',
              }}
            >
              <button
                onClick={() => setSelectedProductRemision(null)}
                style={{
                  padding: '8px 20px',
                  backgroundColor: '#0F172A',
                  color: '#FFFFFF',
                  border: 'none',
                  borderRadius: '8px',
                  fontSize: '13px',
                  fontWeight: 600,
                  cursor: 'pointer',
                }}
              >
                Cerrar
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
