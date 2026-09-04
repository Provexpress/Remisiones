import { useEffect, useMemo, useRef, useState } from 'react';
import {
  ArrowDownRight,
  ArrowUpRight,
  BadgeDollarSign,
  Boxes,
  CalendarRange,
  CheckCircle2,
  ChevronDown,
  CircleDollarSign,
  Cloud,
  Download,
  FileSpreadsheet,
  LayoutDashboard,
  Filter,
  ListChecks,
  LoaderCircle,
  LogIn,
  LogOut,
  PackageCheck,
  RefreshCw,
  RotateCcw,
  Search,
  ShieldCheck,
  TriangleAlert,
  Upload,
  UsersRound,
  X,
} from 'lucide-react';
import {
  Area,
  AreaChart,
  Bar,
  BarChart,
  CartesianGrid,
  Cell,
  Legend,
  Pie,
  PieChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from 'recharts';
import { getExistingProfile, loadSharePointWorkbook, signIn, signOut } from './lib/auth';
import {
  AGE_ORDER,
  aggregateBy,
  buildAgeBreakdown,
  buildDailySeries,
  formatCutoff,
  formatDateTime,
  formatTimeOnly,
  normalizeText,
  parseRemisionesWorkbook,
  summarize,
} from './lib/remisiones';
import type { AgeBreakdownItem, DataSource, FileMetadata, ParsedWorkbook, Remision, Summary, UserProfile } from './types';

type Phase = 'welcome' | 'loading' | 'ready' | 'error';
type View = 'overview' | 'detail';

const currency = new Intl.NumberFormat('es-CO', {
  style: 'currency',
  currency: 'COP',
  maximumFractionDigits: 0,
});
const compactCurrency = new Intl.NumberFormat('es-CO', {
  style: 'currency',
  currency: 'COP',
  notation: 'compact',
  compactDisplay: 'short',
  maximumFractionDigits: 1,
});
const number = new Intl.NumberFormat('es-CO', { maximumFractionDigits: 0 });
const percent = new Intl.NumberFormat('es-CO', { style: 'percent', maximumFractionDigits: 1 });
const PIE_COLORS = ['#0066cc', '#30b96b', '#ff9f0a', '#af52de', '#ff453a', '#64d2ff'];

function App() {
  const [phase, setPhase] = useState<Phase>('welcome');
  const [user, setUser] = useState<UserProfile | null>(null);
  const [workbook, setWorkbook] = useState<ParsedWorkbook | null>(null);
  const [source, setSource] = useState<DataSource>('sharepoint');
  const [metadata, setMetadata] = useState<FileMetadata | null>(null);
  const [error, setError] = useState('');
  const [loadingMessage, setLoadingMessage] = useState('Conectando con Microsoft 365…');
  const fileInput = useRef<HTMLInputElement>(null);

  const parseAndShow = async (buffer: ArrayBuffer, nextSource: DataSource, nextMetadata: FileMetadata) => {
    setLoadingMessage('Leyendo datos y grupos comerciales…');
    const parsed = await parseRemisionesWorkbook(buffer, {
      lastModifiedDateTime: nextMetadata.lastModifiedDateTime,
    });
    if (parsed.cutoffDateTime) {
      nextMetadata.lastModifiedDateTime = parsed.cutoffDateTime;
    }
    setWorkbook(parsed);
    setSource(nextSource);
    setMetadata(nextMetadata);
    setPhase('ready');
  };

  const loadRemote = async (profile?: UserProfile) => {
    setError('');
    setPhase('loading');
    setLoadingMessage('Conectando con Microsoft 365…');
    try {
      const activeProfile = profile || await signIn();
      setUser(activeProfile);
      setLoadingMessage('Descargando Remisiones.xlsx desde SharePoint…');
      const result = await loadSharePointWorkbook();
      await parseAndShow(result.buffer, 'sharepoint', result.metadata);
    } catch (reason) {
      const message = reason instanceof Error ? reason.message : 'No fue posible cargar el archivo.';
      setError(message);
      setPhase('error');
    }
  };

  useEffect(() => {
    let mounted = true;
    getExistingProfile()
      .then((profile) => {
        if (mounted && profile) {
          setUser(profile);
          void loadRemote(profile);
        }
      })
      .catch(() => undefined);
    return () => { mounted = false; };
    // La restauración de sesión solo se ejecuta al montar.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const handleLocalFile = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    setError('');
    setPhase('loading');
    setLoadingMessage('Abriendo el Excel local…');
    try {
      await parseAndShow(await file.arrayBuffer(), 'local', {
        name: file.name,
        size: file.size,
        lastModifiedDateTime: new Date(file.lastModified).toISOString(),
      });
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : 'No fue posible leer el Excel local.');
      setPhase('error');
    } finally {
      event.target.value = '';
    }
  };

  const refresh = async () => {
    if (source === 'sharepoint') await loadRemote(user || undefined);
    else fileInput.current?.click();
  };

  const logout = async () => {
    if (user) await signOut().catch(() => undefined);
    setUser(null);
    setWorkbook(null);
    setMetadata(null);
    setPhase('welcome');
  };

  return (
    <>
      <input
        ref={fileInput}
        className="visually-hidden"
        type="file"
        accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        onChange={handleLocalFile}
      />
      {phase === 'loading' && <LoadingScreen message={loadingMessage} />}
      {(phase === 'welcome' || phase === 'error') && (
        <WelcomeScreen
          error={error}
          onMicrosoft={() => void loadRemote()}
          onLocal={() => fileInput.current?.click()}
        />
      )}
      {phase === 'ready' && workbook && (
        <Dashboard
          data={workbook}
          user={user}
          source={source}
          metadata={metadata}
          onRefresh={() => void refresh()}
          onLogout={() => void logout()}
        />
      )}
    </>
  );
}

function WelcomeScreen({
  error,
  onMicrosoft,
  onLocal,
}: {
  error: string;
  onMicrosoft: () => void;
  onLocal: () => void;
}) {
  return (
    <main className="welcome-shell">
      <section className="welcome-copy">
        <div className="brand-lockup">
          <img src="/Logo.webp" alt="Provexpress" />
          <span>Control operativo</span>
        </div>
        <div className="eyebrow"><ShieldCheck size={16} /> Portal corporativo</div>
        <h1>Remisiones abiertas,<br /><span>bajo control.</span></h1>
        <p>
          Revisa el pendiente por facturar, identifica los casos críticos y sigue la gestión diaria
          de cada equipo desde una sola vista.
        </p>
        {error && (
          <div className="welcome-error" role="alert">
            <TriangleAlert size={19} />
            <div><strong>No pudimos abrir el archivo</strong><span>{error}</span></div>
          </div>
        )}
        <div className="welcome-actions">
          <button className="button button-primary button-large" onClick={onMicrosoft}>
            <LogIn size={19} /> Iniciar con Microsoft
          </button>
          <button className="button button-secondary button-large" onClick={onLocal}>
            <Upload size={19} /> Abrir Excel local
          </button>
        </div>
        <div className="trust-row">
          <span><CheckCircle2 size={16} /> Acceso con cuenta Provexpress</span>
          <span><Cloud size={16} /> Lectura directa desde SharePoint</span>
        </div>
      </section>
      <aside className="welcome-visual" aria-label="Resumen del portal">
        <div className="visual-orb visual-orb-one" />
        <div className="visual-orb visual-orb-two" />
        <div className="preview-window">
          <div className="preview-toolbar"><i /><i /><i /><span>Remisiones · Provexpress</span></div>
          <div className="preview-content">
            <div className="preview-label">PENDIENTE TOTAL</div>
            <div className="preview-amount">$ 3.600 MM</div>
            <div className="preview-change">Corte diario consolidado</div>
            <div className="preview-chart">
              {[34, 46, 39, 58, 51, 68, 62, 79, 72, 88, 82, 94].map((height, index) => (
                <span key={index} style={{ height: `${height}%` }} />
              ))}
            </div>
            <div className="preview-cards">
              <div><span>Remisiones</span><strong>674</strong></div>
              <div><span>&gt;30 días</span><strong>105</strong></div>
              <div><span>Equipos</span><strong>4</strong></div>
            </div>
          </div>
        </div>
      </aside>
    </main>
  );
}

function LoadingScreen({ message }: { message: string }) {
  return (
    <main className="loading-screen" aria-live="polite">
      <div className="loading-card">
        <img src="/Logo.webp" alt="Provexpress" />
        <div className="loading-icon"><LoaderCircle size={26} /></div>
        <h2>Preparando tu vista</h2>
        <p>{message}</p>
        <div className="loading-track"><span /></div>
      </div>
    </main>
  );
}

function Dashboard({
  data,
  user,
  source,
  metadata,
  onRefresh,
  onLogout,
}: {
  data: ParsedWorkbook;
  user: UserProfile | null;
  source: DataSource;
  metadata: FileMetadata | null;
  onRefresh: () => void;
  onLogout: () => void;
}) {
  const [view, setView] = useState<View>('overview');
  const [cutoff, setCutoff] = useState(data.cutoffs.at(-1) || '');
  const [from, setFrom] = useState(data.cutoffs.length > 30 ? data.cutoffs.at(-30)! : data.cutoffs[0]);
  const [to, setTo] = useState(data.cutoffs.at(-1) || '');
  const [director, setDirector] = useState('Todos');
  const [employee, setEmployee] = useState('Todos');
  const [statusFilter, setStatusFilter] = useState('Todos');
  const [ageFilter, setAgeFilter] = useState('Todos');
  const [amountFilter, setAmountFilter] = useState('Todos');
  const [sortBy, setSortBy] = useState<'total-desc' | 'total-asc' | 'age-desc' | 'age-asc'>('total-desc');
  const [query, setQuery] = useState('');
  const [page, setPage] = useState(1);
  const refreshRef = useRef(onRefresh);

  const matchesAgeFilter = (age: number, ageRange: string, filter: string): boolean => {
    if (filter === 'Todos') return true;
    if (filter === '>30 días' || filter === 'Vencidas (>30 días)') return age > 30;
    if (filter === 'Por vencer (16-30 días)' || filter === '16-30 días') return age > 15 && age <= 30;
    if (filter === 'Al día (0-15 días)' || filter === '0-15 días') return age <= 15;
    if (filter === 'Críticas (>60 días)' || filter === 'Crítica (>60 días)' || filter === '>60 días') return age > 60;
    if (filter === 'Vencida (31-60 días)' || filter === '31-60 días') return age > 30 && age <= 60;
    return ageRange === filter;
  };

  const resetAllFilters = () => {
    setDirector('Todos');
    setEmployee('Todos');
    setStatusFilter('Todos');
    setAgeFilter('Todos');
    setAmountFilter('Todos');
    setQuery('');
  };

  const activeFiltersCount =
    (director !== 'Todos' ? 1 : 0) +
    (employee !== 'Todos' ? 1 : 0) +
    (statusFilter !== 'Todos' ? 1 : 0) +
    (ageFilter !== 'Todos' ? 1 : 0) +
    (amountFilter !== 'Todos' ? 1 : 0) +
    (query ? 1 : 0);

  const hasActiveFilters = activeFiltersCount > 0;

  const directors = useMemo(
    () => [...new Set(data.records.map((record) => record.director))].sort((a, b) => a.localeCompare(b, 'es')),
    [data.records],
  );
  const employees = useMemo(() => {
    const scoped = director === 'Todos' ? data.records : data.records.filter((record) => record.director === director);
    return [...new Set(scoped.map((record) => record.employee))].sort((a, b) => a.localeCompare(b, 'es'));
  }, [data.records, director]);

  const baseCutoffRecords = useMemo(
    () => data.records.filter((record) => record.cutoff === cutoff),
    [data.records, cutoff],
  );

  const currentRecords = useMemo(() => {
    return baseCutoffRecords.filter((record) => {
      if (director !== 'Todos' && record.director !== director) return false;
      if (employee !== 'Todos' && record.employee !== employee) return false;
      if (statusFilter !== 'Todos' && record.alert !== statusFilter) return false;
      if (amountFilter !== 'Todos' && record.amountStatus !== amountFilter) return false;
      if (!matchesAgeFilter(record.age, record.ageRange, ageFilter)) return false;
      return true;
    });
  }, [baseCutoffRecords, director, employee, statusFilter, amountFilter, ageFilter]);

  const currentSummary = useMemo(() => summarize(currentRecords), [currentRecords]);
  const previousCutoff = data.cutoffs.filter((date) => date < cutoff).at(-1) || '';
  const previousSummary = useMemo(
    () => summarize(data.records.filter((record) =>
      record.cutoff === previousCutoff &&
      (director === 'Todos' || record.director === director) &&
      (employee === 'Todos' || record.employee === employee) &&
      (statusFilter === 'Todos' || record.alert === statusFilter) &&
      (amountFilter === 'Todos' || record.amountStatus === amountFilter) &&
      matchesAgeFilter(record.age, record.ageRange, ageFilter)
    )),
    [data.records, previousCutoff, director, employee, statusFilter, amountFilter, ageFilter],
  );

  const periodRecords = useMemo(
    () => data.records.filter((record) =>
      record.cutoff >= from && record.cutoff <= to &&
      (director === 'Todos' || record.director === director) &&
      (employee === 'Todos' || record.employee === employee)),
    [data.records, from, to, director, employee],
  );
  const daily = useMemo(() => buildDailySeries(periodRecords), [periodRecords]);

  const ageBreakdownRecords = useMemo(
    () => baseCutoffRecords.filter((r) =>
      (director === 'Todos' || r.director === director) &&
      (employee === 'Todos' || r.employee === employee) &&
      (statusFilter === 'Todos' || r.alert === statusFilter) &&
      (amountFilter === 'Todos' || r.amountStatus === amountFilter)),
    [baseCutoffRecords, director, employee, statusFilter, amountFilter],
  );
  const ageBreakdown = useMemo(() => buildAgeBreakdown(ageBreakdownRecords), [ageBreakdownRecords]);

  const directorFilterRecords = useMemo(
    () => baseCutoffRecords.filter((r) =>
      (statusFilter === 'Todos' || r.alert === statusFilter) &&
      (amountFilter === 'Todos' || r.amountStatus === amountFilter) &&
      matchesAgeFilter(r.age, r.ageRange, ageFilter)),
    [baseCutoffRecords, statusFilter, amountFilter, ageFilter],
  );
  const directorData = useMemo(() => aggregateBy(directorFilterRecords, (record) => record.director), [directorFilterRecords]);

  const sellerFilterRecords = useMemo(
    () => baseCutoffRecords.filter((r) =>
      (director === 'Todos' || r.director === director) &&
      (statusFilter === 'Todos' || r.alert === statusFilter) &&
      (amountFilter === 'Todos' || r.amountStatus === amountFilter) &&
      matchesAgeFilter(r.age, r.ageRange, ageFilter)),
    [baseCutoffRecords, director, statusFilter, amountFilter, ageFilter],
  );
  const sellerData = useMemo(() => aggregateBy(sellerFilterRecords, (record) => record.employee).slice(0, 12), [sellerFilterRecords]);

  const detailRecords = useMemo(() => {
    const normalizedQuery = normalizeText(query);
    return currentRecords
      .filter((record) => !normalizedQuery || normalizeText([
        record.company,
        record.nit,
        record.employee,
        record.director,
        record.document,
        record.order,
      ].join(' ')).includes(normalizedQuery))
      .sort((a, b) => {
        if (sortBy === 'total-desc') return b.total - a.total || b.age - a.age;
        if (sortBy === 'total-asc') return a.total - b.total || b.age - a.age;
        if (sortBy === 'age-desc') return b.age - a.age || b.total - a.total;
        if (sortBy === 'age-asc') return a.age - b.age || b.total - a.total;
        return b.age - a.age;
      });
  }, [currentRecords, sortBy, query]);

  const pageSize = 20;
  const pageCount = Math.max(1, Math.ceil(detailRecords.length / pageSize));
  const visibleRows = detailRecords.slice((page - 1) * pageSize, page * pageSize);

  useEffect(() => setPage(1), [query, statusFilter, amountFilter, ageFilter, sortBy, cutoff, director, employee]);
  useEffect(() => { refreshRef.current = onRefresh; }, [onRefresh]);
  useEffect(() => {
    if (source !== 'sharepoint') return undefined;
    const interval = window.setInterval(() => refreshRef.current(), 10 * 60 * 1000);
    return () => window.clearInterval(interval);
  }, [source]);
  useEffect(() => {
    if (employee !== 'Todos' && !employees.includes(employee)) setEmployee('Todos');
  }, [director, employee, employees]);

  useEffect(() => {
    if (!currentRecords.length) return;
    try {
      const raw = localStorage.getItem('provexpress_remisiones_history');
      const existing = raw ? JSON.parse(raw) : {};
      const point = daily.find((d) => d.cutoff === cutoff) || {
        cutoff,
        pending: currentSummary.pending,
        remissions: currentSummary.remissions,
        clients: currentSummary.clients,
        newValue: currentSummary.pending,
        newCount: currentSummary.remissions,
        previousBalance: 0,
        withdrawn: 0,
        withdrawnCount: 0,
        netManagement: 0,
        grossReduction: 0,
        overdueValue: currentSummary.overdueValue,
        overdueCount: currentSummary.overdueCount,
      };
      existing[cutoff] = point;
      localStorage.setItem('provexpress_remisiones_history', JSON.stringify(existing));
    } catch {
      // safe fallback
    }
  }, [cutoff, currentSummary, daily, currentRecords.length]);

  const exportCsv = () => {
    const headers = [
      'Corte',
      'Hora_Corte',
      'Director',
      'Comercial',
      'NIT',
      'Empresa',
      'Remisión',
      'Pedido',
      'Emisión',
      'Días',
      'Estado_Días',
      'Mercancía',
      'IVA',
      'Total',
      'Estado_Monto',
      'Estado',
    ];
    const rows = detailRecords.map((record) => [
      record.cutoff,
      record.cutoffTime || '',
      record.director,
      record.employee,
      record.nit,
      record.company,
      record.document,
      record.order,
      record.issuedAt,
      record.age,
      record.daysStatus,
      record.merchandise,
      record.tax,
      record.total,
      record.amountStatus,
      record.alert,
    ]);
    const csv = [headers, ...rows].map((row) => row.map((cell) => `"${String(cell).replace(/"/g, '""')}"`).join(';')).join('\r\n');
    const url = URL.createObjectURL(new Blob([`\uFEFF${csv}`], { type: 'text/csv;charset=utf-8' }));
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = `remisiones-${cutoff}.csv`;
    anchor.click();
    URL.revokeObjectURL(url);
  };

  const formattedCutoffWithTime = useMemo(() => {
    const dateLabel = formatCutoff(cutoff);
    return data.cutoffTimeDisplay ? `${dateLabel} · ${data.cutoffTimeDisplay}` : dateLabel;
  }, [cutoff, data.cutoffTimeDisplay]);

  return (
    <div className="app-shell">
      <header className="topbar">
        <div className="topbar-inner">
          <div className="app-brand">
            <img src="/Logo.webp" alt="Provexpress" />
            <span className="app-divider" />
            <div><strong>Remisiones</strong><small>Control de pendientes</small></div>
          </div>
          <nav className="segmented-nav" aria-label="Vistas">
            <button className={view === 'overview' ? 'active' : ''} onClick={() => setView('overview')}>
              <LayoutDashboard size={16} /> Resumen
            </button>
            <button className={view === 'detail' ? 'active' : ''} onClick={() => setView('detail')}>
              <ListChecks size={16} /> Detalle
            </button>
          </nav>
          <div className="topbar-actions">
            <button className="icon-button" onClick={onRefresh} title="Actualizar datos"><RefreshCw size={18} /></button>
            <div className="user-chip">
              <span className="avatar">{initials(user?.name || 'Archivo local')}</span>
              <div><strong>{user?.name?.split(' ')[0] || 'Vista local'}</strong><small>{source === 'sharepoint' ? 'Microsoft 365' : 'Excel local'}</small></div>
            </div>
            <button className="icon-button" onClick={onLogout} title="Salir"><LogOut size={18} /></button>
          </div>
        </div>
      </header>

      <main className="dashboard-main">
        <section className="hero-row">
          <div>
            <div className="eyebrow blue"><PackageCheck size={16} /> Seguimiento diario</div>
            <h1>Remisiones abiertas</h1>
            <p>
              Una lectura clara del pendiente por facturar y de la gestión de cada equipo.
              {data.cutoffTimeDisplay && (
                <span className="hero-time-tag"> · Hora de corte: <b>{data.cutoffTimeDisplay}</b></span>
              )}
            </p>
          </div>
          <div className="source-card">
            <span className={`source-icon ${source}`}><FileSpreadsheet size={21} /></span>
            <div>
              <strong>{metadata?.name || 'Remisiones.xlsx'}</strong>
              <small>
                {source === 'sharepoint' ? 'SharePoint' : 'Archivo local'}
                {data.cutoffTimeDisplay ? ` · Corte ${data.cutoffTimeDisplay}` : ''}
                {data.activeSheetName ? ` · Hoja ${data.activeSheetName}` : ''}
              </small>
            </div>
            <span className="source-status"><i /> Sincronizado</span>
          </div>
        </section>

        <section className="filter-bar" aria-label="Filtros del tablero">
          <SelectFilter
            label="Corte"
            value={cutoff}
            onChange={setCutoff}
            options={data.cutoffs}
            format={(val) => data.cutoffTimeDisplay ? `${formatCutoff(val)} · ${data.cutoffTimeDisplay}` : formatCutoff(val)}
          />
          <SelectFilter label="Director" value={director} onChange={setDirector} options={['Todos', ...directors]} />
          <SelectFilter label="Ejecutivo" value={employee} onChange={setEmployee} options={['Todos', ...employees]} />
          <div className="filter-separator" />
          <label className="date-filter"><span>Desde</span><input type="date" value={from} min={data.cutoffs[0]} max={to} onChange={(event) => setFrom(event.target.value)} /></label>
          <label className="date-filter"><span>Hasta</span><input type="date" value={to} min={from} max={data.cutoffs.at(-1)} onChange={(event) => setTo(event.target.value)} /></label>
        </section>

        {hasActiveFilters && (
          <section className="active-filters-bar" aria-label="Filtros aplicados">
            <div className="active-filters-title">
              <Filter size={14} />
              <span>Filtros activos ({activeFiltersCount}):</span>
            </div>
            <div className="active-pills-wrap">
              {director !== 'Todos' && (
                <button
                  type="button"
                  className="filter-pill-chip"
                  onClick={() => setDirector('Todos')}
                  title="Toca para quitar filtro de director"
                >
                  <span>Director: <b>{director}</b></span>
                  <X size={13} />
                </button>
              )}
              {employee !== 'Todos' && (
                <button
                  type="button"
                  className="filter-pill-chip"
                  onClick={() => setEmployee('Todos')}
                  title="Toca para quitar filtro de comercial"
                >
                  <span>Comercial: <b>{employee}</b></span>
                  <X size={13} />
                </button>
              )}
              {ageFilter !== 'Todos' && (
                <button
                  type="button"
                  className="filter-pill-chip"
                  onClick={() => setAgeFilter('Todos')}
                  title="Toca para quitar filtro de días"
                >
                  <span>Días: <b>{ageFilter}</b></span>
                  <X size={13} />
                </button>
              )}
              {amountFilter !== 'Todos' && (
                <button
                  type="button"
                  className="filter-pill-chip"
                  onClick={() => setAmountFilter('Todos')}
                  title="Toca para quitar filtro de monto"
                >
                  <span>Monto: <b>{amountFilter}</b></span>
                  <X size={13} />
                </button>
              )}
              {statusFilter !== 'Todos' && (
                <button
                  type="button"
                  className="filter-pill-chip"
                  onClick={() => setStatusFilter('Todos')}
                  title="Toca para quitar filtro de estado"
                >
                  <span>Estado: <b>{statusFilter}</b></span>
                  <X size={13} />
                </button>
              )}
              {query && (
                <button
                  type="button"
                  className="filter-pill-chip"
                  onClick={() => setQuery('')}
                  title="Toca para limpiar búsqueda"
                >
                  <span>Búsqueda: <b>"{query}"</b></span>
                  <X size={13} />
                </button>
              )}
              <button
                type="button"
                className="clear-all-pill"
                onClick={resetAllFilters}
                title="Restablecer todos los filtros"
              >
                <RotateCcw size={13} /> Limpiar todo
              </button>
            </div>
          </section>
        )}

        {view === 'overview' ? (
          <>
            <section className="metric-grid">
              <MetricCard
                title="Pendiente total"
                value={currency.format(currentSummary.pending)}
                sub={
                  hasActiveFilters
                    ? `${number.format(currentSummary.remissions)} remisiones · Toca para ver todo`
                    : `${number.format(currentSummary.remissions)} remisiones abiertas`
                }
                icon={<CircleDollarSign />}
                tone="blue"
                current={currentSummary.pending}
                previous={previousSummary.pending}
                onClick={hasActiveFilters ? resetAllFilters : undefined}
                clickable={hasActiveFilters}
              />
              <MetricCard
                title="Clientes con saldo"
                value={number.format(currentSummary.clients)}
                sub={`${number.format(currentSummary.remissions)} remisiones pendientes`}
                icon={<Boxes />}
                tone="purple"
                current={currentSummary.clients}
                previous={previousSummary.clients}
              />
              <MetricCard
                title="Antigüedad promedio"
                value={`${currentSummary.averageAge.toFixed(1)} días`}
                sub={
                  ageFilter === '>30 días'
                    ? 'Filtro >30d activo · Toca para ver todos'
                    : `Máximo: ${currentSummary.maxAge || 0} días · Toca para filtrar >30d`
                }
                icon={<CalendarRange />}
                tone="orange"
                current={currentSummary.averageAge}
                previous={previousSummary.averageAge}
                onClick={() => setAgeFilter((current) => current === '>30 días' ? 'Todos' : '>30 días')}
                clickable
              />
            </section>

            <section className="chart-grid chart-grid-main">
              <ChartCard
                className="chart-wide"
                title="Evolución del pendiente"
                subtitle={
                  daily.length > 1
                    ? `${formatCutoff(from)} — ${formatCutoff(to)} · Toca cualquier punto para cambiar la fecha`
                    : 'Corte diario consolidado · Se construirá la curva histórica con cada corte'
                }
              >
                <ResponsiveContainer width="100%" height={320}>
                  <AreaChart
                    data={daily}
                    margin={{ top: 18, right: 8, left: 0, bottom: 0 }}
                    onClick={(state: any) => {
                      const payloadCutoff = state?.activePayload?.[0]?.payload?.cutoff;
                      if (payloadCutoff && payloadCutoff !== cutoff) {
                        setCutoff(String(payloadCutoff));
                      }
                    }}
                  >
                    <defs>
                      <linearGradient id="pendingFill" x1="0" y1="0" x2="0" y2="1">
                        <stop offset="0%" stopColor="#0071e3" stopOpacity={0.28} />
                        <stop offset="100%" stopColor="#0071e3" stopOpacity={0.02} />
                      </linearGradient>
                    </defs>
                    <CartesianGrid stroke="#e8e8ed" vertical={false} />
                    <XAxis
                      dataKey="cutoff"
                      tickFormatter={(value) => formatCutoff(value, { day: '2-digit', month: 'short', year: undefined })}
                      tickLine={false}
                      axisLine={false}
                    />
                    <YAxis tickFormatter={(value) => compactCurrency.format(value)} tickLine={false} axisLine={false} width={74} />
                    <Tooltip content={<CurrencyTooltip />} />
                    <Area type="monotone" dataKey="pending" name="Pendiente" stroke="#0071e3" strokeWidth={3} fill="url(#pendingFill)" />
                  </AreaChart>
                </ResponsiveContainer>
                <div className="chart-footer-stats">
                  <span><i className="dot blue" />Último saldo <strong>{compactCurrency.format(daily.at(-1)?.pending || 0)}</strong></span>
                  <span><i className="dot green" />Retirado en periodo <strong>{compactCurrency.format(daily.reduce((sum, point) => sum + point.withdrawn, 0))}</strong></span>
                  <span><i className="dot orange" />Nuevas en periodo <strong>{compactCurrency.format(daily.slice(1).reduce((sum, point) => sum + point.newValue, 0))}</strong></span>
                </div>
              </ChartCard>

              <AgeCompositionCard
                ageData={ageBreakdown}
                totalPending={currentSummary.pending}
                activeRange={ageFilter}
                onSelectRange={(range) => {
                  setAgeFilter((current) => current === range ? 'Todos' : range);
                }}
              />
            </section>

            <section className="chart-grid">
              <ChartCard
                title="Pendiente por dirección"
                subtitle={director !== 'Todos' ? `Filtrado por: ${director} · Toca para quitar` : "Toca una dirección para filtrar todo el tablero"}
              >
                <div className="donut-layout">
                  <ResponsiveContainer width="48%" height={260}>
                    <PieChart>
                      <Pie
                        data={directorData}
                        dataKey="value"
                        nameKey="name"
                        innerRadius={64}
                        outerRadius={94}
                        paddingAngle={2}
                        stroke="none"
                        cursor="pointer"
                        onClick={(entry: any) => {
                          if (entry && entry.name) {
                            const dirName = String(entry.name);
                            setDirector((current) => current === dirName ? 'Todos' : dirName);
                          }
                        }}
                      >
                        {directorData.map((entry, index) => {
                          const isSelected = director === entry.name;
                          const isDimmed = director !== 'Todos' && !isSelected;
                          return (
                            <Cell
                              key={entry.name}
                              fill={PIE_COLORS[index % PIE_COLORS.length]}
                              opacity={isDimmed ? 0.35 : 1}
                              stroke={isSelected ? '#1d1d1f' : 'none'}
                              strokeWidth={isSelected ? 2 : 0}
                            />
                          );
                        })}
                      </Pie>
                      <Tooltip formatter={(value) => currency.format(Number(value))} />
                    </PieChart>
                  </ResponsiveContainer>
                  <div className="legend-list">
                    {directorData.map((entry, index) => {
                      const isSelected = director === entry.name;
                      return (
                        <button
                          key={entry.name}
                          type="button"
                          className={`legend-btn ${isSelected ? 'active' : ''}`}
                          onClick={() => setDirector((current) => current === entry.name ? 'Todos' : entry.name)}
                          title={isSelected ? `Toca para quitar filtro de ${entry.name}` : `Toca para filtrar por ${entry.name}`}
                        >
                          <i style={{ background: PIE_COLORS[index % PIE_COLORS.length] }} />
                          <span>{entry.name}<small>{entry.count} remisiones</small></span>
                          <strong>{compactCurrency.format(entry.value)}</strong>
                        </button>
                      );
                    })}
                  </div>
                </div>
              </ChartCard>
              <ChartCard
                title="Ejecutivos con mayor pendiente"
                subtitle={employee !== 'Todos' ? `Filtrado por: ${employee} · Toca para quitar` : "Toca una barra para filtrar por comercial"}
              >
                <ResponsiveContainer width="100%" height={300}>
                  <BarChart
                    data={sellerData}
                    margin={{ top: 8, right: 8, left: 0, bottom: 45 }}
                    onClick={(state: any) => {
                      const executiveName = state?.activePayload?.[0]?.payload?.name;
                      if (executiveName) {
                        const nameStr = String(executiveName);
                        setEmployee((current) => current === nameStr ? 'Todos' : nameStr);
                      }
                    }}
                  >
                    <CartesianGrid stroke="#e8e8ed" vertical={false} />
                    <XAxis
                      dataKey="name"
                      interval={0}
                      angle={-32}
                      textAnchor="end"
                      height={80}
                      tickFormatter={(value) => String(value).split(' ').slice(0, 2).join(' ')}
                      tickLine={false}
                      axisLine={false}
                      cursor="pointer"
                    />
                    <YAxis tickFormatter={(value) => compactCurrency.format(value)} tickLine={false} axisLine={false} width={72} />
                    <Tooltip content={<CurrencyTooltip />} />
                    <Bar dataKey="value" name="Pendiente" radius={[7, 7, 0, 0]} maxBarSize={34} cursor="pointer">
                      {sellerData.map((entry) => {
                        const isSelected = employee === entry.name;
                        const isDimmed = employee !== 'Todos' && !isSelected;
                        return (
                          <Cell
                            key={entry.name}
                            fill={isSelected ? '#7928ca' : '#af52de'}
                            opacity={isDimmed ? 0.35 : 1}
                          />
                        );
                      })}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </ChartCard>
            </section>

            <section className="daily-strip">
              <div className="daily-strip-title"><CalendarRange size={20} /><div><strong>Gestión del corte</strong><span>Entradas y retiros frente al corte anterior</span></div></div>
              <DailyStat label="Saldo anterior" value={currency.format(daily.at(-1)?.previousBalance || 0)} />
              <DailyStat label="Nuevas" value={currency.format(daily.at(-1)?.newValue || 0)} tone="orange" />
              <DailyStat label="Retirado" value={currency.format(daily.at(-1)?.withdrawn || 0)} tone="green" />
              <DailyStat label="Reducción bruta" value={percent.format(daily.at(-1)?.grossReduction || 0)} tone="blue" />
              <button className="button button-secondary" onClick={() => setView('detail')}>Ver detalle</button>
            </section>
          </>
        ) : (
          <section className="detail-card">
            <div className="detail-header">
              <div>
                <h2>Detalle de remisiones abiertas</h2>
                <p>
                  {number.format(detailRecords.length)} registros
                  {statusFilter !== 'Todos' ? ` · Estado: ${statusFilter}` : ''}
                  {ageFilter !== 'Todos' ? ` · Días: ${ageFilter}` : ''}
                  {amountFilter !== 'Todos' ? ` · Monto: ${amountFilter}` : ''}
                  {director !== 'Todos' ? ` · Director: ${director}` : ''}
                  {employee !== 'Todos' ? ` · Comercial: ${employee}` : ''}
                  {` · ${formattedCutoffWithTime}`}
                </p>
              </div>
              <div className="detail-actions">
                <label className="search-box">
                  <Search size={17} />
                  <input
                    value={query}
                    onChange={(event) => setQuery(event.target.value)}
                    placeholder="Cliente, NIT, remisión, pedido…"
                  />
                </label>
                <SelectFilter
                  label="Estado"
                  value={statusFilter}
                  onChange={setStatusFilter}
                  options={[
                    'Todos',
                    'Vencida · Alto valor',
                    'Vencida',
                    'Por vencer · Alto valor',
                    'Por vencer',
                    'Al día · Alto valor',
                    'Al día',
                  ]}
                  compact
                />
                <SelectFilter
                  label="Días"
                  value={ageFilter}
                  onChange={setAgeFilter}
                  options={[
                    'Todos',
                    'Vencidas (>30 días)',
                    'Por vencer (16-30 días)',
                    'Al día (0-15 días)',
                    'Críticas (>60 días)',
                  ]}
                  compact
                />
                <SelectFilter
                  label="Monto"
                  value={amountFilter}
                  onChange={setAmountFilter}
                  options={[
                    'Todos',
                    'Alto valor (> $5M)',
                    'Valor medio ($1M - $5M)',
                    'Menor valor (< $1M)',
                  ]}
                  compact
                />
                <SelectFilter
                  label="Ordenar"
                  value={sortBy}
                  onChange={(val) => setSortBy(val as 'total-desc' | 'total-asc' | 'age-desc' | 'age-asc')}
                  options={['total-desc', 'total-asc', 'age-desc', 'age-asc']}
                  format={(val) => {
                    if (val === 'total-desc') return 'Mayor valor ($)';
                    if (val === 'total-asc') return 'Menor valor ($)';
                    if (val === 'age-desc') return 'Más días (Antiguas)';
                    return 'Menos días (Recientes)';
                  }}
                  compact
                />
                <button className="button button-secondary" onClick={exportCsv}>
                  <Download size={17} /> Exportar CSV
                </button>
              </div>
            </div>
            <div className="table-wrap">
              <table>
                <thead>
                  <tr>
                    <th>Cliente</th>
                    <th>Director</th>
                    <th>Comercial</th>
                    <th>Remisión</th>
                    <th>Pedido</th>
                    <th>Emisión</th>
                    <th className="numeric">Días</th>
                    <th className="numeric">Total</th>
                    <th>Estado</th>
                  </tr>
                </thead>
                <tbody>
                  {visibleRows.map((record) => (
                    <DetailRow
                      key={record.id}
                      record={record}
                      onSelectDirector={(dir) => setDirector((current) => current === dir ? 'Todos' : dir)}
                      onSelectEmployee={(emp) => setEmployee((current) => current === emp ? 'Todos' : emp)}
                      onSelectAge={(age) => setAgeFilter((current) => current === age ? 'Todos' : age)}
                      onSelectAmount={(amt) => setAmountFilter((current) => current === amt ? 'Todos' : amt)}
                      onSelectStatus={(st) => setStatusFilter((current) => current === st ? 'Todos' : st)}
                      onSelectCompany={(company) => setQuery((current) => current === company ? '' : company)}
                    />
                  ))}
                  {!visibleRows.length && (
                    <tr>
                      <td colSpan={9}>
                        <div className="empty-state">
                          <Search size={22} />
                          No hay resultados para estos filtros.
                        </div>
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
            <div className="pagination">
              <span>Mostrando {visibleRows.length ? (page - 1) * pageSize + 1 : 0}–{Math.min(page * pageSize, detailRecords.length)} de {detailRecords.length} remisiones</span>
              <div>
                <button disabled={page <= 1} onClick={() => setPage((value) => value - 1)}>Anterior</button>
                <span>Página {page} de {pageCount}</span>
                <button disabled={page >= pageCount} onClick={() => setPage((value) => value + 1)}>Siguiente</button>
              </div>
            </div>
          </section>
        )}

        <footer className="app-footer">
          <span>Provexpress · Control de remisiones abiertas</span>
          <span>
            {metadata?.lastModifiedDateTime
              ? `Archivo actualizado: ${formatDateTime(metadata.lastModifiedDateTime)}`
              : data.cutoffDateTime
                ? `Corte: ${formatDateTime(data.cutoffDateTime)}`
                : ''}
          </span>
        </footer>
      </main>
    </div>
  );
}

function SelectFilter({
  label,
  value,
  options,
  onChange,
  format,
  compact = false,
}: {
  label: string;
  value: string;
  options: string[];
  onChange: (value: string) => void;
  format?: (value: string) => string;
  compact?: boolean;
}) {
  return (
    <label className={`select-filter ${compact ? 'compact' : ''}`}>
      <span>{label}</span>
      <div>
        <select value={value} onChange={(event) => onChange(event.target.value)}>
          {options.map((option) => (
            <option key={option} value={option}>
              {format ? format(option) : option}
            </option>
          ))}
        </select>
        <ChevronDown size={14} />
      </div>
    </label>
  );
}

function MetricCard({
  title,
  value,
  sub,
  icon,
  tone,
  current,
  previous,
  onClick,
  clickable = false,
}: {
  title: string;
  value: string;
  sub?: string;
  icon: React.ReactNode;
  tone: string;
  current: number;
  previous: number;
  onClick?: () => void;
  clickable?: boolean;
}) {
  const delta = previous ? (current - previous) / Math.abs(previous) : 0;
  const isUp = delta > 0;
  return (
    <article
      className={`metric-card ${tone} ${clickable ? 'clickable' : ''}`}
      onClick={onClick}
      role={clickable ? 'button' : undefined}
      tabIndex={clickable ? 0 : undefined}
    >
      <div className={`metric-icon ${tone}`}>{icon}</div>
      <span className="metric-title">{title}</span>
      <strong className="metric-value">{value}</strong>
      <div className="metric-caption">
        {sub ? <span>{sub}</span> : <span />}
        {previous > 0 && (
          <span className={isUp ? 'delta-up' : 'delta-down'}>
            {isUp ? <ArrowUpRight size={14} /> : <ArrowDownRight size={14} />}
            {percent.format(Math.abs(delta))}
          </span>
        )}
      </div>
    </article>
  );
}

function AgeCompositionCard({
  ageData,
  totalPending,
  activeRange = 'Todos',
  onSelectRange,
}: {
  ageData: AgeBreakdownItem[];
  totalPending: number;
  activeRange?: string;
  onSelectRange: (range: string) => void;
}) {
  const overdueItems = ageData.filter((item) => item.name === '31-60 días' || item.name === '>60 días');
  const overdueTotal = overdueItems.reduce((sum, item) => sum + item.value, 0);
  const overdueCount = overdueItems.reduce((sum, item) => sum + item.count, 0);
  const overduePercent = totalPending > 0 ? (overdueTotal / totalPending) * 100 : 0;
  const maxVal = Math.max(...ageData.map((d) => d.value), 1);
  const isOverdueActive = activeRange === '>30 días' || activeRange === 'Vencidas (>30 días)';

  return (
    <article className="chart-card age-composition-card">
      <header className="age-card-header">
        <div>
          <h2>Composición por antigüedad</h2>
          <p>
            {activeRange !== 'Todos'
              ? `Filtrado por: ${activeRange} · Toca para quitar`
              : 'Toca un rango para filtrar todo el tablero'}
          </p>
        </div>
        <button
          type="button"
          className={`overdue-highlight-chip ${isOverdueActive ? 'active' : ''}`}
          onClick={() => onSelectRange('>30 días')}
          title={isOverdueActive ? 'Toca para quitar filtro de vencidas' : 'Toca para filtrar remisiones vencidas (>30 días)'}
        >
          <TriangleAlert size={16} />
          <div>
            <strong>Vencido &gt;30d: {currency.format(overdueTotal)}</strong>
            <small>{overdueCount} remisiones · {overduePercent.toFixed(1)}%</small>
          </div>
        </button>
      </header>

      <div className="age-bars-container">
        {ageData.map((item) => {
          const isSelected = activeRange === item.name || (activeRange === '>30 días' && (item.name === '31-60 días' || item.name === '>60 días'));
          const isDimmed = activeRange !== 'Todos' && !isSelected;
          const barWidthPercent = Math.max(5, (item.value / maxVal) * 100);
          return (
            <div
              key={item.name}
              className={`age-bar-row tone-${item.tone} ${isSelected ? 'active' : ''} ${isDimmed ? 'dimmed' : ''}`}
              onClick={() => onSelectRange(item.name)}
              role="button"
              tabIndex={0}
              title={isSelected ? `Toca para quitar filtro de ${item.name}` : `Toca para filtrar remisiones de ${item.name}`}
            >
              <div className="age-bar-label">
                <strong>{item.name}</strong>
                <span className={`age-pill ${item.tone}`}>{item.badge}</span>
              </div>
              <div className="age-bar-track">
                <div className={`age-bar-fill ${item.tone}`} style={{ width: `${barWidthPercent}%` }}>
                  <span className="age-bar-inline-val">{compactCurrency.format(item.value)}</span>
                </div>
              </div>
              <div className="age-bar-meta">
                <strong>{currency.format(item.value)}</strong>
                <span>{item.count} rem. · <b>{item.percent.toFixed(1)}%</b></span>
              </div>
            </div>
          );
        })}
      </div>

      <div className="age-card-footer">
        <span className="legend-item blue"><i /> Al día (0 a 15 días)</span>
        <span className="legend-item orange"><i /> Por vencer (16 a 30 días)</span>
        <span className="legend-item red"><i /> Vencido crítico (&gt;30 días)</span>
      </div>
    </article>
  );
}

function ChartCard({ title, subtitle, children, className = '' }: { title: string; subtitle: string; children: React.ReactNode; className?: string }) {
  return <article className={`chart-card ${className}`}><header><div><h2>{title}</h2><p>{subtitle}</p></div></header>{children}</article>;
}

function DailyStat({ label, value, tone = '' }: { label: string; value: string; tone?: string }) {
  return <div className={`daily-stat ${tone}`}><span>{label}</span><strong>{value}</strong></div>;
}

function DetailRow({
  record,
  onSelectDirector,
  onSelectEmployee,
  onSelectAge,
  onSelectAmount,
  onSelectStatus,
  onSelectCompany,
}: {
  record: Remision;
  onSelectDirector?: (director: string) => void;
  onSelectEmployee?: (employee: string) => void;
  onSelectAge?: (age: string) => void;
  onSelectAmount?: (amount: string) => void;
  onSelectStatus?: (status: string) => void;
  onSelectCompany?: (company: string) => void;
}) {
  const isOverdue = record.age > 30;
  const isPriority = record.age > 15 && record.age <= 30;
  return (
    <tr>
      <td>
        <button
          type="button"
          className="cell-touch-btn text-left"
          onClick={() => onSelectCompany?.(record.company)}
          title={`Toca para buscar cliente "${record.company}"`}
        >
          <strong className="cell-company">{record.company || 'Sin empresa'}</strong>
          <small className="cell-nit">NIT {record.nit || '—'}</small>
        </button>
      </td>
      <td>
        <button
          type="button"
          className="cell-touch-btn text-left"
          onClick={() => onSelectDirector?.(record.director)}
          title={`Toca para filtrar por director "${record.director}"`}
        >
          <span className="cell-director">{record.director || 'Sin asignar'}</span>
        </button>
      </td>
      <td>
        <button
          type="button"
          className="cell-touch-btn text-left"
          onClick={() => onSelectEmployee?.(record.employee)}
          title={`Toca para filtrar por comercial "${record.employee}"`}
        >
          <span className="cell-employee">{record.employee}</span>
        </button>
      </td>
      <td>
        <strong className="cell-doc">{record.document || '—'}</strong>
      </td>
      <td>
        <span className="cell-order">{record.order || '—'}</span>
      </td>
      <td>{formatCutoff(record.issuedAt)}</td>
      <td className="numeric">
        <button
          type="button"
          className="cell-touch-btn text-right"
          onClick={() => onSelectAge?.(record.daysStatus)}
          title={`Toca para filtrar por rango "${record.daysStatus}"`}
        >
          <span className={`days-chip ${isOverdue ? 'danger' : isPriority ? 'warning' : 'ok'}`}>
            {record.age} d
          </span>
          <small className="days-range-sub">{record.daysStatus}</small>
        </button>
      </td>
      <td className="numeric">
        <button
          type="button"
          className="cell-touch-btn text-right"
          onClick={() => onSelectAmount?.(record.amountStatus)}
          title={`Toca para filtrar por monto "${record.amountStatus}"`}
        >
          <strong className="cell-total">{currency.format(record.total)}</strong>
          <span className={`amount-chip ${record.total >= 5_000_000 ? 'high' : 'standard'}`}>
            {record.amountStatus}
          </span>
        </button>
      </td>
      <td>
        <button
          type="button"
          className="cell-touch-btn"
          onClick={() => onSelectStatus?.(record.alert)}
          title={`Toca para filtrar por estado "${record.alert}"`}
        >
          <span className={`status-pill ${statusClass(record.alert)}`}>{record.alert}</span>
        </button>
      </td>
    </tr>
  );
}


function CurrencyTooltip({ active, payload, label }: { active?: boolean; payload?: Array<{ value?: number; name?: string; color?: string }>; label?: string }) {
  if (!active || !payload?.length) return null;
  return <div className="chart-tooltip"><strong>{label ? (String(label).match(/^\d{4}-/) ? formatCutoff(label) : label) : ''}</strong>{payload.map((item) => <span key={item.name}><i style={{ background: item.color }} />{item.name}: <b>{currency.format(Number(item.value || 0))}</b></span>)}</div>;
}

function initials(name: string): string {
  return name.split(/\s+/).filter(Boolean).slice(0, 2).map((part) => part[0]).join('').toUpperCase();
}

function statusClass(alert: Remision['alert']): string {
  if (alert === 'Vencida · Alto valor') return 'overdue-high';
  if (alert === 'Vencida') return 'overdue';
  if (alert === 'Por vencer · Alto valor') return 'priority-high';
  if (alert === 'Por vencer') return 'priority';
  if (alert === 'Al día · Alto valor') return 'normal-high';
  return 'normal';
}

export default App;
