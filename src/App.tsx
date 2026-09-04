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
  PlusCircle,
  RefreshCw,
  RotateCcw,
  Search,
  ShieldCheck,
  TrendingUp,
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
import type { AgeBreakdownItem, DailyPoint, DataSource, FileMetadata, ParsedWorkbook, Remision, Summary, UserProfile } from './types';

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
            <div className="preview-change">Seguimiento diario consolidado</div>
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
  const [detailTab, setDetailTab] = useState<'open' | 'withdrawn' | 'new'>('open');
  const latestCutoff = data.cutoffs.at(-1) || '';
  const [cutoff, setCutoff] = useState(latestCutoff);
  const [director, setDirector] = useState('Todos');
  const [employee, setEmployee] = useState('Todos');
  const [statusFilter, setStatusFilter] = useState('Todos');
  const [ageFilter, setAgeFilter] = useState('Todos');
  const [amountFilter, setAmountFilter] = useState('Todos');
  const [sortBy, setSortBy] = useState<'total-desc' | 'total-asc' | 'age-desc' | 'age-asc'>('total-desc');
  const [query, setQuery] = useState('');
  const [page, setPage] = useState(1);
  const refreshRef = useRef(onRefresh);

  useEffect(() => {
    if (data.cutoffs.length && (!cutoff || !data.cutoffs.includes(cutoff))) {
      setCutoff(latestCutoff);
    }
  }, [data.cutoffs, cutoff, latestCutoff]);

  const matchesAgeFilter = (age: number, ageRange: string, filter: string): boolean => {
    if (filter === 'Todos') return true;
    if (filter === '>30 días' || filter === 'Crítico (>30 días)' || filter === 'Crítico' || filter === 'Vencidas (>30 días)') return age > 30;
    if (filter === 'Gestión comercial (16-30 días)' || filter === 'Gestión comercial' || filter === 'Por vencer (16-30 días)' || filter === '16-30 días') return age > 15 && age <= 30;
    if (filter === 'Al día (0-15 días)' || filter === 'Al día' || filter === '0-15 días') return age <= 15;
    if (filter === 'Crítico (>60 días)' || filter === 'Críticas (>60 días)' || filter === '>60 días') return age > 60;
    if (filter === 'Crítico (31-60 días)' || filter === 'Vencida (31-60 días)' || filter === '31-60 días') return age > 30 && age <= 60;
    return ageRange === filter;
  };

  const matchesAmountFilter = (amountStatus: string, filter: string): boolean => {
    if (filter === 'Todos') return true;
    if (filter === 'Alto valor (> $5M)') return amountStatus === 'Alto valor (> $5M)';
    if (filter === 'Cuantía media ($1M - $5M)' || filter === 'Valor medio ($1M - $5M)') {
      return amountStatus === 'Cuantía media ($1M - $5M)' || amountStatus === 'Valor medio ($1M - $5M)';
    }
    if (filter === 'Menor cuantía (< $1M)' || filter === 'Menor valor (< $1M)') {
      return amountStatus === 'Menor cuantía (< $1M)' || amountStatus === 'Menor valor (< $1M)';
    }
    return amountStatus === filter;
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
      if (!matchesAmountFilter(record.amountStatus, amountFilter)) return false;
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
      matchesAmountFilter(record.amountStatus, amountFilter) &&
      matchesAgeFilter(record.age, record.ageRange, ageFilter)
    )),
    [data.records, previousCutoff, director, employee, statusFilter, amountFilter, ageFilter],
  );

  const periodRecords = useMemo(
    () => data.records.filter((record) =>
      (director === 'Todos' || record.director === director) &&
      (employee === 'Todos' || record.employee === employee)),
    [data.records, director, employee],
  );
  const daily = useMemo(() => buildDailySeries(periodRecords), [periodRecords]);
  const currentDailyPoint = useMemo(
    () => daily.find((d) => d.cutoff === cutoff) || daily.at(-1),
    [daily, cutoff],
  );

  const previousCutoffRecords = useMemo(
    () => data.records.filter((record) => record.cutoff === previousCutoff),
    [data.records, previousCutoff],
  );
  const currentKeysSet = useMemo(
    () => new Set(baseCutoffRecords.map((record) => record.stableKey)),
    [baseCutoffRecords],
  );
  const previousKeysSet = useMemo(
    () => new Set(previousCutoffRecords.map((record) => record.stableKey)),
    [previousCutoffRecords],
  );

  const withdrawnRecords = useMemo(() => {
    if (!previousCutoff) return [];
    return previousCutoffRecords.filter((record) => {
      if (currentKeysSet.has(record.stableKey)) return false;
      if (director !== 'Todos' && record.director !== director) return false;
      if (employee !== 'Todos' && record.employee !== employee) return false;
      if (statusFilter !== 'Todos' && record.alert !== statusFilter) return false;
      if (!matchesAmountFilter(record.amountStatus, amountFilter)) return false;
      if (!matchesAgeFilter(record.age, record.ageRange, ageFilter)) return false;
      return true;
    });
  }, [previousCutoff, previousCutoffRecords, currentKeysSet, director, employee, statusFilter, amountFilter, ageFilter]);

  const newRecords = useMemo(() => {
    if (!previousCutoff) return [];
    return currentRecords.filter((record) => !previousKeysSet.has(record.stableKey));
  }, [currentRecords, previousCutoff, previousKeysSet]);

  const managementWithdrawn = useMemo(() => {
    if (!previousCutoff) return [];
    return previousCutoffRecords.filter((record) => {
      if (currentKeysSet.has(record.stableKey)) return false;
      if (director !== 'Todos' && record.director !== director) return false;
      if (statusFilter !== 'Todos' && record.alert !== statusFilter) return false;
      if (!matchesAmountFilter(record.amountStatus, amountFilter)) return false;
      if (!matchesAgeFilter(record.age, record.ageRange, ageFilter)) return false;
      return true;
    });
  }, [previousCutoff, previousCutoffRecords, currentKeysSet, director, statusFilter, amountFilter, ageFilter]);

  const managementNew = useMemo(() => {
    if (!previousCutoff) return [];
    return baseCutoffRecords.filter((record) => {
      if (previousKeysSet.has(record.stableKey)) return false;
      if (director !== 'Todos' && record.director !== director) return false;
      if (statusFilter !== 'Todos' && record.alert !== statusFilter) return false;
      if (!matchesAmountFilter(record.amountStatus, amountFilter)) return false;
      if (!matchesAgeFilter(record.age, record.ageRange, ageFilter)) return false;
      return true;
    });
  }, [baseCutoffRecords, previousCutoff, previousKeysSet, director, statusFilter, amountFilter, ageFilter]);

  const closedBySeller = useMemo(() => {
    const map = new Map<string, { name: string; director: string; total: number; count: number }>();
    for (const r of managementWithdrawn) {
      const cur = map.get(r.employee) || { name: r.employee, director: r.director, total: 0, count: 0 };
      cur.total += r.total;
      cur.count += 1;
      map.set(r.employee, cur);
    }
    return [...map.values()].sort((a, b) => b.total - a.total);
  }, [managementWithdrawn]);

  const newBySeller = useMemo(() => {
    const map = new Map<string, { name: string; director: string; total: number; count: number }>();
    for (const r of managementNew) {
      const cur = map.get(r.employee) || { name: r.employee, director: r.director, total: 0, count: 0 };
      cur.total += r.total;
      cur.count += 1;
      map.set(r.employee, cur);
    }
    return [...map.values()].sort((a, b) => b.total - a.total);
  }, [managementNew]);

  const totalClosed = useMemo(() => managementWithdrawn.reduce((sum, r) => sum + r.total, 0), [managementWithdrawn]);
  const totalNew = useMemo(() => managementNew.reduce((sum, r) => sum + r.total, 0), [managementNew]);
  const netDifference = useMemo(() => totalNew - totalClosed, [totalNew, totalClosed]);

  const ageBreakdownRecords = useMemo(
    () => baseCutoffRecords.filter((r) =>
      (director === 'Todos' || r.director === director) &&
      (employee === 'Todos' || r.employee === employee) &&
      (statusFilter === 'Todos' || r.alert === statusFilter) &&
      matchesAmountFilter(r.amountStatus, amountFilter)),
    [baseCutoffRecords, director, employee, statusFilter, amountFilter],
  );
  const ageBreakdown = useMemo(() => buildAgeBreakdown(ageBreakdownRecords), [ageBreakdownRecords]);

  const directorFilterRecords = useMemo(
    () => baseCutoffRecords.filter((r) =>
      (statusFilter === 'Todos' || r.alert === statusFilter) &&
      matchesAmountFilter(r.amountStatus, amountFilter) &&
      matchesAgeFilter(r.age, r.ageRange, ageFilter)),
    [baseCutoffRecords, statusFilter, amountFilter, ageFilter],
  );
  const directorData = useMemo(() => aggregateBy(directorFilterRecords, (record) => record.director), [directorFilterRecords]);

  const sellerFilterRecords = useMemo(
    () => baseCutoffRecords.filter((r) =>
      (director === 'Todos' || r.director === director) &&
      (statusFilter === 'Todos' || r.alert === statusFilter) &&
      matchesAmountFilter(r.amountStatus, amountFilter) &&
      matchesAgeFilter(r.age, r.ageRange, ageFilter)),
    [baseCutoffRecords, director, statusFilter, amountFilter, ageFilter],
  );
  const sellerData = useMemo(() => aggregateBy(sellerFilterRecords, (record) => record.employee).slice(0, 12), [sellerFilterRecords]);

  const targetRecords = useMemo(() => {
    if (detailTab === 'withdrawn') return withdrawnRecords;
    if (detailTab === 'new') return newRecords;
    return currentRecords;
  }, [detailTab, withdrawnRecords, newRecords, currentRecords]);
  const detailRecords = useMemo(() => {
    const normalizedQuery = normalizeText(query);
    return targetRecords
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
  }, [targetRecords, sortBy, query]);

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
      'Fecha',
      'Hora',
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
              <span className="hero-time-tag"> · Fecha: <b>{formatCutoff(cutoff)}</b>{data.cutoffTimeDisplay ? ` (${data.cutoffTimeDisplay})` : ''}</span>
            </p>
          </div>
          <div className="source-card">
            <span className={`source-icon ${source}`}><FileSpreadsheet size={21} /></span>
            <div>
              <strong>{metadata?.name || 'Remisiones.xlsx'}</strong>
              <small>
                {source === 'sharepoint' ? 'SharePoint' : 'Archivo local'}
                {data.cutoffTimeDisplay ? ` · Actualizado ${data.cutoffTimeDisplay}` : ''}
                {data.activeSheetName ? ` · Hoja ${data.activeSheetName}` : ''}
              </small>
            </div>
            <span className="source-status"><i /> Sincronizado</span>
          </div>
        </section>

        <section className="filter-bar" aria-label="Filtros del tablero">
          <SelectFilter label="Director" value={director} onChange={setDirector} options={['Todos', ...directors]} />
          <SelectFilter label="Ejecutivo" value={employee} onChange={setEmployee} options={['Todos', ...employees]} />
          <div className="filter-date-badge">
            <span className="filter-date-label">Fecha activa:</span>
            <strong className="filter-date-val">{formatCutoff(cutoff)}</strong>
            {cutoff === latestCutoff ? (
              <span className="filter-status-tag current">Al día</span>
            ) : (
              <button
                type="button"
                className="filter-return-today-btn"
                onClick={() => setCutoff(latestCutoff)}
                title="Volver a la fecha actual más reciente"
              >
                <RotateCcw size={12} /> Volver a hoy ({formatCutoff(latestCutoff)})
              </button>
            )}
          </div>
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
            {/* 1. MÓDULO DE GESTIÓN OPERATIVA (ENTRADAS VS SALIDAS DEL DÍA) */}
            <section className="management-overview-card">
              <div className="management-card-header">
                <div className="management-card-title">
                  <div className="management-icon-badge">
                    <TrendingUp size={22} />
                  </div>
                  <div>
                    <div className="management-eyebrow">
                      Proceso de Gestión · {formattedCutoffWithTime}
                    </div>
                    <h2>Gestión Operativa: Cartera, Ingresos y Facturación</h2>
                    <p>
                      {previousCutoff ? (
                        <>
                          Evaluación de flujo frente al día previo del <strong>{formatCutoff(previousCutoff)}</strong>
                          {director !== 'Todos' ? ` · Dirección: ${director}` : ''}
                          {employee !== 'Todos' ? ` · Comercial: ${employee}` : ''}
                        </>
                      ) : (
                        'Estado actual consolidado de la primera fecha registrada'
                      )}
                    </p>
                  </div>
                </div>
                {previousCutoff && (
                  <div className="management-header-actions">
                    <button
                      type="button"
                      className="button button-outline-green"
                      onClick={() => { setView('detail'); setDetailTab('withdrawn'); }}
                      title="Ver detalle de remisiones facturadas o cobradas"
                    >
                      <CheckCircle2 size={15} />
                      Ver facturadas ({number.format(managementWithdrawn.length)})
                    </button>
                    <button
                      type="button"
                      className="button button-outline-orange"
                      onClick={() => { setView('detail'); setDetailTab('new'); }}
                      title="Ver detalle de nuevas remisiones abiertas"
                    >
                      <PlusCircle size={15} />
                      Ver nuevas ({number.format(managementNew.length)})
                    </button>
                  </div>
                )}
              </div>

              <div className="management-kpis-grid">
                {/* 1. Cartera General */}
                <div
                  className="mgmt-kpi-item"
                  role="button"
                  tabIndex={0}
                  onClick={() => { setView('detail'); setDetailTab('open'); }}
                  title="Toca para ver el detalle de todas las remisiones abiertas"
                >
                  <div className="mgmt-kpi-top">
                    <span className="mgmt-kpi-label">Cartera General (Saldo Activo)</span>
                    <span className="mgmt-kpi-badge blue">Día Actual</span>
                  </div>
                  <strong className="mgmt-kpi-value">{currency.format(currentSummary.pending)}</strong>
                  <div className="mgmt-kpi-foot">
                    <CircleDollarSign size={15} />
                    <span><strong>{number.format(currentSummary.remissions)}</strong> remisiones · <strong>{number.format(currentSummary.clients)}</strong> clientes</span>
                  </div>
                </div>

                {/* 2. Ingreso de Nuevas */}
                <div
                  className="mgmt-kpi-item warning"
                  role={previousCutoff ? "button" : undefined}
                  tabIndex={previousCutoff ? 0 : undefined}
                  onClick={previousCutoff ? () => { setView('detail'); setDetailTab('new'); } : undefined}
                  title="Toca para ver el detalle de nuevas remisiones abiertas"
                >
                  <div className="mgmt-kpi-top">
                    <span className="mgmt-kpi-label">Nuevas Remisiones (Ingresos)</span>
                    <span className="mgmt-kpi-badge orange">Entradas</span>
                  </div>
                  <strong className="mgmt-kpi-value text-orange">
                    {previousCutoff ? `+${currency.format(totalNew)}` : currency.format(currentSummary.pending)}
                  </strong>
                  <div className="mgmt-kpi-foot">
                    <ArrowUpRight size={15} />
                    <span>
                      {previousCutoff
                        ? <span><strong>{number.format(managementNew.length)}</strong> remisiones nuevas registradas</span>
                        : 'Punto de partida de cartera'}
                    </span>
                  </div>
                </div>

                {/* 3. Facturadas / Retiradas */}
                <div
                  className="mgmt-kpi-item success"
                  role={previousCutoff ? "button" : undefined}
                  tabIndex={previousCutoff ? 0 : undefined}
                  onClick={previousCutoff ? () => { setView('detail'); setDetailTab('withdrawn'); } : undefined}
                  title="Toca para ver el detalle de remisiones facturadas o cobradas"
                >
                  <div className="mgmt-kpi-top">
                    <span className="mgmt-kpi-label">Remisiones Facturadas (Salidas)</span>
                    <span className="mgmt-kpi-badge green">Retiradas</span>
                  </div>
                  <strong className="mgmt-kpi-value text-green">
                    {previousCutoff ? `-${currency.format(totalClosed)}` : '$ 0'}
                  </strong>
                  <div className="mgmt-kpi-foot">
                    <ArrowDownRight size={15} />
                    <span>
                      {previousCutoff
                        ? <span><strong>{number.format(managementWithdrawn.length)}</strong> remisiones salieron de cartera</span>
                        : 'Disponible con dos o más días registrados'}
                    </span>
                  </div>
                </div>

                {/* 4. Balance Neto */}
                <div className={`mgmt-kpi-item ${netDifference <= 0 ? 'favorable' : 'neutral'}`}>
                  <div className="mgmt-kpi-top">
                    <span className="mgmt-kpi-label">Variación Neta de Cartera</span>
                    <span className={`mgmt-kpi-badge ${netDifference <= 0 ? 'blue' : 'gray'}`}>
                      {previousCutoff ? (netDifference <= 0 ? 'Favorable' : 'Incremento') : 'Base inicial'}
                    </span>
                  </div>
                  <strong className="mgmt-kpi-value">
                    {previousCutoff ? (
                      `${netDifference <= 0 ? '-' : '+'}${currency.format(Math.abs(netDifference))}`
                    ) : (
                      '—'
                    )}
                  </strong>
                  <div className="mgmt-kpi-foot">
                    <span>
                      {previousCutoff ? (
                        netDifference <= 0
                          ? `Desahogo neto (${percent.format(currentDailyPoint?.grossReduction || 0)} reducción bruta)`
                          : 'Crecimiento de saldo frente al día anterior'
                      ) : (
                        `Saldo inicial con ${number.format(currentSummary.remissions)} remisiones`
                      )}
                    </span>
                  </div>
                </div>
              </div>

              {previousCutoff ? (
                <div className="management-sellers-columns">
                  {/* Comerciales que cerraron remisiones */}
                  <div className="mgmt-sellers-card closed-card">
                    <div className="mgmt-sellers-header">
                      <div>
                        <div className="mgmt-sellers-tag green">
                          <CheckCircle2 size={13} /> Facturación y Cierre
                        </div>
                        <h3>Comerciales que cerraron remisiones</h3>
                        <p>
                          {closedBySeller.length > 0
                            ? `${closedBySeller.length} comercial${closedBySeller.length === 1 ? '' : 'es'} lograron facturar o retirar saldo`
                            : 'No hay retiros registrados para los filtros seleccionados'}
                        </p>
                      </div>
                      {employee !== 'Todos' && (
                        <button
                          type="button"
                          className="mgmt-reset-filter-btn"
                          onClick={() => setEmployee('Todos')}
                          title="Quitar filtro de comercial"
                        >
                          Ver todos
                        </button>
                      )}
                    </div>

                    <div className="mgmt-sellers-list">
                      {closedBySeller.length === 0 ? (
                        <div className="mgmt-empty-msg">
                          <span>No se registraron remisiones facturadas para estos filtros.</span>
                        </div>
                      ) : (
                        closedBySeller.map((seller, idx) => {
                          const isSelected = employee === seller.name;
                          const isDimmed = employee !== 'Todos' && !isSelected;
                          const maxTotal = closedBySeller[0]?.total || 1;
                          const pct = Math.round((seller.total / maxTotal) * 100);
                          return (
                            <div
                              key={seller.name}
                              className={`mgmt-seller-row green-row ${isSelected ? 'selected' : ''} ${isDimmed ? 'dimmed' : ''}`}
                              onClick={() => setEmployee((cur) => cur === seller.name ? 'Todos' : seller.name)}
                              role="button"
                              tabIndex={0}
                              title={isSelected ? `Toca para quitar filtro de ${seller.name}` : `Toca para filtrar por ${seller.name}`}
                            >
                              <span className="mgmt-seller-rank">{idx + 1}</span>
                              <div className="mgmt-seller-info">
                                <div className="mgmt-seller-top">
                                  <strong className="mgmt-seller-name">{seller.name}</strong>
                                  <span className="mgmt-seller-count green">{seller.count} remisi{seller.count === 1 ? 'ón facturada' : 'ones facturadas'}</span>
                                </div>
                                <div className="mgmt-seller-bar-track">
                                  <div className="mgmt-seller-bar-fill green" style={{ width: `${pct}%` }} />
                                </div>
                                <div className="mgmt-seller-sub">
                                  <small>{seller.director}</small>
                                  <strong className="mgmt-seller-amount green">{currency.format(seller.total)}</strong>
                                </div>
                              </div>
                            </div>
                          );
                        })
                      )}
                    </div>
                  </div>

                  {/* Comerciales con nuevas remisiones */}
                  <div className="mgmt-sellers-card new-card">
                    <div className="mgmt-sellers-header">
                      <div>
                        <div className="mgmt-sellers-tag orange">
                          <PlusCircle size={13} /> Nuevos Ingresos
                        </div>
                        <h3>Comerciales con nuevas remisiones</h3>
                        <p>
                          {newBySeller.length > 0
                            ? `${newBySeller.length} comercial${newBySeller.length === 1 ? '' : 'es'} abrieron nuevas remisiones`
                            : 'No hay nuevas aperturas registradas para los filtros seleccionados'}
                        </p>
                      </div>
                      {employee !== 'Todos' && (
                        <button
                          type="button"
                          className="mgmt-reset-filter-btn"
                          onClick={() => setEmployee('Todos')}
                          title="Quitar filtro de comercial"
                        >
                          Ver todos
                        </button>
                      )}
                    </div>

                    <div className="mgmt-sellers-list">
                      {newBySeller.length === 0 ? (
                        <div className="mgmt-empty-msg">
                          <span>No se registraron nuevas remisiones para estos filtros.</span>
                        </div>
                      ) : (
                        newBySeller.map((seller, idx) => {
                          const isSelected = employee === seller.name;
                          const isDimmed = employee !== 'Todos' && !isSelected;
                          const maxTotal = newBySeller[0]?.total || 1;
                          const pct = Math.round((seller.total / maxTotal) * 100);
                          return (
                            <div
                              key={seller.name}
                              className={`mgmt-seller-row orange-row ${isSelected ? 'selected' : ''} ${isDimmed ? 'dimmed' : ''}`}
                              onClick={() => setEmployee((cur) => cur === seller.name ? 'Todos' : seller.name)}
                              role="button"
                              tabIndex={0}
                              title={isSelected ? `Toca para quitar filtro de ${seller.name}` : `Toca para filtrar por ${seller.name}`}
                            >
                              <span className="mgmt-seller-rank">{idx + 1}</span>
                              <div className="mgmt-seller-info">
                                <div className="mgmt-seller-top">
                                  <strong className="mgmt-seller-name">{seller.name}</strong>
                                  <span className="mgmt-seller-count orange">{seller.count} remisi{seller.count === 1 ? 'ón nueva' : 'ones nuevas'}</span>
                                </div>
                                <div className="mgmt-seller-bar-track">
                                  <div className="mgmt-seller-bar-fill orange" style={{ width: `${pct}%` }} />
                                </div>
                                <div className="mgmt-seller-sub">
                                  <small>{seller.director}</small>
                                  <strong className="mgmt-seller-amount orange">{currency.format(seller.total)}</strong>
                                </div>
                              </div>
                            </div>
                          );
                        })
                      )}
                    </div>
                  </div>
                </div>
              ) : (
                <div className="management-empty-banner">
                  <FileSpreadsheet size={30} />
                  <div>
                    <strong>Comparación disponible al registrar dos o más días</strong>
                    <p>
                      Para ver qué comerciales cerraron remisiones y quiénes abrieron nuevas, mantén actualizada la hoja <strong>Base-SIS</strong> diariamente. El sistema comparará automáticamente la evolución entre fechas y calculará la trazabilidad completa.
                    </p>
                  </div>
                </div>
              )}
            </section>

            {/* 2. EVOLUCIÓN HISTÓRICA LADO A LADO: VALOR ($) Y # DOCUMENTOS */}
            <DailyEvolutionSideBySide
              daily={daily}
              currentCutoff={cutoff}
              onSelectCutoff={setCutoff}
            />

            {/* 3. SALUD DE CARTERA Y CONCENTRACIÓN */}
            <section className="chart-grid chart-grid-main">
              <AgeCompositionCard
                ageData={ageBreakdown}
                totalPending={currentSummary.pending}
                activeRange={ageFilter}
                onSelectRange={(range) => {
                  setAgeFilter((current) => current === range ? 'Todos' : range);
                }}
              />

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
            </section>

            <section className="chart-grid">
              <ChartCard
                className="chart-wide"
                title="Ejecutivos comerciales con mayor saldo pendiente"
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
          </>
        ) : (
          <section className="detail-card">
            <div className="detail-header">
              <div>
                <div className="detail-title-wrap">
                  <h2>
                    {detailTab === 'withdrawn'
                      ? 'Remisiones retiradas (Cobradas / Facturadas)'
                      : detailTab === 'new'
                        ? 'Nuevas remisiones abiertas (Ingresadas recientemente)'
                        : 'Detalle de remisiones abiertas'}
                  </h2>
                  {previousCutoff && (
                    <div className="detail-tab-buttons">
                      <button
                        type="button"
                        className={`detail-tab-btn ${detailTab === 'open' ? 'active' : ''}`}
                        onClick={() => setDetailTab('open')}
                      >
                        Abiertas ({number.format(currentRecords.length)})
                      </button>
                      <button
                        type="button"
                        className={`detail-tab-btn new ${detailTab === 'new' ? 'active' : ''}`}
                        onClick={() => setDetailTab('new')}
                      >
                        Nuevas ({number.format(newRecords.length)})
                      </button>
                      <button
                        type="button"
                        className={`detail-tab-btn withdrawn ${detailTab === 'withdrawn' ? 'active' : ''}`}
                        onClick={() => setDetailTab('withdrawn')}
                      >
                        Retiradas ({number.format(withdrawnRecords.length)})
                      </button>
                    </div>
                  )}
                </div>
                <p>
                  {number.format(detailRecords.length)}{' '}
                  {detailTab === 'withdrawn'
                    ? 'remisiones facturadas / retiradas frente al día anterior'
                    : detailTab === 'new'
                      ? 'remisiones nuevas abiertas recientemente'
                      : 'registros'}
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
                    'Crítico · Alto valor',
                    'Crítico',
                    'Gestión comercial · Alto valor',
                    'Gestión comercial',
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
                    'Al día (0-15 días)',
                    'Gestión comercial (16-30 días)',
                    'Crítico (>30 días)',
                    'Crítico (>60 días)',
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
                    'Cuantía media ($1M - $5M)',
                    'Menor cuantía (< $1M)',
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
                      isWithdrawn={detailTab === 'withdrawn'}
                      isNew={Boolean((detailTab === 'open' || detailTab === 'new') && previousCutoff && !previousKeysSet.has(record.stableKey))}
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
                ? `Actualizado: ${formatDateTime(data.cutoffDateTime)}`
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
  const isOverdueActive =
    activeRange === '>30 días' ||
    activeRange === 'Crítico (>30 días)' ||
    activeRange === 'Crítico' ||
    activeRange === 'Vencidas (>30 días)';

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
          title={isOverdueActive ? 'Toca para quitar filtro de crítico' : 'Toca para filtrar remisiones críticas (>30 días)'}
        >
          <TriangleAlert size={16} />
          <div>
            <strong>Crítico &gt;30d: {currency.format(overdueTotal)}</strong>
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
        <span className="legend-item orange"><i /> Gestión comercial (16 a 30 días)</span>
        <span className="legend-item red"><i /> Crítico (&gt;30 días)</span>
      </div>
    </article>
  );
}

function EvolutionValueTooltip({
  active,
  payload,
}: {
  active?: boolean;
  payload?: Array<{ payload?: DailyPoint; value?: number; name?: string; color?: string }>;
}) {
  if (!active || !payload?.length) return null;
  const point = payload[0]?.payload;
  if (!point) return null;
  const isInitial = point.pendingDelta === undefined || point.pendingDelta === 0;
  const delta = point.pendingDelta ?? 0;
  const pct = point.pendingDeltaPct ?? 0;
  return (
    <div className="chart-tooltip">
      <strong>{formatCutoff(point.cutoff)}</strong>
      <span>
        <i style={{ background: '#0071e3' }} />
        Saldo Cartera: <b>{currency.format(point.pending)}</b>
      </span>
      {!isInitial && (
        <>
          <span>
            <i style={{ background: delta <= 0 ? '#2fbd68' : '#ff9f0a' }} />
            Variación: <b>{delta > 0 ? '+' : ''}{currency.format(delta)} ({delta > 0 ? '+' : ''}{percent.format(pct)})</b>
          </span>
          {point.newValue > 0 && (
            <span style={{ color: '#c2410c' }}>
              <i style={{ background: '#ff9f0a' }} />
              Nuevas ingresadas: <b>+{currency.format(point.newValue)}</b>
            </span>
          )}
          {point.withdrawn > 0 && (
            <span style={{ color: '#15803d' }}>
              <i style={{ background: '#2fbd68' }} />
              Facturadas / Retiradas: <b>-{currency.format(point.withdrawn)}</b>
            </span>
          )}
        </>
      )}
    </div>
  );
}

function EvolutionDocsTooltip({
  active,
  payload,
}: {
  active?: boolean;
  payload?: Array<{ payload?: DailyPoint; value?: number; name?: string; color?: string }>;
}) {
  if (!active || !payload?.length) return null;
  const point = payload[0]?.payload;
  if (!point) return null;
  const isInitial = point.remissionsDelta === undefined || point.remissionsDelta === 0;
  const delta = point.remissionsDelta ?? 0;
  const pct = point.remissionsDeltaPct ?? 0;
  return (
    <div className="chart-tooltip">
      <strong>{formatCutoff(point.cutoff)}</strong>
      <span>
        <i style={{ background: '#8957d8' }} />
        Total Remisiones: <b>{number.format(point.remissions)} rem.</b>
      </span>
      {!isInitial && (
        <>
          <span>
            <i style={{ background: delta <= 0 ? '#2fbd68' : '#ff9f0a' }} />
            Variación: <b>{delta > 0 ? '+' : ''}{number.format(delta)} rem. ({delta > 0 ? '+' : ''}{percent.format(pct)})</b>
          </span>
          {point.newCount > 0 && (
            <span style={{ color: '#c2410c' }}>
              <i style={{ background: '#ff9f0a' }} />
              Nuevas abiertas: <b>+{number.format(point.newCount)} rem.</b>
            </span>
          )}
          {point.withdrawnCount > 0 && (
            <span style={{ color: '#15803d' }}>
              <i style={{ background: '#2fbd68' }} />
              Facturadas / Retiradas: <b>-{number.format(point.withdrawnCount)} rem.</b>
            </span>
          )}
        </>
      )}
    </div>
  );
}

function DailyEvolutionSideBySide({
  daily,
  currentCutoff,
  onSelectCutoff,
}: {
  daily: DailyPoint[];
  currentCutoff: string;
  onSelectCutoff: (cutoff: string) => void;
}) {
  const [valueChartType, setValueChartType] = useState<'bar' | 'area'>('bar');
  const [docsChartType, setDocsChartType] = useState<'bar' | 'area'>('bar');
  const initialPoint = daily[0];
  const latestPoint = daily.at(-1);
  const totalValueDelta = initialPoint && latestPoint ? latestPoint.pending - initialPoint.pending : 0;
  const totalValuePct = initialPoint && initialPoint.pending > 0 ? totalValueDelta / initialPoint.pending : 0;
  const totalDocsDelta = initialPoint && latestPoint ? latestPoint.remissions - initialPoint.remissions : 0;
  const totalDocsPct = initialPoint && initialPoint.remissions > 0 ? totalDocsDelta / initialPoint.remissions : 0;

  return (
    <section className="evolution-grid" aria-label="Evolución histórica lado a lado">
      {/* Panel 1: Evolución por Valor ($) */}
      <article className="chart-card evolution-card">
        <header className="evolution-header">
          <div>
            <div className="evolution-tag-row">
              <div className="evolution-tag blue">
                <CircleDollarSign size={13} />
                <span>Comportamiento de Saldo</span>
              </div>
              <div className="evolution-chart-type-pill">
                <button
                  type="button"
                  className={valueChartType === 'bar' ? 'active' : ''}
                  onClick={() => setValueChartType('bar')}
                  title="Ver gráfico de barras"
                >
                  Barras
                </button>
                <button
                  type="button"
                  className={valueChartType === 'area' ? 'active' : ''}
                  onClick={() => setValueChartType('area')}
                  title="Ver gráfico de línea continua"
                >
                  Línea
                </button>
              </div>
            </div>
            <h2>Evolución del Valor en Cartera ($)</h2>
            <p>
              {daily.length > 1
                ? `Seguimiento cronológico desde ${formatCutoff(daily[0].cutoff)} hasta ${formatCutoff(daily.at(-1)!.cutoff)}`
                : 'Variación diaria en pesos en el tiempo'}
            </p>
          </div>
          <div className="evolution-kpi-summary">
            <div>
              <span className="evolution-kpi-label">Saldo Inicial</span>
              <strong className="evolution-kpi-val">{compactCurrency.format(initialPoint?.pending || 0)}</strong>
            </div>
            <div className="evolution-kpi-arrow">→</div>
            <div>
              <span className="evolution-kpi-label">Saldo Actual</span>
              <strong className="evolution-kpi-val">{compactCurrency.format(latestPoint?.pending || 0)}</strong>
            </div>
            {daily.length > 1 && (
              <span className={`evolution-badge ${totalValueDelta <= 0 ? 'green' : 'orange'}`}>
                {totalValueDelta <= 0 ? '' : '+'}{percent.format(totalValuePct)}
              </span>
            )}
          </div>
        </header>

        <div className="evolution-chart-wrap">
          <ResponsiveContainer width="100%" height={190}>
            {valueChartType === 'bar' ? (
              <BarChart
                data={daily}
                margin={{ top: 12, right: 8, left: 0, bottom: 0 }}
              >
                <CartesianGrid stroke="#ededf0" vertical={false} />
                <XAxis
                  dataKey="cutoff"
                  tickFormatter={(val) => formatCutoff(val, { day: '2-digit', month: '2-digit', year: undefined })}
                  tickLine={false}
                  axisLine={false}
                />
                <YAxis
                  tickFormatter={(val) => compactCurrency.format(val)}
                  tickLine={false}
                  axisLine={false}
                  width={70}
                />
                <Tooltip content={<EvolutionValueTooltip />} />
                <Bar dataKey="pending" name="Saldo" radius={[5, 5, 0, 0]} maxBarSize={42} cursor="pointer">
                  {daily.map((entry) => (
                    <Cell
                      key={entry.cutoff}
                      fill={entry.cutoff === currentCutoff ? '#0071e3' : '#8ac2ff'}
                      onClick={() => onSelectCutoff(entry.cutoff)}
                    />
                  ))}
                </Bar>
              </BarChart>
            ) : (
              <AreaChart data={daily} margin={{ top: 12, right: 8, left: 0, bottom: 0 }}>
                <defs>
                  <linearGradient id="valGrad" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="0%" stopColor="#0071e3" stopOpacity={0.25} />
                    <stop offset="100%" stopColor="#0071e3" stopOpacity={0.02} />
                  </linearGradient>
                </defs>
                <CartesianGrid stroke="#ededf0" vertical={false} />
                <XAxis
                  dataKey="cutoff"
                  tickFormatter={(val) => formatCutoff(val, { day: '2-digit', month: '2-digit', year: undefined })}
                  tickLine={false}
                  axisLine={false}
                />
                <YAxis
                  tickFormatter={(val) => compactCurrency.format(val)}
                  tickLine={false}
                  axisLine={false}
                  width={70}
                />
                <Tooltip content={<EvolutionValueTooltip />} />
                <Area type="monotone" dataKey="pending" name="Saldo" stroke="#0071e3" strokeWidth={2.5} fill="url(#valGrad)" />
              </AreaChart>
            )}
          </ResponsiveContainer>
        </div>

        <div className="evolution-table-container">
          <div className="evolution-table-title">
            <span>Tabla de variación diaria en pesos:</span>
            <small>Toca una fila para enfocar ese día</small>
          </div>
          <div className="evolution-table-wrap">
            <table className="evolution-table">
              <thead>
                <tr>
                  <th>Fecha</th>
                  <th className="numeric">Saldo Pendiente</th>
                  <th className="numeric">Variación ($)</th>
                  <th className="numeric">Variación (%)</th>
                </tr>
              </thead>
              <tbody>
                {daily.map((pt, idx) => {
                  const isSelected = pt.cutoff === currentCutoff;
                  const isInitial = idx === 0;
                  const delta = pt.pendingDelta ?? 0;
                  const pct = pt.pendingDeltaPct ?? 0;
                  const isDown = delta < 0;
                  const isUp = delta > 0;
                  const hasFlow = !isInitial && (pt.newValue > 0 || pt.withdrawn > 0);
                  return (
                    <tr
                      key={pt.cutoff}
                      className={`evolution-row ${isSelected ? 'selected' : ''}`}
                      onClick={() => onSelectCutoff(pt.cutoff)}
                      role="button"
                      tabIndex={0}
                      title={`Toca para ver datos del ${formatCutoff(pt.cutoff)}`}
                    >
                      <td>
                        <strong>{formatCutoff(pt.cutoff)}</strong>
                        {isSelected && <span className="active-dot-pill">Activo</span>}
                      </td>
                      <td className="numeric">
                        <strong>{currency.format(pt.pending)}</strong>
                      </td>
                      <td className="numeric">
                        {isInitial ? (
                          <span className="text-muted">Punto inicial</span>
                        ) : (
                          <>
                            <span className={isDown ? 'text-green' : isUp ? 'text-orange' : 'text-muted'}>
                              {isUp ? '+' : ''}{currency.format(delta)}
                            </span>
                            {hasFlow && (
                              <small className="evolution-sub-delta">
                                +{compactCurrency.format(pt.newValue)} / -{compactCurrency.format(pt.withdrawn)}
                              </small>
                            )}
                          </>
                        )}
                      </td>
                      <td className="numeric">
                        {isInitial ? (
                          <span className="text-muted">—</span>
                        ) : (
                          <span className={`delta-chip ${isDown ? 'green' : isUp ? 'orange' : 'gray'}`}>
                            {isUp ? '+' : ''}{percent.format(pct)}
                          </span>
                        )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      </article>

      {/* Panel 2: Evolución por Documentos (# Remisiones) */}
      <article className="chart-card evolution-card">
        <header className="evolution-header">
          <div>
            <div className="evolution-tag-row">
              <div className="evolution-tag purple">
                <Boxes size={13} />
                <span>Volumen Documental</span>
              </div>
              <div className="evolution-chart-type-pill">
                <button
                  type="button"
                  className={docsChartType === 'bar' ? 'active' : ''}
                  onClick={() => setDocsChartType('bar')}
                  title="Ver gráfico de barras"
                >
                  Barras
                </button>
                <button
                  type="button"
                  className={docsChartType === 'area' ? 'active' : ''}
                  onClick={() => setDocsChartType('area')}
                  title="Ver gráfico de línea continua"
                >
                  Línea
                </button>
              </div>
            </div>
            <h2>Evolución de Remisiones (# Docs)</h2>
            <p>
              {daily.length > 1
                ? `Conteo de documentos desde ${formatCutoff(daily[0].cutoff)} hasta ${formatCutoff(daily.at(-1)!.cutoff)}`
                : 'Conteo y fluctuación diaria del número de remisiones abiertas'}
            </p>
          </div>
          <div className="evolution-kpi-summary">
            <div>
              <span className="evolution-kpi-label">Docs Inicio</span>
              <strong className="evolution-kpi-val">{number.format(initialPoint?.remissions || 0)} rem.</strong>
            </div>
            <div className="evolution-kpi-arrow">→</div>
            <div>
              <span className="evolution-kpi-label">Docs Hoy</span>
              <strong className="evolution-kpi-val">{number.format(latestPoint?.remissions || 0)} rem.</strong>
            </div>
            {daily.length > 1 && (
              <span className={`evolution-badge ${totalDocsDelta <= 0 ? 'green' : 'orange'}`}>
                {totalDocsDelta <= 0 ? '' : '+'}{percent.format(totalDocsPct)}
              </span>
            )}
          </div>
        </header>

        <div className="evolution-chart-wrap">
          <ResponsiveContainer width="100%" height={190}>
            {docsChartType === 'bar' ? (
              <BarChart
                data={daily}
                margin={{ top: 12, right: 8, left: 0, bottom: 0 }}
              >
                <CartesianGrid stroke="#ededf0" vertical={false} />
                <XAxis
                  dataKey="cutoff"
                  tickFormatter={(val) => formatCutoff(val, { day: '2-digit', month: '2-digit', year: undefined })}
                  tickLine={false}
                  axisLine={false}
                />
                <YAxis
                  tickFormatter={(val) => number.format(val)}
                  tickLine={false}
                  axisLine={false}
                  width={50}
                />
                <Tooltip content={<EvolutionDocsTooltip />} />
                <Bar dataKey="remissions" name="Remisiones" radius={[5, 5, 0, 0]} maxBarSize={42} cursor="pointer">
                  {daily.map((entry) => (
                    <Cell
                      key={entry.cutoff}
                      fill={entry.cutoff === currentCutoff ? '#8957d8' : '#cbb2f5'}
                      onClick={() => onSelectCutoff(entry.cutoff)}
                    />
                  ))}
                </Bar>
              </BarChart>
            ) : (
              <AreaChart data={daily} margin={{ top: 12, right: 8, left: 0, bottom: 0 }}>
                <defs>
                  <linearGradient id="docsGrad" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="0%" stopColor="#8957d8" stopOpacity={0.25} />
                    <stop offset="100%" stopColor="#8957d8" stopOpacity={0.02} />
                  </linearGradient>
                </defs>
                <CartesianGrid stroke="#ededf0" vertical={false} />
                <XAxis
                  dataKey="cutoff"
                  tickFormatter={(val) => formatCutoff(val, { day: '2-digit', month: '2-digit', year: undefined })}
                  tickLine={false}
                  axisLine={false}
                />
                <YAxis
                  tickFormatter={(val) => number.format(val)}
                  tickLine={false}
                  axisLine={false}
                  width={50}
                />
                <Tooltip content={<EvolutionDocsTooltip />} />
                <Area type="monotone" dataKey="remissions" name="Remisiones" stroke="#8957d8" strokeWidth={2.5} fill="url(#docsGrad)" />
              </AreaChart>
            )}
          </ResponsiveContainer>
        </div>

        <div className="evolution-table-container">
          <div className="evolution-table-title">
            <span>Tabla de variación documental:</span>
            <small>Toca una fila para enfocar ese día</small>
          </div>
          <div className="evolution-table-wrap">
            <table className="evolution-table">
              <thead>
                <tr>
                  <th>Fecha</th>
                  <th className="numeric"># Remisiones</th>
                  <th className="numeric">Variación (#)</th>
                  <th className="numeric">Variación (%)</th>
                </tr>
              </thead>
              <tbody>
                {daily.map((pt, idx) => {
                  const isSelected = pt.cutoff === currentCutoff;
                  const isInitial = idx === 0;
                  const delta = pt.remissionsDelta ?? 0;
                  const pct = pt.remissionsDeltaPct ?? 0;
                  const isDown = delta < 0;
                  const isUp = delta > 0;
                  const hasFlow = !isInitial && (pt.newCount > 0 || pt.withdrawnCount > 0);
                  return (
                    <tr
                      key={pt.cutoff}
                      className={`evolution-row ${isSelected ? 'selected' : ''}`}
                      onClick={() => onSelectCutoff(pt.cutoff)}
                      role="button"
                      tabIndex={0}
                      title={`Toca para ver datos del ${formatCutoff(pt.cutoff)}`}
                    >
                      <td>
                        <strong>{formatCutoff(pt.cutoff)}</strong>
                        {isSelected && <span className="active-dot-pill">Activo</span>}
                      </td>
                      <td className="numeric">
                        <strong>{number.format(pt.remissions)} rem.</strong>
                      </td>
                      <td className="numeric">
                        {isInitial ? (
                          <span className="text-muted">Punto inicial</span>
                        ) : (
                          <>
                            <span className={isDown ? 'text-green' : isUp ? 'text-orange' : 'text-muted'}>
                              {isUp ? '+' : ''}{number.format(delta)} rem.
                            </span>
                            {hasFlow && (
                              <small className="evolution-sub-delta">
                                +{pt.newCount} / -{pt.withdrawnCount}
                              </small>
                            )}
                          </>
                        )}
                      </td>
                      <td className="numeric">
                        {isInitial ? (
                          <span className="text-muted">—</span>
                        ) : (
                          <span className={`delta-chip ${isDown ? 'green' : isUp ? 'orange' : 'gray'}`}>
                            {isUp ? '+' : ''}{percent.format(pct)}
                          </span>
                        )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      </article>
    </section>
  );
}

function ChartCard({ title, subtitle, children, className = '' }: { title: string; subtitle: string; children: React.ReactNode; className?: string }) {
  return <article className={`chart-card ${className}`}><header><div><h2>{title}</h2><p>{subtitle}</p></div></header>{children}</article>;
}

function DailyStat({
  label,
  value,
  sub,
  tone = '',
  onClick,
  clickable = false,
}: {
  label: string;
  value: string;
  sub?: string;
  tone?: string;
  onClick?: () => void;
  clickable?: boolean;
}) {
  return (
    <div
      className={`daily-stat ${tone} ${clickable ? 'clickable' : ''}`}
      onClick={onClick}
      role={clickable ? 'button' : undefined}
      tabIndex={clickable ? 0 : undefined}
    >
      <span>{label}</span>
      <strong>{value}</strong>
      {sub && <small style={{ fontSize: '9.5px', color: 'var(--muted)', marginTop: '2px', display: 'block' }}>{sub}</small>}
    </div>
  );
}

function DetailRow({
  record,
  isWithdrawn = false,
  isNew = false,
  onSelectDirector,
  onSelectEmployee,
  onSelectAge,
  onSelectAmount,
  onSelectStatus,
  onSelectCompany,
}: {
  record: Remision;
  isWithdrawn?: boolean;
  isNew?: boolean;
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
    <tr className={isWithdrawn ? 'row-withdrawn' : ''}>
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
        <strong className="cell-doc">
          {record.document || '—'}
          {isNew && <span className="new-badge">Nueva</span>}
        </strong>
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
          <span className={`amount-chip ${record.total >= 5_000_000 ? 'high' : record.total >= 1_000_000 ? 'medium' : 'low'}`}>
            {record.amountStatus}
          </span>
        </button>
      </td>
      <td>
        {isWithdrawn ? (
          <span className="status-pill normal">Facturada / Retirada</span>
        ) : (
          <button
            type="button"
            className="cell-touch-btn"
            onClick={() => onSelectStatus?.(record.alert)}
            title={`Toca para filtrar por estado "${record.alert}"`}
          >
            <span className={`status-pill ${statusClass(record.alert)}`}>{record.alert}</span>
          </button>
        )}
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
  if (alert === 'Crítico · Alto valor' || (alert as string) === 'Vencida · Alto valor') return 'critical-high';
  if (alert === 'Crítico' || (alert as string) === 'Vencida') return 'critical';
  if (alert === 'Gestión comercial · Alto valor' || (alert as string) === 'Por vencer · Alto valor') return 'commercial-high';
  if (alert === 'Gestión comercial' || (alert as string) === 'Por vencer') return 'commercial';
  if (alert === 'Al día · Alto valor') return 'normal-high';
  return 'normal';
}

export default App;
