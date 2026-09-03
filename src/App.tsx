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
  ListChecks,
  LoaderCircle,
  LogIn,
  LogOut,
  PackageCheck,
  RefreshCw,
  Search,
  ShieldCheck,
  TriangleAlert,
  Upload,
  UsersRound,
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
  buildDailySeries,
  formatCutoff,
  normalizeText,
  parseRemisionesWorkbook,
  summarize,
} from './lib/remisiones';
import type { DataSource, FileMetadata, ParsedWorkbook, Remision, Summary, UserProfile } from './types';

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
    setLoadingMessage('Leyendo Base, Diario y Grupos…');
    const parsed = await parseRemisionesWorkbook(buffer);
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
  const [alert, setAlert] = useState('Todas');
  const [query, setQuery] = useState('');
  const [page, setPage] = useState(1);
  const refreshRef = useRef(onRefresh);

  const directors = useMemo(
    () => [...new Set(data.records.map((record) => record.director))].sort((a, b) => a.localeCompare(b, 'es')),
    [data.records],
  );
  const employees = useMemo(() => {
    const scoped = director === 'Todos' ? data.records : data.records.filter((record) => record.director === director);
    return [...new Set(scoped.map((record) => record.employee))].sort((a, b) => a.localeCompare(b, 'es'));
  }, [data.records, director]);
  const scopedRecords = useMemo(
    () => data.records.filter((record) =>
      (director === 'Todos' || record.director === director) &&
      (employee === 'Todos' || record.employee === employee)),
    [data.records, director, employee],
  );
  const currentRecords = useMemo(
    () => scopedRecords.filter((record) => record.cutoff === cutoff),
    [scopedRecords, cutoff],
  );
  const currentSummary = useMemo(() => summarize(currentRecords), [currentRecords]);
  const previousCutoff = data.cutoffs.filter((date) => date < cutoff).at(-1) || '';
  const previousSummary = useMemo(
    () => summarize(scopedRecords.filter((record) => record.cutoff === previousCutoff)),
    [scopedRecords, previousCutoff],
  );
  const periodRecords = useMemo(
    () => scopedRecords.filter((record) => record.cutoff >= from && record.cutoff <= to),
    [scopedRecords, from, to],
  );
  const daily = useMemo(() => buildDailySeries(periodRecords), [periodRecords]);
  const ageData = useMemo(() => {
    const groups = aggregateBy(currentRecords, (record) => record.ageRange);
    return AGE_ORDER.map((name) => ({ name, ...(groups.find((group) => group.name === name) || { value: 0, count: 0 }) }));
  }, [currentRecords]);
  const directorData = useMemo(() => aggregateBy(currentRecords, (record) => record.director), [currentRecords]);
  const sellerData = useMemo(() => aggregateBy(currentRecords, (record) => record.employee).slice(0, 12), [currentRecords]);
  const detailRecords = useMemo(() => {
    const normalizedQuery = normalizeText(query);
    return currentRecords
      .filter((record) => alert === 'Todas' || record.alert === alert)
      .filter((record) => !normalizedQuery || normalizeText([
        record.company,
        record.nit,
        record.employee,
        record.document,
        record.order,
      ].join(' ')).includes(normalizedQuery))
      .sort((a, b) => b.age - a.age || b.total - a.total);
  }, [currentRecords, alert, query]);

  const pageSize = 20;
  const pageCount = Math.max(1, Math.ceil(detailRecords.length / pageSize));
  const visibleRows = detailRecords.slice((page - 1) * pageSize, page * pageSize);

  useEffect(() => setPage(1), [query, alert, cutoff, director, employee]);
  useEffect(() => { refreshRef.current = onRefresh; }, [onRefresh]);
  useEffect(() => {
    if (source !== 'sharepoint') return undefined;
    const interval = window.setInterval(() => refreshRef.current(), 10 * 60 * 1000);
    return () => window.clearInterval(interval);
  }, [source]);
  useEffect(() => {
    if (employee !== 'Todos' && !employees.includes(employee)) setEmployee('Todos');
  }, [director, employee, employees]);

  const exportCsv = () => {
    const headers = ['Corte', 'Director', 'Empleado', 'NIT', 'Empresa', 'Documento', 'Pedido', 'Emisión', 'Días', 'Mercancía', 'IVA', 'Total', 'Cantidad', 'Alerta'];
    const rows = detailRecords.map((record) => [
      record.cutoff,
      record.director,
      record.employee,
      record.nit,
      record.company,
      record.document,
      record.order,
      record.issuedAt,
      record.age,
      record.merchandise,
      record.tax,
      record.total,
      record.quantity,
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
            <p>Una lectura clara del pendiente por facturar y de la gestión de cada equipo.</p>
          </div>
          <div className="source-card">
            <span className={`source-icon ${source}`}><FileSpreadsheet size={21} /></span>
            <div><strong>{metadata?.name || 'Remisiones.xlsx'}</strong><small>{source === 'sharepoint' ? 'SharePoint · actualización automática' : 'Archivo abierto localmente'}</small></div>
            <span className="source-status"><i /> Disponible</span>
          </div>
        </section>

        <section className="filter-bar" aria-label="Filtros del tablero">
          <SelectFilter label="Corte" value={cutoff} onChange={setCutoff} options={data.cutoffs} format={formatCutoff} />
          <SelectFilter label="Director" value={director} onChange={setDirector} options={['Todos', ...directors]} />
          <SelectFilter label="Ejecutivo" value={employee} onChange={setEmployee} options={['Todos', ...employees]} />
          <div className="filter-separator" />
          <label className="date-filter"><span>Desde</span><input type="date" value={from} min={data.cutoffs[0]} max={to} onChange={(event) => setFrom(event.target.value)} /></label>
          <label className="date-filter"><span>Hasta</span><input type="date" value={to} min={from} max={data.cutoffs.at(-1)} onChange={(event) => setTo(event.target.value)} /></label>
        </section>

        {data.unmatchedEmployees.length > 0 && (
          <div className="quality-notice">
            <TriangleAlert size={18} />
            <div><strong>{data.unmatchedEmployees.length} nombres sin grupo asignado</strong><span>Se muestran como “Sin asignar”. Completa la hoja Grupos para incorporarlos al equipo correcto.</span></div>
            <button onClick={() => { setDirector('Sin asignar'); setView('detail'); }}>Revisar</button>
          </div>
        )}

        {view === 'overview' ? (
          <>
            <section className="metric-grid">
              <MetricCard title="Pendiente total" value={currency.format(currentSummary.pending)} icon={<CircleDollarSign />} tone="blue" current={currentSummary.pending} previous={previousSummary.pending} />
              <MetricCard title="Remisiones" value={number.format(currentSummary.remissions)} sub={`${number.format(currentSummary.clients)} clientes`} icon={<Boxes />} tone="purple" current={currentSummary.remissions} previous={previousSummary.remissions} />
              <MetricCard title="Saldo >30 días" value={currency.format(currentSummary.overdueValue)} sub={`${number.format(currentSummary.overdueCount)} remisiones`} icon={<TriangleAlert />} tone="red" current={currentSummary.overdueValue} previous={previousSummary.overdueValue} />
              <MetricCard title="Antigüedad promedio" value={`${currentSummary.averageAge.toFixed(1)} días`} sub={`${number.format(currentSummary.zeroQuantity)} con cantidad cero`} icon={<CalendarRange />} tone="orange" current={currentSummary.averageAge} previous={previousSummary.averageAge} />
            </section>

            <section className="chart-grid chart-grid-main">
              <ChartCard className="chart-wide" title="Evolución del pendiente" subtitle={daily.length > 1 ? `${formatCutoff(from)} — ${formatCutoff(to)}` : 'Agrega cortes diarios en Base para construir la tendencia'}>
                <ResponsiveContainer width="100%" height={300}>
                  <AreaChart data={daily} margin={{ top: 18, right: 8, left: 0, bottom: 0 }}>
                    <defs>
                      <linearGradient id="pendingFill" x1="0" y1="0" x2="0" y2="1">
                        <stop offset="0%" stopColor="#0071e3" stopOpacity={0.28} />
                        <stop offset="100%" stopColor="#0071e3" stopOpacity={0.02} />
                      </linearGradient>
                    </defs>
                    <CartesianGrid stroke="#e8e8ed" vertical={false} />
                    <XAxis dataKey="cutoff" tickFormatter={(value) => formatCutoff(value, { day: '2-digit', month: 'short', year: undefined })} tickLine={false} axisLine={false} />
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
              <ChartCard title="Composición por antigüedad" subtitle="Valor pendiente por rango">
                <ResponsiveContainer width="100%" height={300}>
                  <BarChart data={ageData} layout="vertical" margin={{ top: 10, right: 8, left: 0, bottom: 0 }}>
                    <CartesianGrid stroke="#e8e8ed" horizontal={false} />
                    <XAxis type="number" hide />
                    <YAxis type="category" dataKey="name" tickLine={false} axisLine={false} width={86} />
                    <Tooltip content={<CurrencyTooltip />} />
                    <Bar dataKey="value" name="Pendiente" radius={[0, 7, 7, 0]} barSize={19}>
                      {ageData.map((entry, index) => <Cell key={entry.name} fill={index >= 4 ? '#ff453a' : index === 3 ? '#ff9f0a' : '#0071e3'} />)}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </ChartCard>
            </section>

            <section className="chart-grid">
              <ChartCard title="Pendiente por dirección" subtitle="Distribución del corte seleccionado">
                <div className="donut-layout">
                  <ResponsiveContainer width="48%" height={260}>
                    <PieChart>
                      <Pie data={directorData} dataKey="value" nameKey="name" innerRadius={64} outerRadius={94} paddingAngle={2} stroke="none">
                        {directorData.map((entry, index) => <Cell key={entry.name} fill={PIE_COLORS[index % PIE_COLORS.length]} />)}
                      </Pie>
                      <Tooltip formatter={(value) => currency.format(Number(value))} />
                    </PieChart>
                  </ResponsiveContainer>
                  <div className="legend-list">
                    {directorData.map((entry, index) => (
                      <button key={entry.name} onClick={() => setDirector(entry.name)}>
                        <i style={{ background: PIE_COLORS[index % PIE_COLORS.length] }} />
                        <span>{entry.name}<small>{entry.count} remisiones</small></span>
                        <strong>{compactCurrency.format(entry.value)}</strong>
                      </button>
                    ))}
                  </div>
                </div>
              </ChartCard>
              <ChartCard title="Ejecutivos con mayor pendiente" subtitle="Top 12 por valor total">
                <ResponsiveContainer width="100%" height={300}>
                  <BarChart data={sellerData} margin={{ top: 8, right: 8, left: 0, bottom: 45 }}>
                    <CartesianGrid stroke="#e8e8ed" vertical={false} />
                    <XAxis dataKey="name" interval={0} angle={-32} textAnchor="end" height={80} tickFormatter={(value) => String(value).split(' ').slice(0, 2).join(' ')} tickLine={false} axisLine={false} />
                    <YAxis tickFormatter={(value) => compactCurrency.format(value)} tickLine={false} axisLine={false} width={72} />
                    <Tooltip content={<CurrencyTooltip />} />
                    <Bar dataKey="value" name="Pendiente" fill="#af52de" radius={[7, 7, 0, 0]} maxBarSize={34} />
                  </BarChart>
                </ResponsiveContainer>
              </ChartCard>
            </section>

            <section className="daily-strip">
              <div className="daily-strip-title"><CalendarRange size={20} /><div><strong>Gestión del último corte</strong><span>Entradas y retiros frente al corte anterior</span></div></div>
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
              <div><h2>Detalle de remisiones</h2><p>{number.format(detailRecords.length)} registros en el corte del {formatCutoff(cutoff)}</p></div>
              <div className="detail-actions">
                <label className="search-box"><Search size={17} /><input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="Cliente, NIT, documento…" /></label>
                <SelectFilter label="Alerta" value={alert} onChange={setAlert} options={['Todas', 'Vencida >30 días', 'Prioritaria', 'Cantidad en cero', 'Revisar valor', 'Normal']} compact />
                <button className="button button-secondary" onClick={exportCsv}><Download size={17} /> Exportar</button>
              </div>
            </div>
            <div className="table-wrap">
              <table>
                <thead><tr><th>Cliente</th><th>Responsable</th><th>Documento</th><th>Emisión</th><th className="numeric">Días</th><th className="numeric">Total</th><th>Estado</th></tr></thead>
                <tbody>
                  {visibleRows.map((record) => <DetailRow key={record.id} record={record} />)}
                  {!visibleRows.length && <tr><td colSpan={7}><div className="empty-state"><Search size={22} />No hay resultados para estos filtros.</div></td></tr>}
                </tbody>
              </table>
            </div>
            <div className="pagination">
              <span>Mostrando {visibleRows.length ? (page - 1) * pageSize + 1 : 0}–{Math.min(page * pageSize, detailRecords.length)} de {detailRecords.length}</span>
              <div><button disabled={page <= 1} onClick={() => setPage((value) => value - 1)}>Anterior</button><span>Página {page} de {pageCount}</span><button disabled={page >= pageCount} onClick={() => setPage((value) => value + 1)}>Siguiente</button></div>
            </div>
          </section>
        )}

        <footer className="app-footer">
          <span>Provexpress · Control de remisiones abiertas</span>
          <span>{metadata?.lastModifiedDateTime ? `Archivo actualizado ${new Intl.DateTimeFormat('es-CO', { dateStyle: 'medium', timeStyle: 'short' }).format(new Date(metadata.lastModifiedDateTime))}` : ''}</span>
        </footer>
      </main>
    </div>
  );
}

function SelectFilter({ label, value, options, onChange, format, compact = false }: { label: string; value: string; options: string[]; onChange: (value: string) => void; format?: (value: string) => string; compact?: boolean }) {
  return (
    <label className={`select-filter ${compact ? 'compact' : ''}`}>
      <span>{label}</span>
      <div><select value={value} onChange={(event) => onChange(event.target.value)}>{options.map((option) => <option key={option} value={option}>{format ? format(option) : option}</option>)}</select><ChevronDown size={14} /></div>
    </label>
  );
}

function MetricCard({ title, value, sub, icon, tone, current, previous }: { title: string; value: string; sub?: string; icon: React.ReactNode; tone: string; current: number; previous: number }) {
  const delta = previous ? (current - previous) / Math.abs(previous) : 0;
  const isUp = delta > 0;
  return (
    <article className="metric-card">
      <div className={`metric-icon ${tone}`}>{icon}</div>
      <span className="metric-title">{title}</span>
      <strong className="metric-value">{value}</strong>
      <div className="metric-caption">
        {sub ? <span>{sub}</span> : <span />}
        {previous > 0 && <span className={isUp ? 'delta-up' : 'delta-down'}>{isUp ? <ArrowUpRight size={14} /> : <ArrowDownRight size={14} />}{percent.format(Math.abs(delta))}</span>}
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

function DetailRow({ record }: { record: Remision }) {
  return (
    <tr>
      <td><strong>{record.company || 'Sin empresa'}</strong><small>NIT {record.nit || '—'}</small></td>
      <td><span>{record.employee}</span><small>{record.director}</small></td>
      <td><span>{record.document || '—'}</span><small>Pedido {record.order || '—'}</small></td>
      <td>{formatCutoff(record.issuedAt)}</td>
      <td className="numeric"><strong>{record.age}</strong></td>
      <td className="numeric"><strong>{currency.format(record.total)}</strong><small>{currency.format(record.merchandise)} + IVA</small></td>
      <td><span className={`status-pill ${statusClass(record.alert)}`}>{record.alert}</span></td>
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
  if (alert === 'Normal') return 'normal';
  if (alert === 'Prioritaria') return 'priority';
  if (alert === 'Vencida >30 días') return 'overdue';
  return 'review';
}

export default App;
