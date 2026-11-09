import React from 'react';
import { Calendar, ChevronDown, Clock, Layers, Filter } from 'lucide-react';
import type {
  PeriodFilterState,
  PeriodStructure,
  PeriodMode,
} from '../lib/periodFilter';

interface PeriodFilterControlProps {
  filter: PeriodFilterState;
  onChangeFilter: (next: PeriodFilterState) => void;
  structure: PeriodStructure;
  visibleDaysCount: number;
  totalDaysCount: number;
  theme?: 'green' | 'purple' | 'blue';
  title?: string;
}

export const PeriodFilterControl: React.FC<PeriodFilterControlProps> = ({
  filter,
  onChangeFilter,
  structure,
  visibleDaysCount,
  totalDaysCount,
  theme = 'green',
  title = 'Filtrar periodo',
}) => {
  const {
    months,
    weeks,
    currentWeekKey,
    previousWeekKey,
    currentMonthKey,
    currentMonthLabel,
  } = structure;

  // Semanas disponibles según el mes seleccionado en modo 'custom'
  const availableWeeks = filter.monthKey && filter.monthKey !== 'all'
    ? weeks.filter((w) => w.monthKey === filter.monthKey)
    : weeks;

  const handleModeClick = (mode: PeriodMode) => {
    if (mode === 'current-week') {
      onChangeFilter({ mode: 'current-week', monthKey: 'all', weekKey: currentWeekKey });
    } else if (mode === 'previous-week') {
      onChangeFilter({ mode: 'previous-week', monthKey: 'all', weekKey: previousWeekKey });
    } else if (mode === 'current-month') {
      onChangeFilter({ mode: 'current-month', monthKey: currentMonthKey, weekKey: 'all' });
    } else if (mode === 'all') {
      onChangeFilter({ mode: 'all', monthKey: 'all', weekKey: 'all' });
    }
  };

  const handleMonthChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    const val = e.target.value;
    onChangeFilter({
      mode: 'custom',
      monthKey: val,
      weekKey: 'all', // Reset week when month changes
    });
  };

  const handleWeekChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    const val = e.target.value;
    onChangeFilter({
      ...filter,
      mode: 'custom',
      weekKey: val,
    });
  };

  const isCurrentWeek = filter.mode === 'current-week';
  const isPreviousWeek = filter.mode === 'previous-week';
  const isCurrentMonth = filter.mode === 'current-month';
  const isAll = filter.mode === 'all';
  const isCustom = filter.mode === 'custom';

  return (
    <div className={`period-filter-bar theme-${theme}`} aria-label={title}>
      <div className="period-filter-pills-row">
        <div className="period-filter-label-wrap">
          <Filter size={13} className="period-filter-icon" />
          <span className="period-filter-title">Vista rápida:</span>
        </div>

        <button
          type="button"
          className={`period-pill-btn ${isCurrentWeek ? 'active' : ''}`}
          onClick={() => handleModeClick('current-week')}
          title="Ver exclusivamente los cortes de la semana más reciente"
        >
          <Calendar size={13} />
          <span>Esta Semana</span>
        </button>

        {previousWeekKey && (
          <button
            type="button"
            className={`period-pill-btn ${isPreviousWeek ? 'active' : ''}`}
            onClick={() => handleModeClick('previous-week')}
            title="Ver los cortes de la semana anterior"
          >
            <Clock size={13} />
            <span>Semana Anterior</span>
          </button>
        )}

        <button
          type="button"
          className={`period-pill-btn ${isCurrentMonth ? 'active' : ''}`}
          onClick={() => handleModeClick('current-month')}
          title={`Ver todos los cortes de ${currentMonthLabel}`}
        >
          <Calendar size={13} />
          <span>{currentMonthLabel}</span>
        </button>

        <button
          type="button"
          className={`period-pill-btn ${isAll ? 'active' : ''}`}
          onClick={() => handleModeClick('all')}
          title="Ver todo el historial desde el corte inicial de Septiembre"
        >
          <Layers size={13} />
          <span>Todo el Histórico ({totalDaysCount} cortes)</span>
        </button>
      </div>

      <div className="period-filter-selects-row">
        {/* Selector de Mes */}
        <div className="period-select-wrap">
          <label htmlFor={`period-month-select-${theme}`} className="period-select-label">Mes:</label>
          <div className="period-select-box">
            <select
              id={`period-month-select-${theme}`}
              value={isCustom ? filter.monthKey : isCurrentMonth ? currentMonthKey : 'all'}
              onChange={handleMonthChange}
              className="period-select"
            >
              <option value="all">Todos los meses</option>
              {months.map((m) => (
                <option key={m.key} value={m.key}>
                  {m.label} ({m.count} {m.count === 1 ? 'día' : 'días'})
                </option>
              ))}
            </select>
            <ChevronDown size={13} className="period-select-chevron" />
          </div>
        </div>

        {/* Selector de Semana */}
        <div className="period-select-wrap">
          <label htmlFor={`period-week-select-${theme}`} className="period-select-label">Semana:</label>
          <div className="period-select-box">
            <select
              id={`period-week-select-${theme}`}
              value={isCustom ? filter.weekKey : isCurrentWeek ? currentWeekKey : isPreviousWeek ? previousWeekKey : 'all'}
              onChange={handleWeekChange}
              className="period-select"
            >
              <option value="all">Todas las semanas</option>
              {availableWeeks.map((w) => (
                <option key={w.key} value={w.key}>
                  {w.label} ({w.count} {w.count === 1 ? 'corte' : 'cortes'})
                </option>
              ))}
            </select>
            <ChevronDown size={13} className="period-select-chevron" />
          </div>
        </div>

        <div className="period-counter-badge">
          <span>Mostrando <b>{visibleDaysCount}</b> de {totalDaysCount} cortes</span>
        </div>
      </div>
    </div>
  );
};
