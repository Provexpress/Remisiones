export type DataSource = 'sharepoint' | 'local';

export interface UserProfile {
  name: string;
  email: string;
}

export interface FileMetadata {
  name: string;
  size?: number;
  lastModifiedDateTime?: string;
  webUrl?: string;
}

export interface Remision {
  id: string;
  stableKey: string;
  cutoff: string;
  cutoffTime?: string;
  cutoffDateTime?: string;
  employee: string;
  nit: string;
  company: string;
  merchandise: number;
  tax: number;
  total: number;
  issuedAt: string;
  age: number;
  document: string;
  order: string;
  quantity: number;
  ageRange: string;
  amountStatus: AmountStatus;
  daysStatus: DaysStatus;
  alert: AlertLevel;
  director: string;
  group: number | null;
  matchedGroup: boolean;
}

export type AmountStatus =
  | 'Alto valor (> $5M)'
  | 'Valor medio ($1M - $5M)'
  | 'Menor valor (< $1M)';

export type DaysStatus =
  | 'Crítica (>60 días)'
  | 'Vencida (31-60 días)'
  | 'Por vencer (16-30 días)'
  | 'Al día (0-15 días)';

export type AlertLevel =
  | 'Vencida · Alto valor'
  | 'Vencida'
  | 'Por vencer · Alto valor'
  | 'Por vencer'
  | 'Al día · Alto valor'
  | 'Al día';

export interface GroupEntry {
  group: number;
  director: string;
  member: string;
}

export interface ParsedWorkbook {
  records: Remision[];
  groups: GroupEntry[];
  sheetNames: string[];
  cutoffs: string[];
  unmatchedEmployees: string[];
  activeSheetName?: string;
  cutoffDateTime?: string;
  cutoffTimeDisplay?: string;
}

export interface DailyPoint {
  cutoff: string;
  pending: number;
  remissions: number;
  clients: number;
  newValue: number;
  newCount: number;
  previousBalance: number;
  withdrawn: number;
  withdrawnCount: number;
  netManagement: number;
  grossReduction: number;
  overdueValue: number;
  overdueCount: number;
}

export interface Summary {
  pending: number;
  remissions: number;
  merchandise: number;
  tax: number;
  averageAge: number;
  overdueValue: number;
  overdueCount: number;
  zeroQuantity: number;
  clients: number;
  overduePercentage?: number;
  maxAge?: number;
}

export interface AgeBreakdownItem {
  name: string;
  value: number;
  count: number;
  overdue: number;
  percent: number;
  tone: 'blue' | 'orange' | 'red';
  badge?: string;
}

