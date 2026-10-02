import { serialToDate } from './formula';

// Модель данных листа OpenDesk Sheets (уровень Excel 2010, базовый набор)

export type NumFmt = 'general' | 'number' | 'currency' | 'percent' | 'date' | 'text';

export const NUM_FMT_LABEL: Record<NumFmt, string> = {
  general: 'Общий',
  number: 'Числовой',
  currency: 'Денежный (₽)',
  percent: 'Процентный',
  date: 'Дата',
  text: 'Текстовый',
};

export interface CellStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  font?: string;
  size?: number; // pt
  color?: string; // #rrggbb
  fill?: string; // #rrggbb
  align?: 'left' | 'center' | 'right';
  numFmt?: NumFmt;
  border?: boolean;
}

export interface CellComment {
  author: string;
  text: string;
}

export interface Validation {
  type: 'list' | 'whole' | 'decimal';
  values?: string[];
  min?: number;
  max?: number;
}

export type CFKind = 'scale' | 'bar' | 'icon' | 'cellIs';

export interface CFRule {
  id: number;
  range: string; // "B2:B20"
  kind: CFKind;
  op?: '>' | '<' | '=' | '>=' | '<=';
  value?: number;
  color?: string;
}

export interface SparkDef {
  at: string; // ячейка спарклайна
  range: string; // диапазон данных
  type: 'line' | 'column' | 'winloss';
}

export interface ChartDef {
  id: number;
  title: string;
  type: 'bar' | 'line' | 'pie';
  range: string; // диапазон значений (один столбец/строка)
  labels?: string; // диапазон подписей (опционально)
}

export type Frozen = 'none' | 'row' | 'col' | 'both';

export interface Sheet {
  name: string;
  cells: Record<string, string>;
  styles: Record<string, CellStyle>;
  merges: string[]; // ["B2:C4"]
  comments: Record<string, CellComment>;
  validations: Record<string, Validation>;
  cf: CFRule[];
  sparks: SparkDef[];
  charts: ChartDef[];
  frozen: Frozen;
  filterCol: number | null; // столбец автофильтра
  filterValues: string[]; // выбранные значения (пусто = все)
  protected: boolean;
}

export function newSheet(name: string): Sheet {
  return {
    name,
    cells: {},
    styles: {},
    merges: [],
    comments: {},
    validations: {},
    cf: [],
    sparks: [],
    charts: [],
    frozen: 'none',
    filterCol: null,
    filterValues: [],
    protected: false,
  };
}

// ---------- Форматирование чисел ----------

export function formatValue(raw: number | string, style?: CellStyle): string {
  if (typeof raw === 'string') return raw;
  const fmt = style?.numFmt ?? 'general';
  if (!Number.isFinite(raw)) return String(raw);
  switch (fmt) {
    case 'number':
      return raw.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    case 'currency':
      return raw.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' ₽';
    case 'percent':
      return (Math.round(raw * 10000) / 100).toLocaleString('ru-RU') + '%';
    case 'date': {
      const { y, m, d } = serialToDate(raw);
      return `${String(d).padStart(2, '0')}.${String(m).padStart(2, '0')}.${y}`;
    }
    case 'text':
      return String(raw);
    default: {
      const r = Math.round(raw * 1e10) / 1e10;
      return String(r);
    }
  }
}

export function cssColor(hex?: string): string | undefined {
  return hex;
}

// ---------- Условное форматирование: вычисление ----------

export interface CFAssess {
  fill?: string;
  color?: string;
  barPct?: number; // 0..100 для dataBar
  icon?: '▲' | '●' | '▼';
}

export function assessCF(value: number, all: number[], rule: CFRule): CFAssess | null {
  if (rule.kind === 'cellIs' && rule.value !== undefined) {
    const ok =
      rule.op === '>' ? value > rule.value
      : rule.op === '<' ? value < rule.value
      : rule.op === '>=' ? value >= rule.value
      : rule.op === '<=' ? value <= rule.value
      : value === rule.value;
    if (!ok) return null;
    return { fill: rule.color ?? '#fecaca', color: '#7f1d1d' };
  }
  if (!all.length) return null;
  const min = Math.min(...all);
  const max = Math.max(...all);
  if (rule.kind === 'bar') {
    const pct = max === min ? 100 : Math.round(((value - min) / (max - min)) * 100);
    return { barPct: Math.max(0, Math.min(100, pct)) };
  }
  if (rule.kind === 'scale') {
    const t = max === min ? 1 : (value - min) / (max - min);
    // зелёный -> жёлтый -> красный
    const r = Math.round(99 + t * (248 - 99));
    const g = Math.round(191 + t * (113 - 191));
    return { fill: `rgb(${r},${g},132)` };
  }
  if (rule.kind === 'icon') {
    const t = max === min ? 1 : (value - min) / (max - min);
    return { icon: t >= 0.66 ? '▲' : t >= 0.33 ? '●' : '▼', color: t >= 0.66 ? '#16a34a' : t >= 0.33 ? '#ca8a04' : '#dc2626' };
  }
  return null;
}
