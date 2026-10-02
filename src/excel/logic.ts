import { colToIndex, indexToCol, parseRef, shiftFormulaRefs } from './formula';
import type { Sheet } from './format';

export interface RangeBox {
  sheet: string | null;
  r0: number;
  c0: number;
  r1: number;
  c1: number;
}

const RANGE_RE = /^(?:('(?:[^']|'')+'|[A-Za-zА-Яа-яЁё_][\wА-Яа-яЁё.]*)!)?(\$?)([A-Za-z]+)(\$?)([0-9]+)(?::(\$?)([A-Za-z]+)(\$?)([0-9]+))?$/;

export function parseRange(s: string): RangeBox | null {
  const m = RANGE_RE.exec(s.trim());
  if (!m) return null;
  const sheet = m[1] ? m[1].replace(/^'|'$/g, '').replace(/''/g, "'") : null;
  const c0 = colToIndex(m[3]);
  const r0 = parseInt(m[5], 10) - 1;
  if (r0 < 0) return null;
  if (m[6] === undefined) return { sheet, r0, c0, r1: r0, c1: c0 };
  const c1 = colToIndex(m[7]);
  const r1 = parseInt(m[9], 10) - 1;
  if (r1 < 0) return null;
  return {
    sheet,
    r0: Math.min(r0, r1),
    c0: Math.min(c0, c1),
    r1: Math.max(r0, r1),
    c1: Math.max(c0, c1),
  };
}

export function rangeRefs(box: RangeBox): string[] {
  const out: string[] = [];
  for (let r = box.r0; r <= box.r1; r++) {
    for (let c = box.c0; c <= box.c1; c++) {
      if ((box.r1 - box.r0 + 1) * (box.c1 - box.c0 + 1) > 10000) return out;
      out.push(`${indexToCol(c)}${r + 1}`);
    }
  }
  return out;
}

// Те же ссылки, но с префиксом листа (для межлистовых спарклайнов/диаграмм)
export function rangeRefsFull(box: RangeBox): string[] {
  const prefix = box.sheet ? `${box.sheet}!` : '';
  return rangeRefs(box).map((r) => `${prefix}${r}`);
}

export function refKey(col: number, row: number): string {
  return `${indexToCol(col)}${row + 1}`;
}

// ---------- Сортировка строк диапазона по столбцу ----------

export function sortSheet(
  sheet: Sheet,
  box: RangeBox,
  sortCol: number,
  dir: 1 | -1,
): Sheet {
  const rows: { key: string; vals: (string | undefined)[] }[] = [];
  for (let r = box.r0; r <= box.r1; r++) {
    const vals: (string | undefined)[] = [];
    for (let c = box.c0; c <= box.c1; c++) {
      vals.push(sheet.cells[refKey(c, r)]);
    }
    rows.push({ key: String(r), vals });
  }
  const cmpVal = (v: string | undefined): number | string => {
    if (v == null || v === '') return '';
    const n = Number(v);
    return v.trim() !== '' && Number.isFinite(n) ? n : v.toLowerCase();
  };
  rows.sort((x, y) => {
    const a = cmpVal(x.vals[sortCol - box.c0]);
    const b = cmpVal(y.vals[sortCol - box.c0]);
    if (typeof a === 'number' && typeof b === 'number') return (a - b) * dir;
    if (a === '' && b !== '') return 1;
    if (b === '' && a !== '') return -1;
    return String(a).localeCompare(String(b), 'ru') * dir;
  });
  const cells = { ...sheet.cells };
  const styles = { ...sheet.styles };
  // стираем зону и пишем отсортированное (стили едут вместе со значениями)
  for (let r = box.r0; r <= box.r1; r++) {
    for (let c = box.c0; c <= box.c1; c++) {
      delete cells[refKey(c, r)];
      delete styles[refKey(c, r)];
    }
  }
  rows.forEach((row, i) => {
    const r = box.r0 + i;
    row.vals.forEach((v, j) => {
      const c = box.c0 + j;
      const fromR = box.r0 + Number(row.key);
      const fromRef = refKey(c, fromR);
      if (v != null && v !== '') cells[refKey(c, r)] = v;
      const st = sheet.styles[fromRef];
      if (st) styles[refKey(c, r)] = st;
    });
  });
  return { ...sheet, cells, styles };
}

// ---------- Мини-сводная: группировка по полю строк ----------

export interface Pivot {
  headers: string[];
  rows: (string | number)[][];
}

export function buildPivot(
  data: (string | undefined)[][],
  rowField: number,
  valField: number,
  agg: 'sum' | 'count' | 'avg',
): Pivot {
  const groups = new Map<string, number[]>();
  for (const row of data) {
    const k = row[rowField] ?? '(пусто)';
    const raw = row[valField];
    const n = raw != null && raw !== '' && Number.isFinite(Number(raw)) ? Number(raw) : NaN;
    if (!groups.has(k)) groups.set(k, []);
    if (!Number.isNaN(n)) groups.get(k)!.push(n);
    else if (agg === 'count') groups.get(k)!.push(1);
  }
  const aggName = agg === 'sum' ? 'Сумма' : agg === 'count' ? 'Количество' : 'Среднее';
  const out: Pivot = { headers: ['Строка', aggName], rows: [] };
  for (const [k, vals] of groups) {
    const v =
      agg === 'count'
        ? vals.length
        : vals.length
          ? agg === 'sum'
            ? vals.reduce((a, b) => a + b, 0)
            : vals.reduce((a, b) => a + b, 0) / vals.length
          : 0;
    out.rows.push([k, Math.round(v * 100) / 100]);
  }
  out.rows.push(['ИТОГО', agg === 'avg' ? '' : out.rows.reduce((a, r) => a + (typeof r[1] === 'number' ? r[1] : 0), 0)]);
  return out;
}

// ---------- Вставка/удаление строк/столбцов со сдвигом ссылок ----------

function shiftKey(ref: string, type: 'row' | 'col', index: number, delta: number): string | null {
  const p = parseRef(ref);
  if (!p) return ref;
  // ячейки внутри удалённого диапазона пропадают (null), остальные сдвигаются
  if (type === 'row' && p.row >= index) {
    if (delta < 0 && p.row < index - delta) return null;
    const nr = p.row + delta;
    if (nr < 0) return null;
    return `${indexToCol(p.col)}${nr + 1}`;
  }
  if (type === 'col' && p.col >= index) {
    if (delta < 0 && p.col < index - delta) return null;
    const nc = p.col + delta;
    if (nc < 0) return null;
    return `${indexToCol(nc)}${p.row + 1}`;
  }
  return ref;
}

export function insertDelete(
  sheets: Sheet[],
  sheetIdx: number,
  type: 'row' | 'col',
  index: number,
  delta: number, // +1 вставка, -1 удаление
): Sheet[] {
  return sheets.map((sh, si) => {
    const mapRecord = <T>(rec: Record<string, T>, dropDeleted: boolean): Record<string, T> => {
      const out: Record<string, T> = {};
      for (const [k, v] of Object.entries(rec)) {
        if (si === sheetIdx) {
          const nk = shiftKey(k, type, index, delta);
          if (nk === null) {
            if (!dropDeleted) out[k] = v;
            continue;
          }
          out[nk] = v;
        } else out[k] = v;
      }
      return out;
    };
    const sheetName = si === sheetIdx ? sh.name : null;
    const shiftFormulas = (cells: Record<string, string>): Record<string, string> => {
      const out: Record<string, string> = {};
      for (const [k, v] of Object.entries(cells)) {
        out[k] = v.startsWith('=') ? shiftFormulaRefs(v, { type, sheet: sheetName, index, delta }) : v;
      }
      return out;
    };
    if (si !== sheetIdx) {
      // чужие листы: только правим ссылки на изменённый лист
      const target = sheets[sheetIdx].name;
      const out: Record<string, string> = {};
      for (const [k, v] of Object.entries(sh.cells)) {
        out[k] = v.startsWith('=') ? shiftFormulaRefs(v, { type, sheet: target, index, delta }) : v;
      }
      return { ...sh, cells: out };
    }
    const cells = shiftFormulas(mapRecord(sh.cells, true));
    const styles = mapRecord(sh.styles, false);
    const comments = mapRecord(sh.comments, false);
    const validations = mapRecord(sh.validations, false);
    // объединения, пересекающие границу вставки, проще снять
    const merges = sh.merges.filter((m) => {
      const b = parseRange(m);
      if (!b) return false;
      if (type === 'row') return b.r1 < index || b.r0 > index + (delta > 0 ? 0 : -delta - 0);
      return b.c1 < index || b.c0 > index;
    });
    return { ...sh, cells, styles, comments, validations, merges };
  });
}

// ---------- Проверка данных ----------

export function validateCell(value: string, rule?: { type: string; values?: string[]; min?: number; max?: number }): string | null {
  if (!rule) return null;
  if (value === '') return null;
  if (rule.type === 'list') {
    if (!(rule.values ?? []).includes(value)) return `Допустимы: ${(rule.values ?? []).join(', ')}`;
    return null;
  }
  const n = Number(value);
  if (!Number.isFinite(n) || value.trim() === '') return 'Нужно число';
  if (rule.type === 'whole' && !Number.isInteger(n)) return 'Нужно целое число';
  if (rule.min !== undefined && n < rule.min) return `Минимум ${rule.min}`;
  if (rule.max !== undefined && n > rule.max) return `Максимум ${rule.max}`;
  return null;
}

// ---------- Объединения ----------

export function mergeAt(merges: string[], ref: string): RangeBox | null {
  for (const m of merges) {
    const b = parseRange(m);
    if (b && refInBox(b, ref)) return b;
  }
  return null;
}

export function refInBox(b: RangeBox, ref: string): boolean {
  const p = parseRef(ref);
  if (!p) return false;
  return p.row >= b.r0 && p.row <= b.r1 && p.col >= b.c0 && p.col <= b.c1;
}

export function isMergeHead(merges: string[], ref: string): boolean {
  const b = mergeAt(merges, ref);
  if (!b) return false;
  return ref === refKey(b.c0, b.r0);
}
