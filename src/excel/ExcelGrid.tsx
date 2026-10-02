import { useEffect, useMemo, useRef, useState } from 'react';
import { evaluateFormula, indexToCol, parseRef } from './formula';
import {
  newSheet,
  formatValue,
  assessCF,
  NUM_FMT_LABEL,
  type Sheet,
  type NumFmt,
  type CFKind,
} from './format';
import {
  parseRange,
  rangeRefs,
  refKey,
  sortSheet,
  buildPivot,
  insertDelete,
  validateCell,
  mergeAt,
} from './logic';
import { exportSheetsToXlsx, importXlsx } from './xlsx';
import Sparkline from './Sparkline';
import ChartView from './ChartView';

const ROWS = 500;
const COLS = 52; // A..AZ
const V1_KEY = 'opendesksuite.excel.v1';
const V2_KEY = 'opendesksuite.excel.v2';
const VER_KEY = 'opendesksuite.excel.versions.v1';
const RENDER_STEP = 60;

type Ribbon = 'home' | 'insert' | 'data' | 'review' | 'view' | 'file';

function loadSheets(): { sheets: Sheet[]; active: number } {
  try {
    const v2 = localStorage.getItem(V2_KEY);
    if (v2) {
      const p = JSON.parse(v2) as { sheets: Sheet[]; active: number };
      if (p.sheets?.length) return { sheets: p.sheets, active: Math.min(p.active ?? 0, p.sheets.length - 1) };
    }
    const v1 = localStorage.getItem(V1_KEY);
    if (v1) {
      const cells = JSON.parse(v1) as Record<string, string>;
      return { sheets: [{ ...newSheet('Лист1'), cells }], active: 0 };
    }
  } catch {
    /* ignore */
  }
  return { sheets: [newSheet('Лист1')], active: 0 };
}

function usedBounds(sheet: Sheet): { maxR: number; maxC: number } {
  let maxR = 0;
  let maxC = 0;
  for (const k of Object.keys(sheet.cells)) {
    const p = parseRef(k);
    if (p) {
      maxR = Math.max(maxR, p.row);
      maxC = Math.max(maxC, p.col);
    }
  }
  return { maxR, maxC };
}

export default function ExcelGrid() {
  const [sheets, setSheets] = useState<Sheet[]>(() => loadSheets().sheets);
  const [active, setActive] = useState(() => loadSheets().active);
  const [ribbon, setRibbon] = useState<Ribbon>('home');
  const [sel, setSel] = useState('A1');
  const [draft, setDraft] = useState<string | null>(null);
  const [renderRows, setRenderRows] = useState(RENDER_STEP);
  const [findText, setFindText] = useState('');
  const [found, setFound] = useState<string[]>([]);
  const [fileName, setFileName] = useState('table');
  const [author, setAuthor] = useState('');
  const [gridLines, setGridLines] = useState(true);
  const [showHeaders, setShowHeaders] = useState(true);
  const [frzRows, setFrzRows] = useState(0);
  const [frzCols, setFrzCols] = useState(0);
  const [zoom, setZoom] = useState(100);
  const [versions, setVersions] = useState<{ t: number; data: string }[]>(() => {
    try {
      return JSON.parse(localStorage.getItem(VER_KEY) ?? '[]');
    } catch {
      return [];
    }
  });
  const [valError, setValError] = useState('');
  const tableRef = useRef<HTMLTableElement>(null);
  const [measure, setMeasure] = useState({ hdrH: 37, rowHdrW: 44, rowH: 35, colW: 100 });
  const csvInput = useRef<HTMLInputElement>(null);
  const xlsxInput = useRef<HTMLInputElement>(null);

  const sheet = sheets[active] ?? newSheet('Лист1');

  useEffect(() => {
    localStorage.setItem(V2_KEY, JSON.stringify({ sheets, active }));
  }, [sheets, active]);

  // замер размеров для закрепления областей
  useEffect(() => {
    const t = tableRef.current;
    if (!t) return;
    const thead = t.querySelector('thead');
    const firstTd = t.querySelector<HTMLTableCellElement>('tbody td[data-c]');
    const rowHdr = t.querySelector<HTMLTableCellElement>('tbody th.row-h');
    setMeasure({
      hdrH: thead?.getBoundingClientRect().height ?? 37,
      rowHdrW: rowHdr?.getBoundingClientRect().width ?? 44,
      rowH: firstTd?.getBoundingClientRect().height ?? 35,
      colW: firstTd?.getBoundingClientRect().width ?? 100,
    });
  }, [frzRows, frzCols, renderRows, active, zoom, gridLines, showHeaders]);

  const upd = (fn: (s: Sheet) => Sheet) => {
    setSheets((prev) => prev.map((s, i) => (i === active ? fn(s) : s)));
  };

  // ---------- вычисление ----------

  const getRawEngine = (ref: string): string | undefined => {
    const bang = ref.indexOf('!');
    if (bang >= 0) {
      const sh = sheets.find((s) => s.name.toUpperCase() === ref.slice(0, bang).toUpperCase());
      return sh?.cells[ref.slice(bang + 1)];
    }
    return sheet.cells[ref];
  };

  const computed = useMemo(() => {
    const out: Record<string, number | string> = {};
    for (const [k, v] of Object.entries(sheet.cells)) {
      if (v.startsWith('=')) {
        out[k] = evaluateFormula(v, getRawEngine);
      } else if ((sheet.styles[k]?.numFmt ?? 'general') !== 'text' && v.trim() !== '' && Number.isFinite(Number(v))) {
        out[k] = Number(v);
      } else {
        out[k] = v;
      }
    }
    return out;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [sheets, active]);

  const getNum = (ref: string): number => {
    const plain = ref.includes('!') ? undefined : computed[ref];
    const v = plain !== undefined ? plain : undefined;
    if (typeof v === 'number') return v;
    const raw = getRawEngine(ref.toUpperCase()) ?? '';
    const n = Number(raw);
    return raw.trim() !== '' && Number.isFinite(n) ? n : NaN;
  };

  const getText = (ref: string): string => {
    const v = ref.includes('!') ? undefined : computed[ref];
    if (v !== undefined) return String(v);
    return getRawEngine(ref.toUpperCase()) ?? '';
  };

  const disp = (ref: string): string => {
    const v = computed[ref] ?? sheet.cells[ref] ?? '';
    return formatValue(v, sheet.styles[ref]);
  };

  // ---------- правки ----------

  const commit = (ref: string, value: string) => {
    if (sheet.protected) {
      alert('Лист защищён (Рецензирование → снять защиту).');
      return;
    }
    const err = validateCell(value, sheet.validations[ref]);
    if (err) {
      setValError(`${ref}: ${err}`);
      return;
    }
    setValError('');
    upd((s) => {
      const cells = { ...s.cells };
      if (value === '') delete cells[ref];
      else cells[ref] = value;
      return { ...s, cells };
    });
  };

  const setStyle = (patch: Partial<Sheet['styles'][string]>, toggle?: 'bold' | 'italic' | 'underline') => {
    upd((s) => {
      const cur = s.styles[sel] ?? {};
      const next = { ...cur, ...patch };
      if (toggle) next[toggle] = !cur[toggle];
      const styles = { ...s.styles };
      if (Object.values(next).every((v) => v === undefined || v === false) && !patch.numFmt) delete styles[sel];
      else styles[sel] = next;
      return { ...s, styles };
    });
  };

  // ---------- Автосумма ----------

  const autoSum = () => {
    const p = parseRef(sel);
    if (!p) return;
    const hasNum = (ref: string) => Number.isFinite(getNum(ref));
    // вверх
    let r = p.row - 1;
    const up: string[] = [];
    while (r >= 0 && hasNum(refKey(p.col, r))) {
      up.unshift(refKey(p.col, r));
      r--;
    }
    if (up.length) {
      commit(sel, `=SUM(${up[0]}:${up[up.length - 1]})`);
      return;
    }
    // влево
    let c = p.col - 1;
    const lf: string[] = [];
    while (c >= 0 && hasNum(refKey(c, p.row))) {
      lf.unshift(refKey(c, p.row));
      c--;
    }
    if (lf.length) {
      commit(sel, `=SUM(${lf[0]}:${lf[lf.length - 1]})`);
      return;
    }
    const range = prompt('Диапазон для СУММ (напр. A1:A10):', 'A1:A10');
    if (range) commit(sel, `=SUM(${range})`);
  };

  // ---------- Сортировка ----------

  const sortBySel = (dir: 1 | -1) => {
    const p = parseRef(sel);
    if (!p) return;
    const { maxR, maxC } = usedBounds(sheet);
    const withHeader = maxR > 0 && confirm('Первая строка — заголовок? (ОК — да)');
    const box = { sheet: null, r0: withHeader ? 1 : 0, c0: 0, r1: maxR, c1: Math.max(maxC, p.col) };
    upd((s) => sortSheet(s, box, p.col, dir));
  };

  // ---------- Фильтр ----------

  const { maxR: usedMaxR, maxC: usedMaxC } = usedBounds(sheet);
  const filterVals = useMemo(() => {
    if (sheet.filterCol == null) return [];
    const set = new Set<string>();
    for (let r = 0; r <= Math.min(usedMaxR, ROWS - 1); r++) {
      set.add(String(computed[refKey(sheet.filterCol, r)] ?? sheet.cells[refKey(sheet.filterCol, r)] ?? ''));
    }
    return Array.from(set).sort((a, b) => a.localeCompare(b, 'ru'));
  }, [sheet.filterCol, sheet.cells, computed, usedMaxR]);

  const rowHidden = (r: number): boolean => {
    if (sheet.filterCol == null || sheet.filterValues.length === 0) return false;
    const v = String(computed[refKey(sheet.filterCol, r)] ?? sheet.cells[refKey(sheet.filterCol, r)] ?? '');
    return !sheet.filterValues.includes(v);
  };

  // ---------- Поиск ----------

  const doFind = () => {
    if (!findText) return;
    const out: string[] = [];
    for (const [k, v] of Object.entries(sheet.cells)) {
      const show = String(computed[k] ?? v);
      if (v.toLowerCase().includes(findText.toLowerCase()) || show.toLowerCase().includes(findText.toLowerCase())) {
        out.push(k);
        if (out.length >= 50) break;
      }
    }
    setFound(out);
  };

  const doReplace = () => {
    const rep = prompt('Заменить на:') ?? '';
    upd((s) => {
      const cells = { ...s.cells };
      for (const k of Object.keys(cells)) {
        if (cells[k].includes(findText)) cells[k] = cells[k].split(findText).join(rep);
      }
      return { ...s, cells };
    });
    setFound([]);
  };

  // ---------- Файл ----------

  const toCSV = (): string => {
    const lines: string[] = [];
    for (let r = 0; r < ROWS; r++) {
      let last = -1;
      const row: string[] = [];
      for (let c = 0; c < COLS; c++) {
        const v = sheet.cells[refKey(c, r)] ?? '';
        if (v !== '') last = c;
        row.push(v);
      }
      if (last >= 0) {
        lines.push(
          row
            .slice(0, last + 1)
            .map((v) => (/[",;\n]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v))
            .join(','),
        );
      }
      if (r > usedMaxR + 1 && last < 0) break;
    }
    return lines.join('\n');
  };

  const download = (name: string, content: Blob | string, mime = 'text/plain') => {
    const blob = typeof content === 'string' ? new Blob([content], { type: `${mime};charset=utf-8` }) : content;
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = name;
    a.click();
    URL.revokeObjectURL(url);
  };

  const importCSV = (file: File) => {
    const reader = new FileReader();
    reader.onload = () => {
      const text = String(reader.result ?? '');
      const cells: Record<string, string> = {};
      text.split(/\r?\n/).forEach((line, r) => {
        if (r >= ROWS) return;
        const vals: string[] = [];
        let cur = '';
        let q = false;
        for (let i = 0; i < line.length; i++) {
          const ch = line[i];
          if (q) {
            if (ch === '"' && line[i + 1] === '"') {
              cur += '"';
              i++;
            } else if (ch === '"') q = false;
            else cur += ch;
          } else if (ch === '"') q = true;
          else if (ch === ',' || ch === ';') {
            vals.push(cur);
            cur = '';
          } else cur += ch;
        }
        vals.push(cur);
        vals.forEach((v, c) => {
          if (c < COLS && v !== '') cells[refKey(c, r)] = v;
        });
      });
      upd((s) => ({ ...s, cells }));
    };
    reader.readAsText(file);
  };

  const openXlsx = async (f: File) => {
    const imp = await importXlsx(f);
    if (!imp.length) {
      alert('Листы не найдены');
      return;
    }
    setSheets(
      imp.map((sh) => ({
        ...newSheet(sh.name || 'Лист'),
        cells: sh.cells,
        styles: Object.fromEntries(sh.dateCells.map((r) => [r, { numFmt: 'date' as const }])),
      })),
    );
    setActive(0);
    setRenderRows(RENDER_STEP);
  };

  const snapshot = () => {
    const next = [{ t: Date.now(), data: JSON.stringify(sheets).slice(0, 500000) }, ...versions].slice(0, 5);
    setVersions(next);
    localStorage.setItem(VER_KEY, JSON.stringify(next));
  };

  // ---------- Сводная ----------

  const [pivotRange, setPivotRange] = useState('A1:C20');
  const [pivotRow, setPivotRow] = useState(0);
  const [pivotVal, setPivotVal] = useState(1);
  const [pivotAgg, setPivotAgg] = useState<'sum' | 'count' | 'avg'>('sum');

  const runPivot = () => {
    const box = parseRange(pivotRange);
    if (!box) {
      alert('Неверный диапазон');
      return;
    }
    const data: (string | undefined)[][] = [];
    for (let r = box.r0; r <= box.r1; r++) {
      const row: (string | undefined)[] = [];
      for (let c = box.c0; c <= box.c1; c++) row.push(sheet.cells[refKey(c, r)]);
      data.push(row);
    }
    const headers = data[0] ?? [];
    const piv = buildPivot(data.slice(1), pivotRow, pivotVal, pivotAgg);
    const cells: Record<string, string> = {};
    piv.headers.forEach((h, i) => {
      cells[refKey(i, 0)] = h;
    });
    piv.rows.forEach((row, r) => {
      row.forEach((v, c) => {
        cells[refKey(c, r + 1)] = String(v);
      });
    });
    let n = 1;
    while (sheets.some((s) => s.name === `Сводная${n}`)) n++;
    setSheets((prev) => [...prev, { ...newSheet(`Сводная${n}`), cells }]);
    setActive(sheets.length);
  };

  // ---------- Текст-по-столбцам / дубликаты ----------

  const textToCols = () => {
    const p = parseRef(sel);
    if (!p) return;
    const delim = prompt('Разделитель:', ';') ?? ';';
    upd((s) => {
      const cells = { ...s.cells };
      for (let r = 0; r <= usedMaxR; r++) {
        const v = cells[refKey(p.col, r)];
        if (v == null || v.startsWith('=')) continue;
        const parts = v.split(delim);
        parts.forEach((part, i) => {
          if (part !== '') cells[refKey(p.col + i, r)] = part;
        });
      }
      return { ...s, cells };
    });
  };

  const removeDupes = () => {
    const range = prompt('Диапазон (напр. A1:C50):', `A1:${indexToCol(Math.max(usedMaxC, 0))}${usedMaxR + 1}`);
    if (!range) return;
    const box = parseRange(range);
    if (!box) return;
    const seen = new Set<string>();
    upd((s) => {
      const cells = { ...s.cells };
      for (let r = box.r0; r <= box.r1; r++) {
        const sig: string[] = [];
        for (let c = box.c0; c <= box.c1; c++) sig.push(cells[refKey(c, r)] ?? '');
        const key = sig.join('');
        if (seen.has(key)) {
          for (let c = box.c0; c <= box.c1; c++) delete cells[refKey(c, r)];
        } else seen.add(key);
      }
      return { ...s, cells };
    });
  };

  // ---------- Рендер сетки ----------

  const cols = useMemo(() => Array.from({ length: COLS }, (_, i) => indexToCol(i)), []);
  const rows = useMemo(() => {
    const out: number[] = [];
    for (let r = 1; r <= Math.min(renderRows, ROWS); r++) {
      if (!rowHidden(r - 1)) out.push(r);
    }
    return out;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [renderRows, sheet.filterCol, sheet.filterValues, computed, sheet.cells]);

  const cfNums = useMemo(() => {
    // числовые значения по диапазонам УФ для шкал/баров/значков
    const map = new Map<number, number[]>();
    sheet.cf.forEach((rule) => {
      if (rule.kind === 'cellIs') return;
      const box = parseRange(rule.range);
      if (!box) return;
      const vals: number[] = [];
      for (const ref of rangeRefs(box)) {
        const v = computed[ref];
        const n = typeof v === 'number' ? v : Number(sheet.cells[ref]);
        if (Number.isFinite(n) && (sheet.cells[ref] ?? '').trim() !== '') vals.push(n);
      }
      map.set(rule.id, vals);
    });
    return map;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [sheet.cf, computed, sheet.cells]);

  const cellCF = (ref: string): { fill?: string; color?: string; barPct?: number; icon?: string } | null => {
    const raw = computed[ref] ?? sheet.cells[ref] ?? '';
    const v = typeof raw === 'number' ? raw : Number(raw);
    if (!Number.isFinite(v) || String(raw).trim() === '') return null;
    for (const rule of sheet.cf) {
      const box = parseRange(rule.range);
      if (!box) continue;
      const p = parseRef(ref);
      if (!p || p.row < box.r0 || p.row > box.r1 || p.col < box.c0 || p.col > box.c1) continue;
      const r = assessCF(v, cfNums.get(rule.id) ?? [], rule);
      if (r) return { ...r, icon: r.icon };
    }
    return null;
  };

  const frozenStyle = (r: number, c: number, isHead: boolean): React.CSSProperties => {
    const st: React.CSSProperties = {};
    if (frzRows > 0 && !isHead && r < frzRows) {
      st.top = measure.hdrH + r * measure.rowH;
      st.zIndex = 3;
      st.background = '#fff';
    }
    if (frzCols > 0 && !isHead && c < frzCols) {
      st.left = measure.rowHdrW + c * measure.colW;
      st.zIndex = Number(st.zIndex ?? 1) + 1;
      st.background = '#fff';
    }
    return st;
  };

  const ribbonBtn = (id: Ribbon, label: string) => (
    <button key={id} className={ribbon === id ? 'tab active' : 'tab'} onClick={() => setRibbon(id)}>
      {label}
    </button>
  );

  const zw = zoom / 100;

  return (
    <div>
      <div className="tabs" style={{ marginBottom: 8 }}>
        {ribbonBtn('home', 'Главная')}
        {ribbonBtn('insert', 'Вставка')}
        {ribbonBtn('data', 'Данные')}
        {ribbonBtn('review', 'Рецензирование')}
        {ribbonBtn('view', 'Вид')}
        {ribbonBtn('file', 'Файл')}
      </div>

      {ribbon === 'home' && (
        <div className="toolbar">
          <button onClick={() => setStyle({}, 'bold')} title="Жирный">
            <b>B</b>
          </button>
          <button onClick={() => setStyle({}, 'italic')} title="Курсив">
            <i>I</i>
          </button>
          <button onClick={() => setStyle({}, 'underline')} title="Подчёркнутый">
            <u>U</u>
          </button>
          <select title="Шрифт" defaultValue="" onChange={(e) => e.target.value && setStyle({ font: e.target.value })}>
            <option value="">Шрифт…</option>
            {['Calibri', 'Arial', 'Times New Roman', 'Courier New', 'Verdana'].map((f) => (
              <option key={f} value={f}>
                {f}
              </option>
            ))}
          </select>
          <select title="Кегль" defaultValue="" onChange={(e) => e.target.value && setStyle({ size: Number(e.target.value) })}>
            <option value="">10…</option>
            {[10, 11, 12, 14, 16, 18, 20].map((s) => (
              <option key={s} value={s}>
                {s}
              </option>
            ))}
          </select>
          <label title="Цвет текста">
            🎨
            <input type="color" hidden onChange={(e) => setStyle({ color: e.target.value })} />
          </label>
          <label title="Заливка">
            🪣
            <input type="color" hidden defaultValue="#ffff00" onChange={(e) => setStyle({ fill: e.target.value })} />
          </label>
          <button onClick={() => setStyle({ align: 'left' })}>⬅</button>
          <button onClick={() => setStyle({ align: 'center' })}>⬌</button>
          <button onClick={() => setStyle({ align: 'right' })}>➡</button>
          <button
            title="Границы"
            onClick={() =>
              upd((s) => {
                const cur = s.styles[sel]?.border;
                const styles = { ...s.styles, [sel]: { ...s.styles[sel], border: !cur } };
                return { ...s, styles };
              })
            }
          >
            ▦
          </button>
          <select
            title="Числовой формат"
            value={sheet.styles[sel]?.numFmt ?? 'general'}
            onChange={(e) => setStyle({ numFmt: e.target.value as NumFmt })}
          >
            {(Object.keys(NUM_FMT_LABEL) as NumFmt[]).map((f) => (
              <option key={f} value={f}>
                {NUM_FMT_LABEL[f]}
              </option>
            ))}
          </select>
          <button
            onClick={() => {
              const range = prompt('Объединить диапазон (напр. B2:C4):', `${sel}:${sel}`);
              if (!range) return;
              const box = parseRange(range);
              if (!box || (box.r0 === box.r1 && box.c0 === box.c1)) {
                alert('Нужен диапазон из нескольких ячеек');
                return;
              }
              upd((s) => ({ ...s, merges: [...s.merges, `${indexToCol(box.c0)}${box.r0 + 1}:${indexToCol(box.c1)}${box.r1 + 1}`] }));
            }}
          >
            ⛶ Объединить
          </button>
          <button
            onClick={() =>
              upd((s) => ({ ...s, merges: s.merges.filter((m) => { const b = parseRange(m); return !b || !(parseRef(sel) && b.r0 <= (parseRef(sel)?.row ?? -1) && b.r1 >= (parseRef(sel)?.row ?? -1) && b.c0 <= (parseRef(sel)?.col ?? -1) && b.c1 >= (parseRef(sel)?.col ?? -1)); }) }))
            }
          >
            🗗 Разъединить
          </button>
          <button onClick={autoSum} title="Автосумма">
            Σ
          </button>
          <button onClick={() => sortBySel(1)} title="Сортировка А→Я">
            А↓Я
          </button>
          <button onClick={() => sortBySel(-1)} title="Сортировка Я→А">
            Я↓А
          </button>
          <button
            onClick={() => {
              const p = parseRef(sel);
              if (!p) return;
              setSheets((prev) => insertDelete(prev, active, 'row', p.row, 1));
            }}
          >
            ＋Строка
          </button>
          <button
            onClick={() => {
              const p = parseRef(sel);
              if (!p) return;
              if (!confirm(`Удалить строку ${p.row + 1}?`)) return;
              setSheets((prev) => insertDelete(prev, active, 'row', p.row, -1));
            }}
          >
            －Строка
          </button>
          <button
            onClick={() => {
              const p = parseRef(sel);
              if (!p) return;
              setSheets((prev) => insertDelete(prev, active, 'col', p.col, 1));
            }}
          >
            ＋Столбец
          </button>
          <button
            onClick={() => {
              const p = parseRef(sel);
              if (!p) return;
              if (!confirm(`Удалить столбец ${indexToCol(p.col)}?`)) return;
              setSheets((prev) => insertDelete(prev, active, 'col', p.col, -1));
            }}
          >
            －Столбец
          </button>
          <input placeholder="Найти…" value={findText} onChange={(e) => setFindText(e.target.value)} style={{ width: 90 }} />
          <button onClick={doFind}>🔍{found.length ? ` (${found.length})` : ''}</button>
          <button onClick={doReplace}>⇄ Заменить</button>
        </div>
      )}

      {ribbon === 'insert' && (
        <div className="toolbar">
          <button
            onClick={() => {
              const range = prompt('Диапазон данных для спарклайна:', 'A1:A10');
              if (!range || !parseRange(range)) return;
              const type = (prompt('Тип (line/column/winloss):', 'line') ?? 'line') as 'line' | 'column' | 'winloss';
              upd((s) => ({ ...s, sparks: [...s.sparks.filter((x) => x.at !== sel), { at: sel, range, type }] }));
            }}
          >
            📈 Спарклайн
          </button>
          <button
            onClick={() => {
              const range = prompt('Диапазон значений:', 'A1:A8');
              if (!range || !parseRange(range)) return;
              const type = (prompt('Тип (bar/line/pie):', 'bar') ?? 'bar') as 'bar' | 'line' | 'pie';
              const title = prompt('Название:', 'Диаграмма') ?? 'Диаграмма';
              const labels = prompt('Диапазон подписей (пусто — нет):', '') || undefined;
              upd((s) => ({
                ...s,
                charts: [...s.charts, { id: Date.now(), title, type, range, labels }],
              }));
            }}
          >
            📊 Диаграмма
          </button>
          <button
            onClick={() => {
              const text = prompt('Текст примечания:') ?? '';
              if (!text) return;
              upd((s) => ({ ...s, comments: { ...s.comments, [sel]: { author: author || 'Автор', text } } }));
            }}
          >
            💬 Примечание
          </button>
          <button
            onClick={() => {
              const vals = prompt('Список через запятую:', 'да,нет') ?? '';
              upd((s) => ({
                ...s,
                validations: { ...s.validations, [sel]: { type: 'list', values: vals.split(',').map((x) => x.trim()) } },
              }));
            }}
          >
            📋 Список
          </button>
          <button
            onClick={() => {
              const min = Number(prompt('Минимум:', '0'));
              const max = Number(prompt('Максимум:', '100'));
              upd((s) => ({ ...s, validations: { ...s.validations, [sel]: { type: 'decimal', min, max } } }));
            }}
          >
            🔢 Диапазон
          </button>
          <select
            defaultValue=""
            title="Условное форматирование"
            onChange={(e) => {
              const kind = e.target.value as CFKind | '';
              e.target.value = '';
              if (!kind) return;
              const range = prompt('Диапазон (напр. B2:B20):', `${sel}:${sel}`) ?? sel;
              if (kind === 'cellIs') {
                const op = (prompt('Оператор (>,<,=,>=,<=):', '>') ?? '>') as '>' | '<' | '=' | '>=' | '<=';
                const value = Number(prompt('Значение:', '0') ?? '0');
                upd((s) => ({ ...s, cf: [...s.cf, { id: Date.now(), range, kind, op, value }] }));
              } else {
                upd((s) => ({ ...s, cf: [...s.cf, { id: Date.now(), range, kind }] }));
              }
            }}
          >
            <option value="">УФ…</option>
            <option value="scale">Цветовая шкала</option>
            <option value="bar">Гистограмма</option>
            <option value="icon">Набор значков</option>
            <option value="cellIs">Значение ячейки…</option>
          </select>
          <button onClick={() => upd((s) => ({ ...s, cf: [] }))}>✕ Очистить УФ</button>
        </div>
      )}

      {ribbon === 'data' && (
        <div className="toolbar">
          <button
            onClick={() =>
              upd((s) => {
                const p = parseRef(sel);
                if (p && s.filterCol === p.col) return { ...s, filterCol: null, filterValues: [] };
                return p ? { ...s, filterCol: p.col, filterValues: [] } : s;
              })
            }
          >
            {sheet.filterCol == null ? '🔽 Фильтр' : '✕ Убрать фильтр'}
          </button>
          <button onClick={textToCols}>⇄ Текст-по-столбцам</button>
          <button onClick={removeDupes}>⧉ Убрать дубликаты</button>
          <button
            onClick={() => {
              const range = prompt('Диапазон сводной (первая строка — заголовки):', `A1:${indexToCol(Math.max(usedMaxC, 1))}${usedMaxR + 1}`);
              if (range) setPivotRange(range);
              const box = parseRange(range ?? '');
              if (box) {
                setPivotRow(0);
                setPivotVal(Math.min(1, box.c1 - box.c0));
              }
            }}
          >
            🧮 Сводная…
          </button>
        </div>
      )}

      {ribbon === 'review' && (
        <div className="toolbar">
          <input placeholder="Ваше имя…" value={author} onChange={(e) => setAuthor(e.target.value)} style={{ width: 110 }} />
          <button
            onClick={() => {
              if (sheet.protected) {
                upd((s) => ({ ...s, protected: false }));
              } else {
                // защита локальная (уровень интерфейса), не шифрование файла
                upd((s) => ({ ...s, protected: true }));
              }
            }}
          >
            {sheet.protected ? '🔓 Снять защиту' : '🔒 Защита листа'}
          </button>
          <button onClick={snapshot}>🕘 Версия</button>
        </div>
      )}

      {ribbon === 'view' && (
        <div className="toolbar">
          <select
            value={frzRows > 0 ? (frzCols > 0 ? 'both' : 'row') : frzCols > 0 ? 'col' : 'none'}
            onChange={(e) => {
              const v = e.target.value;
              setFrzRows(v === 'row' || v === 'both' ? 1 : 0);
              setFrzCols(v === 'col' || v === 'both' ? 1 : 0);
            }}
          >
            <option value="none">Без закрепления</option>
            <option value="row">Закрепить строку</option>
            <option value="col">Закрепить столбец</option>
            <option value="both">Строку и столбец</option>
          </select>
          <label>
            Строк:{' '}
            <input
              type="number"
              min={0}
              max={10}
              value={frzRows}
              onChange={(e) => setFrzRows(Math.max(0, Math.min(10, Number(e.target.value))))}
              style={{ width: 50 }}
            />
          </label>
          <label>
            Столбцов:{' '}
            <input
              type="number"
              min={0}
              max={5}
              value={frzCols}
              onChange={(e) => setFrzCols(Math.max(0, Math.min(5, Number(e.target.value))))}
              style={{ width: 50 }}
            />
          </label>
          <label>
            <input type="checkbox" checked={gridLines} onChange={(e) => setGridLines(e.target.checked)} /> Сетка
          </label>
          <label>
            <input type="checkbox" checked={showHeaders} onChange={(e) => setShowHeaders(e.target.checked)} /> Заголовки
          </label>
          <label>
            Масштаб{' '}
            <input type="range" min={70} max={150} value={zoom} onChange={(e) => setZoom(Number(e.target.value))} /> {zoom}%
          </label>
          <button
            onClick={() => {
              setRenderRows(Math.max(renderRows, usedMaxR + 5));
              setTimeout(() => window.print(), 150);
            }}
          >
            🖨 Печать
          </button>
        </div>
      )}

      {ribbon === 'file' && (
        <div className="toolbar">
          <button
            onClick={() => {
              if (!confirm('Новая книга? Текущая будет очищена.')) return;
              setSheets([newSheet('Лист1')]);
              setActive(0);
            }}
          >
            📄 Новая
          </button>
          <button onClick={() => csvInput.current?.click()}>📂 CSV</button>
          <button onClick={() => xlsxInput.current?.click()}>📂 XLSX</button>
          <button onClick={() => download(`${fileName || 'sheet'}.csv`, toCSV(), 'text/csv')}>⬇ CSV</button>
          <button
            onClick={async () => {
              const blob = await exportSheetsToXlsx(sheets, fileName);
              download(`${fileName || 'sheet'}.xlsx`, blob);
            }}
          >
            ⬇ XLSX
          </button>
          <button onClick={snapshot}>🕘 Версия</button>
          <input
            ref={csvInput}
            type="file"
            accept=".csv,.txt"
            hidden
            onChange={(e) => {
              const f = e.target.files?.[0];
              if (f) importCSV(f);
              e.target.value = '';
            }}
          />
          <input
            ref={xlsxInput}
            type="file"
            accept=".xlsx"
            hidden
            onChange={(e) => {
              const f = e.target.files?.[0];
              if (f) openXlsx(f);
              e.target.value = '';
            }}
          />
          <input placeholder="Имя файла…" value={fileName} onChange={(e) => setFileName(e.target.value)} style={{ width: 110 }} />
        </div>
      )}

      {ribbon === 'file' && versions.length > 0 && (
        <div className="card">
          <b>Версии ({versions.length}/5)</b>
          {versions.map((v) => (
            <div key={v.t} style={{ display: 'flex', gap: 8, alignItems: 'center', marginTop: 4 }}>
              <span className="hint">{new Date(v.t).toLocaleString('ru')}</span>
              <button
                className="tab"
                onClick={() => {
                  if (!confirm('Восстановить версию?')) return;
                  try {
                    setSheets(JSON.parse(v.data) as Sheet[]);
                    setActive(0);
                  } catch {
                    alert('Не удалось прочитать версию');
                  }
                }}
              >
                Восстановить
              </button>
            </div>
          ))}
        </div>
      )}

      {ribbon === 'data' && (
        <div className="card">
          <b>Сводная таблица</b>
          <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', marginTop: 8 }}>
            <input value={pivotRange} onChange={(e) => setPivotRange(e.target.value)} style={{ width: 130 }} title="Диапазон" />
            <label>
              Строки — столбец №{' '}
              <input type="number" min={0} value={pivotRow} onChange={(e) => setPivotRow(Number(e.target.value))} style={{ width: 50 }} />
            </label>
            <label>
              Значения — столбец №{' '}
              <input type="number" min={0} value={pivotVal} onChange={(e) => setPivotVal(Number(e.target.value))} style={{ width: 50 }} />
            </label>
            <select value={pivotAgg} onChange={(e) => setPivotAgg(e.target.value as 'sum' | 'count' | 'avg')}>
              <option value="sum">Сумма</option>
              <option value="count">Количество</option>
              <option value="avg">Среднее</option>
            </select>
            <button className="tab" onClick={runPivot}>
              Создать на новом листе
            </button>
          </div>
          {sheet.filterCol != null && (
            <div style={{ marginTop: 8 }}>
              <b>
                Срез: столбец {indexToCol(sheet.filterCol)} ({filterVals.length} значений)
              </b>
              <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap', marginTop: 4 }}>
                {filterVals.map((v) => {
                  const on = sheet.filterValues.includes(v);
                  return (
                    <button
                      key={v}
                      className={on ? 'tab active' : 'tab'}
                      onClick={() =>
                        upd((s) => ({
                          ...s,
                          filterValues: on ? s.filterValues.filter((x) => x !== v) : [...s.filterValues, v],
                        }))
                      }
                    >
                      {v === '' ? '(пустые)' : v}
                    </button>
                  );
                })}
                <button className="tab" onClick={() => upd((s) => ({ ...s, filterValues: [] }))}>
                  Сбросить
                </button>
              </div>
            </div>
          )}
        </div>
      )}

      {ribbon === 'review' && Object.keys(sheet.comments).length > 0 && (
        <div className="card">
          <b>Примечания листа ({Object.keys(sheet.comments).length})</b>
          {Object.entries(sheet.comments).map(([ref, c]) => (
            <div key={ref} style={{ display: 'flex', gap: 8, alignItems: 'center', marginTop: 4, flexWrap: 'wrap' }}>
              <button
                className="tab"
                onClick={() => {
                  setSel(ref);
                  setRenderRows((n) => Math.max(n, (parseRef(ref)?.row ?? 0) + 5));
                }}
              >
                {ref}
              </button>
              <span className="hint">
                {c.author}: {c.text}
              </span>
              <button
                className="tab"
                onClick={() =>
                  upd((s) => {
                    const comments = { ...s.comments };
                    delete comments[ref];
                    return { ...s, comments };
                  })
                }
              >
                ✕
              </button>
            </div>
          ))}
        </div>
      )}

      {found.length > 0 && (
        <div className="card">
          <b>Найдено: {found.length}</b>
          <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap', marginTop: 4 }}>
            {found.map((r) => (
              <button key={r} className="tab" onClick={() => setSel(r)}>
                {r}
              </button>
            ))}
          </div>
        </div>
      )}

      <div className="formula-bar">
        <b style={{ padding: '8px 4px' }}>{sel}</b>
        <input
          value={draft ?? sheet.cells[sel] ?? ''}
          placeholder='=ЕСЛИ(A1>5;"много";"мало") — формулы, Лист2!A1 — другой лист'
          onChange={(e) => setDraft(e.target.value)}
          onBlur={() => {
            if (draft !== null) {
              commit(sel, draft);
              setDraft(null);
            }
          }}
          onKeyDown={(e) => {
            if (e.key === 'Enter' && draft !== null) {
              commit(sel, draft);
              setDraft(null);
            }
          }}
        />
        {sheet.comments[sel] && <span title={`${sheet.comments[sel].author}: ${sheet.comments[sel].text}`}>💬</span>}
      </div>
      {valError && <p className="hint">⚠ {valError}</p>}

      <div className="sheet-tabs">
        {sheets.map((s, i) => (
          <span key={i} style={{ display: 'inline-flex', gap: 2, alignItems: 'center' }}>
            <button
              className={i === active ? 'tab active' : 'tab'}
              onClick={() => {
                setActive(i);
                setSel('A1');
                setDraft(null);
                setRenderRows(RENDER_STEP);
              }}
              onDoubleClick={() => {
                const name = prompt('Имя листа:', s.name);
                if (name) setSheets((prev) => prev.map((x, j) => (j === i ? { ...x, name } : x)));
              }}
              title="Двойной клик — переименовать"
            >
              {s.name}
            </button>
            {sheets.length > 1 && (
              <button
                className="tab"
                title="Удалить лист"
                onClick={() => {
                  if (!confirm(`Удалить лист ${s.name}?`)) return;
                  setSheets((prev) => prev.filter((_, j) => j !== i));
                  setActive((a) => Math.max(0, Math.min(a, sheets.length - 2)));
                }}
              >
                ✕
              </button>
            )}
          </span>
        ))}
        <button
          className="tab"
          title="Добавить лист"
          onClick={() => {
            let n = sheets.length + 1;
            while (sheets.some((s) => s.name === `Лист${n}`)) n++;
            setSheets((prev) => [...prev, newSheet(`Лист${n}`)]);
            setActive(sheets.length);
          }}
        >
          ＋
        </button>
        {sheet.protected && <span className="hint">🔒 Лист защищён</span>}
      </div>

      <div className="grid-wrap">
        <table className="grid" ref={tableRef} style={{ tableLayout: 'fixed', fontSize: `${13 * zw}px` }}>
          {showHeaders && (
            <thead>
              <tr>
                <th className="row-h">#</th>
                {cols.map((c, j) => {
                  const frozen = frzCols > 0 && j < frzCols;
                  return (
                    <th
                      key={c}
                      style={
                        frozen
                          ? { position: 'sticky', left: measure.rowHdrW + j * measure.colW, top: 0, zIndex: 7, background: '#f1f3f5' }
                          : undefined
                      }
                    >
                      {c}
                    </th>
                  );
                })}
              </tr>
            </thead>
          )}
          <tbody>
            {rows.map((r) => (
              <tr key={r}>
                {showHeaders && <th className="row-h">{r}</th>}
                {cols.map((c, j) => {
                  const ref = `${c}${r}`;
                  const m = mergeAt(sheet.merges, ref);
                  if (m && ref !== refKey(m.c0, m.r0)) return null; // перекрыта объединением
                  const isHead = m != null;
                  const raw = sheet.cells[ref] ?? '';
                  const val = computed[ref] ?? raw;
                  const display = formatValue(typeof val === 'number' || typeof val === 'string' ? val : String(val), sheet.styles[ref]);
                  const st = sheet.styles[ref] ?? {};
                  const cf = cellCF(ref);
                  const spark = sheet.sparks.find((x) => x.at === ref);
                  const hasComment = !!sheet.comments[ref];
                  const frozen = frozenStyle(r - 1, j, false);
                  const validation = sheet.validations[ref];
                  const listVals = validation?.type === 'list' ? validation.values ?? [] : null;
                  return (
                    <td
                      key={ref}
                      className={sel === ref ? 'selected' : ''}
                      title={hasComment ? `${sheet.comments[ref].author}: ${sheet.comments[ref].text}` : raw}
                      rowSpan={isHead ? m!.r1 - m!.r0 + 1 : undefined}
                      colSpan={isHead ? m!.c1 - m!.c0 + 1 : undefined}
                      onClick={() => {
                        setSel(ref);
                        setDraft(null);
                        setValError('');
                      }}
                      style={{
                        ...frozen,
                        fontWeight: st.bold ? 700 : undefined,
                        fontStyle: st.italic ? 'italic' : undefined,
                        textDecoration: st.underline ? 'underline' : undefined,
                        fontFamily: st.font,
                        fontSize: st.size ? `${st.size * zw}px` : undefined,
                        color: cf?.color ?? st.color,
                        background: cf?.barPct !== undefined ? `linear-gradient(90deg,#93c5fd ${cf.barPct}%,transparent ${cf.barPct}%)` : (cf?.fill ?? st.fill),
                        textAlign: st.align,
                        border: gridLines ? undefined : '1px solid transparent',
                        minWidth: `${100 * zw}px`,
                        width: `${100 * zw}px`,
                        position: frozen.position ? 'sticky' : undefined,
                      }}
                    >
                      {hasComment && <span className="cmnt-pin" />}
                      {spark ? (
                        <div onClick={() => setSel(ref)}>
                          <Sparkline range={spark.range} type={spark.type} getNum={getNum} />
                        </div>
                      ) : (
                        <input
                          value={sel === ref ? (draft ?? raw) : cf?.icon ? `${cf.icon} ${display}` : display}
                          list={listVals ? `dl-${ref}` : undefined}
                          readOnly={sheet.protected}
                          onFocus={() => {
                            setSel(ref);
                            setDraft(null);
                          }}
                          onChange={(e) => {
                            setSel(ref);
                            setDraft(e.target.value);
                          }}
                          onBlur={() => {
                            if (draft !== null) {
                              commit(ref, draft);
                              setDraft(null);
                            }
                          }}
                          onKeyDown={(e) => {
                            if (e.key === 'Enter' && draft !== null) {
                              commit(ref, draft);
                              setDraft(null);
                            }
                          }}
                        />
                      )}
                      {listVals && (
                        <datalist id={`dl-${ref}`}>
                          {listVals.map((v) => (
                            <option key={v} value={v} />
                          ))}
                        </datalist>
                      )}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {renderRows < ROWS && (
        <button className="tab" style={{ marginTop: 8 }} onClick={() => setRenderRows((n) => Math.min(ROWS, n + RENDER_STEP))}>
          Показать ещё строки ({Math.min(renderRows, ROWS)}/{ROWS})
        </button>
      )}

      {sheet.charts.map((ch) => (
        <ChartView
          key={ch.id}
          chart={ch}
          getNum={getNum}
          getText={getText}
          onDelete={() => upd((s) => ({ ...s, charts: s.charts.filter((x) => x.id !== ch.id) }))}
        />
      ))}

      <p className="hint">
        {sheet.name} • {Object.keys(sheet.cells).length} ячеек • Формулы: <code>=ЕСЛИ(A1&gt;5;"много";"мало")</code>,{' '}
        <code>=ВПР(...)</code>, <code>=СУММЕСЛИ(...)</code> • Ссылки на листы: <code>Лист2!A1</code> • Сетка {ROWS}×{COLS} (A–
        {indexToCol(COLS - 1)}). Ошибки: <code>#DIV/0!</code>, <code>#ЗНАЧ!</code>→<code>#VALUE!</code>, <code>#ССЫЛКА!</code>→
        <code>#REF!</code>, <code>#Н/Д</code>→<code>#N/A</code>.
      </p>
    </div>
  );
}
