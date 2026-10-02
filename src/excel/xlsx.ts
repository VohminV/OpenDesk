import ExcelJS from 'exceljs';
import type { Sheet, NumFmt } from './format';

const NUM_FMT_XLSX: Record<NumFmt, string | undefined> = {
  general: undefined,
  number: '#,##0.00',
  currency: '#,##0.00 [$₽-419]',
  percent: '0.00%',
  date: 'DD.MM.YYYY',
  text: '@',
};

function argb(hex?: string): string | undefined {
  if (!hex) return undefined;
  const m = /^#([0-9a-f]{6})$/i.exec(hex);
  return m ? `FF${m[1].toUpperCase()}` : undefined;
}

export async function exportSheetsToXlsx(sheets: Sheet[], title: string): Promise<Blob> {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'OpenDesk';
  wb.title = title;
  for (const sh of sheets) {
    const ws = wb.addWorksheet(sh.name || 'Лист');
    // значения и формулы
    for (const [ref, raw] of Object.entries(sh.cells)) {
      const cell = ws.getCell(ref);
      if (raw.startsWith('=')) {
        cell.value = { formula: raw.slice(1), result: undefined };
      } else {
        const n = Number(raw);
        cell.value = raw !== '' && Number.isFinite(n) ? n : raw;
      }
      const st = sh.styles[ref];
      if (st) {
        cell.font = {
          bold: st.bold,
          italic: st.italic,
          underline: st.underline,
          name: st.font,
          size: st.size,
          color: st.color ? { argb: argb(st.color)! } : undefined,
        };
        if (st.fill) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: argb(st.fill)! } };
        if (st.align) cell.alignment = { horizontal: st.align, vertical: 'middle' };
        if (st.border) {
          cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' },
          };
        }
        const xf = st.numFmt ? NUM_FMT_XLSX[st.numFmt] : undefined;
        if (xf) cell.numFmt = xf;
      }
      const cm = sh.comments[ref];
      if (cm) {
        cell.note = { texts: [{ text: `${cm.author}: ${cm.text}` }] } as never;
      }
    }
    // объединения
    for (const m of sh.merges) {
      try {
        ws.mergeCells(m);
      } catch {
        /* пересечения пропускаем */
      }
    }
    // проверка данных (списки)
    for (const [ref, v] of Object.entries(sh.validations)) {
      try {
        if (v.type === 'list') {
          ws.getCell(ref).dataValidation = {
            type: 'list',
            formulae: [`"${(v.values ?? []).join(',')}"`],
            showErrorMessage: true,
            errorTitle: 'OpenDesk',
            error: 'Значение не из списка',
          };
        } else {
          ws.getCell(ref).dataValidation = {
            type: v.type,
            operator: 'between',
            formulae: [v.min ?? -1e12, v.max ?? 1e12],
            showErrorMessage: true,
            errorTitle: 'OpenDesk',
            error: 'Число вне диапазона',
          } as never;
        }
      } catch {
        /* ignore */
      }
    }
    // ширины по содержимому (эвристика)
    const widths = new Map<number, number>();
    for (const ref of Object.keys(sh.cells)) {
      const m = /^([A-Z]+)([0-9]+)$/.exec(ref);
      if (!m) continue;
      const col = m[1];
      const len = String(sh.cells[ref]).length;
      const idx = colToNumber(col);
      widths.set(idx, Math.min(60, Math.max(widths.get(idx) ?? 10, len + 2)));
    }
    widths.forEach((w, i) => {
      ws.getColumn(i).width = w;
    });
  }
  const buf = await wb.xlsx.writeBuffer();
  return new Blob([buf as ArrayBuffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
}

function colToNumber(col: string): number {
  let n = 0;
  for (const ch of col) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

export interface ImportedSheet {
  name: string;
  cells: Record<string, string>;
  dateCells: string[];
}

export async function importXlsx(file: File): Promise<ImportedSheet[]> {
  const buf = await file.arrayBuffer();
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf as never);
  const out: ImportedSheet[] = [];
  wb.eachSheet((ws) => {
    const cells: Record<string, string> = {};
    const dateCells: string[] = [];
    ws.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        const ref = (cell as unknown as { address: string }).address;
        const v = cell.value as unknown;
        if (v == null) return;
        if (typeof v === 'object' && v !== null && 'formula' in (v as Record<string, unknown>)) {
          const f = (v as { formula: string; result?: unknown }).formula;
          cells[ref] = `=${f}`;
        } else if (v instanceof Date) {
          const serial = Math.round((Date.UTC(v.getFullYear(), v.getMonth(), v.getDate()) - Date.UTC(1899, 11, 30)) / 86400000);
          cells[ref] = String(serial);
          dateCells.push(ref);
        } else if (typeof v === 'object' && v !== null && 'richText' in (v as Record<string, unknown>)) {
          cells[ref] = ((v as { richText: { text: string }[] }).richText ?? []).map((r) => r.text).join('');
        } else {
          cells[ref] = String(v as string | number | boolean);
        }
      });
    });
    out.push({ name: ws.name, cells, dateCells });
  });
  return out;
}
