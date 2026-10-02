import { describe, expect, it } from 'vitest';
import { exportSheetsToXlsx, importXlsx } from './xlsx';
import { newSheet } from './format';

function xlsxFile(blob: Blob): File {
  return new File([blob], 'test.xlsx', {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
}

describe('xlsx round-trip', () => {
  it('значения, формулы, стили, объединения', async () => {
    const sh = {
      ...newSheet('Лист1'),
      cells: { A1: '10', A2: '20', A3: '=SUM(A1:A2)', B1: 'текст' },
      styles: { A1: { bold: true, numFmt: 'number' as const }, B1: { fill: '#ffff00' } },
      merges: ['C1:D2'],
      comments: { B1: { author: 'Я', text: 'привет' } },
    };
    const blob = await exportSheetsToXlsx([sh, { ...newSheet('Лист2'), cells: { A1: 'x' } }], 't');
    expect(blob.size).toBeGreaterThan(3000);

    const imp = await importXlsx(xlsxFile(blob));
    expect(imp.length).toBe(2);
    const s1 = imp.find((s) => s.name === 'Лист1')!;
    expect(s1.cells['A1']).toBe('10');
    expect(s1.cells['A3']).toBe('=SUM(A1:A2)');
    expect(s1.cells['B1']).toBe('текст');
    expect(imp.find((s) => s.name === 'Лист2')?.cells['A1']).toBe('x');
  });

  it('даты помечаются', async () => {
    const wb = await import('./xlsx');
    void wb;
    const { default: ExcelJS } = await import('exceljs');
    const book = new ExcelJS.Workbook();
    const ws = book.addWorksheet('D');
    ws.getCell('A1').value = new Date(2024, 0, 15);
    const buf = await book.xlsx.writeBuffer();
    const imp = await importXlsx(new File([buf as ArrayBuffer], 'd.xlsx'));
    expect(imp[0].cells['A1']).toBe('45306');
    expect(imp[0].dateCells).toContain('A1');
  });
});
