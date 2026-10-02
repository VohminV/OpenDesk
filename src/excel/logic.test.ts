import { describe, expect, it } from 'vitest';
import { parseRange, sortSheet, buildPivot, insertDelete, validateCell, mergeAt } from './logic';
import { newSheet } from './format';

describe('parseRange', () => {
  it('ячейка и диапазон', () => {
    expect(parseRange('B2')).toMatchObject({ r0: 1, c0: 1, r1: 1, c1: 1 });
    expect(parseRange('A1:C3')).toMatchObject({ r0: 0, c0: 0, r1: 2, c1: 2 });
    expect(parseRange('C3:A1')).toMatchObject({ r0: 0, c0: 0, r1: 2, c1: 2 });
    expect(parseRange('Лист1!B2')).toMatchObject({ sheet: 'Лист1', c0: 1 });
    expect(parseRange('хлам')).toBeNull();
  });
});

describe('sortSheet', () => {
  it('сортирует числа и тянет стили', () => {
    const sh = {
      ...newSheet('S'),
      cells: { A1: '3', A2: '1', A3: '2' },
      styles: { A1: { bold: true } },
    };
    const box = parseRange('A1:A3')!;
    const out = sortSheet(sh, box, 0, 1);
    expect([out.cells['A1'], out.cells['A2'], out.cells['A3']]).toEqual(['1', '2', '3']);
    expect(out.styles['A3']).toEqual({ bold: true });
  });
});

describe('buildPivot', () => {
  it('группирует и суммирует', () => {
    const p = buildPivot(
      [
        ['a', '10'],
        ['b', '5'],
        ['a', '7'],
      ],
      0,
      1,
      'sum',
    );
    expect(p.headers).toEqual(['Строка', 'Сумма']);
    expect(p.rows[0]).toEqual(['a', 17]);
    expect(p.rows[1]).toEqual(['b', 5]);
    expect(p.rows[2][0]).toBe('ИТОГО');
  });
});

describe('insertDelete', () => {
  it('вставка строки сдвигает ячейки и формулы', () => {
    const sh = { ...newSheet('Лист1'), cells: { A1: '1', A2: '=A1+1' } };
    const [out] = insertDelete([sh], 0, 'row', 0, 1);
    expect(out.cells['A1']).toBeUndefined();
    expect(out.cells['A2']).toBe('1');
    expect(out.cells['A3']).toBe('=A2+1');
  });

  it('удаление столбца правит ссылки', () => {
    const sh = { ...newSheet('S'), cells: { A1: '1', C1: '=A1+B1' } };
    const [out] = insertDelete([sh], 0, 'col', 1, -1);
    expect(out.cells['B1']).toBe('=A1+#REF!');
  });
});

describe('validateCell', () => {
  it('список и диапазон', () => {
    expect(validateCell('a', { type: 'list', values: ['a', 'b'] })).toBeNull();
    expect(validateCell('z', { type: 'list', values: ['a', 'b'] })).toContain('Допустимы');
    expect(validateCell('5', { type: 'whole', min: 1, max: 10 })).toBeNull();
    expect(validateCell('5.5', { type: 'whole' })).toContain('целое');
    expect(validateCell('99', { type: 'decimal', max: 10 })).toContain('Максимум');
  });
});

describe('mergeAt', () => {
  it('находит объединение', () => {
    expect(mergeAt(['B2:C4'], 'C3')).not.toBeNull();
    expect(mergeAt(['B2:C4'], 'D5')).toBeNull();
  });
});
