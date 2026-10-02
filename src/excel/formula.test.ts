import { describe, expect, it } from 'vitest';
import { evaluateFormula, shiftFormulaRefs } from './formula';

const table = (cells: Record<string, string>) => (ref: string) =>
  cells[ref.toUpperCase()];

describe('formula engine', () => {
  it('арифметика', () => {
    expect(evaluateFormula('=1+2*3', table({}))).toBe(7);
    expect(evaluateFormula('=(1+2)*3', table({}))).toBe(9);
  });

  it('ссылки на ячейки', () => {
    expect(evaluateFormula('=A1+B1', table({ A1: '2', B1: '3' }))).toBe(5);
    expect(evaluateFormula('=A1*B2+10', table({ A1: '2', B2: '5' }))).toBe(20);
  });

  it('SUM по диапазону', () => {
    const g = table({ A1: '1', A2: '2', A3: '3' });
    expect(evaluateFormula('=SUM(A1:A3)', g)).toBe(6);
  });

  it('AVERAGE / MIN / MAX / COUNT', () => {
    const g = table({ A1: '2', A2: '4' });
    expect(evaluateFormula('=AVERAGE(A1:A2)', g)).toBe(3);
    expect(evaluateFormula('=MIN(A1:A2)', g)).toBe(2);
    expect(evaluateFormula('=MAX(A1:A2)', g)).toBe(4);
    expect(evaluateFormula('=COUNT(A1:A2)', g)).toBe(2);
  });

  it('циклы и деление на ноль', () => {
    expect(evaluateFormula('=1/0', table({}))).toBe('#DIV/0!');
    const g = (ref: string) => (ref === 'A1' ? '=A1+1' : undefined);
    expect(evaluateFormula('=A1', g)).toBe('#CYCLE!');
  });

  it('сравнения и конкатенация', () => {
    expect(evaluateFormula('=2+3=5', table({}))).toBe('TRUE');
    expect(evaluateFormula('=1<>2', table({}))).toBe('TRUE');
    expect(evaluateFormula('=3>5', table({}))).toBe('FALSE');
    expect(evaluateFormula('="a"&"b"', table({}))).toBe('ab');
    expect(evaluateFormula('=2^3', table({}))).toBe(8);
    expect(evaluateFormula('=50%', table({}))).toBe(0.5);
    expect(evaluateFormula('=-2^2', table({}))).toBe(4);
  });

  it('IF / AND / OR / NOT / IFERROR', () => {
    expect(evaluateFormula('=IF(1>2,"да","нет")', table({}))).toBe('нет');
    expect(evaluateFormula('=IF(A1>1,"y","n")', table({ A1: '5' }))).toBe('y');
    expect(evaluateFormula('=AND(TRUE,1)', table({}))).toBe('TRUE');
    expect(evaluateFormula('=OR(FALSE,0)', table({}))).toBe('FALSE');
    expect(evaluateFormula('=NOT(1=1)', table({}))).toBe('FALSE');
    expect(evaluateFormula('=IFERROR(1/0,"ош")', table({}))).toBe('ош');
  });

  it('русские имена', () => {
    const g = table({ A1: '1', A2: '2' });
    expect(evaluateFormula('=СУММ(A1:A2)', g)).toBe(3);
    expect(evaluateFormula('=ЕСЛИ(A1=1,"да","нет")', g)).toBe('да');
  });

  it('разделитель ; и строки', () => {
    expect(evaluateFormula('=SUM(1;2;3)', table({}))).toBe(6);
    expect(evaluateFormula('=ДЛСТР("привет")', table({}))).toBe(6);
    expect(evaluateFormula('=ЛЕВСИМВ("привет";3)', table({}))).toBe('при');
    expect(evaluateFormula('=СЦЕПИТЬ("a";"b")', table({}))).toBe('ab');
  });

  it('COUNTIF / SUMIF / COUNTA', () => {
    const g = table({ A1: '1', A2: '5', A3: '5', B1: 'x' });
    expect(evaluateFormula('=COUNTIF(A1:A3;">3")', g)).toBe(2);
    expect(evaluateFormula('=SUMIF(A1:A3;">3")', g)).toBe(10);
    expect(evaluateFormula('=COUNTA(A1:B1)', g)).toBe(2);
    expect(evaluateFormula('=СЧЁТЕСЛИ(A1:A3;5)', g)).toBe(2);
  });

  it('VLOOKUP / INDEX / MATCH', () => {
    const g = table({ A1: 'a', B1: '1', A2: 'b', B2: '2' });
    expect(evaluateFormula('=VLOOKUP("b";A1:B2;2;FALSE)', g)).toBe(2);
    expect(evaluateFormula('=ВПР("z";A1:B2;2;ЛОЖЬ)', g)).toBe('#N/A');
    expect(evaluateFormula('=INDEX(A1:B2;2;1)', g)).toBe('b');
    expect(evaluateFormula('=MATCH("b";A1:A2;0)', g)).toBe(2);
  });

  it('даты', () => {
    expect(evaluateFormula('=DATE(2024;1;15)', table({}))).toBe(45306);
    expect(evaluateFormula('=YEAR(DATE(2024;5;1))', table({}))).toBe(2024);
    expect(evaluateFormula('=TEXT(0.456;"0.0%")', table({}))).toBe('45.6%');
  });

  it('HYPERLINK', () => {
    expect(evaluateFormula('=HYPERLINK("https://a.b";"Сайт")', table({}))).toBe('Сайт');
    expect(evaluateFormula('=ГИПЕРССЫЛКА("https://a.b")', table({}))).toBe('https://a.b');
  });

  it('ошибки', () => {
    expect(evaluateFormula('=НЕИЗВЕСТНАЯФ(1)', table({}))).toBe('#NAME?');
    expect(evaluateFormula('=SQRT(-1)', table({}))).toBe('#NUM!');
    expect(evaluateFormula('=A1+A2', table({ A1: 'x' }))).toBe('#VALUE!');
    expect(evaluateFormula('=A1:A3+1', table({ A1: '1' }))).toBe('#VALUE!');
  });

  it('абсолютные ссылки и листы', () => {
    // движок канонизирует ссылки ($ снимаются): мок хранит обычные ключи
    const g = (ref: string) => ({ A1: '7', 'ЛИСТ1!B2': '3' })[ref.toUpperCase()];
    expect(evaluateFormula('=$A$1*2', g)).toBe(14);
    expect(evaluateFormula('=Лист1!B2+1', g)).toBe(4);
    expect(evaluateFormula('=ROW(A5)', table({}))).toBe(5);
    expect(evaluateFormula('=COLUMN(C3)', table({}))).toBe(3);
  });
});

describe('shiftFormulaRefs', () => {
  it('сдвигает строки', () => {
    expect(
      shiftFormulaRefs('=SUM(A1:A5)+B2', { type: 'row', sheet: null, index: 1, delta: 1 }),
    ).toBe('=SUM(A1:A6)+B3');
    expect(
      shiftFormulaRefs('=SUM($A$1:A5)', { type: 'row', sheet: null, index: 0, delta: 1 }),
    ).toBe('=SUM($A$1:A6)');
  });

  it('сдвигает столбцы', () => {
    expect(
      shiftFormulaRefs('=A1+B1', { type: 'col', sheet: null, index: 0, delta: 1 }),
    ).toBe('=B1+C1');
  });
});
