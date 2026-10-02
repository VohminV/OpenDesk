import { describe, expect, it } from 'vitest';
import {
  parseMergeData,
  applyMerge,
  thesaurus,
  diffWords,
  shapeSVG,
  smartArtHTML,
  barChartSVG,
  equationHTML,
  coverHTML,
  buildTableHTML,
} from './format';

describe('mail merge', () => {
  it('парсит CSV с ; и ,', () => {
    expect(parseMergeData('Имя;Город\nИван;Москва')).toEqual({
      headers: ['Имя', 'Город'],
      rows: [['Иван', 'Москва']],
    });
    expect(parseMergeData('Name,City\nAnn,Kazan')).toEqual({
      headers: ['Name', 'City'],
      rows: [['Ann', 'Kazan']],
    });
  });

  it('подставляет поля {{}}', () => {
    expect(applyMerge('Привет, {{Имя}} из {{Город}}!', ['Имя', 'Город'], ['Иван', 'Москвы'])).toBe(
      'Привет, Иван из Москвы!',
    );
  });

  it('пустой ввод', () => {
    expect(parseMergeData('')).toEqual({ headers: [], rows: [] });
  });
});

describe('thesaurus', () => {
  it('находит синонимы', () => {
    expect(thesaurus('хороший')).toContain('отличный');
    expect(thesaurus('  БОЛЬШОЙ ')).toContain('огромный');
  });

  it('неизвестное слово — пусто', () => {
    expect(thesaurus('абракадабра')).toEqual([]);
  });
});

describe('diff', () => {
  it('считает вставки/удаления', () => {
    const r = diffWords('один два', 'один три два');
    expect(r.ins).toBe(1);
    expect(r.del).toBe(0);
    expect(r.html).toContain('track-ins');
  });
});

describe('вставки', () => {
  it('фигуры — валидный svg', () => {
    for (const k of ['rect', 'circle', 'arrow', 'star'] as const) {
      expect(shapeSVG(k)).toContain('<svg');
    }
  });

  it('smartart режет до 4 блоков', () => {
    const html = smartArtHTML(['a', 'b', 'c', 'd', 'e']);
    expect(html).toContain('smartart');
    expect((html.match(/<div>/g) ?? []).length).toBe(4);
  });

  it('диаграмма содержит бары', () => {
    expect(barChartSVG([5, 8], ['A', 'B'])).toContain('<rect');
  });

  it('формулы и обложка', () => {
    expect(equationHTML('frac')).toContain('equation');
    expect(coverHTML('T', 'S')).toContain('cover');
  });

  it('таблица NxM', () => {
    const html = buildTableHTML(2, 3);
    expect((html.match(/<tr>/g) ?? []).length).toBe(2);
    expect((html.match(/<td/g) ?? []).length).toBe(6);
  });
});
