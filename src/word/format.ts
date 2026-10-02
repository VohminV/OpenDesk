export function styleWithCSS() {
  try {
    document.execCommand('styleWithCSS', false, 'true');
  } catch {
    /* ignore */
  }
}

export function exec(cmd: string, value?: string) {
  styleWithCSS();
  document.execCommand(cmd, false, value);
}

export function applyFontSizePt(pt: number) {
  styleWithCSS();
  const sel = window.getSelection();
  if (!sel || sel.rangeCount === 0 || sel.isCollapsed) return;
  const range = sel.getRangeAt(0);
  const span = document.createElement('span');
  span.style.fontSize = `${pt}pt`;
  try {
    range.surroundContents(span);
  } catch {
    span.appendChild(range.extractContents());
    range.insertNode(span);
  }
  sel.removeAllRanges();
}

export function applyBlockStyle(css: Partial<CSSStyleDeclaration>) {
  const sel = window.getSelection();
  if (!sel || sel.rangeCount === 0) return;
  let node: Node | null = sel.anchorNode;
  while (node && node.nodeType !== 1) node = node.parentNode;
  let block = node as HTMLElement | null;
  while (block && !/^(P|H1|H2|H3|DIV|LI|BLOCKQUOTE|PRE)$/i.test(block.tagName)) {
    block = block.parentElement;
  }
  if (block) Object.assign(block.style, css);
}

export function insertHTML(html: string) {
  styleWithCSS();
  document.execCommand('insertHTML', false, html);
}

export interface HeadingItem {
  id: string;
  level: number;
  text: string;
}

export function collectHeadings(root: HTMLElement | null): HeadingItem[] {
  if (!root) return [];
  const out: HeadingItem[] = [];
  root.querySelectorAll('h1,h2,h3').forEach((h, i) => {
    const el = h as HTMLElement;
    if (!el.id) el.id = `hd-${i}`;
    out.push({
      id: el.id,
      level: el.tagName === 'H1' ? 1 : el.tagName === 'H2' ? 2 : 3,
      text: (el.textContent ?? '').slice(0, 80) || `Заголовок ${i + 1}`,
    });
  });
  return out;
}

export function buildTableHTML(rows: number, cols: number): string {
  const r = Math.min(Math.max(rows, 1), 20);
  const c = Math.min(Math.max(cols, 1), 10);
  let s = '<table border="1" style="border-collapse:collapse;width:100%;margin:8px 0"><tbody>';
  for (let i = 0; i < r; i++) {
    s += '<tr>';
    for (let j = 0; j < c; j++) s += '<td style="border:1px solid #ccc;padding:6px;min-width:40px">&nbsp;</td>';
    s += '</tr>';
  }
  return s + '</tbody></table><p></p>';
}

// ---------- Новые: фигуры / SmartArt / диаграммы / формулы ----------

export function shapeSVG(kind: 'rect' | 'circle' | 'arrow' | 'star'): string {
  const inner =
    kind === 'rect'
      ? '<rect x="10" y="20" width="180" height="80" rx="10" fill="#dbeafe" stroke="#2563eb" stroke-width="3"/>'
      : kind === 'circle'
        ? '<circle cx="100" cy="60" r="45" fill="#fef9c3" stroke="#ca8a04" stroke-width="3"/>'
        : kind === 'arrow'
          ? '<path d="M10 60 H140 M140 60 L110 35 M140 60 L110 85" fill="none" stroke="#16a34a" stroke-width="6" stroke-linecap="round"/>'
          : '<path d="M100 5 L120 70 L190 70 L133 110 L153 175 L100 135 L47 175 L67 110 L10 70 L80 70 Z" fill="#fce7f3" stroke="#db2777" stroke-width="3"/>';
  return `<svg xmlns="http://www.w3.org/2000/svg" width="200" height="120" viewBox="0 0 200 120" style="max-width:100%">${inner}</svg><p></p>`;
}

export function smartArtHTML(items: string[]): string {
  const cells = items.slice(0, 4).map((t) => `<div>${t}</div>`).join('');
  return `<div class="smartart s1" contenteditable="true">${cells}</div><p></p>`;
}

export function barChartSVG(values: number[], labels: string[]): string {
  const w = 320;
  const h = 160;
  const max = Math.max(...values, 1);
  const bw = Math.floor((w - 20) / Math.max(values.length, 1)) - 8;
  let bars = '';
  values.slice(0, 8).forEach((v, i) => {
    const bh = Math.round(((h - 40) * v) / max);
    const x = 10 + i * (bw + 8);
    const y = h - 20 - bh;
    bars += `<rect x="${x}" y="${y}" width="${bw}" height="${bh}" rx="4" fill="#2563eb"><title>${labels[i] ?? v}</title></rect>`;
    bars += `<text x="${x + bw / 2}" y="${h - 6}" font-size="10" text-anchor="middle" fill="#555">${labels[i] ?? ''}</text>`;
  });
  return `<div class="chart-box"><svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}" style="max-width:100%">${bars}</svg></div><p></p>`;
}

export function equationHTML(kind: 'frac' | 'sum' | 'sqrt' | 'integral'): string {
  if (kind === 'frac')
    return ' <span class="equation"><span style="display:inline-flex;flex-direction:column;text-align:center;vertical-align:middle"><span>a</span><span style="border-top:1px solid #000">b</span></span></span> ';
  if (kind === 'sum') return ' <span class="equation">∑<sub>i=1</sub><sup>n</sup> x<sub>i</sub></span> ';
  if (kind === 'sqrt') return ' <span class="equation">√(x² + y²)</span> ';
  return ' <span class="equation">∫<sub>a</sub><sup>b</sup> f(x)dx</span> ';
}

export function coverHTML(title: string, subtitle: string): string {
  return `<div class="cover"><h1>${title}</h1><p>${subtitle}</p><p style="color:#64748b">OpenDesk • 2010-стиль</p></div><p></p>`;
}

// ---------- Рассылки ----------

export function parseMergeData(csv: string): { headers: string[]; rows: string[][] } {
  const lines = csv.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
  if (!lines.length) return { headers: [], rows: [] };
  const split = (l: string) => l.split(/[;,]/).map((s) => s.trim());
  return { headers: split(lines[0]), rows: lines.slice(1).map(split) };
}

export function applyMerge(templateHtml: string, headers: string[], row: string[]): string {
  let out = templateHtml;
  headers.forEach((h, i) => {
    out = out.split(`{{${h}}}`).join(row[i] ?? '');
  });
  return out;
}

// ---------- Тезаурус (офлайн-мини) ----------

const THESAURUS: Record<string, string[]> = {
  хороший: ['отличный', 'прекрасный', 'замечательный', 'добротный'],
  плохой: ['скверный', 'неважный', 'слабый'],
  большой: ['огромный', 'крупный', 'значительный', 'обширный'],
  маленький: ['крошечный', 'небольшой', 'миниатюрный'],
  быстрый: ['скорый', 'стремительный', 'оперативный'],
  важный: ['значимый', 'существенный', 'ключевой'],
  документ: ['бумага', 'акт', 'материал'],
  good: ['great', 'excellent', 'fine'],
  big: ['large', 'huge', 'vast'],
};

export function thesaurus(word: string): string[] {
  return THESAURUS[word.toLowerCase().trim()] ?? [];
}

// ---------- Сравнение документов (простой diff по словам) ----------

export function diffWords(a: string, b: string): { del: number; ins: number; html: string } {
  const aw = a.split(/\s+/).filter(Boolean);
  const bw = b.split(/\s+/).filter(Boolean);
  const setB = new Set(bw);
  const setA = new Set(aw);
  let del = 0;
  let ins = 0;
  const htmlA = aw.map((w) => (setB.has(w) ? w : (++del, `<del class="track-del">${w}</del>`))).join(' ');
  const htmlB = bw.map((w) => (setA.has(w) ? w : (++ins, `<ins class="track-ins">${w}</ins>`))).join(' ');
  return { del, ins, html: `<p><b>Было:</b> ${htmlA}</p><p><b>Стало:</b> ${htmlB}</p>` };
}
