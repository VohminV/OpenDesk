import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  AlignmentType,
  Table,
  TableRow,
  TableCell,
  WidthType,
  Header,
  Footer,
  PageOrientation,
  ImageRun,
  ExternalHyperlink,
  TableOfContents,
  PageBreak,
  FootnoteReferenceRun,
} from 'docx';
import mammoth from 'mammoth';

export interface DocProps {
  title: string;
  author: string;
  keywords: string;
}

export interface PageOpts {
  orientation: 'portrait' | 'landscape';
  pageSize: 'A4' | 'A5' | 'Letter';
  margins: 'narrow' | 'normal' | 'wide';
  pageNumbers: boolean;
}

const TWIPS = { A4: { w: 11906, h: 16838 }, A5: { w: 8391, h: 11906 }, Letter: { w: 12240, h: 15840 } };
const MARGIN_TWIPS = { narrow: 720, normal: 1440, wide: 2880 };

function dataUrlToBytes(dataUrl: string): { data: Uint8Array; type: 'png' | 'jpg' | 'gif' } | null {
  const m = /^data:image\/(png|jpe?g|gif);base64,(.+)$/.exec(dataUrl);
  if (!m) return null;
  const bin = atob(m[2]);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  const type = m[1] === 'png' ? 'png' : m[1] === 'gif' ? 'gif' : 'jpg';
  return { data: bytes, type };
}

function escapeXml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// ---------- Растеризация SVG/SmartArt в PNG для вставки в DOCX ----------

async function rasterizeSvgString(svg: string, width: number, height: number): Promise<Uint8Array | null> {
  try {
    const blob = new Blob([svg], { type: 'image/svg+xml;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    try {
      const img = new Image();
      await new Promise<void>((res, rej) => {
        img.onload = () => res();
        img.onerror = () => rej(new Error('svg-load'));
        img.src = url;
      });
      const scale = 2;
      const canvas = document.createElement('canvas');
      canvas.width = Math.max(1, Math.round(width * scale));
      canvas.height = Math.max(1, Math.round(height * scale));
      const ctx = canvas.getContext('2d');
      if (!ctx) return null;
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
      const conv = dataUrlToBytes(canvas.toDataURL('image/png'));
      return conv ? conv.data : null;
    } finally {
      URL.revokeObjectURL(url);
    }
  } catch {
    return null;
  }
}

function smartArtToSvg(items: string[]): { svg: string; w: number; h: number } {
  const colors = ['#2563eb', '#7c3aed', '#db2777', '#ea580c'];
  const list = items.length ? items.slice(0, 4) : [''];
  const W = 600;
  const H = 120;
  const bw = W / list.length;
  let parts = '';
  list.forEach((t, i) => {
    parts +=
      `<rect x="${(i * bw).toFixed(1)}" y="0" width="${bw.toFixed(1)}" height="${H}" fill="${colors[i % 4]}"/>` +
      `<text x="${(i * bw + bw / 2).toFixed(1)}" y="${H / 2}" font-size="20" font-family="Arial,sans-serif" fill="#ffffff" text-anchor="middle" dominant-baseline="middle">${escapeXml(t.slice(0, 40))}</text>`;
  });
  return { svg: `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}">${parts}</svg>`, w: W, h: H };
}

function svgSize(el: Element, fallbackW: number, fallbackH: number): { w: number; h: number } {
  const w = Number(el.getAttribute('width')) || fallbackW;
  const h = Number(el.getAttribute('height')) || fallbackH;
  return { w: Math.min(w || fallbackW, 800), h: Math.min(h || fallbackH, 600) };
}

async function imageRunFromSvgElement(svgEl: Element, maxW: number): Promise<ImageRun | null> {
  try {
    const clone = svgEl.cloneNode(true) as Element;
    clone.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
    const { w, h } = svgSize(clone, 400, 200);
    const data = await rasterizeSvgString(new XMLSerializer().serializeToString(clone), w, h);
    if (!data) return null;
    const dispW = Math.min(w, maxW);
    const dispH = Math.round((h * dispW) / Math.max(w, 1));
    return new ImageRun({ data, transformation: { width: dispW, height: dispH }, type: 'png' } as unknown as never);
  } catch {
    return null;
  }
}

// ---------- Inline-форматирование ----------

interface InlineFmt {
  bold?: boolean;
  italics?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: string;
  size?: number; // half-points
  font?: string;
}

function cssColorToDocx(css: string): string | undefined {
  const m = /#([0-9a-f]{6})/i.exec(css);
  if (m) return m[1].toUpperCase();
  const rgb = /rgb\(\s*(\d+),\s*(\d+),\s*(\d+)\s*\)/i.exec(css);
  if (rgb) {
    const h = (n: number) => Math.max(0, Math.min(255, Number(n))).toString(16).padStart(2, '0');
    return (h(Number(rgb[1])) + h(Number(rgb[2])) + h(Number(rgb[3]))).toUpperCase();
  }
  return undefined;
}

function styleOf(el: Element | null): InlineFmt {
  if (!el) return {};
  const cs = (el as HTMLElement).style;
  const out: InlineFmt = {};
  if (cs?.color) {
    const c = cssColorToDocx(cs.color);
    if (c && c !== '000000') out.color = c;
  }
  if (cs?.fontSize) {
    const m = /([\d.]+)pt/.exec(cs.fontSize);
    if (m) out.size = Math.round(Number(m[1]) * 2);
  }
  if (cs?.fontFamily) {
    const f = cs.fontFamily.split(',')[0].replace(/["']/g, '').trim();
    if (f) out.font = f;
  }
  return out;
}

type ParaChild = TextRun | ImageRun | ExternalHyperlink | FootnoteReferenceRun;

function toRun(text: string, fmt: InlineFmt): TextRun {
  return new TextRun({
    text,
    bold: fmt.bold,
    italics: fmt.italics,
    underline: fmt.underline ? {} : undefined,
    strike: fmt.strike,
    color: fmt.color,
    size: fmt.size,
    font: fmt.font,
  });
}

function runsFromInline(
  el: Element | ChildNode,
  base: InlineFmt = {},
  inheritedEl: Element | null = null,
): ParaChild[] {
  const out: ParaChild[] = [];
  el.childNodes.forEach((n) => {
    if (n.nodeType === 3) {
      const text = n.textContent ?? '';
      if (text) out.push(toRun(text, { ...styleOf(inheritedEl), ...base }));
    } else if (n.nodeType === 1) {
      const e = n as Element;
      const tag = e.tagName.toLowerCase();
      if (tag === 'img') {
        const src = e.getAttribute('src') ?? '';
        const conv = dataUrlToBytes(src);
        if (conv) {
          try {
            out.push(
              new ImageRun({
                data: conv.data,
                transformation: { width: 400, height: 300 },
                type: conv.type,
              } as unknown as never),
            );
          } catch {
            out.push(toRun('[рисунок]', base));
          }
        } else {
          out.push(toRun('[рисунок]', base));
        }
        return;
      }
      if (tag === 'sup' && e.classList.contains('footnote-ref')) {
        const m = /\[(\d+)\]/.exec(e.textContent ?? '');
        const id = m ? Number(m[1]) : 0;
        if (id > 0) {
          out.push(new FootnoteReferenceRun(id));
          return;
        }
      }
      if (tag === 'svg') return; // фигуры/диаграммы обрабатываются на уровне блока как PNG
      const st = styleOf(e);
      const next: InlineFmt = {
        ...base,
        ...st,
        bold: base.bold || st.bold || tag === 'b' || tag === 'strong',
        italics: base.italics || st.italics || tag === 'i' || tag === 'em',
        underline: base.underline || st.underline || tag === 'u',
        strike: base.strike || st.strike || tag === 's' || tag === 'strike' || tag === 'del',
        color: st.color ?? base.color,
        size: st.size ?? base.size,
        font: st.font ?? base.font,
      };
      if (tag === 'br') out.push(toRun('\n', base));
      else if (tag === 'a') {
        const href = e.getAttribute('href') ?? '';
        const label = e.textContent || href;
        if (/^https?:\/\//i.test(href)) {
          out.push(
            new ExternalHyperlink({
              children: [toRun(label, { ...next, color: next.color ?? '0563C1', underline: true })],
              link: href,
            }),
          );
        } else {
          out.push(toRun(label, next));
        }
      } else {
        out.push(...runsFromInline(e, next, e));
      }
    }
  });
  return out;
}

// ---------- Блоки ----------

async function blockToParagraphs(el: Element): Promise<(Paragraph | Table | TableOfContents)[]> {
  const tag = el.tagName.toLowerCase();
  if (el.classList.contains('footnotes') || el.classList.contains('footnote')) return []; // идут в w:footnotes
  if (el.classList.contains('toc')) {
    return [new TableOfContents('Оглавление', { hyperlink: true, headingStyleRange: '1-3' })];
  }
  if (el.classList.contains('smartart')) {
    const items = Array.from(el.children).map((c) => c.textContent?.trim() ?? '').filter(Boolean);
    const { svg, w, h } = smartArtToSvg(items);
    const data = await rasterizeSvgString(svg, w, h);
    if (data) {
      return [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new ImageRun({ data, transformation: { width: 500, height: Math.round((h * 500) / w) }, type: 'png' } as unknown as never),
          ],
        }),
      ];
    }
    return [new Paragraph({ children: [toRun(items.join(' → '), {})] })];
  }
  if (el.classList.contains('chart-box') || el.classList.contains('cover')) {
    const svg = el.querySelector('svg');
    if (svg) {
      const pic = await imageRunFromSvgElement(svg, 500);
      if (pic) return [new Paragraph({ alignment: AlignmentType.CENTER, children: [pic] })];
    }
    if (el.classList.contains('cover')) {
      const title = el.querySelector('h1')?.textContent ?? 'Титульный лист';
      return [
        new Paragraph({ heading: HeadingLevel.TITLE, children: [new TextRun({ text: title, size: 56, bold: true })] }),
        new Paragraph({ children: [new TextRun({ text: el.textContent ?? '' })] }),
      ];
    }
  }
  // абзац-обёртка над одиночной фигурой <p><svg/></p>
  if (tag === 'p' && el.querySelector(':scope > svg') && (el.textContent ?? '').trim() === '') {
    const pic = await imageRunFromSvgElement(el.querySelector(':scope > svg')!, 400);
    if (pic) return [new Paragraph({ alignment: AlignmentType.CENTER, children: [pic] })];
  }
  const align = (el as HTMLElement).style?.textAlign;
  const alignment =
    align === 'center'
      ? AlignmentType.CENTER
      : align === 'right'
        ? AlignmentType.RIGHT
        : align === 'justify'
          ? AlignmentType.JUSTIFIED
          : undefined;

  if (tag === 'table') {
    const rows: TableRow[] = [];
    el.querySelectorAll('tr').forEach((tr) => {
      const cells: TableCell[] = [];
      tr.querySelectorAll('th,td').forEach((td) => {
        cells.push(
          new TableCell({
            width: { size: 100 / Math.max(tr.children.length, 1), type: WidthType.PERCENTAGE },
            children: [new Paragraph({ children: runsFromInline(td) as never })],
          }),
        );
      });
      if (cells.length) rows.push(new TableRow({ children: cells }));
    });
    if (rows.length)
      return [new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows })];
    return [];
  }

  if (tag === 'ul' || tag === 'ol') {
    const out: Paragraph[] = [];
    el.querySelectorAll(':scope > li').forEach((li) => {
      out.push(
        new Paragraph({
          children: runsFromInline(li) as never,
          bullet: tag === 'ul' ? { level: 0 } : undefined,
          numbering: tag === 'ol' ? { reference: 'default-numbering', level: 0 } : undefined,
        }),
      );
    });
    return out;
  }

  if (tag === 'div' && el.getAttribute('style')?.includes('page-break-after')) {
    return [new Paragraph({ children: [new PageBreak()] })];
  }

  const heading =
    tag === 'h1'
      ? HeadingLevel.HEADING_1
      : tag === 'h2'
        ? HeadingLevel.HEADING_2
        : tag === 'h3'
          ? HeadingLevel.HEADING_3
          : undefined;

  return [
    new Paragraph({
      heading,
      alignment,
      children: runsFromInline(el) as never,
    }),
  ];
}

// ---------- Экспорт / импорт ----------

export async function exportHtmlToDocxBlob(
  bodyHtml: string,
  props: DocProps,
  headerText: string,
  footerText: string,
  page?: PageOpts,
): Promise<Blob> {
  const parser = new DOMParser();
  const doc = parser.parseFromString(`<div>${bodyHtml}</div>`, 'text/html');
  const root = doc.body.firstElementChild!;
  const children: (Paragraph | Table | TableOfContents)[] = [];

  // Собираем сноски: .footnote[data-n] -> w:footnotes
  const footnoteParas = new Map<number, Paragraph[]>();
  root.querySelectorAll('.footnote').forEach((fn) => {
    const id = Number(fn.getAttribute('data-n')) || 0;
    if (id <= 0) return;
    const text = (fn.textContent ?? '').replace(/^\[\d+\]\s*/, '');
    footnoteParas.set(id, [new Paragraph({ children: [new TextRun({ text, size: 18 })] })]);
  });
  const referencedIds = new Set<number>();
  root.querySelectorAll('.footnote-ref').forEach((s) => {
    const m = /\[(\d+)\]/.exec(s.textContent ?? '');
    if (m) referencedIds.add(Number(m[1]));
  });
  referencedIds.forEach((id) => {
    if (!footnoteParas.has(id)) footnoteParas.set(id, [new Paragraph({ children: [new TextRun('')] })]);
  });

  for (const n of Array.from(root.childNodes)) {
    if (n.nodeType === 3) {
      const t = n.textContent?.trim();
      if (t) children.push(new Paragraph({ children: [new TextRun(t)] }));
    } else if (n.nodeType === 1) {
      const e = n as Element;
      const tag = e.tagName.toLowerCase();
      if (['p', 'h1', 'h2', 'h3', 'ul', 'ol', 'table', 'blockquote', 'pre', 'div'].includes(tag)) {
        if (tag === 'div' && !e.className && !e.querySelector('table') && !e.querySelector('svg')) {
          children.push(new Paragraph({ children: runsFromInline(e) as never }));
        } else {
          children.push(...(await blockToParagraphs(e)));
        }
      } else {
        children.push(new Paragraph({ children: runsFromInline(e) as never }));
      }
    }
  }

  const pg = page ?? { orientation: 'portrait', pageSize: 'A4', margins: 'normal', pageNumbers: false };
  const dims = TWIPS[pg.pageSize];
  const landscape = pg.orientation === 'landscape';
  const m = MARGIN_TWIPS[pg.margins];

  const headers = headerText.trim()
    ? { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: headerText, italics: true, size: 18 })] })] }) }
    : undefined;
  const footers = footerText.trim() || pg.pageNumbers
    ? {
        default: new Footer({
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [new TextRun({ text: `${footerText}${pg.pageNumbers ? ' — стр. ' : ''}`, italics: true, size: 18 })],
            }),
          ],
        }),
      }
    : undefined;

  const wordDoc = new Document({
    creator: props.author || 'OpenDesk',
    title: props.title || 'OpenDesk document',
    keywords: props.keywords,
    footnotes: Object.fromEntries(
      Array.from(footnoteParas.entries()).map(([id, paras]) => [String(id), { children: paras }]),
    ),
    numbering: {
      config: [
        {
          reference: 'default-numbering',
          levels: [{ level: 0, format: 'decimal', text: '%1.', alignment: AlignmentType.LEFT }],
        },
      ],
    },
    sections: [
      {
        headers,
        footers,
        properties: {
          page: {
            size: {
              width: landscape ? dims.h : dims.w,
              height: landscape ? dims.w : dims.h,
              orientation: landscape ? PageOrientation.LANDSCAPE : PageOrientation.PORTRAIT,
            },
            margin: { top: m, bottom: m, left: m, right: m },
          },
        },
        children: children as never,
      },
    ],
  });

  return Packer.toBlob(wordDoc);
}

export async function importDocxToHtml(file: File): Promise<string> {
  const buf = await file.arrayBuffer();
  const res = await mammoth.convertToHtml({ arrayBuffer: buf });
  return res.value;
}
