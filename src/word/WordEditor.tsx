import { useEffect, useMemo, useRef, useState } from 'react';
import { countWords, downloadFile } from './document';
import { exportHtmlToDocxBlob, importDocxToHtml } from './docx';
import {
  exec,
  applyFontSizePt,
  applyBlockStyle,
  insertHTML,
  collectHeadings,
  buildTableHTML,
  shapeSVG,
  smartArtHTML,
  barChartSVG,
  equationHTML,
  coverHTML,
  parseMergeData,
  applyMerge,
  thesaurus,
  diffWords,
} from './format';

const BODY_KEY = 'opendesksuite.word.v1';
const META_KEY = 'opendesksuite.word.meta.v1';
const VER_KEY = 'opendesksuite.word.versions.v1';
const RECENT_KEY = 'opendesksuite.word.recent.v1';

interface DocMeta {
  title: string;
  author: string;
  keywords: string;
  header: string;
  footer: string;
  pageNumbers: boolean;
  orientation: 'portrait' | 'landscape';
  pageSize: 'A4' | 'A5' | 'Letter';
  margins: 'narrow' | 'normal' | 'wide';
  columns: 1 | 2;
  pageColor: string;
  watermark: string;
  hyphens: boolean;
  linenums: boolean;
}

const DEFAULT_META: DocMeta = {
  title: 'Документ OpenDesk',
  author: '',
  keywords: '',
  header: '',
  footer: '',
  pageNumbers: false,
  orientation: 'portrait',
  pageSize: 'A4',
  margins: 'normal',
  columns: 1,
  pageColor: '#ffffff',
  watermark: '',
  hyphens: false,
  linenums: false,
};

const FONTS = ['Calibri', 'Arial', 'Times New Roman', 'Courier New', 'Georgia', 'Verdana', 'Tahoma'];
const PT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48];
const SYMBOLS = ['©', '®', '™', '€', '£', '¥', '§', '¶', '—', '–', '«', '»', '…', '∑', '√', 'π', '±', '×', '÷', '°'];

interface Version { t: number; html: string; }

function loadMeta(): DocMeta {
  try {
    return { ...DEFAULT_META, ...JSON.parse(localStorage.getItem(META_KEY) ?? '{}') };
  } catch {
    return DEFAULT_META;
  }
}

function loadVersions(): Version[] {
  try {
    return JSON.parse(localStorage.getItem(VER_KEY) ?? '[]') as Version[];
  } catch {
    return [];
  }
}

type Ribbon = 'home' | 'insert' | 'layout' | 'refs' | 'mail' | 'review' | 'view' | 'file';

export default function WordEditor() {
  const ref = useRef<HTMLDivElement>(null);
  const fileInput = useRef<HTMLInputElement>(null);
  const imgInput = useRef<HTMLInputElement>(null);
  const [ribbon, setRibbon] = useState<Ribbon>('home');
  const [stats, setStats] = useState({ words: 0, chars: 0 });
  const [meta, setMeta] = useState<DocMeta>(() => loadMeta());
  const [zoom, setZoom] = useState(100);
  const [showRuler, setShowRuler] = useState(true);
  const [showMarks, setShowMarks] = useState(false);
  const [showNav, setShowNav] = useState(true);
  const [navOpen, setNavOpen] = useState(false);
  const [draftMode, setDraftMode] = useState(false);
  const [outlineOnly, setOutlineOnly] = useState(false);
  const [spell, setSpell] = useState(true);
  const [track, setTrack] = useState(false);
  const [readOnly, setReadOnly] = useState(false);
  const [findText, setFindText] = useState('');
  const [replaceText, setReplaceText] = useState('');
  const [foundCount, setFoundCount] = useState(0);
  const [headings, setHeadings] = useState<{ id: string; level: number; text: string }[]>([]);
  const [mergeCsv, setMergeCsv] = useState('Имя;Город\nИван;Москва\nАнна;Казань');
  const [mergeOut, setMergeOut] = useState<string[]>([]);
  const [thesWord, setThesWord] = useState('');
  const [thesOut, setThesOut] = useState<string[]>([]);
  const [diffOut, setDiffOut] = useState('');
  const [versions, setVersions] = useState<Version[]>(() => loadVersions());

  useEffect(() => {
    const saved = localStorage.getItem(BODY_KEY);
    if (saved && ref.current) {
      ref.current.innerHTML = saved;
      setStats(countWords(saved));
      setHeadings(collectHeadings(ref.current));
    }
  }, []);

  useEffect(() => {
    localStorage.setItem(META_KEY, JSON.stringify(meta));
  }, [meta]);

  const update = () => {
    const html = ref.current?.innerHTML ?? '';
    localStorage.setItem(BODY_KEY, html);
    setStats(countWords(html));
    setHeadings(collectHeadings(ref.current));
    try {
      const rec = JSON.parse(localStorage.getItem(RECENT_KEY) ?? '[]') as { name: string; t: number }[];
      rec.unshift({ name: meta.title || 'Без названия', t: Date.now() });
      localStorage.setItem(RECENT_KEY, JSON.stringify(rec.slice(0, 10)));
    } catch { /* ignore */ }
  };

  const cmd = (c: string, v?: string) => {
    if (readOnly) return;
    exec(c, v);
    ref.current?.focus();
    update();
  };

  const snapshot = () => {
    const html = ref.current?.innerHTML ?? '';
    const next = [{ t: Date.now(), html: html.slice(0, 200000) }, ...versions].slice(0, 5);
    setVersions(next);
    localStorage.setItem(VER_KEY, JSON.stringify(next));
  };

  const restoreVersion = (t: number) => {
    const v = versions.find((x) => x.t === t);
    if (v && ref.current) {
      if (!confirm('Восстановить эту версию? Текущий текст будет заменён.')) return;
      ref.current.innerHTML = v.html;
      update();
    }
  };

  // ---------- Файл ----------
  const newDoc = () => {
    if (!confirm('Создать новый документ? Текущий будет очищен.')) return;
    if (ref.current) ref.current.innerHTML = '';
    update();
  };

  const openFile = async (f: File) => {
    if (f.name.toLowerCase().endsWith('.docx')) {
      const html = await importDocxToHtml(f);
      if (ref.current) ref.current.innerHTML = html;
      update();
    } else {
      const text = await f.text();
      if (ref.current) {
        if (f.name.toLowerCase().endsWith('.html') || text.trimStart().startsWith('<')) {
          ref.current.innerHTML = text;
        } else {
          ref.current.innerText = text;
        }
      }
      update();
    }
  };

  const saveDocx = async () => {
    const blob = await exportHtmlToDocxBlob(
      ref.current?.innerHTML ?? '',
      { title: meta.title, author: meta.author, keywords: meta.keywords },
      meta.header,
      meta.footer,
      { orientation: meta.orientation, pageSize: meta.pageSize, margins: meta.margins, pageNumbers: meta.pageNumbers },
    );
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${meta.title || 'document'}.docx`;
    a.click();
    URL.revokeObjectURL(url);
  };

  // ---------- Правка/поиск ----------
  const doFind = () => {
    if (!findText || !ref.current) return;
    const text = ref.current.innerText;
    const count = text.split(findText).length - 1;
    setFoundCount(count);
    try {
      const sel = window.getSelection();
      sel?.removeAllRanges();
      // @ts-expect-error window.find есть в браузерах
      window.find(findText);
    } catch { /* ignore */ }
  };

  const doReplace = () => {
    if (!findText || !ref.current) return;
    const html = ref.current.innerHTML;
    const next = html.split(findText).join(replaceText);
    ref.current.innerHTML = next;
    update();
  };

  const changeCase = (mode: 'upper' | 'lower' | 'cap') => {
    const sel = window.getSelection();
    if (!sel || sel.isCollapsed) return;
    const text = sel.toString();
    const conv =
      mode === 'upper' ? text.toUpperCase() : mode === 'lower' ? text.toLowerCase() : text.replace(/\b\w/g, (c) => c.toUpperCase());
    exec('insertText', conv);
    update();
  };

  // ---------- Вставка ----------
  const insertTable = () => {
    const r = Number(prompt('Строк (1–20):', '3') ?? '3');
    const c = Number(prompt('Столбцов (1–10):', '3') ?? '3');
    insertHTML(buildTableHTML(r || 3, c || 3));
    update();
  };

  const insertImageFile = async (f: File) => {
    const data = await new Promise<string>((res) => {
      const r = new FileReader();
      r.onload = () => res(String(r.result));
      r.readAsDataURL(f);
    });
    insertHTML(`<p><img src="${data}" style="max-width:100%" alt="рисунок" /></p><p></p>`);
    update();
  };

  const onPaste: React.ClipboardEventHandler<HTMLDivElement> = async (e) => {
    const files = Array.from(e.clipboardData.files).filter((f) => f.type.startsWith('image/'));
    if (files.length) {
      e.preventDefault();
      for (const f of files) await insertImageFile(f);
    }
  };

  const insertLink = () => {
    const url = prompt('URL ссылки:', 'https://');
    if (url) cmd('createLink', url);
  };

  const insertSymbol = (s: string) => {
    insertHTML(s);
    update();
  };

  const insertFootnote = () => {
    const n = ref.current?.querySelectorAll('.footnote').length ?? 0;
    const id = n + 1;
    insertHTML(`<sup class="footnote-ref">[${id}]</sup>`);
    if (ref.current && !ref.current.querySelector('.footnotes')) {
      ref.current.insertAdjacentHTML('beforeend', '<hr class="footnotes-sep"/><div class="footnotes"></div>');
    }
    ref.current
      ?.querySelector('.footnotes')
      ?.insertAdjacentHTML('beforeend', `<p class="footnote" data-n="${id}">[${id}] <span>Текст сноски…</span></p>`);
    update();
  };

  const insertComment = () => {
    const author = meta.author || 'Рецензент';
    const text = prompt('Текст примечания:') ?? '';
    if (!text) return;
    insertHTML(` <span class="wcomment" title="${author}: ${text.replace(/"/g, '&quot;')}">💬 ${text}</span> `);
    update();
  };

  const insertTOC = () => {
    const items = collectHeadings(ref.current);
    if (!items.length) {
      alert('Нет заголовков H1–H3 для оглавления.');
      return;
    }
    const html =
      '<div class="toc"><p><b>Оглавление</b></p><ul>' +
      items.map((h) => `<li style="margin-left:${(h.level - 1) * 16}px"><a href="#${h.id}">${h.text}</a></li>`).join('') +
      '</ul></div><p></p>';
    insertHTML(html);
    update();
  };

  const insertShape = (kind: 'rect' | 'circle' | 'arrow' | 'star') => {
    insertHTML(`<p>${shapeSVG(kind)}</p>`);
    update();
  };

  const insertSmartArt = () => {
    const raw = prompt('Элементы через запятую:', 'Идея, План, Результат') ?? '';
    const items = raw.split(',').map((s) => s.trim()).filter(Boolean);
    if (!items.length) return;
    insertHTML(smartArtHTML(items));
    update();
  };

  const insertChart = () => {
    const raw = prompt('Значения через запятую:', '5,8,3,9') ?? '';
    const values = raw.split(',').map((s) => Number(s.trim())).filter((n) => Number.isFinite(n));
    if (!values.length) return;
    insertHTML(barChartSVG(values, values.map((_, i) => `П${i + 1}`)));
    update();
  };

  const insertEquation = (kind: 'frac' | 'sum' | 'sqrt' | 'integral') => {
    insertHTML(equationHTML(kind));
    update();
  };

  const insertCover = () => {
    insertHTML(coverHTML(meta.title || 'Название документа', meta.author || 'Автор'));
    update();
  };

  const insertBookmark = () => {
    const name = prompt('Имя закладки (латиница):', 'chapter1') ?? '';
    if (!/^[A-Za-z][\w-]*$/.test(name)) {
      alert('Имя должно начинаться с латинской буквы.');
      return;
    }
    insertHTML(`<a id="${name}">🔖 ${name}</a> `);
    update();
  };

  const screenshot = async () => {
    try {
      const md = navigator.mediaDevices as MediaDevices & { getDisplayMedia?: (c?: unknown) => Promise<MediaStream> };
      if (!md?.getDisplayMedia) {
        alert('Браузер не поддерживает захват экрана. Вставьте скриншот через Ctrl+V.');
        return;
      }
      const stream = await md.getDisplayMedia({ video: true });
      const video = document.createElement('video');
      video.srcObject = stream;
      await video.play();
      const canvas = document.createElement('canvas');
      canvas.width = video.videoWidth;
      canvas.height = video.videoHeight;
      canvas.getContext('2d')?.drawImage(video, 0, 0);
      stream.getTracks().forEach((t) => t.stop());
      insertHTML(`<p><img src="${canvas.toDataURL('image/png')}" style="max-width:100%" alt="скриншот" /></p><p></p>`);
      update();
    } catch {
      /* отменено пользователем */
    }
  };

  // ---------- Рецензирование ----------
  const authorNow = () => `${meta.author || 'Автор'} ${new Date().toLocaleString('ru')}`;

  const markIns = () => {
    const sel = window.getSelection();
    if (!sel || sel.isCollapsed) return;
    insertHTML(`<ins class="track-ins" data-author="${authorNow()}">${sel.toString()}</ins>`);
    update();
  };
  const markDel = () => {
    const sel = window.getSelection();
    if (!sel || sel.isCollapsed) return;
    insertHTML(`<del class="track-del" data-author="${authorNow()}">${sel.toString()}</del>`);
    update();
  };
  const acceptAll = () => {
    if (!ref.current) return;
    ref.current.querySelectorAll('del.track-del').forEach((d) => d.remove());
    ref.current.querySelectorAll('ins.track-ins').forEach((el) => {
      el.replaceWith(document.createTextNode(el.textContent ?? ''));
    });
    update();
  };

  // ---------- Рассылки ----------
  const mergeInfo = useMemo(() => parseMergeData(mergeCsv), [mergeCsv]);
  const runMerge = () => {
    const tpl = ref.current?.innerHTML ?? '';
    const out = mergeInfo.rows.map((row) => applyMerge(tpl, mergeInfo.headers, row));
    setMergeOut(out);
  };
  const insertMergeField = (h: string) => {
    insertHTML(`<span class="merge-field">{{${h}}}</span>`);
    update();
  };

  // ---------- Стиль страницы ----------
  const pageWidth = meta.pageSize === 'A4' ? 794 : meta.pageSize === 'A5' ? 559 : 816;
  const pad = meta.margins === 'narrow' ? 24 : meta.margins === 'wide' ? 72 : 48;
  const pages = Math.max(1, Math.ceil(stats.words / 500));

  const ribbonBtn = (id: Ribbon, label: string) => (
    <button key={id} className={ribbon === id ? 'tab active' : 'tab'} onClick={() => setRibbon(id)}>
      {label}
    </button>
  );

  return (
    <div>
      <div className="tabs" style={{ marginBottom: 8 }}>
        {ribbonBtn('home', 'Главная')}
        {ribbonBtn('insert', 'Вставка')}
        {ribbonBtn('layout', 'Разметка')}
        {ribbonBtn('refs', 'Ссылки')}
        {ribbonBtn('mail', 'Рассылки')}
        {ribbonBtn('review', 'Рецензирование')}
        {ribbonBtn('view', 'Вид')}
        {ribbonBtn('file', 'Файл')}
        <button className="tab" onClick={() => { setShowNav((v) => !v); setNavOpen((v) => !v); }} title="Навигация">☰</button>
      </div>

      {ribbon === 'home' && (
        <div className="toolbar">
          <button title="Вырезать" onClick={() => cmd('cut')}>✂</button>
          <button title="Копировать" onClick={() => cmd('copy')}>📋</button>
          <button title="Вставить" onClick={() => { navigator.clipboard?.readText().then((t) => { insertHTML(t); update(); }).catch(() => alert('Вставьте через Ctrl+V')); }}>📌</button>
          <select title="Шрифт" defaultValue="Calibri" onChange={(e) => cmd('fontName', e.target.value)}>
            {FONTS.map((f) => <option key={f} value={f}>{f}</option>)}
          </select>
          <select title="Размер (pt)" defaultValue="12" onChange={(e) => applyFontSizePt(Number(e.target.value))}>
            {PT_SIZES.map((s) => <option key={s} value={s}>{s}</option>)}
          </select>
          <button onClick={() => cmd('bold')}><b>B</b></button>
          <button onClick={() => cmd('italic')}><i>I</i></button>
          <button onClick={() => cmd('underline')}><u>U</u></button>
          <button onClick={() => cmd('strikeThrough')}>S̶</button>
          <button onClick={() => cmd('superscript')}>x²</button>
          <button onClick={() => cmd('subscript')}>x₂</button>
          <label title="Цвет текста">🎨<input type="color" hidden onChange={(e) => cmd('foreColor', e.target.value)} /></label>
          <label title="Маркер">🖍<input type="color" hidden defaultValue="#ffff00" onChange={(e) => cmd('hiliteColor', e.target.value)} /></label>
          <button title="Очистить формат" onClick={() => cmd('removeFormat')}>✕</button>
          <select title="Регистр" defaultValue="" onChange={(e) => { if (e.target.value) changeCase(e.target.value as 'upper' | 'lower' | 'cap'); e.target.value = ''; }}>
            <option value="">Аа…</option>
            <option value="upper">ВЕРХНИЙ</option>
            <option value="lower">нижний</option>
            <option value="cap">Как В Предложениях</option>
          </select>
          <select title="Стиль" defaultValue="" onChange={(e) => e.target.value && cmd('formatBlock', e.target.value)}>
            <option value="">Стиль…</option>
            <option value="p">Обычный</option>
            <option value="h1">Заголовок 1</option>
            <option value="h2">Заголовок 2</option>
            <option value="h3">Заголовок 3</option>
            <option value="blockquote">Цитата</option>
            <option value="pre">Код</option>
          </select>
          <button onClick={() => cmd('justifyLeft')}>⬅</button>
          <button onClick={() => cmd('justifyCenter')}>⬌</button>
          <button onClick={() => cmd('justifyRight')}>➡</button>
          <button title="По ширине" onClick={() => cmd('justifyFull')}>☰</button>
          <button title="Уменьшить отступ" onClick={() => cmd('outdent')}>⇤</button>
          <button title="Увеличить отступ" onClick={() => cmd('indent')}>⇥</button>
          <select title="Интервал" defaultValue="" onChange={(e) => { if (e.target.value) applyBlockStyle({ lineHeight: e.target.value }); e.target.value = ''; }}>
            <option value="">↕…</option>
            <option value="1">1.0</option>
            <option value="1.15">1.15</option>
            <option value="1.5">1.5</option>
            <option value="2">2.0</option>
          </select>
          <button onClick={() => cmd('insertUnorderedList')}>• Список</button>
          <button onClick={() => cmd('insertOrderedList')}>1. Нумерация</button>
          <input placeholder="Найти…" value={findText} onChange={(e) => setFindText(e.target.value)} style={{ width: 90 }} />
          <button onClick={doFind}>🔍{foundCount ? ` (${foundCount})` : ''}</button>
          <input placeholder="Заменить на…" value={replaceText} onChange={(e) => setReplaceText(e.target.value)} style={{ width: 90 }} />
          <button onClick={doReplace}>⇄</button>
        </div>
      )}

      {ribbon === 'insert' && (
        <div className="toolbar">
          <button onClick={insertCover}>📕 Титул</button>
          <button onClick={insertTable}>▦ Таблица</button>
          <button onClick={() => imgInput.current?.click()}>🖼 Рисунок</button>
          <button onClick={screenshot}>🖵 Скриншот</button>
          <button onClick={insertLink}>🔗 Ссылка</button>
          <button onClick={insertBookmark}>🔖 Закладка</button>
          <button onClick={() => insertShape('rect')}>▢</button>
          <button onClick={() => insertShape('circle')}>○</button>
          <button onClick={() => insertShape('arrow')}>→</button>
          <button onClick={() => insertShape('star')}>★</button>
          <button onClick={insertSmartArt}>🧩 SmartArt</button>
          <button onClick={insertChart}>📊 Диаграмма</button>
          <button onClick={() => insertEquation('frac')}>a/b</button>
          <button onClick={() => insertEquation('sum')}>∑</button>
          <button onClick={() => insertEquation('sqrt')}>√</button>
          <button onClick={() => insertEquation('integral')}>∫</button>
          <button onClick={() => { insertHTML('<hr/><p></p>'); update(); }}>― Линия</button>
          <button onClick={() => { insertHTML('<div style="border:1px solid #999;padding:12px;margin:8px 0">Надпись…</div><p></p>'); update(); }}>📦 Надпись</button>
          <button onClick={() => { insertHTML('<span style="font-size:28pt;font-weight:800;background:linear-gradient(90deg,#2563eb,#9333ea);-webkit-background-clip:text;color:transparent">WordArt</span> '); update(); }}>🎨 WordArt</button>
          <button onClick={() => { insertHTML('<span style="float:left;font-size:48pt;line-height:1;padding-right:8px;font-weight:800">Б</span>'); update(); }}>Б Буквица</button>
          <button onClick={insertFootnote}>† Сноска</button>
          <button onClick={insertComment}>💬 Комментарий</button>
          <button onClick={() => { insertHTML('<p style="text-align:center">Стр. <span class="pageno">1</span></p>'); update(); }}># Номер стр.</button>
          <select defaultValue="" title="Символ" onChange={(e) => { if (e.target.value) insertSymbol(e.target.value); e.target.value = ''; }}>
            <option value="">Ω Символ…</option>
            {SYMBOLS.map((s) => <option key={s} value={s}>{s}</option>)}
          </select>
          <input ref={imgInput} type="file" accept="image/*" hidden onChange={(e) => { const f = e.target.files?.[0]; if (f) insertImageFile(f); e.target.value = ''; }} />
        </div>
      )}

      {ribbon === 'layout' && (
        <div className="toolbar">
          <select value={meta.margins} onChange={(e) => setMeta({ ...meta, margins: e.target.value as DocMeta['margins'] })}>
            <option value="narrow">Поля: узкие</option>
            <option value="normal">Поля: обычные</option>
            <option value="wide">Поля: широкие</option>
          </select>
          <select value={meta.orientation} onChange={(e) => setMeta({ ...meta, orientation: e.target.value as 'portrait' | 'landscape' })}>
            <option value="portrait">Книжная</option>
            <option value="landscape">Альбомная</option>
          </select>
          <select value={meta.pageSize} onChange={(e) => setMeta({ ...meta, pageSize: e.target.value as DocMeta['pageSize'] })}>
            <option value="A4">A4</option>
            <option value="A5">A5</option>
            <option value="Letter">Letter</option>
          </select>
          <select value={meta.columns} onChange={(e) => setMeta({ ...meta, columns: Number(e.target.value) as 1 | 2 })}>
            <option value={1}>1 колонка</option>
            <option value={2}>2 колонки</option>
          </select>
          <label>Цвет стр. <input type="color" value={meta.pageColor} onChange={(e) => setMeta({ ...meta, pageColor: e.target.value })} /></label>
          <input placeholder="Подложка…" value={meta.watermark} onChange={(e) => setMeta({ ...meta, watermark: e.target.value })} style={{ width: 110 }} />
          <button onClick={() => { insertHTML('<div style="page-break-after:always"></div><p></p>'); update(); }}>⤵ Разрыв стр.</button>
          <button onClick={() => { insertHTML('<div style="page-break-after:always;border:1px dashed #999;padding:4px">— новый раздел —</div><p></p>'); update(); }}>📑 Раздел</button>
          <label><input type="checkbox" checked={meta.hyphens} onChange={(e) => setMeta({ ...meta, hyphens: e.target.checked })} /> Переносы</label>
          <label><input type="checkbox" checked={meta.linenums} onChange={(e) => setMeta({ ...meta, linenums: e.target.checked })} /> Номера строк</label>
          <input placeholder="Верхний колонтитул…" value={meta.header} onChange={(e) => setMeta({ ...meta, header: e.target.value })} style={{ width: 130 }} />
          <input placeholder="Нижний колонтитул…" value={meta.footer} onChange={(e) => setMeta({ ...meta, footer: e.target.value })} style={{ width: 130 }} />
          <label><input type="checkbox" checked={meta.pageNumbers} onChange={(e) => setMeta({ ...meta, pageNumbers: e.target.checked })} /> Номера стр.</label>
        </div>
      )}

      {ribbon === 'refs' && (
        <div className="toolbar">
          <button onClick={insertTOC}>📑 Оглавление</button>
          <button onClick={insertFootnote}>† Вставить сноску</button>
          <button onClick={() => { insertHTML('<p style="text-align:center;font-size:9pt;color:#555">Подпись: Рис. 1 — …</p>'); update(); }}>🖼 Подпись</button>
          <button onClick={() => { const s = window.getSelection(); if (s && !s.isCollapsed) { insertHTML(`<mark>${s.toString()}</mark>`); update(); } }}>✎ В указатель</button>
          <button onClick={() => {
            const marks = ref.current?.querySelectorAll('mark');
            const words = Array.from(new Set(Array.from(marks ?? []).map((m) => (m.textContent ?? '').trim()).filter(Boolean))).sort();
            insertHTML(`<div class="index"><p><b>Указатель</b></p><p>${words.join(', ') || '— нет помеченных слов —'}</p></div>`);
            update();
          }}>📚 Указатель</button>
        </div>
      )}

      {ribbon === 'mail' && (
        <div className="toolbar">
          <span className="hint">Поля:</span>
          {mergeInfo.headers.map((h) => (
            <button key={h} onClick={() => insertMergeField(h)}>{`{{${h}}}`}</button>
          ))}
          <button onClick={runMerge}>✉ Слияние ({mergeInfo.rows.length})</button>
          <button onClick={() => { insertHTML('<div style="border:1px solid #999;padding:24px;margin:8px 0;text-align:center"><p>Кому: {{Имя}}</p><p>{{Город}}</p><p>Индекс, адрес…</p></div><p></p>'); update(); }}>✉ Конверт</button>
          <button onClick={() => { insertHTML('<table border="1" style="border-collapse:collapse;width:100%"><tbody><tr><td style="padding:12px">{{Имя}}<br/>{{Город}}</td><td style="padding:12px">{{Имя}}<br/>{{Город}}</td></tr></tbody></table><p></p>'); update(); }}>🏷 Наклейки</button>
        </div>
      )}

      {ribbon === 'review' && (
        <div className="toolbar">
          <label><input type="checkbox" checked={spell} onChange={(e) => setSpell(e.target.checked)} /> Орфография</label>
          <button onClick={() => alert(`Слов: ${stats.words}, символов: ${stats.chars}, страниц (~500 слов): ${pages}`)}>🔢 Статистика</button>
          <button onClick={() => { const w = window.getSelection()?.toString() || stats.words.toString(); window.open(`https://translate.google.com/?sl=ru&tl=en&text=${encodeURIComponent(w.slice(0, 500))}`, '_blank'); }}>🌍 Перевести</button>
          <button onClick={insertComment}>💬 Примечание</button>
          <label><input type="checkbox" checked={track} onChange={(e) => setTrack(e.target.checked)} /> Исправления</label>
          {track && (<><button onClick={markIns}>➕ Вставка</button><button onClick={markDel}>➖ Удаление</button></>)}
          <button onClick={acceptAll}>✔ Принять все</button>
          <button onClick={() => {
            const cur = ref.current?.innerText ?? '';
            const old = prompt('Вставьте старый текст для сравнения:') ?? '';
            if (old) setDiffOut(diffWords(old, cur).html);
          }}>🔀 Сравнить</button>
          <label><input type="checkbox" checked={readOnly} onChange={(e) => setReadOnly(e.target.checked)} /> Защита</label>
        </div>
      )}

      {ribbon === 'view' && (
        <div className="toolbar">
          <button onClick={() => { setDraftMode(false); setOutlineOnly(false); }}>📄 Разметка</button>
          <button onClick={() => { setDraftMode(true); setOutlineOnly(false); }}>📝 Черновик</button>
          <button onClick={() => setOutlineOnly((v) => !v)}>{outlineOnly ? '👁 Всё' : '🧬 Структура'}</button>
          <label><input type="checkbox" checked={showRuler} onChange={(e) => setShowRuler(e.target.checked)} /> Линейка</label>
          <label><input type="checkbox" checked={showNav} onChange={(e) => setShowNav(e.target.checked)} /> Навигация</label>
          <label><input type="checkbox" checked={showMarks} onChange={(e) => setShowMarks(e.target.checked)} /> ¶ Знаки</label>
          <label>Масштаб <input type="range" min={50} max={200} value={zoom} onChange={(e) => setZoom(Number(e.target.value))} /> {zoom}%</label>
          <button onClick={() => document.documentElement.requestFullscreen?.()}>⛶ Во весь экран</button>
        </div>
      )}

      {ribbon === 'file' && (
        <div className="toolbar">
          <button onClick={newDoc}>📄 Создать</button>
          <button onClick={() => fileInput.current?.click()}>📂 Открыть</button>
          <button onClick={() => { snapshot(); downloadFile(`${meta.title || 'document'}.html`, ref.current?.innerHTML ?? '', 'text/html'); }}>💾 Сохранить HTML</button>
          <button onClick={saveDocx}>⬇ DOCX</button>
          <button onClick={() => downloadFile(`${meta.title || 'document'}.txt`, ref.current?.innerText ?? '', 'text/plain')}>⬇ TXT</button>
          <button onClick={() => window.print()}>🖨 PDF/Печать</button>
          <button onClick={snapshot}>🕘 Версия</button>
          <input ref={fileInput} type="file" accept=".html,.txt,.docx,text/html,text/plain" hidden onChange={(e) => { const f = e.target.files?.[0]; if (f) openFile(f); e.target.value = ''; }} />
        </div>
      )}

      {ribbon === 'file' && (
        <div className="card">
          <b>Свойства документа</b>
          <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', marginTop: 8 }}>
            <input placeholder="Название" value={meta.title} onChange={(e) => setMeta({ ...meta, title: e.target.value })} />
            <input placeholder="Автор" value={meta.author} onChange={(e) => setMeta({ ...meta, author: e.target.value })} />
            <input placeholder="Ключевые слова" value={meta.keywords} onChange={(e) => setMeta({ ...meta, keywords: e.target.value })} />
          </div>
          <p className="hint">Слов: {stats.words} • Символов: {stats.chars} • Страниц (~500 слов): {pages} • Автосохранение ✓</p>
          {versions.length > 0 && (
            <div>
              <b>Версии ({versions.length}/5)</b>
              {versions.map((v) => (
                <div key={v.t} style={{ display: 'flex', gap: 8, alignItems: 'center', marginTop: 4 }}>
                  <span className="hint">{new Date(v.t).toLocaleString('ru')}</span>
                  <button className="tab" onClick={() => restoreVersion(v.t)}>Восстановить</button>
                </div>
              ))}
            </div>
          )}
        </div>
      )}

      {ribbon === 'mail' && (
        <div className="card">
          <b>Данные для слияния (CSV, разделитель ; или ,)</b>
          <textarea value={mergeCsv} onChange={(e) => setMergeCsv(e.target.value)} rows={4} style={{ width: '100%', marginTop: 8, fontSize: 14 }} />
          {mergeOut.length > 0 && (
            <div style={{ marginTop: 8 }}>
              <b>Готово писем: {mergeOut.length}</b>
              <div style={{ display: 'flex', gap: 8, marginTop: 4, flexWrap: 'wrap' }}>
                <button className="tab" onClick={() => {
                  if (ref.current) {
                    ref.current.innerHTML = mergeOut.map((m) => `${m}<div style="page-break-after:always"></div>`).join('');
                    update();
                  }
                }}>Вставить все в документ</button>
                <button className="tab" onClick={() => downloadFile('merge.html', mergeOut.join('<hr/>'), 'text/html')}>Скачать HTML</button>
              </div>
              <div className="hint" dangerouslySetInnerHTML={{ __html: mergeOut[0]?.slice(0, 500) ?? '' }} />
            </div>
          )}
        </div>
      )}

      {(ribbon === 'review') && (
        <div className="card">
          <b>Тезаурус (офлайн)</b>
          <div style={{ display: 'flex', gap: 8, marginTop: 8, flexWrap: 'wrap' }}>
            <input placeholder="Слово…" value={thesWord} onChange={(e) => setThesWord(e.target.value)} />
            <button className="tab" onClick={() => setThesOut(thesaurus(thesWord))}>Синонимы</button>
            {thesOut.length > 0 && <span className="hint">{thesOut.join(', ')}</span>}
            {thesWord && thesOut.length === 0 && <span className="hint">— нет в словаре, попробуйте: хороший, большой, быстрый, важный</span>}
          </div>
          {diffOut && <div style={{ marginTop: 8 }} dangerouslySetInnerHTML={{ __html: diffOut }} />}
        </div>
      )}

      <div className="layout-2col">
        {showNav && (
          <aside className={`card nav-aside${navOpen ? '' : ' collapsed-mobile'}`}>
            <b>Навигация</b>
            {headings.length === 0 && <p className="hint">Нет заголовков</p>}
            {headings.map((h) => (
              <div key={h.id} style={{ marginLeft: (h.level - 1) * 12, fontSize: 13 }}>
                <a href={`#${h.id}`} onClick={(e) => { e.preventDefault(); document.getElementById(h.id)?.scrollIntoView(); }}>{h.text}</a>
              </div>
            ))}
          </aside>
        )}

        <div style={{ flex: 1, minWidth: 0 }}>
          {showRuler && !draftMode && (
            <div className="ruler">
              {Array.from({ length: 20 }, (_, i) => <span key={i}>{i + 1}</span>)}
            </div>
          )}
          {meta.header && <div className="doc-header">{meta.header}</div>}
          <div
            className={draftMode ? 'page draft' : 'page' + (showMarks ? ' show-marks' : '') + (outlineOnly ? ' outline' : '')}
            style={{
              maxWidth: draftMode ? '100%' : meta.orientation === 'portrait' ? pageWidth : Math.round(pageWidth * 1.41),
              padding: pad,
              background: meta.pageColor,
              columnCount: meta.columns,
              position: 'relative',
              overflow: 'hidden',
            }}
          >
            {meta.watermark && <div className="watermark">{meta.watermark}</div>}
            <div
              ref={ref}
              className={'doc' + (meta.hyphens ? ' hyphens' : '') + (meta.linenums ? ' linenums' : '')}
              contentEditable={!readOnly}
              spellCheck={spell}
              lang="ru"
              data-placeholder="Начните писать как в Word…"
              onInput={update}
              onPaste={onPaste}
              suppressContentEditableWarning
              style={{ zoom } as React.CSSProperties}
            />
          </div>
          {meta.footer && <div className="doc-footer">{meta.footer}{meta.pageNumbers ? ' — стр. 1' : ''}</div>}
        </div>
      </div>

      <div className="statusbar">
        <span>Слов: {stats.words}</span>
        <span>Символов: {stats.chars}</span>
        <span>Страниц: {pages}</span>
        <span>{meta.pageSize} {meta.orientation === 'portrait' ? 'книжная' : 'альбомная'}</span>
        <span>{zoom}%</span>
        {track && <span>● Запись исправлений</span>}
        {readOnly && <span>🔒 Только чтение</span>}
      </div>
    </div>
  );
}
