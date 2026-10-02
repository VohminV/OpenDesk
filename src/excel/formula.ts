// Движок формул OpenDesk Sheets — уровень Excel 2010 (базовый набор).
// Поддерживает: числа, строки "...", TRUE/FALSE, + - * / % ^ &,
// сравнения = <> < > <= >=, ссылки A1 (в т.ч. $A$1, Лист1!A1),
// диапазоны A1:B5, разделители аргументов , и ;,
// ~45 функций с английскими и русскими именами.
// Ошибки как в Excel: #DIV/0! #VALUE! #REF! #NAME? #N/A #NUM! #CYCLE!
// Даты — серийные номера Excel (1 = 01.01.1900).

export type GetCell = (ref: string) => number;

export function colToIndex(col: string): number {
  let n = 0;
  for (const ch of col.toUpperCase()) {
    n = n * 26 + (ch.charCodeAt(0) - 64);
  }
  return n - 1; // A -> 0
}

export function indexToCol(i: number): string {
  let s = '';
  let n = i + 1;
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export function parseRef(ref: string): { row: number; col: number } | null {
  const m = /^([A-Z]+)([1-9][0-9]*)$/i.exec(ref.trim());
  if (!m) return null;
  return { col: colToIndex(m[1]), row: parseInt(m[2], 10) - 1 };
}

export function expandRange(a: string, b: string): string[] {
  const pa = parseRef(a);
  const pb = parseRef(b);
  if (!pa || !pb) return [];
  const out: string[] = [];
  const c0 = Math.min(pa.col, pb.col);
  const c1 = Math.max(pa.col, pb.col);
  const r0 = Math.min(pa.row, pb.row);
  const r1 = Math.max(pa.row, pb.row);
  if ((c1 - c0 + 1) * (r1 - r0 + 1) > 10000) return [];
  for (let r = r0; r <= r1; r++) {
    for (let c = c0; c <= c1; c++) {
      out.push(`${indexToCol(c)}${r + 1}`);
    }
  }
  return out;
}

// ================= Модель значений =================

export class CErr {
  constructor(readonly code: string) {}
}

export const BLANK = { blank: true } as const;
export type V = number | string | boolean | typeof BLANK | CErr;

const isBlank = (v: V): v is typeof BLANK => v === BLANK;
const isErr = (v: V): v is CErr => v instanceof CErr;

function num(v: V): number {
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (isBlank(v)) return 0;
  if (isErr(v)) throw v;
  const n = Number(v);
  if (v.trim() !== '' && Number.isFinite(n)) return n;
  throw new CErr('#VALUE!');
}

function str(v: V): string {
  if (typeof v === 'string') return v;
  if (typeof v === 'number') return String(v);
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (isErr(v)) throw v;
  return '';
}

function truthy(v: V): boolean {
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (isErr(v)) throw v;
  if (isBlank(v)) return false;
  throw new CErr('#VALUE!');
}

// Порядок типов Excel: число < текст < логическое
function typeRank(v: V): number {
  if (typeof v === 'number' || isBlank(v)) return 0;
  if (typeof v === 'string') return 1;
  return 2;
}

function cmpEq(a: V, b: V): boolean {
  if (isErr(a)) throw a;
  if (isErr(b)) throw b;
  const ab = isBlank(a);
  const bb = isBlank(b);
  if (ab && bb) return true;
  if (ab) return typeof b === 'number' ? b === 0 : b === '';
  if (bb) return typeof a === 'number' ? a === 0 : a === '';
  const ra = typeRank(a);
  const rb = typeRank(b);
  if (ra !== rb) return false;
  if (typeof a === 'string' && typeof b === 'string') return a.toLowerCase() === b.toLowerCase();
  return a === b;
}

function cmpOrd(a: V, b: V): number {
  if (isErr(a)) throw a;
  if (isErr(b)) throw b;
  const na = isBlank(a) ? 0 : a;
  const nb = isBlank(b) ? 0 : b;
  const ra = typeRank(na as V);
  const rb = typeRank(nb as V);
  if (ra !== rb) return ra - rb;
  if (typeof na === 'string' && typeof nb === 'string') {
    const x = na.toLowerCase();
    const y = nb.toLowerCase();
    return x < y ? -1 : x > y ? 1 : 0;
  }
  if (typeof na === 'number' && typeof nb === 'number') return na - nb;
  return na === nb ? 0 : na ? 1 : -1;
}

// ================= Даты (серийные номера Excel) =================

// Эпоха 30.12.1899: Excel считает, что 1900-02-29 существовал (баг Lotus),
// поэтому все даты после 01.03.1900 сдвинуты на +1. Проверено: 01.01.2024 = 45292.
const EPOCH = Date.UTC(1899, 11, 30); // serial 0

export function dateToSerial(y: number, m: number, d: number): number {
  if (y >= 0 && y < 1900) y += 1900; // quirk Excel: 00-99 -> 1900-1999
  return Math.round((Date.UTC(y, m - 1, d) - EPOCH) / 86400000);
}

export function serialToDate(s: number): { y: number; m: number; d: number } {
  const t = new Date(EPOCH + Math.floor(s) * 86400000);
  return { y: t.getUTCFullYear(), m: t.getUTCMonth() + 1, d: t.getUTCDate() };
}

export function todaySerial(): number {
  const n = new Date();
  return dateToSerial(n.getFullYear(), n.getMonth() + 1, n.getDate());
}

// ================= Токенизатор =================

type Tok =
  | { t: 'num'; v: number }
  | { t: 'str'; v: string }
  | { t: 'ref'; sheet: string | null; col: number; row: number }
  | { t: 'name'; v: string }
  | { t: 'op'; v: string };

function tokenize(src: string): Tok[] {
  const out: Tok[] = [];
  let i = 0;
  const isLetter = (c: string) => /[A-Za-zА-Яа-яЁё_]/.test(c);
  while (i < src.length) {
    const c = src[i];
    if (c === ' ' || c === '\t' || c === '\n') {
      i++;
      continue;
    }
    // число
    if (/[0-9.]/.test(c) && /[0-9.]/.test(c)) {
      const m = /^[0-9]*\.?[0-9]+([eE][+-]?[0-9]+)?/.exec(src.slice(i));
      if (m && /[0-9]/.test(m[0])) {
        out.push({ t: 'num', v: Number(m[0]) });
        i += m[0].length;
        continue;
      }
    }
    // строка
    if (c === '"') {
      let s = '';
      i++;
      while (i < src.length) {
        if (src[i] === '"' && src[i + 1] === '"') {
          s += '"';
          i += 2;
        } else if (src[i] === '"') {
          i++;
          break;
        } else {
          s += src[i];
          i++;
        }
      }
      out.push({ t: 'str', v: s });
      continue;
    }
    // операторы из 2 символов
    const two = src.slice(i, i + 2);
    if (two === '>=' || two === '<=' || two === '<>') {
      out.push({ t: 'op', v: two });
      i += 2;
      continue;
    }
    if ('+-*/^&=<>():;,'.includes(c)) {
      out.push({ t: 'op', v: c === ':' ? ':' : c });
      i++;
      continue;
    }
    if (c === '(' || c === ')') {
      out.push({ t: 'op', v: c });
      i++;
      continue;
    }
    if (c === '%') {
      out.push({ t: 'op', v: '%' });
      i++;
      continue;
    }
    // имя листа в кавычках 'Мой лист'!
    let sheet: string | null = null;
    if (c === "'") {
      const m = /^'((?:[^']|'')+)'!/.exec(src.slice(i));
      if (m) {
        sheet = m[1].replace(/''/g, "'");
        i += m[0].length;
      } else {
        throw new CErr('#NAME?');
      }
    } else if (isLetter(c) || c === '$') {
      // возможно Лист1!A1 — смотрим вперёд до !
      const mSheet = /^((?:\$?[A-Za-zА-Яа-яЁё_][\wА-Яа-яЁё.]*))!/.exec(src.slice(i));
      if (mSheet) {
        sheet = mSheet[1].replace(/\$/g, '');
        i += mSheet[0].length;
      }
    }
    // ссылка $A$1
    const mRef = /^(\$?)([A-Za-z]+)(\$?)([0-9]+)/.exec(src.slice(i));
    if (mRef && !/^[A-Za-z(]/.test(src.slice(i + mRef[0].length).slice(0, 1))) {
      // защита: A1B -> не ссылка (дальше буква) — но A1: следующий ':' ок
      const after = src.slice(i + mRef[0].length, i + mRef[0].length + 1);
      if (after && /[A-Za-z]/.test(after)) {
        // не ссылка, а имя
      } else {
        out.push({ t: 'ref', sheet, col: colToIndex(mRef[2]), row: parseInt(mRef[4], 10) - 1 });
        i += mRef[0].length;
        continue;
      }
    } else if (mRef && sheet) {
      out.push({ t: 'ref', sheet, col: colToIndex(mRef[2]), row: parseInt(mRef[4], 10) - 1 });
      i += mRef[0].length;
      continue;
    }
    // имя (функция / TRUE / именованный диапазон)
    if (isLetter(c)) {
      const m = /^[A-Za-zА-Яа-яЁё_][\wА-Яа-яЁё.]*/.exec(src.slice(i));
      if (m) {
        out.push({ t: 'name', v: m[0] });
        i += m[0].length;
        continue;
      }
    }
    throw new CErr('#NAME?');
  }
  return out;
}

// ================= Парсер + вычислитель =================

type Ctx = {
  get: (ref: string) => V;
};

type Node =
  | { k: 'lit'; v: V }
  | { k: 'ref'; sheet: string | null; col: number; row: number }
  | { k: 'range'; sheet: string | null; r0: number; c0: number; r1: number; c1: number }
  | { k: 'un'; op: '+' | '-'; x: Node }
  | { k: 'bin'; op: string; a: Node; b: Node }
  | { k: 'call'; name: string; args: Node[] };

const BOOLS: Record<string, boolean> = { TRUE: true, FALSE: false, ИСТИНА: true, ЛОЖЬ: false };

class Parser {
  pos = 0;
  constructor(readonly toks: Tok[]) {}

  peek(): Tok | null {
    return this.toks[this.pos] ?? null;
  }

  eatOp(...ops: string[]): string | null {
    const t = this.peek();
    if (t && t.t === 'op' && ops.includes(t.v)) {
      this.pos++;
      return t.v;
    }
    return null;
  }

  parseExpr(): Node {
    return this.parseCmp();
  }

  parseCmp(): Node {
    let a = this.parseConcat();
    for (;;) {
      const t = this.peek();
      if (t && t.t === 'op' && ['=', '<>', '<', '>', '<=', '>='].includes(t.v)) {
        this.pos++;
        const b = this.parseConcat();
        a = { k: 'bin', op: t.v, a, b };
      } else return a;
    }
  }

  parseConcat(): Node {
    let a = this.parseAdd();
    while (this.eatOp('&')) {
      const b = this.parseAdd();
      a = { k: 'bin', op: '&', a, b };
    }
    return a;
  }

  parseAdd(): Node {
    let a = this.parseMul();
    for (;;) {
      const op = this.eatOp('+', '-');
      if (!op) return a;
      const b = this.parseMul();
      a = { k: 'bin', op, a, b };
    }
  }

  parseMul(): Node {
    let a = this.parsePow();
    for (;;) {
      const op = this.eatOp('*', '/');
      if (!op) return a;
      const b = this.parsePow();
      a = { k: 'bin', op, a, b };
    }
  }

  parsePow(): Node {
    const a = this.parseUn();
    if (this.eatOp('^')) {
      const b = this.parsePow();
      return { k: 'bin', op: '^', a, b };
    }
    return a;
  }

  parseUn(): Node {
    const op = this.eatOp('+', '-');
    if (op) {
      const x = this.parseUn();
      return op === '-' ? { k: 'un', op: '-', x } : x;
    }
    return this.parsePost();
  }

  parsePost(): Node {
    let a = this.parsePrimary();
    while (this.eatOp('%')) {
      a = { k: 'bin', op: '%', a, b: { k: 'lit', v: 0 } };
    }
    return a;
  }

  parsePrimary(): Node {
    const t = this.peek();
    if (!t) throw new CErr('#ERROR!');
    if (t.t === 'num') {
      this.pos++;
      return { k: 'lit', v: t.v };
    }
    if (t.t === 'str') {
      this.pos++;
      return { k: 'lit', v: t.v };
    }
    if (t.t === 'ref') {
      this.pos++;
      if (this.eatOp(':')) {
        const t2 = this.peek();
        if (!t2 || t2.t !== 'ref') throw new CErr('#REF!');
        if (t2.sheet && t.sheet && t2.sheet.toUpperCase() !== t.sheet.toUpperCase()) throw new CErr('#REF!');
        this.pos++;
        const sheet = t.sheet ?? t2.sheet;
        return {
          k: 'range',
          sheet,
          r0: Math.min(t.row, t2.row),
          c0: Math.min(t.col, t2.col),
          r1: Math.max(t.row, t2.row),
          c1: Math.max(t.col, t2.col),
        };
      }
      return { k: 'ref', sheet: t.sheet, col: t.col, row: t.row };
    }
    if (t.t === 'name') {
      const up = t.v.toUpperCase();
      if (up in BOOLS) {
        // TRUE/FALSE как значение, кроме вызова TRUE()/FALSE()
        const next = this.toks[this.pos + 1];
        if (!next || next.t !== 'op' || (next as { v: string }).v !== '(') {
          this.pos++;
          return { k: 'lit', v: BOOLS[up] };
        }
      }
      this.pos++;
      if (!this.eatOp('(')) throw new CErr('#NAME?');
      const args: Node[] = [];
      if (!this.eatOp(')')) {
        for (;;) {
          args.push(this.parseExpr());
          if (this.eatOp(',', ';')) continue;
          if (this.eatOp(')')) break;
          throw new CErr('#ERROR!');
        }
      }
      if (up in BOOLS && args.length === 0) return { k: 'lit', v: BOOLS[up] };
      return { k: 'call', name: up, args };
    }
    if (t.t === 'op' && t.v === '(') {
      this.pos++;
      const e = this.parseExpr();
      if (!this.eatOp(')')) throw new CErr('#ERROR!');
      return e;
    }
    throw new CErr('#ERROR!');
  }

  expectEnd() {
    if (this.pos < this.toks.length) throw new CErr('#ERROR!');
  }
}

// ключ ячейки для getRaw / цикла
function cellKey(sheet: string | null, col: number, row: number): string {
  const ref = `${indexToCol(col)}${row + 1}`;
  return sheet ? `${sheet.toUpperCase()}!${ref}` : ref;
}

function evalRange(ctx: Ctx, sheet: string | null, r0: number, c0: number, r1: number, c1: number): V[][] {
  const grid: V[][] = [];
  for (let r = r0; r <= r1; r++) {
    const row: V[] = [];
    for (let c = c0; c <= c1; c++) {
      row.push(ctx.get(cellKey(sheet, c, r)));
    }
    grid.push(row);
  }
  return grid;
}

const flat = (g: V[][]): V[] => g.flat();

function evalNode(n: Node, ctx: Ctx): V | V[][] {
  switch (n.k) {
    case 'lit':
      return n.v;
    case 'ref':
      return ctx.get(cellKey(n.sheet, n.col, n.row));
    case 'range':
      if ((n.r1 - n.r0 + 1) * (n.c1 - n.c0 + 1) > 10000) throw new CErr('#REF!');
      return evalRange(ctx, n.sheet, n.r0, n.c0, n.r1, n.c1);
    case 'un': {
      const v = evalNode(n.x, ctx);
      if (Array.isArray(v)) throw new CErr('#VALUE!');
      return -num(v);
    }
    case 'bin': {
      if (n.op === '%') {
        const v = evalNode(n.a, ctx);
        if (Array.isArray(v)) throw new CErr('#VALUE!');
        return num(v) / 100;
      }
      const av = evalNode(n.a, ctx);
      const bv = evalNode(n.b, ctx);
      if (Array.isArray(av) || Array.isArray(bv)) throw new CErr('#VALUE!');
      switch (n.op) {
        case '+':
          return num(av) + num(bv);
        case '-':
          return num(av) - num(bv);
        case '*':
          return num(av) * num(bv);
        case '/': {
          const d = num(bv);
          if (d === 0) throw new CErr('#DIV/0!');
          return num(av) / d;
        }
        case '^': {
          const r = Math.pow(num(av), num(bv));
          if (!Number.isFinite(r) || Number.isNaN(r)) throw new CErr('#NUM!');
          return Math.round(r * 1e10) / 1e10;
        }
        case '&':
          return str(av) + str(bv);
        case '=':
          return cmpEq(av, bv);
        case '<>':
          return !cmpEq(av, bv);
        case '<':
          return cmpOrd(av, bv) < 0;
        case '>':
          return cmpOrd(av, bv) > 0;
        case '<=':
          return cmpOrd(av, bv) <= 0;
        case '>=':
          return cmpOrd(av, bv) >= 0;
        default:
          throw new CErr('#ERROR!');
      }
    }
    case 'call':
      return evalCall(n.name, n.args, ctx);
  }
}

// ================= Функции =================

function numsOf(vals: V[]): number[] {
  const out: number[] = [];
  for (const v of vals) {
    if (isErr(v)) throw v;
    if (typeof v === 'number') out.push(v);
    else if (typeof v === 'boolean') out.push(v ? 1 : 0);
    // текст и пустые пропускаются (как в Excel для ссылок)
  }
  return out;
}

function resolveArgs(args: Node[], ctx: Ctx): V[] {
  const out: V[] = [];
  for (const a of args) {
    const v = evalNode(a, ctx);
    if (Array.isArray(v)) out.push(...flat(v));
    else out.push(v);
  }
  return out;
}

function matchCrit(v: V, crit: V): boolean {
  if (isErr(crit)) throw crit;
  if (typeof crit === 'string') {
    const m = /^(<=|>=|<>|=|<|>)(.*)$/s.exec(crit);
    if (m) {
      const op = m[1];
      const raw = m[2];
      const rhs: V = raw === '' ? '' : Number.isFinite(Number(raw)) && raw.trim() !== '' ? Number(raw) : raw;
      if (op === '=') return cmpEq(v, rhs);
      if (op === '<>') return !cmpEq(v, rhs);
      const c = cmpOrd(v, rhs);
      return op === '<' ? c < 0 : op === '>' ? c > 0 : op === '<=' ? c <= 0 : c >= 0;
    }
    return cmpEq(v, crit);
  }
  return cmpEq(v, crit);
}

function toTable(v: V | V[][]): V[][] {
  return Array.isArray(v) ? v : [[v]];
}

const FN: Record<string, (args: Node[], ctx: Ctx) => V | V[][]> = {
  SUM(args, ctx) {
    return numsOf(resolveArgs(args, ctx)).reduce((a, b) => a + b, 0);
  },
  СУММ(args, ctx) {
    return FN.SUM(args, ctx);
  },
  AVERAGE(args, ctx) {
    const n = numsOf(resolveArgs(args, ctx));
    if (!n.length) throw new CErr('#DIV/0!');
    return n.reduce((a, b) => a + b, 0) / n.length;
  },
  СРЗНАЧ(args, ctx) {
    return FN.AVERAGE(args, ctx);
  },
  MIN(args, ctx) {
    const n = numsOf(resolveArgs(args, ctx));
    return n.length ? Math.min(...n) : 0;
  },
  МИН(args, ctx) {
    return FN.MIN(args, ctx);
  },
  MAX(args, ctx) {
    const n = numsOf(resolveArgs(args, ctx));
    return n.length ? Math.max(...n) : 0;
  },
  МАКС(args, ctx) {
    return FN.MAX(args, ctx);
  },
  COUNT(args, ctx) {
    return resolveArgs(args, ctx).filter((v) => {
      if (isErr(v)) throw v;
      return typeof v === 'number';
    }).length;
  },
  СЧЁТ(args, ctx) {
    return FN.COUNT(args, ctx);
  },
  COUNTA(args, ctx) {
    return resolveArgs(args, ctx).filter((v) => {
      if (isErr(v)) throw v;
      return !isBlank(v);
    }).length;
  },
  СЧЁТЗ(args, ctx) {
    return FN.COUNTA(args, ctx);
  },
  COUNTBLANK(args, ctx) {
    if (args.length !== 1) throw new CErr('#N/A');
    const v = evalNode(args[0], ctx);
    return flat(toTable(v)).filter(isBlank).length;
  },
  СЧИТАТЬПУСТОТЫ(args, ctx) {
    return FN.COUNTBLANK(args, ctx);
  },
  COUNTIF(args, ctx) {
    if (args.length !== 2) throw new CErr('#N/A');
    const range = flat(toTable(evalNode(args[0], ctx)));
    const crit = evalNode(args[1], ctx);
    if (Array.isArray(crit)) throw new CErr('#VALUE!');
    return range.filter((v) => {
      if (isBlank(v)) return matchCrit('', crit);
      return matchCrit(v, crit);
    }).length;
  },
  СЧЁТЕСЛИ(args, ctx) {
    return FN.COUNTIF(args, ctx);
  },
  SUMIF(args, ctx) {
    if (args.length < 2 || args.length > 3) throw new CErr('#N/A');
    const range = flat(toTable(evalNode(args[0], ctx)));
    const crit = evalNode(args[1], ctx);
    if (Array.isArray(crit)) throw new CErr('#VALUE!');
    const sumRange = args[2] ? flat(toTable(evalNode(args[2], ctx))) : range;
    let s = 0;
    range.forEach((v, i) => {
      const test = isBlank(v) ? '' : v;
      if (matchCrit(test as V, crit)) {
        const sv = sumRange[Math.min(i, sumRange.length - 1)];
        if (isErr(sv)) throw sv;
        if (typeof sv === 'number') s += sv;
      }
    });
    return s;
  },
  СУММЕСЛИ(args, ctx) {
    return FN.SUMIF(args, ctx);
  },
  AVERAGEIF(args, ctx) {
    if (args.length < 2 || args.length > 3) throw new CErr('#N/A');
    const range = flat(toTable(evalNode(args[0], ctx)));
    const crit = evalNode(args[1], ctx);
    if (Array.isArray(crit)) throw new CErr('#VALUE!');
    const avgRange = args[2] ? flat(toTable(evalNode(args[2], ctx))) : range;
    const vals: number[] = [];
    range.forEach((v, i) => {
      const test = isBlank(v) ? '' : v;
      if (matchCrit(test as V, crit)) {
        const sv = avgRange[Math.min(i, avgRange.length - 1)];
        if (isErr(sv)) throw sv;
        if (typeof sv === 'number') vals.push(sv);
      }
    });
    if (!vals.length) throw new CErr('#DIV/0!');
    return vals.reduce((a, b) => a + b, 0) / vals.length;
  },
  СРЗНАЧЕСЛИ(args, ctx) {
    return FN.AVERAGEIF(args, ctx);
  },
  IF(args, ctx) {
    if (args.length < 2 || args.length > 3) throw new CErr('#N/A');
    const c = evalNode(args[0], ctx);
    if (Array.isArray(c)) throw new CErr('#VALUE!');
    const r = evalNode(truthy(c) ? args[1] : args[2] ?? { k: 'lit', v: false }, ctx);
    if (Array.isArray(r)) throw new CErr('#VALUE!');
    return r;
  },
  ЕСЛИ(args, ctx) {
    return FN.IF(args, ctx);
  },
  AND(args, ctx) {
    for (const a of args) {
      const v = evalNode(a, ctx);
      if (Array.isArray(v)) throw new CErr('#VALUE!');
      if (!truthy(v)) return false;
    }
    return true;
  },
  И(args, ctx) {
    return FN.AND(args, ctx);
  },
  OR(args, ctx) {
    for (const a of args) {
      const v = evalNode(a, ctx);
      if (Array.isArray(v)) throw new CErr('#VALUE!');
      if (truthy(v)) return true;
    }
    return false;
  },
  ИЛИ(args, ctx) {
    return FN.OR(args, ctx);
  },
  NOT(args, ctx) {
    if (args.length !== 1) throw new CErr('#N/A');
    const v = evalNode(args[0], ctx);
    if (Array.isArray(v)) throw new CErr('#VALUE!');
    return !truthy(v);
  },
  НЕ(args, ctx) {
    return FN.NOT(args, ctx);
  },
  IFERROR(args, ctx) {
    if (args.length !== 2) throw new CErr('#N/A');
    try {
      const v = evalNode(args[0], ctx);
      if (Array.isArray(v)) throw new CErr('#VALUE!');
      return v;
    } catch (e) {
      if (e instanceof CErr) {
        const fb = evalNode(args[1], ctx);
        if (Array.isArray(fb)) throw new CErr('#VALUE!');
        return fb;
      }
      throw e;
    }
  },
  ЕСЛИОШИБКА(args, ctx) {
    return FN.IFERROR(args, ctx);
  },
  VLOOKUP(args, ctx) {
    if (args.length < 3 || args.length > 4) throw new CErr('#N/A');
    const key = evalNode(args[0], ctx);
    if (Array.isArray(key)) throw new CErr('#VALUE!');
    const table = toTable(evalNode(args[1], ctx));
    const colV = evalNode(args[2], ctx);
    if (Array.isArray(colV)) throw new CErr('#VALUE!');
    const col = num(colV);
    if (!Number.isInteger(col) || col < 1) throw new CErr('#VALUE!');
    if (col > (table[0]?.length ?? 0)) throw new CErr('#REF!');
    const approx = args[3] ? evalNode(args[3], ctx) : true;
    if (Array.isArray(approx)) throw new CErr('#VALUE!');
    const exact = approx === false || approx === 0;
    if (exact) {
      for (const row of table) {
        if (cmpEq(isBlank(row[0]) ? '' : (row[0] as V), key)) return row[col - 1] ?? '';
      }
      throw new CErr('#N/A');
    }
    let best: V | null = null;
    for (const row of table) {
      const first = isBlank(row[0]) ? 0 : row[0];
      if (isErr(first as V)) throw first;
      if (typeof first === 'number' && typeof key === 'number') {
        if ((first as number) <= key) best = row[col - 1] ?? '';
        else break;
      } else if (cmpEq(first as V, key)) return row[col - 1] ?? '';
    }
    if (best == null) throw new CErr('#N/A');
    return best;
  },
  ВПР(args, ctx) {
    return FN.VLOOKUP(args, ctx);
  },
  HLOOKUP(args, ctx) {
    if (args.length < 3 || args.length > 4) throw new CErr('#N/A');
    const key = evalNode(args[0], ctx);
    if (Array.isArray(key)) throw new CErr('#VALUE!');
    const table = toTable(evalNode(args[1], ctx));
    const rowV = evalNode(args[2], ctx);
    if (Array.isArray(rowV)) throw new CErr('#VALUE!');
    const r = num(rowV);
    if (!Number.isInteger(r) || r < 1) throw new CErr('#VALUE!');
    if (r > table.length) throw new CErr('#REF!');
    const firstRow = table[0] ?? [];
    for (let i = 0; i < firstRow.length; i++) {
      if (cmpEq(isBlank(firstRow[i]) ? '' : firstRow[i], key)) return table[r - 1][i] ?? '';
    }
    throw new CErr('#N/A');
  },
  ГПР(args, ctx) {
    return FN.HLOOKUP(args, ctx);
  },
  INDEX(args, ctx) {
    if (args.length < 2 || args.length > 3) throw new CErr('#N/A');
    const table = toTable(evalNode(args[0], ctx));
    const rv = evalNode(args[1], ctx);
    if (Array.isArray(rv)) throw new CErr('#VALUE!');
    const r = num(rv);
    const cv = args[2] ? evalNode(args[2], ctx) : 1;
    if (Array.isArray(cv)) throw new CErr('#VALUE!');
    const c = num(cv);
    if (!Number.isInteger(r) || !Number.isInteger(c) || r < 1 || c < 1) throw new CErr('#VALUE!');
    if (r > table.length || c > (table[0]?.length ?? 0)) throw new CErr('#REF!');
    return table[r - 1][c - 1];
  },
  ИНДЕКС(args, ctx) {
    return FN.INDEX(args, ctx);
  },
  MATCH(args, ctx) {
    if (args.length < 2 || args.length > 3) throw new CErr('#N/A');
    const key = evalNode(args[0], ctx);
    if (Array.isArray(key)) throw new CErr('#VALUE!');
    const arr = flat(toTable(evalNode(args[1], ctx))).map((v) => (isBlank(v) ? '' : v)) as V[];
    const tv = args[2] ? evalNode(args[2], ctx) : 1;
    if (Array.isArray(tv)) throw new CErr('#VALUE!');
    const type = num(tv);
    if (type === 0) {
      const i = arr.findIndex((v) => cmpEq(v, key));
      if (i < 0) throw new CErr('#N/A');
      return i + 1;
    }
    // type 1: наибольшее <= key (возрастание)
    let best = -1;
    for (let i = 0; i < arr.length; i++) {
      try {
        if (cmpOrd(arr[i], key) <= 0) best = i;
        else break;
      } catch {
        throw new CErr('#N/A');
      }
    }
    if (best < 0) throw new CErr('#N/A');
    return best + 1;
  },
  ПОИСКПОЗ(args, ctx) {
    return FN.MATCH(args, ctx);
  },
  ROW(args, ctx) {
    if (args.length > 1) throw new CErr('#N/A');
    if (args.length === 0) throw new CErr('#VALUE!');
    const n = args[0];
    if (n.k === 'ref') return n.row + 1;
    if (n.k === 'range') return n.r0 + 1;
    evalNode(args[0], ctx); // пробросить ошибки значений
    throw new CErr('#VALUE!');
  },
  СТРОКА(args, ctx) {
    return FN.ROW(args, ctx);
  },
  COLUMN(args, ctx) {
    if (args.length > 1) throw new CErr('#N/A');
    if (args.length === 0) throw new CErr('#VALUE!');
    const n = args[0];
    if (n.k === 'range') return n.c0 + 1;
    if (n.k === 'ref') return n.col + 1;
    evalNode(args[0], ctx);
    throw new CErr('#VALUE!');
  },
  СТОЛБЕЦ(args, ctx) {
    return FN.COLUMN(args, ctx);
  },
  LEN(args, ctx) {
    if (args.length !== 1) throw new CErr('#N/A');
    const v = evalNode(args[0], ctx);
    if (Array.isArray(v)) throw new CErr('#VALUE!');
    return str(v).length;
  },
  ДЛСТР(args, ctx) {
    return FN.LEN(args, ctx);
  },
  LEFT(args, ctx) {
    if (args.length < 1 || args.length > 2) throw new CErr('#N/A');
    const v = evalNode(args[0], ctx);
    if (Array.isArray(v)) throw new CErr('#VALUE!');
    const n = args[1] ? num(evalNode(args[1], ctx) as V) : 1;
    return str(v).slice(0, Math.max(0, Math.floor(n)));
  },
  ЛЕВСИМВ(args, ctx) {
    return FN.LEFT(args, ctx);
  },
  RIGHT(args, ctx) {
    if (args.length < 1 || args.length > 2) throw new CErr('#N/A');
    const v = evalNode(args[0], ctx);
    if (Array.isArray(v)) throw new CErr('#VALUE!');
    const n = args[1] ? num(evalNode(args[1], ctx) as V) : 1;
    const s = str(v);
    return s.slice(Math.max(0, s.length - Math.floor(n)));
  },
  ПРАВСИМВ(args, ctx) {
    return FN.RIGHT(args, ctx);
  },
  MID(args, ctx) {
    if (args.length !== 3) throw new CErr('#N/A');
    const v = evalNode(args[0], ctx);
    if (Array.isArray(v)) throw new CErr('#VALUE!');
    const start = num(evalNode(args[1], ctx) as V);
    const n = num(evalNode(args[2], ctx) as V);
    if (start < 1 || n < 0) throw new CErr('#VALUE!');
    return str(v).slice(Math.floor(start) - 1, Math.floor(start) - 1 + Math.floor(n));
  },
  ПСТР(args, ctx) {
    return FN.MID(args, ctx);
  },
  UPPER(args, ctx) {
    if (args.length !== 1) throw new CErr('#N/A');
    const v = evalNode(args[0], ctx);
    if (Array.isArray(v)) throw new CErr('#VALUE!');
    return str(v).toUpperCase();
  },
  ПРОПИСН(args, ctx) {
    return FN.UPPER(args, ctx);
  },
  LOWER(args, ctx) {
    if (args.length !== 1) throw new CErr('#N/A');
    const v = evalNode(args[0], ctx);
    if (Array.isArray(v)) throw new CErr('#VALUE!');
    return str(v).toLowerCase();
  },
  СТРОЧН(args, ctx) {
    return FN.LOWER(args, ctx);
  },
  TRIM(args, ctx) {
    if (args.length !== 1) throw new CErr('#N/A');
    const v = evalNode(args[0], ctx);
    if (Array.isArray(v)) throw new CErr('#VALUE!');
    return str(v).trim().replace(/\s+/g, ' ');
  },
  СЖПРОБЕЛЫ(args, ctx) {
    return FN.TRIM(args, ctx);
  },
  CONCATENATE(args, ctx) {
    return resolveArgs(args, ctx)
      .map((v) => {
        if (isErr(v)) throw v;
        return str(v);
      })
      .join('');
  },
  СЦЕПИТЬ(args, ctx) {
    return FN.CONCATENATE(args, ctx);
  },
  TEXT(args, ctx) {
    if (args.length !== 2) throw new CErr('#N/A');
    const v = evalNode(args[0], ctx);
    const f = evalNode(args[1], ctx);
    if (Array.isArray(v) || Array.isArray(f)) throw new CErr('#VALUE!');
    return applyTextFormat(v, str(f));
  },
  ТЕКСТ(args, ctx) {
    return FN.TEXT(args, ctx);
  },
  HYPERLINK(args, ctx) {
    if (args.length < 1 || args.length > 2) throw new CErr('#N/A');
    const url = evalNode(args[0], ctx);
    if (Array.isArray(url)) throw new CErr('#VALUE!');
    if (args[1]) {
      const text = evalNode(args[1], ctx);
      if (Array.isArray(text)) throw new CErr('#VALUE!');
      return str(text);
    }
    return str(url);
  },
  ГИПЕРССЫЛКА(args, ctx) {
    return FN.HYPERLINK(args, ctx);
  },
  VALUE(args, ctx) {
    if (args.length !== 1) throw new CErr('#N/A');
    const v = evalNode(args[0], ctx);
    if (Array.isArray(v)) throw new CErr('#VALUE!');
    if (typeof v === 'number') return v;
    const n = Number(String(v).trim().replace(',', '.'));
    if (!Number.isFinite(n) || String(v).trim() === '') throw new CErr('#VALUE!');
    return n;
  },
  ЗНАЧЕН(args, ctx) {
    return FN.VALUE(args, ctx);
  },
  ABS(args, ctx) {
    if (args.length !== 1) throw new CErr('#N/A');
    const v = evalNode(args[0], ctx);
    if (Array.isArray(v)) throw new CErr('#VALUE!');
    return Math.abs(num(v));
  },
  ROUND(args, ctx) {
    if (args.length !== 2) throw new CErr('#N/A');
    const x = num(evalNode(args[0], ctx) as V);
    const n = Math.trunc(num(evalNode(args[1], ctx) as V));
    const f = Math.pow(10, n);
    const r = Math.sign(x) * Math.round(Math.abs(x) * f) / f;
    return n > 0 ? r : Math.round(r);
  },
  ОКРУГЛ(args, ctx) {
    return FN.ROUND(args, ctx);
  },
  INT(args, ctx) {
    if (args.length !== 1) throw new CErr('#N/A');
    return Math.floor(num(evalNode(args[0], ctx) as V));
  },
  ЦЕЛОЕ(args, ctx) {
    return FN.INT(args, ctx);
  },
  MOD(args, ctx) {
    if (args.length !== 2) throw new CErr('#N/A');
    const a = num(evalNode(args[0], ctx) as V);
    const b = num(evalNode(args[1], ctx) as V);
    if (b === 0) throw new CErr('#DIV/0!');
    return a - b * Math.floor(a / b);
  },
  ОСТАТ(args, ctx) {
    return FN.MOD(args, ctx);
  },
  POWER(args, ctx) {
    if (args.length !== 2) throw new CErr('#N/A');
    const r = Math.pow(num(evalNode(args[0], ctx) as V), num(evalNode(args[1], ctx) as V));
    if (!Number.isFinite(r) || Number.isNaN(r)) throw new CErr('#NUM!');
    return r;
  },
  СТЕПЕНЬ(args, ctx) {
    return FN.POWER(args, ctx);
  },
  SQRT(args, ctx) {
    if (args.length !== 1) throw new CErr('#N/A');
    const x = num(evalNode(args[0], ctx) as V);
    if (x < 0) throw new CErr('#NUM!');
    return Math.sqrt(x);
  },
  КОРЕНЬ(args, ctx) {
    return FN.SQRT(args, ctx);
  },
  TODAY(args) {
    if (args.length) throw new CErr('#N/A');
    return todaySerial();
  },
  СЕГОДНЯ(args, ctx) {
    return FN.TODAY(args, ctx);
  },
  NOW(args) {
    if (args.length) throw new CErr('#N/A');
    const n = new Date();
    return todaySerial() + (n.getHours() * 3600 + n.getMinutes() * 60 + n.getSeconds()) / 86400;
  },
  ТДАТА(args, ctx) {
    return FN.NOW(args, ctx);
  },
  DATE(args, ctx) {
    if (args.length !== 3) throw new CErr('#N/A');
    return dateToSerial(
      Math.trunc(num(evalNode(args[0], ctx) as V)),
      Math.trunc(num(evalNode(args[1], ctx) as V)),
      Math.trunc(num(evalNode(args[2], ctx) as V)),
    );
  },
  ДАТА(args, ctx) {
    return FN.DATE(args, ctx);
  },
  YEAR(args, ctx) {
    if (args.length !== 1) throw new CErr('#N/A');
    return serialToDate(num(evalNode(args[0], ctx) as V)).y;
  },
  ГОД(args, ctx) {
    return FN.YEAR(args, ctx);
  },
  MONTH(args, ctx) {
    if (args.length !== 1) throw new CErr('#N/A');
    return serialToDate(num(evalNode(args[0], ctx) as V)).m;
  },
  МЕСЯЦ(args, ctx) {
    return FN.MONTH(args, ctx);
  },
  DAY(args, ctx) {
    if (args.length !== 1) throw new CErr('#N/A');
    return serialToDate(num(evalNode(args[0], ctx) as V)).d;
  },
  ДЕНЬ(args, ctx) {
    return FN.DAY(args, ctx);
  },
};

function evalCall(name: string, args: Node[], ctx: Ctx): V | V[][] {
  const fn = FN[name];
  if (!fn) throw new CErr('#NAME?');
  return fn(args, ctx);
}

// Формат ТЕКСТ(): подмножество — decimals, %, даты
function applyTextFormat(v: V, fmt: string): string {
  if (isErr(v)) throw v;
  if (/[yYдДmMдДhH]/.test(fmt) && typeof v === 'number') {
    const { y, m, d } = serialToDate(v);
    const p = (n: number) => String(n).padStart(2, '0');
    return fmt
      .replace(/YYYY|ГГГГ/i, String(y))
      .replace(/YY|ГГ/i, String(y).slice(2))
      .replace(/MM/i, p(m))
      .replace(/DD/i, p(d));
  }
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'string') return v;
  if (isBlank(v)) return '';
  let n = v as number;
  if (fmt.includes('%')) n *= 100;
  const decMatch = /\.([0#]+)/.exec(fmt);
  const dec = decMatch ? decMatch[1].length : 0;
  let s = dec > 0 ? n.toFixed(dec) : String(Math.round(n));
  if (fmt.includes('%')) s += '%';
  return s;
}

// ================= Публичный API =================

function toDisplay(v: V | V[][]): number | string {
  if (Array.isArray(v)) throw new CErr('#VALUE!');
  if (isErr(v)) return v.code;
  if (isBlank(v)) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'number') {
    if (!Number.isFinite(v)) return '#DIV/0!';
    return Math.round(v * 1e10) / 1e10;
  }
  return v;
}

function evalFormulaInner(raw: string, getRaw: (ref: string) => string | undefined, seen: Set<string>): V | V[][] {
  const ctx: Ctx = {
    get: (ref: string) => {
      const key = ref.toUpperCase();
      if (seen.has(key)) throw new CErr('#CYCLE!');
      const rawV = getRaw(key);
      if (rawV == null || rawV === '') return BLANK;
      if (rawV.startsWith('=')) {
        seen.add(key);
        try {
          const inner = evalFormulaInner(rawV, getRaw, seen);
          if (Array.isArray(inner)) throw new CErr('#VALUE!');
          return inner;
        } finally {
          seen.delete(key);
        }
      }
      const t = rawV.trim();
      if (/^(true|истина)$/i.test(t)) return true;
      if (/^(false|ложь)$/i.test(t)) return false;
      const n = Number(t);
      if (t !== '' && Number.isFinite(n)) return n;
      return rawV;
    },
  };
  const toks = tokenize(raw.slice(1));
  if (!toks.length) return '';
  const p = new Parser(toks);
  const node = p.parseExpr();
  p.expectEnd();
  return evalNode(node, ctx);
}

export function evaluateFormula(
  raw: string,
  getRaw: (ref: string) => string | undefined,
  seen: Set<string> = new Set(),
): number | string {
  const formula = raw.trim();
  if (!formula.startsWith('=')) {
    const n = Number(formula);
    return formula === '' ? '' : Number.isFinite(n) && formula !== '' ? n : formula;
  }
  try {
    return toDisplay(evalFormulaInner(formula, getRaw, seen));
  } catch (e: unknown) {
    if (e instanceof CErr) return e.code;
    return '#ERROR!';
  }
}

// Сдвиг ссылок в формуле при вставке/удалении строк/столбцов.
// action: {type:'row'|'col', index: 0-based, delta: +1 вставка / -N удаление}
// Внутренняя группа имени листа — non-capturing, иначе съезжают позиции колбэка.
// Хвост (?![A-Za-z(]) не трогает имена функций с цифрами (LOG10() и т.п.).
const REF_RE = /(?:('(?:[^']|'')+'|[A-Za-zА-Яа-яЁё_][\wА-Яа-яЁё.]*)!)?(\$?)([A-Za-z]+)(\$?)([0-9]+)(?![A-Za-z(])/g;

export function shiftFormulaRefs(
  formula: string,
  action: { type: 'row' | 'col'; sheet: string | null; index: number; delta: number },
): string {
  if (!formula.startsWith('=')) return formula;
  return formula.replace(REF_RE, (full, sheetPart: string | undefined, absCol: string, col: string, absRow: string, row: string) => {
    const refSheet = sheetPart ? sheetPart.replace(/!$/, '').replace(/^'|'$/g, '').replace(/''/g, "'") : null;
    if (action.sheet && refSheet && refSheet.toUpperCase() !== action.sheet.toUpperCase()) return full;
    if (!action.sheet && refSheet) return full;
    const c = colToIndex(col);
    const r = parseInt(row, 10) - 1;
    // удалённый диапазон [index, index-delta) -> #REF! (как в Excel)
    if (action.type === 'row' && absRow === '' && r >= action.index) {
      if (action.delta < 0 && r < action.index - action.delta) return '#REF!';
      const nr = r + action.delta;
      if (nr < 0) return '#REF!';
      return `${sheetPart ?? ''}${absCol}${col}${absRow}${nr + 1}`;
    }
    if (action.type === 'col' && absCol === '' && c >= action.index) {
      if (action.delta < 0 && c < action.index - action.delta) return '#REF!';
      const nc = c + action.delta;
      if (nc < 0) return '#REF!';
      return `${sheetPart ?? ''}${absCol}${indexToCol(nc)}${absRow}${row}`;
    }
    return full;
  });
}
