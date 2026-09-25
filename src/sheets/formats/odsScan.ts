/**
 * A fast reader for the sheets in an OpenDocument spreadsheet's content.xml.
 *
 * Rows and cells are nearly all of a large file, and building a DOM for them is slow. So the reader
 * cuts each sheet's contents out of the XML text (the DOM then only holds styles, validations and
 * names) and streams through them with a small tokenizer; the few other things in a sheet (pictures,
 * charts, conditional formats) are parsed on their own.
 */

export interface CellAttrs {
  repeat: number;
  style?: string;
  validation?: string;
  formula?: string;
  type?: string;
  calcType?: string;
  value?: string;
  dateValue?: string;
  timeValue?: string;
  boolValue?: string;
  stringValue?: string;
  spanC: number;
  spanR: number;
  mxC: number;
  mxR: number;
}

export interface CellData {
  a: CellAttrs;
  /** Number of text paragraphs. */
  paras: number;
  /** True when the paragraphs hold more than plain text (links, spans, spaces). */
  marked: boolean;
  note?: string;
  /** Pictures and charts anchored in the cell. */
  frames: () => Element[];
  text: () => { text: string; link?: string };
}

export interface RowData {
  repeat: number;
  style?: string;
  visibility?: string;
  defaultStyle?: string;
  cells: CellData[];
}

export interface ColData {
  repeat: number;
  style?: string;
  visibility?: string;
  defaultStyle?: string;
}

export interface RowHandlers {
  column(col: ColData): void;
  row(row: RowData): void;
  /** Start and end of the rows repeated on printed pages. */
  headerRows(start: boolean): void;
  /** Anything else in the sheet (shapes, names, conditional formats), parsed on its own. */
  extra(el: Element): void;
}

export const NS_URI = {
  office: 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
  table: 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
  text: 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
  draw: 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0',
  calcext: 'urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0',
  xlink: 'http://www.w3.org/1999/xlink',
};

class Fallback extends Error {}

const ENTITIES: Record<string, string> = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };

export function decodeEntities(s: string): string {
  if (s.indexOf('&') < 0) return s;
  return s.replace(/&(#[xX][0-9a-fA-F]+|#\d+|[A-Za-z]+);/g, (m, e: string) => {
    if (e[0] !== '#') return ENTITIES[e] ?? m;
    const code = e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
    return Number.isFinite(code) && code > 0 && code <= 0x10ffff ? String.fromCodePoint(code) : m;
  });
}

const enum Tok {
  End,
  Open,
  Close,
  Empty,
  Text,
  Skip,
}

/** A minimal XML tokenizer over a string (elements, attributes and text; no DTDs). */
class Scanner {
  pos: number;
  kind: Tok = Tok.End;
  name = '';
  start = 0;
  textStart = 0;
  textEnd = 0;
  cdata = false;
  private names: string[] = [];
  private values: string[] = [];
  count = 0;

  constructor(
    readonly s: string,
    from: number,
    readonly end: number,
    /** Namespace declarations are only expected on the root element. */
    readonly allowXmlns = false,
  ) {
    this.pos = from;
  }

  next(): Tok {
    const s = this.s;
    let i = this.pos;
    if (i >= this.end) return (this.kind = Tok.End);
    this.start = i;
    if (s.charCodeAt(i) !== 60) {
      const j = s.indexOf('<', i);
      this.textStart = i;
      this.textEnd = j < 0 || j > this.end ? this.end : j;
      this.pos = this.textEnd;
      this.cdata = false;
      return (this.kind = Tok.Text);
    }
    const c1 = s.charCodeAt(i + 1);
    if (c1 === 47) {
      const j = s.indexOf('>', i + 2);
      if (j < 0) throw new Fallback();
      this.name = s.slice(i + 2, j).trim();
      this.pos = j + 1;
      return (this.kind = Tok.Close);
    }
    if (c1 === 33) {
      if (s.startsWith('<!--', i)) {
        const j = s.indexOf('-->', i + 4);
        this.pos = j < 0 ? this.end : j + 3;
        return (this.kind = Tok.Skip);
      }
      if (s.startsWith('<![CDATA[', i)) {
        const j = s.indexOf(']]>', i + 9);
        if (j < 0) throw new Fallback();
        this.textStart = i + 9;
        this.textEnd = j;
        this.cdata = true;
        this.pos = j + 3;
        return (this.kind = Tok.Text);
      }
      throw new Fallback();
    }
    if (c1 === 63) {
      const j = s.indexOf('?>', i + 2);
      this.pos = j < 0 ? this.end : j + 2;
      return (this.kind = Tok.Skip);
    }
    // start tag
    let j = i + 1;
    let c = s.charCodeAt(j);
    while (j < this.end && c !== 32 && c !== 9 && c !== 10 && c !== 13 && c !== 47 && c !== 62) c = s.charCodeAt(++j);
    this.name = s.slice(i + 1, j);
    this.count = 0;
    for (;;) {
      while (c === 32 || c === 9 || c === 10 || c === 13) c = s.charCodeAt(++j);
      if (c === 62) {
        this.pos = j + 1;
        return (this.kind = Tok.Open);
      }
      if (c === 47) {
        if (s.charCodeAt(j + 1) !== 62) throw new Fallback();
        this.pos = j + 2;
        return (this.kind = Tok.Empty);
      }
      if (j >= this.end) throw new Fallback();
      const eq = s.indexOf('=', j);
      if (eq < 0) throw new Fallback();
      const an = s.slice(j, eq).trim();
      // namespace declarations inside the body would change what prefixes mean
      if (!this.allowXmlns && an.startsWith('xmlns')) throw new Fallback();
      let q = eq + 1;
      while (s.charCodeAt(q) === 32 || s.charCodeAt(q) === 9 || s.charCodeAt(q) === 10 || s.charCodeAt(q) === 13) q++;
      const quote = s[q];
      if (quote !== '"' && quote !== "'") throw new Fallback();
      const ve = s.indexOf(quote, q + 1);
      if (ve < 0) throw new Fallback();
      this.names[this.count] = an;
      this.values[this.count] = s.slice(q + 1, ve);
      this.count++;
      j = ve + 1;
      c = s.charCodeAt(j);
    }
  }

  attrName(i: number): string {
    return this.names[i];
  }

  attrValue(i: number): string {
    return decodeEntities(this.values[i]);
  }

  attr(qname: string): string | undefined {
    for (let i = 0; i < this.count; i++) if (this.names[i] === qname) return decodeEntities(this.values[i]);
    return undefined;
  }

  text(): string {
    const t = this.s.slice(this.textStart, this.textEnd);
    return this.cdata ? t : decodeEntities(t);
  }

  /** After an opening tag: moves past the matching end tag. */
  skipElement(): void {
    let depth = 1;
    while (depth > 0) {
      const k = this.next();
      if (k === Tok.Open) depth++;
      else if (k === Tok.Close) depth--;
      else if (k === Tok.End) throw new Fallback();
    }
  }
}

/** Namespace prefixes declared on the root element, by namespace. */
export type Prefixes = Record<keyof typeof NS_URI, string>;

export interface SplitContent {
  /** content.xml with the contents of every sheet cut out. */
  reduced: string;
  /** For each sheet in order: where its contents were in the original text. */
  regions: ({ start: number; end: number } | null)[];
  prefixes: Prefixes;
  /** Namespace declarations of the root element (to parse fragments on their own). */
  xmlns: string;
}

/** The next start tag with this exact name (not a longer one such as table:table-row). */
function indexOfTag(s: string, tag: string, from: number, until: number): number {
  const a = s.indexOf(`${tag} `, from);
  const b = s.indexOf(`${tag}>`, from);
  const k = a < 0 ? b : b < 0 ? a : Math.min(a, b);
  return k >= 0 && k < until ? k : -1;
}

/**
 * Finds each sheet's contents in content.xml with plain text searches. Returns null when the file uses
 * something the fast path doesn't handle, so the caller can use the DOM for everything.
 */
export function splitContent(xml: string, flat = false): SplitContent | null {
  try {
    return split(xml, flat);
  } catch (e) {
    if (e instanceof Fallback) return null;
    throw e;
  }
}

function split(xml: string, flat: boolean): SplitContent | null {
  const rootStart = xml.indexOf('<', xml.startsWith('<?xml') ? xml.indexOf('?>') + 2 : 0);
  if (rootStart < 0) return null;
  const root = new Scanner(xml, rootStart, xml.length, true);
  if (root.next() !== Tok.Open) return null;
  const prefixes: Partial<Prefixes> = {};
  const decls: string[] = [];
  for (let i = 0; i < root.count; i++) {
    const name = root.attrName(i);
    if (name === 'xmlns') return null; // a default namespace
    if (!name.startsWith('xmlns:')) continue;
    const uri = root.attrValue(i);
    decls.push(`${name}="${uri.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;')}"`);
    for (const [key, u] of Object.entries(NS_URI)) if (uri === u) prefixes[key as keyof Prefixes] = name.slice(6);
  }
  if (!prefixes.office || !prefixes.table || !prefixes.text) return null;
  const P = { calcext: 'calcext', draw: 'draw', xlink: 'xlink', ...prefixes } as Prefixes;
  const T = P.table;
  const body = indexOfTag(xml, `<${P.office}:spreadsheet`, root.pos, xml.length);
  if (body < 0) return null;
  const spreadsheetEnd = xml.lastIndexOf(`</${P.office}:spreadsheet>`);
  if (spreadsheetEnd < body) return null;
  // charts stored inside flat files carry tables of their own: leave those to the DOM
  if (flat && indexOfTag(xml, `<${P.office}:document`, body, spreadsheetEnd) >= 0) return null;
  const regions: SplitContent['regions'] = [];
  const pieces: string[] = [];
  let copied = 0;
  const close = `</${T}:table>`;
  let k = indexOfTag(xml, `<${T}:table`, body, spreadsheetEnd);
  while (k >= 0) {
    const tag = new Scanner(xml, k, spreadsheetEnd);
    const kind = tag.next();
    if (kind === Tok.Empty) {
      regions.push(null);
      k = indexOfTag(xml, `<${T}:table`, tag.pos, spreadsheetEnd);
      continue;
    }
    if (kind !== Tok.Open) return null;
    const end = xml.indexOf(close, tag.pos);
    if (end < 0 || end > spreadsheetEnd) return null;
    regions.push({ start: tag.pos, end });
    pieces.push(xml.slice(copied, tag.pos));
    copied = end;
    k = indexOfTag(xml, `<${T}:table`, end + close.length, spreadsheetEnd);
  }
  pieces.push(xml.slice(copied));
  return { reduced: pieces.join(''), regions, prefixes: P, xmlns: decls.join(' ') };
}

/**
 * Streams the contents of one sheet: rows and columns go to the handlers as they come, anything else
 * is parsed on its own. Throws when it meets something unexpected; the caller then reads the file
 * through the DOM instead.
 */
export function scanRegion(xml: string, region: { start: number; end: number }, P: Prefixes, xmlns: string, h: RowHandlers): void {
  const sc = new Scanner(xml, region.start, region.end);
  const T = P.table;
  const TEXT = P.text;
  const names = {
    column: `${T}:table-column`,
    row: `${T}:table-row`,
    cell: `${T}:table-cell`,
    covered: `${T}:covered-table-cell`,
    headerRows: `${T}:table-header-rows`,
    rowGroups: new Set([`${T}:table-rows`, `${T}:table-row-group`]),
    colGroups: new Set([`${T}:table-columns`, `${T}:table-column-group`, `${T}:table-header-columns`]),
    p: `${TEXT}:p`,
    h: `${TEXT}:h`,
    s: `${TEXT}:s`,
    tab: `${TEXT}:tab`,
    br: `${TEXT}:line-break`,
    a: `${TEXT}:a`,
    note: `${P.office}:annotation`,
    frame: `${P.draw}:frame`,
    href: `${P.xlink}:href`,
    c: `${TEXT}:c`,
  };
  const A = {
    repeatCols: `${T}:number-columns-repeated`,
    repeatRows: `${T}:number-rows-repeated`,
    style: `${T}:style-name`,
    visibility: `${T}:visibility`,
    defaultStyle: `${T}:default-cell-style-name`,
    validation: `${T}:content-validation-name`,
    formula: `${T}:formula`,
    spanC: `${T}:number-columns-spanned`,
    spanR: `${T}:number-rows-spanned`,
    mxC: `${T}:number-matrix-columns-spanned`,
    mxR: `${T}:number-matrix-rows-spanned`,
    type: `${P.office}:value-type`,
    calcType: `${P.calcext}:value-type`,
    value: `${P.office}:value`,
    dateValue: `${P.office}:date-value`,
    timeValue: `${P.office}:time-value`,
    boolValue: `${P.office}:boolean-value`,
    stringValue: `${P.office}:string-value`,
  };
  const int = (v: string | undefined) => (v === undefined ? 1 : Math.max(1, parseInt(v, 10) || 1));
  const parser = new DOMParser();
  const fragment = (x: string): Element[] => {
    const d = parser.parseFromString(`<r ${xmlns}>${x}</r>`, 'application/xml');
    if (d.getElementsByTagName('parsererror').length) return [];
    const out: Element[] = [];
    for (let n = d.documentElement.firstChild; n; n = n.nextSibling) if (n.nodeType === 1) out.push(n as Element);
    return out;
  };

  /** Text of the element just opened (paragraph, span…), with spaces, tabs, breaks and links. */
  const readText = (into: { text: string; link?: string }) => {
    let depth = 1;
    while (depth > 0) {
      const k = sc.next();
      if (k === Tok.Text) {
        const t = sc.text();
        into.text += /[\r\n\t]/.test(t) ? t.replace(/[\r\n\t]+/g, ' ') : t;
      } else if (k === Tok.Open || k === Tok.Empty) {
        const n = sc.name;
        if (n === names.s) into.text += ' '.repeat(int(sc.attr(names.c)));
        else if (n === names.tab) into.text += '\t';
        else if (n === names.br) into.text += '\n';
        else if (n === names.a && into.link === undefined) into.link = sc.attr(names.href);
        if (k === Tok.Open) {
          if (n === names.note || n === names.frame) sc.skipElement();
          else depth++;
        }
      } else if (k === Tok.Close) depth--;
      else if (k === Tok.End) throw new Fallback();
    }
  };

  const readCell = (): CellData => {
    const a: CellAttrs = { repeat: 1, spanC: 1, spanR: 1, mxC: 1, mxR: 1 };
    for (let i = 0; i < sc.count; i++) {
      const n = sc.attrName(i);
      switch (n) {
        case A.repeatCols:
          a.repeat = int(sc.attrValue(i));
          break;
        case A.style:
          a.style = sc.attrValue(i);
          break;
        case A.validation:
          a.validation = sc.attrValue(i);
          break;
        case A.formula:
          a.formula = sc.attrValue(i);
          break;
        case A.spanC:
          a.spanC = int(sc.attrValue(i));
          break;
        case A.spanR:
          a.spanR = int(sc.attrValue(i));
          break;
        case A.mxC:
          a.mxC = int(sc.attrValue(i));
          break;
        case A.mxR:
          a.mxR = int(sc.attrValue(i));
          break;
        case A.type:
          a.type = sc.attrValue(i);
          break;
        case A.calcType:
          a.calcType = sc.attrValue(i);
          break;
        case A.value:
          a.value = sc.attrValue(i);
          break;
        case A.dateValue:
          a.dateValue = sc.attrValue(i);
          break;
        case A.timeValue:
          a.timeValue = sc.attrValue(i);
          break;
        case A.boolValue:
          a.boolValue = sc.attrValue(i);
          break;
        case A.stringValue:
          a.stringValue = sc.attrValue(i);
          break;
        default:
          break;
      }
    }
    const texts: { text: string; link?: string } = { text: '' };
    let paras = 0;
    let marked = false;
    let note: string | undefined;
    const frames: string[] = [];
    if (sc.kind === Tok.Open) {
      for (;;) {
        const k = sc.next();
        if (k === Tok.Close) break;
        if (k === Tok.End) throw new Fallback();
        if (k !== Tok.Open && k !== Tok.Empty) continue;
        const n = sc.name;
        if (n === names.p || n === names.h) {
          if (paras++) texts.text += '\n';
          if (k === Tok.Open) {
            const before = sc.pos;
            readText(texts);
            // anything but a single run of text counts as markup
            if (!marked && xml.lastIndexOf('<', sc.start - 1) >= before) marked = true;
          }
        } else if (n === names.note) {
          if (k === Tok.Open) {
            const t: { text: string } = { text: '' };
            let lines = 0;
            let depth = 1;
            while (depth > 0) {
              const kk = sc.next();
              if (kk === Tok.Open || kk === Tok.Empty) {
                if (sc.name === names.p || sc.name === names.h) {
                  if (lines++) t.text += '\n';
                  if (kk === Tok.Open) readText(t);
                } else if (kk === Tok.Open) sc.skipElement();
              } else if (kk === Tok.Close) depth--;
              else if (kk === Tok.End) throw new Fallback();
            }
            note = t.text;
          } else note = '';
        } else if (n === names.frame) {
          const from = sc.start;
          if (k === Tok.Open) sc.skipElement();
          frames.push(xml.slice(from, sc.pos));
        } else if (k === Tok.Open) sc.skipElement();
      }
    }
    return {
      a,
      paras,
      marked,
      note,
      frames: () => (frames.length ? frames.flatMap(fragment) : []),
      text: () => texts,
    };
  };

  const readRow = () => {
    const row: RowData = { repeat: int(sc.attr(A.repeatRows)), style: sc.attr(A.style), visibility: sc.attr(A.visibility), defaultStyle: sc.attr(A.defaultStyle), cells: [] };
    if (sc.kind === Tok.Open)
      for (;;) {
        const k = sc.next();
        if (k === Tok.Close) break;
        if (k === Tok.End) throw new Fallback();
        if (k !== Tok.Open && k !== Tok.Empty) continue;
        if (sc.name === names.cell || sc.name === names.covered) row.cells.push(readCell());
        else if (k === Tok.Open) sc.skipElement();
      }
    h.row(row);
  };

  const walk = (depth: number) => {
    for (;;) {
      const k = sc.next();
      if (k === Tok.End) {
        if (depth > 0) throw new Fallback();
        return;
      }
      if (k === Tok.Close) {
        if (depth === 0) throw new Fallback();
        return;
      }
      if (k !== Tok.Open && k !== Tok.Empty) continue;
      const n = sc.name;
      if (n === names.column) {
        h.column({ repeat: int(sc.attr(A.repeatCols)), style: sc.attr(A.style), visibility: sc.attr(A.visibility), defaultStyle: sc.attr(A.defaultStyle) });
        if (k === Tok.Open) sc.skipElement();
      } else if (n === names.row) readRow();
      else if (n === names.headerRows) {
        h.headerRows(true);
        if (k === Tok.Open) walk(depth + 1);
        h.headerRows(false);
      } else if (names.rowGroups.has(n) || names.colGroups.has(n)) {
        if (k === Tok.Open) walk(depth + 1);
      } else if (depth === 0) {
        // shapes, named ranges, conditional formats…
        const from = sc.start;
        if (k === Tok.Open) sc.skipElement();
        for (const el of fragment(xml.slice(from, sc.pos))) h.extra(el);
      } else if (k === Tok.Open) sc.skipElement();
    }
  };
  walk(0);
}

export function isFallback(e: unknown): boolean {
  return e instanceof Fallback;
}

/** An error that sends the reader back to the DOM path. */
export function fallback(): Error {
  return new Fallback();
}
