import type { JSONContent } from '@tiptap/core';
import { bytesToDataUrl, dataUrlToBytes } from '@/lib/utils';
import { normalizeColor } from '@/ui/color';
import { defaultSettings, PAPER_SIZES, resolveStyles, type DocModel, type DocSettings } from '../model';
import { textOf } from './shared';
import type { LoadedDoc } from './index';

/* ================================================================== IMPORT */

interface CharState {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  sup: boolean;
  sub: boolean;
  font: number;
  size: number; // half points
  color: number;
  highlight: number;
  uc: number;
  hidden: boolean;
}

interface ParaState {
  align: string | null;
  li: number;
  ri: number;
  fi: number;
  sb: number | null;
  sa: number | null;
  sl: number | null;
  style: number;
  intbl: boolean;
  list: boolean;
}

type Dest = 'normal' | 'skip' | 'fonttbl' | 'colortbl' | 'stylesheet' | 'pict' | 'fldinst' | 'fldrslt' | 'listtext' | 'info-title' | 'info-author' | 'header' | 'footer';

interface Group {
  ch: CharState;
  dest: Dest;
  fontDef?: { idx: number; name: string };
  styleDef?: { idx: number; name: string };
}

const defaultChar = (): CharState => ({ bold: false, italic: false, underline: false, strike: false, sup: false, sub: false, font: 0, size: 24, color: 0, highlight: 0, uc: 1, hidden: false });
const defaultPara = (): ParaState => ({ align: null, li: 0, ri: 0, fi: 0, sb: null, sa: null, sl: null, style: 0, intbl: false, list: false });

export function importRtf(src: string): LoadedDoc {
  if (!src.startsWith('{\\rtf')) throw new Error('This is not a valid RTF document.');
  const fonts = new Map<number, string>();
  const colors: Array<string | null> = [];
  const styles = new Map<number, string>();
  let codepage = 1252;
  const settings: DocSettings = defaultSettings();
  const paper = { w: 0, h: 0, l: 1440, r: 1440, t: 1440, b: 1440, landscape: false };

  const blocks: JSONContent[] = [];
  let inl: JSONContent[] = [];
  let para = defaultPara();
  let colorAcc = { r: 0, g: 0, b: 0, set: false };
  let pict = { hex: '', type: '', w: 0, h: 0, gw: 0, gh: 0 };
  let fieldInst = '';
  let fieldLink: string | null = null;
  let title = '';
  let author = '';
  let headerText = '';
  let footerText = '';
  // tables
  let row: JSONContent[] = [];
  let cellBlocks: JSONContent[] = [];
  let tableRows: JSONContent[] = [];
  let cellX: number[] = [];
  // lists
  let listItems: JSONContent[] = [];
  let listOrdered = false;
  let listText = '';

  const stack: Group[] = [{ ch: defaultChar(), dest: 'normal' }];
  const top = () => stack[stack.length - 1];

  let bytesPending: number[] = [];
  const flushBytes = () => {
    if (!bytesPending.length) return;
    let dec: TextDecoder;
    try {
      dec = new TextDecoder(`windows-${codepage}`);
    } catch {
      dec = new TextDecoder('windows-1252');
    }
    emitText(dec.decode(new Uint8Array(bytesPending)));
    bytesPending = [];
  };

  const marksFor = (c: CharState): JSONContent['marks'] => {
    const m: NonNullable<JSONContent['marks']> = [];
    if (c.bold) m.push({ type: 'bold' });
    if (c.italic) m.push({ type: 'italic' });
    if (c.underline && !fieldLink) m.push({ type: 'underline' });
    if (c.strike) m.push({ type: 'strike' });
    if (c.sup) m.push({ type: 'superscript' });
    if (c.sub) m.push({ type: 'subscript' });
    const ts: Record<string, string> = {};
    const f = fonts.get(c.font);
    if (f && c.font !== 0) ts.fontFamily = f;
    if (c.size !== 24 && c.size > 0) ts.fontSize = `${c.size / 2}pt`;
    const col = colors[c.color];
    if (c.color && col && !fieldLink) ts.color = col;
    if (Object.keys(ts).length) m.push({ type: 'textStyle', attrs: ts });
    const hl = colors[c.highlight];
    if (c.highlight && hl) m.push({ type: 'highlight', attrs: { color: hl } });
    if (fieldLink) m.push({ type: 'link', attrs: { href: fieldLink } });
    return m;
  };

  function emitText(t: string) {
    const g = top();
    if (!t) return;
    switch (g.dest) {
      case 'fonttbl':
        if (g.fontDef) g.fontDef.name += t;
        return;
      case 'stylesheet':
        if (g.styleDef) g.styleDef.name += t;
        return;
      case 'fldinst':
        fieldInst += t;
        return;
      case 'listtext':
        listText += t;
        return;
      case 'info-title':
        title += t;
        return;
      case 'info-author':
        author += t;
        return;
      case 'header':
        headerText += t;
        return;
      case 'footer':
        footerText += t;
        return;
      case 'pict':
        pict.hex += t.replace(/[^0-9a-f]/gi, '');
        return;
      case 'skip':
      case 'colortbl':
        return;
      default:
        if (g.ch.hidden) return;
        {
          const marks = marksFor(g.ch);
          const last = inl[inl.length - 1];
          if (last && last.type === 'text' && JSON.stringify(last.marks ?? []) === JSON.stringify(marks ?? [])) last.text += t;
          else inl.push(marks && marks.length ? { type: 'text', text: t, marks } : { type: 'text', text: t });
        }
    }
  }

  const paraNode = (): JSONContent => {
    const a: Record<string, unknown> = {};
    if (para.align && para.align !== 'left') a.textAlign = para.align;
    if (para.li && !para.list) a.indent = Math.round(para.li / 15);
    if (para.ri) a.indentRight = Math.round(para.ri / 15);
    if (para.fi && !para.list) a.firstLine = Math.round(para.fi / 15);
    if (para.sb !== null) a.spaceBefore = para.sb / 20;
    if (para.sa !== null) a.spaceAfter = para.sa / 20;
    if (para.sl) a.lineHeight = Math.round((Math.abs(para.sl) / 240) * 100) / 100;
    const sname = (styles.get(para.style) ?? '').toLowerCase().replace(/\s+/g, '');
    const h = /^heading([1-6])$/.exec(sname);
    const content = inl.length ? inl : undefined;
    inl = [];
    if (h) return { type: 'heading', attrs: { level: Number(h[1]), ...a }, content };
    if (sname === 'title') a.styleName = 'title';
    if (sname === 'subtitle') a.styleName = 'subtitle';
    return { type: 'paragraph', attrs: a, content };
  };

  const flushList = () => {
    if (!listItems.length) return;
    blocks.push({ type: listOrdered ? 'orderedList' : 'bulletList', content: listItems });
    listItems = [];
  };

  const endPara = () => {
    flushBytes();
    const node = paraNode();
    if (para.intbl) {
      cellBlocks.push(node);
      return;
    }
    if (para.list) {
      const ordered = /\d|^[a-z]\W|^[ivx]+\W/i.test(listText.trim());
      if (listItems.length && ordered !== listOrdered) flushList();
      listOrdered = ordered;
      listItems.push({ type: 'listItem', content: [node.type === 'heading' ? { ...node, type: 'paragraph' } : node] });
      listText = '';
      return;
    }
    flushList();
    blocks.push(node);
  };

  const endCell = () => {
    flushBytes();
    if (inl.length) cellBlocks.push(paraNode());
    const idx = row.length;
    const right = cellX[idx] ?? (idx + 1) * 1800;
    const left = idx ? cellX[idx - 1] ?? idx * 1800 : 0;
    row.push({ type: 'tableCell', attrs: { colwidth: [Math.max(30, Math.round((right - left) / 15))] }, content: cellBlocks.length ? cellBlocks : [{ type: 'paragraph' }] });
    cellBlocks = [];
  };

  const endRow = () => {
    if (row.length) tableRows.push({ type: 'tableRow', content: row });
    row = [];
  };

  const flushTable = () => {
    if (!tableRows.length) return;
    flushList();
    blocks.push({ type: 'table', content: tableRows });
    tableRows = [];
  };

  const finishPict = () => {
    const type = pict.type;
    if ((type === 'png' || type === 'jpeg') && pict.hex.length > 16) {
      const bytes = new Uint8Array(pict.hex.length / 2);
      for (let i = 0; i < bytes.length; i++) bytes[i] = parseInt(pict.hex.substr(i * 2, 2), 16);
      const w = pict.gw ? Math.round(pict.gw / 15) : pict.w || null;
      const h = pict.gh ? Math.round(pict.gh / 15) : pict.h || null;
      inl.push({ type: 'image', attrs: { src: bytesToDataUrl(bytes, type === 'png' ? 'image/png' : 'image/jpeg'), width: w, height: h, align: 'inline' } });
    }
    pict = { hex: '', type: '', w: 0, h: 0, gw: 0, gh: 0 };
  };

  const len = src.length;
  let i = 0;
  let ignorableNext = false;
  while (i < len) {
    const ch = src[i];
    if (ch === '{') {
      flushBytes();
      const g = top();
      stack.push({ ch: { ...g.ch }, dest: g.dest === 'fonttbl' || g.dest === 'stylesheet' || g.dest === 'colortbl' ? g.dest : g.dest === 'skip' ? 'skip' : g.dest === 'fldinst' ? 'fldinst' : g.dest });
      if (top().dest === 'fonttbl') top().fontDef = { idx: 0, name: '' };
      if (top().dest === 'stylesheet') top().styleDef = { idx: 0, name: '' };
      i++;
      continue;
    }
    if (ch === '}') {
      flushBytes();
      const g = stack.pop()!;
      if (g.fontDef && g.fontDef.name) fonts.set(g.fontDef.idx, g.fontDef.name.replace(/;\s*$/, '').trim());
      if (g.styleDef && g.styleDef.name) styles.set(g.styleDef.idx, g.styleDef.name.replace(/;\s*$/, '').trim());
      if (g.dest === 'pict' && top().dest !== 'pict') finishPict();
      if (g.dest === 'fldinst' && top().dest !== 'fldinst') {
        const m = /HYPERLINK\s+"([^"]+)"/i.exec(fieldInst);
        fieldLink = m ? m[1] : null;
        fieldInst = '';
      }
      if (g.dest === 'fldrslt' && top().dest !== 'fldrslt') fieldLink = null;
      if (!stack.length) break;
      i++;
      continue;
    }
    if (ch === '\\') {
      const next = src[i + 1];
      if (next === '\\' || next === '{' || next === '}') {
        flushBytes();
        emitText(next);
        i += 2;
        continue;
      }
      if (next === "'") {
        bytesPending.push(parseInt(src.substr(i + 2, 2), 16));
        i += 4;
        continue;
      }
      if (next === '*') {
        ignorableNext = true;
        i += 2;
        continue;
      }
      if (next === '~') {
        flushBytes();
        emitText('\u00a0');
        i += 2;
        continue;
      }
      if (next === '_') {
        flushBytes();
        emitText('\u2011');
        i += 2;
        continue;
      }
      if (next === '-') {
        i += 2;
        continue;
      }
      if (next === '\n' || next === '\r') {
        flushBytes();
        endPara();
        i += 2;
        continue;
      }
      // control word
      const m = /^\\([a-zA-Z]+)(-?\d+)? ?/.exec(src.slice(i, i + 40));
      if (!m) {
        i++;
        continue;
      }
      i += m[0].length;
      const word = m[1];
      const param = m[2] !== undefined ? Number(m[2]) : null;
      const g = top();
      if (word !== 'u') flushBytes();

      if (ignorableNext) {
        ignorableNext = false;
        if (!['fldinst', 'pn', 'listtext', 'shppict', 'pntext'].includes(word) && g.dest !== 'fonttbl') {
          g.dest = 'skip';
          continue;
        }
      }
      switch (word) {
        case 'ansicpg':
          if (param) codepage = param;
          break;
        case 'fonttbl':
          g.dest = 'fonttbl';
          break;
        case 'colortbl':
          g.dest = 'colortbl';
          break;
        case 'stylesheet':
          g.dest = 'stylesheet';
          break;
        case 'info':
          g.dest = 'skip';
          break;
        case 'title':
          g.dest = 'info-title';
          break;
        case 'author':
          g.dest = 'info-author';
          break;
        case 'header':
        case 'headerr':
          g.dest = 'header';
          break;
        case 'footer':
        case 'footerr':
          g.dest = 'footer';
          break;
        case 'headerl':
        case 'headerf':
        case 'footerl':
        case 'footerf':
        case 'footnote':
        case 'annotation':
        case 'object':
        case 'nonshppict':
        case 'xmlnstbl':
        case 'listtable':
        case 'listoverridetable':
        case 'rsidtbl':
        case 'generator':
        case 'themedata':
        case 'colorschememapping':
        case 'latentstyles':
        case 'datastore':
        case 'bkmkstart':
        case 'bkmkend':
          g.dest = 'skip';
          break;
        case 'pict':
          g.dest = 'pict';
          break;
        case 'pngblip':
          pict.type = 'png';
          break;
        case 'jpegblip':
          pict.type = 'jpeg';
          break;
        case 'picw':
          pict.w = param ?? 0;
          break;
        case 'pich':
          pict.h = param ?? 0;
          break;
        case 'picwgoal':
          pict.gw = param ?? 0;
          break;
        case 'pichgoal':
          pict.gh = param ?? 0;
          break;
        case 'fldinst':
          g.dest = 'fldinst';
          break;
        case 'fldrslt':
          g.dest = 'fldrslt';
          break;
        case 'listtext':
        case 'pntext':
          g.dest = 'listtext';
          para.list = true;
          break;
        case 'pn':
          g.dest = 'skip';
          para.list = true;
          break;
        case 'ls':
          para.list = true;
          break;
        case 'f':
          if (g.dest === 'fonttbl' && g.fontDef) g.fontDef.idx = param ?? 0;
          else g.ch.font = param ?? 0;
          break;
        case 's':
          if (g.dest === 'stylesheet' && g.styleDef) g.styleDef.idx = param ?? 0;
          else para.style = param ?? 0;
          break;
        case 'red':
          colorAcc.r = param ?? 0;
          colorAcc.set = true;
          break;
        case 'green':
          colorAcc.g = param ?? 0;
          colorAcc.set = true;
          break;
        case 'blue':
          colorAcc.b = param ?? 0;
          colorAcc.set = true;
          break;
        case 'fs':
          g.ch.size = param ?? 24;
          break;
        case 'b':
          g.ch.bold = param !== 0;
          break;
        case 'i':
          g.ch.italic = param !== 0;
          break;
        case 'ul':
          g.ch.underline = param !== 0;
          break;
        case 'ulnone':
          g.ch.underline = false;
          break;
        case 'strike':
        case 'striked':
          g.ch.strike = param !== 0;
          break;
        case 'super':
          g.ch.sup = true;
          break;
        case 'sub':
          g.ch.sub = true;
          break;
        case 'nosupersub':
          g.ch.sup = g.ch.sub = false;
          break;
        case 'v':
          g.ch.hidden = param !== 0;
          break;
        case 'cf':
          g.ch.color = param ?? 0;
          break;
        case 'highlight':
        case 'cb':
        case 'chcbpat':
          g.ch.highlight = param ?? 0;
          break;
        case 'plain':
          g.ch = { ...defaultChar(), uc: g.ch.uc };
          break;
        case 'pard':
          para = { ...defaultPara(), intbl: false };
          break;
        case 'ql':
          para.align = 'left';
          break;
        case 'qc':
          para.align = 'center';
          break;
        case 'qr':
          para.align = 'right';
          break;
        case 'qj':
          para.align = 'justify';
          break;
        case 'li':
          para.li = param ?? 0;
          break;
        case 'ri':
          para.ri = param ?? 0;
          break;
        case 'fi':
          para.fi = param ?? 0;
          break;
        case 'sb':
          para.sb = param ?? 0;
          break;
        case 'sa':
          para.sa = param ?? 0;
          break;
        case 'sl':
          para.sl = param && param > 0 ? param : null;
          break;
        case 'intbl':
          para.intbl = true;
          break;
        case 'par':
          if (g.dest === 'normal' || g.dest === 'fldrslt') endPara();
          else if (g.dest === 'header') headerText += '\n';
          else if (g.dest === 'footer') footerText += '\n';
          break;
        case 'line':
          if (g.dest === 'normal' || g.dest === 'fldrslt') inl.push({ type: 'hardBreak' });
          break;
        case 'tab':
          emitText('\t');
          break;
        case 'emdash':
          emitText('—');
          break;
        case 'endash':
          emitText('–');
          break;
        case 'bullet':
          emitText('•');
          break;
        case 'lquote':
          emitText('‘');
          break;
        case 'rquote':
          emitText('’');
          break;
        case 'ldblquote':
          emitText('“');
          break;
        case 'rdblquote':
          emitText('”');
          break;
        case 'page':
          if (g.dest === 'normal') {
            endPara();
            blocks.push({ type: 'pageBreak' });
          }
          break;
        case 'cell':
          endCell();
          break;
        case 'row':
          endRow();
          break;
        case 'trowd':
          cellX = [];
          break;
        case 'cellx':
          cellX.push(param ?? 0);
          break;
        case 'uc':
          g.ch.uc = param ?? 1;
          break;
        case 'u': {
          flushBytes();
          let code = param ?? 0;
          if (code < 0) code += 65536;
          emitText(String.fromCharCode(code));
          // skip fallback characters
          let skip = g.ch.uc;
          while (skip > 0 && i < len) {
            if (src[i] === '\\' && src[i + 1] === "'") i += 4;
            else if (src[i] === '{' || src[i] === '}') break;
            else i++;
            skip--;
          }
          break;
        }
        case 'paperw':
          paper.w = param ?? 0;
          break;
        case 'paperh':
          paper.h = param ?? 0;
          break;
        case 'margl':
          paper.l = param ?? 1440;
          break;
        case 'margr':
          paper.r = param ?? 1440;
          break;
        case 'margt':
          paper.t = param ?? 1440;
          break;
        case 'margb':
          paper.b = param ?? 1440;
          break;
        case 'landscape':
          paper.landscape = true;
          break;
        default:
          break;
      }
      continue;
    }
    if (ch === '\r' || ch === '\n') {
      i++;
      continue;
    }
    // plain text run
    let j = i;
    while (j < len && src[j] !== '\\' && src[j] !== '{' && src[j] !== '}' && src[j] !== '\r' && src[j] !== '\n') j++;
    const chunk = src.slice(i, j);
    if (top().dest === 'colortbl') {
      for (const c of chunk) {
        if (c === ';') {
          colors.push(colorAcc.set ? normalizeColor(`rgb(${colorAcc.r},${colorAcc.g},${colorAcc.b})`) : null);
          colorAcc = { r: 0, g: 0, b: 0, set: false };
        }
      }
    } else {
      flushBytes();
      emitText(chunk);
    }
    i = j;
    // tables end when a paragraph is not in a table
    if (!para.intbl && tableRows.length && top().dest === 'normal' && chunk.trim()) {
      flushTable();
    }
  }
  flushBytes();
  if (inl.length) endPara();
  flushList();
  flushTable();

  if (paper.w && paper.h) {
    const pw = Math.min(paper.w, paper.h) / 1440;
    const ph = Math.max(paper.w, paper.h) / 1440;
    const known = PAPER_SIZES.find((p) => Math.abs(p.width - pw) < 0.08 && Math.abs(p.height - ph) < 0.08);
    settings.page = {
      size: known?.id ?? 'custom',
      width: known?.width ?? pw,
      height: known?.height ?? ph,
      orientation: paper.landscape || paper.w > paper.h ? 'landscape' : 'portrait',
      margins: { top: paper.t / 1440, right: paper.r / 1440, bottom: paper.b / 1440, left: paper.l / 1440 },
    };
  }
  settings.title = title.trim();
  settings.author = author.trim();
  if (headerText.trim()) settings.header = { left: '', center: headerText.trim().replace(/\s+/g, ' '), right: '' };
  if (footerText.trim()) settings.footer = { left: '', center: footerText.trim().replace(/\s+/g, ' '), right: '' };
  settings.styles = { normal: { font: fonts.get(0) ?? 'Calibri', size: 12, after: 0, lineHeight: 1.15 } };

  const clean = (nodes: JSONContent[]): JSONContent[] =>
    nodes.map((n) => (n.content ? { ...n, content: clean(n.content) } : n)).filter((n) => !(n.type === 'text' && !n.text));
  const content = clean(blocks);
  return { settings, content: { type: 'doc', content: content.length ? content : [{ type: 'paragraph' }] }, comments: [] };
}

/* ================================================================== EXPORT */

function rtfText(s: string): string {
  let out = '';
  for (const ch of s) {
    const code = ch.codePointAt(0)!;
    if (ch === '\\' || ch === '{' || ch === '}') out += `\\${ch}`;
    else if (ch === '\t') out += '\\tab ';
    else if (ch === '\n') out += '\\line ';
    else if (code < 128) out += ch;
    else if (code <= 0xffff) out += `\\u${code > 32767 ? code - 65536 : code}?`;
    else {
      // surrogate pair
      const hi = Math.floor((code - 0x10000) / 0x400) + 0xd800;
      const lo = ((code - 0x10000) % 0x400) + 0xdc00;
      out += `\\u${hi - 65536}?\\u${lo - 65536}?`;
    }
  }
  return out;
}

export function exportRtf(model: DocModel): string {
  const s = model.settings;
  const st = resolveStyles(s);
  const fonts: string[] = [st.normal.font ?? 'Calibri'];
  const colors: string[] = [];
  const fontIdx = (f: string) => {
    const name = f.split(',')[0].replace(/["']/g, '').trim();
    let idx = fonts.indexOf(name);
    if (idx < 0) {
      fonts.push(name);
      idx = fonts.length - 1;
    }
    return idx;
  };
  const colorIdx = (c: string) => {
    const hex = normalizeColor(c);
    if (!hex) return 0;
    let idx = colors.indexOf(hex);
    if (idx < 0) {
      colors.push(hex);
      idx = colors.length - 1;
    }
    return idx + 1;
  };
  const halfPts = (v: string) => {
    const n = parseFloat(v);
    return Math.round((/px$/.test(v) ? (n * 72) / 96 : n) * 2);
  };

  const runs = (nodes: JSONContent[] | undefined): string =>
    (nodes ?? [])
      .map((n) => {
        if (n.type === 'hardBreak') return '\\line ';
        if (n.type === 'image') {
          const d = dataUrlToBytes(String(n.attrs?.src ?? ''));
          if (!d || !/png|jpeg/.test(d.mime)) return '';
          let hex = '';
          for (const b of d.bytes) hex += b.toString(16).padStart(2, '0');
          const w = Math.round((Number(n.attrs?.width) || 300) * 15);
          const h = Math.round((Number(n.attrs?.height) || 200) * 15);
          return `{\\pict\\${d.mime.includes('png') ? 'pngblip' : 'jpegblip'}\\picwgoal${w}\\pichgoal${h}\n${hex.replace(/(.{128})/g, '$1\n')}}`;
        }
        if (n.type === 'mathInline') return `{\\i ${rtfText(String(n.attrs?.latex ?? ''))}}`;
        if (n.type !== 'text') return '';
        let fmt = '';
        let link: string | null = null;
        for (const m of n.marks ?? []) {
          if (m.type === 'bold') fmt += '\\b';
          else if (m.type === 'italic') fmt += '\\i';
          else if (m.type === 'underline') fmt += '\\ul';
          else if (m.type === 'strike') fmt += '\\strike';
          else if (m.type === 'superscript') fmt += '\\super';
          else if (m.type === 'subscript') fmt += '\\sub';
          else if (m.type === 'code') fmt += `\\f${fontIdx('Courier New')}`;
          else if (m.type === 'highlight') fmt += `\\highlight${colorIdx(String(m.attrs?.color ?? '#ffff00'))}`;
          else if (m.type === 'link') link = String(m.attrs?.href ?? '');
          else if (m.type === 'textStyle') {
            const a = m.attrs ?? {};
            if (a.fontFamily) fmt += `\\f${fontIdx(String(a.fontFamily))}`;
            if (a.fontSize) fmt += `\\fs${halfPts(String(a.fontSize))}`;
            if (a.color) fmt += `\\cf${colorIdx(String(a.color))}`;
          }
        }
        const body = `{${fmt}${fmt ? ' ' : ''}${rtfText(n.text ?? '')}}`;
        return link ? `{\\field{\\*\\fldinst{HYPERLINK "${link.replace(/"/g, '')}"}}{\\fldrslt{\\ul\\cf${colorIdx('#0563c1')} ${rtfText(n.text ?? '')}}}}` : body;
      })
      .join('');

  const paraProps = (n: JSONContent, extra = ''): string => {
    const a = n.attrs ?? {};
    let p = '\\pard\\plain';
    const sn = a.styleName as string | undefined;
    const def = n.type === 'heading' ? st[`h${Math.min(6, Number(a.level) || 1)}` as 'h1'] : sn === 'title' ? st.title : sn === 'subtitle' ? st.subtitle : sn === 'quote' ? st.quote : st.normal;
    if (n.type === 'heading') p += `\\s${Math.min(6, Number(a.level) || 1)}\\outlinelevel${(Number(a.level) || 1) - 1}`;
    const al = (a.textAlign as string) ?? def.align;
    if (al === 'center') p += '\\qc';
    else if (al === 'right') p += '\\qr';
    else if (al === 'justify') p += '\\qj';
    if (a.indent) p += `\\li${Math.round(Number(a.indent) * 15)}`;
    if (a.indentRight) p += `\\ri${Math.round(Number(a.indentRight) * 15)}`;
    if (a.firstLine) p += `\\fi${Math.round(Number(a.firstLine) * 15)}`;
    const before = (a.spaceBefore as number | null) ?? def.before ?? 0;
    const after = (a.spaceAfter as number | null) ?? def.after ?? st.normal.after ?? 8;
    p += `\\sb${Math.round(before * 20)}\\sa${Math.round(after * 20)}`;
    const lh = (a.lineHeight as number | null) ?? def.lineHeight ?? st.normal.lineHeight;
    if (lh) p += `\\sl${Math.round(lh * 240)}\\slmult1`;
    // character defaults of the style
    p += `\\f${fontIdx(def.font ?? st.normal.font ?? 'Calibri')}\\fs${Math.round((def.size ?? 11) * 2)}`;
    if (def.bold) p += '\\b';
    if (def.italic) p += '\\i';
    if (def.color && def.color !== '#000000') p += `\\cf${colorIdx(def.color)}`;
    return `${p}${extra} `;
  };

  const blocks = (nodes: JSONContent[] | undefined, listDepth = 0): string =>
    (nodes ?? [])
      .map((n) => {
        switch (n.type) {
          case 'paragraph':
          case 'heading':
            return `${paraProps(n)}${runs(n.content)}\\par\n`;
          case 'bulletList':
          case 'orderedList':
          case 'taskList': {
            let k = 0;
            return (n.content ?? [])
              .map((item) => {
                k++;
                const [first, ...rest] = item.content ?? [];
                const label = n.type === 'orderedList' ? `${(Number(n.attrs?.start) || 1) + k - 1}.` : n.type === 'taskList' ? (item.attrs?.checked ? '\\u9746?' : '\\u9744?') : '\\u8226?';
                const indent = 360 + listDepth * 360;
                const head = first ? `${paraProps(first, `\\fi-360\\li${indent + 360}`)}{\\pntext ${label}\\tab}${runs(first.content)}\\par\n` : '';
                return head + blocks(rest, listDepth + 1);
              })
              .join('');
          }
          case 'blockquote':
            return blocks((n.content ?? []).map((c) => ({ ...c, attrs: { ...(c.attrs ?? {}), indent: 48 } })), listDepth);
          case 'codeBlock':
            return `\\pard\\plain\\f${fontIdx('Courier New')}\\fs20\\sb120\\sa120 ${rtfText(textOf(n))}\\par\n`;
          case 'horizontalRule':
            return '\\pard\\plain\\brdrb\\brdrs\\brdrw10\\brsp20 \\par\n';
          case 'pageBreak':
            return '\\page\n';
          case 'mathBlock':
            return `\\pard\\plain\\qc {\\i ${rtfText(String(n.attrs?.latex ?? ''))}}\\par\n`;
          case 'table': {
            return (n.content ?? [])
              .map((row) => {
                let x = 0;
                const defs = (row.content ?? [])
                  .map((c) => {
                    const w = ((c.attrs?.colwidth as number[] | null) ?? []).reduce((a, b) => a + b, 0) || 150;
                    x += Math.round(w * 15);
                    const bg = c.attrs?.backgroundColor ? `\\clcbpat${colorIdx(String(c.attrs.backgroundColor))}` : '';
                    return `\\clbrdrt\\brdrs\\brdrw10\\clbrdrl\\brdrs\\brdrw10\\clbrdrb\\brdrs\\brdrw10\\clbrdrr\\brdrs\\brdrw10${bg}\\cellx${x}`;
                  })
                  .join('');
                const cells = (row.content ?? [])
                  .map((c) => {
                    const paras = c.content ?? [];
                    return `${paras.map((p, pi) => `${paraProps(p, '\\intbl')}${c.type === 'tableHeader' ? '{\\b ' : '{'}${runs(p.content)}}${pi < paras.length - 1 ? '\\par' : ''}`).join('')}\\cell\n`;
                  })
                  .join('');
                return `\\trowd\\trgaph108\\trleft0${defs}\n${cells}\\row\n`;
              })
              .join('') + '\\pard\\plain \n';
          }
          case 'tableOfContents':
            return `${paraProps({ type: 'heading', attrs: { level: 1 } })}${rtfText(String(n.attrs?.title ?? 'Contents'))}\\par\n`;
          default:
            return n.content ? blocks(n.content, listDepth) : '';
        }
      })
      .join('');

  const body = blocks((model.content as JSONContent).content);
  const page = s.page;
  const pw = Math.round((page.orientation === 'landscape' ? page.height : page.width) * 1440);
  const ph = Math.round((page.orientation === 'landscape' ? page.width : page.height) * 1440);
  const hfText = (h: DocSettings['header']) =>
    [h.left, h.center, h.right]
      .filter(Boolean)
      .join('    ')
      .split(/(\{PAGE\}|\{PAGES\})/i)
      .map((part) => (/^\{PAGE\}$/i.test(part) ? '{\\field{\\*\\fldinst PAGE}{\\fldrslt 1}}' : /^\{PAGES\}$/i.test(part) ? '{\\field{\\*\\fldinst NUMPAGES}{\\fldrslt 1}}' : rtfText(part)))
      .join('');
  const header = s.header.left || s.header.center || s.header.right ? `{\\header\\pard\\plain\\qc\\fs18 ${hfText(s.header)}\\par}` : '';
  const footer = s.footer.left || s.footer.center || s.footer.right ? `{\\footer\\pard\\plain\\qc\\fs18 ${hfText(s.footer)}\\par}` : '';
  const hex2rgb = (h: string) => {
    const n = parseInt(h.slice(1), 16);
    return `\\red${(n >> 16) & 255}\\green${(n >> 8) & 255}\\blue${n & 255};`;
  };
  const stylesheet = `{\\stylesheet{\\s0 Normal;}${[1, 2, 3, 4, 5, 6].map((l) => `{\\s${l}\\sbasedon0\\snext0\\b heading ${l};}`).join('')}}`;
  return `{\\rtf1\\ansi\\ansicpg1252\\deff0\\uc1
{\\fonttbl${fonts.map((f, i) => `{\\f${i}\\fnil\\fcharset0 ${f};}`).join('')}}
{\\colortbl;${colors.map(hex2rgb).join('')}}
${stylesheet}
{\\*\\generator Affice;}
{\\info{\\title ${rtfText(s.title)}}{\\author ${rtfText(s.author)}}}
\\paperw${pw}\\paperh${ph}\\margl${Math.round(page.margins.left * 1440)}\\margr${Math.round(page.margins.right * 1440)}\\margt${Math.round(page.margins.top * 1440)}\\margb${Math.round(page.margins.bottom * 1440)}${page.orientation === 'landscape' ? '\\landscape' : ''}
\\viewkind4\\deftab720
${header}${footer}
${body}}`;
}
