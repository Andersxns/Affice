import { mapEls, removeEls, type SlidesDoc } from './doc';
import {
  cloneEls,
  findEl,
  newId,
  rotatedBounds,
  unionBounds,
  walkEls,
  type El,
  type Fill,
  type GroupEl,
  type Line,
  type Para,
  type Presentation,
  type Run,
  type Shadow,
  type ShapeEl,
  type Slide,
  type TableEl,
  type TextBody,
  type Transition,
} from './model';
import { newSlide, nextLayoutFor } from './themes';

type Frame = { x: number; y: number; w: number; h: number };

/* ================================================================ slides */

export function addSlide(doc: SlidesDoc, layoutId?: string, at = doc.sel.slide + 1): void {
  const pres = doc.pres;
  const s = newSlide(pres, layoutId ?? nextLayoutFor(pres, doc.slide));
  const slides = [...pres.slides];
  slides.splice(at, 0, s);
  doc.commit('New slide', { ...pres, slides }, { slide: at, els: [], slides: [at] });
}

export function duplicateSlides(doc: SlidesDoc, indices = doc.sel.slides): void {
  const pres = doc.pres;
  const sorted = [...indices].sort((a, b) => a - b);
  const copies = sorted.map((i) => {
    const s = pres.slides[i];
    return { ...structuredClone(s), id: newId('s'), elements: cloneEls(s.elements), anims: undefined as Slide['anims'] } as Slide;
  });
  // keep animations pointing at the copied elements
  sorted.forEach((i, k) => {
    const src = pres.slides[i];
    if (!src.anims?.length) return;
    const map = new Map<string, string>();
    src.elements.forEach((e, j) => map.set(e.id, copies[k].elements[j].id));
    copies[k].anims = src.anims.map((a) => ({ ...a, id: newId('a'), el: map.get(a.el) ?? a.el }));
  });
  const at = sorted[sorted.length - 1] + 1;
  const slides = [...pres.slides];
  slides.splice(at, 0, ...copies);
  doc.commit('Duplicate slide', { ...pres, slides }, { slide: at, els: [], slides: copies.map((_, k) => at + k) });
}

export function deleteSlides(doc: SlidesDoc, indices = doc.sel.slides): void {
  const pres = doc.pres;
  const del = new Set(indices);
  if (del.size >= pres.slides.length) {
    // keep at least one slide
    const s = newSlide(pres, pres.layouts.find((l) => l.type === 'title')?.id ?? pres.layouts[0].id);
    doc.commit('Delete slide', { ...pres, slides: [s] }, { slide: 0, els: [], slides: [0] });
    return;
  }
  const slides = pres.slides.filter((_, i) => !del.has(i));
  const first = Math.min(...indices);
  const cur = Math.max(0, Math.min(slides.length - 1, first));
  doc.commit('Delete slide', { ...pres, slides }, { slide: cur, els: [], slides: [cur] });
}

export function moveSlides(doc: SlidesDoc, indices: number[], to: number): void {
  const pres = doc.pres;
  const set = new Set(indices);
  const moving = pres.slides.filter((_, i) => set.has(i));
  const rest = pres.slides.filter((_, i) => !set.has(i));
  const before = indices.filter((i) => i < to).length;
  const at = Math.max(0, Math.min(rest.length, to - before));
  rest.splice(at, 0, ...moving);
  if (rest.every((s, i) => s === pres.slides[i])) return;
  doc.commit('Move slide', { ...pres, slides: rest }, { slide: at, els: [], slides: moving.map((_, k) => at + k) });
}

const phRank = (e: El) => (e.ph === 'title' || e.ph === 'ctrTitle' ? 'title' : e.ph === 'pic' ? 'pic' : e.ph ? 'body' : '');

/** Changes slide layouts, moving placeholders into the new layout's frames (like PowerPoint). */
export function setLayout(doc: SlidesDoc, layoutId: string, indices = doc.sel.slides): void {
  const pres = doc.pres;
  const layout = pres.layouts.find((l) => l.id === layoutId);
  if (!layout) return;
  const slides = pres.slides.map((s, i) => {
    if (!indices.includes(i)) return s;
    const used = new Set<string>();
    const elements = s.elements.map((e) => {
      if (!e.ph) return e;
      const kind = phRank(e);
      const target = layout.placeholders.find((p) => !used.has(p.id) && phRank(p) === kind && (p.phIdx ?? 0) === (e.phIdx ?? 0)) ?? layout.placeholders.find((p) => !used.has(p.id) && phRank(p) === kind);
      if (!target) return e;
      used.add(target.id);
      const next = { ...e, x: target.x, y: target.y, w: target.w, h: target.h, ph: target.ph, phIdx: target.phIdx } as El;
      if (next.type === 'shape' && target.type === 'shape' && target.text && next.text) next.text = { ...next.text, anchor: target.text.anchor, defaults: target.text.defaults, autofit: target.text.autofit };
      return next;
    });
    for (const p of layout.placeholders) if (!used.has(p.id)) elements.push({ ...structuredClone(p), id: newId() });
    return { ...s, layout: layoutId, elements };
  });
  doc.commit('Change layout', { ...pres, slides });
}

/** Resets placeholder positions and removes empty extra placeholders. */
export function resetSlides(doc: SlidesDoc, indices = doc.sel.slides): void {
  for (const i of indices) {
    const s = doc.pres.slides[i];
    if (s) setLayout(doc, s.layout, [i]);
  }
}

export function toggleHidden(doc: SlidesDoc, indices = doc.sel.slides): void {
  const pres = doc.pres;
  const hide = !indices.every((i) => pres.slides[i]?.hidden);
  doc.commit(hide ? 'Hide slide' : 'Show slide', { ...pres, slides: pres.slides.map((s, i) => (indices.includes(i) ? { ...s, hidden: hide || undefined } : s)) });
}

export function setBackground(doc: SlidesDoc, fill: Fill | undefined, all: boolean): void {
  const pres = doc.pres;
  if (all) {
    doc.commit('Background', { ...pres, master: { ...pres.master, background: fill ?? pres.master.background }, slides: pres.slides.map((s) => (s.background ? { ...s, background: undefined } : s)), layouts: pres.layouts.map((l) => (l.background ? { ...l, background: undefined } : l)) });
    return;
  }
  doc.commit('Background', { ...pres, slides: pres.slides.map((s, i) => (doc.sel.slides.includes(i) ? { ...s, background: fill } : s)) });
}

export function setHideDecor(doc: SlidesDoc, hide: boolean): void {
  const pres = doc.pres;
  doc.commit('Background graphics', { ...pres, slides: pres.slides.map((s, i) => (doc.sel.slides.includes(i) ? { ...s, hideDecor: hide || undefined } : s)) });
}

export function setTransition(doc: SlidesDoc, t: Transition | undefined, all = false): void {
  const pres = doc.pres;
  doc.commit('Transition', { ...pres, slides: pres.slides.map((s, i) => (all || doc.sel.slides.includes(i) ? { ...s, transition: t } : s)) });
}

export function patchTransition(doc: SlidesDoc, patch: Partial<Transition>): void {
  const pres = doc.pres;
  doc.commit(
    'Transition',
    { ...pres, slides: pres.slides.map((s, i) => (doc.sel.slides.includes(i) ? { ...s, transition: { type: 'none', ...(s.transition ?? {}), ...patch } } : s)) },
    undefined,
    true,
  );
}

export function setNotes(doc: SlidesDoc, notes: string): void {
  doc.updateSlide('Notes', (s) => (s.notes === notes ? s : { ...s, notes }), doc.sel.slide, true);
}

/* ============================================================== elements */

export function insertEls(doc: SlidesDoc, els: El[], label = 'Insert'): void {
  if (!els.length) return;
  doc.editing = null;
  doc.updateSlide(label, (s) => ({ ...s, elements: [...s.elements, ...els] }));
  doc.setSel({ els: els.map((e) => e.id) });
}

export function deleteEls(doc: SlidesDoc, ids = doc.sel.els): void {
  if (!ids.length) return;
  const set = new Set(ids);
  doc.editing = null;
  doc.updateSlide('Delete', (s) => ({ ...s, elements: removeEls(s.elements, set), anims: s.anims?.filter((a) => !set.has(a.el)) }));
  doc.setSel({ els: [] });
}

export function duplicateEls(doc: SlidesDoc, offset = 16): void {
  const src = doc.selected;
  if (!src.length) return;
  const copies = cloneEls(src).map((e) => moveEl(e, offset, offset));
  insertEls(doc, copies, 'Duplicate');
}

export function moveEl(e: El, dx: number, dy: number): El {
  if (e.type === 'group') return { ...e, x: e.x + dx, y: e.y + dy, children: e.children.map((c) => moveEl(c, dx, dy)) } as GroupEl;
  return { ...e, x: e.x + dx, y: e.y + dy };
}

/** Sets an element's frame; groups scale their children. */
export function setFrame(e: El, f: Frame): El {
  if (e.type === 'group') {
    const sx = e.w ? f.w / e.w : 1;
    const sy = e.h ? f.h / e.h : 1;
    const children = e.children.map((c) => setFrame(c, { x: f.x + (c.x - e.x) * sx, y: f.y + (c.y - e.y) * sy, w: c.w * sx, h: c.h * sy }));
    return { ...e, ...f, children } as GroupEl;
  }
  if (e.type === 'table') {
    const sx = e.w ? f.w / e.w : 1;
    const sy = e.h ? f.h / e.h : 1;
    return { ...e, ...f, cols: e.cols.map((w) => w * sx), rows: e.rows.map((r) => ({ ...r, h: r.h * sy })) } as TableEl;
  }
  return { ...e, ...f };
}

export function nudge(doc: SlidesDoc, dx: number, dy: number): void {
  const ids = doc.sel.els;
  if (!ids.length) return;
  doc.updateEls('Move', ids, (e) => (e.locked ? e : moveEl(e, dx, dy)), true);
}

export function zOrder(doc: SlidesDoc, how: 'front' | 'back' | 'forward' | 'backward'): void {
  const ids = new Set(doc.sel.els);
  if (!ids.size) return;
  doc.updateSlide('Arrange', (s) => {
    const els = [...s.elements];
    const sel = els.filter((e) => ids.has(e.id));
    const rest = els.filter((e) => !ids.has(e.id));
    let out: El[];
    if (how === 'front') out = [...rest, ...sel];
    else if (how === 'back') out = [...sel, ...rest];
    else {
      out = [...els];
      const order = how === 'forward' ? [...out.keys()].reverse() : [...out.keys()];
      for (const i of order) {
        if (!ids.has(out[i].id)) continue;
        const j = how === 'forward' ? i + 1 : i - 1;
        if (j < 0 || j >= out.length || ids.has(out[j].id)) continue;
        [out[i], out[j]] = [out[j], out[i]];
      }
    }
    return { ...s, elements: out };
  });
}

export type AlignHow = 'left' | 'center' | 'right' | 'top' | 'middle' | 'bottom';

export function align(doc: SlidesDoc, how: AlignHow, toSlide: boolean): void {
  const els = doc.selected;
  if (!els.length) return;
  const bounds = els.map((e) => rotatedBounds(e));
  const ref = toSlide || els.length === 1 ? { x: 0, y: 0, w: doc.pres.size.w, h: doc.pres.size.h } : unionBounds(bounds)!;
  const moves = new Map<string, [number, number]>();
  els.forEach((e, i) => {
    const b = bounds[i];
    let dx = 0;
    let dy = 0;
    if (how === 'left') dx = ref.x - b.x;
    if (how === 'center') dx = ref.x + ref.w / 2 - (b.x + b.w / 2);
    if (how === 'right') dx = ref.x + ref.w - (b.x + b.w);
    if (how === 'top') dy = ref.y - b.y;
    if (how === 'middle') dy = ref.y + ref.h / 2 - (b.y + b.h / 2);
    if (how === 'bottom') dy = ref.y + ref.h - (b.y + b.h);
    moves.set(e.id, [dx, dy]);
  });
  doc.updateEls('Align', [...moves.keys()], (e) => {
    const [dx, dy] = moves.get(e.id)!;
    return moveEl(e, Math.round(dx), Math.round(dy));
  });
}

export function distribute(doc: SlidesDoc, axis: 'h' | 'v'): void {
  const els = doc.selected;
  if (els.length < 3) return;
  const items = els.map((e) => ({ e, b: rotatedBounds(e) })).sort((a, b) => (axis === 'h' ? a.b.x - b.b.x : a.b.y - b.b.y));
  const first = items[0].b;
  const last = items[items.length - 1].b;
  const total = items.reduce((acc, it) => acc + (axis === 'h' ? it.b.w : it.b.h), 0);
  const span = axis === 'h' ? last.x + last.w - first.x : last.y + last.h - first.y;
  const gap = (span - total) / (items.length - 1);
  let pos = axis === 'h' ? first.x : first.y;
  const moves = new Map<string, [number, number]>();
  for (const it of items) {
    const d = pos - (axis === 'h' ? it.b.x : it.b.y);
    moves.set(it.e.id, axis === 'h' ? [d, 0] : [0, d]);
    pos += (axis === 'h' ? it.b.w : it.b.h) + gap;
  }
  doc.updateEls('Distribute', [...moves.keys()], (e) => {
    const [dx, dy] = moves.get(e.id)!;
    return moveEl(e, Math.round(dx), Math.round(dy));
  });
}

export function group(doc: SlidesDoc): void {
  const ids = new Set(doc.sel.els);
  if (ids.size < 2) return;
  const s = doc.slide;
  const members = s.elements.filter((e) => ids.has(e.id));
  if (members.length < 2) return;
  const b = unionBounds(members.map(rotatedBounds))!;
  const g: GroupEl = { id: newId(), type: 'group', name: 'Group', x: b.x, y: b.y, w: b.w, h: b.h, children: members };
  const firstIdx = s.elements.findIndex((e) => ids.has(e.id));
  doc.updateSlide('Group', (sl) => {
    const rest = sl.elements.filter((e) => !ids.has(e.id));
    rest.splice(Math.min(firstIdx, rest.length), 0, g);
    return { ...sl, elements: rest };
  });
  doc.setSel({ els: [g.id] });
}

export function ungroup(doc: SlidesDoc): void {
  const groups = doc.selected.filter((e): e is GroupEl => e.type === 'group');
  if (!groups.length) return;
  const ids = new Set(groups.map((g) => g.id));
  const newSel: string[] = [];
  doc.updateSlide('Ungroup', (s) => {
    const out: El[] = [];
    for (const e of s.elements) {
      if (!ids.has(e.id) || e.type !== 'group') {
        out.push(e);
        continue;
      }
      const rot = e.rot ?? 0;
      const cx = e.x + e.w / 2;
      const cy = e.y + e.h / 2;
      for (const c of e.children) {
        let child = c;
        if (rot) {
          const a = (rot * Math.PI) / 180;
          const ccx = c.x + c.w / 2 - cx;
          const ccy = c.y + c.h / 2 - cy;
          const nx = cx + ccx * Math.cos(a) - ccy * Math.sin(a);
          const ny = cy + ccx * Math.sin(a) + ccy * Math.cos(a);
          child = moveEl(c, nx - (c.x + c.w / 2), ny - (c.y + c.h / 2));
          child = { ...child, rot: ((c.rot ?? 0) + rot) % 360 };
        }
        if (e.opacity !== undefined && e.opacity < 1) child = { ...child, opacity: (child.opacity ?? 1) * e.opacity };
        out.push(child);
        newSel.push(child.id);
      }
    }
    return { ...s, elements: out };
  });
  doc.setSel({ els: newSel });
}

export function rotateEls(doc: SlidesDoc, delta: number): void {
  doc.updateEls('Rotate', doc.sel.els, (e) => ({ ...e, rot: ((((e.rot ?? 0) + delta) % 360) + 360) % 360 }));
}

export function flipEls(doc: SlidesDoc, axis: 'h' | 'v'): void {
  doc.updateEls('Flip', doc.sel.els, (e) => (axis === 'h' ? { ...e, flipH: !e.flipH || undefined } : { ...e, flipV: !e.flipV || undefined }));
}

export function patchEls(doc: SlidesDoc, label: string, patch: Partial<El> | ((e: El) => Partial<El> | null), ids = doc.sel.els, coalesce = false): void {
  doc.updateEls(
    label,
    ids,
    (e) => {
      const p = typeof patch === 'function' ? patch(e) : patch;
      if (!p) return e;
      return { ...e, ...p } as El;
    },
    coalesce,
  );
}

export function setShapeFill(doc: SlidesDoc, fill: Fill): void {
  patchEls(doc, 'Shape fill', (e) => (e.type === 'shape' ? { fill } : e.type === 'group' ? null : null));
  // groups: apply to shape children
  const groups = doc.selected.filter((e) => e.type === 'group').map((g) => g.id);
  if (groups.length) doc.updateEls('Shape fill', groups, (g) => ({ ...g, children: (g as GroupEl).children.map((c) => (c.type === 'shape' ? { ...c, fill } : c)) }) as GroupEl);
}

export function setShapeLine(doc: SlidesDoc, patch: Partial<Line> | null): void {
  patchEls(doc, 'Shape outline', (e) => {
    if (e.type !== 'shape' && e.type !== 'image') return null;
    if (patch === null) return { line: null };
    const base: Line = e.line ?? { color: '@tx1', width: 1.33 };
    return { line: { ...base, ...patch } };
  });
}

export function setShadow(doc: SlidesDoc, shadow: Shadow | undefined): void {
  patchEls(doc, 'Shadow', (e) => (e.type === 'shape' || e.type === 'image' ? { shadow } : null));
}

export function changeGeom(doc: SlidesDoc, geom: string): void {
  patchEls(doc, 'Change shape', (e) => (e.type === 'shape' ? { geom, adj: undefined, custom: undefined } : e.type === 'image' ? { geom } : null));
}

/* ================================================================== text */

function mapRuns(body: TextBody, fn: (r: Omit<Run, 'text'>) => Omit<Run, 'text'>): TextBody {
  return {
    ...body,
    paras: body.paras.map((p) => ({ ...p, runs: p.runs.map((r) => ({ ...fn(r), text: r.text }) as Run), endRun: fn(p.endRun ?? {}) })),
  };
}

function mapBodies(e: El, fn: (b: TextBody) => TextBody): El {
  if (e.type === 'shape') return e.text ? { ...e, text: fn(e.text) } : e;
  if (e.type === 'table') return { ...e, rows: e.rows.map((r) => ({ ...r, cells: r.cells.map((c) => ({ ...c, text: fn(c.text) })) })) };
  if (e.type === 'group') return { ...e, children: e.children.map((c) => mapBodies(c, fn)) } as GroupEl;
  return e;
}

/** Applies run formatting to every run in the selected elements (when not editing text). */
export function applyRunStyle(doc: SlidesDoc, patch: Partial<Omit<Run, 'text'>>, label = 'Format text'): void {
  const clean = (r: Omit<Run, 'text'>) => {
    const out = { ...r, ...patch } as Record<string, unknown>;
    for (const k of Object.keys(out)) if (out[k] === undefined || out[k] === false) delete out[k];
    return out as Omit<Run, 'text'>;
  };
  doc.updateEls(label, doc.sel.els, (e) => mapBodies(e, (b) => mapRuns(b, clean)));
}

export function applyParaStyle(doc: SlidesDoc, patch: Partial<Omit<Para, 'runs'>> | ((p: Para) => Partial<Omit<Para, 'runs'>>), label = 'Paragraph'): void {
  doc.updateEls(label, doc.sel.els, (e) =>
    mapBodies(e, (b) => ({
      ...b,
      paras: b.paras.map((p) => {
        const d = typeof patch === 'function' ? patch(p) : patch;
        const out = { ...p, ...d } as Para & Record<string, unknown>;
        for (const k of Object.keys(out)) if (out[k] === undefined) delete out[k];
        return out;
      }),
    })),
  );
}

export function applyBodyStyle(doc: SlidesDoc, patch: Partial<Omit<TextBody, 'paras'>>, label = 'Text options'): void {
  doc.updateEls(label, doc.sel.els, (e) => mapBodies(e, (b) => ({ ...b, ...patch })));
}

/** First run's formatting in the selection (for ribbon state when not editing). */
export function selectionRun(doc: SlidesDoc): Omit<Run, 'text'> & { align?: Para['align']; para?: Para } {
  for (const e of doc.selected) {
    let found: (Omit<Run, 'text'> & { align?: Para['align']; para?: Para }) | null = null;
    walkEls([e], (x) => {
      if (found) return;
      const body = x.type === 'shape' ? x.text : x.type === 'table' ? x.rows[0]?.cells[0]?.text : undefined;
      const p = body?.paras[0];
      if (p) found = { ...(body?.defaults ?? {}), ...(p.runs[0] ?? p.endRun ?? {}), align: p.align, para: p };
    });
    if (found) return found;
  }
  return {};
}

/* ============================================================ clipboard */

let clipboard: { els: El[]; text: string; from: string } | null = null;

export function copyEls(doc: SlidesDoc, cut: boolean): string {
  const els = doc.selected;
  if (!els.length) return '';
  const text = els
    .map((e) => {
      const parts: string[] = [];
      walkEls([e], (x) => {
        if (x.type === 'shape' && x.text) parts.push(x.text.paras.map((p) => p.runs.map((r) => r.text).join('')).join('\n'));
        if (x.type === 'table') parts.push(x.rows.map((r) => r.cells.map((c) => c.text.paras.map((p) => p.runs.map((q) => q.text).join('')).join(' ')).join('\t')).join('\n'));
      });
      return parts.join('\n');
    })
    .filter(Boolean)
    .join('\n');
  clipboard = { els: structuredClone(els), text, from: doc.slide.id };
  if (cut) deleteEls(doc);
  return text;
}

export function hasClipboard(): boolean {
  return !!clipboard;
}

export function clipboardText(): string | null {
  return clipboard?.text ?? null;
}

/** Pastes copied elements (offset when pasting onto the same slide). */
export function pasteEls(doc: SlidesDoc): boolean {
  if (!clipboard) return false;
  const same = clipboard.from === doc.slide.id && clipboard.els.some((e) => findEl(doc.slide.elements, e.id));
  const copies = cloneEls(clipboard.els).map((e) => (same ? moveEl(e, 16, 16) : e));
  insertEls(doc, copies, 'Paste');
  return true;
}

/* ============================================================= creation */

export function textBox(x: number, y: number, w: number, text = '', size?: number): ShapeEl {
  return {
    id: newId(),
    type: 'shape',
    geom: 'rect',
    name: 'TextBox',
    x,
    y,
    w,
    h: 40,
    textbox: true,
    text: { paras: [{ runs: text ? [{ text }] : [], endRun: size ? { size } : undefined }], anchor: 't', autofit: 'resize', wrap: true },
  };
}

export function shapeEl(geom: string, x: number, y: number, w: number, h: number): ShapeEl {
  const lineOnly = geom === 'line' || geom.startsWith('straightConnector') || geom.startsWith('bentConnector') || geom.startsWith('curvedConnector') || geom === 'arc' || /bracket|brace/i.test(geom);
  const [g, variant] = geom.split(':');
  const el: ShapeEl = {
    id: newId(),
    type: 'shape',
    geom: g,
    name: 'Shape',
    x,
    y,
    w,
    h,
    fill: lineOnly ? { type: 'none' } : { type: 'solid', color: '@accent1' },
    line: lineOnly ? { color: '@accent1', width: 2.67, tail: variant ? 'triangle' : undefined, head: variant === 'double' ? 'triangle' : undefined } : { color: '@accent1-25', width: 1.33 },
  };
  if (!lineOnly) el.text = { paras: [{ runs: [], align: 'center', endRun: { color: '@bg1' } }], anchor: 'm', defaults: { color: '@bg1' } };
  return el;
}

/** Makes a presentation-wide change function usable with doc.commit. */
export function mapSlides(pres: Presentation, fn: (s: Slide, i: number) => Slide): Presentation {
  let changed = false;
  const slides = pres.slides.map((s, i) => {
    const n = fn(s, i);
    if (n !== s) changed = true;
    return n;
  });
  return changed ? { ...pres, slides } : pres;
}

export { mapEls };
