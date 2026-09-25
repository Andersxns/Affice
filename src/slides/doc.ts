import { findEl, type El, type GroupEl, type Presentation, type Slide } from './model';

export interface Selection {
  /** Index of the current slide. */
  slide: number;
  /** Selected element ids on the current slide. */
  els: string[];
  /** Slides selected in the slide panel (always contains the current slide). */
  slides: number[];
}

interface Entry {
  label: string;
  pres: Presentation;
  sel: Selection;
  at: number;
}

export interface EditTarget {
  el: string;
  /** Table cell being edited. */
  cell?: [number, number];
}

/**
 * Presentation document: immutable snapshots with undo/redo. Every change produces a new Presentation;
 * unchanged slides and elements keep their identity so React can skip re-rendering them.
 */
export class SlidesDoc {
  pres: Presentation;
  sel: Selection = { slide: 0, els: [], slides: [0] };
  editing: EditTarget | null = null;
  /** Transient element changes while dragging (not in the undo history). */
  preview: Map<string, Partial<El>> | null = null;
  version = 0;
  onChange?: () => void;
  private undoStack: Entry[] = [];
  private redoStack: Entry[] = [];
  private listeners = new Set<() => void>();
  private lastLabel = '';
  private lastAt = 0;

  constructor(pres: Presentation) {
    this.pres = pres;
  }

  subscribe = (fn: () => void) => {
    this.listeners.add(fn);
    return () => this.listeners.delete(fn);
  };
  getVersion = () => this.version;
  emit(): void {
    this.version++;
    for (const l of this.listeners) l();
  }

  get slide(): Slide {
    return this.pres.slides[Math.min(this.sel.slide, this.pres.slides.length - 1)];
  }

  get selected(): El[] {
    const s = this.slide;
    if (!s) return [];
    return this.sel.els.map((id) => findEl(s.elements, id)).filter((e): e is El => !!e);
  }

  canUndo(): boolean {
    return this.undoStack.length > 0;
  }
  canRedo(): boolean {
    return this.redoStack.length > 0;
  }
  undoLabel(): string | undefined {
    return this.undoStack[this.undoStack.length - 1]?.label;
  }
  redoLabel(): string | undefined {
    return this.redoStack[this.redoStack.length - 1]?.label;
  }

  /**
   * Replaces the presentation, recording the previous state for undo. Rapid repeats of the same
   * action (typing, nudging) within a second merge into one undo step when `coalesce` is set.
   */
  commit(label: string, next: Presentation, sel?: Partial<Selection>, coalesce = false): void {
    if (next === this.pres && !sel) return;
    const now = Date.now();
    const merge = coalesce && this.lastLabel === label && now - this.lastAt < 1000 && this.undoStack.length > 0;
    if (!merge && next !== this.pres) {
      this.undoStack.push({ label, pres: this.pres, sel: this.sel, at: now });
      if (this.undoStack.length > 300) this.undoStack.shift();
    }
    this.redoStack = [];
    this.lastLabel = label;
    this.lastAt = now;
    this.pres = { ...next, props: { ...next.props, modified: new Date().toISOString() } };
    if (sel) this.setSel(sel, false);
    else this.clampSel();
    this.onChange?.();
    this.emit();
  }

  /** Applies a change to the current slide. */
  updateSlide(label: string, fn: (s: Slide) => Slide, index = this.sel.slide, coalesce = false): void {
    const cur = this.pres.slides[index];
    if (!cur) return;
    const next = fn(cur);
    if (next === cur) return;
    const slides = [...this.pres.slides];
    slides[index] = next;
    this.commit(label, { ...this.pres, slides }, undefined, coalesce);
  }

  /** Applies a change to elements (by id, including elements inside groups) on the current slide. */
  updateEls(label: string, ids: string[], fn: (e: El) => El, coalesce = false): void {
    const want = new Set(ids);
    this.updateSlide(label, (s) => ({ ...s, elements: mapEls(s.elements, want, fn) }), this.sel.slide, coalesce);
  }

  undo(): string | null {
    const e = this.undoStack.pop();
    if (!e) return null;
    this.redoStack.push({ label: e.label, pres: this.pres, sel: this.sel, at: Date.now() });
    this.pres = e.pres;
    this.sel = e.sel;
    this.editing = null;
    this.lastLabel = '';
    this.clampSel();
    this.onChange?.();
    this.emit();
    return e.label;
  }

  redo(): string | null {
    const e = this.redoStack.pop();
    if (!e) return null;
    this.undoStack.push({ label: e.label, pres: this.pres, sel: this.sel, at: Date.now() });
    this.pres = e.pres;
    this.sel = e.sel;
    this.editing = null;
    this.lastLabel = '';
    this.clampSel();
    this.onChange?.();
    this.emit();
    return e.label;
  }

  setSel(sel: Partial<Selection>, emit = true): void {
    const next = { ...this.sel, ...sel };
    if (sel.slide !== undefined && sel.slides === undefined) next.slides = [sel.slide];
    if (sel.slide !== undefined && sel.slide !== this.sel.slide && sel.els === undefined) next.els = [];
    this.sel = next;
    this.clampSel();
    if (emit) this.emit();
  }

  selectEls(ids: string[]): void {
    this.setSel({ els: ids });
  }

  goToSlide(i: number): void {
    const n = Math.max(0, Math.min(this.pres.slides.length - 1, i));
    this.editing = null;
    this.setSel({ slide: n, els: [], slides: [n] });
  }

  setPreview(map: Map<string, Partial<El>> | null): void {
    this.preview = map;
    this.emit();
  }

  private clampSel(): void {
    const n = this.pres.slides.length;
    let slide = Math.max(0, Math.min(n - 1, this.sel.slide));
    if (n === 0) slide = 0;
    const s = this.pres.slides[slide];
    const els = s ? this.sel.els.filter((id) => !!findEl(s.elements, id)) : [];
    const slides = this.sel.slides.filter((i) => i >= 0 && i < n);
    if (!slides.includes(slide)) slides.push(slide);
    this.sel = { slide, els, slides: slides.sort((a, b) => a - b) };
    if (this.editing && !(s && findEl(s.elements, this.editing.el))) this.editing = null;
  }
}

export function mapEls(els: El[], want: Set<string>, fn: (e: El) => El): El[] {
  let changed = false;
  const out = els.map((e) => {
    if (want.has(e.id)) {
      const n = fn(e);
      if (n !== e) changed = true;
      return n;
    }
    if (e.type === 'group') {
      const kids = mapEls(e.children, want, fn);
      if (kids !== e.children) {
        changed = true;
        return { ...e, children: kids } as GroupEl;
      }
    }
    return e;
  });
  return changed ? out : els;
}

export function removeEls(els: El[], ids: Set<string>): El[] {
  return els.filter((e) => !ids.has(e.id)).map((e) => (e.type === 'group' ? ({ ...e, children: removeEls(e.children, ids) } as GroupEl) : e));
}
