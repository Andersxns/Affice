import type { Editor } from '@tiptap/core';
import type { Run } from '../model';

export type InsertMode =
  | { kind: 'shape'; geom: string }
  | { kind: 'textbox' }
  | { kind: 'table'; rows: number; cols: number };

export type ViewMode = 'normal' | 'sorter' | 'outline';

export interface Guides {
  x: number[];
  y: number[];
}

/** View state of a Slides editor (not saved with the file). */
export class SlidesUI {
  version = 0;
  zoom = 1;
  fit = true;
  /** Zoom actually used on screen (fit or manual). */
  effectiveZoom = 1;
  insert: InsertMode | null = null;
  guides: Guides | null = null;
  marquee: { x: number; y: number; w: number; h: number } | null = null;
  textEditor: Editor | null = null;
  /** Where to put the caret when a text editor opens: client point, "all", or "end". */
  pendingCaret: { x: number; y: number } | 'all' | 'end' | null = null;
  /** Text typed into a selected shape before its editor opened. */
  pendingText: string | null = null;
  view: ViewMode = 'normal';
  notes = true;
  panel = true;
  showGrid = false;
  snap = true;
  painter: { run?: Omit<Run, 'text'>; shape?: Record<string, unknown> } | null = null;
  /** Selected table cells while a table is selected: r1, c1, r2, c2. */
  cellSel: [number, number, number, number] | null = null;
  private listeners = new Set<() => void>();

  subscribe = (fn: () => void) => {
    this.listeners.add(fn);
    return () => this.listeners.delete(fn);
  };
  getVersion = () => this.version;
  emit(): void {
    this.version++;
    for (const l of this.listeners) l();
  }

  setZoom(z: number | 'fit'): void {
    if (z === 'fit') this.fit = true;
    else {
      this.fit = false;
      this.zoom = Math.max(0.1, Math.min(4, z));
    }
    this.emit();
  }

  setInsert(m: InsertMode | null): void {
    this.insert = m;
    this.emit();
  }
}
