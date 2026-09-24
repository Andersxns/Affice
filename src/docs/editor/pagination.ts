import { Extension } from '@tiptap/core';
import type { Node as PMNode } from '@tiptap/pm/model';
import { Plugin, PluginKey, type EditorState } from '@tiptap/pm/state';
import { Decoration, DecorationSet, type EditorView } from '@tiptap/pm/view';

/**
 * Print-layout pagination.
 *
 * Content flows in one editable column; this plugin inserts invisible spacer widgets so that
 * each page's content fits between the page margins. Spacers are computed from a fresh,
 * spacer-free layout (removed and re-added within the same task, so nothing flickers).
 *
 * Break strategy (closest to what word processors do):
 *  - top-level blocks move to the next page when they don't fit,
 *  - long paragraphs split between lines (keeping at least 2 lines on each side),
 *  - lists break between items,
 *  - headings stay with the paragraph that follows them,
 *  - manual page breaks always start a new page.
 */

export interface PageGeometry {
  /** Full page height, px */
  height: number;
  /** Gap between pages on screen, px */
  gap: number;
  marginTop: number;
  marginBottom: number;
}

interface Break {
  pos: number;
  height: number;
  inline: boolean;
}

interface PagState {
  geometry: PageGeometry | null;
  breaks: Break[];
  deco: DecorationSet;
  pages: number;
  headingPages: Map<number, number>;
}

export const paginationKey = new PluginKey<PagState>('pagination');

function spacer(height: number, inline: boolean): HTMLElement {
  const el = document.createElement(inline ? 'span' : 'div');
  el.className = inline ? 'page-spacer page-spacer-inline' : 'page-spacer';
  el.style.height = `${height}px`;
  el.contentEditable = 'false';
  el.setAttribute('aria-hidden', 'true');
  return el;
}

function buildDeco(doc: PMNode, breaks: Break[]): DecorationSet {
  if (!breaks.length) return DecorationSet.empty;
  return DecorationSet.create(
    doc,
    breaks.map((b) =>
      Decoration.widget(b.pos, () => spacer(b.height, b.inline), {
        side: -1,
        key: `pg-${b.inline ? 'i' : 'b'}-${b.pos}-${Math.round(b.height)}`,
        ignoreSelection: true,
      }),
    ),
  );
}

interface Unit {
  pos: number;
  node: PMNode;
  dom: HTMLElement;
  kind: 'text' | 'block' | 'break' | 'item';
  isHeading: boolean;
}

/** Flattens the document into layout units (top-level blocks, list items). */
function collectUnits(view: EditorView): Unit[] {
  const units: Unit[] = [];
  const doc = view.state.doc;
  doc.forEach((node, offset) => {
    const name = node.type.name;
    if (name === 'bulletList' || name === 'orderedList' || name === 'taskList') {
      node.forEach((item, itemOffset) => {
        const pos = offset + 1 + itemOffset;
        const dom = view.nodeDOM(pos) as HTMLElement | null;
        if (dom instanceof HTMLElement) units.push({ pos, node: item, dom, kind: 'item', isHeading: false });
      });
      return;
    }
    const dom = view.nodeDOM(offset) as HTMLElement | null;
    if (!(dom instanceof HTMLElement)) return;
    const kind: Unit['kind'] = name === 'pageBreak' ? 'break' : name === 'paragraph' || name === 'heading' ? 'text' : 'block';
    units.push({ pos: offset, node, dom, kind, isHeading: name === 'heading' });
  });
  return units;
}

interface Line {
  top: number;
  bottom: number;
  left: number;
  midY: number;
  clientLeft: number;
}

/** Visual lines of a textblock, in layout px relative to the editor top. */
function lineBoxes(el: HTMLElement, originTop: number, zoom: number): Line[] {
  const range = document.createRange();
  range.selectNodeContents(el);
  const rects = Array.from(range.getClientRects()).filter((r) => r.height > 0 && r.width > 0);
  rects.sort((a, b) => a.top - b.top || a.left - b.left);
  const lines: Line[] = [];
  for (const r of rects) {
    const top = (r.top - originTop) / zoom;
    const bottom = (r.bottom - originTop) / zoom;
    const last = lines[lines.length - 1];
    if (last && top < last.bottom - 2) {
      // same line (overlapping vertically)
      last.bottom = Math.max(last.bottom, bottom);
      last.top = Math.min(last.top, top);
      if (r.left < last.clientLeft) last.clientLeft = r.left;
      last.midY = (last.top + last.bottom) / 2;
    } else {
      lines.push({ top, bottom, left: r.left, midY: (top + bottom) / 2, clientLeft: r.left });
    }
  }
  return lines;
}

function computeBreaks(view: EditorView, g: PageGeometry): { breaks: Break[]; pages: number; headingPages: Map<number, number> } {
  const units = collectUnits(view);
  const root = view.dom as HTMLElement;
  const rootRect = root.getBoundingClientRect();
  const zoom = rootRect.height && root.offsetHeight ? rootRect.height / root.offsetHeight : 1;
  const S = g.height + g.gap;
  const C = Math.max(40, g.height - g.marginTop - g.marginBottom);
  const breaks: Break[] = [];
  const headingPages = new Map<number, number>();
  let add = 0; // total spacer height inserted so far
  let page = 0;
  let forceNewPage = false;

  const unitTop = (u: Unit) => (u.dom.getBoundingClientRect().top - rootRect.top) / zoom;

  for (let i = 0; i < units.length; i++) {
    const u = units[i];
    const rect = u.dom.getBoundingClientRect();
    const natTop = (rect.top - rootRect.top) / zoom;
    const h = rect.height / zoom;
    let y = natTop + add;

    // Skip into the right page if this unit starts in the gap/margins or beyond the current page
    while (y >= page * S + C) {
      const nextStart = (page + 1) * S;
      if (y < nextStart) {
        const sp = nextStart - y;
        breaks.push({ pos: u.pos, height: sp, inline: false });
        add += sp;
        y = nextStart;
      }
      page++;
    }

    if (forceNewPage && y > page * S + 0.5) {
      const nextStart = (page + 1) * S;
      const sp = nextStart - y;
      breaks.push({ pos: u.pos, height: sp, inline: false });
      add += sp;
      y = nextStart;
      page++;
    }
    forceNewPage = false;

    if (u.kind === 'break') {
      forceNewPage = true;
      continue;
    }

    const pageEnd = page * S + C;
    if (y + h > pageEnd + 0.5 && y > page * S + 0.5) {
      // Doesn't fit on this page
      let handled = false;
      if (u.kind === 'text' && !u.isHeading) {
        const lines = lineBoxes(u.dom, rootRect.top, zoom);
        if (lines.length >= 4) {
          // lines in "shifted" coordinates
          let k = lines.findIndex((l) => l.bottom + add > pageEnd + 0.5);
          if (k >= 2 && lines.length - k >= 2) {
            // split: move lines k.. to next page(s)
            let splitLine = k;
            let curPage = page;
            while (splitLine >= 0) {
              const line = lines[splitLine];
              const coords = view.posAtCoords({ left: line.clientLeft + 1, top: rootRect.top + line.midY * zoom });
              if (!coords) break;
              const pos = coords.pos;
              const nextStart = (curPage + 1) * S;
              const sp = nextStart - (line.top + add);
              if (sp <= 0 || pos <= u.pos + 1 || pos >= u.pos + u.node.nodeSize - 1) break;
              breaks.push({ pos, height: sp, inline: true });
              add += sp;
              curPage++;
              handled = true;
              // does the rest overflow the new page too? (very long paragraphs)
              const newEnd = curPage * S + C;
              const next = lines.findIndex((l, idx) => idx > splitLine && l.bottom + add > newEnd + 0.5);
              if (next > 0 && lines.length - next >= 2 && next - splitLine >= 2) splitLine = next;
              else break;
            }
            if (handled) {
              page = curPage;
              const endY = natTop + h + add;
              while (endY > page * S + C + 0.5) page++;
              continue;
            }
          }
        }
      }
      if (!handled && h <= C) {
        // Move the whole unit — or the heading before it (keep with next)
        let target = u;
        let targetY = y;
        const prev = units[i - 1];
        if (prev && prev.isHeading && !breaks.some((b) => b.pos === prev.pos)) {
          const prevY = unitTop(prev) + add;
          if (prevY > page * S + 0.5 && prevY < pageEnd) {
            target = prev;
            targetY = prevY;
          }
        }
        const nextStart = (page + 1) * S;
        const sp = nextStart - targetY;
        breaks.push({ pos: target.pos, height: sp, inline: false });
        add += sp;
        page++;
        if (target !== u) {
          headingPages.set(target.pos, page + 1);
        }
        y = natTop + add;
      }
    }

    if (u.isHeading) headingPages.set(u.pos, page + 1);
    // a tall unit may run over several pages
    const endY = y + h;
    while (endY > page * S + C + 0.5) page++;
  }

  breaks.sort((a, b) => a.pos - b.pos);
  return { breaks, pages: page + 1, headingPages };
}

function sameBreaks(a: Break[], b: Break[]): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) {
    if (a[i].pos !== b[i].pos || a[i].inline !== b[i].inline || Math.abs(a[i].height - b[i].height) > 0.5) return false;
  }
  return true;
}

export interface PaginationOptions {
  onPaginate: (info: { pages: number; headingPages: Map<number, number> }) => void;
}

export const Pagination = Extension.create<PaginationOptions>({
  name: 'pagination',
  addOptions() {
    return { onPaginate: () => undefined };
  },
  addProseMirrorPlugins() {
    const ext = this;
    return [
      new Plugin<PagState>({
        key: paginationKey,
        state: {
          init: () => ({ geometry: null, breaks: [], deco: DecorationSet.empty, pages: 1, headingPages: new Map() }),
          apply(tr, prev, _old, newState) {
            const meta = tr.getMeta(paginationKey) as Partial<PagState> | undefined;
            if (meta) {
              const geometry = meta.geometry !== undefined ? meta.geometry : prev.geometry;
              const breaks = geometry ? (meta.breaks ?? prev.breaks) : [];
              return {
                geometry,
                breaks,
                deco: buildDeco(newState.doc, breaks),
                pages: meta.pages ?? prev.pages,
                headingPages: meta.headingPages ?? prev.headingPages,
              };
            }
            if (tr.docChanged && prev.breaks.length) {
              // keep spacers roughly in place until the next measurement
              const breaks = prev.breaks
                .map((b) => ({ ...b, pos: tr.mapping.map(b.pos, b.inline ? 1 : -1) }))
                .filter((b) => b.pos >= 0 && b.pos <= newState.doc.content.size);
              return { ...prev, breaks, deco: prev.deco.map(tr.mapping, tr.doc) };
            }
            return prev;
          },
        },
        props: {
          decorations(state) {
            return paginationKey.getState(state)?.deco;
          },
        },
        view(view) {
          let timer: ReturnType<typeof setTimeout> | undefined;
          let passes = 0;
          let destroyed = false;

          const measure = () => {
            timer = undefined;
            if (destroyed) return;
            const st = paginationKey.getState(view.state);
            const g = st?.geometry;
            if (!g) return;
            if (view.composing || !(view.dom as HTMLElement).offsetParent) {
              schedule(400);
              return;
            }
            // 1) remove spacers, 2) measure natural layout, 3) apply new spacers — same task, no flicker
            const old = st!.breaks;
            if (old.length) view.dispatch(view.state.tr.setMeta(paginationKey, { breaks: [] }).setMeta('addToHistory', false));
            const res = computeBreaks(view, g);
            view.dispatch(
              view.state.tr
                .setMeta(paginationKey, { breaks: res.breaks, pages: res.pages, headingPages: res.headingPages })
                .setMeta('addToHistory', false),
            );
            ext.options.onPaginate({ pages: res.pages, headingPages: res.headingPages });
            (ext.editor as unknown as { emit: (e: string) => void }).emit('paginate');
            // a second pass settles margin-collapsing differences
            if (!sameBreaks(old, res.breaks) && passes < 2) {
              passes++;
              schedule(30);
            } else passes = 0;
          };

          const schedule = (ms: number) => {
            if (timer) clearTimeout(timer);
            timer = setTimeout(measure, ms);
          };

          // images and fonts change heights after they load
          const ro = typeof ResizeObserver !== 'undefined' ? new ResizeObserver(() => schedule(120)) : null;
          ro?.observe(view.dom);
          const onFonts = () => schedule(50);
          document.fonts?.addEventListener?.('loadingdone', onFonts);
          (view.dom as HTMLElement & { __paginate?: () => void }).__paginate = () => schedule(0);
          schedule(60);

          return {
            update(v, prevState: EditorState) {
              const prev = paginationKey.getState(prevState);
              const cur = paginationKey.getState(v.state);
              if (cur?.geometry !== prev?.geometry) {
                passes = 0;
                schedule(0);
              } else if (v.state.doc !== prevState.doc) {
                passes = 0;
                schedule(v.state.doc.content.size > 60000 ? 450 : 220);
              }
            },
            destroy() {
              destroyed = true;
              if (timer) clearTimeout(timer);
              ro?.disconnect();
              document.fonts?.removeEventListener?.('loadingdone', onFonts);
            },
          };
        },
      }),
    ];
  },
});

/** Sets (or clears, with null) the page geometry. */
export function setPageGeometry(view: EditorView, g: PageGeometry | null): void {
  view.dispatch(view.state.tr.setMeta(paginationKey, { geometry: g, breaks: g ? undefined : [] }).setMeta('addToHistory', false));
}

export function repaginate(view: EditorView): void {
  (view.dom as HTMLElement & { __paginate?: () => void }).__paginate?.();
}
