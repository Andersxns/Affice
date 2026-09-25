import { Editor } from '@tiptap/core';
import { TextSelection } from '@tiptap/pm/state';
import { useEffect, useLayoutEffect, useRef } from 'react';
import { useSettings } from '@/app/settings';
import type { EditTarget, SlidesDoc } from '../doc';
import { findEl, type El, type ShapeEl, type TableEl, type TextBody } from '../model';
import { roleOf, DEFAULT_INSET, type TextContext } from '../render/text';
import { normalizeTable, styledCellText } from '../tables';
import { bodyToDoc, buildSlideTextExtensions, docToParas } from './textEditor';
import type { SlidesUI } from './ui';

export function slotKey(t: EditTarget): string {
  return t.cell ? `${t.el}:${t.cell[0]}:${t.cell[1]}` : t.el;
}

export function bodyOf(el: El | undefined, cell?: [number, number]): TextBody | undefined {
  if (!el) return undefined;
  if (el.type === 'shape') return el.text ?? { paras: [{ runs: [] }], anchor: el.textbox ? 't' : 'm' };
  if (el.type === 'table' && cell) return el.rows[cell[0]]?.cells[cell[1]]?.text;
  return undefined;
}

/** Writes a text body back into a shape or table cell. */
export function withBody(el: El, cell: [number, number] | undefined, body: TextBody): El {
  if (el.type === 'shape') return { ...el, text: body } as ShapeEl;
  if (el.type === 'table' && cell) {
    const [r, c] = cell;
    return { ...el, rows: el.rows.map((row, ri) => (ri !== r ? row : { ...row, cells: row.cells.map((x, ci) => (ci === c ? { ...x, text: body } : x)) })) } as TableEl;
  }
  return el;
}

interface Props {
  doc: SlidesDoc;
  ui: SlidesUI;
  target: EditTarget;
  root: () => HTMLElement | null;
  onExit: () => void;
  /** Tab pressed in a table cell. */
  onTab: (dir: 1 | -1) => void;
}

/**
 * Mounts a TipTap editor into the slot the renderer leaves in the edited shape (or table cell),
 * commits every change to the document (coalesced into one undo step per burst of typing) and keeps
 * auto-fit text boxes sized to their content.
 */
export function TextEditSession({ doc, ui, target, root, onExit, onTab }: Props) {
  const editorRef = useRef<Editor | null>(null);
  const bodyRef = useRef<TextBody | null>(null);
  const attached = useRef(false);
  const exitRef = useRef(onExit);
  exitRef.current = onExit;
  const tabRef = useRef(onTab);
  tabRef.current = onTab;
  const key = slotKey(target);

  useEffect(() => {
    const el = findEl(doc.slide.elements, target.el);
    const body = bodyOf(el, target.cell);
    if (!el || !body) {
      exitRef.current();
      return;
    }
    bodyRef.current = body;
    const ctx = (): TextContext => {
      const cur = findEl(doc.slide.elements, target.el) ?? el;
      return { pres: doc.pres, role: target.cell ? 'other' : roleOf(cur.ph), ph: target.cell ? undefined : cur.ph, slideNumber: doc.sel.slide + 1 };
    };
    // table cells take their text colour and weight from the table style
    const styled = (b: TextBody): TextBody => {
      const cur = findEl(doc.slide.elements, target.el);
      return target.cell && cur?.type === 'table' ? styledCellText(cur, target.cell[0], target.cell[1], b) : b;
    };
    const editor = new Editor({
      extensions: buildSlideTextExtensions({ theme: () => doc.pres.theme, ctx, body: () => styled(bodyRef.current ?? body) }),
      content: bodyToDoc(body),
      editorProps: {
        attributes: { class: 'sl-pm', spellcheck: String(useSettings.getState().settings.spellcheck) },
        handleKeyDown: (_view, ev) => {
          if (ev.key === 'Escape') {
            ev.preventDefault();
            exitRef.current();
            return true;
          }
          if (ev.key === 'Tab' && target.cell && !ev.ctrlKey && !ev.metaKey && !ev.altKey) {
            ev.preventDefault();
            tabRef.current(ev.shiftKey ? -1 : 1);
            return true;
          }
          return false;
        },
      },
      onUpdate: ({ editor: ed }) => commit(docToParas(ed.state.doc)),
      onSelectionUpdate: () => ui.emit(),
      onTransaction: ({ transaction }) => {
        if (!transaction.docChanged && transaction.storedMarksSet) ui.emit();
      },
    });
    editorRef.current = editor;
    attached.current = false;
    ui.textEditor = editor;
    ui.emit();

    // new paragraphs go into the element's current body, so ribbon changes made while typing (anchor, autofit…) stick
    const commit = (paras: TextBody['paras']) => {
      const slot = root()?.querySelector<HTMLElement>(`[data-slot="${CSS.escape(key)}"]`);
      doc.updateEls(
        'Typing',
        [target.el],
        (cur) => {
          const next: TextBody = { ...(bodyOf(cur, target.cell) ?? body), paras };
          bodyRef.current = next;
          let out = withBody(cur, target.cell, next);
          if (slot) out = fitToText(out, target.cell, slot, ui.effectiveZoom);
          return out;
        },
        true,
      );
    };
    return () => {
      editor.destroy();
      if (ui.textEditor === editor) {
        ui.textEditor = null;
        ui.emit();
      }
      editorRef.current = null;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [key]);

  // attach the ProseMirror DOM to the slot after every render (the slot can be re-created)
  useLayoutEffect(() => {
    const editor = editorRef.current;
    if (!editor || editor.isDestroyed) return;
    const cur = bodyOf(findEl(doc.slide.elements, target.el), target.cell);
    if (cur && cur !== bodyRef.current) {
      bodyRef.current = cur;
      // paragraph styling depends on body defaults; re-run the decorations
      editor.view.dispatch(editor.state.tr.setMeta('slideRestyle', true));
    }
    const slot = root()?.querySelector<HTMLElement>(`[data-slot="${CSS.escape(key)}"]`);
    if (!slot) return;
    const dom = editor.view.dom;
    if (dom.parentElement !== slot) {
      slot.appendChild(dom);
      if (!attached.current) {
        attached.current = true;
        placeCaret(editor, ui);
      }
    }
  });

  return null;
}

function placeCaret(editor: Editor, ui: SlidesUI): void {
  const pc = ui.pendingCaret;
  const text = ui.pendingText;
  ui.pendingCaret = null;
  ui.pendingText = null;
  editor.view.focus();
  // a text selection (not an AllSelection) keeps the first paragraph's alignment and bullets when typed over
  const allText = () => TextSelection.create(editor.state.doc, 1, Math.max(1, editor.state.doc.content.size - 1));
  if (text !== null) {
    editor.view.dispatch(editor.state.tr.setSelection(allText()).insertText(text));
    return;
  }
  if (pc === 'all') {
    editor.view.dispatch(editor.state.tr.setSelection(allText()));
    return;
  }
  if (pc && typeof pc === 'object') {
    const pos = editor.view.posAtCoords({ left: pc.x, top: pc.y });
    if (pos) {
      editor.view.dispatch(editor.state.tr.setSelection(TextSelection.create(editor.state.doc, pos.pos)));
      return;
    }
  }
  editor.commands.focus('end');
}

/** Grows auto-fit text boxes and shrinks overflowing text, like PowerPoint's AutoFit options. */
export function fitToText(el: El, cell: [number, number] | undefined, slot: HTMLElement, zoom: number): El {
  if (el.type === 'table') {
    // rows grow with their content: sync the model with the rendered heights
    const table = slot.closest('table');
    if (!table) return el;
    const trs = Array.from(table.querySelectorAll(':scope > tbody > tr')) as HTMLElement[];
    let changed = false;
    const rows = el.rows.map((r, i) => {
      const hgt = trs[i] ? trs[i].getBoundingClientRect().height / zoom : r.h;
      if (hgt > r.h + 0.5) {
        changed = true;
        return { ...r, h: Math.round(hgt) };
      }
      return r;
    });
    return changed ? normalizeTable({ ...el, rows }) : el;
  }
  if (el.type !== 'shape' || !el.text || cell) return el;
  const body = el.text;
  const inset = body.inset ?? DEFAULT_INSET;
  if (body.autofit === 'resize') {
    const contentH = slot.getBoundingClientRect().height / zoom;
    const h = Math.max(20, Math.ceil(contentH + inset[1] + inset[3]));
    let out: El = el;
    if (Math.abs(h - el.h) > 0.5) out = { ...out, h };
    if (body.wrap === false) {
      const w = Math.ceil(slot.scrollWidth + inset[0] + inset[2] + 2);
      if (Math.abs(w - el.w) > 0.5) out = { ...out, w };
    }
    return out;
  }
  if (body.autofit === 'shrink') {
    const avail = el.h - inset[1] - inset[3];
    const measure = (z: number) => {
      slot.style.zoom = String(z);
      return slot.getBoundingClientRect().height / zoom;
    };
    const prev = slot.style.zoom;
    let scale = 1;
    if (measure(1) > avail + 1) {
      let lo = 0.25;
      let hi = 1;
      for (let i = 0; i < 6; i++) {
        const mid = (lo + hi) / 2;
        if (measure(mid) > avail + 1) hi = mid;
        else lo = mid;
      }
      scale = Math.floor(lo * 40) / 40;
    }
    slot.style.zoom = prev;
    const cur = body.fontScale ?? 1;
    if (Math.abs(cur - scale) > 0.001) return { ...el, text: { ...body, fontScale: scale < 1 ? scale : undefined } };
  }
  return el;
}
