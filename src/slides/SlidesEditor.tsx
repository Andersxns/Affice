import './slides.css';
import { EyeOff, LayoutGrid, NotebookPen, Presentation as PresIcon, StickyNote } from 'lucide-react';
import { useCallback, useEffect, useMemo, useRef, useState, useSyncExternalStore } from 'react';
import { SLIDES_EXPORT_TARGETS, SLIDES_SAVE_TARGETS } from '@/lib/formats';
import { api, isElectron } from '@/lib/platform';
import { Backstage } from '@/app/Backstage';
import { exportBytes, exportPdf, markDirty, printJob, saveDocument } from '@/app/fileOps';
import { useSettings } from '@/app/settings';
import { SaveState, ZoomControl } from '@/app/StatusBits';
import { registerController, useWorkspace, type Command, type Tab } from '@/app/workspace';
import { isMod } from '@/lib/utils';
import { nextFontSize } from '@/ui/fontControls';
import { openContextMenu, type MenuItem } from '@/ui/menu';
import { toast } from '@/ui/toast';
import { SlidesDoc } from './doc';
import { findEl, newId, type Anim, type AnimClass, type AnimEffect, type El, type ImageEl, type ShapeEl, type SlideChartType, type TableEl } from './model';
import * as ops from './ops';
import { applyDesign, applyFonts, applyPalette, buildLayouts, designById, FONT_PAIRS, PALETTES } from './themes';
import { createTable, deleteCol, deleteRow, insertCol, insertRow, mergeCells, splitCell } from './tables';
import { SAMPLE_CHART } from './render/chart';
import { SlideCanvas, type CanvasHandlers } from './editor/Canvas';
import { SlidePanel, NotesPane } from './editor/panels';
import { SlidesRibbon, QUICK_STYLES, ANIMATIONS, TRANSITIONS, type SlideActions, type SelKind } from './editor/SlidesRibbon';
import { SlidesUI } from './editor/ui';
import { AnimationPane, SlideFind, SlideSorter } from './editor/extras';
import { activeEditor, changeCase, clearFormatting, paraStyle, readTextState, runStyle, toggleMark } from './editor/textActions';
import { chartDataDialog, cropDialog, formatBackgroundDialog, formatShapeDialog, headerFooterDialog, iconPickerDialog, imageSize, insertTableDialog, linkDialog, pickImageDataUrl, slideSizeDialog } from './editor/dialogs';
import { exportPresentation, loadPresentation } from './formats';
import { Slideshow, SlidePreview } from './show/Slideshow';

export default function SlidesEditor({ tab, active }: { tab: Tab; active: boolean }) {
  const [state, setState] = useState<{ doc: SlidesDoc; ui: SlidesUI } | null>(null);
  useEffect(() => {
    let cancelled = false;
    void (async () => {
      try {
        const { pres, warning } = await loadPresentation(tab.source);
        if (cancelled) return;
        setState({ doc: new SlidesDoc(pres), ui: new SlidesUI() });
        if (warning) toast.warning(warning, { duration: 8000 });
      } catch (e) {
        console.error(e);
        toast.error(`Couldn't open ${tab.title}`, { detail: String((e as Error)?.message ?? e) });
        const { createPresentation } = await import('./themes');
        if (!cancelled) setState({ doc: new SlidesDoc(createPresentation('affice')), ui: new SlidesUI() });
      } finally {
        if (!cancelled) useWorkspace.getState().updateTab(tab.id, { loading: false, dirty: tab.source?.type === 'recovery', source: undefined });
      }
    })();
    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
  if (!state) return <div className="editor-loading">Opening {tab.title}…</div>;
  return <SlidesView tab={tab} active={active} doc={state.doc} ui={state.ui} />;
}

function selKind(els: El[]): SelKind {
  if (!els.length) return 'none';
  if (els.length > 1) return 'multi';
  const e = els[0];
  return e.type === 'shape' ? 'shape' : e.type;
}

/** Places a new object into the selected (or else the first) empty content placeholder, like PowerPoint does. */
function takePlaceholder(doc: SlidesDoc, kinds: string[]): El | null {
  const fits = (e: El) => !!e.ph && kinds.includes(e.ph) && e.type === 'shape' && !e.text?.paras.some((p) => p.runs.some((r) => r.text));
  const sel = doc.selected;
  if (sel.length === 1 && fits(sel[0])) return sel[0];
  return doc.slide.elements.find(fits) ?? null;
}

function fitInto(w: number, h: number, box: { x: number; y: number; w: number; h: number }): { x: number; y: number; w: number; h: number } {
  const s = Math.min(box.w / w, box.h / h, 1.5);
  const nw = Math.round(w * s);
  const nh = Math.round(h * s);
  return { x: Math.round(box.x + (box.w - nw) / 2), y: Math.round(box.y + (box.h - nh) / 2), w: nw, h: nh };
}

function SlidesView({ tab, active, doc, ui }: { tab: Tab; active: boolean; doc: SlidesDoc; ui: SlidesUI }) {
  useSyncExternalStore(doc.subscribe, doc.getVersion);
  useSyncExternalStore(ui.subscribe, ui.getVersion);
  const appSettings = useSettings((s) => s.settings);
  const [backstage, setBackstage] = useState(false);
  const [show, setShow] = useState<{ start: number; presenter: boolean } | null>(null);
  const [preview, setPreview] = useState<'transition' | 'anims' | null>(null);
  const [animPane, setAnimPane] = useState(false);
  const [find, setFind] = useState<null | 'find' | 'replace'>(null);
  const [panelW, setPanelW] = useState(210);
  const [notesH, setNotesH] = useState(110);
  const rootRef = useRef<HTMLDivElement>(null);
  const pres = doc.pres;

  useEffect(() => {
    doc.onChange = () => markDirty(tab.id);
    return () => {
      doc.onChange = undefined;
    };
  }, [doc, tab.id]);

  const focusCanvas = useCallback(() => {
    requestAnimationFrame(() => rootRef.current?.querySelector<HTMLTextAreaElement>('.sl-sink')?.focus({ preventScroll: true }));
  }, []);

  const title = () => useWorkspace.getState().tabs.find((t) => t.id === tab.id)?.title ?? 'Presentation';

  /* ------------------------------------------------------------ images */

  const insertImageSrc = useCallback(
    async (src: string, name: string, at?: { x: number; y: number }, svg?: string) => {
      const size = await imageSize(src);
      const ph = at ? null : takePlaceholder(doc, ['pic', 'obj']);
      const box = ph ? { x: ph.x, y: ph.y, w: ph.w, h: ph.h } : { x: pres.size.w * 0.15, y: pres.size.h * 0.15, w: pres.size.w * 0.7, h: pres.size.h * 0.7 };
      let f = fitInto(size.w, size.h, box);
      if (at) f = { ...f, x: Math.round(at.x - f.w / 2), y: Math.round(at.y - f.h / 2) };
      const img: ImageEl = { id: newId(), type: 'image', name: svg ? 'Icon' : 'Picture', src, svg, alt: name.replace(/\.[^.]+$/, ''), ...f };
      if (ph) {
        doc.updateSlide('Insert picture', (s) => ({ ...s, elements: s.elements.map((e) => (e.id === ph.id ? img : e)) }));
        doc.selectEls([img.id]);
      } else ops.insertEls(doc, [img], 'Insert picture');
    },
    [doc, pres],
  );

  const insertFiles = useCallback(
    async (files: File[], at?: { x: number; y: number }) => {
      for (const f of files) {
        if (!f.type.startsWith('image/')) continue;
        const src = await new Promise<string>((res) => {
          const r = new FileReader();
          r.onload = () => res(String(r.result));
          r.readAsDataURL(f);
        });
        await insertImageSrc(src, f.name, at, f.type === 'image/svg+xml' ? await f.text() : undefined);
      }
    },
    [insertImageSrc],
  );

  /* ----------------------------------------------------------- actions */

  const selectedTable = (): TableEl | null => {
    const e = doc.editing ? findEl(doc.slide.elements, doc.editing.el) : doc.selected[0];
    return e && e.type === 'table' ? e : null;
  };
  const currentCell = (): [number, number] => doc.editing?.cell ?? (ui.cellSel ? [ui.cellSel[0], ui.cellSel[1]] : [0, 0]);

  const updateTable = (label: string, fn: (t: TableEl) => TableEl | null) => {
    const t = selectedTable();
    if (!t) return;
    const next = fn(t);
    doc.editing = null;
    if (!next) ops.deleteEls(doc, [t.id]);
    else doc.updateEls(label, [t.id], () => next);
  };

  const a: SlideActions = useMemo(() => {
    const sel = doc.selected;
    const slide = doc.slide;
    const firstAnim = slide?.anims?.find((x) => sel.some((e) => e.id === x.el));
    const text = readTextState(doc, ui);
    const setAnims = (label: string, fn: (anims: Anim[]) => Anim[]) => doc.updateSlide(label, (s) => ({ ...s, anims: fn(s.anims ?? []) }));
    return {
      undo: () => {
        const ed = activeEditor(ui);
        if (ed && ed.can().undo()) ed.commands.undo();
        else if (!doc.undo()) toast.info('Nothing to undo');
      },
      redo: () => {
        const ed = activeEditor(ui);
        if (ed && ed.can().redo()) ed.commands.redo();
        else if (!doc.redo()) toast.info('Nothing to redo');
      },
      cut: () => {
        if (activeEditor(ui)) return api.app.editCommand('cut');
        ops.copyEls(doc, true);
        void navigator.clipboard?.writeText(ops.clipboardText() ?? '').catch(() => undefined);
      },
      copy: () => {
        if (activeEditor(ui)) return api.app.editCommand('copy');
        const t = ops.copyEls(doc, false);
        void navigator.clipboard?.writeText(t || ' ').catch(() => undefined);
      },
      paste: () => {
        if (activeEditor(ui)) return api.app.editCommand('paste');
        if (!ops.pasteEls(doc)) void navigator.clipboard?.readText().then((t) => t && ops.insertEls(doc, [ops.textBox(pres.size.w / 2 - 300, pres.size.h / 2 - 40, 600, t)], 'Paste'));
      },
      painter: () => {
        if (ui.painter) {
          ui.painter = null;
          ui.emit();
          return;
        }
        const e = doc.selected[0];
        if (!e) return toast.info('Select an object to copy its formatting.');
        const shape: Record<string, unknown> = {};
        if (e.type === 'shape') Object.assign(shape, { fill: e.fill, line: e.line, shadow: e.shadow });
        if (e.type === 'image') Object.assign(shape, { line: e.line, shadow: e.shadow });
        ui.painter = { run: ops.selectionRun(doc), shape };
        ui.emit();
        toast.info('Click an object to apply the formatting.', { duration: 2500 });
      },
      painterOn: !!ui.painter,
      newSlide: (layoutId) => ops.addSlide(doc, layoutId),
      duplicateSlide: () => ops.duplicateSlides(doc),
      deleteSlide: () => ops.deleteSlides(doc),
      setLayout: (id) => ops.setLayout(doc, id),
      resetSlide: () => ops.resetSlides(doc),
      hideSlide: () => ops.toggleHidden(doc),
      text,
      runStyle: (patch) => runStyle(doc, ui, patch),
      toggle: (m) => toggleMark(doc, ui, m, text),
      growFont: (dir) => runStyle(doc, ui, { size: nextFontSize(text.size ?? 18, dir) }),
      clearFormatting: () => clearFormatting(doc, ui),
      changeCase: (m) => changeCase(doc, ui, m),
      align: (al) => paraStyle(doc, ui, { align: al === 'left' ? undefined : al }),
      bullets: (b) => paraStyle(doc, ui, { bullet: b, marL: undefined, indent: undefined }),
      level: (dir) => paraStyle(doc, ui, (p) => ({ level: Math.max(0, Math.min(8, (p.level ?? 0) + dir)) || undefined, marL: undefined, indent: undefined })),
      lineSpacing: (v) => paraStyle(doc, ui, { lineSpacing: v }),
      anchor: (v) => {
        const t = selectedTable();
        if (t && doc.editing?.cell) {
          const [r, c] = doc.editing.cell;
          doc.updateEls('Cell alignment', [t.id], (e) => ({ ...(e as TableEl), rows: (e as TableEl).rows.map((row, ri) => (ri !== r ? row : { ...row, cells: row.cells.map((x, ci) => (ci === c ? { ...x, text: { ...x.text, anchor: v } } : x)) })) }));
        } else ops.applyBodyStyle(doc, { anchor: v });
      },
      columns: (n) => ops.applyBodyStyle(doc, { columns: n > 1 ? n : undefined }),
      insertShape: (geom) => ui.setInsert({ kind: 'shape', geom }),
      insertTextBox: () => ui.setInsert({ kind: 'textbox' }),
      insertTable: (rows, cols) => {
        const ph = takePlaceholder(doc, ['obj']);
        const box = ph ? { x: ph.x, y: ph.y, w: ph.w } : { x: 88, y: Math.round(pres.size.h * 0.27), w: pres.size.w - 176 };
        const t = createTable(rows, cols, box.x, box.y, box.w, 44);
        if (ph) {
          doc.updateSlide('Insert table', (s) => ({ ...s, elements: s.elements.map((e) => (e.id === ph.id ? t : e)) }));
          doc.selectEls([t.id]);
        } else ops.insertEls(doc, [t], 'Insert table');
      },
      insertImage: () => {
        void pickImageDataUrl().then((img) => img && insertImageSrc(img.src, img.name));
      },
      insertIcon: () => {
        const ph = takePlaceholder(doc, ['pic', 'obj']);
        void iconPickerDialog(pres.theme).then((r) => {
          if (!r) return;
          const src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(r.svg)}`;
          const s = ph ? Math.round(Math.min(ph.w, ph.h, 300) * 0.8) : 150;
          const cx = ph ? ph.x + ph.w / 2 : pres.size.w / 2;
          const cy = ph ? ph.y + ph.h / 2 : pres.size.h / 2;
          const img: ImageEl = { id: newId(), type: 'image', name: 'Icon', src, svg: r.svg, alt: r.name.replace(/([a-z])([A-Z])/g, '$1 $2'), x: Math.round(cx - s / 2), y: Math.round(cy - s / 2), w: s, h: s };
          if (ph && findEl(doc.slide.elements, ph.id)) {
            doc.updateSlide('Insert icon', (sl) => ({ ...sl, elements: sl.elements.map((e) => (e.id === ph.id ? img : e)) }));
            doc.selectEls([img.id]);
          } else ops.insertEls(doc, [img], 'Insert icon');
        });
      },
      insertChart: (type: SlideChartType) => {
        const ph = takePlaceholder(doc, ['obj']);
        const box = ph ? { x: ph.x, y: ph.y, w: ph.w, h: ph.h } : { x: Math.round(pres.size.w * 0.15), y: Math.round(pres.size.h * 0.22), w: Math.round(pres.size.w * 0.7), h: Math.round(pres.size.h * 0.66) };
        const el: El = { id: newId(), type: 'chart', name: 'Chart', ...box, chart: { ...structuredClone(SAMPLE_CHART), type } };
        if (ph) {
          doc.updateSlide('Insert chart', (s) => ({ ...s, elements: s.elements.map((e) => (e.id === ph.id ? el : e)) }));
          doc.selectEls([el.id]);
        } else ops.insertEls(doc, [el], 'Insert chart');
        void chartDataDialog((el as { chart: typeof SAMPLE_CHART }).chart, pres.theme).then((c) => c && doc.updateEls('Chart data', [el.id], (e) => ({ ...e, chart: c }) as El));
      },
      insertLink: () => {
        const ed = activeEditor(ui);
        if (ed) {
          const cur = ed.getAttributes('link').href as string | undefined;
          void linkDialog(doc, cur).then((href) => {
            if (href === undefined) return;
            if (href === null) ed.chain().focus().extendMarkRange('link').unsetLink().run();
            else if (ed.state.selection.empty && !cur) ed.chain().focus().insertContent({ type: 'text', text: href.replace(/^#slide:/, 'Slide '), marks: [{ type: 'link', attrs: { href } }] }).run();
            else ed.chain().focus().extendMarkRange('link').setLink({ href }).run();
          });
          return;
        }
        const e = doc.selected[0];
        if (!e) return toast.info('Select an object or some text first.');
        void linkDialog(doc, e.link).then((href) => href !== undefined && ops.patchEls(doc, 'Link', { link: href ?? undefined }));
      },
      insertField: (kind) => {
        const ed = activeEditor(ui);
        const label = kind === 'slidenum' ? String(doc.sel.slide + 1) : new Date().toLocaleDateString();
        if (ed) {
          ed.chain().focus().insertContent({ type: 'text', text: label, marks: [{ type: 'field', attrs: { kind } }] }).run();
          return;
        }
        const tb = ops.textBox(pres.size.w - 300, pres.size.h - 60, 220, '', 14);
        tb.text!.paras = [{ runs: [{ text: label, field: kind, size: 14 }], align: 'right' }];
        ops.insertEls(doc, [tb], kind === 'slidenum' ? 'Slide number' : 'Date');
      },
      headerFooter: () => void headerFooterDialog(doc).then((f) => f && doc.commit('Header & footer', { ...doc.pres, footer: f })),
      shapeFill: (fill) => {
        if (selectedTable()) return a.cellFill(fill);
        ops.setShapeFill(doc, fill ?? { type: 'none' });
      },
      shapeLine: (patch) => ops.setShapeLine(doc, patch),
      shadow: (on) => ops.setShadow(doc, on ? { color: '#000000/35', blur: 14, dist: 5, angle: 60 } : undefined),
      quickStyle: (i) => {
        const q = QUICK_STYLES[i];
        ops.patchEls(doc, 'Shape style', (e) => {
          if (e.type !== 'shape') return null;
          const text = e.text ? { ...e.text, defaults: { ...(e.text.defaults ?? {}), color: q.text } } : { paras: [{ runs: [], align: 'center' as const }], anchor: 'm' as const, defaults: { color: q.text } };
          return { fill: q.fill ? { type: 'solid', color: q.fill } : { type: 'none' }, line: q.line ? { color: q.line, width: 1.33 } : null, text } as Partial<ShapeEl>;
        });
      },
      arrange: (how) => ops.zOrder(doc, how),
      alignEls: (how) => ops.align(doc, how, doc.sel.els.length === 1),
      distribute: (axis) => ops.distribute(doc, axis),
      group: () => ops.group(doc),
      ungroup: () => ops.ungroup(doc),
      rotate: (deg) => ops.rotateEls(doc, deg),
      flip: (axis) => ops.flipEls(doc, axis),
      setSize: (w, h) => doc.updateEls('Size', doc.sel.els, (e) => ops.setFrame(e, { x: e.x, y: e.y, w: w ?? e.w, h: h ?? e.h })),
      changeShape: (geom) => ops.changeGeom(doc, geom.split(':')[0]),
      formatShape: () => {
        const e = doc.selected[0];
        if (!e) return toast.info('Select a shape or picture first.');
        void formatShapeDialog(doc, e).then((f) => {
          if (!f) return;
          doc.updateEls('Format shape', [e.id], (x) => {
            let out: El = { ...x };
            if (f.frame) out = { ...ops.setFrame(out, { x: f.frame.x, y: f.frame.y, w: f.frame.w, h: f.frame.h }), rot: f.frame.rot || undefined };
            if (out.type === 'shape') {
              if (f.fill) out.fill = f.fill;
              if (f.line !== undefined) out.line = f.line;
              if (f.shadow !== undefined) out.shadow = f.shadow ?? undefined;
              if (f.text && out.text) out.text = { ...out.text, ...f.text };
              else if (f.text && !out.text) out.text = { paras: [{ runs: [] }], ...f.text };
            }
            if (out.type === 'image') {
              if (f.line !== undefined) out.line = f.line;
              if (f.shadow !== undefined) out.shadow = f.shadow ?? undefined;
            }
            out.alt = f.alt || undefined;
            out.opacity = f.opacity;
            return out;
          });
        });
      },
      design: (id) => doc.commit('Design', applyDesign(doc.pres, id)),
      palette: (i) => doc.commit('Theme colours', applyPalette(doc.pres, PALETTES[i].colors)),
      fonts: (i) => doc.commit('Theme fonts', applyFonts(doc.pres, FONT_PAIRS[i])),
      slideSize: (kind) => {
        const apply = (w: number, h: number, scale: boolean) => {
          const p = doc.pres;
          const sx = w / p.size.w;
          const sy = h / p.size.h;
          const s = Math.min(sx, sy);
          const ox = (w - p.size.w * s) / 2;
          const oy = (h - p.size.h * s) / 2;
          const fit = (e: El): El => (scale ? ops.setFrame(e, { x: ox + e.x * s, y: oy + e.y * s, w: e.w * s, h: e.h * s }) : e);
          const design = designById(p.design);
          const layouts = design ? buildLayouts(design, { w, h }) : p.layouts.map((l) => ({ ...l, decor: l.decor.map((d) => ops.setFrame(d, { x: d.x * sx, y: d.y * sy, w: d.w * sx, h: d.h * sy })), placeholders: l.placeholders.map((d) => ops.setFrame(d, { x: d.x * sx, y: d.y * sy, w: d.w * sx, h: d.h * sy })) }));
          doc.commit('Slide size', { ...p, size: { w, h }, layouts, slides: p.slides.map((sl) => ({ ...sl, elements: sl.elements.map(fit) })) });
        };
        if (kind === '16:9') apply(1280, 720, true);
        else if (kind === '4:3') apply(960, 720, true);
        else void slideSizeDialog(doc).then((r) => r && apply(r.w, r.h, r.scale));
      },
      formatBackground: () =>
        void formatBackgroundDialog(doc, (fill, all, hide) => {
          ops.setBackground(doc, fill, all);
          ops.setHideDecor(doc, hide);
        }),
      transition: (type) => {
        const def = TRANSITIONS.find((t) => t.type === type);
        ops.setTransition(doc, type === 'none' ? undefined : { ...(doc.slide.transition ?? {}), type, dir: def?.dirs?.[0], dur: doc.slide.transition?.dur ?? (type === 'morph' ? 1000 : 700) });
        if (type !== 'none') setPreview('transition');
      },
      transitionPatch: (p) => ops.patchTransition(doc, p),
      transitionAll: () => {
        ops.setTransition(doc, doc.slide.transition, true);
        toast.success('Applied to all slides', { duration: 1500 });
      },
      previewTransition: () => setPreview('transition'),
      animate: (cls: AnimClass, effect: AnimEffect) => {
        const ids = doc.sel.els;
        if (!ids.length) return;
        const def = ANIMATIONS.find((x) => x.cls === cls && x.effect === effect);
        setAnims('Animation', (anims) => {
          const out = [...anims];
          for (const id of ids) {
            const i = out.findIndex((x) => x.el === id);
            const next: Anim = { id: newId('a'), el: id, cls, effect, dir: def?.dirs?.[0], start: i >= 0 ? out[i].start : 'click', dur: i >= 0 ? out[i].dur : effect === 'appear' || effect === 'disappear' ? 1 : 500, delay: i >= 0 ? out[i].delay : 0, byPara: i >= 0 ? out[i].byPara : undefined };
            if (i >= 0) out[i] = next;
            else out.push(next);
          }
          return out;
        });
        setPreview('anims');
      },
      animPatch: (p) => {
        const ids = new Set(doc.sel.els);
        setAnims('Animation', (anims) => anims.map((x) => (ids.has(x.el) ? { ...x, ...p } : x)));
      },
      removeAnim: () => {
        const ids = new Set(doc.sel.els);
        setAnims('Remove animation', (anims) => anims.filter((x) => !ids.has(x.el)));
      },
      moveAnim: (dir) => {
        const ids = new Set(doc.sel.els);
        setAnims('Reorder animation', (anims) => {
          const out = [...anims];
          const i = out.findIndex((x) => ids.has(x.el));
          const j = i + dir;
          if (i < 0 || j < 0 || j >= out.length) return anims;
          [out[i], out[j]] = [out[j], out[i]];
          return out;
        });
      },
      previewAnim: () => setPreview('anims'),
      animPaneOn: animPane,
      toggleAnimPane: () => setAnimPane((v) => !v),
      startShow: (fromCurrent) => setShow({ start: fromCurrent ? doc.sel.slide : 0, presenter: false }),
      presenter: () => setShow({ start: doc.sel.slide, presenter: true }),
      view: (mode) => {
        ui.view = mode;
        ui.emit();
      },
      toggleNotes: () => {
        ui.notes = !ui.notes;
        ui.emit();
      },
      toggleGrid: () => {
        ui.showGrid = !ui.showGrid;
        ui.emit();
      },
      toggleSnap: () => {
        ui.snap = !ui.snap;
        ui.emit();
      },
      zoom: (z) => ui.setZoom(z),
      tableStyle: (p) => updateTable('Table style', (t) => ({ ...t, style: { ...(t.style ?? {}), ...p } })),
      tableInsert: (where) => {
        const [r, c] = currentCell();
        updateTable('Insert', (t) => (where === 'above' ? insertRow(t, r) : where === 'below' ? insertRow(t, r + 1) : where === 'left' ? insertCol(t, c) : insertCol(t, c + 1)));
      },
      tableDelete: (what) => {
        const [r, c] = currentCell();
        updateTable('Delete', (t) => (what === 'table' ? null : what === 'row' ? deleteRow(t, r) : deleteCol(t, c)));
      },
      mergeCells: () => {
        const cs = ui.cellSel;
        if (!cs || (cs[0] === cs[2] && cs[1] === cs[3])) return toast.info('Drag across the cells you want to merge first.');
        updateTable('Merge cells', (t) => mergeCells(t, Math.min(cs[0], cs[2]), Math.min(cs[1], cs[3]), Math.max(cs[0], cs[2]), Math.max(cs[1], cs[3])));
      },
      splitCell: () => {
        const [r, c] = currentCell();
        updateTable('Split cell', (t) => splitCell(t, r, c));
      },
      cellFill: (fill) => {
        const cs = ui.cellSel;
        const cell = doc.editing?.cell;
        updateTable('Cell shading', (t) => {
          const inRange = (r: number, c: number) => (cell ? r === cell[0] && c === cell[1] : cs ? r >= Math.min(cs[0], cs[2]) && r <= Math.max(cs[0], cs[2]) && c >= Math.min(cs[1], cs[3]) && c <= Math.max(cs[1], cs[3]) : true);
          return { ...t, rows: t.rows.map((row, r) => ({ ...row, cells: row.cells.map((x, c) => (inRange(r, c) ? { ...x, fill: fill ?? { type: 'none' } } : x)) })) };
        });
      },
      editChart: () => {
        const e = doc.selected[0];
        if (e?.type !== 'chart') return;
        void chartDataDialog(e.chart, pres.theme).then((c) => c && doc.updateEls('Chart data', [e.id], (x) => ({ ...x, chart: c }) as El));
      },
      chartType: (t) => doc.updateEls('Chart type', doc.sel.els, (e) => (e.type === 'chart' ? { ...e, chart: { ...e.chart, type: t } } : e)),
      chartPatch: (p) => doc.updateEls('Chart', doc.sel.els, (e) => (e.type === 'chart' ? { ...e, chart: { ...e.chart, ...p } } : e)),
      pictureShape: (geom) => ops.patchEls(doc, 'Picture shape', (e) => (e.type === 'image' ? { geom: geom === 'rect' ? undefined : geom } : null)),
      crop: () => {
        const e = doc.selected[0];
        if (e?.type !== 'image') return;
        void cropDialog(e).then((crop) => {
          if (!crop) return;
          const [l0, t0, r0, b0] = e.crop ?? [0, 0, 0, 0];
          // keep the picture's scale: the frame shrinks or grows with the crop
          const fullW = e.w / Math.max(0.01, 1 - l0 - r0);
          const fullH = e.h / Math.max(0.01, 1 - t0 - b0);
          const [l, t, r, b] = crop;
          doc.updateEls('Crop', [e.id], (x) => ({ ...x, crop: crop.some((v) => v > 0.0005) ? crop : undefined, x: Math.round(e.x - (l0 - l) * fullW), y: Math.round(e.y - (t0 - t) * fullH), w: Math.round(fullW * (1 - l - r)), h: Math.round(fullH * (1 - t - b)) }) as El);
        });
      },
      resetPicture: () => ops.patchEls(doc, 'Reset picture', (e) => (e.type === 'image' ? { crop: undefined, geom: undefined, line: undefined, shadow: undefined, rot: undefined, flipH: undefined, flipV: undefined, opacity: undefined } : null)),
      find: (replace) => setFind(replace ? 'replace' : 'find'),
      selectAll: () => doc.selectEls(doc.slide.elements.filter((e) => !e.hidden).map((e) => e.id)),
      selKind: selKind(sel),
      selected: sel,
      currentAnim: firstAnim ? { start: firstAnim.start, dur: firstAnim.dur, delay: firstAnim.delay, dir: firstAnim.dir, effect: firstAnim.effect, cls: firstAnim.cls } : undefined,
      currentTransition: slide?.transition,
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [doc, ui, doc.version, ui.version, animPane, pres, insertImageSrc]);

  /* ------------------------------------------------------ context menus */

  const slideMenu = useCallback(
    (e: React.MouseEvent | { clientX: number; clientY: number }, index: number | null): void => {
      const items: MenuItem[] = [
        { label: 'New slide', shortcut: 'Ctrl+M', onSelect: () => ops.addSlide(doc, undefined, index === null ? doc.pres.slides.length : index + 1) },
        ...(index !== null
          ? [
              { label: 'Duplicate slide', shortcut: 'Ctrl+D', onSelect: () => ops.duplicateSlides(doc) },
              { label: 'Delete slide', shortcut: 'Del', danger: true, onSelect: () => ops.deleteSlides(doc) },
              { separator: true },
              { label: 'Layout', submenu: doc.pres.layouts.map((l) => ({ label: l.name, checked: doc.slide.layout === l.id, onSelect: () => ops.setLayout(doc, l.id) })) },
              { label: 'Reset slide', onSelect: () => ops.resetSlides(doc) },
              { label: 'Format background…', onSelect: a.formatBackground },
              { separator: true },
              { label: doc.pres.slides[index]?.hidden ? 'Show slide' : 'Hide slide', icon: <EyeOff size={15} />, onSelect: () => ops.toggleHidden(doc) },
              { label: 'Start show from here', icon: <PresIcon size={15} />, onSelect: () => setShow({ start: index, presenter: false }) },
            ]
          : []),
      ];
      openContextMenu(e, items);
    },
    [doc, a],
  );

  const canvasHandlers: CanvasHandlers = useMemo(
    () => ({
      onContextMenu: (e, target) => {
        if (!target) {
          openContextMenu(e, [
            { label: 'Paste', shortcut: 'Ctrl+V', disabled: !ops.hasClipboard(), onSelect: () => ops.pasteEls(doc) },
            { separator: true },
            { label: 'Layout', submenu: doc.pres.layouts.map((l) => ({ label: l.name, checked: doc.slide.layout === l.id, onSelect: () => ops.setLayout(doc, l.id) })) },
            { label: 'Reset slide', onSelect: () => ops.resetSlides(doc) },
            { label: 'Format background…', onSelect: a.formatBackground },
            { separator: true },
            { label: 'New slide', shortcut: 'Ctrl+M', onSelect: () => ops.addSlide(doc) },
            { label: 'Hide slide', checked: !!doc.slide.hidden, onSelect: () => ops.toggleHidden(doc) },
          ]);
          return;
        }
        const isText = target.type === 'shape';
        openContextMenu(e, [
          { label: 'Cut', shortcut: 'Ctrl+X', onSelect: a.cut },
          { label: 'Copy', shortcut: 'Ctrl+C', onSelect: a.copy },
          { label: 'Paste', shortcut: 'Ctrl+V', disabled: !ops.hasClipboard(), onSelect: () => ops.pasteEls(doc) },
          { label: 'Duplicate', shortcut: 'Ctrl+D', onSelect: () => ops.duplicateEls(doc) },
          { label: 'Delete', shortcut: 'Del', danger: true, onSelect: () => ops.deleteEls(doc) },
          { separator: true },
          ...(isText ? [{ label: 'Edit text', shortcut: 'Enter', onSelect: () => ((doc.editing = { el: target.id }), (ui.pendingCaret = 'end'), doc.emit()) }] : []),
          ...(target.type === 'chart' ? [{ label: 'Edit data…', onSelect: a.editChart }] : []),
          ...(target.type === 'image' ? [{ label: 'Crop…', onSelect: a.crop }, { label: 'Reset picture', onSelect: a.resetPicture }] : []),
          { label: 'Group', shortcut: 'Ctrl+G', disabled: doc.sel.els.length < 2, onSelect: () => ops.group(doc) },
          { label: 'Ungroup', shortcut: 'Ctrl+Shift+G', disabled: target.type !== 'group', onSelect: () => ops.ungroup(doc) },
          {
            label: 'Order',
            submenu: [
              { label: 'Bring to front', onSelect: () => ops.zOrder(doc, 'front') },
              { label: 'Bring forward', onSelect: () => ops.zOrder(doc, 'forward') },
              { label: 'Send backward', onSelect: () => ops.zOrder(doc, 'backward') },
              { label: 'Send to back', onSelect: () => ops.zOrder(doc, 'back') },
            ],
          },
          { separator: true },
          { label: 'Link…', shortcut: 'Ctrl+K', onSelect: a.insertLink },
          { label: target.locked ? 'Unlock' : 'Lock position', onSelect: () => ops.patchEls(doc, target.locked ? 'Unlock' : 'Lock', { locked: !target.locked || undefined }) },
          { label: 'Format shape…', onSelect: a.formatShape },
        ]);
      },
      onEditChart: () => a.editChart(),
      onPlaceholderInsert: (kind) => {
        if (kind === 'table') void insertTableDialog().then((r) => r && a.insertTable(r.rows, r.cols));
        else if (kind === 'chart') a.insertChart('column');
        else if (kind === 'picture') a.insertImage();
        else a.insertIcon();
      },
      onDropFiles: (files, at) => void insertFiles(files, at),
      onPasteImage: (blob) => void insertFiles([new File([blob], 'Pasted image.png', { type: blob.type || 'image/png' })]),
      onShortcut: (e) => {
        if (!isMod(e)) return false;
        const k = e.key.toLowerCase();
        if (k === 'z' && !e.shiftKey) return a.undo(), true;
        if (k === 'y' || (k === 'z' && e.shiftKey)) return a.redo(), true;
        if (k === 'm') return ops.addSlide(doc), true;
        if (k === 'b') return a.toggle('b'), true;
        if (k === 'i') return a.toggle('i'), true;
        if (k === 'u') return a.toggle('u'), true;
        if (k === 'e') return a.align('center'), true;
        if (k === 'l') return a.align('left'), true;
        if (k === 'r') return a.align('right'), true;
        if (k === 'j') return a.align('justify'), true;
        if (k === 'k') return a.insertLink(), true;
        if (k === 'f') return setFind('find'), true;
        if (k === 'h') return setFind('replace'), true;
        if (e.key === ']' || e.key === '>') return e.shiftKey ? ops.zOrder(doc, 'front') : a.growFont(1), true;
        if (e.key === '[' || e.key === '<') return e.shiftKey ? ops.zOrder(doc, 'back') : a.growFont(-1), true;
        if (e.key === ' ') return a.clearFormatting(), true;
        return false;
      },
      onStartShow: (fromCurrent) => setShow({ start: fromCurrent ? doc.sel.slide : 0, presenter: false }),
      onPainterApply: (ids) => {
        const p = ui.painter;
        ui.painter = null;
        if (!p) return;
        doc.selectEls(ids);
        if (p.shape && Object.keys(p.shape).length) ops.patchEls(doc, 'Format painter', (e) => (e.type === 'shape' ? p.shape! : e.type === 'image' ? { line: p.shape!.line, shadow: p.shape!.shadow } : null) as Partial<El>);
        if (p.run) ops.applyRunStyle(doc, { font: p.run.font, size: p.run.size, color: p.run.color, b: p.run.b, i: p.run.i, u: p.run.u }, 'Format painter');
        ui.emit();
      },
    }),
    [doc, ui, a, insertFiles],
  );

  // shortcuts while editing text (TipTap keeps its own for bold/italic/underline)
  const onKeyDownCapture = (e: React.KeyboardEvent) => {
    if (e.key === 'F5') {
      e.preventDefault();
      setShow({ start: e.shiftKey ? doc.sel.slide : 0, presenter: false });
      return;
    }
    if (!activeEditor(ui) || !isMod(e)) return;
    const k = e.key.toLowerCase();
    const map: Record<string, () => void> = {
      e: () => a.align('center'),
      l: () => a.align('left'),
      r: () => a.align('right'),
      j: () => a.align('justify'),
      k: a.insertLink,
      ']': () => a.growFont(1),
      '[': () => a.growFont(-1),
      m: () => ops.addSlide(doc),
    };
    const fn = map[k];
    if (fn) {
      e.preventDefault();
      e.stopPropagation();
      fn();
    }
  };

  /* --------------------------------------------------------- controller */

  const serialize = useCallback((formatId: string) => exportPresentation(doc.pres, formatId, title()), [doc]); // eslint-disable-line react-hooks/exhaustive-deps

  const exportTarget = useCallback(
    async (id: string) => {
      if (id === 'pdf') {
        const { buildSlidesPdfJob } = await import('./formats/raster');
        return exportPdf(title(), await buildSlidesPdfJob(doc.pres, title()));
      }
      if (id === 'png' || id === 'png-all') {
        const { slidePng, slidesZip } = await import('./formats/raster');
        const target = SLIDES_EXPORT_TARGETS.find((t) => t.id === id)!;
        await exportBytes(title(), target, () => (id === 'png' ? slidePng(doc.pres, doc.sel.slide) : slidesZip(doc.pres, title())));
        return;
      }
      const target = SLIDES_EXPORT_TARGETS.find((t) => t.id === id) ?? SLIDES_SAVE_TARGETS.find((t) => t.id === id);
      if (target) await exportBytes(title(), target, () => serialize(id));
    },
    [doc, serialize], // eslint-disable-line react-hooks/exhaustive-deps
  );

  const commands = useCallback((): Command[] => {
    const cmd = (id: string, t: string, category: string, run: () => void, keys?: string): Command => ({ id: `slides-${id}`, title: t, category, run, keys });
    return [
      cmd('new', 'New slide', 'Slides', () => ops.addSlide(doc), 'Ctrl+M'),
      cmd('dup', 'Duplicate slide', 'Slides', () => ops.duplicateSlides(doc)),
      cmd('del', 'Delete slide', 'Slides', () => ops.deleteSlides(doc)),
      cmd('hide', 'Hide / show slide', 'Slides', () => ops.toggleHidden(doc)),
      cmd('show', 'Start slide show', 'Slide Show', () => setShow({ start: 0, presenter: false }), 'F5'),
      cmd('show-cur', 'Start from current slide', 'Slide Show', () => setShow({ start: doc.sel.slide, presenter: false }), 'Shift+F5'),
      cmd('presenter', 'Presenter view', 'Slide Show', () => setShow({ start: doc.sel.slide, presenter: true })),
      cmd('textbox', 'Insert text box', 'Insert', a.insertTextBox),
      cmd('picture', 'Insert picture', 'Insert', a.insertImage),
      cmd('icon', 'Insert icon', 'Insert', a.insertIcon),
      cmd('table', 'Insert table', 'Insert', () => a.insertTable(3, 3)),
      cmd('chart', 'Insert chart', 'Insert', () => a.insertChart('column')),
      cmd('bg', 'Format background…', 'Design', a.formatBackground),
      cmd('size', 'Slide size…', 'Design', () => a.slideSize('custom')),
      cmd('hf', 'Header & footer…', 'Insert', a.headerFooter),
      cmd('sorter', 'Slide sorter view', 'View', () => a.view(ui.view === 'sorter' ? 'normal' : 'sorter')),
      cmd('notes', 'Toggle notes', 'View', a.toggleNotes),
      cmd('find', 'Find', 'Edit', () => setFind('find'), 'Ctrl+F'),
      cmd('replace', 'Replace', 'Edit', () => setFind('replace'), 'Ctrl+H'),
      cmd('pdf', 'Export as PDF', 'File', () => void exportTarget('pdf')),
      cmd('print', 'Print', 'File', () => void printCurrent(), 'Ctrl+P'),
    ];
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [a, doc, ui, exportTarget]);

  const printCurrent = useCallback(async () => {
    const { buildSlidesPdfJob } = await import('./formats/raster');
    await printJob(await buildSlidesPdfJob(doc.pres, title()));
  }, [doc]); // eslint-disable-line react-hooks/exhaustive-deps

  useEffect(
    () =>
      registerController(tab.id, {
        save: () => saveDocument({ tabId: tab.id, kind: 'slides', targets: SLIDES_SAVE_TARGETS, defaultFormat: useSettings.getState().settings.defaultSlidesFormat, serialize }),
        saveAs: (formatId) => saveDocument({ tabId: tab.id, kind: 'slides', targets: SLIDES_SAVE_TARGETS, defaultFormat: useSettings.getState().settings.defaultSlidesFormat, serialize, saveAs: true, formatId }),
        exportAs: exportTarget,
        print: printCurrent,
        snapshot: () => JSON.stringify(doc.pres),
        focus: focusCanvas,
        commands,
        info: () => [
          { label: 'Slides', value: String(doc.pres.slides.length) },
          { label: 'Size', value: `${Math.round((doc.pres.size.w / 96) * 2.54 * 10) / 10} × ${Math.round((doc.pres.size.h / 96) * 2.54 * 10) / 10} cm` },
        ],
      }),
    [doc, tab.id, serialize, exportTarget, printCurrent, commands, focusCanvas],
  );

  useEffect(() => {
    if (active) focusCanvas();
  }, [active, focusCanvas]);

  // development builds expose the document to automated UI tests
  useEffect(() => {
    if (import.meta.env.DEV && active) (window as unknown as { __afficeSlides?: unknown }).__afficeSlides = { doc, ui, a };
  }, [active, doc, ui, a]);

  /* ------------------------------------------------------------- render */

  const panelKey = (e: React.KeyboardEvent) => {
    const k = e.key;
    if (k === 'Delete' || k === 'Backspace') {
      e.preventDefault();
      ops.deleteSlides(doc);
    } else if (k === 'ArrowDown' || k === 'ArrowUp') {
      e.preventDefault();
      doc.goToSlide(doc.sel.slide + (k === 'ArrowDown' ? 1 : -1));
    } else if (k === 'Enter') {
      e.preventDefault();
      ops.addSlide(doc);
    } else if (isMod(e) && k.toLowerCase() === 'd') {
      e.preventDefault();
      ops.duplicateSlides(doc);
    } else if (isMod(e) && k.toLowerCase() === 'z') {
      e.preventDefault();
      a.undo();
    } else if (isMod(e) && k.toLowerCase() === 'y') {
      e.preventDefault();
      a.redo();
    }
  };

  const startPanelResize = (e: React.PointerEvent) => {
    const x0 = e.clientX;
    const w0 = panelW;
    const move = (ev: PointerEvent) => setPanelW(Math.max(130, Math.min(420, w0 + ev.clientX - x0)));
    const up = () => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
  };

  const hiddenCount = pres.slides.filter((s) => s.hidden).length;

  return (
    <div className="editor slides-editor" data-tab={tab.id} ref={rootRef} onKeyDownCapture={onKeyDownCapture}>
      <SlidesRibbon doc={doc} ui={ui} a={a} onFile={() => setBackstage(true)} />
      <div className="sl-body">
        {ui.view === 'normal' ? (
          <>
            {ui.panel && (
              <>
                <SlidePanel doc={doc} width={panelW} onContextMenu={slideMenu} onAdd={() => ops.addSlide(doc)} onKey={panelKey} />
                <div className="sl-panel-grip" onPointerDown={startPanelResize} />
              </>
            )}
            <div className="sl-main">
              <div className="sl-canvas-wrap">
                <SlideCanvas doc={doc} ui={ui} active={active && !show} h={canvasHandlers} />
                {find && <SlideFind doc={doc} mode={find} onClose={() => (setFind(null), focusCanvas())} />}
                {preview && <SlidePreview key={`${preview}-${doc.sel.slide}-${doc.version}`} pres={pres} index={doc.sel.slide} what={preview} onDone={() => setPreview(null)} />}
              </div>
              {ui.notes && <NotesPane doc={doc} height={notesH} onResize={setNotesH} />}
            </div>
            {animPane && <AnimationPane doc={doc} onClose={() => setAnimPane(false)} onPlay={() => setPreview('anims')} />}
          </>
        ) : (
          <SlideSorter doc={doc} onOpen={(i) => (doc.goToSlide(i), a.view('normal'))} onContextMenu={slideMenu} />
        )}
      </div>
      <footer className="statusbar sl-status">
        <span className="status-item">
          Slide {doc.sel.slide + 1} of {pres.slides.length}
        </span>
        {hiddenCount > 0 && <span className="status-item muted">{hiddenCount} hidden</span>}
        {ui.insert && <span className="status-item active">Click and drag on the slide to draw · Esc to cancel</span>}
        {ui.painter && <span className="status-item active">Format painter: click an object</span>}
        <span className="spacer" />
        <SaveState dirty={tab.dirty} path={tab.path} />
        <span className="status-sep" />
        <button className={`status-item${ui.notes ? ' active' : ''}`} data-tip="Notes" onClick={a.toggleNotes}>
          <StickyNote size={14} /> Notes
        </button>
        <button className={`status-item${ui.view === 'normal' ? ' active' : ''}`} data-tip="Normal view" onClick={() => a.view('normal')}>
          <NotebookPen size={14} />
        </button>
        <button className={`status-item${ui.view === 'sorter' ? ' active' : ''}`} data-tip="Slide sorter" onClick={() => a.view('sorter')}>
          <LayoutGrid size={14} />
        </button>
        <button className="status-item" data-tip="Slide show (F5)" onClick={() => setShow({ start: doc.sel.slide, presenter: false })}>
          <PresIcon size={14} />
        </button>
        <span className="status-sep" />
        <ZoomControl zoom={ui.effectiveZoom} onZoom={(z) => ui.setZoom(z)} min={0.1} max={4} presets={[0.33, 0.5, 0.66, 1, 1.5, 2]} extra={[{ label: 'Fit slide to window', onSelect: () => ui.setZoom('fit') }]} />
      </footer>
      {show && (
        <Slideshow
          pres={pres}
          start={show.start}
          presenter={show.presenter}
          onExit={(i) => {
            setShow(null);
            doc.goToSlide(i);
            focusCanvas();
          }}
        />
      )}
      {backstage && (
        <Backstage
          app="slides"
          tab={tab}
          onClose={() => (setBackstage(false), focusCanvas())}
          onSave={() => void saveDocument({ tabId: tab.id, kind: 'slides', targets: SLIDES_SAVE_TARGETS, defaultFormat: appSettings.defaultSlidesFormat, serialize })}
          onSaveAs={(fid) => void saveDocument({ tabId: tab.id, kind: 'slides', targets: SLIDES_SAVE_TARGETS, defaultFormat: appSettings.defaultSlidesFormat, serialize, saveAs: true, formatId: fid })}
          onExport={(id) => void exportTarget(id)}
          onPrint={() => void printCurrent()}
          saveTargets={SLIDES_SAVE_TARGETS}
          exportTargets={SLIDES_EXPORT_TARGETS}
          libreOffice={isElectron}
          info={[
            { label: 'Slides', value: String(pres.slides.length) },
            { label: 'Design', value: designById(pres.design)?.name ?? pres.theme.name },
            { label: 'Author', value: pres.props.author || appSettings.authorName || '—' },
          ]}
        />
      )}
    </div>
  );
}

