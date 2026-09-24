import './sheets.css';
import { AlertTriangle, Calculator, FileSpreadsheet, Grid3x3, Lock, Sigma } from 'lucide-react';
import { useCallback, useEffect, useMemo, useRef, useState, useSyncExternalStore } from 'react';
import { api, isElectron } from '@/lib/platform';
import { SHEET_EXPORT_TARGETS, SHEET_SAVE_TARGETS } from '@/lib/formats';
import { Backstage } from '@/app/Backstage';
import { exportBytes, exportPdf, markDirty, printJob, saveDocument } from '@/app/fileOps';
import { useSettings } from '@/app/settings';
import { SaveState, ZoomControl } from '@/app/StatusBits';
import { registerController, useWorkspace, type Command, type Tab } from '@/app/workspace';
import { alertDialog, confirmDialog, promptDialog } from '@/ui/dialog';
import { openContextMenu, type MenuItem } from '@/ui/menu';
import { toast } from '@/ui/toast';
import { isMod } from '@/lib/utils';
import { SheetDoc } from './doc';
import { cellKey, colName, keyCol, keyRow, MAX_COLS, MAX_ROWS, addrName, rangeName, type Range } from './model/address';
import type { CellStyle, ChartType, CondFormat } from './model/types';
import { workbookToJSON } from './model/workbook';
import { copySelection, paste, fillDirection, autoFill, moveBlock, flashFill, type PasteMode } from './ops/clipboard';
import { deleteCellsShift, deleteCols, deleteRows, insertCellsShift, insertCols, insertRows, selectedIndexes, setColWidth, setFreeze, setHidden, setRowHeight, addSheet, deleteSheet } from './ops/structure';
import { applyFilter, clearFilters, currentRegion, dataRange, deleteChart, deleteCF, deleteImage, looksLikeHeader, newId, setCellMeta, sortRange, toggleFilter, upsertChart, upsertCF, upsertImage } from './ops/data';
import { formatValue } from './format/numfmt';
import { Grid, type GridHandlers, type GridTarget } from './grid/Grid';
import { autoFitRows, autoFitWidth, autoFitHeight } from './grid/autofit';
import { SheetUI } from './ui/controller';
import { FormulaBar } from './ui/FormulaBar';
import { SheetTabs } from './ui/SheetTabs';
import { SheetsRibbon, type SheetActions } from './ui/SheetsRibbon';
import { FindPanel } from './ui/FindPanel';
import { FilterMenu, ListPicker } from './ui/FilterMenu';
import {
  chartDialog,
  conditionalRuleDialog,
  formatCellsDialog,
  goalSeekDialog,
  insertFunctionDialog,
  manageRulesDialog,
  nameManagerDialog,
  newChartSpec,
  pageSetupDialog,
  pasteSpecialDialog,
  protectSheetDialog,
  removeDuplicatesDialog,
  shiftCellsDialog,
  sortDialog,
  textToColumnsDialog,
  validationDialog,
  CF_PRESET_STYLES,
} from './ui/dialogs';
import { stepDecimals } from './ui/numberFormats';
import { CELL_STYLES, STYLE_KEYS, TABLE_STYLES } from './ui/tableStyles';
import { exportWorkbook, fillFromRows, loadWorkbook, parseDelimited, sniffDelimiter, decodeText } from './formats';
import { buildSheetPdfJob } from './formats/print';

export default function SheetsEditor({ tab, active }: { tab: Tab; active: boolean }) {
  const [state, setState] = useState<{ doc: SheetDoc; ui: SheetUI } | null>(null);
  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const { wb, warning } = await loadWorkbook(tab.source);
        if (cancelled) return;
        const doc = new SheetDoc(wb);
        for (const s of wb.sheets) if (s.filter) applyFilter(doc, s);
        setState({ doc, ui: new SheetUI(doc) });
        if (warning) toast.warning(warning, { duration: 8000 });
      } catch (e) {
        console.error(e);
        toast.error(`Couldn't open ${tab.title}`, { detail: String((e as Error)?.message ?? e) });
        const doc = new SheetDoc(new (await import('./model/workbook')).Workbook());
        if (!cancelled) setState({ doc, ui: new SheetUI(doc) });
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
  return <SheetsView tab={tab} active={active} doc={state.doc} ui={state.ui} />;
}

function SheetsView({ tab, active, doc, ui }: { tab: Tab; active: boolean; doc: SheetDoc; ui: SheetUI }) {
  const docVersion = useSyncExternalStore(doc.subscribe, doc.getVersion);
  useSyncExternalStore(ui.subscribe, ui.getVersion);
  const appSettings = useSettings((s) => s.settings);
  const [backstage, setBackstage] = useState(false);
  const [find, setFind] = useState<null | 'find' | 'replace'>(null);
  const [filterMenu, setFilterMenu] = useState<{ col: number; anchor: DOMRect } | null>(null);
  const [listPick, setListPick] = useState<DOMRect | null>(null);
  const [formulaBar, setFormulaBar] = useState(true);
  const [painter, setPainter] = useState<CellStyle | null>(null);
  const pendingPaste = useRef<PasteMode | null>(null);
  const sheet = doc.sheet;

  /* ------------------------------------------------------------ wiring */
  useEffect(() => {
    doc.onChange = (kind) => {
      if (kind === 'edit') markDirty(tab.id);
    };
    ui.onError = (title, detail) => void alertDialog(title, detail);
    ui.onNotify = (msg) => toast.info(msg, { duration: 2500 });
    ui.onEditChart = (id) => {
      const spec = doc.sheet.charts.find((c) => c.id === id);
      if (spec) void chartDialog(doc, spec);
    };
    return () => {
      doc.onChange = undefined;
    };
  }, [doc, ui, tab.id]);

  const focusGrid = useCallback(() => {
    requestAnimationFrame(() => document.querySelector<HTMLTextAreaElement>(`.sheets-editor[data-tab="${tab.id}"] .grid-sink`)?.focus({ preventScroll: true }));
  }, [tab.id]);

  /** Rows touched by the last change get their automatic height refreshed. */
  const refitRows = useCallback(
    (rows?: Iterable<number>) => {
      const s = doc.sheet;
      const target = rows ?? (() => {
        const out = new Set<number>();
        for (const rg of doc.sel.ranges) {
          const b = doc.clampToUsed(s, rg);
          for (let r = b.r1; r <= Math.min(b.r2, b.r1 + 500); r++) out.add(r);
        }
        return out;
      })();
      if (autoFitRows(doc, s, target)) {
        doc.version++;
        doc.emit('view');
      }
    },
    [doc],
  );

  const title = () => useWorkspace.getState().tabs.find((t) => t.id === tab.id)?.title ?? 'Workbook';

  /* ---------------------------------------------------------- clipboard */
  const doCopy = useCallback(
    (cut: boolean) => {
      const sink = document.querySelector<HTMLTextAreaElement>(`.sheets-editor[data-tab="${tab.id}"] .grid-sink`);
      sink?.focus({ preventScroll: true });
      sink?.select();
      if (isElectron) api.app.editCommand(cut ? 'cut' : 'copy');
      else if (!document.execCommand(cut ? 'cut' : 'copy')) {
        const res = copySelection(doc, cut);
        if (res) void navigator.clipboard?.writeText(res.text);
      }
    },
    [doc, tab.id],
  );

  const doPaste = useCallback(
    async (mode: PasteMode = 'all') => {
      const sink = document.querySelector<HTMLTextAreaElement>(`.sheets-editor[data-tab="${tab.id}"] .grid-sink`);
      sink?.focus({ preventScroll: true });
      if (isElectron) {
        pendingPaste.current = mode;
        api.app.editCommand('paste');
        return;
      }
      // browsers: read the clipboard directly
      let text: string | undefined;
      let html: string | undefined;
      try {
        const items = await navigator.clipboard.read();
        for (const it of items) {
          if (it.types.includes('text/html')) html = await (await it.getType('text/html')).text();
          if (it.types.includes('text/plain')) text = await (await it.getType('text/plain')).text();
        }
      } catch {
        try {
          text = await navigator.clipboard.readText();
        } catch {
          text = undefined;
        }
      }
      if (text === undefined && !html && doc.clip) text = doc.clip.text;
      const rg = paste(doc, { text, html }, mode);
      if (rg) refitRows(range(rg.r1, Math.min(rg.r2, rg.r1 + 500)));
    },
    [doc, tab.id, refitRows],
  );

  /* ------------------------------------------------------------ actions */
  const selRange = () => doc.sel.ranges[doc.sel.ranges.length - 1];

  const a: SheetActions = useMemo(() => {
    const style = (patch: Partial<CellStyle>, label?: string) => {
      doc.applyStyle(patch, label);
      if ('size' in patch || 'wrap' in patch || 'font' in patch || 'rotation' in patch) refitRows();
    };
    const cf = (kind: string) => {
      const rg = doc.sel.ranges.map((r) => doc.clampToUsed(doc.sheet, r, 100));
      const base = { id: newId('cf'), ranges: rg };
      if (kind.startsWith('bar:')) return upsertCF(doc, { ...base, type: 'dataBar', colors: [kind.slice(4)] });
      if (kind.startsWith('scale:')) return upsertCF(doc, { ...base, type: 'colorScale', colors: kind.slice(6).split(',') });
      if (kind.startsWith('icons:')) return upsertCF(doc, { ...base, type: 'iconSet', iconSet: kind.slice(6) as CondFormat['iconSet'] });
      const style = CF_PRESET_STYLES[0].style;
      switch (kind) {
        case 'greater':
        case 'less':
        case 'between':
        case 'equal':
          return void conditionalRuleDialog(doc, undefined, { type: 'cell', op: kind, style, values: [''] });
        case 'text':
          return void conditionalRuleDialog(doc, undefined, { type: 'text', text: '', style });
        case 'date':
          return void conditionalRuleDialog(doc, undefined, { type: 'date', datePeriod: 'today', style });
        case 'duplicate':
        case 'unique':
          return upsertCF(doc, { ...base, type: kind, style });
        case 'top10':
          return upsertCF(doc, { ...base, type: 'top', rank: 10, style: CF_PRESET_STYLES[2].style });
        case 'top10p':
          return upsertCF(doc, { ...base, type: 'top', rank: 10, percent: true, style: CF_PRESET_STYLES[2].style });
        case 'bottom10':
          return upsertCF(doc, { ...base, type: 'top', rank: 10, bottom: true, style });
        case 'above':
          return upsertCF(doc, { ...base, type: 'average', above: true, style: CF_PRESET_STYLES[2].style });
        case 'below':
          return upsertCF(doc, { ...base, type: 'average', above: false, style });
      }
    };
    const insertChart = (type: ChartType) => {
      const spec = newChartSpec(doc, type);
      const rg = dataRange(doc);
      if (rg.r1 === rg.r2 && rg.c1 === rg.c2 && doc.value(doc.sheet, rg.r1, rg.c1) === undefined) {
        toast.info('Select the data you want to chart first (for example A1:C10).');
        return;
      }
      upsertChart(doc, spec);
      ui.selectedObject = { kind: 'chart', id: spec.id };
      ui.emit();
    };
    return {
      undo: () => {
        const l = doc.undo();
        if (!l) toast.info('Nothing to undo');
      },
      redo: () => {
        const l = doc.redo();
        if (!l) toast.info('Nothing to redo');
      },
      paste: (mode) => void doPaste(mode ?? 'all'),
      pasteSpecial: () =>
        void pasteSpecialDialog().then((m) => {
          if (m) void doPaste(m as PasteMode);
        }),
      cut: () => doCopy(true),
      copy: () => doCopy(false),
      painter: () => {
        if (painter) setPainter(null);
        else {
          setPainter(doc.activeStyle());
          toast.info('Select the cells to apply the formatting to.', { duration: 2500 });
        }
      },
      painterOn: !!painter,
      style,
      toggle: (key) => style({ [key]: doc.activeStyle()[key] ? undefined : true } as Partial<CellStyle>, key === 'wrap' ? 'Wrap text' : 'Format'),
      underline: () => style({ underline: doc.activeStyle().underline ? undefined : 'single' }, 'Underline'),
      borders: (kind, st, color) => doc.applyBorders(kind, st, color),
      merge: (mode) => {
        const msg = doc.merge(mode);
        if (msg) toast.info(msg);
      },
      numFmt: (fmt) => style({ numFmt: fmt }, 'Number format'),
      decimals: (dir) => {
        const v = doc.value(doc.sheet, doc.sel.active.r, doc.sel.active.c);
        style({ numFmt: stepDecimals(doc.activeStyle().numFmt, dir, typeof v === 'number' ? v : undefined) }, 'Decimals');
      },
      indent: (dir) => style({ indent: Math.max(0, (doc.activeStyle().indent ?? 0) + dir) || undefined, hAlign: doc.activeStyle().hAlign ?? 'left' }, 'Indent'),
      formatCells: (t) => void formatCellsDialog(doc, t).then(() => refitRows()),
      cfPreset: (kind) => void cf(kind),
      cfNew: () => void conditionalRuleDialog(doc),
      cfManage: () => void manageRulesDialog(doc),
      cfClear: (scope) => deleteCF(doc, scope),
      tableStyle: (id) => {
        const t = TABLE_STYLES.find((x) => x.id === id) ?? TABLE_STYLES[0];
        const rg = dataRange(doc);
        if (rg.r1 === rg.r2) {
          toast.info('Select a block of data with a header row first.');
          return;
        }
        const s = doc.sheet;
        // the header is formatted directly; banding is a formula rule so it stays in place when rows are sorted
        const body: CellStyle = { ...t.band2 };
        if (body.fill === '#ffffff') delete body.fill;
        const inside = (x: Range) => x.r1 >= rg.r1 && x.r2 <= rg.r2 && x.c1 >= rg.c1 && x.c2 <= rg.c2;
        doc.transact('Format as table', (tx) => {
          for (let r = rg.r1; r <= rg.r2; r++) {
            const st = r === rg.r1 ? t.header : body;
            for (let c = rg.c1; c <= rg.c2; c++) {
              const k = cellKey(r, c);
              const cur = s.cells.get(k);
              tx.patch(s, k, { s: doc.wb.patchStyle(cur?.s ?? doc.wb.cellStyleId(s, r, c), { fill: undefined, color: undefined, bb: undefined, ...st }) });
            }
          }
          tx.prop(s, 'cf');
          s.cf = s.cf.filter((x) => !(x.id.startsWith('band') && x.ranges.every(inside)));
          if (rg.r2 > rg.r1 && (t.band1.fill || t.band1.color)) {
            const first = `$${colName(rg.c1)}$${rg.r1 + 2}`;
            s.cf.push({ id: newId('band'), ranges: [{ r1: rg.r1 + 1, c1: rg.c1, r2: rg.r2, c2: rg.c2 }], type: 'formula', values: [`=MOD(ROW()-ROW(${first}),2)=0`], style: { fill: t.band1.fill, color: t.band1.color } });
          }
          tx.prop(s, 'filter');
          s.filter = { range: rg, columns: {} };
        });
        applyFilter(doc, s);
      },
      cellStyle: (id) => {
        const cs = CELL_STYLES.find((c) => c.id === id);
        if (!cs) return;
        const patch: Partial<CellStyle> = {};
        for (const k of STYLE_KEYS) (patch as Record<string, unknown>)[k] = cs.style[k];
        if (cs.style.numFmt) patch.numFmt = cs.style.numFmt;
        style(patch, `${cs.label} style`);
      },
      insert: (what) => {
        const rg = selRange();
        if (what === 'rows') insertRows(doc, rg.r1, rg.r2 - rg.r1 + 1);
        else if (what === 'cols') insertCols(doc, rg.c1, rg.c2 - rg.c1 + 1);
        else if (what === 'sheet') addSheet(doc);
        else
          void shiftCellsDialog('insert').then((m) => {
            if (m === 'row') insertRows(doc, rg.r1, rg.r2 - rg.r1 + 1);
            else if (m === 'col') insertCols(doc, rg.c1, rg.c2 - rg.c1 + 1);
            else if (m === 'down' || m === 'right') insertCellsShift(doc, doc.clampToUsed(doc.sheet, rg), m);
          });
      },
      remove: (what) => {
        const rg = selRange();
        if (what === 'rows') deleteRows(doc, rg.r1, Math.min(rg.r2, MAX_ROWS - 1) - rg.r1 + 1);
        else if (what === 'cols') deleteCols(doc, rg.c1, Math.min(rg.c2, MAX_COLS - 1) - rg.c1 + 1);
        else if (what === 'sheet') {
          void confirmDialog({ title: `Delete “${doc.sheet.name}”?`, message: 'You can undo this with Ctrl+Z.', confirmLabel: 'Delete', danger: true }).then((ok) => {
            if (ok && !deleteSheet(doc, doc.wb.activeSheet)) toast.info('A workbook needs at least one visible sheet.');
          });
        } else
          void shiftCellsDialog('delete').then((m) => {
            if (m === 'row') deleteRows(doc, rg.r1, rg.r2 - rg.r1 + 1);
            else if (m === 'col') deleteCols(doc, rg.c1, rg.c2 - rg.c1 + 1);
            else if (m === 'up' || m === 'left') deleteCellsShift(doc, doc.clampToUsed(doc.sheet, rg), m);
          });
      },
      rowHeight: () =>
        void promptDialog({ title: 'Row height', label: 'Height (pixels)', value: String(doc.sheet.rowHeight(doc.sel.active.r)), validate: (v) => (Number(v) >= 0 && Number(v) <= 600 ? null : 'Enter 0–600') }).then((v) => {
          if (v !== null) setRowHeight(doc, selectedIndexes(doc, 'row'), Number(v));
        }),
      colWidth: () =>
        void promptDialog({ title: 'Column width', label: 'Width (pixels)', value: String(doc.sheet.colWidth(doc.sel.active.c)), validate: (v) => (Number(v) >= 0 && Number(v) <= 2000 ? null : 'Enter 0–2000') }).then((v) => {
          if (v !== null) setColWidth(doc, selectedIndexes(doc, 'col'), Number(v));
        }),
      autofit: (axis) => h.onAutoFit(axis, selectedIndexes(doc, axis)),
      hide: (axis, hidden) => {
        let idx = selectedIndexes(doc, axis);
        if (!hidden) {
          // unhide includes hidden neighbours of a single selected row/column
          const lo = Math.max(0, idx[0] - 1);
          const hi = idx[idx.length - 1] + 1;
          idx = range(lo, hi);
        }
        setHidden(doc, axis, idx, hidden);
      },
      renameSheet: () =>
        void promptDialog({ title: 'Rename sheet', value: doc.sheet.name }).then(async (v) => {
          if (v === null) return;
          const { renameSheet } = await import('./ops/structure');
          const err = renameSheet(doc, doc.wb.activeSheet, v);
          if (err) toast.error(err);
        }),
      protectSheet: () =>
        void protectSheetDialog(doc, doc.wb.activeSheet).then((m) => {
          if (m) toast.info(m);
        }),
      autoSum: (fn) => {
        const rg = selRange();
        const s = doc.sheet;
        const num = (r: number, c: number) => typeof doc.value(s, r, c) === 'number';
        if (rg.r1 === rg.r2 && rg.c1 === rg.c2) {
          const { r, c } = doc.sel.active;
          let r0 = r - 1;
          while (r0 >= 0 && num(r0, c)) r0--;
          let text: string;
          if (r0 < r - 1) text = `=${fn}(${addrName(r0 + 1, c)}:${addrName(r - 1, c)})`;
          else {
            let c0 = c - 1;
            while (c0 >= 0 && num(r, c0)) c0--;
            text = c0 < c - 1 ? `=${fn}(${addrName(r, c0 + 1)}:${addrName(r, c - 1)})` : `=${fn}()`;
          }
          ui.startEdit(text, 'edit');
          return;
        }
        // totals below each selected column
        const b = doc.clampToUsed(s, rg);
        doc.transact('AutoSum', (tx) => {
          for (let c = b.c1; c <= b.c2; c++) tx.set(s, cellKey(b.r2 + 1, c), { ...(s.get(b.r2 + 1, c) ?? {}), f: `${fn}(${addrName(b.r1, c)}:${addrName(b.r2, c)})`, v: undefined });
        });
      },
      fill: (dir) => {
        if (dir === 'series') return;
        fillDirection(doc, dir);
      },
      clear: (what) => doc.clear(what),
      sort: (desc) => {
        const rg = dataRange(doc);
        const header = looksLikeHeader(doc, doc.sheet, rg);
        sortRange(doc, rg, [{ col: doc.sel.active.c, desc }], header);
      },
      customSort: () => void sortDialog(doc),
      filter: () => toggleFilter(doc),
      clearFilter: () => clearFilters(doc),
      reapplyFilter: () => {
        applyFilter(doc, doc.sheet);
        doc.version++;
        doc.emit('view');
      },
      find: (m) => setFind(m),
      goTo: () =>
        void promptDialog({ title: 'Go to', label: 'Reference (e.g. B12, A1:D20, Sheet2!C5 or a name)', value: '' }).then((v) => {
          if (v && !ui.goTo(v)) toast.error('That reference isn’t valid.');
        }),
      selectSpecial: (kind) => {
        const s = doc.sheet;
        const ext = doc.engine.extent(s.id);
        const hits: Range[] = [];
        const push = (r: number, c: number) => {
          if (hits.length < 5000) hits.push({ r1: r, c1: c, r2: r, c2: c });
        };
        if (kind === 'blanks') {
          const rg = doc.clampToUsed(s, selRange().r1 === selRange().r2 && selRange().c1 === selRange().c2 ? { r1: 0, c1: 0, r2: ext.rows, c2: ext.cols } : selRange());
          for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) if (doc.value(s, r, c) === undefined || doc.value(s, r, c) === null) push(r, c);
        } else if (kind === 'cf' || kind === 'validation') {
          for (const x of kind === 'cf' ? s.cf : s.validations) hits.push(...x.ranges);
        } else {
          for (const [k, cell] of s.cells) {
            const ok = kind === 'formulas' ? cell.f !== undefined : kind === 'constants' ? cell.f === undefined && cell.v !== undefined : !!cell.note;
            if (ok) push(keyRow(k), keyCol(k));
          }
        }
        if (!hits.length) {
          toast.info('No cells were found.');
          return;
        }
        doc.sel = { ranges: hits, active: { r: hits[0].r1, c: hits[0].c1 }, anchor: { r: hits[0].r1, c: hits[0].c1 } };
        ui.ensureVisible(hits[0].r1, hits[0].c1);
        ui.emit();
      },
      insertChart,
      insertImage: async () => {
        const f = await api.files.pickImage();
        if (!f) return;
        const ext = f.name.split('.').pop()?.toLowerCase() ?? 'png';
        const mime = ext === 'svg' ? 'image/svg+xml' : `image/${ext === 'jpg' ? 'jpeg' : ext}`;
        const blob = new Blob([f.data as BlobPart], { type: mime });
        const src = await new Promise<string>((res) => {
          const r = new FileReader();
          r.onload = () => res(String(r.result));
          r.readAsDataURL(blob);
        });
        const img = new Image();
        img.src = src;
        await img.decode().catch(() => undefined);
        const scale = Math.min(1, 480 / Math.max(1, img.naturalWidth), 360 / Math.max(1, img.naturalHeight));
        const { r, c } = doc.sel.active;
        upsertImage(doc, { id: newId('img'), src, anchor: { r, c, dx: 0, dy: 0 }, w: Math.round((img.naturalWidth || 320) * scale), h: Math.round((img.naturalHeight || 240) * scale), alt: f.name });
      },
      insertLink: () => {
        const { r, c } = doc.sel.active;
        const cur = doc.sheet.get(r, c);
        void promptDialog({ title: 'Insert link', label: 'Address (web page, email or #Sheet2!A1)', value: cur?.link ?? 'https://' }).then((url) => {
          if (url === null) return;
          let link = url.trim();
          if (link && !/^[a-z]+:|^#/i.test(link)) link = /@/.test(link) ? `mailto:${link}` : `https://${link}`;
          doc.transact('Link', (tx) => tx.patch(doc.sheet, cellKey(r, c), { link: link || undefined, v: cur?.v ?? (cur?.f === undefined ? url.trim() : undefined) }));
        });
      },
      insertNote: () => {
        const { r, c } = doc.sel.active;
        const cur = doc.sheet.get(r, c);
        void promptDialog({ title: cur?.note ? 'Edit note' : 'New note', label: `Note for ${addrName(r, c)}`, value: cur?.note ?? '', multiline: true }).then((v) => {
          if (v !== null) setCellMeta(doc, r, c, { note: v.trim() || undefined });
        });
      },
      insertFunction: (cat) =>
        void insertFunctionDialog(cat).then((name) => {
          if (!name) return;
          const e = ui.edit;
          if (e) ui.updateEdit(e.text.slice(0, e.caret) + name + '(' + e.text.slice(e.caret), e.caret + name.length + 1);
          else ui.startEdit(`=${name}(`, 'edit');
        }),
      insertDateTime: (kind) => {
        const d = new Date();
        const text = kind === 'date' ? `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}` : `${d.getHours()}:${String(d.getMinutes()).padStart(2, '0')}`;
        doc.enter(text);
      },
      insertCheckbox: () => {
        const s = doc.sheet;
        const rgs = doc.sel.ranges.map((r) => doc.clampToUsed(s, r, 50));
        doc.transact('Checkboxes', (tx) => {
          tx.prop(s, 'validations');
          s.validations.push({ id: newId('dv'), ranges: rgs, type: 'checkbox', values: [], allowBlank: true });
          for (const rg of rgs)
            for (let r = rg.r1; r <= rg.r2; r++)
              for (let c = rg.c1; c <= rg.c2; c++) {
                const cur = s.get(r, c);
                if (cur?.v === undefined && cur?.f === undefined) tx.patch(s, cellKey(r, c), { v: false });
              }
        });
      },
      nameManager: () => void nameManagerDialog(doc),
      showFormulas: () => {
        ui.showFormulas = !ui.showFormulas;
        ui.emit();
      },
      calcNow: () => {
        doc.engine.recalcAll();
        doc.emit('view');
        toast.success(`Recalculated ${doc.engine.stats.formulas.toLocaleString()} formulas in ${Math.round(doc.engine.stats.lastRecalcMs)} ms`, { duration: 2000 });
      },
      setCalcAuto: (auto) => {
        doc.wb.calcAuto = auto;
        if (auto) doc.engine.recalcAll();
        doc.emit('edit');
      },
      trace: (kind) => {
        const { r, c } = doc.sel.active;
        const s = doc.sheet;
        if (kind === 'clear') {
          ui.findHits = null;
          ui.emit();
          return;
        }
        const ranges: Range[] = [];
        if (kind === 'precedents') {
          const p = doc.engine.precedentsOf(s.id, r, c);
          for (const x of p.cells) if (x.sheet === s.id) ranges.push({ r1: x.r, c1: x.c, r2: x.r, c2: x.c });
          for (const x of p.ranges) if (x.sheet === s.id) ranges.push({ r1: x.r1, c1: x.c1, r2: Math.min(x.r2, doc.engine.extent(s.id).rows), c2: x.c2 });
        } else for (const x of doc.engine.dependentsOf(s.id, r, c)) if (x.sheet === s.id) ranges.push({ r1: x.r, c1: x.c, r2: x.r, c2: x.c });
        if (!ranges.length) {
          toast.info(kind === 'precedents' ? 'This cell doesn’t refer to other cells on this sheet.' : 'No formulas on this sheet use this cell.');
          return;
        }
        const keys = new Set<number>();
        for (const rg of ranges) for (let rr = rg.r1; rr <= Math.min(rg.r2, rg.r1 + 2000); rr++) for (let cc = rg.c1; cc <= rg.c2; cc++) keys.add(cellKey(rr, cc));
        ui.findHits = keys;
        ui.emit();
        toast.info(`${kind === 'precedents' ? 'Precedents' : 'Dependents'} of ${addrName(r, c)}: ${ranges.map(rangeName).slice(0, 6).join(', ')}${ranges.length > 6 ? '…' : ''}`, { duration: 5000 });
      },
      textToColumns: () => void textToColumnsDialog(doc),
      flashFill: () => {
        const n = flashFill(doc);
        toast.info(n ? `Flash Fill filled ${n} ${n === 1 ? 'cell' : 'cells'}.` : 'Type one or two examples of the result next to your data, then try Flash Fill again.');
      },
      removeDuplicates: () =>
        void removeDuplicatesDialog(doc).then((msg) => {
          if (msg) toast.info(msg);
        }),
      validation: () => void validationDialog(doc),
      goalSeek: () =>
        void goalSeekDialog(doc).then((msg) => {
          if (msg) toast.info(msg, { duration: 5000 });
        }),
      importData: async () => {
        const files = await api.files.openDialog({ title: 'Import data', filters: [{ name: 'Text data', extensions: ['csv', 'tsv', 'txt'] }] });
        const f = files[0];
        if (!f) return;
        const text = decodeText(f.data);
        const rows = parseDelimited(text, f.name.toLowerCase().endsWith('.tsv') ? '\t' : sniffDelimiter(text));
        const { r, c } = doc.sel.active;
        doc.transact('Import data', (tx) => {
          tx.prop(doc.sheet, 'cells');
          tx.prop(doc.sheet, 'cols');
          fillFromRows(doc.wb, doc.sheet, rows, r, c);
        });
        toast.success(`Imported ${rows.length.toLocaleString()} rows`);
      },
      freeze: (kind) => {
        const { r, c } = doc.sel.active;
        if (kind === 'panes') setFreeze(doc, r, c);
        else if (kind === 'row') setFreeze(doc, 1, 0);
        else if (kind === 'col') setFreeze(doc, 0, 1);
        else setFreeze(doc, 0, 0);
      },
      zoom: (z) => {
        if (z === 'selection') {
          const rg = doc.clampToUsed(doc.sheet, selRange());
          const g = ui.geo();
          const w = g.cols.pos(rg.c2 + 1) - g.cols.pos(rg.c1);
          const hh = g.rows.pos(rg.r2 + 1) - g.rows.pos(rg.r1);
          const vz = Math.min((ui.viewport.width - ui.viewport.headerW - 20) / Math.max(1, w), (ui.viewport.height - ui.viewport.headerH - 20) / Math.max(1, hh));
          ui.setZoom(Math.max(0.25, Math.min(4, vz)));
          ui.ensureVisible(rg.r1, rg.c1);
        } else ui.setZoom(z);
      },
      toggleView: (key) => {
        doc.sheet.view[key] = !doc.sheet.view[key];
        doc.emit('edit');
      },
      toggleFormulaBar: () => setFormulaBar((x) => !x),
      formulaBarOn: formulaBar,
      pageSetup: () => void pageSetupDialog(doc),
      print: async () => printJob(await buildSheetPdfJob(doc, 'sheet', title())),
      exportPdf: async () => exportPdf(title(), await buildSheetPdfJob(doc, 'sheet', title())),
      chart: (action, type) => {
        const id = ui.selectedObject?.kind === 'chart' ? ui.selectedObject.id : null;
        const spec = id ? doc.sheet.charts.find((c) => c.id === id) : undefined;
        if (!spec) return;
        if (action === 'delete') {
          deleteChart(doc, spec.id);
          ui.selectedObject = null;
          ui.emit();
        } else if (action === 'type' && type) upsertChart(doc, { ...spec, type });
        else void chartDialog(doc, spec);
      },
      stats: () => {
        const cells = doc.wb.sheets.reduce((n, s) => n + s.cells.size, 0);
        void alertDialog(
          'Workbook statistics',
          `${doc.wb.sheets.length} sheets · ${cells.toLocaleString()} non-empty cells · ${doc.engine.stats.formulas.toLocaleString()} formulas · ${doc.wb.sheets.reduce((n, s) => n + s.charts.length, 0)} charts · ${doc.wb.names.length} names · last full recalculation ${Math.round(doc.engine.stats.lastRecalcMs)} ms`,
        );
      },
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [doc, ui, doCopy, doPaste, painter, formulaBar, refitRows]);

  // development builds expose the active workbook to automated UI tests
  useEffect(() => {
    if (import.meta.env.DEV && active) (window as unknown as { __afficeSheet?: unknown }).__afficeSheet = { doc, ui, a };
  }, [active, doc, ui, a]);

  /* ------------------------------------------------------ grid handlers */
  const buildMenu = useCallback(
    (target: GridTarget): MenuItem[] => {
      const clip: MenuItem[] = [
        { label: 'Cut', shortcut: 'Ctrl+X', onSelect: () => a.cut() },
        { label: 'Copy', shortcut: 'Ctrl+C', onSelect: () => a.copy() },
        { label: 'Paste', shortcut: 'Ctrl+V', onSelect: () => a.paste() },
        {
          label: 'Paste special',
          submenu: [
            { label: 'Values', shortcut: 'Ctrl+Shift+V', onSelect: () => a.paste('values') },
            { label: 'Formulas', onSelect: () => a.paste('formulas') },
            { label: 'Formatting', onSelect: () => a.paste('formats') },
            { label: 'Transpose', onSelect: () => a.paste('transpose') },
            { label: 'Paste link', onSelect: () => a.paste('link') },
            { separator: true },
            { label: 'Paste special…', onSelect: a.pasteSpecial },
          ],
        },
        { separator: true },
      ];
      if (target.kind === 'row' || target.kind === 'col') {
        const axis = target.kind;
        const n = selectedIndexes(doc, axis).length;
        const noun = axis === 'row' ? (n === 1 ? 'row' : `${n} rows`) : n === 1 ? 'column' : `${n} columns`;
        return [
          ...clip,
          { label: `Insert ${noun} ${axis === 'row' ? 'above' : 'left'}`, onSelect: () => a.insert(axis === 'row' ? 'rows' : 'cols') },
          { label: `Delete ${noun}`, onSelect: () => a.remove(axis === 'row' ? 'rows' : 'cols') },
          { label: 'Clear contents', onSelect: () => a.clear('contents') },
          { separator: true },
          { label: 'Format cells…', shortcut: 'Ctrl+1', onSelect: () => a.formatCells() },
          axis === 'row' ? { label: 'Row height…', onSelect: a.rowHeight } : { label: 'Column width…', onSelect: a.colWidth },
          { label: axis === 'row' ? 'Fit row height' : 'Fit column width', onSelect: () => a.autofit(axis) },
          { label: 'Hide', shortcut: axis === 'row' ? 'Ctrl+9' : 'Ctrl+0', onSelect: () => a.hide(axis, true) },
          { label: 'Unhide', onSelect: () => a.hide(axis, false) },
          ...(axis === 'col' ? [{ separator: true } as MenuItem, { label: 'Sort sheet A → Z', onSelect: () => a.sort(false) }, { label: 'Sort sheet Z → A', onSelect: () => a.sort(true) }] : []),
        ];
      }
      const { r, c } = target.kind === 'cell' ? target : doc.sel.active;
      const cell = doc.sheet.get(r, c);
      const hasList = doc.sheet.validations.some((v) => v.type === 'list' && v.ranges.some((rg) => r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2));
      const items: MenuItem[] = [
        ...clip,
        {
          label: 'Insert',
          submenu: [
            { label: 'Row above', onSelect: () => insertRows(doc, doc.sel.active.r, 1) },
            { label: 'Row below', onSelect: () => insertRows(doc, doc.sel.active.r + 1, 1) },
            { label: 'Column left', onSelect: () => insertCols(doc, doc.sel.active.c, 1) },
            { label: 'Column right', onSelect: () => insertCols(doc, doc.sel.active.c + 1, 1) },
            { label: 'Cells…', onSelect: () => a.insert('cells') },
          ],
        },
        {
          label: 'Delete',
          submenu: [
            { label: 'Row', onSelect: () => a.remove('rows') },
            { label: 'Column', onSelect: () => a.remove('cols') },
            { label: 'Cells…', onSelect: () => a.remove('cells') },
          ],
        },
        { label: 'Clear contents', shortcut: 'Delete', onSelect: () => a.clear('contents') },
        { separator: true },
        {
          label: 'Sort',
          submenu: [
            { label: 'Sort A → Z', onSelect: () => a.sort(false) },
            { label: 'Sort Z → A', onSelect: () => a.sort(true) },
            { label: 'Custom sort…', onSelect: a.customSort },
          ],
        },
        {
          label: 'Filter',
          submenu: [
            { label: doc.sheet.filter ? 'Remove filter' : 'Turn on filter', onSelect: a.filter },
            { label: 'Clear filter', disabled: !doc.sheet.filter, onSelect: a.clearFilter },
          ],
        },
        { separator: true },
        { label: cell?.note ? 'Edit note' : 'New note', shortcut: 'Shift+F2', onSelect: a.insertNote },
        ...(cell?.note ? [{ label: 'Delete note', onSelect: () => setCellMeta(doc, r, c, { note: undefined }) } as MenuItem] : []),
        { label: cell?.link ? 'Edit link…' : 'Link…', shortcut: 'Ctrl+K', onSelect: a.insertLink },
        ...(cell?.link ? [{ label: 'Open link', onSelect: () => openLink(cell.link!) } as MenuItem, { label: 'Remove link', onSelect: () => setCellMeta(doc, r, c, { link: undefined }) } as MenuItem] : []),
        ...(hasList ? [{ label: 'Pick from list…', onSelect: () => setListPick(new DOMRect(window.innerWidth / 2, window.innerHeight / 2, 120, 24)) } as MenuItem] : []),
        { separator: true },
        { label: 'Format cells…', shortcut: 'Ctrl+1', onSelect: () => a.formatCells() },
        { label: 'Conditional formatting…', onSelect: a.cfNew },
        { label: 'Data validation…', onSelect: a.validation },
        { label: 'Define name…', onSelect: a.nameManager },
      ];
      return items;
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [a, doc],
  );

  const openLink = useCallback(
    (url: string) => {
      if (url.startsWith('#')) {
        if (!ui.goTo(url.slice(1))) toast.error('The link points to a place that doesn’t exist.');
      } else api.app.openExternal(url);
    },
    [ui],
  );

  const shortcut = useCallback(
    (e: KeyboardEvent | React.KeyboardEvent): boolean => {
      const mod = isMod(e);
      const k = e.key;
      const low = k.toLowerCase();
      const shift = e.shiftKey;
      if (ui.selectedObject && (k === 'Delete' || k === 'Backspace')) {
        const o = ui.selectedObject;
        if (o.kind === 'chart') deleteChart(doc, o.id);
        else deleteImage(doc, o.id);
        ui.selectedObject = null;
        ui.emit();
        return true;
      }
      if (k === 'Escape' && ui.selectedObject) {
        ui.selectedObject = null;
        ui.emit();
        return true;
      }
      if (mod && !e.altKey) {
        if (low === 'z' && !shift) return a.undo(), true;
        if ((low === 'y' && !shift) || (low === 'z' && shift)) return a.redo(), true;
        if (low === 'v' && shift) {
          pendingPaste.current = 'values';
          if (!isElectron) {
            void doPaste('values');
            return true;
          }
          return false;
        }
        if (low === 'v' && !isElectron) return void doPaste('all'), true;
        if (low === 'b') return a.toggle('bold'), true;
        if (low === 'i') return a.toggle('italic'), true;
        if (low === 'u') return a.underline(), true;
        if (k === '5') return a.toggle('strike'), true;
        if (k === '1') return a.formatCells(), true;
        if (low === 'd') return a.fill('down'), true;
        if (low === 'r') return a.fill('right'), true;
        if (low === 'e') return a.flashFill(), true;
        if (low === 'k') return a.insertLink(), true;
        if (low === 'f') return setFind('find'), true;
        if (low === 'h') return setFind('replace'), true;
        if (low === 'g') return a.goTo(), true;
        if (low === 't') return a.tableStyle('blue'), true;
        if (k === '`') return a.showFormulas(), true;
        if (k === ';' && !shift) return a.insertDateTime('date'), true;
        if ((k === ':' || (k === ';' && shift))) return a.insertDateTime('time'), true;
        if (low === 'l' && shift) return a.filter(), true;
        if (k === ' ') {
          const c = doc.sel.active.c;
          ui.select({ r1: 0, c1: c, r2: MAX_ROWS - 1, c2: c }, doc.sel.active);
          return true;
        }
        if (low === 'a') {
          const rg = selRange();
          const reg = currentRegion(doc, doc.sheet, doc.sel.active.r, doc.sel.active.c);
          const same = rg.r1 === reg.r1 && rg.r2 === reg.r2 && rg.c1 === reg.c1 && rg.c2 === reg.c2;
          if (!same && (reg.r2 > reg.r1 || reg.c2 > reg.c1)) ui.select(reg, doc.sel.active);
          else ui.select({ r1: 0, c1: 0, r2: MAX_ROWS - 1, c2: MAX_COLS - 1 }, doc.sel.active);
          return true;
        }
        if (k === 'PageDown' || k === 'PageUp') {
          const dir = k === 'PageDown' ? 1 : -1;
          let i = doc.wb.activeSheet + dir;
          while (doc.wb.sheets[i]?.hidden) i += dir;
          if (i >= 0 && i < doc.wb.sheets.length) doc.setActiveSheet(i);
          return true;
        }
        if (k === '-' ) return a.remove('cells'), true;
        if (k === '+' || (k === '=' && shift)) return a.insert('cells'), true;
        if (k === '9' && !shift) return a.hide('row', true), true;
        if (k === '0' && !shift) return a.hide('col', true), true;
        if (shift) {
          const fmts: Record<string, string | undefined> = { '$': '$#,##0.00', '%': '0%', '#': 'd-mmm-yy', '@': 'h:mm AM/PM', '!': '#,##0.00', '~': undefined, '^': '0.00E+00' };
          if (k in fmts) return a.numFmt(fmts[k]), true;
          if (k === '&') return a.borders('outside'), true;
          if (k === '_') return a.borders('none'), true;
        }
      }
      if (e.altKey && (k === '=' || k === '+')) return a.autoSum('SUM'), true;
      if (shift && k === ' ' && !mod) {
        const r = doc.sel.active.r;
        ui.select({ r1: r, c1: 0, r2: r, c2: MAX_COLS - 1 }, doc.sel.active);
        return true;
      }
      if (shift && k === 'F2') return a.insertNote(), true;
      if (shift && k === 'F3') return a.insertFunction(), true;
      if (shift && k === 'F11') return a.insert('sheet'), true;
      if (k === 'F9') return a.calcNow(), true;
      if (k === 'F5') return a.goTo(), true;
      if (k === 'F4' && !mod) return false;
      if (e.altKey && k === 'ArrowDown') {
        const v = doc.sheet.validations.find((x) => x.type === 'list' && x.ranges.some((rg) => doc.sel.active.r >= rg.r1 && doc.sel.active.r <= rg.r2 && doc.sel.active.c >= rg.c1 && doc.sel.active.c <= rg.c2));
        if (v) {
          const el = document.querySelector(`.sheets-editor[data-tab="${tab.id}"] .grid-sink`) as HTMLElement | null;
          const r = el?.getBoundingClientRect();
          if (r) setListPick(new DOMRect(r.left, r.top, 120, 22));
          return true;
        }
      }
      // checkbox cells toggle with Space
      if (k === ' ' && !mod && !shift) {
        const { r, c } = doc.sel.active;
        const isCheck = doc.sheet.validations.some((v) => v.type === 'checkbox' && v.ranges.some((rg) => r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2));
        if (isCheck) {
          doc.transact('Checkbox', (tx) => tx.patch(doc.sheet, cellKey(r, c), { v: !(doc.sheet.get(r, c)?.v === true) }));
          return true;
        }
      }
      return false;
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [a, doc, ui, doPaste, tab.id],
  );

  const h: GridHandlers = useMemo(
    () => ({
      onContextMenu: (x, y, target) => openContextMenu({ x, y }, buildMenu(target)),
      onFilterMenu: (col, anchor) => setFilterMenu({ col, anchor }),
      onListPick: (anchor) => setListPick(anchor),
      onOpenLink: openLink,
      onShortcut: shortcut,
      onCopy: (e, cut) => {
        if (ui.edit) return;
        const res = copySelection(doc, cut);
        if (!res) {
          e.preventDefault();
          toast.error('The selection is too large to copy.');
          return;
        }
        e.clipboardData?.setData('text/plain', res.text);
        e.clipboardData?.setData('text/html', res.html);
        e.preventDefault();
      },
      onPaste: (e) => {
        if (ui.edit) return;
        e.preventDefault();
        const mode = pendingPaste.current ?? 'all';
        pendingPaste.current = null;
        const text = e.clipboardData?.getData('text/plain');
        const html = e.clipboardData?.getData('text/html');
        const files = [...(e.clipboardData?.files ?? [])].filter((f) => f.type.startsWith('image/'));
        if (files.length && !text) {
          const reader = new FileReader();
          reader.onload = () => {
            const { r, c } = doc.sel.active;
            upsertImage(doc, { id: newId('img'), src: String(reader.result), anchor: { r, c, dx: 0, dy: 0 }, w: 320, h: 240 });
          };
          reader.readAsDataURL(files[0]);
          return;
        }
        const rg = paste(doc, { text: text || undefined, html: html || undefined }, mode);
        if (rg) refitRows(range(rg.r1, Math.min(rg.r2, rg.r1 + 500)));
      },
      onResize: (axis, idx, size) => (axis === 'col' ? setColWidth(doc, idx, size) : setRowHeight(doc, idx, size)),
      onAutoFit: (axis, idx) => {
        if (axis === 'col') {
          const widths = idx.map((c) => autoFitWidth(doc, doc.sheet, c));
          doc.transact('AutoFit columns', (tx) => {
            tx.prop(doc.sheet, 'cols');
            idx.forEach((c, i) => doc.sheet.cols.set(c, { ...(doc.sheet.cols.get(c) ?? {}), w: widths[i] }));
          });
        } else {
          doc.transact('AutoFit rows', (tx) => {
            tx.prop(doc.sheet, 'rows');
            for (const r of idx) {
              const info = { ...(doc.sheet.rows.get(r) ?? {}) };
              const hh = autoFitHeight(doc, doc.sheet, r);
              delete info.custom;
              if (hh === null) delete info.h;
              else info.h = hh;
              doc.sheet.rows.set(r, info);
            }
          });
        }
      },
      onFill: (src, target, copyOnly) => {
        autoFill(doc, src, target, copyOnly);
        refitRows(range(target.r1, Math.min(target.r2, target.r1 + 500)));
      },
      onMove: (src, dr, dc, copy) => moveBlock(doc, src, dr, dc, copy),
      onCommitted: () => refitRows([doc.sel.active.r]),
      onSelectionEnd: () => {
        if (painter) {
          const patch: Partial<CellStyle> = {};
          const keys: (keyof CellStyle)[] = ['font', 'size', 'bold', 'italic', 'underline', 'strike', 'color', 'fill', 'hAlign', 'vAlign', 'wrap', 'indent', 'rotation', 'numFmt', 'bt', 'bb', 'bl', 'br'];
          for (const k of keys) (patch as Record<string, unknown>)[k] = painter[k];
          doc.applyStyle(patch, 'Format painter');
          setPainter(null);
        }
      },
    }),
    [buildMenu, openLink, shortcut, doc, ui, refitRows, painter],
  );

  /* ---------------------------------------------------------- controller */
  const serialize = useCallback((formatId: string) => exportWorkbook(doc, formatId, title()), [doc]);

  const commands = useCallback((): Command[] => {
    const cmd = (id: string, t: string, category: string, run: () => void, keys?: string): Command => ({ id: `sheet-${id}`, title: t, category, run, keys });
    return [
      cmd('bold', 'Bold', 'Format', () => a.toggle('bold'), 'Ctrl+B'),
      cmd('italic', 'Italic', 'Format', () => a.toggle('italic'), 'Ctrl+I'),
      cmd('format', 'Format cells…', 'Format', () => a.formatCells(), 'Ctrl+1'),
      cmd('cf', 'Conditional formatting…', 'Format', a.cfNew),
      cmd('table', 'Format as table', 'Format', () => a.tableStyle('blue'), 'Ctrl+T'),
      cmd('merge', 'Merge & centre', 'Format', () => a.merge('center')),
      cmd('wrap', 'Wrap text', 'Format', () => a.toggle('wrap')),
      cmd('freeze', 'Freeze panes', 'View', () => a.freeze('panes')),
      cmd('unfreeze', 'Unfreeze panes', 'View', () => a.freeze('none')),
      cmd('formulas', 'Show formulas', 'View', a.showFormulas, 'Ctrl+`'),
      cmd('gridlines', 'Toggle gridlines', 'View', () => a.toggleView('showGrid')),
      cmd('sort-az', 'Sort A → Z', 'Data', () => a.sort(false)),
      cmd('sort-za', 'Sort Z → A', 'Data', () => a.sort(true)),
      cmd('sort', 'Custom sort…', 'Data', a.customSort),
      cmd('filter', 'Toggle filter', 'Data', a.filter, 'Ctrl+Shift+L'),
      cmd('dupes', 'Remove duplicates…', 'Data', a.removeDuplicates),
      cmd('t2c', 'Text to columns…', 'Data', a.textToColumns),
      cmd('flash', 'Flash Fill', 'Data', a.flashFill, 'Ctrl+E'),
      cmd('dv', 'Data validation / dropdown list…', 'Data', a.validation),
      cmd('goal', 'Goal seek…', 'Data', a.goalSeek),
      cmd('import', 'Import CSV / text…', 'Data', () => void a.importData()),
      cmd('fx', 'Insert function…', 'Formulas', () => a.insertFunction(), 'Shift+F3'),
      cmd('names', 'Name manager', 'Formulas', a.nameManager),
      cmd('sum', 'AutoSum', 'Formulas', () => a.autoSum('SUM'), 'Alt+='),
      cmd('calc', 'Calculate now', 'Formulas', a.calcNow, 'F9'),
      cmd('chart-col', 'Insert column chart', 'Insert', () => a.insertChart('column')),
      cmd('chart-line', 'Insert line chart', 'Insert', () => a.insertChart('line')),
      cmd('chart-pie', 'Insert pie chart', 'Insert', () => a.insertChart('pie')),
      cmd('image', 'Insert picture', 'Insert', () => void a.insertImage()),
      cmd('link', 'Insert link', 'Insert', a.insertLink, 'Ctrl+K'),
      cmd('note', 'Insert note', 'Insert', a.insertNote, 'Shift+F2'),
      cmd('rows', 'Insert rows', 'Edit', () => a.insert('rows')),
      cmd('cols', 'Insert columns', 'Edit', () => a.insert('cols')),
      cmd('sheet', 'Insert sheet', 'Edit', () => a.insert('sheet'), 'Shift+F11'),
      cmd('find', 'Find', 'Edit', () => setFind('find'), 'Ctrl+F'),
      cmd('replace', 'Replace', 'Edit', () => setFind('replace'), 'Ctrl+H'),
      cmd('goto', 'Go to…', 'Edit', a.goTo, 'Ctrl+G'),
      cmd('protect', 'Protect sheet…', 'Review', a.protectSheet),
      cmd('setup', 'Page setup…', 'File', a.pageSetup),
      cmd('pdf', 'Export as PDF', 'File', () => void a.exportPdf()),
      cmd('print', 'Print', 'File', () => void a.print(), 'Ctrl+P'),
    ];
  }, [a]);

  useEffect(
    () =>
      registerController(tab.id, {
        save: () => saveDocument({ tabId: tab.id, kind: 'sheet', targets: SHEET_SAVE_TARGETS, defaultFormat: useSettings.getState().settings.defaultSheetFormat, serialize }),
        saveAs: (formatId) => saveDocument({ tabId: tab.id, kind: 'sheet', targets: SHEET_SAVE_TARGETS, defaultFormat: useSettings.getState().settings.defaultSheetFormat, serialize, saveAs: true, formatId }),
        exportAs: async (id) => {
          if (id === 'pdf') return exportPdf(title(), await buildSheetPdfJob(doc, 'sheet', title()));
          const target = SHEET_EXPORT_TARGETS.find((t) => t.id === id) ?? SHEET_SAVE_TARGETS.find((t) => t.id === id);
          if (target) await exportBytes(title(), target, () => serialize(id));
        },
        print: async () => printJob(await buildSheetPdfJob(doc, 'sheet', title())),
        snapshot: () => JSON.stringify(workbookToJSON(doc.wb)),
        focus: focusGrid,
        commands,
        info: () => [
          { label: 'Sheets', value: String(doc.wb.sheets.length) },
          { label: 'Cells', value: doc.wb.sheets.reduce((n, s) => n + s.cells.size, 0).toLocaleString() },
          { label: 'Formulas', value: doc.engine.stats.formulas.toLocaleString() },
        ],
      }),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [doc, tab.id, serialize, commands, focusGrid],
  );

  useEffect(() => {
    if (active) focusGrid();
  }, [active, focusGrid]);

  /* ------------------------------------------------------------ status */
  const stats = useMemo(() => (doc.selectionSize() > 1 ? doc.selectionStats() : null), [doc, docVersion, ui.version]); // eslint-disable-line react-hooks/exhaustive-deps
  const fmtNum = (n: number) => {
    const fmt = doc.activeStyle().numFmt;
    return fmt && !/^general$/i.test(fmt) && !/[dmyhs]/i.test(fmt.replace(/"[^"]*"/g, '')) ? formatStat(n, fmt) : n.toLocaleString(undefined, { maximumFractionDigits: 6 });
  };
  const circular = doc.engine.stats.circular.length > 0;
  const validationPrompt = (() => {
    const { r, c } = doc.sel.active;
    return doc.sheet.validations.find((v) => v.prompt && v.ranges.some((rg) => r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2))?.prompt;
  })();

  return (
    <div className={`editor sheets-editor${painter ? ' painter-on' : ''}`} data-tab={tab.id}>
      <SheetsRibbon doc={doc} ui={ui} a={a} onFile={() => setBackstage(true)} />
      {formulaBar && <FormulaBar doc={doc} ui={ui} onInsertFunction={() => a.insertFunction()} onCommitted={() => refitRows([doc.sel.active.r])} onBlurToGrid={focusGrid} />}
      <div className="sheet-body">
        <Grid doc={doc} ui={ui} active={active} h={h} />
        {find && <FindPanel doc={doc} ui={ui} mode={find} onClose={() => (setFind(null), focusGrid())} />}
        {validationPrompt && <div className="validation-prompt">{validationPrompt}</div>}
      </div>
      <div className="sheet-footer">
        <SheetTabs doc={doc} onChanged={focusGrid} onProtect={(i) => void protectSheetDialog(doc, i).then((m) => m && toast.info(m))} />
        <footer className="statusbar sheet-status">
          {sheet.protection?.enabled && (
            <span className="status-item">
              <Lock size={13} /> Protected
            </span>
          )}
          {circular && (
            <button className="status-item warn" data-tip="Formulas that refer to their own result. Click to select one." onClick={() => {
              const g = doc.engine.stats.circular[0];
              const sheetId = Math.floor(g / 2 ** 34);
              const idx = doc.wb.sheets.findIndex((s) => s.id === sheetId);
              if (idx >= 0) doc.setActiveSheet(idx);
              const k = g % 2 ** 34;
              ui.selectCell(Math.floor(k / MAX_COLS), k % MAX_COLS);
            }}>
              <AlertTriangle size={13} /> Circular reference
            </button>
          )}
          {!doc.wb.calcAuto && (
            <button className="status-item warn" onClick={a.calcNow}>
              <Calculator size={13} /> Calculate (F9)
            </button>
          )}
          {painter && <span className="status-item active">Format painter: select cells</span>}
          {doc.clip && <span className="status-item muted">Select a destination and press Enter or Ctrl+V</span>}
          <span className="spacer" />
          {stats && stats.count > 0 && (
            <div className="sel-stats">
              {stats.numCount > 0 && <span data-tip="Average of the selected numbers">Average: {fmtNum(stats.avg!)}</span>}
              <span data-tip="Non-empty cells">Count: {stats.count.toLocaleString()}</span>
              {stats.numCount > 1 && <span data-tip="Smallest number">Min: {fmtNum(stats.min!)}</span>}
              {stats.numCount > 1 && <span data-tip="Largest number">Max: {fmtNum(stats.max!)}</span>}
              {stats.numCount > 0 && (
                <button className="status-item strong" data-tip="Click to copy the sum" onClick={() => void navigator.clipboard?.writeText(String(stats.sum)).then(() => toast.success('Sum copied', { duration: 1500 }))}>
                  <Sigma size={13} /> {fmtNum(stats.sum)}
                </button>
              )}
            </div>
          )}
          <SaveState dirty={tab.dirty} path={tab.path} />
          <span className="status-sep" />
          <button className={`status-item${sheet.view.showGrid ? ' active' : ''}`} data-tip="Gridlines" onClick={() => a.toggleView('showGrid')}>
            <Grid3x3 size={14} />
          </button>
          <button className="status-item" data-tip="Page setup" onClick={a.pageSetup}>
            <FileSpreadsheet size={14} />
          </button>
          <span className="status-sep" />
          <ZoomControl zoom={ui.zoom} onZoom={(z) => ui.setZoom(z)} extra={[{ label: 'Zoom to selection', onSelect: () => a.zoom('selection') }]} />
        </footer>
      </div>
      {filterMenu && doc.sheet.filter && <FilterMenu doc={doc} col={filterMenu.col} anchor={filterMenu.anchor} onClose={() => (setFilterMenu(null), focusGrid())} />}
      {listPick && (
        <ListPicker
          doc={doc}
          anchor={listPick}
          onClose={() => (setListPick(null), focusGrid())}
          onPick={(v) => {
            doc.enter(v);
            setListPick(null);
            focusGrid();
          }}
        />
      )}
      {backstage && (
        <Backstage
          app="sheet"
          tab={tab}
          onClose={() => (setBackstage(false), focusGrid())}
          onSave={() => void saveDocument({ tabId: tab.id, kind: 'sheet', targets: SHEET_SAVE_TARGETS, defaultFormat: appSettings.defaultSheetFormat, serialize })}
          onSaveAs={(fid) => void saveDocument({ tabId: tab.id, kind: 'sheet', targets: SHEET_SAVE_TARGETS, defaultFormat: appSettings.defaultSheetFormat, serialize, saveAs: true, formatId: fid })}
          onExport={(id) => {
            void (async () => {
              if (id === 'pdf') return exportPdf(title(), await buildSheetPdfJob(doc, 'sheet', title()));
              const target = SHEET_EXPORT_TARGETS.find((t) => t.id === id);
              if (target) await exportBytes(title(), target, () => serialize(id));
            })();
          }}
          onPrint={() => void a.print()}
          saveTargets={SHEET_SAVE_TARGETS}
          exportTargets={SHEET_EXPORT_TARGETS}
          libreOffice={isElectron}
          info={[
            { label: 'Sheets', value: String(doc.wb.sheets.length) },
            { label: 'Formulas', value: doc.engine.stats.formulas.toLocaleString() },
            { label: 'Author', value: doc.wb.props.author || appSettings.authorName || '—' },
          ]}
        />
      )}
    </div>
  );
}

function range(a: number, b: number): number[] {
  const out: number[] = [];
  for (let i = a; i <= b; i++) out.push(i);
  return out;
}

function formatStat(n: number, fmt: string): string {
  return formatValue(n, fmt).text;
}
