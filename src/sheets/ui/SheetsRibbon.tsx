import {
  AArrowDown,
  AArrowUp,
  AlignCenter,
  AlignLeft,
  AlignRight,
  AlignVerticalJustifyCenter,
  AlignVerticalJustifyEnd,
  AlignVerticalJustifyStart,
  ArrowDownAZ,
  ArrowUpZA,
  AreaChart,
  Baseline,
  BarChart3,
  BarChartHorizontal,
  Bold,
  BookA,
  Calculator,
  ChartScatter,
  ClipboardPaste,
  Columns3,
  Copy,
  DatabaseZap,
  DollarSign,
  Eraser,
  Eye,
  FileSpreadsheet,
  Filter,
  FilterX,
  FunctionSquare,
  Grid3x3,
  Image as ImageIcon,
  Indent,
  Italic,
  LayoutGrid,
  LineChart,
  Link2,
  ListChecks,
  Lock,
  Maximize2,
  MessageSquarePlus,
  Minus,
  Outdent,
  Paintbrush,
  Palette,
  Percent,
  PieChart,
  Plus,
  Printer,
  Redo2,
  RefreshCw,
  Rows3,
  Scissors,
  Search,
  Sigma,
  Snowflake,
  SplitSquareHorizontal,
  Strikethrough,
  Table as TableIcon,
  Target,
  TableCellsMerge,
  Trash2,
  Type,
  Underline,
  Undo2,
  WandSparkles,
  WrapText,
  ZoomIn,
  PaintBucket,
  Hash,
  Radar,
  Sparkles,
  CopyX,
  Workflow,
} from 'lucide-react';
import { useState, useSyncExternalStore, type ReactNode } from 'react';
import { ColorSplitButton } from '@/ui/color';
import { FontFamilyCombo, FontSizeCombo, nextFontSize } from '@/ui/fontControls';
import type { MenuItem } from '@/ui/menu';
import { RBigButton, RButton, RDropdown, RibbonGroup, RRow, RRows, RSep, RSplit, Ribbon, type RibbonTab } from '@/ui/ribbon';
import type { BorderKind, SheetDoc } from '../doc';
import type { PasteMode } from '../ops/clipboard';
import type { BorderStyle, CellStyle, ChartType } from '../model/types';
import type { SheetUI } from './controller';
import { BorderIcon, CHART_TYPES } from './dialogs';
import { CELL_STYLES, TABLE_STYLES } from './tableStyles';

export interface SheetActions {
  undo(): void;
  redo(): void;
  paste(mode?: PasteMode): void;
  pasteSpecial(): void;
  cut(): void;
  copy(): void;
  painter(): void;
  painterOn: boolean;
  style(patch: Partial<CellStyle>, label?: string): void;
  toggle(key: 'bold' | 'italic' | 'strike' | 'wrap'): void;
  underline(): void;
  borders(kind: BorderKind, style?: BorderStyle, color?: string): void;
  merge(mode: 'merge' | 'center' | 'across' | 'unmerge'): void;
  numFmt(fmt: string | undefined): void;
  decimals(dir: 1 | -1): void;
  indent(dir: 1 | -1): void;
  formatCells(tab?: 'number' | 'alignment' | 'font' | 'border' | 'fill' | 'protection'): void;
  cfPreset(kind: string, styleIdx?: number): void;
  cfNew(): void;
  cfManage(): void;
  cfClear(scope: 'selection' | 'sheet'): void;
  tableStyle(id: string): void;
  cellStyle(id: string): void;
  insert(what: 'rows' | 'cols' | 'cells' | 'sheet'): void;
  remove(what: 'rows' | 'cols' | 'cells' | 'sheet'): void;
  rowHeight(): void;
  colWidth(): void;
  autofit(axis: 'row' | 'col'): void;
  hide(axis: 'row' | 'col', hidden: boolean): void;
  renameSheet(): void;
  protectSheet(): void;
  autoSum(fn: string): void;
  fill(dir: 'down' | 'right' | 'series'): void;
  clear(what: 'all' | 'contents' | 'formats' | 'notes' | 'links'): void;
  sort(desc: boolean): void;
  customSort(): void;
  filter(): void;
  clearFilter(): void;
  reapplyFilter(): void;
  find(mode: 'find' | 'replace'): void;
  goTo(): void;
  selectSpecial(kind: 'formulas' | 'constants' | 'blanks' | 'notes' | 'cf' | 'validation'): void;
  insertChart(type: ChartType): void;
  insertImage(): void;
  insertLink(): void;
  insertNote(): void;
  insertFunction(cat?: string): void;
  insertDateTime(kind: 'date' | 'time'): void;
  insertCheckbox(): void;
  nameManager(): void;
  showFormulas(): void;
  calcNow(): void;
  setCalcAuto(auto: boolean): void;
  trace(kind: 'precedents' | 'dependents' | 'clear'): void;
  textToColumns(): void;
  flashFill(): void;
  removeDuplicates(): void;
  validation(): void;
  goalSeek(): void;
  importData(): void;
  freeze(kind: 'panes' | 'row' | 'col' | 'none'): void;
  zoom(z: number | 'selection'): void;
  toggleView(key: 'showGrid' | 'showHeaders'): void;
  toggleFormulaBar(): void;
  formulaBarOn: boolean;
  pageSetup(): void;
  print(): void;
  exportPdf(): void;
  chart(action: 'edit' | 'delete' | 'type', type?: ChartType): void;
  stats(): void;
}

const I = 17;

export function SheetsRibbon({ doc, ui, a, onFile }: { doc: SheetDoc; ui: SheetUI; a: SheetActions; onFile: () => void }) {
  useSyncExternalStore(doc.subscribe, doc.getVersion);
  useSyncExternalStore(ui.subscribe, ui.getVersion);
  const [active, setActive] = useState('home');
  const st = doc.activeStyle();
  const chartSelected = ui.selectedObject?.kind === 'chart';
  const tabs: RibbonTab[] = [
    { id: 'home', label: 'Home', content: <HomeTab doc={doc} a={a} st={st} /> },
    { id: 'insert', label: 'Insert', content: <InsertTab a={a} /> },
    { id: 'layout', label: 'Page layout', content: <LayoutTab doc={doc} a={a} /> },
    { id: 'formulas', label: 'Formulas', content: <FormulasTab doc={doc} ui={ui} a={a} /> },
    { id: 'data', label: 'Data', content: <DataTab doc={doc} a={a} /> },
    { id: 'review', label: 'Review', content: <ReviewTab doc={doc} a={a} /> },
    { id: 'view', label: 'View', content: <ViewTab doc={doc} ui={ui} a={a} /> },
    { id: 'chart', label: 'Chart', content: <ChartTab a={a} />, contextual: true, hidden: !chartSelected },
  ];
  const effective = chartSelected && active === 'chart' ? 'chart' : active === 'chart' ? 'home' : active;
  return (
    <Ribbon
      app="sheet"
      tabs={tabs}
      active={effective}
      onActive={setActive}
      onFile={onFile}
      right={
        <div className="ribbon-right">
          <button type="button" className="icon-btn icon-btn-sm" data-tip="Undo" data-tip-key="Ctrl+Z" disabled={!doc.canUndo()} onMouseDown={(e) => e.preventDefault()} onClick={a.undo}>
            <Undo2 size={16} />
          </button>
          <button type="button" className="icon-btn icon-btn-sm" data-tip="Redo" data-tip-key="Ctrl+Y" disabled={!doc.canRedo()} onMouseDown={(e) => e.preventDefault()} onClick={a.redo}>
            <Redo2 size={16} />
          </button>
        </div>
      }
    />
  );
}

/* =================================================================== HOME */

const NUMBER_MENU = (a: SheetActions): MenuItem[] => [
  { label: 'General', hint: 'No specific format', onSelect: () => a.numFmt(undefined) },
  { label: 'Number', hint: '1,234.00', onSelect: () => a.numFmt('#,##0.00') },
  { label: 'Currency', hint: '$1,234.00', onSelect: () => a.numFmt('$#,##0.00') },
  { label: 'Accounting', hint: '$ 1,234.00', onSelect: () => a.numFmt('_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)') },
  { label: 'Short date', hint: '3/14/2025', onSelect: () => a.numFmt('m/d/yyyy') },
  { label: 'Long date', hint: 'Friday, March 14, 2025', onSelect: () => a.numFmt('dddd, mmmm d, yyyy') },
  { label: 'Time', hint: '1:30:00 PM', onSelect: () => a.numFmt('h:mm:ss AM/PM') },
  { label: 'Percentage', hint: '12.50%', onSelect: () => a.numFmt('0.00%') },
  { label: 'Fraction', hint: '1 1/4', onSelect: () => a.numFmt('# ?/?') },
  { label: 'Scientific', hint: '1.23E+03', onSelect: () => a.numFmt('0.00E+00') },
  { label: 'Text', hint: 'Keeps what you type', onSelect: () => a.numFmt('@') },
  { separator: true },
  { label: 'More number formats…', shortcut: 'Ctrl+1', onSelect: () => a.formatCells('number') },
];

function numLabel(fmt: string | undefined): string {
  if (!fmt || /^general$/i.test(fmt)) return 'General';
  if (fmt === '@') return 'Text';
  if (/%/.test(fmt)) return 'Percentage';
  if (/^_\(/.test(fmt)) return 'Accounting';
  if (/[$€£¥₹]|"[^"]+"#/.test(fmt)) return 'Currency';
  if (/[dmy]{1,4}/i.test(fmt.replace(/"[^"]*"/g, '')) && !/0/.test(fmt)) return /[hs]/.test(fmt) ? 'Date & time' : 'Date';
  if (/[hs]/i.test(fmt) && !/0/.test(fmt.replace(/\.0+/, ''))) return 'Time';
  if (/E\+/.test(fmt)) return 'Scientific';
  if (/\?\//.test(fmt)) return 'Fraction';
  return 'Number';
}

function HomeTab({ doc, a, st }: { doc: SheetDoc; a: SheetActions; st: CellStyle }) {
  const bordersMenu: MenuItem[] = (
    [
      ['bottom', 'Bottom border'],
      ['top', 'Top border'],
      ['left', 'Left border'],
      ['right', 'Right border'],
      ['none', 'No border'],
      ['all', 'All borders'],
      ['outside', 'Outside borders'],
      ['thickOutside', 'Thick outside borders'],
      ['doubleBottom', 'Bottom double border'],
      ['thickBottom', 'Thick bottom border'],
      ['topBottom', 'Top and bottom border'],
      ['topThickBottom', 'Top and thick bottom border'],
      ['inside', 'Inside borders'],
    ] as [BorderKind, string][]
  ).map(([k, label]) => ({ label, icon: <BorderIcon kind={k} />, onSelect: () => a.borders(k) }));
  bordersMenu.push({ separator: true }, { label: 'More borders…', onSelect: () => a.formatCells('border') });
  return (
    <>
      <RibbonGroup label="Clipboard">
        <RSplit
          big
          icon={<ClipboardPaste />}
          label="Paste"
          tip="Paste (Ctrl+V)"
          onClick={() => a.paste()}
          menu={[
            { label: 'Paste', shortcut: 'Ctrl+V', onSelect: () => a.paste() },
            { label: 'Paste values', shortcut: 'Ctrl+Shift+V', onSelect: () => a.paste('values') },
            { label: 'Paste formulas', onSelect: () => a.paste('formulas') },
            { label: 'Paste formatting', onSelect: () => a.paste('formats') },
            { label: 'Transpose', onSelect: () => a.paste('transpose') },
            { label: 'Keep column widths', onSelect: () => a.paste('colWidths') },
            { separator: true },
            { label: 'Paste special…', shortcut: 'Ctrl+Alt+V', onSelect: a.pasteSpecial },
          ]}
        />
        <RRows>
          <RButton icon={<Scissors size={I} />} label="Cut" keys="Ctrl+X" onClick={a.cut} />
          <RButton icon={<Copy size={I} />} label="Copy" keys="Ctrl+C" onClick={a.copy} />
          <RButton icon={<Paintbrush size={I} />} label="Format painter" tip="Format painter — copy formatting to other cells" active={a.painterOn} onClick={a.painter} />
        </RRows>
      </RibbonGroup>

      <RibbonGroup label="Font">
        <RRows>
          <RRow>
            <FontFamilyCombo value={st.font ?? 'Calibri'} onChange={(font) => a.style({ font: font === 'Calibri' ? undefined : font }, 'Font')} />
            <FontSizeCombo value={st.size ?? 11} onChange={(pt) => a.style({ size: pt === 11 ? undefined : pt }, 'Font size')} />
            <RButton icon={<AArrowUp size={I} />} tip="Increase font size" onClick={() => a.style({ size: nextFontSize(st.size ?? 11, 1) }, 'Font size')} />
            <RButton icon={<AArrowDown size={I} />} tip="Decrease font size" onClick={() => a.style({ size: nextFontSize(st.size ?? 11, -1) }, 'Font size')} />
          </RRow>
          <RRow>
            <RButton icon={<Bold size={I} />} tip="Bold" keys="Ctrl+B" active={!!st.bold} onClick={() => a.toggle('bold')} />
            <RButton icon={<Italic size={I} />} tip="Italic" keys="Ctrl+I" active={!!st.italic} onClick={() => a.toggle('italic')} />
            <RButton icon={<Underline size={I} />} tip="Underline" keys="Ctrl+U" active={!!st.underline} onClick={a.underline} />
            <RButton icon={<Strikethrough size={I} />} tip="Strikethrough" keys="Ctrl+5" active={!!st.strike} onClick={() => a.toggle('strike')} />
            <RSep />
            <RSplit icon={<Grid3x3 size={I} />} tip="Borders" onClick={() => a.borders('bottom')} menu={bordersMenu} />
            <ColorSplitButton icon={<PaintBucket size={I} />} color={st.fill ?? '#ffff00'} palette="theme" tip="Fill colour" noneLabel="No fill" onApply={(c) => a.style({ fill: c ?? undefined }, 'Fill colour')} />
            <ColorSplitButton icon={<Baseline size={I} />} color={st.color ?? '#e5484d'} tip="Font colour" noneLabel="Automatic" onApply={(c) => a.style({ color: c ?? undefined }, 'Font colour')} />
          </RRow>
        </RRows>
      </RibbonGroup>

      <RibbonGroup label="Alignment" collapse={2} icon={<AlignLeft />}>
        <RRows>
          <RRow>
            <RButton icon={<AlignVerticalJustifyStart size={I} />} tip="Top align" active={st.vAlign === 'top'} onClick={() => a.style({ vAlign: 'top' }, 'Align')} />
            <RButton icon={<AlignVerticalJustifyCenter size={I} />} tip="Middle align" active={st.vAlign === 'middle'} onClick={() => a.style({ vAlign: 'middle' }, 'Align')} />
            <RButton icon={<AlignVerticalJustifyEnd size={I} />} tip="Bottom align" active={!st.vAlign || st.vAlign === 'bottom'} onClick={() => a.style({ vAlign: undefined }, 'Align')} />
            <RSep />
            <RDropdown
              icon={<Type size={I} />}
              tip="Orientation"
              showLabel={false}
              menu={[
                { label: 'Angle counterclockwise', onSelect: () => a.style({ rotation: 45 }, 'Orientation') },
                { label: 'Angle clockwise', onSelect: () => a.style({ rotation: -45 }, 'Orientation') },
                { label: 'Vertical text', onSelect: () => a.style({ rotation: 255 }, 'Orientation') },
                { label: 'Rotate text up', onSelect: () => a.style({ rotation: 90 }, 'Orientation') },
                { label: 'Rotate text down', onSelect: () => a.style({ rotation: -90 }, 'Orientation') },
                { label: 'Horizontal (reset)', onSelect: () => a.style({ rotation: undefined }, 'Orientation') },
                { separator: true },
                { label: 'Format cell alignment…', onSelect: () => a.formatCells('alignment') },
              ]}
            />
            <RButton icon={<WrapText size={I} />} tip="Wrap text" active={!!st.wrap} onClick={() => a.toggle('wrap')} />
          </RRow>
          <RRow>
            <RButton icon={<AlignLeft size={I} />} tip="Align left" active={st.hAlign === 'left'} onClick={() => a.style({ hAlign: st.hAlign === 'left' ? undefined : 'left' }, 'Align')} />
            <RButton icon={<AlignCenter size={I} />} tip="Centre" active={st.hAlign === 'center'} onClick={() => a.style({ hAlign: st.hAlign === 'center' ? undefined : 'center' }, 'Align')} />
            <RButton icon={<AlignRight size={I} />} tip="Align right" active={st.hAlign === 'right'} onClick={() => a.style({ hAlign: st.hAlign === 'right' ? undefined : 'right' }, 'Align')} />
            <RSep />
            <RButton icon={<Outdent size={I} />} tip="Decrease indent" onClick={() => a.indent(-1)} />
            <RButton icon={<Indent size={I} />} tip="Increase indent" onClick={() => a.indent(1)} />
            <RSplit
              icon={<TableCellsMerge size={I} />}
              label="Merge"
              tip="Merge & centre"
              onClick={() => a.merge('center')}
              menu={[
                { label: 'Merge & centre', onSelect: () => a.merge('center') },
                { label: 'Merge across', onSelect: () => a.merge('across') },
                { label: 'Merge cells', onSelect: () => a.merge('merge') },
                { label: 'Unmerge cells', onSelect: () => a.merge('unmerge') },
              ]}
            />
          </RRow>
        </RRows>
      </RibbonGroup>

      <RibbonGroup label="Number">
        <RRows>
          <RRow>
            <RDropdown label={numLabel(st.numFmt)} tip="Number format" width={132} menu={() => NUMBER_MENU(a)} />
          </RRow>
          <RRow>
            <RDropdown
              icon={<DollarSign size={I} />}
              tip="Accounting number format"
              showLabel={false}
              menu={[
                ['$', '$ English (United States)'],
                ['€', '€ Euro'],
                ['£', '£ English (United Kingdom)'],
                ['¥', '¥ Japanese / Chinese'],
                ['₹', '₹ Indian rupee'],
                ['CHF', 'CHF Swiss franc'],
              ].map(([s, label]) => ({ label, onSelect: () => a.numFmt(`_(${s === '$' ? '$' : `"${s}"`}* #,##0.00_);_(${s === '$' ? '$' : `"${s}"`}* (#,##0.00);_(${s === '$' ? '$' : `"${s}"`}* "-"??_);_(@_)`) }))}
            />
            <RButton icon={<Percent size={I} />} tip="Percent style" keys="Ctrl+Shift+%" onClick={() => a.numFmt('0%')} />
            <RButton icon={<Hash size={I} />} tip="Comma style" onClick={() => a.numFmt('#,##0.00')} />
            <RButton icon={<span className="dec-icon">.0→.00</span>} tip="Increase decimal" onClick={() => a.decimals(1)} />
            <RButton icon={<span className="dec-icon">.00→.0</span>} tip="Decrease decimal" onClick={() => a.decimals(-1)} />
          </RRow>
        </RRows>
      </RibbonGroup>

      <RibbonGroup label="Styles" collapse={2} icon={<Palette />}>
        <RBigButton
          icon={<Sparkles />}
          label="Conditional"
          tip="Conditional formatting — highlight cells with rules, data bars, colour scales and icons"
          menu={() => [
            {
              label: 'Highlight cells rules',
              submenu: [
                { label: 'Greater than…', onSelect: () => a.cfPreset('greater') },
                { label: 'Less than…', onSelect: () => a.cfPreset('less') },
                { label: 'Between…', onSelect: () => a.cfPreset('between') },
                { label: 'Equal to…', onSelect: () => a.cfPreset('equal') },
                { label: 'Text that contains…', onSelect: () => a.cfPreset('text') },
                { label: 'A date occurring…', onSelect: () => a.cfPreset('date') },
                { label: 'Duplicate values', onSelect: () => a.cfPreset('duplicate') },
                { label: 'Unique values', onSelect: () => a.cfPreset('unique') },
              ],
            },
            {
              label: 'Top / bottom rules',
              submenu: [
                { label: 'Top 10 items', onSelect: () => a.cfPreset('top10') },
                { label: 'Top 10%', onSelect: () => a.cfPreset('top10p') },
                { label: 'Bottom 10 items', onSelect: () => a.cfPreset('bottom10') },
                { label: 'Above average', onSelect: () => a.cfPreset('above') },
                { label: 'Below average', onSelect: () => a.cfPreset('below') },
              ],
            },
            {
              label: 'Data bars',
              submenu: ['#638ec6', '#63be7b', '#ff555a', '#ffb628', '#008aef', '#d6007b'].map((c) => ({
                label: <span className="swatch-row"><span className="swatch-bar" style={{ background: `linear-gradient(90deg, ${c}, ${c}33)` }} /> </span>,
                onSelect: () => a.cfPreset(`bar:${c}`),
              })),
            },
            {
              label: 'Colour scales',
              submenu: [
                ['#f8696b,#ffeb84,#63be7b', 'Red – yellow – green'],
                ['#63be7b,#ffeb84,#f8696b', 'Green – yellow – red'],
                ['#f8696b,#ffffff,#5a8ac6', 'Red – white – blue'],
                ['#ffffff,#63be7b', 'White – green'],
                ['#ffffff,#f8696b', 'White – red'],
                ['#fcfcff,#5a8ac6', 'White – blue'],
              ].map(([cols, label]) => ({
                label: (
                  <span className="swatch-row">
                    <span className="swatch-bar" style={{ background: `linear-gradient(90deg, ${cols})` }} /> {label}
                  </span>
                ),
                onSelect: () => a.cfPreset(`scale:${cols}`),
              })),
            },
            {
              label: 'Icon sets',
              submenu: [
                ['3Arrows', 'Arrows'],
                ['3TrafficLights', 'Traffic lights'],
                ['3Symbols', 'Symbols'],
                ['3Stars', 'Stars'],
                ['3Flags', 'Flags'],
                ['5Ratings', 'Ratings'],
              ].map(([id, label]) => ({ label, onSelect: () => a.cfPreset(`icons:${id}`) })),
            },
            { separator: true },
            { label: 'New rule…', onSelect: a.cfNew },
            { label: 'Clear rules from selected cells', onSelect: () => a.cfClear('selection') },
            { label: 'Clear rules from entire sheet', onSelect: () => a.cfClear('sheet') },
            { label: 'Manage rules…', onSelect: a.cfManage },
          ]}
        />
        <RDropdown
          icon={<TableIcon size={I} />}
          label="Format as table"
          tip="Format as table — banded rows, header row and filter buttons"
          panel={(close) => (
            <div className="table-style-gallery">
              {TABLE_STYLES.map((t) => (
                <button key={t.id} className="table-style" aria-label={t.label} data-tip={t.label} onClick={() => (a.tableStyle(t.id), close())}>
                  <span style={{ background: t.header.fill ?? 'transparent', borderBottom: t.header.bb ? `2px solid ${t.header.bb.color}` : undefined }} />
                  <span style={{ background: t.band1.fill ?? 'transparent' }} />
                  <span style={{ background: t.band2.fill ?? 'transparent' }} />
                  <span style={{ background: t.band1.fill ?? 'transparent' }} />
                </button>
              ))}
            </div>
          )}
        />
        <RDropdown
          icon={<Palette size={I} />}
          label="Cell styles"
          tip="Cell styles"
          panel={(close) => (
            <div className="cell-style-gallery">
              {[...new Set(CELL_STYLES.map((c) => c.group))].map((g) => (
                <div key={g}>
                  <div className="gallery-head">{g}</div>
                  <div className="cell-style-grid">
                    {CELL_STYLES.filter((c) => c.group === g).map((c) => (
                      <button
                        key={c.id}
                        className="cell-style"
                        onClick={() => (a.cellStyle(c.id), close())}
                        style={{ background: c.style.fill, color: c.style.color, fontWeight: c.style.bold ? 700 : 400, fontStyle: c.style.italic ? 'italic' : 'normal', borderBottom: c.style.bb ? `2px ${c.style.bb.style === 'double' ? 'double' : 'solid'} ${c.style.bb.color}` : undefined }}
                      >
                        {c.label}
                      </button>
                    ))}
                  </div>
                </div>
              ))}
            </div>
          )}
        />
      </RibbonGroup>

      <RibbonGroup label="Cells" collapse={1} icon={<Rows3 />}>
        <RRows>
          <RDropdown
            icon={<Plus size={I} />}
            label="Insert"
            menu={[
              { label: 'Insert cells…', shortcut: 'Ctrl+Shift+=', onSelect: () => a.insert('cells') },
              { label: 'Insert sheet rows', onSelect: () => a.insert('rows') },
              { label: 'Insert sheet columns', onSelect: () => a.insert('cols') },
              { label: 'Insert sheet', shortcut: 'Shift+F11', onSelect: () => a.insert('sheet') },
            ]}
          />
          <RDropdown
            icon={<Minus size={I} />}
            label="Delete"
            menu={[
              { label: 'Delete cells…', shortcut: 'Ctrl+-', onSelect: () => a.remove('cells') },
              { label: 'Delete sheet rows', onSelect: () => a.remove('rows') },
              { label: 'Delete sheet columns', onSelect: () => a.remove('cols') },
              { label: 'Delete sheet', danger: true, onSelect: () => a.remove('sheet') },
            ]}
          />
          <RDropdown
            icon={<LayoutGrid size={I} />}
            label="Format"
            menu={[
              { header: 'Cell size' },
              { label: 'Row height…', onSelect: a.rowHeight },
              { label: 'AutoFit row height', onSelect: () => a.autofit('row') },
              { label: 'Column width…', onSelect: a.colWidth },
              { label: 'AutoFit column width', onSelect: () => a.autofit('col') },
              { header: 'Visibility' },
              { label: 'Hide rows', shortcut: 'Ctrl+9', onSelect: () => a.hide('row', true) },
              { label: 'Hide columns', shortcut: 'Ctrl+0', onSelect: () => a.hide('col', true) },
              { label: 'Unhide rows', onSelect: () => a.hide('row', false) },
              { label: 'Unhide columns', onSelect: () => a.hide('col', false) },
              { header: 'Organise sheets' },
              { label: 'Rename sheet', onSelect: a.renameSheet },
              { header: 'Protection' },
              { label: doc.sheet.protection?.enabled ? 'Unprotect sheet…' : 'Protect sheet…', onSelect: a.protectSheet },
              { label: 'Lock / unlock cells', onSelect: () => a.formatCells('protection') },
              { separator: true },
              { label: 'Format cells…', shortcut: 'Ctrl+1', onSelect: () => a.formatCells() },
            ]}
          />
        </RRows>
      </RibbonGroup>

      <RibbonGroup label="Editing" collapse={1} icon={<Search />}>
        <RRows>
          <RRow>
            <RSplit
              icon={<Sigma size={I} />}
              label="AutoSum"
              tip="AutoSum (Alt+=)"
              onClick={() => a.autoSum('SUM')}
              menu={['SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN', 'MEDIAN'].map((f) => ({ label: f[0] + f.slice(1).toLowerCase(), onSelect: () => a.autoSum(f) })).concat([{ label: 'More functions…', onSelect: () => a.insertFunction() }])}
            />
          </RRow>
          <RRow>
            <RDropdown
              icon={<ArrowDownAZ size={I} />}
              label="Sort & filter"
              menu={[
                { label: 'Sort A to Z', icon: <ArrowDownAZ size={15} />, onSelect: () => a.sort(false) },
                { label: 'Sort Z to A', icon: <ArrowUpZA size={15} />, onSelect: () => a.sort(true) },
                { label: 'Custom sort…', onSelect: a.customSort },
                { separator: true },
                { label: 'Filter', icon: <Filter size={15} />, shortcut: 'Ctrl+Shift+L', checked: !!doc.sheet.filter, onSelect: a.filter },
                { label: 'Clear', icon: <FilterX size={15} />, disabled: !doc.sheet.filter, onSelect: a.clearFilter },
                { label: 'Reapply', disabled: !doc.sheet.filter, onSelect: a.reapplyFilter },
              ]}
            />
          </RRow>
          <RRow>
            <RDropdown
              icon={<Search size={I} />}
              label="Find"
              menu={[
                { label: 'Find…', shortcut: 'Ctrl+F', onSelect: () => a.find('find') },
                { label: 'Replace…', shortcut: 'Ctrl+H', onSelect: () => a.find('replace') },
                { label: 'Go to…', shortcut: 'Ctrl+G', onSelect: a.goTo },
                { separator: true },
                { label: 'Formulas', onSelect: () => a.selectSpecial('formulas') },
                { label: 'Constants', onSelect: () => a.selectSpecial('constants') },
                { label: 'Blanks', onSelect: () => a.selectSpecial('blanks') },
                { label: 'Notes', onSelect: () => a.selectSpecial('notes') },
                { label: 'Conditional formatting', onSelect: () => a.selectSpecial('cf') },
                { label: 'Data validation', onSelect: () => a.selectSpecial('validation') },
              ]}
            />
            <RDropdown
              icon={<Eraser size={I} />}
              tip="Clear"
              showLabel={false}
              menu={[
                { label: 'Clear all', onSelect: () => a.clear('all') },
                { label: 'Clear formats', onSelect: () => a.clear('formats') },
                { label: 'Clear contents', shortcut: 'Delete', onSelect: () => a.clear('contents') },
                { label: 'Clear notes', onSelect: () => a.clear('notes') },
                { label: 'Clear hyperlinks', onSelect: () => a.clear('links') },
              ]}
            />
            <RDropdown
              icon={<Rows3 size={I} />}
              tip="Fill"
              showLabel={false}
              menu={[
                { label: 'Down', shortcut: 'Ctrl+D', onSelect: () => a.fill('down') },
                { label: 'Right', shortcut: 'Ctrl+R', onSelect: () => a.fill('right') },
                { label: 'Flash Fill', shortcut: 'Ctrl+E', onSelect: a.flashFill },
              ]}
            />
          </RRow>
        </RRows>
      </RibbonGroup>
    </>
  );
}

/* ================================================================= INSERT */

const CHART_ICONS: Partial<Record<ChartType, ReactNode>> = {
  column: <BarChart3 />,
  bar: <BarChartHorizontal />,
  line: <LineChart />,
  area: <AreaChart />,
  pie: <PieChart />,
  scatter: <ChartScatter />,
  radar: <Radar />,
};

function InsertTab({ a }: { a: SheetActions }) {
  return (
    <>
      <RibbonGroup label="Tables">
        <RBigButton icon={<TableIcon />} label="Table" tip="Format the data as a table with banded rows and filters (Ctrl+T)" onClick={() => a.tableStyle('blue')} />
      </RibbonGroup>
      <RibbonGroup label="Illustrations">
        <RBigButton icon={<ImageIcon />} label="Pictures" tip="Insert a picture from your computer" onClick={a.insertImage} />
      </RibbonGroup>
      <RibbonGroup label="Charts">
        {(['column', 'line', 'pie', 'bar'] as ChartType[]).map((t) => (
          <RBigButton key={t} icon={CHART_ICONS[t]} label={CHART_TYPES.find((c) => c.value === t)!.label} tip={`Insert a ${t} chart from the selected data`} onClick={() => a.insertChart(t)} />
        ))}
        <RBigButton
          icon={<AreaChart />}
          label="More"
          tip="More chart types"
          menu={CHART_TYPES.map((c) => ({ label: c.label, icon: CHART_ICONS[c.value] ? <span className="menu-icon-sm">{CHART_ICONS[c.value]}</span> : undefined, onSelect: () => a.insertChart(c.value) }))}
        />
      </RibbonGroup>
      <RibbonGroup label="Controls">
        <RBigButton icon={<ListChecks />} label="Checkbox" tip="Insert checkboxes in the selected cells" onClick={a.insertCheckbox} />
        <RBigButton icon={<Workflow />} label="Dropdown" tip="Add a dropdown list (data validation)" onClick={a.validation} />
      </RibbonGroup>
      <RibbonGroup label="Links & notes">
        <RBigButton icon={<Link2 />} label="Link" tip="Insert link (Ctrl+K)" onClick={a.insertLink} />
        <RBigButton icon={<MessageSquarePlus />} label="Note" tip="Add a note to the cell (Shift+F2)" onClick={a.insertNote} />
      </RibbonGroup>
      <RibbonGroup label="Formulas">
        <RBigButton icon={<FunctionSquare />} label="Function" tip="Insert function (Shift+F3)" onClick={() => a.insertFunction()} />
        <RRows>
          <RButton icon={<BookA size={I} />} label="Today's date" showLabel keys="Ctrl+;" onClick={() => a.insertDateTime('date')} />
          <RButton icon={<Target size={I} />} label="Current time" showLabel keys="Ctrl+Shift+;" onClick={() => a.insertDateTime('time')} />
        </RRows>
      </RibbonGroup>
    </>
  );
}

/* ============================================================ PAGE LAYOUT */

function LayoutTab({ doc, a }: { doc: SheetDoc; a: SheetActions }) {
  return (
    <>
      <RibbonGroup label="Page setup">
        <RBigButton icon={<FileSpreadsheet />} label="Page setup" tip="Orientation, paper size, margins, scaling and what to print" onClick={a.pageSetup} />
        <RBigButton icon={<Printer />} label="Print" tip="Print (Ctrl+P)" onClick={a.print} />
        <RBigButton icon={<Maximize2 />} label="PDF" tip="Export this sheet as PDF" onClick={a.exportPdf} />
      </RibbonGroup>
      <RibbonGroup label="Sheet options">
        <RRows>
          <RButton icon={<Grid3x3 size={I} />} label="Gridlines" showLabel active={doc.sheet.view.showGrid} onClick={() => a.toggleView('showGrid')} />
          <RButton icon={<Columns3 size={I} />} label="Headings" showLabel active={doc.sheet.view.showHeaders} onClick={() => a.toggleView('showHeaders')} />
        </RRows>
      </RibbonGroup>
    </>
  );
}

/* =============================================================== FORMULAS */

function FormulasTab({ doc, ui, a }: { doc: SheetDoc; ui: SheetUI; a: SheetActions }) {
  const cat = (label: string, id: string, icon: ReactNode) => <RBigButton key={id} icon={icon} label={label} tip={`${label} functions`} onClick={() => a.insertFunction(id)} />;
  return (
    <>
      <RibbonGroup label="Function library">
        <RBigButton icon={<FunctionSquare />} label="Insert function" tip="Browse and search all functions (Shift+F3)" onClick={() => a.insertFunction()} />
        <RSplit big icon={<Sigma />} label="AutoSum" tip="AutoSum (Alt+=)" onClick={() => a.autoSum('SUM')} menu={['SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN'].map((f) => ({ label: f, onSelect: () => a.autoSum(f) }))} />
        {cat('Financial', 'Financial', <DollarSign />)}
        {cat('Logical', 'Logical', <Workflow />)}
        {cat('Text', 'Text', <Type />)}
        {cat('Date & time', 'Date & time', <BookA />)}
        {cat('Lookup', 'Lookup', <Search />)}
        {cat('Math', 'Math', <Calculator />)}
        {cat('Statistics', 'Statistical', <BarChart3 />)}
      </RibbonGroup>
      <RibbonGroup label="Defined names">
        <RBigButton icon={<BookA />} label="Name manager" tip="Create, edit and delete named ranges and LAMBDA functions" onClick={a.nameManager} />
      </RibbonGroup>
      <RibbonGroup label="Formula auditing">
        <RRows>
          <RButton icon={<Workflow size={I} />} label="Trace precedents" showLabel onClick={() => a.trace('precedents')} />
          <RButton icon={<Workflow size={I} />} label="Trace dependents" showLabel onClick={() => a.trace('dependents')} />
          <RButton icon={<Eye size={I} />} label="Show formulas" showLabel keys="Ctrl+`" active={ui.showFormulas} onClick={a.showFormulas} />
        </RRows>
      </RibbonGroup>
      <RibbonGroup label="Calculation">
        <RBigButton icon={<RefreshCw />} label="Calculate now" tip="Recalculate the workbook (F9)" onClick={a.calcNow} />
        <RDropdown label={doc.wb.calcAuto ? 'Automatic' : 'Manual'} tip="Calculation options" menu={[{ label: 'Automatic', checked: doc.wb.calcAuto, onSelect: () => a.setCalcAuto(true) }, { label: 'Manual', checked: !doc.wb.calcAuto, onSelect: () => a.setCalcAuto(false) }]} />
      </RibbonGroup>
    </>
  );
}

/* =================================================================== DATA */

function DataTab({ doc, a }: { doc: SheetDoc; a: SheetActions }) {
  return (
    <>
      <RibbonGroup label="Get data">
        <RBigButton icon={<DatabaseZap />} label="From text/CSV" tip="Import a CSV, TSV or text file into this sheet" onClick={a.importData} />
      </RibbonGroup>
      <RibbonGroup label="Sort & filter">
        <RRows>
          <RButton icon={<ArrowDownAZ size={I} />} label="Sort A → Z" showLabel onClick={() => a.sort(false)} />
          <RButton icon={<ArrowUpZA size={I} />} label="Sort Z → A" showLabel onClick={() => a.sort(true)} />
        </RRows>
        <RBigButton icon={<ArrowDownAZ />} label="Sort" tip="Sort by several columns" onClick={a.customSort} />
        <RBigButton icon={<Filter />} label="Filter" tip="Turn filter buttons on or off (Ctrl+Shift+L)" active={!!doc.sheet.filter} onClick={a.filter} />
        <RRows>
          <RButton icon={<FilterX size={I} />} label="Clear" showLabel disabled={!doc.sheet.filter} onClick={a.clearFilter} />
          <RButton icon={<RefreshCw size={I} />} label="Reapply" showLabel disabled={!doc.sheet.filter} onClick={a.reapplyFilter} />
        </RRows>
      </RibbonGroup>
      <RibbonGroup label="Data tools">
        <RBigButton icon={<SplitSquareHorizontal />} label="Text to columns" tip="Split one column into several" onClick={a.textToColumns} />
        <RBigButton icon={<WandSparkles />} label="Flash fill" tip="Fill the column by example (Ctrl+E)" onClick={a.flashFill} />
        <RBigButton icon={<CopyX />} label="Remove duplicates" tip="Delete duplicate rows" onClick={a.removeDuplicates} />
        <RBigButton icon={<ListChecks />} label="Validation" tip="Dropdown lists and input rules" onClick={a.validation} />
      </RibbonGroup>
      <RibbonGroup label="Forecast">
        <RBigButton icon={<Target />} label="Goal seek" tip="Find the input that gives the result you want" onClick={a.goalSeek} />
      </RibbonGroup>
    </>
  );
}

/* ================================================================= REVIEW */

function ReviewTab({ doc, a }: { doc: SheetDoc; a: SheetActions }) {
  return (
    <>
      <RibbonGroup label="Notes">
        <RBigButton icon={<MessageSquarePlus />} label="New note" tip="Add a note to the selected cell (Shift+F2)" onClick={a.insertNote} />
        <RBigButton icon={<Trash2 />} label="Delete notes" tip="Delete the notes in the selection" onClick={() => a.clear('notes')} />
      </RibbonGroup>
      <RibbonGroup label="Protect">
        <RBigButton icon={<Lock />} label={doc.sheet.protection?.enabled ? 'Unprotect' : 'Protect sheet'} tip="Prevent changes to locked cells" active={!!doc.sheet.protection?.enabled} onClick={a.protectSheet} />
      </RibbonGroup>
      <RibbonGroup label="Insights">
        <RBigButton icon={<Calculator />} label="Statistics" tip="Workbook statistics" onClick={a.stats} />
      </RibbonGroup>
    </>
  );
}

/* =================================================================== VIEW */

function ViewTab({ doc, ui, a }: { doc: SheetDoc; ui: SheetUI; a: SheetActions }) {
  const v = doc.sheet.view;
  return (
    <>
      <RibbonGroup label="Show">
        <RRows>
          <RButton icon={<Grid3x3 size={I} />} label="Gridlines" showLabel active={v.showGrid} onClick={() => a.toggleView('showGrid')} />
          <RButton icon={<Columns3 size={I} />} label="Headings" showLabel active={v.showHeaders} onClick={() => a.toggleView('showHeaders')} />
          <RButton icon={<FunctionSquare size={I} />} label="Formula bar" showLabel active={a.formulaBarOn} onClick={a.toggleFormulaBar} />
        </RRows>
        <RBigButton icon={<Eye />} label="Formulas" tip="Show formulas instead of results (Ctrl+`)" active={ui.showFormulas} onClick={a.showFormulas} />
      </RibbonGroup>
      <RibbonGroup label="Zoom">
        <RBigButton icon={<ZoomIn />} label="100%" tip="Zoom to 100%" onClick={() => a.zoom(1)} />
        <RBigButton icon={<Maximize2 />} label="Zoom to selection" tip="Zoom so the selection fills the window" onClick={() => a.zoom('selection')} />
      </RibbonGroup>
      <RibbonGroup label="Window">
        <RBigButton
          icon={<Snowflake />}
          label="Freeze panes"
          tip="Keep rows and columns visible while you scroll"
          active={v.freezeRows > 0 || v.freezeCols > 0}
          menu={[
            { label: 'Freeze panes', hint: 'Above and left of the active cell', onSelect: () => a.freeze('panes') },
            { label: 'Freeze top row', onSelect: () => a.freeze('row') },
            { label: 'Freeze first column', onSelect: () => a.freeze('col') },
            { label: 'Unfreeze panes', disabled: !(v.freezeRows || v.freezeCols), onSelect: () => a.freeze('none') },
          ]}
        />
      </RibbonGroup>
    </>
  );
}

/* ================================================================== CHART */

function ChartTab({ a }: { a: SheetActions }) {
  return (
    <>
      <RibbonGroup label="Type">
        <RBigButton icon={<BarChart3 />} label="Change type" tip="Change chart type" menu={CHART_TYPES.map((c) => ({ label: c.label, onSelect: () => a.chart('type', c.value) }))} />
      </RibbonGroup>
      <RibbonGroup label="Chart">
        <RBigButton icon={<Palette />} label="Edit chart" tip="Titles, legend, colours, data range" onClick={() => a.chart('edit')} />
        <RBigButton icon={<Trash2 />} label="Delete" tip="Delete the chart" onClick={() => a.chart('delete')} />
      </RibbonGroup>
    </>
  );
}

