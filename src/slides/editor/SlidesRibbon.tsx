import {
  AArrowDown,
  AArrowUp,
  AlignCenter,
  AlignCenterHorizontal,
  AlignCenterVertical,
  AlignEndHorizontal,
  AlignEndVertical,
  AlignHorizontalDistributeCenter,
  AlignJustify,
  AlignLeft,
  AlignRight,
  AlignStartHorizontal,
  AlignStartVertical,
  AlignVerticalDistributeCenter,
  AlignVerticalJustifyCenter,
  AlignVerticalJustifyEnd,
  AlignVerticalJustifyStart,
  ArrowDown,
  ArrowUp,
  Baseline,
  BetweenHorizontalStart,
  BetweenVerticalStart,
  Bold,
  BringToFront,
  Calendar,
  CaseSensitive,
  ChartColumn,
  ChartLine,
  ChartPie,
  ClipboardPaste,
  Columns2,
  Copy,
  Crop,
  Eye,
  EyeOff,
  FlipHorizontal2,
  FlipVertical2,
  Group,
  Hash,
  Highlighter,
  Image,
  Italic,
  LayoutGrid,
  LayoutTemplate,
  Link,
  List,
  ListIndentDecrease,
  ListIndentIncrease,
  ListOrdered,
  MonitorPlay,
  PaintBucket,
  Paintbrush,
  PanelBottom,
  PenLine,
  Play,
  Plus,
  Presentation,
  Ratio,
  Redo2,
  RemoveFormatting,
  Replace,
  RotateCcw,
  RotateCw,
  Rows3,
  Scissors,
  Search,
  SendToBack,
  Shapes,
  Smile,
  Sparkles,
  Square,
  StickyNote,
  Strikethrough,
  Subscript,
  Superscript,
  Table,
  TableCellsMerge,
  TableCellsSplit,
  Trash2,
  Type,
  Underline,
  Undo2,
  Ungroup,
  WandSparkles,
  ZoomIn,
} from 'lucide-react';
import { useState, useSyncExternalStore, type ReactNode } from 'react';
import { ColorSplitButton } from '@/ui/color';
import { Checkbox, NumberField, Select } from '@/ui/controls';
import { FontFamilyCombo, FontSizeCombo } from '@/ui/fontControls';
import type { MenuItem } from '@/ui/menu';
import { Ribbon, RibbonGroup, RBigButton, RButton, RDropdown, RRow, RRows, RSep, RSplit, type RibbonTab } from '@/ui/ribbon';
import type { SlidesDoc } from '../doc';
import { SHAPE_GALLERY, presetGeometry } from '../geometry';
import { BULLET_CHARS, resolveColor, type AnimClass, type AnimEffect, type Bullet, type Dir, type El, type Fill, type Line, type NumStyle, type Para, type Run, type SlideChartType, type TableStyleOpts, type ThemeColorName, type Transition, type TransitionType } from '../model';
import { DESIGNS, FONT_PAIRS, PALETTES } from '../themes';
import { colorProps } from './colors';
import { DesignThumb } from './DesignThumb';
import type { SlidesUI } from './ui';

export interface TextState {
  editing: boolean;
  font?: string;
  size?: number;
  color?: string;
  hl?: string;
  b: boolean;
  i: boolean;
  u: boolean;
  s: boolean;
  sup: boolean;
  sub: boolean;
  align?: Para['align'];
  bullet?: Bullet;
  level: number;
  lineSpacing?: number;
  anchor?: 't' | 'm' | 'b';
  hasText: boolean;
}

export type SelKind = 'none' | 'shape' | 'image' | 'table' | 'chart' | 'group' | 'multi';

export interface SlideActions {
  undo(): void;
  redo(): void;
  cut(): void;
  copy(): void;
  paste(): void;
  painter(): void;
  painterOn: boolean;
  newSlide(layoutId?: string): void;
  duplicateSlide(): void;
  deleteSlide(): void;
  setLayout(id: string): void;
  resetSlide(): void;
  hideSlide(): void;
  text: TextState;
  runStyle(patch: Partial<Omit<Run, 'text'>>): void;
  toggle(mark: 'b' | 'i' | 'u' | 's' | 'sup' | 'sub'): void;
  growFont(dir: 1 | -1): void;
  clearFormatting(): void;
  changeCase(mode: 'sentence' | 'lower' | 'upper' | 'title' | 'toggle'): void;
  align(a: NonNullable<Para['align']>): void;
  bullets(b: Bullet): void;
  level(dir: 1 | -1): void;
  lineSpacing(v: number): void;
  anchor(v: 't' | 'm' | 'b'): void;
  columns(n: number): void;
  insertShape(geom: string): void;
  insertTextBox(): void;
  insertTable(rows: number, cols: number): void;
  insertImage(): void;
  insertIcon(): void;
  insertChart(type: SlideChartType): void;
  insertLink(): void;
  insertField(kind: 'slidenum' | 'date'): void;
  headerFooter(): void;
  shapeFill(fill: Fill | null): void;
  shapeLine(patch: Partial<Line> | null): void;
  shadow(on: boolean): void;
  quickStyle(i: number): void;
  arrange(how: 'front' | 'back' | 'forward' | 'backward'): void;
  alignEls(how: 'left' | 'center' | 'right' | 'top' | 'middle' | 'bottom'): void;
  distribute(axis: 'h' | 'v'): void;
  group(): void;
  ungroup(): void;
  rotate(deg: number): void;
  flip(axis: 'h' | 'v'): void;
  setSize(w?: number, h?: number): void;
  changeShape(geom: string): void;
  formatShape(): void;
  design(id: string): void;
  palette(i: number): void;
  fonts(i: number): void;
  slideSize(kind: '16:9' | '4:3' | 'custom'): void;
  formatBackground(): void;
  transition(type: TransitionType): void;
  transitionPatch(p: Partial<Transition>): void;
  transitionAll(): void;
  previewTransition(): void;
  animate(cls: AnimClass, effect: AnimEffect): void;
  animPatch(p: { start?: 'click' | 'with' | 'after'; dur?: number; delay?: number; dir?: Dir }): void;
  removeAnim(): void;
  moveAnim(dir: -1 | 1): void;
  previewAnim(): void;
  animPaneOn: boolean;
  toggleAnimPane(): void;
  startShow(fromCurrent: boolean): void;
  presenter(): void;
  view(mode: 'normal' | 'sorter'): void;
  toggleNotes(): void;
  toggleGrid(): void;
  toggleSnap(): void;
  zoom(z: number | 'fit'): void;
  tableStyle(p: Partial<TableStyleOpts>): void;
  tableInsert(where: 'above' | 'below' | 'left' | 'right'): void;
  tableDelete(what: 'row' | 'col' | 'table'): void;
  mergeCells(): void;
  splitCell(): void;
  cellFill(fill: Fill | null): void;
  editChart(): void;
  chartType(t: SlideChartType): void;
  chartPatch(p: { legend?: 'bottom' | 'right' | 'top' | 'none'; dataLabels?: boolean; title?: string }): void;
  pictureShape(geom: string): void;
  crop(): void;
  resetPicture(): void;
  find(replace: boolean): void;
  selectAll(): void;
  selKind: SelKind;
  selected: El[];
  currentAnim?: { start: 'click' | 'with' | 'after'; dur: number; delay: number; dir?: Dir; effect: AnimEffect; cls: AnimClass };
  currentTransition?: Transition;
}

const I = 17;

/* ================================================================ icons */

export function ShapeIcon({ geom, size = 22 }: { geom: string; size?: number }) {
  const [g, variant] = geom.split(':');
  const line = /^(line|straightConnector|bentConnector|curvedConnector|arc)/.test(g) || /Bracket|Brace/i.test(g);
  const w = size;
  const h = size;
  const geo = presetGeometry(g, w - 4, h - 4);
  return (
    <svg width={w} height={h} viewBox={`-2 -2 ${w} ${h}`} className="shape-icon" aria-hidden>
      {variant && (
        <defs>
          <marker id={`mi-${geom}`} markerWidth="5" markerHeight="5" refX="4" refY="2.5" orient="auto-start-reverse">
            <path d="M0 0L5 2.5L0 5Z" fill="currentColor" />
          </marker>
        </defs>
      )}
      {geo.paths.map((p, i) => (
        <path
          key={i}
          d={p.d}
          fill={line || p.fill === 'none' ? 'none' : 'var(--shape-icon-fill)'}
          stroke="currentColor"
          strokeWidth={1.3}
          markerEnd={variant ? `url(#mi-${geom})` : undefined}
          markerStart={variant === 'double' ? `url(#mi-${geom})` : undefined}
        />
      ))}
    </svg>
  );
}

function ShapesPanel({ onPick, close }: { onPick: (g: string) => void; close: () => void }) {
  return (
    <div className="shape-gallery thin-scroll">
      {SHAPE_GALLERY.map((grp) => (
        <div key={grp.label} className="shape-group">
          <div className="shape-group-label">{grp.label}</div>
          <div className="shape-grid">
            {grp.shapes.map((s) => (
              <button
                key={s.geom}
                type="button"
                className="shape-cell"
                data-tip={s.label}
                onClick={() => {
                  onPick(s.geom);
                  close();
                }}
              >
                <ShapeIcon geom={s.geom} />
              </button>
            ))}
          </div>
        </div>
      ))}
    </div>
  );
}

function TableGridPicker({ onPick, close }: { onPick: (r: number, c: number) => void; close: () => void }) {
  const [hover, setHover] = useState<[number, number]>([0, 0]);
  return (
    <div className="table-picker">
      <div className="table-picker-label">{hover[0] ? `${hover[1]} × ${hover[0]} table` : 'Insert table'}</div>
      <div className="table-picker-grid" onMouseLeave={() => setHover([0, 0])}>
        {Array.from({ length: 8 }, (_, r) =>
          Array.from({ length: 10 }, (__, c) => (
            <button
              key={`${r}-${c}`}
              type="button"
              className={`table-picker-cell${r < hover[0] && c < hover[1] ? ' on' : ''}`}
              onMouseEnter={() => setHover([r + 1, c + 1])}
              onClick={() => {
                onPick(r + 1, c + 1);
                close();
              }}
            />
          )),
        )}
      </div>
    </div>
  );
}

/* ================================================================ ribbon */

export function SlidesRibbon({ doc, ui, a, onFile }: { doc: SlidesDoc; ui: SlidesUI; a: SlideActions; onFile: () => void }) {
  useSyncExternalStore(doc.subscribe, doc.getVersion);
  useSyncExternalStore(ui.subscribe, ui.getVersion);
  const [active, setActive] = useState('home');
  const k = a.selKind;
  const shapeLike = k === 'shape' || k === 'group' || k === 'multi';
  const tabs: RibbonTab[] = [
    { id: 'home', label: 'Home', content: <HomeTab doc={doc} a={a} /> },
    { id: 'insert', label: 'Insert', content: <InsertTab doc={doc} a={a} /> },
    { id: 'design', label: 'Design', content: <DesignTab doc={doc} a={a} /> },
    { id: 'transitions', label: 'Transitions', content: <TransitionsTab a={a} /> },
    { id: 'animations', label: 'Animations', content: <AnimationsTab a={a} /> },
    { id: 'show', label: 'Slide Show', content: <ShowTab doc={doc} a={a} /> },
    { id: 'view', label: 'View', content: <ViewTab ui={ui} a={a} /> },
    { id: 'shape', label: 'Shape Format', content: <ShapeTab doc={doc} a={a} />, contextual: true, hidden: !shapeLike },
    { id: 'picture', label: 'Picture Format', content: <PictureTab doc={doc} a={a} />, contextual: true, hidden: k !== 'image' },
    { id: 'table', label: 'Table Design', content: <TableTab doc={doc} a={a} />, contextual: true, hidden: k !== 'table' },
    { id: 'chart', label: 'Chart', content: <ChartTab a={a} />, contextual: true, hidden: k !== 'chart' },
  ];
  const ctxIds = new Set(['shape', 'picture', 'table', 'chart']);
  const effective = tabs.find((t) => t.id === active && !t.hidden) ? active : ctxIds.has(active) ? 'home' : active;
  return (
    <Ribbon
      app="slides"
      tabs={tabs}
      active={effective}
      onActive={setActive}
      onFile={onFile}
      right={
        <div className="ribbon-right">
          <button type="button" className="icon-btn icon-btn-sm" data-tip="Undo" data-tip-key="Ctrl+Z" disabled={!doc.canUndo() && !ui.textEditor} onMouseDown={(e) => e.preventDefault()} onClick={a.undo}>
            <Undo2 size={16} />
          </button>
          <button type="button" className="icon-btn icon-btn-sm" data-tip="Redo" data-tip-key="Ctrl+Y" disabled={!doc.canRedo() && !ui.textEditor} onMouseDown={(e) => e.preventDefault()} onClick={a.redo}>
            <Redo2 size={16} />
          </button>
          <button type="button" className="btn btn-sm btn-accent ribbon-present" data-tip="Start the slide show from the beginning" data-tip-key="F5" onMouseDown={(e) => e.preventDefault()} onClick={() => a.startShow(false)}>
            <Play size={14} /> Present
          </button>
        </div>
      }
    />
  );
}

/* ================================================================== HOME */

function layoutMenu(doc: SlidesDoc, pick: (id: string) => void): MenuItem[] {
  return doc.pres.layouts.map((l) => ({ label: l.name, onSelect: () => pick(l.id), checked: doc.slide?.layout === l.id }));
}

function FontGroup({ doc, a }: { doc: SlidesDoc; a: SlideActions }) {
  const t = a.text;
  const cp = colorProps(doc.pres.theme);
  return (
    <RibbonGroup label="Font">
      <RRows>
        <RRow>
          <FontFamilyCombo value={t.font ?? '+minor'} theme={doc.pres.theme.fonts} onChange={(f) => a.runStyle({ font: f })} />
          <FontSizeCombo value={t.size ?? null} onChange={(pt) => a.runStyle({ size: pt })} />
          <RButton icon={<AArrowUp size={I} />} tip="Increase font size" keys="Ctrl+]" onClick={() => a.growFont(1)} />
          <RButton icon={<AArrowDown size={I} />} tip="Decrease font size" keys="Ctrl+[" onClick={() => a.growFont(-1)} />
          <RButton icon={<RemoveFormatting size={I} />} tip="Clear formatting" keys="Ctrl+Space" onClick={a.clearFormatting} />
        </RRow>
        <RRow>
          <RButton icon={<Bold size={I} />} tip="Bold" keys="Ctrl+B" active={t.b} onClick={() => a.toggle('b')} />
          <RButton icon={<Italic size={I} />} tip="Italic" keys="Ctrl+I" active={t.i} onClick={() => a.toggle('i')} />
          <RButton icon={<Underline size={I} />} tip="Underline" keys="Ctrl+U" active={t.u} onClick={() => a.toggle('u')} />
          <RButton icon={<Strikethrough size={I} />} tip="Strikethrough" active={t.s} onClick={() => a.toggle('s')} />
          <RButton icon={<Subscript size={I} />} tip="Subscript" keys="Ctrl+=" active={t.sub} onClick={() => a.toggle('sub')} />
          <RButton icon={<Superscript size={I} />} tip="Superscript" keys="Ctrl+Shift++" active={t.sup} onClick={() => a.toggle('sup')} />
          <RDropdown
            icon={<CaseSensitive size={I} />}
            tip="Change case"
            showLabel={false}
            menu={[
              { label: 'Sentence case.', onSelect: () => a.changeCase('sentence') },
              { label: 'lowercase', onSelect: () => a.changeCase('lower') },
              { label: 'UPPERCASE', onSelect: () => a.changeCase('upper') },
              { label: 'Capitalize Each Word', onSelect: () => a.changeCase('title') },
              { label: 'tOGGLE cASE', onSelect: () => a.changeCase('toggle') },
            ]}
          />
          <ColorSplitButton icon={<Highlighter size={I} />} color={t.hl ?? '#ffff00'} palette="highlight" tip="Text highlight colour" noneLabel="No colour" onApply={(c) => a.runStyle({ hl: c ?? undefined })} />
          <ColorSplitButton icon={<Baseline size={I} />} color={t.color ?? '#e5484d'} tip="Font colour" noneLabel="Automatic" onApply={(c) => a.runStyle({ color: c ?? undefined })} {...cp} />
        </RRow>
      </RRows>
    </RibbonGroup>
  );
}

const NUM_STYLES: { style: NumStyle; label: string }[] = [
  { style: 'arabicPeriod', label: '1. 2. 3.' },
  { style: 'arabicParenR', label: '1) 2) 3)' },
  { style: 'alphaUcPeriod', label: 'A. B. C.' },
  { style: 'alphaLcPeriod', label: 'a. b. c.' },
  { style: 'alphaLcParenR', label: 'a) b) c)' },
  { style: 'romanUcPeriod', label: 'I. II. III.' },
  { style: 'romanLcPeriod', label: 'i. ii. iii.' },
];

function ParagraphGroup({ a }: { a: SlideActions }) {
  const t = a.text;
  const isChar = t.bullet?.type === 'char';
  const isNum = t.bullet?.type === 'num';
  return (
    <RibbonGroup label="Paragraph" collapse={2} icon={<AlignLeft />}>
      <RRows>
        <RRow>
          <RSplit
            icon={<List size={I} />}
            tip="Bullets"
            active={isChar}
            onClick={() => a.bullets(isChar ? { type: 'none' } : { type: 'char', char: '•' })}
            menu={[{ label: 'None', onSelect: () => a.bullets({ type: 'none' }) }, ...BULLET_CHARS.map((c) => ({ label: <span className="bullet-choice">{c}</span>, onSelect: () => a.bullets({ type: 'char', char: c }) }))]}
          />
          <RSplit
            icon={<ListOrdered size={I} />}
            tip="Numbering"
            active={isNum}
            onClick={() => a.bullets(isNum ? { type: 'none' } : { type: 'num', style: 'arabicPeriod' })}
            menu={[{ label: 'None', onSelect: () => a.bullets({ type: 'none' }) }, ...NUM_STYLES.map((n) => ({ label: n.label, onSelect: () => a.bullets({ type: 'num', style: n.style }) }))]}
          />
          <RButton icon={<ListIndentDecrease size={I} />} tip="Decrease list level" keys="Shift+Tab" onClick={() => a.level(-1)} />
          <RButton icon={<ListIndentIncrease size={I} />} tip="Increase list level" keys="Tab" onClick={() => a.level(1)} />
          <RDropdown
            icon={<Rows3 size={I} />}
            tip="Line spacing"
            showLabel={false}
            menu={[0.9, 1, 1.15, 1.5, 2, 2.5, 3].map((v) => ({ label: v.toFixed(v % 1 ? 2 : 1).replace(/0$/, ''), checked: Math.abs((t.lineSpacing ?? 0) - v) < 0.01, onSelect: () => a.lineSpacing(v) }))}
          />
          <RDropdown
            icon={<Columns2 size={I} />}
            tip="Columns"
            showLabel={false}
            menu={[1, 2, 3].map((n) => ({ label: `${n} column${n > 1 ? 's' : ''}`, onSelect: () => a.columns(n) }))}
          />
        </RRow>
        <RRow>
          <RButton icon={<AlignLeft size={I} />} tip="Align left" keys="Ctrl+L" active={!t.align || t.align === 'left'} onClick={() => a.align('left')} />
          <RButton icon={<AlignCenter size={I} />} tip="Centre" keys="Ctrl+E" active={t.align === 'center'} onClick={() => a.align('center')} />
          <RButton icon={<AlignRight size={I} />} tip="Align right" keys="Ctrl+R" active={t.align === 'right'} onClick={() => a.align('right')} />
          <RButton icon={<AlignJustify size={I} />} tip="Justify" keys="Ctrl+J" active={t.align === 'justify'} onClick={() => a.align('justify')} />
          <RSep />
          <RButton icon={<AlignVerticalJustifyStart size={I} />} tip="Align text to top" active={t.anchor === 't'} onClick={() => a.anchor('t')} />
          <RButton icon={<AlignVerticalJustifyCenter size={I} />} tip="Align text to middle" active={t.anchor === 'm'} onClick={() => a.anchor('m')} />
          <RButton icon={<AlignVerticalJustifyEnd size={I} />} tip="Align text to bottom" active={t.anchor === 'b'} onClick={() => a.anchor('b')} />
        </RRow>
      </RRows>
    </RibbonGroup>
  );
}

function arrangeMenu(a: SlideActions): MenuItem[] {
  return [
    { header: 'Order objects' },
    { label: 'Bring to front', icon: <BringToFront size={15} />, shortcut: 'Ctrl+Shift+]', onSelect: () => a.arrange('front') },
    { label: 'Send to back', icon: <SendToBack size={15} />, shortcut: 'Ctrl+Shift+[', onSelect: () => a.arrange('back') },
    { label: 'Bring forward', icon: <ArrowUp size={15} />, shortcut: 'Ctrl+]', onSelect: () => a.arrange('forward') },
    { label: 'Send backward', icon: <ArrowDown size={15} />, shortcut: 'Ctrl+[', onSelect: () => a.arrange('backward') },
    { separator: true },
    { header: 'Group objects' },
    { label: 'Group', icon: <Group size={15} />, shortcut: 'Ctrl+G', onSelect: a.group, disabled: a.selected.length < 2 },
    { label: 'Ungroup', icon: <Ungroup size={15} />, shortcut: 'Ctrl+Shift+G', onSelect: a.ungroup, disabled: !a.selected.some((e) => e.type === 'group') },
    { separator: true },
    {
      label: 'Align',
      submenu: [
        { label: 'Align left', icon: <AlignStartVertical size={15} />, onSelect: () => a.alignEls('left') },
        { label: 'Align centre', icon: <AlignCenterVertical size={15} />, onSelect: () => a.alignEls('center') },
        { label: 'Align right', icon: <AlignEndVertical size={15} />, onSelect: () => a.alignEls('right') },
        { label: 'Align top', icon: <AlignStartHorizontal size={15} />, onSelect: () => a.alignEls('top') },
        { label: 'Align middle', icon: <AlignCenterHorizontal size={15} />, onSelect: () => a.alignEls('middle') },
        { label: 'Align bottom', icon: <AlignEndHorizontal size={15} />, onSelect: () => a.alignEls('bottom') },
        { separator: true },
        { label: 'Distribute horizontally', icon: <AlignHorizontalDistributeCenter size={15} />, onSelect: () => a.distribute('h'), disabled: a.selected.length < 3 },
        { label: 'Distribute vertically', icon: <AlignVerticalDistributeCenter size={15} />, onSelect: () => a.distribute('v'), disabled: a.selected.length < 3 },
      ],
    },
    {
      label: 'Rotate',
      submenu: [
        { label: 'Rotate right 90°', icon: <RotateCw size={15} />, onSelect: () => a.rotate(90) },
        { label: 'Rotate left 90°', icon: <RotateCcw size={15} />, onSelect: () => a.rotate(-90) },
        { label: 'Flip vertical', icon: <FlipVertical2 size={15} />, onSelect: () => a.flip('v') },
        { label: 'Flip horizontal', icon: <FlipHorizontal2 size={15} />, onSelect: () => a.flip('h') },
      ],
    },
  ];
}

export const QUICK_STYLES: { fill: string | null; line: string | null; text: string; label: string }[] = [
  { fill: '@accent1', line: '@accent1-25', text: '@bg1', label: 'Colored fill – Accent 1' },
  { fill: '@accent2', line: '@accent2-25', text: '@bg1', label: 'Colored fill – Accent 2' },
  { fill: '@accent3', line: '@accent3-25', text: '@bg1', label: 'Colored fill – Accent 3' },
  { fill: '@accent4', line: '@accent4-25', text: '@bg1', label: 'Colored fill – Accent 4' },
  { fill: '@accent5', line: '@accent5-25', text: '@bg1', label: 'Colored fill – Accent 5' },
  { fill: '@accent6', line: '@accent6-25', text: '@bg1', label: 'Colored fill – Accent 6' },
  { fill: '@accent1+80', line: '@accent1', text: '@accent1-50', label: 'Light fill – Accent 1' },
  { fill: '@accent2+80', line: '@accent2', text: '@accent2-50', label: 'Light fill – Accent 2' },
  { fill: '@accent3+80', line: '@accent3', text: '@accent3-50', label: 'Light fill – Accent 3' },
  { fill: '@accent4+80', line: '@accent4', text: '@accent4-50', label: 'Light fill – Accent 4' },
  { fill: '@accent5+80', line: '@accent5', text: '@accent5-50', label: 'Light fill – Accent 5' },
  { fill: '@accent6+80', line: '@accent6', text: '@accent6-50', label: 'Light fill – Accent 6' },
  { fill: null, line: '@accent1', text: '@accent1', label: 'Outline – Accent 1' },
  { fill: null, line: '@accent2', text: '@accent2', label: 'Outline – Accent 2' },
  { fill: null, line: '@tx1', text: '@tx1', label: 'Outline – Dark' },
  { fill: '@tx1', line: null, text: '@bg1', label: 'Dark fill' },
  { fill: '@bg2', line: null, text: '@tx1', label: 'Subtle fill' },
  { fill: null, line: null, text: '@tx1', label: 'No fill, no line' },
];


function QuickStyles({ doc, a, close }: { doc: SlidesDoc; a: SlideActions; close: () => void }) {
  const theme = doc.pres.theme;
  return (
    <div className="quick-styles">
      {QUICK_STYLES.map((q, i) => (
        <button
          key={i}
          type="button"
          className="quick-style"
          data-tip={q.label}
          style={{ background: q.fill ? resolveColor(q.fill, theme) : resolveColor('@bg1', theme), borderColor: q.line ? resolveColor(q.line, theme) : 'transparent', color: resolveColor(q.text, theme) }}
          onClick={() => {
            a.quickStyle(i);
            close();
          }}
        >
          Abc
        </button>
      ))}
    </div>
  );
}

function DrawingGroup({ doc, a, fillOnly }: { doc: SlidesDoc; a: SlideActions; fillOnly?: boolean }) {
  const cp = colorProps(doc.pres.theme);
  const sel = a.selected.find((e) => e.type === 'shape');
  const fillColor = sel && sel.type === 'shape' && sel.fill?.type === 'solid' ? sel.fill.color : '@accent1';
  const lineColor = sel && sel.type === 'shape' && sel.line ? sel.line.color : '@accent1-25';
  return (
    <RibbonGroup label={fillOnly ? 'Fill' : 'Drawing'} collapse={fillOnly ? undefined : 1} icon={<Shapes />}>
      {!fillOnly && (
        <>
          <RDropdown icon={<Shapes size={I} />} label="Shapes" tip="Insert a shape" panel={(close) => <ShapesPanel onPick={a.insertShape} close={close} />} />
          <RDropdown icon={<LayoutGrid size={I} />} label="Arrange" tip="Order, group, align and rotate objects" menu={() => arrangeMenu(a)} />
          <RDropdown icon={<WandSparkles size={I} />} label="Styles" tip="Quick shape styles" panel={(close) => <QuickStyles doc={doc} a={a} close={close} />} />
        </>
      )}
      <RRows>
        <ColorSplitButton icon={<PaintBucket size={I} />} color={fillColor} tip="Shape fill" noneLabel="No fill" onApply={(c) => a.shapeFill(c ? { type: 'solid', color: c } : { type: 'none' })} {...cp} extra={(close) => <button className="menu-extra-btn" onClick={() => (close(), a.formatShape())}>More fill options…</button>} />
        <ColorSplitButton icon={<PenLine size={I} />} color={lineColor} tip="Shape outline" noneLabel="No outline" onApply={(c) => a.shapeLine(c ? { color: c } : null)} {...cp} extra={(close) => <LineOptions a={a} close={close} />} />
      </RRows>
    </RibbonGroup>
  );
}

function LineOptions({ a, close }: { a: SlideActions; close: () => void }) {
  return (
    <div className="line-options">
      <div className="color-section">Weight</div>
      <div className="line-weights">
        {[0.75, 1, 1.5, 2.25, 3, 4.5, 6].map((pt) => (
          <button key={pt} type="button" onClick={() => (a.shapeLine({ width: (pt * 4) / 3 }), close())}>
            <span style={{ height: Math.max(1, (pt * 4) / 3) }} /> {pt} pt
          </button>
        ))}
      </div>
      <div className="color-section">Dashes</div>
      <div className="line-weights">
        {(['solid', 'dash', 'dot', 'dashDot', 'longDash'] as const).map((d) => (
          <button key={d} type="button" onClick={() => (a.shapeLine({ dash: d }), close())}>
            <span className={`dash-sample dash-${d}`} /> {d === 'dashDot' ? 'Dash dot' : d === 'longDash' ? 'Long dash' : d[0].toUpperCase() + d.slice(1)}
          </button>
        ))}
      </div>
      <div className="color-section">Arrows</div>
      <div className="line-weights">
        <button type="button" onClick={() => (a.shapeLine({ head: 'none', tail: 'none' }), close())}>None</button>
        <button type="button" onClick={() => (a.shapeLine({ head: 'none', tail: 'triangle' }), close())}>End arrow</button>
        <button type="button" onClick={() => (a.shapeLine({ head: 'triangle', tail: 'triangle' }), close())}>Both ends</button>
        <button type="button" onClick={() => (a.shapeLine({ head: 'oval', tail: 'triangle' }), close())}>Dot to arrow</button>
      </div>
    </div>
  );
}

function HomeTab({ doc, a }: { doc: SlidesDoc; a: SlideActions }) {
  return (
    <>
      <RibbonGroup label="Clipboard">
        <RBigButton icon={<ClipboardPaste />} label="Paste" tip="Paste (Ctrl+V)" onClick={a.paste} />
        <RRows>
          <RButton icon={<Scissors size={I} />} label="Cut" keys="Ctrl+X" onClick={a.cut} />
          <RButton icon={<Copy size={I} />} label="Copy" keys="Ctrl+C" onClick={a.copy} />
          <RButton icon={<Paintbrush size={I} />} label="Format painter" tip="Format painter — copy formatting to another object" active={a.painterOn} onClick={a.painter} />
        </RRows>
      </RibbonGroup>
      <RibbonGroup label="Slides">
        <RSplit big icon={<Plus />} label="New slide" tip="New slide (Ctrl+M)" onClick={() => a.newSlide()} menu={() => layoutMenu(doc, (id) => a.newSlide(id))} />
        <RRows>
          <RDropdown icon={<LayoutTemplate size={I} />} label="Layout" tip="Change the slide layout" menu={() => layoutMenu(doc, a.setLayout)} />
          <RButton icon={<RotateCcw size={I} />} label="Reset" showLabel tip="Reset placeholders to the layout" onClick={a.resetSlide} />
          <RButton icon={<EyeOff size={I} />} label="Hide" showLabel tip="Hide the slide in the slide show" active={!!doc.slide?.hidden} onClick={a.hideSlide} />
        </RRows>
      </RibbonGroup>
      <FontGroup doc={doc} a={a} />
      <ParagraphGroup a={a} />
      <DrawingGroup doc={doc} a={a} />
      <RibbonGroup label="Editing" collapse={1} icon={<Search />}>
        <RRows>
          <RButton icon={<Search size={I} />} label="Find" showLabel keys="Ctrl+F" onClick={() => a.find(false)} />
          <RButton icon={<Replace size={I} />} label="Replace" showLabel keys="Ctrl+H" onClick={() => a.find(true)} />
          <RButton icon={<Square size={I} />} label="Select all" showLabel keys="Ctrl+A" onClick={a.selectAll} />
        </RRows>
      </RibbonGroup>
    </>
  );
}

/* ================================================================ INSERT */

const CHARTS: { type: SlideChartType; label: string; icon: ReactNode }[] = [
  { type: 'column', label: 'Column', icon: <ChartColumn size={15} /> },
  { type: 'bar', label: 'Bar', icon: <ChartColumn size={15} style={{ transform: 'rotate(90deg)' }} /> },
  { type: 'line', label: 'Line', icon: <ChartLine size={15} /> },
  { type: 'area', label: 'Area', icon: <ChartLine size={15} /> },
  { type: 'pie', label: 'Pie', icon: <ChartPie size={15} /> },
  { type: 'doughnut', label: 'Doughnut', icon: <ChartPie size={15} /> },
  { type: 'scatter', label: 'Scatter', icon: <ChartLine size={15} /> },
  { type: 'radar', label: 'Radar', icon: <ChartPie size={15} /> },
];

function InsertTab({ doc, a }: { doc: SlidesDoc; a: SlideActions }) {
  return (
    <>
      <RibbonGroup label="Slides">
        <RSplit big icon={<Plus />} label="New slide" tip="New slide (Ctrl+M)" onClick={() => a.newSlide()} menu={() => layoutMenu(doc, (id) => a.newSlide(id))} />
      </RibbonGroup>
      <RibbonGroup label="Tables">
        <RDropdown icon={<Table size={I} />} label="Table" tip="Insert a table" panel={(close) => <TableGridPicker onPick={a.insertTable} close={close} />} />
      </RibbonGroup>
      <RibbonGroup label="Images">
        <RBigButton icon={<Image />} label="Pictures" tip="Insert a picture from your computer" onClick={a.insertImage} />
        <RBigButton icon={<Smile />} label="Icons" tip="Insert a vector icon (over 1,500 to choose from)" onClick={a.insertIcon} />
      </RibbonGroup>
      <RibbonGroup label="Illustrations">
        <RDropdown icon={<Shapes size={I} />} label="Shapes" tip="Insert a shape" panel={(close) => <ShapesPanel onPick={a.insertShape} close={close} />} />
        <RDropdown icon={<ChartColumn size={I} />} label="Chart" tip="Insert a chart" menu={CHARTS.map((c) => ({ label: c.label, icon: c.icon, onSelect: () => a.insertChart(c.type) }))} />
      </RibbonGroup>
      <RibbonGroup label="Links">
        <RBigButton icon={<Link />} label="Link" tip="Link to a web page or another slide (Ctrl+K)" onClick={a.insertLink} />
      </RibbonGroup>
      <RibbonGroup label="Text">
        <RBigButton icon={<Type />} label="Text box" tip="Draw a text box" onClick={a.insertTextBox} />
        <RRows>
          <RButton icon={<PanelBottom size={I} />} label="Header & footer" showLabel onClick={a.headerFooter} />
          <RButton icon={<Hash size={I} />} label="Slide number" showLabel onClick={() => a.insertField('slidenum')} />
          <RButton icon={<Calendar size={I} />} label="Date & time" showLabel onClick={() => a.insertField('date')} />
        </RRows>
      </RibbonGroup>
    </>
  );
}

/* ================================================================ DESIGN */

function DesignTab({ doc, a }: { doc: SlidesDoc; a: SlideActions }) {
  return (
    <>
      <RibbonGroup label="Themes" className="design-group">
        <div className="design-gallery thin-scroll">
          {DESIGNS.map((d) => (
            <button key={d.id} type="button" className={`design-item${doc.pres.design === d.id ? ' active' : ''}`} data-tip={d.name} onMouseDown={(e) => e.preventDefault()} onClick={() => a.design(d.id)}>
              <DesignThumb id={d.id} />
            </button>
          ))}
        </div>
      </RibbonGroup>
      <RibbonGroup label="Variants">
        <RRows>
          <RDropdown
            icon={<Sparkles size={I} />}
            label="Colours"
            tip="Theme colours"
            menu={PALETTES.map((p, i) => ({
              label: (
                <span className="palette-row">
                  {(['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'] as ThemeColorName[]).map((n) => (
                    <i key={n} style={{ background: p.colors[n] }} />
                  ))}
                  {p.name}
                </span>
              ),
              onSelect: () => a.palette(i),
            }))}
          />
          <RDropdown
            icon={<Type size={I} />}
            label="Fonts"
            tip="Theme fonts"
            menu={FONT_PAIRS.map((f, i) => ({
              label: (
                <span className="font-pair">
                  <b style={{ fontFamily: f.major }}>{f.major}</b> <span style={{ fontFamily: f.minor }}>{f.minor}</span>
                </span>
              ),
              checked: doc.pres.theme.fonts.major === f.major && doc.pres.theme.fonts.minor === f.minor,
              onSelect: () => a.fonts(i),
            }))}
          />
        </RRows>
      </RibbonGroup>
      <RibbonGroup label="Customise">
        <RDropdown
          icon={<Ratio size={I} />}
          label="Slide size"
          tip="Widescreen, standard or a custom size"
          menu={[
            { label: 'Widescreen (16:9)', checked: Math.abs(doc.pres.size.w / doc.pres.size.h - 16 / 9) < 0.01, onSelect: () => a.slideSize('16:9') },
            { label: 'Standard (4:3)', checked: Math.abs(doc.pres.size.w / doc.pres.size.h - 4 / 3) < 0.01, onSelect: () => a.slideSize('4:3') },
            { separator: true },
            { label: 'Custom slide size…', onSelect: () => a.slideSize('custom') },
          ]}
        />
        <RBigButton icon={<PaintBucket />} label="Format background" tip="Colour, gradient or picture background" onClick={a.formatBackground} />
      </RibbonGroup>
    </>
  );
}

/* =========================================================== TRANSITIONS */

export const TRANSITIONS: { type: TransitionType; label: string; dirs?: Dir[] }[] = [
  { type: 'none', label: 'None' },
  { type: 'fade', label: 'Fade' },
  { type: 'push', label: 'Push', dirs: ['u', 'd', 'l', 'r'] },
  { type: 'wipe', label: 'Wipe', dirs: ['l', 'r', 'u', 'd'] },
  { type: 'split', label: 'Split', dirs: ['l', 'u'] },
  { type: 'reveal', label: 'Reveal', dirs: ['l', 'r'] },
  { type: 'cover', label: 'Cover', dirs: ['l', 'r', 'u', 'd'] },
  { type: 'zoom', label: 'Zoom' },
  { type: 'morph', label: 'Morph' },
  { type: 'circle', label: 'Circle' },
  { type: 'dissolve', label: 'Dissolve' },
  { type: 'flip', label: 'Flip', dirs: ['l', 'r'] },
];

const DIR_LABEL: Record<Dir, string> = { l: 'From right', r: 'From left', u: 'From bottom', d: 'From top' };

function TransitionIcon({ type }: { type: TransitionType }) {
  return <span className={`fx-icon fx-${type}`} aria-hidden />;
}

function TransitionsTab({ a }: { a: SlideActions }) {
  const t = a.currentTransition;
  const def = TRANSITIONS.find((x) => x.type === (t?.type ?? 'none'));
  return (
    <>
      <RibbonGroup label="Preview">
        <RBigButton icon={<Eye />} label="Preview" tip="Play the transition on this slide" onClick={a.previewTransition} disabled={!t || t.type === 'none'} />
      </RibbonGroup>
      <RibbonGroup label="Transition to this slide" className="fx-group">
        <div className="fx-gallery thin-scroll">
          {TRANSITIONS.map((x) => (
            <button key={x.type} type="button" className={`fx-item${(t?.type ?? 'none') === x.type ? ' active' : ''}`} onClick={() => a.transition(x.type)}>
              <TransitionIcon type={x.type} />
              <span>{x.label}</span>
            </button>
          ))}
        </div>
        <RDropdown label="Effect options" tip="Direction of the transition" menu={(def?.dirs ?? []).map((d) => ({ label: DIR_LABEL[d], checked: t?.dir === d, onSelect: () => a.transitionPatch({ dir: d }) }))} />
      </RibbonGroup>
      <RibbonGroup label="Timing">
        <RRows>
          <label className="ribbon-field">
            <span>Duration</span>
            <NumberField value={(t?.dur ?? 700) / 1000} min={0.1} max={10} step={0.1} precision={2} width={70} suffix="s" onChange={(v) => a.transitionPatch({ dur: Math.round(v * 1000) })} />
          </label>
          <button type="button" className="btn btn-sm" onMouseDown={(e) => e.preventDefault()} onClick={a.transitionAll}>
            Apply to all
          </button>
        </RRows>
        <RRows>
          <Checkbox checked={t?.onClick !== false} onChange={(v) => a.transitionPatch({ onClick: v })} label="On mouse click" />
          <label className="ribbon-field">
            <Checkbox checked={!!t?.after} onChange={(v) => a.transitionPatch({ after: v ? 5000 : undefined })} label="After" />
            <NumberField value={(t?.after ?? 5000) / 1000} min={0} max={3600} step={0.5} precision={2} width={70} suffix="s" onChange={(v) => a.transitionPatch({ after: Math.round(v * 1000) })} />
          </label>
        </RRows>
      </RibbonGroup>
    </>
  );
}

/* ============================================================ ANIMATIONS */

export const ANIMATIONS: { cls: AnimClass; effect: AnimEffect; label: string; dirs?: Dir[] }[] = [
  { cls: 'entr', effect: 'appear', label: 'Appear' },
  { cls: 'entr', effect: 'fade', label: 'Fade' },
  { cls: 'entr', effect: 'fly', label: 'Fly in', dirs: ['d', 'u', 'l', 'r'] },
  { cls: 'entr', effect: 'float', label: 'Float in', dirs: ['u', 'd'] },
  { cls: 'entr', effect: 'split', label: 'Split' },
  { cls: 'entr', effect: 'wipe', label: 'Wipe', dirs: ['d', 'u', 'l', 'r'] },
  { cls: 'entr', effect: 'zoom', label: 'Zoom' },
  { cls: 'entr', effect: 'wheel', label: 'Wheel' },
  { cls: 'entr', effect: 'bounce', label: 'Bounce' },
  { cls: 'emph', effect: 'pulse', label: 'Pulse' },
  { cls: 'emph', effect: 'spin', label: 'Spin' },
  { cls: 'emph', effect: 'grow', label: 'Grow/Shrink' },
  { cls: 'emph', effect: 'teeter', label: 'Teeter' },
  { cls: 'emph', effect: 'transparency', label: 'Transparency' },
  { cls: 'exit', effect: 'disappear', label: 'Disappear' },
  { cls: 'exit', effect: 'fade', label: 'Fade out' },
  { cls: 'exit', effect: 'fly', label: 'Fly out', dirs: ['d', 'u', 'l', 'r'] },
  { cls: 'exit', effect: 'zoom', label: 'Zoom out' },
  { cls: 'exit', effect: 'wipe', label: 'Wipe out', dirs: ['d', 'u', 'l', 'r'] },
];

const ANIM_DIR: Record<Dir, string> = { d: 'From bottom', u: 'From top', l: 'From left', r: 'From right' };

function AnimationsTab({ a }: { a: SlideActions }) {
  const cur = a.currentAnim;
  const hasSel = a.selected.length > 0;
  const def = ANIMATIONS.find((x) => x.cls === cur?.cls && x.effect === cur?.effect);
  return (
    <>
      <RibbonGroup label="Preview">
        <RBigButton icon={<Eye />} label="Preview" tip="Play this slide's animations" onClick={a.previewAnim} />
      </RibbonGroup>
      <RibbonGroup label="Animation" className="fx-group">
        <div className="fx-gallery anim thin-scroll">
          <button type="button" className={`fx-item${!cur ? ' active' : ''}`} disabled={!hasSel} onClick={a.removeAnim}>
            <span className="fx-icon fx-none" />
            <span>None</span>
          </button>
          {ANIMATIONS.map((x) => (
            <button key={`${x.cls}-${x.effect}`} type="button" disabled={!hasSel} className={`fx-item anim-${x.cls}${cur?.cls === x.cls && cur?.effect === x.effect ? ' active' : ''}`} onClick={() => a.animate(x.cls, x.effect)}>
              <span className={`fx-icon anim-icon anim-${x.cls}`}>
                <Sparkles size={16} />
              </span>
              <span>{x.label}</span>
            </button>
          ))}
        </div>
        <RDropdown label="Effect options" tip="Direction" menu={(def?.dirs ?? []).map((d) => ({ label: ANIM_DIR[d], checked: cur?.dir === d, onSelect: () => a.animPatch({ dir: d }) }))} />
      </RibbonGroup>
      <RibbonGroup label="Timing">
        <RRows>
          <label className={`ribbon-field${cur ? '' : ' is-disabled'}`}>
            <span>Start</span>
            <Select
              label="Start"
              width={140}
              value={cur?.start ?? 'click'}
              onChange={(v) => cur && a.animPatch({ start: v })}
              options={[
                { value: 'click', label: 'On click' },
                { value: 'with', label: 'With previous' },
                { value: 'after', label: 'After previous' },
              ]}
            />
          </label>
          <label className="ribbon-field">
            <span>Duration</span>
            <NumberField value={(cur?.dur ?? 500) / 1000} min={0.05} max={20} step={0.1} precision={2} width={70} suffix="s" onChange={(v) => a.animPatch({ dur: Math.round(v * 1000) })} />
          </label>
          <label className="ribbon-field">
            <span>Delay</span>
            <NumberField value={(cur?.delay ?? 0) / 1000} min={0} max={60} step={0.1} precision={2} width={70} suffix="s" onChange={(v) => a.animPatch({ delay: Math.round(v * 1000) })} />
          </label>
        </RRows>
        <RRows>
          <RButton icon={<ArrowUp size={I} />} label="Move earlier" showLabel disabled={!cur} onClick={() => a.moveAnim(-1)} />
          <RButton icon={<ArrowDown size={I} />} label="Move later" showLabel disabled={!cur} onClick={() => a.moveAnim(1)} />
          <RButton icon={<Trash2 size={I} />} label="Remove" showLabel disabled={!cur} onClick={a.removeAnim} />
        </RRows>
      </RibbonGroup>
      <RibbonGroup label="Advanced">
        <RBigButton icon={<Sparkles />} label="Animation pane" tip="See and reorder all animations on this slide" active={a.animPaneOn} onClick={a.toggleAnimPane} />
      </RibbonGroup>
    </>
  );
}

/* ================================================================== SHOW */

function ShowTab({ doc, a }: { doc: SlidesDoc; a: SlideActions }) {
  return (
    <>
      <RibbonGroup label="Start slide show">
        <RBigButton icon={<Play />} label="From beginning" tip="Start from the first slide" keys="F5" onClick={() => a.startShow(false)} />
        <RBigButton icon={<MonitorPlay />} label="From current" tip="Start from this slide" keys="Shift+F5" onClick={() => a.startShow(true)} />
        <RBigButton icon={<Presentation />} label="Presenter view" tip="Notes, timer and next slide for you — the slides for your audience" onClick={a.presenter} />
      </RibbonGroup>
      <RibbonGroup label="Set up">
        <RBigButton icon={doc.slide?.hidden ? <Eye /> : <EyeOff />} label={doc.slide?.hidden ? 'Unhide slide' : 'Hide slide'} tip="Hidden slides are skipped in the slide show" onClick={a.hideSlide} />
      </RibbonGroup>
    </>
  );
}

/* ================================================================== VIEW */

function ViewTab({ ui, a }: { ui: SlidesUI; a: SlideActions }) {
  return (
    <>
      <RibbonGroup label="Presentation views">
        <RBigButton icon={<Presentation />} label="Normal" active={ui.view === 'normal'} onClick={() => a.view('normal')} />
        <RBigButton icon={<LayoutGrid />} label="Slide sorter" active={ui.view === 'sorter'} onClick={() => a.view('sorter')} />
      </RibbonGroup>
      <RibbonGroup label="Show">
        <RRows>
          <Checkbox checked={ui.notes} onChange={a.toggleNotes} label="Notes" />
          <Checkbox checked={ui.showGrid} onChange={a.toggleGrid} label="Gridlines" />
          <Checkbox checked={ui.snap} onChange={a.toggleSnap} label="Smart guides" />
        </RRows>
      </RibbonGroup>
      <RibbonGroup label="Zoom">
        <RDropdown icon={<ZoomIn size={I} />} label="Zoom" menu={[0.5, 0.75, 1, 1.5, 2].map((z) => ({ label: `${z * 100}%`, onSelect: () => a.zoom(z) }))} />
        <RBigButton icon={<Square />} label="Fit to window" active={ui.fit} onClick={() => a.zoom('fit')} />
      </RibbonGroup>
      <RibbonGroup label="Notes">
        <RBigButton icon={<StickyNote />} label="Notes pane" active={ui.notes} onClick={a.toggleNotes} />
      </RibbonGroup>
    </>
  );
}

/* ========================================================= SHAPE FORMAT */

function SizeGroup({ a }: { a: SlideActions }) {
  const one = a.selected.length === 1 ? a.selected[0] : null;
  return (
    <RibbonGroup label="Size">
      <RRows>
        <label className="ribbon-field">
          <span>Height</span>
          <NumberField value={one ? Math.round((one.h / 96) * 2.54 * 100) / 100 : 0} disabled={!one} min={0.01} max={200} step={0.1} precision={2} width={80} suffix="cm" onChange={(v) => a.setSize(undefined, (v / 2.54) * 96)} />
        </label>
        <label className="ribbon-field">
          <span>Width</span>
          <NumberField value={one ? Math.round((one.w / 96) * 2.54 * 100) / 100 : 0} disabled={!one} min={0.01} max={200} step={0.1} precision={2} width={80} suffix="cm" onChange={(v) => a.setSize((v / 2.54) * 96, undefined)} />
        </label>
      </RRows>
    </RibbonGroup>
  );
}

function ShapeTab({ doc, a }: { doc: SlidesDoc; a: SlideActions }) {
  return (
    <>
      <RibbonGroup label="Insert shapes">
        <RDropdown icon={<Shapes size={I} />} label="Shapes" panel={(close) => <ShapesPanel onPick={a.insertShape} close={close} />} />
        <RDropdown icon={<PenLine size={I} />} label="Change shape" tip="Swap the shape, keeping its text and formatting" panel={(close) => <ShapesPanel onPick={a.changeShape} close={close} />} />
      </RibbonGroup>
      <RibbonGroup label="Shape styles">
        <div className="quick-styles inline thin-scroll">
          <QuickStyles doc={doc} a={a} close={() => undefined} />
        </div>
      </RibbonGroup>
      <DrawingGroup doc={doc} a={a} fillOnly />
      <RibbonGroup label="Effects">
        <RRows>
          <RButton icon={<Square size={I} />} label="Shadow" showLabel active={a.selected.some((e) => (e.type === 'shape' || e.type === 'image') && !!e.shadow)} onClick={() => a.shadow(!a.selected.some((e) => (e.type === 'shape' || e.type === 'image') && !!e.shadow))} />
          <RButton icon={<PaintBucket size={I} />} label="Format shape…" showLabel onClick={a.formatShape} />
        </RRows>
      </RibbonGroup>
      <RibbonGroup label="Arrange">
        <RRows>
          <RButton icon={<BringToFront size={I} />} label="Bring forward" showLabel onClick={() => a.arrange('forward')} />
          <RButton icon={<SendToBack size={I} />} label="Send backward" showLabel onClick={() => a.arrange('backward')} />
        </RRows>
        <RDropdown icon={<LayoutGrid size={I} />} label="Arrange" menu={() => arrangeMenu(a)} />
      </RibbonGroup>
      <SizeGroup a={a} />
    </>
  );
}

/* ======================================================= PICTURE FORMAT */

const MASKS = ['rect', 'roundRect', 'ellipse', 'triangle', 'diamond', 'hexagon', 'octagon', 'star5', 'heart', 'teardrop', 'snip1Rect', 'round2SameRect'];

function PictureTab({ doc, a }: { doc: SlidesDoc; a: SlideActions }) {
  const cp = colorProps(doc.pres.theme);
  const img = a.selected.find((e) => e.type === 'image');
  return (
    <>
      <RibbonGroup label="Adjust">
        <RBigButton icon={<RotateCcw />} label="Reset picture" tip="Remove crop, border, shadow and shape" onClick={a.resetPicture} />
      </RibbonGroup>
      <RibbonGroup label="Picture styles">
        <RDropdown
          icon={<Shapes size={I} />}
          label="Picture shape"
          tip="Crop the picture to a shape"
          panel={(close) => (
            <div className="shape-grid mask-grid">
              {MASKS.map((g) => (
                <button key={g} type="button" className="shape-cell" onClick={() => (a.pictureShape(g), close())}>
                  <ShapeIcon geom={g} />
                </button>
              ))}
            </div>
          )}
        />
        <ColorSplitButton icon={<PenLine size={I} />} color={img && img.type === 'image' && img.line ? img.line.color : '@tx1'} tip="Picture border" noneLabel="No border" onApply={(c) => a.shapeLine(c ? { color: c } : null)} {...cp} extra={(close) => <LineOptions a={a} close={close} />} />
        <RButton icon={<Square size={I} />} label="Shadow" showLabel active={!!(img && img.type === 'image' && img.shadow)} onClick={() => a.shadow(!(img && img.type === 'image' && img.shadow))} />
      </RibbonGroup>
      <RibbonGroup label="Size">
        <RBigButton icon={<Crop />} label="Crop" tip="Crop the picture" onClick={a.crop} />
      </RibbonGroup>
      <RibbonGroup label="Arrange">
        <RDropdown icon={<LayoutGrid size={I} />} label="Arrange" menu={() => arrangeMenu(a)} />
      </RibbonGroup>
      <SizeGroup a={a} />
    </>
  );
}

/* ========================================================== TABLE DESIGN */

const TABLE_FAMILIES: { family: NonNullable<TableStyleOpts['family']>; label: string }[] = [
  { family: 'medium2', label: 'Medium' },
  { family: 'light1', label: 'Light' },
  { family: 'light2', label: 'Light grid' },
  { family: 'medium1', label: 'Medium grid' },
  { family: 'dark1', label: 'Dark' },
  { family: 'none', label: 'No style' },
];

function TableTab({ doc, a }: { doc: SlidesDoc; a: SlideActions }) {
  const t = a.selected.find((e) => e.type === 'table');
  const st = t && t.type === 'table' ? (t.style ?? {}) : {};
  const theme = doc.pres.theme;
  const cp = colorProps(theme);
  return (
    <>
      <RibbonGroup label="Table style options">
        <RRows>
          <Checkbox checked={!!st.firstRow} onChange={(v) => a.tableStyle({ firstRow: v })} label="Header row" />
          <Checkbox checked={!!st.lastRow} onChange={(v) => a.tableStyle({ lastRow: v })} label="Total row" />
          <Checkbox checked={!!st.bandRows} onChange={(v) => a.tableStyle({ bandRows: v })} label="Banded rows" />
        </RRows>
        <RRows>
          <Checkbox checked={!!st.firstCol} onChange={(v) => a.tableStyle({ firstCol: v })} label="First column" />
          <Checkbox checked={!!st.lastCol} onChange={(v) => a.tableStyle({ lastCol: v })} label="Last column" />
          <Checkbox checked={!!st.bandCols} onChange={(v) => a.tableStyle({ bandCols: v })} label="Banded columns" />
        </RRows>
      </RibbonGroup>
      <RibbonGroup label="Table styles" className="fx-group">
        <div className="table-style-gallery thin-scroll">
          {TABLE_FAMILIES.flatMap((f) =>
            (f.family === 'none' ? (['accent1'] as ThemeColorName[]) : (['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'] as ThemeColorName[])).map((acc) => (
              <button
                key={`${f.family}-${acc}`}
                type="button"
                className={`table-style-item${(st.family ?? 'medium2') === f.family && (f.family === 'none' || (st.accent ?? 'accent1') === acc) ? ' active' : ''}`}
                data-tip={`${f.label}${f.family === 'none' ? '' : ` – ${acc.replace('accent', 'Accent ')}`}`}
                onClick={() => a.tableStyle({ family: f.family, accent: acc })}
                style={{ '--ts-accent': resolveColor(`@${acc}`, theme), '--ts-light': resolveColor(`@${acc}+80`, theme), '--ts-mid': resolveColor(`@${acc}+60`, theme) } as React.CSSProperties}
              >
                <span className={`ts-sample ts-${f.family}`} />
              </button>
            )),
          )}
        </div>
      </RibbonGroup>
      <RibbonGroup label="Cells">
        <ColorSplitButton icon={<PaintBucket size={I} />} color="@accent1+60" tip="Cell shading" noneLabel="No fill" onApply={(c) => a.cellFill(c ? { type: 'solid', color: c } : null)} {...cp} />
        <RRows>
          <RButton icon={<TableCellsMerge size={I} />} label="Merge cells" showLabel onClick={a.mergeCells} />
          <RButton icon={<TableCellsSplit size={I} />} label="Split cell" showLabel onClick={a.splitCell} />
        </RRows>
      </RibbonGroup>
      <RibbonGroup label="Rows & columns">
        <RRows>
          <RButton icon={<BetweenHorizontalStart size={I} />} label="Insert above" showLabel onClick={() => a.tableInsert('above')} />
          <RButton icon={<BetweenHorizontalStart size={I} style={{ transform: 'scaleY(-1)' }} />} label="Insert below" showLabel onClick={() => a.tableInsert('below')} />
        </RRows>
        <RRows>
          <RButton icon={<BetweenVerticalStart size={I} />} label="Insert left" showLabel onClick={() => a.tableInsert('left')} />
          <RButton icon={<BetweenVerticalStart size={I} style={{ transform: 'scaleX(-1)' }} />} label="Insert right" showLabel onClick={() => a.tableInsert('right')} />
        </RRows>
        <RDropdown
          icon={<Trash2 size={I} />}
          label="Delete"
          menu={[
            { label: 'Delete row', onSelect: () => a.tableDelete('row') },
            { label: 'Delete column', onSelect: () => a.tableDelete('col') },
            { label: 'Delete table', onSelect: () => a.tableDelete('table') },
          ]}
        />
      </RibbonGroup>
      <RibbonGroup label="Alignment">
        <RRows>
          <RButton icon={<AlignVerticalJustifyStart size={I} />} tip="Align top" onClick={() => a.anchor('t')} />
          <RButton icon={<AlignVerticalJustifyCenter size={I} />} tip="Centre vertically" onClick={() => a.anchor('m')} />
          <RButton icon={<AlignVerticalJustifyEnd size={I} />} tip="Align bottom" onClick={() => a.anchor('b')} />
        </RRows>
      </RibbonGroup>
    </>
  );
}

/* ================================================================= CHART */

function ChartTab({ a }: { a: SlideActions }) {
  const c = a.selected.find((e) => e.type === 'chart');
  const chart = c && c.type === 'chart' ? c.chart : undefined;
  return (
    <>
      <RibbonGroup label="Data">
        <RBigButton icon={<Table />} label="Edit data" tip="Edit the chart's numbers and labels" onClick={a.editChart} />
      </RibbonGroup>
      <RibbonGroup label="Type">
        <RDropdown icon={<ChartColumn size={I} />} label="Change chart type" menu={CHARTS.map((x) => ({ label: x.label, icon: x.icon, checked: chart?.type === x.type, onSelect: () => a.chartType(x.type) }))} />
      </RibbonGroup>
      <RibbonGroup label="Chart elements">
        <RRows>
          <RDropdown
            label="Legend"
            menu={(['bottom', 'right', 'top', 'none'] as const).map((l) => ({ label: l === 'none' ? 'None' : l[0].toUpperCase() + l.slice(1), checked: (chart?.legend ?? 'bottom') === l, onSelect: () => a.chartPatch({ legend: l }) }))}
          />
          <Checkbox checked={!!chart?.dataLabels} onChange={(v) => a.chartPatch({ dataLabels: v })} label="Data labels" />
        </RRows>
      </RibbonGroup>
    </>
  );
}
