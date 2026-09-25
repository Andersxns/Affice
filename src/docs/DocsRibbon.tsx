import { useEditorState, type Editor } from '@tiptap/react';
import { NodeSelection } from '@tiptap/pm/state';
import {
  AArrowDown,
  AArrowUp,
  AlignCenter,
  AlignJustify,
  AlignLeft,
  AlignRight,
  ArrowDownToLine,
  ArrowUpToLine,
  Baseline,
  BetweenHorizontalStart,
  BetweenVerticalStart,
  Bold,
  BookOpen,
  CaseSensitive,
  CalendarDays,
  ChevronsUpDown,
  ClipboardPaste,
  Code2,
  Columns3,
  Copy,
  Eraser,
  Eye,
  FileText,
  Focus,
  FoldVertical,
  Heading1,
  Highlighter,
  Image as ImageIcon,
  Indent,
  Italic,
  Languages,
  LayoutTemplate,
  Link2,
  List,
  ListOrdered,
  ListTodo,
  ListTree,
  Maximize2,
  MessageSquarePlus,
  MessagesSquare,
  Minus,
  Moon,
  Omega,
  Outdent,
  Paintbrush,
  PanelLeft,
  Pilcrow,
  Quote,
  Redo2,
  Replace,
  RectangleHorizontal,
  RectangleVertical,
  Ruler,
  Scissors,
  Search,
  SeparatorHorizontal,
  Sigma,
  SpellCheck,
  SquareSplitHorizontal,
  Strikethrough,
  Subscript as SubIcon,
  Superscript as SupIcon,
  Table as TableIcon,
  TableCellsMerge,
  TableCellsSplit,
  TextCursorInput,
  Trash2,
  Type,
  Underline as UnderlineIcon,
  Undo2,
  Volume2,
  WholeWord,
  Droplet,
  Stamp,
  PanelTop,
  PanelBottom,
  Hash,
  Monitor,
  Square,
  Circle,
} from 'lucide-react';
import { useState, type ReactNode } from 'react';
import { api } from '@/lib/platform';
import { CORE_FONTS } from '@/lib/fonts';
import { ColorPicker, ColorSplitButton } from '@/ui/color';
import { NumberField } from '@/ui/controls';
import { promptDialog } from '@/ui/dialog';
import { FontFamilyCombo, FontSizeCombo, nextFontSize } from '@/ui/fontControls';
import type { MenuItem } from '@/ui/menu';
import { RBigButton, RButton, RDropdown, RibbonGroup, RRow, RRows, RSep, RSplit, Ribbon, type RibbonTab } from '@/ui/ribbon';
import { toast } from '@/ui/toast';
import { selectWordAtCursor } from './editor/extensions';
import { MARGIN_PRESETS, PAPER_SIZES, resolveStyles, STYLE_SETS, type DocSettings, type StyleId } from './model';
import type { DocView } from './DocPages';

/* ---------------------------------------------------------- format state */

export interface FormatState {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  sub: boolean;
  sup: boolean;
  code: boolean;
  fontFamily: string | null;
  fontSize: number | null;
  color: string | null;
  highlight: string | null;
  align: string;
  block: string;
  styleName: string | null;
  bullet: boolean;
  ordered: boolean;
  task: boolean;
  quote: boolean;
  codeBlock: boolean;
  link: boolean;
  inTable: boolean;
  image: null | { align: string; width: number | null; border: boolean; shadow: boolean; rounded: boolean; alt: string };
  canUndo: boolean;
  canRedo: boolean;
  lineHeight: number | null;
  indent: number;
  spaceBefore: number | null;
  spaceAfter: number | null;
  empty: boolean;
}

function sizeToPt(v: string | undefined | null): number | null {
  if (!v) return null;
  const n = parseFloat(v);
  if (!Number.isFinite(n)) return null;
  if (/px$/.test(v)) return Math.round(((n * 72) / 96) * 10) / 10;
  return n;
}

export function selectFormat(editor: Editor, settings: DocSettings): FormatState {
  const ts = editor.getAttributes('textStyle');
  const { selection } = editor.state;
  const $from = selection.$from;
  let para: Record<string, unknown> = {};
  let block = 'paragraph';
  for (let d = $from.depth; d >= 0; d--) {
    const n = $from.node(d);
    if (n.type.name === 'paragraph' || n.type.name === 'heading') {
      para = n.attrs;
      block = n.type.name === 'heading' ? `h${n.attrs.level}` : 'paragraph';
      break;
    }
  }
  const styles = resolveStyles(settings);
  const styleName = (para.styleName as string | null) ?? null;
  const styleKey = (block !== 'paragraph' ? block : styleName && ['title', 'subtitle', 'quote'].includes(styleName) ? styleName : 'normal') as StyleId;
  const st = styles[styleKey] ?? styles.normal;
  const img = selection instanceof NodeSelection && selection.node.type.name === 'image' ? selection.node.attrs : null;
  return {
    bold: editor.isActive('bold'),
    italic: editor.isActive('italic'),
    underline: editor.isActive('underline'),
    strike: editor.isActive('strike'),
    sub: editor.isActive('subscript'),
    sup: editor.isActive('superscript'),
    code: editor.isActive('code'),
    fontFamily: (ts.fontFamily as string | undefined)?.replace(/^["']|["']$/g, '').split(',')[0].replace(/["']/g, '').trim() || st.font || 'Calibri',
    fontSize: sizeToPt(ts.fontSize as string | undefined) ?? st.size ?? 11,
    color: (ts.color as string | undefined) ?? null,
    highlight: (editor.getAttributes('highlight').color as string | undefined) ?? null,
    align: (para.textAlign as string | undefined) ?? st.align ?? 'left',
    block,
    styleName,
    bullet: editor.isActive('bulletList'),
    ordered: editor.isActive('orderedList'),
    task: editor.isActive('taskList'),
    quote: editor.isActive('blockquote'),
    codeBlock: editor.isActive('codeBlock'),
    link: editor.isActive('link'),
    inTable: editor.isActive('table'),
    image: img
      ? { align: img.align, width: img.width, border: Boolean(img.border), shadow: Boolean(img.shadow), rounded: Boolean(img.rounded), alt: img.alt ?? '' }
      : null,
    canUndo: editor.can().undo(),
    canRedo: editor.can().redo(),
    lineHeight: (para.lineHeight as number | null) ?? null,
    indent: (para.indent as number) ?? 0,
    spaceBefore: (para.spaceBefore as number | null) ?? null,
    spaceAfter: (para.spaceAfter as number | null) ?? null,
    empty: selection.empty,
  };
}

/* --------------------------------------------------------------- props */

export interface DocsRibbonProps {
  editor: Editor;
  settings: DocSettings;
  updateSettings: (patch: Partial<DocSettings>) => void;
  onFile: () => void;
  openFind: (mode: 'find' | 'replace') => void;
  insertImage: () => void;
  insertLink: () => void;
  addComment: () => void;
  pageSetup: () => void;
  headerFooter: (which: 'header' | 'footer') => void;
  wordCount: () => void;
  insertSymbol: () => void;
  insertEquation: () => void;
  paragraph: () => void;
  view: DocView;
  setView: (v: DocView) => void;
  zoom: number;
  setZoom: (z: number) => void;
  fitWidth: () => void;
  onePage: () => void;
  panel: string;
  togglePanel: (p: 'nav' | 'comments') => void;
  showMarks: boolean;
  setShowMarks: (v: boolean) => void;
  focusMode: boolean;
  setFocusMode: (v: boolean) => void;
  darkPage: boolean;
  setDarkPage: (v: boolean) => void;
  readAloud: () => void;
  reading: boolean;
  painter: boolean;
  startPainter: () => void;
  spellcheck: boolean;
  toggleSpell: () => void;
}

const I = 17;

/** Applies a character command; with a collapsed cursor inside a word, formats the whole word (like Word). */
function charCmd(editor: Editor, fn: (c: ReturnType<Editor['chain']>) => ReturnType<Editor['chain']>) {
  const { selection } = editor.state;
  if (selection.empty) {
    const cursor = selection.from;
    const text = selection.$from.parent.textContent;
    const off = selection.$from.parentOffset;
    const inWord = off > 0 && off < text.length && /[\p{L}\p{N}]/u.test(text[off - 1]) && /[\p{L}\p{N}]/u.test(text[off]);
    if (inWord && selectWordAtCursor(editor)) {
      fn(editor.chain().focus()).setTextSelection(cursor).run();
      return;
    }
  }
  fn(editor.chain().focus()).run();
}

function changeCase(editor: Editor, mode: 'upper' | 'lower' | 'sentence' | 'title' | 'toggle') {
  const { from, to, empty } = editor.state.selection;
  if (empty) {
    toast.info('Select some text first');
    return;
  }
  const tr = editor.state.tr;
  editor.state.doc.nodesBetween(from, to, (node, pos) => {
    if (!node.isText) return true;
    const s = Math.max(from, pos);
    const e = Math.min(to, pos + node.nodeSize);
    const t = node.text!.slice(s - pos, e - pos);
    let out = t;
    if (mode === 'upper') out = t.toUpperCase();
    else if (mode === 'lower') out = t.toLowerCase();
    else if (mode === 'title') out = t.toLowerCase().replace(/(^|[\s\-–—(“"'])(\p{L})/gu, (_m, a, b) => a + b.toUpperCase());
    else if (mode === 'sentence') out = t.toLowerCase().replace(/(^\s*|[.!?]\s+)(\p{L})/gu, (_m, a, b) => a + b.toUpperCase());
    else out = Array.from(t, (c) => (c === c.toUpperCase() ? c.toLowerCase() : c.toUpperCase())).join('');
    if (out !== t) tr.insertText(out, tr.mapping.map(s), tr.mapping.map(e));
    return false;
  });
  editor.view.dispatch(tr);
  editor.commands.setTextSelection({ from, to });
}

export function DocsRibbon(p: DocsRibbonProps) {
  const { editor, settings } = p;
  const f = useEditorState({ editor, selector: ({ editor: e }) => selectFormat(e as Editor, settings) }) as FormatState;
  const [active, setActive] = useState('home');

  const effectiveTab = f.image && active === 'picture' ? 'picture' : f.inTable && active === 'table' ? 'table' : active;

  const tabs: RibbonTab[] = [
    { id: 'home', label: 'Home', content: <HomeTab {...p} f={f} /> },
    { id: 'insert', label: 'Insert', content: <InsertTab {...p} f={f} /> },
    { id: 'layout', label: 'Layout', content: <LayoutTab {...p} f={f} /> },
    { id: 'design', label: 'Design', content: <DesignTab {...p} /> },
    { id: 'review', label: 'Review', content: <ReviewTab {...p} /> },
    { id: 'view', label: 'View', content: <ViewTab {...p} /> },
    { id: 'table', label: 'Table', content: <TableTab {...p} f={f} />, contextual: true, hidden: !f.inTable },
    { id: 'picture', label: 'Picture', content: <PictureTab {...p} f={f} />, contextual: true, hidden: !f.image },
  ];

  return (
    <Ribbon
      app="doc"
      tabs={tabs}
      active={tabs.find((t) => t.id === effectiveTab && !t.hidden) ? effectiveTab : 'home'}
      onActive={setActive}
      onFile={p.onFile}
      right={
        <div className="ribbon-right">
          <button type="button" className="icon-btn icon-btn-sm" data-tip="Undo" data-tip-key="Ctrl+Z" disabled={!f.canUndo} onMouseDown={(e) => e.preventDefault()} onClick={() => editor.chain().focus().undo().run()}>
            <Undo2 size={16} />
          </button>
          <button type="button" className="icon-btn icon-btn-sm" data-tip="Redo" data-tip-key="Ctrl+Y" disabled={!f.canRedo} onMouseDown={(e) => e.preventDefault()} onClick={() => editor.chain().focus().redo().run()}>
            <Redo2 size={16} />
          </button>
        </div>
      }
    />
  );
}

type TabProps = DocsRibbonProps & { f: FormatState };

/* =================================================================== HOME */

function HomeTab(p: TabProps) {
  const { editor, f } = p;
  const bulletStyles: Array<[string | null, string]> = [
    [null, '● Disc'],
    ['circle', '○ Circle'],
    ['square', '■ Square'],
    ['dash', '– Dash'],
    ['check', '✓ Check'],
    ['arrow', '➤ Arrow'],
  ];
  const numberStyles: Array<[string | null, string]> = [
    [null, '1. 2. 3.'],
    ['lower-alpha', 'a. b. c.'],
    ['upper-alpha', 'A. B. C.'],
    ['lower-roman', 'i. ii. iii.'],
    ['upper-roman', 'I. II. III.'],
  ];
  const setList = (type: 'bulletList' | 'orderedList', style: string | null) => {
    const chain = editor.chain().focus();
    if (!editor.isActive(type)) (type === 'bulletList' ? chain.toggleBulletList() : chain.toggleOrderedList()).run();
    editor.chain().focus().updateAttributes(type, { listStyle: style }).run();
  };
  const lineSpacing = (v: number | null) => editor.chain().focus().setParagraphFormat({ lineHeight: v }).run();

  return (
    <>
      <RibbonGroup label="Clipboard">
        <RSplit
          big
          icon={<ClipboardPaste />}
          label="Paste"
          tip="Paste (Ctrl+V)"
          onClick={() => {
            editor.commands.focus();
            api.app.editCommand('paste');
          }}
          menu={[
            { label: 'Paste', shortcut: 'Ctrl+V', onSelect: () => (editor.commands.focus(), api.app.editCommand('paste')) },
            { label: 'Paste as plain text', shortcut: 'Ctrl+Shift+V', onSelect: () => (editor.commands.focus(), api.app.editCommand('pasteAndMatchStyle')) },
          ]}
        />
        <RRows>
          <RButton icon={<Scissors size={I} />} label="Cut" keys="Ctrl+X" onClick={() => (editor.commands.focus(), api.app.editCommand('cut'))} />
          <RButton icon={<Copy size={I} />} label="Copy" keys="Ctrl+C" onClick={() => (editor.commands.focus(), api.app.editCommand('copy'))} />
          <RButton icon={<Paintbrush size={I} />} label="Format painter" tip="Format painter — copy formatting, then select text to apply it" active={p.painter} onClick={p.startPainter} />
        </RRows>
      </RibbonGroup>

      <RibbonGroup label="Font">
        <RRows>
          <RRow>
            <FontFamilyCombo value={f.fontFamily ?? 'Calibri'} onChange={(font) => charCmd(editor, (c) => c.setFontFamily(font))} />
            <FontSizeCombo value={f.fontSize} onChange={(pt) => charCmd(editor, (c) => c.setFontSize(`${pt}pt`))} />
            <RButton icon={<AArrowUp size={I} />} tip="Increase font size" keys="Ctrl+Shift+>" onClick={() => charCmd(editor, (c) => c.setFontSize(`${nextFontSize(f.fontSize ?? 11, 1)}pt`))} />
            <RButton icon={<AArrowDown size={I} />} tip="Decrease font size" keys="Ctrl+Shift+<" onClick={() => charCmd(editor, (c) => c.setFontSize(`${nextFontSize(f.fontSize ?? 11, -1)}pt`))} />
            <RDropdown
              icon={<CaseSensitive size={I} />}
              tip="Change case"
              showLabel={false}
              menu={[
                { label: 'Sentence case.', onSelect: () => changeCase(editor, 'sentence') },
                { label: 'lowercase', onSelect: () => changeCase(editor, 'lower') },
                { label: 'UPPERCASE', onSelect: () => changeCase(editor, 'upper') },
                { label: 'Capitalize Each Word', onSelect: () => changeCase(editor, 'title') },
                { label: 'tOGGLE cASE', onSelect: () => changeCase(editor, 'toggle') },
              ]}
            />
            <RButton
              icon={<Eraser size={I} />}
              tip="Clear all formatting"
              keys="Ctrl+Space"
              onClick={() => editor.chain().focus().unsetAllMarks().clearParagraphFormat().setParagraph().setStyleName(null).run()}
            />
          </RRow>
          <RRow>
            <RButton icon={<Bold size={I} />} tip="Bold" keys="Ctrl+B" active={f.bold} onClick={() => charCmd(editor, (c) => c.toggleBold())} />
            <RButton icon={<Italic size={I} />} tip="Italic" keys="Ctrl+I" active={f.italic} onClick={() => charCmd(editor, (c) => c.toggleItalic())} />
            <RButton icon={<UnderlineIcon size={I} />} tip="Underline" keys="Ctrl+U" active={f.underline} onClick={() => charCmd(editor, (c) => c.toggleUnderline())} />
            <RButton icon={<Strikethrough size={I} />} tip="Strikethrough" active={f.strike} onClick={() => charCmd(editor, (c) => c.toggleStrike())} />
            <RButton icon={<SubIcon size={I} />} tip="Subscript" keys="Ctrl+=" active={f.sub} onClick={() => charCmd(editor, (c) => c.toggleSubscript())} />
            <RButton icon={<SupIcon size={I} />} tip="Superscript" keys="Ctrl+Shift+=" active={f.sup} onClick={() => charCmd(editor, (c) => c.toggleSuperscript())} />
            <RSep />
            <ColorSplitButton
              icon={<Highlighter size={I} />}
              color={f.highlight ?? '#ffff00'}
              palette="highlight"
              tip="Highlight"
              noneLabel="No highlight"
              onApply={(c) => charCmd(editor, (ch) => (c ? ch.setHighlight({ color: c }) : ch.unsetHighlight()))}
            />
            <ColorSplitButton
              icon={<Baseline size={I} />}
              color={f.color ?? '#e5484d'}
              tip="Font colour"
              noneLabel="Automatic"
              onApply={(c) => charCmd(editor, (ch) => (c ? ch.setColor(c) : ch.unsetColor()))}
            />
            <RButton icon={<Code2 size={I} />} tip="Inline code" active={f.code} onClick={() => charCmd(editor, (c) => c.toggleCode())} />
          </RRow>
        </RRows>
      </RibbonGroup>

      <RibbonGroup label="Paragraph" collapse={2} icon={<AlignLeft />}>
        <RRows>
          <RRow>
            <RSplit
              icon={<List size={I} />}
              tip="Bullets"
              keys="Ctrl+Shift+8"
              active={f.bullet}
              onClick={() => editor.chain().focus().toggleBulletList().run()}
              menu={bulletStyles.map(([s, label]) => ({ label, onSelect: () => setList('bulletList', s) }))}
            />
            <RSplit
              icon={<ListOrdered size={I} />}
              tip="Numbering"
              keys="Ctrl+Shift+7"
              active={f.ordered}
              onClick={() => editor.chain().focus().toggleOrderedList().run()}
              menu={numberStyles.map(([s, label]) => ({ label, onSelect: () => setList('orderedList', s) }))}
            />
            <RButton icon={<ListTodo size={I} />} tip="Checklist" active={f.task} onClick={() => editor.chain().focus().toggleTaskList().run()} />
            <RButton icon={<Outdent size={I} />} tip="Decrease indent" keys="Ctrl+[" onClick={() => editor.chain().focus().outdent().run()} />
            <RButton icon={<Indent size={I} />} tip="Increase indent" keys="Ctrl+]" onClick={() => editor.chain().focus().indent().run()} />
            <RButton icon={<Pilcrow size={I} />} tip="Show paragraph marks" active={p.showMarks} onClick={() => p.setShowMarks(!p.showMarks)} />
          </RRow>
          <RRow>
            <RButton icon={<AlignLeft size={I} />} tip="Align left" keys="Ctrl+L" active={f.align === 'left'} onClick={() => editor.chain().focus().setTextAlign('left').run()} />
            <RButton icon={<AlignCenter size={I} />} tip="Centre" keys="Ctrl+E" active={f.align === 'center'} onClick={() => editor.chain().focus().setTextAlign('center').run()} />
            <RButton icon={<AlignRight size={I} />} tip="Align right" keys="Ctrl+R" active={f.align === 'right'} onClick={() => editor.chain().focus().setTextAlign('right').run()} />
            <RButton icon={<AlignJustify size={I} />} tip="Justify" keys="Ctrl+J" active={f.align === 'justify'} onClick={() => editor.chain().focus().setTextAlign('justify').run()} />
            <RDropdown
              icon={<ChevronsUpDown size={I} />}
              tip="Line and paragraph spacing"
              showLabel={false}
              menu={[
                ...[1, 1.15, 1.5, 2, 2.5, 3].map((v) => ({ label: v.toFixed(v % 1 ? 2 : 1).replace(/0$/, ''), checked: f.lineHeight === v, onSelect: () => lineSpacing(v) })),
                { label: 'Style default', checked: f.lineHeight === null, onSelect: () => lineSpacing(null) },
                { separator: true },
                {
                  label: f.spaceBefore ? 'Remove space before paragraph' : 'Add space before paragraph',
                  onSelect: () => editor.chain().focus().setParagraphFormat({ spaceBefore: f.spaceBefore ? null : 12 }).run(),
                },
                {
                  label: f.spaceAfter === 0 ? 'Add space after paragraph' : 'Remove space after paragraph',
                  onSelect: () => editor.chain().focus().setParagraphFormat({ spaceAfter: f.spaceAfter === 0 ? null : 0 }).run(),
                },
                { separator: true },
                { label: 'Paragraph options…', onSelect: p.paragraph },
              ]}
            />
            <RDropdown
              icon={<Droplet size={I} />}
              tip="Paragraph shading"
              showLabel={false}
              panel={(close) => (
                <ColorPicker
                  noneLabel="No shading"
                  onPick={(c) => {
                    editor.chain().focus().setParagraphFormat({ shading: c }).run();
                    close();
                  }}
                />
              )}
            />
            <RButton icon={<Quote size={I} />} tip="Block quote" active={f.quote} onClick={() => editor.chain().focus().toggleBlockquote().run()} />
          </RRow>
        </RRows>
      </RibbonGroup>

      <RibbonGroup label="Styles" className="rgroup-styles">
        <StyleGallery editor={editor} settings={p.settings} f={f} />
      </RibbonGroup>

      <RibbonGroup label="Editing" collapse={1} icon={<Search />}>
        <RRows>
          <RButton icon={<Search size={I} />} label="Find" keys="Ctrl+F" showLabel onClick={() => p.openFind('find')} />
          <RButton icon={<Replace size={I} />} label="Replace" keys="Ctrl+H" showLabel onClick={() => p.openFind('replace')} />
          <RButton icon={<TextCursorInput size={I} />} label="Select all" keys="Ctrl+A" showLabel onClick={() => editor.chain().focus().selectAll().run()} />
        </RRows>
      </RibbonGroup>
    </>
  );
}

const GALLERY: Array<{ id: string; label: string; apply: (e: Editor) => void; isActive: (f: FormatState) => boolean; key: StyleId }> = [
  { id: 'normal', label: 'Normal', key: 'normal', apply: (e) => e.chain().focus().setParagraph().setStyleName(null).run(), isActive: (f) => f.block === 'paragraph' && !f.styleName },
  { id: 'nospacing', label: 'No Spacing', key: 'normal', apply: (e) => e.chain().focus().setParagraph().setStyleName('nospacing').run(), isActive: (f) => f.styleName === 'nospacing' },
  { id: 'title', label: 'Title', key: 'title', apply: (e) => e.chain().focus().setParagraph().setStyleName('title').run(), isActive: (f) => f.styleName === 'title' && f.block === 'paragraph' },
  { id: 'subtitle', label: 'Subtitle', key: 'subtitle', apply: (e) => e.chain().focus().setParagraph().setStyleName('subtitle').run(), isActive: (f) => f.styleName === 'subtitle' },
  { id: 'h1', label: 'Heading 1', key: 'h1', apply: (e) => e.chain().focus().setHeading({ level: 1 }).setStyleName(null).run(), isActive: (f) => f.block === 'h1' },
  { id: 'h2', label: 'Heading 2', key: 'h2', apply: (e) => e.chain().focus().setHeading({ level: 2 }).setStyleName(null).run(), isActive: (f) => f.block === 'h2' },
  { id: 'h3', label: 'Heading 3', key: 'h3', apply: (e) => e.chain().focus().setHeading({ level: 3 }).setStyleName(null).run(), isActive: (f) => f.block === 'h3' },
  { id: 'quote', label: 'Quote', key: 'quote', apply: (e) => e.chain().focus().setParagraph().setStyleName('quote').run(), isActive: (f) => f.styleName === 'quote' },
];

function StyleGallery({ editor, settings, f }: { editor: Editor; settings: DocSettings; f: FormatState }) {
  const styles = resolveStyles(settings);
  return (
    <div className="style-gallery thin-scroll">
      {GALLERY.map((g) => {
        const s = styles[g.key];
        return (
          <button
            key={g.id}
            type="button"
            className={`style-tile${g.isActive(f) ? ' active' : ''}`}
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => g.apply(editor)}
            data-tip={`Apply the ${g.label} style`}
          >
            <span
              className="style-sample"
              style={{
                fontFamily: `"${s.font ?? 'Calibri'}"`,
                color: s.color,
                fontWeight: s.bold ? 700 : 400,
                fontStyle: s.italic ? 'italic' : 'normal',
                fontSize: Math.min(20, Math.max(12, (s.size ?? 11) * 0.9)),
                textTransform: s.caps ? 'uppercase' : 'none',
              }}
            >
              AaBbCc
            </span>
            <span className="style-name">{g.label}</span>
          </button>
        );
      })}
    </div>
  );
}

/* ================================================================= INSERT */

function TableGrid({ onPick }: { onPick: (rows: number, cols: number) => void }) {
  const [hover, setHover] = useState<[number, number]>([0, 0]);
  return (
    <div className="table-grid-picker">
      <div className="tgp-label">{hover[0] ? `${hover[1]} × ${hover[0]} table` : 'Insert table'}</div>
      <div className="tgp-grid" onMouseLeave={() => setHover([0, 0])}>
        {Array.from({ length: 8 }, (_, r) =>
          Array.from({ length: 10 }, (_, c) => (
            <span
              key={`${r}-${c}`}
              className={`tgp-cell${r < hover[0] && c < hover[1] ? ' on' : ''}`}
              onMouseEnter={() => setHover([r + 1, c + 1])}
              onClick={() => onPick(r + 1, c + 1)}
            />
          )),
        )}
      </div>
    </div>
  );
}

function InsertTab(p: TabProps) {
  const { editor } = p;
  const insertTable = (rows: number, cols: number, style: string | null = null) => {
    editor.chain().focus().insertTable({ rows, cols, withHeaderRow: true }).run();
    if (style) editor.chain().setTableStyle(style).run();
  };
  const dateItems: MenuItem[] = [
    new Date().toLocaleDateString(undefined, { year: 'numeric', month: 'long', day: 'numeric' }),
    new Date().toLocaleDateString(),
    new Date().toISOString().slice(0, 10),
    new Date().toLocaleDateString(undefined, { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' }),
    new Date().toLocaleTimeString(undefined, { hour: '2-digit', minute: '2-digit' }),
    new Date().toLocaleString(),
  ].map((s) => ({ label: s, onSelect: () => editor.chain().focus().insertContent(s).run() }));
  return (
    <>
      <RibbonGroup label="Pages">
        <RBigButton icon={<SeparatorHorizontal />} label="Page break" keys="Ctrl+Enter" tip="Start the next page here" onClick={() => editor.chain().focus().setPageBreak().run()} />
        <RBigButton icon={<ListTree />} label="Contents" tip="Insert a table of contents built from your headings" onClick={() => editor.chain().focus().insertToc().run()} />
      </RibbonGroup>
      <RibbonGroup label="Tables">
        <RDropdown
          icon={<TableIcon size={24} />}
          label="Table"
          className="rbig-like"
          panel={(close) => (
            <div className="table-insert-panel">
              <TableGrid
                onPick={(r, c) => {
                  insertTable(r, c);
                  close();
                }}
              />
              <div className="menu-sep" />
              <div className="quick-tables">
                {[
                  ['accent', 'Blue header'],
                  ['banded', 'Banded rows'],
                  ['minimal', 'Minimal'],
                  ['dark', 'Dark header'],
                  ['green', 'Green'],
                  ['orange', 'Orange'],
                ].map(([id, label]) => (
                  <button
                    key={id}
                    type="button"
                    className="btn btn-sm"
                    onClick={() => {
                      insertTable(4, 3, id);
                      close();
                    }}
                  >
                    {label}
                  </button>
                ))}
              </div>
            </div>
          )}
        />
      </RibbonGroup>
      <RibbonGroup label="Illustrations">
        <RBigButton icon={<ImageIcon />} label="Picture" tip="Insert a picture from your computer (you can also paste or drag images in)" onClick={p.insertImage} />
        <RBigButton icon={<Sigma />} label="Equation" tip="Insert a maths equation" onClick={p.insertEquation} />
      </RibbonGroup>
      <RibbonGroup label="Links & notes">
        <RBigButton icon={<Link2 />} label="Link" keys="Ctrl+K" onClick={p.insertLink} />
        <RBigButton icon={<MessageSquarePlus />} label="Comment" tip="Add a comment to the selected text" onClick={p.addComment} />
      </RibbonGroup>
      <RibbonGroup label="Header & footer">
        <RRows>
          <RButton icon={<PanelTop size={I} />} label="Header" showLabel onClick={() => p.headerFooter('header')} />
          <RButton icon={<PanelBottom size={I} />} label="Footer" showLabel onClick={() => p.headerFooter('footer')} />
          <RDropdown
            icon={<Hash size={I} />}
            label="Page number"
            menu={[
              { label: 'Bottom centre', onSelect: () => p.updateSettings({ footer: { ...p.settings.footer, center: '{PAGE}' } }) },
              { label: 'Bottom right', onSelect: () => p.updateSettings({ footer: { ...p.settings.footer, right: 'Page {PAGE} of {PAGES}' } }) },
              { label: 'Top right', onSelect: () => p.updateSettings({ header: { ...p.settings.header, right: '{PAGE}' } }) },
              { separator: true },
              {
                label: 'Remove page numbers',
                onSelect: () => {
                  const strip = (s: string) => s.replace(/Page \{PAGE\} of \{PAGES\}|\{PAGE\}|\{PAGES\}/g, '').trim();
                  const clean = (h: DocSettings['header']) => ({ left: strip(h.left), center: strip(h.center), right: strip(h.right) });
                  p.updateSettings({ header: clean(p.settings.header), footer: clean(p.settings.footer) });
                },
              },
            ]}
          />
        </RRows>
      </RibbonGroup>
      <RibbonGroup label="Text">
        <RRows>
          <RDropdown icon={<CalendarDays size={I} />} label="Date & time" menu={dateItems} />
          <RButton icon={<Minus size={I} />} label="Horizontal line" showLabel onClick={() => editor.chain().focus().setHorizontalRule().run()} />
          <RButton icon={<Omega size={I} />} label="Symbol" showLabel onClick={p.insertSymbol} />
        </RRows>
        <RRows>
          <RButton icon={<Code2 size={I} />} label="Code block" showLabel onClick={() => editor.chain().focus().toggleCodeBlock().run()} />
          <RButton icon={<ListTodo size={I} />} label="Checklist" showLabel onClick={() => editor.chain().focus().toggleTaskList().run()} />
          <RButton icon={<Quote size={I} />} label="Quote" showLabel onClick={() => editor.chain().focus().toggleBlockquote().run()} />
        </RRows>
      </RibbonGroup>
    </>
  );
}

/* ================================================================= LAYOUT */

function LayoutTab(p: TabProps) {
  const { editor, settings, f } = p;
  const page = settings.page;
  return (
    <>
      <RibbonGroup label="Page setup">
        <RBigButton
          icon={<Maximize2 />}
          label="Margins"
          menu={[
            ...MARGIN_PRESETS.map((m) => ({
              label: m.label,
              hint: `Top ${m.m.top}" · Bottom ${m.m.bottom}" · Left ${m.m.left}" · Right ${m.m.right}"`,
              checked: JSON.stringify(m.m) === JSON.stringify(page.margins),
              onSelect: () => p.updateSettings({ page: { ...page, margins: { ...m.m } } }),
            })),
            { separator: true },
            { label: 'Custom margins…', onSelect: p.pageSetup },
          ]}
        />
        <RBigButton
          icon={page.orientation === 'portrait' ? <RectangleVertical /> : <RectangleHorizontal />}
          label="Orientation"
          menu={[
            { label: 'Portrait', icon: <RectangleVertical size={16} />, checked: page.orientation === 'portrait', onSelect: () => p.updateSettings({ page: { ...page, orientation: 'portrait' } }) },
            { label: 'Landscape', icon: <RectangleHorizontal size={16} />, checked: page.orientation === 'landscape', onSelect: () => p.updateSettings({ page: { ...page, orientation: 'landscape' } }) },
          ]}
        />
        <RBigButton
          icon={<FileText />}
          label="Size"
          menu={[
            ...PAPER_SIZES.map((s) => ({ label: s.label, checked: page.size === s.id, onSelect: () => p.updateSettings({ page: { ...page, size: s.id, width: s.width, height: s.height } }) })),
            { separator: true },
            { label: 'More paper sizes…', onSelect: p.pageSetup },
          ]}
        />
        <RBigButton icon={<SeparatorHorizontal />} label="Breaks" onClick={() => editor.chain().focus().setPageBreak().run()} tip="Insert a page break" />
      </RibbonGroup>
      <RibbonGroup label="Indent">
        <RRows>
          <div className="rfield">
            <span>Left</span>
            <NumberField value={f.indent / 96} min={0} max={6} step={0.1} suffix="in" width={84} onChange={(v) => editor.chain().focus().setParagraphFormat({ indent: Math.round(v * 96) }).run()} />
          </div>
          <div className="rfield">
            <span>Right</span>
            <NumberField value={0} min={0} max={6} step={0.1} suffix="in" width={84} onChange={(v) => editor.chain().focus().setParagraphFormat({ indentRight: Math.round(v * 96) }).run()} />
          </div>
        </RRows>
      </RibbonGroup>
      <RibbonGroup label="Spacing">
        <RRows>
          <div className="rfield">
            <span>Before</span>
            <NumberField value={f.spaceBefore ?? 0} min={0} max={200} step={2} suffix="pt" width={84} onChange={(v) => editor.chain().focus().setParagraphFormat({ spaceBefore: v }).run()} />
          </div>
          <div className="rfield">
            <span>After</span>
            <NumberField value={f.spaceAfter ?? resolveStyles(settings).normal.after ?? 8} min={0} max={200} step={2} suffix="pt" width={84} onChange={(v) => editor.chain().focus().setParagraphFormat({ spaceAfter: v }).run()} />
          </div>
        </RRows>
      </RibbonGroup>
      <RibbonGroup label="Paragraph">
        <RBigButton icon={<Ruler />} label="Paragraph settings" onClick={p.paragraph} />
        <RBigButton
          icon={<Languages />}
          label="Direction"
          tip="Text direction (for right-to-left languages)"
          menu={[
            { label: 'Left to right', onSelect: () => editor.chain().focus().setParagraphFormat({ dir: 'ltr' }).run() },
            { label: 'Right to left', onSelect: () => editor.chain().focus().setParagraphFormat({ dir: 'rtl' }).setTextAlign('right').run() },
            { label: 'Automatic', onSelect: () => editor.chain().focus().setParagraphFormat({ dir: null }).run() },
          ]}
        />
      </RibbonGroup>
    </>
  );
}

/* ================================================================= DESIGN */

function DesignTab(p: DocsRibbonProps) {
  const { settings } = p;
  return (
    <>
      <RibbonGroup label="Document formatting">
        <div className="styleset-gallery thin-scroll">
          {STYLE_SETS.map((s) => (
            <button
              key={s.id}
              type="button"
              className={`styleset-tile${settings.styleSet === s.id ? ' active' : ''}`}
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => p.updateSettings({ styleSet: s.id, styles: {} })}
              data-tip={`${s.label} style set`}
            >
              <span className="ss-title" style={{ fontFamily: `"${s.styles.title.font}"`, color: s.styles.title.color, fontWeight: s.styles.title.bold ? 700 : 400, fontStyle: s.styles.title.italic ? 'italic' : 'normal' }}>
                Title
              </span>
              <span className="ss-h" style={{ fontFamily: `"${s.styles.h1.font}"`, color: s.styles.h1.color, fontWeight: s.styles.h1.bold ? 700 : 400 }}>
                Heading 1
              </span>
              <span className="ss-lines" />
              <span className="ss-name">{s.label}</span>
            </button>
          ))}
        </div>
      </RibbonGroup>
      <RibbonGroup label="Fonts">
        <RDropdown
          icon={<Type size={24} />}
          label="Body font"
          className="rbig-like"
          menu={CORE_FONTS.filter((f) => f.category !== 'script').map((f) => ({
            label: f.name,
            onSelect: () => p.updateSettings({ styles: { ...settings.styles, normal: { ...settings.styles.normal, font: f.name } } }),
          }))}
        />
        <RDropdown
          icon={<Heading1 size={24} />}
          label="Heading font"
          className="rbig-like"
          menu={CORE_FONTS.filter((f) => f.category !== 'script').map((f) => ({
            label: f.name,
            onSelect: () =>
              p.updateSettings({
                styles: {
                  ...settings.styles,
                  title: { ...settings.styles.title, font: f.name },
                  h1: { ...settings.styles.h1, font: f.name },
                  h2: { ...settings.styles.h2, font: f.name },
                  h3: { ...settings.styles.h3, font: f.name },
                },
              }),
          }))}
        />
      </RibbonGroup>
      <RibbonGroup label="Page background">
        <RDropdown
          icon={<Square size={24} />}
          label="Page colour"
          className="rbig-like"
          panel={(close) => (
            <ColorPicker
              noneLabel="No colour"
              value={settings.pageColor}
              onPick={(c) => {
                p.updateSettings({ pageColor: c });
                close();
              }}
            />
          )}
        />
        <RBigButton
          icon={<Stamp />}
          label="Watermark"
          menu={[
            ...['CONFIDENTIAL', 'DRAFT', 'DO NOT COPY', 'SAMPLE', 'URGENT'].map((w) => ({ label: w, checked: settings.watermark === w, onSelect: () => p.updateSettings({ watermark: w }) })),
            { separator: true },
            {
              label: 'Custom watermark…',
              onSelect: async () => {
                const t = await promptDialog({ title: 'Custom watermark', label: 'Text', value: settings.watermark });
                if (t !== null) p.updateSettings({ watermark: t });
              },
            },
            { label: 'Remove watermark', onSelect: () => p.updateSettings({ watermark: '' }) },
          ]}
        />
      </RibbonGroup>
    </>
  );
}

/* ================================================================= REVIEW */

function ReviewTab(p: DocsRibbonProps) {
  return (
    <>
      <RibbonGroup label="Proofing">
        <RBigButton icon={<SpellCheck />} label="Spelling" tip="Turn spell check on or off" active={p.spellcheck} onClick={p.toggleSpell} />
        <RBigButton icon={<WholeWord />} label="Word count" onClick={p.wordCount} />
      </RibbonGroup>
      <RibbonGroup label="Speech">
        <RBigButton icon={<Volume2 />} label={p.reading ? 'Stop reading' : 'Read aloud'} active={p.reading} tip="Read the document (or selection) aloud" onClick={p.readAloud} />
      </RibbonGroup>
      <RibbonGroup label="Comments">
        <RBigButton icon={<MessageSquarePlus />} label="New comment" onClick={p.addComment} />
        <RBigButton icon={<MessagesSquare />} label="Show comments" active={p.panel === 'comments'} onClick={() => p.togglePanel('comments')} />
      </RibbonGroup>
    </>
  );
}

/* =================================================================== VIEW */

function ViewTab(p: DocsRibbonProps) {
  return (
    <>
      <RibbonGroup label="Views">
        <RBigButton icon={<FileText />} label="Print layout" active={p.view === 'print'} onClick={() => p.setView('print')} />
        <RBigButton icon={<Monitor />} label="Web layout" active={p.view === 'web'} onClick={() => p.setView('web')} />
        <RBigButton icon={<BookOpen />} label="Read mode" active={p.view === 'read'} onClick={() => p.setView('read')} />
        <RBigButton icon={<Focus />} label="Focus" tip="Hide everything except your writing (Esc to exit)" active={p.focusMode} onClick={() => p.setFocusMode(!p.focusMode)} />
      </RibbonGroup>
      <RibbonGroup label="Show">
        <RRows>
          <RButton icon={<PanelLeft size={I} />} label="Navigation pane" showLabel active={p.panel === 'nav'} onClick={() => p.togglePanel('nav')} />
          <RButton icon={<Pilcrow size={I} />} label="Formatting marks" showLabel active={p.showMarks} onClick={() => p.setShowMarks(!p.showMarks)} />
          <RButton icon={<Moon size={I} />} label="Dark page" tip="Show the page in dark colours (doesn't change the file)" showLabel active={p.darkPage} onClick={() => p.setDarkPage(!p.darkPage)} />
        </RRows>
      </RibbonGroup>
      <RibbonGroup label="Zoom">
        <RBigButton icon={<Eye />} label="100%" onClick={() => p.setZoom(1)} />
        <RBigButton icon={<FoldVertical />} label="One page" onClick={p.onePage} />
        <RBigButton icon={<Columns3 />} label="Page width" onClick={p.fitWidth} />
      </RibbonGroup>
    </>
  );
}

/* ================================================================== TABLE */

const TABLE_STYLES: Array<[string | null, string]> = [
  [null, 'Grid'],
  ['accent', 'Blue header'],
  ['banded', 'Banded'],
  ['minimal', 'Minimal'],
  ['plain', 'No borders'],
  ['dark', 'Dark'],
  ['green', 'Green'],
  ['orange', 'Orange'],
];

function TableTab(p: TabProps) {
  const { editor } = p;
  const c = () => editor.chain().focus();
  return (
    <>
      <RibbonGroup label="Table style">
        <div className="table-style-gallery">
          {TABLE_STYLES.map(([id, label]) => (
            <button key={label} type="button" className={`tsg-tile tsg-${id ?? 'grid'}`} onMouseDown={(e) => e.preventDefault()} onClick={() => c().setTableStyle(id).run()} data-tip={label}>
              <span />
              <span />
              <span />
              <span />
              <span />
              <span />
              <span />
              <span />
              <span />
            </button>
          ))}
        </div>
      </RibbonGroup>
      <RibbonGroup label="Rows & columns">
        <RRows>
          <RButton icon={<ArrowUpToLine size={I} />} label="Insert above" showLabel onClick={() => c().addRowBefore().run()} />
          <RButton icon={<ArrowDownToLine size={I} />} label="Insert below" showLabel onClick={() => c().addRowAfter().run()} />
          <RButton icon={<Trash2 size={I} />} label="Delete row" showLabel onClick={() => c().deleteRow().run()} />
        </RRows>
        <RRows>
          <RButton icon={<BetweenVerticalStart size={I} />} label="Insert left" showLabel onClick={() => c().addColumnBefore().run()} />
          <RButton icon={<BetweenHorizontalStart size={I} />} label="Insert right" showLabel onClick={() => c().addColumnAfter().run()} />
          <RButton icon={<Trash2 size={I} />} label="Delete column" showLabel onClick={() => c().deleteColumn().run()} />
        </RRows>
      </RibbonGroup>
      <RibbonGroup label="Merge">
        <RBigButton icon={<TableCellsMerge />} label="Merge cells" onClick={() => c().mergeCells().run()} />
        <RBigButton icon={<TableCellsSplit />} label="Split cell" onClick={() => c().splitCell().run()} />
      </RibbonGroup>
      <RibbonGroup label="Cells">
        <RRows>
          <RButton icon={<LayoutTemplate size={I} />} label="Header row" showLabel onClick={() => c().toggleHeaderRow().run()} />
          <RButton icon={<SquareSplitHorizontal size={I} />} label="Header column" showLabel onClick={() => c().toggleHeaderColumn().run()} />
          <RDropdown
            icon={<Droplet size={I} />}
            label="Shading"
            panel={(close) => (
              <ColorPicker
                noneLabel="No colour"
                onPick={(col) => {
                  c().setCellBackground(col).run();
                  close();
                }}
              />
            )}
          />
        </RRows>
        <RRows>
          <RButton icon={<ArrowUpToLine size={I} />} label="Align top" showLabel onClick={() => c().setCellVAlign('top').run()} />
          <RButton icon={<AlignCenter size={I} />} label="Align middle" showLabel onClick={() => c().setCellVAlign('middle').run()} />
          <RButton icon={<ArrowDownToLine size={I} />} label="Align bottom" showLabel onClick={() => c().setCellVAlign('bottom').run()} />
        </RRows>
      </RibbonGroup>
      <RibbonGroup label="Table">
        <RBigButton icon={<Trash2 />} label="Delete table" onClick={() => c().deleteTable().run()} />
      </RibbonGroup>
    </>
  );
}

/* ================================================================ PICTURE */

function PictureTab(p: TabProps) {
  const { editor, f } = p;
  const img = f.image;
  if (!img) return null;
  const set = (a: Parameters<Editor['commands']['setImageAttrs']>[0]) => editor.chain().focus().setImageAttrs(a).run();
  const wrap: Array<[string, string, ReactNode]> = [
    ['inline', 'In line with text', <AlignLeft key="i" />],
    ['left', 'Wrap: image on left', <AlignLeft key="l" />],
    ['right', 'Wrap: image on right', <AlignRight key="r" />],
    ['center', 'Centred on its own line', <AlignCenter key="c" />],
  ];
  return (
    <>
      <RibbonGroup label="Position">
        {wrap.map(([id, label, icon]) => (
          <RBigButton key={id} icon={icon} label={label.split(':')[0].replace('In line with text', 'In line').replace('Centred on its own line', 'Centre')} tip={label} active={img.align === id} onClick={() => set({ align: id as 'inline' })} />
        ))}
      </RibbonGroup>
      <RibbonGroup label="Size">
        <RRows>
          <div className="rfield">
            <span>Width</span>
            <NumberField value={img.width ?? 300} min={16} max={2400} step={10} suffix="px" width={96} onChange={(v) => set({ width: v })} />
          </div>
          <RButton icon={<Maximize2 size={I} />} label="Fit to page width" showLabel onClick={() => set({ width: 624 })} />
        </RRows>
      </RibbonGroup>
      <RibbonGroup label="Style">
        <RBigButton icon={<Square />} label="Border" active={img.border} onClick={() => set({ border: !img.border })} />
        <RBigButton icon={<Circle />} label="Rounded" active={img.rounded} onClick={() => set({ rounded: !img.rounded })} />
        <RBigButton icon={<Droplet />} label="Shadow" active={img.shadow} onClick={() => set({ shadow: !img.shadow })} />
      </RibbonGroup>
      <RibbonGroup label="Accessibility">
        <RBigButton
          icon={<Type />}
          label="Alt text"
          tip="Describe the picture for screen readers"
          onClick={async () => {
            const t = await promptDialog({ title: 'Alternative text', label: 'Describe this picture for people who can’t see it', value: img.alt, multiline: true });
            if (t !== null) set({ alt: t });
          }}
        />
      </RibbonGroup>
      <RibbonGroup label="Picture">
        <RBigButton icon={<Trash2 />} label="Delete" onClick={() => editor.chain().focus().deleteSelection().run()} />
      </RibbonGroup>
    </>
  );
}
