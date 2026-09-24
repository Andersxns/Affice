import type { Editor } from '@tiptap/react';
import { useMemo, useState } from 'react';
import { Field, NumberField, Segmented, Select, Checkbox } from '@/ui/controls';
import { Modal, openDialog } from '@/ui/dialog';
import { cmToIn, inToCm } from '@/lib/utils';
import { renderMathToHtml } from './editor/extensions';
import { MARGIN_PRESETS, PAPER_SIZES, type DocSettings, type HeaderFooter, type PageSetup } from './model';

/* ------------------------------------------------------------------ link */

export function linkDialog(editor: Editor): void {
  const prev = (editor.getAttributes('link').href as string | undefined) ?? '';
  const { from, to } = editor.state.selection;
  const selected = editor.state.doc.textBetween(from, to, ' ');
  void openDialog<{ href: string; text: string } | null>((close) => <LinkBody close={close} href={prev} text={selected} />).then((r) => {
    if (!r) return;
    if (!r.href) {
      editor.chain().focus().extendMarkRange('link').unsetLink().run();
      return;
    }
    let href = r.href.trim();
    if (!/^[a-z][a-z0-9+.-]*:/i.test(href) && !href.startsWith('#')) href = href.includes('@') && !href.includes('/') ? `mailto:${href}` : `https://${href}`;
    if (editor.state.selection.empty && !editor.isActive('link')) {
      const text = r.text || href;
      editor.chain().focus().insertContent({ type: 'text', text, marks: [{ type: 'link', attrs: { href } }] }).run();
    } else {
      editor.chain().focus().extendMarkRange('link').setLink({ href }).run();
    }
  });
}

function LinkBody({ close, href, text }: { close: (v?: { href: string; text: string } | null) => void; href: string; text: string }) {
  const [h, setH] = useState(href);
  const [t, setT] = useState(text);
  return (
    <Modal
      title={href ? 'Edit link' : 'Insert link'}
      onClose={() => close(null)}
      width={460}
      footer={
        <>
          {href && (
            <button className="btn btn-ghost" onClick={() => close({ href: '', text: t })}>
              Remove link
            </button>
          )}
          <span className="spacer" />
          <button className="btn" onClick={() => close(null)}>
            Cancel
          </button>
          <button className="btn btn-primary" onClick={() => close({ href: h, text: t })} disabled={!h.trim()}>
            OK
          </button>
        </>
      }
    >
      <div className="col" style={{ gap: 12 }}>
        {!text && (
          <Field label="Text to display">
            <input className="input" value={t} onChange={(e) => setT(e.target.value)} placeholder="Link text" />
          </Field>
        )}
        <Field label="Address" hint="A web address, an email address, or #heading-name.">
          <input
            className="input"
            data-autofocus
            value={h}
            onChange={(e) => setH(e.target.value)}
            placeholder="https://example.com"
            onKeyDown={(e) => e.key === 'Enter' && h.trim() && close({ href: h, text: t })}
          />
        </Field>
      </div>
    </Modal>
  );
}

/* ------------------------------------------------------------ page setup */

export function pageSetupDialog(settings: DocSettings, apply: (patch: Partial<DocSettings>) => void): void {
  void openDialog<PageSetup | null>((close) => <PageSetupBody close={close} page={settings.page} />).then((page) => {
    if (page) apply({ page });
  });
}

function PageSetupBody({ close, page }: { close: (v?: PageSetup | null) => void; page: PageSetup }) {
  const [p, setP] = useState<PageSetup>(JSON.parse(JSON.stringify(page)));
  const metric = !['letter', 'legal', 'executive', 'tabloid'].includes(p.size);
  const [unit, setUnit] = useState<'in' | 'cm'>(metric ? 'cm' : 'in');
  const toU = (v: number) => (unit === 'cm' ? Math.round(inToCm(v) * 100) / 100 : Math.round(v * 100) / 100);
  const fromU = (v: number) => (unit === 'cm' ? cmToIn(v) : v);
  const setM = (k: keyof PageSetup['margins'], v: number) => setP({ ...p, margins: { ...p.margins, [k]: fromU(v) } });
  return (
    <Modal
      title="Page setup"
      onClose={() => close(null)}
      width={520}
      footer={
        <>
          <button className="btn" onClick={() => close(null)}>
            Cancel
          </button>
          <button className="btn btn-primary" onClick={() => close(p)}>
            Apply
          </button>
        </>
      }
    >
      <div className="col" style={{ gap: 16 }}>
        <div className="row" style={{ justifyContent: 'space-between' }}>
          <Segmented
            value={p.orientation}
            onChange={(o) => setP({ ...p, orientation: o })}
            options={[
              { value: 'portrait', label: 'Portrait' },
              { value: 'landscape', label: 'Landscape' },
            ]}
          />
          <Segmented
            size="sm"
            value={unit}
            onChange={setUnit}
            options={[
              { value: 'in', label: 'Inches' },
              { value: 'cm', label: 'Centimetres' },
            ]}
          />
        </div>
        <Field label="Paper size">
          <Select
            value={p.size}
            onChange={(id) => {
              const s = PAPER_SIZES.find((x) => x.id === id);
              if (s) setP({ ...p, size: id, width: s.width, height: s.height });
              else setP({ ...p, size: 'custom' });
            }}
            options={[...PAPER_SIZES.map((s) => ({ value: s.id, label: s.label })), { value: 'custom', label: 'Custom size' }]}
          />
        </Field>
        {p.size === 'custom' && (
          <div className="row" style={{ gap: 16 }}>
            <Field label={`Width (${unit})`}>
              <NumberField value={toU(p.width)} onChange={(v) => setP({ ...p, width: fromU(v) })} min={1} max={unit === 'cm' ? 120 : 48} step={unit === 'cm' ? 0.5 : 0.1} width={110} />
            </Field>
            <Field label={`Height (${unit})`}>
              <NumberField value={toU(p.height)} onChange={(v) => setP({ ...p, height: fromU(v) })} min={1} max={unit === 'cm' ? 120 : 48} step={unit === 'cm' ? 0.5 : 0.1} width={110} />
            </Field>
          </div>
        )}
        <div>
          <div className="field-label" style={{ marginBottom: 6 }}>
            Margins
          </div>
          <div className="row" style={{ flexWrap: 'wrap', gap: 6, marginBottom: 10 }}>
            {MARGIN_PRESETS.map((mp) => (
              <button key={mp.id} className="btn btn-sm" onClick={() => setP({ ...p, margins: { ...mp.m } })}>
                {mp.label}
              </button>
            ))}
          </div>
          <div className="margin-grid">
            {(['top', 'bottom', 'left', 'right'] as const).map((k) => (
              <Field key={k} label={`${k[0].toUpperCase()}${k.slice(1)} (${unit})`}>
                <NumberField value={toU(p.margins[k])} onChange={(v) => setM(k, v)} min={0} max={unit === 'cm' ? 15 : 6} step={unit === 'cm' ? 0.25 : 0.1} width={110} />
              </Field>
            ))}
          </div>
        </div>
      </div>
    </Modal>
  );
}

/* --------------------------------------------------------- paragraph */

export interface ParagraphValues {
  indent: number;
  indentRight: number;
  firstLine: number;
  lineHeight: number | null;
  spaceBefore: number | null;
  spaceAfter: number | null;
  align: string;
  dir: 'ltr' | 'rtl' | null;
}

export function paragraphDialog(editor: Editor): void {
  const { $from } = editor.state.selection;
  let node = $from.parent;
  for (let d = $from.depth; d > 0; d--) {
    const n = $from.node(d);
    if (n.type.name === 'paragraph' || n.type.name === 'heading') {
      node = n;
      break;
    }
  }
  const a = node.attrs;
  const init: ParagraphValues = {
    indent: a.indent ?? 0,
    indentRight: a.indentRight ?? 0,
    firstLine: a.firstLine ?? 0,
    lineHeight: a.lineHeight ?? null,
    spaceBefore: a.spaceBefore ?? null,
    spaceAfter: a.spaceAfter ?? null,
    align: a.textAlign ?? 'left',
    dir: a.dir ?? null,
  };
  void openDialog<ParagraphValues | null>((close) => <ParagraphBody close={close} init={init} />).then((v) => {
    if (!v) return;
    editor
      .chain()
      .focus()
      .setParagraphFormat({ indent: v.indent, indentRight: v.indentRight, firstLine: v.firstLine, lineHeight: v.lineHeight, spaceBefore: v.spaceBefore, spaceAfter: v.spaceAfter, dir: v.dir })
      .setTextAlign(v.align)
      .run();
  });
}

function ParagraphBody({ close, init }: { close: (v?: ParagraphValues | null) => void; init: ParagraphValues }) {
  const [v, setV] = useState(init);
  const set = <K extends keyof ParagraphValues>(k: K, x: ParagraphValues[K]) => setV((o) => ({ ...o, [k]: x }));
  const special = v.firstLine > 0 ? 'first' : v.firstLine < 0 ? 'hanging' : 'none';
  return (
    <Modal
      title="Paragraph"
      onClose={() => close(null)}
      width={520}
      footer={
        <>
          <button className="btn" onClick={() => close(null)}>
            Cancel
          </button>
          <button className="btn btn-primary" onClick={() => close(v)}>
            OK
          </button>
        </>
      }
    >
      <div className="dialog-grid">
        <Field label="Alignment">
          <Select
            value={v.align}
            onChange={(x) => set('align', x)}
            options={[
              { value: 'left', label: 'Left' },
              { value: 'center', label: 'Centred' },
              { value: 'right', label: 'Right' },
              { value: 'justify', label: 'Justified' },
            ]}
          />
        </Field>
        <Field label="Text direction">
          <Select
            value={v.dir ?? 'auto'}
            onChange={(x) => set('dir', x === 'auto' ? null : (x as 'ltr' | 'rtl'))}
            options={[
              { value: 'auto', label: 'Automatic' },
              { value: 'ltr', label: 'Left to right' },
              { value: 'rtl', label: 'Right to left' },
            ]}
          />
        </Field>
        <Field label="Indent left (in)">
          <NumberField value={v.indent / 96} onChange={(x) => set('indent', Math.round(x * 96))} min={0} max={6} step={0.1} width={120} />
        </Field>
        <Field label="Indent right (in)">
          <NumberField value={v.indentRight / 96} onChange={(x) => set('indentRight', Math.round(x * 96))} min={0} max={6} step={0.1} width={120} />
        </Field>
        <Field label="Special">
          <Select
            value={special}
            onChange={(x) => set('firstLine', x === 'none' ? 0 : x === 'first' ? Math.abs(v.firstLine) || 48 : -(Math.abs(v.firstLine) || 48))}
            options={[
              { value: 'none', label: '(none)' },
              { value: 'first', label: 'First line' },
              { value: 'hanging', label: 'Hanging' },
            ]}
          />
        </Field>
        <Field label="By (in)">
          <NumberField
            value={Math.abs(v.firstLine) / 96}
            onChange={(x) => set('firstLine', (special === 'hanging' ? -1 : 1) * Math.round(x * 96))}
            min={0}
            max={4}
            step={0.1}
            width={120}
            disabled={special === 'none'}
          />
        </Field>
        <Field label="Spacing before (pt)">
          <NumberField value={v.spaceBefore ?? 0} onChange={(x) => set('spaceBefore', x)} min={0} max={200} step={2} width={120} />
        </Field>
        <Field label="Spacing after (pt)">
          <NumberField value={v.spaceAfter ?? 8} onChange={(x) => set('spaceAfter', x)} min={0} max={200} step={2} width={120} />
        </Field>
        <Field label="Line spacing">
          <Select
            value={String(v.lineHeight ?? 'default')}
            onChange={(x) => set('lineHeight', x === 'default' ? null : Number(x))}
            options={[
              { value: 'default', label: 'Style default' },
              { value: '1', label: 'Single' },
              { value: '1.15', label: '1.15' },
              { value: '1.5', label: '1.5 lines' },
              { value: '2', label: 'Double' },
              { value: '2.5', label: '2.5' },
              { value: '3', label: 'Triple' },
            ]}
          />
        </Field>
      </div>
    </Modal>
  );
}

/* ------------------------------------------------------- header & footer */

export function headerFooterDialog(settings: DocSettings, apply: (patch: Partial<DocSettings>) => void, focus: 'header' | 'footer' = 'header'): void {
  void openDialog<Partial<DocSettings> | null>((close) => <HeaderFooterBody close={close} settings={settings} focus={focus} />).then((r) => {
    if (r) apply(r);
  });
}

function HFRow({ label, value, onChange }: { label: string; value: HeaderFooter; onChange: (v: HeaderFooter) => void }) {
  const [focused, setFocused] = useState<keyof HeaderFooter>('center');
  const insert = (token: string) => onChange({ ...value, [focused]: `${value[focused]}${value[focused] ? ' ' : ''}${token}` });
  return (
    <div className="hf-row">
      <div className="row" style={{ justifyContent: 'space-between', marginBottom: 6 }}>
        <span className="field-label">{label}</span>
        <div className="row" style={{ gap: 4 }}>
          <button className="btn btn-sm" onClick={() => insert('{PAGE}')}>
            Page number
          </button>
          <button className="btn btn-sm" onClick={() => insert('Page {PAGE} of {PAGES}')}>
            Page X of Y
          </button>
          <button className="btn btn-sm" onClick={() => insert('{DATE}')}>
            Date
          </button>
          <button className="btn btn-sm" onClick={() => insert('{TITLE}')}>
            Title
          </button>
        </div>
      </div>
      <div className="hf-inputs">
        {(['left', 'center', 'right'] as const).map((k) => (
          <input
            key={k}
            className="input"
            style={{ textAlign: k }}
            placeholder={k === 'left' ? 'Left' : k === 'center' ? 'Centre' : 'Right'}
            value={value[k]}
            onFocus={() => setFocused(k)}
            onChange={(e) => onChange({ ...value, [k]: e.target.value })}
          />
        ))}
      </div>
    </div>
  );
}

function HeaderFooterBody({ close, settings, focus }: { close: (v?: Partial<DocSettings> | null) => void; settings: DocSettings; focus: string }) {
  const [header, setHeader] = useState(settings.header);
  const [footer, setFooter] = useState(settings.footer);
  const [first, setFirst] = useState(settings.differentFirstPage);
  return (
    <Modal
      title="Header & footer"
      onClose={() => close(null)}
      width={620}
      footer={
        <>
          <button
            className="btn btn-ghost"
            onClick={() => {
              setHeader({ left: '', center: '', right: '' });
              setFooter({ left: '', center: '', right: '' });
            }}
          >
            Clear all
          </button>
          <span className="spacer" />
          <button className="btn" onClick={() => close(null)}>
            Cancel
          </button>
          <button className="btn btn-primary" onClick={() => close({ header, footer, differentFirstPage: first })}>
            Apply
          </button>
        </>
      }
    >
      <div className="col" style={{ gap: 18 }} data-focus={focus}>
        <HFRow label="Header" value={header} onChange={setHeader} />
        <HFRow label="Footer" value={footer} onChange={setFooter} />
        <Checkbox checked={first} onChange={setFirst} label="Different first page (no header/footer on page 1)" />
        <p className="muted small">Fields like {'{PAGE}'} and {'{PAGES}'} are filled in automatically on every page, in print and in exported files.</p>
      </div>
    </Modal>
  );
}

/* ------------------------------------------------------------ word count */

export function wordCountDialog(editor: Editor, pages: number): void {
  const text = editor.state.doc.textBetween(0, editor.state.doc.content.size, '\n', ' ');
  const { from, to, empty } = editor.state.selection;
  const selText = empty ? '' : editor.state.doc.textBetween(from, to, '\n', ' ');
  const count = (t: string) => {
    const words = (t.match(/[\p{L}\p{N}][\p{L}\p{N}'’_-]*/gu) ?? []).length;
    const chars = t.replace(/\n/g, '').length;
    const charsNoSpace = t.replace(/\s/g, '').length;
    const paras = t.split(/\n+/).filter((p) => p.trim()).length;
    const sentences = (t.match(/[^.!?…]+[.!?…]+/g) ?? []).length || (words ? 1 : 0);
    return { words, chars, charsNoSpace, paras, sentences };
  };
  const all = count(text);
  const sel = selText ? count(selText) : null;
  const minutes = (n: number, wpm: number) => {
    const m = n / wpm;
    return m < 1 ? '< 1 min' : `${Math.round(m)} min`;
  };
  const grade = (() => {
    // Flesch reading ease (English approximation)
    const syll = (text.toLowerCase().match(/[aeiouy]+/g) ?? []).length;
    if (!all.words || !all.sentences) return null;
    const score = 206.835 - 1.015 * (all.words / all.sentences) - 84.6 * (syll / all.words);
    const s = Math.max(0, Math.min(100, Math.round(score)));
    const label = s >= 80 ? 'Very easy' : s >= 60 ? 'Easy' : s >= 50 ? 'Fairly difficult' : s >= 30 ? 'Difficult' : 'Very difficult';
    return `${s} / 100 · ${label}`;
  })();
  void openDialog((close) => (
    <Modal
      title="Word count"
      onClose={() => close()}
      width={420}
      footer={
        <button className="btn btn-primary" onClick={() => close()}>
          Close
        </button>
      }
    >
      <div className="stats-grid selectable">
        <span>Pages</span>
        <b>{pages.toLocaleString()}</b>
        <span>Words</span>
        <b>
          {all.words.toLocaleString()}
          {sel && <span className="muted"> ({sel.words.toLocaleString()} selected)</span>}
        </b>
        <span>Characters (no spaces)</span>
        <b>{all.charsNoSpace.toLocaleString()}</b>
        <span>Characters (with spaces)</span>
        <b>{all.chars.toLocaleString()}</b>
        <span>Paragraphs</span>
        <b>{all.paras.toLocaleString()}</b>
        <span>Sentences</span>
        <b>{all.sentences.toLocaleString()}</b>
        <span>Reading time</span>
        <b>{minutes(all.words, 238)}</b>
        <span>Speaking time</span>
        <b>{minutes(all.words, 140)}</b>
        {grade && (
          <>
            <span>Readability</span>
            <b>{grade}</b>
          </>
        )}
      </div>
    </Modal>
  ));
}

/* --------------------------------------------------------------- symbols */

const SYMBOL_SETS: Record<string, string> = {
  Common: '©®™§¶†‡•…–—‘’“”«»‹›°±×÷≈≠≤≥∞√∑∏∫∂µΩπ€£¥¢₹₽₩฿¤½⅓⅔¼¾⅛',
  Arrows: '←↑→↓↔↕⇐⇑⇒⇓⇔↖↗↘↙↩↪⟵⟶⟷➔➜➤▶◀▲▼',
  Math: '∀∁∂∃∄∅∆∇∈∉∋∌∏∑−∓∗∘∙√∛∜∝∞∟∠∡∢∣∤∥∦∧∨∩∪∫∬∭∮∴∵∶∷∼≃≅≈≡≢≤≥≪≫⊂⊃⊆⊇⊕⊗⊥⋅',
  Greek: 'ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩαβγδεζηθικλμνξοπρστυφχψω',
  Shapes: '■□▪▫▬▭▮▯▰▱▲△▴▵▶▷▸▹►▻▼▽▾▿◀◁◂◃◄◅◆◇◈◉◊○◌◍◎●◐◑◒◓◔◕◖◗★☆✓✔✕✖✗✘',
  Emoji: '😀😂😊😍🤔😎😢😡👍👎👏🙏💪🎉🎂🎁❤️💔⭐🔥💡📌📎📅📈📉✅❌⚠️❓❗➕➖🚀🌍🌟☀️🌙☕🍕🏆🎯',
};

export function symbolDialog(editor: Editor): void {
  void openDialog<string | null>((close) => <SymbolBody close={close} />).then((s) => {
    if (s) editor.chain().focus().insertContent(s).run();
  });
}

function SymbolBody({ close }: { close: (v?: string | null) => void }) {
  const [set, setSet] = useState('Common');
  const chars = useMemo(() => Array.from(SYMBOL_SETS[set]), [set]);
  return (
    <Modal title="Insert symbol" onClose={() => close(null)} width={560}>
      <div className="row" style={{ marginBottom: 12, flexWrap: 'wrap' }}>
        {Object.keys(SYMBOL_SETS).map((k) => (
          <button key={k} className={`btn btn-sm${k === set ? ' btn-primary' : ''}`} onClick={() => setSet(k)}>
            {k}
          </button>
        ))}
      </div>
      <div className="symbol-grid">
        {chars.map((c, i) => (
          <button key={`${c}${i}`} className="symbol-cell" title={`U+${c.codePointAt(0)!.toString(16).toUpperCase().padStart(4, '0')}`} onClick={() => close(c)}>
            {c}
          </button>
        ))}
      </div>
    </Modal>
  );
}

/* ----------------------------------------------------------------- math */

const MATH_SNIPPETS: Array<[string, string]> = [
  ['Fraction', '\\frac{a}{b}'],
  ['Square root', '\\sqrt{x}'],
  ['Power', 'x^{2}'],
  ['Subscript', 'x_{i}'],
  ['Sum', '\\sum_{i=1}^{n} x_i'],
  ['Integral', '\\int_{a}^{b} f(x)\\,dx'],
  ['Limit', '\\lim_{x \\to \\infty} f(x)'],
  ['Matrix', '\\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix}'],
  ['Quadratic', 'x = \\frac{-b \\pm \\sqrt{b^2-4ac}}{2a}'],
  ['Pythagoras', 'a^2 + b^2 = c^2'],
  ['Euler', 'e^{i\\pi} + 1 = 0'],
  ['Greek', '\\alpha \\beta \\gamma \\theta \\pi'],
];

export function equationDialog(editor: Editor, initial = '', onDone?: (latex: string, display: boolean) => void): void {
  void openDialog<{ latex: string; display: boolean } | null>((close) => <EquationBody close={close} initial={initial} />).then((r) => {
    if (!r) return;
    if (onDone) onDone(r.latex, r.display);
    else editor.chain().focus().insertMath(r.latex, r.display).run();
  });
}

function EquationBody({ close, initial }: { close: (v?: { latex: string; display: boolean } | null) => void; initial: string }) {
  const [latex, setLatex] = useState(initial || 'x = \\frac{-b \\pm \\sqrt{b^2-4ac}}{2a}');
  const [display, setDisplay] = useState(true);
  const preview = useMemo(() => renderMathToHtml(latex, true), [latex]);
  return (
    <Modal
      title="Equation"
      onClose={() => close(null)}
      width={620}
      footer={
        <>
          <Checkbox checked={display} onChange={setDisplay} label="On its own line" />
          <span className="spacer" />
          <button className="btn" onClick={() => close(null)}>
            Cancel
          </button>
          <button className="btn btn-primary" onClick={() => close({ latex, display })}>
            Insert
          </button>
        </>
      }
    >
      <div className="col" style={{ gap: 12 }}>
        <div className="math-snippets">
          {MATH_SNIPPETS.map(([label, tex]) => (
            <button key={label} className="btn btn-sm" onClick={() => setLatex((l) => (l && l !== initial ? `${l} ${tex}` : tex))}>
              {label}
            </button>
          ))}
        </div>
        <Field label="LaTeX" hint="Type LaTeX math — the preview updates as you type.">
          <textarea className="input mono" rows={3} value={latex} onChange={(e) => setLatex(e.target.value)} spellCheck={false} />
        </Field>
        <div className="math-preview" dangerouslySetInnerHTML={{ __html: preview }} />
      </div>
    </Modal>
  );
}
