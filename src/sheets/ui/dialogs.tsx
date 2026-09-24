import { ArrowDown, ArrowUp, Plus, Search, Trash2 } from 'lucide-react';
import { useMemo, useState, type ReactNode } from 'react';
import { ColorField } from '@/ui/color';
import { Checkbox, Field, NumberField, Segmented, Select } from '@/ui/controls';
import { Modal, openDialog } from '@/ui/dialog';
import { FontFamilyCombo, FontSizeCombo } from '@/ui/fontControls';
import { allFunctions, type Category } from '../engine/functions';
import { isMatrix } from '../engine/values';
import { formatValue } from '../format/numfmt';
import { addrName, colName, parseRange, rangeName, type Range } from '../model/address';
import { DEFAULT_PRINT, type BorderStyle, type CellStyle, type ChartSpec, type ChartType, type CondFormat, type CondType, type CompareOp, type DefinedName, type SheetPrint, type Validation, type ValidationType } from '../model/types';
import type { SheetDoc } from '../doc';
import { currentRegion, dataRange, looksLikeHeader, newId, removeDuplicates, setValidation, sortRange, textToColumns, upsertCF, deleteCF, reorderCF, setNames, upsertChart, type SortKey } from '../ops/data';
import { PALETTES } from '../charts/chartConfig';
import { buildFormat, categorize, CATEGORIES, CURRENCIES, CUSTOM_PRESETS, DATE_FORMATS, FRACTION_FORMATS, SPECIAL_FORMATS, TIME_FORMATS, type NegativeStyle, type NumCategory } from './numberFormats';

const Footer = ({ onCancel, onOk, okLabel = 'OK', disabled, extra }: { onCancel: () => void; onOk: () => void; okLabel?: string; disabled?: boolean; extra?: ReactNode }) => (
  <>
    {extra}
    <span className="spacer" />
    <button className="btn" onClick={onCancel}>
      Cancel
    </button>
    <button className="btn btn-primary" disabled={disabled} onClick={onOk}>
      {okLabel}
    </button>
  </>
);

/* ============================================================ format cells */

type FmtTab = 'number' | 'alignment' | 'font' | 'border' | 'fill' | 'protection';

export function formatCellsDialog(doc: SheetDoc, tab: FmtTab = 'number'): Promise<void> {
  return openDialog<void>((close) => <FormatCells doc={doc} initialTab={tab} close={() => close()} />).then(() => undefined);
}

function FormatCells({ doc, initialTab, close }: { doc: SheetDoc; initialTab: FmtTab; close: () => void }) {
  const base = doc.activeStyle();
  const [tab, setTab] = useState<FmtTab>(initialTab);
  const [st, setSt] = useState<CellStyle>({ ...base });
  const initial = categorize(base.numFmt);
  const [cat, setCat] = useState<NumCategory>(initial.cat);
  const [decimals, setDecimals] = useState(initial.opts.decimals || (initial.cat === 'general' ? 2 : 0));
  const [thousands, setThousands] = useState(initial.opts.thousands);
  const [negative, setNegative] = useState<NegativeStyle>(initial.opts.negative);
  const [symbol, setSymbol] = useState(initial.opts.symbol);
  const [pick, setPick] = useState(base.numFmt ?? 'General');
  const [custom, setCustom] = useState(base.numFmt ?? 'General');
  const [borderStyle, setBorderStyle] = useState<BorderStyle>('thin');
  const [borderColor, setBorderColor] = useState<string | null>(null);
  const [borderOps, setBorderOps] = useState<string[]>([]);
  const sample = doc.value(doc.sheet, doc.sel.active.r, doc.sel.active.c);
  const fmt = cat === 'custom' ? custom : ['date', 'time', 'fraction', 'special'].includes(cat) ? pick : buildFormat(cat, { decimals, thousands, negative, symbol });
  const preview = sample === undefined || sample === null ? '' : formatValue(sample, fmt);
  const set = (p: Partial<CellStyle>) => setSt((s) => ({ ...s, ...p }));

  const apply = () => {
    const patch: Partial<CellStyle> = {};
    const keys: (keyof CellStyle)[] = ['font', 'size', 'bold', 'italic', 'underline', 'strike', 'color', 'fill', 'hAlign', 'vAlign', 'wrap', 'shrink', 'indent', 'rotation', 'locked', 'hideFormula'];
    for (const k of keys) if (st[k] !== base[k]) (patch as Record<string, unknown>)[k] = st[k] ?? undefined;
    if (fmt !== (base.numFmt ?? 'General')) patch.numFmt = fmt === 'General' ? undefined : fmt;
    if (Object.keys(patch).length) doc.applyStyle(patch, 'Format cells');
    const bOps = borderOps as Parameters<SheetDoc['applyBorders']>[0][];
    for (const op of bOps) doc.applyBorders(op, borderStyle, borderColor ?? undefined);
    close();
  };

  const tabs: [FmtTab, string][] = [
    ['number', 'Number'],
    ['alignment', 'Alignment'],
    ['font', 'Font'],
    ['border', 'Border'],
    ['fill', 'Fill'],
    ['protection', 'Protection'],
  ];

  return (
    <Modal title="Format cells" onClose={close} width={640} footer={<Footer onCancel={close} onOk={apply} />}>
      <div className="dlg-tabs">
        {tabs.map(([id, label]) => (
          <button key={id} className={`dlg-tab${tab === id ? ' active' : ''}`} onClick={() => setTab(id)}>
            {label}
          </button>
        ))}
      </div>
      <div className="dlg-tab-body fmt-cells">
        {tab === 'number' && (
          <div className="fmt-number">
            <div className="fmt-cats" role="listbox">
              {CATEGORIES.map((c) => (
                <button key={c.id} className={`fmt-cat${cat === c.id ? ' active' : ''}`} onClick={() => setCat(c.id)}>
                  {c.label}
                </button>
              ))}
            </div>
            <div className="fmt-opts">
              <div className="fmt-sample">
                <span className="muted">Sample</span>
                <span className="fmt-sample-value" style={{ color: preview && preview.color ? preview.color : undefined }}>
                  {preview ? preview.text || ' ' : '—'}
                </span>
              </div>
              {(cat === 'number' || cat === 'currency' || cat === 'percent' || cat === 'scientific' || cat === 'accounting') && (
                <Field label="Decimal places" inline>
                  <NumberField value={decimals} min={0} max={15} precision={0} onChange={(v) => setDecimals(Math.round(v))} />
                </Field>
              )}
              {cat === 'number' && <Checkbox checked={thousands} onChange={setThousands} label="Use 1000 separator (,)" />}
              {(cat === 'currency' || cat === 'accounting') && (
                <Field label="Symbol" inline>
                  <Select value={symbol} onChange={setSymbol} options={CURRENCIES.map((c) => ({ value: c.symbol, label: c.label }))} width={200} />
                </Field>
              )}
              {(cat === 'number' || cat === 'currency') && (
                <Field label="Negative numbers">
                  <div className="neg-list">
                    {(['minus', 'red', 'parens', 'redParens'] as NegativeStyle[]).map((n) => (
                      <button key={n} className={`neg-opt${negative === n ? ' active' : ''}${n.startsWith('red') ? ' red' : ''}`} onClick={() => setNegative(n)}>
                        {n === 'minus' || n === 'red' ? '-1,234.10' : '(1,234.10)'}
                      </button>
                    ))}
                  </div>
                </Field>
              )}
              {(cat === 'date' || cat === 'time') && (
                <div className="fmt-list">
                  {(cat === 'date' ? DATE_FORMATS : TIME_FORMATS).map((f) => (
                    <button key={f} className={`fmt-list-item${pick === f ? ' active' : ''}`} onClick={() => setPick(f)}>
                      {formatValue(typeof sample === 'number' ? sample : 45300.5625, f).text}
                      <span className="muted">{f}</span>
                    </button>
                  ))}
                </div>
              )}
              {cat === 'fraction' && (
                <div className="fmt-list">
                  {FRACTION_FORMATS.map(([f, label]) => (
                    <button key={f} className={`fmt-list-item${pick === f ? ' active' : ''}`} onClick={() => setPick(f)}>
                      {label}
                    </button>
                  ))}
                </div>
              )}
              {cat === 'special' && (
                <div className="fmt-list">
                  {SPECIAL_FORMATS.map(([f, label]) => (
                    <button key={f} className={`fmt-list-item${pick === f ? ' active' : ''}`} onClick={() => setPick(f)}>
                      {label}
                      <span className="muted">{f}</span>
                    </button>
                  ))}
                </div>
              )}
              {cat === 'custom' && (
                <>
                  <Field label="Type">
                    <input className="input mono" value={custom} onChange={(e) => setCustom(e.target.value)} />
                  </Field>
                  <div className="fmt-list compact">
                    {CUSTOM_PRESETS.map((f) => (
                      <button key={f} className={`fmt-list-item mono${custom === f ? ' active' : ''}`} onClick={() => setCustom(f)}>
                        {f}
                      </button>
                    ))}
                  </div>
                </>
              )}
              <p className="muted small">{CATEGORIES.find((c) => c.id === cat)?.hint}</p>
            </div>
          </div>
        )}
        {tab === 'alignment' && (
          <div className="fmt-grid">
            <Field label="Horizontal">
              <Select
                value={st.hAlign ?? 'general'}
                onChange={(v) => set({ hAlign: v === 'general' ? undefined : v })}
                options={[
                  { value: 'general', label: 'General' },
                  { value: 'left', label: 'Left (indent)' },
                  { value: 'center', label: 'Center' },
                  { value: 'right', label: 'Right (indent)' },
                  { value: 'fill', label: 'Fill' },
                  { value: 'justify', label: 'Justify' },
                  { value: 'centerContinuous', label: 'Center across selection' },
                  { value: 'distributed', label: 'Distributed' },
                ]}
              />
            </Field>
            <Field label="Vertical">
              <Select value={st.vAlign ?? 'bottom'} onChange={(v) => set({ vAlign: v === 'bottom' ? undefined : v })} options={[{ value: 'top', label: 'Top' }, { value: 'middle', label: 'Center' }, { value: 'bottom', label: 'Bottom' }]} />
            </Field>
            <Field label="Indent" inline>
              <NumberField value={st.indent ?? 0} min={0} max={15} precision={0} onChange={(v) => set({ indent: v || undefined })} />
            </Field>
            <Field label="Rotation (degrees)" inline>
              <NumberField value={st.rotation === 255 ? 90 : st.rotation ?? 0} min={-90} max={90} precision={0} onChange={(v) => set({ rotation: v || undefined })} />
            </Field>
            <Checkbox checked={st.rotation === 255} onChange={(v) => set({ rotation: v ? 255 : undefined })} label="Stack text vertically" />
            <Checkbox checked={!!st.wrap} onChange={(v) => set({ wrap: v || undefined })} label="Wrap text" />
            <Checkbox checked={!!st.shrink} onChange={(v) => set({ shrink: v || undefined })} label="Shrink to fit" />
          </div>
        )}
        {tab === 'font' && (
          <div className="fmt-grid">
            <Field label="Font">
              <FontFamilyCombo value={st.font ?? 'Calibri'} onChange={(f) => set({ font: f === 'Calibri' ? undefined : f })} width={220} />
            </Field>
            <Field label="Size">
              <FontSizeCombo value={st.size ?? 11} onChange={(s) => set({ size: s === 11 ? undefined : s })} width={80} />
            </Field>
            <Field label="Style">
              <Select
                value={st.bold && st.italic ? 'bi' : st.bold ? 'b' : st.italic ? 'i' : 'r'}
                onChange={(v) => set({ bold: v.includes('b') || undefined, italic: v.includes('i') || undefined })}
                options={[
                  { value: 'r', label: 'Regular' },
                  { value: 'i', label: 'Italic' },
                  { value: 'b', label: 'Bold' },
                  { value: 'bi', label: 'Bold italic' },
                ]}
              />
            </Field>
            <Field label="Underline">
              <Select value={st.underline ?? 'none'} onChange={(v) => set({ underline: v === 'none' ? undefined : v })} options={[{ value: 'none', label: 'None' }, { value: 'single', label: 'Single' }, { value: 'double', label: 'Double' }]} />
            </Field>
            <Field label="Colour">
              <ColorField value={st.color ?? null} onChange={(c) => set({ color: c ?? undefined })} noneLabel="Automatic" />
            </Field>
            <Checkbox checked={!!st.strike} onChange={(v) => set({ strike: v || undefined })} label="Strikethrough" />
            <div className="font-preview" style={{ fontFamily: st.font ?? 'Calibri, Carlito', fontSize: `${st.size ?? 11}pt`, fontWeight: st.bold ? 700 : 400, fontStyle: st.italic ? 'italic' : 'normal', textDecoration: `${st.underline ? 'underline' : ''} ${st.strike ? 'line-through' : ''}`, color: st.color }}>
              AaBbCcYyZz 123
            </div>
          </div>
        )}
        {tab === 'border' && (
          <div className="fmt-border">
            <div className="border-presets">
              {[
                ['none', 'None'],
                ['outside', 'Outline'],
                ['inside', 'Inside'],
                ['all', 'All'],
                ['top', 'Top'],
                ['bottom', 'Bottom'],
                ['left', 'Left'],
                ['right', 'Right'],
                ['insideH', 'Inside horizontal'],
                ['insideV', 'Inside vertical'],
                ['thickOutside', 'Thick outline'],
                ['doubleBottom', 'Double bottom'],
              ].map(([id, label]) => (
                <button key={id} className={`border-preset${borderOps.includes(id) ? ' active' : ''}`} onClick={() => setBorderOps((ops) => (id === 'none' ? ['none'] : ops.includes(id) ? ops.filter((o) => o !== id) : [...ops.filter((o) => o !== 'none'), id]))}>
                  <BorderIcon kind={id} />
                  <span>{label}</span>
                </button>
              ))}
            </div>
            <div className="fmt-grid">
              <Field label="Line style">
                <Select
                  value={borderStyle}
                  onChange={setBorderStyle}
                  options={[
                    { value: 'thin', label: 'Thin' },
                    { value: 'medium', label: 'Medium' },
                    { value: 'thick', label: 'Thick' },
                    { value: 'dashed', label: 'Dashed' },
                    { value: 'dotted', label: 'Dotted' },
                    { value: 'double', label: 'Double' },
                    { value: 'hair', label: 'Hairline' },
                    { value: 'dashDot', label: 'Dash dot' },
                    { value: 'mediumDashed', label: 'Medium dashed' },
                  ]}
                />
              </Field>
              <Field label="Colour">
                <ColorField value={borderColor} onChange={setBorderColor} noneLabel="Automatic" />
              </Field>
              <p className="muted small">Pick one or more presets; they are applied to the selected cells when you press OK.</p>
            </div>
          </div>
        )}
        {tab === 'fill' && (
          <div className="fmt-grid">
            <Field label="Background colour">
              <ColorField value={st.fill ?? null} onChange={(c) => set({ fill: c ?? undefined })} noneLabel="No fill" />
            </Field>
            <div className="fill-preview" style={{ background: st.fill ?? 'transparent' }}>
              Sample
            </div>
          </div>
        )}
        {tab === 'protection' && (
          <div className="fmt-grid">
            <Checkbox checked={st.locked !== false} onChange={(v) => set({ locked: v ? undefined : false })} label="Locked" />
            <Checkbox checked={!!st.hideFormula} onChange={(v) => set({ hideFormula: v || undefined })} label="Hidden (hide formulas)" />
            <p className="muted small">Locking cells or hiding formulas has no effect until you protect the sheet (Review ▸ Protect sheet).</p>
          </div>
        )}
      </div>
    </Modal>
  );
}

export function BorderIcon({ kind }: { kind: string }) {
  const s = 16;
  const thin = 'var(--text-3)';
  const on = 'var(--text)';
  const line = (x1: number, y1: number, x2: number, y2: number, c: string, w = 1.4, dash?: string) => <line x1={x1} y1={y1} x2={x2} y2={y2} stroke={c} strokeWidth={w} strokeDasharray={dash} />;
  const has = (side: string) => {
    switch (kind) {
      case 'all':
        return true;
      case 'outside':
      case 'thickOutside':
        return ['t', 'b', 'l', 'r'].includes(side);
      case 'inside':
        return ['h', 'v'].includes(side);
      case 'insideH':
        return side === 'h';
      case 'insideV':
        return side === 'v';
      case 'top':
        return side === 't';
      case 'bottom':
      case 'thickBottom':
      case 'doubleBottom':
        return side === 'b';
      case 'left':
        return side === 'l';
      case 'right':
        return side === 'r';
      case 'topBottom':
      case 'topThickBottom':
      case 'topDoubleBottom':
        return side === 't' || side === 'b';
    }
    return false;
  };
  const w = (side: string) => (kind === 'thickOutside' || (side === 'b' && (kind === 'thickBottom' || kind === 'topThickBottom')) ? 2.4 : 1.4);
  return (
    <svg width={s} height={s} viewBox="0 0 16 16" aria-hidden>
      {line(1, 1, 15, 1, has('t') ? on : thin, has('t') ? w('t') : 1, has('t') ? undefined : '1 1.5')}
      {line(1, 15, 15, 15, has('b') ? on : thin, has('b') ? w('b') : 1, has('b') ? undefined : '1 1.5')}
      {kind.includes('double') && line(1, 12.8, 15, 12.8, on, 1)}
      {line(1, 1, 1, 15, has('l') ? on : thin, has('l') ? w('l') : 1, has('l') ? undefined : '1 1.5')}
      {line(15, 1, 15, 15, has('r') ? on : thin, has('r') ? w('r') : 1, has('r') ? undefined : '1 1.5')}
      {line(1, 8, 15, 8, has('h') ? on : thin, has('h') ? 1.4 : 1, has('h') ? undefined : '1 1.5')}
      {line(8, 1, 8, 15, has('v') ? on : thin, has('v') ? 1.4 : 1, has('v') ? undefined : '1 1.5')}
    </svg>
  );
}

/* ============================================================ sort */

export function sortDialog(doc: SheetDoc): Promise<void> {
  return openDialog<void>((close) => <SortDialog doc={doc} close={() => close()} />).then(() => undefined);
}

function SortDialog({ doc, close }: { doc: SheetDoc; close: () => void }) {
  const rg = dataRange(doc);
  const [header, setHeader] = useState(looksLikeHeader(doc, doc.sheet, rg));
  const [levels, setLevels] = useState<SortKey[]>([{ col: doc.sel.active.c >= rg.c1 && doc.sel.active.c <= rg.c2 ? doc.sel.active.c : rg.c1, desc: false }]);
  const [caseSensitive, setCase] = useState(false);
  const colLabel = (c: number) => (header ? `${doc.displayText(doc.sheet, rg.r1, c).text || `Column ${colName(c)}`}` : `Column ${colName(c)}`);
  const cols = Array.from({ length: rg.c2 - rg.c1 + 1 }, (_, i) => rg.c1 + i);
  return (
    <Modal
      title="Sort"
      onClose={close}
      width={560}
      footer={
        <Footer
          onCancel={close}
          onOk={() => {
            sortRange(doc, rg, levels, header, { caseSensitive });
            close();
          }}
          extra={<Checkbox checked={caseSensitive} onChange={setCase} label="Case sensitive" />}
        />
      }
    >
      <p className="muted small">
        Sorting {rangeName(rg)} ({rg.r2 - rg.r1 + 1} rows).
      </p>
      <Checkbox checked={header} onChange={setHeader} label="My data has headers" />
      <div className="sort-levels">
        {levels.map((lv, i) => (
          <div key={i} className="sort-level">
            <span className="muted">{i === 0 ? 'Sort by' : 'Then by'}</span>
            <Select value={String(lv.col)} onChange={(v) => setLevels((ls) => ls.map((l, k) => (k === i ? { ...l, col: Number(v) } : l)))} options={cols.map((c) => ({ value: String(c), label: colLabel(c) }))} width={200} />
            <Select
              value={lv.by === 'fill' ? 'fill' : lv.desc ? 'desc' : 'asc'}
              onChange={(v) => setLevels((ls) => ls.map((l, k) => (k === i ? { ...l, desc: v === 'desc', by: v === 'fill' ? 'fill' : 'value', color: v === 'fill' ? doc.styleOf(doc.sheet, doc.sel.active.r, l.col).fill : undefined } : l)))}
              options={[
                { value: 'asc', label: 'A → Z / smallest first' },
                { value: 'desc', label: 'Z → A / largest first' },
                { value: 'fill', label: 'Cell colour on top' },
              ]}
              width={200}
            />
            <button className="icon-btn icon-btn-sm" aria-label="Remove level" disabled={levels.length === 1} onClick={() => setLevels((ls) => ls.filter((_, k) => k !== i))}>
              <Trash2 size={14} />
            </button>
          </div>
        ))}
      </div>
      <button className="btn btn-sm" onClick={() => setLevels((ls) => [...ls, { col: cols.find((c) => !ls.some((l) => l.col === c)) ?? rg.c1, desc: false }])}>
        <Plus size={14} /> Add level
      </button>
    </Modal>
  );
}

/* ============================================================ validation */

export function validationDialog(doc: SheetDoc): Promise<void> {
  return openDialog<void>((close) => <ValidationDialog doc={doc} close={() => close()} />).then(() => undefined);
}

const OPS: { value: CompareOp; label: string }[] = [
  { value: 'between', label: 'between' },
  { value: 'notBetween', label: 'not between' },
  { value: 'equal', label: 'equal to' },
  { value: 'notEqual', label: 'not equal to' },
  { value: 'greater', label: 'greater than' },
  { value: 'less', label: 'less than' },
  { value: 'greaterEqual', label: 'greater than or equal to' },
  { value: 'lessEqual', label: 'less than or equal to' },
];

function ValidationDialog({ doc, close }: { doc: SheetDoc; close: () => void }) {
  const { r, c } = doc.sel.active;
  const existing = doc.sheet.validations.find((v) => v.ranges.some((rg) => r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2));
  const [type, setType] = useState<ValidationType>(existing?.type ?? 'list');
  const [op, setOp] = useState<CompareOp>(existing?.op ?? 'between');
  const [v1, setV1] = useState(existing?.values[0] ?? '');
  const [v2, setV2] = useState(existing?.values[1] ?? '');
  const [blank, setBlank] = useState(existing?.allowBlank !== false);
  const [dropdown, setDropdown] = useState(existing?.showDropdown !== false);
  const [prompt, setPrompt] = useState(existing?.prompt ?? '');
  const [error, setError] = useState(existing?.error ?? '');
  const [style, setStyle] = useState<Validation['errorStyle']>(existing?.errorStyle ?? 'stop');
  const needsOp = ['whole', 'decimal', 'date', 'time', 'textLength'].includes(type);
  const two = needsOp && (op === 'between' || op === 'notBetween');
  const ok = () => {
    if (type === 'any') setValidation(doc, null);
    else
      setValidation(doc, {
        id: existing?.id ?? newId('dv'),
        ranges: [],
        type,
        op: needsOp ? op : undefined,
        values: type === 'list' ? (v1.startsWith('=') ? [v1] : v1.split(/[,;]/).map((x) => x.trim()).filter(Boolean)) : type === 'checkbox' ? [] : two ? [v1, v2] : [v1],
        allowBlank: blank,
        showDropdown: dropdown,
        prompt: prompt || undefined,
        error: error || undefined,
        errorStyle: style,
      });
    close();
  };
  return (
    <Modal
      title="Data validation"
      onClose={close}
      width={520}
      footer={
        <Footer
          onCancel={close}
          onOk={ok}
          extra={
            <button
              className="btn btn-ghost"
              onClick={() => {
                setValidation(doc, null);
                close();
              }}
            >
              Clear all
            </button>
          }
        />
      }
    >
      <Field label="Allow">
        <Select
          value={type}
          onChange={setType}
          options={[
            { value: 'any', label: 'Any value' },
            { value: 'list', label: 'List (dropdown)' },
            { value: 'whole', label: 'Whole number' },
            { value: 'decimal', label: 'Decimal' },
            { value: 'date', label: 'Date' },
            { value: 'time', label: 'Time' },
            { value: 'textLength', label: 'Text length' },
            { value: 'checkbox', label: 'Checkbox (TRUE/FALSE)' },
            { value: 'custom', label: 'Custom formula' },
          ]}
        />
      </Field>
      {needsOp && (
        <Field label="Data">
          <Select value={op} onChange={setOp} options={OPS} />
        </Field>
      )}
      {type === 'list' && (
        <Field label="Source" hint="Items separated by commas (Yes, No, Maybe) or a range like =$A$1:$A$10">
          <input className="input" value={v1} onChange={(e) => setV1(e.target.value)} placeholder="Yes, No, Maybe" />
        </Field>
      )}
      {(needsOp || type === 'custom') && (
        <div className="row gap">
          <Field label={type === 'custom' ? 'Formula' : two ? 'Minimum' : 'Value'}>
            <input className="input" value={v1} onChange={(e) => setV1(e.target.value)} placeholder={type === 'custom' ? '=A1>0' : type === 'date' ? '1/1/2024' : '0'} />
          </Field>
          {two && (
            <Field label="Maximum">
              <input className="input" value={v2} onChange={(e) => setV2(e.target.value)} placeholder="100" />
            </Field>
          )}
        </div>
      )}
      <div className="row gap wrap">
        <Checkbox checked={blank} onChange={setBlank} label="Ignore blank" />
        {type === 'list' && <Checkbox checked={dropdown} onChange={setDropdown} label="Show dropdown arrow" />}
      </div>
      <Field label="Input message (optional)">
        <input className="input" value={prompt} onChange={(e) => setPrompt(e.target.value)} placeholder="Shown when the cell is selected" />
      </Field>
      <div className="row gap">
        <Field label="When invalid">
          <Select value={style ?? 'stop'} onChange={(v) => setStyle(v)} options={[{ value: 'stop', label: 'Stop (reject)' }, { value: 'warning', label: 'Warning' }, { value: 'information', label: 'Information' }]} />
        </Field>
        <Field label="Error message">
          <input className="input" value={error} onChange={(e) => setError(e.target.value)} placeholder="Please pick a value from the list." />
        </Field>
      </div>
    </Modal>
  );
}

/* ============================================================ insert function */

const CATS: (Category | 'All' | 'Popular')[] = ['Popular', 'All', 'Math', 'Statistical', 'Logical', 'Text', 'Lookup', 'Date & time', 'Financial', 'Information', 'Engineering', 'Database', 'Web'];
const POPULAR = ['SUM', 'AVERAGE', 'IF', 'COUNT', 'COUNTIF', 'SUMIF', 'VLOOKUP', 'XLOOKUP', 'INDEX', 'MATCH', 'IFERROR', 'MAX', 'MIN', 'ROUND', 'CONCAT', 'TEXT', 'TODAY', 'FILTER', 'SORT', 'UNIQUE'];

export function insertFunctionDialog(initialCat?: string): Promise<string | undefined> {
  return openDialog<string>((close) => <InsertFunction close={close} initialCat={initialCat} />);
}

function InsertFunction({ close, initialCat }: { close: (v?: string) => void; initialCat?: string }) {
  const [q, setQ] = useState('');
  const [cat, setCat] = useState<string>(initialCat ?? 'Popular');
  const all = useMemo(() => allFunctions(), []);
  const list = useMemo(() => {
    const query = q.trim().toLowerCase();
    let fns = all;
    if (query) fns = fns.filter((f) => f.name.toLowerCase().includes(query) || f.desc.toLowerCase().includes(query));
    else if (cat === 'Popular') fns = POPULAR.map((n) => all.find((f) => f.name === n)!).filter(Boolean);
    else if (cat !== 'All') fns = fns.filter((f) => f.category === cat);
    return fns;
  }, [q, cat, all]);
  const [sel, setSel] = useState<string | null>(null);
  const current = list.find((f) => f.name === sel) ?? list[0];
  return (
    <Modal title="Insert function" onClose={() => close()} width={620} footer={<Footer onCancel={() => close()} onOk={() => current && close(current.name)} okLabel="Insert" disabled={!current} />}>
      <div className="fn-search">
        <Search size={15} />
        <input className="input" autoFocus placeholder="Search 480+ functions (e.g. “average”, “lookup”, “date”)" value={q} onChange={(e) => setQ(e.target.value)} />
      </div>
      <div className="fn-browser">
        <div className="fn-cats">
          {CATS.map((c) => (
            <button key={c} className={`fmt-cat${cat === c && !q ? ' active' : ''}`} onClick={() => (setCat(c), setQ(''))}>
              {c}
            </button>
          ))}
        </div>
        <div className="fn-list" role="listbox">
          {list.map((f) => (
            <button key={f.name} className={`fn-item${current?.name === f.name ? ' active' : ''}`} onClick={() => setSel(f.name)} onDoubleClick={() => close(f.name)}>
              {f.name}
            </button>
          ))}
          {!list.length && <p className="muted small">No functions match.</p>}
        </div>
      </div>
      {current && (
        <div className="fn-detail">
          <div className="mono">
            <b>{current.name}</b>({current.syntax})
          </div>
          <p>{current.desc}</p>
          <span className="chip">{current.category}</span>
        </div>
      )}
    </Modal>
  );
}

/* ============================================================ conditional formatting */

const CF_TYPES: { value: CondType; label: string }[] = [
  { value: 'cell', label: 'Cell value' },
  { value: 'text', label: 'Text that contains' },
  { value: 'begins', label: 'Text that begins with' },
  { value: 'ends', label: 'Text that ends with' },
  { value: 'notText', label: 'Text that does not contain' },
  { value: 'date', label: 'A date occurring' },
  { value: 'top', label: 'Top / bottom ranked values' },
  { value: 'average', label: 'Above / below average' },
  { value: 'duplicate', label: 'Duplicate values' },
  { value: 'unique', label: 'Unique values' },
  { value: 'blank', label: 'Blank cells' },
  { value: 'notBlank', label: 'Cells that are not blank' },
  { value: 'error', label: 'Errors' },
  { value: 'formula', label: 'Use a formula' },
  { value: 'colorScale', label: 'Colour scale' },
  { value: 'dataBar', label: 'Data bars' },
  { value: 'iconSet', label: 'Icon set' },
];

export const CF_PRESET_STYLES: { label: string; style: NonNullable<CondFormat['style']> }[] = [
  { label: 'Light red fill with dark red text', style: { fill: '#ffc7ce', color: '#9c0006' } },
  { label: 'Yellow fill with dark yellow text', style: { fill: '#ffeb9c', color: '#9c5700' } },
  { label: 'Green fill with dark green text', style: { fill: '#c6efce', color: '#006100' } },
  { label: 'Light red fill', style: { fill: '#ffc7ce' } },
  { label: 'Red text', style: { color: '#c00000' } },
  { label: 'Bold', style: { bold: true } },
  { label: 'Blue fill with white text', style: { fill: '#2f6dff', color: '#ffffff' } },
];

export function conditionalRuleDialog(doc: SheetDoc, rule?: CondFormat, preset?: Partial<CondFormat>): Promise<void> {
  return openDialog<void>((close) => <CfRuleDialog doc={doc} rule={rule} preset={preset} close={() => close()} />).then(() => undefined);
}

function CfRuleDialog({ doc, rule, preset, close }: { doc: SheetDoc; rule?: CondFormat; preset?: Partial<CondFormat>; close: () => void }) {
  const initial: CondFormat = rule ?? { id: newId('cf'), ranges: doc.sel.ranges.map((r) => doc.clampToUsed(doc.sheet, r, 200)), type: 'cell', op: 'greater', values: [''], style: CF_PRESET_STYLES[0].style, ...preset };
  const [r, setR] = useState<CondFormat>(initial);
  const [rangeText, setRangeText] = useState(initial.ranges.map(rangeName).join(', '));
  const set = (p: Partial<CondFormat>) => setR((x) => ({ ...x, ...p }));
  const ranges = rangeText
    .split(/[,;]\s*/)
    .map((t) => parseRange(t))
    .filter((x): x is Range => !!x);
  const presetIdx = CF_PRESET_STYLES.findIndex((p) => JSON.stringify(p.style) === JSON.stringify(r.style));
  const visual = r.type === 'colorScale' || r.type === 'dataBar' || r.type === 'iconSet';
  return (
    <Modal
      title={rule ? 'Edit formatting rule' : 'New formatting rule'}
      onClose={close}
      width={560}
      footer={
        <Footer
          onCancel={close}
          disabled={!ranges.length}
          onOk={() => {
            upsertCF(doc, { ...r, ranges });
            close();
          }}
        />
      }
    >
      <Field label="Apply to">
        <input className="input mono" value={rangeText} onChange={(e) => setRangeText(e.target.value)} />
      </Field>
      <Field label="Format cells where">
        <Select value={r.type} onChange={(t) => set({ type: t, values: r.values ?? [''], colors: t === 'colorScale' ? ['#f8696b', '#ffeb84', '#63be7b'] : t === 'dataBar' ? ['#638ec6'] : r.colors })} options={CF_TYPES} />
      </Field>
      {r.type === 'cell' && (
        <div className="row gap">
          <Select value={r.op ?? 'greater'} onChange={(op) => set({ op })} options={OPS} width={200} />
          <input className="input" value={r.values?.[0] ?? ''} placeholder="Value or =formula" onChange={(e) => set({ values: [e.target.value, r.values?.[1] ?? ''] })} />
          {(r.op === 'between' || r.op === 'notBetween') && <input className="input" value={r.values?.[1] ?? ''} placeholder="and" onChange={(e) => set({ values: [r.values?.[0] ?? '', e.target.value] })} />}
        </div>
      )}
      {(r.type === 'text' || r.type === 'begins' || r.type === 'ends' || r.type === 'notText') && <input className="input" value={r.text ?? ''} placeholder="Text" onChange={(e) => set({ text: e.target.value })} />}
      {r.type === 'date' && (
        <Select
          value={r.datePeriod ?? 'today'}
          onChange={(d) => set({ datePeriod: d })}
          options={[
            { value: 'yesterday', label: 'Yesterday' },
            { value: 'today', label: 'Today' },
            { value: 'tomorrow', label: 'Tomorrow' },
            { value: 'last7Days', label: 'In the last 7 days' },
            { value: 'lastWeek', label: 'Last week' },
            { value: 'thisWeek', label: 'This week' },
            { value: 'nextWeek', label: 'Next week' },
            { value: 'lastMonth', label: 'Last month' },
            { value: 'thisMonth', label: 'This month' },
            { value: 'nextMonth', label: 'Next month' },
          ]}
        />
      )}
      {r.type === 'top' && (
        <div className="row gap">
          <Segmented value={r.bottom ? 'bottom' : 'top'} onChange={(v) => set({ bottom: v === 'bottom' })} options={[{ value: 'top', label: 'Top' }, { value: 'bottom', label: 'Bottom' }]} size="sm" />
          <NumberField value={r.rank ?? 10} min={1} max={1000} precision={0} onChange={(v) => set({ rank: v })} />
          <Checkbox checked={!!r.percent} onChange={(v) => set({ percent: v })} label="% of the range" />
        </div>
      )}
      {r.type === 'average' && <Segmented value={r.above === false ? 'below' : 'above'} onChange={(v) => set({ above: v === 'above' })} options={[{ value: 'above', label: 'Above average' }, { value: 'below', label: 'Below average' }]} size="sm" />}
      {r.type === 'formula' && (
        <Field label="Formula (relative to the first cell of the range)" hint="Example: =$C2>100 formats whole rows where column C is over 100.">
          <input className="input mono" value={r.values?.[0] ?? '='} onChange={(e) => set({ values: [e.target.value] })} />
        </Field>
      )}
      {r.type === 'colorScale' && (
        <div className="row gap">
          {(r.colors ?? []).map((c, i) => (
            <ColorField key={i} value={c} onChange={(v) => set({ colors: (r.colors ?? []).map((x, k) => (k === i ? v ?? x : x)) })} />
          ))}
          <Segmented value={String(r.colors?.length ?? 3)} onChange={(v) => set({ colors: v === '2' ? [r.colors?.[0] ?? '#ffffff', r.colors?.[r.colors.length - 1] ?? '#63be7b'] : ['#f8696b', '#ffeb84', '#63be7b'] })} options={[{ value: '2', label: '2 colours' }, { value: '3', label: '3 colours' }]} size="sm" />
        </div>
      )}
      {r.type === 'dataBar' && (
        <div className="row gap">
          <ColorField value={r.colors?.[0] ?? '#638ec6'} onChange={(v) => set({ colors: [v ?? '#638ec6'] })} />
          <Checkbox checked={r.showValue !== false} onChange={(v) => set({ showValue: v })} label="Show the value" />
        </div>
      )}
      {r.type === 'iconSet' && (
        <div className="row gap">
          <Select
            value={r.iconSet ?? '3Arrows'}
            onChange={(v) => set({ iconSet: v })}
            options={[
              { value: '3Arrows', label: '3 arrows' },
              { value: '4Arrows', label: '4 arrows' },
              { value: '5Arrows', label: '5 arrows' },
              { value: '3TrafficLights', label: 'Traffic lights' },
              { value: '3Symbols', label: 'Symbols (✓ ! ✗)' },
              { value: '3Stars', label: 'Stars' },
              { value: '3Flags', label: 'Flags' },
              { value: '5Ratings', label: 'Ratings' },
            ]}
          />
          <Checkbox checked={r.showValue !== false} onChange={(v) => set({ showValue: v })} label="Show the value" />
        </div>
      )}
      {!visual && (
        <Field label="Format with">
          <Select
            value={presetIdx >= 0 ? String(presetIdx) : 'custom'}
            onChange={(v) => {
              if (v !== 'custom') set({ style: CF_PRESET_STYLES[Number(v)].style });
            }}
            options={[...CF_PRESET_STYLES.map((p, i) => ({ value: String(i), label: p.label })), { value: 'custom', label: 'Custom format…' }]}
          />
          <div className="row gap cf-custom">
            <ColorField value={r.style?.fill ?? null} onChange={(c) => set({ style: { ...r.style, fill: c ?? undefined } })} noneLabel="No fill" />
            <ColorField value={r.style?.color ?? null} onChange={(c) => set({ style: { ...r.style, color: c ?? undefined } })} noneLabel="Text colour" />
            <Checkbox checked={!!r.style?.bold} onChange={(v) => set({ style: { ...r.style, bold: v || undefined } })} label="Bold" />
            <Checkbox checked={!!r.style?.italic} onChange={(v) => set({ style: { ...r.style, italic: v || undefined } })} label="Italic" />
          </div>
          <div className="cf-preview" style={{ background: r.style?.fill, color: r.style?.color, fontWeight: r.style?.bold ? 700 : 400, fontStyle: r.style?.italic ? 'italic' : 'normal' }}>
            AaBbCc 123
          </div>
        </Field>
      )}
      <Checkbox checked={!!r.stop} onChange={(v) => set({ stop: v })} label="Stop if true (later rules are skipped)" />
    </Modal>
  );
}

export function manageRulesDialog(doc: SheetDoc): Promise<void> {
  return openDialog<void>((close) => <ManageRules doc={doc} close={() => close()} />).then(() => undefined);
}

function ManageRules({ doc, close }: { doc: SheetDoc; close: () => void }) {
  const [, force] = useState(0);
  const rules = doc.sheet.cf;
  const describe = (r: CondFormat) => {
    const t = CF_TYPES.find((x) => x.value === r.type)?.label ?? r.type;
    if (r.type === 'cell') return `${t} ${OPS.find((o) => o.value === r.op)?.label ?? ''} ${r.values?.filter(Boolean).join(' and ') ?? ''}`;
    if (r.text) return `${t} “${r.text}”`;
    if (r.type === 'top') return `${r.bottom ? 'Bottom' : 'Top'} ${r.rank ?? 10}${r.percent ? '%' : ''}`;
    if (r.type === 'formula') return `Formula: ${r.values?.[0] ?? ''}`;
    return t;
  };
  return (
    <Modal
      title="Conditional formatting rules"
      onClose={close}
      width={640}
      footer={
        <>
          <button
            className="btn"
            onClick={async () => {
              await conditionalRuleDialog(doc);
              force((x) => x + 1);
            }}
          >
            <Plus size={14} /> New rule
          </button>
          <span className="spacer" />
          <button className="btn btn-primary" onClick={close}>
            Done
          </button>
        </>
      }
    >
      {!rules.length && <p className="muted">This sheet has no conditional formatting rules yet.</p>}
      <div className="cf-rules">
        {rules.map((r, i) => (
          <div key={r.id} className="cf-rule">
            <span className="cf-swatch" style={{ background: r.style?.fill ?? (r.colors ? `linear-gradient(90deg, ${r.colors.join(',')})` : undefined), color: r.style?.color }}>
              Aa
            </span>
            <span className="cf-desc">{describe(r)}</span>
            <span className="mono muted">{r.ranges.map(rangeName).join(', ')}</span>
            <button className="icon-btn icon-btn-sm" aria-label="Move up" disabled={i === 0} onClick={() => (reorderCF(doc, r.id, -1), force((x) => x + 1))}>
              <ArrowUp size={14} />
            </button>
            <button className="icon-btn icon-btn-sm" aria-label="Move down" disabled={i === rules.length - 1} onClick={() => (reorderCF(doc, r.id, 1), force((x) => x + 1))}>
              <ArrowDown size={14} />
            </button>
            <button
              className="btn btn-sm"
              onClick={async () => {
                await conditionalRuleDialog(doc, r);
                force((x) => x + 1);
              }}
            >
              Edit
            </button>
            <button className="icon-btn icon-btn-sm" aria-label="Delete rule" onClick={() => (deleteCF(doc, [r.id]), force((x) => x + 1))}>
              <Trash2 size={14} />
            </button>
          </div>
        ))}
      </div>
    </Modal>
  );
}

/* ============================================================ names */

export function nameManagerDialog(doc: SheetDoc): Promise<void> {
  return openDialog<void>((close) => <NameManager doc={doc} close={() => close()} />).then(() => undefined);
}

function NameManager({ doc, close }: { doc: SheetDoc; close: () => void }) {
  const [names, setList] = useState<DefinedName[]>(structuredClone(doc.wb.names));
  const sel = doc.sel.ranges[doc.sel.ranges.length - 1];
  const defaultRef = `'${doc.sheet.name.replace(/'/g, "''")}'!$${colName(sel.c1)}$${sel.r1 + 1}${sel.r1 !== sel.r2 || sel.c1 !== sel.c2 ? `:$${colName(sel.c2)}$${sel.r2 + 1}` : ''}`;
  const valid = (n: string) => /^[A-Za-z_\\][\w.]*$/.test(n) && !/^[A-Za-z]{1,3}\d+$/.test(n) && !/^(R|C)$/i.test(n);
  const dup = (n: string, i: number) => names.some((x, k) => k !== i && x.name.toLowerCase() === n.toLowerCase());
  const bad = names.some((n, i) => !valid(n.name) || dup(n.name, i));
  return (
    <Modal
      title="Name manager"
      onClose={close}
      width={680}
      footer={
        <Footer
          onCancel={close}
          disabled={bad}
          onOk={() => {
            setNames(doc, names.map((n) => ({ ...n, ref: n.ref.replace(/^=/, '') })));
            close();
          }}
          extra={
            <button className="btn" onClick={() => setList((l) => [...l, { name: `Name${l.length + 1}`, ref: defaultRef }])}>
              <Plus size={14} /> New
            </button>
          }
        />
      }
    >
      <p className="muted small">Names make formulas easier to read: =SUM(Sales) instead of =SUM(Data!$B$2:$B$200). A name can also hold a constant (=0.2) or a LAMBDA to create your own function.</p>
      <div className="name-list">
        {names.map((n, i) => {
          const value = doc.engine.evaluateText(n.ref, doc.sheet.id, 0, 0);
          const preview = isMatrix(value) ? `{${value.data.length}×${value.data[0]?.length ?? 0} values}` : value === null ? '' : typeof value === 'object' && 'error' in value ? (value as { error: string }).error : String(value);
          return (
            <div key={i} className="name-row">
              <input className={`input${!valid(n.name) || dup(n.name, i) ? ' invalid' : ''}`} value={n.name} onChange={(e) => setList((l) => l.map((x, k) => (k === i ? { ...x, name: e.target.value } : x)))} />
              <input className="input mono" value={n.ref.startsWith('=') ? n.ref : '=' + n.ref} onChange={(e) => setList((l) => l.map((x, k) => (k === i ? { ...x, ref: e.target.value.replace(/^=/, '') } : x)))} />
              <Select value={n.sheet === undefined ? 'wb' : String(n.sheet)} onChange={(v) => setList((l) => l.map((x, k) => (k === i ? { ...x, sheet: v === 'wb' ? undefined : Number(v) } : x)))} options={[{ value: 'wb', label: 'Workbook' }, ...doc.wb.sheets.map((s) => ({ value: String(s.id), label: s.name }))]} width={130} />
              <span className="name-value muted" title={preview}>
                {preview}
              </span>
              <button className="icon-btn icon-btn-sm" aria-label="Delete name" onClick={() => setList((l) => l.filter((_, k) => k !== i))}>
                <Trash2 size={14} />
              </button>
            </div>
          );
        })}
        {!names.length && <p className="muted">No names defined yet.</p>}
      </div>
    </Modal>
  );
}

/* ============================================================ charts */

export const CHART_TYPES: { value: ChartType; label: string }[] = [
  { value: 'column', label: 'Column' },
  { value: 'bar', label: 'Bar' },
  { value: 'line', label: 'Line' },
  { value: 'area', label: 'Area' },
  { value: 'pie', label: 'Pie' },
  { value: 'doughnut', label: 'Doughnut' },
  { value: 'scatter', label: 'Scatter (X Y)' },
  { value: 'radar', label: 'Radar' },
  { value: 'polarArea', label: 'Polar area' },
  { value: 'bubble', label: 'Bubble' },
  { value: 'combo', label: 'Combo (column + line)' },
];

export function newChartSpec(doc: SheetDoc, type: ChartType): ChartSpec {
  const rg = dataRange(doc);
  const src = `'${doc.sheet.name.replace(/'/g, "''")}'!${addrName(rg.r1, rg.c1, true, true)}:${addrName(rg.r2, rg.c2, true, true)}`;
  const title = looksLikeHeader(doc, doc.sheet, rg) && rg.c2 - rg.c1 === 1 ? doc.displayText(doc.sheet, rg.r1, rg.c2).text : '';
  return { id: newId('chart'), type, source: src, title, anchor: { r: rg.r1, c: rg.c2 + 2, dx: 0, dy: 0 }, w: 480, h: 290, palette: 'affice' };
}

export function chartDialog(doc: SheetDoc, spec: ChartSpec): Promise<void> {
  return openDialog<void>((close) => <ChartEditor doc={doc} initial={spec} close={() => close()} />).then(() => undefined);
}

function ChartEditor({ doc, initial, close }: { doc: SheetDoc; initial: ChartSpec; close: () => void }) {
  const [s, setS] = useState<ChartSpec>(initial);
  const set = (p: Partial<ChartSpec>) => setS((x) => ({ ...x, ...p }));
  return (
    <Modal
      title="Chart"
      onClose={close}
      width={560}
      footer={
        <Footer
          onCancel={close}
          onOk={() => {
            upsertChart(doc, s);
            close();
          }}
        />
      }
    >
      <div className="chart-types">
        {CHART_TYPES.map((t) => (
          <button key={t.value} className={`chart-type${s.type === t.value ? ' active' : ''}`} onClick={() => set({ type: t.value })}>
            {t.label}
          </button>
        ))}
      </div>
      <Field label="Data range">
        <input className="input mono" value={s.source} onChange={(e) => set({ source: e.target.value })} />
      </Field>
      <div className="row gap">
        <Field label="Title">
          <input className="input" value={s.title ?? ''} onChange={(e) => set({ title: e.target.value })} />
        </Field>
        <Field label="Legend">
          <Select value={s.legend ?? 'bottom'} onChange={(v) => set({ legend: v })} options={[{ value: 'bottom', label: 'Bottom' }, { value: 'top', label: 'Top' }, { value: 'right', label: 'Right' }, { value: 'left', label: 'Left' }, { value: 'none', label: 'None' }]} />
        </Field>
      </div>
      <div className="row gap">
        <Field label="Horizontal axis title">
          <input className="input" value={s.xTitle ?? ''} onChange={(e) => set({ xTitle: e.target.value || undefined })} />
        </Field>
        <Field label="Vertical axis title">
          <input className="input" value={s.yTitle ?? ''} onChange={(e) => set({ yTitle: e.target.value || undefined })} />
        </Field>
      </div>
      <Field label="Colours">
        <div className="palette-row">
          {Object.entries(PALETTES).map(([id, cols]) => (
            <button key={id} className={`palette${(s.palette ?? 'affice') === id ? ' active' : ''}`} onClick={() => set({ palette: id })} aria-label={id}>
              {cols.slice(0, 5).map((c) => (
                <span key={c} style={{ background: c }} />
              ))}
            </button>
          ))}
        </div>
      </Field>
      <div className="row gap wrap">
        <Select value={String(s.stacked ?? false)} onChange={(v) => set({ stacked: v === 'false' ? false : v === 'percent' ? 'percent' : true })} options={[{ value: 'false', label: 'Not stacked' }, { value: 'true', label: 'Stacked' }, { value: 'percent', label: '100% stacked' }]} width={150} />
        <Select value={s.seriesInRows === undefined ? 'auto' : s.seriesInRows ? 'rows' : 'cols'} onChange={(v) => set({ seriesInRows: v === 'auto' ? undefined : v === 'rows' })} options={[{ value: 'auto', label: 'Series: automatic' }, { value: 'cols', label: 'Series in columns' }, { value: 'rows', label: 'Series in rows' }]} width={170} />
        <Checkbox checked={!!s.smooth} onChange={(v) => set({ smooth: v })} label="Smooth lines" />
        <Checkbox checked={!!s.dataLabels} onChange={(v) => set({ dataLabels: v })} label="Data labels" />
        <Checkbox checked={s.gridlines !== false} onChange={(v) => set({ gridlines: v })} label="Gridlines" />
      </div>
    </Modal>
  );
}

/* ============================================================ data tools */

export function removeDuplicatesDialog(doc: SheetDoc): Promise<string | undefined> {
  return openDialog<string>((close) => <RemoveDup doc={doc} close={close} />);
}

function RemoveDup({ doc, close }: { doc: SheetDoc; close: (msg?: string) => void }) {
  const rg = dataRange(doc);
  const [header, setHeader] = useState(looksLikeHeader(doc, doc.sheet, rg));
  const cols = Array.from({ length: rg.c2 - rg.c1 + 1 }, (_, i) => rg.c1 + i);
  const [picked, setPicked] = useState<number[]>(cols);
  return (
    <Modal
      title="Remove duplicates"
      onClose={() => close()}
      width={460}
      footer={
        <Footer
          onCancel={() => close()}
          disabled={!picked.length}
          onOk={() => {
            const res = removeDuplicates(doc, rg, picked, header);
            close(res.removed ? `Removed ${res.removed} duplicate ${res.removed === 1 ? 'row' : 'rows'}; ${res.kept} unique rows remain.` : 'No duplicate rows were found.');
          }}
        />
      }
    >
      <Checkbox checked={header} onChange={setHeader} label="My data has headers" />
      <p className="muted small">Rows count as duplicates when all the ticked columns match.</p>
      <div className="check-list">
        {cols.map((c) => (
          <Checkbox key={c} checked={picked.includes(c)} onChange={(v) => setPicked((p) => (v ? [...p, c] : p.filter((x) => x !== c)))} label={header ? doc.displayText(doc.sheet, rg.r1, c).text || `Column ${colName(c)}` : `Column ${colName(c)}`} />
        ))}
      </div>
    </Modal>
  );
}

export function textToColumnsDialog(doc: SheetDoc): Promise<void> {
  return openDialog<void>((close) => <TextToCols doc={doc} close={() => close()} />).then(() => undefined);
}

function TextToCols({ doc, close }: { doc: SheetDoc; close: () => void }) {
  const sel = doc.sel.ranges[doc.sel.ranges.length - 1];
  const rg = sel.r1 === sel.r2 && sel.c1 === sel.c2 ? currentRegion(doc, doc.sheet, sel.r1, sel.c1) : doc.clampToUsed(doc.sheet, sel);
  const col: Range = { ...rg, c2: rg.c1 };
  const [delims, setDelims] = useState<string[]>(['comma']);
  const [other, setOther] = useState('');
  const [consec, setConsec] = useState(false);
  const map: Record<string, string> = { comma: ',', tab: 'tab', semicolon: ';', space: 'space' };
  const sample = Array.from({ length: Math.min(5, col.r2 - col.r1 + 1) }, (_, i) => doc.displayText(doc.sheet, col.r1 + i, col.c1).text);
  return (
    <Modal
      title="Text to columns"
      onClose={close}
      width={520}
      footer={
        <Footer
          onCancel={close}
          onOk={() => {
            textToColumns(doc, col, { delimiters: delims.map((d) => map[d]), other: other || undefined, treatConsecutive: consec });
            close();
          }}
        />
      }
    >
      <p className="muted small">Splits column {colName(col.c1)} into several columns.</p>
      <div className="row gap wrap">
        {Object.keys(map).map((d) => (
          <Checkbox key={d} checked={delims.includes(d)} onChange={(v) => setDelims((x) => (v ? [...x, d] : x.filter((y) => y !== d)))} label={d[0].toUpperCase() + d.slice(1)} />
        ))}
        <input className="input" style={{ width: 80 }} placeholder="Other" value={other} maxLength={3} onChange={(e) => setOther(e.target.value)} />
      </div>
      <Checkbox checked={consec} onChange={setConsec} label="Treat consecutive delimiters as one" />
      <div className="t2c-preview mono">
        {sample.map((s, i) => (
          <div key={i}>{s || ' '}</div>
        ))}
      </div>
    </Modal>
  );
}

export function goalSeekDialog(doc: SheetDoc): Promise<string | undefined> {
  return openDialog<string>((close) => <GoalSeek doc={doc} close={close} />);
}

function GoalSeek({ doc, close }: { doc: SheetDoc; close: (msg?: string) => void }) {
  const a = doc.sel.active;
  const [setCell, setSetCell] = useState(addrName(a.r, a.c));
  const [target, setTarget] = useState('');
  const [changing, setChanging] = useState('');
  const run = () => {
    const sc = parseRange(setCell);
    const cc = parseRange(changing);
    const goal = Number(target);
    if (!sc || !cc || Number.isNaN(goal)) return close('Goal Seek needs a formula cell, a target number and an input cell.');
    const sheet = doc.sheet;
    const inputKey = { r: cc.r1, c: cc.c1 };
    const original = sheet.get(inputKey.r, inputKey.c);
    if (original?.f !== undefined) return close('The changing cell must contain a value, not a formula.');
    const f = (x: number) => {
      sheet.put(inputKey.r * 16384 + inputKey.c, { ...(sheet.get(inputKey.r, inputKey.c) ?? {}), v: x });
      doc.engine.cellsChanged([{ sheet: sheet.id, key: inputKey.r * 16384 + inputKey.c }]);
      const v = doc.value(sheet, sc.r1, sc.c1);
      return typeof v === 'number' ? v - goal : NaN;
    };
    let x = typeof original?.v === 'number' ? original.v : 0;
    let found: number | null = null;
    for (let i = 0; i < 100; i++) {
      const y = f(x);
      if (Number.isNaN(y)) break;
      if (Math.abs(y) < 1e-9) {
        found = x;
        break;
      }
      const h = Math.max(1e-6, Math.abs(x) * 1e-6);
      const d = (f(x + h) - y) / h;
      if (!d || !Number.isFinite(d)) break;
      x -= y / d;
    }
    // restore, then apply as one undoable edit
    sheet.put(inputKey.r * 16384 + inputKey.c, original ? { ...original } : undefined);
    doc.engine.cellsChanged([{ sheet: sheet.id, key: inputKey.r * 16384 + inputKey.c }]);
    if (found === null) return close('Goal Seek couldn’t find a solution.');
    const val = parseFloat(found.toPrecision(12));
    doc.transact('Goal seek', (tx) => tx.patch(sheet, inputKey.r * 16384 + inputKey.c, { v: val }));
    close(`Found a solution: ${changing} = ${val}`);
  };
  return (
    <Modal title="Goal seek" onClose={() => close()} width={420} footer={<Footer onCancel={() => close()} onOk={run} />}>
      <p className="muted small">Finds the input value that makes a formula reach the result you want.</p>
      <Field label="Set cell (formula)">
        <input className="input mono" value={setCell} onChange={(e) => setSetCell(e.target.value)} />
      </Field>
      <Field label="To value">
        <input className="input" value={target} onChange={(e) => setTarget(e.target.value)} placeholder="e.g. 1000" />
      </Field>
      <Field label="By changing cell">
        <input className="input mono" value={changing} onChange={(e) => setChanging(e.target.value)} placeholder="e.g. B3" />
      </Field>
    </Modal>
  );
}

/* ============================================================ misc prompts */

export function pasteSpecialDialog(): Promise<string | undefined> {
  const opts: [string, string][] = [
    ['all', 'All'],
    ['formulas', 'Formulas'],
    ['values', 'Values'],
    ['formats', 'Formats'],
    ['valuesAndFormats', 'Values and number formats'],
    ['noBorders', 'All except borders'],
    ['colWidths', 'Column widths'],
    ['transpose', 'Transpose'],
    ['link', 'Paste link'],
  ];
  return openDialog<string>((close) => {
    let pick = 'values';
    return (
      <Modal title="Paste special" onClose={() => close()} width={380} footer={<Footer onCancel={() => close()} onOk={() => close(pick)} />}>
        <div className="radio-list">
          {opts.map(([v, l]) => (
            <label key={v} className="radio">
              <input type="radio" name="ps" defaultChecked={v === 'values'} onChange={() => (pick = v)} /> {l}
            </label>
          ))}
        </div>
      </Modal>
    );
  });
}

export function shiftCellsDialog(kind: 'insert' | 'delete'): Promise<string | undefined> {
  const opts: [string, string][] =
    kind === 'insert'
      ? [
          ['right', 'Shift cells right'],
          ['down', 'Shift cells down'],
          ['row', 'Entire row'],
          ['col', 'Entire column'],
        ]
      : [
          ['left', 'Shift cells left'],
          ['up', 'Shift cells up'],
          ['row', 'Entire row'],
          ['col', 'Entire column'],
        ];
  return openDialog<string>((close) => {
    let pick = opts[1][0];
    return (
      <Modal title={kind === 'insert' ? 'Insert' : 'Delete'} onClose={() => close()} width={320} footer={<Footer onCancel={() => close()} onOk={() => close(pick)} />}>
        <div className="radio-list">
          {opts.map(([v, l], i) => (
            <label key={v} className="radio">
              <input type="radio" name="sc" defaultChecked={i === 1} onChange={() => (pick = v)} /> {l}
            </label>
          ))}
        </div>
      </Modal>
    );
  });
}

export function protectSheetDialog(doc: SheetDoc, index: number): Promise<string | undefined> {
  const sheet = doc.wb.sheets[index];
  const on = !!sheet?.protection?.enabled;
  return openDialog<string>((close) => {
    let pw = '';
    return (
      <Modal
        title={on ? 'Unprotect sheet' : 'Protect sheet'}
        onClose={() => close()}
        width={400}
        footer={
          <Footer
            onCancel={() => close()}
            onOk={() => {
              if (on && sheet.protection?.password && sheet.protection.password !== hashPw(pw)) return close('The password is not correct.');
              doc.transact(on ? 'Unprotect sheet' : 'Protect sheet', (tx) => {
                tx.prop(sheet, 'protection');
                sheet.protection = on ? undefined : { enabled: true, password: pw ? hashPw(pw) : undefined };
              });
              close(on ? 'Sheet unprotected.' : 'Sheet protected. Only unlocked cells can be edited.');
            }}
          />
        }
      >
        <p className="muted small">{on ? 'Enter the password to unprotect this sheet (leave empty if none was set).' : 'Protected sheets stop changes to locked cells. Use Format cells ▸ Protection to unlock cells people may edit.'}</p>
        <Field label="Password (optional)">
          <input className="input" type="password" onChange={(e) => (pw = e.target.value)} />
        </Field>
      </Modal>
    );
  });
}

export function hashPw(pw: string): string {
  // Excel-compatible 16-bit legacy password hash
  let h = 0;
  for (let i = pw.length - 1; i >= 0; i--) {
    h = ((h >> 14) & 1) | ((h << 1) & 0x7fff);
    h ^= pw.charCodeAt(i);
  }
  h = ((h >> 14) & 1) | ((h << 1) & 0x7fff);
  h ^= pw.length;
  h ^= 0xce4b;
  return h.toString(16).toUpperCase().padStart(4, '0');
}

/* ============================================================ page setup */

export function pageSetupDialog(doc: SheetDoc): Promise<void> {
  return openDialog<void>((close) => <PageSetup doc={doc} close={() => close()} />).then(() => undefined);
}

function PageSetup({ doc, close }: { doc: SheetDoc; close: () => void }) {
  const sheet = doc.sheet;
  const [p, setP] = useState<SheetPrint>({ ...DEFAULT_PRINT, ...(sheet.print ?? {}) });
  const set = (x: Partial<SheetPrint>) => setP((cur) => ({ ...cur, ...x }));
  const sel = doc.sel.ranges[doc.sel.ranges.length - 1];
  return (
    <Modal
      title="Page setup"
      onClose={close}
      width={520}
      footer={
        <Footer
          onCancel={close}
          onOk={() => {
            doc.transact('Page setup', (tx) => {
              tx.prop(sheet, 'print');
              sheet.print = p;
            });
            close();
          }}
        />
      }
    >
      <div className="row gap">
        <Field label="Orientation">
          <Segmented value={p.orientation} onChange={(v) => set({ orientation: v })} options={[{ value: 'portrait', label: 'Portrait' }, { value: 'landscape', label: 'Landscape' }]} size="sm" />
        </Field>
        <Field label="Paper">
          <Select value={p.paper} onChange={(v) => set({ paper: v })} options={[{ value: 'letter', label: 'Letter' }, { value: 'a4', label: 'A4' }, { value: 'legal', label: 'Legal' }, { value: 'a3', label: 'A3' }, { value: 'a5', label: 'A5' }, { value: 'tabloid', label: 'Tabloid' }]} />
        </Field>
        <Field label="Margins">
          <Select value={p.margins} onChange={(v) => set({ margins: v })} options={[{ value: 'normal', label: 'Normal' }, { value: 'narrow', label: 'Narrow' }, { value: 'wide', label: 'Wide' }]} />
        </Field>
      </div>
      <div className="row gap">
        <Field label="Scaling">
          <Select value={p.fit} onChange={(v) => set({ fit: v })} options={[{ value: 'width', label: 'Fit all columns on one page' }, { value: 'page', label: 'Fit sheet on one page' }, { value: 'none', label: 'Actual size (custom %)' }]} />
        </Field>
        {p.fit === 'none' && (
          <Field label="Scale %" inline>
            <NumberField value={p.scale} min={10} max={400} precision={0} onChange={(v) => set({ scale: v })} />
          </Field>
        )}
      </div>
      <div className="row gap wrap">
        <Checkbox checked={p.gridlines} onChange={(v) => set({ gridlines: v })} label="Print gridlines" />
        <Checkbox checked={p.headings} onChange={(v) => set({ headings: v })} label="Print row and column headings" />
        <Checkbox checked={!!p.centerH} onChange={(v) => set({ centerH: v })} label="Centre horizontally" />
      </div>
      <Field label="Print area">
        <div className="row gap">
          <input className="input mono" value={p.area ? rangeName(p.area) : ''} placeholder="Whole sheet" onChange={(e) => set({ area: parseRange(e.target.value) ?? undefined })} />
          <button className="btn btn-sm" onClick={() => set({ area: doc.clampToUsed(sheet, sel) })}>
            Use selection
          </button>
        </div>
      </Field>
      <Field label="Rows to repeat at the top of each page" inline>
        <NumberField value={p.repeatRows ?? 0} min={0} max={20} precision={0} onChange={(v) => set({ repeatRows: v || undefined })} />
      </Field>
      <div className="row gap">
        <Field label="Header" hint="&P page, &N pages, &D date, &F file, &A sheet">
          <input className="input" value={p.header ?? ''} onChange={(e) => set({ header: e.target.value || undefined })} />
        </Field>
        <Field label="Footer">
          <input className="input" value={p.footer ?? ''} onChange={(e) => set({ footer: e.target.value || undefined })} />
        </Field>
      </div>
    </Modal>
  );
}
