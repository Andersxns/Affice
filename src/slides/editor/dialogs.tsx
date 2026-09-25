import { Image as ImageIcon, Plus, Search, Trash2 } from 'lucide-react';
import { useEffect, useMemo, useRef, useState, type ReactNode } from 'react';
import { flushSync } from 'react-dom';
import { createRoot } from 'react-dom/client';
import { api } from '@/lib/platform';
import { bytesToDataUrl, mimeFromExt } from '@/lib/utils';
import { ColorField } from '@/ui/color';
import { Checkbox, NumberField, Segmented, Select, Slider } from '@/ui/controls';
import { Modal, openDialog } from '@/ui/dialog';
import type { SlidesDoc } from '../doc';
import { fillCss, resolveColor, type Dash, type El, type Fill, type ImageEl, type Line, type Presentation, type Shadow, type SlideChart, type SlideChartType, type TextBody, type Theme } from '../model';
import { DEFAULT_INSET } from '../render/text';
import { colorProps, describeColor } from './colors';

const Footer = ({ onCancel, onOk, okLabel = 'OK', extra, disabled }: { onCancel: () => void; onOk: () => void; okLabel?: string; extra?: ReactNode; disabled?: boolean }) => (
  <>
    {extra}
    <span className="spacer" />
    <button className="btn" onClick={onCancel}>
      Cancel
    </button>
    <button className="btn btn-primary" onClick={onOk} disabled={disabled}>
      {okLabel}
    </button>
  </>
);

export async function pickImageDataUrl(): Promise<{ src: string; name: string; w: number; h: number } | null> {
  const f = await api.files.pickImage();
  if (!f) return null;
  const ext = f.name.split('.').pop()?.toLowerCase() ?? 'png';
  const src = bytesToDataUrl(f.data, mimeFromExt(ext));
  const size = await imageSize(src);
  return { src, name: f.name, ...size };
}

export function imageSize(src: string): Promise<{ w: number; h: number }> {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => resolve({ w: img.naturalWidth || 400, h: img.naturalHeight || 300 });
    img.onerror = () => resolve({ w: 400, h: 300 });
    img.src = src;
  });
}

/* ============================================================ fill editor */

type FillKind = 'none' | 'solid' | 'gradient' | 'image';

function FillEditor({ fill, onChange, theme, allowNone = true }: { fill: Fill; onChange: (f: Fill) => void; theme: Theme; allowNone?: boolean }) {
  const cp = colorProps(theme);
  const kind: FillKind = fill.type;
  const setKind = (k: FillKind) => {
    if (k === kind) return;
    if (k === 'none') onChange({ type: 'none' });
    else if (k === 'solid') onChange({ type: 'solid', color: fill.type === 'gradient' ? fill.stops[0].color : '@accent1' });
    else if (k === 'gradient') {
      const base = fill.type === 'solid' ? fill.color : '@accent1';
      onChange({ type: 'gradient', angle: 90, stops: [{ pos: 0, color: base }, { pos: 1, color: base.startsWith('@') ? `${base.replace(/[+-]\d+/, '')}-50` : '#000000' }] });
    } else void pickImageDataUrl().then((img) => img && onChange({ type: 'image', src: img.src, mode: 'cover' }));
  };
  const opts: { value: FillKind; label: string }[] = [
    ...(allowNone ? [{ value: 'none' as const, label: 'None' }] : []),
    { value: 'solid', label: 'Solid' },
    { value: 'gradient', label: 'Gradient' },
    { value: 'image', label: 'Picture' },
  ];
  return (
    <div className="fill-editor">
      <Segmented size="sm" value={kind} onChange={setKind} options={opts} />
      {fill.type === 'solid' && (
        <div className="form-grid">
          <span>Colour</span>
          <ColorField value={fill.color.replace(/\/\d+$/, '')} onChange={(c) => onChange({ type: 'solid', color: c ?? '@accent1' })} describe={describeColor} {...cp} />
          <span>Transparency</span>
          <AlphaSlider color={fill.color} onChange={(c) => onChange({ type: 'solid', color: c })} />
        </div>
      )}
      {fill.type === 'gradient' && (
        <div className="form-grid">
          {fill.stops.map((s, i) => (
            <div key={i} className="contents">
              <span>{i === 0 ? 'From' : i === fill.stops.length - 1 ? 'To' : 'Middle'}</span>
              <div className="row gap-6">
                <ColorField value={s.color} onChange={(c) => onChange({ ...fill, stops: fill.stops.map((x, j) => (j === i ? { ...x, color: c ?? '#ffffff' } : x)) })} describe={describeColor} {...cp} />
                {fill.stops.length > 2 && (
                  <button className="icon-btn icon-btn-sm" aria-label="Remove stop" onClick={() => onChange({ ...fill, stops: fill.stops.filter((_, j) => j !== i) })}>
                    <Trash2 size={14} />
                  </button>
                )}
              </div>
            </div>
          ))}
          <span />
          <button className="btn btn-sm" disabled={fill.stops.length >= 4} onClick={() => onChange({ ...fill, stops: [...fill.stops.slice(0, -1), { pos: 0.5, color: '@accent2' }, fill.stops[fill.stops.length - 1]].map((x, j, arr) => ({ ...x, pos: j / (arr.length - 1) })) })}>
            <Plus size={14} /> Add colour
          </button>
          <span>Type</span>
          <Segmented size="sm" value={fill.radial ? 'radial' : 'linear'} onChange={(v) => onChange({ ...fill, radial: v === 'radial' })} options={[{ value: 'linear', label: 'Linear' }, { value: 'radial', label: 'Radial' }]} />
          {!fill.radial && (
            <>
              <span>Angle</span>
              <NumberField value={fill.angle} min={0} max={359} step={15} suffix="°" onChange={(v) => onChange({ ...fill, angle: v })} />
            </>
          )}
        </div>
      )}
      {fill.type === 'image' && (
        <div className="form-grid">
          <span>Picture</span>
          <button className="btn btn-sm" onClick={() => void pickImageDataUrl().then((img) => img && onChange({ ...fill, src: img.src }))}>
            <ImageIcon size={14} /> Choose picture…
          </button>
          <span>Fit</span>
          <Select value={fill.mode ?? 'stretch'} onChange={(v) => onChange({ ...fill, mode: v })} options={[{ value: 'cover', label: 'Fill (crop to fit)' }, { value: 'stretch', label: 'Stretch' }, { value: 'tile', label: 'Tile' }]} />
        </div>
      )}
      <div className="fill-preview" style={{ background: fillCss(fill, theme) ?? 'transparent' }} />
    </div>
  );
}

function AlphaSlider({ color, onChange }: { color: string; onChange: (c: string) => void }) {
  const m = /\/(\d+)$/.exec(color);
  const alpha = m ? parseInt(m[1], 10) : 100;
  const base = color.replace(/\/\d+$/, '');
  const set = (transparency: number) => {
    const a = Math.round(100 - transparency);
    if (base.startsWith('@')) onChange(a >= 100 ? base : `${base}/${a}`);
    else onChange(a >= 100 ? base.slice(0, 7) : `${base.slice(0, 7)}${Math.round((a / 100) * 255).toString(16).padStart(2, '0')}`);
  };
  const hexAlpha = !base.startsWith('@') && base.length === 9 ? Math.round((parseInt(base.slice(7, 9), 16) / 255) * 100) : alpha;
  return (
    <div className="row gap-8">
      <Slider value={100 - hexAlpha} min={0} max={100} onChange={set} />
      <span className="muted small">{100 - hexAlpha}%</span>
    </div>
  );
}

/* ====================================================== format background */

export function formatBackgroundDialog(doc: SlidesDoc, apply: (fill: Fill | undefined, all: boolean, hideDecor: boolean) => void): Promise<void> {
  const slide = doc.slide;
  const pres = doc.pres;
  const current: Fill = slide.background ?? pres.layouts.find((l) => l.id === slide.layout)?.background ?? pres.master.background;
  return openDialog<void>((close) => <BackgroundDlg pres={pres} initial={current} hideDecor={!!slide.hideDecor} apply={apply} close={() => close()} />).then(() => undefined);
}

function BackgroundDlg({ pres, initial, hideDecor: hd0, apply, close }: { pres: Presentation; initial: Fill; hideDecor: boolean; apply: (f: Fill | undefined, all: boolean, hide: boolean) => void; close: () => void }) {
  const [fill, setFill] = useState<Fill>(initial.type === 'none' ? { type: 'solid', color: '@bg1' } : initial);
  const [hide, setHide] = useState(hd0);
  return (
    <Modal
      title="Format background"
      onClose={close}
      width={460}
      footer={
        <Footer
          onCancel={close}
          okLabel="Apply"
          onOk={() => {
            apply(fill, false, hide);
            close();
          }}
          extra={
            <button
              className="btn"
              onClick={() => {
                apply(fill, true, hide);
                close();
              }}
            >
              Apply to all
            </button>
          }
        />
      }
    >
      <FillEditor fill={fill} onChange={setFill} theme={pres.theme} allowNone={false} />
      <div style={{ marginTop: 12 }}>
        <Checkbox checked={hide} onChange={setHide} label="Hide background graphics" />
      </div>
    </Modal>
  );
}

/* =========================================================== format shape */

export interface ShapeFormat {
  fill?: Fill;
  line?: Line | null;
  shadow?: Shadow | null;
  frame?: { x: number; y: number; w: number; h: number; rot: number };
  text?: Partial<Pick<TextBody, 'autofit' | 'wrap' | 'inset' | 'anchor' | 'columns'>>;
  alt?: string;
  opacity?: number;
}

export function formatShapeDialog(doc: SlidesDoc, el: El): Promise<ShapeFormat | undefined> {
  return openDialog<ShapeFormat>((close) => <ShapeDlg theme={doc.pres.theme} el={el} close={close} />);
}

const DASHES: { value: Dash; label: string }[] = [
  { value: 'solid', label: 'Solid' },
  { value: 'dash', label: 'Dash' },
  { value: 'dot', label: 'Round dot' },
  { value: 'sysDash', label: 'Square dot' },
  { value: 'dashDot', label: 'Dash dot' },
  { value: 'longDash', label: 'Long dash' },
];

function ShapeDlg({ theme, el, close }: { theme: Theme; el: El; close: (v?: ShapeFormat) => void }) {
  const cp = colorProps(theme);
  const [tab, setTab] = useState<'fill' | 'effects' | 'size' | 'text' | 'alt'>('fill');
  const isShape = el.type === 'shape';
  const [fill, setFill] = useState<Fill>(isShape && el.fill ? el.fill : { type: 'none' });
  const initialLine = el.type === 'shape' || el.type === 'image' ? (el.line ?? null) : null;
  const [lineOn, setLineOn] = useState(!!initialLine);
  const [line, setLine] = useState<Line>(initialLine ?? { color: '@tx1', width: 1.33 });
  const initialShadow = el.type === 'shape' || el.type === 'image' ? (el.shadow ?? null) : null;
  const [shadowOn, setShadowOn] = useState(!!initialShadow);
  const [shadow, setShadow] = useState<Shadow>(initialShadow ?? { color: '#000000/40', blur: 12, dist: 4, angle: 45 });
  const [frame, setFrame] = useState({ x: el.x, y: el.y, w: el.w, h: el.h, rot: el.rot ?? 0 });
  const body = el.type === 'shape' ? el.text : undefined;
  const [autofit, setAutofit] = useState(body?.autofit ?? 'none');
  const [wrap, setWrap] = useState(body?.wrap !== false);
  const [inset, setInset] = useState<[number, number, number, number]>(body?.inset ?? DEFAULT_INSET);
  const [cols, setCols] = useState(body?.columns ?? 1);
  const [alt, setAlt] = useState(el.alt ?? '');
  const [opacity, setOpacity] = useState(Math.round((1 - (el.opacity ?? 1)) * 100));
  const cm = (px: number) => Math.round((px / 96) * 2.54 * 100) / 100;
  const px = (c: number) => (c / 2.54) * 96;
  const tabs: [typeof tab, string][] = [
    ...(isShape ? [['fill', 'Fill & line'] as [typeof tab, string]] : el.type === 'image' ? [['fill', 'Border'] as [typeof tab, string]] : []),
    ['effects', 'Effects'],
    ['size', 'Size & position'],
    ...(isShape ? [['text', 'Text box'] as [typeof tab, string]] : []),
    ['alt', 'Alt text'],
  ];
  const ok = () =>
    close({
      fill: isShape ? fill : undefined,
      line: el.type === 'shape' || el.type === 'image' ? (lineOn ? line : null) : undefined,
      shadow: el.type === 'shape' || el.type === 'image' ? (shadowOn ? shadow : null) : undefined,
      frame,
      text: isShape ? { autofit, wrap, inset, columns: cols > 1 ? cols : undefined } : undefined,
      alt,
      opacity: opacity ? 1 - opacity / 100 : undefined,
    });
  return (
    <Modal title="Format shape" onClose={() => close()} width={520} footer={<Footer onCancel={() => close()} onOk={ok} />}>
      <div className="dlg-tabs">
        {tabs.map(([id, label]) => (
          <button key={id} className={`dlg-tab${tab === id ? ' active' : ''}`} onClick={() => setTab(id)}>
            {label}
          </button>
        ))}
      </div>
      <div className="dlg-tab-body">
        {tab === 'fill' && (
          <>
            {isShape && (
              <>
                <div className="section-label">Fill</div>
                <FillEditor fill={fill} onChange={setFill} theme={theme} />
              </>
            )}
            <div className="section-label">{isShape ? 'Line' : 'Border'}</div>
            <Checkbox checked={lineOn} onChange={setLineOn} label={isShape ? 'Show line' : 'Show border'} />
            {lineOn && (
              <div className="form-grid">
                <span>Colour</span>
                <ColorField value={line.color} onChange={(c) => setLine({ ...line, color: c ?? '@tx1' })} describe={describeColor} {...cp} />
                <span>Width</span>
                <NumberField value={Math.round(((line.width * 72) / 96) * 100) / 100} min={0.25} max={100} step={0.25} suffix="pt" onChange={(v) => setLine({ ...line, width: (v * 96) / 72 })} />
                <span>Dash</span>
                <Select value={line.dash ?? 'solid'} onChange={(v) => setLine({ ...line, dash: v })} options={DASHES} />
                {isShape && (
                  <>
                    <span>Begin arrow</span>
                    <Select value={line.head ?? 'none'} onChange={(v) => setLine({ ...line, head: v })} options={ARROWS} />
                    <span>End arrow</span>
                    <Select value={line.tail ?? 'none'} onChange={(v) => setLine({ ...line, tail: v })} options={ARROWS} />
                  </>
                )}
              </div>
            )}
          </>
        )}
        {tab === 'effects' && (
          <>
            <div className="section-label">Shadow</div>
            <Checkbox checked={shadowOn} onChange={setShadowOn} label="Drop shadow" />
            {shadowOn && (
              <div className="form-grid">
                <span>Colour</span>
                <ColorField value={shadow.color} onChange={(c) => setShadow({ ...shadow, color: c ?? '#000000' })} describe={describeColor} {...cp} />
                <span>Blur</span>
                <NumberField value={shadow.blur} min={0} max={100} suffix="px" onChange={(v) => setShadow({ ...shadow, blur: v })} />
                <span>Distance</span>
                <NumberField value={shadow.dist} min={0} max={100} suffix="px" onChange={(v) => setShadow({ ...shadow, dist: v })} />
                <span>Angle</span>
                <NumberField value={shadow.angle} min={0} max={359} step={15} suffix="°" onChange={(v) => setShadow({ ...shadow, angle: v })} />
              </div>
            )}
            <div className="section-label">Transparency</div>
            <div className="row gap-8">
              <Slider value={opacity} min={0} max={95} onChange={setOpacity} />
              <span className="muted small">{opacity}%</span>
            </div>
          </>
        )}
        {tab === 'size' && (
          <div className="form-grid">
            <span>Width</span>
            <NumberField value={cm(frame.w)} min={0.01} max={500} step={0.1} suffix="cm" onChange={(v) => setFrame({ ...frame, w: px(v) })} />
            <span>Height</span>
            <NumberField value={cm(frame.h)} min={0.01} max={500} step={0.1} suffix="cm" onChange={(v) => setFrame({ ...frame, h: px(v) })} />
            <span>Rotation</span>
            <NumberField value={frame.rot} min={-360} max={360} step={15} suffix="°" onChange={(v) => setFrame({ ...frame, rot: ((v % 360) + 360) % 360 })} />
            <span>Horizontal position</span>
            <NumberField value={cm(frame.x)} min={-500} max={500} step={0.1} suffix="cm" onChange={(v) => setFrame({ ...frame, x: px(v) })} />
            <span>Vertical position</span>
            <NumberField value={cm(frame.y)} min={-500} max={500} step={0.1} suffix="cm" onChange={(v) => setFrame({ ...frame, y: px(v) })} />
          </div>
        )}
        {tab === 'text' && (
          <>
            <div className="section-label">AutoFit</div>
            <div className="radio-list">
              {(
                [
                  ['none', 'Do not autofit'],
                  ['shrink', 'Shrink text on overflow'],
                  ['resize', 'Resize shape to fit text'],
                ] as const
              ).map(([v, label]) => (
                <label key={v} className="radio">
                  <input type="radio" checked={autofit === v} onChange={() => setAutofit(v)} /> {label}
                </label>
              ))}
            </div>
            <div className="section-label">Margins</div>
            <div className="form-grid">
              {(['Left', 'Top', 'Right', 'Bottom'] as const).map((label, i) => (
                <div key={label} className="contents">
                  <span>{label}</span>
                  <NumberField value={cm(inset[i])} min={0} max={50} step={0.05} suffix="cm" onChange={(v) => setInset(inset.map((x, j) => (j === i ? px(v) : x)) as [number, number, number, number])} />
                </div>
              ))}
              <span>Columns</span>
              <NumberField value={cols} min={1} max={6} step={1} precision={0} onChange={setCols} />
            </div>
            <div style={{ marginTop: 10 }}>
              <Checkbox checked={wrap} onChange={setWrap} label="Wrap text in shape" />
            </div>
          </>
        )}
        {tab === 'alt' && (
          <div className="field">
            <label className="field-label">Description (read aloud by screen readers)</label>
            <textarea className="input" rows={5} value={alt} onChange={(e) => setAlt(e.target.value)} placeholder="Describe the picture or shape" />
          </div>
        )}
      </div>
    </Modal>
  );
}

const ARROWS: { value: NonNullable<Line['head']>; label: string }[] = [
  { value: 'none', label: 'None' },
  { value: 'triangle', label: 'Arrow' },
  { value: 'stealth', label: 'Stealth arrow' },
  { value: 'arrow', label: 'Open arrow' },
  { value: 'oval', label: 'Oval' },
  { value: 'diamond', label: 'Diamond' },
];

/* ========================================================= header footer */

export function headerFooterDialog(doc: SlidesDoc): Promise<Presentation['footer'] | undefined> {
  return openDialog<Presentation['footer']>((close) => <HeaderFooterDlg initial={doc.pres.footer ?? {}} close={close} />);
}

function HeaderFooterDlg({ initial, close }: { initial: NonNullable<Presentation['footer']>; close: (v?: Presentation['footer']) => void }) {
  const [f, setF] = useState(initial);
  const [on, setOn] = useState(!!initial.text);
  return (
    <Modal title="Header and footer" onClose={() => close()} width={420} footer={<Footer onCancel={() => close()} okLabel="Apply to all" onOk={() => close({ ...f, text: on ? f.text || ' ' : undefined })} />}>
      <div className="stack-10">
        <Checkbox checked={!!f.date} onChange={(v) => setF({ ...f, date: v })} label="Date and time (updates automatically)" />
        <Checkbox checked={!!f.slideNumber} onChange={(v) => setF({ ...f, slideNumber: v })} label="Slide number" />
        <Checkbox checked={on} onChange={setOn} label="Footer" />
        {on && <input className="input" value={f.text ?? ''} placeholder="Footer text" onChange={(e) => setF({ ...f, text: e.target.value })} data-autofocus />}
        <Checkbox checked={!!f.skipTitle} onChange={(v) => setF({ ...f, skipTitle: v })} label="Don't show on title slide" />
      </div>
    </Modal>
  );
}

/* ========================================================== insert table */

export function insertTableDialog(): Promise<{ rows: number; cols: number } | undefined> {
  return openDialog<{ rows: number; cols: number }>((close) => <TableDlg close={close} />);
}

function TableDlg({ close }: { close: (v?: { rows: number; cols: number }) => void }) {
  const [cols, setCols] = useState(4);
  const [rows, setRows] = useState(3);
  return (
    <Modal title="Insert table" onClose={() => close()} width={320} footer={<Footer onCancel={() => close()} onOk={() => close({ rows, cols })} />}>
      <div className="form-grid">
        <span>Columns</span>
        <NumberField value={cols} min={1} max={30} step={1} onChange={(v) => setCols(Math.round(v))} />
        <span>Rows</span>
        <NumberField value={rows} min={1} max={60} step={1} onChange={(v) => setRows(Math.round(v))} />
      </div>
    </Modal>
  );
}

/* ============================================================= slide size */

export function slideSizeDialog(doc: SlidesDoc): Promise<{ w: number; h: number; scale: boolean } | undefined> {
  return openDialog<{ w: number; h: number; scale: boolean }>((close) => <SizeDlg w={doc.pres.size.w} h={doc.pres.size.h} close={close} />);
}

function SizeDlg({ w: w0, h: h0, close }: { w: number; h: number; close: (v?: { w: number; h: number; scale: boolean }) => void }) {
  const [w, setW] = useState(Math.round((w0 / 96) * 2.54 * 100) / 100);
  const [h, setH] = useState(Math.round((h0 / 96) * 2.54 * 100) / 100);
  const [scale, setScale] = useState(true);
  const presets: [string, number, number][] = [
    ['Widescreen (16:9)', 33.867, 19.05],
    ['Standard (4:3)', 25.4, 19.05],
    ['On-screen show (16:10)', 25.4, 15.875],
    ['A4 paper', 27.517, 19.05],
    ['Letter paper', 25.4, 19.05],
    ['Square (1:1)', 19.05, 19.05],
    ['Portrait story (9:16)', 19.05, 33.867],
  ];
  return (
    <Modal title="Slide size" onClose={() => close()} width={420} footer={<Footer onCancel={() => close()} onOk={() => close({ w: Math.round((w / 2.54) * 96), h: Math.round((h / 2.54) * 96), scale })} />}>
      <div className="form-grid">
        <span>Preset</span>
        <Select
          value=""
          onChange={(v) => {
            const p = presets.find((x) => x[0] === v);
            if (p) {
              setW(p[1]);
              setH(p[2]);
            }
          }}
          options={[{ value: '', label: 'Choose…' }, ...presets.map((p) => ({ value: p[0], label: p[0] }))]}
        />
        <span>Width</span>
        <NumberField value={w} min={2.54} max={142.24} step={0.1} suffix="cm" onChange={setW} />
        <span>Height</span>
        <NumberField value={h} min={2.54} max={142.24} step={0.1} suffix="cm" onChange={setH} />
      </div>
      <div style={{ marginTop: 12 }}>
        <Checkbox checked={scale} onChange={setScale} label="Scale content to fit the new size" />
      </div>
    </Modal>
  );
}

/* ================================================================== link */

export function linkDialog(doc: SlidesDoc, current?: string): Promise<string | null | undefined> {
  return openDialog<string | null>((close) => <LinkDlg doc={doc} current={current} close={close} />);
}

function LinkDlg({ doc, current, close }: { doc: SlidesDoc; current?: string; close: (v?: string | null) => void }) {
  const internal = current?.startsWith('#');
  const [mode, setMode] = useState<'web' | 'slide'>(internal ? 'slide' : 'web');
  const [url, setUrl] = useState(internal ? 'https://' : (current ?? 'https://'));
  const [target, setTarget] = useState(internal ? current! : '#next');
  const slides = doc.pres.slides;
  const titleOf = (i: number) => {
    const t = slides[i].elements.find((e) => e.ph === 'title' || e.ph === 'ctrTitle');
    const text = t && t.type === 'shape' ? t.text?.paras.map((p) => p.runs.map((r) => r.text).join('')).join(' ') : '';
    return `${i + 1}. ${text || 'Slide'}`;
  };
  const ok = () => {
    if (mode === 'slide') return close(target);
    let link = url.trim();
    if (!link || link === 'https://') return close(null);
    if (!/^[a-z]+:/i.test(link)) link = /@/.test(link) ? `mailto:${link}` : `https://${link}`;
    close(link);
  };
  return (
    <Modal
      title="Insert link"
      onClose={() => close()}
      width={440}
      footer={<Footer onCancel={() => close()} onOk={ok} extra={current ? <button className="btn btn-danger" onClick={() => close(null)}>Remove link</button> : undefined} />}
    >
      <Segmented value={mode} onChange={setMode} options={[{ value: 'web', label: 'Web page or email' }, { value: 'slide', label: 'Place in this presentation' }]} />
      <div style={{ marginTop: 12 }}>
        {mode === 'web' ? (
          <input className="input" value={url} onChange={(e) => setUrl(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && ok()} data-autofocus />
        ) : (
          <select className="select" size={8} value={target} onChange={(e) => setTarget(e.target.value)} style={{ width: '100%', height: 200 }}>
            <option value="#first">First slide</option>
            <option value="#last">Last slide</option>
            <option value="#next">Next slide</option>
            <option value="#prev">Previous slide</option>
            {slides.map((_, i) => (
              <option key={i} value={`#slide:${i + 1}`}>
                {titleOf(i)}
              </option>
            ))}
          </select>
        )}
      </div>
    </Modal>
  );
}

/* ============================================================ chart data */

export function chartDataDialog(chart: SlideChart, theme: Theme): Promise<SlideChart | undefined> {
  return openDialog<SlideChart>((close) => <ChartDataDlg initial={chart} theme={theme} close={close} />);
}

const CHART_TYPES: { value: SlideChartType; label: string }[] = [
  { value: 'column', label: 'Column' },
  { value: 'bar', label: 'Bar' },
  { value: 'line', label: 'Line' },
  { value: 'area', label: 'Area' },
  { value: 'pie', label: 'Pie' },
  { value: 'doughnut', label: 'Doughnut' },
  { value: 'scatter', label: 'Scatter' },
  { value: 'radar', label: 'Radar' },
];

function ChartDataDlg({ initial, theme, close }: { initial: SlideChart; theme: Theme; close: (v?: SlideChart) => void }) {
  const [c, setC] = useState<SlideChart>(structuredClone(initial));
  const setCell = (r: number, col: number, text: string) => {
    const next = structuredClone(c);
    if (r === -1 && col >= 0) next.series[col].name = text;
    else if (col === -1 && r >= 0) next.categories[r] = text;
    else if (r >= 0 && col >= 0) {
      const v = text.trim() === '' ? null : Number(text.replace(',', '.'));
      next.series[col].values[r] = v === null || Number.isNaN(v) ? null : v;
    }
    setC(next);
  };
  const addRow = () => setC({ ...c, categories: [...c.categories, `Category ${c.categories.length + 1}`], series: c.series.map((s) => ({ ...s, values: [...s.values, 0] })) });
  const addCol = () => setC({ ...c, series: [...c.series, { name: `Series ${c.series.length + 1}`, values: c.categories.map(() => 0) }] });
  const delRow = (r: number) => c.categories.length > 1 && setC({ ...c, categories: c.categories.filter((_, i) => i !== r), series: c.series.map((s) => ({ ...s, values: s.values.filter((_, i) => i !== r) })) });
  const delCol = (i: number) => c.series.length > 1 && setC({ ...c, series: c.series.filter((_, j) => j !== i) });
  const onPaste = (e: React.ClipboardEvent) => {
    const text = e.clipboardData.getData('text/plain');
    if (!text.includes('\t') && !text.includes('\n')) return;
    e.preventDefault();
    const rows = text.replace(/\r/g, '').split('\n').filter((l) => l.trim()).map((l) => l.split('\t'));
    if (rows.length < 2) return;
    const head = rows[0].slice(1);
    const body = rows.slice(1);
    setC({ ...c, categories: body.map((r) => r[0]), series: head.map((name, j) => ({ name, values: body.map((r) => (r[j + 1] === undefined || r[j + 1] === '' ? null : Number(r[j + 1].replace(',', '.')))) })) });
  };
  return (
    <Modal title="Edit chart data" onClose={() => close()} width={720} footer={<Footer onCancel={() => close()} onOk={() => close(c)} />}>
      <div className="form-grid wide">
        <span>Chart type</span>
        <Select value={c.type} onChange={(v) => setC({ ...c, type: v })} options={CHART_TYPES} />
        <span>Title</span>
        <input className="input" value={c.title ?? ''} placeholder="No title" onChange={(e) => setC({ ...c, title: e.target.value || undefined })} />
      </div>
      <p className="muted small" style={{ margin: '10px 0 6px' }}>
        Tip: paste a table from Sheets or Excel into any cell to replace all the data.
      </p>
      <div className="chart-data thin-scroll" onPaste={onPaste}>
        <table>
          <thead>
            <tr>
              <th />
              {c.series.map((s, i) => (
                <th key={i}>
                  <div className="row gap-4">
                    <span className="series-dot" style={{ background: resolveColor(s.color ?? `@accent${(i % 6) + 1}`, theme) }} />
                    <input value={s.name} onChange={(e) => setCell(-1, i, e.target.value)} />
                    <button className="icon-btn icon-btn-sm" aria-label="Delete series" onClick={() => delCol(i)}>
                      <Trash2 size={13} />
                    </button>
                  </div>
                </th>
              ))}
              <th>
                <button className="btn btn-sm" onClick={addCol}>
                  <Plus size={13} /> Series
                </button>
              </th>
            </tr>
          </thead>
          <tbody>
            {c.categories.map((cat, r) => (
              <tr key={r}>
                <td>
                  <div className="row gap-4">
                    <button className="icon-btn icon-btn-sm" aria-label="Delete row" onClick={() => delRow(r)}>
                      <Trash2 size={13} />
                    </button>
                    <input value={cat} onChange={(e) => setCell(r, -1, e.target.value)} />
                  </div>
                </td>
                {c.series.map((s, i) => (
                  <td key={i}>
                    <input className="num" value={s.values[r] ?? ''} onChange={(e) => setCell(r, i, e.target.value)} />
                  </td>
                ))}
                <td />
              </tr>
            ))}
          </tbody>
        </table>
        <button className="btn btn-sm" style={{ marginTop: 8 }} onClick={addRow}>
          <Plus size={13} /> Category
        </button>
      </div>
    </Modal>
  );
}

/* ================================================================= icons */

export function iconPickerDialog(theme: Theme): Promise<{ svg: string; name: string } | undefined> {
  return openDialog<{ svg: string; name: string }>((close) => <IconDlg theme={theme} close={close} />);
}

type IconComp = React.ComponentType<{ size?: number; color?: string; strokeWidth?: number }>;

/** Serialises a Lucide icon component to standalone SVG markup. */
function iconSvg(Icon: IconComp, color: string): string {
  const host = document.createElement('div');
  const root = createRoot(host);
  flushSync(() => root.render(<Icon size={96} color={color} strokeWidth={2} />));
  const svg = host.innerHTML.replace('<svg ', '<svg xmlns="http://www.w3.org/2000/svg" ');
  root.unmount();
  return svg;
}

function IconDlg({ theme, close }: { theme: Theme; close: (v?: { svg: string; name: string }) => void }) {
  const [icons, setIcons] = useState<Record<string, IconComp> | null>(null);
  const [q, setQ] = useState('');
  const [color, setColor] = useState('@accent1');
  const [sel, setSel] = useState<string | null>(null);
  useEffect(() => {
    void import('./iconSet').then((m) => setIcons(m.icons as unknown as Record<string, IconComp>));
  }, []);
  const names = useMemo(() => {
    if (!icons) return [];
    const all = Object.keys(icons);
    const needle = q.trim().toLowerCase().replace(/\s+/g, '');
    return (needle ? all.filter((n) => n.toLowerCase().includes(needle)) : all).slice(0, 500);
  }, [icons, q]);
  const css = resolveColor(color, theme);
  const pick = (n: string) => icons && close({ svg: iconSvg(icons[n], css), name: n });
  const cp = colorProps(theme);
  return (
    <Modal title="Insert icon" onClose={() => close()} width={680} footer={<Footer onCancel={() => close()} okLabel="Insert" disabled={!sel} onOk={() => sel && pick(sel)} />}>
      <div className="row gap-8" style={{ marginBottom: 10 }}>
        <div className="input-icon grow">
          <Search size={15} />
          <input className="input" placeholder="Search icons (e.g. rocket, chart, user)" value={q} onChange={(e) => setQ(e.target.value)} data-autofocus />
        </div>
        <ColorField value={color} onChange={(c) => setColor(c ?? '@accent1')} describe={describeColor} {...cp} />
      </div>
      <div className="icon-grid thin-scroll">
        {!icons && <div className="muted">Loading icons…</div>}
        {icons &&
          names.map((n) => {
            const Icon = icons[n];
            return (
              <button key={n} type="button" className={`icon-cell${sel === n ? ' selected' : ''}`} data-tip={n.replace(/([a-z])([A-Z0-9])/g, '$1 $2')} onClick={() => setSel(n)} onDoubleClick={() => pick(n)}>
                <Icon size={26} color={css} strokeWidth={2} />
              </button>
            );
          })}
      </div>
      {icons && <div className="muted small" style={{ marginTop: 6 }}>{names.length === 500 ? 'Showing the first 500 matches — type to narrow down.' : `${names.length} icons`}</div>}
    </Modal>
  );
}

/* ================================================================== crop */

export function cropDialog(img: ImageEl): Promise<[number, number, number, number] | undefined> {
  return openDialog<[number, number, number, number]>((close) => <CropDlg img={img} close={close} />);
}

function CropDlg({ img, close }: { img: ImageEl; close: (v?: [number, number, number, number]) => void }) {
  const [crop, setCrop] = useState<[number, number, number, number]>(img.crop ?? [0, 0, 0, 0]);
  const box = useRef<HTMLDivElement>(null);
  const drag = (side: number) => (e: React.PointerEvent) => {
    e.preventDefault();
    const r = box.current!.getBoundingClientRect();
    const move = (ev: PointerEvent) => {
      const fx = Math.max(0, Math.min(1, (ev.clientX - r.left) / r.width));
      const fy = Math.max(0, Math.min(1, (ev.clientY - r.top) / r.height));
      setCrop((c) => {
        const n = [...c] as [number, number, number, number];
        if (side === 0) n[0] = Math.min(fx, 1 - n[2] - 0.05);
        if (side === 1) n[1] = Math.min(fy, 1 - n[3] - 0.05);
        if (side === 2) n[2] = Math.min(1 - fx, 1 - n[0] - 0.05);
        if (side === 3) n[3] = Math.min(1 - fy, 1 - n[1] - 0.05);
        return n;
      });
    };
    const up = () => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
  };
  const [l, t, rr, b] = crop;
  return (
    <Modal title="Crop picture" onClose={() => close()} width={620} footer={<Footer onCancel={() => close()} onOk={() => close(crop)} extra={<button className="btn" onClick={() => setCrop([0, 0, 0, 0])}>Reset</button>} />}>
      <div className="crop-stage">
        <div className="crop-box" ref={box}>
          <img src={img.svg ? `data:image/svg+xml;charset=utf-8,${encodeURIComponent(img.svg)}` : img.src} alt="" draggable={false} />
          <div className="crop-shade" style={{ left: 0, top: 0, width: `${l * 100}%`, bottom: 0 }} />
          <div className="crop-shade" style={{ right: 0, top: 0, width: `${rr * 100}%`, bottom: 0 }} />
          <div className="crop-shade" style={{ left: `${l * 100}%`, right: `${rr * 100}%`, top: 0, height: `${t * 100}%` }} />
          <div className="crop-shade" style={{ left: `${l * 100}%`, right: `${rr * 100}%`, bottom: 0, height: `${b * 100}%` }} />
          <div className="crop-frame" style={{ left: `${l * 100}%`, top: `${t * 100}%`, right: `${rr * 100}%`, bottom: `${b * 100}%` }}>
            <span className="crop-h l" onPointerDown={drag(0)} />
            <span className="crop-h t" onPointerDown={drag(1)} />
            <span className="crop-h r" onPointerDown={drag(2)} />
            <span className="crop-h b" onPointerDown={drag(3)} />
          </div>
        </div>
      </div>
      <p className="muted small">Drag the edges to crop. The cropped parts are kept, so you can undo the crop later.</p>
    </Modal>
  );
}
