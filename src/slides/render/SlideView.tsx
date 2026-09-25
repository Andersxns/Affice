import { BarChart3, Image as ImageIcon, Shapes, Table2 } from 'lucide-react';
import { memo, useEffect, useRef, type CSSProperties, type ReactNode } from 'react';
import { cssFontStack } from '@/lib/fonts';
import { customGeometry, isLineGeom, presetGeometry, type Geometry, type PathFill } from '../geometry';
import {
  fillCss,
  isEmptyText,
  layoutOf,
  PT,
  resolveColor,
  shiftLum,
  slideBackground,
  type ArrowType,
  type ChartEl,
  type El,
  type Fill,
  type GroupEl,
  type ImageEl,
  type Line,
  type Presentation,
  type Shadow,
  type ShapeEl,
  type Slide,
  type TableEl,
  type TextBody,
  type Theme,
} from '../model';
import { cellLook, styledCellText } from '../tables';
import { PROMPTS, resolveFont } from '../themes';
import { slideChartConfig } from './chart';
import { roleOf, textFrameStyles, TextView, type TextContext } from './text';

export type RenderMode = 'edit' | 'thumb' | 'show' | 'print';

export interface RenderCtx {
  pres: Presentation;
  mode: RenderMode;
  slideNumber: number;
  /** Element whose text is being edited (rendered without its static text). */
  editing?: string | null;
  /** Transient geometry while dragging (id → frame). */
  preview?: Map<string, Partial<El>>;
}

/* ------------------------------------------------------------- helpers */

function dashArray(line: Line): string | undefined {
  const w = Math.max(1, line.width);
  switch (line.dash) {
    case 'dash':
      return `${w * 4} ${w * 3}`;
    case 'dot':
    case 'sysDot':
      return `${w} ${w * 2}`;
    case 'sysDash':
      return `${w * 3} ${w}`;
    case 'dashDot':
      return `${w * 4} ${w * 3} ${w} ${w * 3}`;
    case 'longDash':
      return `${w * 8} ${w * 3}`;
    default:
      return undefined;
  }
}

function shadowFilter(s: Shadow | undefined, theme: Theme): string | undefined {
  if (!s) return undefined;
  const rad = (s.angle * Math.PI) / 180;
  return `drop-shadow(${Math.round(Math.cos(rad) * s.dist * 10) / 10}px ${Math.round(Math.sin(rad) * s.dist * 10) / 10}px ${s.blur / 2}px ${resolveColor(s.color, theme)})`;
}

function modFill(base: string, mod: PathFill): string {
  switch (mod) {
    case 'darken':
      return shiftLum(base, -40);
    case 'darkenLess':
      return shiftLum(base, -20);
    case 'lighten':
      return shiftLum(base, 40);
    case 'lightenLess':
      return shiftLum(base, 20);
    default:
      return base;
  }
}

function Marker({ id, type, color, width, start }: { id: string; type: ArrowType; color: string; width: number; start: boolean }) {
  const s = Math.max(7, width * 3.2);
  const common = { id, markerUnits: 'userSpaceOnUse' as const, markerWidth: s, markerHeight: s, refX: s / 2, refY: s / 2, orient: start ? 'auto-start-reverse' : 'auto', viewBox: `0 0 ${s} ${s}` };
  switch (type) {
    case 'oval':
      return (
        <marker {...common}>
          <circle cx={s / 2} cy={s / 2} r={s / 2.6} fill={color} />
        </marker>
      );
    case 'diamond':
      return (
        <marker {...common}>
          <path d={`M0 ${s / 2}L${s / 2} 0L${s} ${s / 2}L${s / 2} ${s}Z`} fill={color} />
        </marker>
      );
    case 'arrow':
      return (
        <marker {...common} refX={s * 0.9}>
          <path d={`M0 0L${s} ${s / 2}L0 ${s}`} fill="none" stroke={color} strokeWidth={Math.max(1, width)} />
        </marker>
      );
    case 'stealth':
      return (
        <marker {...common} refX={s * 0.8}>
          <path d={`M0 0L${s} ${s / 2}L0 ${s}L${s * 0.3} ${s / 2}Z`} fill={color} />
        </marker>
      );
    default:
      return (
        <marker {...common} refX={s * 0.8}>
          <path d={`M0 0L${s} ${s / 2}L0 ${s}Z`} fill={color} />
        </marker>
      );
  }
}

function fillPaint(fill: Fill | undefined, id: string, theme: Theme, defs: ReactNode[], mod: PathFill = 'norm'): string {
  if (!fill || fill.type === 'none' || mod === 'none') return 'none';
  if (fill.type === 'solid') return modFill(resolveColor(fill.color, theme), mod);
  if (fill.type === 'gradient') {
    const gid = `${id}-g`;
    const stops = [...fill.stops].sort((a, b) => a.pos - b.pos).map((s, i) => <stop key={i} offset={s.pos} stopColor={resolveColor(s.color, theme)} />);
    defs.push(
      fill.radial ? (
        <radialGradient key="g" id={gid} cx="0.5" cy="0.5" r="0.5">
          {stops}
        </radialGradient>
      ) : (
        <linearGradient key="g" id={gid} x1="0" y1="0" x2="1" y2="0" gradientTransform={`rotate(${fill.angle} 0.5 0.5)`}>
          {stops}
        </linearGradient>
      ),
    );
    return `url(#${gid})`;
  }
  const pid = `${id}-p`;
  defs.push(
    <pattern key="p" id={pid} patternContentUnits="objectBoundingBox" width="1" height="1">
      <image href={fill.src} width="1" height="1" preserveAspectRatio={fill.mode === 'cover' ? 'xMidYMid slice' : 'none'} />
    </pattern>,
  );
  return `url(#${pid})`;
}

function frameStyle(e: El, extra?: CSSProperties): CSSProperties {
  const tr: string[] = [];
  if (e.rot) tr.push(`rotate(${e.rot}deg)`);
  return { left: e.x, top: e.y, width: e.w, height: e.h, transform: tr.length ? tr.join(' ') : undefined, opacity: e.opacity !== undefined && e.opacity < 1 ? e.opacity : undefined, ...extra };
}

/* -------------------------------------------------------------- shapes */

function geometryOf(e: ShapeEl): Geometry {
  if (e.geom === 'custom' && e.custom?.length) return customGeometry(e.custom, e.w, e.h);
  return presetGeometry(e.geom, Math.max(e.w, 0.01), Math.max(e.h, 0.01), e.adj);
}

function textCtx(e: El, rc: RenderCtx): TextContext {
  return { pres: rc.pres, role: roleOf(e.ph), ph: e.ph, slideNumber: rc.slideNumber };
}

const ShapeView = memo(function ShapeView({ e, rc }: { e: ShapeEl; rc: RenderCtx }) {
  const theme = rc.pres.theme;
  const g = geometryOf(e);
  const defs: ReactNode[] = [];
  const fill = e.fill ?? (isLineGeom(e.geom) ? { type: 'none' as const } : undefined);
  const line = e.line ?? undefined;
  const stroke = line && line.width > 0 ? resolveColor(line.color, theme) : 'none';
  const hasHead = line && line.head && line.head !== 'none';
  const hasTail = line && line.tail && line.tail !== 'none';
  if (line && hasHead) defs.push(<Marker key="h" id={`${e.id}-mh`} type={line.head!} color={stroke} width={line.width} start />);
  if (line && hasTail) defs.push(<Marker key="t" id={`${e.id}-mt`} type={line.tail!} color={stroke} width={line.width} start={false} />);
  const flip = e.flipH || e.flipV ? `scale(${e.flipH ? -1 : 1},${e.flipV ? -1 : 1})` : undefined;
  const editingHere = rc.editing === e.id;
  const showText = e.text && !editingHere;
  const lineLike = isLineGeom(e.geom) && !e.text;
  const empty = isEmptyText(e.text);
  const prompt = rc.mode === 'edit' && e.ph && empty ? PROMPTS[e.ph] : undefined;
  const hideEmptyPh = rc.mode !== 'edit' && e.ph && empty && (!e.fill || e.fill.type === 'none') && !line;
  if (hideEmptyPh) return null;
  const tr = g.text ?? [0, 0, 0, 0];
  const paints = g.paths.map((p) => (p.fill === 'none' ? 'none' : fillPaint(fill, e.id, theme, defs, p.fill === 'norm' ? 'norm' : p.fill)));
  // gradient/pattern defs are keyed per element: keep only the first of each kind
  const seen = new Set<string>();
  const uniqueDefs = defs.filter((d) => {
    const k = String((d as { key?: string }).key);
    if (seen.has(k)) return false;
    seen.add(k);
    return true;
  });
  return (
    <div className={`sl-el sl-shape${e.ph ? ' sl-ph' : ''}${prompt ? ' sl-ph-empty' : ''}${lineLike ? ' is-line' : ''}`} data-el={e.id} style={frameStyle(e)}>
      {(fill?.type !== 'none' || stroke !== 'none') && (
        // at least 1 px each way: a zero-height (horizontal) or zero-width (vertical) line would otherwise scale to nothing
        <svg className="sl-geom" width={Math.max(e.w, 1)} height={Math.max(e.h, 1)} viewBox={`0 0 ${Math.max(e.w, 1)} ${Math.max(e.h, 1)}`} style={{ transform: flip, filter: shadowFilter(e.shadow, theme) }} aria-hidden>
          {uniqueDefs.length > 0 && <defs>{uniqueDefs}</defs>}
          {g.paths.map((p, i) => (
            <path
              key={i}
              d={p.d}
              fill={paints[i]}
              fillRule="evenodd"
              stroke={p.stroke ? stroke : 'none'}
              strokeWidth={line?.width ?? 0}
              strokeDasharray={line ? dashArray(line) : undefined}
              strokeLinecap={line?.cap === 'round' ? 'round' : line?.cap === 'square' ? 'square' : 'butt'}
              strokeLinejoin="round"
              markerStart={hasHead ? `url(#${e.id}-mh)` : undefined}
              markerEnd={hasTail ? `url(#${e.id}-mt)` : undefined}
              vectorEffect="non-scaling-stroke"
            />
          ))}
          {lineLike && rc.mode === 'edit' && g.paths.map((p, i) => <path key={`hit${i}`} className="sl-hit" d={p.d} fill="none" stroke="transparent" strokeWidth={Math.max(12, (line?.width ?? 0) + 8)} />)}
        </svg>
      )}
      {showText && e.ph !== 'pic' && (
        <div className="sl-text-box" style={{ left: tr[0], top: tr[1], right: tr[2], bottom: tr[3] }}>
          <TextView body={e.text!} ctx={textCtx(e, rc)} prompt={prompt} />
        </div>
      )}
      {prompt && !editingHere && (e.ph === 'obj' || e.ph === 'pic') && <PlaceholderInsert kinds={e.ph === 'pic' ? ['picture'] : ['table', 'chart', 'picture', 'icon']} label={e.ph === 'pic' ? prompt : undefined} />}
      {editingHere && <EditSlot body={e.text ?? { paras: [{ runs: [] }] }} id={e.id} rect={tr} />}
    </div>
  );
});

export type PlaceholderInsertKind = 'table' | 'chart' | 'picture' | 'icon';

const PH_INSERT: Record<PlaceholderInsertKind, { icon: typeof Table2; label: string }> = {
  table: { icon: Table2, label: 'Insert table' },
  chart: { icon: BarChart3, label: 'Insert chart' },
  picture: { icon: ImageIcon, label: 'Insert picture' },
  icon: { icon: Shapes, label: 'Insert icon' },
};

/** Quick-insert buttons in the middle of empty content and picture placeholders (the canvas handles the clicks). */
function PlaceholderInsert({ kinds, label }: { kinds: PlaceholderInsertKind[]; label?: string }) {
  return (
    <div className="sl-ph-insert">
      <div className="sl-ph-insert-row">
        {kinds.map((k) => {
          const I = PH_INSERT[k].icon;
          return (
            <button key={k} type="button" data-ph-insert={k} title={PH_INSERT[k].label} aria-label={PH_INSERT[k].label}>
              <I size={30} strokeWidth={1.6} />
            </button>
          );
        })}
      </div>
      {label && <div className="sl-ph-insert-label">{label}</div>}
    </div>
  );
}

/** Where the text editor mounts while an element's text is being edited. */
function EditSlot({ body, id, rect }: { body: TextBody; id: string; rect?: [number, number, number, number] }) {
  const { outer, inner } = textFrameStyles(body);
  const content = (
    <div className="sl-text sl-editing" style={outer}>
      <div className="sl-text-inner" style={inner} data-slot={id} />
    </div>
  );
  return rect ? (
    <div className="sl-text-box" style={{ left: rect[0], top: rect[1], right: rect[2], bottom: rect[3] }}>
      {content}
    </div>
  ) : (
    content
  );
}

/* -------------------------------------------------------------- images */

const ImageView = memo(function ImageView({ e, rc }: { e: ImageEl; rc: RenderCtx }) {
  const theme = rc.pres.theme;
  const [l, t, r, b] = e.crop ?? [0, 0, 0, 0];
  const sw = 1 - l - r || 1;
  const sh = 1 - t - b || 1;
  const clip = e.geom && e.geom !== 'rect' ? `path('${presetGeometry(e.geom, e.w, e.h).paths[0].d}')` : undefined;
  const flip = e.flipH || e.flipV ? `scale(${e.flipH ? -1 : 1},${e.flipV ? -1 : 1})` : undefined;
  const line = e.line;
  const src = e.svg ? `data:image/svg+xml;charset=utf-8,${encodeURIComponent(e.svg)}` : e.src;
  return (
    <div className="sl-el sl-image" data-el={e.id} style={frameStyle(e, { filter: shadowFilter(e.shadow, theme) })}>
      <div className="sl-img-clip" style={{ clipPath: clip, transform: flip }}>
        <img src={src} alt={e.alt ?? ''} draggable={false} style={{ left: `${(-l / sw) * 100}%`, top: `${(-t / sh) * 100}%`, width: `${100 / sw}%`, height: `${100 / sh}%` }} />
      </div>
      {line && line.width > 0 && (
        <svg className="sl-geom" width={e.w} height={e.h} viewBox={`0 0 ${e.w} ${e.h}`} aria-hidden>
          <path d={presetGeometry(e.geom ?? 'rect', e.w, e.h).paths[0].d} fill="none" stroke={resolveColor(line.color, theme)} strokeWidth={line.width} strokeDasharray={dashArray(line)} />
        </svg>
      )}
    </div>
  );
});

/* -------------------------------------------------------------- tables */

function borderCss(l: Line | null | undefined, theme: Theme): string | undefined {
  if (!l || l.width <= 0) return undefined;
  return `${Math.max(0.5, l.width)}px ${l.dash && l.dash !== 'solid' ? 'dashed' : 'solid'} ${resolveColor(l.color, theme)}`;
}

const TableView = memo(function TableView({ e, rc }: { e: TableEl; rc: RenderCtx; editingCell?: [number, number] | null }) {
  const theme = rc.pres.theme;
  const ctx: TextContext = { pres: rc.pres, role: 'other', slideNumber: rc.slideNumber };
  return (
    <div className="sl-el sl-table" data-el={e.id} style={frameStyle(e)}>
      <table>
        <colgroup>
          {e.cols.map((w, i) => (
            <col key={i} style={{ width: w }} />
          ))}
        </colgroup>
        <tbody>
          {e.rows.map((row, r) => (
            <tr key={r} style={{ height: row.h }}>
              {row.cells.map((cell, c) => {
                if (cell.merged) return null;
                const look = cellLook(e, r, c);
                const fill = cell.fill ?? look.fill;
                const borders = cell.borders ?? look.borders;
                const body = styledCellText(e, r, c, cell.text);
                return (
                  <td
                    key={c}
                    rowSpan={cell.rowSpan}
                    colSpan={cell.colSpan}
                    data-cell={`${r}:${c}`}
                    style={{ background: fillCss(fill, theme), borderTop: borderCss(borders[0], theme), borderRight: borderCss(borders[1], theme), borderBottom: borderCss(borders[2], theme), borderLeft: borderCss(borders[3], theme) }}
                  >
                    {rc.editing !== `${e.id}:${r}:${c}` ? <TextView body={body} ctx={ctx} /> : <EditSlot body={body} id={`${e.id}:${r}:${c}`} />}
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
});

/* -------------------------------------------------------------- charts */

const ChartView = memo(function ChartView({ e, rc }: { e: ChartEl; rc: RenderCtx }) {
  const canvas = useRef<HTMLCanvasElement>(null);
  const animate = rc.mode === 'show';
  useEffect(() => {
    let chart: { destroy(): void } | null = null;
    let cancelled = false;
    void import('chart.js/auto').then(({ default: Chart }) => {
      if (cancelled || !canvas.current) return;
      const cfg = slideChartConfig(e.chart, rc.pres.theme, animate);
      if (rc.mode === 'thumb' || rc.mode === 'print') cfg.options = { ...cfg.options, devicePixelRatio: rc.mode === 'print' ? 2 : 1, responsive: false };
      chart = new Chart(canvas.current, cfg);
      canvas.current.dataset.ready = '1';
    });
    return () => {
      cancelled = true;
      chart?.destroy();
    };
  }, [e.chart, e.w, e.h, rc.pres.theme, rc.mode, animate]);
  return (
    <div className="sl-el sl-chart" data-el={e.id} style={frameStyle(e)}>
      <canvas ref={canvas} width={e.w} height={e.h} />
    </div>
  );
});

/* -------------------------------------------------------------- groups */

const GroupView = memo(function GroupView({ e, rc }: { e: GroupEl; rc: RenderCtx }) {
  return (
    <div className="sl-el sl-group" data-el={e.id} style={frameStyle(e)}>
      {e.children.map((c) => (
        <ElementView key={c.id} e={{ ...c, x: c.x - e.x, y: c.y - e.y } as El} rc={rc} />
      ))}
    </div>
  );
});

const touches = (editing: string | null | undefined, id: string) => !!editing && (editing === id || editing.startsWith(`${id}:`));

function sameRender(a: { e: El; rc: RenderCtx }, b: { e: El; rc: RenderCtx }): boolean {
  return (
    a.e === b.e &&
    a.rc.mode === b.rc.mode &&
    a.rc.slideNumber === b.rc.slideNumber &&
    a.rc.pres.theme === b.rc.pres.theme &&
    a.rc.pres.master === b.rc.pres.master &&
    (a.rc.editing === b.rc.editing || (!touches(a.rc.editing, a.e.id) && !touches(b.rc.editing, a.e.id))) &&
    a.rc.preview?.get(a.e.id) === b.rc.preview?.get(b.e.id)
  );
}

export const ElementView = memo(function ElementView({ e, rc }: { e: El; rc: RenderCtx }) {
  const p = rc.preview?.get(e.id);
  const el = p ? ({ ...e, ...p } as El) : e;
  if (el.hidden && rc.mode !== 'edit') return null;
  switch (el.type) {
    case 'shape':
      return <ShapeView e={el} rc={rc} />;
    case 'image':
      return <ImageView e={el} rc={rc} />;
    case 'table':
      return <TableView e={el} rc={rc} />;
    case 'chart':
      return <ChartView e={el} rc={rc} />;
    case 'group':
      return <GroupView e={el} rc={rc} />;
    default:
      return null;
  }
}, sameRender);

/* --------------------------------------------------------------- slide */

function Footers({ pres, slide, slideNumber }: { pres: Presentation; slide: Slide; slideNumber: number }) {
  const f = pres.footer;
  if (!f || (!f.text && !f.slideNumber && !f.date)) return null;
  const layout = layoutOf(pres, slide);
  if (f.skipTitle && layout?.type === 'title') return null;
  const theme = pres.theme;
  const sx = pres.size.w / 1280;
  const style: CSSProperties = { fontFamily: cssFontStack(resolveFont('+minor', theme)), fontSize: `${12 * PT}px`, color: resolveColor('@tx1+45', theme) };
  const y = pres.size.h - 53;
  return (
    <>
      {f.date && (
        <div className="sl-footer" style={{ ...style, left: 88 * sx, top: y, width: 288 * sx }}>
          {new Date().toLocaleDateString()}
        </div>
      )}
      {f.text && (
        <div className="sl-footer" style={{ ...style, left: 424 * sx, top: y, width: 432 * sx, textAlign: 'center' }}>
          {f.text}
        </div>
      )}
      {f.slideNumber && (
        <div className="sl-footer" style={{ ...style, left: 904 * sx, top: y, width: 288 * sx, textAlign: 'right' }}>
          {slideNumber}
        </div>
      )}
    </>
  );
}

export interface SlideViewProps {
  pres: Presentation;
  slide: Slide;
  slideNumber: number;
  mode: RenderMode;
  editing?: string | null;
  preview?: Map<string, Partial<El>>;
  /** Extra layers rendered inside the slide (editor overlays). */
  children?: ReactNode;
  className?: string;
  style?: CSSProperties;
}

export const SlideView = memo(function SlideView({ pres, slide, slideNumber, mode, editing, preview, children, className = '', style }: SlideViewProps) {
  const rc: RenderCtx = { pres, mode, slideNumber, editing, preview };
  const layout = layoutOf(pres, slide);
  const bg = slideBackground(pres, slide);
  return (
    <div className={`sl-slide mode-${mode} ${className}`} style={{ width: pres.size.w, height: pres.size.h, background: fillCss(bg, pres.theme) ?? resolveColor('@bg1', pres.theme), ...style }} data-slide={slide.id}>
      {!slide.hideDecor && (
        <div className="sl-decor" aria-hidden>
          {pres.master.decor.map((d) => (
            <ElementView key={d.id} e={d} rc={{ ...rc, mode: mode === 'edit' ? 'thumb' : mode }} />
          ))}
          {layout?.decor.map((d) => (
            <ElementView key={d.id} e={d} rc={{ ...rc, mode: mode === 'edit' ? 'thumb' : mode }} />
          ))}
        </div>
      )}
      {slide.elements.map((e) => (
        <ElementView key={e.id} e={e} rc={rc} />
      ))}
      <Footers pres={pres} slide={slide} slideNumber={slideNumber} />
      {children}
    </div>
  );
});
