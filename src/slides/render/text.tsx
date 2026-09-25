import { memo, type CSSProperties, type ReactNode } from 'react';
import { cssFontStack } from '@/lib/fonts';
import { paragraphNumbers, PT, resolveColor, type Para, type Presentation, type PlaceholderType, type Run, type TextBody, type Theme } from '../model';
import { resolveFont } from '../themes';

export type TextRole = 'title' | 'body' | 'other';

export function roleOf(ph: PlaceholderType | undefined): TextRole {
  if (ph === 'title' || ph === 'ctrTitle') return 'title';
  if (ph === 'body' || ph === 'subTitle' || ph === 'obj') return 'body';
  return 'other';
}

export const DEFAULT_INSET: [number, number, number, number] = [9.6, 4.8, 9.6, 4.8];
const BODY_SIZES = [28, 24, 20, 20, 20, 20, 20, 20, 20];
const LEVEL_BULLETS = ['•', '–', '•', '–', '»', '•', '–', '•', '–'];

export interface TextContext {
  pres: Presentation;
  role: TextRole;
  ph?: PlaceholderType;
  /** Title slide titles default to 60 pt (ctrTitle). */
  slideNumber?: number;
}

/** Effective run formatting after placeholder, body and paragraph defaults. */
export function effectiveRun(run: Omit<Run, 'text'>, para: Para, body: TextBody, ctx: TextContext): Required<Pick<Run, 'font' | 'size' | 'color'>> & Omit<Run, 'text'> {
  const t = ctx.pres.master.text;
  const lvl = para.level ?? 0;
  const base =
    ctx.role === 'title'
      ? { font: t.title.font, size: ctx.ph === 'ctrTitle' ? Math.round((t.title.size * 60) / 44) : t.title.size, color: t.title.color, b: t.title.bold }
      : ctx.role === 'body'
        ? { font: t.body.font, size: ctx.ph === 'subTitle' ? 24 : (BODY_SIZES[lvl] ?? 20) * (t.body.size / 28), color: t.body.color }
        : { font: t.other.font, size: t.other.size, color: t.other.color };
  return { ...base, ...(body.defaults ?? {}), ...run } as ReturnType<typeof effectiveRun>;
}

export function effectivePara(para: Para, ctx: TextContext): { bullet: Para['bullet']; marL: number; indent: number; lineSpacing: number; spaceBefore: number; spaceAfter: number } {
  const lvl = para.level ?? 0;
  const bodyBullets = ctx.role === 'body' && ctx.ph !== 'subTitle';
  const bullet = para.bullet ?? (bodyBullets ? { type: 'char' as const, char: LEVEL_BULLETS[lvl] ?? '•' } : { type: 'none' as const });
  const hasBullet = bullet.type !== 'none';
  // numbers need a wider hanging indent (PowerPoint's Numbering button uses 54 px in placeholders, 36 px in text boxes)
  const hang = bullet.type === 'num' ? (ctx.role === 'body' ? 54 : 36) : 24;
  const marL = para.marL ?? (hasBullet || (bodyBullets && lvl > 0) ? hang + lvl * 48 : lvl * 48);
  const indent = para.indent ?? (hasBullet ? -hang : 0);
  const defSpacing = ctx.role === 'title' ? 0.9 : ctx.role === 'body' ? (ctx.pres.master.text.body.lineSpacing ?? 0.9) : 1;
  return {
    bullet,
    marL,
    indent,
    lineSpacing: para.lineSpacing ?? defSpacing,
    spaceBefore: para.spaceBefore ?? (ctx.role === 'body' && ctx.ph !== 'subTitle' ? (lvl === 0 ? 10 : 5) : 0),
    spaceAfter: para.spaceAfter ?? 0,
  };
}

export function runCss(r: ReturnType<typeof effectiveRun>, theme: Theme): CSSProperties {
  const css: CSSProperties = {
    fontFamily: cssFontStack(resolveFont(r.font, theme)),
    fontSize: `${Math.round(r.size * PT * 100) / 100}px`,
    color: resolveColor(r.color, theme),
  };
  if (r.b) css.fontWeight = 700;
  if (r.i) css.fontStyle = 'italic';
  const deco = [r.u ? 'underline' : '', r.s ? 'line-through' : ''].filter(Boolean).join(' ');
  if (deco) css.textDecoration = deco;
  if (r.hl) css.backgroundColor = resolveColor(r.hl, theme);
  if (r.sup || r.sub) {
    css.verticalAlign = r.sup ? 'super' : 'sub';
    css.fontSize = `${Math.round(r.size * PT * 0.66 * 100) / 100}px`;
  }
  if (r.caps) css.textTransform = 'uppercase';
  if (r.spacing) css.letterSpacing = `${r.spacing * PT}px`;
  return css;
}

function fieldText(r: Run, ctx: TextContext): string {
  if (r.field === 'slidenum') return ctx.slideNumber ? String(ctx.slideNumber) : r.text;
  if (r.field === 'date') return new Date().toLocaleDateString();
  return r.text;
}

/** Paragraph CSS shared by the static renderer and the text editor (font size is the default run's size). */
export function paraCss(p: Para, body: TextBody, ctx: TextContext, theme: Theme, hasText: boolean, number?: string): { style: Record<string, string | number | undefined>; bullet?: string } {
  const e = effectivePara(p, ctx);
  const def = effectiveRun(p.runs.length ? {} : (p.endRun ?? {}), p, body, ctx);
  const first = effectiveRun(p.runs.find((r) => r.text) ?? p.endRun ?? {}, p, body, ctx);
  // the paragraph's own font size sets the minimum line height (the CSS strut), so it must not exceed the
  // text actually on the line: PowerPoint spaces lines by the fonts used in them
  const runSizes = p.runs.filter((r) => r.text).map((r) => effectiveRun(r, p, body, ctx).size);
  const strut = runSizes.length ? Math.min(def.size, ...runSizes) : def.size;
  const style: Record<string, string | number | undefined> = {
    textAlign: p.align ?? 'left',
    paddingLeft: `${e.marL}px`,
    textIndent: `${e.indent}px`,
    marginTop: e.spaceBefore ? `${Math.round(e.spaceBefore * PT * 100) / 100}px` : undefined,
    marginBottom: e.spaceAfter ? `${Math.round(e.spaceAfter * PT * 100) / 100}px` : undefined,
    lineHeight: Math.round(e.lineSpacing * 1.2 * 1000) / 1000,
    fontSize: `${Math.round(strut * PT * 100) / 100}px`,
    fontFamily: cssFontStack(resolveFont(def.font, theme)),
    color: resolveColor(def.color, theme),
    fontWeight: def.b ? 700 : undefined,
    fontStyle: def.i ? 'italic' : undefined,
  };
  let bullet: string | undefined;
  if (hasText && e.bullet?.type === 'char') {
    bullet = e.bullet.char;
    if (e.bullet.color) style['--bullet-color'] = resolveColor(e.bullet.color, theme);
    if (e.bullet.font) style['--bullet-font'] = cssFontStack(e.bullet.font);
  } else if (hasText && e.bullet?.type === 'num') bullet = number;
  if (bullet !== undefined) {
    style['--bullet-size'] = `${Math.round(first.size * PT * 100) / 100}px`;
    if (!style['--bullet-color']) style['--bullet-color'] = resolveColor(first.color, theme);
    if (e.indent < 0) style['--bullet-w'] = `${-e.indent}px`;
  }
  return { style, bullet };
}

interface TextViewProps {
  body: TextBody;
  ctx: TextContext;
  /** Shown in grey when the body is empty (placeholders in the editor). */
  prompt?: string;
  vertical?: boolean;
}

/** Outer (padding, anchoring) and inner (shrink-to-fit, columns) styles of a text body. */
export function textFrameStyles(body: TextBody): { outer: CSSProperties; inner: CSSProperties } {
  const inset = body.inset ?? DEFAULT_INSET;
  const anchor = body.anchor ?? 't';
  const outer: CSSProperties = {
    padding: `${inset[1]}px ${inset[2]}px ${inset[3]}px ${inset[0]}px`,
    justifyContent: anchor === 'm' ? 'center' : anchor === 'b' ? 'flex-end' : 'flex-start',
    whiteSpace: body.wrap === false ? 'pre' : 'pre-wrap',
  };
  if (body.vert === 'vert' || body.vert === 'vert270') outer.writingMode = 'vertical-rl';
  const inner: CSSProperties = {};
  if (body.fontScale && body.fontScale < 1) inner.zoom = body.fontScale;
  if (body.columns && body.columns > 1) {
    inner.columnCount = body.columns;
    inner.columnGap = '24px';
  }
  if (body.vert === 'vert270') inner.transform = 'rotate(180deg)';
  return { outer, inner };
}

export const TextView = memo(function TextView({ body, ctx, prompt }: TextViewProps) {
  const theme = ctx.pres.theme;
  const empty = body.paras.every((p) => p.runs.every((r) => !r.text));
  const numbers = paragraphNumbers(body.paras);
  const { outer, inner } = textFrameStyles(body);

  let content: ReactNode;
  if (empty && prompt) {
    const p0 = body.paras[0] ?? { runs: [] };
    const pc = paraCss({ ...p0, bullet: { type: 'none' } }, body, ctx, theme, false);
    content = (
      <p className="sl-p sl-prompt" style={{ ...(pc.style as CSSProperties), textIndent: 0, paddingLeft: 0 }}>
        {prompt}
      </p>
    );
  } else {
    content = body.paras.map((p, i) => {
      const hasText = p.runs.some((r) => r.text);
      const pc = paraCss(p, body, ctx, theme, hasText, numbers[i]);
      return (
        <p key={i} className="sl-p" style={pc.style as CSSProperties} data-bullet={pc.bullet}>
          {hasText ? (
            p.runs.map((r, j) => {
              const er = effectiveRun(r, p, body, ctx);
              const css = runCss(er, theme);
              const text = fieldText(r, ctx);
              return r.link ? (
                <a key={j} className="sl-link" href={r.link} style={css} onClick={(ev) => ev.preventDefault()}>
                  {text}
                </a>
              ) : (
                <span key={j} style={css}>
                  {text}
                </span>
              );
            })
          ) : (
            <br />
          )}
        </p>
      );
    });
  }

  return (
    <div className={`sl-text${empty && prompt ? ' is-empty' : ''}`} style={outer}>
      <div className="sl-text-inner" style={inner}>
        {content}
      </div>
    </div>
  );
});
