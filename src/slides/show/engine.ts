/**
 * Slide show logic that does not depend on React: the build (click steps) of a slide's animations,
 * which objects are visible at each step, and the keyframes for animations and transitions.
 */
import type { Anim, Dir, El, Presentation, Slide, Transition } from '../model';

export interface BuildItem {
  anim: Anim;
  /** Paragraph index when the text is animated by paragraph. */
  para?: number;
  /** Start offset within its step (ms). */
  offset: number;
}

export interface Build {
  /** Steps played automatically when the slide appears (before the first click). */
  auto: BuildItem[] | null;
  /** Click steps. */
  steps: BuildItem[][];
}

function paraCount(el: El | undefined): number {
  if (!el || el.type !== 'shape' || !el.text) return 0;
  return el.text.paras.length;
}

/** Groups a slide's animations into click steps with start offsets (With / After previous). */
export function buildOf(slide: Slide): Build {
  const anims = slide.anims ?? [];
  const els = new Map(slide.elements.map((e) => [e.id, e]));
  // expand paragraph-by-paragraph animations
  const flat: { anim: Anim; para?: number; start: Anim['start'] }[] = [];
  for (const a of anims) {
    if (!els.has(a.el)) continue;
    const n = a.byPara ? paraCount(els.get(a.el)) : 0;
    if (a.byPara && n > 1) {
      const el = els.get(a.el);
      for (let i = 0; i < n; i++) {
        if (el?.type === 'shape' && el.text && !el.text.paras[i].runs.some((r) => r.text)) continue;
        flat.push({ anim: a, para: i, start: flat.length && i > 0 ? (a.start === 'click' ? 'click' : 'after') : a.start });
      }
    } else flat.push({ anim: a, start: a.start });
  }
  const steps: BuildItem[][] = [];
  let auto: BuildItem[] | null = null;
  let cur: BuildItem[] | null = null;
  let prevStart = 0;
  let prevEnd = 0;
  for (const f of flat) {
    if (f.start === 'click' || !cur) {
      if (f.start === 'click') {
        cur = [];
        steps.push(cur);
      } else {
        auto ??= [];
        cur = auto;
      }
      prevStart = 0;
      prevEnd = 0;
    }
    const start = f.start === 'after' ? prevEnd : prevStart;
    const offset = start + f.anim.delay;
    cur.push({ anim: f.anim, para: f.para, offset });
    prevStart = offset;
    prevEnd = Math.max(prevEnd, offset + f.anim.dur * Math.max(1, f.anim.repeat ?? 1));
  }
  return { auto, steps };
}

export function stepDuration(items: BuildItem[]): number {
  return items.reduce((m, it) => Math.max(m, it.offset + it.anim.dur * Math.max(1, it.anim.repeat ?? 1)), 0);
}

const key = (el: string, para?: number) => (para === undefined ? el : `${el}#${para}`);

/**
 * Hidden objects (and paragraphs) after `done` click steps, with the automatic step applied or not.
 * While `playing` is a step index (-1 for the automatic step), exits in that step are still visible
 * because they are animating out.
 */
export function hiddenAt(build: Build, done: number, playing: number | null = null, autoApplied = true): Set<string> {
  const all: { item: BuildItem; step: number }[] = [];
  build.auto?.forEach((item) => all.push({ item, step: -1 }));
  build.steps.forEach((s, i) => s.forEach((item) => all.push({ item, step: i })));
  const hidden = new Set<string>();
  const seen = new Set<string>();
  // initial state: objects whose first animation is an entrance start hidden
  for (const { item } of all) {
    const k = key(item.anim.el, item.para);
    if (seen.has(k)) continue;
    seen.add(k);
    if (item.anim.cls === 'entr') hidden.add(k);
  }
  for (const { item, step } of all) {
    if (step === -1 ? !autoApplied : step >= done) continue;
    const k = key(item.anim.el, item.para);
    if (item.anim.cls === 'entr') hidden.delete(k);
    else if (item.anim.cls === 'exit' && !(playing !== null && step === playing)) hidden.add(k);
  }
  return hidden;
}

export function hiddenCss(scope: string, hidden: Set<string>): string {
  return [...hidden]
    .map((k) => {
      const [el, para] = k.split('#');
      const sel = `.${scope} [data-el="${CSS.escape(el)}"]`;
      return para === undefined ? `${sel}{visibility:hidden!important}` : `${sel} .sl-p:nth-of-type(${Number(para) + 1}){visibility:hidden!important}`;
    })
    .join('\n');
}

/* ------------------------------------------------------------ keyframes */

export interface Frame {
  x: number;
  y: number;
  w: number;
  h: number;
}

function offscreen(dir: Dir | undefined, f: Frame, pres: Presentation): string {
  switch (dir ?? 'd') {
    case 'u':
      return `translate(0px, ${-(f.y + f.h + 20)}px)`;
    case 'l':
      return `translate(${-(f.x + f.w + 20)}px, 0px)`;
    case 'r':
      return `translate(${pres.size.w - f.x + 20}px, 0px)`;
    default:
      return `translate(0px, ${pres.size.h - f.y + 20}px)`;
  }
}

function wipeFrom(dir: Dir | undefined): string {
  switch (dir ?? 'd') {
    case 'u':
      return 'inset(0 0 100% 0)';
    case 'l':
      return 'inset(0 100% 0 0)';
    case 'r':
      return 'inset(0 0 0 100%)';
    default:
      return 'inset(100% 0 0 0)';
  }
}

/** Keyframes and timing for an animation on an element with base transform `base`. */
export function animKeyframes(a: Anim, base: string, f: Frame, pres: Presentation): { frames: Keyframe[]; easing: string; duration: number } {
  const b = base ? `${base} ` : '';
  const dur = Math.max(1, a.dur);
  const out = (frames: Keyframe[], easing = 'cubic-bezier(.2,.7,.2,1)') => ({ frames: a.cls === 'exit' ? [...frames].reverse() : frames, easing: a.cls === 'exit' ? 'cubic-bezier(.6,0,.8,.3)' : easing, duration: dur });
  switch (a.effect) {
    case 'appear':
    case 'disappear':
      return { ...out([{ opacity: 0 }, { opacity: 1 }], 'steps(1, jump-end)'), duration: 1 };
    case 'fade':
      return out([{ opacity: 0 }, { opacity: 1 }], 'ease');
    case 'fly':
      return out([{ transform: `${offscreen(a.dir, f, pres)} ${b}`.trim() }, { transform: base || 'none' }]);
    case 'float':
      return out([{ opacity: 0, transform: `translate(0px, ${a.dir === 'd' ? -60 : 60}px) ${b}`.trim() }, { opacity: 1, transform: base || 'none' }]);
    case 'zoom':
      return out([{ opacity: 0, transform: `${b}scale(0.3)` }, { opacity: 1, transform: `${b}scale(1)` }]);
    case 'wipe':
      return out([{ clipPath: wipeFrom(a.dir) }, { clipPath: 'inset(0 0 0 0)' }], 'ease-out');
    case 'split':
      return out([{ clipPath: 'inset(0 50% 0 50%)' }, { clipPath: 'inset(0 0% 0 0%)' }], 'ease-out');
    case 'wheel':
      return out([{ opacity: 0, transform: `${b}rotate(-200deg) scale(0.2)` }, { opacity: 1, transform: `${b}rotate(0deg) scale(1)` }]);
    case 'bounce':
      return out(
        [
          { opacity: 0, transform: `${b}translate(0px, -160px)`, offset: 0 },
          { opacity: 1, transform: `${b}translate(0px, 0px)`, offset: 0.55 },
          { transform: `${b}translate(0px, -40px)`, offset: 0.72 },
          { transform: `${b}translate(0px, 0px)`, offset: 0.86 },
          { transform: `${b}translate(0px, -10px)`, offset: 0.93 },
          { transform: `${b}translate(0px, 0px)`, offset: 1 },
        ],
        'ease-in-out',
      );
    case 'pulse':
      return { frames: [{ transform: `${b}scale(1)` }, { transform: `${b}scale(1.1)` }, { transform: `${b}scale(1)` }], easing: 'ease-in-out', duration: dur };
    case 'spin':
      return { frames: [{ transform: `${b}rotate(0deg)` }, { transform: `${b}rotate(360deg)` }], easing: 'ease-in-out', duration: dur };
    case 'grow':
      return { frames: [{ transform: `${b}scale(1)` }, { transform: `${b}scale(1.35)` }, { transform: `${b}scale(1)` }], easing: 'ease-in-out', duration: dur };
    case 'teeter':
      return { frames: [0, -6, 6, -6, 6, 0].map((d) => ({ transform: `${b}rotate(${d}deg)` })), easing: 'ease-in-out', duration: dur };
    case 'transparency':
      return { frames: [{ opacity: 1 }, { opacity: 0.35 }, { opacity: 1 }], easing: 'ease-in-out', duration: dur };
    default:
      return out([{ opacity: 0 }, { opacity: 1 }]);
  }
}

/** Keyframes for the outgoing and incoming slide layers. */
export function transitionKeyframes(t: Transition): { out?: Keyframe[]; in?: Keyframe[]; easing: string; outOnTop?: boolean } {
  const dir = t.dir;
  const vec = (d: Dir | undefined, flip = false) => {
    const s = flip ? -1 : 1;
    switch (d) {
      case 'r':
        return [`${-100 * s}%`, '0%'];
      case 'u':
        return ['0%', `${100 * s}%`];
      case 'd':
        return ['0%', `${-100 * s}%`];
      default:
        return [`${100 * s}%`, '0%'];
    }
  };
  switch (t.type) {
    case 'fade':
      return { in: [{ opacity: 0 }, { opacity: 1 }], easing: 'ease' };
    case 'dissolve':
      return { in: [{ opacity: 0, filter: 'blur(8px)' }, { opacity: 1, filter: 'blur(0px)' }], easing: 'ease' };
    case 'push': {
      const [x, y] = vec(dir);
      const [ox, oy] = vec(dir, true);
      return { in: [{ transform: `translate(${x}, ${y})` }, { transform: 'translate(0, 0)' }], out: [{ transform: 'translate(0, 0)' }, { transform: `translate(${ox}, ${oy})` }], easing: 'cubic-bezier(.4,0,.2,1)' };
    }
    case 'cover': {
      const [x, y] = vec(dir);
      return { in: [{ transform: `translate(${x}, ${y})` }, { transform: 'translate(0, 0)' }], easing: 'cubic-bezier(.4,0,.2,1)' };
    }
    case 'reveal': {
      const [ox, oy] = vec(dir, true);
      return { out: [{ transform: 'translate(0, 0)', opacity: 1 }, { transform: `translate(${ox}, ${oy})`, opacity: 0.4 }], easing: 'cubic-bezier(.4,0,.2,1)', outOnTop: true };
    }
    case 'wipe': {
      const from = dir === 'r' ? 'inset(0 100% 0 0)' : dir === 'u' ? 'inset(100% 0 0 0)' : dir === 'd' ? 'inset(0 0 100% 0)' : 'inset(0 0 0 100%)';
      return { in: [{ clipPath: from }, { clipPath: 'inset(0 0 0 0)' }], easing: 'ease-in-out' };
    }
    case 'split':
      return { in: [{ clipPath: dir === 'u' ? 'inset(50% 0 50% 0)' : 'inset(0 50% 0 50%)' }, { clipPath: 'inset(0 0 0 0)' }], easing: 'ease-in-out' };
    case 'circle':
      return { in: [{ clipPath: 'circle(0% at 50% 50%)' }, { clipPath: 'circle(75% at 50% 50%)' }], easing: 'ease-in-out' };
    case 'zoom':
      return { in: [{ opacity: 0, transform: 'scale(0.4)' }, { opacity: 1, transform: 'scale(1)' }], out: [{ opacity: 1, transform: 'scale(1)' }, { opacity: 0, transform: 'scale(1.3)' }], easing: 'cubic-bezier(.2,.7,.2,1)' };
    case 'flip':
      return {
        out: [{ transform: 'perspective(1600px) rotateY(0deg)', opacity: 1, offset: 0 }, { transform: `perspective(1600px) rotateY(${dir === 'r' ? '' : '-'}90deg)`, opacity: 1, offset: 0.5 }, { opacity: 0, offset: 0.51 }, { opacity: 0, offset: 1 }],
        in: [{ opacity: 0, offset: 0 }, { opacity: 0, transform: `perspective(1600px) rotateY(${dir === 'r' ? '-' : ''}90deg)`, offset: 0.5 }, { opacity: 1, transform: `perspective(1600px) rotateY(${dir === 'r' ? '-' : ''}90deg)`, offset: 0.51 }, { opacity: 1, transform: 'perspective(1600px) rotateY(0deg)', offset: 1 }],
        easing: 'ease-in-out',
      };
    case 'morph':
      return { out: [{ opacity: 1 }, { opacity: 0 }], easing: 'cubic-bezier(.4,0,.2,1)' };
    default:
      return { easing: 'linear' };
  }
}

/** Visible slide indices in show order (hidden slides skipped unless it's the start slide). */
export function showOrder(pres: Presentation, start: number): number[] {
  return pres.slides.map((s, i) => (s.hidden && i !== start ? -1 : i)).filter((i) => i >= 0);
}
