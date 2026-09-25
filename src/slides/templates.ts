import { newId, textFromPlain, type Anim, type AnimEffect, type El, type Fill, type GroupEl, type LayoutType, type Line, type Para, type Presentation, type ShapeEl, type Slide, type SlideChart, type TextBody, type Transition } from './model';
import { createPresentation, newSlide } from './themes';
import { createTable } from './tables';

type BodyLine = string | { text: string; level?: number; b?: boolean };

class Deck {
  pres: Presentation;
  constructor(design: string) {
    this.pres = createPresentation(design);
    this.pres.slides = [];
  }

  layout(type: LayoutType): string {
    return this.pres.layouts.find((l) => l.type === type)?.id ?? this.pres.layouts[0].id;
  }

  slide(type: LayoutType, content: { title?: string; subtitle?: string; body?: BodyLine[]; body2?: BodyLine[]; head1?: string; head2?: string; numbered?: boolean } = {}, extras: El[] = [], opts: { transition?: Transition; notes?: string; anims?: Omit<Anim, 'id'>[] } = {}): Slide {
    const s = newSlide(this.pres, this.layout(type));
    const bodies = s.elements.filter((e) => e.ph === 'body' || e.ph === 'obj');
    const setText = (e: El | undefined, lines: BodyLine[] | string | undefined, numbered = false) => {
      if (!e || e.type !== 'shape' || !e.text || lines === undefined) return;
      const arr = typeof lines === 'string' ? [lines] : lines;
      const template = e.text.paras[0] ?? { runs: [] };
      e.text = {
        ...e.text,
        paras: arr.map((l): Para => {
          const o = typeof l === 'string' ? { text: l } : l;
          return { ...template, runs: [{ text: o.text, ...(o.b ? { b: true } : {}) }], level: o.level, ...(numbered ? { bullet: { type: 'num', style: 'arabicPeriod' } } : {}) };
        }),
      };
    };
    setText(s.elements.find((e) => e.ph === 'title' || e.ph === 'ctrTitle'), content.title);
    setText(s.elements.find((e) => e.ph === 'subTitle'), content.subtitle);
    if (type === 'twoTxTwoObj') {
      setText(bodies[0], content.head1);
      setText(bodies[1], content.body);
      setText(bodies[2], content.head2);
      setText(bodies[3], content.body2);
    } else if (type === 'secHead') setText(bodies[0], content.subtitle);
    else {
      setText(bodies[0], content.body, content.numbered);
      setText(bodies[1], content.body2);
    }
    // drop placeholders left empty by the template
    s.elements = s.elements.filter((e) => !(e.ph && e.type === 'shape' && e.ph !== 'pic' && e.text && e.text.paras.every((p) => p.runs.every((r) => !r.text))));
    s.elements.push(...extras);
    if (opts.transition) s.transition = opts.transition;
    if (opts.notes) s.notes = opts.notes;
    if (opts.anims) s.anims = opts.anims.map((a) => ({ ...a, id: newId('a') }));
    this.pres.slides.push(s);
    return s;
  }
}

/* ------------------------------------------------------------- helpers */

const solid = (color: string): Fill => ({ type: 'solid', color });

function text(x: number, y: number, w: number, h: number, value: string, o: { size?: number; color?: string; b?: boolean; i?: boolean; align?: Para['align']; anchor?: TextBody['anchor']; font?: string; lineSpacing?: number } = {}): ShapeEl {
  return {
    id: newId(),
    type: 'shape',
    geom: 'rect',
    name: 'TextBox',
    x,
    y,
    w,
    h,
    textbox: true,
    text: {
      paras: textFromPlain(value, { ...(o.size ? { size: o.size } : {}), ...(o.color ? { color: o.color } : {}), ...(o.b ? { b: true } : {}), ...(o.i ? { i: true } : {}), ...(o.font ? { font: o.font } : {}) }, { align: o.align, lineSpacing: o.lineSpacing }),
      anchor: o.anchor ?? 't',
      wrap: true,
    },
  };
}

function box(geom: string, x: number, y: number, w: number, h: number, fill: Fill, o: { line?: Line | null; text?: string; size?: number; color?: string; b?: boolean; adj?: Record<string, number>; name?: string; anchor?: TextBody['anchor']; shadow?: boolean } = {}): ShapeEl {
  return {
    id: newId(),
    type: 'shape',
    geom,
    name: o.name ?? 'Shape',
    x,
    y,
    w,
    h,
    fill,
    line: o.line ?? null,
    adj: o.adj,
    shadow: o.shadow ? { color: '#000000/22', blur: 18, dist: 6, angle: 90 } : undefined,
    text: o.text !== undefined ? { paras: textFromPlain(o.text, { ...(o.size ? { size: o.size } : {}), ...(o.color ? { color: o.color } : {}), ...(o.b ? { b: true } : {}) }, { align: 'center' }), anchor: o.anchor ?? 'm' } : undefined,
  };
}

/** A big number with a caption underneath. */
function stat(x: number, y: number, value: string, caption: string): ShapeEl {
  const tb = text(x, y, 320, 100, value, { size: 30, b: true, font: '+major', color: '@accent2' });
  tb.text!.paras.push({ runs: [{ text: caption, size: 18, color: '@tx2' }] });
  return tb;
}

/** Groups elements so they move and animate together. */
function grp(children: El[], name: string): GroupEl {
  const x = Math.min(...children.map((c) => c.x));
  const y = Math.min(...children.map((c) => c.y));
  const w = Math.max(...children.map((c) => c.x + c.w)) - x;
  const h = Math.max(...children.map((c) => c.y + c.h)) - y;
  return { id: newId(), type: 'group', name, x, y, w, h, children };
}

function arrow(x1: number, y1: number, x2: number, y2: number, color = '@tx1+40', width = 3): ShapeEl {
  return { id: newId(), type: 'shape', geom: 'straightConnector1', name: 'Arrow', x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1), flipH: x2 < x1 || undefined, flipV: y2 < y1 || undefined, fill: { type: 'none' }, line: { color, width, tail: 'triangle' } };
}

function chart(x: number, y: number, w: number, h: number, c: SlideChart): El {
  return { id: newId(), type: 'chart', name: 'Chart', x, y, w, h, chart: c };
}

const anim = (el: El, effect: AnimEffect, start: Anim['start'] = 'click', extra: Partial<Anim> = {}): Omit<Anim, 'id'> => ({ el: el.id, cls: 'entr', effect, start, dur: 500, delay: 0, ...extra });

/* ============================================================== decks */

function pitch(): Presentation {
  const d = new Deck('pitch');
  d.slide('title', { title: 'Nova Robotics', subtitle: 'Warehouse automation for everyone · Seed round 2026' }, [], { transition: { type: 'fade', dur: 800 }, notes: 'Welcome everyone. Thanks for taking the time today.' });
  const probBody = d.slide('obj', { title: 'The problem', body: ['Pickers walk up to 15 km a day in a typical warehouse', 'Today’s robots cost over $250k and take months to set up', 'Small and mid-sized businesses are left behind'] }, [], { transition: { type: 'push', dir: 'u', dur: 700 } });
  const body = probBody.elements.find((e) => e.ph === 'obj');
  if (body) probBody.anims = [{ ...anim(body, 'fade'), id: newId('a'), byPara: true }];
  const cards = [
    ['Plug & play', 'Unbox, scan a QR code and start picking the same day.'],
    ['Learns by watching', 'Shows the robot a task once — it figures out the rest.'],
    ['Pay monthly', 'No upfront cost. Cancel any time.'],
  ];
  const cardEls: El[] = [];
  cards.forEach(([h, t], i) => {
    const x = 88 + i * 376;
    // heading and description share one text box so a wrapped heading pushes the description down
    const tb = text(x + 28, 348, 300, 230, h, { size: 26, b: true, font: '+major' });
    tb.text!.paras.push({ runs: [{ text: t, size: 18, color: '@tx2' }], spaceBefore: 8 });
    cardEls.push(
      grp(
        [
          box('roundRect', x, 220, 352, 380, solid('@bg2'), { adj: { adj: 8000 }, name: 'Card', line: { color: '@accent1/40', width: 1.5 } }),
          box('ellipse', x + 32, 256, 72, 72, solid('@accent1'), { text: String(i + 1), size: 26, b: true, color: '#FFFFFF' }),
          tb,
        ],
        `Card ${i + 1}`,
      ),
    );
  });
  const sol = d.slide('titleOnly', { title: 'Our solution' }, cardEls, { transition: { type: 'fade', dur: 600 } });
  sol.anims = cardEls.map((c, i) => ({ ...anim(c, 'float', i ? 'after' : 'click', { dir: 'u' as const }), id: newId('a') }));
  const circles = [
    box('ellipse', 380, 170, 520, 520, solid('@accent1/25'), { name: 'TAM' }),
    box('ellipse', 480, 330, 320, 320, solid('@accent1/45'), { name: 'SAM' }),
    box('ellipse', 545, 455, 190, 190, solid('@accent1'), { name: 'SOM', text: '$600M', size: 22, b: true, color: '#FFFFFF' }),
    stat(900, 220, '$48B', 'Total market'),
    stat(900, 370, '$9B', 'Serviceable market'),
    stat(900, 520, '$600M', 'Our target in 5 years'),
  ];
  const mk = d.slide('titleOnly', { title: 'Market opportunity' }, circles, { transition: { type: 'zoom', dur: 700 } });
  mk.anims = [0, 1, 2].map((i) => ({ ...anim(circles[i], 'zoom', i ? 'after' : 'click'), id: newId('a') }));
  d.slide(
    'titleOnly',
    { title: 'Traction' },
    [
      chart(88, 200, 760, 460, { type: 'column', title: 'Monthly revenue ($k)', categories: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'], series: [{ name: 'Revenue', values: [12, 19, 28, 41, 55, 78] }], legend: 'none', dataLabels: true }),
      box('roundRect', 890, 280, 300, 200, solid('@accent1'), { text: '6× growth\nin six months', size: 28, b: true, color: '#FFFFFF', adj: { adj: 10000 } }),
    ],
    { transition: { type: 'wipe', dir: 'l', dur: 700 } },
  );
  const team = [
    ['Ada Chen', 'CEO · ex-Amazon Robotics'],
    ['Leo Park', 'CTO · PhD, robotics'],
    ['Maya Singh', 'Head of Sales'],
    ['Tom Weber', 'Lead Engineer'],
  ];
  const teamEls: El[] = [];
  team.forEach(([n, r], i) => {
    const x = 120 + i * 270;
    teamEls.push(box('ellipse', x + 40, 220, 160, 160, { type: 'gradient', angle: 45, stops: [{ pos: 0, color: '@accent1' }, { pos: 1, color: '@accent3' }] }, { text: n.split(' ').map((w) => w[0]).join(''), size: 40, b: true, color: '#FFFFFF' }));
    teamEls.push(text(x, 400, 240, 40, n, { size: 22, b: true, align: 'center' }));
    teamEls.push(text(x, 440, 240, 60, r, { size: 16, color: '@tx2', align: 'center' }));
  });
  d.slide('titleOnly', { title: 'Team' }, teamEls, { transition: { type: 'push', dir: 'l', dur: 700 } });
  d.slide('secHead', { title: 'Let’s automate the other 90%', subtitle: 'Raising $2.5M · hello@nova.example' }, [], { transition: { type: 'fade', dur: 900 } });
  return d.pres;
}

function lesson(): Presentation {
  const d = new Deck('lesson');
  d.slide('title', { title: 'Photosynthesis', subtitle: 'How plants turn sunlight into food · Science, Year 7' }, [], { transition: { type: 'fade', dur: 700 } });
  d.slide('obj', { title: 'Today we will learn', body: ['What photosynthesis is', 'What plants need to make food', 'Where it happens inside a leaf', 'Why it matters for every living thing'], numbered: true }, [], { transition: { type: 'push', dir: 'l', dur: 600 } });
  const inputs = [
    box('roundRect', 88, 250, 220, 90, solid('@accent2'), { text: '☀ Sunlight', size: 22, b: true, color: '#FFFFFF', adj: { adj: 30000 } }),
    box('roundRect', 88, 370, 220, 90, solid('@accent3'), { text: '💧 Water', size: 22, b: true, color: '#FFFFFF', adj: { adj: 30000 } }),
    box('roundRect', 88, 490, 220, 90, solid('@tx2+30'), { text: 'CO₂', size: 22, b: true, color: '#FFFFFF', adj: { adj: 30000 } }),
  ];
  const leaf = box('ellipse', 480, 320, 300, 200, solid('@accent1'), { text: 'Leaf\n(chloroplasts)', size: 22, b: true, color: '#FFFFFF' });
  const outputs = [box('roundRect', 960, 300, 230, 90, solid('@accent6'), { text: 'Glucose', size: 22, b: true, color: '#FFFFFF', adj: { adj: 30000 } }), box('roundRect', 960, 450, 230, 90, solid('@accent5'), { text: 'Oxygen', size: 22, b: true, color: '#FFFFFF', adj: { adj: 30000 } })];
  const arrows = [arrow(318, 295, 470, 380), arrow(318, 415, 470, 420), arrow(318, 535, 470, 460), arrow(790, 390, 950, 345), arrow(790, 450, 950, 495)];
  const recipe = d.slide('titleOnly', { title: 'The recipe' }, [...inputs, leaf, ...arrows, ...outputs], { transition: { type: 'fade', dur: 600 }, notes: 'Ask the class: what do you think the plant does with the glucose?' });
  recipe.anims = [
    { ...anim(leaf, 'zoom'), id: newId('a') },
    ...outputs.map((o, i) => ({ ...anim(o, 'fly', i ? 'with' : 'click', { dir: 'r' as const }), id: newId('a') })),
  ];
  const t = createTable(5, 2, 88, 200, 1104, 64, 'accent1');
  const rows = [
    ['Term', 'Meaning'],
    ['Chlorophyll', 'The green pigment that captures light energy'],
    ['Stomata', 'Tiny pores that let gases in and out of the leaf'],
    ['Glucose', 'A sugar the plant uses for energy and growth'],
    ['Oxygen', 'The gas released into the air that we breathe'],
  ];
  t.cols = [300, 804];
  t.rows.forEach((r, i) => r.cells.forEach((c, j) => (c.text = { paras: textFromPlain(rows[i][j], { size: 20 }), anchor: 'm' })));
  d.slide('titleOnly', { title: 'Key words' }, [t], { transition: { type: 'wipe', dir: 'u', dur: 600 } });
  d.slide(
    'twoObj',
    { title: 'Quick quiz', body: ['What are the three things a plant needs?', 'Which gas do plants release?', 'Where in the leaf does it happen?'] },
    [box('wedgeRoundRectCallout', 700, 260, 440, 240, solid('@accent2+60'), { text: 'Talk with a partner\nfor 2 minutes!', size: 28, b: true, color: '@tx1', adj: { adj1: -30000, adj2: 70000, adj3: 16667 } })],
    { transition: { type: 'cover', dir: 'l', dur: 600 } },
  );
  d.slide('secHead', { title: 'Great work!', subtitle: 'Next lesson: respiration' }, [], { transition: { type: 'fade', dur: 800 } });
  return d.pres;
}

function minimal(): Presentation {
  const d = new Deck('minimal');
  d.slide('title', { title: 'Annual Report', subtitle: '2026 · Prepared for the board' }, [], { transition: { type: 'fade', dur: 1000 } });
  d.slide('obj', { title: 'Agenda', body: ['Highlights of the year', 'Financial results', 'What we learned', 'Plans for next year'], numbered: true }, [], { transition: { type: 'fade', dur: 800 } });
  const q = text(160, 220, 960, 200, '“Simplicity is the ultimate sophistication.”', { size: 48, i: true, align: 'center', font: '+major', anchor: 'm' });
  const who = text(160, 440, 960, 50, '— Leonardo da Vinci', { size: 20, align: 'center', color: '@tx2' });
  const qs = d.slide('blank', {}, [q, who], { transition: { type: 'fade', dur: 1000 } });
  qs.anims = [{ ...anim(who, 'fade', 'after', { delay: 300 }), id: newId('a') }];
  d.slide(
    'titleOnly',
    { title: 'Revenue' },
    [
      text(88, 230, 520, 170, '$4.2M', { size: 120, font: '+major', b: false }),
      text(96, 430, 500, 50, '+38% year over year', { size: 24, color: '@tx2' }),
      chart(640, 200, 560, 420, { type: 'line', categories: ['2022', '2023', '2024', '2025', '2026'], series: [{ name: 'Revenue ($M)', values: [1.4, 1.9, 2.4, 3.05, 4.2] }], legend: 'none', smooth: true }),
    ],
    { transition: { type: 'push', dir: 'u', dur: 700 } },
  );
  d.slide('twoTxTwoObj', { title: 'Then and now', head1: '2025', body: ['12 people', 'One product', 'Local customers'], head2: '2026', body2: ['31 people', 'Three products', 'Customers in 14 countries'] }, [], { transition: { type: 'fade', dur: 700 } });
  d.slide('title', { title: 'Thank you', subtitle: 'questions@company.example' }, [], { transition: { type: 'fade', dur: 1000 } });
  return d.pres;
}

function midnight(): Presentation {
  const d = new Deck('midnight');
  d.slide('title', { title: 'Introducing Aurora', subtitle: 'The assistant that respects your privacy' }, [], { transition: { type: 'fade', dur: 1000 } });
  d.slide('secHead', { title: 'One more thing…' }, [], { transition: { type: 'zoom', dur: 800 } });
  const feats = [
    ['Faster', 'Answers in under 100 ms, even offline.'],
    ['Smarter', 'Understands context across your apps.'],
    ['Private', 'Everything stays on your device.'],
  ];
  const els: El[] = [];
  feats.forEach(([h, t], i) => {
    const x = 88 + i * 376;
    const tb = text(x + 32, 260, 290, 280, h, { size: 34, b: true, font: '+major' });
    tb.text!.paras.push({ runs: [{ text: t, size: 20, color: '@tx2' }], spaceBefore: 14 });
    els.push(
      grp(
        [box('roundRect', x, 220, 352, 360, { type: 'gradient', angle: 135, stops: [{ pos: 0, color: '@accent1/35' }, { pos: 1, color: '@accent2/15' }] }, { adj: { adj: 8000 }, line: { color: '@accent1/60', width: 1.5 }, name: 'Card' }), tb],
        `Feature ${i + 1}`,
      ),
    );
  });
  const fs = d.slide('titleOnly', { title: 'What’s new' }, els, { transition: { type: 'morph', dur: 900 } });
  fs.anims = els.map((c, i) => ({ ...anim(c, 'float', i ? 'after' : 'click', { dir: 'u' as const }), id: newId('a') }));
  const steps = [
    ['Q1', 'Private beta'],
    ['Q2', 'Public launch'],
    ['Q3', 'Teams & sharing'],
    ['Q4', 'Aurora for desktop'],
  ];
  const tl: El[] = [box('rect', 120, 408, 1040, 4, solid('@tx1/30'), { name: 'Timeline' })];
  steps.forEach(([q, t], i) => {
    const x = 150 + i * 320;
    tl.push(grp([box('ellipse', x - 10, 370, 80, 80, solid(`@accent${i + 1}`), { text: q, size: 16, b: true, color: '#FFFFFF' }), text(x - 70, 470, 200, 80, t, { size: 20, align: 'center' })], q));
  });
  const road = d.slide('titleOnly', { title: 'Roadmap' }, tl, { transition: { type: 'push', dir: 'l', dur: 700 } });
  road.anims = steps.map((_, i) => ({ ...anim(tl[1 + i], 'zoom', i ? 'after' : 'click'), id: newId('a') }));
  d.slide(
    'titleOnly',
    { title: 'Performance' },
    [chart(88, 190, 1104, 470, { type: 'line', categories: ['1k', '10k', '100k', '1M', '10M'], series: [{ name: 'Aurora', values: [18, 22, 30, 41, 60] }, { name: 'Others', values: [45, 80, 160, 310, 700] }], title: 'Response time (ms) by document size', smooth: true })],
    { transition: { type: 'fade', dur: 700 } },
  );
  d.slide('title', { title: 'Available today', subtitle: 'aurora.example' }, [], { transition: { type: 'zoom', dur: 900 } });
  return d.pres;
}

function sunset(): Presentation {
  const d = new Deck('sunset');
  d.slide('title', { title: 'Our Story', subtitle: 'Ten years of making things by hand' }, [], { transition: { type: 'fade', dur: 900 } });
  d.slide('obj', { title: 'Chapters', body: ['A kitchen table in 2016', 'Our first shop', 'Growing without losing our soul', 'What comes next'] }, [], { transition: { type: 'wipe', dir: 'l', dur: 700 } });
  d.slide('picTx', { title: 'Where it began', body: ['Two friends, one kitchen table and a box of tools. Everything we make still starts the same way.'] }, [], { transition: { type: 'fade', dur: 700 } });
  const q = text(140, 200, 1000, 240, 'Make it with love,\nand people will feel it.', { size: 54, b: true, align: 'center', font: '+major', anchor: 'm', color: '@tx1' });
  const who = text(140, 470, 1000, 50, '— Our founder', { size: 22, align: 'center', color: '@tx2' });
  const qs = d.slide('blank', {}, [q, who], { transition: { type: 'fade', dur: 1000 } });
  qs.anims = [{ ...anim(q, 'fade'), id: newId('a') }, { ...anim(who, 'fade', 'after'), id: newId('a') }];
  const nums = [
    ['10', 'years'],
    ['38k', 'happy customers'],
    ['1', 'kitchen table'],
  ];
  const ne: El[] = [];
  nums.forEach(([n, l], i) => {
    const x = 128 + i * 360;
    ne.push(box('ellipse', x + 30, 210, 260, 260, { type: 'gradient', angle: 90, stops: [{ pos: 0, color: '@accent2' }, { pos: 1, color: '@accent1' }] }, { text: n, size: 60, b: true, color: '#FFFFFF' }));
    ne.push(text(x, 490, 320, 60, l, { size: 24, align: 'center' }));
  });
  const ns = d.slide('titleOnly', { title: 'By the numbers' }, ne, { transition: { type: 'push', dir: 'u', dur: 700 } });
  ns.anims = [0, 1, 2].map((i) => ({ ...anim(ne[i * 2], 'bounce', i ? 'after' : 'click'), id: newId('a') }));
  d.slide('secHead', { title: 'Thank you', subtitle: 'Let’s create something together' }, [], { transition: { type: 'fade', dur: 900 } });
  return d.pres;
}

export function slideTemplate(id: string): Presentation {
  switch (id) {
    case 'slides-pitch':
      return pitch();
    case 'slides-lesson':
      return lesson();
    case 'slides-minimal':
      return minimal();
    case 'slides-midnight':
      return midnight();
    case 'slides-sunset':
      return sunset();
    default:
      return createPresentation('affice');
  }
}

