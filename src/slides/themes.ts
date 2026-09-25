/**
 * Built-in designs (theme colours, fonts, backgrounds and decorative shapes) and the standard slide layouts.
 * Placeholder frames follow PowerPoint's default Office layouts so imported and exported decks line up.
 */
import {
  DEFAULT_H,
  DEFAULT_W,
  newId,
  type El,
  type Fill,
  type Layout,
  type LayoutType,
  type Para,
  type PlaceholderType,
  type Presentation,
  type ShapeEl,
  type Slide,
  type TextBody,
  type TextDefaults,
  type Theme,
  type ThemeColorName,
} from './model';

export interface Design {
  id: string;
  name: string;
  colors: Record<ThemeColorName, string>;
  fonts: { major: string; minor: string };
  background: Fill;
  /** Title slide background (defaults to background). */
  titleBackground?: Fill;
  text: {
    titleSize?: number;
    ctrTitleSize?: number;
    titleBold?: boolean;
    titleColor?: string;
    subtitleColor?: string;
    bodyColor?: string;
    titleCaps?: boolean;
    /** Left-aligns the title slide (for designs whose decoration sits on the right). */
    titleLeft?: boolean;
    /** Moves the title slide's subtitle down to make room for a rule between title and subtitle (px). */
    subtitleGap?: number;
  };
  decor: { title: El[]; section: El[]; content: El[] };
}

/* ------------------------------------------------------------- builders */

function shape(geom: string, x: number, y: number, w: number, h: number, fill: Fill, extra: Partial<ShapeEl> = {}): ShapeEl {
  return { id: newId('d'), type: 'shape', geom, x, y, w, h, fill, line: null, ...extra };
}
const solid = (color: string): Fill => ({ type: 'solid', color });
const radial = (inner: string, outer: string): Fill => ({ type: 'gradient', radial: true, angle: 0, stops: [{ pos: 0, color: inner }, { pos: 1, color: outer }] });
const linear = (angle: number, a: string, b: string, c?: string): Fill => ({
  type: 'gradient',
  angle,
  stops: c ? [{ pos: 0, color: a }, { pos: 0.5, color: b }, { pos: 1, color: c }] : [{ pos: 0, color: a }, { pos: 1, color: b }],
});

/* ------------------------------------------------------------- designs */

export const DESIGNS: Design[] = [
  {
    id: 'office',
    name: 'Office',
    colors: { bg1: '#FFFFFF', tx1: '#000000', bg2: '#E8E8E8', tx2: '#0E2841', accent1: '#156082', accent2: '#E97132', accent3: '#196B24', accent4: '#0F9ED5', accent5: '#A02B93', accent6: '#4EA72E', hlink: '#467886', folHlink: '#96607D' },
    fonts: { major: 'Calibri', minor: 'Calibri' },
    background: solid('@bg1'),
    text: {},
    decor: { title: [], section: [], content: [] },
  },
  {
    id: 'affice',
    name: 'Affice',
    colors: { bg1: '#FFFFFF', tx1: '#1B1F2A', bg2: '#F2F5FB', tx2: '#243B6B', accent1: '#2F6DFF', accent2: '#17A35A', accent3: '#F26A26', accent4: '#8E4EC6', accent5: '#12A594', accent6: '#E5484D', hlink: '#2F6DFF', folHlink: '#8E4EC6' },
    fonts: { major: 'Poppins', minor: 'Open Sans' },
    background: solid('@bg1'),
    text: { titleBold: true, subtitleColor: '@tx1+35', titleLeft: true, subtitleGap: 22 },
    decor: {
      title: [shape('ellipse', 860, -260, 680, 680, solid('@accent1/12')), shape('ellipse', 1010, 420, 360, 360, solid('@accent2/14')), shape('rect', 170, 384, 90, 6, solid('@accent1'))],
      section: [shape('rect', 0, 0, 22, 720, solid('@accent1')), shape('ellipse', 1040, -140, 380, 380, solid('@accent1/10'))],
      content: [shape('rect', 88, 172, 72, 5, solid('@accent1'))],
    },
  },
  {
    id: 'pitch',
    name: 'Pitch',
    colors: { bg1: '#0B1B4D', tx1: '#FFFFFF', bg2: '#13266B', tx2: '#C7D2FE', accent1: '#3B82F6', accent2: '#22D3EE', accent3: '#A78BFA', accent4: '#F472B6', accent5: '#FBBF24', accent6: '#34D399', hlink: '#93C5FD', folHlink: '#C4B5FD' },
    fonts: { major: 'Montserrat', minor: 'Open Sans' },
    background: linear(135, '#0B1B4D', '#07102E'),
    text: { titleBold: true, subtitleColor: '@tx2', titleLeft: true },
    decor: {
      title: [shape('ellipse', 700, -300, 820, 820, radial('@accent1/55', '@accent1/0')), shape('ellipse', -200, 440, 520, 520, radial('@accent3/35', '@accent3/0')), shape('roundRect', 170, 560, 150, 34, solid('@accent1'), { adj: { adj: 50000 } })],
      section: [shape('ellipse', 760, -200, 700, 700, radial('@accent2/40', '@accent2/0'))],
      content: [shape('ellipse', 980, -220, 520, 520, radial('@accent1/40', '@accent1/0')), shape('rect', 88, 172, 56, 4, solid('@accent2'))],
    },
  },
  {
    id: 'lesson',
    name: 'Lesson',
    colors: { bg1: '#F3FAF5', tx1: '#14532D', bg2: '#DCF2E3', tx2: '#166534', accent1: '#17A35A', accent2: '#F59E0B', accent3: '#0EA5E9', accent4: '#EF4444', accent5: '#8B5CF6', accent6: '#14B8A6', hlink: '#0E7490', folHlink: '#6D28D9' },
    fonts: { major: 'Nunito', minor: 'Nunito' },
    background: solid('@bg1'),
    text: { titleBold: true, subtitleColor: '@tx2' },
    decor: {
      title: [shape('ellipse', 860, 320, 620, 560, solid('@accent1/16')), shape('ellipse', -120, -140, 360, 320, solid('@accent2/22')), shape('ellipse', 1060, 80, 90, 90, solid('@accent3/30'))],
      section: [shape('ellipse', 900, 300, 560, 520, solid('@accent1/16'))],
      content: [shape('ellipse', 1120, -90, 260, 240, solid('@accent2/22')), shape('roundRect', 88, 168, 64, 8, solid('@accent1'), { adj: { adj: 50000 } })],
    },
  },
  {
    id: 'minimal',
    name: 'Minimal',
    colors: { bg1: '#FFFFFF', tx1: '#111111', bg2: '#F5F5F5', tx2: '#444444', accent1: '#111111', accent2: '#8C8C8C', accent3: '#C8C8C8', accent4: '#525252', accent5: '#A3A3A3', accent6: '#737373', hlink: '#111111', folHlink: '#525252' },
    fonts: { major: 'Playfair Display', minor: 'Lato' },
    background: solid('@bg1'),
    text: { subtitleColor: '@tx1+45', subtitleGap: 22 },
    decor: {
      title: [shape('rect', 590, 386, 100, 2, solid('@tx1'))],
      section: [shape('rect', 88, 470, 100, 2, solid('@tx1'))],
      content: [shape('rect', 88, 170, 48, 2, solid('@tx1'))],
    },
  },
  {
    id: 'midnight',
    name: 'Midnight',
    colors: { bg1: '#0F0F1A', tx1: '#F5F3FF', bg2: '#1C1B2E', tx2: '#C4B5FD', accent1: '#A855F7', accent2: '#3B82F6', accent3: '#EC4899', accent4: '#22D3EE', accent5: '#F59E0B', accent6: '#10B981', hlink: '#C4B5FD', folHlink: '#F0ABFC' },
    fonts: { major: 'Raleway', minor: 'Lato' },
    background: solid('@bg1'),
    text: { titleBold: true, subtitleColor: '@tx2' },
    decor: {
      title: [shape('ellipse', 760, -180, 640, 640, radial('@accent1/70', '@accent2/0')), shape('ellipse', 980, 360, 420, 420, radial('@accent3/45', '@accent3/0'))],
      section: [shape('ellipse', 840, 180, 560, 560, radial('@accent1/60', '@accent2/0'))],
      content: [shape('ellipse', 1030, 470, 420, 420, radial('@accent1/45', '@accent2/0'))],
    },
  },
  {
    id: 'sunset',
    name: 'Sunset',
    colors: { bg1: '#FFF4EC', tx1: '#7C2D12', bg2: '#FFE4D2', tx2: '#C2410C', accent1: '#F26A26', accent2: '#F59E0B', accent3: '#E11D48', accent4: '#A16207', accent5: '#FB7185', accent6: '#D97706', hlink: '#C2410C', folHlink: '#9A3412' },
    fonts: { major: 'Poppins', minor: 'Nunito' },
    background: linear(90, '#FFF4EC', '#FFE1CC'),
    text: { titleBold: true, subtitleColor: '@tx2', titleLeft: true },
    decor: {
      title: [shape('ellipse', 860, 150, 420, 420, linear(90, '@accent2', '@accent1')), shape('rect', 0, 600, 1280, 120, solid('@accent1/12'))],
      section: [shape('ellipse', 980, 420, 360, 360, linear(90, '@accent2', '@accent1'))],
      content: [shape('ellipse', 1150, -70, 200, 200, linear(90, '@accent2', '@accent1'))],
    },
  },
  {
    id: 'ocean',
    name: 'Ocean',
    colors: { bg1: '#FFFFFF', tx1: '#0B3B4F', bg2: '#E6F4F8', tx2: '#0E7490', accent1: '#0891B2', accent2: '#0EA5E9', accent3: '#14B8A6', accent4: '#6366F1', accent5: '#F59E0B', accent6: '#84CC16', hlink: '#0E7490', folHlink: '#4F46E5' },
    fonts: { major: 'Montserrat', minor: 'Lato' },
    background: solid('@bg1'),
    text: { titleBold: true, subtitleColor: '@tx2' },
    decor: {
      title: [shape('wave', -40, 560, 1360, 200, linear(0, '@accent1', '@accent2'), { adj: { adj1: 12500, adj2: 0 } }), shape('ellipse', 1060, 60, 140, 140, solid('@accent2/18'))],
      section: [shape('wave', -40, 600, 1360, 160, linear(0, '@accent1', '@accent3'), { adj: { adj1: 12500, adj2: 0 } })],
      content: [shape('rect', 0, 706, 1280, 14, linear(0, '@accent1', '@accent2'))],
    },
  },
  {
    id: 'slate',
    name: 'Slate',
    colors: { bg1: '#1F2937', tx1: '#F9FAFB', bg2: '#374151', tx2: '#FCD34D', accent1: '#F59E0B', accent2: '#10B981', accent3: '#3B82F6', accent4: '#EF4444', accent5: '#A78BFA', accent6: '#EC4899', hlink: '#FCD34D', folHlink: '#FDBA74' },
    fonts: { major: 'Oswald', minor: 'Roboto' },
    background: solid('@bg1'),
    text: { titleCaps: true, subtitleColor: '@tx2', titleLeft: true, subtitleGap: 22 },
    decor: {
      title: [shape('rect', 0, 0, 24, 720, solid('@accent1')), shape('rect', 170, 384, 120, 6, solid('@accent1'))],
      section: [shape('rect', 0, 0, 24, 720, solid('@accent1'))],
      content: [shape('rect', 0, 0, 12, 720, solid('@accent1'))],
    },
  },
  {
    id: 'paper',
    name: 'Paper',
    colors: { bg1: '#FAF7F0', tx1: '#2B2118', bg2: '#F0E9DA', tx2: '#7C2D12', accent1: '#B45309', accent2: '#7C2D12', accent3: '#4D7C0F', accent4: '#1D4ED8', accent5: '#9D174D', accent6: '#57534E', hlink: '#9A3412', folHlink: '#7C2D12' },
    fonts: { major: 'Merriweather', minor: 'Lora' },
    background: solid('@bg1'),
    text: { subtitleColor: '@tx2' },
    decor: {
      title: [shape('rect', 60, 48, 1160, 2, solid('@tx1/60')), shape('rect', 60, 54, 1160, 1, solid('@tx1/40')), shape('rect', 60, 666, 1160, 1, solid('@tx1/40')), shape('rect', 60, 670, 1160, 2, solid('@tx1/60'))],
      section: [shape('rect', 60, 48, 1160, 2, solid('@tx1/60')), shape('rect', 60, 670, 1160, 2, solid('@tx1/60'))],
      content: [shape('rect', 60, 30, 1160, 1, solid('@tx1/40'))],
    },
  },
  {
    id: 'berry',
    name: 'Berry',
    colors: { bg1: '#FFFFFF', tx1: '#3B0764', bg2: '#FCE7F3', tx2: '#9D174D', accent1: '#DB2777', accent2: '#9333EA', accent3: '#F97316', accent4: '#0EA5E9', accent5: '#10B981', accent6: '#EAB308', hlink: '#9333EA', folHlink: '#DB2777' },
    fonts: { major: 'Poppins', minor: 'Poppins' },
    background: linear(120, '#FDF2F8', '#F5F3FF'),
    text: { titleBold: true, subtitleColor: '@tx2' },
    decor: {
      title: [shape('parallelogram', 820, -40, 620, 800, linear(90, '@accent1', '@accent2'), { adj: { adj: 40000 } })],
      section: [shape('parallelogram', 1000, -40, 420, 800, linear(90, '@accent1', '@accent2'), { adj: { adj: 40000 } })],
      content: [shape('rect', 88, 172, 60, 6, linear(0, '@accent1', '@accent2'))],
    },
  },
  {
    id: 'forest',
    name: 'Forest',
    colors: { bg1: '#0F2E1F', tx1: '#ECFDF5', bg2: '#14532D', tx2: '#BEF264', accent1: '#A3E635', accent2: '#34D399', accent3: '#FDE047', accent4: '#38BDF8', accent5: '#FB923C', accent6: '#F472B6', hlink: '#BEF264', folHlink: '#6EE7B7' },
    fonts: { major: 'Montserrat', minor: 'Open Sans' },
    background: solid('@bg1'),
    text: { titleBold: true, subtitleColor: '@tx2' },
    decor: {
      title: [shape('teardrop', 1000, 380, 300, 300, solid('@accent2/25'), { rot: 180 }), shape('teardrop', 1080, 470, 200, 200, solid('@accent1/35'), { rot: 180 })],
      section: [shape('teardrop', 1060, 440, 240, 240, solid('@accent2/25'), { rot: 180 })],
      content: [shape('rect', 88, 172, 56, 4, solid('@accent1'))],
    },
  },
];

export const designById = (id: string | undefined) => DESIGNS.find((d) => d.id === id);

/* ------------------------------------------------------ colour palettes */

export const PALETTES: { name: string; colors: Partial<Record<ThemeColorName, string>> }[] = [
  { name: 'Office', colors: { accent1: '#156082', accent2: '#E97132', accent3: '#196B24', accent4: '#0F9ED5', accent5: '#A02B93', accent6: '#4EA72E' } },
  { name: 'Office 2013', colors: { accent1: '#4472C4', accent2: '#ED7D31', accent3: '#A5A5A5', accent4: '#FFC000', accent5: '#5B9BD5', accent6: '#70AD47' } },
  { name: 'Blue', colors: { accent1: '#0F6FC6', accent2: '#009DD9', accent3: '#0BD0D9', accent4: '#10CF9B', accent5: '#7CCA62', accent6: '#A5C249' } },
  { name: 'Blue Green', colors: { accent1: '#3494BA', accent2: '#58B6C0', accent3: '#75BDA7', accent4: '#7A8C8E', accent5: '#84ACB6', accent6: '#2683C6' } },
  { name: 'Green', colors: { accent1: '#549E39', accent2: '#8AB833', accent3: '#C0CF3A', accent4: '#029676', accent5: '#4AB5C4', accent6: '#0989B1' } },
  { name: 'Yellow Orange', colors: { accent1: '#F0A22E', accent2: '#A5644E', accent3: '#B58B80', accent4: '#C3986D', accent5: '#A19574', accent6: '#C17529' } },
  { name: 'Orange Red', colors: { accent1: '#D34817', accent2: '#9B2D1F', accent3: '#A28E6A', accent4: '#956251', accent5: '#918485', accent6: '#855D5D' } },
  { name: 'Red Violet', colors: { accent1: '#E32D91', accent2: '#C830CC', accent3: '#4EA6DC', accent4: '#4775E7', accent5: '#8971E1', accent6: '#D54773' } },
  { name: 'Violet', colors: { accent1: '#AD84C6', accent2: '#8784C7', accent3: '#5D739A', accent4: '#6997AF', accent5: '#84ACB6', accent6: '#6F8183' } },
  { name: 'Marquee', colors: { accent1: '#418AB3', accent2: '#A6B727', accent3: '#F69200', accent4: '#838383', accent5: '#FEC306', accent6: '#DF5327' } },
  { name: 'Grayscale', colors: { accent1: '#DDDDDD', accent2: '#B2B2B2', accent3: '#969696', accent4: '#808080', accent5: '#5F5F5F', accent6: '#4D4D4D' } },
  { name: 'Vivid', colors: { accent1: '#2F6DFF', accent2: '#17A35A', accent3: '#F26A26', accent4: '#8E4EC6', accent5: '#12A594', accent6: '#E5484D' } },
];

export const FONT_PAIRS: { name: string; major: string; minor: string }[] = [
  { name: 'Office', major: 'Calibri', minor: 'Calibri' },
  { name: 'Classic', major: 'Cambria', minor: 'Calibri' },
  { name: 'Arial', major: 'Arial', minor: 'Arial' },
  { name: 'Montserrat · Open Sans', major: 'Montserrat', minor: 'Open Sans' },
  { name: 'Poppins · Open Sans', major: 'Poppins', minor: 'Open Sans' },
  { name: 'Playfair · Lato', major: 'Playfair Display', minor: 'Lato' },
  { name: 'Merriweather · Lora', major: 'Merriweather', minor: 'Lora' },
  { name: 'Oswald · Roboto', major: 'Oswald', minor: 'Roboto' },
  { name: 'Raleway · Lato', major: 'Raleway', minor: 'Lato' },
  { name: 'Nunito', major: 'Nunito', minor: 'Nunito' },
  { name: 'Bebas Neue · Montserrat', major: 'Bebas Neue', minor: 'Montserrat' },
  { name: 'EB Garamond', major: 'EB Garamond', minor: 'EB Garamond' },
];

/* ------------------------------------------------------------- layouts */

interface PhSpec {
  ph: PlaceholderType;
  idx?: number;
  x: number;
  y: number;
  w: number;
  h: number;
  anchor: 't' | 'm' | 'b';
  align?: Para['align'];
  size?: number;
  bold?: boolean;
  color?: string;
  noBullets?: boolean;
  name: string;
}

const L = (type: LayoutType, name: string, phs: PhSpec[]) => ({ type, name, phs });

/** PowerPoint's default 16:9 layouts (EMU positions converted to px). */
export const LAYOUT_SPECS: { type: LayoutType; name: string; phs: PhSpec[] }[] = [
  L('title', 'Title Slide', [
    { ph: 'ctrTitle', x: 160, y: 118, w: 960, h: 251, anchor: 'b', align: 'center', name: 'Title' },
    { ph: 'subTitle', idx: 1, x: 160, y: 378, w: 960, h: 174, anchor: 't', align: 'center', size: 24, name: 'Subtitle' },
  ]),
  L('obj', 'Title and Content', [
    { ph: 'title', x: 88, y: 38, w: 1104, h: 139, anchor: 'm', name: 'Title' },
    { ph: 'obj', idx: 1, x: 88, y: 192, w: 1104, h: 457, anchor: 't', name: 'Content' },
  ]),
  L('secHead', 'Section Header', [
    { ph: 'title', x: 87, y: 180, w: 1104, h: 300, anchor: 'b', size: 60, name: 'Title' },
    { ph: 'body', idx: 1, x: 87, y: 482, w: 1104, h: 158, anchor: 't', size: 24, noBullets: true, color: '@tx1+30', name: 'Text' },
  ]),
  L('twoObj', 'Two Content', [
    { ph: 'title', x: 88, y: 38, w: 1104, h: 139, anchor: 'm', name: 'Title' },
    { ph: 'obj', idx: 1, x: 88, y: 192, w: 544, h: 457, anchor: 't', name: 'Content left' },
    { ph: 'obj', idx: 2, x: 648, y: 192, w: 544, h: 457, anchor: 't', name: 'Content right' },
  ]),
  L('twoTxTwoObj', 'Comparison', [
    { ph: 'title', x: 88, y: 38, w: 1104, h: 139, anchor: 'm', name: 'Title' },
    { ph: 'body', idx: 1, x: 88, y: 177, w: 542, h: 87, anchor: 'b', size: 24, bold: true, noBullets: true, name: 'Heading left' },
    { ph: 'obj', idx: 2, x: 88, y: 263, w: 542, h: 387, anchor: 't', name: 'Content left' },
    { ph: 'body', idx: 3, x: 648, y: 177, w: 544, h: 87, anchor: 'b', size: 24, bold: true, noBullets: true, name: 'Heading right' },
    { ph: 'obj', idx: 4, x: 648, y: 263, w: 544, h: 387, anchor: 't', name: 'Content right' },
  ]),
  L('titleOnly', 'Title Only', [{ ph: 'title', x: 88, y: 38, w: 1104, h: 139, anchor: 'm', name: 'Title' }]),
  L('blank', 'Blank', []),
  L('objTx', 'Content with Caption', [
    { ph: 'title', x: 88, y: 48, w: 413, h: 168, anchor: 'b', size: 32, name: 'Title' },
    { ph: 'obj', idx: 1, x: 544, y: 104, w: 648, h: 512, anchor: 't', name: 'Content' },
    { ph: 'body', idx: 2, x: 88, y: 216, w: 413, h: 400, anchor: 't', size: 16, noBullets: true, name: 'Caption' },
  ]),
  L('picTx', 'Picture with Caption', [
    { ph: 'title', x: 88, y: 48, w: 413, h: 168, anchor: 'b', size: 32, name: 'Title' },
    { ph: 'pic', idx: 1, x: 544, y: 104, w: 648, h: 512, anchor: 't', name: 'Picture' },
    { ph: 'body', idx: 2, x: 88, y: 216, w: 413, h: 400, anchor: 't', size: 16, noBullets: true, name: 'Caption' },
  ]),
];

export const PROMPTS: Record<PlaceholderType, string> = {
  title: 'Click to add title',
  ctrTitle: 'Click to add title',
  subTitle: 'Click to add subtitle',
  body: 'Click to add text',
  obj: 'Click to add text',
  pic: 'Click the icon to add a picture',
  dt: 'Date',
  ftr: 'Footer',
  sldNum: '#',
};

function placeholderEl(spec: PhSpec, sx: number, design: Design): ShapeEl {
  const left = design.text.titleLeft && (spec.ph === 'ctrTitle' || spec.ph === 'subTitle');
  const para: Para = { runs: [], align: left ? undefined : spec.align };
  if (spec.noBullets) para.bullet = { type: 'none' };
  const defaults: TextBody['defaults'] = {};
  if (spec.size) defaults.size = spec.size;
  if (spec.bold) defaults.b = true;
  if (spec.color) defaults.color = spec.color;
  if (spec.ph === 'subTitle' && design.text.subtitleColor) defaults.color = design.text.subtitleColor;
  const text: TextBody = { paras: [para], anchor: spec.anchor, autofit: 'shrink', defaults: Object.keys(defaults).length ? defaults : undefined };
  const gap = spec.ph === 'subTitle' ? (design.text.subtitleGap ?? 0) : 0;
  return { id: newId('p'), type: 'shape', geom: 'rect', name: spec.name, x: Math.round(spec.x * sx), y: spec.y + gap, w: Math.round(spec.w * sx), h: spec.h - gap, ph: spec.ph, phIdx: spec.idx, text };
}

function scaleEls(els: El[], sx: number, sy: number): El[] {
  return els.map((e) => {
    const c = structuredClone(e) as El;
    c.id = newId('d');
    c.x = Math.round(c.x * sx);
    c.y = Math.round(c.y * sy);
    c.w = Math.round(c.w * sx);
    c.h = Math.round(c.h * sy);
    return c;
  });
}

export function buildLayouts(design: Design, size: { w: number; h: number }): Layout[] {
  const sx = size.w / DEFAULT_W;
  const sy = size.h / DEFAULT_H;
  return LAYOUT_SPECS.map((spec) => {
    const decor = spec.type === 'title' ? design.decor.title : spec.type === 'secHead' ? design.decor.section : spec.type === 'blank' ? [] : design.decor.content;
    return {
      id: `layout-${spec.type}`,
      name: spec.name,
      type: spec.type,
      background: spec.type === 'title' ? design.titleBackground : undefined,
      decor: scaleEls(decor, sx, sy),
      placeholders: spec.phs.map((p) => {
        const el = placeholderEl(p, sx, design);
        el.y = Math.round(el.y * sy);
        el.h = Math.round(el.h * sy);
        return el;
      }),
    };
  });
}

export function textDefaults(design: Design): TextDefaults {
  return {
    title: { font: '+major', size: design.text.titleSize ?? 44, color: design.text.titleColor ?? '@tx1', bold: design.text.titleBold },
    body: { font: '+minor', size: 28, color: design.text.bodyColor ?? '@tx1', lineSpacing: 0.9 },
    other: { font: '+minor', size: 18, color: '@tx1' },
  };
}

export function themeOf(design: Design): Theme {
  return { id: design.id, name: design.name, colors: { ...design.colors }, fonts: { ...design.fonts } };
}

export function createPresentation(designId = 'affice', size = { w: DEFAULT_W, h: DEFAULT_H }): Presentation {
  const design = designById(designId) ?? DESIGNS[1];
  const layouts = buildLayouts(design, size);
  const pres: Presentation = {
    format: 'affice-slides',
    version: 1,
    size,
    theme: themeOf(design),
    design: design.id,
    master: { background: design.background, decor: [], text: textDefaults(design) },
    layouts,
    slides: [],
    props: { created: new Date().toISOString() },
  };
  pres.slides.push(newSlide(pres, 'layout-title'));
  return pres;
}

/** A new slide with fresh, empty placeholders from a layout. */
export function newSlide(pres: Presentation, layoutId: string): Slide {
  const layout = pres.layouts.find((l) => l.id === layoutId) ?? pres.layouts.find((l) => l.type === 'obj') ?? pres.layouts[0];
  return {
    id: newId('s'),
    layout: layout?.id ?? layoutId,
    elements: (layout?.placeholders ?? []).map((p) => ({ ...structuredClone(p), id: newId() })),
  };
}

/** The layout PowerPoint picks for "New slide" after a given slide. */
export function nextLayoutFor(pres: Presentation, slide: Slide | undefined): string {
  const cur = slide ? pres.layouts.find((l) => l.id === slide.layout) : undefined;
  if (!cur || cur.type === 'title') return pres.layouts.find((l) => l.type === 'obj')?.id ?? pres.layouts[0].id;
  return cur.id;
}

/** Switches the design: theme, fonts, backgrounds and decorations. Slide content stays. */
export function applyDesign(pres: Presentation, designId: string): Presentation {
  const design = designById(designId);
  if (!design) return pres;
  const layouts = buildLayouts(design, pres.size);
  // keep custom (imported) layouts that have no built-in equivalent
  const byType = new Map(layouts.map((l) => [l.type, l]));
  const oldToNew = new Map<string, string>();
  for (const l of pres.layouts) {
    const n = byType.get(l.type) ?? byType.get('obj');
    if (n) oldToNew.set(l.id, n.id);
  }
  return {
    ...pres,
    design: design.id,
    theme: themeOf(design),
    master: { background: design.background, decor: [], text: textDefaults(design) },
    layouts,
    slides: pres.slides.map((s) => ({ ...s, layout: oldToNew.get(s.layout) ?? layouts[1].id })),
  };
}

export function applyPalette(pres: Presentation, colors: Partial<Record<ThemeColorName, string>>): Presentation {
  return { ...pres, theme: { ...pres.theme, colors: { ...pres.theme.colors, ...colors } } };
}

export function applyFonts(pres: Presentation, fonts: { major: string; minor: string }): Presentation {
  return { ...pres, theme: { ...pres.theme, fonts: { ...fonts } } };
}

/** Resolves "+major"/"+minor" theme font references. */
export function resolveFont(font: string | undefined, theme: Theme): string | undefined {
  if (!font) return undefined;
  if (font === '+major' || font === '+mj-lt') return theme.fonts.major;
  if (font === '+minor' || font === '+mn-lt') return theme.fonts.minor;
  return font;
}
