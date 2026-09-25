import { beforeAll, describe, expect, it } from 'vitest';
import { DOMParser } from '@xmldom/xmldom';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { exportOdp } from '../src/slides/formats/odp-export';
import { importOdp } from '../src/slides/formats/odp-import';
import { arcToCubics, enhancedPath, evaluateEnhancedGeometry, svgPathSegments } from '../src/slides/formats/odp-geometry';
import { slideTemplate } from '../src/slides/templates';
import { createPresentation, newSlide } from '../src/slides/themes';
import { plainText, resolveHex, type El, type Presentation, type ShapeEl, type TableEl } from '../src/slides/model';

beforeAll(() => {
  // the reader uses the browser's DOMParser
  (globalThis as { DOMParser?: unknown }).DOMParser = DOMParser;
});

const TEMPLATES = ['slides-pitch', 'slides-lesson', 'slides-minimal', 'slides-midnight', 'slides-sunset'];

/* ============================================================== geometry */

describe('ODP geometry', () => {
  it('parses relative, smooth and arc path commands', () => {
    const segs = svgPathSegments('m10 10 h20 v10 l-5 5 z M0 0 C1 1 2 2 3 3 S5 5 6 6 Q7 7 8 8 T10 10');
    expect(segs.map((s) => s.c).join('')).toBe('MLLLZMCCCC');
    expect(segs[1].p).toEqual([30, 10]);
    expect(segs[2].p).toEqual([30, 20]);
    expect(segs[3].p).toEqual([25, 25]);
    // S reflects the previous control point
    expect(segs[7].p.slice(0, 2)).toEqual([4, 4]);
  });

  it('turns arcs into cubic curves that end where the arc ends', () => {
    const c = arcToCubics(0, 50, 50, 50, 0, false, true, 100, 50);
    expect(c.length).toBe(2);
    expect(c[1][4]).toBeCloseTo(100, 6);
    expect(c[1][5]).toBeCloseTo(50, 6);
    // the half circle passes through the top (sweep flag 1 goes clockwise on screen)
    expect(c[0][5]).toBeCloseTo(0, 6);
  });

  it('writes standard enhanced paths with LibreOffice shading and logical text areas', () => {
    const out = enhancedPath(
      [
        { d: 'M0 0L96 0L96 48Z', fill: 'norm', stroke: true },
        { d: 'M0 0L10 10', fill: 'darken', stroke: false },
      ],
      96,
      48,
      [9.6, 4.8, 9.6, 4.8],
      true,
    );
    expect(out.viewBox).toBe('0 0 2540 1270');
    expect(out.path).toBe('M 0 0 L 2540 0 L 2540 1270 Z N S M 0 0 L 265 265 N');
    expect(out.extended).toBe('M 0 0 L 2540 0 L 2540 1270 Z N H S M 0 0 L 265 265 N');
    expect(out.textAreas).toBe('?affice0 ?affice1 ?affice2 ?affice3');
    expect(out.equations).toContain('logwidth*0.1');
  });

  it('evaluates LibreOffice shape geometry (equations, modifiers and arcs)', () => {
    // LibreOffice's block arc: two arcs joined into a ring segment
    const equations = new Map(
      Object.entries({
        f0: '10800*cos($0 *(pi/180))',
        f1: '10800*sin($0 *(pi/180))',
        f2: '?f0 +10800',
        f3: '?f1 +10800',
        f4: '21600-?f2 ',
        f5: '$1 ',
        f6: '21600-$1 ',
      }),
    );
    const g = evaluateEnhancedGeometry({ viewBox: [0, 0, 21600, 21600], path: 'B 0 0 21600 21600 ?f4 ?f3 ?f2 ?f3 W ?f5 ?f5 ?f6 ?f6 ?f2 ?f3 ?f4 ?f3 Z N', equations, modifiers: [180, 5400] }, 200, 200);
    expect(g.paths.length).toBe(1);
    const cmds = g.paths[0].cmds;
    expect(cmds[0].c).toBe('M');
    // with a 180° modifier the outer arc runs from the left edge over the top to the right edge
    const ys = cmds.filter((c) => c.c === 'C').map((c) => c.p[5]);
    expect(Math.min(...ys)).toBeLessThan(100);
    expect(cmds[cmds.length - 1].c).toBe('Z');
  });

  it('evaluates OOXML-style arcs (G) with visual angles', () => {
    const g = evaluateEnhancedGeometry({ viewBox: [0, 0, 0, 0], path: 'M 0 ?f0 G ?f1 ?f0 180 180 Z N', equations: new Map([['f0', 'logheight/2'], ['f1', 'logwidth/2']]), modifiers: [] }, 96, 48);
    const last = g.paths[0].cmds.filter((c) => c.c === 'C').pop()!;
    // a half ellipse from the left edge over the top ends at the right edge (in 1/100 mm)
    expect(last.p[4]).toBeCloseTo(2540, 0);
    expect(last.p[5]).toBeCloseTo(635, 0);
  });
});

/* ============================================================== writing */

function unpack(bytes: Uint8Array) {
  const files = unzipSync(bytes);
  const xml = (p: string) => new DOMParser().parseFromString(strFromU8(files[p]), 'application/xml');
  return { files, xml };
}

describe('ODP writer', () => {
  it('writes a valid package: mimetype first and stored, manifest, masters and pages', async () => {
    const pres = slideTemplate('slides-pitch');
    const bytes = await exportOdp(pres);
    // the mimetype entry is the first file, uncompressed, right after its 30-byte header and name
    expect(strFromU8(bytes.subarray(30, 38))).toBe('mimetype');
    expect(strFromU8(bytes.subarray(38, 38 + 47))).toBe('application/vnd.oasis.opendocument.presentation');
    const { files, xml } = unpack(bytes);
    const manifest = strFromU8(files['META-INF/manifest.xml']);
    for (const f of Object.keys(files)) if (f !== 'mimetype' && f !== 'META-INF/manifest.xml') expect(manifest).toContain(`full-path="${f}"`);
    const content = xml('content.xml');
    const styles = xml('styles.xml');
    const P = 'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0';
    const S = 'urn:oasis:names:tc:opendocument:xmlns:style:1.0';
    expect(content.getElementsByTagNameNS(P, 'page').length).toBe(pres.slides.length);
    expect(styles.getElementsByTagNameNS(S, 'master-page').length).toBe(pres.layouts.length);
    // every slide's master exists
    const masters = new Set(Array.from(styles.getElementsByTagNameNS(S, 'master-page')).map((m) => m.getAttributeNS(S, 'name')));
    for (const p of Array.from(content.getElementsByTagNameNS(P, 'page'))) expect(masters.has(p.getAttributeNS(P, 'master-page-name'))).toBe(true);
    // charts are embedded chart documents with their own data
    expect(files['Object 1/content.xml']).toBeTruthy();
    expect(strFromU8(files['Object 1/content.xml'])).toContain('local-table');
  });

  it('keeps white space, tabs and line breaks', async () => {
    const pres = createPresentation('affice');
    const s = newSlide(pres, pres.layouts.find((l) => l.type === 'blank')!.id);
    s.elements.push({ id: 't1', type: 'shape', geom: 'rect', textbox: true, x: 10, y: 10, w: 400, h: 100, text: { paras: [{ runs: [{ text: '  two  spaces\tand a\nbreak ' }] }] } });
    pres.slides = [s];
    const back = await importOdp(await exportOdp(pres));
    const el = back.pres.slides[0].elements[0] as ShapeEl;
    expect(plainText(el.text)).toBe('  two  spaces\tand a\nbreak ');
  });
});

/* ============================================================ round trip */

type Summary = ReturnType<typeof summarize>;

function elSummary(e: El): unknown {
  const base = { type: e.type, x: Math.round(e.x), y: Math.round(e.y), w: Math.round(e.w), h: Math.round(e.h), rot: Math.round(e.rot ?? 0), ph: e.ph };
  switch (e.type) {
    case 'shape':
      return { ...base, geom: e.geom === 'textBox' ? 'rect' : e.geom, text: plainText(e.text), textbox: !!e.textbox };
    case 'image':
      return { ...base, geom: e.geom ?? 'rect' };
    case 'table':
      return { ...base, rows: e.rows.length, cols: e.cols.length, text: e.rows.map((r) => r.cells.map((c) => plainText(c.text)).join('|')).join('/'), style: e.style?.family };
    case 'chart':
      return { ...base, chart: { type: e.chart.type, categories: e.chart.categories, series: e.chart.series.map((s) => ({ name: s.name, values: s.values })) } };
    case 'group':
      return { ...base, children: e.children.map(elSummary) };
  }
}

function summarize(p: Presentation) {
  const layoutIndex = new Map(p.layouts.map((l, i) => [l.id, i]));
  return {
    size: p.size,
    theme: { colors: Object.fromEntries(Object.entries(p.theme.colors).map(([k, v]) => [k, v.toLowerCase()])), fonts: p.theme.fonts },
    layouts: p.layouts.map((l) => ({ name: l.name, type: l.type, decor: l.decor.length, ph: l.placeholders.map((x) => x.ph) })),
    slides: p.slides.map((s) => ({
      layout: layoutIndex.get(s.layout),
      els: s.elements.map(elSummary),
      transition: s.transition ? { type: s.transition.type, dir: s.transition.dir, dur: s.transition.dur, after: s.transition.after } : undefined,
      anims: (s.anims ?? []).map((a) => ({ cls: a.cls, effect: a.effect, start: a.start, dur: a.dur, delay: a.delay, byPara: !!a.byPara, dir: a.dir })),
      notes: s.notes,
      hidden: !!s.hidden,
    })),
  };
}

describe('ODP round trip', () => {
  for (const id of TEMPLATES) {
    it(`keeps the ${id} template`, async () => {
      const pres = slideTemplate(id);
      const { pres: back, warning } = await importOdp(await exportOdp(pres));
      expect(warning).toBeUndefined();
      const a: Summary = summarize(pres);
      const b: Summary = summarize(back);
      expect(b).toEqual(a);
      expect(back.design).toBe(pres.design);
      expect(back.master.text).toEqual(pres.master.text);
    });
  }

  it('keeps table styles, theme colours and fills', async () => {
    const pres = slideTemplate('slides-lesson');
    const { pres: back } = await importOdp(await exportOdp(pres));
    const find = (p: Presentation) => p.slides.flatMap((s) => s.elements).find((e): e is TableEl => e.type === 'table')!;
    const t1 = find(pres);
    const t2 = find(back);
    expect(t2.style).toEqual(t1.style);
    // cells keep only what differs from the table style
    expect(t2.rows.flatMap((r) => r.cells).filter((c) => c.fill || c.borders).length).toBe(t1.rows.flatMap((r) => r.cells).filter((c) => c.fill || c.borders).length);
    // theme references survive (layout decor uses theme colours)
    const decor = JSON.stringify(back.layouts.map((l) => l.decor));
    expect(decor).toContain('@accent');
  });

  it('keeps footers, hidden slides, sections and Morph', async () => {
    const pres = slideTemplate('slides-minimal');
    pres.footer = { text: 'Board meeting', slideNumber: true, date: true, skipTitle: true };
    pres.slides[1].hidden = true;
    pres.slides[2].section = 'Numbers';
    pres.slides[3].transition = { type: 'morph', dur: 1200, after: 4000 };
    const { pres: back } = await importOdp(await exportOdp(pres));
    expect(back.footer).toEqual(pres.footer);
    expect(back.slides[1].hidden).toBe(true);
    expect(back.slides[2].section).toBe('Numbers');
    expect(back.slides[3].transition).toEqual({ type: 'morph', dur: 1200, after: 4000 });
  });
});

/* ======================================================= foreign files */

const NS = [
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"',
  'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"',
  'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"',
  'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"',
  'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"',
  'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"',
  'xmlns:xlink="http://www.w3.org/1999/xlink"',
  'xmlns:dc="http://purl.org/dc/elements/1.1/"',
  'xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"',
  'xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"',
  'xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"',
  'xmlns:smil="urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0"',
  'xmlns:anim="urn:oasis:names:tc:opendocument:xmlns:animation:1.0"',
  'xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"',
  'xmlns:drawooo="http://openoffice.org/2010/draw"',
].join(' ');

const PIXEL = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

// the parts of a presentation as LibreOffice Impress writes them
const LO_STYLES = `<office:styles>
<draw:gradient draw:name="Sky" draw:style="linear" draw:start-color="#ffffff" draw:end-color="#0000ff" draw:start-intensity="100%" draw:end-intensity="100%" draw:angle="0deg" draw:border="0%"/>
<draw:marker draw:name="Arrow" svg:viewBox="0 0 20 30" svg:d="M10 0l-10 30h20z"/>
<style:default-style style:family="graphic"><style:graphic-properties svg:stroke-color="#3465a4" draw:fill-color="#729fcf"/><style:text-properties style:font-name="Liberation Sans" fo:font-size="18pt"/></style:default-style>
<style:style style:name="standard" style:family="graphic"><style:graphic-properties draw:stroke="solid" svg:stroke-width="0cm" svg:stroke-color="#3465a4" draw:fill="solid" draw:fill-color="#729fcf"/></style:style>
<style:style style:name="Default-title" style:family="presentation"><style:graphic-properties draw:stroke="none" draw:fill="none" draw:textarea-vertical-align="middle"/><style:paragraph-properties fo:text-align="center"/><style:text-properties style:font-name="DejaVu Serif" fo:font-size="40pt" fo:color="#1c1c1c"/></style:style>
<style:style style:name="Default-outline1" style:family="presentation"><style:graphic-properties draw:stroke="none" draw:fill="none"><text:list-style style:name="Default-outline1"><text:list-level-style-bullet text:level="1" text:bullet-char="●"><style:list-level-properties text:space-before="0.3cm" text:min-label-width="0.9cm"/><style:text-properties fo:font-family="OpenSymbol" fo:font-size="45%"/></text:list-level-style-bullet><text:list-level-style-bullet text:level="2" text:bullet-char="–"><style:list-level-properties text:space-before="1.5cm" text:min-label-width="0.9cm"/></text:list-level-style-bullet></text:list-style></style:graphic-properties><style:paragraph-properties fo:margin-top="0.5cm" fo:margin-bottom="0cm"/><style:text-properties style:font-name="Liberation Sans" fo:font-size="32pt" fo:color="#333333"/></style:style>
<style:style style:name="Default-outline2" style:family="presentation" style:parent-style-name="Default-outline1"><style:text-properties fo:font-size="28pt"/></style:style>
<style:style style:name="Default-background" style:family="presentation"><style:graphic-properties draw:stroke="none" draw:fill="none"/></style:style>
<style:style style:name="Default-backgroundobjects" style:family="presentation"/>
<style:style style:name="Default-notes" style:family="presentation"><style:text-properties fo:font-size="20pt"/></style:style>
<table:table-template table:name="blue"><table:first-row table:style-name="head"/><table:body table:style-name="body"/></table:table-template>
<style:style style:name="head" style:family="table-cell"><loext:graphic-properties draw:fill="solid" draw:fill-color="#1c4587"/><style:text-properties fo:color="#ffffff" fo:font-weight="bold"/></style:style>
<style:style style:name="body" style:family="table-cell"><loext:graphic-properties draw:fill="solid" draw:fill-color="#dde6f5"/></style:style>
</office:styles>`;

const LO_AUTO_STYLES = `<office:automatic-styles>
<style:page-layout style:name="PM1"><style:page-layout-properties fo:page-width="28cm" fo:page-height="15.75cm"/></style:page-layout>
<style:style style:name="Mdp1" style:family="drawing-page"><style:drawing-page-properties draw:background-size="border" draw:fill="gradient" draw:fill-gradient-name="Sky"/></style:style>
<style:style style:name="Mgr1" style:family="graphic" style:parent-style-name="standard"><style:graphic-properties draw:fill-color="#ff0000"/></style:style>
</office:automatic-styles>`;

const LO_MASTER = `<office:master-styles><draw:layer-set><draw:layer draw:name="layout"/><draw:layer draw:name="backgroundobjects"/></draw:layer-set>
<style:master-page style:name="Default" style:page-layout-name="PM1" draw:style-name="Mdp1">
<loext:theme loext:name="Office"><loext:theme-colors loext:name="Office"><loext:color loext:name="dark1" loext:color="#000000"/><loext:color loext:name="light1" loext:color="#ffffff"/><loext:color loext:name="dark2" loext:color="#44546a"/><loext:color loext:name="light2" loext:color="#e7e6e6"/><loext:color loext:name="accent1" loext:color="#4472c4"/><loext:color loext:name="accent2" loext:color="#ed7d31"/><loext:color loext:name="accent3" loext:color="#a5a5a5"/><loext:color loext:name="accent4" loext:color="#ffc000"/><loext:color loext:name="accent5" loext:color="#5b9bd5"/><loext:color loext:name="accent6" loext:color="#70ad47"/><loext:color loext:name="hyperlink" loext:color="#0563c1"/><loext:color loext:name="followed-hyperlink" loext:color="#954f72"/></loext:theme-colors></loext:theme>
<draw:custom-shape draw:style-name="Mgr1" draw:layer="backgroundobjects" svg:width="2cm" svg:height="1cm" svg:x="25cm" svg:y="14cm"><text:p/><draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:type="rectangle" draw:enhanced-path="M 0 0 L 21600 0 21600 21600 0 21600 0 0 Z N"/></draw:custom-shape>
<draw:frame presentation:style-name="Default-title" draw:layer="backgroundobjects" svg:width="25cm" svg:height="2.5cm" svg:x="1.5cm" svg:y="0.6cm" presentation:class="title" presentation:placeholder="true"><draw:text-box/></draw:frame>
<draw:frame presentation:style-name="Default-outline1" draw:layer="backgroundobjects" svg:width="25cm" svg:height="9cm" svg:x="1.5cm" svg:y="3.7cm" presentation:class="outline" presentation:placeholder="true"><draw:text-box/></draw:frame>
<draw:frame presentation:style-name="Default-backgroundobjects" draw:layer="backgroundobjects" svg:width="6cm" svg:height="1cm" svg:x="20cm" svg:y="14.3cm" presentation:class="page-number"><draw:text-box><text:p><text:page-number>&lt;number&gt;</text:page-number></text:p></draw:text-box></draw:frame>
</style:master-page></office:master-styles>`;

const LO_CONTENT_STYLES = `<office:automatic-styles>
<style:style style:name="dp1" style:family="drawing-page"><style:drawing-page-properties presentation:transition-type="automatic" presentation:duration="PT00H00M05S" presentation:transition-speed="medium" smil:type="pushWipe" smil:subtype="fromRight" presentation:display-page-number="true"/></style:style>
<style:style style:name="dp2" style:family="drawing-page"><style:drawing-page-properties presentation:display-page-number="true" draw:fill="solid" draw:fill-color="#fff2cc"/></style:style>
<style:style style:name="pr1" style:family="presentation" style:parent-style-name="Default-title"><style:graphic-properties fo:min-height="2.5cm"/></style:style>
<style:style style:name="pr2" style:family="presentation" style:parent-style-name="Default-outline1"/>
<style:style style:name="gr1" style:family="graphic" style:parent-style-name="standard"><style:graphic-properties draw:fill-color="#4472c4" draw:shadow="visible" draw:shadow-offset-x="0.2cm" draw:shadow-offset-y="0.2cm" draw:shadow-color="#000000" draw:shadow-opacity="50%"><loext:fill-complex-color loext:theme-type="accent1" loext:color-type="theme"/></style:graphic-properties></style:style>
<style:style style:name="gr2" style:family="graphic" style:parent-style-name="standard"><style:graphic-properties draw:stroke="dash" draw:stroke-dash="Fine_20_Dashed" svg:stroke-width="0.05cm" svg:stroke-color="#ff0000" draw:marker-end="Arrow" draw:marker-end-width="0.3cm" draw:fill="none"/></style:style>
<style:style style:name="gr3" style:family="graphic"><style:graphic-properties draw:stroke="none" draw:fill="none" fo:clip="rect(0in, 0in, 0in, 0in)"/></style:style>
<style:style style:name="gr4" style:family="graphic"><style:graphic-properties draw:stroke="none" draw:fill="none"/></style:style>
<style:style style:name="P1" style:family="paragraph"><style:paragraph-properties fo:text-align="start"/></style:style>
<style:style style:name="T1" style:family="text"><style:text-properties fo:font-weight="bold" fo:color="#c00000"/></style:style>
<style:style style:name="T2" style:family="text"><style:text-properties style:text-underline-style="solid" fo:font-size="150%"/></style:style>
<style:style style:name="co1" style:family="table-column"><style:table-column-properties style:column-width="6cm"/></style:style>
<style:style style:name="ro1" style:family="table-row"><style:table-row-properties style:row-height="1cm"/></style:style>
<text:list-style style:name="L1"><text:list-level-style-number text:level="1" style:num-format="1" style:num-suffix="."><style:list-level-properties text:space-before="0cm" text:min-label-width="1cm"/></text:list-level-style-number></text:list-style>
</office:automatic-styles>`;

const LO_BODY = `<office:body><office:presentation>
<draw:page draw:name="Intro" draw:style-name="dp1" draw:master-page-name="Default" xml:id="pg1" draw:id="pg1">
<draw:frame presentation:style-name="pr1" draw:layer="layout" svg:width="25cm" svg:height="2.5cm" svg:x="1.5cm" svg:y="0.6cm" presentation:class="title" presentation:user-transformed="true"><draw:text-box><text:p>Quarterly <text:span text:style-name="T1">results</text:span></text:p></draw:text-box></draw:frame>
<draw:frame presentation:style-name="pr2" draw:layer="layout" svg:width="25cm" svg:height="9cm" svg:x="1.5cm" svg:y="3.7cm" presentation:class="outline" presentation:user-transformed="true"><draw:text-box><text:list text:style-name="Default-outline1"><text:list-item><text:p xml:id="p1">Revenue grew</text:p><text:list><text:list-item><text:p xml:id="p2">Mostly <text:span text:style-name="T2">online</text:span></text:p></text:list-item></text:list></text:list-item><text:list-item><text:p xml:id="p3">Costs fell</text:p></text:list-item></text:list></draw:text-box></draw:frame>
<draw:custom-shape draw:style-name="gr1" draw:layer="layout" svg:width="4cm" svg:height="2cm" draw:transform="rotate (-0.523598775598299) translate (18cm 12cm)" xml:id="s1" draw:id="s1"><text:p text:style-name="P1">Badge</text:p><draw:enhanced-geometry svg:viewBox="0 0 0 0" draw:type="ooxml-roundRect" draw:modifiers="25000" draw:enhanced-path="M 0 0 Z N"/></draw:custom-shape>
<draw:line draw:style-name="gr2" draw:layer="layout" svg:x1="2cm" svg:y1="14cm" svg:x2="10cm" svg:y2="13cm"><text:p/></draw:line>
<draw:frame draw:style-name="gr3" draw:layer="layout" svg:width="2cm" svg:height="2cm" svg:x="12cm" svg:y="12cm"><draw:image xlink:href="Pictures/pixel.png" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/><svg:desc>A dot</svg:desc></draw:frame>
<anim:par presentation:node-type="timing-root"><anim:par smil:begin="pg1.begin"><anim:transitionFilter smil:dur="0.4s" smil:type="pushWipe" smil:subtype="fromRight"/></anim:par>
<anim:seq presentation:node-type="main-sequence">
<anim:par smil:begin="next"><anim:par smil:begin="0s"><anim:par smil:begin="0s" presentation:node-type="on-click" presentation:preset-class="entrance" presentation:preset-id="ooo-entrance-fade-in"><anim:set smil:dur="0.001s" smil:targetElement="p1" smil:attributeName="visibility" smil:to="visible"/><anim:transitionFilter smil:dur="0.6s" smil:targetElement="p1" smil:type="fade" smil:subtype="crossfade"/></anim:par></anim:par></anim:par>
<anim:par smil:begin="next"><anim:par smil:begin="0s"><anim:par smil:begin="0s" presentation:node-type="on-click" presentation:preset-class="entrance" presentation:preset-id="ooo-entrance-fade-in"><anim:set smil:dur="0.001s" smil:targetElement="p2" smil:attributeName="visibility" smil:to="visible"/><anim:transitionFilter smil:dur="0.6s" smil:targetElement="p2" smil:type="fade" smil:subtype="crossfade"/></anim:par></anim:par></anim:par>
<anim:par smil:begin="next"><anim:par smil:begin="0s"><anim:par smil:begin="0.25s" presentation:node-type="on-click" presentation:preset-class="entrance" presentation:preset-id="ooo-entrance-fly-in" presentation:preset-sub-type="from-left"><anim:set smil:dur="0.001s" smil:targetElement="s1" smil:attributeName="visibility" smil:to="visible"/><anim:animate smil:dur="0.8s" smil:targetElement="s1" smil:attributeName="x" smil:values="0-width/2;x"/></anim:par><anim:par smil:begin="0s" presentation:node-type="with-previous" presentation:preset-class="emphasis" presentation:preset-id="ooo-emphasis-spin"><anim:animateTransform smil:dur="2s" smil:targetElement="s1" smil:by="360" svg:type="rotate"/></anim:par></anim:par></anim:par>
</anim:seq></anim:par>
<presentation:notes><draw:frame presentation:class="notes" svg:width="17cm" svg:height="13cm" svg:x="2cm" svg:y="13cm"><draw:text-box><text:p>Say hello</text:p><text:p>Then the numbers</text:p></draw:text-box></draw:frame></presentation:notes>
</draw:page>
<draw:page draw:name="Table" draw:style-name="dp2" draw:master-page-name="Default">
<draw:frame draw:style-name="gr4" draw:layer="layout" svg:width="12cm" svg:height="2cm" svg:x="2cm" svg:y="4cm"><table:table table:template-name="blue" table:use-first-row-styles="true"><table:table-column table:style-name="co1" table:number-columns-repeated="2"/><table:table-row table:style-name="ro1"><table:table-cell><text:p>Region</text:p></table:table-cell><table:table-cell><text:p>Sales</text:p></table:table-cell></table:table-row><table:table-row table:style-name="ro1"><table:table-cell table:number-columns-spanned="2"><text:p>North <text:a xlink:href="#Intro">see intro</text:a></text:p></table:table-cell><table:covered-table-cell/></table:table-row></table:table></draw:frame>
<draw:frame draw:layer="layout" svg:width="10cm" svg:height="2cm" svg:x="2cm" svg:y="8cm"><draw:text-box><text:list text:style-name="L1"><text:list-item><text:p>First</text:p></text:list-item><text:list-item><text:p>Second</text:p></text:list-item></text:list></draw:text-box></draw:frame>
</draw:page>
</office:presentation></office:body>`;

function loPackage(master = LO_MASTER, body = LO_BODY): Uint8Array {
  const head = `<?xml version="1.0" encoding="UTF-8"?>`;
  const files: Record<string, Uint8Array | [Uint8Array, { level: 0 }]> = {
    mimetype: [strToU8('application/vnd.oasis.opendocument.presentation'), { level: 0 }],
    'content.xml': strToU8(`${head}<office:document-content ${NS} office:version="1.3">${LO_CONTENT_STYLES}${body}</office:document-content>`),
    'styles.xml': strToU8(`${head}<office:document-styles ${NS} office:version="1.3">${LO_STYLES}${LO_AUTO_STYLES}${master}</office:document-styles>`),
    'meta.xml': strToU8(`${head}<office:document-meta ${NS}><office:meta><meta:generator>LibreOffice/24.2.7.2$Linux_X86_64</meta:generator><dc:title>Q3 review</dc:title><meta:initial-creator>Sam</meta:initial-creator></office:meta></office:document-meta>`),
    'Pictures/pixel.png': Uint8Array.from(atob(PIXEL), (c) => c.charCodeAt(0)),
  };
  return zipSync(files);
}

function checkLoFixture(p: Presentation) {
  // theme and master text
  expect(p.theme.colors.accent1.toLowerCase()).toBe('#4472c4');
  expect(p.theme.fonts.major).toBe('DejaVu Serif');
  expect(p.theme.fonts.minor).toBe('Liberation Sans');
  expect(p.master.text.title.size).toBe(40);
  expect(p.master.text.body.size).toBe(32);
  expect(p.size).toEqual({ w: 1058, h: 595 });
  // master page → layout with decor, placeholders and the gradient background
  expect(p.layouts).toHaveLength(1);
  expect(p.layouts[0].placeholders.map((x) => x.ph)).toEqual(['title', 'obj']);
  expect(p.layouts[0].decor).toHaveLength(1);
  expect(p.master.background.type).toBe('gradient');
  const s1 = p.slides[0];
  // title: plain first word, bold red second
  const title = s1.elements[0] as ShapeEl;
  expect(title.ph).toBe('title');
  expect(title.text!.paras[0].runs.map((r) => r.text)).toEqual(['Quarterly ', 'results']);
  expect(title.text!.paras[0].runs[1]).toMatchObject({ b: true, color: '#c00000' });
  expect(title.text!.paras[0].runs[0].size).toBeUndefined();
  // outline: nested list levels, relative font size
  const body = s1.elements[1] as ShapeEl;
  expect(body.text!.paras.map((q) => [q.level ?? 0, q.runs.map((r) => r.text).join('')])).toEqual([
    [0, 'Revenue grew'],
    [1, 'Mostly online'],
    [0, 'Costs fell'],
  ]);
  // LibreOffice's "●" differs from our default bullet; its level-2 "–" is our default, so it stays implicit
  expect(body.text!.paras[0].bullet).toMatchObject({ type: 'char', char: '●' });
  expect(body.text!.paras[1].bullet).toBeUndefined();
  // the second level uses the master's "outline2" size (28pt), and "online" is 150 % of it
  expect(body.text!.paras[1].runs[0].size).toBe(28);
  expect(body.text!.paras[1].runs[1]).toMatchObject({ u: true, size: 42 });
  // rotated OOXML rounded rectangle with a theme fill and shadow
  const badge = s1.elements[2] as ShapeEl;
  expect(badge.geom).toBe('roundRect');
  expect(badge.adj).toEqual({ adj: 25000 });
  expect(badge.rot).toBeCloseTo(30, 1);
  expect(badge.fill).toEqual({ type: 'solid', color: '@accent1' });
  expect(badge.shadow).toMatchObject({ angle: 45 });
  expect(plainText(badge.text)).toBe('Badge');
  // a dashed arrow line going up to the right
  const line = s1.elements[3] as ShapeEl;
  expect(line.geom).toBe('line');
  expect(line.flipV).toBe(true);
  expect(line.line).toMatchObject({ color: '#ff0000', tail: 'triangle' });
  expect(line.line!.dash).toBeTruthy();
  // picture with alt text
  const img = s1.elements[4];
  expect(img.type).toBe('image');
  expect(img.alt).toBe('A dot');
  // transition with auto advance, and animations (paragraph by paragraph, then fly-in with a spin)
  expect(s1.transition).toEqual({ type: 'push', dir: 'l', dur: 400, after: 5000 });
  expect(s1.anims!.map((a) => [a.cls, a.effect, a.start, a.byPara ?? false, a.dir, a.dur, a.delay])).toEqual([
    ['entr', 'fade', 'click', true, undefined, 600, 0],
    ['entr', 'fly', 'click', false, 'l', 800, 250],
    ['emph', 'spin', 'with', false, undefined, 2000, 0],
  ]);
  expect(s1.anims![0].el).toBe(body.id);
  expect(s1.anims![1].el).toBe(badge.id);
  expect(s1.notes).toBe('Say hello\nThen the numbers');
  // second slide: its own background, a table from a LibreOffice table design, a numbered list
  const s2 = p.slides[1];
  expect(s2.background).toEqual({ type: 'solid', color: '#fff2cc' });
  const t = s2.elements[0] as TableEl;
  expect(t.cols).toHaveLength(2);
  expect(t.rows[0].cells[0].fill).toEqual({ type: 'solid', color: '#1c4587' });
  expect(t.rows[0].cells[0].text.paras[0].runs[0]).toMatchObject({ b: true, color: '#ffffff' });
  expect(t.rows[1].cells[0].colSpan).toBe(2);
  expect(t.rows[1].cells[1].merged).toBe(true);
  expect(t.rows[1].cells[0].text.paras[0].runs[1]).toMatchObject({ text: 'see intro', link: '#slide:1' });
  const list = s2.elements[1] as ShapeEl;
  expect(list.text!.paras.map((q) => q.bullet)).toEqual([
    { type: 'num', style: 'arabicPeriod' },
    { type: 'num', style: 'arabicPeriod' },
  ]);
  // the page number is shown (the master has a page-number frame)
  expect(p.footer).toEqual({ slideNumber: true });
  expect(p.props).toMatchObject({ title: 'Q3 review', author: 'Sam' });
}

describe('ODP reader', () => {
  it('reads a presentation written by LibreOffice Impress', async () => {
    const { pres, warning } = await importOdp(loPackage());
    expect(warning).toBeUndefined();
    checkLoFixture(pres);
  });

  it('reads flat OpenDocument presentations (.fodp)', async () => {
    const flat = `<?xml version="1.0" encoding="UTF-8"?><office:document ${NS} office:version="1.3" office:mimetype="application/vnd.oasis.opendocument.presentation"><office:meta><meta:generator>LibreOffice/24.2.7.2$Linux_X86_64</meta:generator><dc:title>Q3 review</dc:title><meta:initial-creator>Sam</meta:initial-creator></office:meta>${LO_STYLES}${LO_AUTO_STYLES.replace('</office:automatic-styles>', '')}${LO_CONTENT_STYLES.replace('<office:automatic-styles>', '')}${LO_MASTER}${LO_BODY.replace('xlink:href="Pictures/pixel.png" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>', `><office:binary-data>${PIXEL}</office:binary-data></draw:image>`)}</office:document>`;
    const { pres } = await importOdp(strToU8(flat));
    checkLoFixture(pres);
  });

  it('recognises the PowerPoint layout names LibreOffice keeps on converted masters', async () => {
    const layoutOf = async (display: string) => {
      const master = LO_MASTER.replace('style:name="Default"', `style:name="M1" style:display-name="${display}"`);
      const { pres } = await importOdp(loPackage(master, LO_BODY.replaceAll('draw:master-page-name="Default"', 'draw:master-page-name="M1"')));
      return pres.layouts[0];
    };
    const title = await layoutOf('Title Slide');
    expect(title.type).toBe('title');
    expect(title.placeholders.map((x) => x.ph)).toEqual(['ctrTitle', 'subTitle']);
    expect((await layoutOf('Title Only')).type).toBe('titleOnly');
    expect((await layoutOf('Section Header')).type).toBe('secHead');
    // other names keep the type the placeholders suggest
    const other = await layoutOf('Ocean');
    expect(other.type).toBe('obj');
    expect(other.placeholders.map((x) => x.ph)).toEqual(['title', 'obj']);
  });

  it('rejects files that are not presentations', async () => {
    const text = zipSync({ mimetype: strToU8('application/vnd.oasis.opendocument.text'), 'content.xml': strToU8(`<office:document-content ${NS}><office:body><office:text/></office:body></office:document-content>`) });
    await expect(importOdp(text)).rejects.toThrow(/not a presentation/);
    await expect(importOdp(strToU8('not a zip'))).rejects.toThrow();
  });

  it('keeps colours that equal the theme as the plain colours they are', async () => {
    const pres = slideTemplate('slides-pitch');
    const { pres: back } = await importOdp(await exportOdp(pres));
    const theme = back.theme;
    // white text on the pitch deck is written as plain white, not as the (white) theme text colour
    const runs = back.slides.flatMap((s) => s.elements).flatMap((e) => (e.type === 'shape' && e.text ? e.text.paras.flatMap((p) => p.runs) : []));
    for (const r of runs) if (r.color) expect(resolveHex(r.color, theme).hex).toMatch(/^#[0-9a-f]{6}$/);
  });
});
