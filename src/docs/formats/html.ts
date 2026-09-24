import contentCss from '../content.css?inline';
import katexCss from 'katex/dist/katex.min.css?inline';
import { escapeHtml } from '@/lib/utils';
import { renderMathToHtml, collectHeadings } from '../editor/extensions';
import { defaultSettings, styleVarsCss, type DocModel } from '../model';
import type { LoadedDoc } from './index';

export function slugify(s: string): string {
  return (
    s
      .toLowerCase()
      .normalize('NFKD')
      .replace(/[̀-ͯ]/g, '')
      .replace(/[^a-z0-9]+/g, '-')
      .replace(/^-+|-+$/g, '')
      .slice(0, 60) || 'section'
  );
}

/**
 * Turns the editor's HTML into export-ready HTML: renders equations, fills in the table of
 * contents, adds heading anchors and removes editor-only markup (comments).
 */
export function prepareExportHtml(html: string, _model: DocModel, headingPages: Map<number, number>, opts: { forMarkdown?: boolean } = {}): string {
  const doc = new DOMParser().parseFromString(`<body>${html}</body>`, 'text/html');
  const body = doc.body;

  // comments → plain text
  body.querySelectorAll('span[data-comment]').forEach((s) => s.replaceWith(...Array.from(s.childNodes)));

  // headings get stable ids
  const used = new Set<string>();
  const headingEls = Array.from(body.querySelectorAll('h1,h2,h3,h4,h5,h6'));
  headingEls.forEach((h) => {
    let id = slugify(h.textContent ?? '');
    let n = 2;
    while (used.has(id)) id = `${slugify(h.textContent ?? '')}-${n++}`;
    used.add(id);
    h.id = id;
  });

  // equations
  body.querySelectorAll('[data-math-inline],[data-math-block]').forEach((el) => {
    const latex = el.getAttribute('data-latex') ?? el.textContent ?? '';
    const display = el.hasAttribute('data-math-block');
    if (opts.forMarkdown) {
      el.replaceWith(doc.createTextNode(display ? `\n$$${latex}$$\n` : `$${latex}$`));
      return;
    }
    const wrap = doc.createElement(display ? 'div' : 'span');
    wrap.className = display ? 'math-block' : 'math-node';
    wrap.innerHTML = renderMathToHtml(latex, display);
    el.replaceWith(wrap);
  });

  // table of contents
  body.querySelectorAll('[data-toc]').forEach((el) => {
    const title = el.getAttribute('data-title') || 'Contents';
    const pageList = Array.from(headingPages.entries())
      .sort((a, b) => a[0] - b[0])
      .map((e) => e[1]);
    const levelHeads = headingEls.map((h, i) => ({ h, page: pageList[i] ?? '' })).filter((x) => Number(x.h.tagName[1]) <= 3 && x.h.textContent?.trim());
    const toc = doc.createElement('div');
    toc.className = 'toc';
    toc.innerHTML =
      `<div class="toc-title">${escapeHtml(title)}</div>` +
      levelHeads
        .map(({ h, page }) =>
          opts.forMarkdown
            ? `<p><a href="#${h.id}">${escapeHtml(h.textContent ?? '')}</a></p>`
            : `<a class="toc-entry" data-level="${h.tagName[1]}" href="#${h.id}"><span>${escapeHtml(h.textContent ?? '')}</span><span class="toc-dots"></span><span>${page}</span></a>`,
        )
        .join('');
    el.replaceWith(toc);
  });

  // page breaks for markdown are meaningless
  if (opts.forMarkdown) body.querySelectorAll('.page-break').forEach((el) => el.remove());

  return body.innerHTML;
}

function cssString(s: string): string {
  return `"${s.replace(/\\/g, '\\\\').replace(/"/g, '\\"')}"`;
}

/** CSS content value for header/footer text with {PAGE}/{PAGES} fields. */
function marginContent(text: string, title: string): string {
  if (!text.trim()) return 'none';
  const parts: string[] = [];
  const re = /\{(PAGE|PAGES|DATE|TITLE)\}/gi;
  let last = 0;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text))) {
    if (m.index > last) parts.push(cssString(text.slice(last, m.index)));
    const f = m[1].toUpperCase();
    if (f === 'PAGE') parts.push('counter(page)');
    else if (f === 'PAGES') parts.push('counter(pages)');
    else if (f === 'DATE') parts.push(cssString(new Date().toLocaleDateString()));
    else parts.push(cssString(title));
    last = m.index + m[0].length;
  }
  if (last < text.length) parts.push(cssString(text.slice(last)));
  return parts.join(' ');
}

/** Stylesheet for printing/PDF: document styles, equations, page colour and header/footer margin boxes. */
export function printCss(model: DocModel, title = model.settings.title): string {
  const s = model.settings;
  const font = `font-family: "${(s.styles.normal?.font ?? 'Calibri').replace(/"/g, '')}", Carlito, sans-serif; font-size: 9pt; color: #555;`;
  const box = (name: string, text: string, align: string) =>
    text.trim() ? `@${name} { content: ${marginContent(text, title)}; ${font} text-align: ${align}; vertical-align: middle; }` : '';
  const boxes = [
    box('top-left', s.header.left, 'left'),
    box('top-center', s.header.center, 'center'),
    box('top-right', s.header.right, 'right'),
    box('bottom-left', s.footer.left, 'left'),
    box('bottom-center', s.footer.center, 'center'),
    box('bottom-right', s.footer.right, 'right'),
  ].join('\n');
  const firstPage = s.differentFirstPage
    ? `@page :first { @top-left{content:none} @top-center{content:none} @top-right{content:none} @bottom-left{content:none} @bottom-center{content:none} @bottom-right{content:none} }`
    : '';
  const watermark = s.watermark
    ? `.print-doc::before{content:${cssString(s.watermark)};position:fixed;left:50%;top:50%;transform:translate(-50%,-50%) rotate(-40deg);font:700 90pt Calibri,Carlito,sans-serif;color:rgba(0,0,0,0.07);white-space:nowrap;z-index:0;pointer-events:none}`
    : '';
  return `${contentCss}
${katexCss}
:root{${styleVarsCss(s)}}
body.print-doc{margin:0;background:${s.pageColor ?? '#fff'}!important;-webkit-print-color-adjust:exact;print-color-adjust:exact}
.doc-content{${styleVarsCss(s)}}
@page { ${boxes} }
${firstPage}
${watermark}`;
}

export function buildStandaloneHtml(model: DocModel, html: string, title: string, headingPages: Map<number, number>): string {
  const s = model.settings;
  const body = prepareExportHtml(html, model, headingPages);
  const meta = [
    s.author && `<meta name="author" content="${escapeHtml(s.author)}">`,
    s.subject && `<meta name="description" content="${escapeHtml(s.subject)}">`,
    s.keywords && `<meta name="keywords" content="${escapeHtml(s.keywords)}">`,
  ]
    .filter(Boolean)
    .join('\n');
  return `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="generator" content="Affice">
${meta}
<title>${escapeHtml(s.title || title.replace(/\.[^.]+$/, ''))}</title>
<style>
${contentCss}
${katexCss.replace(/url\([^)]*\)/g, 'local("KaTeX")')}
body{margin:0;background:${s.pageColor ?? '#f3f4f7'}}
.doc-content{${styleVarsCss(s)};max-width:${(s.page.width - s.page.margins.left - s.page.margins.right).toFixed(2)}in;margin:0 auto;padding:1in 0.75in;background:#fff;min-height:100vh;box-sizing:content-box}
@media (max-width:760px){.doc-content{padding:24px 18px}}
@media print{body{background:#fff}.doc-content{padding:0;max-width:none}}
</style>
</head>
<body>
<article class="doc-content">
${body}
</article>
</body>
</html>`;
}

export function importHtml(text: string): LoadedDoc {
  const doc = new DOMParser().parseFromString(text, 'text/html');
  doc.querySelectorAll('script,style,noscript,iframe,object,embed,link,meta').forEach((n) => n.remove());
  // Word "Save as HTML" uses MsoTitle/MsoSubtitle classes
  doc.querySelectorAll('p.MsoTitle').forEach((p) => p.setAttribute('data-style', 'title'));
  doc.querySelectorAll('p.MsoSubtitle').forEach((p) => p.setAttribute('data-style', 'subtitle'));
  doc.querySelectorAll('p.MsoQuote').forEach((p) => p.setAttribute('data-style', 'quote'));
  const settings = defaultSettings();
  settings.title = doc.title || '';
  const root = doc.querySelector('article.doc-content') ?? doc.body;
  return { settings, content: root.innerHTML, comments: [] };
}

/** Collects heading positions for callers that need a TOC without an editor view. */
export { collectHeadings };
