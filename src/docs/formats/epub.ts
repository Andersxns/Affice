import { strToU8, zipSync, type Zippable } from 'fflate';
import contentCss from '../content.css?inline';
import { dataUrlToBytes, escapeXml, extFromMime, uid } from '@/lib/utils';
import { styleVarsCss, type DocModel } from '../model';

/** Exports an EPUB 3 e-book: one XHTML file per top-level (H1) section with a navigation document. */
export function exportEpub(model: DocModel, html: string, title: string): Uint8Array {
  const s = model.settings;
  const bookTitle = s.title || title.replace(/\.[^.]+$/, '');
  const doc = new DOMParser().parseFromString(`<body>${html}</body>`, 'text/html');

  // move embedded images to separate files
  const images: Array<{ name: string; mime: string; data: Uint8Array }> = [];
  doc.querySelectorAll('img').forEach((img) => {
    const src = img.getAttribute('src') ?? '';
    const d = src.startsWith('data:') ? dataUrlToBytes(src) : null;
    if (!d) return;
    const name = `images/img${images.length + 1}.${extFromMime(d.mime)}`;
    images.push({ name, mime: d.mime, data: d.bytes });
    img.setAttribute('src', name);
  });
  doc.querySelectorAll('.page-break').forEach((el) => el.remove());

  // split into chapters at H1
  const chapters: Array<{ title: string; html: string }> = [];
  let current: { title: string; nodes: string[] } = { title: bookTitle, nodes: [] };
  for (const node of Array.from(doc.body.childNodes)) {
    if (node instanceof HTMLElement && node.tagName === 'H1' && current.nodes.length) {
      chapters.push({ title: current.title, html: current.nodes.join('') });
      current = { title: node.textContent?.trim() || `Chapter ${chapters.length + 1}`, nodes: [] };
    } else if (node instanceof HTMLElement && node.tagName === 'H1') {
      current.title = node.textContent?.trim() || current.title;
    }
    current.nodes.push(serializeXhtml(node));
  }
  chapters.push({ title: current.title, html: current.nodes.join('') });

  const id = `urn:uuid:${crypto.randomUUID?.() ?? uid()}`;
  const modified = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  const css = `${contentCss}\n.doc-content{${styleVarsCss(s)}}\nbody{margin:0 5%;}\nimg{max-width:100%}`;

  const files: Zippable = {
    mimetype: [strToU8('application/epub+zip'), { level: 0 }],
    'META-INF/container.xml': strToU8(
      `<?xml version="1.0" encoding="UTF-8"?><container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container"><rootfiles><rootfile full-path="OEBPS/content.opf" media-type="application/oebps-package+xml"/></rootfiles></container>`,
    ),
    'OEBPS/style.css': strToU8(css),
  };
  chapters.forEach((c, i) => {
    files[`OEBPS/chapter${i + 1}.xhtml`] = strToU8(
      `<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE html>\n<html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops" xml:lang="en"><head><meta charset="UTF-8"/><title>${escapeXml(c.title)}</title><link rel="stylesheet" type="text/css" href="style.css"/></head><body><div class="doc-content">${c.html}</div></body></html>`,
    );
  });
  images.forEach((im) => {
    files[`OEBPS/${im.name}`] = im.data;
  });
  const nav = `<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE html>\n<html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops"><head><title>${escapeXml(bookTitle)}</title></head><body><nav epub:type="toc" id="toc"><h1>Contents</h1><ol>${chapters
    .map((c, i) => `<li><a href="chapter${i + 1}.xhtml">${escapeXml(c.title)}</a></li>`)
    .join('')}</ol></nav></body></html>`;
  files['OEBPS/nav.xhtml'] = strToU8(nav);
  const manifest = [
    `<item id="nav" href="nav.xhtml" media-type="application/xhtml+xml" properties="nav"/>`,
    `<item id="css" href="style.css" media-type="text/css"/>`,
    ...chapters.map((_c, i) => `<item id="ch${i + 1}" href="chapter${i + 1}.xhtml" media-type="application/xhtml+xml"/>`),
    ...images.map((im, i) => `<item id="img${i + 1}" href="${im.name}" media-type="${im.mime}"/>`),
  ].join('');
  const spine = chapters.map((_c, i) => `<itemref idref="ch${i + 1}"/>`).join('');
  files['OEBPS/content.opf'] = strToU8(
    `<?xml version="1.0" encoding="UTF-8"?><package xmlns="http://www.idpf.org/2007/opf" version="3.0" unique-identifier="bookid"><metadata xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:identifier id="bookid">${id}</dc:identifier><dc:title>${escapeXml(bookTitle)}</dc:title><dc:language>en</dc:language>${
      s.author ? `<dc:creator>${escapeXml(s.author)}</dc:creator>` : ''
    }<meta property="dcterms:modified">${modified}</meta></metadata><manifest>${manifest}</manifest><spine>${spine}</spine></package>`,
  );
  return zipSync(files, { level: 6 });
}

/** Serialises a DOM node as well-formed XHTML. */
function serializeXhtml(node: Node): string {
  return new XMLSerializer().serializeToString(node).replace(/ xmlns="http:\/\/www\.w3\.org\/1999\/xhtml"/g, '');
}
