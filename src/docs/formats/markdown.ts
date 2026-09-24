import { marked } from 'marked';
import TurndownService from 'turndown';
import { gfm } from 'turndown-plugin-gfm';
import { defaultSettings } from '../model';
import type { LoadedDoc } from './index';

export function importMarkdown(md: string): LoadedDoc {
  // front matter (--- title: … ---) becomes the document title
  let title = '';
  const fm = /^---\n([\s\S]*?)\n---\n?/.exec(md);
  if (fm) {
    title = /^title:\s*(.+)$/m.exec(fm[1])?.[1]?.replace(/^["']|["']$/g, '') ?? '';
    md = md.slice(fm[0].length);
  }
  // $$…$$ and $…$ maths → math nodes
  md = md
    .replace(/\$\$([\s\S]+?)\$\$/g, (_m, tex) => `<div data-math-block data-latex="${escapeAttr(tex.trim())}"></div>`)
    .replace(/(^|[^\\$])\$([^\s$][^$\n]*?[^\s$]|[^\s$])\$(?!\d)/g, (_m, pre, tex) => `${pre}<span data-math-inline data-latex="${escapeAttr(tex)}"></span>`);
  const html = marked.parse(md, { gfm: true, breaks: false, async: false }) as string;
  // GFM task lists → TipTap task list markup
  const doc = new DOMParser().parseFromString(`<body>${html}</body>`, 'text/html');
  doc.querySelectorAll('ul').forEach((ul) => {
    const items = Array.from(ul.children).filter((c) => c.tagName === 'LI');
    if (items.length && items.every((li) => li.querySelector(':scope > input[type=checkbox], :scope > p > input[type=checkbox]'))) {
      ul.setAttribute('data-type', 'taskList');
      items.forEach((li) => {
        const box = li.querySelector('input[type=checkbox]') as HTMLInputElement;
        li.setAttribute('data-type', 'taskItem');
        li.setAttribute('data-checked', String(box.checked || box.hasAttribute('checked')));
        box.remove();
      });
    }
  });
  const settings = defaultSettings();
  settings.title = title;
  return { settings, content: doc.body.innerHTML, comments: [] };
}

function escapeAttr(s: string) {
  return s.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');
}

export function exportMarkdown(html: string): string {
  const td = new TurndownService({
    headingStyle: 'atx',
    codeBlockStyle: 'fenced',
    bulletListMarker: '-',
    emDelimiter: '*',
    strongDelimiter: '**',
    hr: '---',
  });
  td.use(gfm);
  td.addRule('taskItem', {
    filter: (node) => node.nodeName === 'LI' && (node as HTMLElement).getAttribute('data-type') === 'taskItem',
    replacement: (content, node) => {
      const checked = (node as HTMLElement).getAttribute('data-checked') === 'true';
      return `- [${checked ? 'x' : ' '}] ${content.trim().replace(/\n+/g, ' ')}\n`;
    },
  });
  td.addRule('image', {
    filter: (node) => node.nodeName === 'SPAN' && (node as HTMLElement).classList.contains('doc-img'),
    replacement: (_c, node) => {
      const img = (node as HTMLElement).querySelector('img');
      return img ? `![${img.getAttribute('alt') ?? ''}](${img.getAttribute('src') ?? ''})` : '';
    },
  });
  td.addRule('highlight', { filter: ['mark'], replacement: (c) => `==${c}==` });
  td.addRule('underline', { filter: ['u'], replacement: (c) => `<u>${c}</u>` });
  td.addRule('sub', { filter: ['sub'], replacement: (c) => `<sub>${c}</sub>` });
  td.addRule('sup', { filter: ['sup'], replacement: (c) => `<sup>${c}</sup>` });
  td.addRule('pageBreak', { filter: (n) => (n as HTMLElement).classList?.contains('page-break'), replacement: () => '\n\n' });
  const out = td.turndown(html);
  return `${out.replace(/\n{3,}/g, '\n\n').trim()}\n`;
}
