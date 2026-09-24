import type { JSONContent } from '@tiptap/core';
import type { PdfJob } from '@/shared/types';
import type { TabSource } from '@/app/workspace';
import { api } from '@/lib/platform';
import { encoder } from '@/lib/utils';
import { defaultSettings, EMPTY_DOC, pageDims, parseAfdoc, serializeAfdoc, type DocModel } from '../model';
import { buildStandaloneHtml, printCss, prepareExportHtml } from './html';

export type LoadedDoc = Omit<DocModel, 'content'> & {
  content: JSONContent | string;
  warning?: string;
};

function textToHtml(text: string): string {
  const esc = (s: string) => s.replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' })[c]!);
  return text
    .replace(/\r\n?/g, '\n')
    .split('\n')
    .map((line) => `<p>${esc(line) || ''}</p>`)
    .join('');
}

function decodeText(bytes: Uint8Array): string {
  // BOM sniffing: UTF-8, UTF-16 LE/BE; fall back to windows-1252 for legacy files
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder('utf-16le').decode(bytes.subarray(2));
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return new TextDecoder('utf-16be').decode(bytes.subarray(2));
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(bytes);
  } catch {
    return new TextDecoder('windows-1252').decode(bytes);
  }
}

export async function loadDocument(source: TabSource | undefined): Promise<LoadedDoc> {
  const blank = (): LoadedDoc => ({ settings: defaultSettings(), content: EMPTY_DOC, comments: [] });
  if (!source || source.type === 'blank') return blank();
  if (source.type === 'template') {
    const { docTemplate } = await import('../templates');
    return docTemplate(source.id);
  }
  if (source.type === 'recovery') {
    const raw = JSON.parse(source.data);
    return { settings: { ...defaultSettings(), ...raw.settings }, content: raw.content ?? EMPTY_DOC, comments: raw.comments ?? [] };
  }
  if (source.type === 'data') return source.data as LoadedDoc;
  const { file, formatId } = source;
  const bytes = file.data;
  switch (formatId) {
    case 'docx': {
      const { importDocx } = await import('./docx-import');
      return importDocx(bytes);
    }
    case 'odt': {
      const { importOdt } = await import('./odt');
      return importOdt(bytes);
    }
    case 'rtf': {
      const { importRtf } = await import('./rtf');
      return importRtf(new TextDecoder('windows-1252').decode(bytes));
    }
    case 'html': {
      const { importHtml } = await import('./html');
      return importHtml(decodeText(bytes));
    }
    case 'md': {
      const { importMarkdown } = await import('./markdown');
      return importMarkdown(decodeText(bytes));
    }
    case 'afdoc':
      return parseAfdoc(bytes);
    case 'doc-text': {
      const text = await api.convert.legacyDocText(bytes);
      return {
        ...blank(),
        content: textToHtml(text),
        warning: 'This older Word (.doc) file was opened as text only. Install LibreOffice to keep its formatting, or save it as .docx.',
      };
    }
    case 'txt':
    default:
      return { ...blank(), content: textToHtml(decodeText(bytes)) };
  }
}

export interface ExportExtras {
  html: string;
  text: string;
  title: string;
  headingPages: Map<number, number>;
}

export async function exportDocument(model: DocModel, formatId: string, extra: ExportExtras): Promise<Uint8Array> {
  switch (formatId) {
    case 'docx': {
      const { exportDocx } = await import('./docx-export');
      return exportDocx(model, extra.title);
    }
    case 'odt': {
      const { exportOdt } = await import('./odt');
      return exportOdt(model, extra.title);
    }
    case 'rtf': {
      const { exportRtf } = await import('./rtf');
      return encoder.encode(exportRtf(model));
    }
    case 'html':
      return encoder.encode(buildStandaloneHtml(model, extra.html, extra.title, extra.headingPages));
    case 'md': {
      const { exportMarkdown } = await import('./markdown');
      return encoder.encode(exportMarkdown(prepareExportHtml(extra.html, model, extra.headingPages, { forMarkdown: true })));
    }
    case 'epub': {
      const { exportEpub } = await import('./epub');
      return exportEpub(model, prepareExportHtml(extra.html, model, extra.headingPages), extra.title);
    }
    case 'afdoc':
      return serializeAfdoc(model);
    case 'txt':
    default:
      return encoder.encode(extra.text.replace(/\n/g, '\r\n'));
  }
}

export function buildPdfJob(model: DocModel, html: string, title: string, headingPages: Map<number, number>): PdfJob {
  const { w, h } = pageDims(model.settings.page);
  const m = model.settings.page.margins;
  return {
    html: `<div class="doc-content">${prepareExportHtml(html, model, headingPages)}</div>`,
    css: printCss(model),
    bodyClass: 'print-doc',
    pageWidthIn: w,
    pageHeightIn: h,
    marginsIn: { top: m.top, right: m.right, bottom: m.bottom, left: m.left },
    title,
  };
}
