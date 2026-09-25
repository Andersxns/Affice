import { gunzipSync, gzipSync, strFromU8, strToU8 } from 'fflate';
import type { TabSource } from '@/app/workspace';
import type { Presentation } from '../model';
import { createPresentation } from '../themes';

export interface LoadedPresentation {
  pres: Presentation;
  warning?: string;
}

export function serializeAfslides(pres: Presentation): Uint8Array {
  return gzipSync(strToU8(JSON.stringify(pres)), { level: 6 });
}

export function parseAfslides(bytes: Uint8Array): Presentation {
  const text = bytes[0] === 0x1f && bytes[1] === 0x8b ? strFromU8(gunzipSync(bytes)) : new TextDecoder().decode(bytes);
  const p = JSON.parse(text) as Presentation;
  if (p.format !== 'affice-slides') throw new Error('This is not an Affice presentation.');
  return p;
}

export async function loadPresentation(source: TabSource | undefined): Promise<LoadedPresentation> {
  if (!source || source.type === 'blank') return { pres: createPresentation('affice') };
  if (source.type === 'template') {
    const { slideTemplate } = await import('../templates');
    return { pres: slideTemplate(source.id) };
  }
  if (source.type === 'recovery') return { pres: JSON.parse(source.data) as Presentation };
  if (source.type === 'data') return source.data as LoadedPresentation;
  const { file, formatId } = source;
  switch (formatId) {
    case 'afslides':
      return { pres: parseAfslides(file.data) };
    default: {
      const { importPptx } = await import('./pptx-import');
      return importPptx(file.data);
    }
  }
}

export async function exportPresentation(pres: Presentation, formatId: string, title: string): Promise<Uint8Array> {
  switch (formatId) {
    case 'afslides':
      return serializeAfslides(pres);
    case 'pptx': {
      const { exportPptx } = await import('./pptx-export');
      return exportPptx(pres);
    }
    case 'md': {
      const { outlineMarkdown } = await import('./outline');
      return new TextEncoder().encode(outlineMarkdown(pres, title));
    }
    case 'html': {
      const { exportHtmlShow } = await import('./html');
      return exportHtmlShow(pres, title);
    }
    default:
      throw new Error(`Unsupported format ${formatId}`);
  }
}
