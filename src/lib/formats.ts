import type { AppKind, FileFilter } from '@/shared/types';

export interface FormatInfo {
  id: string;
  label: string;
  exts: string[];
  kind: AppKind;
  open: 'native' | 'libreoffice' | 'none';
  save: 'native' | 'libreoffice' | 'none';
  /** Shown before saving: what gets lost in this format. */
  lossy?: string;
}

export const FORMATS: FormatInfo[] = [
  // Documents
  { id: 'docx', label: 'Word Document', exts: ['docx', 'docm', 'dotx', 'dotm'], kind: 'doc', open: 'native', save: 'native' },
  { id: 'odt', label: 'OpenDocument Text', exts: ['odt', 'ott'], kind: 'doc', open: 'native', save: 'native' },
  { id: 'rtf', label: 'Rich Text Format', exts: ['rtf'], kind: 'doc', open: 'native', save: 'native' },
  { id: 'doc', label: 'Word 97-2003 Document', exts: ['doc', 'dot'], kind: 'doc', open: 'native', save: 'libreoffice' },
  { id: 'html', label: 'Web Page', exts: ['html', 'htm', 'xhtml'], kind: 'doc', open: 'native', save: 'native' },
  { id: 'md', label: 'Markdown', exts: ['md', 'markdown', 'mdown'], kind: 'doc', open: 'native', save: 'native', lossy: 'Markdown keeps headings, lists, links, tables and emphasis — fonts, colours and page layout are not saved.' },
  { id: 'txt', label: 'Plain Text', exts: ['txt', 'text', 'log'], kind: 'doc', open: 'native', save: 'native', lossy: 'Plain text keeps only the words — all formatting, images and tables are removed.' },
  { id: 'wpd', label: 'WordPerfect / Pages / other', exts: ['wpd', 'pages', 'wps', 'abw', 'fodt', 'sxw'], kind: 'doc', open: 'libreoffice', save: 'none' },
  { id: 'afdoc', label: 'Affice Document', exts: ['afdoc'], kind: 'doc', open: 'native', save: 'native' },
  // Sheets
  { id: 'xlsx', label: 'Excel Workbook', exts: ['xlsx', 'xlsm', 'xltx', 'xltm'], kind: 'sheet', open: 'native', save: 'native' },
  { id: 'xls', label: 'Excel 97-2003 Workbook', exts: ['xls', 'xlt'], kind: 'sheet', open: 'native', save: 'native', lossy: 'The old .xls format has limited styling — some formatting may be simplified.' },
  { id: 'xlsb', label: 'Excel Binary Workbook', exts: ['xlsb'], kind: 'sheet', open: 'native', save: 'none' },
  { id: 'ods', label: 'OpenDocument Spreadsheet', exts: ['ods', 'ots', 'fods'], kind: 'sheet', open: 'native', save: 'native' },
  { id: 'numbers', label: 'Apple Numbers', exts: ['numbers'], kind: 'sheet', open: 'native', save: 'none' },
  { id: 'csv', label: 'CSV (Comma separated)', exts: ['csv'], kind: 'sheet', open: 'native', save: 'native', lossy: 'CSV saves only the values of the current sheet — formatting, formulas and other sheets are not saved.' },
  { id: 'tsv', label: 'Tab separated values', exts: ['tsv', 'tab'], kind: 'sheet', open: 'native', save: 'native', lossy: 'TSV saves only the values of the current sheet.' },
  { id: 'afsheet', label: 'Affice Sheet', exts: ['afsheet'], kind: 'sheet', open: 'native', save: 'native' },
  // Slides
  { id: 'pptx', label: 'PowerPoint Presentation', exts: ['pptx', 'pptm', 'potx', 'ppsx', 'potm', 'ppsm'], kind: 'slides', open: 'native', save: 'native' },
  { id: 'ppt', label: 'PowerPoint 97-2003', exts: ['ppt', 'pps', 'pot'], kind: 'slides', open: 'libreoffice', save: 'libreoffice' },
  { id: 'odp', label: 'OpenDocument Presentation', exts: ['odp', 'otp', 'fodp'], kind: 'slides', open: 'libreoffice', save: 'libreoffice' },
  { id: 'key', label: 'Apple Keynote', exts: ['key'], kind: 'slides', open: 'libreoffice', save: 'none' },
  { id: 'afslides', label: 'Affice Presentation', exts: ['afslides'], kind: 'slides', open: 'native', save: 'native' },
];

const byExt = new Map<string, FormatInfo>();
for (const f of FORMATS) for (const e of f.exts) byExt.set(e, f);

export function formatForExt(ext: string): FormatInfo | undefined {
  return byExt.get(ext.toLowerCase().replace(/^\./, ''));
}

export function kindForExt(ext: string): AppKind | undefined {
  return formatForExt(ext)?.kind;
}

export function formatById(id: string): FormatInfo | undefined {
  return FORMATS.find((f) => f.id === id);
}

export const KIND_LABEL: Record<AppKind, string> = {
  doc: 'Document',
  sheet: 'Spreadsheet',
  slides: 'Presentation',
};

export const KIND_APP: Record<AppKind, string> = {
  doc: 'Documents',
  sheet: 'Sheets',
  slides: 'Slides',
};

export function openFilters(kind?: AppKind, withLibreOffice = true): FileFilter[] {
  const list = FORMATS.filter((f) => (!kind || f.kind === kind) && (f.open === 'native' || (withLibreOffice && f.open === 'libreoffice')));
  const all = list.flatMap((f) => f.exts);
  const filters: FileFilter[] = [
    { name: kind ? `All ${KIND_APP[kind]} files` : 'All supported files', extensions: all },
    ...list.map((f) => ({ name: `${f.label} (${f.exts.map((e) => `.${e}`).join(', ')})`, extensions: f.exts })),
  ];
  if (!kind) {
    filters.splice(1, 0,
      { name: 'Documents', extensions: FORMATS.filter((f) => f.kind === 'doc' && f.open !== 'none').flatMap((f) => f.exts) },
      { name: 'Spreadsheets', extensions: FORMATS.filter((f) => f.kind === 'sheet' && f.open !== 'none').flatMap((f) => f.exts) },
      { name: 'Presentations', extensions: FORMATS.filter((f) => f.kind === 'slides' && f.open !== 'none').flatMap((f) => f.exts) },
    );
  }
  filters.push({ name: 'All files', extensions: ['*'] });
  return filters;
}

export interface ExportTarget {
  id: string;
  label: string;
  ext: string;
  description: string;
  requiresLibreOffice?: boolean;
  desktopOnly?: boolean;
}

export const DOC_SAVE_TARGETS: ExportTarget[] = [
  { id: 'docx', label: 'Word Document', ext: 'docx', description: 'Best for sharing with Microsoft Word users.' },
  { id: 'odt', label: 'OpenDocument Text', ext: 'odt', description: 'Open standard used by LibreOffice and others.' },
  { id: 'rtf', label: 'Rich Text Format', ext: 'rtf', description: 'Opens in almost every word processor, even very old ones.' },
  { id: 'doc', label: 'Word 97-2003', ext: 'doc', description: 'For very old versions of Word (needs LibreOffice).', requiresLibreOffice: true },
  { id: 'afdoc', label: 'Affice Document', ext: 'afdoc', description: 'Keeps every Affice feature exactly.' },
];

export const DOC_EXPORT_TARGETS: ExportTarget[] = [
  { id: 'pdf', label: 'PDF', ext: 'pdf', description: 'Fixed layout for printing and sharing, with bookmarks.' },
  { id: 'html', label: 'Web Page', ext: 'html', description: 'A single self-contained .html file.' },
  { id: 'md', label: 'Markdown', ext: 'md', description: 'Plain-text formatting for notes, wikis and code hosting.' },
  { id: 'epub', label: 'EPUB e-book', ext: 'epub', description: 'Reflowable e-book for readers and phones.' },
  { id: 'txt', label: 'Plain Text', ext: 'txt', description: 'Just the text.' },
];

export const SHEET_SAVE_TARGETS: ExportTarget[] = [
  { id: 'xlsx', label: 'Excel Workbook', ext: 'xlsx', description: 'Best for sharing with Microsoft Excel users.' },
  { id: 'ods', label: 'OpenDocument Spreadsheet', ext: 'ods', description: 'Open standard used by LibreOffice and others.' },
  { id: 'xls', label: 'Excel 97-2003', ext: 'xls', description: 'For very old versions of Excel.' },
  { id: 'afsheet', label: 'Affice Sheet', ext: 'afsheet', description: 'Keeps every Affice feature exactly (charts, rules…).' },
];

export const SHEET_EXPORT_TARGETS: ExportTarget[] = [
  { id: 'pdf', label: 'PDF', ext: 'pdf', description: 'Print-ready copy of the current sheet.' },
  { id: 'csv', label: 'CSV', ext: 'csv', description: 'Values of the current sheet, comma separated.' },
  { id: 'tsv', label: 'TSV', ext: 'tsv', description: 'Values of the current sheet, tab separated.' },
  { id: 'html', label: 'Web Page', ext: 'html', description: 'The current sheet as an HTML table.' },
  { id: 'json', label: 'JSON', ext: 'json', description: 'Rows as JSON objects, using the first row as keys.' },
  { id: 'md', label: 'Markdown table', ext: 'md', description: 'The current sheet as a Markdown table.' },
];

export const SLIDES_SAVE_TARGETS: ExportTarget[] = [
  { id: 'pptx', label: 'PowerPoint Presentation', ext: 'pptx', description: 'Best for sharing with Microsoft PowerPoint users.' },
  { id: 'odp', label: 'OpenDocument Presentation', ext: 'odp', description: 'For LibreOffice Impress (needs LibreOffice).', requiresLibreOffice: true },
  { id: 'afslides', label: 'Affice Presentation', ext: 'afslides', description: 'Keeps every Affice feature exactly (animations…).' },
];

export const SLIDES_EXPORT_TARGETS: ExportTarget[] = [
  { id: 'pdf', label: 'PDF', ext: 'pdf', description: 'One page per slide.' },
  { id: 'png', label: 'PNG image (current slide)', ext: 'png', description: 'High-resolution picture of the current slide.' },
  { id: 'png-all', label: 'PNG images (all slides, .zip)', ext: 'zip', description: 'Every slide as a picture, in a zip file.' },
  { id: 'html', label: 'Web presentation', ext: 'html', description: 'A self-contained .html slideshow that runs in any browser.' },
  { id: 'md', label: 'Outline (Markdown)', ext: 'md', description: 'Slide titles and text as a Markdown outline.' },
];
