import type { AppKind, FileFilter, PdfJob } from '@/shared/types';
import { api, isElectron } from '@/lib/platform';
import { formatById, formatForExt, type ExportTarget } from '@/lib/formats';
import { basename, extname, stripExt } from '@/lib/utils';
import { alertDialog } from '@/ui/dialog';
import { toast } from '@/ui/toast';
import { useRecent } from './actions';
import { getTab, useWorkspace } from './workspace';

export type Serializer = (formatId: string) => Promise<Uint8Array>;

const MODERN: Record<AppKind, string> = { doc: 'docx', sheet: 'xlsx', slides: 'pptx' };

async function produce(kind: AppKind, formatId: string, serialize: Serializer): Promise<Uint8Array> {
  const fmt = formatById(formatId);
  if (fmt?.save === 'libreoffice') {
    const lo = await api.convert.libreOfficePath();
    if (!lo) throw new Error(`Saving as ${fmt.label} needs LibreOffice, which isn't installed.`);
    const modern = await serialize(MODERN[kind]);
    return api.convert.withLibreOffice(modern, MODERN[kind], fmt.exts[0]);
  }
  return serialize(formatId);
}

function filtersFor(targets: ExportTarget[], first?: string): FileFilter[] {
  const list = [...targets];
  if (first) list.sort((a, b) => (a.id === first ? -1 : b.id === first ? 1 : 0));
  return list.map((t) => ({ name: `${t.label} (*.${t.ext})`, extensions: [t.ext] }));
}

export interface SaveOptions {
  tabId: string;
  kind: AppKind;
  targets: ExportTarget[];
  defaultFormat: string;
  serialize: Serializer;
  /** Force the Save As dialog, optionally preselecting a format. */
  saveAs?: boolean;
  formatId?: string;
}

/** Save / Save As shared by all editors. Returns true when the document was written. */
export async function saveDocument(o: SaveOptions): Promise<boolean> {
  const tab = getTab(o.tabId);
  if (!tab) return false;
  const current = tab.format ? formatById(tab.format) : undefined;
  const canWriteCurrent = Boolean(tab.path && current && current.save !== 'none' && (current.save === 'native' || current.save === 'libreoffice'));

  let path = tab.path;
  let formatId = tab.format;
  if (o.saveAs || !canWriteCurrent || !path || !isElectron) {
    const pref = o.formatId ?? (current && current.save !== 'none' && o.targets.some((t) => t.id === current.id) ? current.id : o.defaultFormat);
    const target = o.targets.find((t) => t.id === pref) ?? o.targets[0];
    const name = `${stripExt(tab.title)}.${target.ext}`;
    const chosen = await api.files.saveDialog({
      title: o.saveAs ? 'Save a copy' : 'Save',
      defaultPath: tab.path ? tab.path.replace(/[^\\/]*$/, name) : name,
      filters: filtersFor(o.targets, target.id),
    });
    if (!chosen) return false;
    let ext = extname(chosen.replace(/^download:/, ''));
    let t = o.targets.find((x) => x.ext === ext);
    if (!t) {
      t = target;
      ext = t.ext;
      path = `${chosen}.${ext}`;
    } else path = chosen;
    formatId = t.id;
    if (t.requiresLibreOffice && !(await api.convert.libreOfficePath())) {
      await alertDialog('LibreOffice needed', `Saving as ${t.label} requires LibreOffice. Install it for free from libreoffice.org, or choose another format.`);
      return false;
    }
  }

  const fmt = formatById(formatId!);
  const pending = toast.loading(`Saving ${basename(path!.replace(/^download:/, ''))}…`);
  try {
    const bytes = await produce(o.kind, formatId!, o.serialize);
    await api.files.write(path!, bytes);
    toast.dismiss(pending);
    const isDownload = path!.startsWith('download:');
    const title = basename(path!.replace(/^download:/, ''));
    useWorkspace.getState().updateTab(o.tabId, {
      dirty: false,
      title,
      path: isDownload ? tab.path : path,
      format: isDownload ? tab.format : formatId,
    });
    if (!isDownload) {
      void useRecent.getState().add({ path: path!, name: title, kind: o.kind, size: bytes.length });
      void api.recovery.remove(tab.recoveryId);
    }
    if (fmt?.lossy) toast.warning(`Saved as ${fmt.label}`, { detail: fmt.lossy, duration: 6000 });
    else toast.success(isDownload ? `Downloaded ${title}` : `Saved ${title}`, { duration: 2200 });
    return true;
  } catch (e) {
    toast.update(pending, 'error', 'Could not save the file', { detail: String((e as Error)?.message ?? e) });
    return false;
  }
}

/** Export to a file without changing the document's own path/format. */
export async function exportBytes(title: string, target: ExportTarget, bytes: Uint8Array | (() => Promise<Uint8Array>)): Promise<boolean> {
  const name = `${stripExt(title)}.${target.ext}`;
  const p = await api.files.saveDialog({ title: `Export as ${target.label}`, defaultPath: name, filters: [{ name: target.label, extensions: [target.ext] }] });
  if (!p) return false;
  const pending = toast.loading(`Exporting ${target.label}…`);
  try {
    const data = typeof bytes === 'function' ? await bytes() : bytes;
    await api.files.write(p, data);
    toast.update(pending, 'success', `Exported ${basename(p.replace(/^download:/, ''))}`, {
      duration: 4000,
      action: isElectron && !p.startsWith('download:') ? { label: 'Show', onClick: () => api.files.showInFolder(p) } : undefined,
    });
    return true;
  } catch (e) {
    toast.update(pending, 'error', `Export failed`, { detail: String((e as Error)?.message ?? e) });
    return false;
  }
}

/** PDF export via Chromium's print engine (desktop) or the print dialog (browser). */
export async function exportPdf(title: string, job: PdfJob): Promise<void> {
  if (!isElectron) {
    toast.info('Choose “Save as PDF” as the printer to create a PDF.', { duration: 6000 });
    await api.output.print(job);
    return;
  }
  const target: ExportTarget = { id: 'pdf', label: 'PDF', ext: 'pdf', description: '' };
  await exportBytes(title, target, () => api.output.pdf(job));
}

export async function printJob(job: PdfJob): Promise<void> {
  try {
    await api.output.print(job);
  } catch (e) {
    toast.error('Printing failed', { detail: String((e as Error)?.message ?? e) });
  }
}

export function markDirty(tabId: string, dirty = true): void {
  const t = getTab(tabId);
  if (t && t.dirty !== dirty) useWorkspace.getState().updateTab(tabId, { dirty });
}

export function formatLabelFor(tabFormat?: string): string {
  if (!tabFormat) return 'Not saved yet';
  return formatById(tabFormat)?.label ?? tabFormat.toUpperCase();
}

export { formatForExt };
