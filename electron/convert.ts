import { execFile } from 'node:child_process';
import fs from 'node:fs';
import fsp from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { pathToFileURL } from 'node:url';

let cachedSoffice: string | null | undefined;

function candidates(): string[] {
  const list: string[] = [];
  if (process.platform === 'win32') {
    for (const base of [process.env['ProgramFiles'], process.env['ProgramFiles(x86)'], 'C:\\Program Files', 'C:\\Program Files (x86)']) {
      if (base) list.push(path.join(base, 'LibreOffice', 'program', 'soffice.exe'));
    }
  } else {
    list.push(
      '/usr/bin/soffice',
      '/usr/local/bin/soffice',
      '/usr/lib/libreoffice/program/soffice',
      '/usr/lib64/libreoffice/program/soffice',
      '/opt/libreoffice/program/soffice',
      '/snap/bin/libreoffice',
      '/var/lib/flatpak/exports/bin/org.libreoffice.LibreOffice',
      path.join(os.homedir(), '.local/share/flatpak/exports/bin/org.libreoffice.LibreOffice'),
    );
    try {
      for (const d of fs.readdirSync('/opt')) {
        if (d.toLowerCase().startsWith('libreoffice')) list.push(path.join('/opt', d, 'program', 'soffice'));
      }
    } catch {
      /* no /opt */
    }
  }
  return list;
}

export function findLibreOffice(): string | null {
  if (cachedSoffice !== undefined) return cachedSoffice;
  cachedSoffice = candidates().find((c) => {
    try {
      fs.accessSync(c, fs.constants.X_OK);
      return true;
    } catch {
      return false;
    }
  }) ?? null;
  return cachedSoffice;
}

const safeExt = (e: string) => e.replace(/[^a-z0-9]/gi, '').toLowerCase();

/** Converts a file with a headless LibreOffice using a throw-away profile (so a running LibreOffice is not disturbed). */
export async function convertWithLibreOffice(data: Uint8Array, fromExt: string, toExt: string): Promise<Uint8Array> {
  const soffice = findLibreOffice();
  if (!soffice) throw new Error('LibreOffice is not installed.');
  const work = await fsp.mkdtemp(path.join(os.tmpdir(), 'affice-convert-'));
  const profile = path.join(os.tmpdir(), 'affice-lo-profile');
  const input = path.join(work, `input.${safeExt(fromExt)}`);
  const outDir = path.join(work, 'out');
  try {
    await fsp.mkdir(outDir, { recursive: true });
    await fsp.writeFile(input, data);
    await new Promise<void>((resolve, reject) => {
      execFile(
        soffice,
        [
          `-env:UserInstallation=${pathToFileURL(profile).href}`,
          '--headless',
          '--norestore',
          '--nolockcheck',
          '--nodefault',
          '--convert-to',
          safeExt(toExt),
          '--outdir',
          outDir,
          input,
        ],
        { timeout: 120_000, windowsHide: true },
        (err) => (err ? reject(err) : resolve()),
      );
    });
    const produced = (await fsp.readdir(outDir)).find((f) => f.toLowerCase().endsWith(`.${safeExt(toExt)}`));
    if (!produced) throw new Error('LibreOffice did not produce an output file.');
    return new Uint8Array(await fsp.readFile(path.join(outDir, produced)));
  } finally {
    fsp.rm(work, { recursive: true, force: true }).catch(() => undefined);
  }
}

/** Plain text extraction for legacy Word 97-2003 (.doc) files when LibreOffice is unavailable. */
export async function legacyDocText(data: Uint8Array): Promise<string> {
  // word-extractor is CommonJS; esbuild bundles it into the main process.
  const mod = await import('word-extractor');
  const WordExtractor = (mod as unknown as { default: new () => { extract(b: Buffer): Promise<{ getBody(): string; getHeaders?(): string; getFootnotes?(): string }> } }).default;
  const extractor = new WordExtractor();
  const doc = await extractor.extract(Buffer.from(data));
  return doc.getBody();
}
