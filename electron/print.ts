import { BrowserWindow, ipcMain, type WebContents } from 'electron';
import type { PdfJob } from '../src/shared/types';

type Loader = (win: BrowserWindow, route: string) => Promise<void>;

const jobs = new Map<number, PdfJob>();
const readyWaiters = new Map<number, () => void>();

export function registerPrintIpc(preloadPath: string, load: Loader): void {
  ipcMain.handle('output:take-job', (e) => {
    const job = jobs.get(e.sender.id) ?? null;
    return job;
  });
  ipcMain.on('output:ready', (e) => {
    readyWaiters.get(e.sender.id)?.();
  });

  async function prepare(job: PdfJob): Promise<BrowserWindow> {
    const win = new BrowserWindow({
      show: false,
      width: Math.round(job.pageWidthIn * 96) + 40,
      height: Math.round(job.pageHeightIn * 96) + 40,
      webPreferences: {
        preload: preloadPath,
        contextIsolation: true,
        nodeIntegration: false,
        sandbox: true,
        spellcheck: false,
      },
    });
    const id = win.webContents.id;
    jobs.set(id, job);
    const ready = new Promise<void>((resolve, reject) => {
      const timer = setTimeout(() => reject(new Error('Timed out while preparing the document for printing.')), 90_000);
      readyWaiters.set(id, () => {
        clearTimeout(timer);
        resolve();
      });
    });
    win.on('closed', () => {
      jobs.delete(id);
      readyWaiters.delete(id);
    });
    await load(win, 'print');
    await ready;
    return win;
  }

  ipcMain.handle('output:pdf', async (_e, job: PdfJob) => {
    const win = await prepare(job);
    try {
      const m = job.marginsIn ?? { top: 0, right: 0, bottom: 0, left: 0 };
      const hasHF = Boolean(job.headerHtml || job.footerHtml);
      const buf = await win.webContents.printToPDF({
        printBackground: true,
        preferCSSPageSize: false,
        pageSize: { width: job.pageWidthIn, height: job.pageHeightIn },
        landscape: false,
        margins: { top: m.top, right: m.right, bottom: m.bottom, left: m.left },
        displayHeaderFooter: hasHF,
        headerTemplate: job.headerHtml || '<span></span>',
        footerTemplate: job.footerHtml || '<span></span>',
        generateDocumentOutline: true,
        generateTaggedPDF: true,
      });
      return new Uint8Array(buf);
    } finally {
      win.destroy();
    }
  });

  ipcMain.handle('output:print', async (_e, job: PdfJob) => {
    const win = await prepare(job);
    return new Promise<boolean>((resolve) => {
      const wc: WebContents = win.webContents;
      wc.print({ silent: false, printBackground: true }, (success) => {
        resolve(success);
        setTimeout(() => win.destroy(), 500);
      });
    });
  });
}
