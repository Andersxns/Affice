import { app } from 'electron';
import fs from 'node:fs';
import fsp from 'node:fs/promises';
import path from 'node:path';
import { DEFAULT_SETTINGS, type RecentFile, type RecoveryEntry, type Settings } from '../src/shared/types';

const MAX_RECENT = 60;

function dataDir(): string {
  return app.getPath('userData');
}

function readJsonSync<T>(file: string, fallback: T): T {
  try {
    const raw = fs.readFileSync(file, 'utf8');
    return JSON.parse(raw) as T;
  } catch {
    return fallback;
  }
}

async function writeJson(file: string, value: unknown): Promise<void> {
  await fsp.mkdir(path.dirname(file), { recursive: true });
  const tmp = `${file}.${process.pid}.tmp`;
  await fsp.writeFile(tmp, JSON.stringify(value, null, 2), 'utf8');
  try {
    await fsp.rename(tmp, file);
  } catch {
    await fsp.writeFile(file, JSON.stringify(value, null, 2), 'utf8');
    await fsp.rm(tmp, { force: true });
  }
}

/* ---------------------------------------------------------------- settings */

let settingsCache: Settings | null = null;
const settingsFile = () => path.join(dataDir(), 'settings.json');

export function getSettings(): Settings {
  if (!settingsCache) {
    const stored = readJsonSync<Partial<Settings>>(settingsFile(), {});
    settingsCache = { ...DEFAULT_SETTINGS, ...stored };
  }
  return settingsCache;
}

export async function updateSettings(patch: Partial<Settings>): Promise<Settings> {
  const next: Settings = { ...getSettings(), ...patch };
  settingsCache = next;
  await writeJson(settingsFile(), next);
  return next;
}

/* ------------------------------------------------------------ recent files */

const recentFile = () => path.join(dataDir(), 'recent.json');
let recentCache: RecentFile[] | null = null;

export function listRecent(): RecentFile[] {
  if (!recentCache) recentCache = readJsonSync<RecentFile[]>(recentFile(), []);
  return recentCache;
}

async function saveRecent(list: RecentFile[]): Promise<RecentFile[]> {
  recentCache = list;
  await writeJson(recentFile(), list);
  return list;
}

const samePath = (a: string, b: string) =>
  process.platform === 'win32' ? a.toLowerCase() === b.toLowerCase() : a === b;

export async function addRecent(entry: Omit<RecentFile, 'openedAt'>): Promise<RecentFile[]> {
  const existing = listRecent().find((r) => samePath(r.path, entry.path));
  const rest = listRecent().filter((r) => !samePath(r.path, entry.path));
  const item: RecentFile = { ...entry, pinned: entry.pinned ?? existing?.pinned, openedAt: Date.now() };
  const next = [item, ...rest];
  // keep pinned entries no matter what, trim the rest
  const pinned = next.filter((r) => r.pinned);
  const unpinned = next.filter((r) => !r.pinned).slice(0, Math.max(0, MAX_RECENT - pinned.length));
  const merged = next.filter((r) => r.pinned || unpinned.includes(r));
  return saveRecent(merged);
}

export async function removeRecent(p: string): Promise<RecentFile[]> {
  return saveRecent(listRecent().filter((r) => !samePath(r.path, p)));
}

export async function pinRecent(p: string, pinned: boolean): Promise<RecentFile[]> {
  return saveRecent(listRecent().map((r) => (samePath(r.path, p) ? { ...r, pinned } : r)));
}

export async function clearRecent(): Promise<RecentFile[]> {
  return saveRecent(listRecent().filter((r) => r.pinned));
}

/* ------------------------------------------------------------ window state */

export interface WindowState {
  x?: number;
  y?: number;
  width: number;
  height: number;
  maximized: boolean;
}

const windowStateFile = () => path.join(dataDir(), 'window-state.json');

export function loadWindowState(): WindowState {
  return readJsonSync<WindowState>(windowStateFile(), { width: 1320, height: 860, maximized: false });
}

export function saveWindowStateSync(state: WindowState): void {
  try {
    fs.mkdirSync(dataDir(), { recursive: true });
    fs.writeFileSync(windowStateFile(), JSON.stringify(state));
  } catch {
    /* non-fatal */
  }
}

/* ---------------------------------------------------------------- recovery */

const recoveryDir = () => path.join(dataDir(), 'recovery');
const safeId = (id: string) => id.replace(/[^a-zA-Z0-9_-]/g, '');

export async function saveRecovery(entry: RecoveryEntry, data: string): Promise<void> {
  await fsp.mkdir(recoveryDir(), { recursive: true });
  const file = path.join(recoveryDir(), `${safeId(entry.id)}.json`);
  await writeJson(file, { entry: { ...entry, size: data.length }, data });
}

export async function listRecovery(): Promise<RecoveryEntry[]> {
  try {
    const files = await fsp.readdir(recoveryDir());
    const out: RecoveryEntry[] = [];
    for (const f of files) {
      if (!f.endsWith('.json')) continue;
      try {
        const raw = JSON.parse(await fsp.readFile(path.join(recoveryDir(), f), 'utf8'));
        if (raw?.entry?.id) out.push(raw.entry as RecoveryEntry);
      } catch {
        /* skip corrupt entries */
      }
    }
    return out.sort((a, b) => b.savedAt - a.savedAt);
  } catch {
    return [];
  }
}

export async function loadRecovery(id: string): Promise<string | null> {
  try {
    const raw = JSON.parse(await fsp.readFile(path.join(recoveryDir(), `${safeId(id)}.json`), 'utf8'));
    return typeof raw?.data === 'string' ? raw.data : null;
  } catch {
    return null;
  }
}

export async function removeRecovery(id: string): Promise<void> {
  await fsp.rm(path.join(recoveryDir(), `${safeId(id)}.json`), { force: true });
}
