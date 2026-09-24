import { create } from 'zustand';
import { api } from '@/lib/platform';
import { DEFAULT_SETTINGS, type Settings } from '@/shared/types';

interface SettingsState {
  settings: Settings;
  loaded: boolean;
  systemDark: boolean;
  load(): Promise<void>;
  update(patch: Partial<Settings>): Promise<void>;
}

const mq = typeof window !== 'undefined' ? window.matchMedia('(prefers-color-scheme: dark)') : null;
const reducedMq = typeof window !== 'undefined' ? window.matchMedia('(prefers-reduced-motion: reduce)') : null;

export const useSettings = create<SettingsState>((set, get) => ({
  settings: DEFAULT_SETTINGS,
  loaded: false,
  systemDark: mq?.matches ?? false,
  async load() {
    const s = await api.settings.get();
    set({ settings: { ...DEFAULT_SETTINGS, ...s }, loaded: true });
    applySettings(get().settings, get().systemDark);
  },
  async update(patch) {
    const optimistic = { ...get().settings, ...patch };
    set({ settings: optimistic });
    applySettings(optimistic, get().systemDark);
    const saved = await api.settings.set(patch);
    set({ settings: { ...DEFAULT_SETTINGS, ...saved } });
  },
}));

mq?.addEventListener('change', (e) => {
  useSettings.setState({ systemDark: e.matches });
  applySettings(useSettings.getState().settings, e.matches);
});

export function isDark(s: Settings, systemDark: boolean): boolean {
  return s.theme === 'dark' || (s.theme === 'system' && systemDark);
}

export function useIsDark(): boolean {
  return useSettings((s) => isDark(s.settings, s.systemDark));
}

function shade(hex: string, amt: number): string {
  const h = hex.replace('#', '');
  const n = parseInt(h.length === 3 ? h.split('').map((c) => c + c).join('') : h, 16);
  const r = (n >> 16) & 255;
  const g = (n >> 8) & 255;
  const b = n & 255;
  const f = (c: number) => Math.round(amt < 0 ? c * (1 + amt) : c + (255 - c) * amt);
  return `#${[f(r), f(g), f(b)].map((c) => c.toString(16).padStart(2, '0')).join('')}`;
}

function rgba(hex: string, a: number): string {
  const h = hex.replace('#', '');
  const n = parseInt(h, 16);
  return `rgba(${(n >> 16) & 255}, ${(n >> 8) & 255}, ${n & 255}, ${a})`;
}

export function applySettings(s: Settings, systemDark: boolean): void {
  const root = document.documentElement;
  const dark = isDark(s, systemDark);
  root.dataset.theme = dark ? 'dark' : 'light';
  root.dataset.motion = reducedMq?.matches && s.motion === 'full' ? 'reduced' : s.motion;
  root.dataset.density = s.density;
  root.dataset.perf = s.performanceMode ? 'on' : 'off';

  const accent = /^#[0-9a-f]{6}$/i.test(s.accent) ? s.accent : DEFAULT_SETTINGS.accent;
  root.style.setProperty('--accent', accent);
  root.style.setProperty('--accent-hover', shade(accent, -0.12));
  root.style.setProperty('--accent-press', shade(accent, -0.25));
  root.style.setProperty('--accent-soft', rgba(accent, dark ? 0.18 : 0.1));
  root.style.setProperty('--accent-soft-2', rgba(accent, dark ? 0.26 : 0.16));
  root.style.setProperty('--accent-text', dark ? shade(accent, 0.45) : shade(accent, -0.1));
  root.style.setProperty('--focus-ring', `0 0 0 3px ${rgba(accent, dark ? 0.45 : 0.3)}`);
  root.style.setProperty('--selection', rgba(accent, dark ? 0.35 : 0.22));

  try {
    localStorage.setItem('affice.boot', JSON.stringify({ theme: s.theme }));
  } catch {
    /* ignore */
  }
  api.window.setTitleBarColors(dark ? '#15171c' : '#eef1f8', dark ? '#e8eaf0' : '#1d2433');
}
