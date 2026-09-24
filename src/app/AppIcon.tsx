import { FileText, House, Presentation, Table } from 'lucide-react';
import type { TabKind } from './workspace';
import logoMark from '@/assets/logo-mark.png';
import logoMarkWhite from '@/assets/logo-mark-white.png';
import appIcon from '@/assets/app-icon.png';

export { logoMark, logoMarkWhite, appIcon };

const GLYPH = {
  doc: FileText,
  sheet: Table,
  slides: Presentation,
  home: House,
} as const;

/** Rounded gradient square with a white glyph — the icon for each Affice app. */
export function AppIcon({ kind, size = 18, className = '' }: { kind: TabKind; size?: number; className?: string }) {
  const Glyph = GLYPH[kind];
  if (kind === 'home') {
    return <img src={logoMark} width={size} height={size} alt="" className={`app-icon-img ${className}`} draggable={false} />;
  }
  return (
    <span className={`app-icon app-icon-${kind} ${className}`} style={{ width: size, height: size, borderRadius: Math.round(size * 0.28) }}>
      <Glyph size={Math.round(size * 0.62)} strokeWidth={2.2} color="#fff" />
    </span>
  );
}

export function GithubIcon({ size = 16 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 16 16" fill="currentColor" aria-hidden>
      <path d="M8 0C3.58 0 0 3.58 0 8c0 3.54 2.29 6.53 5.47 7.59.4.07.55-.17.55-.38 0-.19-.01-.82-.01-1.49-2.01.37-2.53-.49-2.69-.94-.09-.23-.48-.94-.82-1.13-.28-.15-.68-.52-.01-.53.63-.01 1.08.58 1.23.82.72 1.21 1.87.87 2.33.66.07-.52.28-.87.51-1.07-1.78-.2-3.64-.89-3.64-3.95 0-.87.31-1.59.82-2.15-.08-.2-.36-1.02.08-2.12 0 0 .67-.21 2.2.82.64-.18 1.32-.27 2-.27.68 0 1.36.09 2 .27 1.53-1.04 2.2-.82 2.2-.82.44 1.1.16 1.92.08 2.12.51.56.82 1.27.82 2.15 0 3.07-1.87 3.75-3.65 3.95.29.25.54.73.54 1.48 0 1.07-.01 1.93-.01 2.2 0 .21.15.46.55.38A8.013 8.013 0 0016 8c0-4.42-3.58-8-8-8z" />
    </svg>
  );
}
