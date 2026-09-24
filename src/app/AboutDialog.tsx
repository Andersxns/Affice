import { ExternalLink, Heart } from 'lucide-react';
import { useEffect, useState } from 'react';
import { api } from '@/lib/platform';
import { Modal, openDialog } from '@/ui/dialog';
import type { AppInfo } from '@/shared/types';
import { appIcon, GithubIcon } from './AppIcon';

export const REPO_URL = 'https://github.com/andersxns/affice';

export function openAbout() {
  void openDialog((close) => <AboutDialog onClose={() => close()} />);
}

function AboutDialog({ onClose }: { onClose: () => void }) {
  const [info, setInfo] = useState<AppInfo | null>(null);
  useEffect(() => {
    void api.app.info().then(setInfo);
  }, []);
  return (
    <Modal title="About Affice" onClose={onClose} width={480}>
      <div className="about">
        <img src={appIcon} width={88} height={88} alt="Affice" className="about-logo" />
        <h2 className="about-name">Affice</h2>
        <p className="about-tag">Documents, Sheets and Slides — free and open source, forever.</p>
        <div className="about-grid selectable">
          <span>Version</span>
          <span>{info?.version ?? '…'}</span>
          <span>Platform</span>
          <span>
            {info?.platform} {info?.arch}
          </span>
          <span>Engine</span>
          <span>
            Electron {info?.electron} · Chromium {info?.chrome}
          </span>
          <span>LibreOffice</span>
          <span>{info?.libreOffice ? 'Available for legacy formats' : 'Not installed (optional)'}</span>
          <span>Data folder</span>
          <span className="ellipsis" title={info?.userData}>
            {info?.userData}
          </span>
        </div>
        <p className="about-license">
          Released under the MIT License. Affice bundles open-source fonts (SIL Open Font License) and libraries — see the project page for credits.
        </p>
        <div className="row" style={{ justifyContent: 'center', marginTop: 6 }}>
          <button type="button" className="btn" onClick={() => api.app.openExternal(REPO_URL)}>
            <GithubIcon size={15} /> Source code
          </button>
          <button type="button" className="btn" onClick={() => api.app.openExternal(`${REPO_URL}/issues`)}>
            <ExternalLink size={15} /> Report an issue
          </button>
        </div>
        <p className="about-heart">
          Made with <Heart size={12} fill="currentColor" /> by the Affice community
        </p>
      </div>
    </Modal>
  );
}

const SHORTCUTS: Array<[string, Array<[string, string]>]> = [
  [
    'Everywhere',
    [
      ['New document', 'Ctrl+N'],
      ['Open', 'Ctrl+O'],
      ['Save', 'Ctrl+S'],
      ['Save as', 'Ctrl+Shift+S'],
      ['Print', 'Ctrl+P'],
      ['Close tab', 'Ctrl+W'],
      ['Next / previous tab', 'Ctrl+Tab / Ctrl+Shift+Tab'],
      ['Search commands', 'Ctrl+Shift+P or Alt+Q'],
      ['Settings', 'Ctrl+,'],
      ['Full screen', 'F11'],
      ['Undo / redo', 'Ctrl+Z / Ctrl+Y'],
    ],
  ],
  [
    'Documents',
    [
      ['Bold / italic / underline', 'Ctrl+B / I / U'],
      ['Headings 1–3', 'Ctrl+Alt+1…3'],
      ['Normal text', 'Ctrl+Alt+0'],
      ['Align left / centre / right / justify', 'Ctrl+L / E / R / J'],
      ['Bulleted / numbered list', 'Ctrl+Shift+8 / 7'],
      ['Insert link', 'Ctrl+K'],
      ['Find / replace', 'Ctrl+F / Ctrl+H'],
      ['Page break', 'Ctrl+Enter'],
      ['Focus mode', 'F9'],
    ],
  ],
  [
    'Sheets',
    [
      ['Edit cell', 'F2'],
      ['Fill down / right', 'Ctrl+D / Ctrl+R'],
      ['Jump to edge of data', 'Ctrl+Arrow'],
      ['Select row / column', 'Shift+Space / Ctrl+Space'],
      ['AutoSum', 'Alt+='],
      ['Insert today’s date', 'Ctrl+;'],
      ['Go to cell', 'Ctrl+G'],
      ['Find / replace', 'Ctrl+F / Ctrl+H'],
    ],
  ],
  [
    'Slides',
    [
      ['New slide', 'Ctrl+M'],
      ['Duplicate', 'Ctrl+D'],
      ['Start slideshow', 'F5'],
      ['From current slide', 'Shift+F5'],
      ['Group / ungroup', 'Ctrl+G / Ctrl+Shift+G'],
      ['Nudge', 'Arrow keys (Shift for bigger steps)'],
      ['During show: laser / pen / black screen', 'L / P / B'],
    ],
  ],
];

export function openShortcuts() {
  void openDialog((close) => (
    <Modal title="Keyboard shortcuts" onClose={() => close()} width={680}>
      <div className="shortcuts-grid">
        {SHORTCUTS.map(([group, rows]) => (
          <div key={group} className="shortcuts-group">
            <h4>{group}</h4>
            {rows.map(([label, keys]) => (
              <div key={label} className="shortcut-row">
                <span>{label}</span>
                <span className="kbd">{keys}</span>
              </div>
            ))}
          </div>
        ))}
      </div>
    </Modal>
  ));
}
