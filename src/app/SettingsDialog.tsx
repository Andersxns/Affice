import { Check, FileCog, Gauge, Palette, PenLine, RotateCcw } from 'lucide-react';
import { useEffect, useState, type ReactNode } from 'react';
import { api, isElectron } from '@/lib/platform';
import { Checkbox, Segmented, Select, Switch } from '@/ui/controls';
import { Modal, openDialog } from '@/ui/dialog';
import { toast } from '@/ui/toast';
import type { Settings } from '@/shared/types';
import { useSettings } from './settings';

const ACCENTS = ['#004fff', '#2f6dff', '#7c3aed', '#db2777', '#e5484d', '#f26a26', '#d97706', '#17a35a', '#0d9488', '#0891b2', '#475569'];

type Section = 'appearance' | 'performance' | 'editing' | 'files';

export function openSettings(section: Section = 'appearance') {
  void openDialog((close) => <SettingsDialog onClose={() => close()} initial={section} />);
}

function Row({ title, desc, children }: { title: ReactNode; desc?: ReactNode; children: ReactNode }) {
  return (
    <div className="set-row">
      <div className="set-text">
        <div className="set-title">{title}</div>
        {desc && <div className="set-desc">{desc}</div>}
      </div>
      <div className="set-control">{children}</div>
    </div>
  );
}

function SettingsDialog({ onClose, initial }: { onClose: () => void; initial: Section }) {
  const s = useSettings((st) => st.settings);
  const update = useSettings((st) => st.update);
  const [section, setSection] = useState<Section>(initial);
  const [needsRestart, setNeedsRestart] = useState(false);
  const [lo, setLo] = useState<string | null | undefined>(undefined);
  const [langs, setLangs] = useState<string[]>([]);

  useEffect(() => {
    void api.convert.libreOfficePath().then(setLo);
    void api.spell.languages().then(setLangs);
  }, []);

  const set = <K extends keyof Settings>(k: K, v: Settings[K], restart = false) => {
    void update({ [k]: v } as Partial<Settings>);
    if (restart) setNeedsRestart(true);
  };

  const nav: Array<{ id: Section; label: string; icon: ReactNode }> = [
    { id: 'appearance', label: 'Appearance', icon: <Palette size={16} /> },
    { id: 'performance', label: 'Performance', icon: <Gauge size={16} /> },
    { id: 'editing', label: 'Editing', icon: <PenLine size={16} /> },
    { id: 'files', label: 'Files & formats', icon: <FileCog size={16} /> },
  ];

  return (
    <Modal title="Settings" onClose={onClose} width={760} className="settings-modal" bodyClassName="settings-body">
      <nav className="settings-nav">
        {nav.map((n) => (
          <button key={n.id} type="button" className={`settings-nav-item${section === n.id ? ' active' : ''}`} onClick={() => setSection(n.id)}>
            {n.icon}
            {n.label}
          </button>
        ))}
      </nav>
      <div className="settings-content thin-scroll" key={section}>
        {section === 'appearance' && (
          <>
            <h3 className="set-h">Appearance</h3>
            <Row title="Theme" desc="Follow your system, or pick light or dark.">
              <Segmented
                value={s.theme}
                onChange={(v) => set('theme', v)}
                options={[
                  { value: 'system', label: 'System' },
                  { value: 'light', label: 'Light' },
                  { value: 'dark', label: 'Dark' },
                ]}
              />
            </Row>
            <Row title="Accent colour" desc="Used for highlights, selections and buttons.">
              <div className="accent-row">
                {ACCENTS.map((c) => (
                  <button key={c} type="button" className={`accent-dot${s.accent === c ? ' on' : ''}`} style={{ background: c }} aria-label={c} onClick={() => set('accent', c)}>
                    {s.accent === c && <Check size={13} color="#fff" strokeWidth={3} />}
                  </button>
                ))}
              </div>
            </Row>
            <Row title="Density" desc="Compact fits more on small or low-resolution screens.">
              <Segmented
                value={s.density}
                onChange={(v) => set('density', v)}
                options={[
                  { value: 'comfortable', label: 'Comfortable' },
                  { value: 'compact', label: 'Compact' },
                ]}
              />
            </Row>
            <Row title="Interface size" desc="Scale all menus and toolbars.">
              <Select
                value={String(s.uiScale)}
                onChange={(v) => {
                  set('uiScale', Number(v));
                  if (!isElectron) api.app.zoom(Number(v));
                }}
                width={120}
                options={['0.8', '0.9', '1', '1.1', '1.25', '1.5'].map((v) => ({ value: v, label: `${Math.round(Number(v) * 100)}%` }))}
              />
            </Row>
            <Row title="Animations" desc="Smooth transitions throughout the app.">
              <Segmented
                value={s.motion}
                onChange={(v) => set('motion', v)}
                options={[
                  { value: 'full', label: 'Full' },
                  { value: 'reduced', label: 'Reduced' },
                  { value: 'off', label: 'Off' },
                ]}
              />
            </Row>
            <Row title="Ribbon style" desc="Classic shows labelled groups; simplified is a single compact row.">
              <Segmented
                value={s.ribbonMode}
                onChange={(v) => set('ribbonMode', v)}
                options={[
                  { value: 'full', label: 'Classic' },
                  { value: 'simplified', label: 'Simplified' },
                ]}
              />
            </Row>
          </>
        )}
        {section === 'performance' && (
          <>
            <h3 className="set-h">Performance & older computers</h3>
            <Row title="Performance mode" desc="Turns off shadows, blur and other effects that are slow on older graphics cards. Recommended for older or low-power PCs.">
              <Switch checked={s.performanceMode} onChange={(v) => set('performanceMode', v)} label="Performance mode" />
            </Row>
            <Row title="Hardware acceleration" desc="Uses your graphics card to draw the app. Turn off if you see flickering, black windows or glitches. Requires restart.">
              <Switch checked={s.hardwareAcceleration} onChange={(v) => set('hardwareAcceleration', v, true)} label="Hardware acceleration" disabled={!isElectron} />
            </Row>
            <Row title="Use system title bar" desc="Show the operating system's own window frame (helpful on some Linux desktops). Requires restart.">
              <Switch checked={s.nativeTitleBar} onChange={(v) => set('nativeTitleBar', v, true)} label="System title bar" disabled={!isElectron} />
            </Row>
            <Row title="Quick setup for an old computer" desc="Performance mode on, animations off, compact layout.">
              <button
                type="button"
                className="btn"
                onClick={() => {
                  void update({ performanceMode: true, motion: 'off', density: 'compact' });
                  toast.success('Optimised for older hardware');
                }}
              >
                Optimise
              </button>
            </Row>
          </>
        )}
        {section === 'editing' && (
          <>
            <h3 className="set-h">Editing</h3>
            <Row title="Your name" desc="Used as the author of new documents and comments.">
              <input className="input" style={{ width: 220 }} value={s.authorName} placeholder="Your name" onChange={(e) => set('authorName', e.target.value)} />
            </Row>
            <Row title="Spell check" desc="Underline misspelled words while you type (right-click a word for suggestions).">
              <Switch checked={s.spellcheck} onChange={(v) => set('spellcheck', v)} label="Spell check" />
            </Row>
            {isElectron && langs.length > 0 && (
              <Row title="Spelling languages" desc="Choose up to 3 languages.">
                <div className="lang-list thin-scroll">
                  {langs.map((l) => (
                    <Checkbox
                      key={l}
                      label={l}
                      checked={s.spellcheckLanguages.includes(l)}
                      onChange={(on) => {
                        const next = on ? [...s.spellcheckLanguages, l].slice(-3) : s.spellcheckLanguages.filter((x) => x !== l);
                        set('spellcheckLanguages', next);
                      }}
                    />
                  ))}
                </div>
              </Row>
            )}
            <Row title="Crash protection" desc="Keep a recovery copy of unsaved work every few minutes.">
              <Select
                value={String(s.autosaveMinutes)}
                onChange={(v) => set('autosaveMinutes', Number(v))}
                width={150}
                options={[
                  { value: '0', label: 'Off' },
                  { value: '1', label: 'Every minute' },
                  { value: '2', label: 'Every 2 minutes' },
                  { value: '5', label: 'Every 5 minutes' },
                  { value: '10', label: 'Every 10 minutes' },
                ]}
              />
            </Row>
            <Row title="AutoSave to file" desc="Automatically save changes to files that are already saved on your computer.">
              <Switch checked={s.autoSaveToFile} onChange={(v) => set('autoSaveToFile', v)} label="AutoSave" />
            </Row>
            <Row title="Ask before closing unsaved files">
              <Switch checked={s.confirmOnClose} onChange={(v) => set('confirmOnClose', v)} label="Confirm on close" />
            </Row>
          </>
        )}
        {section === 'files' && (
          <>
            <h3 className="set-h">Files & formats</h3>
            <Row title="Documents save as" desc="Default format for new documents.">
              <Select
                value={s.defaultDocFormat}
                onChange={(v) => set('defaultDocFormat', v)}
                width={210}
                options={[
                  { value: 'docx', label: 'Word Document (.docx)' },
                  { value: 'odt', label: 'OpenDocument (.odt)' },
                  { value: 'afdoc', label: 'Affice Document (.afdoc)' },
                ]}
              />
            </Row>
            <Row title="Sheets save as">
              <Select
                value={s.defaultSheetFormat}
                onChange={(v) => set('defaultSheetFormat', v)}
                width={210}
                options={[
                  { value: 'xlsx', label: 'Excel Workbook (.xlsx)' },
                  { value: 'ods', label: 'OpenDocument (.ods)' },
                  { value: 'afsheet', label: 'Affice Sheet (.afsheet)' },
                ]}
              />
            </Row>
            <Row title="Slides save as">
              <Select
                value={s.defaultSlidesFormat}
                onChange={(v) => set('defaultSlidesFormat', v)}
                width={210}
                options={[
                  { value: 'pptx', label: 'PowerPoint (.pptx)' },
                  { value: 'afslides', label: 'Affice Presentation (.afslides)' },
                ]}
              />
            </Row>
            <Row
              title="LibreOffice integration"
              desc={
                lo === undefined
                  ? 'Checking…'
                  : lo
                    ? `Found at ${lo}. Affice uses it to open older formats like .doc, .ppt, .odp, WordPerfect and Keynote.`
                    : 'Not found. Install LibreOffice (free) to open older formats like .ppt, .odp, WordPerfect and Keynote.'
              }
            >
              <Switch checked={s.useLibreOfficeForLegacy} onChange={(v) => set('useLibreOfficeForLegacy', v)} label="Use LibreOffice" disabled={!lo} />
            </Row>
          </>
        )}
        {needsRestart && (
          <div className="restart-banner anim-rise">
            <RotateCcw size={16} />
            <span>Some changes take effect after restarting Affice.</span>
            <span className="spacer" />
            <button type="button" className="btn btn-primary btn-sm" onClick={() => api.app.relaunch()}>
              Restart now
            </button>
          </div>
        )}
      </div>
    </Modal>
  );
}
