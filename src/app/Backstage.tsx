import { ArrowLeft, Copy, Download, FilePlus2, FolderOpen, Info, Printer, Save, SaveAll, X } from 'lucide-react';
import { useState, type ReactNode } from 'react';
import type { AppKind } from '@/shared/types';
import { api, isElectron } from '@/lib/platform';
import { KIND_LABEL, type ExportTarget } from '@/lib/formats';
import { timeAgo } from '@/lib/utils';
import { TEMPLATES } from '@/templates/catalog';
import { TemplatePreview } from '@/home/TemplatePreview';
import { closeTab, newDocument, openPaths, showOpenDialog, useRecent } from './actions';
import { AppIcon } from './AppIcon';
import { formatLabelFor } from './fileOps';
import type { Tab } from './workspace';

type Section = 'info' | 'new' | 'open' | 'saveas' | 'export' | 'print';

export interface BackstageProps {
  app: AppKind;
  tab: Tab;
  onClose: () => void;
  onSave: () => void;
  onSaveAs: (formatId: string) => void;
  onExport: (targetId: string) => void;
  onPrint: () => void;
  saveTargets: ExportTarget[];
  exportTargets: ExportTarget[];
  info?: Array<{ label: string; value: string }>;
  infoExtra?: ReactNode;
  printExtra?: ReactNode;
  libreOffice?: boolean;
}

export function Backstage(p: BackstageProps) {
  const [section, setSection] = useState<Section>('info');
  const recent = useRecent((s) => s.items).filter((r) => r.kind === p.app).slice(0, 12);
  const run = (fn: () => void) => {
    p.onClose();
    setTimeout(fn, 30);
  };
  const Item = ({ id, icon, label, onClick }: { id?: Section; icon: ReactNode; label: string; onClick?: () => void }) => (
    <button type="button" className={`backstage-item${id && section === id ? ' active' : ''}`} onClick={onClick ?? (() => id && setSection(id))}>
      {icon}
      {label}
    </button>
  );

  return (
    <div className={`backstage app-${p.app}`} onKeyDown={(e) => e.key === 'Escape' && p.onClose()} tabIndex={-1}>
      <nav className="backstage-nav">
        <button type="button" className="backstage-back" onClick={p.onClose}>
          <ArrowLeft size={17} /> Back
        </button>
        <Item id="info" icon={<Info size={17} />} label="Info" />
        <Item id="new" icon={<FilePlus2 size={17} />} label="New" />
        <Item id="open" icon={<FolderOpen size={17} />} label="Open" />
        <Item icon={<Save size={17} />} label="Save" onClick={() => run(p.onSave)} />
        <Item id="saveas" icon={<SaveAll size={17} />} label="Save as" />
        <Item id="export" icon={<Download size={17} />} label="Export" />
        <Item id="print" icon={<Printer size={17} />} label="Print" />
        <span className="spacer" />
        <Item icon={<X size={17} />} label="Close" onClick={() => run(() => void closeTab(p.tab.id))} />
      </nav>
      <div className="backstage-content thin-scroll">
        {section === 'info' && (
          <div className="bs-panel" key="info">
            <h2>Info</h2>
            <div className="bs-info selectable">
              <span>Name</span>
              <span>{p.tab.title}</span>
              <span>Location</span>
              <span>{p.tab.path ?? 'Not saved yet'}</span>
              <span>Format</span>
              <span>{formatLabelFor(p.tab.format)}</span>
              <span>Status</span>
              <span>{p.tab.dirty ? 'Unsaved changes' : 'All changes saved'}</span>
              {p.info?.map((r) => (
                <InfoRow key={r.label} {...r} />
              ))}
            </div>
            {p.tab.path && isElectron && (
              <div className="row" style={{ marginTop: 18 }}>
                <button type="button" className="btn" onClick={() => api.files.showInFolder(p.tab.path!)}>
                  <FolderOpen size={15} /> Show in folder
                </button>
                <button type="button" className="btn" onClick={() => void navigator.clipboard.writeText(p.tab.path!)}>
                  <Copy size={15} /> Copy path
                </button>
              </div>
            )}
            {p.infoExtra}
          </div>
        )}
        {section === 'new' && (
          <div className="bs-panel" key="new">
            <h2>New</h2>
            <div className="bs-grid">
              <button type="button" className="bs-card" onClick={() => run(() => newDocument(p.app))}>
                <span className="bs-card-icon">
                  <AppIcon kind={p.app} size={24} />
                </span>
                <span>
                  <div className="bs-card-title">Blank {KIND_LABEL[p.app].toLowerCase()}</div>
                  <div className="bs-card-desc">Start from scratch.</div>
                </span>
              </button>
            </div>
            <div className="bs-section-title">Templates</div>
            <div className="template-grid">
              {TEMPLATES.filter((t) => t.kind === p.app).map((t) => (
                <button key={t.id} type="button" className={`template-card kind-${t.kind}`} onClick={() => run(() => newDocument(t.kind, t.id, t.name))}>
                  <div className={`template-thumb thumb-${t.kind}`}>
                    <TemplatePreview t={t} />
                  </div>
                  <div className="template-meta">
                    <div className="grow">
                      <div className="template-name">{t.name}</div>
                      <div className="template-desc">{t.description}</div>
                    </div>
                  </div>
                </button>
              ))}
            </div>
          </div>
        )}
        {section === 'open' && (
          <div className="bs-panel" key="open">
            <h2>Open</h2>
            <button type="button" className="btn btn-primary btn-lg" onClick={() => run(() => void showOpenDialog(p.app))}>
              <FolderOpen size={17} /> Browse…
            </button>
            <div className="bs-section-title">Recent {KIND_LABEL[p.app].toLowerCase()}s</div>
            {recent.length === 0 && <p className="muted">Nothing here yet.</p>}
            <div className="recent-list" style={{ maxWidth: 760 }}>
              {recent.map((r) => (
                <div key={r.path} className="recent-row" style={{ gridTemplateColumns: '34px 1fr 140px' }} onClick={() => run(() => void openPaths([r.path]))}>
                  <AppIcon kind={r.kind} size={24} />
                  <span className="recent-name">
                    <span className="ellipsis">{r.name}</span>
                  </span>
                  <span className="recent-when">{timeAgo(r.openedAt)}</span>
                </div>
              ))}
            </div>
          </div>
        )}
        {section === 'saveas' && (
          <div className="bs-panel" key="saveas">
            <h2>Save as</h2>
            <div className="bs-grid">
              {p.saveTargets.map((t) => (
                <button
                  key={t.id}
                  type="button"
                  className="bs-card"
                  disabled={t.requiresLibreOffice && !p.libreOffice}
                  onClick={() => run(() => p.onSaveAs(t.id))}
                >
                  <span className="bs-card-icon">.{t.ext.toUpperCase()}</span>
                  <span>
                    <div className="bs-card-title">{t.label}</div>
                    <div className="bs-card-desc">{t.requiresLibreOffice && !p.libreOffice ? 'Install LibreOffice to enable this format.' : t.description}</div>
                  </span>
                </button>
              ))}
            </div>
          </div>
        )}
        {section === 'export' && (
          <div className="bs-panel" key="export">
            <h2>Export</h2>
            <div className="bs-grid">
              {p.exportTargets.map((t) => (
                <button key={t.id} type="button" className="bs-card" onClick={() => run(() => p.onExport(t.id))}>
                  <span className="bs-card-icon">.{t.ext.toUpperCase()}</span>
                  <span>
                    <div className="bs-card-title">{t.label}</div>
                    <div className="bs-card-desc">{t.description}</div>
                  </span>
                </button>
              ))}
            </div>
          </div>
        )}
        {section === 'print' && (
          <div className="bs-panel" key="print">
            <h2>Print</h2>
            <p className="muted" style={{ marginBottom: 16 }}>
              Opens your system’s print dialog, where you can pick a printer, pages and copies — or save as PDF.
            </p>
            <button type="button" className="btn btn-primary btn-lg" onClick={() => run(p.onPrint)}>
              <Printer size={17} /> Print…
            </button>
            {p.printExtra}
          </div>
        )}
      </div>
    </div>
  );
}

function InfoRow({ label, value }: { label: string; value: string }) {
  return (
    <>
      <span>{label}</span>
      <span>{value}</span>
    </>
  );
}
