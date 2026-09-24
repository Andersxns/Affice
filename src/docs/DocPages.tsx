import type { Editor } from '@tiptap/core';
import { EditorHost } from './useStableEditor';
import { forwardRef, useMemo, type CSSProperties } from 'react';
import { expandFields, pageDims, styleVars, type DocSettings } from './model';

export const PAGE_GAP = 22;

export type DocView = 'print' | 'web' | 'read';

interface Props {
  editor: Editor;
  settings: DocSettings;
  pages: number;
  zoom: number;
  view: DocView;
  title: string;
  darkPage: boolean;
  onHeaderFooter: (which: 'header' | 'footer') => void;
  showMarks: boolean;
  focusMode: boolean;
}

export const DocPages = forwardRef<HTMLDivElement, Props>(function DocPages(
  { editor, settings, pages, zoom, view, title, darkPage, onHeaderFooter, showMarks, focusMode },
  scrollRef,
) {
  const { w, h } = pageDims(settings.page);
  const W = w * 96;
  const H = h * 96;
  const m = settings.page.margins;
  const T = m.top * 96;
  const B = m.bottom * 96;
  const L = m.left * 96;
  const R = m.right * 96;
  const S = H + PAGE_GAP;
  const vars = useMemo(() => styleVars(settings), [settings]);
  const print = view === 'print';
  const stackH = print ? pages * S - PAGE_GAP : undefined;
  const hf = (t: { left: string; center: string; right: string }, page: number) =>
    [t.left, t.center, t.right].map((s) => expandFields(s, page, pages, title));

  const bodyStyle: CSSProperties = print
    ? { padding: `${T}px ${R}px ${B}px ${L}px`, minHeight: stackH }
    : view === 'read'
      ? { padding: '56px 64px 120px' }
      : { padding: '48px 56px 120px' };

  const onDoubleClick = (e: React.MouseEvent) => {
    if (!print) return;
    const stack = e.currentTarget as HTMLElement;
    const r = stack.getBoundingClientRect();
    const y = (e.clientY - r.top) / zoom;
    const within = y - Math.floor(y / S) * S;
    if (within < T - 4) {
      e.preventDefault();
      onHeaderFooter('header');
    } else if (within > H - B + 4 && within < H) {
      e.preventDefault();
      onHeaderFooter('footer');
    }
  };

  return (
    <div className={`doc-scroll thin-scroll${darkPage ? ' dark-page' : ''}${focusMode ? ' focus-mode' : ''}`} ref={scrollRef} data-accepts-drop>
      <div className="doc-zoom" style={{ zoom } as CSSProperties}>
        <div
          className={`doc-stack view-${view}${showMarks ? ' show-marks' : ''}`}
          style={{ width: print ? W : view === 'read' ? 820 : Math.max(W, 780), height: stackH, ['--page-color' as string]: settings.pageColor ?? '#ffffff' }}
          onDoubleClick={onDoubleClick}
        >
          {print &&
            Array.from({ length: pages }, (_, p) => {
              const hideHF = settings.differentFirstPage && p === 0;
              const [hl, hc, hr] = hf(settings.header, p + 1);
              const [fl, fc, fr] = hf(settings.footer, p + 1);
              return (
                <div key={p} className="page-bg" style={{ top: p * S, height: H }}>
                  {!hideHF && (hl || hc || hr) && (
                    <div className="page-hf page-header" style={{ height: T, padding: `0 ${R}px 0 ${L}px` }}>
                      <span>{hl}</span>
                      <span>{hc}</span>
                      <span>{hr}</span>
                    </div>
                  )}
                  {!hideHF && (fl || fc || fr) && (
                    <div className="page-hf page-footer" style={{ height: B, padding: `0 ${R}px 0 ${L}px` }}>
                      <span>{fl}</span>
                      <span>{fc}</span>
                      <span>{fr}</span>
                    </div>
                  )}
                  {settings.watermark && <div className="page-watermark">{settings.watermark}</div>}
                  <span className="page-num-badge">{p + 1}</span>
                </div>
              );
            })}
          {!print && <div className="page-bg page-bg-web" />}
          <div className="doc-body" style={{ ...bodyStyle, ...(vars as CSSProperties) }}>
            <EditorHost editor={editor} className="doc-editor-host" />
          </div>
        </div>
      </div>
    </div>
  );
});
