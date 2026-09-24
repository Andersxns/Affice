import { ChevronLeft, ChevronRight, List, Plus } from 'lucide-react';
import { useEffect, useRef, useState, useSyncExternalStore } from 'react';
import { ColorPicker } from '@/ui/color';
import { Menu, openContextMenu, type MenuItem } from '@/ui/menu';
import { confirmDialog } from '@/ui/dialog';
import { toast } from '@/ui/toast';
import type { SheetDoc } from '../doc';
import { addSheet, deleteSheet, duplicateSheet, moveSheet, renameSheet, setSheetProp } from '../ops/structure';

export function SheetTabs({ doc, onChanged, onProtect }: { doc: SheetDoc; onChanged: () => void; onProtect: (index: number) => void }) {
  useSyncExternalStore(doc.subscribe, doc.getVersion);
  const [renaming, setRenaming] = useState<number | null>(null);
  const [draft, setDraft] = useState('');
  const [listOpen, setListOpen] = useState(false);
  const [drag, setDrag] = useState<{ from: number; over: number; x: number } | null>(null);
  const stripRef = useRef<HTMLDivElement>(null);
  const listRef = useRef<HTMLButtonElement>(null);
  const wb = doc.wb;
  const active = wb.activeSheet;

  useEffect(() => {
    // keep the active tab in view
    const el = stripRef.current?.querySelector<HTMLElement>('.sheet-tab.active');
    el?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }, [active, wb.sheets.length]);

  const commitRename = () => {
    if (renaming === null) return;
    const err = renameSheet(doc, renaming, draft);
    if (err) toast.error(err);
    else onChanged();
    setRenaming(null);
  };

  const removeSheet = async (i: number) => {
    const s = wb.sheets[i];
    const hasData = s.cells.size > 0 || s.charts.length > 0;
    if (hasData && !(await confirmDialog({ title: `Delete “${s.name}”?`, message: 'The sheet and everything on it will be deleted. You can undo this with Ctrl+Z.', confirmLabel: 'Delete', danger: true }))) return;
    if (!deleteSheet(doc, i)) toast.info('A workbook must keep at least one visible sheet.');
    else onChanged();
  };

  const menuFor = (i: number): MenuItem[] => {
    const s = wb.sheets[i];
    const hidden = wb.sheets.map((x, k) => [x, k] as const).filter(([x]) => x.hidden);
    return [
      { label: 'Insert sheet', shortcut: 'Shift+F11', onSelect: () => (addSheet(doc, i + 1), onChanged()) },
      { label: 'Delete', danger: true, onSelect: () => void removeSheet(i) },
      {
        label: 'Rename',
        onSelect: () => {
          setDraft(s.name);
          setRenaming(i);
        },
      },
      { label: 'Duplicate', onSelect: () => (duplicateSheet(doc, i), onChanged()) },
      { separator: true },
      { label: 'Move left', disabled: i === 0, onSelect: () => (moveSheet(doc, i, i - 1), onChanged()) },
      { label: 'Move right', disabled: i === wb.sheets.length - 1, onSelect: () => (moveSheet(doc, i, i + 1), onChanged()) },
      {
        label: 'Tab colour',
        submenu: [
          {
            custom: (close) => (
              <ColorPicker
                value={s.tabColor ?? null}
                noneLabel="No colour"
                onPick={(c) => {
                  setSheetProp(doc, i, 'tabColor', c ?? undefined);
                  onChanged();
                  close();
                }}
              />
            ),
          },
        ],
      },
      { separator: true },
      { label: s.protection?.enabled ? 'Unprotect sheet…' : 'Protect sheet…', onSelect: () => onProtect(i) },
      { label: 'Hide', onSelect: () => (setSheetProp(doc, i, 'hidden', true), onChanged()) },
      {
        label: 'Unhide',
        disabled: !hidden.length,
        submenu: hidden.map(([x, k]) => ({ label: x.name, onSelect: () => (setSheetProp(doc, k, 'hidden', false), doc.setActiveSheet(k), onChanged()) })),
      },
    ];
  };

  const scrollBy = (dx: number) => stripRef.current?.scrollBy({ left: dx, behavior: 'smooth' });

  const onPointerDown = (e: React.PointerEvent, i: number) => {
    if (e.button !== 0 || renaming !== null) return;
    doc.setActiveSheet(i);
    setDrag({ from: i, over: i, x: e.clientX });
    (e.currentTarget as HTMLElement).setPointerCapture(e.pointerId);
  };

  const onPointerMove = (e: React.PointerEvent) => {
    if (!drag) return;
    if (Math.abs(e.clientX - drag.x) < 6 && drag.over === drag.from) return;
    const tabs = [...(stripRef.current?.querySelectorAll<HTMLElement>('.sheet-tab') ?? [])];
    let over = drag.from;
    for (const t of tabs) {
      const r = t.getBoundingClientRect();
      const idx = Number(t.dataset.index);
      if (e.clientX >= r.left && e.clientX <= r.right) over = idx;
    }
    if (over !== drag.over) setDrag({ ...drag, over });
  };

  const onPointerUp = () => {
    if (drag && drag.over !== drag.from) {
      moveSheet(doc, drag.from, drag.over);
      onChanged();
    }
    setDrag(null);
  };

  return (
    <div className="sheet-tabs">
      <button className="icon-btn icon-btn-sm" data-tip="New sheet" data-tip-key="Shift+F11" onClick={() => (addSheet(doc), onChanged())}>
        <Plus size={16} />
      </button>
      <button ref={listRef} className="icon-btn icon-btn-sm" data-tip="All sheets" onClick={() => setListOpen((o) => !o)}>
        <List size={15} />
      </button>
      <Menu
        open={listOpen}
        anchor={listRef.current}
        placement="top-start"
        onClose={() => setListOpen(false)}
        items={wb.sheets.map((s, i) => ({ label: s.name + (s.hidden ? ' (hidden)' : ''), checked: i === active, onSelect: () => (s.hidden ? (setSheetProp(doc, i, 'hidden', false), doc.setActiveSheet(i)) : doc.setActiveSheet(i)) }))}
      />
      <button className="icon-btn icon-btn-sm tabs-nav" aria-label="Scroll sheet tabs left" onClick={() => scrollBy(-160)}>
        <ChevronLeft size={15} />
      </button>
      <div className="sheet-tab-strip thin-scroll" ref={stripRef} role="tablist" onPointerMove={onPointerMove} onPointerUp={onPointerUp}>
        {wb.sheets.map((s, i) =>
          s.hidden ? null : (
            <div
              key={s.id}
              role="tab"
              aria-selected={i === active}
              data-index={i}
              className={`sheet-tab${i === active ? ' active' : ''}${drag && drag.over === i && drag.from !== i ? ' drop' : ''}`}
              style={s.tabColor ? ({ '--tab-color': s.tabColor } as React.CSSProperties) : undefined}
              onPointerDown={(e) => onPointerDown(e, i)}
              onDoubleClick={() => {
                setDraft(s.name);
                setRenaming(i);
              }}
              onContextMenu={(e) => {
                e.preventDefault();
                doc.setActiveSheet(i);
                openContextMenu(e, menuFor(i));
              }}
            >
              {renaming === i ? (
                <input
                  className="sheet-tab-input"
                  value={draft}
                  autoFocus
                  onFocus={(e) => e.target.select()}
                  onChange={(e) => setDraft(e.target.value)}
                  onBlur={commitRename}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') commitRename();
                    if (e.key === 'Escape') setRenaming(null);
                    e.stopPropagation();
                  }}
                  style={{ width: Math.max(40, draft.length * 8 + 16) }}
                />
              ) : (
                <span className="sheet-tab-name">
                  {s.protection?.enabled && '🔒 '}
                  {s.name}
                </span>
              )}
              {s.tabColor && <span className="sheet-tab-color" />}
            </div>
          ),
        )}
      </div>
      <button className="icon-btn icon-btn-sm tabs-nav" aria-label="Scroll sheet tabs right" onClick={() => scrollBy(160)}>
        <ChevronRight size={15} />
      </button>
    </div>
  );
}
