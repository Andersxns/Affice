import { Check, ChevronDown, ChevronUp, FunctionSquare, X } from 'lucide-react';
import { useEffect, useLayoutEffect, useRef, useState, useSyncExternalStore } from 'react';
import { Menu } from '@/ui/menu';
import { colName } from '../model/address';
import type { SheetDoc } from '../doc';
import type { SheetUI } from './controller';
import { analyzeFormula, applySuggestion, FormulaAssist } from './formulaAssist';

export function FormulaBar({ doc, ui, onInsertFunction, onCommitted, onBlurToGrid }: { doc: SheetDoc; ui: SheetUI; onInsertFunction: () => void; onCommitted: () => void; onBlurToGrid: () => void }) {
  useSyncExternalStore(doc.subscribe, doc.getVersion);
  useSyncExternalStore(ui.subscribe, ui.getVersion);
  const [expanded, setExpanded] = useState(false);
  const [nameText, setNameText] = useState<string | null>(null);
  const [namesOpen, setNamesOpen] = useState(false);
  const [assistIndex, setAssistIndex] = useState(0);
  const inputRef = useRef<HTMLTextAreaElement>(null);
  const hlRef = useRef<HTMLDivElement>(null);
  const nameRef = useRef<HTMLDivElement>(null);

  const sheet = doc.sheet;
  const { r, c } = doc.sel.active;
  const edit = ui.edit;
  const barEditing = !!edit && edit.source === 'bar';
  const text = edit ? edit.text : doc.editText(sheet, r, c);
  const hideFormula = !edit && sheet.protection?.enabled && doc.styleOf(sheet, r, c).hideFormula;
  const shown = hideFormula ? '' : text;

  const assist = barEditing && edit ? analyzeFormula(edit.text, edit.caret, doc.wb.names.map((n) => n.name)) : null;
  useEffect(() => setAssistIndex(0), [assist?.prefix?.text]);

  // keep the textarea value in sync when not typing in it
  useLayoutEffect(() => {
    const el = inputRef.current;
    if (!el) return;
    if (el.value !== shown) el.value = shown;
    if (barEditing && edit && document.activeElement === el && el.selectionStart !== edit.caret) el.setSelectionRange(edit.caret, edit.caret);
  });

  // syntax highlighting of references
  const refs = edit ? ui.editRefs() : [];
  const pieces: { text: string; color?: string }[] = [];
  if (edit && edit.text.startsWith('=')) {
    let pos = 0;
    for (const rf of refs) {
      if (rf.start > pos) pieces.push({ text: edit.text.slice(pos, rf.start) });
      pieces.push({ text: edit.text.slice(rf.start, rf.end), color: rf.color });
      pos = rf.end;
    }
    pieces.push({ text: edit.text.slice(pos) });
  }

  const begin = () => {
    if (!ui.edit) {
      if (!ui.startEdit(null, 'edit', 'bar')) {
        inputRef.current?.blur();
        return;
      }
    } else if (ui.edit.source !== 'bar') {
      ui.edit = { ...ui.edit, source: 'bar', mode: 'edit' };
      ui.emit();
    }
  };

  const commit = (move?: [number, number]) => {
    if (!ui.commitEdit()) return;
    onCommitted();
    if (move) ui.cycle(move[0], move[1]);
    onBlurToGrid();
  };

  const onKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    const ed = ui.edit;
    if (!ed || e.nativeEvent.isComposing) return;
    const sugg = assist?.suggestions ?? [];
    if (sugg.length && (e.key === 'ArrowDown' || e.key === 'ArrowUp')) {
      e.preventDefault();
      setAssistIndex((i) => (i + (e.key === 'ArrowDown' ? 1 : -1) + sugg.length) % sugg.length);
      return;
    }
    if (sugg.length && (e.key === 'Tab' || e.key === 'Enter')) {
      e.preventDefault();
      const res = applySuggestion(ed.text, assist!, sugg[assistIndex] ?? sugg[0]);
      ui.updateEdit(res.text, res.caret, 'bar');
      return;
    }
    if (e.key === 'Enter' && !e.altKey) {
      e.preventDefault();
      commit([e.shiftKey ? -1 : 1, 0]);
    } else if (e.key === 'Enter' && e.altKey) {
      e.preventDefault();
      const el = e.currentTarget;
      const s = el.selectionStart;
      ui.updateEdit(ed.text.slice(0, s) + '\n' + ed.text.slice(el.selectionEnd), s + 1, 'bar');
    } else if (e.key === 'Tab') {
      e.preventDefault();
      commit([0, e.shiftKey ? -1 : 1]);
    } else if (e.key === 'Escape') {
      e.preventDefault();
      ui.cancelEdit();
      onBlurToGrid();
    } else if (e.key === 'F4') {
      e.preventDefault();
      ui.toggleAbsolute();
    }
  };

  const names = doc.wb.names.filter((n) => n.sheet === undefined || n.sheet === sheet.id);

  return (
    <div className={`formula-bar${expanded ? ' expanded' : ''}`}>
      <div className="fb-name" ref={nameRef}>
        <input
          className="fb-name-input"
          aria-label="Name box"
          value={nameText ?? ui.cellAddress()}
          onFocus={(e) => {
            setNameText(ui.cellAddress());
            requestAnimationFrame(() => e.target.select());
          }}
          onChange={(e) => setNameText(e.target.value)}
          onBlur={() => setNameText(null)}
          onKeyDown={(e) => {
            if (e.key === 'Enter') {
              e.preventDefault();
              const ok = ui.goTo(nameText ?? '');
              if (!ok && nameText && /^[A-Za-z_\\][\w.]*$/.test(nameText) && !/^[A-Za-z]{1,3}\d+$/.test(nameText)) {
                // typing a new name defines it for the selection (like Excel)
                const rg = doc.sel.ranges[doc.sel.ranges.length - 1];
                const ref = `'${sheet.name.replace(/'/g, "''")}'!$${colName(rg.c1)}$${rg.r1 + 1}${rg.r1 !== rg.r2 || rg.c1 !== rg.c2 ? `:$${colName(rg.c2)}$${rg.r2 + 1}` : ''}`;
                doc.transact('Define name', (tx) => {
                  tx.names();
                  doc.wb.names.push({ name: nameText, ref });
                });
                ui.onNotify?.(`Named ${nameText}`);
              } else if (!ok) ui.onError?.('Invalid reference', 'Type a cell like B7, a range like A1:C20, or a defined name.');
              setNameText(null);
              onBlurToGrid();
            } else if (e.key === 'Escape') {
              setNameText(null);
              onBlurToGrid();
            }
          }}
        />
        <button className="fb-name-drop" aria-label="Defined names" onMouseDown={(e) => e.preventDefault()} onClick={() => setNamesOpen((o) => !o)}>
          <ChevronDown size={12} />
        </button>
        <Menu
          open={namesOpen}
          anchor={nameRef.current}
          onClose={() => setNamesOpen(false)}
          items={
            names.length
              ? names.map((n) => ({ label: n.name, hint: n.ref, onSelect: () => ui.goTo(n.name) }))
              : [{ label: 'No named ranges yet', disabled: true }, { label: 'Type a name in this box and press Enter to name the selection.', disabled: true }]
          }
        />
      </div>
      <div className="fb-buttons">
        <button className="icon-btn icon-btn-sm" data-tip="Cancel" disabled={!edit} onMouseDown={(e) => e.preventDefault()} onClick={() => (ui.cancelEdit(), onBlurToGrid())}>
          <X size={15} />
        </button>
        <button className="icon-btn icon-btn-sm" data-tip="Enter" disabled={!edit} onMouseDown={(e) => e.preventDefault()} onClick={() => commit()}>
          <Check size={15} />
        </button>
        <button className="icon-btn icon-btn-sm fx" data-tip="Insert function" data-tip-key="Shift+F3" onMouseDown={(e) => e.preventDefault()} onClick={onInsertFunction}>
          <FunctionSquare size={16} />
        </button>
      </div>
      <div className="fb-input-wrap">
        {pieces.length > 0 && (
          <div className="fb-highlight" ref={hlRef} aria-hidden>
            {pieces.map((p, i) => (
              <span key={i} style={p.color ? { color: p.color } : undefined}>
                {p.text}
              </span>
            ))}
            {'\n'}
          </div>
        )}
        <textarea
          ref={inputRef}
          className={`fb-input${pieces.length ? ' colored' : ''}`}
          aria-label="Formula bar"
          spellCheck={false}
          defaultValue={shown}
          readOnly={hideFormula}
          rows={1}
          onFocus={begin}
          onScroll={(e) => {
            if (hlRef.current) hlRef.current.scrollTop = e.currentTarget.scrollTop;
          }}
          onChange={(e) => {
            begin();
            ui.updateEdit(e.target.value, e.target.selectionStart, 'bar');
          }}
          onSelect={(e) => {
            const t = e.currentTarget;
            if (ui.edit && ui.edit.source === 'bar' && t.selectionStart !== ui.edit.caret) {
              ui.edit = { ...ui.edit, caret: t.selectionStart };
              ui.emit();
            }
          }}
          onKeyDown={onKeyDown}
          onBlur={(e) => {
            const to = e.relatedTarget as HTMLElement | null;
            if (to?.closest('.formula-assist, .grid-scroller, .cell-editor, .fb-buttons')) return;
            if (ui.edit && ui.edit.source === 'bar' && !to?.closest('.dialog, .modal, .popover, .menu')) {
              if (ui.commitEdit()) onCommitted();
            }
          }}
        />
        {barEditing && assist && (
          <FormulaAssist
            info={assist}
            index={assistIndex}
            anchor={inputRef.current?.getBoundingClientRect() ?? null}
            onHover={setAssistIndex}
            onPick={(s) => {
              const ed = ui.edit;
              if (!ed) return;
              const res = applySuggestion(ed.text, assist, s);
              ui.updateEdit(res.text, res.caret, 'bar');
              inputRef.current?.focus();
            }}
          />
        )}
      </div>
      <button className="icon-btn icon-btn-sm fb-expand" data-tip={expanded ? 'Collapse formula bar' : 'Expand formula bar'} onClick={() => setExpanded((x) => !x)}>
        {expanded ? <ChevronUp size={15} /> : <ChevronDown size={15} />}
      </button>
    </div>
  );
}
