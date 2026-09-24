import { useEffect, useRef } from 'react';
import { createPortal } from 'react-dom';
import { allFunctions, getFunction, type FnDef } from '../engine/functions';
import { tokenize } from '../engine/parser';

export interface Suggestion {
  kind: 'fn' | 'name';
  name: string;
  syntax?: string;
  desc?: string;
}

export interface AssistInfo {
  prefix: { start: number; end: number; text: string } | null;
  suggestions: Suggestion[];
  signature: { def: FnDef; arg: number } | null;
}

const POPULAR = ['SUM', 'AVERAGE', 'IF', 'COUNT', 'COUNTA', 'MAX', 'MIN', 'VLOOKUP', 'XLOOKUP', 'SUMIF', 'COUNTIF', 'IFERROR', 'INDEX', 'MATCH', 'ROUND', 'TODAY', 'CONCAT', 'TEXT', 'LEFT', 'FILTER', 'SORT', 'UNIQUE'];

/** Works out autocomplete suggestions and the signature hint at the caret. */
export function analyzeFormula(text: string, caret: number, names: string[]): AssistInfo {
  const none: AssistInfo = { prefix: null, suggestions: [], signature: null };
  if (!text.startsWith('=')) return none;
  const before = text.slice(0, caret);
  // inside a string literal?
  if ((before.match(/"/g)?.length ?? 0) % 2 === 1) return none;
  let prefix: AssistInfo['prefix'] = null;
  const m = /(^|[^A-Za-z0-9_.$!'"\]])([A-Za-z_][A-Za-z0-9_.]*)$/.exec(before.slice(1));
  if (m && !/^[A-Za-z]{1,3}\d+$/.test(m[2])) {
    const start = caret - m[2].length;
    const after = text.slice(caret);
    const tail = /^[A-Za-z0-9_.]*/.exec(after)?.[0] ?? '';
    if (!after.slice(tail.length).startsWith('(')) prefix = { start, end: caret + tail.length, text: m[2] };
  }
  let suggestions: Suggestion[] = [];
  if (prefix) {
    const up = prefix.text.toUpperCase();
    const fns = allFunctions().filter((f) => f.name.startsWith(up));
    fns.sort((a, b) => {
      const pa = POPULAR.indexOf(a.name);
      const pb = POPULAR.indexOf(b.name);
      if (a.name === up) return -1;
      if (b.name === up) return 1;
      if (pa >= 0 || pb >= 0) return (pa < 0 ? 999 : pa) - (pb < 0 ? 999 : pb);
      return a.name.length - b.name.length || a.name.localeCompare(b.name);
    });
    suggestions = [
      ...names.filter((n) => n.toUpperCase().startsWith(up)).map((n) => ({ kind: 'name' as const, name: n, desc: 'Defined name' })),
      ...fns.map((f) => ({ kind: 'fn' as const, name: f.name, syntax: f.syntax, desc: f.desc })),
    ].slice(0, 12);
    if (suggestions.length === 1 && suggestions[0].name.toUpperCase() === up && suggestions[0].kind === 'fn') suggestions = [];
  }
  // innermost open function call
  let signature: AssistInfo['signature'] = null;
  const toks = tokenize(before.slice(1));
  const stack: { name: string | null; arg: number }[] = [];
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (t.type === 'func') {
      stack.push({ name: t.text.toUpperCase().replace(/^_XLFN\./, ''), arg: 0 });
      // skip the '(' that follows
      while (i + 1 < toks.length && toks[i + 1].type === 'ws') i++;
      if (toks[i + 1]?.type === 'lparen') i++;
    } else if (t.type === 'lparen') stack.push({ name: null, arg: 0 });
    else if (t.type === 'rparen') stack.pop();
    else if ((t.type === 'sep' || t.type === 'semi') && stack.length) stack[stack.length - 1].arg++;
    else if (t.type === 'lbrace') stack.push({ name: null, arg: 0 });
    else if (t.type === 'rbrace') stack.pop();
  }
  for (let i = stack.length - 1; i >= 0; i--) {
    const s = stack[i];
    if (s.name) {
      const def = getFunction(s.name);
      if (def) signature = { def, arg: s.arg };
      break;
    }
  }
  return { prefix, suggestions, signature };
}

/** Splits a syntax string into arguments, handling repeating "…" parts. */
export function syntaxParts(syntax: string): string[] {
  if (!syntax) return [];
  return syntax.split(/,\s*/);
}

export function highlightArg(syntax: string, arg: number): { parts: string[]; active: number } {
  const parts = syntaxParts(syntax);
  let active = arg;
  if (active >= parts.length) {
    // repeating tail ("value2, …") keeps the last named argument highlighted
    const ell = parts.findIndex((p) => p.includes('…'));
    active = ell > 0 ? ell - 1 : parts.length - 1;
  }
  return { parts, active };
}

export function FormulaAssist({
  info,
  index,
  anchor,
  onPick,
  onHover,
}: {
  info: AssistInfo;
  index: number;
  anchor: DOMRect | null;
  onPick: (s: Suggestion) => void;
  onHover?: (i: number) => void;
}) {
  const listRef = useRef<HTMLDivElement>(null);
  useEffect(() => {
    listRef.current?.querySelector('.fa-item.active')?.scrollIntoView({ block: 'nearest' });
  }, [index]);
  if (!anchor) return null;
  const showList = info.suggestions.length > 0;
  const showSig = !showList && info.signature;
  if (!showList && !showSig) return null;
  const top = anchor.bottom + 4;
  const left = Math.max(8, Math.min(anchor.left, window.innerWidth - 420));
  const current = showList ? info.suggestions[index] : null;
  return createPortal(
    <div className="formula-assist" style={{ top, left }} onMouseDown={(e) => e.preventDefault()}>
      {showList && (
        <div className="fa-list" ref={listRef} role="listbox">
          {info.suggestions.map((s, i) => (
            <div
              key={s.kind + s.name}
              role="option"
              aria-selected={i === index}
              className={`fa-item${i === index ? ' active' : ''}`}
              onMouseEnter={() => onHover?.(i)}
              onClick={() => onPick(s)}
            >
              <span className={`fa-kind ${s.kind}`}>{s.kind === 'fn' ? 'ƒ' : '≡'}</span>
              <span className="fa-name">{s.name}</span>
            </div>
          ))}
        </div>
      )}
      {showList && current && (
        <div className="fa-detail">
          <div className="fa-sig">
            <b>{current.name}</b>
            {current.kind === 'fn' ? `(${current.syntax ?? ''})` : ''}
          </div>
          {current.desc && <div className="fa-desc">{current.desc}</div>}
          <div className="fa-hint">Tab to insert</div>
        </div>
      )}
      {showSig && info.signature && <Signature def={info.signature.def} arg={info.signature.arg} />}
    </div>,
    document.body,
  );
}

function Signature({ def, arg }: { def: FnDef; arg: number }) {
  const { parts, active } = highlightArg(def.syntax, arg);
  return (
    <div className="fa-signature">
      <div className="fa-sig">
        <b>{def.name}</b>(
        {parts.map((p, i) => (
          <span key={i}>
            {i > 0 && ', '}
            <span className={i === active ? 'fa-arg active' : 'fa-arg'}>{p}</span>
          </span>
        ))}
        )
      </div>
      <div className="fa-desc">{def.desc}</div>
    </div>
  );
}

/** Applies a picked suggestion to the text; returns new text and caret. */
export function applySuggestion(text: string, info: AssistInfo, s: Suggestion): { text: string; caret: number } {
  if (!info.prefix) return { text, caret: text.length };
  const insert = s.kind === 'fn' ? `${s.name}(` : s.name;
  const next = text.slice(0, info.prefix.start) + insert + text.slice(info.prefix.end);
  return { text: next, caret: info.prefix.start + insert.length };
}
