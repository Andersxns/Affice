import { Extension } from '@tiptap/core';
import type { Node as PMNode } from '@tiptap/pm/model';
import { Plugin, PluginKey, TextSelection, type EditorState, type Transaction } from '@tiptap/pm/state';
import { Decoration, DecorationSet } from '@tiptap/pm/view';

export interface SearchOptions {
  caseSensitive: boolean;
  wholeWord: boolean;
  regex: boolean;
}

export interface SearchState {
  query: string;
  opts: SearchOptions;
  results: Array<{ from: number; to: number }>;
  current: number;
  deco: DecorationSet;
}

export const searchKey = new PluginKey<SearchState>('search');

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    search: {
      setSearch: (query: string, opts?: Partial<SearchOptions>) => ReturnType;
      clearSearch: () => ReturnType;
      findNext: () => ReturnType;
      findPrev: () => ReturnType;
      replaceCurrent: (text: string) => ReturnType;
      replaceAll: (text: string) => ReturnType;
    };
  }
}

function buildRegex(query: string, o: SearchOptions): RegExp | null {
  if (!query) return null;
  let src = o.regex ? query : query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  if (o.wholeWord) src = `(?<![\\p{L}\\p{N}_])(?:${src})(?![\\p{L}\\p{N}_])`;
  try {
    return new RegExp(src, `gu${o.caseSensitive ? '' : 'i'}`);
  } catch {
    return null;
  }
}

/** Finds matches inside each textblock, mapping string offsets back to document positions. */
export function findMatches(doc: PMNode, re: RegExp): Array<{ from: number; to: number }> {
  const out: Array<{ from: number; to: number }> = [];
  doc.descendants((node, pos) => {
    if (!node.isTextblock) return true;
    let text = '';
    const map: number[] = [];
    node.forEach((child, offset) => {
      const start = pos + 1 + offset;
      if (child.isText) {
        for (let i = 0; i < child.text!.length; i++) map.push(start + i);
        text += child.text;
      } else {
        map.push(start);
        text += '￼';
      }
    });
    map.push(pos + 1 + node.content.size);
    re.lastIndex = 0;
    let m: RegExpExecArray | null;
    while ((m = re.exec(text))) {
      if (m[0].length === 0) {
        re.lastIndex++;
        continue;
      }
      out.push({ from: map[m.index], to: map[m.index + m[0].length - 1] + 1 });
      if (out.length > 5000) return false;
    }
    return false;
  });
  return out;
}

function decorate(doc: PMNode, results: SearchState['results'], current: number): DecorationSet {
  return DecorationSet.create(
    doc,
    results.map((r, i) => Decoration.inline(r.from, r.to, { class: i === current ? 'search-hit search-current' : 'search-hit' })),
  );
}

function compute(state: EditorState, query: string, opts: SearchOptions, prevCurrent = 0): SearchState {
  const re = buildRegex(query, opts);
  const results = re ? findMatches(state.doc, re) : [];
  let current = -1;
  if (results.length) {
    // start at the first match after the cursor
    const pos = state.selection.from;
    current = results.findIndex((r) => r.from >= pos);
    if (current < 0) current = 0;
    if (prevCurrent >= 0 && prevCurrent < results.length && Math.abs((results[prevCurrent]?.from ?? 0) - pos) < 2) current = prevCurrent;
  }
  return { query, opts, results, current, deco: decorate(state.doc, results, current) };
}

const EMPTY: SearchState = { query: '', opts: { caseSensitive: false, wholeWord: false, regex: false }, results: [], current: -1, deco: DecorationSet.empty };

export const Search = Extension.create({
  name: 'search',
  addProseMirrorPlugins() {
    return [
      new Plugin<SearchState>({
        key: searchKey,
        state: {
          init: () => EMPTY,
          apply(tr: Transaction, prev: SearchState, _old, newState): SearchState {
            const meta = tr.getMeta(searchKey) as Partial<SearchState> | { clear: true } | undefined;
            if (meta && 'clear' in meta) return EMPTY;
            if (meta) {
              const q = meta.query ?? prev.query;
              const o = meta.opts ?? prev.opts;
              if (meta.current !== undefined && meta.query === undefined) {
                return { ...prev, current: meta.current, deco: decorate(newState.doc, prev.results, meta.current) };
              }
              return compute(newState, q, o);
            }
            if (tr.docChanged && prev.query) {
              const next = compute(newState, prev.query, prev.opts, prev.current);
              return next;
            }
            return prev;
          },
        },
        props: {
          decorations(state) {
            return searchKey.getState(state)?.deco;
          },
        },
      }),
    ];
  },
  addCommands() {
    const goTo = (dir: 1 | -1) => () => ({ state, dispatch, view }: { state: EditorState; dispatch?: (tr: Transaction) => void; view: import('@tiptap/pm/view').EditorView }) => {
      const s = searchKey.getState(state);
      if (!s || !s.results.length) return false;
      const n = s.results.length;
      const current = ((s.current < 0 ? (dir === 1 ? -1 : 0) : s.current) + dir + n) % n;
      const r = s.results[current];
      if (dispatch) {
        const tr = state.tr.setMeta(searchKey, { current }).setSelection(TextSelection.create(state.doc, r.from, r.to)).scrollIntoView();
        dispatch(tr);
        requestAnimationFrame(() => {
          const el = view.dom.querySelector('.search-current');
          el?.scrollIntoView({ block: 'center', behavior: 'smooth' });
        });
      }
      return true;
    };
    return {
      setSearch:
        (query, opts) =>
        ({ state, dispatch }) => {
          const prev = searchKey.getState(state) ?? EMPTY;
          if (dispatch) dispatch(state.tr.setMeta(searchKey, { query, opts: { ...prev.opts, ...opts } }));
          return true;
        },
      clearSearch:
        () =>
        ({ state, dispatch }) => {
          if (dispatch) dispatch(state.tr.setMeta(searchKey, { clear: true }));
          return true;
        },
      findNext: goTo(1) as never,
      findPrev: goTo(-1) as never,
      replaceCurrent:
        (text) =>
        ({ state, dispatch }) => {
          const s = searchKey.getState(state);
          if (!s || s.current < 0 || !s.results[s.current]) return false;
          const r = s.results[s.current];
          if (dispatch) {
            const replacement = s.opts.regex ? expandRegexReplacement(state.doc.textBetween(r.from, r.to, '￼'), s.query, s.opts, text) : text;
            const tr = state.tr.insertText(replacement, r.from, r.to);
            dispatch(tr);
          }
          return true;
        },
      replaceAll:
        (text) =>
        ({ state, dispatch }) => {
          const s = searchKey.getState(state);
          if (!s || !s.results.length) return false;
          if (dispatch) {
            const tr = state.tr;
            // replace from the end so earlier positions stay valid
            for (let i = s.results.length - 1; i >= 0; i--) {
              const r = s.results[i];
              const replacement = s.opts.regex ? expandRegexReplacement(state.doc.textBetween(r.from, r.to, '￼'), s.query, s.opts, text) : text;
              tr.insertText(replacement, r.from, r.to);
            }
            dispatch(tr);
          }
          return true;
        },
    };
  },
});

function expandRegexReplacement(matched: string, query: string, opts: SearchOptions, replacement: string): string {
  const re = buildRegex(query, opts);
  if (!re) return replacement;
  re.lastIndex = 0;
  return matched.replace(new RegExp(re.source, re.flags.replace('g', '')), replacement);
}

export function getSearchState(state: EditorState): SearchState {
  return searchKey.getState(state) ?? EMPTY;
}
