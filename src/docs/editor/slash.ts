import { Extension } from '@tiptap/core';
import { Plugin, PluginKey } from '@tiptap/pm/state';

export interface SlashState {
  active: boolean;
  from: number;
  query: string;
}

export const slashKey = new PluginKey<SlashState>('slash');

const INACTIVE: SlashState = { active: false, from: 0, query: '' };

export interface SlashOptions {
  /** Called whenever the menu state changes. */
  onChange: (s: SlashState) => void;
  /** Keyboard handler while active; return true to swallow the key. */
  onKey: (e: KeyboardEvent) => boolean;
}

export const SlashCommands = Extension.create<SlashOptions>({
  name: 'slashCommands',
  addOptions() {
    return { onChange: () => undefined, onKey: () => false };
  },
  addProseMirrorPlugins() {
    const opts = this.options;
    return [
      new Plugin<SlashState>({
        key: slashKey,
        state: {
          init: () => INACTIVE,
          apply(tr, prev, _old, state) {
            const meta = tr.getMeta(slashKey) as SlashState | undefined;
            if (meta) return meta;
            if (!prev.active) return prev;
            const sel = state.selection;
            if (!sel.empty) return INACTIVE;
            const from = tr.mapping.map(prev.from);
            const $pos = state.doc.resolve(sel.from);
            if (sel.from <= from || $pos.start() > from) return INACTIVE;
            const text = state.doc.textBetween(from, sel.from, '\n', ' ');
            if (!text.startsWith('/') || /\s/.test(text.slice(1)) || text.length > 24) return INACTIVE;
            return { active: true, from, query: text.slice(1) };
          },
        },
        props: {
          handleTextInput(view, from, _to, text) {
            if (text !== '/') return false;
            const $from = view.state.doc.resolve(from);
            if (!$from.parent.isTextblock || $from.parent.type.name === 'codeBlock') return false;
            const before = $from.parent.textBetween(0, $from.parentOffset, undefined, ' ');
            if (before.length && !/\s$/.test(before)) return false;
            // let the "/" be inserted, then activate
            setTimeout(() => {
              const tr = view.state.tr.setMeta(slashKey, { active: true, from, query: '' });
              view.dispatch(tr);
            }, 0);
            return false;
          },
          handleKeyDown(view, event) {
            const s = slashKey.getState(view.state);
            if (!s?.active) return false;
            if (event.key === 'Escape') {
              view.dispatch(view.state.tr.setMeta(slashKey, INACTIVE));
              return true;
            }
            return opts.onKey(event);
          },
        },
        view() {
          let last: SlashState = INACTIVE;
          return {
            update(v) {
              const s = slashKey.getState(v.state) ?? INACTIVE;
              if (s.active !== last.active || s.query !== last.query || s.from !== last.from) {
                last = s;
                opts.onChange(s);
              }
            },
          };
        },
      }),
    ];
  },
});
