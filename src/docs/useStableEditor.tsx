import type { Editor } from '@tiptap/core';
import { useEffect, useLayoutEffect, useRef, useState } from 'react';
import { getTab } from '@/app/workspace';

/**
 * Creates one TipTap editor per tab and keeps it alive for the tab's lifetime.
 *
 * TipTap's own `useEditor` destroys the editor whenever React temporarily disconnects
 * effects (StrictMode, Suspense), which would throw away the loaded document. Here the
 * editor is only destroyed once the tab has really been closed.
 */
export function useStableEditor(tabId: string, create: () => Editor): Editor {
  const [editor] = useState(create);
  useEffect(
    () => () => {
      setTimeout(() => {
        if (!getTab(tabId) && !editor.isDestroyed) editor.destroy();
      }, 200);
    },
    [editor, tabId],
  );
  return editor;
}

/** Places the editor's ProseMirror DOM inside this component. */
export function EditorHost({ editor, className }: { editor: Editor; className?: string }) {
  const ref = useRef<HTMLDivElement>(null);
  useLayoutEffect(() => {
    const host = ref.current;
    if (!host || editor.isDestroyed) return;
    const dom = editor.view.dom;
    if (dom.parentElement !== host) host.appendChild(dom);
  }, [editor]);
  return <div ref={ref} className={className} />;
}
