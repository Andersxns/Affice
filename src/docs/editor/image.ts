import { Node } from '@tiptap/core';
import type { Node as PMNode } from '@tiptap/pm/model';
import { NodeSelection, Plugin, PluginKey } from '@tiptap/pm/state';
import type { EditorView } from '@tiptap/pm/view';
import { cssToPx } from './extensions';

export type ImageAlign = 'inline' | 'left' | 'right' | 'center';

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    docImage: {
      insertImage: (attrs: { src: string; alt?: string; width?: number; height?: number; align?: ImageAlign }) => ReturnType;
      setImageAttrs: (attrs: Partial<{ align: ImageAlign; width: number; height: number; alt: string; border: boolean; shadow: boolean; rounded: boolean }>) => ReturnType;
    };
  }
}

const MAX_INITIAL_WIDTH = 624; // 6.5in text column

export const DocImage = Node.create({
  name: 'image',
  group: 'inline',
  inline: true,
  atom: true,
  draggable: true,
  selectable: true,

  addAttributes() {
    return {
      src: { default: null },
      alt: { default: null },
      title: { default: null },
      width: { default: null },
      height: { default: null },
      align: { default: 'inline' },
      border: { default: false },
      shadow: { default: false },
      rounded: { default: false },
    };
  },

  parseHTML() {
    const fromImg = (img: HTMLImageElement, wrapper?: HTMLElement) => {
      const style = img.style;
      const w = cssToPx(style.width) ?? (img.getAttribute('width') ? Number(img.getAttribute('width')) : null);
      const h = cssToPx(style.height) ?? (img.getAttribute('height') ? Number(img.getAttribute('height')) : null);
      const float = (wrapper?.getAttribute('data-align') as ImageAlign | null) ?? (style.cssFloat === 'left' ? 'left' : style.cssFloat === 'right' ? 'right' : null);
      return {
        src: img.getAttribute('src'),
        alt: img.getAttribute('alt'),
        title: img.getAttribute('title'),
        width: w && Number.isFinite(w) ? Math.round(w) : null,
        height: h && Number.isFinite(h) ? Math.round(h) : null,
        align: float ?? 'inline',
        border: wrapper?.getAttribute('data-border') === 'true',
        shadow: wrapper?.getAttribute('data-shadow') === 'true',
        rounded: wrapper?.getAttribute('data-rounded') === 'true',
      };
    };
    return [
      {
        tag: 'span.doc-img',
        getAttrs: (el) => {
          const img = (el as HTMLElement).querySelector('img');
          return img ? fromImg(img, el as HTMLElement) : false;
        },
      },
      { tag: 'img[src]', getAttrs: (el) => fromImg(el as HTMLImageElement) },
    ];
  },

  renderHTML({ node }) {
    const a = node.attrs;
    const imgAttrs: Record<string, string> = { src: a.src };
    if (a.alt) imgAttrs.alt = a.alt;
    if (a.title) imgAttrs.title = a.title;
    if (a.width) imgAttrs.width = String(Math.round(a.width));
    if (a.height) imgAttrs.height = String(Math.round(a.height));
    const style = a.width ? `width:${Math.round(a.width)}px` : '';
    return [
      'span',
      {
        class: 'doc-img',
        'data-align': a.align,
        'data-border': a.border ? 'true' : null,
        'data-shadow': a.shadow ? 'true' : null,
        'data-rounded': a.rounded ? 'true' : null,
        style,
      },
      ['img', imgAttrs],
    ];
  },

  addCommands() {
    return {
      insertImage:
        (attrs) =>
        ({ commands }) =>
          commands.insertContent({ type: this.name, attrs: { align: 'inline', ...attrs } }),
      setImageAttrs:
        (attrs) =>
        ({ state, tr, dispatch }) => {
          const sel = state.selection;
          if (!(sel instanceof NodeSelection) || sel.node.type.name !== this.name) return false;
          tr.setNodeMarkup(sel.from, undefined, { ...sel.node.attrs, ...attrs });
          tr.setSelection(NodeSelection.create(tr.doc, sel.from));
          dispatch?.(tr);
          return true;
        },
    };
  },

  addNodeView() {
    return ({ node, getPos, editor }) => new ImageView(node, editor.view, getPos as () => number | undefined);
  },
});

class ImageView {
  dom: HTMLElement;
  img: HTMLImageElement;
  node: PMNode;
  private handles: HTMLElement[] = [];

  constructor(
    node: PMNode,
    private view: EditorView,
    private getPos: () => number | undefined,
  ) {
    this.node = node;
    this.dom = document.createElement('span');
    this.dom.className = 'doc-img';
    this.img = document.createElement('img');
    this.img.draggable = false;
    this.dom.appendChild(this.img);
    for (const corner of ['nw', 'ne', 'sw', 'se']) {
      const h = document.createElement('span');
      h.className = `img-handle img-handle-${corner}`;
      h.addEventListener('pointerdown', (e) => this.startResize(e, corner));
      this.dom.appendChild(h);
      this.handles.push(h);
    }
    this.apply(node);
  }

  private apply(node: PMNode) {
    const a = node.attrs;
    if (this.img.getAttribute('src') !== a.src) this.img.src = a.src ?? '';
    this.img.alt = a.alt ?? '';
    this.dom.setAttribute('data-align', a.align ?? 'inline');
    this.dom.toggleAttribute('data-border', false);
    if (a.border) this.dom.setAttribute('data-border', 'true');
    if (a.shadow) this.dom.setAttribute('data-shadow', 'true');
    else this.dom.removeAttribute('data-shadow');
    if (a.rounded) this.dom.setAttribute('data-rounded', 'true');
    else this.dom.removeAttribute('data-rounded');
    if (a.width) {
      this.dom.style.width = `${a.width}px`;
    } else {
      this.dom.style.width = '';
      // first render: size to natural width, capped to the text column
      const setNatural = () => {
        const w = Math.min(MAX_INITIAL_WIDTH, this.img.naturalWidth || MAX_INITIAL_WIDTH);
        if (!this.node.attrs.width && this.img.naturalWidth) this.dom.style.width = `${w}px`;
      };
      if (this.img.complete) setNatural();
      else this.img.onload = setNatural;
    }
  }

  update(node: PMNode) {
    if (node.type !== this.node.type) return false;
    this.node = node;
    this.apply(node);
    return true;
  }

  selectNode() {
    this.dom.classList.add('selected');
  }

  deselectNode() {
    this.dom.classList.remove('selected');
  }

  stopEvent(e: Event) {
    return (e.target as HTMLElement).classList?.contains('img-handle') ?? false;
  }

  ignoreMutation() {
    return true;
  }

  private startResize(e: PointerEvent, corner: string) {
    e.preventDefault();
    e.stopPropagation();
    const startX = e.clientX;
    const startW = this.dom.getBoundingClientRect().width;
    const zoom = this.dom.getBoundingClientRect().width / (this.dom.offsetWidth || 1) || 1;
    const ratio = (this.img.naturalHeight || 1) / (this.img.naturalWidth || 1);
    const dir = corner.includes('w') ? -1 : 1;
    const maxW = (this.dom.closest('.ProseMirror') as HTMLElement | null)?.clientWidth ?? 2000;
    let w = startW / zoom;
    const move = (ev: PointerEvent) => {
      const dx = ((ev.clientX - startX) * dir) / zoom;
      w = Math.max(24, Math.min(maxW, startW / zoom + dx));
      this.dom.style.width = `${w}px`;
    };
    const up = () => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
      const pos = this.getPos();
      if (pos === undefined) return;
      const tr = this.view.state.tr.setNodeMarkup(pos, undefined, { ...this.node.attrs, width: Math.round(w), height: Math.round(w * ratio) });
      tr.setSelection(NodeSelection.create(tr.doc, pos));
      this.view.dispatch(tr);
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
  }
}

/* ------------------------------------------------- paste / drop of images */

const MAX_EMBED_BYTES = 6 * 1024 * 1024;

export function readImageFile(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = async () => {
      const url = String(r.result);
      if (file.size <= MAX_EMBED_BYTES || file.type === 'image/svg+xml' || file.type === 'image/gif') resolve(url);
      else resolve(await downscale(url));
    };
    r.onerror = () => reject(r.error);
    r.readAsDataURL(file);
  });
}

/** Large photos are re-encoded (max 2400px) to keep documents light. */
async function downscale(url: string): Promise<string> {
  const img = new Image();
  img.src = url;
  await img.decode();
  const scale = Math.min(1, 2400 / Math.max(img.naturalWidth, img.naturalHeight));
  const c = document.createElement('canvas');
  c.width = Math.round(img.naturalWidth * scale);
  c.height = Math.round(img.naturalHeight * scale);
  c.getContext('2d')!.drawImage(img, 0, 0, c.width, c.height);
  return c.toDataURL('image/jpeg', 0.88);
}

export const ImageDropPaste = new Plugin({
  key: new PluginKey('imageDropPaste'),
  props: {
    handlePaste(view, event) {
      const files = Array.from(event.clipboardData?.files ?? []).filter((f) => f.type.startsWith('image/'));
      // If HTML is also present (e.g. copied from a web page), let the HTML paste handle it.
      if (!files.length || event.clipboardData?.getData('text/html')) return false;
      event.preventDefault();
      void insertFiles(view, files, view.state.selection.from);
      return true;
    },
    handleDrop(view, event) {
      const files = Array.from(event.dataTransfer?.files ?? []).filter((f) => f.type.startsWith('image/'));
      if (!files.length) return false;
      event.preventDefault();
      const pos = view.posAtCoords({ left: event.clientX, top: event.clientY })?.pos ?? view.state.selection.from;
      void insertFiles(view, files, pos);
      return true;
    },
  },
});

async function insertFiles(view: EditorView, files: File[], pos: number) {
  for (const f of files) {
    const src = await readImageFile(f);
    const node = view.state.schema.nodes.image.create({ src, alt: f.name, align: 'inline' });
    view.dispatch(view.state.tr.insert(Math.min(pos, view.state.doc.content.size), node));
  }
}
