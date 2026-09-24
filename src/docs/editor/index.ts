import { Extension, type AnyExtension } from '@tiptap/core';
import StarterKit from '@tiptap/starter-kit';
import { TextStyleKit } from '@tiptap/extension-text-style';
import Highlight from '@tiptap/extension-highlight';
import TextAlign from '@tiptap/extension-text-align';
import Subscript from '@tiptap/extension-subscript';
import Superscript from '@tiptap/extension-superscript';
import Typography from '@tiptap/extension-typography';
import { TaskItem, TaskList } from '@tiptap/extension-list';
import { CharacterCount, Focus, Placeholder } from '@tiptap/extensions';
import {
  CommentMark,
  DocTable,
  DocTableCell,
  DocTableHeader,
  LinkClick,
  MathBlock,
  MathInline,
  PageBreak,
  ParagraphFormat,
  TableOfContents,
  TableRow,
  TabKey,
} from './extensions';
import { DocImage, ImageDropPaste } from './image';
import { Pagination } from './pagination';
import { Search } from './search';
import { SlashCommands, type SlashOptions } from './slash';

/** Bullet/number style for lists (disc, square, dash, check… / decimal, alpha, roman). */
const ListStyle = Extension.create({
  name: 'listStyle',
  addGlobalAttributes() {
    return [
      {
        types: ['bulletList', 'orderedList'],
        attributes: {
          listStyle: {
            default: null,
            parseHTML: (el) => {
              const e = el as HTMLElement;
              const d = e.getAttribute('data-list-style');
              if (d) return d;
              const t = e.style.listStyleType;
              if (t && t !== 'disc' && t !== 'decimal') return t;
              const type = e.getAttribute('type');
              if (type) return { a: 'lower-alpha', A: 'upper-alpha', i: 'lower-roman', I: 'upper-roman', '1': null }[type] ?? null;
              return null;
            },
            renderHTML: (a) => (a.listStyle ? { 'data-list-style': a.listStyle } : {}),
          },
        },
      },
    ];
  },
});

export interface BuildOptions {
  onPaginate: (info: { pages: number; headingPages: Map<number, number> }) => void;
  slash: SlashOptions;
  openLink: (href: string) => void;
  placeholder?: string;
}

export function buildExtensions(o: BuildOptions): AnyExtension[] {
  return [
    StarterKit.configure({
      heading: { levels: [1, 2, 3, 4, 5, 6] },
      link: {
        openOnClick: false,
        autolink: true,
        linkOnPaste: true,
        defaultProtocol: 'https',
        HTMLAttributes: { rel: 'noopener noreferrer nofollow', target: null },
      },
      undoRedo: { depth: 400, newGroupDelay: 700 },
      dropcursor: { color: '#2f6dff', width: 2 },
      bulletList: { keepMarks: true, keepAttributes: true },
      orderedList: { keepMarks: true, keepAttributes: true },
    }),
    TextStyleKit.configure({ lineHeight: false }),
    Highlight.configure({ multicolor: true }),
    TextAlign.configure({ types: ['heading', 'paragraph'], alignments: ['left', 'center', 'right', 'justify'] }),
    Subscript,
    Superscript,
    TaskList,
    TaskItem.configure({ nested: true }),
    DocTable.configure({ resizable: true, cellMinWidth: 28, allowTableNodeSelection: true, lastColumnResizable: true }),
    TableRow,
    DocTableHeader,
    DocTableCell,
    DocImage,
    Placeholder.configure({
      placeholder: ({ node, editor }) => {
        if (node.type.name === 'heading') return `Heading ${node.attrs.level}`;
        return editor.state.doc.childCount <= 1 ? (o.placeholder ?? 'Start typing, or press “/” for quick insert…') : '';
      },
      showOnlyCurrent: true,
    }),
    CharacterCount,
    Typography.configure({
      oneHalf: false,
      oneQuarter: false,
      threeQuarters: false,
      laquo: false,
      raquo: false,
      multiplication: false,
      plusMinus: false,
      notEqual: false,
      superscriptTwo: false,
      superscriptThree: false,
      leftArrow: false,
    }),
    Focus.configure({ className: 'has-focus', mode: 'shallowest' }),
    ParagraphFormat,
    ListStyle,
    PageBreak,
    TabKey,
    CommentMark,
    MathInline,
    MathBlock,
    TableOfContents,
    Search,
    Pagination.configure({ onPaginate: o.onPaginate }),
    SlashCommands.configure(o.slash),
    LinkClick.configure({ open: o.openLink }),
    Extension.create({ name: 'imageDropPaste', addProseMirrorPlugins: () => [ImageDropPaste] }),
  ];
}
