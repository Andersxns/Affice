/**
 * Table styles modelled on PowerPoint's built-in families. A table stores only the style options;
 * cell appearance is derived here, so tables recolour with the theme, and explicit cell formatting wins.
 */
import { newId, type Fill, type Line, type TableCell, type TableEl, type TableStyleOpts, type TextBody, type ThemeColorName } from './model';

export interface CellLook {
  fill?: Fill;
  color?: string;
  bold?: boolean;
  borders: [Line | null, Line | null, Line | null, Line | null];
}

const thin = (color: string, width = 1.33): Line => ({ color, width });

/** Built-in PowerPoint table style GUIDs by family and accent (used when exporting). */
export const TABLE_STYLE_IDS: Record<string, Partial<Record<ThemeColorName | 'none', string>>> = {
  medium2: {
    tx1: '{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}',
    accent1: '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}',
    accent2: '{21E4AEA4-8DFA-4A89-87EB-49C32662AFE8}',
    accent3: '{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}',
    accent4: '{00A15C55-8517-42AA-B614-E9B94910E393}',
    accent5: '{7DF18680-E054-41AD-8BC1-D1AEF772440D}',
    accent6: '{93296810-A885-4BE3-A3E7-6D5BEEA58F35}',
  },
  light1: { accent1: '{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}' },
  light2: { accent1: '{69012ECD-51FC-41F1-AA8D-1B2483CD663E}' },
  medium1: { accent1: '{B301B821-A1FF-4177-AEE7-76D212191A09}' },
  dark1: { accent1: '{E8034E78-7F5D-4C2E-B375-FC64B27BC917}' },
};

export function styleFromId(id: string | undefined): Pick<TableStyleOpts, 'family' | 'accent' | 'styleId'> {
  if (!id) return { family: 'none' };
  for (const [family, accents] of Object.entries(TABLE_STYLE_IDS))
    for (const [accent, gid] of Object.entries(accents)) if (gid?.toUpperCase() === id.toUpperCase()) return { family: family as TableStyleOpts['family'], accent: accent as ThemeColorName, styleId: id };
  // unknown styles render like the default so tables still look structured
  return { family: 'medium2', accent: 'accent1', styleId: id };
}

export function tableStyleId(style: TableStyleOpts | undefined): string | undefined {
  if (!style || style.family === 'none') return style?.styleId;
  const fam = TABLE_STYLE_IDS[style.family ?? 'medium2'];
  return fam?.[style.accent ?? 'accent1'] ?? style.styleId ?? TABLE_STYLE_IDS.medium2.accent1;
}

/** Visual look of a cell from the table style (before explicit cell formatting). */
export function cellLook(t: TableEl, r: number, c: number): CellLook {
  const st = t.style ?? {};
  const fam = st.family ?? 'medium2';
  const acc = `@${st.accent ?? 'accent1'}`;
  const nR = t.rows.length;
  const nC = t.cols.length;
  const isHeader = !!st.firstRow && r === 0;
  const isTotal = !!st.lastRow && r === nR - 1 && nR > 1;
  const isFirstCol = !!st.firstCol && c === 0;
  const isLastCol = !!st.lastCol && c === nC - 1 && nC > 1;
  const dataRow = r - (st.firstRow ? 1 : 0);
  const banded = !!st.bandRows && dataRow % 2 === 0 && !isHeader && !isTotal;
  const bandedCol = !!st.bandCols && (c - (st.firstCol ? 1 : 0)) % 2 === 0;
  const none: CellLook['borders'] = [null, null, null, null];
  switch (fam) {
    case 'none':
      return { borders: none };
    case 'light1': {
      const border = thin(acc);
      return {
        fill: banded || bandedCol ? { type: 'solid', color: `${acc}+80` } : undefined,
        bold: isHeader || isTotal || isFirstCol || isLastCol,
        borders: [r === 0 || isTotal ? border : null, null, r === nR - 1 || isHeader ? border : null, null],
      };
    }
    case 'light2': {
      const border = thin(acc);
      return {
        fill: isHeader ? { type: 'solid', color: acc } : undefined,
        color: isHeader ? '@bg1' : undefined,
        bold: isHeader || isTotal || isFirstCol || isLastCol,
        borders: [border, border, border, border],
      };
    }
    case 'medium1': {
      const border = thin(acc);
      return {
        fill: isHeader ? { type: 'solid', color: acc } : banded || bandedCol ? { type: 'solid', color: `${acc}+80` } : { type: 'solid', color: '@bg1' },
        color: isHeader ? '@bg1' : undefined,
        bold: isHeader || isTotal || isFirstCol || isLastCol,
        borders: [r === 0 ? border : null, c === nC - 1 ? border : null, border, c === 0 ? border : null],
      };
    }
    case 'dark1': {
      const white = thin('@bg1');
      return {
        fill: { type: 'solid', color: isHeader || isTotal || isFirstCol || isLastCol ? '@tx1' : banded || bandedCol ? `${acc}-25` : acc },
        color: '@bg1',
        bold: isHeader || isTotal || isFirstCol || isLastCol,
        borders: [isTotal ? thin('@bg1', 2.66) : null, null, isHeader ? white : null, null],
      };
    }
    default: {
      // Medium Style 2: tinted body, stronger bands, solid accent header, white grid lines
      const white = thin('@bg1');
      const accentCell = isHeader || isTotal || isFirstCol || isLastCol;
      return {
        fill: { type: 'solid', color: accentCell ? acc : banded || bandedCol ? `${acc}+60` : `${acc}+80` },
        color: accentCell ? '@bg1' : undefined,
        bold: accentCell,
        borders: [isTotal ? thin('@bg1', 4) : white, white, isHeader ? thin('@bg1', 4) : white, white],
      };
    }
  }
}

/** A cell's text with the table style's text colour and weight as defaults (explicit formatting still wins). */
export function styledCellText(t: TableEl, r: number, c: number, text: TextBody = t.rows[r]?.cells[c]?.text ?? { paras: [{ runs: [] }] }): TextBody {
  const look = cellLook(t, r, c);
  if (!look.color && !look.bold) return text;
  return { ...text, defaults: { ...(look.color ? { color: look.color } : {}), ...(look.bold ? { b: true } : {}), ...(text.defaults ?? {}) } };
}

export function emptyCell(): TableCell {
  return { text: { paras: [{ runs: [] }], anchor: 't' } as TextBody };
}

export function createTable(rows: number, cols: number, x: number, y: number, w: number, rowH = 40, accent: ThemeColorName = 'accent1'): TableEl {
  const cw = Math.round(w / cols);
  return {
    id: newId(),
    type: 'table',
    name: 'Table',
    x,
    y,
    w: cw * cols,
    h: rowH * rows,
    cols: Array.from({ length: cols }, () => cw),
    rows: Array.from({ length: rows }, () => ({ h: rowH, cells: Array.from({ length: cols }, emptyCell) })),
    style: { family: 'medium2', accent, firstRow: true, bandRows: true },
  };
}

/** Keeps the table frame in sync with its rows and columns. */
export function normalizeTable(t: TableEl): TableEl {
  return { ...t, w: t.cols.reduce((a, b) => a + b, 0), h: t.rows.reduce((a, r) => a + r.h, 0) };
}

export function insertRow(t: TableEl, at: number): TableEl {
  const ref = t.rows[Math.max(0, Math.min(t.rows.length - 1, at - 1))];
  const row = { h: ref?.h ?? 40, cells: t.cols.map(() => emptyCell()) };
  const rows = [...t.rows];
  rows.splice(at, 0, row);
  return normalizeTable({ ...t, rows });
}

export function insertCol(t: TableEl, at: number): TableEl {
  const w = t.cols[Math.max(0, Math.min(t.cols.length - 1, at - 1))] ?? 120;
  const cols = [...t.cols];
  cols.splice(at, 0, w);
  return normalizeTable({ ...t, cols, rows: t.rows.map((r) => ({ ...r, cells: [...r.cells.slice(0, at), emptyCell(), ...r.cells.slice(at)] })) });
}

export function deleteRow(t: TableEl, at: number): TableEl | null {
  if (t.rows.length <= 1) return null;
  return normalizeTable({ ...t, rows: t.rows.filter((_, i) => i !== at) });
}

export function deleteCol(t: TableEl, at: number): TableEl | null {
  if (t.cols.length <= 1) return null;
  return normalizeTable({ ...t, cols: t.cols.filter((_, i) => i !== at), rows: t.rows.map((r) => ({ ...r, cells: r.cells.filter((_, i) => i !== at) })) });
}

export function mergeCells(t: TableEl, r1: number, c1: number, r2: number, c2: number): TableEl {
  const rows = t.rows.map((row, r) => ({
    ...row,
    cells: row.cells.map((cell, c) => {
      if (r < r1 || r > r2 || c < c1 || c > c2) return cell;
      if (r === r1 && c === c1) return { ...cell, rowSpan: r2 - r1 + 1, colSpan: c2 - c1 + 1, merged: false };
      return { ...cell, merged: true, rowSpan: undefined, colSpan: undefined };
    }),
  }));
  return { ...t, rows };
}

export function splitCell(t: TableEl, r: number, c: number): TableEl {
  const cell = t.rows[r]?.cells[c];
  if (!cell) return t;
  const rs = cell.rowSpan ?? 1;
  const cs = cell.colSpan ?? 1;
  const rows = t.rows.map((row, ri) => ({
    ...row,
    cells: row.cells.map((x, ci) => {
      if (ri < r || ri >= r + rs || ci < c || ci >= c + cs) return x;
      return { ...x, merged: false, rowSpan: undefined, colSpan: undefined };
    }),
  }));
  return { ...t, rows };
}
