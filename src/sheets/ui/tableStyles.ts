import type { CellStyle } from '../model/types';

export interface TableStyle {
  id: string;
  label: string;
  header: CellStyle;
  band1: CellStyle;
  band2: CellStyle;
  total?: CellStyle;
}

const mk = (id: string, label: string, accent: string, light: string, lighter = '#ffffff', headerText = '#ffffff'): TableStyle => ({
  id,
  label,
  header: { fill: accent, color: headerText, bold: true, bb: { style: 'thin', color: accent } },
  band1: { fill: light, bb: { style: 'thin', color: light } },
  band2: { fill: lighter, bb: { style: 'thin', color: light } },
  total: { bold: true, bt: { style: 'double', color: accent } },
});

export const TABLE_STYLES: TableStyle[] = [
  mk('blue', 'Blue', '#2f6dff', '#dce7ff'),
  mk('green', 'Green', '#17a35a', '#d7f2e3'),
  mk('orange', 'Orange', '#f26a26', '#fde3d4'),
  mk('purple', 'Purple', '#8e4ec6', '#ecdff8'),
  mk('red', 'Red', '#e5484d', '#fbdadb'),
  mk('teal', 'Teal', '#12a594', '#d2f1ec'),
  mk('gray', 'Gray', '#5b6475', '#e6e8ec'),
  mk('navy', 'Navy', '#1f3b73', '#dbe2f0'),
  mk('gold', 'Gold', '#c28a00', '#fbefcc', '#ffffff', '#ffffff'),
  { id: 'plain', label: 'Plain', header: { bold: true, bb: { style: 'medium', color: '#1a1d24' } }, band1: { bb: { style: 'thin', color: '#d9dce3' } }, band2: { bb: { style: 'thin', color: '#d9dce3' } } },
  { id: 'dark', label: 'Dark', header: { fill: '#1f2128', color: '#ffffff', bold: true }, band1: { fill: '#3a3d47', color: '#ffffff' }, band2: { fill: '#2b2e36', color: '#ffffff' } },
];

export interface NamedCellStyle {
  id: string;
  label: string;
  group: 'Good, bad and neutral' | 'Titles and headings' | 'Data and model' | 'Themed' | 'Number format';
  style: CellStyle;
}

export const CELL_STYLES: NamedCellStyle[] = [
  { id: 'normal', label: 'Normal', group: 'Good, bad and neutral', style: {} },
  { id: 'good', label: 'Good', group: 'Good, bad and neutral', style: { fill: '#c6efce', color: '#006100' } },
  { id: 'bad', label: 'Bad', group: 'Good, bad and neutral', style: { fill: '#ffc7ce', color: '#9c0006' } },
  { id: 'neutral', label: 'Neutral', group: 'Good, bad and neutral', style: { fill: '#ffeb9c', color: '#9c5700' } },
  { id: 'title', label: 'Title', group: 'Titles and headings', style: { size: 18, bold: true, color: '#1f3b73', font: 'Cambria' } },
  { id: 'h1', label: 'Heading 1', group: 'Titles and headings', style: { size: 15, bold: true, color: '#1f3b73', bb: { style: 'thick', color: '#4472c4' } } },
  { id: 'h2', label: 'Heading 2', group: 'Titles and headings', style: { size: 13, bold: true, color: '#1f3b73', bb: { style: 'thick', color: '#a9c4eb' } } },
  { id: 'h3', label: 'Heading 3', group: 'Titles and headings', style: { size: 11, bold: true, color: '#1f3b73', bb: { style: 'medium', color: '#95b3d7' } } },
  { id: 'total', label: 'Total', group: 'Titles and headings', style: { bold: true, bt: { style: 'thin', color: '#4472c4' }, bb: { style: 'double', color: '#4472c4' } } },
  { id: 'input', label: 'Input', group: 'Data and model', style: { fill: '#ffcc99', color: '#3f3f76', bt: { style: 'thin', color: '#7f7f7f' }, bb: { style: 'thin', color: '#7f7f7f' }, bl: { style: 'thin', color: '#7f7f7f' }, br: { style: 'thin', color: '#7f7f7f' } } },
  { id: 'output', label: 'Output', group: 'Data and model', style: { fill: '#f2f2f2', color: '#3f3f3f', bold: true, bt: { style: 'thin', color: '#3f3f3f' }, bb: { style: 'thin', color: '#3f3f3f' }, bl: { style: 'thin', color: '#3f3f3f' }, br: { style: 'thin', color: '#3f3f3f' } } },
  { id: 'calc', label: 'Calculation', group: 'Data and model', style: { fill: '#f2f2f2', color: '#fa7d00', bold: true, bt: { style: 'thin', color: '#7f7f7f' }, bb: { style: 'thin', color: '#7f7f7f' }, bl: { style: 'thin', color: '#7f7f7f' }, br: { style: 'thin', color: '#7f7f7f' } } },
  { id: 'check', label: 'Check cell', group: 'Data and model', style: { fill: '#a5a5a5', color: '#ffffff', bold: true, bt: { style: 'double', color: '#3f3f3f' }, bb: { style: 'double', color: '#3f3f3f' }, bl: { style: 'double', color: '#3f3f3f' }, br: { style: 'double', color: '#3f3f3f' } } },
  { id: 'note', label: 'Note', group: 'Data and model', style: { fill: '#ffffcc', bt: { style: 'thin', color: '#b2b2b2' }, bb: { style: 'thin', color: '#b2b2b2' }, bl: { style: 'thin', color: '#b2b2b2' }, br: { style: 'thin', color: '#b2b2b2' } } },
  { id: 'warn', label: 'Warning text', group: 'Data and model', style: { color: '#ff0000' } },
  { id: 'explan', label: 'Explanatory', group: 'Data and model', style: { italic: true, color: '#7f7f7f' } },
  { id: 'accent1', label: 'Accent 1', group: 'Themed', style: { fill: '#2f6dff', color: '#ffffff' } },
  { id: 'accent2', label: 'Accent 2', group: 'Themed', style: { fill: '#17a35a', color: '#ffffff' } },
  { id: 'accent3', label: 'Accent 3', group: 'Themed', style: { fill: '#f26a26', color: '#ffffff' } },
  { id: 'accent4', label: 'Accent 4', group: 'Themed', style: { fill: '#8e4ec6', color: '#ffffff' } },
  { id: 'accent1l', label: '20% Accent 1', group: 'Themed', style: { fill: '#dce7ff' } },
  { id: 'accent2l', label: '20% Accent 2', group: 'Themed', style: { fill: '#d7f2e3' } },
  { id: 'accent3l', label: '20% Accent 3', group: 'Themed', style: { fill: '#fde3d4' } },
  { id: 'accent4l', label: '20% Accent 4', group: 'Themed', style: { fill: '#ecdff8' } },
  { id: 'comma', label: 'Comma', group: 'Number format', style: { numFmt: '#,##0.00' } },
  { id: 'currency', label: 'Currency', group: 'Number format', style: { numFmt: '$#,##0.00' } },
  { id: 'percent', label: 'Percent', group: 'Number format', style: { numFmt: '0%' } },
];

/** Style properties a named style controls (so applying one resets the others). */
export const STYLE_KEYS: (keyof CellStyle)[] = ['font', 'size', 'bold', 'italic', 'underline', 'strike', 'color', 'fill', 'bt', 'bb', 'bl', 'br'];
