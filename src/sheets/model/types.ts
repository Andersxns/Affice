import type { Scalar } from '../engine/values';
import type { Range } from './address';

export type BorderStyle =
  | 'thin'
  | 'medium'
  | 'thick'
  | 'dashed'
  | 'dotted'
  | 'double'
  | 'hair'
  | 'mediumDashed'
  | 'dashDot'
  | 'mediumDashDot'
  | 'dashDotDot'
  | 'mediumDashDotDot'
  | 'slantDashDot';

export interface Border {
  style: BorderStyle;
  color?: string;
}

export type HAlign = 'general' | 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'distributed';
export type VAlign = 'top' | 'middle' | 'bottom';

/** Cell formatting. Everything optional; `{}` is the workbook default. */
export interface CellStyle {
  font?: string;
  size?: number; // points
  bold?: boolean;
  italic?: boolean;
  underline?: 'single' | 'double';
  strike?: boolean;
  color?: string;
  fill?: string;
  hAlign?: HAlign;
  vAlign?: VAlign;
  wrap?: boolean;
  shrink?: boolean;
  indent?: number;
  rotation?: number; // -90..90, 255 = vertical text
  numFmt?: string;
  bt?: Border;
  br?: Border;
  bb?: Border;
  bl?: Border;
  locked?: boolean;
  hideFormula?: boolean;
}

export interface Cell {
  /** Literal value, or the cached result of the formula. */
  v?: Scalar;
  /** Formula text without the leading '='. */
  f?: string;
  /** Style id in the workbook style table (0 = default). */
  s?: number;
  link?: string;
  note?: string;
}

export interface ColInfo {
  w?: number; // px at 100%
  hidden?: boolean;
  s?: number;
  level?: number; // outline level
}

export interface RowInfo {
  h?: number; // px at 100%
  hidden?: boolean;
  s?: number;
  custom?: boolean; // height set by the user (no auto-fit)
  level?: number;
}

export type CompareOp = 'between' | 'notBetween' | 'equal' | 'notEqual' | 'greater' | 'less' | 'greaterEqual' | 'lessEqual';

export type CondType =
  | 'cell'
  | 'text'
  | 'begins'
  | 'ends'
  | 'notText'
  | 'top'
  | 'average'
  | 'duplicate'
  | 'unique'
  | 'blank'
  | 'notBlank'
  | 'error'
  | 'notError'
  | 'date'
  | 'formula'
  | 'colorScale'
  | 'dataBar'
  | 'iconSet';

export interface CondFormat {
  id: string;
  ranges: Range[];
  type: CondType;
  op?: CompareOp;
  /** Operands: numbers/text or formulas (starting with '=') relative to the top-left cell. */
  values?: string[];
  text?: string;
  rank?: number;
  percent?: boolean;
  bottom?: boolean;
  above?: boolean;
  equalAverage?: boolean;
  datePeriod?: 'yesterday' | 'today' | 'tomorrow' | 'last7Days' | 'lastWeek' | 'thisWeek' | 'nextWeek' | 'lastMonth' | 'thisMonth' | 'nextMonth';
  style?: Pick<CellStyle, 'bold' | 'italic' | 'underline' | 'strike' | 'color' | 'fill' | 'numFmt' | 'bt' | 'br' | 'bb' | 'bl'>;
  /** Color scale stops (2 or 3), data bar color, icon set name. */
  colors?: string[];
  iconSet?: '3Arrows' | '3TrafficLights' | '3Symbols' | '3Stars' | '4Arrows' | '5Arrows' | '3Flags' | '5Ratings';
  showValue?: boolean;
  stop?: boolean;
}

export type ValidationType = 'any' | 'list' | 'whole' | 'decimal' | 'date' | 'time' | 'textLength' | 'custom' | 'checkbox';

export interface Validation {
  id: string;
  ranges: Range[];
  type: ValidationType;
  op?: CompareOp;
  /** list: items or a range formula ('=$A$1:$A$5'); others: formula1 [, formula2]. */
  values: string[];
  allowBlank?: boolean;
  showDropdown?: boolean;
  errorStyle?: 'stop' | 'warning' | 'information';
  errorTitle?: string;
  error?: string;
  promptTitle?: string;
  prompt?: string;
}

export type ChartType = 'column' | 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'scatter' | 'radar' | 'polarArea' | 'bubble' | 'combo';

export interface ChartSpec {
  id: string;
  type: ChartType;
  /** Source data, e.g. "Sheet1!$A$1:$C$10". */
  source: string;
  seriesInRows?: boolean;
  headerRow?: boolean;
  headerCol?: boolean;
  stacked?: boolean | 'percent';
  title?: string;
  xTitle?: string;
  yTitle?: string;
  legend?: 'top' | 'bottom' | 'left' | 'right' | 'none';
  palette?: string;
  smooth?: boolean;
  dataLabels?: boolean;
  gridlines?: boolean;
  /** Anchored at a cell with a px offset; size in px. */
  anchor: { r: number; c: number; dx: number; dy: number };
  w: number;
  h: number;
}

export interface ImageSpec {
  id: string;
  src: string; // data URL
  anchor: { r: number; c: number; dx: number; dy: number };
  w: number;
  h: number;
  alt?: string;
}

export interface ColumnFilter {
  /** Allowed display values (null = everything). */
  values?: string[];
  blanks?: boolean;
  condition?: { op: CompareOp | 'contains' | 'begins' | 'ends' | 'notContains'; value: string; value2?: string };
  top?: { n: number; bottom?: boolean; percent?: boolean };
  color?: string;
}

export interface AutoFilter {
  range: Range;
  columns: Record<number, ColumnFilter>;
  sort?: { col: number; desc: boolean };
}

export interface DefinedName {
  name: string;
  /** Formula text without '=' (usually an absolute reference). */
  ref: string;
  /** Sheet id for sheet-scoped names. */
  sheet?: number;
  comment?: string;
  hidden?: boolean;
}

export interface SheetView {
  zoom: number;
  showGrid: boolean;
  showHeaders: boolean;
  showFormulas: boolean;
  rtl: boolean;
  freezeRows: number;
  freezeCols: number;
  scrollRow: number;
  scrollCol: number;
  selection?: { ranges: Range[]; active: { r: number; c: number } };
}

export interface SheetProtection {
  enabled: boolean;
  password?: string; // simple hash
  allowFormat?: boolean;
  allowInsert?: boolean;
  allowDelete?: boolean;
  allowSort?: boolean;
  allowFilter?: boolean;
}

export const DEFAULT_COL_WIDTH = 88;
export const DEFAULT_ROW_HEIGHT = 22;
