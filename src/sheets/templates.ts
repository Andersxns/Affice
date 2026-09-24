import { partsToSerial, SHORT_DATE } from './format/numfmt';
import { cellKey, parseAddr, parseRange } from './model/address';
import { Workbook, type Sheet } from './model/workbook';
import type { CellStyle, CondFormat, Validation } from './model/types';

class Builder {
  constructor(
    public wb: Workbook,
    public sheet: Sheet,
  ) {}
  set(addr: string, v: string | number | boolean | null, st?: CellStyle): this {
    const p = parseAddr(addr)!;
    const k = cellKey(p.r, p.c);
    const cur = this.sheet.cells.get(k) ?? {};
    const cell = { ...cur };
    if (typeof v === 'string' && v.startsWith('=')) {
      cell.f = v.slice(1);
      delete cell.v;
    } else if (v !== null) cell.v = v;
    if (st) cell.s = this.wb.patchStyle(cell.s, st);
    this.sheet.put(k, cell);
    return this;
  }
  row(addr: string, values: (string | number | boolean | null)[], st?: CellStyle): this {
    const p = parseAddr(addr)!;
    values.forEach((v, i) => {
      if (v === null && !st) return;
      const k = cellKey(p.r, p.c + i);
      const cur = this.sheet.cells.get(k) ?? {};
      const cell = { ...cur };
      if (typeof v === 'string' && v.startsWith('=')) cell.f = v.slice(1);
      else if (v !== null) cell.v = v;
      if (st) cell.s = this.wb.patchStyle(cell.s, st);
      this.sheet.put(k, cell);
    });
    return this;
  }
  style(range: string, st: CellStyle): this {
    const rg = parseRange(range)!;
    for (let r = rg.r1; r <= rg.r2; r++)
      for (let c = rg.c1; c <= rg.c2; c++) {
        const k = cellKey(r, c);
        const cur = this.sheet.cells.get(k) ?? {};
        this.sheet.put(k, { ...cur, s: this.wb.patchStyle(cur.s, st) });
      }
    return this;
  }
  width(col: number, px: number): this {
    this.sheet.cols.set(col, { ...(this.sheet.cols.get(col) ?? {}), w: px });
    return this;
  }
  height(row: number, px: number): this {
    this.sheet.rows.set(row, { ...(this.sheet.rows.get(row) ?? {}), h: px, custom: true });
    return this;
  }
  merge(range: string): this {
    this.sheet.merges.push(parseRange(range)!);
    return this;
  }
  cf(rule: Omit<CondFormat, 'id' | 'ranges'> & { range: string }): this {
    const { range, ...rest } = rule;
    this.sheet.cf.push({ id: `t${this.sheet.cf.length}${Math.random().toString(36).slice(2, 6)}`, ranges: range.split(',').map((x) => parseRange(x.trim())!), ...rest });
    return this;
  }
  validate(range: string, v: Omit<Validation, 'id' | 'ranges'>): this {
    this.sheet.validations.push({ id: `v${this.sheet.validations.length}`, ranges: [parseRange(range)!], ...v });
    return this;
  }
}

const thin = (color = '#d9dce3') => ({ style: 'thin' as const, color });
const MONEY = '$#,##0.00';

function budget(): Workbook {
  const wb = new Workbook();
  const s = wb.addSheet('Budget');
  const b = new Builder(wb, s);
  const green = '#17a35a';
  s.view.showGrid = false;
  [28, 190, 120, 120, 120, 28, 170, 120].forEach((w, i) => b.width(i, w));
  b.height(1, 44);
  b.set('B2', 'Monthly budget', { size: 22, bold: true, color: '#0f3d27', font: 'Montserrat' });
  b.set('B3', 'Plan your month and track where the money goes.', { color: '#6b7280', italic: true });
  // income
  b.row('B5', ['Income', 'Planned', 'Actual', 'Difference'], { bold: true, color: '#ffffff', fill: green });
  const income = [
    ['Salary', 4200, 4200],
    ['Side projects', 600, 740],
    ['Other', 100, 45],
  ];
  income.forEach(([n, p, a], i) => {
    const r = 6 + i;
    b.row(`B${r}`, [n, p, a, `=D${r}-C${r}`]);
  });
  b.row('B9', ['Total income', '=SUM(C6:C8)', '=SUM(D6:D8)', '=D9-C9'], { bold: true, bt: thin('#17a35a') });
  // expenses
  b.row('B11', ['Expenses', 'Planned', 'Actual', 'Difference'], { bold: true, color: '#ffffff', fill: '#0f3d27' });
  const expenses: [string, number, number][] = [
    ['Housing', 1400, 1400],
    ['Groceries', 520, 588],
    ['Transport', 240, 196],
    ['Utilities', 180, 204],
    ['Insurance', 150, 150],
    ['Eating out', 200, 265],
    ['Health', 90, 40],
    ['Entertainment', 120, 98],
    ['Savings', 800, 800],
    ['Other', 100, 64],
  ];
  expenses.forEach(([n, p, a], i) => {
    const r = 12 + i;
    b.row(`B${r}`, [n, p, a, `=C${r}-D${r}`]);
    if (i % 2 === 1) b.style(`B${r}:E${r}`, { fill: '#f1faf5' });
  });
  const last = 12 + expenses.length - 1;
  b.row(`B${last + 1}`, ['Total expenses', `=SUM(C12:C${last})`, `=SUM(D12:D${last})`, `=C${last + 1}-D${last + 1}`], { bold: true, bt: thin('#0f3d27') });
  b.style(`C6:E${last + 1}`, { numFmt: MONEY });
  // summary
  b.set('H5', 'Summary', { bold: true, size: 13, color: '#0f3d27' });
  b.row('H6', ['Balance', `=D9-D${last + 1}`]);
  b.row('H7', ['Saved this month', '=D20']);
  b.row('H8', ['Spent vs plan', `=D${last + 1}/C${last + 1}`]);
  b.row('H9', ['Biggest expense', `=INDEX(B12:B${last},MATCH(MAX(D12:D${last}),D12:D${last},0))`]);
  b.style('H6:H9', { color: '#374151' });
  b.style('I6:I7', { numFmt: MONEY, bold: true });
  b.style('I8', { numFmt: '0.0%', bold: true });
  b.style('I9', { bold: true, hAlign: 'right' });
  b.style('H6:I9', { bb: thin() });
  b.cf({ range: `E6:E8,E12:E${last}`, type: 'cell', op: 'less', values: ['0'], style: { color: '#c0392b', bold: true } });
  b.cf({ range: `D12:D${last}`, type: 'dataBar', colors: ['#17a35a'] });
  s.charts.push({ id: 'budget-chart', type: 'bar', source: `Budget!$B$11:$D$${last}`, title: 'Planned vs actual', anchor: { r: 11, c: 7, dx: 0, dy: 0 }, w: 420, h: 330, legend: 'bottom', palette: 'affice', seriesInRows: false });
  return wb;
}

function invoice(): Workbook {
  const wb = new Workbook();
  const s = wb.addSheet('Invoice');
  const b = new Builder(wb, s);
  const blue = '#2f6dff';
  s.view.showGrid = false;
  [24, 250, 80, 110, 120, 24].forEach((w, i) => b.width(i, w));
  b.height(1, 46);
  b.set('B2', 'INVOICE', { size: 26, bold: true, color: blue, font: 'Montserrat' });
  b.set('E2', 'Your Company', { size: 14, bold: true, hAlign: 'right' });
  b.set('E3', '123 Main Street, Springfield', { hAlign: 'right', color: '#6b7280' });
  b.set('E4', 'hello@yourcompany.com', { hAlign: 'right', color: '#6b7280' });
  b.row('B6', ['Bill to', null, 'Invoice #', 'INV-0042'], { bold: false });
  b.style('B6', { bold: true, color: blue });
  b.style('D6:D8', { bold: true, color: '#374151', hAlign: 'right' });
  b.set('B7', 'Customer Name');
  b.set('B8', 'Customer address');
  b.row('D7', ['Date', '=TODAY()']);
  b.row('D8', ['Due', '=E7+30']);
  b.style('E7:E8', { numFmt: SHORT_DATE, hAlign: 'right' });
  b.style('E6', { hAlign: 'right' });
  b.row('B10', ['Description', 'Qty', 'Unit price', 'Amount'], { bold: true, color: '#ffffff', fill: blue });
  const items: [string, number, number][] = [
    ['Website design', 1, 1800],
    ['Hosting (12 months)', 12, 15],
    ['Logo refresh', 1, 450],
    ['Content writing (per page)', 6, 60],
    ['', 0, 0],
    ['', 0, 0],
  ];
  items.forEach(([d, q, p], i) => {
    const r = 11 + i;
    if (d) b.row(`B${r}`, [d, q, p, `=IF(C${r}="","",C${r}*D${r})`]);
    else b.set(`E${r}`, `=IF(C${r}="","",C${r}*D${r})`);
    b.style(`B${r}:E${r}`, { bb: thin() });
  });
  b.style('D11:E16', { numFmt: MONEY });
  b.style('C11:C16', { hAlign: 'center' });
  b.row('D18', ['Subtotal', '=SUM(E11:E16)']);
  b.row('D19', ['Tax rate', 0.08]);
  b.row('D20', ['Tax', '=E18*E19']);
  b.row('D21', ['Total', '=E18+E20'], { bold: true, size: 13, color: blue });
  b.style('D18:D21', { hAlign: 'right' });
  b.style('E18', { numFmt: MONEY });
  b.style('E19', { numFmt: '0.0%' });
  b.style('E20:E21', { numFmt: MONEY });
  b.style('D21:E21', { bt: { style: 'medium', color: blue } });
  b.set('B23', 'Thank you for your business!', { italic: true, color: '#6b7280' });
  b.set('B24', 'Payment within 30 days. Bank: 0000 0000 0000', { color: '#9ca3af', size: 9 });
  return wb;
}

function tracker(): Workbook {
  const wb = new Workbook();
  const s = wb.addSheet('Tasks');
  const b = new Builder(wb, s);
  const orange = '#f26a26';
  [250, 110, 110, 90, 105, 110, 90].forEach((w, i) => b.width(i, w));
  b.row('A1', ['Task', 'Owner', 'Status', 'Priority', 'Due date', 'Progress', 'Days left'], { bold: true, color: '#ffffff', fill: orange });
  b.height(0, 28);
  const today = Math.floor(partsToSerial(new Date().getFullYear(), new Date().getMonth() + 1, new Date().getDate()));
  const tasks: [string, string, string, string, number, number][] = [
    ['Kick-off meeting', 'Alex', 'Done', 'High', -10, 1],
    ['Collect requirements', 'Sam', 'Done', 'High', -4, 1],
    ['Design mock-ups', 'Priya', 'In progress', 'Medium', 3, 0.6],
    ['Build prototype', 'Jordan', 'In progress', 'High', 9, 0.3],
    ['User testing', 'Sam', 'Not started', 'Medium', 16, 0],
    ['Fix feedback', 'Jordan', 'Not started', 'Low', 21, 0],
    ['Launch', 'Alex', 'Blocked', 'High', 28, 0],
  ];
  tasks.forEach(([t, o, st, p, d, pr], i) => {
    const r = i + 2;
    b.row(`A${r}`, [t, o, st, p, today + d, pr, `=IF(C${r}="Done","",E${r}-TODAY())`]);
  });
  const lastRow = tasks.length + 1;
  b.style(`E2:E${lastRow + 20}`, { numFmt: 'd mmm yyyy' });
  b.style(`F2:F${lastRow + 20}`, { numFmt: '0%' });
  b.style(`A2:G${lastRow}`, { bb: thin() });
  b.validate(`C2:C${lastRow + 50}`, { type: 'list', values: ['Not started', 'In progress', 'Blocked', 'Done'], allowBlank: true, showDropdown: true });
  b.validate(`D2:D${lastRow + 50}`, { type: 'list', values: ['High', 'Medium', 'Low'], allowBlank: true, showDropdown: true });
  b.cf({ range: `C2:C${lastRow + 50}`, type: 'cell', op: 'equal', values: ['Done'], style: { fill: '#d7f2e3', color: '#0a6b38' } });
  b.cf({ range: `C2:C${lastRow + 50}`, type: 'cell', op: 'equal', values: ['Blocked'], style: { fill: '#fbdadb', color: '#9c0006' } });
  b.cf({ range: `C2:C${lastRow + 50}`, type: 'cell', op: 'equal', values: ['In progress'], style: { fill: '#fff1d6', color: '#9c5700' } });
  b.cf({ range: `F2:F${lastRow + 50}`, type: 'dataBar', colors: [orange] });
  b.cf({ range: `G2:G${lastRow + 50}`, type: 'formula', values: ['=AND(ISNUMBER(G2),G2<0)'], style: { color: '#c0392b', bold: true } });
  b.cf({ range: `D2:D${lastRow + 50}`, type: 'cell', op: 'equal', values: ['High'], style: { bold: true } });
  s.view.freezeRows = 1;
  s.filter = { range: { r1: 0, c1: 0, r2: lastRow, c2: 6 }, columns: {} };
  // summary sheet
  const sum = wb.addSheet('Summary');
  const b2 = new Builder(wb, sum);
  b2.width(0, 160).width(1, 90);
  b2.row('A1', ['Status', 'Tasks'], { bold: true, color: '#ffffff', fill: orange });
  ['Not started', 'In progress', 'Blocked', 'Done'].forEach((st, i) => b2.row(`A${i + 2}`, [st, `=COUNTIF(Tasks!C:C,A${i + 2})`]));
  b2.row('A7', ['Overall progress', '=AVERAGE(Tasks!F2:F100)'], { bold: true });
  b2.style('B7', { numFmt: '0%' });
  sum.charts.push({ id: 'status-chart', type: 'pie', source: 'Summary!$A$1:$B$5', title: 'Tasks by status', anchor: { r: 0, c: 3, dx: 0, dy: 0 }, w: 360, h: 260, legend: 'right', palette: 'affice' });
  return wb;
}

function grades(): Workbook {
  const wb = new Workbook();
  const s = wb.addSheet('Gradebook');
  const b = new Builder(wb, s);
  const purple = '#7c3aed';
  [170, 90, 90, 90, 90, 90, 70, 60, 24, 90, 60].forEach((w, i) => b.width(i, w));
  b.row('A1', ['Student', 'Essay', 'Quiz 1', 'Quiz 2', 'Project', 'Average', 'Grade', 'Rank'], { bold: true, color: '#ffffff', fill: purple, hAlign: 'center' });
  b.style('A1', { hAlign: 'left' });
  const names = ['Ava Martin', 'Ben Carter', 'Chloe Nguyen', 'Diego Alvarez', 'Emma Johnson', 'Farah Khan', 'George Smith', 'Hana Suzuki', 'Isaac Brown', 'Julia Rossi'];
  const scores = [
    [92, 88, 95, 90],
    [75, 80, 68, 82],
    [88, 94, 91, 97],
    [64, 72, 70, 66],
    [95, 90, 98, 93],
    [81, 77, 85, 88],
    [58, 65, 61, 70],
    [90, 96, 89, 94],
    [72, 69, 78, 74],
    [85, 83, 80, 91],
  ];
  names.forEach((n, i) => {
    const r = i + 2;
    b.row(`A${r}`, [n, ...scores[i], `=AVERAGE(B${r}:E${r})`, `=XLOOKUP(F${r},$J$2:$J$6,$K$2:$K$6,,-1)`, `=RANK(F${r},$F$2:$F$11)`]);
  });
  b.style('B2:E11', { hAlign: 'center' });
  b.style('F2:F11', { numFmt: '0.0', bold: true, hAlign: 'center' });
  b.style('G2:H11', { hAlign: 'center' });
  b.style('A2:H11', { bb: thin() });
  b.row('A13', ['Class average', '=AVERAGE(B2:B11)', '=AVERAGE(C2:C11)', '=AVERAGE(D2:D11)', '=AVERAGE(E2:E11)', '=AVERAGE(F2:F11)'], { bold: true, bt: { style: 'medium', color: purple } });
  b.row('A14', ['Highest', '=MAX(B2:B11)', '=MAX(C2:C11)', '=MAX(D2:D11)', '=MAX(E2:E11)', '=MAX(F2:F11)']);
  b.row('A15', ['Lowest', '=MIN(B2:B11)', '=MIN(C2:C11)', '=MIN(D2:D11)', '=MIN(E2:E11)', '=MIN(F2:F11)']);
  b.style('B13:F15', { numFmt: '0.0', hAlign: 'center' });
  b.row('J1', ['From', 'Grade'], { bold: true, color: '#ffffff', fill: '#4c1d95', hAlign: 'center' });
  [
    [0, 'F'],
    [60, 'D'],
    [70, 'C'],
    [80, 'B'],
    [90, 'A'],
  ].forEach(([f, g], i) => b.row(`J${i + 2}`, [f, g], { hAlign: 'center' }));
  b.cf({ range: 'F2:F11', type: 'colorScale', colors: ['#f8696b', '#ffeb84', '#63be7b'] });
  b.cf({ range: 'G2:G11', type: 'cell', op: 'equal', values: ['F'], style: { color: '#c0392b', bold: true } });
  s.view.freezeRows = 1;
  s.view.freezeCols = 1;
  s.charts.push({ id: 'grades-chart', type: 'column', source: 'Gradebook!$A$1:$E$11', title: 'Scores by student', anchor: { r: 16, c: 0, dx: 0, dy: 8 }, w: 620, h: 280, legend: 'bottom', palette: 'affice', seriesInRows: false });
  return wb;
}

function schedule(): Workbook {
  const wb = new Workbook();
  const s = wb.addSheet('Week');
  const b = new Builder(wb, s);
  const teal = '#0891b2';
  s.view.showGrid = false;
  b.width(0, 70);
  for (let c = 1; c <= 7; c++) b.width(c, 120);
  b.set('A1', 'Weekly schedule', { size: 20, bold: true, color: '#0e4f5f', font: 'Montserrat' });
  b.merge('A1:H1');
  b.height(0, 40);
  const days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
  b.row('B3', days, { bold: true, color: '#ffffff', fill: teal, hAlign: 'center' });
  b.set('A3', 'Time', { bold: true, color: '#ffffff', fill: teal, hAlign: 'center' });
  for (let h = 7; h <= 21; h++) {
    const r = h - 7 + 4;
    b.set(`A${r}`, h / 24, { numFmt: 'h:mm AM/PM', color: '#6b7280', hAlign: 'right', vAlign: 'top' });
    b.style(`A${r}:H${r}`, { bb: thin('#e5e7eb') });
    b.height(r - 1, 30);
  }
  const events: [string, number, number, string, string][] = [
    ['B', 9, 10, 'Team stand-up', '#cffafe'],
    ['C', 9, 10, 'Team stand-up', '#cffafe'],
    ['D', 9, 10, 'Team stand-up', '#cffafe'],
    ['E', 9, 10, 'Team stand-up', '#cffafe'],
    ['F', 9, 10, 'Team stand-up', '#cffafe'],
    ['B', 7, 8, 'Gym', '#dcfce7'],
    ['D', 7, 8, 'Gym', '#dcfce7'],
    ['F', 7, 8, 'Gym', '#dcfce7'],
    ['C', 13, 15, 'Project work', '#ede9fe'],
    ['E', 13, 15, 'Project work', '#ede9fe'],
    ['G', 10, 12, 'Farmers market', '#fef3c7'],
    ['H', 18, 20, 'Family dinner', '#ffe4e6'],
    ['F', 17, 18, 'Weekly review', '#e0f2fe'],
  ];
  for (const [col, from, to, what, fill] of events) {
    const r1 = from - 7 + 4;
    const r2 = to - 7 + 3;
    b.set(`${col}${r1}`, what, { fill, bold: true, color: '#1f2937', vAlign: 'top', wrap: true });
    for (let r = r1; r <= r2; r++) b.style(`${col}${r}`, { fill });
    if (r2 > r1) b.merge(`${col}${r1}:${col}${r2}`);
  }
  s.view.freezeRows = 3;
  s.view.freezeCols = 1;
  return wb;
}

export function sheetTemplate(id: string): Workbook {
  switch (id) {
    case 'sheet-budget':
      return budget();
    case 'sheet-invoice':
      return invoice();
    case 'sheet-tracker':
      return tracker();
    case 'sheet-grades':
      return grades();
    case 'sheet-schedule':
      return schedule();
  }
  const wb = new Workbook();
  wb.addSheet();
  return wb;
}
