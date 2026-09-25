import { beforeAll, describe, expect, it } from 'vitest';
import { DOMParser } from '@xmldom/xmldom';
import { strToU8, zipSync } from 'fflate';
import { fromOdsFormula, toOdsFormula } from '../src/sheets/formats/odsFormula';
import { numFmtToOds, odsToNumFmt, odsValueType } from '../src/sheets/formats/odsNumber';
import { exportOds, importOds, odsDateToSerial, odsTimeToSerial, serialToOdsDate, serialToOdsTime } from '../src/sheets/formats/ods';
import { splitContent } from '../src/sheets/formats/odsScan';
import { strFromU8, unzipSync } from 'fflate';
import { formatValue } from '../src/sheets/format/numfmt';
import { Workbook, workbookToJSON, type Sheet } from '../src/sheets/model/workbook';
import { cellKey, parseAddr } from '../src/sheets/model/address';
import { SheetDoc } from '../src/sheets/doc';
import type { Cell } from '../src/sheets/model/types';

beforeAll(() => {
  // the reader uses the browser's DOMParser
  (globalThis as { DOMParser?: unknown }).DOMParser = DOMParser;
});

const NS_NUMBER = 'urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0';
const NS_STYLE = 'urn:oasis:names:tc:opendocument:xmlns:style:1.0';

describe('OpenFormula conversion', () => {
  const cases: [excel: string, ods: string][] = [
    ['B2*2', 'of:=[.B2]*2'],
    ['IF(B3>100,"big","small")', 'of:=IF([.B3]>100;"big";"small")'],
    ['SUM(B:B)', 'of:=SUM([.B:.B])'],
    ["SUM('Other sheet'!A1:A3)+Rate", "of:=SUM([$'Other sheet'.A1:.A3])+Rate"],
    ['Data!$B$4*2', 'of:=[$Data.$B$4]*2'],
    ['$B$2+B$3+$B4', 'of:=[.$B$2]+[.B$3]+[.$B4]'],
    ['SUMPRODUCT({1,2;3,4}*2)', 'of:=SUMPRODUCT({1;2|3;4}*2)'],
    ['IFS(B6>0,"pos",TRUE,"zero")', 'of:=COM.MICROSOFT.IFS([.B6]>0;"pos";TRUE();"zero")'],
    ['TEXTJOIN(", ",TRUE,A2:A4)', 'of:=COM.MICROSOFT.TEXTJOIN(", ";TRUE();[.A2:.A4])'],
    ['STDEV.S(B2:B6)', 'of:=COM.MICROSOFT.STDEV.S([.B2:.B6])'],
    ['FORMULATEXT(A1)', 'of:=FORMULA([.A1])'],
    ['SUM((A1,B1))', 'of:=SUM(([.A1]~[.B1]))'],
    ['SUM(A1:B5 B2:C3)', 'of:=SUM([.A1:.B5]![.B2:.C3])'],
    ['"say ""hi"""&A1', 'of:="say ""hi"""&[.A1]'],
    ['IFERROR(1/0,#N/A)', 'of:=IFERROR(1/0;#N/A)'],
  ];
  it.each(cases)('writes %s', (excel, ods) => expect(toOdsFormula(excel)).toBe(ods));
  it.each(cases)('reads back %s', (excel, ods) => expect(fromOdsFormula(ods)).toBe(excel));

  it('reads what LibreOffice writes', () => {
    expect(fromOdsFormula('of:=_xlfn.xlookup("Figs";[.A2:.A6];[.B2:.B6])')).toBe('XLOOKUP("Figs",A2:A6,B2:B6)');
    expect(fromOdsFormula('of:=[.A1:.B2]![.B2:.C3]')).toBe('A1:B2 B2:C3');
    expect(fromOdsFormula('of:=SUM([.A1]~[.B1])')).toBe('SUM(A1,B1)');
    expect(fromOdsFormula('of:=IF([.A1]=1;TRUE();FALSE())')).toBe('IF(A1=1,TRUE,FALSE)');
    expect(fromOdsFormula('of:=LEGACY.CHIDIST(1;2)')).toBe('CHIDIST(1,2)');
    expect(fromOdsFormula('of:=[.#REF!]+1')).toBe('#REF!+1');
    expect(fromOdsFormula('msoxl:=SUM(A1:A2)')).toBe('SUM(A1:A2)');
    expect(fromOdsFormula("of:=[$'It''s'.A1]")).toBe("'It''s'!A1");
  });
});

describe('OpenDocument number styles', () => {
  const roundTrip = (code: string): string | undefined => {
    const xml = numFmtToOds(code, 'N1');
    const doc = new DOMParser().parseFromString(`<r xmlns:number="${NS_NUMBER}" xmlns:style="${NS_STYLE}" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0">${xml}</r>`, 'application/xml');
    const styles = new Map<string, Element>();
    for (let n = doc.documentElement.firstChild; n; n = n.nextSibling) if (n.nodeType === 1) styles.set((n as Element).getAttributeNS(NS_STYLE, 'name') ?? '', n as Element);
    return odsToNumFmt(styles.get('N1')!, (name) => styles.get(name));
  };
  const codes = [
    '0',
    '0.00',
    '#,##0',
    '#,##0.00',
    '0%',
    '0.0%',
    '$#,##0.00;[Red]-$#,##0.00',
    '#,##0.00_);[Red](#,##0.00)',
    '0.00E+00',
    '# ?/?',
    '# ??/16',
    'h:mm AM/PM',
    '[h]:mm:ss',
    'mm:ss.0',
    'd-mmm-yyyy',
    'dddd, mmmm d, yyyy',
    'yyyy-mm-dd hh:mm:ss',
    '@',
    '#,##0,"K"',
    '0 "items"',
    '[Blue][>=100]0;[Red][<0]-0;0.00',
  ];
  const samples = [1234.5678, -1234.5678, 0, 0.5, 7.25, 46290.5625, 1.5104166, 150];
  it.each(codes)('keeps %s', (code) => {
    const back = roundTrip(code);
    expect(back).toBeDefined();
    for (const v of samples) expect(formatValue(v, back).text).toBe(formatValue(v, code).text);
  });
  it('treats General as no style', () => {
    expect(numFmtToOds('General', 'N1')).toBe('');
  });
  it('picks value types', () => {
    expect(odsValueType('0.0%')).toBe('percentage');
    expect(odsValueType('$#,##0.00')).toBe('currency');
    expect(odsValueType('d-mmm-yyyy')).toBe('date');
    expect(odsValueType('h:mm')).toBe('time');
    expect(odsValueType('[h]:mm:ss')).toBe('time');
    expect(odsValueType('0.00')).toBe('float');
    expect(odsValueType(undefined)).toBe('float');
  });
  it('reads decimals LibreOffice gives only as a minimum', () => {
    const doc = new DOMParser().parseFromString(
      `<r xmlns:number="${NS_NUMBER}" xmlns:style="${NS_STYLE}"><number:currency-style style:name="N1"><number:currency-symbol/><number:number number:min-decimal-places="2" number:min-integer-digits="1" number:grouping="true"/></number:currency-style></r>`,
      'application/xml',
    );
    const el = doc.getElementsByTagNameNS(NS_NUMBER, 'currency-style')[0];
    expect(odsToNumFmt(el, () => undefined)).toBe('$#,##0.00');
  });
});

describe('OpenDocument dates and times', () => {
  it('converts dates', () => {
    expect(serialToOdsDate(46290)).toBe('2026-09-25');
    expect(serialToOdsDate(46290.5625)).toBe('2026-09-25T13:30:00');
    expect(odsDateToSerial('2026-09-25')).toBe(46290);
    expect(odsDateToSerial('2026-09-25T13:30:00')).toBe(46290.5625);
    // times of day sit on LibreOffice's day 0, 1899-12-30
    expect(odsDateToSerial('1899-12-30')).toBe(0);
    expect(odsDateToSerial('1899-12-30T12:00:00')).toBe(0.5);
    expect(serialToOdsDate(0.5)).toBe('1899-12-30T12:00:00');
    // early dates keep their calendar date (Excel counts a 29 February 1900)
    expect(serialToOdsDate(1)).toBe('1900-01-01');
    expect(odsDateToSerial('1900-01-01')).toBe(1);
    expect(serialToOdsDate(61)).toBe('1900-03-01');
    expect(odsDateToSerial('1900-03-01')).toBe(61);
    expect(serialToOdsDate(60)).toBeUndefined();
  });
  it('converts durations', () => {
    expect(serialToOdsTime(0.5625)).toBe('PT13H30M00S');
    expect(serialToOdsTime(1.5104166)).toBe('PT36H14M59.99424S');
    expect(odsTimeToSerial('PT13H30M00S')).toBe(0.5625);
    expect(odsTimeToSerial('PT36H14M59.99424S')).toBe(1.5104166);
    expect(odsTimeToSerial('-PT1H')).toBeCloseTo(-1 / 24, 12);
    expect(odsTimeToSerial('P1DT12H')).toBe(1.5);
  });
});

function put(sheet: Sheet, a1: string, cell: Cell): void {
  const p = parseAddr(a1)!;
  sheet.put(cellKey(p.r, p.c), cell);
}

function sampleBook(): Workbook {
  const wb = new Workbook();
  const s = wb.addSheet('Data');
  const o = wb.addSheet('Other sheet');
  const head = wb.styleId({ bold: true, fill: '#1f6feb', color: '#ffffff', hAlign: 'center', bb: { style: 'medium', color: '#0a3d91' } });
  ['Name', 'Amount', 'Date', 'Share', 'Flag'].forEach((t, i) => put(s, `${String.fromCharCode(65 + i)}1`, { v: t, s: head }));
  const cur = wb.styleId({ numFmt: '$#,##0.00;[Red]-$#,##0.00' });
  const date = wb.styleId({ numFmt: 'd-mmm-yyyy', hAlign: 'center' });
  const pct = wb.styleId({ numFmt: '0.0%' });
  ['Apples', 'Pears', 'Plums'].forEach((n, i) => {
    put(s, `A${i + 2}`, { v: n });
    put(s, `B${i + 2}`, { v: [120.5, -45.25, 88][i], s: cur });
    put(s, `C${i + 2}`, { v: 46290 + i, s: date });
    put(s, `D${i + 2}`, { v: [0.125, 0.5, 0.05][i], s: pct });
    put(s, `E${i + 2}`, { v: i % 2 === 0 });
  });
  put(s, 'F2', { f: 'SUM(B2:B4)' });
  put(s, 'F3', { f: 'IF(B3>100,"big","small")' });
  put(s, 'F4', { f: "SUM('Other sheet'!A1:A2)+Rate" });
  put(s, 'F5', { f: '1/0' });
  put(s, 'A2', { v: 'Apples', note: 'A note\nover two lines' });
  put(s, 'A3', { v: 'Pears', link: 'https://example.com/pears' });
  put(s, 'A6', { v: '  two  spaces\tand a tab', s: wb.styleId({ wrap: true, vAlign: 'top', font: 'Georgia', size: 14, italic: true, underline: 'double', strike: true, rotation: -30, indent: 2, hAlign: 'left', locked: false }) });
  put(s, 'A8', { v: 'Merged', s: wb.styleId({ hAlign: 'center', bt: { style: 'thin' }, bb: { style: 'double', color: '#ff0000' }, bl: { style: 'dashed' }, br: { style: 'thick' } }) });
  s.merges.push({ r1: 7, c1: 0, r2: 8, c2: 2 });
  s.cols.set(0, { w: 140 });
  s.cols.set(6, { hidden: true });
  s.rows.set(7, { h: 40, custom: true });
  s.rows.set(20, { hidden: true });
  s.view.freezeRows = 1;
  s.view.freezeCols = 1;
  s.view.zoom = 1.25;
  s.cf.push({ id: 'a', ranges: [{ r1: 1, c1: 1, r2: 3, c2: 1 }], type: 'cell', op: 'greater', values: ['100'], style: { fill: '#ffc7ce', color: '#9c0006' } });
  s.cf.push({ id: 'b', ranges: [{ r1: 1, c1: 3, r2: 3, c2: 3 }], type: 'dataBar', colors: ['#638ec6'] });
  s.cf.push({ id: 'c', ranges: [{ r1: 1, c1: 2, r2: 3, c2: 2 }], type: 'colorScale', colors: ['#f8696b', '#ffeb84', '#63be7b'] });
  s.cf.push({ id: 'd', ranges: [{ r1: 1, c1: 1, r2: 3, c2: 1 }], type: 'iconSet', iconSet: '3TrafficLights', showValue: false });
  s.cf.push({ id: 'e', ranges: [{ r1: 1, c1: 0, r2: 3, c2: 0 }], type: 'formula', values: ['=LEN(A2)>4'], style: { bold: true } });
  s.cf.push({ id: 'f', ranges: [{ r1: 1, c1: 0, r2: 3, c2: 0 }], type: 'text', text: 'ea', style: { italic: true } });
  s.cf.push({ id: 'g', ranges: [{ r1: 1, c1: 1, r2: 3, c2: 1 }], type: 'cell', op: 'between', values: ['0', '=$B$2'], style: { underline: 'single' } });
  s.cf.push({ id: 'h', ranges: [{ r1: 1, c1: 2, r2: 3, c2: 2 }], type: 'date', datePeriod: 'last7Days', style: { color: '#00aa00' } });
  s.cf.push({ id: 'i', ranges: [{ r1: 1, c1: 1, r2: 3, c2: 1 }], type: 'top', rank: 2, bottom: true, style: { fill: '#eeeeee' } });
  s.cf.push({ id: 'j', ranges: [{ r1: 1, c1: 0, r2: 9, c2: 0 }], type: 'blank', style: { fill: '#dddddd' } });
  s.validations.push({ id: 'v1', ranges: [{ r1: 1, c1: 4, r2: 3, c2: 4 }], type: 'list', values: ['Yes', 'No', 'Maybe'], showDropdown: true, allowBlank: true, error: 'Pick one', errorTitle: 'Oops', errorStyle: 'warning' });
  s.validations.push({ id: 'v2', ranges: [{ r1: 13, c1: 1, r2: 13, c2: 1 }], type: 'whole', op: 'between', values: ['1', '10'], prompt: 'Whole number 1-10', promptTitle: 'Tip' });
  s.validations.push({ id: 'v3', ranges: [{ r1: 13, c1: 2, r2: 13, c2: 2 }], type: 'list', values: ['=$A$2:$A$4'] });
  s.validations.push({ id: 'v4', ranges: [{ r1: 0, c1: 10, r2: 1_048_575, c2: 10 }], type: 'textLength', op: 'lessEqual', values: ['5'] });
  s.filter = { range: { r1: 0, c1: 0, r2: 3, c2: 5 }, columns: { 0: { values: ['Apples', 'Plums'] } } };
  s.filterHidden.add(2);
  s.print = { orientation: 'landscape', paper: 'a4', margins: 'narrow', fit: 'page', scale: 100, gridlines: true, headings: false, header: '&A', footer: 'Page &P of &N', repeatRows: 1, area: { r1: 0, c1: 0, r2: 9, c2: 5 } };
  wb.names.push({ name: 'Rate', ref: 'Data!$B$4' });
  wb.names.push({ name: 'Double', ref: 'Data!$B$2*2' });
  wb.names.push({ name: 'Local', ref: "'Other sheet'!$A$1:$A$2", sheet: o.id });
  o.tabColor = '#e5484d';
  o.hidden = true;
  o.protection = { enabled: true };
  o.view.showGrid = false;
  put(o, 'A1', { v: 1 });
  put(o, 'A2', { v: 2 });
  put(o, 'B1', { v: 'Line one\nLine two', s: wb.styleId({ wrap: true }) });
  s.charts.push({ id: 'c1', type: 'column', source: 'Data!$A$1:$B$4', title: 'Amounts', legend: 'bottom', palette: 'office', anchor: { r: 15, c: 1, dx: 4, dy: 3 }, w: 420, h: 260 });
  s.charts.push({ id: 'c2', type: 'line', source: "'Other sheet'!$A$1:$A$2", legend: 'none', smooth: true, xTitle: 'Step', yTitle: 'Value', anchor: { r: 15, c: 8, dx: 0, dy: 0 }, w: 360, h: 240 });
  s.charts.push({ id: 'c3', type: 'pie', source: 'Data!$A$1:$B$4', legend: 'right', dataLabels: true, palette: 'vivid', anchor: { r: 30, c: 1, dx: 0, dy: 0 }, w: 320, h: 240 });
  s.charts.push({ id: 'c4', type: 'bar', source: 'Data!$A$1:$D$4', stacked: 'percent', gridlines: false, anchor: { r: 30, c: 8, dx: 0, dy: 0 }, w: 360, h: 240 });
  s.images.push({ id: 'i1', src: PIXEL, anchor: { r: 12, c: 3, dx: 6, dy: 2 }, w: 80, h: 60, alt: 'A dot' });
  wb.props = { title: 'ODS test', author: 'Affice' };
  wb.activeSheet = 0;
  return wb;
}

const PIXEL = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';

describe('OpenDocument spreadsheet files', () => {
  it('writes a valid package', async () => {
    const doc = new SheetDoc(sampleBook());
    const bytes = exportOds(doc);
    // the mimetype comes first and uncompressed, as the format requires
    expect(new TextDecoder().decode(bytes.subarray(30, 38))).toBe('mimetype');
    expect(new TextDecoder().decode(bytes.subarray(38, 84))).toBe('application/vnd.oasis.opendocument.spreadsheet');
  });

  it('round-trips a workbook', async () => {
    const original = sampleBook();
    const doc = new SheetDoc(original);
    const { wb } = await importOds(exportOds(doc));
    const [s, o] = wb.sheets;
    expect(wb.sheets.map((x) => x.name)).toEqual(['Data', 'Other sheet']);
    const cell = (sheet: Sheet, a1: string) => {
      const p = parseAddr(a1)!;
      const c = sheet.get(p.r, p.c);
      return { ...c, style: wb.style(c?.s) };
    };
    // values and formats
    expect(cell(s, 'A1')).toMatchObject({ v: 'Name', style: { bold: true, fill: '#1f6feb', color: '#ffffff', hAlign: 'center', bb: { style: 'medium', color: '#0a3d91' } } });
    expect(cell(s, 'B3')).toMatchObject({ v: -45.25, style: { numFmt: '$#,##0.00;[Red]-$#,##0.00' } });
    expect(cell(s, 'C2')).toMatchObject({ v: 46290, style: { numFmt: 'd-mmm-yyyy', hAlign: 'center' } });
    expect(cell(s, 'D2')).toMatchObject({ v: 0.125, style: { numFmt: '0.0%' } });
    expect(cell(s, 'E2').v).toBe(true);
    expect(cell(s, 'E3').v).toBe(false);
    // formulas with cached results
    expect(cell(s, 'F2')).toMatchObject({ f: 'SUM(B2:B4)', v: 163.25 });
    expect(cell(s, 'F3')).toMatchObject({ f: 'IF(B3>100,"big","small")', v: 'small' });
    expect(cell(s, 'F4')).toMatchObject({ f: "SUM('Other sheet'!A1:A2)+Rate", v: 91 });
    expect(cell(s, 'F5').f).toBe('1/0');
    expect((cell(s, 'F5').v as { error: string }).error).toBe('#DIV/0!');
    // text, notes, links
    expect(cell(s, 'A2')).toMatchObject({ v: 'Apples', note: 'A note\nover two lines' });
    expect(cell(s, 'A3')).toMatchObject({ v: 'Pears', link: 'https://example.com/pears' });
    expect(cell(s, 'A6')).toMatchObject({
      v: '  two  spaces\tand a tab',
      style: { wrap: true, vAlign: 'top', font: 'Georgia', size: 14, italic: true, underline: 'double', strike: true, rotation: -30, indent: 2, hAlign: 'left', locked: false },
    });
    expect(cell(s, 'A8').style).toMatchObject({ bt: { style: 'thin' }, bb: { style: 'double', color: '#ff0000' }, bl: { style: 'dashed' }, br: { style: 'thick' } });
    expect(s.merges).toEqual([{ r1: 7, c1: 0, r2: 8, c2: 2 }]);
    // layout
    expect(s.cols.get(0)?.w).toBe(140);
    expect(s.cols.get(6)?.hidden).toBe(true);
    expect(s.rows.get(7)).toMatchObject({ h: 40, custom: true });
    expect(s.rows.get(20)?.hidden).toBe(true);
    expect(s.defaultColWidth).toBe(original.sheets[0].defaultColWidth);
    expect(s.defaultRowHeight).toBe(original.sheets[0].defaultRowHeight);
    expect(s.view).toMatchObject({ freezeRows: 1, freezeCols: 1, zoom: 1.25 });
    // rules
    const strip = (list: { id: string }[]) => list.map(({ id: _id, ...rest }) => rest);
    expect(strip(s.cf)).toEqual(strip(original.sheets[0].cf));
    expect(s.validations.map((v) => ({ type: v.type, op: v.op, values: v.values, ranges: v.ranges }))).toEqual(original.sheets[0].validations.map((v) => ({ type: v.type, op: v.op, values: v.values, ranges: v.ranges })));
    expect(s.validations[0]).toMatchObject({ error: 'Pick one', errorTitle: 'Oops', errorStyle: 'warning' });
    expect(s.validations[1]).toMatchObject({ prompt: 'Whole number 1-10', promptTitle: 'Tip' });
    expect(s.filter).toEqual(original.sheets[0].filter);
    expect([...s.filterHidden]).toEqual([2]);
    expect(s.print).toEqual(original.sheets[0].print);
    // names and other sheet
    expect(wb.names).toEqual([
      { name: 'Local', ref: "'Other sheet'!$A$1:$A$2", sheet: o.id },
      { name: 'Rate', ref: 'Data!$B$4', sheet: undefined },
      { name: 'Double', ref: 'Data!$B$2*2', sheet: undefined },
    ]);
    // charts and pictures
    const chartKeys = (c: object) => {
      const { id: _id, ...rest } = c as { id: string };
      return rest;
    };
    expect(s.charts.map(chartKeys)).toEqual([
      { type: 'column', source: 'Data!$A$1:$B$4', seriesInRows: false, title: 'Amounts', legend: 'bottom', palette: 'office', anchor: { r: 15, c: 1, dx: 4, dy: 3 }, w: 420, h: 260 },
      { type: 'line', source: "'Other sheet'!$A$1:$A$2", seriesInRows: false, legend: 'none', smooth: true, xTitle: 'Step', yTitle: 'Value', palette: 'affice', anchor: { r: 15, c: 8, dx: 0, dy: 0 }, w: 360, h: 240 },
      { type: 'pie', source: 'Data!$A$1:$B$4', seriesInRows: false, legend: 'right', dataLabels: true, palette: 'vivid', anchor: { r: 30, c: 1, dx: 0, dy: 0 }, w: 320, h: 240 },
      { type: 'bar', source: 'Data!$A$1:$D$4', seriesInRows: false, stacked: 'percent', gridlines: false, legend: 'bottom', palette: 'affice', anchor: { r: 30, c: 8, dx: 0, dy: 0 }, w: 360, h: 240 },
    ]);
    expect(s.images.map(chartKeys)).toEqual([{ src: PIXEL, anchor: { r: 12, c: 3, dx: 6, dy: 2 }, w: 80, h: 60, alt: 'A dot' }]);
    expect(o).toMatchObject({ tabColor: '#e5484d', hidden: true, protection: { enabled: true } });
    expect(o.view.showGrid).toBe(false);
    expect(cell(o, 'B1')).toMatchObject({ v: 'Line one\nLine two', style: { wrap: true } });
    expect(wb.props).toMatchObject({ title: 'ODS test', author: 'Affice' });
  });

  it('streams rows the same way the DOM reads them', async () => {
    const bytes = exportOds(new SheetDoc(sampleBook()));
    // the fast path applies to our files…
    expect(splitContent(strFromU8(unzipSync(bytes)['content.xml']))).not.toBeNull();
    // …and gives the same workbook as reading everything through the DOM
    const fast = await importOds(bytes);
    const dom = await importOds(bytes, { fast: false });
    const json = (wb: Workbook) => JSON.stringify({ ...workbookToJSON(wb), sheets: workbookToJSON(wb).sheets.map((sh) => ({ ...sh, cf: sh.cf.map((c) => ({ ...c, id: '' })), validations: sh.validations.map((v) => ({ ...v, id: '' })), charts: sh.charts.map((c) => ({ ...c, id: '' })), images: sh.images.map((i) => ({ ...i, id: '' })) })) });
    expect(json(fast.wb)).toBe(json(dom.wb));
  });

  it('reads LibreOffice-style repeats, inherited styles and legacy conditions', async () => {
    const ns =
      'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0"';
    const styles = `<?xml version="1.0" encoding="UTF-8"?><office:document-styles ${ns}><office:styles>
      <style:default-style style:family="table-cell"><style:text-properties style:font-name="Liberation Sans" fo:font-size="10pt"/></style:default-style>
      <style:style style:name="Default" style:family="table-cell"/>
      <style:style style:name="Good" style:family="table-cell" style:parent-style-name="Default"><style:table-cell-properties fo:background-color="#ccffcc"/><style:text-properties fo:color="#006600"/></style:style>
      <style:style style:name="Bad" style:family="table-cell" style:parent-style-name="Default"><style:text-properties fo:color="#cc0000"/></style:style>
    </office:styles></office:document-styles>`;
    const content = `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${ns}><office:automatic-styles>
      <style:style style:name="co1" style:family="table-column"><style:table-column-properties style:column-width="0.889in"/></style:style>
      <style:style style:name="co2" style:family="table-column"><style:table-column-properties style:column-width="2in"/></style:style>
      <style:style style:name="ro1" style:family="table-row"><style:table-row-properties style:row-height="0.178in" style:use-optimal-row-height="true"/></style:style>
      <number:number-style style:name="N2"><number:number number:decimal-places="2" number:min-integer-digits="1"/></number:number-style>
      <style:style style:name="ce1" style:family="table-cell" style:parent-style-name="Good" style:data-style-name="N2"><style:text-properties fo:font-weight="bold"/></style:style>
      <style:style style:name="ce2" style:family="table-cell" style:parent-style-name="Default"><style:map style:condition="cell-content()&gt;5" style:apply-style-name="Bad" style:base-cell-address="Sheet1.C1"/></style:style>
    </office:automatic-styles><office:body><office:spreadsheet><table:table table:name="Sheet1">
      <table:table-column table:style-name="co2"/>
      <table:table-column table:style-name="co1" table:number-columns-repeated="1023" table:default-cell-style-name="Default"/>
      <table:table-row table:style-name="ro1" table:number-rows-repeated="2">
        <table:table-cell table:style-name="ce1" office:value-type="float" office:value="1.5" table:number-columns-repeated="2"><text:p>1.50</text:p></table:table-cell>
        <table:table-cell table:style-name="ce2" office:value-type="float" office:value="7"><text:p>7</text:p></table:table-cell>
      </table:table-row>
      <table:table-row table:style-name="ro1">
        <table:table-cell office:value-type="string"><text:p>a<text:s text:c="2"/>b<text:tab/>c</text:p><text:p>second</text:p></table:table-cell>
        <table:table-cell table:formula="of:=TRUE()" office:value-type="boolean" office:boolean-value="true"/>
        <table:table-cell table:formula="of:=[.A1:.B1]*2" table:number-matrix-columns-spanned="2" table:number-matrix-rows-spanned="1" office:value-type="float" office:value="3"/>
        <table:table-cell office:value-type="float" office:value="3"/>
      </table:table-row>
      <table:table-row table:style-name="ro1" table:number-rows-repeated="1048573"><table:table-cell table:number-columns-repeated="1024"/></table:table-row>
    </table:table></office:spreadsheet></office:body></office:document-content>`;
    const zip = zipSync({ mimetype: strToU8('application/vnd.oasis.opendocument.spreadsheet'), 'content.xml': strToU8(content), 'styles.xml': strToU8(styles) });
    const { wb } = await importOds(zip);
    expect(JSON.stringify(workbookToJSON((await importOds(zip, { fast: false })).wb))).toBe(JSON.stringify(workbookToJSON(wb)));
    const s = wb.sheets[0];
    const at = (a1: string) => {
      const p = parseAddr(a1)!;
      return s.get(p.r, p.c);
    };
    // repeated rows and cells hold copies
    for (const a of ['A1', 'B1', 'A2', 'B2']) expect(at(a)?.v).toBe(1.5);
    expect(wb.style(at('B2')?.s)).toEqual({ font: 'Liberation Sans', size: 10, bold: true, fill: '#ccffcc', color: '#006600', numFmt: '0.00' });
    expect(at('A3')?.v).toBe('a  b\tc\nsecond');
    expect(at('B3')).toEqual({ v: true, s: wb.styleId({ font: 'Liberation Sans', size: 10 }) });
    // a matrix formula keeps its formula; the values it spilled are dropped
    expect(at('C3')?.f).toBe('A1:B1*2');
    expect(at('D3')).toBeUndefined();
    // the most common sizes become the defaults
    expect(s.defaultColWidth).toBe(85);
    expect(s.defaultRowHeight).toBe(17);
    expect(s.cols.get(0)?.w).toBe(192);
    // conditions stored on cell styles become conditional formats
    expect(s.cf).toHaveLength(1);
    expect(s.cf[0]).toMatchObject({ type: 'cell', op: 'greater', values: ['5'], ranges: [{ r1: 0, c1: 2, r2: 1, c2: 2 }], style: { color: '#cc0000' } });
  });
});
