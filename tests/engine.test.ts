import { describe, expect, it } from 'vitest';
import { parseFormula, tokenize } from '../src/sheets/engine/parser';
import { CalcEngine } from '../src/sheets/engine/calc';
import { Workbook } from '../src/sheets/model/workbook';
import { cellKey, parseAddr } from '../src/sheets/model/address';
import { isErr, asErr, type Scalar } from '../src/sheets/engine/values';
import { adjustFormulaStructural, translateFormula, renameSheetInFormula, adjustFormulaMove, cycleAbsolute } from '../src/sheets/engine/refshift';
import { parseInput } from '../src/sheets/format/numfmt';

function book(data: Record<string, string | number | boolean>, sheets = ['Sheet1']) {
  const wb = new Workbook();
  for (const n of sheets) wb.addSheet(n);
  for (const [addr, raw] of Object.entries(data)) {
    const [sheetName, a] = addr.includes('!') ? addr.split('!') : [sheets[0], addr];
    const sheet = wb.sheetByName(sheetName)!;
    const p = parseAddr(a)!;
    if (typeof raw === 'string' && raw.startsWith('=')) sheet.put(cellKey(p.r, p.c), { f: raw.slice(1) });
    else if (typeof raw === 'string') {
      const v = parseInput(raw).value;
      sheet.put(cellKey(p.r, p.c), { v });
    } else sheet.put(cellKey(p.r, p.c), { v: raw });
  }
  const eng = new CalcEngine(wb);
  eng.clock = () => new Date(2024, 0, 15, 12, 0, 0);
  eng.random = () => 0.5;
  eng.rebuild();
  const get = (addr: string): Scalar => {
    const [sheetName, a] = addr.includes('!') ? addr.split('!') : [sheets[0], addr];
    const p = parseAddr(a)!;
    const v = eng.displayValue(wb.sheetByName(sheetName)!.id, p.r, p.c);
    return v === undefined ? null : v;
  };
  const set = (addr: string, raw: string | number | null) => {
    const [sheetName, a] = addr.includes('!') ? addr.split('!') : [sheets[0], addr];
    const sheet = wb.sheetByName(sheetName)!;
    const p = parseAddr(a)!;
    const k = cellKey(p.r, p.c);
    if (raw === null) sheet.put(k, undefined);
    else if (typeof raw === 'string' && raw.startsWith('=')) sheet.put(k, { f: raw.slice(1) });
    else sheet.put(k, { v: typeof raw === 'string' ? parseInput(raw).value : raw });
    eng.cellsChanged([{ sheet: sheet.id, key: k }]);
  };
  return { wb, eng, get, set };
}

/** Evaluates a single formula on a sheet with some data. */
function calc(formula: string, data: Record<string, string | number | boolean> = {}): Scalar {
  const b = book({ ...data, Z100: formula.startsWith('=') ? formula : '=' + formula });
  return b.get('Z100');
}

const err = (v: Scalar) => (isErr(v) ? asErr(v).error : v);

describe('parser', () => {
  it('handles precedence', () => {
    expect(calc('=-2^2')).toBe(4);
    expect(calc('=2+3*4')).toBe(14);
    expect(calc('=2^3^2')).toBe(64);
    expect(calc('=10-2-3')).toBe(5);
    expect(calc('=50%*4')).toBe(2);
    expect(calc('="a"&1+2')).toBe('a3');
    expect(calc('=1+2=3')).toBe(true);
  });
  it('parses references', () => {
    expect(parseFormula("'My Sheet'!A1")).toMatchObject({ t: 'ref', sheet: 'My Sheet', r: 0, c: 0 });
    expect(parseFormula('A:C')).toMatchObject({ t: 'range', kind: 'cols', c1: 0, c2: 2 });
    expect(parseFormula('1:3')).toMatchObject({ t: 'range', kind: 'rows', r1: 0, r2: 2 });
    expect(parseFormula('$B$2:C3')).toMatchObject({ t: 'range', r1: 1, c1: 1, r2: 2, c2: 2, ra1: true, ca1: true, ra2: false });
    expect(parseFormula('A1#')).toMatchObject({ t: 'spill' });
    expect(parseFormula('{1,2;3,4}')).toMatchObject({ t: 'array', rows: [[{ v: 1 }, { v: 2 }], [{ v: 3 }, { v: 4 }]] });
    expect(parseFormula('IF(A1,,1)')).toMatchObject({ t: 'func', name: 'IF', args: [{ t: 'ref' }, { t: 'missing' }, { t: 'num' }] });
    expect(parseFormula('_xlfn._xlws.SORT(A1:A3)')).toMatchObject({ t: 'func', name: 'SORT' });
    expect(parseFormula('LOG10(100)')).toMatchObject({ t: 'func', name: 'LOG10' });
  });
  it('rejects malformed formulas', () => {
    expect(() => parseFormula('1+')).toThrow();
    expect(() => parseFormula('SUM(1,2')).toThrow();
    expect(() => parseFormula('"abc')).toThrow();
  });
  it('tokenizes sheet refs', () => {
    const t = tokenize("Sheet2!A1+'Q1 Data'!B2:C3");
    expect(t.filter((x) => x.type === 'ref').map((x) => x.sheet)).toEqual(['Sheet2', 'Q1 Data', undefined]);
  });
});

describe('math & stats', () => {
  const data = { A1: 1, A2: 2, A3: 3, A4: 'x', A5: true, B1: 10, B2: 20, B3: 30 };
  it('aggregates', () => {
    expect(calc('=SUM(A1:A5)', data)).toBe(6);
    expect(calc('=SUM(A1:A3,10,"5",TRUE)', data)).toBe(22);
    expect(calc('=AVERAGE(A1:A5)', data)).toBe(2);
    expect(calc('=COUNT(A1:A5)', data)).toBe(3);
    expect(calc('=COUNTA(A1:A5)', data)).toBe(5);
    expect(calc('=COUNTBLANK(A1:A6)', data)).toBe(1);
    expect(calc('=MAX(A1:A5)', data)).toBe(3);
    expect(calc('=MIN(B1:B3)', data)).toBe(10);
    expect(calc('=PRODUCT(A1:A3)', data)).toBe(6);
    expect(calc('=SUMPRODUCT(A1:A3,B1:B3)', data)).toBe(140);
    expect(calc('=SUMPRODUCT((A1:A3>1)*B1:B3)', data)).toBe(50);
    expect(calc('=MEDIAN(1,2,3,4)')).toBe(2.5);
    expect(calc('=ROUND(STDEV(2,4,4,4,5,5,7,9),6)')).toBe(2.13809);
    expect(calc('=STDEV.P(2,4,4,4,5,5,7,9)')).toBe(2);
    expect(calc('=LARGE(B1:B3,2)', data)).toBe(20);
    expect(calc('=RANK(20,B1:B3)', data)).toBe(2);
    expect(calc('=PERCENTILE(B1:B3,0.5)', data)).toBe(20);
    expect(calc('=QUARTILE({1,2,3,4,5,6,7,8},1)')).toBe(2.75);
    expect(calc('=MODE(1,2,2,3)')).toBe(2);
    expect(calc('=CORREL(A1:A3,B1:B3)', data)).toBeCloseTo(1);
    expect(calc('=SLOPE(B1:B3,A1:A3)', data)).toBeCloseTo(10);
    expect(calc('=FORECAST(4,B1:B3,A1:A3)', data)).toBeCloseTo(40);
  });
  it('conditional aggregates', () => {
    const d = { A1: 'apple', A2: 'banana', A3: 'apple', A4: 'cherry', B1: 5, B2: 7, B3: 3, B4: 1 };
    expect(calc('=SUMIF(A1:A4,"apple",B1:B4)', d)).toBe(8);
    expect(calc('=SUMIF(B1:B4,">4")', d)).toBe(12);
    expect(calc('=COUNTIF(A1:A4,"a*")', d)).toBe(2);
    expect(calc('=COUNTIF(A1:A4,"<>apple")', d)).toBe(2);
    expect(calc('=COUNTIFS(A1:A4,"apple",B1:B4,">3")', d)).toBe(1);
    expect(calc('=SUMIFS(B1:B4,A1:A4,"?????",B1:B4,"<>1")', d)).toBe(8);
    expect(calc('=AVERAGEIF(A1:A4,"apple",B1:B4)', d)).toBe(4);
    expect(calc('=MAXIFS(B1:B4,A1:A4,"apple")', d)).toBe(5);
    expect(calc('=COUNTIF(A1:A10,"")', d)).toBe(6);
  });
  it('rounds like Excel', () => {
    expect(calc('=ROUND(2.5,0)')).toBe(3);
    expect(calc('=ROUND(-2.5,0)')).toBe(-3);
    expect(calc('=ROUND(1.005,2)')).toBe(1.01);
    expect(calc('=ROUND(1234.5678,-2)')).toBe(1200);
    expect(calc('=ROUNDUP(3.2,0)')).toBe(4);
    expect(calc('=ROUNDDOWN(-3.7,0)')).toBe(-3);
    expect(calc('=INT(-4.3)')).toBe(-5);
    expect(calc('=TRUNC(-4.3)')).toBe(-4);
    expect(calc('=MOD(-3,2)')).toBe(1);
    expect(calc('=MOD(3,-2)')).toBe(-1);
    expect(calc('=CEILING(2.5,1)')).toBe(3);
    expect(calc('=FLOOR(-2.5,-2)')).toBe(-2);
    expect(calc('=MROUND(10,3)')).toBe(9);
    expect(calc('=0.1+0.2-0.3')).toBe(0);
    expect(calc('=0.1+0.2=0.3')).toBe(true);
  });
  it('errors', () => {
    expect(err(calc('=1/0'))).toBe('#DIV/0!');
    expect(err(calc('=SQRT(-1)'))).toBe('#NUM!');
    expect(err(calc('="a"+1'))).toBe('#VALUE!');
    expect(err(calc('=FOO(1)'))).toBe('#NAME?');
    expect(err(calc('=SUM(A1:A2)', { A1: 1, A2: '=1/0' }))).toBe('#DIV/0!');
  });
  it('matrices', () => {
    expect(calc('=MDETERM({1,2;3,4})')).toBe(-2);
    expect(calc('=INDEX(MINVERSE({4,7;2,6}),1,1)')).toBeCloseTo(0.6);
    expect(calc('=INDEX(MMULT({1,2;3,4},{5;6}),2,1)')).toBe(39);
    expect(calc('=SUM(SEQUENCE(10))')).toBe(55);
  });
});

describe('text', () => {
  it('string functions', () => {
    expect(calc('=LEFT("Hello",2)&RIGHT("World",3)&MID("abcdef",2,3)')).toBe('Herldbcd');
    expect(calc('=LEN(1/3)')).toBe(17);
    expect(calc('=UPPER("abc")&LOWER("DEF")&PROPER("hello wORLD")')).toBe('ABCdefHello World');
    expect(calc('=TRIM("  a   b  ")')).toBe('a b');
    expect(calc('=SUBSTITUTE("a-b-c","-","+")')).toBe('a+b+c');
    expect(calc('=SUBSTITUTE("a-b-c","-","+",2)')).toBe('a-b+c');
    expect(calc('=FIND("b","abcb")')).toBe(2);
    expect(calc('=SEARCH("B*D","abcd")')).toBe(2);
    expect(calc('=TEXT(1234.567,"#,##0.00")')).toBe('1,234.57');
    expect(calc('=TEXT(45000,"yyyy-mm-dd")')).toBe('2023-03-15');
    expect(calc('=VALUE("1,234.5")')).toBe(1234.5);
    expect(calc('=TEXTJOIN(", ",TRUE,"a","","b")')).toBe('a, b');
    expect(calc('=CONCAT(A1:B1)', { A1: 'x', B1: 2 })).toBe('x2');
    expect(calc('=REPT("ab",3)')).toBe('ababab');
    expect(calc('=FIXED(1234.567,1)')).toBe('1,234.6');
    expect(calc('=DOLLAR(-1234.5)')).toBe('($1,234.50)');
    expect(calc('=TEXTBEFORE("a.b.c",".",-1)')).toBe('a.b');
    expect(calc('=TEXTAFTER("a.b.c",".")')).toBe('b.c');
    expect(calc('=CHAR(65)&CODE("a")')).toBe('A97');
    expect(calc('=REGEXEXTRACT("Order 1234 done","[0-9]+")')).toBe('1234');
    expect(calc('=REGEXREPLACE("a1b2","[0-9]","#")')).toBe('a#b#');
    expect(calc('=EXACT("a","A")')).toBe(false);
  });
});

describe('logical & lookup', () => {
  const table = { A1: 'id', B1: 'name', A2: 1, B2: 'Ann', A3: 2, B3: 'Bob', A4: 3, B4: 'Cid' };
  it('logical', () => {
    expect(calc('=IF(1>2,"y","n")')).toBe('n');
    expect(calc('=IF(TRUE,,1)')).toBe(0);
    expect(calc('=IF(FALSE,1)')).toBe(false);
    expect(calc('=IFERROR(1/0,"bad")')).toBe('bad');
    expect(calc('=IFNA(NA(),5)')).toBe(5);
    expect(calc('=IFS(1>2,"a",2>1,"b")')).toBe('b');
    expect(calc('=SWITCH(2,1,"one",2,"two","other")')).toBe('two');
    expect(calc('=AND(TRUE,1,A1)', { A1: true })).toBe(true);
    expect(calc('=OR(FALSE,0)')).toBe(false);
    expect(calc('=XOR(TRUE,TRUE,TRUE)')).toBe(true);
    expect(calc('=CHOOSE(2,"a","b","c")')).toBe('b');
    expect(calc('=LET(x,5,y,x*2,x+y)')).toBe(15);
    expect(calc('=LAMBDA(a,b,a*b)(3,4)')).toBe(12);
    expect(calc('=SUM(MAP({1,2,3},LAMBDA(v,v*10)))')).toBe(60);
    expect(calc('=REDUCE(0,{1,2,3},LAMBDA(a,v,a+v))')).toBe(6);
    expect(calc('=INDEX(BYROW({1,2;3,4},LAMBDA(r,SUM(r))),2,1)')).toBe(7);
    expect(calc('=INDEX(BYROW({1,2;3,4},SUM),1,1)')).toBe(3);
  });
  it('lookup', () => {
    expect(calc('=VLOOKUP(2,A2:B4,2,FALSE)', table)).toBe('Bob');
    expect(calc('=VLOOKUP(2.5,A2:B4,2)', table)).toBe('Bob');
    expect(err(calc('=VLOOKUP(9,A2:B4,2,FALSE)', table))).toBe('#N/A');
    expect(calc('=HLOOKUP("name",A1:B4,3,FALSE)', table)).toBe('Bob');
    expect(calc('=MATCH("bob",B1:B4,0)', table)).toBe(3);
    expect(calc('=MATCH(2.5,A2:A4)', table)).toBe(2);
    expect(calc('=INDEX(B1:B4,MATCH(3,A1:A4,0))', table)).toBe('Cid');
    expect(calc('=XLOOKUP("Cid",B2:B4,A2:A4)', table)).toBe(3);
    expect(calc('=XLOOKUP("Zed",B2:B4,A2:A4,"none")', table)).toBe('none');
    expect(calc('=XLOOKUP(2.5,A2:A4,B2:B4,,1)', table)).toBe('Cid');
    expect(calc('=XMATCH("B*",B2:B4,2)', table)).toBe(2);
    expect(calc('=LOOKUP(2,A2:A4,B2:B4)', table)).toBe('Bob');
    expect(calc('=SUM(OFFSET(A2,0,0,3,1))', table)).toBe(6);
    expect(calc('=INDIRECT("B"&3)', table)).toBe('Bob');
    expect(calc('=ROWS(A1:B4)*COLUMNS(A1:B4)', table)).toBe(8);
    expect(calc('=ADDRESS(2,3)')).toBe('$C$2');
    expect(calc('=ROW()+COLUMN()')).toBe(100 + 26);
    expect(calc('=SUM(INDEX(A2:B4,0,1))', table)).toBe(6);
  });
});

describe('dates', () => {
  it('date math', () => {
    expect(calc('=DATE(2024,1,31)')).toBe(45322);
    expect(calc('=DATE(2024,14,1)')).toBe(calc('=DATE(2025,2,1)'));
    expect(calc('=YEAR(45322)&"-"&MONTH(45322)&"-"&DAY(45322)')).toBe('2024-1-31');
    expect(calc('=EDATE(DATE(2024,1,31),1)')).toBe(calc('=DATE(2024,2,29)'));
    expect(calc('=EOMONTH(DATE(2024,1,15),0)')).toBe(45322);
    expect(calc('=WEEKDAY(DATE(2024,1,15))')).toBe(2);
    expect(calc('=WEEKDAY(DATE(2024,1,15),2)')).toBe(1);
    expect(calc('=DATEDIF(DATE(2020,5,10),DATE(2024,1,15),"Y")')).toBe(3);
    expect(calc('=DATEDIF(DATE(2020,5,10),DATE(2024,1,15),"YM")')).toBe(8);
    expect(calc('=NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,31))')).toBe(23);
    expect(calc('=WORKDAY(DATE(2024,1,5),1)')).toBe(calc('=DATE(2024,1,8)'));
    expect(calc('=TODAY()')).toBe(45306);
    expect(calc('=HOUR(TIME(13,45,10))*100+MINUTE("13:45")')).toBe(1345);
    expect(calc('=ISOWEEKNUM(DATE(2021,1,3))')).toBe(53);
    expect(calc('=ROUND(YEARFRAC(DATE(2024,1,1),DATE(2024,7,1)),4)')).toBe(0.5);
    expect(calc('=DATEVALUE("2024-03-01")')).toBe(45352);
  });
});

describe('financial', () => {
  it('loans', () => {
    expect(calc('=ROUND(PMT(0.05/12,360,200000),2)')).toBe(-1073.64);
    expect(calc('=ROUND(FV(0.06/12,10,-200,-500,1),2)')).toBe(2581.4);
    expect(calc('=ROUND(PV(0.08/12,240,500),2)')).toBe(-59777.15);
    expect(calc('=ROUND(NPER(0.01,-100,1000),4)')).toBe(10.5886);
    expect(calc('=ROUND(RATE(48,-200,8000),6)')).toBe(0.007701);
    expect(calc('=ROUND(IPMT(0.1/12,1,36,8000),2)')).toBe(-66.67);
    expect(calc('=ROUND(NPV(0.1,-10000,3000,4200,6800),2)')).toBe(1188.44);
    expect(calc('=ROUND(IRR({-70000,12000,15000,18000,21000,26000}),4)')).toBe(0.0866);
    expect(calc('=SLN(30000,7500,10)')).toBe(2250);
    expect(calc('=ROUND(DDB(2400,300,10,1),2)')).toBe(480);
    expect(calc('=ROUND(DB(1000000,100000,6,1,7),2)')).toBe(186083.33);
  });
});

describe('dynamic arrays & spills', () => {
  it('spills into neighbours', () => {
    const b = book({ A1: 3, A2: 1, A3: 2, C1: '=SORT(A1:A3)', D1: '=C2*10', E1: '=SUM(C1#)' });
    expect(b.get('C1')).toBe(1);
    expect(b.get('C2')).toBe(2);
    expect(b.get('C3')).toBe(3);
    expect(b.get('D1')).toBe(20);
    expect(b.get('E1')).toBe(6);
    b.set('A2', 10);
    expect(b.get('C3')).toBe(10);
    expect(b.get('D1')).toBe(30);
    expect(b.get('E1')).toBe(15);
  });
  it('reports #SPILL! when blocked and recovers', () => {
    const b = book({ A1: '=SEQUENCE(3)', A3: 'x' });
    expect(err(b.get('A1'))).toBe('#SPILL!');
    b.set('A3', null);
    expect(b.get('A3')).toBe(3);
  });
  it('filter/unique/sortby', () => {
    const b = book({ A1: 'a', A2: 'b', A3: 'a', B1: 1, B2: 2, B3: 3, D1: '=UNIQUE(A1:A3)', E1: '=FILTER(B1:B3,A1:A3="a")', F1: '=SORTBY(A1:A3,B1:B3,-1)' });
    expect([b.get('D1'), b.get('D2'), b.get('D3')]).toEqual(['a', 'b', null]);
    expect([b.get('E1'), b.get('E2')]).toEqual([1, 3]);
    expect([b.get('F1'), b.get('F2'), b.get('F3')]).toEqual(['a', 'b', 'a']);
  });
  it('array arithmetic', () => {
    const b = book({ A1: 1, A2: 2, B1: '=A1:A2*{10,100}' });
    expect([b.get('B1'), b.get('C1'), b.get('B2'), b.get('C2')]).toEqual([10, 100, 20, 200]);
  });
});

describe('recalc', () => {
  it('propagates through chains and ranges', () => {
    const b = book({ A1: 1, A2: '=A1+1', A3: '=A2+1', A4: '=SUM(A1:A3)', B1: '=A4*2' });
    expect(b.get('B1')).toBe(12);
    b.set('A1', 10);
    expect(b.get('A3')).toBe(12);
    expect(b.get('B1')).toBe(66);
  });
  it('detects circular references', () => {
    const b = book({ A1: '=B1+1', B1: '=A1+1' });
    expect(err(b.get('A1'))).toBe('#CIRC!');
    b.set('B1', 5);
    expect(b.get('A1')).toBe(6);
  });
  it('handles long chains without stack overflow', () => {
    const data: Record<string, string | number> = { A1: 1 };
    for (let i = 2; i <= 5000; i++) data[`A${i}`] = `=A${i - 1}+1`;
    const b = book(data);
    expect(b.get('A5000')).toBe(5000);
    b.set('A1', 2);
    expect(b.get('A5000')).toBe(5001);
  });
  it('works across sheets and names', () => {
    const b = book({ 'Data!A1': 5, 'Main!A1': '=Data!A1*2', 'Main!A2': "='Data'!A1+Main!A1" }, ['Main', 'Data']);
    expect(b.get('Main!A1')).toBe(10);
    expect(b.get('Main!A2')).toBe(15);
    b.wb.names.push({ name: 'Rate', ref: 'Data!$A$1' });
    b.set('Main!A3', '=Rate*3');
    expect(b.get('Main!A3')).toBe(15);
    b.set('Data!A1', 1);
    expect(b.get('Main!A3')).toBe(3);
  });
  it('whole-column references', () => {
    const b = book({ A1: 1, A2: 2, B1: '=SUM(A:A)' });
    expect(b.get('B1')).toBe(3);
    b.set('A500', 10);
    expect(b.get('B1')).toBe(13);
  });
});

describe('reference adjustment', () => {
  it('translates relative references', () => {
    expect(translateFormula('A1+$B$2+B$3+$C4', 1, 1)).toBe('B2+$B$2+C$3+$C5');
    expect(translateFormula('SUM(A1:A3)', 2, 0)).toBe('SUM(A3:A5)');
    expect(translateFormula('A1', -1, 0)).toBe('#REF!');
    expect(translateFormula('SUM(A:A)', 5, 1)).toBe('SUM(B:B)');
    expect(translateFormula("'My Sheet'!A1*2", 0, 2)).toBe("'My Sheet'!C1*2");
  });
  it('shifts for inserted and deleted rows', () => {
    expect(adjustFormulaStructural('A1+A5', 'Sheet1', { sheet: 'Sheet1', axis: 'row', at: 2, count: 2 })).toBe('A1+A7');
    expect(adjustFormulaStructural('SUM(A1:A5)', 'Sheet1', { sheet: 'Sheet1', axis: 'row', at: 2, count: 1 })).toBe('SUM(A1:A6)');
    expect(adjustFormulaStructural('SUM(A1:A5)', 'Sheet1', { sheet: 'Sheet1', axis: 'row', at: 1, count: -2 })).toBe('SUM(A1:A3)');
    expect(adjustFormulaStructural('A3', 'Sheet1', { sheet: 'Sheet1', axis: 'row', at: 2, count: -1 })).toBe('#REF!');
    expect(adjustFormulaStructural('Other!A3+A3', 'Sheet1', { sheet: 'Other', axis: 'row', at: 0, count: 1 })).toBe('Other!A4+A3');
    expect(adjustFormulaStructural('SUM(B:D)', 'Sheet1', { sheet: 'Sheet1', axis: 'col', at: 2, count: -1 })).toBe('SUM(B:C)');
  });
  it('moves and renames', () => {
    expect(adjustFormulaMove('A1+B1', 'Sheet1', { srcSheet: 'Sheet1', src: { r1: 0, c1: 0, r2: 0, c2: 0 }, dstSheet: 'Sheet1', dr: 2, dc: 2 })).toBe('C3+B1');
    expect(renameSheetInFormula("Data!A1+'Data'!B2", 'Data', 'Q1 Data')).toBe("'Q1 Data'!A1+'Q1 Data'!B2");
    expect(cycleAbsolute('A1+B2', 1).text).toBe('$A$1+B2');
    expect(cycleAbsolute('$A$1', 1).text).toBe('A$1');
  });
});
