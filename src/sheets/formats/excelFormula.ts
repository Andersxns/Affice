import { parseFormula, tokenize, type Node } from '../engine/parser';

/** Functions newer than Excel 2007 are stored with a _xlfn. prefix in files. */
const XLFN = new Set(
  `ACOT ACOTH AGGREGATE ARABIC BASE BETA.DIST BETA.INV BINOM.DIST BINOM.DIST.RANGE BINOM.INV BITAND BITLSHIFT BITOR BITRSHIFT BITXOR CEILING.MATH CEILING.PRECISE
CHISQ.DIST CHISQ.DIST.RT CHISQ.INV CHISQ.INV.RT CHISQ.TEST COMBINA CONCAT CONFIDENCE.NORM CONFIDENCE.T COT COTH COVARIANCE.P COVARIANCE.S CSC CSCH DAYS DECIMAL
ERF.PRECISE ERFC.PRECISE EXPON.DIST F.DIST F.DIST.RT F.INV F.INV.RT F.TEST FLOOR.MATH FLOOR.PRECISE FORECAST.LINEAR FORMULATEXT GAMMA GAMMA.DIST GAMMA.INV
GAMMALN.PRECISE GAUSS HYPGEOM.DIST IFNA IFS IMCOSH IMCOT IMCSC IMCSCH IMSEC IMSECH IMSINH IMTAN ISFORMULA ISOWEEKNUM LOGNORM.DIST LOGNORM.INV MAXIFS MINIFS
MODE.MULT MODE.SNGL MUNIT NEGBINOM.DIST NORM.DIST NORM.INV NORM.S.DIST NORM.S.INV NUMBERVALUE PDURATION PERCENTILE.EXC PERCENTILE.INC PERCENTRANK.EXC
PERCENTRANK.INC PERMUTATIONA PHI POISSON.DIST QUARTILE.EXC QUARTILE.INC RANK.AVG RANK.EQ RRI SEC SECH SHEET SHEETS SKEW.P STDEV.P STDEV.S SWITCH T.DIST
T.DIST.2T T.DIST.RT T.INV T.INV.2T T.TEST TEXTJOIN UNICHAR UNICODE VAR.P VAR.S WEIBULL.DIST XOR Z.TEST XLOOKUP XMATCH LET LAMBDA SEQUENCE RANDARRAY UNIQUE
SORTBY TEXTBEFORE TEXTAFTER TEXTSPLIT VSTACK HSTACK TOCOL TOROW WRAPROWS WRAPCOLS TAKE DROP CHOOSEROWS CHOOSECOLS EXPAND MAP REDUCE SCAN MAKEARRAY BYROW
BYCOL ISOMITTED VALUETOTEXT ARRAYTOTEXT REGEXTEST REGEXEXTRACT REGEXREPLACE GROUPBY PIVOTBY ENCODEURL`.split(/\s+/),
);
const XLWS = new Set(['SORT', 'FILTER']);

function lambdaParams(node: Node, out: Set<string>): void {
  switch (node.t) {
    case 'func':
      if (node.name === 'LET') node.args.forEach((a, i) => i % 2 === 0 && i < node.args.length - 1 && a.t === 'name' && out.add(a.name));
      if (node.name === 'LAMBDA') node.args.slice(0, -1).forEach((a) => a.t === 'name' && out.add(a.name));
      node.args.forEach((a) => lambdaParams(a, out));
      break;
    case 'call':
      lambdaParams(node.callee, out);
      node.args.forEach((a) => lambdaParams(a, out));
      break;
    case 'unary':
    case 'postfix':
      lambdaParams(node.arg, out);
      break;
    case 'bin':
      lambdaParams(node.left, out);
      lambdaParams(node.right, out);
      break;
    case 'spill':
      lambdaParams(node.ref, out);
      break;
    default:
      break;
  }
}

/** Formula text as Excel stores it (future-function prefixes, _xlpm. parameters). */
export function toExcelFormula(f: string): string {
  const toks = tokenize(f);
  let params = new Set<string>();
  if (/\b(LET|LAMBDA)\s*\(/i.test(f)) {
    try {
      lambdaParams(parseFormula(f), params);
    } catch {
      params = new Set();
    }
  }
  let out = '';
  for (const t of toks) {
    if (t.type === 'func') {
      const up = t.text.toUpperCase();
      if (XLWS.has(up)) out += `_xlfn._xlws.${up}`;
      else if (XLFN.has(up)) out += `_xlfn.${up}`;
      else out += t.text;
    } else if (t.type === 'name' && !t.sheet && params.has(t.text.toUpperCase())) out += `_xlpm.${t.text}`;
    else out += t.text;
  }
  return out;
}

/** Strips Excel's storage prefixes (the parser tolerates them, but the UI shouldn't show them). */
export function fromExcelFormula(f: string): string {
  return f.replace(/_xlfn\._xlws\./gi, '').replace(/_xlfn\./gi, '').replace(/_xlpm\./gi, '').replace(/_xlws\./gi, '');
}
