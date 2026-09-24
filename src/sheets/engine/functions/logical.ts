import { ERR, isErr, asErr, isLambda, isMatrix, isRef, matrix, toBool, compare, type LambdaValue, type Matrix, type Scalar, type Value } from '../values';
import { at, broadcastDims, callLambda, deref, evaluate, isOmitted, makeLambda, numEquals, toGrid, toScalar, type EvalContext, type Scope } from '../evaluator';
import type { Node } from '../parser';
import { define } from './registry';
import { bool, eachArg, fn, grid, int, num } from './helpers';

const L = 'Logical' as const;

/* ------------------------------------------------------------ helpers */

function evalArg(node: Node | undefined, ctx: EvalContext, scope: Scope | undefined, omitted: Value): Value {
  if (!node) return omitted;
  if (node.t === 'missing') return 0;
  return evaluate(node, ctx, scope);
}

/** Evaluated value as scalar-or-matrix (references resolved). */
function sm(v: Value, ctx: EvalContext): Scalar | Matrix {
  const d = deref(v, ctx);
  if (isLambda(d)) return ERR.CALC;
  return d;
}

function truthy(v: Scalar): boolean | Scalar {
  if (isErr(v)) return asErr(v);
  if (v === null) return false;
  const b = toBool(v);
  return typeof b === 'boolean' ? b : ERR.VALUE;
}

/* ------------------------------------------------------------------ IF */

define({
  name: 'IF',
  category: L,
  min: 1,
  max: 3,
  syntax: 'logical_test, [value_if_true], [value_if_false]',
  desc: 'Checks whether a condition is met, and returns one value if TRUE, and another value if FALSE.',
  lazy: (nodes, ctx, scope) => {
    const cond = sm(evaluate(nodes[0], ctx, scope), ctx);
    if (isMatrix(cond)) {
      const a = sm(evalArg(nodes[1], ctx, scope, true), ctx);
      const b = sm(evalArg(nodes[2], ctx, scope, false), ctx);
      const [R, C] = broadcastDims([cond, a, b]);
      const out: Scalar[][] = [];
      for (let r = 0; r < R; r++) {
        const row: Scalar[] = [];
        for (let c = 0; c < C; c++) {
          const t = truthy(at(cond, r, c));
          if (typeof t !== 'boolean') row.push(t);
          else {
            const pick = t ? a : b;
            row.push(isMatrix(pick) ? at(pick, r, c) : pick);
          }
        }
        out.push(row);
      }
      return matrix(out);
    }
    const t = truthy(cond);
    if (typeof t !== 'boolean') return t;
    return t ? evalArg(nodes[1], ctx, scope, true) : evalArg(nodes[2], ctx, scope, false);
  },
});

define({
  name: 'IFS',
  category: L,
  min: 2,
  max: 254,
  syntax: 'logical_test1, value_if_true1, [logical_test2, value_if_true2], …',
  desc: 'Checks whether one or more conditions are met and returns the value of the first TRUE condition.',
  lazy: (nodes, ctx, scope) => {
    if (nodes.length % 2) return ERR.VALUE;
    const conds: (Scalar | Matrix)[] = [];
    for (let i = 0; i < nodes.length; i += 2) {
      const c = sm(evaluate(nodes[i], ctx, scope), ctx);
      if (isMatrix(c)) {
        // array mode: evaluate everything and pick element-wise
        const allC: (Scalar | Matrix)[] = [];
        const allV: (Scalar | Matrix)[] = [];
        for (let k = 0; k < nodes.length; k += 2) {
          allC.push(k < i ? conds[k / 2] : k === i ? c : sm(evaluate(nodes[k], ctx, scope), ctx));
          allV.push(sm(evalArg(nodes[k + 1], ctx, scope, 0), ctx));
        }
        const [R, C] = broadcastDims([...allC, ...allV]);
        const out: Scalar[][] = [];
        for (let r = 0; r < R; r++) {
          const row: Scalar[] = [];
          for (let cc = 0; cc < C; cc++) {
            let res: Scalar = ERR.NA;
            for (let k = 0; k < allC.length; k++) {
              const cv = allC[k];
              const t = truthy(isMatrix(cv) ? at(cv, r, cc) : cv);
              if (typeof t !== 'boolean') {
                res = t;
                break;
              }
              if (t) {
                const v = allV[k];
                res = isMatrix(v) ? at(v, r, cc) : v;
                break;
              }
            }
            row.push(res);
          }
          out.push(row);
        }
        return matrix(out);
      }
      conds.push(c);
      const t = truthy(c);
      if (typeof t !== 'boolean') return t;
      if (t) return evalArg(nodes[i + 1], ctx, scope, 0);
    }
    return ERR.NA;
  },
});

function errorSwap(nodes: Node[], ctx: EvalContext, scope: Scope | undefined, only: string | null): Value {
  const v = evaluate(nodes[0], ctx, scope);
  const hit = (x: Scalar) => isErr(x) && (only === null || asErr(x).error === only);
  if (isRef(v) || isMatrix(v)) {
    const g = toGrid(v, ctx);
    if (g.length === 1 && g[0].length === 1) {
      if (!hit(g[0][0])) return v;
      return evalArg(nodes[1], ctx, scope, 0);
    }
    if (!g.some((row) => row.some(hit))) return v;
    const alt = sm(evalArg(nodes[1], ctx, scope, 0), ctx);
    const src = matrix(g);
    const [R, C] = broadcastDims([src, alt]);
    const out: Scalar[][] = [];
    for (let r = 0; r < R; r++) {
      const row: Scalar[] = [];
      for (let c = 0; c < C; c++) {
        const x = at(src, r, c);
        row.push(hit(x) ? (isMatrix(alt) ? at(alt, r, c) : alt) : x === null ? 0 : x);
      }
      out.push(row);
    }
    return matrix(out);
  }
  if (isLambda(v)) return v;
  if (hit(v)) return evalArg(nodes[1], ctx, scope, 0);
  return v;
}

define({
  name: 'IFERROR',
  category: L,
  min: 2,
  max: 2,
  syntax: 'value, value_if_error',
  desc: 'Returns value_if_error if the expression is an error and the value of the expression itself otherwise.',
  lazy: (nodes, ctx, scope) => errorSwap(nodes, ctx, scope, null),
});

define({
  name: 'IFNA',
  category: L,
  min: 2,
  max: 2,
  syntax: 'value, value_if_na',
  desc: 'Returns the value you specify if the expression resolves to #N/A, otherwise returns the result of the expression.',
  lazy: (nodes, ctx, scope) => errorSwap(nodes, ctx, scope, '#N/A'),
});

/* ------------------------------------------------------ AND / OR / XOR */

function logicalValues(args: Value[], ctx: EvalContext): boolean[] {
  const out: boolean[] = [];
  eachArg(args, ctx, (v, direct) => {
    if (isErr(v)) throw asErr(v);
    if (typeof v === 'boolean') out.push(v);
    else if (typeof v === 'number') out.push(v !== 0);
    else if (v === null) {
      if (direct) out.push(false);
    } else if (direct) out.push(bool(v));
  });
  if (!out.length) throw ERR.VALUE;
  return out;
}

fn('AND', L, 1, 255, 'r', 'logical1, [logical2], …', 'Returns TRUE if all of its arguments are TRUE.', (args, ctx) => logicalValues(args, ctx).every(Boolean));
fn('OR', L, 1, 255, 'r', 'logical1, [logical2], …', 'Returns TRUE if any argument is TRUE.', (args, ctx) => logicalValues(args, ctx).some(Boolean));
fn('XOR', L, 1, 254, 'r', 'logical1, [logical2], …', 'Returns a logical exclusive OR of all arguments.', (args, ctx) => logicalValues(args, ctx).filter(Boolean).length % 2 === 1);
fn('NOT', L, 1, 1, 'v', 'logical', 'Changes FALSE to TRUE, or TRUE to FALSE.', ([v]) => !bool(v));
fn('TRUE', L, 0, 0, 'v', '', 'Returns the logical value TRUE.', () => true);
fn('FALSE', L, 0, 0, 'v', '', 'Returns the logical value FALSE.', () => false);

/* ------------------------------------------------------ SWITCH/CHOOSE */

function sameValue(a: Scalar, b: Scalar): boolean {
  if (typeof a === 'number' && typeof b === 'number') return numEquals(a, b);
  if (typeof a !== typeof b && !(a === null || b === null)) return false;
  return compare(a, b) === 0;
}

define({
  name: 'SWITCH',
  category: L,
  min: 3,
  max: 254,
  syntax: 'expression, value1, result1, [value2, result2], …, [default]',
  desc: 'Evaluates an expression against a list of values and returns the result corresponding to the first matching value.',
  lazy: (nodes, ctx, scope) => {
    const x = sm(evaluate(nodes[0], ctx, scope), ctx);
    const rest = nodes.slice(1);
    const hasDefault = rest.length % 2 === 1;
    const pairs = Math.floor(rest.length / 2);
    const pick = (val: Scalar): Value => {
      if (isErr(val)) return val;
      for (let i = 0; i < pairs; i++) {
        const cand = toScalar(evaluate(rest[i * 2], ctx, scope), ctx);
        if (isErr(cand)) return cand;
        if (sameValue(val, cand)) return evalArg(rest[i * 2 + 1], ctx, scope, 0);
      }
      return hasDefault ? evalArg(rest[rest.length - 1], ctx, scope, 0) : ERR.NA;
    };
    if (isMatrix(x)) return matrix(x.data.map((row) => row.map((v) => toScalar(pick(v), ctx))));
    return pick(x);
  },
});

define({
  name: 'CHOOSE',
  category: 'Lookup',
  min: 2,
  max: 255,
  syntax: 'index_num, value1, [value2], …',
  desc: 'Chooses a value or action to perform from a list of values, based on an index number.',
  lazy: (nodes, ctx, scope) => {
    const idx = sm(evaluate(nodes[0], ctx, scope), ctx);
    const n = nodes.length - 1;
    if (isMatrix(idx)) {
      const vals = nodes.slice(1).map((nd) => sm(evalArg(nd, ctx, scope, 0), ctx));
      const [R, C] = broadcastDims([idx, ...vals]);
      const out: Scalar[][] = [];
      for (let r = 0; r < R; r++) {
        const row: Scalar[] = [];
        for (let c = 0; c < C; c++) {
          const iv = at(idx, r, c);
          if (isErr(iv)) {
            row.push(iv);
            continue;
          }
          const k = Math.trunc(num(iv));
          if (k < 1 || k > n) row.push(ERR.VALUE);
          else {
            const v = vals[k - 1];
            row.push(isMatrix(v) ? at(v, r, c) : v);
          }
        }
        out.push(row);
      }
      return matrix(out);
    }
    if (isErr(idx)) return idx;
    const k = Math.trunc(num(idx));
    if (k < 1 || k > n) return ERR.VALUE;
    return evalArg(nodes[k], ctx, scope, 0);
  },
});

/* ------------------------------------------------------- LET / LAMBDA */

define({
  name: 'LET',
  category: L,
  min: 3,
  max: 253,
  syntax: 'name1, name_value1, [name2, name_value2], …, calculation',
  desc: 'Assigns names to calculation results so they can be reused inside the formula.',
  lazy: (nodes, ctx, scope) => {
    if (nodes.length % 2 === 0) return ERR.VALUE;
    const s: Scope = new Map(scope ?? []);
    for (let i = 0; i < nodes.length - 1; i += 2) {
      const nm = nodes[i];
      if (nm.t !== 'name' || nm.sheet) return ERR.VALUE;
      s.set(nm.name, evaluate(nodes[i + 1], ctx, s));
    }
    return evaluate(nodes[nodes.length - 1], ctx, s);
  },
});

define({
  name: 'LAMBDA',
  category: L,
  min: 1,
  max: 254,
  syntax: '[parameter1, parameter2, …], calculation',
  desc: 'Creates a custom, reusable function that can be called by a friendly name (define it in the Name Manager).',
  lazy: (nodes, _ctx, scope) => {
    const params: string[] = [];
    for (const nd of nodes.slice(0, -1)) {
      if (nd.t !== 'name' || nd.sheet) return ERR.VALUE;
      params.push(nd.name);
    }
    return makeLambda(params, nodes[nodes.length - 1], scope);
  },
});

define({
  name: 'ISOMITTED',
  category: 'Information',
  min: 1,
  max: 1,
  syntax: 'argument',
  desc: 'Checks whether a LAMBDA parameter was left out when the function was called.',
  lazy: (nodes, _ctx, scope) => {
    const nd = nodes[0];
    if (nd.t !== 'name') return false;
    return isOmitted(nd.name, scope);
  },
});

function lambdaArg(v: Value): LambdaValue {
  if (!isLambda(v)) throw ERR.VALUE;
  return v;
}

function scalarResult(v: Value, ctx: EvalContext): Scalar {
  if (isRef(v) || isMatrix(v)) {
    const g = toGrid(v, ctx);
    if (g.length !== 1 || g[0].length !== 1) return ERR.CALC;
    return g[0][0] ?? 0;
  }
  if (isLambda(v)) return ERR.CALC;
  return v ?? 0;
}

define({
  name: 'MAP',
  category: L,
  min: 2,
  max: 254,
  args: 'r',
  syntax: 'array1, [array2], …, lambda',
  desc: 'Returns an array formed by mapping each value in the array(s) to a new value by applying a LAMBDA.',
  fn: (args: Value[], ctx: EvalContext) => {
    const f = lambdaArg(args[args.length - 1]);
    const arrays = args.slice(0, -1).map((a) => matrix(grid(a, ctx)));
    const [R, C] = broadcastDims(arrays);
    const out: Scalar[][] = [];
    for (let r = 0; r < R; r++) {
      const row: Scalar[] = [];
      for (let c = 0; c < C; c++) row.push(scalarResult(callLambda(f, arrays.map((m) => at(m, r, c)), ctx), ctx));
      out.push(row);
    }
    return matrix(out);
  },
});

define({
  name: 'REDUCE',
  category: L,
  min: 2,
  max: 3,
  args: 'r',
  syntax: '[initial_value], array, lambda(accumulator, value)',
  desc: 'Reduces an array to an accumulated value by applying a LAMBDA to each value.',
  fn: (args: Value[], ctx: EvalContext) => {
    const f = lambdaArg(args[args.length - 1]);
    let acc: Value = args.length === 3 ? args[0] : null;
    const g = grid(args[args.length - 2], ctx);
    for (const row of g) for (const v of row) acc = callLambda(f, [acc, v], ctx);
    return acc;
  },
});

define({
  name: 'SCAN',
  category: L,
  min: 2,
  max: 3,
  args: 'r',
  syntax: '[initial_value], array, lambda(accumulator, value)',
  desc: 'Scans an array by applying a LAMBDA to each value and returns an array of each intermediate value.',
  fn: (args: Value[], ctx: EvalContext) => {
    const f = lambdaArg(args[args.length - 1]);
    let acc: Value = args.length === 3 ? args[0] : null;
    const g = grid(args[args.length - 2], ctx);
    return matrix(
      g.map((row) =>
        row.map((v) => {
          acc = callLambda(f, [acc, v], ctx);
          return scalarResult(acc, ctx);
        }),
      ),
    );
  },
});

define({
  name: 'MAKEARRAY',
  category: L,
  min: 3,
  max: 3,
  args: 'vvr',
  syntax: 'rows, cols, lambda(row, col)',
  desc: 'Returns a calculated array of a specified row and column size, by applying a LAMBDA.',
  fn: (args: Value[], ctx: EvalContext) => {
    const R = int(args[0] as Scalar);
    const C = int(args[1] as Scalar);
    if (R < 1 || C < 1) throw ERR.VALUE;
    const f = lambdaArg(args[2]);
    const out: Scalar[][] = [];
    for (let r = 1; r <= R; r++) {
      const row: Scalar[] = [];
      for (let c = 1; c <= C; c++) row.push(scalarResult(callLambda(f, [r, c], ctx), ctx));
      out.push(row);
    }
    return matrix(out);
  },
});

define({
  name: 'BYROW',
  category: L,
  min: 2,
  max: 2,
  args: 'r',
  syntax: 'array, lambda(row)',
  desc: 'Applies a LAMBDA to each row and returns an array of the results.',
  fn: ([a, f]: Value[], ctx: EvalContext) => {
    const fnv = lambdaArg(f);
    return matrix(grid(a, ctx).map((row) => [scalarResult(callLambda(fnv, [matrix([row])], ctx), ctx)]));
  },
});

define({
  name: 'BYCOL',
  category: L,
  min: 2,
  max: 2,
  args: 'r',
  syntax: 'array, lambda(column)',
  desc: 'Applies a LAMBDA to each column and returns an array of the results.',
  fn: ([a, f]: Value[], ctx: EvalContext) => {
    const fnv = lambdaArg(f);
    const g = grid(a, ctx);
    const C = g[0]?.length ?? 0;
    const row: Scalar[] = [];
    for (let c = 0; c < C; c++) row.push(scalarResult(callLambda(fnv, [matrix(g.map((r) => [r[c]]))], ctx), ctx));
    return matrix([row]);
  },
});
