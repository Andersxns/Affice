import { ERR, isErr, asErr, makeCriteria, type Scalar, type Value } from '../values';
import type { EvalContext } from '../evaluator';
import { fn, grid, mean, str, sum, variance } from './helpers';

const DB = 'Database' as const;

/** Rows of a database range that satisfy a criteria range (rows are OR-ed, columns AND-ed). */
function select(args: Value[], ctx: EvalContext, needField: boolean): Scalar[] {
  const data = grid(args[0], ctx);
  const crit = grid(args[2], ctx);
  if (data.length < 1 || crit.length < 1) throw ERR.VALUE;
  const headers = data[0].map((h) => (h === null ? '' : String(h)).trim().toLowerCase());
  let col = -1;
  const field = args[1] as Scalar;
  if (field !== null && field !== undefined) {
    if (typeof field === 'number') col = Math.trunc(field) - 1;
    else col = headers.indexOf(str(field).trim().toLowerCase());
    if (col < 0 || col >= headers.length) throw ERR.VALUE;
  } else if (needField) throw ERR.VALUE;
  const critHeaders = crit[0].map((h) => (h === null ? '' : String(h)).trim().toLowerCase());
  const rules = crit.slice(1).map((row) =>
    row
      .map((v, j) => {
        if (v === null || v === '') return null;
        const target = headers.indexOf(critHeaders[j]);
        if (target < 0) return () => false;
        // plain text in database criteria means "begins with"
        const c: Scalar = typeof v === 'string' && !/^[<>=]/.test(v) && !/[*?]/.test(v) ? v + '*' : v;
        const test = makeCriteria(c);
        return (rowVals: Scalar[]) => test(rowVals[target]);
      })
      .filter((x): x is (r: Scalar[]) => boolean => x !== null),
  );
  const out: Scalar[] = [];
  for (const row of data.slice(1)) {
    const hit = !rules.length || rules.some((conds) => conds.every((t) => t(row)));
    if (hit) out.push(col >= 0 ? row[col] : row.some((v) => v !== null) ? 1 : null);
  }
  return out;
}

function nums(vals: Scalar[]): number[] {
  const out: number[] = [];
  for (const v of vals) {
    if (isErr(v)) throw asErr(v);
    if (typeof v === 'number') out.push(v);
  }
  return out;
}

const SYNTAX = 'database, field, criteria';

fn('DSUM', DB, 3, 3, 'r', SYNTAX, 'Adds the numbers in the field column of records in the database that match the criteria.', (a, ctx) => sum(nums(select(a, ctx, true))));
fn('DAVERAGE', DB, 3, 3, 'r', SYNTAX, 'Averages the values in a column of a list or database that match conditions you specify.', (a, ctx) => mean(nums(select(a, ctx, true))));
fn('DCOUNT', DB, 3, 3, 'r', SYNTAX, 'Counts the cells that contain numbers in a database.', (a, ctx) => nums(select(a, ctx, false)).length);
fn('DCOUNTA', DB, 3, 3, 'r', SYNTAX, 'Counts nonblank cells in a database.', (a, ctx) => select(a, ctx, false).filter((v) => v !== null).length);
fn('DMAX', DB, 3, 3, 'r', SYNTAX, 'Returns the maximum value from selected database entries.', (a, ctx) => {
  const xs = nums(select(a, ctx, true));
  return xs.length ? Math.max(...xs) : 0;
});
fn('DMIN', DB, 3, 3, 'r', SYNTAX, 'Returns the minimum value from selected database entries.', (a, ctx) => {
  const xs = nums(select(a, ctx, true));
  return xs.length ? Math.min(...xs) : 0;
});
fn('DPRODUCT', DB, 3, 3, 'r', SYNTAX, 'Multiplies the values in a particular field of records that match the criteria.', (a, ctx) => {
  const xs = nums(select(a, ctx, true));
  return xs.length ? xs.reduce((p, x) => p * x, 1) : 0;
});
fn('DSTDEV', DB, 3, 3, 'r', SYNTAX, 'Estimates the standard deviation based on a sample from selected database entries.', (a, ctx) => Math.sqrt(variance(nums(select(a, ctx, true)), true)));
fn('DSTDEVP', DB, 3, 3, 'r', SYNTAX, 'Calculates the standard deviation based on the entire population of selected database entries.', (a, ctx) => Math.sqrt(variance(nums(select(a, ctx, true)), false)));
fn('DVAR', DB, 3, 3, 'r', SYNTAX, 'Estimates variance based on a sample from selected database entries.', (a, ctx) => variance(nums(select(a, ctx, true)), true));
fn('DVARP', DB, 3, 3, 'r', SYNTAX, 'Calculates variance based on the entire population of selected database entries.', (a, ctx) => variance(nums(select(a, ctx, true)), false));
fn('DGET', DB, 3, 3, 'r', SYNTAX, 'Extracts from a database a single record that matches the specified criteria.', (a, ctx) => {
  const vals = select(a, ctx, true);
  if (!vals.length) throw ERR.VALUE;
  if (vals.length > 1) throw ERR.NUM;
  return vals[0] ?? 0;
});

