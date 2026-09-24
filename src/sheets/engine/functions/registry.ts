import type { Node } from '../parser';
import type { EvalContext, Scope } from '../evaluator';
import type { Value } from '../values';

export type Category =
  | 'Math'
  | 'Statistical'
  | 'Text'
  | 'Logical'
  | 'Lookup'
  | 'Date & time'
  | 'Financial'
  | 'Information'
  | 'Engineering'
  | 'Database'
  | 'Web';

/* eslint-disable @typescript-eslint/no-explicit-any */
export interface FnDef {
  name: string;
  category: Category;
  min: number;
  max: number;
  /**
   * Parameter kinds, one letter per parameter (the last one repeats):
   * v = scalar (arrays are lifted element-wise, errors propagate),
   * e = scalar that may be an error, r = raw value (references/arrays passed through).
   */
  args?: string;
  fn?: (args: any, ctx: EvalContext) => Value;
  lazy?: (args: Node[], ctx: EvalContext, scope?: Scope) => Value;
  volatile?: boolean;
  syntax: string;
  desc: string;
}

const REGISTRY = new Map<string, FnDef>();

export function define(def: FnDef): void {
  REGISTRY.set(def.name, def);
}

/** Registers another name for an existing function (e.g. legacy STDEV → STDEV.S). */
export function alias(name: string, target: string, desc?: string): void {
  const def = REGISTRY.get(target);
  if (!def) throw new Error(`alias target ${target} missing`);
  REGISTRY.set(name, { ...def, name, desc: desc ?? def.desc, syntax: def.syntax });
}

export function getFunction(name: string): FnDef | undefined {
  return REGISTRY.get(name);
}

export function allFunctions(): FnDef[] {
  return [...REGISTRY.values()].sort((a, b) => a.name.localeCompare(b.name));
}

export function hasFunction(name: string): boolean {
  return REGISTRY.has(name);
}
