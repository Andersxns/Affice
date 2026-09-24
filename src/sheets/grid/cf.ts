import { isErr, isMatrix, toBool, type Scalar } from '../engine/values';
import { parseInput, serialToParts } from '../format/numfmt';
import type { CondFormat, CellStyle } from '../model/types';
import type { Sheet } from '../model/workbook';
import type { SheetDoc } from '../doc';

export interface CfEffect {
  style?: CellStyle;
  scale?: string;
  bar?: { from: number; to: number; color: string; negative: boolean; axis: number };
  icon?: { set: NonNullable<CondFormat['iconSet']>; index: number; total: number };
  hideValue?: boolean;
}

interface RuleStats {
  nums: number[];
  sorted: number[];
  min: number;
  max: number;
  avg: number;
  std: number;
  counts: Map<string, number>;
}

const keyOf = (v: Scalar | undefined) => (typeof v === 'string' ? 's' + v.toLowerCase() : typeof v + ':' + String(v));

function hexToRgb(h: string): [number, number, number] {
  let s = h.replace('#', '');
  if (s.length === 3) s = s.split('').map((c) => c + c).join('');
  const n = parseInt(s.slice(0, 6), 16);
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
}

function mix(a: string, b: string, t: number): string {
  const x = hexToRgb(a);
  const y = hexToRgb(b);
  const c = x.map((v, i) => Math.round(v + (y[i] - v) * Math.min(1, Math.max(0, t))));
  return '#' + c.map((v) => v.toString(16).padStart(2, '0')).join('');
}

export class CfEvaluator {
  private stats = new Map<string, RuleStats>();
  private version = -1;
  private sheetId = -1;

  constructor(private doc: SheetDoc) {}

  private reset(sheet: Sheet): void {
    if (this.version !== this.doc.version || this.sheetId !== sheet.id) {
      this.stats.clear();
      this.version = this.doc.version;
      this.sheetId = sheet.id;
    }
  }

  private statsFor(sheet: Sheet, rule: CondFormat): RuleStats {
    let st = this.stats.get(rule.id);
    if (st) return st;
    const nums: number[] = [];
    const counts = new Map<string, number>();
    for (const rg0 of rule.ranges) {
      const rg = this.doc.clampToUsed(sheet, rg0);
      for (let r = rg.r1; r <= rg.r2; r++)
        for (let c = rg.c1; c <= rg.c2; c++) {
          const v = this.doc.value(sheet, r, c);
          if (v === undefined || v === null || v === '') continue;
          if (typeof v === 'number') nums.push(v);
          const k = keyOf(v);
          counts.set(k, (counts.get(k) ?? 0) + 1);
        }
    }
    const sorted = nums.slice().sort((a, b) => a - b);
    const avg = nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : 0;
    const std = nums.length > 1 ? Math.sqrt(nums.reduce((s, x) => s + (x - avg) ** 2, 0) / (nums.length - 1)) : 0;
    st = { nums, sorted, min: sorted[0] ?? 0, max: sorted[sorted.length - 1] ?? 0, avg, std, counts };
    this.stats.set(rule.id, st);
    return st;
  }

  private operand(text: string | undefined, sheet: Sheet, r: number, c: number, rule: CondFormat): Scalar {
    if (text === undefined) return null;
    if (text.startsWith('=')) {
      const anchor = { r: rule.ranges[0].r1, c: rule.ranges[0].c1 };
      const v = this.doc.engine.evaluateText(text, sheet.id, r, c, anchor);
      return isMatrix(v) ? v.data[0][0] : v;
    }
    const p = parseInput(text);
    return p.value;
  }

  private matches(rule: CondFormat, sheet: Sheet, r: number, c: number, v: Scalar | undefined): boolean {
    const text = typeof v === 'string' ? v.toLowerCase() : v === undefined || v === null ? '' : String(v).toLowerCase();
    switch (rule.type) {
      case 'cell': {
        if (v === undefined || v === null) return false;
        const a = this.operand(rule.values?.[0], sheet, r, c, rule);
        const b = this.operand(rule.values?.[1], sheet, r, c, rule);
        const cmp = (x: Scalar, y: Scalar) => {
          if (typeof x === 'number' && typeof y === 'number') return x - y;
          const p = String(x ?? '').toLowerCase();
          const q = String(y ?? '').toLowerCase();
          return p === q ? 0 : p < q ? -1 : 1;
        };
        switch (rule.op) {
          case 'between':
            return cmp(v, a) >= 0 && cmp(v, b) <= 0;
          case 'notBetween':
            return cmp(v, a) < 0 || cmp(v, b) > 0;
          case 'equal':
            return cmp(v, a) === 0;
          case 'notEqual':
            return cmp(v, a) !== 0;
          case 'greater':
            return cmp(v, a) > 0;
          case 'less':
            return cmp(v, a) < 0;
          case 'greaterEqual':
            return cmp(v, a) >= 0;
          case 'lessEqual':
            return cmp(v, a) <= 0;
        }
        return false;
      }
      case 'text':
        return text.includes((rule.text ?? '').toLowerCase());
      case 'notText':
        return !text.includes((rule.text ?? '').toLowerCase());
      case 'begins':
        return text.startsWith((rule.text ?? '').toLowerCase());
      case 'ends':
        return text.endsWith((rule.text ?? '').toLowerCase());
      case 'blank':
        return v === undefined || v === null || v === '';
      case 'notBlank':
        return !(v === undefined || v === null || v === '');
      case 'error':
        return isErr(v);
      case 'notError':
        return !isErr(v);
      case 'duplicate':
      case 'unique': {
        if (v === undefined || v === null || v === '') return false;
        const n = this.statsFor(sheet, rule).counts.get(keyOf(v)) ?? 0;
        return rule.type === 'duplicate' ? n > 1 : n === 1;
      }
      case 'top': {
        if (typeof v !== 'number') return false;
        const s = this.statsFor(sheet, rule).sorted;
        if (!s.length) return false;
        const n = rule.percent ? Math.max(1, Math.floor((s.length * (rule.rank ?? 10)) / 100)) : rule.rank ?? 10;
        if (rule.bottom) return v <= s[Math.min(s.length - 1, n - 1)];
        return v >= s[Math.max(0, s.length - n)];
      }
      case 'average': {
        if (typeof v !== 'number') return false;
        const st = this.statsFor(sheet, rule);
        if (rule.above === false) return rule.equalAverage ? v <= st.avg : v < st.avg;
        return rule.equalAverage ? v >= st.avg : v > st.avg;
      }
      case 'date': {
        if (typeof v !== 'number') return false;
        const today = Math.floor(this.doc.engine.nowSerial());
        const d = Math.floor(v);
        const tp = serialToParts(today);
        const vp = serialToParts(d);
        const dow = (tp.dow + 6) % 7; // Monday = 0
        const weekStart = today - dow;
        switch (rule.datePeriod) {
          case 'yesterday':
            return d === today - 1;
          case 'today':
            return d === today;
          case 'tomorrow':
            return d === today + 1;
          case 'last7Days':
            return d > today - 7 && d <= today;
          case 'lastWeek':
            return d >= weekStart - 7 && d < weekStart;
          case 'thisWeek':
            return d >= weekStart && d < weekStart + 7;
          case 'nextWeek':
            return d >= weekStart + 7 && d < weekStart + 14;
          case 'lastMonth': {
            const lm = tp.m === 1 ? 12 : tp.m - 1;
            const ly = tp.m === 1 ? tp.y - 1 : tp.y;
            return vp.m === lm && vp.y === ly;
          }
          case 'thisMonth':
            return vp.m === tp.m && vp.y === tp.y;
          case 'nextMonth': {
            const nm = tp.m === 12 ? 1 : tp.m + 1;
            const ny = tp.m === 12 ? tp.y + 1 : tp.y;
            return vp.m === nm && vp.y === ny;
          }
        }
        return false;
      }
      case 'formula': {
        const res = this.operand(rule.values?.[0], sheet, r, c, rule);
        if (isErr(res)) return false;
        const b = toBool(res);
        return b === true;
      }
    }
    return false;
  }

  /** Combined conditional-formatting effect for a cell (undefined when no rule applies). */
  effect(sheet: Sheet, r: number, c: number): CfEffect | undefined {
    if (!sheet.cf.length) return undefined;
    this.reset(sheet);
    let out: CfEffect | undefined;
    for (const rule of sheet.cf) {
      if (!rule.ranges.some((rg) => r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2)) continue;
      const v = this.doc.value(sheet, r, c);
      if (rule.type === 'colorScale') {
        if (typeof v !== 'number' || out?.scale) continue;
        const st = this.statsFor(sheet, rule);
        const colors = rule.colors?.length ? rule.colors : ['#f8696b', '#ffeb84', '#63be7b'];
        const t = st.max === st.min ? 0.5 : (v - st.min) / (st.max - st.min);
        let color: string;
        if (colors.length >= 3) {
          const mid = st.sorted[Math.floor((st.sorted.length - 1) / 2)] ?? (st.min + st.max) / 2;
          const tm = st.max === st.min ? 0.5 : (mid - st.min) / (st.max - st.min);
          color = t <= tm ? mix(colors[0], colors[1], tm ? t / tm : 0) : mix(colors[1], colors[2], tm < 1 ? (t - tm) / (1 - tm) : 1);
        } else color = mix(colors[0], colors[1], t);
        out = { ...(out ?? {}), scale: color };
        continue;
      }
      if (rule.type === 'dataBar') {
        if (typeof v !== 'number' || out?.bar) continue;
        const st = this.statsFor(sheet, rule);
        const lo = Math.min(0, st.min);
        const hi = Math.max(0, st.max);
        const span = hi - lo || 1;
        const axis = (0 - lo) / span;
        const pos = (v - lo) / span;
        out = {
          ...(out ?? {}),
          bar: { from: Math.min(axis, pos), to: Math.max(axis, pos), color: v < 0 ? '#e5484d' : rule.colors?.[0] ?? '#638ec6', negative: v < 0, axis },
          hideValue: rule.showValue === false,
        };
        continue;
      }
      if (rule.type === 'iconSet') {
        if (typeof v !== 'number' || out?.icon) continue;
        const st = this.statsFor(sheet, rule);
        const set = rule.iconSet ?? '3Arrows';
        const total = Number(set[0]);
        const t = st.max === st.min ? 1 : (v - st.min) / (st.max - st.min);
        const index = Math.min(total - 1, Math.floor(t * total - 1e-9 + (t === 1 ? 0 : 0)));
        out = { ...(out ?? {}), icon: { set, index: Math.max(0, index), total }, hideValue: rule.showValue === false };
        continue;
      }
      if (!this.matches(rule, sheet, r, c, v)) continue;
      const style = { ...(rule.style ?? {}) };
      out = { ...(out ?? {}), style: { ...style, ...(out?.style ?? {}) } };
      if (rule.stop) break;
    }
    return out;
  }
}
