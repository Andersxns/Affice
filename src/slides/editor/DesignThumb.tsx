import type { Presentation, ShapeEl } from '../model';
import { createPresentation } from '../themes';
import { Thumb } from './panels';

const cache = new Map<string, Presentation>();

/** A title slide of a design showing "Aa" in its fonts, like PowerPoint's theme gallery. */
function samplePres(id: string): Presentation {
  let p = cache.get(id);
  if (p) return p;
  p = createPresentation(id);
  const s = p.slides[0];
  s.elements = s.elements
    .filter((e) => e.ph !== 'subTitle')
    .map((e) => (e.ph === 'ctrTitle' && e.type === 'shape' ? ({ ...e, y: 220, h: 280, text: { ...(e as ShapeEl).text!, anchor: 'm', paras: [{ runs: [{ text: 'Aa', size: 110 }], align: 'center' }] } } as ShapeEl) : e));
  cache.set(id, p);
  return p;
}

export function DesignThumb({ id, width = 104 }: { id: string; width?: number }) {
  const p = samplePres(id);
  return <Thumb pres={p} slide={p.slides[0]} index={0} width={width} eager />;
}
