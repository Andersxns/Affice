import type { CSSProperties } from 'react';
import type { TemplateMeta } from '@/templates/catalog';

/** Lightweight CSS illustrations of each template for the dashboard. */
export function TemplatePreview({ t }: { t: TemplateMeta }) {
  const style = { ['--tp-accent' as string]: t.preview.accent, ['--tp-bg' as string]: t.preview.bg ?? '#fff' } as CSSProperties;
  if (t.kind === 'doc') return <DocPreview variant={t.preview.variant} style={style} />;
  if (t.kind === 'sheet') return <SheetPreview variant={t.preview.variant} style={style} />;
  return <SlidePreview variant={t.preview.variant} style={style} />;
}

const L = ({ w, h = 3, c, mt = 3 }: { w: number; h?: number; c?: string; mt?: number }) => (
  <span className="tp-line" style={{ width: `${w}%`, height: h, background: c, marginTop: mt }} />
);

function DocPreview({ variant, style }: { variant: string; style: CSSProperties }) {
  return (
    <div className={`tp-page tp-${variant}`} style={style}>
      {variant === 'resume' && (
        <>
          <div className="tp-row">
            <span className="tp-avatar" />
            <div style={{ flex: 1 }}>
              <L w={70} h={6} c="var(--tp-accent)" />
              <L w={45} />
            </div>
          </div>
          <L w={30} h={4} c="var(--tp-accent)" mt={9} />
          <L w={100} />
          <L w={92} />
          <L w={80} />
          <L w={30} h={4} c="var(--tp-accent)" mt={8} />
          <L w={96} />
          <L w={88} />
          <div className="tp-chips">
            <span />
            <span />
            <span />
          </div>
        </>
      )}
      {variant === 'letter' && (
        <>
          <div className="tp-bar" />
          <L w={40} h={5} c="var(--tp-accent)" mt={8} />
          <L w={30} />
          <L w={25} mt={10} />
          <L w={100} mt={8} />
          <L w={96} />
          <L w={100} />
          <L w={70} />
          <L w={100} mt={6} />
          <L w={90} />
          <L w={35} mt={10} />
        </>
      )}
      {variant === 'report' && (
        <>
          <div className="tp-cover" />
          <L w={65} h={6} c="var(--tp-accent)" mt={8} />
          <L w={100} />
          <L w={94} />
          <div className="tp-table">
            {Array.from({ length: 9 }, (_, i) => (
              <span key={i} className={i < 3 ? 'h' : ''} />
            ))}
          </div>
          <L w={98} mt={6} />
          <L w={60} />
        </>
      )}
      {variant === 'notes' && (
        <>
          <L w={60} h={6} c="var(--tp-accent)" />
          <L w={35} />
          {[0, 1, 2, 3].map((i) => (
            <div key={i} className="tp-check">
              <span />
              <L w={60 + ((i * 13) % 30)} mt={0} />
            </div>
          ))}
          <L w={30} h={4} c="var(--tp-accent)" mt={8} />
          <L w={96} />
          <L w={84} />
        </>
      )}
      {variant === 'essay' && (
        <>
          <L w={30} />
          <L w={28} />
          <L w={32} />
          <L w={50} h={4} c="#333" mt={10} />
          {[100, 98, 100, 95, 100, 70, 100, 96].map((w, i) => (
            <L key={i} w={w} mt={6} />
          ))}
        </>
      )}
      {variant === 'newsletter' && (
        <>
          <div className="tp-banner">
            <L w={60} h={7} c="#fff" mt={0} />
          </div>
          <div className="tp-cols">
            <div>
              <L w={80} h={4} c="var(--tp-accent)" />
              <L w={100} />
              <L w={95} />
              <L w={100} />
              <span className="tp-img" />
            </div>
            <div>
              <span className="tp-img" />
              <L w={80} h={4} c="var(--tp-accent)" />
              <L w={100} />
              <L w={90} />
            </div>
          </div>
        </>
      )}
    </div>
  );
}

function SheetPreview({ variant, style }: { variant: string; style: CSSProperties }) {
  const cols = variant === 'schedule' ? 6 : 5;
  const rows = 7;
  return (
    <div className={`tp-sheet tp-${variant}`} style={style}>
      <div className="tp-sheet-title">
        <L w={55} h={6} c="var(--tp-accent)" mt={0} />
      </div>
      <div className="tp-grid" style={{ gridTemplateColumns: `repeat(${cols}, 1fr)` }}>
        {Array.from({ length: cols * rows }, (_, i) => {
          const r = Math.floor(i / cols);
          const c = i % cols;
          let cls = '';
          if (r === 0) cls = 'h';
          else if (variant === 'schedule' && c > 0 && (r + c) % 3 === 0) cls = 'block';
          else if (variant === 'tracker' && c === cols - 1) cls = 'bar';
          else if (variant === 'grades' && c === cols - 1) cls = r % 3 === 0 ? 'good' : 'ok';
          else if (c === cols - 1) cls = 'num';
          return (
            <span key={i} className={cls} style={cls === 'bar' ? { ['--w' as string]: `${30 + ((r * 23) % 70)}%` } : undefined} />
          );
        })}
      </div>
      {(variant === 'budget' || variant === 'invoice') && (
        <div className="tp-mini">
          {variant === 'budget' ? (
            <div className="tp-donut" />
          ) : (
            <div className="tp-total">
              <L w={100} h={5} c="var(--tp-accent)" mt={0} />
            </div>
          )}
        </div>
      )}
    </div>
  );
}

function SlidePreview({ variant, style }: { variant: string; style: CSSProperties }) {
  return (
    <div className={`tp-slide tp-${variant}`} style={style}>
      {variant === 'pitch' && (
        <>
          <div className="tp-glow" />
          <L w={55} h={9} c="#fff" mt={0} />
          <L w={35} h={4} c="rgba(255,255,255,.6)" mt={6} />
          <div className="tp-pill" />
        </>
      )}
      {variant === 'lesson' && (
        <>
          <div className="tp-blob" />
          <L w={50} h={8} c="#14532d" mt={0} />
          <L w={60} h={3} c="#4b7a5a" mt={8} />
          <L w={52} h={3} c="#4b7a5a" />
          <L w={56} h={3} c="#4b7a5a" />
        </>
      )}
      {variant === 'minimal' && (
        <>
          <L w={45} h={8} c="#111" mt={0} />
          <L w={25} h={3} c="#999" mt={8} />
          <div className="tp-rule" />
        </>
      )}
      {variant === 'midnight' && (
        <>
          <div className="tp-orb" />
          <L w={50} h={8} c="#fff" mt={0} />
          <L w={30} h={3} c="#a78bfa" mt={8} />
        </>
      )}
      {variant === 'sunset' && (
        <>
          <div className="tp-sun" />
          <L w={50} h={8} c="#7c2d12" mt={0} />
          <L w={32} h={3} c="#c2410c" mt={8} />
        </>
      )}
    </div>
  );
}
