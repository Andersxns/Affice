import { paragraphNumbers, walkEls, type Presentation, type TextBody } from '../model';

function paraText(body: TextBody, i: number): string {
  return body.paras[i].runs.map((r) => (r.link ? `[${r.text}](${r.link})` : r.b && r.text.trim() ? `**${r.text}**` : r.i && r.text.trim() ? `*${r.text}*` : r.text)).join('');
}

/** Slide titles and text as a Markdown outline (with speaker notes as quotes). */
export function outlineMarkdown(pres: Presentation, title: string): string {
  const out: string[] = [`# ${title.replace(/\.[^.]+$/, '')}`, ''];
  pres.slides.forEach((s, i) => {
    const titleEl = s.elements.find((e) => e.ph === 'title' || e.ph === 'ctrTitle');
    const t = titleEl && titleEl.type === 'shape' && titleEl.text ? titleEl.text.paras.map((p) => p.runs.map((r) => r.text).join('')).join(' ').trim() : '';
    out.push(`## ${i + 1}. ${t || 'Slide'}${s.hidden ? ' (hidden)' : ''}`, '');
    walkEls(s.elements, (e) => {
      if (e === titleEl) return;
      if (e.type === 'shape' && e.text) {
        const body = e.text;
        const numbers = paragraphNumbers(body.paras);
        body.paras.forEach((p, j) => {
          const text = paraText(body, j).trim();
          if (!text) return;
          const lvl = p.level ?? 0;
          // an explicit bullet setting wins; otherwise content placeholders are bulleted (like the renderer)
          const bullet = p.bullet ? p.bullet.type !== 'none' : e.ph === 'body' || e.ph === 'obj';
          const num = p.bullet?.type === 'num' ? /^(\d+)[.)]$/.exec(numbers[j] ?? '') : null;
          out.push(num ? `${'   '.repeat(lvl)}${num[1]}. ${text}` : bullet ? `${'  '.repeat(lvl)}- ${text}` : text);
        });
        out.push('');
      } else if (e.type === 'table') {
        const rows = e.rows.map((r) => r.cells.map((c) => c.text.paras.map((p) => p.runs.map((x) => x.text).join('')).join(' ').replace(/\|/g, '\\|')));
        if (rows.length) {
          out.push(`| ${rows[0].join(' | ')} |`, `| ${rows[0].map(() => '---').join(' | ')} |`, ...rows.slice(1).map((r) => `| ${r.join(' | ')} |`), '');
        }
      } else if (e.type === 'image' && e.alt) out.push(`![${e.alt}](image)`, '');
      else if (e.type === 'chart') {
        const c = e.chart;
        out.push(`*Chart${c.title ? `: ${c.title}` : ''}*`, '');
        out.push(`| | ${c.series.map((x) => x.name).join(' | ')} |`, `| --- | ${c.series.map(() => '---').join(' | ')} |`, ...c.categories.map((cat, k) => `| ${cat} | ${c.series.map((x) => x.values[k] ?? '').join(' | ')} |`), '');
      }
    });
    if (s.notes?.trim()) out.push(...s.notes.trim().split('\n').map((l) => `> ${l}`), '');
  });
  return out.join('\n').replace(/\n{3,}/g, '\n\n');
}
