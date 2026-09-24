import '@fontsource-variable/inter';
import './styles/tokens.css';
import './styles/base.css';
import './styles/ui.css';
import './styles/shell.css';
import './styles/home.css';
import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { installFonts } from './lib/fonts';
import { api } from './lib/platform';
import { App } from './app/App';

installFonts();

async function waitForImages(root: ParentNode): Promise<void> {
  const imgs = Array.from(root.querySelectorAll('img'));
  await Promise.all(
    imgs.map((img) =>
      img.complete
        ? Promise.resolve()
        : new Promise<void>((resolve) => {
            img.addEventListener('load', () => resolve(), { once: true });
            img.addEventListener('error', () => resolve(), { once: true });
          }),
    ),
  );
}

/** Hidden print window: renders the job's HTML with the app's fonts, then tells main it's ready. */
async function runPrintMode(): Promise<void> {
  document.documentElement.dataset.theme = 'light';
  const job = await api.output.takeJob();
  if (!job) return;
  document.title = job.title ?? 'Affice';
  document.body.className = `print-root ${job.bodyClass ?? ''}`;
  const style = document.createElement('style');
  style.textContent = `html,body,#root{overflow:visible!important;height:auto!important;background:#fff!important;user-select:text}
@page{size:${job.pageWidthIn}in ${job.pageHeightIn}in;margin:${job.marginsIn ? `${job.marginsIn.top}in ${job.marginsIn.right}in ${job.marginsIn.bottom}in ${job.marginsIn.left}in` : '0'}}
${job.css ?? ''}`;
  document.head.appendChild(style);
  const root = document.getElementById('root')!;
  root.innerHTML = job.html;
  // Make sure every font used in the content is actually loaded before printing
  const families = new Set<string>();
  root.querySelectorAll<HTMLElement>('*').forEach((el) => {
    const ff = getComputedStyle(el).fontFamily;
    if (ff) families.add(ff);
  });
  await Promise.all(
    Array.from(families).flatMap((ff) => ['400', '700', 'italic 400'].map((w) => document.fonts.load(`${w} 16px ${ff}`).catch(() => undefined))),
  );
  await document.fonts.ready;
  await waitForImages(root);
  await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
  api.output.ready();
}

if (location.hash === '#print') {
  void runPrintMode();
} else {
  createRoot(document.getElementById('root')!).render(
    <StrictMode>
      <App />
    </StrictMode>,
  );
}
