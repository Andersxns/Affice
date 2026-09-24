/// <reference types="vitest/config" />
import { defineConfig, type Plugin } from 'vite';
import react from '@vitejs/plugin-react';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

const pkg = JSON.parse(readFileSync(new URL('./package.json', import.meta.url), 'utf8')) as { version: string };

const CSP = [
  "default-src 'self'",
  "script-src 'self'",
  "style-src 'self' 'unsafe-inline'",
  "img-src 'self' data: blob: https: http:",
  "font-src 'self' data:",
  "media-src 'self' data: blob:",
  "connect-src 'self' data: blob:",
  "worker-src 'self' blob:",
  "object-src 'none'",
  "frame-src 'self' blob: data:",
].join('; ');

/** Adds a strict Content-Security-Policy to production builds (dev needs inline HMR scripts). */
function cspPlugin(): Plugin {
  return {
    name: 'affice-csp',
    apply: 'build',
    transformIndexHtml(html) {
      return html.replace('<head>', `<head>\n    <meta http-equiv="Content-Security-Policy" content="${CSP}" />`);
    },
  };
}

export default defineConfig({
  // Relative asset URLs so the packaged app can load dist/index.html over file://
  base: './',
  plugins: [react(), cspPlugin()],
  define: {
    __APP_VERSION__: JSON.stringify(pkg.version),
  },
  resolve: {
    alias: { '@': fileURLToPath(new URL('./src', import.meta.url)) },
  },
  build: {
    outDir: 'dist',
    emptyOutDir: true,
    target: 'chrome120',
    chunkSizeWarningLimit: 4000,
    sourcemap: false,
    assetsInlineLimit: 0,
  },
  server: { port: 5173, strictPort: false },
  test: {
    environment: 'node',
    include: ['tests/**/*.test.ts'],
  },
});
