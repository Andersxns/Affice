// Bundles the Electron main process and preload script with esbuild.
import { build, context } from 'esbuild';

const watch = process.argv.includes('--watch');
const dev = watch || process.env.NODE_ENV === 'development';

const common = {
  bundle: true,
  platform: 'node',
  target: 'node22',
  format: 'cjs',
  external: ['electron'],
  sourcemap: dev ? 'inline' : false,
  minify: !dev,
  logLevel: 'info',
  outExtension: { '.js': '.cjs' },
  outdir: 'dist-electron',
  define: { 'process.env.NODE_ENV': JSON.stringify(dev ? 'development' : 'production') },
};

const entries = { main: 'electron/main.ts', preload: 'electron/preload.ts' };

if (watch) {
  const ctx = await context({ ...common, entryPoints: entries });
  await ctx.watch();
} else {
  await build({ ...common, entryPoints: entries });
}
