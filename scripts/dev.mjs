// Development runner: Vite dev server + esbuild watch for the main process + Electron with auto-restart.
import { spawn } from 'node:child_process';
import { createServer } from 'vite';
import { context } from 'esbuild';
import electronPath from 'electron';

const server = await createServer({ configFile: 'vite.config.ts' });
await server.listen();
const url = server.resolvedUrls.local[0];
console.log(`\n  renderer: ${url}\n`);

let child = null;
let restarting = false;
function startElectron() {
  const args = ['.'];
  if (process.getuid && process.getuid() === 0) args.push('--no-sandbox');
  child = spawn(electronPath, args, { stdio: 'inherit', env: { ...process.env, VITE_DEV_SERVER_URL: url } });
  child.on('exit', () => {
    if (!restarting) {
      server.close();
      process.exit(0);
    }
  });
}

const ctx = await context({
  bundle: true,
  platform: 'node',
  target: 'node22',
  format: 'cjs',
  external: ['electron'],
  sourcemap: 'inline',
  outExtension: { '.js': '.cjs' },
  outdir: 'dist-electron',
  entryPoints: { main: 'electron/main.ts', preload: 'electron/preload.ts' },
  plugins: [
    {
      name: 'restart-electron',
      setup(b) {
        b.onEnd((result) => {
          if (result.errors.length) return;
          if (child) {
            restarting = true;
            child.once('exit', () => {
              restarting = false;
              startElectron();
            });
            child.kill();
          } else startElectron();
        });
      },
    },
  ],
});
await ctx.watch();
