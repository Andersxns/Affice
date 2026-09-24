/// <reference types="vite/client" />

declare const __APP_VERSION__: string;

declare module '*.css?inline' {
  const css: string;
  export default css;
}

declare module 'turndown-plugin-gfm' {
  import type TurndownService from 'turndown';
  export const gfm: TurndownService.Plugin;
  export const tables: TurndownService.Plugin;
  export const strikethrough: TurndownService.Plugin;
  export const taskListItems: TurndownService.Plugin;
}

declare module 'jstat' {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const jStat: any;
  export default jStat;
  export { jStat };
}

declare module 'bessel' {
  export function besselj(x: number, n: number): number;
  export function bessely(x: number, n: number): number;
  export function besseli(x: number, n: number): number;
  export function besselk(x: number, n: number): number;
}
