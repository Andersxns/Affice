# Contributing to Affice

Thanks for helping make Affice better! This guide explains how to run the app, how the code is organised and what we look for in a pull request.

## Getting started

You need [Node.js](https://nodejs.org/) 22 or newer.

```bash
npm install
npm run dev        # Electron app with hot reload
npm run dev:web    # just the interface, in your browser (no file system or printing)
```

Before opening a pull request, run:

```bash
npm run typecheck
npm test
npm run build
```

The [Build workflow](.github/workflows/build.yml) runs the same checks on every push and pull request.

## How the code is organised

| Folder | What lives there |
| --- | --- |
| `electron/` | The main process: windows, file dialogs and file associations, the settings store (`store.ts`), printing and PDF export (`print.ts`) and LibreOffice conversions for legacy formats (`convert.ts`). `preload.ts` exposes a small, typed API to the interface. |
| `src/app/` | The app shell: title bar and tabs, the workspace store, File backstage, settings, command palette and save/open logic (`fileOps.ts`, `actions.ts`). |
| `src/home/` | The dashboard. |
| `src/docs/` | **Documents**: the TipTap/ProseMirror editor, pagination, ribbon, dialogs and the readers and writers in `formats/` (DOCX, ODT, RTF, HTML, Markdown, EPUB). |
| `src/sheets/` | **Sheets**: the formula engine (`engine/`, with functions grouped by category in `engine/functions/`), the workbook model, the canvas grid (`grid/`), charts, number formats and file formats (`formats/`). |
| `src/slides/` | **Slides**: the presentation model (`model.ts`), preset shapes (`geometry.ts`), designs and layouts (`themes.ts`), the renderer (`render/`), the editing canvas and ribbon (`editor/`), the slide show (`show/`) and PPTX/HTML/PDF/PNG formats (`formats/`). |
| `src/ui/` | Shared controls: ribbon, menus, popovers, dialogs, colour pickers, toasts and tooltips. |
| `src/lib/` | Format registry (`formats.ts`), bundled font registry, platform bridge and helpers. |
| `src/shared/` | Types shared by the main process and the interface. |
| `src/styles/` | Design tokens (`tokens.css`), the shell and shared component styles. |
| `tests/` | Unit tests (Vitest). |
| `build/` | Icons, installer artwork and file-type icons used by electron-builder. |
| `scripts/` | Build helpers: bundling the main process, the dev launcher, the font registry generator and the icon generator. |

Each editor is loaded on demand, so a change to Sheets doesn't make Documents slower to open.

## Guidelines

- **Match the surrounding code.** TypeScript is strict (no unused locals or parameters). Two-space indentation, single quotes, semicolons and trailing commas.
- **Keep the interface easy.** Use plain, friendly words in labels and messages, give every icon-only button a tooltip, and make sure new ribbon groups still fit at 1024 pixels wide (give wide groups a `collapse` priority and an `icon` so they can fold into a button).
- **Respect older machines.** Animations must honour the motion setting (`data-motion` on the root element) and performance mode (`data-perf`). Avoid work on every keystroke that grows with document size.
- **File formats are the heart of Affice.** When you change a reader or writer, round-trip real files: open them, save them, and check the result in Microsoft Office or LibreOffice (`soffice --headless --convert-to pdf file.pptx` makes a quick visual check easy). Never drop content silently; if a format can't keep something, say so to the user.
- **No network calls.** Affice works offline and must never send documents or usage data anywhere.
- **Tests.** Add unit tests for engine code (formulas, number formats, parsers). For editor changes, describe how you tested in the pull request.

## Adding things

- **A spreadsheet function:** add a `define({ name, category, min, max, syntax, desc, … })` call to the matching file in `src/sheets/engine/functions/`, then add cases to `tests/engine.test.ts`.
- **A template:** add it to `src/templates/catalog.ts` and build it in the editor's `templates.ts`.
- **A slide design:** add an entry to `DESIGNS` in `src/slides/themes.ts` (colours, fonts, background and decoration shapes).
- **A font:** add the `@fontsource` package and run `node scripts/gen-fonts.mjs` to regenerate `src/lib/fonts.generated.ts`.
- **Icons:** edit the sources in `build/logo-source/` and run `npm run icons`.

## Releasing

Update `version` in `package.json`, commit, then push a tag such as `v1.2.0`. The Build workflow creates the Windows installers (x64, 32-bit and ARM64, plus a portable app) and the Linux packages (AppImage, .deb, .rpm and .tar.gz) and publishes them as a GitHub release with checksums.

## Reporting bugs

Please include your operating system, the Affice version (**Home → About**), what you did, what you expected and what happened. If a file doesn't open or save correctly, attach a small sample file if you can share it.
