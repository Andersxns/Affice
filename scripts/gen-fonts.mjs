// Generates src/lib/fonts.generated.ts: a registry of bundled font faces (woff2 only).
// Office fonts (Calibri, Cambria, Arial, Times New Roman, Courier New, Georgia) and LibreOffice's
// defaults (Liberation Sans, Serif and Mono) are aliased to metric-compatible open fonts, but prefer
// the real font when it is installed locally.
import fs from 'node:fs';
import path from 'node:path';

const root = path.resolve(path.dirname(new URL(import.meta.url).pathname), '..');
const nm = path.join(root, 'node_modules', '@fontsource');

const ALL_SUBSETS = ['latin', 'latin-ext', 'cyrillic', 'cyrillic-ext', 'greek', 'greek-ext', 'vietnamese'];
const BASIC_SUBSETS = ['latin', 'latin-ext'];
const STYLES = [
  [400, 'normal'],
  [700, 'normal'],
  [400, 'italic'],
  [700, 'italic'],
];

const L = (regular, bold, italic, boldItalic) => ({ '400normal': regular, '700normal': bold, '400italic': italic, '700italic': boldItalic });

/** pkg, family (as registered), category, subsets, aliases: [{ family, locals }] */
const FONTS = [
  {
    pkg: 'carlito', family: 'Carlito', category: 'sans', subsets: ALL_SUBSETS,
    aliases: [{ family: 'Calibri', locals: L(['Calibri'], ['Calibri Bold', 'Calibri-Bold'], ['Calibri Italic', 'Calibri-Italic'], ['Calibri Bold Italic', 'Calibri-BoldItalic']) }],
  },
  {
    pkg: 'caladea', family: 'Caladea', category: 'serif', subsets: ALL_SUBSETS,
    aliases: [{ family: 'Cambria', locals: L(['Cambria'], ['Cambria Bold', 'Cambria-Bold'], ['Cambria Italic', 'Cambria-Italic'], ['Cambria Bold Italic', 'Cambria-BoldItalic']) }],
  },
  {
    pkg: 'arimo', family: 'Arimo', category: 'sans', subsets: ALL_SUBSETS,
    aliases: [
      { family: 'Arial', locals: L(['Arial', 'ArialMT', 'Liberation Sans', 'LiberationSans'], ['Arial Bold', 'Arial-BoldMT', 'Liberation Sans Bold', 'LiberationSans-Bold'], ['Arial Italic', 'Arial-ItalicMT', 'Liberation Sans Italic', 'LiberationSans-Italic'], ['Arial Bold Italic', 'Arial-BoldItalicMT', 'Liberation Sans Bold Italic', 'LiberationSans-BoldItalic']) },
      { family: 'Helvetica', locals: L(['Helvetica', 'Arial', 'ArialMT'], ['Helvetica Bold', 'Helvetica-Bold', 'Arial Bold'], ['Helvetica Oblique', 'Helvetica-Oblique', 'Arial Italic'], ['Helvetica Bold Oblique', 'Helvetica-BoldOblique', 'Arial Bold Italic']) },
      { family: 'Liberation Sans', locals: L(['Liberation Sans', 'LiberationSans'], ['Liberation Sans Bold', 'LiberationSans-Bold'], ['Liberation Sans Italic', 'LiberationSans-Italic'], ['Liberation Sans Bold Italic', 'LiberationSans-BoldItalic']) },
    ],
  },
  {
    pkg: 'tinos', family: 'Tinos', category: 'serif', subsets: ALL_SUBSETS,
    aliases: [
      { family: 'Times New Roman', locals: L(['Times New Roman', 'TimesNewRomanPSMT', 'Liberation Serif', 'LiberationSerif'], ['Times New Roman Bold', 'TimesNewRomanPS-BoldMT', 'Liberation Serif Bold', 'LiberationSerif-Bold'], ['Times New Roman Italic', 'TimesNewRomanPS-ItalicMT', 'Liberation Serif Italic', 'LiberationSerif-Italic'], ['Times New Roman Bold Italic', 'TimesNewRomanPS-BoldItalicMT', 'Liberation Serif Bold Italic', 'LiberationSerif-BoldItalic']) },
      { family: 'Times', locals: L(['Times', 'Times New Roman', 'TimesNewRomanPSMT'], ['Times Bold', 'Times-Bold', 'Times New Roman Bold'], ['Times Italic', 'Times-Italic', 'Times New Roman Italic'], ['Times Bold Italic', 'Times-BoldItalic', 'Times New Roman Bold Italic']) },
      { family: 'Liberation Serif', locals: L(['Liberation Serif', 'LiberationSerif'], ['Liberation Serif Bold', 'LiberationSerif-Bold'], ['Liberation Serif Italic', 'LiberationSerif-Italic'], ['Liberation Serif Bold Italic', 'LiberationSerif-BoldItalic']) },
    ],
  },
  {
    pkg: 'cousine', family: 'Cousine', category: 'mono', subsets: ALL_SUBSETS,
    aliases: [
      { family: 'Courier New', locals: L(['Courier New', 'CourierNewPSMT', 'Liberation Mono', 'LiberationMono'], ['Courier New Bold', 'CourierNewPS-BoldMT', 'Liberation Mono Bold', 'LiberationMono-Bold'], ['Courier New Italic', 'CourierNewPS-ItalicMT', 'Liberation Mono Italic', 'LiberationMono-Italic'], ['Courier New Bold Italic', 'CourierNewPS-BoldItalicMT', 'Liberation Mono Bold Italic', 'LiberationMono-BoldItalic']) },
      { family: 'Liberation Mono', locals: L(['Liberation Mono', 'LiberationMono'], ['Liberation Mono Bold', 'LiberationMono-Bold'], ['Liberation Mono Italic', 'LiberationMono-Italic'], ['Liberation Mono Bold Italic', 'LiberationMono-BoldItalic']) },
    ],
  },
  {
    pkg: 'gelasio', family: 'Gelasio', category: 'serif', subsets: ALL_SUBSETS,
    aliases: [{ family: 'Georgia', locals: L(['Georgia'], ['Georgia Bold', 'Georgia-Bold'], ['Georgia Italic', 'Georgia-Italic'], ['Georgia Bold Italic', 'Georgia-BoldItalic']) }],
  },
  { pkg: 'roboto', family: 'Roboto', category: 'sans', subsets: BASIC_SUBSETS },
  { pkg: 'open-sans', family: 'Open Sans', category: 'sans', subsets: BASIC_SUBSETS },
  { pkg: 'lato', family: 'Lato', category: 'sans', subsets: BASIC_SUBSETS },
  { pkg: 'montserrat', family: 'Montserrat', category: 'sans', subsets: BASIC_SUBSETS },
  { pkg: 'poppins', family: 'Poppins', category: 'sans', subsets: BASIC_SUBSETS },
  { pkg: 'nunito', family: 'Nunito', category: 'sans', subsets: BASIC_SUBSETS },
  { pkg: 'raleway', family: 'Raleway', category: 'sans', subsets: BASIC_SUBSETS },
  { pkg: 'noto-sans', family: 'Noto Sans', category: 'sans', subsets: ALL_SUBSETS },
  { pkg: 'oswald', family: 'Oswald', category: 'display', subsets: BASIC_SUBSETS },
  { pkg: 'bebas-neue', family: 'Bebas Neue', category: 'display', subsets: BASIC_SUBSETS },
  { pkg: 'merriweather', family: 'Merriweather', category: 'serif', subsets: BASIC_SUBSETS },
  { pkg: 'lora', family: 'Lora', category: 'serif', subsets: BASIC_SUBSETS },
  { pkg: 'playfair-display', family: 'Playfair Display', category: 'serif', subsets: BASIC_SUBSETS },
  { pkg: 'eb-garamond', family: 'EB Garamond', category: 'serif', subsets: BASIC_SUBSETS },
  { pkg: 'noto-serif', family: 'Noto Serif', category: 'serif', subsets: ALL_SUBSETS },
  { pkg: 'source-code-pro', family: 'Source Code Pro', category: 'mono', subsets: BASIC_SUBSETS },
  { pkg: 'dancing-script', family: 'Dancing Script', category: 'script', subsets: BASIC_SUBSETS },
  { pkg: 'caveat', family: 'Caveat', category: 'script', subsets: BASIC_SUBSETS },
];

const imports = [];
const faces = [];
let n = 0;

for (const font of FONTS) {
  const dir = path.join(nm, font.pkg);
  const unicode = JSON.parse(fs.readFileSync(path.join(dir, 'unicode.json'), 'utf8'));
  for (const [weight, style] of STYLES) {
    for (const subset of font.subsets) {
      const file = `${font.pkg}-${subset}-${weight}-${style}.woff2`;
      if (!fs.existsSync(path.join(dir, 'files', file))) continue;
      const range = unicode[subset];
      if (!range) continue;
      const v = `f${n++}`;
      imports.push(`import ${v} from '@fontsource/${font.pkg}/files/${file}?url';`);
      faces.push(`  { family: ${JSON.stringify(font.family)}, weight: ${weight}, style: '${style}', url: ${v}, range: ${JSON.stringify(range)}, locals: [] },`);
      for (const alias of font.aliases ?? []) {
        const locals = alias.locals[`${weight}${style}`] ?? [];
        faces.push(`  { family: ${JSON.stringify(alias.family)}, weight: ${weight}, style: '${style}', url: ${v}, range: ${JSON.stringify(range)}, locals: ${JSON.stringify(locals)} },`);
      }
    }
  }
}

const families = FONTS.map((f) => ({ name: f.family, category: f.category, aliases: (f.aliases ?? []).map((a) => a.family) }));

const out = `// AUTO-GENERATED by scripts/gen-fonts.mjs — do not edit by hand.
/* eslint-disable */
${imports.join('\n')}

export interface BundledFace {
  family: string;
  weight: number;
  style: 'normal' | 'italic';
  url: string;
  range: string;
  /** Local font names tried before the bundled file (so installed Office fonts win). */
  locals: string[];
}

export const BUNDLED_FACES: BundledFace[] = [
${faces.join('\n')}
];

export const BUNDLED_FAMILIES: { name: string; category: string; aliases: string[] }[] = ${JSON.stringify(families, null, 2)};
`;

fs.writeFileSync(path.join(root, 'src', 'lib', 'fonts.generated.ts'), out);
console.log(`fonts.generated.ts: ${faces.length} faces from ${n} files`);
