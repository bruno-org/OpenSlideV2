/**
 * Decides which text can become a native PowerPoint text box and measures where
 * the browser put each line. Anything that cannot be reproduced faithfully is
 * not promoted: it stays baked into the background image.
 */

export type TextRun = {
  text: string;
  fontFamily: string;
  typeface: string;
  bold: boolean;
  measureFamily: string;
  measureWeight: number;
  sizePx: number;
  weight: number;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  color: string;
  alpha: number;
  letterSpacingPx: number;
};

export type TextLine = {
  runs: TextRun[];
  x: number;
  y: number;
  w: number;
  h: number;
  /**
   * Per-character letter-spacing correction (px) so the line occupies the same
   * width in PowerPoint as it does on screen. See `calibrateSpacing`.
   */
  spcAdjustPx: number;
};

export type TextBlock = {
  /** Union of the line boxes, relative to the page canvas, in canvas px. */
  x: number;
  y: number;
  w: number;
  h: number;
  lines: TextLine[];
  align: 'l' | 'ctr' | 'r' | 'just';
  /** Computed line-height in px; drives the fixed line spacing in the export. */
  lineHeightPx: number;
  /** The element whose text we promoted — hidden before the background capture. */
  el: HTMLElement;
};

const INLINE_TEXT_TAGS = new Set([
  'B',
  'CODE',
  'DEL',
  'EM',
  'I',
  'INS',
  'MARK',
  'S',
  'SMALL',
  'SPAN',
  'STRONG',
  'SUB',
  'SUP',
  'U',
  'BR',
]);

const LINE_GROUP_TOLERANCE_PX = 2;

export function collectTextBlocks(root: HTMLElement): TextBlock[] {
  const hostRect = root.getBoundingClientRect();
  const out: TextBlock[] = [];
  walk(root, hostRect, out, 1);
  return out;
}

function walk(el: HTMLElement, hostRect: DOMRect, out: TextBlock[], inheritedAlpha: number): void {
  const cs = getComputedStyle(el);
  if (cs.display === 'none' || cs.visibility === 'hidden') return;

  const alpha = inheritedAlpha * (Number.parseFloat(cs.opacity) || 0);
  if (alpha === 0) return;

  // A transform, clip or filter on an ancestor changes where or how the text
  // paints in ways a flat text box can't reproduce. Stop descending: everything
  // below stays in the background image, pixel-identical.
  if (breaksTextPromotion(cs)) return;

  if (isTextLeaf(el)) {
    if (canPromote(el, cs)) {
      const block = measureBlock(el, cs, hostRect, alpha);
      if (block) out.push(block);
    }
    return;
  }

  for (const child of Array.from(el.children)) {
    if (child instanceof HTMLElement) walk(child, hostRect, out, alpha);
  }
}

function breaksTextPromotion(cs: CSSStyleDeclaration): boolean {
  if (cs.clipPath && cs.clipPath !== 'none') return true;
  if (cs.filter && cs.filter !== 'none') return true;
  if (cs.backdropFilter && cs.backdropFilter !== 'none') return true;
  if (cs.mixBlendMode && cs.mixBlendMode !== 'normal') return true;
  if (cs.maskImage && cs.maskImage !== 'none') return true;
  if (cs.writingMode && cs.writingMode !== 'horizontal-tb') return true;
  if (hasNonTranslateTransform(cs.transform)) return true;
  if (cs.rotate && cs.rotate !== 'none') return true;
  if (cs.scale && cs.scale !== 'none') return true;
  return false;
}

/**
 * `matrix(a, b, c, d, e, f)` is a pure translation when a=d=1 and b=c=0 — that
 * only shifts the box, which the measured rect already accounts for.
 */
function hasNonTranslateTransform(transform: string): boolean {
  if (!transform || transform === 'none') return false;
  const nums = transform
    .match(/matrix\(([^)]+)\)/)?.[1]
    ?.split(',')
    .map(Number);
  if (!nums || nums.length < 6) return true;
  const [a, b, c, d] = nums;
  return (
    Math.abs(a - 1) > 1e-6 || Math.abs(b) > 1e-6 || Math.abs(c) > 1e-6 || Math.abs(d - 1) > 1e-6
  );
}

function canPromote(el: HTMLElement, cs: CSSStyleDeclaration): boolean {
  if (!el.textContent?.trim()) return false;
  if (cs.textShadow && cs.textShadow !== 'none') return false;
  // Gradient-filled text: the glyphs are painted by the background, so making
  // the colour transparent wouldn't hide them and the fill can't be restated.
  const clip = cs.backgroundClip || (cs as unknown as Record<string, string>).webkitBackgroundClip;
  if (clip === 'text') return false;
  if (cs.webkitTextStrokeWidth && Number.parseFloat(cs.webkitTextStrokeWidth) > 0) return false;
  // Clipped overflow can't be honoured by a floating text box.
  if (overflowsClippingAncestor(el)) return false;
  return true;
}

function overflowsClippingAncestor(el: HTMLElement): boolean {
  const rect = el.getBoundingClientRect();
  for (let p = el.parentElement; p; p = p.parentElement) {
    const cs = getComputedStyle(p);
    const clips =
      cs.overflow !== 'visible' || cs.overflowX !== 'visible' || cs.overflowY !== 'visible';
    if (!clips) continue;
    const pr = p.getBoundingClientRect();
    if (
      rect.left < pr.left - 0.5 ||
      rect.right > pr.right + 0.5 ||
      rect.top < pr.top - 0.5 ||
      rect.bottom > pr.bottom + 0.5
    ) {
      return true;
    }
  }
  return false;
}

function isTextLeaf(el: HTMLElement): boolean {
  if (!el.textContent?.trim()) return false;
  for (const child of Array.from(el.childNodes)) {
    if (child.nodeType === Node.TEXT_NODE) continue;
    if (child instanceof HTMLElement) {
      if (INLINE_TEXT_TAGS.has(child.tagName) && isInlineOnly(child)) continue;
    }
    return false;
  }
  return true;
}

function isInlineOnly(el: HTMLElement): boolean {
  for (const child of Array.from(el.childNodes)) {
    if (child.nodeType === Node.TEXT_NODE) continue;
    if (
      child instanceof HTMLElement &&
      INLINE_TEXT_TAGS.has(child.tagName) &&
      isInlineOnly(child)
    ) {
      continue;
    }
    return false;
  }
  return true;
}

type Token = { node: Text; start: number; end: number; rect: DOMRect; text: string };

function measureBlock(
  el: HTMLElement,
  cs: CSSStyleDeclaration,
  hostRect: DOMRect,
  alpha: number,
): TextBlock | null {
  const tokens = tokenise(el);
  if (tokens.length === 0) return null;

  // Group by line box: same visual top within tolerance.
  const lines: Token[][] = [];
  for (const tk of tokens) {
    const last = lines[lines.length - 1];
    const prev = last?.[last.length - 1];
    if (prev && Math.abs(prev.rect.top - tk.rect.top) <= LINE_GROUP_TOLERANCE_PX) last.push(tk);
    else lines.push([tk]);
  }

  const textLines: TextLine[] = [];
  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;

  for (const line of lines) {
    // Leading/trailing whitespace is collapsed at line ends by the browser and
    // would shift a PowerPoint run.
    let from = 0;
    let to = line.length - 1;
    while (from <= to && !line[from].text.trim()) from++;
    while (to >= from && !line[to].text.trim()) to--;
    if (from > to) continue;

    const runs: TextRun[] = [];
    let lx = Infinity;
    let ly = Infinity;
    let rx = -Infinity;
    let by = -Infinity;
    for (let i = from; i <= to; i++) {
      const tk = line[i];
      const parent = tk.node.parentElement;
      if (!parent) continue;
      const run = runFor(parent, tk.text, alpha);
      const prev = runs[runs.length - 1];
      if (prev && sameFormat(prev, run)) prev.text += run.text;
      else runs.push(run);

      lx = Math.min(lx, tk.rect.left);
      ly = Math.min(ly, tk.rect.top);
      rx = Math.max(rx, tk.rect.right);
      by = Math.max(by, tk.rect.bottom);
    }
    if (runs.length === 0 || lx === Infinity) continue;

    minX = Math.min(minX, lx);
    minY = Math.min(minY, ly);
    maxX = Math.max(maxX, rx);
    maxY = Math.max(maxY, by);
    textLines.push({
      runs,
      x: lx - hostRect.left,
      y: ly - hostRect.top,
      w: rx - lx,
      h: by - ly,
      spcAdjustPx: calibrateSpacing(runs, rx - lx),
    });
  }

  if (textLines.length === 0 || minX === Infinity) return null;

  const fontSize = Number.parseFloat(cs.fontSize) || 16;
  const lineHeightPx = resolveLineHeight(cs.lineHeight, fontSize);

  return {
    x: minX - hostRect.left,
    y: minY - hostRect.top,
    w: maxX - minX,
    h: maxY - minY,
    lines: textLines,
    align: mapAlign(cs.textAlign, cs.direction),
    lineHeightPx,
    el,
  };
}

/**
 * One rect per token so lines can be detected exactly as the browser laid them
 * out. Single-line blocks skip straight to whole text nodes — the common case,
 * and measuring every word there is wasted work.
 */
function tokenise(el: HTMLElement): Token[] {
  const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
  const nodes: Text[] = [];
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    if (n.nodeValue && n.nodeValue.length > 0) nodes.push(n as Text);
  }
  if (nodes.length === 0) return [];

  const range = document.createRange();
  range.selectNodeContents(el);
  const singleLine = range.getClientRects().length <= 1;

  const out: Token[] = [];
  for (const node of nodes) {
    const value = node.nodeValue ?? '';
    if (singleLine) {
      if (!value.trim()) continue;
      const r = rectOf(node, 0, value.length);
      if (r) out.push({ node, start: 0, end: value.length, rect: r, text: value });
      continue;
    }
    // Split on whitespace boundaries, keeping the separators as their own
    // tokens so line grouping can drop them at the edges.
    const re = /(\s+|\S+)/g;
    let m = re.exec(value);
    while (m) {
      const start = m.index;
      const end = start + m[0].length;
      const r = rectOf(node, start, end);
      if (r && (r.width > 0 || r.height > 0)) {
        out.push({ node, start, end, rect: r, text: m[0] });
      }
      m = re.exec(value);
    }
  }
  return out;
}

function rectOf(node: Text, start: number, end: number): DOMRect | null {
  const range = document.createRange();
  range.setStart(node, start);
  range.setEnd(node, end);
  const rects = range.getClientRects();
  if (rects.length === 0) return null;
  // A token can straddle a line only when it contains the break itself
  // (collapsed whitespace); the first rect is where it starts.
  return rects[0];
}

function runFor(parent: HTMLElement, text: string, alpha: number): TextRun {
  const cs = getComputedStyle(parent);
  const { hex, a } = parseColor(cs.color);
  const decorations = cs.textDecorationLine || '';
  const spacing = Number.parseFloat(cs.letterSpacing);
  const weight = Number.parseInt(cs.fontWeight, 10) || 400;
  const resolved = resolveFont(cs.fontFamily, weight);
  return {
    text,
    fontFamily: primaryFont(cs.fontFamily),
    typeface: resolved.typeface,
    bold: resolved.bold,
    measureFamily: resolved.measureFamily,
    measureWeight: resolved.measureWeight,
    sizePx: Number.parseFloat(cs.fontSize) || 16,
    weight,
    italic: cs.fontStyle === 'italic' || cs.fontStyle === 'oblique',
    underline: decorations.includes('underline'),
    strike: decorations.includes('line-through'),
    color: hex,
    alpha: a * alpha,
    letterSpacingPx: Number.isFinite(spacing) ? spacing : 0,
  };
}

let measureCtx: CanvasRenderingContext2D | null = null;

function getMeasureCtx(): CanvasRenderingContext2D | null {
  if (measureCtx) return measureCtx;
  measureCtx = document.createElement('canvas').getContext('2d');
  if (measureCtx) {
    // PowerPoint reads the legacy `kern` table only, so fonts that carry their
    // kerning in GPOS (Geist and most modern faces) render unkerned there and
    // every line comes out wider. Measuring unkerned reproduces that, and the
    // calibration below turns the difference into letter spacing.
    measureCtx.fontKerning = 'none';
  }
  return measureCtx;
}

/**
 * PowerPoint only has bold on or off. A CSS weight of 800 (or 500) has no
 * counterpart, so the browser synthesises it — wider glyphs — while PowerPoint
 * renders the family's real bold (or regular), which is narrower.
 *
 * That drift is predictable: measure the line again in the weight PowerPoint
 * will actually use, and spread the difference across the characters as letter
 * spacing. The line then ends on the same x it does on screen.
 *
 * Returns 0 when the gap is implausibly large (a missing font, not a weight
 * mismatch) — squeezing the text then would look worse than leaving it be.
 */
function calibrateSpacing(runs: TextRun[], targetWidthPx: number): number {
  const ctx = getMeasureCtx();
  if (!ctx || targetWidthPx <= 0) return 0;

  let measured = 0;
  let chars = 0;
  for (const run of runs) {
    // Measure exactly what PowerPoint will render: the substitute family when
    // there is one, at the weight the resolved face actually carries.
    const style = run.italic ? 'italic ' : '';
    const name = run.measureFamily;
    const family = /[^\w-]/.test(name) ? `"${name}"` : name;
    ctx.font = `${style}${run.measureWeight} ${run.sizePx}px ${family}`;
    measured += ctx.measureText(run.text).width + run.letterSpacingPx * run.text.length;
    chars += run.text.length;
  }
  if (chars === 0 || measured <= 0) return 0;

  const delta = targetWidthPx - measured;
  if (Math.abs(delta) > targetWidthPx * 0.25) return 0;
  return delta / chars;
}

function sameFormat(a: TextRun, b: TextRun): boolean {
  return (
    a.typeface === b.typeface &&
    a.bold === b.bold &&
    a.fontFamily === b.fontFamily &&
    a.sizePx === b.sizePx &&
    a.weight === b.weight &&
    a.italic === b.italic &&
    a.underline === b.underline &&
    a.strike === b.strike &&
    a.color === b.color &&
    Math.abs(a.alpha - b.alpha) < 0.01 &&
    a.letterSpacingPx === b.letterSpacingPx
  );
}

/**
 * Which family name goes into the PPTX, and at which weight it will render.
 *
 * PowerPoint has no weight axis: a run is bold or it isn't. A CSS weight of 500
 * or 800 therefore has nowhere to go — unless the family ships that weight as
 * its own named face, which is exactly how the fonts bundled with this fork are
 * installed (`tools/install-fonts.mjs`). Point the run at "Geist Black" and the
 * weight survives.
 *
 * Safety net: a family we can't vouch for on the target machine is mapped to a
 * face every Windows install has, chosen here by category, instead of letting
 * PowerPoint pick a substitute on its own. `calibrateSpacing` then measures
 * against that same substitute, so the line still occupies its original width.
 */
function resolveFont(
  familyList: string,
  weight: number,
): { typeface: string; bold: boolean; measureFamily: string; measureWeight: number } {
  const family = primaryFont(familyList);
  // Webfont packages ship variable families under a "… Variable" name ("Geist
  // Variable"), while the installed static faces are plain ("Geist Black").
  const base = family.replace(/\s+Variable$/i, '');
  const named = NAMED_WEIGHT_FAMILIES[base];
  if (named) {
    const face = named[nearestWeight(weight, Object.keys(named).map(Number))];
    if (face) {
      // The base family carries regular and bold; the rest are separate faces.
      const bold = face === base && weight >= 600;
      // A static face and its variable sibling are not metrically identical, so
      // measure against the installed face itself when the system exposes it.
      // Its weight is baked in, hence 400 (or the family's own bold).
      const installed = isFamilyInstalled(face);
      return {
        typeface: face,
        bold,
        measureFamily: installed ? face : family,
        measureWeight: installed ? (bold ? 700 : 400) : weight,
      };
    }
  }
  const bold = weight >= 600;
  const measureWeight = bold ? 700 : 400;
  if (SYSTEM_SAFE_FAMILIES.has(family.toLowerCase())) {
    return { typeface: family, bold, measureFamily: family, measureWeight };
  }
  const substitute = fallbackFamily(familyList);
  return { typeface: substitute, bold, measureFamily: substitute, measureWeight };
}

const installedCache = new Map<string, boolean>();

/**
 * Whether the system exposes this family to the browser. Comparing against a
 * deliberately missing family is the only reliable probe: `fonts.check()`
 * answers about loaded webfonts, not about what the OS has.
 */
function isFamilyInstalled(name: string): boolean {
  const cached = installedCache.get(name);
  if (cached !== undefined) return cached;
  const ctx = getMeasureCtx();
  if (!ctx) return false;
  const probe = 'mmmwwwiiilll0123';
  ctx.font = '72px "__osd_missing_family__"';
  const fallback = ctx.measureText(probe).width;
  ctx.font = `72px "${name}", "__osd_missing_family__"`;
  const candidate = ctx.measureText(probe).width;
  const available = Math.abs(candidate - fallback) > 0.5;
  installedCache.set(name, available);
  return available;
}

/**
 * Named faces installed by `tools/install-fonts.mjs`, per CSS weight.
 *
 * These are the family names the files declare internally, which is what
 * PowerPoint matches on — not the file names. Geist ships `UltraBlack.ttf`
 * declaring itself as "Geist ExtraBold", so trusting the file name here would
 * point 800 at the wrong face. Verified with PrivateFontCollection.
 */
const NAMED_WEIGHT_FAMILIES: Record<string, Record<number, string>> = {
  Geist: {
    100: 'Geist Thin',
    200: 'Geist ExtraLight',
    300: 'Geist Light',
    400: 'Geist',
    500: 'Geist Medium',
    600: 'Geist SemiBold',
    700: 'Geist',
    800: 'Geist ExtraBold',
    900: 'Geist Black',
  },
  'Geist Mono': {
    100: 'Geist Mono Thin',
    200: 'Geist Mono ExtraLight',
    300: 'Geist Mono Light',
    400: 'Geist Mono',
    500: 'Geist Mono Medium',
    600: 'Geist Mono SemiBold',
    700: 'Geist Mono',
    800: 'Geist Mono ExtraBold',
    900: 'Geist Mono Black',
  },
};

/** Present on any stock Windows install, so they need no mapping. */
const SYSTEM_SAFE_FAMILIES = new Set([
  'arial',
  'arial black',
  'calibri',
  'cambria',
  'candara',
  'consolas',
  'constantia',
  'corbel',
  'courier new',
  'georgia',
  'impact',
  'lucida console',
  'palatino linotype',
  'segoe ui',
  'tahoma',
  'times new roman',
  'trebuchet ms',
  'verdana',
]);

function nearestWeight(weight: number, available: number[]): number {
  return available.reduce((best, w) => (Math.abs(w - weight) < Math.abs(best - weight) ? w : best));
}

/** Category of the CSS stack decides the substitute, not the first name. */
function fallbackFamily(familyList: string): string {
  const list = familyList.toLowerCase();
  if (list.includes('monospace') || list.includes('mono')) return 'Consolas';
  if (list.includes('serif') && !list.includes('sans-serif')) return 'Georgia';
  return 'Segoe UI';
}

function primaryFont(family: string): string {
  const first = family.split(',')[0]?.trim() ?? '';
  return first.replace(/^["']|["']$/g, '') || 'Arial';
}

export function parseColor(value: string): { hex: string; a: number } {
  const m = value.match(/rgba?\(([^)]+)\)/);
  if (!m) return { hex: '000000', a: 1 };
  const parts = m[1].split(',').map((p) => Number.parseFloat(p.trim()));
  const [r, g, b] = parts;
  const a = parts.length > 3 ? parts[3] : 1;
  const hex = [r, g, b]
    .map((n) =>
      Math.max(0, Math.min(255, Math.round(n || 0)))
        .toString(16)
        .padStart(2, '0'),
    )
    .join('')
    .toUpperCase();
  return { hex, a: Number.isFinite(a) ? a : 1 };
}

function resolveLineHeight(lineHeight: string, fontSize: number): number {
  if (lineHeight === 'normal') return fontSize * 1.2;
  const px = Number.parseFloat(lineHeight);
  return Number.isFinite(px) ? px : fontSize * 1.2;
}

function mapAlign(textAlign: string, direction: string): 'l' | 'ctr' | 'r' | 'just' {
  switch (textAlign) {
    case 'center':
      return 'ctr';
    case 'right':
      return 'r';
    case 'justify':
      return 'just';
    case 'start':
      return direction === 'rtl' ? 'r' : 'l';
    case 'end':
      return direction === 'rtl' ? 'l' : 'r';
    default:
      return 'l';
  }
}
