import { createElement } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { designToCssVars } from './design';
import { SlidePageProvider } from './page-context';
import { collectTextBlocks, type TextLine } from './pptx/text-extract';
import { isFrameAnimationSettled, waitForDataWaitfor, waitForFonts } from './print-ready';
import type { SlideModule } from './sdk';

const SLIDE_W = 1920;
const SLIDE_H = 1080;
// 16:9 widescreen in EMU (914400 per inch → 13.333in × 7.5in).
const EMU_W = 12192000;
const EMU_H = 6858000;
const EMU_PER_PX = EMU_W / SLIDE_W;
// The canvas is 1920px across 13.333in → 144 px per inch, so 1px = 0.5pt.
const PT_PER_PX = 0.5;
const CAPTURE_PIXEL_RATIO = 2;

const ANIMATION_TIMEOUT_MS = 15_000;
const POLL_INTERVAL_MS = 100;

const CAPTURE_CLASS = 'os-pptx-capture';

export type EditablePptxProgress = {
  phase: 'processing' | 'generating' | 'done';
  current: number;
  total: number;
  percent: number;
  /** Text boxes promoted so far — surfaced so the UI can report editability. */
  textBoxes: number;
};

/**
 * Export a deck as a PPTX whose text is real, editable PowerPoint text.
 *
 * Fidelity first: each page is captured as a background image with every
 * promoted text made transparent (layout, gradients, shadows, SVG and effects
 * all still paint), then native text boxes are placed on top at the exact
 * pixels the browser measured. Text we can't restate faithfully — gradient
 * fills, text shadows, rotated or clipped text — is left painted in the
 * background instead of being approximated, so the slide always matches the
 * screen; only its editability degrades.
 *
 * One text box per rendered line, positioned on that line's measured rect, with
 * wrap disabled. A substituted font can then change glyph widths but can never
 * re-flow the deck.
 */
export async function exportSlideAsEditablePptx(
  slide: SlideModule,
  slideId: string,
  onProgress?: (progress: EditablePptxProgress) => void,
): Promise<void> {
  const pages = slide.default ?? [];
  if (pages.length === 0) return;

  const total = pages.length;
  let textBoxes = 0;
  onProgress?.({ phase: 'processing', current: 0, total, percent: 0, textBoxes });

  const container = document.createElement('div');
  container.className = CAPTURE_CLASS;
  container.setAttribute('aria-hidden', 'true');
  Object.assign(container.style, {
    position: 'fixed',
    left: '-99999px',
    top: '0',
    pointerEvents: 'none',
  });
  document.body.appendChild(container);

  // html-to-image copies computed styles into its clone, which would replay
  // intro keyframes from their hidden 0% frame. Fast-forward everything.
  const captureStyle = document.createElement('style');
  captureStyle.textContent = `.${CAPTURE_CLASS} *, .${CAPTURE_CLASS} *::before, .${CAPTURE_CLASS} *::after {
    animation-delay: -1s !important;
    animation-duration: 1ms !important;
    animation-iteration-count: 1 !important;
    animation-fill-mode: forwards !important;
    transition: none !important;
  }`;
  document.head.appendChild(captureStyle);

  const designVars = slide.design ? designToCssVars(slide.design) : null;
  const reactRoots: Root[] = [];
  const frames: HTMLElement[] = [];

  for (let i = 0; i < pages.length; i++) {
    const Page = pages[i];
    if (!Page) continue;
    const host = document.createElement('div');
    host.setAttribute('data-osd-canvas', '');
    host.style.width = `${SLIDE_W}px`;
    host.style.height = `${SLIDE_H}px`;
    host.style.overflow = 'hidden';
    host.style.background = '#fff';
    if (designVars) {
      for (const [k, v] of Object.entries(designVars)) host.style.setProperty(k, v);
    }
    container.appendChild(host);
    frames.push(host);
    const r = createRoot(host);
    r.render(
      createElement(SlidePageProvider, { index: i, total: pages.length }, createElement(Page)),
    );
    reactRoots.push(r);
  }
  await nextPaint();

  try {
    await waitForFonts();
    const deadline = performance.now() + ANIMATION_TIMEOUT_MS;
    while (performance.now() < deadline) {
      if (frames.every((frame) => isFrameAnimationSettled(frame))) break;
      await sleep(POLL_INTERVAL_MS);
    }
    await waitForDataWaitfor(container);

    const { toBlob } = await import('html-to-image');
    const slides: SlidePlan[] = [];

    for (let i = 0; i < frames.length; i++) {
      const frame = frames[i];
      freezeForCapture(frame);

      const blocks = collectTextBlocks(frame);
      // Hide only the glyphs: the element keeps its box, background and border,
      // so the captured background stays pixel-identical minus the text.
      const restore = blocks.map((b) => hideGlyphs(b.el));
      const blob = await toBlob(frame, {
        width: SLIDE_W,
        height: SLIDE_H,
        pixelRatio: CAPTURE_PIXEL_RATIO,
        backgroundColor: '#ffffff',
        cacheBust: true,
      });
      for (const undo of restore) undo();
      if (!blob) throw new Error(`failed to capture page ${i + 1}`);

      const boxes = blocks.flatMap(toLineBoxes);
      textBoxes += boxes.length;
      slides.push({ background: new Uint8Array(await blob.arrayBuffer()), boxes });

      onProgress?.({
        phase: 'processing',
        current: i + 1,
        total,
        percent: Math.min(95, ((i + 1) / total) * 95),
        textBoxes,
      });
    }

    onProgress?.({ phase: 'generating', current: total, total, percent: 98, textBoxes });
    const pptx = await buildPptx(slides);
    downloadBlob(
      new Blob([pptx as BlobPart], {
        type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      }),
      `${slideId}-editavel.pptx`,
    );
  } finally {
    onProgress?.({ phase: 'done', current: total, total, percent: 100, textBoxes });
    for (const r of reactRoots) r.unmount();
    container.remove();
    captureStyle.remove();
  }
}

/** A single PowerPoint text box: one rendered line, at its measured rect. */
type LineBox = {
  x: number;
  y: number;
  w: number;
  h: number;
  align: 'l' | 'ctr' | 'r' | 'just';
  lineHeightPx: number;
  line: TextLine;
};

type SlidePlan = { background: Uint8Array; boxes: LineBox[] };

/**
 * One box per rendered line, on that line's own measured rect. Alignment is
 * preserved by widening the box away from the anchoring edge: if a substituted
 * font renders wider, a centred line still grows around the same axis and a
 * right-aligned line still ends on the same x.
 */
function toLineBoxes(block: ReturnType<typeof collectTextBlocks>[number]): LineBox[] {
  return block.lines.map((line) => {
    const slack = Math.max(24, line.w * 0.25);
    const x =
      block.align === 'ctr' ? line.x - slack / 2 : block.align === 'r' ? line.x - slack : line.x;
    return {
      x,
      y: line.y,
      w: line.w + slack,
      h: line.h,
      align: block.align === 'just' ? 'l' : block.align,
      // Spacing comes from the measured line box, not the CSS line-height: the
      // box is positioned on that same rect, so the two have to agree.
      lineHeightPx: line.h,
      line,
    };
  });
}

/** Make text invisible without touching layout, background or borders. */
function hideGlyphs(el: HTMLElement): () => void {
  const prevColor = el.style.color;
  const prevFill = el.style.webkitTextFillColor;
  const prevPriority = el.style.getPropertyPriority('color');
  el.style.setProperty('color', 'transparent', 'important');
  el.style.setProperty('-webkit-text-fill-color', 'transparent', 'important');
  return () => {
    el.style.removeProperty('-webkit-text-fill-color');
    if (prevColor) el.style.setProperty('color', prevColor, prevPriority);
    else el.style.removeProperty('color');
    if (prevFill) el.style.webkitTextFillColor = prevFill;
  };
}

function freezeForCapture(root: HTMLElement): void {
  for (const el of root.querySelectorAll<HTMLElement>('*')) {
    const cs = getComputedStyle(el);
    for (const prop of ['opacity', 'transform', 'filter', 'clip-path'] as const) {
      el.style.setProperty(prop, cs.getPropertyValue(prop), 'important');
    }
    el.style.setProperty('animation', 'none', 'important');
    el.style.setProperty('transition', 'none', 'important');
  }
}

// --- OOXML -----------------------------------------------------------------
// ponytail: the package skeleton mirrors export-pptx.ts on purpose. That file
// hardcodes one full-bleed picture per slide and is the shipping path for image
// export; forking the few shared strings beats refactoring a working exporter.

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OD_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main';

async function buildPptx(slides: SlidePlan[]): Promise<Uint8Array> {
  const { zipSync, strToU8 } = await import('fflate');
  const n = slides.length;
  const files: Record<string, Uint8Array> = {};

  files['[Content_Types].xml'] = strToU8(contentTypesXml(n));
  files['_rels/.rels'] = strToU8(rootRelsXml());
  files['ppt/presentation.xml'] = strToU8(presentationXml(n));
  files['ppt/_rels/presentation.xml.rels'] = strToU8(presentationRelsXml(n));
  files['ppt/presProps.xml'] = strToU8(presPropsXml());
  files['ppt/theme/theme1.xml'] = strToU8(themeXml());
  files['ppt/slideMasters/slideMaster1.xml'] = strToU8(slideMasterXml());
  files['ppt/slideMasters/_rels/slideMaster1.xml.rels'] = strToU8(slideMasterRelsXml());
  files['ppt/slideLayouts/slideLayout1.xml'] = strToU8(slideLayoutXml());
  files['ppt/slideLayouts/_rels/slideLayout1.xml.rels'] = strToU8(slideLayoutRelsXml());

  for (let i = 0; i < n; i++) {
    const idx = i + 1;
    files[`ppt/slides/slide${idx}.xml`] = strToU8(slideXml(slides[i]));
    files[`ppt/slides/_rels/slide${idx}.xml.rels`] = strToU8(slideRelsXml(idx));
    files[`ppt/media/image${idx}.png`] = slides[i].background;
  }

  return zipSync(files);
}

function slideXml(plan: SlidePlan): string {
  const shapes = plan.boxes.map((box, i) => textBoxXml(box, i + 3)).join('');
  return `${XML_DECL}<p:sld xmlns:a="${A_NS}" xmlns:r="${OD_REL}" xmlns:p="${P_NS}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:pic><p:nvPicPr><p:cNvPr id="2" name="Background"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${EMU_W}" cy="${EMU_H}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>${shapes}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
}

function textBoxXml(box: LineBox, id: number): string {
  const off = `<a:off x="${emu(box.x)}" y="${emu(box.y)}"/><a:ext cx="${emu(Math.max(box.w, 1))}" cy="${emu(Math.max(box.h, 1))}"/>`;
  // Fixed line spacing in points reproduces the CSS line box exactly, instead of
  // letting PowerPoint derive it from the substituted font's metrics.
  const lnSpc = `<a:lnSpc><a:spcPts val="${Math.round(box.lineHeightPx * PT_PER_PX * 100)}"/></a:lnSpc>`;
  const runs = box.line.runs.map((run) => runXml(run, box.line.spcAdjustPx)).join('');
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Text ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm>${off}</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t" anchorCtr="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:pPr algn="${box.align}" marL="0" indent="0">${lnSpc}<a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft></a:pPr>${runs}</a:p></p:txBody></p:sp>`;
}

function runXml(run: TextLine['runs'][number], spcAdjustPx: number): string {
  const sz = Math.round(run.sizePx * PT_PER_PX * 100);
  // The weight lives in the resolved typeface when the family ships named
  // faces, so the bold flag comes from the resolution, not from the CSS weight.
  const bold = run.bold ? '1' : '0';
  const italic = run.italic ? '1' : '0';
  const underline = run.underline ? 'sng' : 'none';
  const strike = run.strike ? 'sngStrike' : 'noStrike';
  // OOXML letter spacing is in hundredths of a point. The calibration term
  // absorbs the width drift from weights PowerPoint can't express.
  const spc = Math.round((run.letterSpacingPx + spcAdjustPx) * PT_PER_PX * 100);
  const alpha =
    run.alpha < 0.999 ? `<a:alpha val="${Math.round(Math.max(0, run.alpha) * 100000)}"/>` : '';
  const font = escapeXml(run.typeface);
  return `<a:r><a:rPr lang="${documentLang()}" sz="${sz}" b="${bold}" i="${italic}" u="${underline}" strike="${strike}" spc="${spc}" kern="100" dirty="0"><a:solidFill><a:srgbClr val="${run.color}">${alpha}</a:srgbClr></a:solidFill><a:latin typeface="${font}"/><a:cs typeface="${font}"/></a:rPr><a:t>${escapeXml(run.text)}</a:t></a:r>`;
}

function emu(px: number): number {
  return Math.round(px * EMU_PER_PX);
}

/**
 * Language tag stamped on every run. PowerPoint uses it for spell checking, so
 * it should follow the document rather than being hardcoded to one locale.
 */
function documentLang(): string {
  const lang = document.documentElement.lang?.trim();
  return lang && /^[a-z]{2}(-[A-Za-z0-9]+)*$/.test(lang) ? lang : 'en-US';
}

// biome-ignore lint/suspicious/noControlCharactersInRegex: XML 1.0 forbids these outright.
const XML_CONTROL_CHARS = /[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g;

function escapeXml(value: string): string {
  return value
    .replace(XML_CONTROL_CHARS, '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function contentTypesXml(n: number): string {
  const slideOverrides = Array.from(
    { length: n },
    (_, i) =>
      `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`,
  ).join('');
  return `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>${slideOverrides}</Types>`;
}

function rootRelsXml(): string {
  return `${XML_DECL}<Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="${OD_REL}/officeDocument" Target="ppt/presentation.xml"/></Relationships>`;
}

function presentationXml(n: number): string {
  const sldIds = Array.from(
    { length: n },
    (_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 3}"/>`,
  ).join('');
  return `${XML_DECL}<p:presentation xmlns:a="${A_NS}" xmlns:r="${OD_REL}" xmlns:p="${P_NS}"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>${sldIds}</p:sldIdLst><p:sldSz cx="${EMU_W}" cy="${EMU_H}"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`;
}

function presentationRelsXml(n: number): string {
  const rels = [
    `<Relationship Id="rId1" Type="${OD_REL}/slideMaster" Target="slideMasters/slideMaster1.xml"/>`,
    `<Relationship Id="rId2" Type="${OD_REL}/presProps" Target="presProps.xml"/>`,
  ];
  for (let i = 0; i < n; i++) {
    rels.push(
      `<Relationship Id="rId${i + 3}" Type="${OD_REL}/slide" Target="slides/slide${i + 1}.xml"/>`,
    );
  }
  return `${XML_DECL}<Relationships xmlns="${REL_NS}">${rels.join('')}</Relationships>`;
}

function presPropsXml(): string {
  return `${XML_DECL}<p:presentationPr xmlns:a="${A_NS}" xmlns:r="${OD_REL}" xmlns:p="${P_NS}"/>`;
}

function slideMasterXml(): string {
  return `${XML_DECL}<p:sldMaster xmlns:a="${A_NS}" xmlns:r="${OD_REL}" xmlns:p="${P_NS}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>`;
}

function slideMasterRelsXml(): string {
  return `${XML_DECL}<Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="${OD_REL}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="${OD_REL}/theme" Target="../theme/theme1.xml"/></Relationships>`;
}

function slideLayoutXml(): string {
  return `${XML_DECL}<p:sldLayout xmlns:a="${A_NS}" xmlns:r="${OD_REL}" xmlns:p="${P_NS}" type="blank" preserve="1"><p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;
}

function slideLayoutRelsXml(): string {
  return `${XML_DECL}<Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="${OD_REL}/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`;
}

function slideRelsXml(idx: number): string {
  return `${XML_DECL}<Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="${OD_REL}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="${OD_REL}/image" Target="../media/image${idx}.png"/></Relationships>`;
}

function themeXml(): string {
  return `${XML_DECL}<a:theme xmlns:a="${A_NS}" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`;
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function nextPaint(): Promise<void> {
  return new Promise((resolve) => {
    let settled = false;
    const settle = () => {
      if (settled) return;
      settled = true;
      resolve();
    };
    requestAnimationFrame(settle);
    setTimeout(settle, 50);
  });
}

function downloadBlob(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.rel = 'noopener';
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 0);
}
