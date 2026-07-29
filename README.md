# OpenSlideV2

A fork of [open-slide](https://github.com/1weiho/open-slide) by [Yiwei Ho](https://github.com/1weiho), adding the three things a deck needs before it goes to a client: direct editing on the canvas, export to editable PowerPoint, and automatic resolution of the fonts and assets a deck depends on.

*Read this in [Português (Brasil)](README.pt-BR.md).*

## What it is

open-slide is a presentation framework written for coding agents. You describe the deck in natural language, the agent writes the React components, and the framework handles the fixed 1920x1080 canvas, scaling, navigation, hot reload and presenter mode. Every credit for that belongs to the original author.

This fork changes none of it. It adds three features on top, and everything else tracks upstream.

## Quick start

**Requirements**

- Node.js 18 or newer
- pnpm (the repository is a pnpm and Turbo monorepo)

**Install**

```bash
git clone https://github.com/bruno-org/OpenSlideV2.git
cd OpenSlideV2
pnpm install
pnpm setup:fonts
```

`pnpm setup:fonts` installs the project fonts into the current user profile, no administrator rights required. It matters because the editable PPTX export names fonts by family: without them installed, whoever opens the file gets PowerPoint's automatic substitution and the letterforms change.

**Publish this fork into an existing open-slide workspace**

```bash
pnpm install:workspace ../MyWorkspace
```

The command builds, packs with `npm pack` and reinstalls. It does not copy loose files into `node_modules`, and that is deliberate: Vite caches transforms for files under `node_modules` and would keep serving the previous version even after a restart.

## What this fork adds

### Drag and resize on the canvas

Upstream's inspector edits properties of the selected element (text, typography, colour, image). Here it also moves and resizes:

- dragging the body of an element repositions it;
- the four corner handles resize it, with the opposite edge anchored.

The gesture ends up as an ordinary `set-style` operation (`translate`, `width`, `height`), the same one the panel already used, so Save, undo and redo work with nothing extra and the result is written back into the `.tsx`. Small adjustments no longer require a round trip through the agent.

### Export to editable PPTX

The `Download` menu gains **Export as editable PPTX**. The text comes out as real, editable PowerPoint text, and the slide still matches what is on screen.

The order of operations is what guarantees that:

1. each page is captured as a background image with the promoted text made transparent. Gradients, shadows, SVG, `clip-path` and filters all still paint exactly where they were, because it is a capture of the page itself;
2. native text boxes go on top, one per rendered line, at the rect the browser measured, with wrapping disabled. A substituted font can change the letterforms, but it can never reflow the deck;
3. anything that cannot be restated faithfully as PowerPoint text (gradient-filled text, text shadows, rotated or clipped text) is not promoted: it stays painted in the background. You lose the ability to edit that fragment, never the appearance.

Three details separate "close" from "identical":

- **weight**: PowerPoint has no weight axis, only bold on or off. A weight of 500 or 800 is pointed at the matching named face ("Geist SemiBold", "Geist ExtraBold"). Those names come from what each font file declares internally, not from the file name: `Geist-UltraBlack.ttf` declares itself as "Geist ExtraBold";
- **kerning**: PowerPoint reads only the legacy `kern` table, so fonts that carry kerning in GPOS (most modern ones) render unkerned and every line comes out wider. Calibration measures with kerning disabled precisely to reproduce that, and turns the difference into letter spacing;
- **safety net**: a family that cannot be guaranteed on the target machine is mapped to a face every Windows install has (Segoe UI, Georgia, Consolas), chosen here rather than by PowerPoint, with the line width calibrated against it.

### Dependency preflight

Every `open-slide dev` starts by checking whether the machine has what the decks ask for, and resolving what it does not:

- reads the decks in `slides/` and the themes in `themes/` and collects the font families in use, including the one the runtime applies by default;
- checks which are already installed on the system;
- fetches what is missing: first from the matching npm package, then from the [google/fonts](https://github.com/google/fonts) repository, which is the stable public source of installable TTFs (the download endpoint on the website answers with HTML, and the CSS API serves woff2, which no operating system can install);
- installs for the current user on Windows, macOS and Linux;
- assets: a deck can ship `slides/<id>/assets.manifest.json` mapping file names to URLs. Anything missing on disk is downloaded.

None of it can take the server down. A network failure becomes a line in the report. Each deck opened this way leaves the machine slightly better equipped than it found it.

To run it on its own: `open-slide preflight`, or `--no-install` to only list what is missing.

## Verification

Both features ship with an executable check. Run them from `packages/core`, with the workspace dev server up:

```bash
pnpm dev:demo                               # in another terminal, from the repo root
node tools/verify-drag-resize.mjs           # drags, resizes, checks the .tsx
node tools/verify-pptx-editable.mjs         # exports, opens in PowerPoint, compares
```

Both target fixtures that live in this repository (`apps/demo/slides/verify-*`), so they run on any clone with no setup. The second one exports the same deck in both formats, uses the image PPTX as the reference, opens the editable PPTX in actual PowerPoint (through COM, on Windows), exports the slides and compares them. It measures geometry, which is the real proof of fidelity, and keeps a diverging-pixel count as a guard against gross regression.

Results on the stress fixture, which carries gradients, a radial glow, shadows, inline SVG, multi-line headings, mixed weights and every text alignment:

| measurement | result |
| --- | --- |
| horizontal offset | up to 4.9px |
| vertical offset | up to 14.7px |
| text width | within 0.3% |
| text height | within 0.5% |
| diverging pixels | 2.1% to 3.9% |

The remaining pixels are edge antialiasing and PowerPoint's own rasterisation. The vertical offset concentrates on large text whose CSS line-height is tighter than the font's natural line: PowerPoint will not compress a line below that minimum, so the text sits a few pixels higher. On a 1080px canvas that is around 1%.

Both scripts launch their own browser with a temporary profile and never touch a browser you already have open. Step 3 of the PPTX check needs PowerPoint installed, so it only runs on Windows; everything else works anywhere.

## Known limits

- Text with a gradient fill, a shadow, rotation or clipping does not become editable text. It stays in the background image, visually identical.
- The editable PPTX depends on the fonts being installed on the machine that opens the file. The preflight handles that for OpenSlideV2 users; for someone who only receives the `.pptx`, the system-font safety net applies.
- Shapes and images do not become native objects: they stay in the background. Only text is promoted.

## Attribution

[open-slide](https://github.com/1weiho/open-slide) was created by [Yiwei Ho](https://github.com/1weiho) and is the entire foundation of this project. This repository is an independent fork that adds features on top of that work.

If OpenSlideV2 is useful to you, go to the source first: [star the original repository](https://github.com/1weiho/open-slide) and consider [supporting its development](https://ko-fi.com/D1D11YPUP1).

## License

MIT, for both. Free to use, copy, modify and distribute, commercially or not, keeping the copyright notice. See [LICENSE](LICENSE).

## Contributing

Issues and pull requests are welcome. Changes to the framework core are usually better sent upstream, where everyone benefits from them.
