import type { Page, SlideMeta } from '@open-slide/core';

// Stress fixture for tools/verify-pptx-editable.mjs. It deliberately gathers
// what usually breaks an HTML to OOXML conversion: gradients, radial glows,
// shadows, border radius, inline SVG, multi-line headings, mixed weights and
// every text alignment. Not presentation material, a test case.

const Cover: Page = () => (
  <div
    style={{
      width: '100%',
      height: '100%',
      background: 'linear-gradient(135deg, #0b0b14 0%, #1a1033 55%, #3b1d5e 100%)',
      color: '#f4f4f5',
      display: 'flex',
      flexDirection: 'column',
      justifyContent: 'center',
      padding: '0 160px',
      gap: 28,
      position: 'relative',
      overflow: 'hidden',
    }}
  >
    <div
      style={{
        position: 'absolute',
        right: -120,
        top: -120,
        width: 520,
        height: 520,
        borderRadius: '50%',
        background: 'radial-gradient(circle, rgba(139,92,246,0.55) 0%, rgba(139,92,246,0) 70%)',
      }}
    />
    <p style={{ fontSize: 26, letterSpacing: 8, color: '#a78bfa', margin: 0, fontWeight: 600 }}>
      QUARTERLY REPORT
    </p>
    <h1 style={{ fontSize: 116, fontWeight: 800, lineHeight: 1.04, margin: 0, maxWidth: 1350 }}>
      Steady growth across three revenue lines
    </h1>
    <p style={{ fontSize: 38, color: '#c4b5fd', margin: 0, lineHeight: 1.4, maxWidth: 1150 }}>
      A quarter of <strong style={{ color: '#ffffff' }}>accelerated expansion</strong>, with stable
      margin and <em>record retention</em> across the base.
    </p>
    <p
      style={{
        position: 'absolute',
        right: 160,
        bottom: 72,
        fontSize: 22,
        color: 'rgba(244,244,245,0.65)',
        margin: 0,
        textAlign: 'right',
      }}
    >
      Right aligned footer
    </p>
  </div>
);

const Numbers: Page = () => (
  <div
    style={{
      width: '100%',
      height: '100%',
      background: '#fbfaf7',
      color: '#16161a',
      display: 'flex',
      flexDirection: 'column',
      padding: '110px 150px',
      gap: 56,
    }}
  >
    <h2 style={{ fontSize: 62, fontWeight: 700, margin: 0, letterSpacing: -1 }}>
      Three numbers that sum up the period
    </h2>
    <div style={{ display: 'flex', gap: 40 }}>
      {[
        { n: '+38%', k: 'Recurring revenue', c: '#7c3aed' },
        { n: '94.2%', k: 'Net retention', c: '#0f766e' },
        { n: '$1.4M', k: 'Qualified pipeline', c: '#b45309' },
      ].map((item) => (
        <div
          key={item.k}
          style={{
            flex: 1,
            background: '#ffffff',
            borderRadius: 22,
            padding: '48px 40px',
            boxShadow: '0 18px 40px rgba(15,15,25,0.10)',
            border: '1px solid rgba(15,15,25,0.06)',
            display: 'flex',
            flexDirection: 'column',
            gap: 14,
          }}
        >
          <span style={{ fontSize: 78, fontWeight: 800, color: item.c, lineHeight: 1 }}>
            {item.n}
          </span>
          <span style={{ fontSize: 26, color: '#52525b', lineHeight: 1.35 }}>{item.k}</span>
        </div>
      ))}
    </div>
    <div style={{ display: 'flex', alignItems: 'center', gap: 22, marginTop: 'auto' }}>
      <svg width="46" height="46" viewBox="0 0 24 24" fill="none" aria-hidden="true">
        <title>Warning icon</title>
        <circle cx="12" cy="12" r="10" stroke="#b45309" strokeWidth="2" />
        <path d="M12 7v6" stroke="#b45309" strokeWidth="2" strokeLinecap="round" />
        <circle cx="12" cy="16.5" r="1.2" fill="#b45309" />
      </svg>
      <p style={{ fontSize: 24, color: '#52525b', margin: 0, maxWidth: 1300, lineHeight: 1.5 }}>
        Inline SVG next to a paragraph that wraps onto a second line, so the export has to place two
        lines from the same block independently.
      </p>
    </div>
  </div>
);

const Centred: Page = () => (
  <div
    style={{
      width: '100%',
      height: '100%',
      background: '#111827',
      color: '#f9fafb',
      display: 'flex',
      flexDirection: 'column',
      alignItems: 'center',
      justifyContent: 'center',
      gap: 34,
      padding: '0 220px',
      textAlign: 'center',
    }}
  >
    <p style={{ fontSize: 24, letterSpacing: 6, color: '#6ee7b7', margin: 0 }}>NEXT STEP</p>
    <h2 style={{ fontSize: 88, fontWeight: 700, lineHeight: 1.1, margin: 0 }}>
      Close the quarter with the whole operation looking at the same dashboard
    </h2>
    <p style={{ fontSize: 30, color: 'rgba(249,250,251,0.72)', margin: 0, lineHeight: 1.5 }}>
      One source of truth, updated daily, with no spreadsheet on the side.
    </p>
  </div>
);

export const meta: SlideMeta = { title: 'Verify PPTX export', createdAt: '2026-07-29' };
export default [Cover, Numbers, Centred] satisfies Page[];
