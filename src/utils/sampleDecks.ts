import JSZip from 'jszip';
import { ParsedSlide } from '../types/presentation';

function escapeXml(str: string): string {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/**
 * Builds a valid OpenXML `.pptx` ArrayBuffer from a list of `ParsedSlide` objects
 * so the resulting file is a true ZIP-packaged PowerPoint `.pptx` archive that
 * can be parsed by standard OpenXML parsers AND downloaded by the user.
 */
export async function buildValidPptxBuffer(
  slides: ParsedSlide[],
  aspectRatio = 16 / 9
): Promise<ArrayBuffer> {
  const zip = new JSZip();

  const cx = 12192000;
  const cy = Math.round(cx / aspectRatio);

  // 1. [Content_Types].xml
  const slideOverrides = slides
    .map(
      (_, i) =>
        `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`
    )
    .join('\n  ');

  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  ${slideOverrides}
</Types>`
  );

  // 2. _rels/.rels
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`
  );

  // 3. ppt/presentation.xml
  const sldIdLst = slides
    .map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 1}"/>`)
    .join('\n    ');

  zip.file(
    'ppt/presentation.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldIdLst>
    ${sldIdLst}
  </p:sldIdLst>
  <p:sldSz cx="${cx}" cy="${cy}"/>
</p:presentation>`
  );

  // 4. ppt/_rels/presentation.xml.rels
  const presRels = slides
    .map(
      (_, i) =>
        `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`
    )
    .join('\n  ');

  zip.file(
    'ppt/_rels/presentation.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${presRels}
</Relationships>`
  );

  // 5. Each ppt/slides/slideN.xml in standard OpenXML format
  const canvasWidth = 960;
  const canvasHeight = Math.round(canvasWidth / aspectRatio);

  slides.forEach((slide, idx) => {
    const bgHex = (slide.background.color || '#FAF8FF').replace('#', '');
    const spXmlList: string[] = [];

    slide.elements.forEach((el, elIdx) => {
      const offX = Math.round((el.x / canvasWidth) * cx);
      const offY = Math.round((el.y / canvasHeight) * cy);
      const extCx = Math.round((el.w / canvasWidth) * cx);
      const extCy = Math.round((el.h / canvasHeight) * cy);

      if (el.type === 'table' && el.table) {
        const allRows = [el.table.headers, ...el.table.rows];
        const trXml = allRows
          .map(
            (row) =>
              `<a:tr h="370840">${row
                .map(
                  (cell) =>
                    `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${escapeXml(cell)}</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>`
                )
                .join('')}</a:tr>`
          )
          .join('');
        spXmlList.push(`
      <p:graphicFrame>
        <p:xfrm><a:off x="${offX}" y="${offY}"/><a:ext cx="${extCx}" cy="${extCy}"/></p:xfrm>
        <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>${trXml}</a:tbl></a:graphicData></a:graphic>
      </p:graphicFrame>`);
        return;
      }

      const fillXml = el.fillColor
        ? `<a:solidFill><a:srgbClr val="${el.fillColor.replace('#', '')}"/></a:solidFill>`
        : `<a:noFill/>`;

      const prst =
        el.shapeKind === 'ellipse'
          ? 'ellipse'
          : el.shapeKind === 'roundRect' || el.shapeKind === 'card' || el.shapeKind === 'pill'
            ? 'roundRect'
            : 'rect';

      const pXml = (el.paragraphs || [])
        .map((p) => {
          const algn = p.align === 'center' ? 'ctr' : p.align === 'right' ? 'r' : 'l';
          const bu = p.bullet ? '<a:buChar char="•"/>' : '';
          const runsXml = p.runs
            .map((r) => {
              const sz = Math.round((r.fontSize || 16) * 100);
              const b = r.bold ? ' b="1"' : '';
              const i = r.italic ? ' i="1"' : '';
              const clr = (r.color || '#2A2244').replace('#', '');
              return `<a:r><a:rPr sz="${sz}"${b}${i}><a:solidFill><a:srgbClr val="${clr}"/></a:solidFill></a:rPr><a:t>${escapeXml(r.text)}</a:t></a:r>`;
            })
            .join('');
          return `<a:p><a:pPr algn="${algn}">${bu}</a:pPr>${runsXml}</a:p>`;
        })
        .join('');

      const phAttr = elIdx === 0 ? '<p:ph type="title"/>' : '';

      spXmlList.push(`
      <p:sp>
        <p:nvSpPr><p:cNvPr id="${elIdx + 2}" name="Shape ${elIdx + 1}"/><p:cNvSpPr/><p:nvPr>${phAttr}</p:nvPr></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="${offX}" y="${offY}"/><a:ext cx="${extCx}" cy="${extCy}"/></a:xfrm>
          <a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>
          ${fillXml}
        </p:spPr>
        ${pXml ? `<p:txBody><a:bodyPr/><a:lstStyle/>${pXml}</p:txBody>` : ''}
      </p:sp>`);
    });

    zip.file(
      `ppt/slides/slide${idx + 1}.xml`,
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg><p:bgPr><a:solidFill><a:srgbClr val="${bgHex}"/></a:solidFill></p:bgPr></p:bg>
    <p:spTree>
      ${spXmlList.join('\n')}
    </p:spTree>
  </p:cSld>
</p:sld>`
    );
  });

  // Also store the rich vector manifest for chart/card fidelity
  zip.file(
    'ppt/slideclay-manifest.json',
    JSON.stringify({ aspectRatio, slides })
  );

  return zip.generateAsync({ type: 'arraybuffer' });
}

export interface SampleDeckBlueprint {
  fileName: string;
  accentColor: string;
  uploadedAtOffsetMs: number;
  slides: ParsedSlide[];
}

export const SAMPLE_DECK_BLUEPRINTS: SampleDeckBlueprint[] = [
  {
    fileName: 'Claymorphic Design System 2026.pptx',
    accentColor: '#B9A7FF',
    uploadedAtOffsetMs: 1000 * 60 * 45, // 45 mins ago
    slides: [
      {
        index: 0,
        slideNumber: 1,
        title: 'Claymorphic Design System 2026',
        subtitle: 'Tactile, Puffy 3D Interfaces Built for Speed & Clarity',
        background: { color: '#F3EFFF', accentBlobColor: '#B9A7FF' },
        elements: [
          {
            id: 's1-card',
            type: 'shape',
            x: 56,
            y: 52,
            w: 848,
            h: 436,
            shapeKind: 'roundRect',
            fillColor: '#FFFFFF',
            borderRadius: 32,
          },
          {
            id: 's1-badge',
            type: 'shape',
            x: 96,
            y: 94,
            w: 220,
            h: 38,
            shapeKind: 'pill',
            fillColor: '#EBE4FF',
            borderRadius: 999,
            paragraphs: [
              {
                align: 'center',
                runs: [{ text: 'DESIGN LANGUAGE v3.2', bold: true, fontSize: 12, color: '#6448D6' }],
              },
            ],
          },
          {
            id: 's1-title',
            type: 'text',
            x: 96,
            y: 150,
            w: 540,
            h: 130,
            paragraphs: [
              {
                align: 'left',
                runs: [
                  {
                    text: 'Soft 3D Clay Surfaces That Invite Touch',
                    bold: true,
                    fontSize: 38,
                    color: '#251D3D',
                  },
                ],
              },
            ],
          },
          {
            id: 's1-sub',
            type: 'text',
            x: 96,
            y: 294,
            w: 480,
            h: 80,
            paragraphs: [
              {
                align: 'left',
                runs: [
                  {
                    text: 'Combining outer ambient drop shadows with dual top-left & bottom-right inset highlights for instant visual hierarchy.',
                    fontSize: 16,
                    color: '#5E5380',
                  },
                ],
              },
            ],
          },
          {
            id: 's1-orb1',
            type: 'shape',
            x: 640,
            y: 120,
            w: 210,
            h: 210,
            shapeKind: 'ellipse',
            fillColor: '#B9A7FF',
          },
          {
            id: 's1-orb2',
            type: 'shape',
            x: 720,
            y: 250,
            w: 145,
            h: 145,
            shapeKind: 'ellipse',
            fillColor: '#FFC8A8',
          },
          {
            id: 's1-orb3',
            type: 'shape',
            x: 610,
            y: 285,
            w: 125,
            h: 125,
            shapeKind: 'ellipse',
            fillColor: '#A8EBD0',
          },
        ],
      },
      {
        index: 1,
        slideNumber: 2,
        title: 'The Four Pillars of Claymorphism',
        subtitle: 'Why tactile interfaces boost comprehension and retention',
        background: { color: '#FAF8FF', accentBlobColor: '#FFC8A8' },
        elements: [
          {
            id: 's2-title',
            type: 'text',
            x: 64,
            y: 44,
            w: 832,
            h: 60,
            paragraphs: [
              {
                align: 'left',
                runs: [{ text: 'The Four Pillars of Claymorphism', bold: true, fontSize: 32, color: '#251D3D' }],
              },
            ],
          },
          {
            id: 's2-c1',
            type: 'shape',
            x: 64,
            y: 124,
            w: 400,
            h: 172,
            shapeKind: 'roundRect',
            fillColor: '#EFEAFF',
            borderRadius: 24,
            paragraphs: [
              {
                align: 'left',
                runs: [{ text: '01 · Pastel Luminance', bold: true, fontSize: 20, color: '#39286E' }],
              },
              {
                align: 'left',
                runs: [
                  {
                    text: 'Warm off-white canvases (#F1EEFF) paired with lavender, peach, mint, and sky accents prevent eye strain.',
                    fontSize: 14,
                    color: '#54487A',
                  },
                ],
              },
            ],
          },
          {
            id: 's2-c2',
            type: 'shape',
            x: 496,
            y: 124,
            w: 400,
            h: 172,
            shapeKind: 'roundRect',
            fillColor: '#FFF0E6',
            borderRadius: 24,
            paragraphs: [
              {
                align: 'left',
                runs: [{ text: '02 · Dual Inset Depth', bold: true, fontSize: 20, color: '#6E3216' }],
              },
              {
                align: 'left',
                runs: [
                  {
                    text: 'A top-left specular inner highlight and bottom-right inner shade create unmistakable 3D volume.',
                    fontSize: 14,
                    color: '#7A4930',
                  },
                ],
              },
            ],
          },
          {
            id: 's2-c3',
            type: 'shape',
            x: 64,
            y: 318,
            w: 400,
            h: 172,
            shapeKind: 'roundRect',
            fillColor: '#E6F9F1',
            borderRadius: 24,
            paragraphs: [
              {
                align: 'left',
                runs: [{ text: '03 · Super-Rounded Radii', bold: true, fontSize: 20, color: '#14523B' }],
              },
              {
                align: 'left',
                runs: [
                  {
                    text: '24px–40px card corners and 999px pill controls eliminate harsh visual edges across the workspace.',
                    fontSize: 14,
                    color: '#286B52',
                  },
                ],
              },
            ],
          },
          {
            id: 's2-c4',
            type: 'shape',
            x: 496,
            y: 318,
            w: 400,
            h: 172,
            shapeKind: 'roundRect',
            fillColor: '#E8F3FF',
            borderRadius: 24,
            paragraphs: [
              {
                align: 'left',
                runs: [{ text: '04 · Spring Physics', bold: true, fontSize: 20, color: '#184780' }],
              },
              {
                align: 'left',
                runs: [
                  {
                    text: 'Buttons lift on hover and squish to 0.96 scale on press, mimicking real sculpted polymer clay.',
                    fontSize: 14,
                    color: '#305C91',
                  },
                ],
              },
            ],
          },
        ],
      },
      {
        index: 2,
        slideNumber: 3,
        title: 'Color Palette & Contrast Ratios',
        subtitle: 'WCAG AA compliant typography on pastel surfaces',
        background: { color: '#F6F3FF', accentBlobColor: '#A8EBD0' },
        elements: [
          {
            id: 's3-title',
            type: 'text',
            x: 64,
            y: 44,
            w: 832,
            h: 56,
            paragraphs: [
              {
                align: 'left',
                runs: [{ text: 'Curated Clay Palette & Contrast Benchmarks', bold: true, fontSize: 30, color: '#251D3D' }],
              },
            ],
          },
          {
            id: 's3-tbl',
            type: 'table',
            x: 64,
            y: 126,
            w: 832,
            h: 356,
            table: {
              headers: ['Token Name', 'Hex Code', 'Primary Role', 'Text Contrast Ratio'],
              rows: [
                ['Clay Lavender', '#B9A7FF', 'Primary Actions & Focus States', '7.4 : 1 (AAA)'],
                ['Soft Peach', '#FFC8A8', 'Upload Highlights & Warm Callouts', '8.1 : 1 (AAA)'],
                ['Puffy Mint', '#A8EBD0', 'Completed States & Fast Cache Indicators', '8.6 : 1 (AAA)'],
                ['Sky Pastel', '#A8D4FF', 'Viewer Stage Accents & Zoom Controls', '7.9 : 1 (AAA)'],
                ['Butter Yellow', '#FFE9A8', 'Keyboard Shortcut Badges & Tips', '9.2 : 1 (AAA)'],
              ],
              accentColor: '#B9A7FF',
            },
          },
        ],
      },
      {
        index: 3,
        slideNumber: 4,
        title: 'Rendering Speed & Frame Budget',
        subtitle: '60fps navigation powered by Web Workers & IndexedDB',
        background: { color: '#FAF8FF', accentBlobColor: '#A8D4FF' },
        elements: [
          {
            id: 's4-title',
            type: 'text',
            x: 64,
            y: 44,
            w: 832,
            h: 56,
            paragraphs: [
              {
                align: 'left',
                runs: [{ text: 'First-Slide Latency Benchmark (Lower is Better)', bold: true, fontSize: 30, color: '#251D3D' }],
              },
            ],
          },
          {
            id: 's4-chart',
            type: 'chart',
            x: 64,
            y: 120,
            w: 832,
            h: 364,
            chartTitle: 'Time to First Interactive Slide (ms) — 25-Slide Deck',
            chartBars: [
              { label: 'SlideClay Worker + IDB Cache', value: 85, color: '#A8EBD0' },
              { label: 'SlideClay Cold Upload', value: 340, color: '#B9A7FF' },
              { label: 'Standard Main-Thread Parser', value: 1420, color: '#FFC8A8' },
              { label: 'Legacy Cloud Iframe Viewer', value: 2890, color: '#FFB4C8' },
            ],
          },
        ],
      },
      {
        index: 4,
        slideNumber: 5,
        title: 'Typography & Scannable Hierarchy',
        subtitle: 'Fredoka Display paired with Nunito Body and JetBrains Mono',
        background: { color: '#F3EFFF', accentBlobColor: '#FFE9A8' },
        elements: [
          {
            id: 's5-title',
            type: 'text',
            x: 64,
            y: 44,
            w: 832,
            h: 56,
            paragraphs: [
              {
                align: 'left',
                runs: [{ text: 'Rounded Typography With Tabular Precision', bold: true, fontSize: 30, color: '#251D3D' }],
              },
            ],
          },
          {
            id: 's5-box1',
            type: 'shape',
            x: 64,
            y: 124,
            w: 832,
            h: 112,
            shapeKind: 'roundRect',
            fillColor: '#FFFFFF',
            borderRadius: 24,
            paragraphs: [
              {
                align: 'left',
                runs: [{ text: 'Fredoka Bold — Display & Deck Titles (24px–40px)', bold: true, fontSize: 22, color: '#251D3D' }],
              },
              {
                align: 'left',
                runs: [{ text: 'Warm, friendly rounded terminals that echo the 32px border radius of our clay cards.', fontSize: 15, color: '#5E5380' }],
              },
            ],
          },
          {
            id: 's5-box2',
            type: 'shape',
            x: 64,
            y: 256,
            w: 832,
            h: 112,
            shapeKind: 'roundRect',
            fillColor: '#FFFFFF',
            borderRadius: 24,
            paragraphs: [
              {
                align: 'left',
                runs: [{ text: 'Nunito SemiBold — UI Controls & Slide Copy (15px–18px)', bold: true, fontSize: 20, color: '#251D3D' }],
              },
              {
                align: 'left',
                runs: [{ text: 'Balanced x-height and open apertures keep long presentation notes effortless to read.', fontSize: 15, color: '#5E5380' }],
              },
            ],
          },
          {
            id: 's5-box3',
            type: 'shape',
            x: 64,
            y: 388,
            w: 832,
            h: 100,
            shapeKind: 'roundRect',
            fillColor: '#EBE4FF',
            borderRadius: 24,
            paragraphs: [
              {
                align: 'left',
                runs: [{ text: 'JetBrains Mono — Slide Counters (05 / 24) & File Sizes (2.48 MB)', bold: true, fontSize: 18, color: '#452CB8' }],
              },
              {
                align: 'left',
                runs: [{ text: 'Tabular numerals prevent horizontal jitter during rapid 60fps keyboard slide scrubbing.', fontSize: 14, color: '#5E5380' }],
              },
            ],
          },
        ],
      },
      {
        index: 5,
        slideNumber: 6,
        title: 'Dark Clay Theme Architecture',
        subtitle: 'Deep muted purples with preserved specular highlights',
        background: { color: '#1E1931', accentBlobColor: '#7C63E6' },
        elements: [
          {
            id: 's6-card',
            type: 'shape',
            x: 64,
            y: 56,
            w: 832,
            h: 428,
            shapeKind: 'roundRect',
            fillColor: '#282142',
            borderRadius: 32,
          },
          {
            id: 's6-title',
            type: 'text',
            x: 104,
            y: 96,
            w: 752,
            h: 60,
            paragraphs: [
              {
                align: 'left',
                runs: [{ text: 'Nighttime Viewing Without Losing Puffy Depth', bold: true, fontSize: 32, color: '#F4F1FF' }],
              },
            ],
          },
          {
            id: 's6-bullets',
            type: 'text',
            x: 104,
            y: 178,
            w: 740,
            h: 260,
            paragraphs: [
              {
                align: 'left',
                bullet: true,
                runs: [{ text: 'Deep plum canvas (#171326) replaces harsh pure black to preserve outer shadow contrast.', fontSize: 18, color: '#D7CFF2' }],
              },
              {
                align: 'left',
                bullet: true,
                runs: [{ text: 'Elevated clay surfaces (#251F3D) carry soft violet top-left inner highlights.', fontSize: 18, color: '#D7CFF2' }],
              },
              {
                align: 'left',
                bullet: true,
                runs: [{ text: 'Pastel badges maintain high luminance for instant status recognition in dim rooms.', fontSize: 18, color: '#A8EBD0' }],
              },
            ],
          },
        ],
      },
      {
        index: 6,
        slideNumber: 7,
        title: 'Keyboard-First Presenter Workflow',
        subtitle: 'Zero-mouse control from library to fullscreen stage',
        background: { color: '#FAF8FF', accentBlobColor: '#B9A7FF' },
        elements: [
          {
            id: 's7-title',
            type: 'text',
            x: 64,
            y: 44,
            w: 832,
            h: 56,
            paragraphs: [
              {
                align: 'left',
                runs: [{ text: 'Essential Presenter Shortcuts', bold: true, fontSize: 30, color: '#251D3D' }],
              },
            ],
          },
          {
            id: 's7-tbl',
            type: 'table',
            x: 64,
            y: 120,
            w: 832,
            h: 360,
            table: {
              headers: ['Shortcut Key', 'Action in Viewer', 'Performance Guarantee'],
              rows: [
                ['← / → Arrow Keys', 'Previous / Next Slide', 'Preloaded adjacent slides (< 5ms swap)'],
                ['Home / End', 'Jump to First / Last Slide', 'Virtualized sidebar auto-scrolls smoothly'],
                ['F Key', 'Toggle Distraction-Free Fullscreen', 'Scales 16:9 stage to 100% viewport'],
                ['Ctrl/Cmd + K', 'Open Quick Switcher Palette', 'Fuzzy search across all uploaded decks'],
                ['+ / - / 0', 'Zoom In, Zoom Out, Fit to Screen', 'Crisp vector scaling at up to 200%'],
              ],
              accentColor: '#A8D4FF',
            },
          },
        ],
      },
      {
        index: 7,
        slideNumber: 8,
        title: 'Ready to Present With SlideClay',
        subtitle: 'Drop any .pptx file to experience instant client-side rendering',
        background: { color: '#EFEAFF', accentBlobColor: '#A8EBD0' },
        elements: [
          {
            id: 's8-card',
            type: 'shape',
            x: 100,
            y: 80,
            w: 760,
            h: 380,
            shapeKind: 'roundRect',
            fillColor: '#FFFFFF',
            borderRadius: 36,
          },
          {
            id: 's8-title',
            type: 'text',
            x: 140,
            y: 150,
            w: 680,
            h: 100,
            paragraphs: [
              {
                align: 'center',
                runs: [{ text: '100% Private. 100% In-Browser.', bold: true, fontSize: 38, color: '#251D3D' }],
              },
            ],
          },
          {
            id: 's8-sub',
            type: 'text',
            x: 160,
            y: 255,
            w: 640,
            h: 120,
            paragraphs: [
              {
                align: 'center',
                runs: [
                  {
                    text: 'Your presentations never leave your device. IndexedDB keeps your library and thumbnails ready even after browser restarts.',
                    fontSize: 18,
                    color: '#5E5380',
                  },
                ],
              },
            ],
          },
        ],
      },
    ],
  },
  {
    fileName: 'Q4 Product Velocity & Roadmap.pptx',
    accentColor: '#FFC8A8',
    uploadedAtOffsetMs: 1000 * 60 * 60 * 5, // 5 hours ago
    slides: [
      {
        index: 0,
        slideNumber: 1,
        title: 'Q4 Product Velocity & Roadmap',
        subtitle: 'Quarterly Engineering & Customer Adoption Review',
        background: { color: '#FFF6F0', accentBlobColor: '#FFC8A8' },
        elements: [
          {
            id: 'q1-card',
            type: 'shape',
            x: 56,
            y: 52,
            w: 848,
            h: 436,
            shapeKind: 'roundRect',
            fillColor: '#FFFFFF',
            borderRadius: 32,
          },
          {
            id: 'q1-pill',
            type: 'shape',
            x: 96,
            y: 96,
            w: 210,
            h: 38,
            shapeKind: 'pill',
            fillColor: '#FFE5D4',
            borderRadius: 999,
            paragraphs: [
              {
                align: 'center',
                runs: [{ text: 'Q4 EXECUTIVE BRIEF', bold: true, fontSize: 12, color: '#B84A14' }],
              },
            ],
          },
          {
            id: 'q1-title',
            type: 'text',
            x: 96,
            y: 158,
            w: 620,
            h: 120,
            paragraphs: [
              {
                align: 'left',
                runs: [{ text: 'Q4 Product Velocity & 2027 Strategic Roadmap', bold: true, fontSize: 36, color: '#2B1B12' }],
              },
            ],
          },
          {
            id: 'q1-sub',
            type: 'text',
            x: 96,
            y: 292,
            w: 560,
            h: 90,
            paragraphs: [
              {
                align: 'left',
                runs: [
                  {
                    text: 'Highlighting 99.98% client-side reliability, 3.4x faster deck indexing, and enterprise offline workspace adoption.',
                    fontSize: 17,
                    color: '#6E4F3D',
                  },
                ],
              },
            ],
          },
          {
            id: 'q1-orb',
            type: 'shape',
            x: 700,
            y: 160,
            w: 160,
            h: 160,
            shapeKind: 'ellipse',
            fillColor: '#FFC8A8',
          },
        ],
      },
      {
        index: 1,
        slideNumber: 2,
        title: 'Quarterly KPI Highlights',
        subtitle: 'Core performance metrics across active workspaces',
        background: { color: '#FFFBF8', accentBlobColor: '#A8EBD0' },
        elements: [
          {
            id: 'q2-title',
            type: 'text',
            x: 64,
            y: 44,
            w: 832,
            h: 56,
            paragraphs: [
              {
                align: 'left',
                runs: [{ text: 'Quarterly Product KPI Summary', bold: true, fontSize: 30, color: '#2B1B12' }],
              },
            ],
          },
          {
            id: 'q2-k1',
            type: 'shape',
            x: 64,
            y: 130,
            w: 260,
            h: 340,
            shapeKind: 'roundRect',
            fillColor: '#E6F9F1',
            borderRadius: 28,
            paragraphs: [
              {
                align: 'left',
                runs: [{ text: '94 ms', bold: true, fontSize: 42, color: '#14523B' }],
              },
              {
                align: 'left',
                runs: [{ text: 'Median Library Boot', bold: true, fontSize: 18, color: '#1E6E50' }],
              },
              {
                align: 'left',
                runs: [{ text: '3.2x faster than our 300ms target thanks to split IndexedDB metadata stores.', fontSize: 14, color: '#2D7A5E' }],
              },
            ],
          },
          {
            id: 'q2-k2',
            type: 'shape',
            x: 350,
            y: 130,
            w: 260,
            h: 340,
            shapeKind: 'roundRect',
            fillColor: '#EFEAFF',
            borderRadius: 28,
            paragraphs: [
              {
                align: 'left',
                runs: [{ text: '60 fps', bold: true, fontSize: 42, color: '#39286E' }],
              },
              {
                align: 'left',
                runs: [{ text: 'Slide Transitions', bold: true, fontSize: 18, color: '#4B368C' }],
              },
              {
                align: 'left',
                runs: [{ text: 'Adjacent N-1 and N+1 slides stay hot in memory for zero-jank scrubbing.', fontSize: 14, color: '#5D489C' }],
              },
            ],
          },
          {
            id: 'q2-k3',
            type: 'shape',
            x: 636,
            y: 130,
            w: 260,
            h: 340,
            shapeKind: 'roundRect',
            fillColor: '#FFF0E6',
            borderRadius: 28,
            paragraphs: [
              {
                align: 'left',
                runs: [{ text: '100%', bold: true, fontSize: 42, color: '#6E3216' }],
              },
              {
                align: 'left',
                runs: [{ text: 'Offline Capable', bold: true, fontSize: 18, color: '#8A411E' }],
              },
              {
                align: 'left',
                runs: [{ text: 'Zero server roundtrips required for parsing, searching, or presenting.', fontSize: 14, color: '#8F5234' }],
              },
            ],
          },
        ],
      },
      {
        index: 2,
        slideNumber: 3,
        title: 'Engineering Delivery Velocity',
        subtitle: 'Story points shipped per sprint across core modules',
        background: { color: '#FFFBF8', accentBlobColor: '#B9A7FF' },
        elements: [
          {
            id: 'q3-title',
            type: 'text',
            x: 64,
            y: 44,
            w: 832,
            h: 56,
            paragraphs: [
              {
                align: 'left',
                runs: [{ text: 'Engineering Throughput by Subsystem', bold: true, fontSize: 30, color: '#2B1B12' }],
              },
            ],
          },
          {
            id: 'q3-chart',
            type: 'chart',
            x: 64,
            y: 120,
            w: 832,
            h: 360,
            chartTitle: 'Completed Feature Milestones (% Readiness)',
            chartBars: [
              { label: 'Web Worker XML Parser', value: 100, color: '#A8EBD0' },
              { label: 'IndexedDB Instant Cache', value: 100, color: '#B9A7FF' },
              { label: 'Virtualized Slide Sidebar', value: 98, color: '#A8D4FF' },
              { label: 'Claymorphic Dual Theme', value: 100, color: '#FFC8A8' },
            ],
          },
        ],
      },
      {
        index: 3,
        slideNumber: 4,
        title: '2027 Release Schedule',
        subtitle: 'Upcoming milestones for Q1–Q2',
        background: { color: '#FFF6F0', accentBlobColor: '#A8D4FF' },
        elements: [
          {
            id: 'q4-title',
            type: 'text',
            x: 64,
            y: 44,
            w: 832,
            h: 56,
            paragraphs: [
              {
                align: 'left',
                runs: [{ text: 'Upcoming Milestone Delivery Matrix', bold: true, fontSize: 30, color: '#2B1B12' }],
              },
            ],
          },
          {
            id: 'q4-tbl',
            type: 'table',
            x: 64,
            y: 120,
            w: 832,
            h: 360,
            table: {
              headers: ['Milestone', 'Target Window', 'Lead Team', 'Status'],
              rows: [
                ['Multi-Deck Batch Uploader', 'Shipped Q4', 'Core Runtime', 'Complete'],
                ['Virtualized 500-Slide Sidebar', 'Shipped Q4', 'UI Performance', 'Complete'],
                ['Command Palette (Ctrl/Cmd+K)', 'Shipped Q4', 'Design Systems', 'Complete'],
                ['Speaker Notes Split Screen', 'Q1 Next Year', 'Presenter Experience', 'In Progress'],
              ],
              accentColor: '#FFC8A8',
            },
          },
        ],
      },
      {
        index: 4,
        slideNumber: 5,
        title: 'Customer Feedback & NPS',
        subtitle: 'What design and product teams love most',
        background: { color: '#FFFBF8', accentBlobColor: '#FFE9A8' },
        elements: [
          {
            id: 'q5-title',
            type: 'text',
            x: 64,
            y: 44,
            w: 832,
            h: 56,
            paragraphs: [
              {
                align: 'left',
                runs: [{ text: 'What Teams Say About SlideClay', bold: true, fontSize: 30, color: '#2B1B12' }],
              },
            ],
          },
          {
            id: 'q5-quote1',
            type: 'shape',
            x: 64,
            y: 126,
            w: 832,
            h: 160,
            shapeKind: 'roundRect',
            fillColor: '#FFFFFF',
            borderRadius: 28,
            paragraphs: [
              {
                align: 'left',
                runs: [
                  {
                    text: '“Opening a 20-slide product review deck takes less than half a second, and the clay bezel makes every slide look like a studio showcase.”',
                    italic: true,
                    fontSize: 18,
                    color: '#2A2244',
                  },
                ],
              },
              {
                align: 'left',
                runs: [{ text: '— Elena Rostova, VP of Product Design at Kinetix Labs', bold: true, fontSize: 14, color: '#7C63E6' }],
              },
            ],
          },
          {
            id: 'q5-quote2',
            type: 'shape',
            x: 64,
            y: 312,
            w: 832,
            h: 160,
            shapeKind: 'roundRect',
            fillColor: '#FFFFFF',
            borderRadius: 28,
            paragraphs: [
              {
                align: 'left',
                runs: [
                  {
                    text: '“Ctrl+K quick switching between our pitch deck and technical appendix during live Q&A completely transformed our customer demos.”',
                    italic: true,
                    fontSize: 18,
                    color: '#2A2244',
                  },
                ],
              },
              {
                align: 'left',
                runs: [{ text: '— Marcus Vance, Principal Solutions Architect at LatticeCloud', bold: true, fontSize: 14, color: '#E8824B' }],
              },
            ],
          },
        ],
      },
    ],
  },
  {
    fileName: 'Neural Architecture & Edge AI.pptx',
    accentColor: '#A8EBD0',
    uploadedAtOffsetMs: 1000 * 60 * 60 * 26, // 1 day ago
    slides: [
      {
        index: 0,
        slideNumber: 1,
        title: 'Neural Architecture & Edge AI',
        subtitle: 'Low-Power Inference Pipelines on Consumer Hardware',
        background: { color: '#EEFCF6', accentBlobColor: '#A8EBD0' },
        elements: [
          {
            id: 'n1-card',
            type: 'shape',
            x: 56,
            y: 52,
            w: 848,
            h: 436,
            shapeKind: 'roundRect',
            fillColor: '#FFFFFF',
            borderRadius: 32,
          },
          {
            id: 'n1-pill',
            type: 'shape',
            x: 96,
            y: 96,
            w: 230,
            h: 38,
            shapeKind: 'pill',
            fillColor: '#D4F7E8',
            borderRadius: 999,
            paragraphs: [
              {
                align: 'center',
                runs: [{ text: 'SYSTEMS ARCHITECTURE', bold: true, fontSize: 12, color: '#146B4B' }],
              },
            ],
          },
          {
            id: 'n1-title',
            type: 'text',
            x: 96,
            y: 158,
            w: 620,
            h: 120,
            paragraphs: [
              {
                align: 'left',
                runs: [{ text: 'Neural Architecture & Sub-10ms Edge Inference', bold: true, fontSize: 36, color: '#133629' }],
              },
            ],
          },
          {
            id: 'n1-sub',
            type: 'text',
            x: 96,
            y: 294,
            w: 580,
            h: 90,
            paragraphs: [
              {
                align: 'left',
                runs: [
                  {
                    text: 'Quantized INT4 attention blocks, zero-copy WebAssembly buffers, and deterministic memory scheduling.',
                    fontSize: 17,
                    color: '#346653',
                  },
                ],
              },
            ],
          },
          {
            id: 'n1-orb',
            type: 'shape',
            x: 705,
            y: 165,
            w: 155,
            h: 155,
            shapeKind: 'ellipse',
            fillColor: '#A8EBD0',
          },
        ],
      },
      {
        index: 1,
        slideNumber: 2,
        title: 'Quantization Throughput Comparison',
        subtitle: 'Tokens per second across precision formats',
        background: { color: '#F6FEFB', accentBlobColor: '#A8D4FF' },
        elements: [
          {
            id: 'n2-title',
            type: 'text',
            x: 64,
            y: 44,
            w: 832,
            h: 56,
            paragraphs: [
              {
                align: 'left',
                runs: [{ text: 'On-Device Inference Speed (Tokens / Sec)', bold: true, fontSize: 30, color: '#133629' }],
              },
            ],
          },
          {
            id: 'n2-chart',
            type: 'chart',
            x: 64,
            y: 120,
            w: 832,
            h: 360,
            chartTitle: 'Sustained Generation Throughput (Tokens/s, Higher is Better)',
            chartBars: [
              { label: 'INT4 AWQ Group-128', value: 142, color: '#A8EBD0' },
              { label: 'INT8 Dynamic Quant', value: 88, color: '#A8D4FF' },
              { label: 'FP16 Native Tensor', value: 46, color: '#B9A7FF' },
              { label: 'FP32 Baseline CPU', value: 14, color: '#FFC8A8' },
            ],
          },
        ],
      },
      {
        index: 2,
        slideNumber: 3,
        title: 'Memory Footprint Breakdown',
        subtitle: 'VRAM allocation for 3B parameter edge model',
        background: { color: '#F6FEFB', accentBlobColor: '#B9A7FF' },
        elements: [
          {
            id: 'n3-title',
            type: 'text',
            x: 64,
            y: 44,
            w: 832,
            h: 56,
            paragraphs: [
              {
                align: 'left',
                runs: [{ text: 'Memory Footprint by Precision Tier', bold: true, fontSize: 30, color: '#133629' }],
              },
            ],
          },
          {
            id: 'n3-tbl',
            type: 'table',
            x: 64,
            y: 120,
            w: 832,
            h: 360,
            table: {
              headers: ['Precision Format', 'Weights VRAM', 'KV Cache (8K ctx)', 'Perplexity Delta'],
              rows: [
                ['FP16 Uncompressed', '6.20 GB', '1.40 GB', 'Baseline (0.00)'],
                ['INT8 Symmetric', '3.15 GB', '0.70 GB', '+0.03 (Negligible)'],
                ['INT4 Group-128', '1.68 GB', '0.38 GB', '+0.11 (Production Optimal)'],
                ['NF3 Extreme', '1.25 GB', '0.38 GB', '+0.48 (Degraded)'],
              ],
              accentColor: '#A8EBD0',
            },
          },
        ],
      },
      {
        index: 3,
        slideNumber: 4,
        title: 'Key Takeaways & Deployment Checklist',
        subtitle: 'Production readiness gates for edge rollout',
        background: { color: '#EEFCF6', accentBlobColor: '#FFE9A8' },
        elements: [
          {
            id: 'n4-card',
            type: 'shape',
            x: 64,
            y: 56,
            w: 832,
            h: 428,
            shapeKind: 'roundRect',
            fillColor: '#FFFFFF',
            borderRadius: 32,
          },
          {
            id: 'n4-title',
            type: 'text',
            x: 104,
            y: 96,
            w: 752,
            h: 60,
            paragraphs: [
              {
                align: 'left',
                runs: [{ text: 'Edge Rollout Verification Gates', bold: true, fontSize: 32, color: '#133629' }],
              },
            ],
          },
          {
            id: 'n4-bullets',
            type: 'text',
            x: 104,
            y: 176,
            w: 740,
            h: 260,
            paragraphs: [
              {
                align: 'left',
                bullet: true,
                runs: [{ text: 'Run all tokenization and tensor unpacking inside dedicated Web Workers.', fontSize: 18, color: '#275242' }],
              },
              {
                align: 'left',
                bullet: true,
                runs: [{ text: 'Persist compiled model shards in IndexedDB to achieve < 200ms warm boots.', fontSize: 18, color: '#275242' }],
              },
              {
                align: 'left',
                bullet: true,
                runs: [{ text: 'Evict inactive context buffers outside the active sliding window.', fontSize: 18, color: '#275242' }],
              },
            ],
          },
        ],
      },
    ],
  },
];
