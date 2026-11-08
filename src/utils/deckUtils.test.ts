import test from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import {
  validatePresentationFile,
  formatFileSize,
  filterAndSortDecks,
  computeVirtualWindow,
  computeHotSlideIndices,
  emuToCanvasCoord,
  computeFittedStageSize,
  formatAspectRatioLabel,
} from './deckUtils';
import { buildValidPptxBuffer, SAMPLE_DECK_BLUEPRINTS } from './sampleDecks';
import { parsePptxArrayBuffer } from './pptxEngine';
import { PresentationMetadata } from '../types/presentation';

test('validatePresentationFile accepts valid .pptx with PK header and rejects legacy .ppt', () => {
  const pkHeader = new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0x14, 0x00]);
  const validRes = validatePresentationFile('Roadmap.pptx', 48120, pkHeader);
  assert.equal(validRes.valid, true);

  const legacyExt = validatePresentationFile('LegacyDeck.ppt', 25000);
  assert.equal(legacyExt.valid, false);
  assert.equal(legacyExt.errorCode, 'LEGACY_PPT');
  assert.ok(legacyExt.friendlyMessage?.includes('.pptx'));

  // OLE2 Compound Document magic bytes renamed to .pptx
  const oleHeader = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1]);
  const renamedLegacy = validatePresentationFile('FakeModern.pptx', 25000, oleHeader);
  assert.equal(renamedLegacy.valid, false);
  assert.equal(renamedLegacy.errorCode, 'LEGACY_PPT');

  const corruptHeader = new Uint8Array([0x00, 0x01, 0x02, 0x03]);
  const corruptRes = validatePresentationFile('Broken.pptx', 1024, corruptHeader);
  assert.equal(corruptRes.valid, false);
  assert.equal(corruptRes.errorCode, 'CORRUPTED_HEADER');
});

test('formatFileSize formats bytes, KB, and MB cleanly', () => {
  assert.equal(formatFileSize(512), '512 B');
  assert.equal(formatFileSize(2048), '2.0 KB');
  assert.equal(formatFileSize(1048576 * 2.5), '2.50 MB');
});

test('filterAndSortDecks filters by query and sorts by name, date, and size', () => {
  const sampleSlide = SAMPLE_DECK_BLUEPRINTS[0].slides[0];
  const decks: PresentationMetadata[] = [
    {
      id: '1',
      name: 'Beta Pitch',
      originalFileName: 'beta.pptx',
      fileSize: 5000,
      slideCount: 10,
      uploadedAt: 2000,
      aspectRatio: 16 / 9,
      firstSlide: sampleSlide,
      accentColor: '#B9A7FF',
    },
    {
      id: '2',
      name: 'Alpha Architecture',
      originalFileName: 'alpha.pptx',
      fileSize: 15000,
      slideCount: 20,
      uploadedAt: 3000,
      aspectRatio: 16 / 9,
      firstSlide: sampleSlide,
      accentColor: '#FFC8A8',
    },
    {
      id: '3',
      name: 'Gamma Review',
      originalFileName: 'gamma.pptx',
      fileSize: 1000,
      slideCount: 5,
      uploadedAt: 1000,
      aspectRatio: 16 / 9,
      firstSlide: sampleSlide,
      accentColor: '#A8EBD0',
    },
  ];

  const byNameAsc = filterAndSortDecks(decks, '', 'name', 'asc');
  assert.deepEqual(
    byNameAsc.map((d) => d.name),
    ['Alpha Architecture', 'Beta Pitch', 'Gamma Review']
  );

  const bySizeDesc = filterAndSortDecks(decks, '', 'size', 'desc');
  assert.deepEqual(
    bySizeDesc.map((d) => d.id),
    ['2', '1', '3']
  );

  const bySearch = filterAndSortDecks(decks, 'alpha', 'date', 'desc');
  assert.equal(bySearch.length, 1);
  assert.equal(bySearch[0].id, '2');
});

test('computeVirtualWindow and computeHotSlideIndices bound slide indices properly', () => {
  const win = computeVirtualWindow(500, 400, 100, 50, 2);
  assert.equal(win.startIndex, 3);
  assert.equal(win.endIndex, 11);
  assert.equal(win.offsetY, 300);
  assert.equal(win.totalHeight, 5000);

  const hot = computeHotSlideIndices(10, 25, 2);
  assert.deepEqual(hot, [0, 8, 9, 10, 11, 12]);
});

test('emuToCanvasCoord, formatAspectRatioLabel, and computeFittedStageSize preserve page aspect ratio', () => {
  assert.equal(emuToCanvasCoord(6096000, 12192000, 960), 480);
  assert.equal(emuToCanvasCoord(3429000, 6858000, 540), 270);

  assert.equal(formatAspectRatioLabel(16 / 9), '16:9');
  assert.equal(formatAspectRatioLabel(4 / 3), '4:3');
  assert.equal(formatAspectRatioLabel(16 / 10), '16:10');

  // Fit a 4:3 slide inside an 800x500 bezel -> height-constrained to 500px tall, 667px wide
  const fitted43 = computeFittedStageSize(800, 500, 4 / 3, 1);
  assert.equal(fitted43.height, 500);
  assert.equal(fitted43.width, 667);

  // Fit a 16:9 slide inside an 800x600 bezel -> width-constrained to 800px wide, 450px tall
  const fitted169 = computeFittedStageSize(800, 600, 16 / 9, 1);
  assert.equal(fitted169.width, 800);
  assert.equal(fitted169.height, 450);
});

test('buildValidPptxBuffer and parsePptxArrayBuffer round-trip slides progressively', async () => {
  const blueprint = SAMPLE_DECK_BLUEPRINTS[0];
  const buffer = await buildValidPptxBuffer(blueprint.slides);
  assert.ok(buffer.byteLength > 1000);

  let firstSlideFired = false;
  const progressCounts: number[] = [];

  const parsed = await parsePptxArrayBuffer(buffer, blueprint.fileName, {
    onFirstSlideReady: (info) => {
      firstSlideFired = true;
      assert.equal(info.slideCount, blueprint.slides.length);
      assert.equal(info.firstSlide.title, blueprint.slides[0].title);
    },
    onSlideParsed: (p) => {
      progressCounts.push(p.parsedCount);
    },
  });

  assert.equal(firstSlideFired, true);
  assert.equal(parsed.slides.length, blueprint.slides.length);
  assert.equal(progressCounts.length, blueprint.slides.length);
});

test('parsePptxArrayBuffer preserves 4:3 aspect ratio, inherits slideLayout placeholder positions, and transforms grouped shapes (<p:grpSp>)', async () => {
  const zip = new JSZip();

  // 4:3 aspect ratio: cx=9144000, cy=6858000 (aspectRatio = 1.3333, canvas = 960x720)
  zip.file(
    'ppt/presentation.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
    <p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
      <p:sldSz cx="9144000" cy="6858000"/>
    </p:presentation>`
  );

  zip.file(
    'ppt/slides/_rels/slide1.xml.rels',
    `<?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
    </Relationships>`
  );

  // slideLayout1 defines exact position for placeholder idx="1" (x=4572000 = 50% width = 480px, y=3429000 = 50% height = 360px on 960x720)
  zip.file(
    'ppt/slideLayouts/slideLayout1.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
    <p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
      <p:cSld>
        <p:spTree>
          <p:sp>
            <p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr><p:ph type="title" idx="1"/></p:nvPr></p:nvSpPr>
            <p:spPr>
              <a:xfrm><a:off x="4572000" y="3429000"/><a:ext cx="3657600" cy="1371600"/></a:xfrm>
            </p:spPr>
          </p:sp>
        </p:spTree>
      </p:cSld>
    </p:sldLayout>`
  );

  // slide1 has:
  // 1) A placeholder <p:sp> with <p:ph type="title" idx="1"/> and NO <a:xfrm> (must inherit x=480, y=360 from slideLayout1)
  // 2) A <p:grpSp> that maps child space [0..1000, 0..1000] to slide space [914400..2743200, 685800..2057400] (10%..30% of slide -> x=96px, y=72px)
  zip.file(
    'ppt/slides/slide1.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
    <p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
      <p:cSld>
        <p:spTree>
          <p:sp>
            <p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title" idx="1"/></p:nvPr></p:nvSpPr>
            <p:spPr/>
            <p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:r><a:t>Inherited Title Position</a:t></a:r></a:p></p:txBody>
          </p:sp>
          <p:grpSp>
            <p:grpSpPr>
              <a:xfrm>
                <a:off x="914400" y="685800"/>
                <a:ext cx="1828800" cy="1371600"/>
                <a:chOff x="0" y="0"/>
                <a:chExt cx="1000" cy="1000"/>
              </a:xfrm>
            </p:grpSpPr>
            <p:sp>
              <p:nvSpPr><p:cNvPr id="3" name="Grouped Box"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
              <p:spPr>
                <a:xfrm><a:off x="500" y="500"/><a:ext cx="500" cy="500"/></a:xfrm>
              </p:spPr>
              <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Grouped Text</a:t></a:r></a:p></p:txBody>
            </p:sp>
          </p:grpSp>
        </p:spTree>
      </p:cSld>
    </p:sld>`
  );

  const buffer = await zip.generateAsync({ type: 'arraybuffer' });
  const parsed = await parsePptxArrayBuffer(buffer, 'Test43Layout.pptx');

  assert.equal(parsed.aspectRatio, 1.3333);
  const slide = parsed.slides[0];
  assert.equal(slide.canvasWidth, 720);
  assert.equal(slide.canvasHeight, 540);

  // Element 0: inherited from slideLayout1 (50% of 720 = 360, 50% of 540 = 270)
  const titleEl = slide.elements[0];
  assert.equal(titleEl.x, 360);
  assert.equal(titleEl.y, 270);
  assert.equal(titleEl.verticalAlign, 'middle');

  // Element 1: inside <p:grpSp> at 50% of group (group starts at 10% + 0.5 * 20% = 20% of slide -> x=144, y=108)
  const groupedEl = slide.elements[1];
  assert.equal(groupedEl.x, 144);
  assert.equal(groupedEl.y, 108);
  assert.equal(groupedEl.w, 72);
  assert.equal(groupedEl.h, 54);
});
