/* ============================================================================
 * SLIDECLAY — NATIVE `pptx-preview` RENDERING & MODEL NORMALIZATION SERVICE
 * ============================================================================
 * Uses the `pptx-preview` library's internal PPTX reader + HtmlRender engine
 * to render individual slides (SlideMaster + SlideLayout + Slide layers,
 * 50+ preset geometries, custom SVG paths, SmartArt diagrams, merged-cell
 * tables, ECharts, and embedded media) at their exact 1:1 OpenXML point
 * dimensions (`pptx.width` x `pptx.height`).
 *
 * Also normalizes known `pptx-preview` quirks before and after `load()` so that:
 * 1. Single-item XML lists in `theme1.xml` (`<a:lnStyleLst>`) or `[Content_Types].xml`
 *    never trigger `e.map is not a function` inside `pptx-preview`'s XML reader.
 * 2. Latin fonts (`<a:latin typeface="...">`) are preserved without negative
 *    letter-spacing compression (`-0.04em` -> `normal`).
 * 3. Text boxes (`txBox="1"`) and body/title placeholders default to left
 *    alignment (`"l"`) rather than centering when `<a:pPr algn>` is omitted.
 * 4. Body placeholders inherit font size and color from `slideMaster.textStyles.bodyStyle`.
 * 5. Slide background images (`blipFill`) include valid CSS pixel/percentage units.
 * ============================================================================ */

import JSZip from 'jszip';
import { init } from 'pptx-preview';

interface PptxInternalRun {
  isBr?: boolean;
  props?: {
    size?: number;
    bold?: boolean;
    italic?: boolean;
    underline?: string;
    color?: unknown;
    typeface?: string;
  };
  text?: string;
}

interface PptxInternalParagraph {
  props?: {
    align?: string;
    marginLeft?: number;
    indent?: number;
    level?: string | number;
    lineHeight?: number;
    buAutoNum?: string;
    buChar?: string;
  };
  inheritProps?: {
    align?: string;
    marginLeft?: number;
    indent?: number;
    level?: string | number;
    lineHeight?: number;
  };
  inheritRProps?: {
    size?: number;
    bold?: boolean;
    italic?: boolean;
    color?: unknown;
    typeface?: string;
  };
  rows?: PptxInternalRun[];
}

interface PptxInternalNode {
  isTextBox?: boolean;
  type?: string;
  idx?: string;
  source?: Record<string, unknown>;
  nodes?: PptxInternalNode[];
  textBody?: {
    props?: Record<string, unknown>;
    inheritProps?: Record<string, unknown>;
    paragraphs?: PptxInternalParagraph[];
  };
  tr?: Array<{
    td?: Array<{
      paragraphs?: PptxInternalParagraph[];
    }>;
  }>;
}

interface PptxInternalSlideMaster {
  nodes?: PptxInternalNode[];
  background?: { type: string };
  textStyles?: {
    titleStyle?: Record<string, { props?: Record<string, unknown>; defRPr?: Record<string, unknown> }>;
    bodyStyle?: Record<string, { props?: Record<string, unknown>; defRPr?: Record<string, unknown> }>;
    otherStyle?: Record<string, { props?: Record<string, unknown>; defRPr?: Record<string, unknown> }>;
  };
  defaultTextStyle?: Record<string, { props?: Record<string, unknown>; defRPr?: Record<string, unknown> }>;
  tableStyles?: Record<string, unknown>;
  getColorThemeName?: (t: string) => string;
  getNodeByType?: (t: string) => PptxInternalNode | undefined;
  getNodeByIdx?: (idx: string) => PptxInternalNode | undefined;
  getNodeInheritAttrsByType?: (t: string, path: string[]) => unknown;
  getNodeInheritAttrsByIdx?: (idx: string, path: string[]) => unknown;
}

interface PptxInternalSlideLayout {
  nodes?: PptxInternalNode[];
  background?: { type: string };
  slideMaster?: PptxInternalSlideMaster;
  getColorThemeName?: (t: string) => string;
  getNodeByType?: (t: string) => PptxInternalNode | undefined;
  getNodeByIdx?: (idx: string) => PptxInternalNode | undefined;
  getNodeInheritAttrsByType?: (t: string, path: string[]) => unknown;
  getNodeInheritAttrsByIdx?: (idx: string, path: string[]) => unknown;
}

interface PptxInternalSlide {
  nodes?: PptxInternalNode[];
  background?: { type: string };
  slideLayout?: PptxInternalSlideLayout;
  slideMaster?: PptxInternalSlideMaster;
}

interface PptxInternalModel {
  width: number;
  height: number;
  slides: PptxInternalSlide[];
  slideLayouts: PptxInternalSlideLayout[];
  slideMaster: PptxInternalSlideMaster[];
}

interface PptxPreviewerInternal {
  pptx: PptxInternalModel;
  htmlRender: {
    wrapper: HTMLElement;
    options: {
      viewPort: { width: number; height: number };
      mode?: 'list' | 'slide';
    };
    _calcScaleAndRenderPort: () => void;
    renderSlide: (slideIndex: number) => void;
  };
  load: (buffer: ArrayBuffer) => Promise<PptxInternalModel>;
}

export interface CachedDeckPreviewer {
  previewer: PptxPreviewerInternal;
  width: number;
  height: number;
  slideCount: number;
}

const previewerCacheByKey = new Map<string, Promise<CachedDeckPreviewer>>();
const previewerCacheByBuffer = new WeakMap<ArrayBuffer, Promise<CachedDeckPreviewer>>();

const FALLBACK_THEME_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:srgbClr val="1E293B"/></a:dk1>
      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="334155"/></a:dk2>
      <a:lt2><a:srgbClr val="F8FAFC"/></a:lt2>
      <a:accent1><a:srgbClr val="6366F1"/></a:accent1>
      <a:accent2><a:srgbClr val="F59E0B"/></a:accent2>
      <a:accent3><a:srgbClr val="10B981"/></a:accent3>
      <a:accent4><a:srgbClr val="3B82F6"/></a:accent4>
      <a:accent5><a:srgbClr val="8B5CF6"/></a:accent5>
      <a:accent6><a:srgbClr val="EC4899"/></a:accent6>
      <a:hlink><a:srgbClr val="2563EB"/></a:hlink>
      <a:folHlink><a:srgbClr val="7C3AED"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="25400"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
      <a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>`;

const FALLBACK_MASTER_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg><p:bgPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></p:bgPr></p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:txStyles>
    <p:titleStyle><a:lvl1pPr algn="l"><a:defRPr sz="3200" b="1"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:defRPr></a:lvl1pPr></p:titleStyle>
    <p:bodyStyle><a:lvl1pPr algn="l"><a:defRPr sz="1800"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:defRPr></a:lvl1pPr></p:bodyStyle>
    <p:otherStyle><a:lvl1pPr algn="l"><a:defRPr sz="1400"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:defRPr></a:lvl1pPr></p:otherStyle>
  </p:txStyles>
</p:sldMaster>`;

const FALLBACK_LAYOUT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="obj">
  <p:cSld name="Title and Content">
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
    </p:spTree>
  </p:cSld>
</p:sldLayout>`;

/**
 * Ensures any `.pptx` archive is 100% compatible with `pptx-preview`'s internal
 * XML loader (`_loadContentTypes`, `Theme._parseLineStyleLst`, etc.).
 * If the archive already conforms to all expectations, returns `rawBuffer` with zero re-zipping.
 */
async function ensurePptxArchiveCompatibility(
  rawBuffer: ArrayBuffer
): Promise<ArrayBuffer> {
  try {
    const zip = await JSZip.loadAsync(rawBuffer);
    let modified = false;

    const allPaths = Object.keys(zip.files);
    const themePaths = allPaths.filter((p) => /^ppt\/theme\/theme\d+\.xml$/i.test(p));
    const masterPaths = allPaths.filter((p) =>
      /^ppt\/slideMasters\/slideMaster\d+\.xml$/i.test(p)
    );
    const layoutPaths = allPaths.filter((p) =>
      /^ppt\/slideLayouts\/slideLayout\d+\.xml$/i.test(p)
    );
    const slidePaths = allPaths.filter((p) => /^ppt\/slides\/slide\d+\.xml$/i.test(p));

    // 0. Ensure `ppt/presentation.xml` contains `<p:defaultTextStyle>` so `SlideMaster._parseDefaultTextStyle`
    // in `pptx-preview` never fails with `Cannot convert undefined or null to object` before `_loadNodes()`.
    const presFile = zip.file('ppt/presentation.xml');
    if (presFile) {
      const presXml = await presFile.async('string');
      if (!presXml.includes('<p:defaultTextStyle')) {
        const patchedPresXml = presXml.replace(
          /<\/p:presentation>/,
          `  <p:defaultTextStyle><a:defPPr/></p:defaultTextStyle>\n</p:presentation>`
        );
        zip.file('ppt/presentation.xml', patchedPresXml);
        modified = true;
      }
    }

    // 1. Ensure `ppt/theme/theme*.xml` has >= 2 `<a:ln>` tags in `<a:lnStyleLst>`
    // so `pptx-preview`'s `H.prototype._parseLineStyleLst` never fails with `e.map is not a function`.
    if (themePaths.length === 0) {
      zip.file('ppt/theme/theme1.xml', FALLBACK_THEME_XML);
      themePaths.push('ppt/theme/theme1.xml');
      modified = true;
    } else {
      for (const tp of themePaths) {
        const tf = zip.file(tp);
        if (!tf) continue;
        const xml = await tf.async('string');
        const lnListMatch = xml.match(/<a:lnStyleLst>([\s\S]*?)<\/a:lnStyleLst>/);
        if (lnListMatch) {
          const inner = lnListMatch[1];
          const lnMatches = inner.match(/<a:ln\b[\s\S]*?<\/a:ln>|<a:ln\b[^>]*\/>/g) || [];
          if (lnMatches.length === 1) {
            const patched = xml.replace(
              /<a:lnStyleLst>[\s\S]*?<\/a:lnStyleLst>/,
              `<a:lnStyleLst>${lnMatches[0]}${lnMatches[0]}</a:lnStyleLst>`
            );
            zip.file(tp, patched);
            modified = true;
          }
        }
      }
    }

    // 2. Ensure at least one SlideMaster and SlideLayout exist, and every Master/Layout has a `.rels` file and valid `<p:txStyles>`
    if (masterPaths.length === 0) {
      zip.file('ppt/slideMasters/slideMaster1.xml', FALLBACK_MASTER_XML);
      masterPaths.push('ppt/slideMasters/slideMaster1.xml');
      modified = true;
    }
    for (const mp of masterPaths) {
      const masterFile = zip.file(mp);
      if (masterFile) {
        let masterXml = await masterFile.async('string');
        let masterPatched = false;
        if (!masterXml.includes('<p:clrMap')) {
          masterXml = masterXml.replace(
            /<\/p:cSld>/,
            `</p:cSld>\n  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>`
          );
          masterPatched = true;
        }
        if (!masterXml.includes('<p:txStyles')) {
          masterXml = masterXml.replace(
            /<\/p:sldMaster>/,
            `  <p:txStyles><p:titleStyle><a:lvl1pPr/></p:titleStyle><p:bodyStyle><a:lvl1pPr/></p:bodyStyle><p:otherStyle><a:lvl1pPr/></p:otherStyle></p:txStyles>\n</p:sldMaster>`
          );
          masterPatched = true;
        } else {
          if (!masterXml.includes('<p:otherStyle')) {
            masterXml = masterXml.replace(
              /<\/p:txStyles>/,
              `  <p:otherStyle><a:lvl1pPr/></p:otherStyle>\n  </p:txStyles>`
            );
            masterPatched = true;
          }
          if (!masterXml.includes('<p:titleStyle')) {
            masterXml = masterXml.replace(
              /<\/p:txStyles>/,
              `  <p:titleStyle><a:lvl1pPr/></p:titleStyle>\n  </p:txStyles>`
            );
            masterPatched = true;
          }
          if (!masterXml.includes('<p:bodyStyle')) {
            masterXml = masterXml.replace(
              /<\/p:txStyles>/,
              `  <p:bodyStyle><a:lvl1pPr/></p:bodyStyle>\n  </p:txStyles>`
            );
            masterPatched = true;
          }
        }
        if (masterPatched) {
          zip.file(mp, masterXml);
          modified = true;
        }
      }

      const masterFileName = mp.split('/').pop()!;
      const relsPath = `ppt/slideMasters/_rels/${masterFileName}.rels`;
      if (!zip.file(relsPath)) {
        zip.file(
          relsPath,
          `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`
        );
        modified = true;
      }
    }

    if (layoutPaths.length === 0) {
      zip.file('ppt/slideLayouts/slideLayout1.xml', FALLBACK_LAYOUT_XML);
      layoutPaths.push('ppt/slideLayouts/slideLayout1.xml');
      modified = true;
    }
    for (const lp of layoutPaths) {
      const layoutFileName = lp.split('/').pop()!;
      const relsPath = `ppt/slideLayouts/_rels/${layoutFileName}.rels`;
      if (!zip.file(relsPath)) {
        zip.file(
          relsPath,
          `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`
        );
        modified = true;
      }
    }

    // 3. Ensure every slide has a `.rels` linking to a `slideLayout`
    for (const sp of slidePaths) {
      const slideFileName = sp.split('/').pop()!;
      const relsPath = `ppt/slides/_rels/${slideFileName}.rels`;
      const existingRelsFile = zip.file(relsPath);
      if (!existingRelsFile) {
        zip.file(
          relsPath,
          `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdLayout1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`
        );
        modified = true;
      }
    }

    // 4. Ensure `[Content_Types].xml` lists all theme, master, layout, and slide parts
    const ctFile = zip.file('[Content_Types].xml');
    if (ctFile) {
      let ctXml = await ctFile.async('string');
      const missingOverrides: string[] = [];

      for (const tp of themePaths) {
        if (!ctXml.includes(`/${tp}`)) {
          missingOverrides.push(
            `<Override PartName="/${tp}" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>`
          );
        }
      }
      for (const mp of masterPaths) {
        if (!ctXml.includes(`/${mp}`)) {
          missingOverrides.push(
            `<Override PartName="/${mp}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>`
          );
        }
      }
      for (const lp of layoutPaths) {
        if (!ctXml.includes(`/${lp}`)) {
          missingOverrides.push(
            `<Override PartName="/${lp}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>`
          );
        }
      }
      for (const sp of slidePaths) {
        if (!ctXml.includes(`/${sp}`)) {
          missingOverrides.push(
            `<Override PartName="/${sp}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`
          );
        }
      }

      if (missingOverrides.length > 0) {
        ctXml = ctXml.replace(
          /<\/Types>/i,
          `  ${missingOverrides.join('\n  ')}\n</Types>`
        );
        zip.file('[Content_Types].xml', ctXml);
        modified = true;
      }
    }

    if (!modified) {
      return rawBuffer;
    }

    return await zip.generateAsync({ type: 'arraybuffer' });
  } catch {
    return rawBuffer;
  }
}

/**
 * Extracts `<a:latin typeface="...">` from a raw parsed XML node if present.
 */
function extractLatinTypeface(rawRunSource: unknown): string | undefined {
  if (!rawRunSource || typeof rawRunSource !== 'object') return undefined;
  const rPr = (rawRunSource as Record<string, unknown>)['a:rPr'] as
    | Record<string, unknown>
    | undefined;
  const latin = (rPr?.['a:latin'] ||
    (rawRunSource as Record<string, unknown>)['a:latin']) as
    | { attrs?: { typeface?: string } }
    | undefined;
  const tf = latin?.attrs?.typeface;
  if (!tf || tf.startsWith('+')) return undefined;
  return tf;
}

/**
 * Recursively normalizes a node tree inside `pptx-preview` so text alignment,
 * bodyStyle inheritance, and font families match PowerPoint's OpenXML defaults.
 */
function normalizeNodeTree(
  nodes: PptxInternalNode[] | undefined,
  master: PptxInternalSlideMaster | undefined,
  isMasterOrLayout = false
): void {
  if (!Array.isArray(nodes)) return;

  for (const node of nodes) {
    // Non-placeholder shapes on SlideMaster or SlideLayout are always drawn in PowerPoint
    if (isMasterOrLayout && !node.type && !node.idx) {
      (node as PptxInternalNode & { userDrawn?: boolean }).userDrawn = true;
    }

    if (Array.isArray(node.nodes) && node.nodes.length > 0) {
      normalizeNodeTree(node.nodes, master, isMasterOrLayout);
    }

    // Normalize table cell paragraphs: default to left align when unspecified
    if (Array.isArray(node.tr)) {
      for (const row of node.tr) {
        if (!Array.isArray(row.td)) continue;
        for (const cell of row.td) {
          if (!Array.isArray(cell.paragraphs)) continue;
          for (const para of cell.paragraphs) {
            para.inheritProps = para.inheritProps || {};
            if (!para.props?.align && !para.inheritProps.align) {
              para.inheritProps.align = 'l';
            }
            if (Array.isArray(para.rows)) {
              for (const r of para.rows) {
                r.props = r.props || {};
                if (!r.props.typeface) {
                  r.props.typeface = 'Calibri, Nunito, sans-serif';
                }
              }
            }
          }
        }
      }
    }

    const textBody = node.textBody;
    if (!textBody || !Array.isArray(textBody.paragraphs)) continue;

    // Determine if this shape is a text box or a non-centered placeholder
    const phType = node.type || '';
    const isCenteredPlaceholder = phType === 'ctrTitle';
    const isPlaceholder = Boolean(node.type || node.idx);
    const shouldDefaultLeftAlign =
      !isCenteredPlaceholder && (Boolean(node.isTextBox) || isPlaceholder);

    const isBodyPlaceholder =
      phType === 'body' ||
      phType === 'obj' ||
      (!phType && Boolean(node.idx) && node.idx !== '0');

    // Inspect raw XML paragraphs for `<a:latin typeface>` if available
    const rawTxBody = node.source?.['p:txBody'] as Record<string, unknown> | undefined;
    const rawParasRaw = rawTxBody?.['a:p'];
    const rawParas = Array.isArray(rawParasRaw)
      ? rawParasRaw
      : rawParasRaw
        ? [rawParasRaw]
        : [];

    for (let pIdx = 0; pIdx < textBody.paragraphs.length; pIdx++) {
      const para = textBody.paragraphs[pIdx];
      para.props = para.props || {};
      para.inheritProps = para.inheritProps || {};
      para.inheritRProps = para.inheritRProps || {};

      const lvlNum = para.props.level ? Number(para.props.level) + 1 : 1;
      const lvlKey = `lvl${Math.max(1, Math.min(9, lvlNum))}`;

      // Inherit from `slideMaster.textStyles.bodyStyle` for body/content placeholders
      if (isBodyPlaceholder && master?.textStyles?.bodyStyle?.[lvlKey]) {
        const bodyLvl = master.textStyles.bodyStyle[lvlKey];
        if (bodyLvl.props) {
          para.inheritProps = {
            ...bodyLvl.props,
            ...para.inheritProps,
          };
        }
        if (bodyLvl.defRPr) {
          para.inheritRProps = {
            ...bodyLvl.defRPr,
            ...para.inheritRProps,
          };
        }
      }

      // Fix OpenXML default alignment: text boxes and placeholders default to left ('l')
      if (
        shouldDefaultLeftAlign &&
        !para.props.align &&
        !para.inheritProps.align
      ) {
        para.inheritProps.align = 'l';
      }

      if (!para.inheritRProps.typeface) {
        para.inheritRProps.typeface = 'Calibri, Nunito, sans-serif';
      }

      const rawPara = rawParas[pIdx] as Record<string, unknown> | undefined;
      const rawRunsRaw = rawPara?.['a:r'];
      const rawRuns = Array.isArray(rawRunsRaw)
        ? rawRunsRaw
        : rawRunsRaw
          ? [rawRunsRaw]
          : [];

      if (Array.isArray(para.rows)) {
        let runCursor = 0;
        for (const r of para.rows) {
          if (r.isBr) continue;
          r.props = r.props || {};
          const rawRun = rawRuns[runCursor++];
          const latinFont = extractLatinTypeface(rawRun);
          if (latinFont && !r.props.typeface) {
            r.props.typeface = `${latinFont}, Calibri, Nunito, sans-serif`;
          } else if (!r.props.typeface) {
            r.props.typeface =
              para.inheritRProps.typeface || 'Calibri, Nunito, sans-serif';
          }
        }
      }
    }
  }
}

/**
 * Normalizes the parsed `PPTX` model in-place after `previewer.load()` completes.
 */
function normalizeParsedPptxModel(pptx: PptxInternalModel): void {
  const defaultMaster: PptxInternalSlideMaster = pptx.slideMaster?.[0] || {
    nodes: [],
    background: { type: 'none' },
    textStyles: { titleStyle: {}, bodyStyle: {}, otherStyle: {} },
    defaultTextStyle: {},
    tableStyles: {},
    getColorThemeName: (t: string) => t,
    getNodeByType: () => undefined,
    getNodeByIdx: () => undefined,
    getNodeInheritAttrsByType: () => undefined,
    getNodeInheritAttrsByIdx: () => undefined,
  };

  const defaultLayout: PptxInternalSlideLayout = pptx.slideLayouts?.[0] || {
    nodes: [],
    background: { type: 'none' },
    slideMaster: defaultMaster,
    getColorThemeName: (t: string) => t,
    getNodeByType: () => undefined,
    getNodeByIdx: () => undefined,
    getNodeInheritAttrsByType: () => undefined,
    getNodeInheritAttrsByIdx: () => undefined,
  };

  if (Array.isArray(pptx.slideMaster)) {
    for (const master of pptx.slideMaster) {
      master.nodes = master.nodes || [];
      master.background = master.background || { type: 'none' };
      normalizeNodeTree(master.nodes, master, true);
    }
  }

  if (Array.isArray(pptx.slideLayouts)) {
    for (const layout of pptx.slideLayouts) {
      layout.nodes = layout.nodes || [];
      layout.background = layout.background || { type: 'none' };
      if (!layout.slideMaster) {
        layout.slideMaster = defaultMaster;
      }
      normalizeNodeTree(layout.nodes, layout.slideMaster, true);
    }
  }

  if (Array.isArray(pptx.slides)) {
    for (const slide of pptx.slides) {
      slide.nodes = slide.nodes || [];
      slide.background = slide.background || { type: 'none' };
      if (!slide.slideLayout) {
        slide.slideLayout = defaultLayout;
      }
      if (!slide.slideLayout.slideMaster) {
        slide.slideLayout.slideMaster = defaultMaster;
      }
      if (!slide.slideMaster) {
        slide.slideMaster = slide.slideLayout.slideMaster || defaultMaster;
      }
      normalizeNodeTree(slide.nodes, slide.slideMaster || defaultMaster, false);
    }
  }
}

async function createAndLoadPreviewer(
  rawBuffer: ArrayBuffer
): Promise<CachedDeckPreviewer> {
  const compatibleBuffer = await ensurePptxArchiveCompatibility(rawBuffer);

  const offscreenDom = document.createElement('div');
  const previewer = init(offscreenDom, {
    width: 960,
    height: 540,
    mode: 'list',
  }) as unknown as PptxPreviewerInternal;

  const pptx = await previewer.load(compatibleBuffer.slice(0));

  // Allow any asynchronous SlideMaster._loadNodes() calls inside pptx-preview to settle
  await new Promise((resolve) => setTimeout(resolve, 25));

  if (!pptx || !Array.isArray(pptx.slides) || pptx.slides.length === 0) {
    throw new Error('No slides parsed by pptx-preview');
  }

  normalizeParsedPptxModel(pptx);

  const width = pptx.width && pptx.width > 0 ? pptx.width : 960;
  const height = pptx.height && pptx.height > 0 ? pptx.height : 540;

  return {
    previewer,
    width,
    height,
    slideCount: pptx.slides.length,
  };
}

/**
 * Loads and caches a `pptx-preview` instance for a deck so subsequent slide
 * transitions or thumbnail renders execute in < 2ms without re-parsing the ZIP.
 */
export function loadDeckPreviewer(
  deckId: string | undefined,
  rawBuffer: ArrayBuffer
): Promise<CachedDeckPreviewer> {
  if (deckId) {
    const existing = previewerCacheByKey.get(deckId);
    if (existing) return existing;
    const promise = createAndLoadPreviewer(rawBuffer).catch((err) => {
      previewerCacheByKey.delete(deckId);
      throw err;
    });
    previewerCacheByKey.set(deckId, promise);
    return promise;
  }

  const existingByBuf = previewerCacheByBuffer.get(rawBuffer);
  if (existingByBuf) return existingByBuf;
  const promise = createAndLoadPreviewer(rawBuffer);
  previewerCacheByBuffer.set(rawBuffer, promise);
  return promise;
}

/**
 * Evicts a deck from the in-memory `pptx-preview` cache when deleted.
 */
export function evictDeckPreviewer(deckId: string): void {
  previewerCacheByKey.delete(deckId);
}

/**
 * Renders a single slide (`slideIndex`) into `targetContainer` at the presentation's
 * exact native 1:1 coordinate dimensions (`width` x `height`).
 */
export async function renderSlideWithPptxPreview(
  deckId: string | undefined,
  rawBuffer: ArrayBuffer,
  slideIndex: number,
  targetContainer: HTMLElement
): Promise<{ width: number; height: number }> {
  const cached = await loadDeckPreviewer(deckId, rawBuffer);
  const { previewer, width, height, slideCount } = cached;

  const safeIndex = Math.max(0, Math.min(slideCount - 1, slideIndex));

  // Configure HtmlRender to render at 1:1 native scale into a temporary fragment
  const tempWrapper = document.createElement('div');
  previewer.htmlRender.wrapper = tempWrapper;
  previewer.htmlRender.options = {
    viewPort: { width, height },
    mode: 'list',
  };
  previewer.htmlRender._calcScaleAndRenderPort();
  previewer.htmlRender.renderSlide(safeIndex);

  const slideWrapper = tempWrapper.firstElementChild as HTMLElement | null;
  if (slideWrapper) {
    slideWrapper.style.margin = '0';
    slideWrapper.style.position = 'relative';
    slideWrapper.style.boxSizing = 'border-box';

    // Fix pptx-preview blipFill background-size missing 'px'/'%' units
    const bgEl = slideWrapper.querySelector(
      '.slide-background'
    ) as HTMLElement | null;
    if (bgEl && bgEl.style.backgroundImage && !bgEl.style.backgroundSize) {
      bgEl.style.backgroundSize = '100% 100%';
    }

    // Remove pptx-preview's hardcoded `-0.04em` negative letter-spacing so Latin text widths match PowerPoint 1:1
    const textSpans = slideWrapper.querySelectorAll<HTMLElement>('.slide-text');
    for (let i = 0; i < textSpans.length; i++) {
      textSpans[i].style.letterSpacing = 'normal';
    }
  }

  targetContainer.innerHTML = '';
  if (slideWrapper) {
    targetContainer.appendChild(slideWrapper);
  }

  return { width, height };
}
