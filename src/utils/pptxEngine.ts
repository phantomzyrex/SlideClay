import JSZip from 'jszip';
import {
  ParsedSlide,
  SlideElement,
  SlideParagraph,
  SlideTextRun,
} from '../types/presentation';
import { emuToCanvasCoord, validatePresentationFile } from './deckUtils';

const DEFAULT_SLIDE_WIDTH_EMU = 12192000; // 16:9 widescreen (13.333 in)
const DEFAULT_SLIDE_HEIGHT_EMU = 6858000; // 7.5 in
const BASE_CANVAS_WIDTH = 960;

const PASTEL_ACCENTS = [
  '#B9A7FF',
  '#FFC8A8',
  '#A8EBD0',
  '#A8D4FF',
  '#FFE9A8',
];

/**
 * Decodes XML character entities into plain text.
 */
export function decodeXmlEntities(raw: string): string {
  return raw
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, dec) => String.fromCharCode(Number(dec)))
    .replace(/&#x([0-9a-fA-F]+);/g, (_, hex) =>
      String.fromCharCode(parseInt(hex, 16))
    )
    .replace(/&amp;/g, '&');
}

/* ============================================================================
 * HIERARCHICAL XML TREE PARSER (Works in Web Workers, Browser, and Node)
 * ============================================================================ */

export interface XmlNode {
  tag: string;
  attrs: Record<string, string>;
  children: XmlNode[];
  text: string;
}

const ATTR_REGEX = /([a-zA-Z0-9_:-]+)\s*=\s*"([^"]*)"/g;

function parseAttrs(attrString: string): Record<string, string> {
  const attrs: Record<string, string> = {};
  if (!attrString) return attrs;
  let m: RegExpExecArray | null;
  ATTR_REGEX.lastIndex = 0;
  while ((m = ATTR_REGEX.exec(attrString)) !== null) {
    attrs[m[1]] = decodeXmlEntities(m[2]);
  }
  return attrs;
}

/**
 * Parses an OpenXML string into a lightweight hierarchical tree of `XmlNode` objects.
 * Correctly handles nested tags of the same name (e.g. `<p:grpSp>` inside `<p:grpSp>`,
 * `<a:xfrm>` inside `<p:grpSpPr>` vs `<p:spPr>`) and self-closing elements.
 */
export function parseXmlTree(xml: string): XmlNode {
  const root: XmlNode = { tag: 'root', attrs: {}, children: [], text: '' };
  const stack: XmlNode[] = [root];

  const tokenRegex = /<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<!\[CDATA\[([\s\S]*?)\]\]>|<([/?!]?[a-zA-Z0-9_:-]+)([^>]*?)(\/?)>/g;
  let lastIndex = 0;
  let match: RegExpExecArray | null;

  while ((match = tokenRegex.exec(xml)) !== null) {
    const parent = stack[stack.length - 1];
    if (match.index > lastIndex) {
      const rawText = xml.slice(lastIndex, match.index);
      if (rawText) {
        parent.text += decodeXmlEntities(rawText);
      }
    }
    lastIndex = tokenRegex.lastIndex;

    if (match[1] !== undefined) {
      // CDATA content
      parent.text += match[1];
      continue;
    }

    const rawTag = match[2];
    if (!rawTag || rawTag.startsWith('!') || rawTag.startsWith('?')) {
      continue;
    }

    if (rawTag.startsWith('/')) {
      const closeTagName = rawTag.slice(1);
      // Pop up to matching tag
      for (let i = stack.length - 1; i >= 1; i--) {
        if (stack[i].tag === closeTagName) {
          stack.length = i;
          break;
        }
      }
    } else {
      const attrStr = match[3] || '';
      const selfClosing = match[4] === '/' || attrStr.trim().endsWith('/');
      const node: XmlNode = {
        tag: rawTag,
        attrs: parseAttrs(attrStr),
        children: [],
        text: '',
      };
      parent.children.push(node);
      if (!selfClosing) {
        stack.push(node);
      }
    }
  }

  return root;
}

export function findChild(node: XmlNode | undefined, tagName: string): XmlNode | undefined {
  if (!node) return undefined;
  return node.children.find((c) => c.tag === tagName);
}

export function findChildren(node: XmlNode | undefined, tagName: string): XmlNode[] {
  if (!node) return [];
  return node.children.filter((c) => c.tag === tagName);
}

export function findDescendant(node: XmlNode | undefined, tagName: string): XmlNode | undefined {
  if (!node) return undefined;
  for (const child of node.children) {
    if (child.tag === tagName) return child;
    const found = findDescendant(child, tagName);
    if (found) return found;
  }
  return undefined;
}

export function findDescendants(node: XmlNode | undefined, tagName: string): XmlNode[] {
  const out: XmlNode[] = [];
  if (!node) return out;
  for (const child of node.children) {
    if (child.tag === tagName) out.push(child);
    if (child.children.length > 0) {
      out.push(...findDescendants(child, tagName));
    }
  }
  return out;
}

/* ============================================================================
 * THEME COLOR & PLACEHOLDER LAYOUT RESOLUTION
 * ============================================================================ */

export type ThemeColorMap = Record<string, string>;

const DEFAULT_THEME_COLORS: ThemeColorMap = {
  dk1: '#231D38',
  tx1: '#231D38',
  lt1: '#FFFFFF',
  bg1: '#FFFFFF',
  dk2: '#4A3F6E',
  tx2: '#4A3F6E',
  lt2: '#F4F1FF',
  bg2: '#F4F1FF',
  accent1: '#7C63E6',
  accent2: '#E8824B',
  accent3: '#2B9E74',
  accent4: '#3B82D6',
  accent5: '#C9971C',
  accent6: '#D94872',
};

/**
 * Parses `ppt/theme/theme1.xml` to extract exact presentation color scheme tokens.
 */
function parseThemeXml(themeXml: string): ThemeColorMap {
  const map: ThemeColorMap = { ...DEFAULT_THEME_COLORS };
  if (!themeXml) return map;
  const tree = parseXmlTree(themeXml);
  const clrScheme = findDescendant(tree, 'a:clrScheme');
  if (!clrScheme) return map;

  for (const child of clrScheme.children) {
    const tokenName = child.tag.replace(/^a:/, '');
    const srgb = findDescendant(child, 'a:srgbClr');
    const sys = findDescendant(child, 'a:sysClr');
    const hex = srgb?.attrs.val || sys?.attrs.lastClr;
    if (hex && /^[0-9A-Fa-f]{6}$/.test(hex)) {
      map[tokenName] = `#${hex.toUpperCase()}`;
    }
  }
  // Map aliases
  if (map.dk1) map.tx1 = map.dk1;
  if (map.lt1) map.bg1 = map.lt1;
  if (map.dk2) map.tx2 = map.dk2;
  if (map.lt2) map.bg2 = map.lt2;
  return map;
}

/**
 * Extracts a `#RRGGBB` color from a node containing `<a:srgbClr>`, `<a:schemeClr>`, or `<a:sysClr>`.
 */
function extractColorFromNode(
  node: XmlNode | undefined,
  themeColors: ThemeColorMap,
  fallback = ''
): string {
  if (!node) return fallback;
  const srgb = node.tag === 'a:srgbClr' ? node : findDescendant(node, 'a:srgbClr');
  if (srgb?.attrs.val && /^[0-9A-Fa-f]{6}$/.test(srgb.attrs.val)) {
    return `#${srgb.attrs.val.toUpperCase()}`;
  }
  const sys = node.tag === 'a:sysClr' ? node : findDescendant(node, 'a:sysClr');
  if (sys?.attrs.lastClr && /^[0-9A-Fa-f]{6}$/.test(sys.attrs.lastClr)) {
    return `#${sys.attrs.lastClr.toUpperCase()}`;
  }
  const scheme = node.tag === 'a:schemeClr' ? node : findDescendant(node, 'a:schemeClr');
  if (scheme?.attrs.val) {
    const key = scheme.attrs.val;
    return themeColors[key] || DEFAULT_THEME_COLORS[key] || fallback;
  }
  return fallback;
}

export interface PlaceholderTemplate {
  phType?: string;
  phIdx?: string;
  xEmu?: number;
  yEmu?: number;
  wEmu?: number;
  hEmu?: number;
  verticalAlign?: 'top' | 'middle' | 'bottom';
  paddingEmu?: { top?: number; right?: number; bottom?: number; left?: number };
  defaultAlign?: 'left' | 'center' | 'right' | 'justify';
  levelFontSizesPt?: Record<number, number>;
  levelTextColors?: Record<number, string>;
  fillColor?: string;
}

export interface SlideInheritanceContext {
  placeholdersByIdx: Map<string, PlaceholderTemplate>;
  placeholdersByType: Map<string, PlaceholderTemplate>;
  backgroundColor?: string;
  themeColors: ThemeColorMap;
}

function parseLstStyleNode(
  lstStyle: XmlNode | undefined,
  themeColors: ThemeColorMap
): {
  defaultAlign?: 'left' | 'center' | 'right' | 'justify';
  levelFontSizesPt: Record<number, number>;
  levelTextColors: Record<number, string>;
} {
  const levelFontSizesPt: Record<number, number> = {};
  const levelTextColors: Record<number, string> = {};
  let defaultAlign: 'left' | 'center' | 'right' | 'justify' | undefined;

  if (!lstStyle) return { defaultAlign, levelFontSizesPt, levelTextColors };

  for (let lvl = 1; lvl <= 9; lvl++) {
    const lvlNode = findChild(lstStyle, `a:lvl${lvl}pPr`);
    if (!lvlNode) continue;
    const algn = lvlNode.attrs.algn;
    if (lvl === 1 && algn) {
      defaultAlign =
        algn === 'ctr'
          ? 'center'
          : algn === 'r'
            ? 'right'
            : algn === 'just'
              ? 'justify'
              : 'left';
    }
    const defRPr = findChild(lvlNode, 'a:defRPr');
    if (defRPr) {
      const sz = Number(defRPr.attrs.sz ?? NaN);
      if (!Number.isNaN(sz) && sz > 0) {
        levelFontSizesPt[lvl - 1] = sz / 100;
      }
      const clr = extractColorFromNode(defRPr, themeColors, '');
      if (clr) {
        levelTextColors[lvl - 1] = clr;
      }
    }
  }

  return { defaultAlign, levelFontSizesPt, levelTextColors };
}

/**
 * Extracts placeholder coordinate and typography templates from a `slideLayout` or `slideMaster` XML tree.
 */
function extractTemplatesFromLayoutOrMaster(
  xmlStr: string,
  themeColors: ThemeColorMap
): {
  byIdx: Map<string, PlaceholderTemplate>;
  byType: Map<string, PlaceholderTemplate>;
  bgColor?: string;
  masterTitleStyle?: ReturnType<typeof parseLstStyleNode>;
  masterBodyStyle?: ReturnType<typeof parseLstStyleNode>;
} {
  const byIdx = new Map<string, PlaceholderTemplate>();
  const byType = new Map<string, PlaceholderTemplate>();
  if (!xmlStr) return { byIdx, byType };

  const tree = parseXmlTree(xmlStr);
  const bgNode = findDescendant(tree, 'p:bg');
  const bgColor = bgNode ? extractColorFromNode(bgNode, themeColors, '') : undefined;

  // Check `<p:txStyles>` in slideMaster (`<p:titleStyle>`, `<p:bodyStyle>`, `<p:otherStyle>`)
  const txStyles = findDescendant(tree, 'p:txStyles');
  const masterTitleStyle = txStyles
    ? parseLstStyleNode(findChild(txStyles, 'p:titleStyle'), themeColors)
    : undefined;
  const masterBodyStyle = txStyles
    ? parseLstStyleNode(findChild(txStyles, 'p:bodyStyle'), themeColors)
    : undefined;

  const spNodes = findDescendants(tree, 'p:sp');
  for (const sp of spNodes) {
    const nvPr = findDescendant(sp, 'p:nvPr');
    const ph = findChild(nvPr, 'p:ph');
    if (!ph) continue;

    const phType = ph.attrs.type;
    const phIdx = ph.attrs.idx;

    const spPr = findChild(sp, 'p:spPr');
    const xfrm = findChild(spPr, 'a:xfrm');
    const off = findChild(xfrm, 'a:off');
    const ext = findChild(xfrm, 'a:ext');

    const xEmu = off?.attrs.x !== undefined ? Number(off.attrs.x) : undefined;
    const yEmu = off?.attrs.y !== undefined ? Number(off.attrs.y) : undefined;
    const wEmu = ext?.attrs.cx !== undefined ? Number(ext.attrs.cx) : undefined;
    const hEmu = ext?.attrs.cy !== undefined ? Number(ext.attrs.cy) : undefined;

    const txBody = findChild(sp, 'p:txBody');
    const bodyPr = findChild(txBody, 'a:bodyPr');
    const anchorAttr = bodyPr?.attrs.anchor;
    const verticalAlign: 'top' | 'middle' | 'bottom' | undefined =
      anchorAttr === 'ctr'
        ? 'middle'
        : anchorAttr === 'b'
          ? 'bottom'
          : anchorAttr === 't'
            ? 'top'
            : undefined;

    const paddingEmu = bodyPr
      ? {
          left: bodyPr.attrs.lIns !== undefined ? Number(bodyPr.attrs.lIns) : undefined,
          top: bodyPr.attrs.tIns !== undefined ? Number(bodyPr.attrs.tIns) : undefined,
          right: bodyPr.attrs.rIns !== undefined ? Number(bodyPr.attrs.rIns) : undefined,
          bottom: bodyPr.attrs.bIns !== undefined ? Number(bodyPr.attrs.bIns) : undefined,
        }
      : undefined;

    const lstStyle = findChild(txBody, 'a:lstStyle');
    const parsedLst = parseLstStyleNode(lstStyle, themeColors);

    const solidFill = findChild(spPr, 'a:solidFill');
    const fillColor = solidFill
      ? extractColorFromNode(solidFill, themeColors, '')
      : undefined;

    const tpl: PlaceholderTemplate = {
      phType,
      phIdx,
      xEmu: !Number.isNaN(xEmu ?? NaN) ? xEmu : undefined,
      yEmu: !Number.isNaN(yEmu ?? NaN) ? yEmu : undefined,
      wEmu: !Number.isNaN(wEmu ?? NaN) ? wEmu : undefined,
      hEmu: !Number.isNaN(hEmu ?? NaN) ? hEmu : undefined,
      verticalAlign,
      paddingEmu,
      defaultAlign: parsedLst.defaultAlign,
      levelFontSizesPt: parsedLst.levelFontSizesPt,
      levelTextColors: parsedLst.levelTextColors,
      fillColor: fillColor || undefined,
    };

    if (phIdx !== undefined) {
      byIdx.set(phIdx, tpl);
    }
    if (phType) {
      byType.set(phType, tpl);
      // Map ctrTitle <-> title and subTitle <-> body fallbacks
      if (phType === 'ctrTitle' && !byType.has('title')) byType.set('title', tpl);
      if (phType === 'title' && !byType.has('ctrTitle')) byType.set('ctrTitle', tpl);
    }
  }

  return { byIdx, byType, bgColor, masterTitleStyle, masterBodyStyle };
}

function mergeTemplates(
  layoutTpl: PlaceholderTemplate | undefined,
  masterTpl: PlaceholderTemplate | undefined,
  masterTitleStyle?: ReturnType<typeof parseLstStyleNode>,
  masterBodyStyle?: ReturnType<typeof parseLstStyleNode>,
  phType?: string
): PlaceholderTemplate {
  const isTitle = phType === 'title' || phType === 'ctrTitle';
  const fallbackMasterStyle = isTitle ? masterTitleStyle : masterBodyStyle;

  return {
    phType: layoutTpl?.phType ?? masterTpl?.phType ?? phType,
    phIdx: layoutTpl?.phIdx ?? masterTpl?.phIdx,
    xEmu: layoutTpl?.xEmu ?? masterTpl?.xEmu,
    yEmu: layoutTpl?.yEmu ?? masterTpl?.yEmu,
    wEmu: layoutTpl?.wEmu ?? masterTpl?.wEmu,
    hEmu: layoutTpl?.hEmu ?? masterTpl?.hEmu,
    verticalAlign: layoutTpl?.verticalAlign ?? masterTpl?.verticalAlign,
    paddingEmu: {
      top: layoutTpl?.paddingEmu?.top ?? masterTpl?.paddingEmu?.top,
      right: layoutTpl?.paddingEmu?.right ?? masterTpl?.paddingEmu?.right,
      bottom: layoutTpl?.paddingEmu?.bottom ?? masterTpl?.paddingEmu?.bottom,
      left: layoutTpl?.paddingEmu?.left ?? masterTpl?.paddingEmu?.left,
    },
    defaultAlign:
      layoutTpl?.defaultAlign ??
      masterTpl?.defaultAlign ??
      fallbackMasterStyle?.defaultAlign,
    levelFontSizesPt: {
      ...(fallbackMasterStyle?.levelFontSizesPt || {}),
      ...(masterTpl?.levelFontSizesPt || {}),
      ...(layoutTpl?.levelFontSizesPt || {}),
    },
    levelTextColors: {
      ...(fallbackMasterStyle?.levelTextColors || {}),
      ...(masterTpl?.levelTextColors || {}),
      ...(layoutTpl?.levelTextColors || {}),
    },
    fillColor: layoutTpl?.fillColor ?? masterTpl?.fillColor,
  };
}

/* ============================================================================
 * GROUP COORDINATE TRANSFORMATIONS (`<p:grpSp>`)
 * ============================================================================ */

interface GroupTransform {
  /** Maps an EMU X coordinate from local space into root slide EMU space */
  mapX: (xEmu: number) => number;
  /** Maps an EMU Y coordinate from local space into root slide EMU space */
  mapY: (yEmu: number) => number;
  /** Scales an EMU width from local space into root slide EMU space */
  scaleW: (wEmu: number) => number;
  /** Scales an EMU height from local space into root slide EMU space */
  scaleH: (hEmu: number) => number;
}

const IDENTITY_TRANSFORM: GroupTransform = {
  mapX: (x) => x,
  mapY: (y) => y,
  scaleW: (w) => w,
  scaleH: (h) => h,
};

function composeGroupTransform(
  parent: GroupTransform,
  grpSpPr: XmlNode | undefined
): GroupTransform {
  const xfrm = findChild(grpSpPr, 'a:xfrm');
  if (!xfrm) return parent;

  const off = findChild(xfrm, 'a:off');
  const ext = findChild(xfrm, 'a:ext');
  const chOff = findChild(xfrm, 'a:chOff');
  const chExt = findChild(xfrm, 'a:chExt');

  const gx = Number(off?.attrs.x ?? NaN);
  const gy = Number(off?.attrs.y ?? NaN);
  const gw = Number(ext?.attrs.cx ?? NaN);
  const gh = Number(ext?.attrs.cy ?? NaN);
  const cx = Number(chOff?.attrs.x ?? NaN);
  const cy = Number(chOff?.attrs.y ?? NaN);
  const cw = Number(chExt?.attrs.cx ?? NaN);
  const ch = Number(chExt?.attrs.cy ?? NaN);

  if (
    Number.isNaN(gx) ||
    Number.isNaN(gy) ||
    Number.isNaN(gw) ||
    Number.isNaN(gh) ||
    Number.isNaN(cx) ||
    Number.isNaN(cy) ||
    Number.isNaN(cw) ||
    Number.isNaN(ch) ||
    cw <= 0 ||
    ch <= 0
  ) {
    return parent;
  }

  return {
    mapX: (localX) => parent.mapX(gx + ((localX - cx) / cw) * gw),
    mapY: (localY) => parent.mapY(gy + ((localY - cy) / ch) * gh),
    scaleW: (localW) => parent.scaleW((localW / cw) * gw),
    scaleH: (localH) => parent.scaleH((localH / ch) * gh),
  };
}

/**
 * Standard OpenXML fallback coordinates (in EMUs) when a placeholder has no layout/master match.
 */
function getStandardPlaceholderEmuBox(
  phType: string | undefined,
  phIdx: string | undefined,
  slideWidthEmu: number,
  slideHeightEmu: number
): { x: number; y: number; w: number; h: number } {
  if (phType === 'ctrTitle') {
    return {
      x: slideWidthEmu * 0.1,
      y: slideHeightEmu * 0.25,
      w: slideWidthEmu * 0.8,
      h: slideHeightEmu * 0.24,
    };
  }
  if (phType === 'title') {
    return {
      x: slideWidthEmu * 0.065,
      y: slideHeightEmu * 0.06,
      w: slideWidthEmu * 0.87,
      h: slideHeightEmu * 0.16,
    };
  }
  if (phType === 'subTitle') {
    return {
      x: slideWidthEmu * 0.12,
      y: slideHeightEmu * 0.52,
      w: slideWidthEmu * 0.76,
      h: slideHeightEmu * 0.18,
    };
  }
  if (phType === 'dt') {
    return {
      x: slideWidthEmu * 0.065,
      y: slideHeightEmu * 0.91,
      w: slideWidthEmu * 0.22,
      h: slideHeightEmu * 0.05,
    };
  }
  if (phType === 'ftr') {
    return {
      x: slideWidthEmu * 0.34,
      y: slideHeightEmu * 0.91,
      w: slideWidthEmu * 0.32,
      h: slideHeightEmu * 0.05,
    };
  }
  if (phType === 'sldNum') {
    return {
      x: slideWidthEmu * 0.71,
      y: slideHeightEmu * 0.91,
      w: slideWidthEmu * 0.22,
      h: slideHeightEmu * 0.05,
    };
  }
  // Body / content placeholder (check if two-column idx="1" vs idx="2")
  if (phIdx === '2' && (phType === 'body' || phType === 'obj' || !phType)) {
    return {
      x: slideWidthEmu * 0.51,
      y: slideHeightEmu * 0.24,
      w: slideWidthEmu * 0.425,
      h: slideHeightEmu * 0.65,
    };
  }
  return {
    x: slideWidthEmu * 0.065,
    y: slideHeightEmu * 0.24,
    w: slideWidthEmu * 0.87,
    h: slideHeightEmu * 0.65,
  };
}

/**
 * Converts font size in OpenXML points (`pt`) to exact canvas pixels based on the slide's physical height in EMUs.
 * In OpenXML: `1 pt = 12,700 EMUs`.
 */
function ptToCanvasPx(
  fontPt: number,
  slideHeightEmu: number,
  canvasHeight: number
): number {
  const fontEmu = fontPt * 12700;
  const px = (fontEmu / slideHeightEmu) * canvasHeight;
  return Number(Math.max(8, Math.min(96, px)).toFixed(1));
}

/**
 * Parses a single `<p:sp>` node into a `SlideElement`, applying group transforms and layout/master inheritance.
 */
function parseShapeNode(
  spNode: XmlNode,
  idSuffix: string,
  slideWidthEmu: number,
  slideHeightEmu: number,
  canvasWidth: number,
  canvasHeight: number,
  groupTransform: GroupTransform,
  inheritance: SlideInheritanceContext
): {
  element: SlideElement | null;
  detectedTitle?: string;
  detectedSubtitle?: string;
} {
  const nvSpPr = findChild(spNode, 'p:nvSpPr');
  const nvPr = findChild(nvSpPr, 'p:nvPr');
  const phNode = findChild(nvPr, 'p:ph');
  const phType = phNode?.attrs.type;
  const phIdx = phNode?.attrs.idx;

  // Lookup inherited placeholder template from slideLayout / slideMaster
  const inheritedTpl =
    (phIdx !== undefined ? inheritance.placeholdersByIdx.get(phIdx) : undefined) ||
    (phType ? inheritance.placeholdersByType.get(phType) : undefined) ||
    (phNode ? inheritance.placeholdersByType.get('body') : undefined);

  const isTitle = phType === 'title' || phType === 'ctrTitle';
  const isSubtitle = phType === 'subTitle';
  const isFooterOrMeta = phType === 'dt' || phType === 'ftr' || phType === 'sldNum';

  const spPr = findChild(spNode, 'p:spPr');
  const xfrm = findChild(spPr, 'a:xfrm');
  const off = findChild(xfrm, 'a:off');
  const ext = findChild(xfrm, 'a:ext');

  const rawX = off?.attrs.x !== undefined ? Number(off.attrs.x) : NaN;
  const rawY = off?.attrs.y !== undefined ? Number(off.attrs.y) : NaN;
  const rawW = ext?.attrs.cx !== undefined ? Number(ext.attrs.cx) : NaN;
  const rawH = ext?.attrs.cy !== undefined ? Number(ext.attrs.cy) : NaN;

  const fallbackBox = getStandardPlaceholderEmuBox(
    phType,
    phIdx,
    slideWidthEmu,
    slideHeightEmu
  );

  const localXEmu = !Number.isNaN(rawX)
    ? rawX
    : inheritedTpl?.xEmu !== undefined
      ? inheritedTpl.xEmu
      : fallbackBox.x;
  const localYEmu = !Number.isNaN(rawY)
    ? rawY
    : inheritedTpl?.yEmu !== undefined
      ? inheritedTpl.yEmu
      : fallbackBox.y;
  const localWEmu = !Number.isNaN(rawW) && rawW > 0
    ? rawW
    : inheritedTpl?.wEmu !== undefined && inheritedTpl.wEmu > 0
      ? inheritedTpl.wEmu
      : fallbackBox.w;
  const localHEmu = !Number.isNaN(rawH) && rawH > 0
    ? rawH
    : inheritedTpl?.hEmu !== undefined && inheritedTpl.hEmu > 0
      ? inheritedTpl.hEmu
      : fallbackBox.h;

  const slideXEmu = !Number.isNaN(rawX)
    ? groupTransform.mapX(localXEmu)
    : localXEmu;
  const slideYEmu = !Number.isNaN(rawY)
    ? groupTransform.mapY(localYEmu)
    : localYEmu;
  const slideWEmu = !Number.isNaN(rawW)
    ? groupTransform.scaleW(localWEmu)
    : localWEmu;
  const slideHEmu = !Number.isNaN(rawH)
    ? groupTransform.scaleH(localHEmu)
    : localHEmu;

  const x = emuToCanvasCoord(slideXEmu, slideWidthEmu, canvasWidth);
  const y = emuToCanvasCoord(slideYEmu, slideHeightEmu, canvasHeight);
  const w = Math.max(4, emuToCanvasCoord(slideWEmu, slideWidthEmu, canvasWidth));
  const h = Math.max(4, emuToCanvasCoord(slideHEmu, slideHeightEmu, canvasHeight));

  // Optional rotation (`rot` on `<a:xfrm>` is in 1/60,000ths of a degree)
  const rotAttr = xfrm?.attrs.rot !== undefined ? Number(xfrm.attrs.rot) : 0;
  const rotation =
    !Number.isNaN(rotAttr) && rotAttr !== 0
      ? Number((rotAttr / 60000).toFixed(1))
      : undefined;

  // Shape Geometry & Fill
  const prstGeom = findChild(spPr, 'a:prstGeom');
  const prst = prstGeom?.attrs.prst || 'rect';
  const hasExplicitNoFill = Boolean(findChild(spPr, 'a:noFill'));
  const solidFill = findChild(spPr, 'a:solidFill');
  const gradFill = findChild(spPr, 'a:gradFill');

  let fillColor = '';
  if (!hasExplicitNoFill) {
    if (solidFill) {
      fillColor = extractColorFromNode(solidFill, inheritance.themeColors, '');
    } else if (gradFill) {
      fillColor = extractColorFromNode(gradFill, inheritance.themeColors, '');
    } else if (inheritedTpl?.fillColor) {
      fillColor = inheritedTpl.fillColor;
    }
  }

  // Stroke (`<a:ln>`)
  const lnNode = findChild(spPr, 'a:ln');
  const lnHasNoFill = Boolean(findChild(lnNode, 'a:noFill'));
  const strokeColor =
    lnNode && !lnHasNoFill
      ? extractColorFromNode(lnNode, inheritance.themeColors, '')
      : undefined;
  const lnWidthEmu = Number(lnNode?.attrs.w ?? 0);
  const strokeWidth =
    lnWidthEmu > 0
      ? Math.max(1, Number(((lnWidthEmu / slideHeightEmu) * canvasHeight).toFixed(1)))
      : undefined;

  // Text Body (`<p:txBody>`)
  const txBody = findChild(spNode, 'p:txBody');
  const bodyPr = findChild(txBody, 'a:bodyPr');
  const lstStyle = findChild(txBody, 'a:lstStyle');
  const localLstStyle = parseLstStyleNode(lstStyle, inheritance.themeColors);

  const anchorAttr = bodyPr?.attrs.anchor;
  const verticalAlign: 'top' | 'middle' | 'bottom' =
    anchorAttr === 'ctr'
      ? 'middle'
      : anchorAttr === 'b'
        ? 'bottom'
        : anchorAttr === 't'
          ? 'top'
          : inheritedTpl?.verticalAlign || (phType === 'ctrTitle' ? 'middle' : 'top');

  // Internal text box margins in EMUs (OpenXML defaults: lIns=91440, rIns=91440, tIns=45720, bIns=45720)
  const lInsEmu =
    bodyPr?.attrs.lIns !== undefined
      ? Number(bodyPr.attrs.lIns)
      : inheritedTpl?.paddingEmu?.left ?? 91440;
  const rInsEmu =
    bodyPr?.attrs.rIns !== undefined
      ? Number(bodyPr.attrs.rIns)
      : inheritedTpl?.paddingEmu?.right ?? 91440;
  const tInsEmu =
    bodyPr?.attrs.tIns !== undefined
      ? Number(bodyPr.attrs.tIns)
      : inheritedTpl?.paddingEmu?.top ?? 45720;
  const bInsEmu =
    bodyPr?.attrs.bIns !== undefined
      ? Number(bodyPr.attrs.bIns)
      : inheritedTpl?.paddingEmu?.bottom ?? 45720;

  const padding = {
    left: emuToCanvasCoord(Math.max(0, lInsEmu), slideWidthEmu, canvasWidth),
    right: emuToCanvasCoord(Math.max(0, rInsEmu), slideWidthEmu, canvasWidth),
    top: emuToCanvasCoord(Math.max(0, tInsEmu), slideHeightEmu, canvasHeight),
    bottom: emuToCanvasCoord(Math.max(0, bInsEmu), slideHeightEmu, canvasHeight),
  };

  const wordWrap = bodyPr?.attrs.wrap !== 'none';

  // Parse paragraphs (`<a:p>`)
  const pNodes = findChildren(txBody, 'a:p');
  const paragraphs: SlideParagraph[] = [];
  let plainTextCombined = '';

  for (const pNode of pNodes) {
    const pPr = findChild(pNode, 'a:pPr');
    const lvl = Number(pPr?.attrs.lvl ?? 0);
    const algnAttr = pPr?.attrs.algn;

    const align: 'left' | 'center' | 'right' | 'justify' = algnAttr
      ? algnAttr === 'ctr'
        ? 'center'
        : algnAttr === 'r'
          ? 'right'
          : algnAttr === 'just'
            ? 'justify'
            : 'left'
      : localLstStyle.defaultAlign ||
        inheritedTpl?.defaultAlign ||
        (phType === 'ctrTitle' ? 'center' : 'left');

    const hasNoBullet = Boolean(findChild(pPr, 'a:buNone'));
    const buCharNode = findChild(pPr, 'a:buChar');
    const buAutoNumNode = findChild(pPr, 'a:buAutoNum');
    const hasBullet = !hasNoBullet && Boolean(buCharNode || buAutoNumNode);
    const bulletChar = buCharNode?.attrs.char || '•';

    const defRPr = findChild(pPr, 'a:defRPr');
    const pDefaultSzAttr = defRPr?.attrs.sz ? Number(defRPr.attrs.sz) / 100 : undefined;
    const pDefaultColor = defRPr
      ? extractColorFromNode(defRPr, inheritance.themeColors, '')
      : '';

    const defaultPtForLevel =
      pDefaultSzAttr ||
      localLstStyle.levelFontSizesPt[lvl] ||
      inheritedTpl?.levelFontSizesPt?.[lvl] ||
      (isTitle ? 36 : isSubtitle ? 20 : isFooterOrMeta ? 11 : 18 - Math.min(lvl * 2, 6));

    const defaultColorForLevel =
      pDefaultColor ||
      localLstStyle.levelTextColors[lvl] ||
      inheritedTpl?.levelTextColors?.[lvl] ||
      inheritance.themeColors.tx1 ||
      '#231D38';

    const runs: SlideTextRun[] = [];

    for (const child of pNode.children) {
      if (child.tag === 'a:r' || child.tag === 'a:fld') {
        const tNode = findChild(child, 'a:t');
        const text = tNode?.text ?? '';
        if (!text) continue;

        const rPr = findChild(child, 'a:rPr');
        const szAttr = rPr?.attrs.sz !== undefined ? Number(rPr.attrs.sz) : NaN;
        const fontPt =
          !Number.isNaN(szAttr) && szAttr > 0 ? szAttr / 100 : defaultPtForLevel;
        const fontPx = ptToCanvasPx(fontPt, slideHeightEmu, canvasHeight);

        const boldAttr = rPr?.attrs.b;
        const italicAttr = rPr?.attrs.i;
        const uAttr = rPr?.attrs.u;
        const runColor = extractColorFromNode(rPr, inheritance.themeColors, '');
        const latinNode = findChild(rPr, 'a:latin') || findChild(defRPr, 'a:latin');
        const rawTypeface = latinNode?.attrs.typeface;
        const fontFamily =
          rawTypeface && !rawTypeface.startsWith('+') ? rawTypeface : undefined;

        runs.push({
          text,
          bold: boldAttr === '1' || boldAttr === 'true' || (boldAttr === undefined && isTitle),
          italic: italicAttr === '1' || italicAttr === 'true',
          underline: Boolean(uAttr && uAttr !== 'none'),
          fontSize: fontPx,
          fontFamily,
          color: runColor || defaultColorForLevel,
        });
        plainTextCombined += text;
      } else if (child.tag === 'a:br') {
        runs.push({
          text: '\n',
          fontSize: ptToCanvasPx(defaultPtForLevel, slideHeightEmu, canvasHeight),
        });
      }
    }

    if (runs.length > 0) {
      paragraphs.push({
        align,
        bullet: hasBullet,
        bulletChar,
        level: lvl,
        runs,
      });
    }
  }

  if (paragraphs.length === 0 && !fillColor && !strokeColor) {
    return { element: null };
  }

  const shapeKind: SlideElement['shapeKind'] =
    prst === 'ellipse'
      ? 'ellipse'
      : prst === 'roundRect'
        ? 'roundRect'
        : prst === 'line' || prst === 'straightConnector1'
          ? 'line'
          : 'rect';

  return {
    element: {
      id: `sp-${idSuffix}`,
      type: paragraphs.length > 0 && !fillColor ? 'text' : 'shape',
      x,
      y,
      w,
      h,
      rotation,
      shapeKind,
      fillColor: fillColor || undefined,
      strokeColor: strokeColor || undefined,
      strokeWidth,
      borderRadius:
        shapeKind === 'roundRect'
          ? Math.round(Math.min(w, h) * 0.14)
          : shapeKind === 'ellipse'
            ? 9999
            : 0,
      verticalAlign,
      padding,
      wordWrap,
      paragraphs: paragraphs.length > 0 ? paragraphs : undefined,
    },
    detectedTitle:
      isTitle && plainTextCombined.trim() ? plainTextCombined.trim() : undefined,
    detectedSubtitle:
      isSubtitle && plainTextCombined.trim()
        ? plainTextCombined.trim()
        : undefined,
  };
}

/**
 * Parses `<p:graphicFrame>` tables (`<a:tbl>`) into a `SlideElement` with accurate column widths.
 */
function parseTableNode(
  gfNode: XmlNode,
  idSuffix: string,
  slideWidthEmu: number,
  slideHeightEmu: number,
  canvasWidth: number,
  canvasHeight: number,
  groupTransform: GroupTransform,
  themeColors: ThemeColorMap
): SlideElement | null {
  const tblNode = findDescendant(gfNode, 'a:tbl');
  if (!tblNode) return null;

  const xfrm = findChild(gfNode, 'p:xfrm') || findDescendant(gfNode, 'a:xfrm');
  const off = findChild(xfrm, 'a:off');
  const ext = findChild(xfrm, 'a:ext');

  const rawX = Number(off?.attrs.x ?? slideWidthEmu * 0.065);
  const rawY = Number(off?.attrs.y ?? slideHeightEmu * 0.24);
  const rawW = Number(ext?.attrs.cx ?? slideWidthEmu * 0.87);
  const rawH = Number(ext?.attrs.cy ?? slideHeightEmu * 0.5);

  const x = emuToCanvasCoord(groupTransform.mapX(rawX), slideWidthEmu, canvasWidth);
  const y = emuToCanvasCoord(groupTransform.mapY(rawY), slideHeightEmu, canvasHeight);
  const w = Math.max(80, emuToCanvasCoord(groupTransform.scaleW(rawW), slideWidthEmu, canvasWidth));
  const h = Math.max(40, emuToCanvasCoord(groupTransform.scaleH(rawH), slideHeightEmu, canvasHeight));

  // Column widths from `<a:tblGrid><a:gridCol w="..."/>`
  const tblGrid = findChild(tblNode, 'a:tblGrid');
  const gridCols = findChildren(tblGrid, 'a:gridCol');
  const colWidthsPx = gridCols
    .map((c) => emuToCanvasCoord(Number(c.attrs.w ?? 0), slideWidthEmu, canvasWidth))
    .filter((val) => val > 0);

  const trNodes = findChildren(tblNode, 'a:tr');
  const allRows: string[][] = [];

  for (const tr of trNodes) {
    const tcNodes = findChildren(tr, 'a:tc');
    const rowCells: string[] = [];
    for (const tc of tcNodes) {
      const tNodes = findDescendants(tc, 'a:t');
      const cellText = tNodes
        .map((t) => t.text)
        .join(' ')
        .trim();
      rowCells.push(cellText);
    }
    if (rowCells.length > 0) {
      allRows.push(rowCells);
    }
  }

  if (allRows.length === 0) return null;

  return {
    id: `tbl-${idSuffix}`,
    type: 'table',
    x,
    y,
    w,
    h,
    table: {
      headers: allRows[0],
      rows: allRows.slice(1),
      colWidthsPx: colWidthsPx.length === allRows[0].length ? colWidthsPx : undefined,
      accentColor: themeColors.accent1 || '#B9A7FF',
    },
  };
}

/**
 * Parses `<p:pic>` embedded images into a `SlideElement`.
 */
async function parsePicNode(
  picNode: XmlNode,
  idSuffix: string,
  zip: JSZip,
  relMap: Map<string, string>,
  slideWidthEmu: number,
  slideHeightEmu: number,
  canvasWidth: number,
  canvasHeight: number,
  groupTransform: GroupTransform,
  inheritance: SlideInheritanceContext
): Promise<SlideElement | null> {
  const blip = findDescendant(picNode, 'a:blip');
  const embedId = blip?.attrs['r:embed'] || blip?.attrs['embed'];
  if (!embedId) return null;

  const mediaPath = relMap.get(embedId);
  if (!mediaPath) return null;
  const mediaFile = zip.file(mediaPath);
  if (!mediaFile) return null;

  try {
    const base64 = await mediaFile.async('base64');
    const ext = mediaPath.split('.').pop()?.toLowerCase() || 'png';
    const mime =
      ext === 'jpg' || ext === 'jpeg'
        ? 'image/jpeg'
        : ext === 'svg'
          ? 'image/svg+xml'
          : ext === 'gif'
            ? 'image/gif'
            : ext === 'webp'
              ? 'image/webp'
              : 'image/png';

    const nvPicPr = findChild(picNode, 'p:nvPicPr');
    const nvPr = findChild(nvPicPr, 'p:nvPr');
    const phNode = findChild(nvPr, 'p:ph');
    const phIdx = phNode?.attrs.idx;
    const phType = phNode?.attrs.type;
    const inheritedTpl =
      (phIdx !== undefined ? inheritance.placeholdersByIdx.get(phIdx) : undefined) ||
      (phType ? inheritance.placeholdersByType.get(phType) : undefined);

    const spPr = findChild(picNode, 'p:spPr');
    const xfrm = findChild(spPr, 'a:xfrm');
    const off = findChild(xfrm, 'a:off');
    const extNode = findChild(xfrm, 'a:ext');

    const rawX = Number(off?.attrs.x ?? inheritedTpl?.xEmu ?? slideWidthEmu * 0.1);
    const rawY = Number(off?.attrs.y ?? inheritedTpl?.yEmu ?? slideHeightEmu * 0.2);
    const rawW = Number(extNode?.attrs.cx ?? inheritedTpl?.wEmu ?? slideWidthEmu * 0.4);
    const rawH = Number(extNode?.attrs.cy ?? inheritedTpl?.hEmu ?? slideHeightEmu * 0.45);

    const x = emuToCanvasCoord(groupTransform.mapX(rawX), slideWidthEmu, canvasWidth);
    const y = emuToCanvasCoord(groupTransform.mapY(rawY), slideHeightEmu, canvasHeight);
    const w = Math.max(16, emuToCanvasCoord(groupTransform.scaleW(rawW), slideWidthEmu, canvasWidth));
    const h = Math.max(16, emuToCanvasCoord(groupTransform.scaleH(rawH), slideHeightEmu, canvasHeight));

    return {
      id: `pic-${idSuffix}`,
      type: 'image',
      x,
      y,
      w,
      h,
      imageUrl: `data:${mime};base64,${base64}`,
      borderRadius: 0,
    };
  } catch {
    return null;
  }
}

/**
 * Resolves a relative OpenXML relationship target path (e.g. `../slideLayouts/slideLayout1.xml`)
 * against the directory of the source part (`ppt/slides`).
 */
function resolveRelTarget(baseDir: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const stack = baseDir.split('/').filter(Boolean);
  for (const part of target.split('/')) {
    if (part === '.' || !part) continue;
    if (part === '..') stack.pop();
    else stack.push(part);
  }
  return stack.join('/');
}

/**
 * Reads an OpenXML `.rels` file into a Map of `RelationshipId -> NormalizedZipPath`.
 */
async function readRelsMap(
  zip: JSZip,
  partPath: string
): Promise<Map<string, string>> {
  const map = new Map<string, string>();
  const segments = partPath.split('/');
  const fileName = segments.pop() || '';
  const baseDir = segments.join('/');
  const relsPath = `${baseDir}/_rels/${fileName}.rels`;
  const relsFile = zip.file(relsPath);
  if (!relsFile) return map;

  const relsXml = await relsFile.async('string');
  const tree = parseXmlTree(relsXml);
  const relNodes = findDescendants(tree, 'Relationship');
  for (const r of relNodes) {
    const id = r.attrs.Id;
    const target = r.attrs.Target;
    if (id && target) {
      map.set(id, resolveRelTarget(baseDir, target));
    }
  }
  return map;
}

/**
 * Extracts a background image data URL from a `<p:bg>` node if it uses `<a:blipFill>`.
 */
async function extractBackgroundImageDataUrl(
  tree: XmlNode | undefined,
  zip: JSZip,
  relMap: Map<string, string>
): Promise<string | undefined> {
  if (!tree) return undefined;
  const bgNode = findDescendant(tree, 'p:bg');
  const blip = findDescendant(bgNode, 'a:blip');
  const embedId = blip?.attrs['r:embed'] || blip?.attrs['embed'];
  if (!embedId) return undefined;
  const mediaPath = relMap.get(embedId);
  if (!mediaPath) return undefined;
  const mediaFile = zip.file(mediaPath);
  if (!mediaFile) return undefined;
  try {
    const base64 = await mediaFile.async('base64');
    const ext = mediaPath.split('.').pop()?.toLowerCase() || 'png';
    const mime =
      ext === 'jpg' || ext === 'jpeg'
        ? 'image/jpeg'
        : ext === 'svg'
          ? 'image/svg+xml'
          : ext === 'gif'
            ? 'image/gif'
            : ext === 'webp'
              ? 'image/webp'
              : 'image/png';
    return `data:${mime};base64,${base64}`;
  } catch {
    return undefined;
  }
}

/**
 * Recursively traverses `<p:spTree>` or `<p:grpSp>` in exact z-order,
 * resolving shapes, text boxes, tables, pictures, and nested group transforms.
 */
async function traverseShapeTree(
  containerNode: XmlNode | undefined,
  prefix: string,
  zip: JSZip,
  relMap: Map<string, string>,
  slideWidthEmu: number,
  slideHeightEmu: number,
  canvasWidth: number,
  canvasHeight: number,
  groupTransform: GroupTransform,
  inheritance: SlideInheritanceContext,
  elementsOut: SlideElement[],
  titlesOut: { title?: string; subtitle?: string },
  skipPlaceholders = false
): Promise<void> {
  if (!containerNode) return;

  for (let i = 0; i < containerNode.children.length; i++) {
    const child = containerNode.children[i];
    const itemId = `${prefix}-${i}`;

    if (child.tag === 'p:sp' || child.tag === 'p:cxnSp') {
      if (skipPlaceholders) {
        const nvPr = findDescendant(child, 'p:nvPr');
        if (findChild(nvPr, 'p:ph')) continue;
      }
      const { element, detectedTitle, detectedSubtitle } = parseShapeNode(
        child,
        itemId,
        slideWidthEmu,
        slideHeightEmu,
        canvasWidth,
        canvasHeight,
        groupTransform,
        inheritance
      );
      if (!skipPlaceholders) {
        if (detectedTitle && !titlesOut.title) titlesOut.title = detectedTitle;
        if (detectedSubtitle && !titlesOut.subtitle) titlesOut.subtitle = detectedSubtitle;
      }
      if (element) elementsOut.push(element);
    } else if (child.tag === 'p:graphicFrame') {
      if (skipPlaceholders) {
        const nvPr = findDescendant(child, 'p:nvPr');
        if (findChild(nvPr, 'p:ph')) continue;
      }
      const tblEl = parseTableNode(
        child,
        itemId,
        slideWidthEmu,
        slideHeightEmu,
        canvasWidth,
        canvasHeight,
        groupTransform,
        inheritance.themeColors
      );
      if (tblEl) elementsOut.push(tblEl);
    } else if (child.tag === 'p:pic') {
      if (skipPlaceholders) {
        const nvPr = findDescendant(child, 'p:nvPr');
        if (findChild(nvPr, 'p:ph')) continue;
      }
      const picEl = await parsePicNode(
        child,
        itemId,
        zip,
        relMap,
        slideWidthEmu,
        slideHeightEmu,
        canvasWidth,
        canvasHeight,
        groupTransform,
        inheritance
      );
      if (picEl) elementsOut.push(picEl);
    } else if (child.tag === 'p:grpSp') {
      const grpSpPr = findChild(child, 'p:grpSpPr');
      const nextTransform = composeGroupTransform(groupTransform, grpSpPr);
      await traverseShapeTree(
        child,
        itemId,
        zip,
        relMap,
        slideWidthEmu,
        slideHeightEmu,
        canvasWidth,
        canvasHeight,
        nextTransform,
        inheritance,
        elementsOut,
        titlesOut,
        skipPlaceholders
      );
    }
  }
}

/**
 * Parses a single slide XML file (`ppt/slides/slideN.xml`), its `.rels`, its linked
 * `slideLayout`, and its `slideMaster` into an accurately positioned `ParsedSlide`.
 */
export async function parseSingleSlideXml(
  zip: JSZip,
  slidePath: string,
  slideIndex: number,
  slideWidthEmu: number,
  slideHeightEmu: number,
  themeColors: ThemeColorMap = DEFAULT_THEME_COLORS
): Promise<ParsedSlide> {
  // Use exact OpenXML point coordinates (1 pt = 12,700 EMU), matching pptx-preview's native coordinate space
  const canvasWidth = Math.max(100, Math.round(slideWidthEmu / 12700));
  const canvasHeight = Math.max(100, Math.round(slideHeightEmu / 12700));

  const slideFile = zip.file(slidePath);
  const slideXml = slideFile ? await slideFile.async('string') : '';
  const slideTree = parseXmlTree(slideXml);

  // 1. Read slide relationships (images + linked slideLayout)
  const relMap = await readRelsMap(zip, slidePath);

  let layoutPath = '';
  for (const target of relMap.values()) {
    if (target.includes('slideLayouts/slideLayout')) {
      layoutPath = target;
      break;
    }
  }

  // 2. Read slideLayout and linked slideMaster for placeholder inheritance & master/layout shapes
  let layoutXml = '';
  let masterXml = '';
  let masterPath = '';
  let layoutRels = new Map<string, string>();
  let masterRels = new Map<string, string>();

  if (layoutPath) {
    const layoutFile = zip.file(layoutPath);
    if (layoutFile) {
      layoutXml = await layoutFile.async('string');
      layoutRels = await readRelsMap(zip, layoutPath);
      for (const mTarget of layoutRels.values()) {
        if (mTarget.includes('slideMasters/slideMaster')) {
          masterPath = mTarget;
          const masterFile = zip.file(mTarget);
          if (masterFile) {
            masterXml = await masterFile.async('string');
            masterRels = await readRelsMap(zip, mTarget);
          }
          break;
        }
      }
    }
  }

  const masterTree = masterXml ? parseXmlTree(masterXml) : undefined;
  const layoutTree = layoutXml ? parseXmlTree(layoutXml) : undefined;

  const masterExtracted = extractTemplatesFromLayoutOrMaster(masterXml, themeColors);
  const layoutExtracted = extractTemplatesFromLayoutOrMaster(layoutXml, themeColors);

  const placeholdersByIdx = new Map<string, PlaceholderTemplate>();
  const placeholdersByType = new Map<string, PlaceholderTemplate>();

  const allIdxKeys = new Set([
    ...masterExtracted.byIdx.keys(),
    ...layoutExtracted.byIdx.keys(),
  ]);
  for (const idxKey of allIdxKeys) {
    const lTpl = layoutExtracted.byIdx.get(idxKey);
    const mTpl =
      masterExtracted.byIdx.get(idxKey) ||
      (lTpl?.phType ? masterExtracted.byType.get(lTpl.phType) : undefined);
    placeholdersByIdx.set(
      idxKey,
      mergeTemplates(
        lTpl,
        mTpl,
        masterExtracted.masterTitleStyle,
        masterExtracted.masterBodyStyle,
        lTpl?.phType ?? mTpl?.phType
      )
    );
  }

  const allTypeKeys = new Set([
    ...masterExtracted.byType.keys(),
    ...layoutExtracted.byType.keys(),
  ]);
  for (const typeKey of allTypeKeys) {
    placeholdersByType.set(
      typeKey,
      mergeTemplates(
        layoutExtracted.byType.get(typeKey),
        masterExtracted.byType.get(typeKey),
        masterExtracted.masterTitleStyle,
        masterExtracted.masterBodyStyle,
        typeKey
      )
    );
  }

  const slideBgNode = findDescendant(slideTree, 'p:bg');
  const slideBgColor = slideBgNode
    ? extractColorFromNode(slideBgNode, themeColors, '')
    : '';
  const bgColor =
    slideBgColor ||
    layoutExtracted.bgColor ||
    masterExtracted.bgColor ||
    themeColors.bg1 ||
    '#FFFFFF';

  const bgImageUrl =
    (await extractBackgroundImageDataUrl(slideTree, zip, relMap)) ||
    (await extractBackgroundImageDataUrl(layoutTree, zip, layoutRels)) ||
    (await extractBackgroundImageDataUrl(masterTree, zip, masterRels));

  const inheritance: SlideInheritanceContext = {
    placeholdersByIdx,
    placeholdersByType,
    backgroundColor: bgColor,
    themeColors,
  };

  // 3. Traverse Master, Layout (non-placeholders), and Slide `<p:spTree>` in exact z-order
  const elements: SlideElement[] = [];
  const titlesOut: { title?: string; subtitle?: string } = {};

  const sldRoot = findChild(slideTree, 'p:sld');
  const layoutRoot = findChild(layoutTree, 'p:sldLayout');
  const showMasterOnSlide = sldRoot?.attrs.showMasterSp !== '0';
  const showMasterOnLayout = layoutRoot?.attrs.showMasterSp !== '0';

  if (showMasterOnSlide && showMasterOnLayout && masterTree && masterPath) {
    const masterSpTree = findDescendant(masterTree, 'p:spTree');
    await traverseShapeTree(
      masterSpTree,
      `m${slideIndex + 1}`,
      zip,
      masterRels,
      slideWidthEmu,
      slideHeightEmu,
      canvasWidth,
      canvasHeight,
      IDENTITY_TRANSFORM,
      inheritance,
      elements,
      titlesOut,
      true
    );
  }

  if (showMasterOnSlide && layoutTree && layoutPath) {
    const layoutSpTree = findDescendant(layoutTree, 'p:spTree');
    await traverseShapeTree(
      layoutSpTree,
      `l${slideIndex + 1}`,
      zip,
      layoutRels,
      slideWidthEmu,
      slideHeightEmu,
      canvasWidth,
      canvasHeight,
      IDENTITY_TRANSFORM,
      inheritance,
      elements,
      titlesOut,
      true
    );
  }

  const spTree = findDescendant(slideTree, 'p:spTree');
  await traverseShapeTree(
    spTree,
    `s${slideIndex + 1}`,
    zip,
    relMap,
    slideWidthEmu,
    slideHeightEmu,
    canvasWidth,
    canvasHeight,
    IDENTITY_TRANSFORM,
    inheritance,
    elements,
    titlesOut,
    false
  );

  let slideTitle = titlesOut.title || '';
  const slideSubtitle = titlesOut.subtitle || '';

  // Fallback title detection from first text element if no explicit `<p:ph type="title">`
  if (!slideTitle) {
    for (const el of elements) {
      const firstText = el.paragraphs?.[0]?.runs
        .map((r) => r.text)
        .join('')
        .trim();
      if (firstText) {
        slideTitle = firstText.slice(0, 80);
        break;
      }
    }
  }

  if (!slideTitle) {
    slideTitle = `Slide ${slideIndex + 1}`;
  }

  return {
    index: slideIndex,
    slideNumber: slideIndex + 1,
    title: slideTitle,
    subtitle: slideSubtitle || undefined,
    canvasWidth,
    canvasHeight,
    background: {
      color: bgColor,
      imageUrl: bgImageUrl,
      accentBlobColor: PASTEL_ACCENTS[slideIndex % PASTEL_ACCENTS.length],
    },
    elements,
  };
}

export interface ParsePptxCallbacks {
  onFirstSlideReady?: (info: {
    slideCount: number;
    aspectRatio: number;
    firstSlide: ParsedSlide;
  }) => void;
  onSlideParsed?: (info: {
    parsedCount: number;
    totalCount: number;
    slide: ParsedSlide;
  }) => void;
}

/**
 * Full progressive `.pptx` binary parser.
 * Parses Slide 1 first and invokes `onFirstSlideReady` immediately, then streams remaining slides.
 */
export async function parsePptxArrayBuffer(
  buffer: ArrayBuffer,
  fileName = 'presentation.pptx',
  callbacks?: ParsePptxCallbacks
): Promise<{
  slides: ParsedSlide[];
  aspectRatio: number;
}> {
  const headerBytes = new Uint8Array(buffer.slice(0, 8));
  const validation = validatePresentationFile(
    fileName,
    buffer.byteLength,
    headerBytes
  );

  if (!validation.valid) {
    const err = new Error(
      validation.friendlyMessage || 'Invalid .pptx presentation file.'
    ) as Error & { code?: string };
    err.code =
      validation.errorCode === 'LEGACY_PPT' ? 'LEGACY_PPT' : 'CORRUPTED_ZIP';
    throw err;
  }

  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(buffer);
  } catch {
    const err = new Error(
      `"${fileName}" could not be unpacked. The .pptx ZIP archive appears corrupted.`
    ) as Error & { code?: string };
    err.code = 'CORRUPTED_ZIP';
    throw err;
  }

  // Fast-path check for embedded SlideClay rich slide manifest (for generated sample decks)
  const manifestFile = zip.file('ppt/slideclay-manifest.json');
  if (manifestFile) {
    try {
      const jsonStr = await manifestFile.async('string');
      const parsed = JSON.parse(jsonStr) as {
        aspectRatio: number;
        slides: ParsedSlide[];
      };
      if (Array.isArray(parsed.slides) && parsed.slides.length > 0) {
        const ratio = parsed.aspectRatio || 16 / 9;
        const cw = BASE_CANVAS_WIDTH;
        const ch = Math.round(BASE_CANVAS_WIDTH / ratio);
        const normalizedSlides = parsed.slides.map((s) => ({
          ...s,
          canvasWidth: s.canvasWidth || cw,
          canvasHeight: s.canvasHeight || ch,
        }));
        callbacks?.onFirstSlideReady?.({
          slideCount: normalizedSlides.length,
          aspectRatio: ratio,
          firstSlide: normalizedSlides[0],
        });
        for (let i = 0; i < normalizedSlides.length; i++) {
          callbacks?.onSlideParsed?.({
            parsedCount: i + 1,
            totalCount: normalizedSlides.length,
            slide: normalizedSlides[i],
          });
        }
        return {
          slides: normalizedSlides,
          aspectRatio: ratio,
        };
      }
    } catch {
      // Fall through to standard OpenXML slide parsing
    }
  }

  // 1. Read presentation.xml for exact slide dimensions (`<p:sldSz cx="..." cy="..."/>`)
  let slideWidthEmu = DEFAULT_SLIDE_WIDTH_EMU;
  let slideHeightEmu = DEFAULT_SLIDE_HEIGHT_EMU;
  const presFile = zip.file('ppt/presentation.xml');
  if (presFile) {
    const presXml = await presFile.async('string');
    const presTree = parseXmlTree(presXml);
    const sldSz = findDescendant(presTree, 'p:sldSz');
    const cx = Number(sldSz?.attrs.cx ?? NaN);
    const cy = Number(sldSz?.attrs.cy ?? NaN);
    if (!Number.isNaN(cx) && cx > 0 && !Number.isNaN(cy) && cy > 0) {
      slideWidthEmu = cx;
      slideHeightEmu = cy;
    }
  }

  const aspectRatio = Number((slideWidthEmu / slideHeightEmu).toFixed(4));

  // 2. Read theme1.xml for presentation color scheme
  let themeColors = DEFAULT_THEME_COLORS;
  const themeFile = zip.file('ppt/theme/theme1.xml');
  if (themeFile) {
    try {
      const themeXml = await themeFile.async('string');
      themeColors = parseThemeXml(themeXml);
    } catch {
      // Keep default theme colors
    }
  }

  // 3. Locate all slide XML files in `ppt/slides/slide*.xml` sorted by slide number
  const slidePaths = Object.keys(zip.files)
    .filter((p) => /^ppt\/slides\/slide\d+\.xml$/i.test(p))
    .sort((a, b) => {
      const numA = Number((a.match(/slide(\d+)\.xml/i) || [])[1] || 0);
      const numB = Number((b.match(/slide(\d+)\.xml/i) || [])[1] || 0);
      return numA - numB;
    });

  if (slidePaths.length === 0) {
    const err = new Error(
      `"${fileName}" does not contain any PowerPoint slides (ppt/slides/slide*.xml).`
    ) as Error & { code?: string };
    err.code = 'NO_SLIDES';
    throw err;
  }

  const slides: ParsedSlide[] = [];

  // Parse Slide 1 immediately for instant open & thumbnail preview
  const firstSlide = await parseSingleSlideXml(
    zip,
    slidePaths[0],
    0,
    slideWidthEmu,
    slideHeightEmu,
    themeColors
  );
  slides.push(firstSlide);

  callbacks?.onFirstSlideReady?.({
    slideCount: slidePaths.length,
    aspectRatio,
    firstSlide,
  });
  callbacks?.onSlideParsed?.({
    parsedCount: 1,
    totalCount: slidePaths.length,
    slide: firstSlide,
  });

  // Parse remaining slides sequentially and stream progress
  for (let i = 1; i < slidePaths.length; i++) {
    const slide = await parseSingleSlideXml(
      zip,
      slidePaths[i],
      i,
      slideWidthEmu,
      slideHeightEmu,
      themeColors
    );
    slides.push(slide);
    callbacks?.onSlideParsed?.({
      parsedCount: i + 1,
      totalCount: slidePaths.length,
      slide,
    });
  }

  return { slides, aspectRatio };
}
