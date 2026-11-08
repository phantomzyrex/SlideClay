export type SlideElementType =
  | 'shape'
  | 'text'
  | 'image'
  | 'table'
  | 'chart';

export interface SlideTextRun {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontSize?: number; // in canvas px (scaled accurately from OpenXML pt/EMU)
  fontFamily?: string;
  color?: string;
  bullet?: boolean;
}

export interface SlideParagraph {
  align?: 'left' | 'center' | 'right' | 'justify';
  runs: SlideTextRun[];
  bullet?: boolean;
  bulletChar?: string;
  level?: number; // 0-based indentation level
  spaceBeforePx?: number;
  spaceAfterPx?: number;
  lineHeight?: number;
}

export interface SlideTableData {
  headers: string[];
  rows: string[][];
  colWidthsPx?: number[];
  accentColor?: string;
}

export interface SlideChartBar {
  label: string;
  value: number;
  color: string;
}

export interface SlideElement {
  id: string;
  type: SlideElementType;
  /** X position in slide canvas coordinates */
  x: number;
  /** Y position in slide canvas coordinates */
  y: number;
  /** Width in slide canvas coordinates */
  w: number;
  /** Height in slide canvas coordinates */
  h: number;
  /** Optional rotation in degrees */
  rotation?: number;
  shapeKind?: 'rect' | 'roundRect' | 'ellipse' | 'line' | 'pill' | 'card';
  fillColor?: string;
  strokeColor?: string;
  strokeWidth?: number;
  borderRadius?: number;
  opacity?: number;
  /** Vertical text alignment inside shape body (`a:bodyPr anchor`) */
  verticalAlign?: 'top' | 'middle' | 'bottom';
  /** Internal text box padding in canvas px (`a:bodyPr lIns, tIns, rIns, bIns`) */
  padding?: {
    top: number;
    right: number;
    bottom: number;
    left: number;
  };
  /** Whether text should wrap inside shape (`a:bodyPr wrap`) */
  wordWrap?: boolean;
  /** Whether shape has clay elevation shadow (used for sample showcase cards) */
  clayShadow?: boolean;
  paragraphs?: SlideParagraph[];
  imageUrl?: string;
  table?: SlideTableData;
  chartBars?: SlideChartBar[];
  chartTitle?: string;
}

export interface ParsedSlide {
  index: number; // 0-based slide index
  slideNumber: number; // 1-based slide number
  title: string;
  subtitle?: string;
  notes?: string;
  /** Logical canvas width (default 960) */
  canvasWidth?: number;
  /** Logical canvas height matching exact slide aspect ratio (e.g. 540 for 16:9, 720 for 4:3) */
  canvasHeight?: number;
  background: {
    color: string;
    secondaryColor?: string;
    accentBlobColor?: string;
    imageUrl?: string;
  };
  elements: SlideElement[];
}

export interface PresentationMetadata {
  id: string;
  name: string;
  originalFileName: string;
  fileSize: number;
  slideCount: number;
  uploadedAt: number;
  lastOpenedAt?: number;
  aspectRatio: number; // e.g., 16/9 = 1.7778, 4/3 = 1.3333
  firstSlide: ParsedSlide;
  accentColor: string;
  isSample?: boolean;
}

export interface PresentationFullData {
  id: string;
  slides: ParsedSlide[];
  rawBuffer?: ArrayBuffer;
}

export type SortField = 'name' | 'date' | 'size';
export type SortOrder = 'asc' | 'desc';

export interface UploadProgressItem {
  id: string;
  fileName: string;
  fileSize: number;
  progress: number; // 0 to 100
  status: 'reading' | 'parsing-first' | 'parsing-slides' | 'saving' | 'done' | 'error';
  errorMessage?: string;
}

export interface ClayToast {
  id: string;
  title: string;
  message: string;
  variant: 'info' | 'success' | 'warning' | 'error';
}

/* Web Worker Message Protocol */
export type WorkerRequestMessage =
  | {
      type: 'PARSE_DECK';
      id: string;
      fileName: string;
      buffer: ArrayBuffer;
    };

export type WorkerResponseMessage =
  | {
      type: 'FIRST_SLIDE_READY';
      id: string;
      slideCount: number;
      aspectRatio: number;
      firstSlide: ParsedSlide;
    }
  | {
      type: 'SLIDE_PROGRESS';
      id: string;
      parsedCount: number;
      totalCount: number;
      slide: ParsedSlide;
    }
  | {
      type: 'DECK_COMPLETE';
      id: string;
      slides: ParsedSlide[];
      aspectRatio: number;
    }
  | {
      type: 'PARSE_ERROR';
      id: string;
      error: string;
      code?: 'LEGACY_PPT' | 'CORRUPTED_ZIP' | 'NO_SLIDES';
    };
