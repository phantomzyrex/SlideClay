import {
  PresentationMetadata,
  SortField,
  SortOrder,
} from '../types/presentation';

export interface FileValidationResult {
  valid: boolean;
  errorCode?: 'LEGACY_PPT' | 'UNSUPPORTED_EXT' | 'EMPTY_FILE' | 'CORRUPTED_HEADER';
  friendlyTitle?: string;
  friendlyMessage?: string;
}

/**
 * Validates file extension, size, and optional binary header (magic bytes)
 * to distinguish valid .pptx archives from legacy .ppt OLE2 files or corrupted files.
 */
export function validatePresentationFile(
  fileName: string,
  fileSize: number,
  headerBytes?: Uint8Array
): FileValidationResult {
  const lower = fileName.trim().toLowerCase();

  if (lower.endsWith('.ppt')) {
    return {
      valid: false,
      errorCode: 'LEGACY_PPT',
      friendlyTitle: 'Legacy .ppt Format Not Supported',
      friendlyMessage: `"${fileName}" uses the legacy PowerPoint 97-2003 (.ppt) format. Please re-save or convert it to modern .pptx in PowerPoint, Keynote, or Google Slides.`,
    };
  }

  if (!lower.endsWith('.pptx')) {
    return {
      valid: false,
      errorCode: 'UNSUPPORTED_EXT',
      friendlyTitle: 'Unsupported File Format',
      friendlyMessage: `"${fileName}" is not a PowerPoint .pptx presentation. Please drop a valid .pptx file.`,
    };
  }

  if (fileSize <= 0) {
    return {
      valid: false,
      errorCode: 'EMPTY_FILE',
      friendlyTitle: 'Empty Presentation File',
      friendlyMessage: `"${fileName}" is empty (0 bytes) and cannot be opened.`,
    };
  }

  if (headerBytes && headerBytes.length >= 4) {
    // Check for OLE2 Compound Document magic bytes: D0 CF 11 E0 (legacy .ppt renamed to .pptx)
    if (
      headerBytes[0] === 0xd0 &&
      headerBytes[1] === 0xcf &&
      headerBytes[2] === 0x11 &&
      headerBytes[3] === 0xe0
    ) {
      return {
        valid: false,
        errorCode: 'LEGACY_PPT',
        friendlyTitle: 'Legacy .ppt Detected Inside File',
        friendlyMessage: `"${fileName}" appears to be a legacy binary .ppt file renamed to .pptx. Please convert it to a modern OpenXML .pptx file.`,
      };
    }

    // Check for ZIP magic bytes: 50 4B (PK)
    if (headerBytes[0] !== 0x50 || headerBytes[1] !== 0x4b) {
      return {
        valid: false,
        errorCode: 'CORRUPTED_HEADER',
        friendlyTitle: 'Corrupted .pptx Archive',
        friendlyMessage: `"${fileName}" is not a valid OpenXML ZIP archive or is corrupted.`,
      };
    }
  }

  return { valid: true };
}

/**
 * Formats byte count into human-readable KB / MB string with tabular precision.
 */
export function formatFileSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  const kb = bytes / 1024;
  if (kb < 1024) return `${kb.toFixed(1)} KB`;
  const mb = kb / 1024;
  return `${mb.toFixed(2)} MB`;
}

/**
 * Formats a timestamp into a clean date string.
 */
export function formatUploadDate(timestamp: number): string {
  const d = new Date(timestamp);
  return d.toLocaleDateString('en-US', {
    month: 'short',
    day: 'numeric',
    year: 'numeric',
  });
}

/**
 * Returns a human-friendly aspect ratio label (e.g. "16:9", "4:3", "16:10", "A4").
 */
export function formatAspectRatioLabel(ratio: number): string {
  if (!ratio || ratio <= 0) return '16:9';
  if (Math.abs(ratio - 16 / 9) < 0.04) return '16:9';
  if (Math.abs(ratio - 4 / 3) < 0.04) return '4:3';
  if (Math.abs(ratio - 16 / 10) < 0.04) return '16:10';
  if (Math.abs(ratio - 1) < 0.03) return '1:1';
  if (Math.abs(ratio - 1.414) < 0.04) return 'A4 Landscape';
  if (Math.abs(ratio - 0.707) < 0.04) return 'A4 Portrait';
  if (Math.abs(ratio - 9 / 16) < 0.04) return '9:16';
  return `${ratio.toFixed(2)}:1`;
}

/**
 * Computes exact fitted stage width and height (in px) so a slide of `aspectRatio`
 * fits 100% inside `(availableWidth, availableHeight)` without clipping or distortion,
 * scaled by `zoom`.
 */
export function computeFittedStageSize(
  availableWidth: number,
  availableHeight: number,
  aspectRatio: number,
  zoom = 1
): { width: number; height: number } {
  const safeRatio = aspectRatio && aspectRatio > 0 ? aspectRatio : 16 / 9;
  const maxW = Math.max(160, availableWidth);
  const maxH = Math.max(120, availableHeight);

  let baseW = maxW;
  let baseH = baseW / safeRatio;

  if (baseH > maxH) {
    baseH = maxH;
    baseW = baseH * safeRatio;
  }

  return {
    width: Math.round(baseW * zoom),
    height: Math.round(baseH * zoom),
  };
}

/**
 * Filters and sorts presentation decks in the library.
 */
export function filterAndSortDecks(
  decks: PresentationMetadata[],
  searchQuery: string,
  sortBy: SortField,
  sortOrder: SortOrder
): PresentationMetadata[] {
  const q = searchQuery.trim().toLowerCase();

  const filtered = q
    ? decks.filter(
        (d) =>
          d.name.toLowerCase().includes(q) ||
          d.originalFileName.toLowerCase().includes(q) ||
          d.firstSlide.title.toLowerCase().includes(q)
      )
    : [...decks];

  filtered.sort((a, b) => {
    let cmp = 0;
    if (sortBy === 'name') {
      cmp = a.name.localeCompare(b.name, undefined, { sensitivity: 'base' });
    } else if (sortBy === 'date') {
      cmp = a.uploadedAt - b.uploadedAt;
    } else if (sortBy === 'size') {
      cmp = a.fileSize - b.fileSize;
    }
    return sortOrder === 'asc' ? cmp : -cmp;
  });

  return filtered;
}

/**
 * Converts PowerPoint English Metric Units (EMUs) into canvas pixels with floating-point precision
 * (rounded to 2 decimal places so text boxes and table columns never drift or shift).
 */
export function emuToCanvasCoord(
  emuValue: number,
  slideMaxEmu: number,
  canvasMaxPx: number
): number {
  if (!slideMaxEmu || slideMaxEmu <= 0) return 0;
  return Number(((emuValue / slideMaxEmu) * canvasMaxPx).toFixed(2));
}

/**
 * Computes the virtualized window of items to render in a scrollable thumbnail list.
 */
export function computeVirtualWindow(
  scrollTop: number,
  viewportHeight: number,
  itemHeight: number,
  totalCount: number,
  overscan = 4
): { startIndex: number; endIndex: number; offsetY: number; totalHeight: number } {
  const totalHeight = totalCount * itemHeight;
  if (totalCount === 0 || itemHeight <= 0) {
    return { startIndex: 0, endIndex: -1, offsetY: 0, totalHeight: 0 };
  }

  const rawStart = Math.floor(Math.max(0, scrollTop) / itemHeight);
  const visibleCount = Math.ceil(viewportHeight / itemHeight);

  const startIndex = Math.max(0, rawStart - overscan);
  const endIndex = Math.min(totalCount - 1, rawStart + visibleCount + overscan);
  const offsetY = startIndex * itemHeight;

  return {
    startIndex,
    endIndex,
    offsetY,
    totalHeight,
  };
}

/**
 * Computes the set of nearby slide indices to keep warm in the active viewer memory cache.
 * Always includes slide 0 (cover), currentSlideIndex, and +/- radius slides for zero-lag navigation.
 */
export function computeHotSlideIndices(
  currentSlideIndex: number,
  totalSlides: number,
  radius = 3
): number[] {
  if (totalSlides <= 0) return [];
  const set = new Set<number>();
  set.add(0);
  const start = Math.max(0, currentSlideIndex - radius);
  const end = Math.min(totalSlides - 1, currentSlideIndex + radius);
  for (let i = start; i <= end; i++) {
    set.add(i);
  }
  return Array.from(set).sort((a, b) => a - b);
}
