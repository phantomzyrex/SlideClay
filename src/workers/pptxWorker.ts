/* ============================================================================
 * SLIDECLAY — DEDICATED WEB WORKER FOR NON-BLOCKING PPTX PARSING
 * ============================================================================
 * Unpacks .pptx ZIP archives and parses OpenXML slides off the main thread.
 * Emits `FIRST_SLIDE_READY` immediately after parsing Slide 1 so the library
 * thumbnail and instant viewer open with zero UI freeze, then streams remaining
 * slides via `SLIDE_PROGRESS` and finishes with `DECK_COMPLETE`.
 * ============================================================================ */

import { parsePptxArrayBuffer } from '../utils/pptxEngine';
import {
  WorkerRequestMessage,
  WorkerResponseMessage,
} from '../types/presentation';

const ctx: DedicatedWorkerGlobalScope = self as unknown as DedicatedWorkerGlobalScope;

ctx.onmessage = async (event: MessageEvent<WorkerRequestMessage>) => {
  const msg = event.data;
  if (!msg || msg.type !== 'PARSE_DECK') return;

  const { id, fileName, buffer } = msg;

  try {
    const result = await parsePptxArrayBuffer(buffer, fileName, {
      onFirstSlideReady: ({ slideCount, aspectRatio, firstSlide }) => {
        const response: WorkerResponseMessage = {
          type: 'FIRST_SLIDE_READY',
          id,
          slideCount,
          aspectRatio,
          firstSlide,
        };
        ctx.postMessage(response);
      },
      onSlideParsed: ({ parsedCount, totalCount, slide }) => {
        const response: WorkerResponseMessage = {
          type: 'SLIDE_PROGRESS',
          id,
          parsedCount,
          totalCount,
          slide,
        };
        ctx.postMessage(response);
      },
    });

    const completeMsg: WorkerResponseMessage = {
      type: 'DECK_COMPLETE',
      id,
      slides: result.slides,
      aspectRatio: result.aspectRatio,
    };
    ctx.postMessage(completeMsg);
  } catch (err: unknown) {
    const errorObj = err as Error & { code?: 'LEGACY_PPT' | 'CORRUPTED_ZIP' | 'NO_SLIDES' };
    const errorMsg: WorkerResponseMessage = {
      type: 'PARSE_ERROR',
      id,
      error: errorObj?.message || 'Failed to parse .pptx presentation.',
      code: errorObj?.code,
    };
    ctx.postMessage(errorMsg);
  }
};
