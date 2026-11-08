import {
  ParsedSlide,
  WorkerRequestMessage,
  WorkerResponseMessage,
} from '../types/presentation';
import { parsePptxArrayBuffer } from '../utils/pptxEngine';

interface PendingJob {
  onFirstSlide?: (info: {
    slideCount: number;
    aspectRatio: number;
    firstSlide: ParsedSlide;
  }) => void;
  onProgress?: (info: {
    parsedCount: number;
    totalCount: number;
    slide: ParsedSlide;
  }) => void;
  resolve: (res: { slides: ParsedSlide[]; aspectRatio: number }) => void;
  reject: (err: Error & { code?: string }) => void;
}

class PptxWorkerClient {
  private worker: Worker | null = null;
  private jobs = new Map<string, PendingJob>();

  private getWorker(): Worker | null {
    if (typeof Worker === 'undefined') return null;
    if (!this.worker) {
      try {
        this.worker = new Worker(
          new URL('../workers/pptxWorker.ts', import.meta.url),
          { type: 'module' }
        );
        this.worker.onmessage = (e: MessageEvent<WorkerResponseMessage>) => {
          const data = e.data;
          if (!data || !data.id) return;
          const job = this.jobs.get(data.id);
          if (!job) return;

          if (data.type === 'FIRST_SLIDE_READY') {
            job.onFirstSlide?.({
              slideCount: data.slideCount,
              aspectRatio: data.aspectRatio,
              firstSlide: data.firstSlide,
            });
          } else if (data.type === 'SLIDE_PROGRESS') {
            job.onProgress?.({
              parsedCount: data.parsedCount,
              totalCount: data.totalCount,
              slide: data.slide,
            });
          } else if (data.type === 'DECK_COMPLETE') {
            this.jobs.delete(data.id);
            job.resolve({
              slides: data.slides,
              aspectRatio: data.aspectRatio,
            });
          } else if (data.type === 'PARSE_ERROR') {
            this.jobs.delete(data.id);
            const err = new Error(data.error) as Error & { code?: string };
            err.code = data.code;
            job.reject(err);
          }
        };
      } catch {
        this.worker = null;
      }
    }
    return this.worker;
  }

  public async parseDeck(
    id: string,
    fileName: string,
    buffer: ArrayBuffer,
    callbacks?: {
      onFirstSlide?: (info: {
        slideCount: number;
        aspectRatio: number;
        firstSlide: ParsedSlide;
      }) => void;
      onProgress?: (info: {
        parsedCount: number;
        totalCount: number;
        slide: ParsedSlide;
      }) => void;
    }
  ): Promise<{ slides: ParsedSlide[]; aspectRatio: number }> {
    const worker = this.getWorker();
    if (worker) {
      return new Promise((resolve, reject) => {
        this.jobs.set(id, {
          onFirstSlide: callbacks?.onFirstSlide,
          onProgress: callbacks?.onProgress,
          resolve,
          reject,
        });
        // Clone buffer so original buffer can still be saved to IndexedDB for download/preview
        const cloned = buffer.slice(0);
        const msg: WorkerRequestMessage = {
          type: 'PARSE_DECK',
          id,
          fileName,
          buffer: cloned,
        };
        worker.postMessage(msg, [cloned]);
      });
    }

    // Fallback if Worker is unavailable in current environment
    return parsePptxArrayBuffer(buffer, fileName, {
      onFirstSlideReady: callbacks?.onFirstSlide,
      onSlideParsed: callbacks?.onProgress,
    });
  }
}

export const pptxWorkerClient = new PptxWorkerClient();
