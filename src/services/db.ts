/* ============================================================================
 * SLIDECLAY — INDEXEDDB STORAGE & IN-MEMORY HOT CACHE (via `idb`)
 * ============================================================================
 * Separates lightweight presentation metadata + first-slide thumbnail
 * (`metadata` object store) from full parsed slide decks and raw .pptx binaries
 * (`decks` object store). This guarantees < 50ms library boot times even when
 * dozens of large presentations are stored locally.
 * ============================================================================ */

import { openDB, DBSchema, IDBPDatabase } from 'idb';
import {
  PresentationFullData,
  PresentationMetadata,
  ParsedSlide,
} from '../types/presentation';
import {
  buildValidPptxBuffer,
  SAMPLE_DECK_BLUEPRINTS,
} from '../utils/sampleDecks';
import { evictDeckPreviewer } from './pptxPreviewService';

interface SlideClayDB extends DBSchema {
  metadata: {
    key: string;
    value: PresentationMetadata;
    indexes: { 'by-date': number };
  };
  decks: {
    key: string;
    value: PresentationFullData;
  };
}

const DB_NAME = 'slideclay-db-v1';
const DB_VERSION = 1;

// Hot in-memory cache for zero-latency re-opens during the same session
const hotSlidesCache = new Map<string, ParsedSlide[]>();
const hotBufferCache = new Map<string, ArrayBuffer>();

export function getHotRawBuffer(id: string): ArrayBuffer | undefined {
  return hotBufferCache.get(id);
}

export function setHotRawBuffer(id: string, buffer: ArrayBuffer): void {
  hotBufferCache.set(id, buffer);
}

let dbPromise: Promise<IDBPDatabase<SlideClayDB>> | null = null;

function getDb(): Promise<IDBPDatabase<SlideClayDB>> {
  if (!dbPromise) {
    dbPromise = openDB<SlideClayDB>(DB_NAME, DB_VERSION, {
      upgrade(db) {
        if (!db.objectStoreNames.contains('metadata')) {
          const store = db.createObjectStore('metadata', { keyPath: 'id' });
          store.createIndex('by-date', 'uploadedAt');
        }
        if (!db.objectStoreNames.contains('decks')) {
          db.createObjectStore('decks', { keyPath: 'id' });
        }
      },
    });
  }
  return dbPromise;
}

export async function getAllPresentationMetadata(): Promise<PresentationMetadata[]> {
  const db = await getDb();
  return db.getAll('metadata');
}

export async function savePresentationMetadata(
  meta: PresentationMetadata
): Promise<void> {
  const db = await getDb();
  await db.put('metadata', meta);
}

export async function savePresentationFull(
  meta: PresentationMetadata,
  slides: ParsedSlide[],
  rawBuffer?: ArrayBuffer
): Promise<void> {
  hotSlidesCache.set(meta.id, slides);
  if (rawBuffer) {
    hotBufferCache.set(meta.id, rawBuffer);
  }
  const db = await getDb();
  const tx = db.transaction(['metadata', 'decks'], 'readwrite');
  await Promise.all([
    tx.objectStore('metadata').put(meta),
    tx.objectStore('decks').put({
      id: meta.id,
      slides,
      rawBuffer,
    }),
    tx.done,
  ]);
}

export async function getPresentationFull(
  id: string
): Promise<PresentationFullData | undefined> {
  const cachedSlides = hotSlidesCache.get(id);
  const cachedBuffer = hotBufferCache.get(id);
  if (cachedSlides && cachedSlides.length > 0 && cachedBuffer) {
    return {
      id,
      slides: cachedSlides,
      rawBuffer: cachedBuffer,
    };
  }

  const db = await getDb();
  const data = await db.get('decks', id);
  if (data) {
    hotSlidesCache.set(id, data.slides);
    if (data.rawBuffer) {
      hotBufferCache.set(id, data.rawBuffer);
    }
  }
  return data;
}

export async function renamePresentationInDb(
  id: string,
  newName: string
): Promise<PresentationMetadata | undefined> {
  const db = await getDb();
  const existing = await db.get('metadata', id);
  if (!existing) return undefined;
  const updated: PresentationMetadata = {
    ...existing,
    name: newName.trim() || existing.name,
  };
  await db.put('metadata', updated);
  return updated;
}

export async function touchPresentationOpened(id: string): Promise<void> {
  const db = await getDb();
  const existing = await db.get('metadata', id);
  if (!existing) return;
  existing.lastOpenedAt = Date.now();
  await db.put('metadata', existing);
}

export async function deletePresentationFromDb(id: string): Promise<void> {
  hotSlidesCache.delete(id);
  hotBufferCache.delete(id);
  evictDeckPreviewer(id);
  const db = await getDb();
  const tx = db.transaction(['metadata', 'decks'], 'readwrite');
  await Promise.all([
    tx.objectStore('metadata').delete(id),
    tx.objectStore('decks').delete(id),
    tx.done,
  ]);
}

export async function clearAllPresentationsFromDb(): Promise<void> {
  for (const id of hotBufferCache.keys()) {
    evictDeckPreviewer(id);
  }
  hotSlidesCache.clear();
  hotBufferCache.clear();
  const db = await getDb();
  const tx = db.transaction(['metadata', 'decks'], 'readwrite');
  await Promise.all([
    tx.objectStore('metadata').clear(),
    tx.objectStore('decks').clear(),
    tx.done,
  ]);
}

/**
 * Generates and persists the 3 built-in sample `.pptx` decks into IndexedDB.
 */
export async function seedSampleDecksToDb(): Promise<PresentationMetadata[]> {
  const now = Date.now();
  const created: PresentationMetadata[] = [];

  for (let i = 0; i < SAMPLE_DECK_BLUEPRINTS.length; i++) {
    const bp = SAMPLE_DECK_BLUEPRINTS[i];
    const buffer = await buildValidPptxBuffer(bp.slides);
    const id = `sample-deck-${i + 1}-${now}`;
    const cleanName = bp.fileName.replace(/\.pptx$/i, '');

    const meta: PresentationMetadata = {
      id,
      name: cleanName,
      originalFileName: bp.fileName,
      fileSize: buffer.byteLength,
      slideCount: bp.slides.length,
      uploadedAt: now - bp.uploadedAtOffsetMs,
      aspectRatio: 16 / 9,
      firstSlide: bp.slides[0],
      accentColor: bp.accentColor,
      isSample: true,
    };

    await savePresentationFull(meta, bp.slides, buffer);
    created.push(meta);
  }

  return created;
}
