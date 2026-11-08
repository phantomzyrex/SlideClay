/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import {
  Sun,
  Moon,
  Command,
  Upload,
  CheckCircle2,
  AlertTriangle,
  Info,
  X,
} from 'lucide-react';
import {
  ClayToast,
  ParsedSlide,
  PresentationMetadata,
  SortField,
  SortOrder,
  UploadProgressItem,
} from './types/presentation';
import {
  filterAndSortDecks,
  validatePresentationFile,
} from './utils/deckUtils';
import {
  deletePresentationFromDb,
  getAllPresentationMetadata,
  getHotRawBuffer,
  getPresentationFull,
  renamePresentationInDb,
  savePresentationFull,
  savePresentationMetadata,
  seedSampleDecksToDb,
  setHotRawBuffer,
  touchPresentationOpened,
} from './services/db';
import { loadDeckPreviewer } from './services/pptxPreviewService';
import { pptxWorkerClient } from './services/workerClient';
import { buildValidPptxBuffer } from './utils/sampleDecks';
import { UploadZone } from './components/UploadZone';
import { LibraryView } from './components/LibraryView';
import { ViewerView } from './components/ViewerView';
import { QuickSwitcher } from './components/QuickSwitcher';

const ACCENT_PALETTE = [
  '#B9A7FF',
  '#FFC8A8',
  '#A8EBD0',
  '#A8D4FF',
  '#FFE9A8',
];

const SEEDED_STORAGE_KEY = 'slideclay-seeded-v1';
const THEME_STORAGE_KEY = 'slideclay-theme-v1';

export default function App() {
  // Theme state (Light / Dark Clay)
  const [isDark, setIsDark] = useState<boolean>(() => {
    try {
      return localStorage.getItem(THEME_STORAGE_KEY) === 'dark';
    } catch {
      return false;
    }
  });

  // Library & Active Viewer state
  const [decks, setDecks] = useState<PresentationMetadata[]>([]);
  const [isLoadingLibrary, setIsLoadingLibrary] = useState(true);
  const [searchQuery, setSearchQuery] = useState('');
  const [sortBy, setSortBy] = useState<SortField>('date');
  const [sortOrder, setSortOrder] = useState<SortOrder>('desc');

  // Active open presentation state (Instant Open)
  const [activeDeck, setActiveDeck] = useState<PresentationMetadata | null>(null);
  const [activeSlides, setActiveSlides] = useState<ParsedSlide[]>([]);
  const [activeRawBuffer, setActiveRawBuffer] = useState<ArrayBuffer | undefined>(
    undefined
  );
  const [isParsingRemaining, setIsParsingRemaining] = useState(false);

  // Upload queue & Quick Switcher & Clay Pill Toasts
  const [uploadQueue, setUploadQueue] = useState<UploadProgressItem[]>([]);
  const [quickSwitcherOpen, setQuickSwitcherOpen] = useState(false);
  const [toasts, setToasts] = useState<ClayToast[]>([]);
  const [isGeneratingSample, setIsGeneratingSample] = useState(false);

  const headerFileInputRef = useRef<HTMLInputElement>(null);

  // Sync dark mode class on <html>
  useEffect(() => {
    const root = document.documentElement;
    if (isDark) {
      root.classList.add('dark');
    } else {
      root.classList.remove('dark');
    }
    try {
      localStorage.setItem(THEME_STORAGE_KEY, isDark ? 'dark' : 'light');
    } catch {
      // Ignore storage error
    }
  }, [isDark]);

  const addToast = useCallback(
    (
      title: string,
      message: string,
      variant: ClayToast['variant'] = 'info'
    ) => {
      const id = `toast-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`;
      setToasts((prev) => [...prev.slice(-3), { id, title, message, variant }]);
      setTimeout(() => {
        setToasts((prev) => prev.filter((t) => t.id !== id));
      }, 5200);
    },
    []
  );

  const dismissToast = useCallback((id: string) => {
    setToasts((prev) => prev.filter((t) => t.id !== id));
  }, []);

  // Load library from IndexedDB on mount (< 50ms warm boot)
  useEffect(() => {
    let mounted = true;
    (async () => {
      try {
        const stored = await getAllPresentationMetadata();
        if (stored.length > 0) {
          if (mounted) {
            setDecks(stored);
            setIsLoadingLibrary(false);
          }
          return;
        }

        const alreadySeeded = localStorage.getItem(SEEDED_STORAGE_KEY);
        if (!alreadySeeded) {
          const seeded = await seedSampleDecksToDb();
          localStorage.setItem(SEEDED_STORAGE_KEY, '1');
          if (mounted) {
            setDecks(seeded);
          }
        }
      } catch {
        // Fallback if IndexedDB is blocked
      } finally {
        if (mounted) {
          setIsLoadingLibrary(false);
        }
      }
    })();
    return () => {
      mounted = false;
    };
  }, []);

  // Global Ctrl/Cmd + K listener for Quick Switcher
  useEffect(() => {
    const onKeyDown = (e: KeyboardEvent) => {
      if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'k') {
        e.preventDefault();
        setQuickSwitcherOpen((open) => !open);
      }
    };
    window.addEventListener('keydown', onKeyDown);
    return () => window.removeEventListener('keydown', onKeyDown);
  }, []);

  // Instant Open Handler: Immediately show Slide 1 from metadata, then load/stream remaining slides
  const handleOpenDeck = useCallback(
    async (deck: PresentationMetadata) => {
      setActiveDeck(deck);
      // Render Slide 1 immediately (< 16ms) before awaiting full slide list
      setActiveSlides([deck.firstSlide]);
      setActiveRawBuffer(getHotRawBuffer(deck.id));

      touchPresentationOpened(deck.id).catch(() => {});

      try {
        const full = await getPresentationFull(deck.id);
        if (full) {
          if (full.slides && full.slides.length > 0) {
            setActiveSlides(full.slides);
          }
          if (full.rawBuffer) {
            setActiveRawBuffer(full.rawBuffer);
          }
          setIsParsingRemaining(false);
        }
      } catch {
        addToast(
          'Slide Cache Notice',
          'Showing cached cover slide while re-indexing presentation.',
          'info'
        );
      }
    },
    [addToast]
  );

  // Process multi-file uploads via Web Worker + IndexedDB
  const handleFilesSelected = useCallback(
    async (files: File[]) => {
      if (!files || files.length === 0) return;

      for (let i = 0; i < files.length; i++) {
        const file = files[i];
        const uploadId = `up-${Date.now()}-${i}-${Math.random().toString(36).slice(2, 6)}`;

        // 1. Pre-check extension & size
        const extCheck = validatePresentationFile(file.name, file.size);
        if (!extCheck.valid) {
          setUploadQueue((prev) => [
            {
              id: uploadId,
              fileName: file.name,
              fileSize: file.size,
              progress: 100,
              status: 'error',
              errorMessage: extCheck.friendlyMessage,
            },
            ...prev,
          ]);
          addToast(
            extCheck.friendlyTitle || 'Unsupported File',
            extCheck.friendlyMessage || 'Please upload a valid .pptx file.',
            'error'
          );
          continue;
        }

        // Add item to upload progress queue
        setUploadQueue((prev) => [
          {
            id: uploadId,
            fileName: file.name,
            fileSize: file.size,
            progress: 12,
            status: 'reading',
          },
          ...prev,
        ]);

        try {
          const buffer = await file.arrayBuffer();
          const headerBytes = new Uint8Array(buffer.slice(0, 8));
          const binaryCheck = validatePresentationFile(
            file.name,
            file.size,
            headerBytes
          );

          if (!binaryCheck.valid) {
            setUploadQueue((prev) =>
              prev.map((item) =>
                item.id === uploadId
                  ? {
                      ...item,
                      progress: 100,
                      status: 'error',
                      errorMessage: binaryCheck.friendlyMessage,
                    }
                  : item
              )
            );
            addToast(
              binaryCheck.friendlyTitle || 'Unsupported Format',
              binaryCheck.friendlyMessage ||
                'Legacy .ppt or corrupted file detected.',
              'error'
            );
            continue;
          }

          setUploadQueue((prev) =>
            prev.map((item) =>
              item.id === uploadId
                ? { ...item, progress: 28, status: 'parsing-first' }
                : item
            )
          );

          const deckId = `deck-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;
          const cleanName = file.name.replace(/\.pptx$/i, '');
          const accentColor =
            ACCENT_PALETTE[(decks.length + i) % ACCENT_PALETTE.length];

          // Immediately cache rawBuffer in memory and pre-warm `pptx-preview` model
          setHotRawBuffer(deckId, buffer);
          loadDeckPreviewer(deckId, buffer).catch(() => {});

          let initialMeta: PresentationMetadata | null = null;
          const streamedSlides: ParsedSlide[] = [];

          const parsedResult = await pptxWorkerClient.parseDeck(
            deckId,
            file.name,
            buffer,
            {
              onFirstSlide: ({ slideCount, aspectRatio, firstSlide }) => {
                initialMeta = {
                  id: deckId,
                  name: cleanName,
                  originalFileName: file.name,
                  fileSize: file.size,
                  slideCount,
                  uploadedAt: Date.now(),
                  aspectRatio,
                  firstSlide,
                  accentColor,
                };
                // Immediately add card to Library with Slide 1 thumbnail!
                setDecks((prev) => [initialMeta!, ...prev]);
                savePresentationMetadata(initialMeta).catch(() => {});

                setUploadQueue((prev) =>
                  prev.map((item) =>
                    item.id === uploadId
                      ? { ...item, progress: 50, status: 'parsing-slides' }
                      : item
                  )
                );
              },
              onProgress: ({ parsedCount, totalCount, slide }) => {
                streamedSlides[slide.index] = slide;
                const pct = Math.min(
                  95,
                  50 + Math.round((parsedCount / Math.max(1, totalCount)) * 45)
                );
                setUploadQueue((prev) =>
                  prev.map((item) =>
                    item.id === uploadId ? { ...item, progress: pct } : item
                  )
                );

                // If user already clicked into this deck while background parsing is running, stream slides live!
                setActiveDeck((curr) => {
                  if (curr && curr.id === deckId) {
                    setIsParsingRemaining(parsedCount < totalCount);
                    setActiveSlides([...streamedSlides.filter(Boolean)]);
                  }
                  return curr;
                });
              },
            }
          );

          const finalMeta: PresentationMetadata = initialMeta || {
            id: deckId,
            name: cleanName,
            originalFileName: file.name,
            fileSize: file.size,
            slideCount: parsedResult.slides.length,
            uploadedAt: Date.now(),
            aspectRatio: parsedResult.aspectRatio,
            firstSlide: parsedResult.slides[0],
            accentColor,
          };

          await savePresentationFull(finalMeta, parsedResult.slides, buffer);

          setDecks((prev) => {
            const exists = prev.some((d) => d.id === deckId);
            return exists
              ? prev.map((d) => (d.id === deckId ? finalMeta : d))
              : [finalMeta, ...prev];
          });

          setUploadQueue((prev) =>
            prev.map((item) =>
              item.id === uploadId
                ? { ...item, progress: 100, status: 'done' }
                : item
            )
          );

          addToast(
            'Presentation Ready',
            `“${finalMeta.name}” (${finalMeta.slideCount} slides) cached in IndexedDB.`,
            'success'
          );
        } catch (err: unknown) {
          const message =
            (err as Error)?.message ||
            `Could not parse "${file.name}". Please ensure it is a valid .pptx file.`;
          setUploadQueue((prev) =>
            prev.map((item) =>
              item.id === uploadId
                ? {
                    ...item,
                    progress: 100,
                    status: 'error',
                    errorMessage: message,
                  }
                : item
            )
          );
          addToast('Upload Error', message, 'error');
        }
      }
    },
    [addToast, decks.length]
  );

  const handleRenameDeck = useCallback(
    async (id: string, newName: string) => {
      const updated = await renamePresentationInDb(id, newName);
      if (updated) {
        setDecks((prev) => prev.map((d) => (d.id === id ? updated : d)));
        setActiveDeck((curr) => (curr && curr.id === id ? updated : curr));
        addToast('Deck Renamed', `Renamed to “${updated.name}”.`, 'success');
      }
    },
    [addToast]
  );

  const handleDeleteDeck = useCallback(
    async (deck: PresentationMetadata) => {
      await deletePresentationFromDb(deck.id);
      setDecks((prev) => prev.filter((d) => d.id !== deck.id));
      if (activeDeck?.id === deck.id) {
        setActiveDeck(null);
      }
      addToast(
        'Presentation Deleted',
        `Removed “${deck.name}” from local storage.`,
        'info'
      );
    },
    [activeDeck, addToast]
  );

  const handleDownloadDeck = useCallback(
    async (deck: PresentationMetadata) => {
      try {
        const full = await getPresentationFull(deck.id);
        let buffer = full?.rawBuffer;
        if (!buffer) {
          const slides = full?.slides || [deck.firstSlide];
          buffer = await buildValidPptxBuffer(slides, deck.aspectRatio);
        }
        const blob = new Blob([buffer], {
          type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = deck.name.endsWith('.pptx')
          ? deck.name
          : `${deck.name}.pptx`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        addToast(
          'Download Started',
          `Downloading “${deck.name}.pptx”.`,
          'success'
        );
      } catch {
        addToast('Download Failed', 'Could not export .pptx archive.', 'error');
      }
    },
    [addToast]
  );

  const handleRestoreSamples = useCallback(async () => {
    setIsGeneratingSample(true);
    try {
      const seeded = await seedSampleDecksToDb();
      const all = await getAllPresentationMetadata();
      setDecks(all.length > 0 ? all : seeded);
      addToast(
        'Sample Decks Restored',
        'Added 3 claymorphic .pptx presentations to your library.',
        'success'
      );
    } finally {
      setIsGeneratingSample(false);
    }
  }, [addToast]);

  const filteredDecks = useMemo(
    () => filterAndSortDecks(decks, searchQuery, sortBy, sortOrder),
    [decks, searchQuery, sortBy, sortOrder]
  );

  return (
    <div className="min-h-screen relative flex flex-col selection:bg-[#B9A7FF] selection:text-[#1F1638]">
      {/* Subtle Blurred Floating Clay Background Blobs */}
      <div
        aria-hidden="true"
        className="fixed inset-0 pointer-events-none overflow-hidden -z-10"
      >
        <div
          className="absolute -top-24 -left-24 w-96 h-96 rounded-full opacity-45 dark:opacity-20 blur-3xl"
          style={{ backgroundColor: '#B9A7FF' }}
        />
        <div
          className="absolute top-1/3 -right-28 w-80 h-80 rounded-full opacity-40 dark:opacity-15 blur-3xl"
          style={{ backgroundColor: '#FFC8A8' }}
        />
        <div
          className="absolute -bottom-28 left-1/4 w-96 h-96 rounded-full opacity-40 dark:opacity-15 blur-3xl"
          style={{ backgroundColor: '#A8EBD0' }}
        />
      </div>

      {/* Top Navigation Bar (Strict 3-Zone Contract) */}
      <header className="w-full max-w-[1440px] mx-auto px-4 sm:px-8 pt-5 pb-3 flex items-center justify-between gap-8">
        {/* Zone 1: Single-line Brand Wordmark */}
        <a
          href="#library"
          onClick={(e) => {
            e.preventDefault();
            setActiveDeck(null);
          }}
          className="text-2xl md:text-3xl font-bold font-display tracking-tight text-[var(--clay-text)] whitespace-nowrap shrink-0"
        >
          SlideClay
        </a>

        {/* Zone 2: Concise Navigation & Utility Links */}
        <nav
          aria-label="Primary Navigation"
          className="hidden md:flex items-center gap-6 text-sm font-bold text-[var(--clay-text-muted)]"
        >
          <button
            type="button"
            onClick={() => setActiveDeck(null)}
            className={`hover:text-[var(--clay-text)] transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
              !activeDeck ? 'text-[#7C63E6] dark:text-[#B9A7FF] underline underline-offset-4' : ''
            }`}
          >
            Library ({decks.length})
          </button>

          <button
            type="button"
            onClick={() => setQuickSwitcherOpen(true)}
            className="hover:text-[var(--clay-text)] transition-colors whitespace-nowrap shrink-0 flex items-center gap-1.5 cursor-pointer"
          >
            <Command className="w-4 h-4 text-[#7C63E6] dark:text-[#B9A7FF]" />
            <span>Quick Switch (⌘K)</span>
          </button>

          <button
            type="button"
            onClick={() => setIsDark((d) => !d)}
            aria-label={
              isDark ? 'Switch to Light Clay theme' : 'Switch to Dark Clay theme'
            }
            className="hover:text-[var(--clay-text)] transition-colors whitespace-nowrap shrink-0 flex items-center gap-1.5 cursor-pointer"
          >
            {isDark ? (
              <>
                <Sun className="w-4 h-4 text-[#FFE9A8]" />
                <span>Light Clay</span>
              </>
            ) : (
              <>
                <Moon className="w-4 h-4 text-[#7C63E6]" />
                <span>Dark Clay</span>
              </>
            )}
          </button>
        </nav>

        {/* Zone 3: Primary Action */}
        <div className="flex items-center gap-2.5 shrink-0">
          {/* Mobile Theme Toggle */}
          <button
            type="button"
            onClick={() => setIsDark((d) => !d)}
            aria-label={
              isDark
                ? 'Switch mobile view to Light Clay theme'
                : 'Switch mobile view to Dark Clay theme'
            }
            className="md:hidden clay-btn p-2.5"
          >
            {isDark ? (
              <Sun className="w-4 h-4 text-[#FFE9A8]" />
            ) : (
              <Moon className="w-4 h-4 text-[#7C63E6]" />
            )}
          </button>

          <input
            ref={headerFileInputRef}
            type="file"
            accept=".pptx,.ppt"
            multiple
            onChange={(e) => {
              if (e.target.files && e.target.files.length > 0) {
                handleFilesSelected(Array.from(e.target.files));
                e.target.value = '';
              }
            }}
            className="hidden"
            aria-label="Upload PowerPoint files from header"
          />

          <button
            type="button"
            onClick={() => headerFileInputRef.current?.click()}
            className="clay-btn-lavender px-5 py-2.5 text-xs md:text-sm flex items-center gap-2 whitespace-nowrap shrink-0"
          >
            <Upload className="w-4 h-4 stroke-[2.5]" />
            <span>Upload .pptx</span>
          </button>
        </div>
      </header>

      {/* Main Content Container */}
      <main className="flex-1 w-full max-w-[1440px] mx-auto px-4 sm:px-8 py-4">
        <AnimatePresence mode="wait">
          {activeDeck ? (
            <motion.div
              key={`viewer-${activeDeck.id}`}
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
              transition={{ duration: 0.18 }}
            >
              <ViewerView
                deck={activeDeck}
                slides={activeSlides}
                isParsingRemaining={isParsingRemaining}
                rawBuffer={activeRawBuffer}
                onBackToLibrary={() => setActiveDeck(null)}
                onOpenQuickSwitcher={() => setQuickSwitcherOpen(true)}
              />
            </motion.div>
          ) : (
            <motion.div
              key="library-view"
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
              transition={{ duration: 0.18 }}
            >
              <UploadZone
                onFilesSelected={handleFilesSelected}
                uploadQueue={uploadQueue}
                onClearFinishedUploads={() =>
                  setUploadQueue((prev) =>
                    prev.filter(
                      (i) => i.status !== 'done' && i.status !== 'error'
                    )
                  )
                }
                onGenerateSampleDeck={handleRestoreSamples}
                isGeneratingSample={isGeneratingSample}
              />

              {isLoadingLibrary ? (
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-7">
                  {[1, 2, 3].map((n) => (
                    <div
                      key={n}
                      className="clay-card p-5 h-80 animate-pulse flex flex-col justify-between"
                    >
                      <div className="rounded-[20px] bg-black/5 dark:bg-white/5 h-44 w-full" />
                      <div className="space-y-2">
                        <div className="h-5 w-2/3 rounded-full bg-black/5 dark:bg-white/5" />
                        <div className="h-3.5 w-1/2 rounded-full bg-black/5 dark:bg-white/5" />
                      </div>
                    </div>
                  ))}
                </div>
              ) : (
                <LibraryView
                  decks={filteredDecks}
                  totalDeckCount={decks.length}
                  searchQuery={searchQuery}
                  onSearchChange={setSearchQuery}
                  sortBy={sortBy}
                  sortOrder={sortOrder}
                  onSortChange={(field) => setSortBy(field)}
                  onToggleSortOrder={() =>
                    setSortOrder((o) => (o === 'asc' ? 'desc' : 'asc'))
                  }
                  onOpenDeck={handleOpenDeck}
                  onRenameDeck={handleRenameDeck}
                  onDeleteDeck={handleDeleteDeck}
                  onDownloadDeck={handleDownloadDeck}
                  onRestoreSamples={handleRestoreSamples}
                />
              )}
            </motion.div>
          )}
        </AnimatePresence>
      </main>

      {/* Quick Switcher Modal (Ctrl/Cmd + K) */}
      <QuickSwitcher
        isOpen={quickSwitcherOpen}
        decks={decks}
        activeDeckId={activeDeck?.id}
        onClose={() => setQuickSwitcherOpen(false)}
        onSelectDeck={(deck) => handleOpenDeck(deck)}
      />

      {/* Clay Pill Toast Notifications */}
      <div
        aria-live="polite"
        className="fixed bottom-5 right-5 z-50 flex flex-col gap-2.5 max-w-sm w-full pointer-events-none"
      >
        <AnimatePresence>
          {toasts.map((toast) => (
            <motion.div
              key={toast.id}
              initial={{ opacity: 0, y: 20, scale: 0.92 }}
              animate={{ opacity: 1, y: 0, scale: 1 }}
              exit={{ opacity: 0, y: 12, scale: 0.92 }}
              className="pointer-events-auto clay-card px-4 py-3 rounded-[24px] flex items-start gap-3 shadow-xl"
            >
              <span
                className="clay-icon-circle w-8 h-8 shrink-0 mt-0.5"
                style={{
                  backgroundColor:
                    toast.variant === 'error'
                      ? '#FFB4C8'
                      : toast.variant === 'success'
                        ? '#A8EBD0'
                        : '#B9A7FF',
                }}
              >
                {toast.variant === 'error' ? (
                  <AlertTriangle className="w-4 h-4 text-[#591227]" />
                ) : toast.variant === 'success' ? (
                  <CheckCircle2 className="w-4 h-4 text-[#113D2B]" />
                ) : (
                  <Info className="w-4 h-4 text-[#231942]" />
                )}
              </span>

              <div className="flex-1 min-w-0">
                <div className="text-xs md:text-sm font-bold text-[var(--clay-text)]">
                  {toast.title}
                </div>
                <div className="text-xs text-[var(--clay-text-muted)] leading-snug">
                  {toast.message}
                </div>
              </div>

              <button
                type="button"
                onClick={() => dismissToast(toast.id)}
                aria-label="Dismiss notification"
                className="p-1 rounded-full hover:bg-black/5 dark:hover:bg-white/10 text-[var(--clay-text-muted)] cursor-pointer"
              >
                <X className="w-3.5 h-3.5" />
              </button>
            </motion.div>
          ))}
        </AnimatePresence>
      </div>
    </div>
  );
}
