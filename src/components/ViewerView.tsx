import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import {
  ArrowLeft,
  ChevronLeft,
  ChevronRight,
  Maximize2,
  Minimize2,
  ZoomIn,
  ZoomOut,
  RotateCcw,
  PanelLeft,
  Loader2,
  Keyboard,
  Command,
  Layers,
} from 'lucide-react';
import {
  ParsedSlide,
  PresentationMetadata,
} from '../types/presentation';
import {
  computeFittedStageSize,
  computeHotSlideIndices,
  computeVirtualWindow,
  formatAspectRatioLabel,
} from '../utils/deckUtils';
import { SlideCanvas } from './SlideCanvas';

interface ViewerViewProps {
  deck: PresentationMetadata;
  slides: ParsedSlide[];
  isParsingRemaining: boolean;
  rawBuffer?: ArrayBuffer;
  onBackToLibrary: () => void;
  onOpenQuickSwitcher: () => void;
}

const THUMB_ITEM_HEIGHT = 152;

export const ViewerView: React.FC<ViewerViewProps> = ({
  deck,
  slides,
  isParsingRemaining,
  rawBuffer,
  onBackToLibrary,
  onOpenQuickSwitcher,
}) => {
  const [currentIndex, setCurrentIndex] = useState(0);
  const [zoom, setZoom] = useState(1);
  const [isPresentationMode, setIsPresentationMode] = useState(false);
  const [mobileDrawerOpen, setMobileDrawerOpen] = useState(false);
  const [showShortcutsModal, setShowShortcutsModal] = useState(false);

  // Virtualized sidebar state
  const sidebarScrollRef = useRef<HTMLDivElement>(null);
  const viewerRootRef = useRef<HTMLDivElement>(null);
  const stageBezelRef = useRef<HTMLDivElement>(null);
  const [scrollTop, setScrollTop] = useState(0);
  const [sidebarHeight, setSidebarHeight] = useState(640);
  const [bezelBounds, setBezelBounds] = useState<{ width: number; height: number }>({
    width: 880,
    height: 540,
  });

  // Touch swipe tracking for mobile slide navigation
  const touchStartXRef = useRef<number | null>(null);

  const totalSlides = Math.max(slides.length, deck.slideCount || 1);
  const activeSlide: ParsedSlide =
    slides[currentIndex] || slides[0] || deck.firstSlide;

  // Determine exact slide aspect ratio from active slide canvas dimensions or deck metadata
  const resolvedAspectRatio = useMemo(() => {
    if (activeSlide?.canvasWidth && activeSlide?.canvasHeight) {
      return activeSlide.canvasWidth / activeSlide.canvasHeight;
    }
    return deck.aspectRatio && deck.aspectRatio > 0 ? deck.aspectRatio : 16 / 9;
  }, [activeSlide?.canvasWidth, activeSlide?.canvasHeight, deck.aspectRatio]);

  // Reset slide index when switching decks
  useEffect(() => {
    setCurrentIndex(0);
    setZoom(1);
  }, [deck.id]);

  // Measure sidebar viewport height for virtualization
  useEffect(() => {
    const el = sidebarScrollRef.current;
    if (!el) return;
    const updateHeight = () => {
      setSidebarHeight(el.clientHeight || 640);
    };
    updateHeight();
    const ro = new ResizeObserver(updateHeight);
    ro.observe(el);
    return () => ro.disconnect();
  }, []);

  // Measure available stage bezel inner area so any aspect ratio (16:9, 4:3, A4) fits 100% without clipping
  useEffect(() => {
    const el = stageBezelRef.current;
    if (!el) return;
    const updateBounds = () => {
      const pad = isPresentationMode ? 48 : 44;
      const w = Math.max(240, el.clientWidth - pad);
      const h = Math.max(180, el.clientHeight - pad);
      setBezelBounds((prev) =>
        Math.abs(prev.width - w) > 2 || Math.abs(prev.height - h) > 2
          ? { width: w, height: h }
          : prev
      );
    };
    updateBounds();
    const ro = new ResizeObserver(updateBounds);
    ro.observe(el);
    return () => ro.disconnect();
  }, [isPresentationMode]);

  const fittedStageSize = useMemo(
    () =>
      computeFittedStageSize(
        bezelBounds.width,
        bezelBounds.height,
        resolvedAspectRatio,
        zoom
      ),
    [bezelBounds.width, bezelBounds.height, resolvedAspectRatio, zoom]
  );

  // Auto-scroll virtualized sidebar so current slide stays visible
  useEffect(() => {
    const el = sidebarScrollRef.current;
    if (!el) return;
    const targetTop = currentIndex * THUMB_ITEM_HEIGHT;
    const targetBottom = targetTop + THUMB_ITEM_HEIGHT;
    if (targetTop < el.scrollTop) {
      el.scrollTo({ top: Math.max(0, targetTop - 24), behavior: 'smooth' });
    } else if (targetBottom > el.scrollTop + el.clientHeight) {
      el.scrollTo({
        top: targetBottom - el.clientHeight + 24,
        behavior: 'smooth',
      });
    }
  }, [currentIndex]);

  const goToSlide = useCallback(
    (index: number) => {
      const maxAvailable = Math.max(0, slides.length - 1);
      const clamped = Math.max(0, Math.min(maxAvailable, index));
      setCurrentIndex(clamped);
    },
    [slides.length]
  );

  const goNext = useCallback(() => {
    goToSlide(currentIndex + 1);
  }, [currentIndex, goToSlide]);

  const goPrev = useCallback(() => {
    goToSlide(currentIndex - 1);
  }, [currentIndex, goToSlide]);

  const toggleFullscreen = useCallback(async () => {
    if (!isPresentationMode) {
      setIsPresentationMode(true);
      try {
        if (
          document.fullscreenElement === null &&
          viewerRootRef.current?.requestFullscreen
        ) {
          await viewerRootRef.current.requestFullscreen();
        }
      } catch {
        // Fallback to in-window distraction-free presentation mode if iframe blocks native fullscreen
      }
    } else {
      setIsPresentationMode(false);
      try {
        if (document.fullscreenElement) {
          await document.exitFullscreen();
        }
      } catch {
        // Ignore fullscreen exit errors
      }
    }
  }, [isPresentationMode]);

  // Sync fullscreen exit via Esc key
  useEffect(() => {
    const onFsChange = () => {
      if (!document.fullscreenElement && isPresentationMode) {
        setIsPresentationMode(false);
      }
    };
    document.addEventListener('fullscreenchange', onFsChange);
    return () => document.removeEventListener('fullscreenchange', onFsChange);
  }, [isPresentationMode]);

  // Keyboard shortcuts: ← →, Home/End, F, Esc, +/-/0
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      const tag = (e.target as HTMLElement)?.tagName;
      if (tag === 'INPUT' || tag === 'TEXTAREA') return;
      if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'k') return;

      if (e.key === 'ArrowRight' || e.key === 'PageDown' || e.key === ' ') {
        e.preventDefault();
        goNext();
      } else if (e.key === 'ArrowLeft' || e.key === 'PageUp') {
        e.preventDefault();
        goPrev();
      } else if (e.key === 'Home') {
        e.preventDefault();
        goToSlide(0);
      } else if (e.key === 'End') {
        e.preventDefault();
        goToSlide(slides.length - 1);
      } else if (e.key === 'f' || e.key === 'F') {
        e.preventDefault();
        toggleFullscreen();
      } else if (e.key === 'Escape') {
        if (showShortcutsModal) {
          setShowShortcutsModal(false);
        } else if (isPresentationMode) {
          setIsPresentationMode(false);
          if (document.fullscreenElement) {
            document.exitFullscreen().catch(() => {});
          }
        } else {
          onBackToLibrary();
        }
      } else if (e.key === '=' || e.key === '+') {
        e.preventDefault();
        setZoom((z) => Math.min(2, Number((z + 0.15).toFixed(2))));
      } else if (e.key === '-') {
        e.preventDefault();
        setZoom((z) => Math.max(0.5, Number((z - 0.15).toFixed(2))));
      } else if (e.key === '0') {
        e.preventDefault();
        setZoom(1);
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [
    goNext,
    goPrev,
    goToSlide,
    slides.length,
    toggleFullscreen,
    isPresentationMode,
    showShortcutsModal,
    onBackToLibrary,
  ]);

  // Virtualized thumbnail range
  const virtualWindow = useMemo(
    () =>
      computeVirtualWindow(
        scrollTop,
        sidebarHeight,
        THUMB_ITEM_HEIGHT,
        slides.length,
        4
      ),
    [scrollTop, sidebarHeight, slides.length]
  );

  const visibleThumbnailSlides = useMemo(() => {
    if (virtualWindow.endIndex < virtualWindow.startIndex) return [];
    return slides.slice(virtualWindow.startIndex, virtualWindow.endIndex + 1);
  }, [slides, virtualWindow.startIndex, virtualWindow.endIndex]);

  // Nearby hot slides kept preloaded in memory (N-1, N+1) for zero-lag transitions
  const hotPreloadSlides = useMemo(() => {
    const indices = computeHotSlideIndices(currentIndex, slides.length, 1).filter(
      (i) => i !== currentIndex
    );
    return indices.map((i) => slides[i]).filter(Boolean);
  }, [currentIndex, slides]);

  const progressPercentage =
    totalSlides > 1
      ? Math.round(((currentIndex + 1) / totalSlides) * 100)
      : 100;

  const handleTouchStart = (e: React.TouchEvent) => {
    touchStartXRef.current = e.touches[0]?.clientX ?? null;
  };

  const handleTouchEnd = (e: React.TouchEvent) => {
    if (touchStartXRef.current === null) return;
    const endX = e.changedTouches[0]?.clientX ?? touchStartXRef.current;
    const deltaX = endX - touchStartXRef.current;
    touchStartXRef.current = null;
    if (Math.abs(deltaX) > 48) {
      if (deltaX < 0) goNext();
      else goPrev();
    }
  };

  return (
    <div
      ref={viewerRootRef}
      className={`w-full flex flex-col ${
        isPresentationMode
          ? 'fixed inset-0 z-50 bg-[#141022] p-3 md:p-6 h-screen'
          : 'h-[calc(100vh-108px)] min-h-[580px]'
      }`}
    >
      {/* Viewer Top Control Bar */}
      {!isPresentationMode && (
        <div className="clay-card px-4 py-3 mb-4 flex flex-wrap items-center justify-between gap-3 shrink-0">
          {/* Left: Back to Library + Deck Name */}
          <div className="flex items-center gap-3 min-w-0">
            <button
              type="button"
              onClick={onBackToLibrary}
              aria-label="Back to Library"
              className="clay-btn px-4 py-2 text-xs md:text-sm flex items-center gap-2 whitespace-nowrap shrink-0"
            >
              <ArrowLeft className="w-4 h-4 stroke-[2.5]" />
              <span>Library</span>
            </button>

            <div className="min-w-0">
              <h2 className="text-base md:text-lg font-bold font-display text-[var(--clay-text)] truncate">
                {deck.name}
              </h2>
              <div className="flex items-center gap-2 text-xs text-[var(--clay-text-muted)] font-mono-num">
                <span>
                  Slide {currentIndex + 1} of {totalSlides}
                </span>
                <span aria-hidden="true">·</span>
                <span>{formatAspectRatioLabel(resolvedAspectRatio)}</span>
                {isParsingRemaining && (
                  <>
                    <span aria-hidden="true">·</span>
                    <span className="inline-flex items-center gap-1 text-[#7C63E6] dark:text-[#B9A7FF]">
                      <Loader2 className="w-3 h-3 animate-spin" />
                      Indexing slides ({slides.length}/{totalSlides})
                    </span>
                  </>
                )}
              </div>
            </div>
          </div>

          {/* Center: Zoom & Fit Controls */}
          <div className="flex items-center gap-2">
            <div className="clay-inset px-2 py-1 flex items-center gap-1">
              <button
                type="button"
                onClick={() =>
                  setZoom((z) => Math.max(0.5, Number((z - 0.15).toFixed(2))))
                }
                aria-label="Zoom out"
                title="Zoom Out (-)"
                className="p-1.5 rounded-full hover:bg-black/5 dark:hover:bg-white/10 text-[var(--clay-text)] cursor-pointer"
              >
                <ZoomOut className="w-4 h-4" />
              </button>
              <span className="px-2 text-xs font-bold font-mono-num text-[var(--clay-text)] min-w-[52px] text-center">
                {Math.round(zoom * 100)}%
              </span>
              <button
                type="button"
                onClick={() =>
                  setZoom((z) => Math.min(2, Number((z + 0.15).toFixed(2))))
                }
                aria-label="Zoom in"
                title="Zoom In (+)"
                className="p-1.5 rounded-full hover:bg-black/5 dark:hover:bg-white/10 text-[var(--clay-text)] cursor-pointer"
              >
                <ZoomIn className="w-4 h-4" />
              </button>
              {zoom !== 1 && (
                <button
                  type="button"
                  onClick={() => setZoom(1)}
                  aria-label="Fit slide to screen"
                  title="Fit to screen (0)"
                  className="p-1.5 rounded-full hover:bg-black/5 dark:hover:bg-white/10 text-[#7C63E6] dark:text-[#B9A7FF] cursor-pointer"
                >
                  <RotateCcw className="w-3.5 h-3.5" />
                </button>
              )}
            </div>
          </div>

          {/* Right: Quick Switcher, Shortcuts, Mobile Thumbnails Toggle, Presentation Mode */}
          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={() => setMobileDrawerOpen((o) => !o)}
              aria-label="Toggle slide thumbnails drawer"
              className="md:hidden clay-btn px-3.5 py-2 text-xs flex items-center gap-1.5"
            >
              <PanelLeft className="w-4 h-4" />
              <span>Slides</span>
            </button>

            <button
              type="button"
              onClick={onOpenQuickSwitcher}
              aria-label="Quick Switch Deck (Ctrl+K)"
              className="hidden sm:flex clay-btn px-3.5 py-2 text-xs items-center gap-1.5 whitespace-nowrap shrink-0"
            >
              <Command className="w-3.5 h-3.5 text-[#7C63E6] dark:text-[#B9A7FF]" />
              <span>Switch Deck</span>
            </button>

            <button
              type="button"
              onClick={() => setShowShortcutsModal((s) => !s)}
              aria-label="Keyboard shortcuts"
              title="Keyboard shortcuts"
              className="clay-btn p-2.5 text-[var(--clay-text-muted)] hover:text-[var(--clay-text)]"
            >
              <Keyboard className="w-4 h-4" />
            </button>

            <button
              type="button"
              onClick={toggleFullscreen}
              aria-label="Enter Presentation Mode (F)"
              className="clay-btn-lavender px-4 py-2 text-xs md:text-sm flex items-center gap-2 whitespace-nowrap shrink-0"
            >
              <Maximize2 className="w-4 h-4" />
              <span>Present</span>
            </button>
          </div>
        </div>
      )}

      {/* Main Split Workspace: Left Virtualized Sidebar + Right Clay Stage */}
      <div className="flex-1 flex flex-col md:flex-row gap-5 min-h-0 relative">
        {/* Left Sidebar: Virtualized Scrollable Slide Thumbnails (Desktop) */}
        {!isPresentationMode && (
          <aside
            aria-label="Slide Thumbnails"
            className="hidden md:flex flex-col w-64 lg:w-72 shrink-0 clay-card p-3.5 h-full min-h-0"
          >
            <div className="flex items-center justify-between px-2 pb-2.5 mb-2 border-b border-[var(--clay-border)] shrink-0">
              <span className="text-xs font-bold uppercase tracking-wider text-[var(--clay-text-muted)] flex items-center gap-1.5">
                <Layers className="w-3.5 h-3.5 text-[#7C63E6] dark:text-[#B9A7FF]" />
                Slides ({slides.length})
              </span>
              <span className="text-xs font-mono-num font-bold text-[var(--clay-text-subtle)]">
                {formatAspectRatioLabel(resolvedAspectRatio)}
              </span>
            </div>

            <div
              ref={sidebarScrollRef}
              onScroll={(e) => setScrollTop(e.currentTarget.scrollTop)}
              className="flex-1 overflow-y-auto pr-1 relative min-h-0"
            >
              <div
                style={{ height: `${virtualWindow.totalHeight}px` }}
                className="relative w-full"
              >
                <div
                  style={{
                    transform: `translateY(${virtualWindow.offsetY}px)`,
                  }}
                  className="flex flex-col gap-3"
                >
                  {visibleThumbnailSlides.map((s) => {
                    const isSelected = s.index === currentIndex;
                    return (
                      <button
                        key={s.index}
                        type="button"
                        onClick={() => goToSlide(s.index)}
                        aria-label={`Go to slide ${s.slideNumber}: ${s.title}`}
                        aria-current={isSelected ? 'step' : undefined}
                        style={{ height: `${THUMB_ITEM_HEIGHT - 12}px` }}
                        className={`w-full text-left p-2 rounded-[20px] transition-all flex flex-col justify-between cursor-pointer ${
                          isSelected
                            ? 'bg-[#EBE4FF] dark:bg-[#352B59] ring-3 ring-[#7C63E6] dark:ring-[#B9A7FF] shadow-md'
                            : 'hover:bg-black/5 dark:hover:bg-white/5'
                        }`}
                      >
                        <div className="w-full flex-1 min-h-0 rounded-[12px] bg-black/5 dark:bg-black/25 flex items-center justify-center overflow-hidden p-1">
                          <div
                            style={{ aspectRatio: `${resolvedAspectRatio}` }}
                            className="max-h-full max-w-full rounded-[8px] overflow-hidden shadow-xs border border-black/10 dark:border-white/10"
                          >
                            <SlideCanvas
                              slide={s}
                              aspectRatio={resolvedAspectRatio}
                              isThumbnail={true}
                              rawBuffer={rawBuffer}
                              deckId={deck.id}
                              isSample={deck.isSample}
                            />
                          </div>
                        </div>
                        <div className="flex items-center justify-between gap-2 px-1 pt-1 shrink-0">
                          <span className="text-[11px] font-bold font-mono-num text-[#7C63E6] dark:text-[#B9A7FF] shrink-0">
                            {String(s.slideNumber).padStart(2, '0')}
                          </span>
                          <span className="text-[11px] font-semibold text-[var(--clay-text)] truncate">
                            {s.title}
                          </span>
                        </div>
                      </button>
                    );
                  })}
                </div>
              </div>
            </div>
          </aside>
        )}

        {/* Right / Center: Main Slide Stage inside a Clay "Screen" Bezel */}
        <div
          onTouchStart={handleTouchStart}
          onTouchEnd={handleTouchEnd}
          className="flex-1 flex flex-col justify-between min-w-0 min-h-0"
        >
          {/* Clay Screen Bezel */}
          <div
            ref={stageBezelRef}
            className={`clay-bezel flex-1 min-h-0 flex items-center justify-center p-4 sm:p-5 relative ${
              zoom > 1 ? 'overflow-auto' : 'overflow-hidden'
            } ${isPresentationMode ? 'bg-[#1F1934] border-white/15' : ''}`}
          >
            {/* Distraction-free Presentation Mode Exit Pill */}
            {isPresentationMode && (
              <div className="absolute top-4 right-4 z-30 flex items-center gap-2">
                <button
                  type="button"
                  onClick={toggleFullscreen}
                  aria-label="Exit Presentation Mode (Esc)"
                  className="clay-btn px-4 py-2 text-xs flex items-center gap-1.5 shadow-lg"
                >
                  <Minimize2 className="w-3.5 h-3.5" />
                  <span>Exit Fullscreen (Esc)</span>
                </button>
              </div>
            )}

            {/* Exact Aspect-Ratio Fitted Slide Stage Container */}
            <div
              style={{
                width: `${fittedStageSize.width}px`,
                height: `${fittedStageSize.height}px`,
              }}
              className="mx-auto rounded-[20px] overflow-hidden shadow-2xl border border-white/70 dark:border-white/15 shrink-0"
            >
              <SlideCanvas
                slide={activeSlide}
                aspectRatio={resolvedAspectRatio}
                rawBuffer={rawBuffer}
                deckId={deck.id}
                isSample={deck.isSample}
              />
            </div>

            {/* Hidden Preload Layer for Adjacent Slides (Zero-Lag Navigation) */}
            <div
              aria-hidden="true"
              className="fixed -left-[9999px] top-0 w-40 h-24 pointer-events-none opacity-0 overflow-hidden"
            >
              {hotPreloadSlides.map((preloadSlide) => (
                <SlideCanvas
                  key={preloadSlide.index}
                  slide={preloadSlide}
                  aspectRatio={resolvedAspectRatio}
                  isThumbnail={true}
                  rawBuffer={rawBuffer}
                  deckId={deck.id}
                  isSample={deck.isSample}
                />
              ))}
            </div>
          </div>

          {/* Bottom Navigation Bar: Prev/Next Buttons + Slide Counter + Clay Progress Bar */}
          <div className="clay-card mt-4 px-5 py-3 flex flex-col sm:flex-row items-center justify-between gap-4 shrink-0">
            {/* Previous / Next Controls */}
            <div className="flex items-center gap-2.5">
              <button
                type="button"
                onClick={goPrev}
                disabled={currentIndex <= 0}
                aria-label="Previous slide"
                className="clay-btn px-4 py-2 text-xs md:text-sm flex items-center gap-1.5 whitespace-nowrap shrink-0"
              >
                <ChevronLeft className="w-4 h-4 stroke-[2.5]" />
                <span>Previous</span>
              </button>

              <div
                aria-live="polite"
                data-testid="slide-counter"
                className="clay-inset px-4 py-2 text-xs md:text-sm font-bold font-mono-num text-[var(--clay-text)] min-w-[84px] text-center"
              >
                {currentIndex + 1} / {totalSlides}
              </div>

              <button
                type="button"
                onClick={goNext}
                disabled={currentIndex >= slides.length - 1}
                aria-label="Next slide"
                className="clay-btn-lavender px-4 py-2 text-xs md:text-sm flex items-center gap-1.5 whitespace-nowrap shrink-0"
              >
                <span>Next</span>
                <ChevronRight className="w-4 h-4 stroke-[2.5]" />
              </button>
            </div>

            {/* Interactive Slide Progress Bar + Active Slide Title */}
            <div className="flex-1 w-full max-w-xl flex flex-col gap-1.5">
              <div className="flex items-center justify-between text-xs font-semibold text-[var(--clay-text-muted)]">
                <span className="truncate max-w-xs md:max-w-md font-bold text-[var(--clay-text)]">
                  {activeSlide.slideNumber}. {activeSlide.title}
                </span>
                <span className="font-mono-num shrink-0">
                  {progressPercentage}%
                </span>
              </div>

              {/* Clickable Progress Track */}
              <div
                role="progressbar"
                aria-valuenow={currentIndex + 1}
                aria-valuemin={1}
                aria-valuemax={totalSlides}
                aria-label="Presentation slide progress"
                onClick={(e) => {
                  const rect = e.currentTarget.getBoundingClientRect();
                  const ratio = (e.clientX - rect.left) / rect.width;
                  const targetIdx = Math.round(ratio * (slides.length - 1));
                  goToSlide(targetIdx);
                }}
                className="w-full h-3.5 clay-inset p-0.5 cursor-pointer overflow-hidden"
              >
                <div
                  style={{
                    width: `${progressPercentage}%`,
                    background: 'linear-gradient(90deg, #B9A7FF, #A8EBD0)',
                  }}
                  className="h-full rounded-full transition-all duration-200"
                />
              </div>
            </div>
          </div>
        </div>

        {/* Mobile Bottom Drawer for Slide Thumbnails */}
        <AnimatePresence>
          {mobileDrawerOpen && !isPresentationMode && (
            <motion.div
              initial={{ opacity: 0, y: 80 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: 80 }}
              className="md:hidden fixed inset-x-3 bottom-3 z-40 clay-card p-4 max-h-[55vh] flex flex-col"
            >
              <div className="flex items-center justify-between mb-3">
                <span className="text-sm font-bold font-display text-[var(--clay-text)]">
                  Jump to Slide ({slides.length})
                </span>
                <button
                  type="button"
                  onClick={() => setMobileDrawerOpen(false)}
                  className="clay-btn px-3 py-1 text-xs"
                >
                  Close
                </button>
              </div>
              <div className="grid grid-cols-2 gap-3 overflow-y-auto pr-1">
                {slides.map((s) => (
                  <button
                    key={s.index}
                    type="button"
                    onClick={() => {
                      goToSlide(s.index);
                      setMobileDrawerOpen(false);
                    }}
                    className={`p-2 rounded-[16px] text-left ${
                      s.index === currentIndex
                        ? 'bg-[#EBE4FF] dark:bg-[#352B59] ring-2 ring-[#7C63E6]'
                        : 'bg-black/5 dark:bg-white/5'
                    }`}
                  >
                    <div className="rounded-[10px] overflow-hidden mb-1">
                      <SlideCanvas
                        slide={s}
                        aspectRatio={resolvedAspectRatio}
                        isThumbnail={true}
                        rawBuffer={rawBuffer}
                        deckId={deck.id}
                        isSample={deck.isSample}
                      />
                    </div>
                    <div className="text-xs font-bold truncate text-[var(--clay-text)]">
                      {s.slideNumber}. {s.title}
                    </div>
                  </button>
                ))}
              </div>
            </motion.div>
          )}
        </AnimatePresence>

        {/* Keyboard Shortcuts Modal */}
        <AnimatePresence>
          {showShortcutsModal && (
            <motion.div
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setShowShortcutsModal(false)}
              className="fixed inset-0 z-50 bg-black/40 backdrop-blur-xs flex items-center justify-center p-4"
            >
              <motion.div
                initial={{ scale: 0.94, opacity: 0 }}
                animate={{ scale: 1, opacity: 1 }}
                exit={{ scale: 0.94, opacity: 0 }}
                onClick={(e) => e.stopPropagation()}
                role="dialog"
                aria-modal="true"
                aria-label="Keyboard Shortcuts"
                className="clay-card p-6 md:p-8 max-w-md w-full"
              >
                <div className="flex items-center justify-between mb-4">
                  <h3 className="text-xl font-bold font-display text-[var(--clay-text)]">
                    Keyboard Shortcuts
                  </h3>
                  <button
                    type="button"
                    onClick={() => setShowShortcutsModal(false)}
                    className="clay-btn px-3 py-1 text-xs"
                  >
                    Close
                  </button>
                </div>

                <div className="space-y-2.5 text-sm">
                  {[
                    ['← / → or Space', 'Previous / Next slide'],
                    ['Home / End', 'Jump to first / last slide'],
                    ['F', 'Toggle fullscreen presentation'],
                    ['Esc', 'Exit fullscreen or viewer'],
                    ['+ / - / 0', 'Zoom in, zoom out, reset fit'],
                    ['Ctrl / Cmd + K', 'Quick switcher command palette'],
                  ].map(([keyLabel, desc]) => (
                    <div
                      key={keyLabel}
                      className="flex items-center justify-between py-2 border-b border-[var(--clay-border)]"
                    >
                      <span className="text-[var(--clay-text-muted)] font-semibold">
                        {desc}
                      </span>
                      <kbd className="clay-inset px-3 py-1 text-xs font-mono-num font-bold text-[var(--clay-text)]">
                        {keyLabel}
                      </kbd>
                    </div>
                  ))}
                </div>
              </motion.div>
            </motion.div>
          )}
        </AnimatePresence>
      </div>
    </div>
  );
};
