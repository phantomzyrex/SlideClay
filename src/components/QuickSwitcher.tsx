import React, { useEffect, useMemo, useState } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import {
  Search,
  CornerDownLeft,
  Layers,
  X,
} from 'lucide-react';
import { PresentationMetadata } from '../types/presentation';
import { formatFileSize, formatUploadDate } from '../utils/deckUtils';

interface QuickSwitcherProps {
  isOpen: boolean;
  decks: PresentationMetadata[];
  activeDeckId?: string;
  onClose: () => void;
  onSelectDeck: (deck: PresentationMetadata) => void;
}

export const QuickSwitcher: React.FC<QuickSwitcherProps> = ({
  isOpen,
  decks,
  activeDeckId,
  onClose,
  onSelectDeck,
}) => {
  const [query, setQuery] = useState('');
  const [selectedIndex, setSelectedIndex] = useState(0);

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return decks;
    return decks.filter(
      (d) =>
        d.name.toLowerCase().includes(q) ||
        d.originalFileName.toLowerCase().includes(q) ||
        d.firstSlide.title.toLowerCase().includes(q)
    );
  }, [decks, query]);

  useEffect(() => {
    if (isOpen) {
      setQuery('');
      setSelectedIndex(0);
    }
  }, [isOpen]);

  useEffect(() => {
    setSelectedIndex(0);
  }, [query]);

  useEffect(() => {
    if (!isOpen) return;
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.preventDefault();
        onClose();
      } else if (e.key === 'ArrowDown') {
        e.preventDefault();
        setSelectedIndex((prev) =>
          filtered.length > 0 ? (prev + 1) % filtered.length : 0
        );
      } else if (e.key === 'ArrowUp') {
        e.preventDefault();
        setSelectedIndex((prev) =>
          filtered.length > 0
            ? (prev - 1 + filtered.length) % filtered.length
            : 0
        );
      } else if (e.key === 'Enter' && filtered[selectedIndex]) {
        e.preventDefault();
        onSelectDeck(filtered[selectedIndex]);
        onClose();
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [isOpen, filtered, selectedIndex, onSelectDeck, onClose]);

  return (
    <AnimatePresence>
      {isOpen && (
        <motion.div
          initial={{ opacity: 0 }}
          animate={{ opacity: 1 }}
          exit={{ opacity: 0 }}
          onClick={onClose}
          className="fixed inset-0 z-50 bg-black/45 backdrop-blur-xs flex items-start justify-center pt-20 px-4"
        >
          <motion.div
            initial={{ scale: 0.94, opacity: 0, y: -12 }}
            animate={{ scale: 1, opacity: 1, y: 0 }}
            exit={{ scale: 0.94, opacity: 0, y: -12 }}
            transition={{ type: 'spring', stiffness: 360, damping: 26 }}
            onClick={(e) => e.stopPropagation()}
            role="dialog"
            aria-modal="true"
            aria-label="Quick Switcher Command Palette"
            className="clay-card w-full max-w-xl p-5 overflow-hidden"
          >
            {/* Search Input Well */}
            <div className="clay-inset flex items-center gap-3 px-4 py-3 mb-4">
              <Search className="w-5 h-5 text-[#7C63E6] dark:text-[#B9A7FF] shrink-0" />
              <input
                type="text"
                value={query}
                onChange={(e) => setQuery(e.target.value)}
                placeholder="Jump to any uploaded presentation..."
                aria-label="Search presentations in Quick Switcher"
                autoFocus
                className="w-full bg-transparent text-sm md:text-base font-bold text-[var(--clay-text)] placeholder:text-[var(--clay-text-subtle)] focus:outline-none"
              />
              <button
                type="button"
                onClick={onClose}
                aria-label="Close Quick Switcher"
                className="p-1 rounded-full hover:bg-black/5 dark:hover:bg-white/10 text-[var(--clay-text-muted)] cursor-pointer"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            {/* Deck Results List */}
            <div className="max-h-80 overflow-y-auto space-y-2 pr-1">
              {filtered.length === 0 ? (
                <div className="py-8 text-center text-sm font-semibold text-[var(--clay-text-muted)]">
                  No presentations match “{query}”
                </div>
              ) : (
                filtered.map((deck, idx) => {
                  const isHighlighted = idx === selectedIndex;
                  const isCurrent = deck.id === activeDeckId;
                  return (
                    <button
                      key={deck.id}
                      type="button"
                      onMouseEnter={() => setSelectedIndex(idx)}
                      onClick={() => {
                        onSelectDeck(deck);
                        onClose();
                      }}
                      className={`w-full text-left px-4 py-3 rounded-[20px] flex items-center justify-between gap-3 transition-all cursor-pointer ${
                        isHighlighted
                          ? 'bg-[#EBE4FF] dark:bg-[#352B59] shadow-xs'
                          : 'hover:bg-black/5 dark:hover:bg-white/5'
                      }`}
                    >
                      <div className="flex items-center gap-3 min-w-0">
                        <span
                          className="clay-icon-circle w-10 h-10 shrink-0"
                          style={{ backgroundColor: deck.accentColor || '#B9A7FF' }}
                        >
                          <Layers className="w-5 h-5 text-[#231942]" />
                        </span>
                        <div className="min-w-0">
                          <div className="text-sm md:text-base font-bold text-[var(--clay-text)] truncate">
                            {deck.name}
                            {isCurrent && (
                              <span className="ml-2 text-xs font-semibold text-[#7C63E6] dark:text-[#B9A7FF]">
                                · Currently Open
                              </span>
                            )}
                          </div>
                          <div className="text-xs font-mono-num text-[var(--clay-text-muted)]">
                            {deck.slideCount} slides · {formatFileSize(deck.fileSize)} ·{' '}
                            {formatUploadDate(deck.uploadedAt)}
                          </div>
                        </div>
                      </div>

                      <span className="text-xs font-bold text-[#7C63E6] dark:text-[#B9A7FF] flex items-center gap-1 shrink-0">
                        <span>Open</span>
                        <CornerDownLeft className="w-3.5 h-3.5" />
                      </span>
                    </button>
                  );
                })
              )}
            </div>

            {/* Footer Hints */}
            <div className="mt-4 pt-3 border-t border-[var(--clay-border)] flex items-center justify-between text-xs font-semibold text-[var(--clay-text-muted)]">
              <span>Use ↑↓ arrows to navigate, Enter to open</span>
              <span className="font-mono-num">Esc to close</span>
            </div>
          </motion.div>
        </motion.div>
      )}
    </AnimatePresence>
  );
};
