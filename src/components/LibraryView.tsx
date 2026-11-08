import React, { useState } from 'react';
import { motion } from 'motion/react';
import {
  Search,
  ArrowUpDown,
  Play,
  Pencil,
  Trash2,
  Download,
  Check,
  X,
  FolderOpen,
  Layers,
  HardDrive,
} from 'lucide-react';
import {
  PresentationMetadata,
  SortField,
  SortOrder,
} from '../types/presentation';
import {
  formatAspectRatioLabel,
  formatFileSize,
  formatUploadDate,
} from '../utils/deckUtils';
import { SlideCanvas } from './SlideCanvas';

interface LibraryViewProps {
  decks: PresentationMetadata[];
  totalDeckCount: number;
  searchQuery: string;
  onSearchChange: (q: string) => void;
  sortBy: SortField;
  sortOrder: SortOrder;
  onSortChange: (field: SortField) => void;
  onToggleSortOrder: () => void;
  onOpenDeck: (deck: PresentationMetadata) => void;
  onRenameDeck: (id: string, newName: string) => void;
  onDeleteDeck: (deck: PresentationMetadata) => void;
  onDownloadDeck: (deck: PresentationMetadata) => void;
  onRestoreSamples: () => void;
}

export const LibraryView: React.FC<LibraryViewProps> = ({
  decks,
  totalDeckCount,
  searchQuery,
  onSearchChange,
  sortBy,
  sortOrder,
  onSortChange,
  onToggleSortOrder,
  onOpenDeck,
  onRenameDeck,
  onDeleteDeck,
  onDownloadDeck,
  onRestoreSamples,
}) => {
  const [editingId, setEditingId] = useState<string | null>(null);
  const [editingValue, setEditingValue] = useState('');

  const startRename = (deck: PresentationMetadata, e: React.MouseEvent) => {
    e.stopPropagation();
    setEditingId(deck.id);
    setEditingValue(deck.name);
  };

  const commitRename = (id: string, e?: React.FormEvent | React.MouseEvent) => {
    e?.stopPropagation();
    e?.preventDefault();
    if (editingValue.trim()) {
      onRenameDeck(id, editingValue.trim());
    }
    setEditingId(null);
  };

  const cancelRename = (e: React.MouseEvent) => {
    e.stopPropagation();
    setEditingId(null);
  };

  const totalSlidesCount = decks.reduce((acc, d) => acc + d.slideCount, 0);
  const totalBytesCount = decks.reduce((acc, d) => acc + d.fileSize, 0);

  return (
    <section aria-label="Presentation Library" className="w-full">
      {/* Toolbar: Search + Sort Controls + Library Summary */}
      <div className="flex flex-col lg:flex-row items-stretch lg:items-center justify-between gap-4 mb-7">
        {/* Search Input Well */}
        <div className="clay-inset flex items-center gap-3 px-5 py-3 flex-1 max-w-xl">
          <Search className="w-5 h-5 text-[var(--clay-text-muted)] shrink-0" />
          <input
            type="search"
            value={searchQuery}
            onChange={(e) => onSearchChange(e.target.value)}
            placeholder="Search presentations by file name or slide title..."
            aria-label="Search presentations"
            className="w-full bg-transparent text-sm md:text-base font-semibold text-[var(--clay-text)] placeholder:text-[var(--clay-text-subtle)] focus:outline-none"
          />
          {searchQuery && (
            <button
              type="button"
              onClick={() => onSearchChange('')}
              aria-label="Clear search query"
              className="p-1 rounded-full hover:bg-black/5 dark:hover:bg-white/10 text-[var(--clay-text-muted)] cursor-pointer"
            >
              <X className="w-4 h-4" />
            </button>
          )}
        </div>

        {/* Sort Segmented Controls + Stats */}
        <div className="flex flex-wrap items-center justify-between lg:justify-end gap-3">
          <div className="hidden xl:flex items-center gap-2 text-xs font-bold text-[var(--clay-text-muted)] mr-2 font-mono-num">
            <span className="inline-flex items-center gap-1">
              <Layers className="w-3.5 h-3.5 text-[#7C63E6] dark:text-[#B9A7FF]" />
              {decks.length} {decks.length === 1 ? 'deck' : 'decks'}
            </span>
            <span aria-hidden="true">·</span>
            <span>{totalSlidesCount} slides</span>
            <span aria-hidden="true">·</span>
            <span className="inline-flex items-center gap-1">
              <HardDrive className="w-3.5 h-3.5 text-[#2B9E74] dark:text-[#A8EBD0]" />
              {formatFileSize(totalBytesCount)} cached
            </span>
          </div>

          <div
            role="group"
            aria-label="Sort presentations by"
            className="clay-inset p-1.5 flex items-center gap-1"
          >
            {(['date', 'name', 'size'] as SortField[]).map((field) => {
              const active = sortBy === field;
              const label =
                field === 'date' ? 'Date' : field === 'name' ? 'Name' : 'Size';
              return (
                <button
                  key={field}
                  type="button"
                  onClick={() => onSortChange(field)}
                  className={`px-4 py-1.5 rounded-full text-xs md:text-sm font-bold transition-all whitespace-nowrap shrink-0 cursor-pointer ${
                    active
                      ? 'clay-btn-lavender text-[#1F1638]'
                      : 'text-[var(--clay-text-muted)] hover:text-[var(--clay-text)]'
                  }`}
                >
                  Sort: {label}
                </button>
              );
            })}
          </div>

          <button
            type="button"
            onClick={onToggleSortOrder}
            aria-label={`Sort order: ${
              sortOrder === 'asc' ? 'Ascending' : 'Descending'
            }`}
            className="clay-btn px-4 py-2.5 text-xs md:text-sm flex items-center gap-1.5 whitespace-nowrap shrink-0"
          >
            <ArrowUpDown className="w-4 h-4 text-[#7C63E6] dark:text-[#B9A7FF]" />
            <span>{sortOrder === 'asc' ? 'Asc' : 'Desc'}</span>
          </button>
        </div>
      </div>

      {/* Empty State */}
      {decks.length === 0 ? (
        <motion.div
          initial={{ opacity: 0, y: 12 }}
          animate={{ opacity: 1, y: 0 }}
          className="clay-card p-10 md:p-14 text-center max-w-2xl mx-auto my-6 flex flex-col items-center"
        >
          {/* Cute Clay Empty Illustration */}
          <div
            className="w-28 h-28 rounded-[34px] mb-6 flex items-center justify-center relative"
            style={{
              background: 'linear-gradient(145deg, #FFE9A8, #FFC8A8)',
              boxShadow:
                '12px 12px 24px rgba(190,130,90,0.25), -8px -8px 20px rgba(255,255,255,0.9), inset 6px 6px 12px rgba(255,255,255,0.75), inset -6px -6px 12px rgba(190,115,70,0.28)',
            }}
            aria-hidden="true"
          >
            <FolderOpen className="w-14 h-14 text-[#4A210B] stroke-[2.2]" />
          </div>

          <h3 className="text-2xl font-bold font-display text-[var(--clay-text)] mb-2">
            {totalDeckCount === 0
              ? 'Drop your presentations here'
              : 'No presentations match your search'}
          </h3>
          <p className="text-sm md:text-base text-[var(--clay-text-muted)] max-w-md mb-6">
            {totalDeckCount === 0
              ? 'Your IndexedDB presentation library is currently empty. Drag and drop any .pptx file above or load our sample showcase decks.'
              : `We couldn't find any deck matching "${searchQuery}". Try clearing your search filter.`}
          </p>

          {totalDeckCount === 0 ? (
            <button
              type="button"
              onClick={onRestoreSamples}
              className="clay-btn-lavender px-6 py-3 text-sm md:text-base"
            >
              Load Sample .pptx Decks
            </button>
          ) : (
            <button
              type="button"
              onClick={() => onSearchChange('')}
              className="clay-btn px-6 py-3 text-sm md:text-base"
            >
              Reset Search Filter
            </button>
          )}
        </motion.div>
      ) : (
        /* Responsive Presentation Card Grid */
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-7">
          {decks.map((deck, idx) => {
            const isEditing = editingId === deck.id;

            return (
              <motion.article
                key={deck.id}
                initial={{ opacity: 0, y: 14 }}
                animate={{ opacity: 1, y: 0 }}
                transition={{
                  duration: 0.22,
                  delay: Math.min(idx * 0.04, 0.2),
                }}
                onClick={() => {
                  if (!isEditing) onOpenDeck(deck);
                }}
                className="clay-card clay-card-interactive p-5 flex flex-col justify-between cursor-pointer group"
              >
                <div>
                  {/* First-Slide Thumbnail Preview Frame (preserves 16:9, 4:3, or custom page aspect ratio) */}
                  <div className="relative aspect-video rounded-[20px] overflow-hidden mb-4 border border-white/60 dark:border-white/10 shadow-inner bg-[#EAE4FA] dark:bg-[#1B162E] flex items-center justify-center p-1.5">
                    <div
                      style={{
                        aspectRatio: `${deck.aspectRatio || 16 / 9}`,
                      }}
                      className="max-h-full max-w-full w-full rounded-[14px] overflow-hidden shadow-sm"
                    >
                      <SlideCanvas
                        slide={deck.firstSlide}
                        aspectRatio={deck.aspectRatio || 16 / 9}
                        isThumbnail={true}
                        deckId={deck.id}
                        isSample={deck.isSample}
                      />
                    </div>

                    {/* Hover Play Overlay */}
                    <div className="absolute inset-0 bg-[#1F1638]/0 group-hover:bg-[#1F1638]/25 transition-colors flex items-center justify-center opacity-0 group-hover:opacity-100">
                      <span className="clay-btn-lavender px-5 py-2.5 text-sm flex items-center gap-2 shadow-lg">
                        <Play className="w-4 h-4 fill-current" />
                        <span>Open Instant Viewer</span>
                      </span>
                    </div>
                  </div>

                  {/* Card Title or Inline Rename Form */}
                  {isEditing ? (
                    <form
                      onSubmit={(e) => commitRename(deck.id, e)}
                      onClick={(e) => e.stopPropagation()}
                      className="flex items-center gap-2 mb-2"
                    >
                      <input
                        type="text"
                        value={editingValue}
                        onChange={(e) => setEditingValue(e.target.value)}
                        aria-label={`New name for ${deck.name}`}
                        autoFocus
                        className="clay-inset px-3.5 py-1.5 text-sm font-bold text-[var(--clay-text)] flex-1 min-w-0"
                      />
                      <button
                        type="submit"
                        aria-label="Save presentation name"
                        className="clay-btn-mint p-2 rounded-full"
                      >
                        <Check className="w-4 h-4" />
                      </button>
                      <button
                        type="button"
                        onClick={cancelRename}
                        aria-label="Cancel renaming"
                        className="clay-btn p-2 rounded-full"
                      >
                        <X className="w-4 h-4" />
                      </button>
                    </form>
                  ) : (
                    <h3
                      title={deck.name}
                      className="text-lg md:text-xl font-bold font-display text-[var(--clay-text)] truncate mb-1.5 group-hover:text-[#7C63E6] dark:group-hover:text-[#B9A7FF] transition-colors"
                    >
                      {deck.name}
                    </h3>
                  )}

                  {/* Unboxed Metadata Line (Zero-Pill Discipline) */}
                  <div className="flex flex-wrap items-center gap-1.5 text-xs font-semibold text-[var(--clay-text-muted)] font-mono-num mb-4">
                    <span>{deck.slideCount} slides</span>
                    <span aria-hidden="true">·</span>
                    <span>{formatAspectRatioLabel(deck.aspectRatio || 16 / 9)}</span>
                    <span aria-hidden="true">·</span>
                    <span>{formatFileSize(deck.fileSize)}</span>
                    <span aria-hidden="true">·</span>
                    <span>{formatUploadDate(deck.uploadedAt)}</span>
                  </div>
                </div>

                {/* Card Footer Actions */}
                <div
                  onClick={(e) => e.stopPropagation()}
                  className="pt-3 border-t border-[var(--clay-border)] flex items-center justify-between gap-2"
                >
                  <button
                    type="button"
                    onClick={() => onOpenDeck(deck)}
                    aria-label={`Open ${deck.name}`}
                    className="clay-btn-lavender px-4 py-2 text-xs md:text-sm flex items-center gap-1.5 whitespace-nowrap shrink-0"
                  >
                    <Play className="w-3.5 h-3.5 fill-current" />
                    <span>Open Deck</span>
                  </button>

                  <div className="flex items-center gap-1.5">
                    <button
                      type="button"
                      onClick={(e) => startRename(deck, e)}
                      aria-label={`Rename ${deck.name}`}
                      title="Rename presentation"
                      className="clay-btn p-2.5 text-[var(--clay-text-muted)] hover:text-[var(--clay-text)]"
                    >
                      <Pencil className="w-4 h-4" />
                    </button>

                    <button
                      type="button"
                      onClick={() => onDownloadDeck(deck)}
                      aria-label={`Download ${deck.name} as .pptx`}
                      title="Download .pptx file"
                      className="clay-btn p-2.5 text-[var(--clay-text-muted)] hover:text-[var(--clay-text)]"
                    >
                      <Download className="w-4 h-4" />
                    </button>

                    <button
                      type="button"
                      onClick={() => onDeleteDeck(deck)}
                      aria-label={`Delete ${deck.name}`}
                      title="Delete presentation"
                      className="clay-btn p-2.5 text-[#D94872] dark:text-[#FFB4C8]"
                    >
                      <Trash2 className="w-4 h-4" />
                    </button>
                  </div>
                </div>
              </motion.article>
            );
          })}
        </div>
      )}
    </section>
  );
};
