import React, { useRef, useState } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import {
  UploadCloud,
  Sparkles,
  CheckCircle2,
  AlertCircle,
  Loader2,
  FileSpreadsheet,
} from 'lucide-react';
import { UploadProgressItem } from '../types/presentation';
import { formatFileSize } from '../utils/deckUtils';

interface UploadZoneProps {
  onFilesSelected: (files: File[]) => void;
  uploadQueue: UploadProgressItem[];
  onClearFinishedUploads: () => void;
  onGenerateSampleDeck: () => void;
  isGeneratingSample?: boolean;
}

export const UploadZone: React.FC<UploadZoneProps> = ({
  onFilesSelected,
  uploadQueue,
  onClearFinishedUploads,
  onGenerateSampleDeck,
  isGeneratingSample = false,
}) => {
  const [isDraggingOver, setIsDraggingOver] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (!isDraggingOver) setIsDraggingOver(true);
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDraggingOver(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDraggingOver(false);

    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      const filesArray = Array.from(e.dataTransfer.files);
      onFilesSelected(filesArray);
    }
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      const filesArray = Array.from(e.target.files);
      onFilesSelected(filesArray);
      e.target.value = '';
    }
  };

  return (
    <section aria-label="Upload Presentations" className="w-full mb-8">
      <motion.div
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onDrop={handleDrop}
        animate={
          isDraggingOver
            ? {
                scale: 1.015,
                rotate: [-0.6, 0.6, -0.4, 0],
              }
            : {
                scale: 1,
                rotate: 0,
              }
        }
        transition={{ type: 'spring', stiffness: 320, damping: 20 }}
        className={`clay-card relative overflow-hidden p-6 md:p-8 transition-all ${
          isDraggingOver
            ? 'ring-4 ring-[#B9A7FF] dark:ring-[#9E86FF] bg-[#EBE4FF]/70 dark:bg-[#332A56]'
            : ''
        }`}
      >
        {/* Puffy dashed inner perimeter */}
        <div
          className={`rounded-[24px] border-3 border-dashed p-6 md:p-8 transition-colors flex flex-col lg:flex-row items-center justify-between gap-6 ${
            isDraggingOver
              ? 'border-[#7C63E6] bg-[#F4F0FF]/80 dark:bg-[#2A2248]/80'
              : 'border-[#B9A7FF]/65 dark:border-[#7C63E6]/45 bg-[#F4F1FF]/55 dark:bg-[#1F1A33]/60'
          }`}
        >
          {/* Left: Custom Clay Illustration + Copy */}
          <div className="flex flex-col sm:flex-row items-center sm:items-start text-center sm:text-left gap-5">
            {/* Soft 3D Clay Illustration */}
            <motion.div
              animate={
                isDraggingOver
                  ? { y: [-4, 4, -4], scale: 1.08 }
                  : { y: 0, scale: 1 }
              }
              transition={{ repeat: isDraggingOver ? Infinity : 0, duration: 1.2 }}
              className="w-20 h-20 rounded-[26px] shrink-0 flex items-center justify-center relative"
              style={{
                background: 'linear-gradient(145deg, #D5C8FF, #B9A7FF)',
                boxShadow:
                  '8px 8px 18px rgba(120,100,180,0.28), -5px -5px 14px rgba(255,255,255,0.85), inset 5px 5px 10px rgba(255,255,255,0.7), inset -5px -5px 10px rgba(105,80,205,0.32)',
              }}
              aria-hidden="true"
            >
              <UploadCloud className="w-10 h-10 text-[#231942] stroke-[2.4]" />
              <span
                className="w-6 h-6 rounded-full absolute -bottom-1 -right-1 flex items-center justify-center"
                style={{
                  backgroundColor: '#FFC8A8',
                  boxShadow:
                    'inset 2px 2px 4px rgba(255,255,255,0.8), inset -2px -2px 4px rgba(190,100,50,0.3)',
                }}
              >
                <FileSpreadsheet className="w-3.5 h-3.5 text-[#4A210B]" />
              </span>
            </motion.div>

            <div>
              <h2 className="text-xl md:text-2xl font-bold font-display text-[var(--clay-text)] mb-1.5">
                {isDraggingOver
                  ? 'Release to sculpted-open your .pptx decks!'
                  : 'Drop your PowerPoint (.pptx) presentations here'}
              </h2>
              <p className="text-sm md:text-base text-[var(--clay-text-muted)] max-w-xl">
                Upload multiple <span className="font-bold">.pptx</span> files at once.
                Slide 1 renders immediately while our Web Worker indexes the remaining
                slides in the background and caches everything in IndexedDB.
              </p>
            </div>
          </div>

          {/* Right: Chunky Action Buttons */}
          <div className="flex flex-wrap items-center justify-center gap-3 shrink-0">
            <input
              ref={fileInputRef}
              type="file"
              accept=".pptx,.ppt"
              multiple
              onChange={handleFileChange}
              className="hidden"
              aria-label="Choose PowerPoint files to upload"
            />
            <button
              type="button"
              onClick={() => fileInputRef.current?.click()}
              className="clay-btn-lavender px-6 py-3.5 text-sm md:text-base flex items-center gap-2.5 whitespace-nowrap shrink-0"
            >
              <UploadCloud className="w-5 h-5 stroke-[2.5]" />
              <span>Choose .pptx Files</span>
            </button>

            <button
              type="button"
              onClick={onGenerateSampleDeck}
              disabled={isGeneratingSample}
              className="clay-btn px-5 py-3.5 text-sm md:text-base flex items-center gap-2 whitespace-nowrap shrink-0"
            >
              {isGeneratingSample ? (
                <Loader2 className="w-4 h-4 animate-spin text-[#7C63E6]" />
              ) : (
                <Sparkles className="w-4 h-4 text-[#7C63E6] dark:text-[#B9A7FF]" />
              )}
              <span>Restore Sample Decks</span>
            </button>
          </div>
        </div>

        {/* Per-File Upload & Parsing Progress Queue */}
        <AnimatePresence>
          {uploadQueue.length > 0 && (
            <motion.div
              initial={{ opacity: 0, height: 0 }}
              animate={{ opacity: 1, height: 'auto' }}
              exit={{ opacity: 0, height: 0 }}
              className="mt-5 pt-4 border-t border-[var(--clay-border)]"
            >
              <div className="flex items-center justify-between mb-3">
                <span className="text-xs font-bold uppercase tracking-wider text-[var(--clay-text-muted)]">
                  Upload & Worker Parsing Queue ({uploadQueue.length})
                </span>
                <button
                  type="button"
                  onClick={onClearFinishedUploads}
                  className="text-xs font-bold text-[#7C63E6] dark:text-[#B9A7FF] hover:underline cursor-pointer"
                >
                  Dismiss Completed
                </button>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                {uploadQueue.map((item) => (
                  <div
                    key={item.id}
                    className="clay-inset-panel px-4 py-3 flex flex-col gap-2"
                  >
                    <div className="flex items-center justify-between gap-2">
                      <div className="flex items-center gap-2 min-w-0">
                        {item.status === 'done' ? (
                          <CheckCircle2 className="w-4 h-4 text-[#2B9E74] shrink-0" />
                        ) : item.status === 'error' ? (
                          <AlertCircle className="w-4 h-4 text-[#D94872] shrink-0" />
                        ) : (
                          <Loader2 className="w-4 h-4 text-[#7C63E6] animate-spin shrink-0" />
                        )}
                        <span className="text-sm font-bold truncate text-[var(--clay-text)]">
                          {item.fileName}
                        </span>
                      </div>
                      <span className="text-xs font-mono-num text-[var(--clay-text-muted)] shrink-0">
                        {formatFileSize(item.fileSize)} · {item.progress}%
                      </span>
                    </div>

                    {/* Progress Pill Bar */}
                    <div className="w-full h-2.5 rounded-full bg-white/60 dark:bg-black/30 overflow-hidden p-0.5">
                      <div
                        style={{
                          width: `${Math.max(5, item.progress)}%`,
                          backgroundColor:
                            item.status === 'error'
                              ? '#FFB4C8'
                              : item.status === 'done'
                                ? '#A8EBD0'
                                : '#B9A7FF',
                        }}
                        className="h-full rounded-full transition-all duration-200"
                      />
                    </div>

                    {item.errorMessage && (
                      <p className="text-xs font-semibold text-[#D94872] dark:text-[#FFB4C8]">
                        {item.errorMessage}
                      </p>
                    )}
                  </div>
                ))}
              </div>
            </motion.div>
          )}
        </AnimatePresence>
      </motion.div>
    </section>
  );
};
