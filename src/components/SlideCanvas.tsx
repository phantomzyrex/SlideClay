import React, { useEffect, useRef, useState } from 'react';
import { ParsedSlide, SlideElement } from '../types/presentation';
import { getHotRawBuffer, getPresentationFull } from '../services/db';
import { renderSlideWithPptxPreview } from '../services/pptxPreviewService';

interface SlideCanvasProps {
  slide: ParsedSlide;
  aspectRatio?: number;
  className?: string;
  isThumbnail?: boolean;
  rawBuffer?: ArrayBuffer;
  deckId?: string;
  isSample?: boolean;
}

/**
 * Renders a single presentation slide at its exact original aspect ratio and
 * coordinate space.
 *
 * - For user-uploaded `.pptx` files (`!isSample`), uses `pptx-preview`'s native
 *   OpenXML engine (`SlideMaster` + `SlideLayout` + `Slide` layers, custom vector
 *   geometries, SmartArt, tables, ECharts, and embedded media) rendered at 1:1
 *   native point dimensions (`pptx.width` x `pptx.height`) and scaled uniformly
 *   via CSS `transform: scale(...)` so text positions, sizes, and layouts never shift.
 * - For built-in sample decks (`isSample`) or before the binary buffer finishes
 *   loading from IndexedDB, renders the responsive SVG `viewBox` vector layer.
 */
export const SlideCanvas: React.FC<SlideCanvasProps> = React.memo(
  ({
    slide,
    aspectRatio = 16 / 9,
    className = '',
    isThumbnail = false,
    rawBuffer,
    deckId,
    isSample = false,
  }) => {
    const outerContainerRef = useRef<HTMLDivElement>(null);
    const nativeHostRef = useRef<HTMLDivElement>(null);

    const [resolvedBuffer, setResolvedBuffer] = useState<ArrayBuffer | undefined>(
      () => rawBuffer || (deckId ? getHotRawBuffer(deckId) : undefined)
    );
    const [nativeDims, setNativeDims] = useState<{
      width: number;
      height: number;
    } | null>(null);
    const [nativeReady, setNativeReady] = useState(false);
    const [containerWidth, setContainerWidth] = useState<number>(0);

    const safeRatio = aspectRatio && aspectRatio > 0 ? aspectRatio : 16 / 9;
    const baseWidth = slide.canvasWidth || 960;
    const baseHeight =
      slide.canvasHeight || Math.round(baseWidth / safeRatio);
    const activeWidth = nativeDims?.width || baseWidth;
    const activeHeight = nativeDims?.height || baseHeight;
    const computedRatio = activeWidth / activeHeight;

    // Resolve binary ArrayBuffer from props, hot memory cache, or IndexedDB
    useEffect(() => {
      if (isSample) {
        setResolvedBuffer(undefined);
        setNativeReady(false);
        return;
      }
      if (rawBuffer) {
        setResolvedBuffer(rawBuffer);
        return;
      }
      if (!deckId) return;

      const hot = getHotRawBuffer(deckId);
      if (hot) {
        setResolvedBuffer(hot);
        return;
      }

      let cancelled = false;
      getPresentationFull(deckId)
        .then((full) => {
          if (!cancelled && full?.rawBuffer) {
            setResolvedBuffer(full.rawBuffer);
          }
        })
        .catch(() => {});

      return () => {
        cancelled = true;
      };
    }, [rawBuffer, deckId, isSample]);

    // Track outer container width for uniform 1:1 CSS transform scaling
    useEffect(() => {
      const el = outerContainerRef.current;
      if (!el) return;
      const updateWidth = () => {
        const w = el.clientWidth;
        if (w > 0) {
          setContainerWidth((prev) => (Math.abs(prev - w) > 0.5 ? w : prev));
        }
      };
      updateWidth();
      const ro = new ResizeObserver(updateWidth);
      ro.observe(el);
      return () => ro.disconnect();
    }, []);

    // Render exact native slide via `pptx-preview` when buffer is available
    useEffect(() => {
      if (isSample || !resolvedBuffer || !nativeHostRef.current) {
        setNativeReady(false);
        return;
      }

      let cancelled = false;
      const targetEl = nativeHostRef.current;

      renderSlideWithPptxPreview(
        deckId,
        resolvedBuffer,
        slide.index,
        targetEl
      )
        .then((dims) => {
          if (cancelled) return;
          setNativeDims(dims);
          setNativeReady(true);
        })
        .catch(() => {
          if (!cancelled) {
            setNativeReady(false);
          }
        });

      return () => {
        cancelled = true;
      };
    }, [isSample, resolvedBuffer, deckId, slide.index]);

    const renderElement = (el: SlideElement) => {
      const commonStyle: React.CSSProperties = {
        position: 'absolute',
        left: `${el.x}px`,
        top: `${el.y}px`,
        width: `${el.w}px`,
        height: `${el.h}px`,
        boxSizing: 'border-box',
        transform: el.rotation ? `rotate(${el.rotation}deg)` : undefined,
        transformOrigin: 'center center',
      };

      if (el.type === 'image' && el.imageUrl) {
        return (
          <div
            key={el.id}
            style={{
              ...commonStyle,
              borderRadius: `${el.borderRadius ?? 0}px`,
              overflow: 'hidden',
            }}
          >
            <img
              src={el.imageUrl}
              alt={slide.title || 'Slide graphic'}
              referrerPolicy="no-referrer"
              className="w-full h-full object-contain"
            />
          </div>
        );
      }

      if (el.type === 'table' && el.table) {
        const { headers, rows, colWidthsPx, accentColor = '#B9A7FF' } = el.table;
        return (
          <div
            key={el.id}
            style={{
              ...commonStyle,
              borderRadius: '14px',
              backgroundColor: '#FFFFFF',
              boxShadow:
                '6px 6px 16px rgba(120, 100, 180, 0.12), inset 2px 2px 4px rgba(255,255,255,0.8)',
              overflow: 'hidden',
              border: '1px solid rgba(185, 167, 255, 0.35)',
            }}
          >
            <table className="w-full h-full border-collapse text-left table-fixed">
              {colWidthsPx && colWidthsPx.length === headers.length && (
                <colgroup>
                  {colWidthsPx.map((cw, i) => (
                    <col key={i} style={{ width: `${cw}px` }} />
                  ))}
                </colgroup>
              )}
              <thead>
                <tr style={{ backgroundColor: accentColor }}>
                  {headers.map((h, i) => (
                    <th
                      key={i}
                      className="px-4 py-2.5 font-bold text-[#1F1735] text-[14px] border-b border-black/10 break-words"
                    >
                      {h}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {rows.map((row, rIdx) => (
                  <tr
                    key={rIdx}
                    style={{
                      backgroundColor: rIdx % 2 === 0 ? '#FFFFFF' : '#F8F6FF',
                    }}
                  >
                    {row.map((cell, cIdx) => (
                      <td
                        key={cIdx}
                        className={`px-4 py-2 text-[13px] text-[#332952] border-b border-[#EDE8FF] break-words ${
                          cIdx === 0 ? 'font-bold' : 'font-mono-num'
                        }`}
                      >
                        {cell}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        );
      }

      if (el.type === 'chart' && el.chartBars) {
        const maxVal = Math.max(...el.chartBars.map((b) => b.value), 1);
        return (
          <div
            key={el.id}
            style={{
              ...commonStyle,
              borderRadius: '26px',
              backgroundColor: '#FFFFFF',
              padding: '28px 34px',
              boxShadow:
                '10px 10px 24px rgba(120, 100, 180, 0.14), inset 4px 4px 8px rgba(255,255,255,0.9), inset -4px -4px 8px rgba(150,130,200,0.16)',
              border: '1px solid rgba(185, 167, 255, 0.3)',
              display: 'flex',
              flexDirection: 'column',
              justifyContent: 'space-between',
            }}
          >
            {el.chartTitle && (
              <div className="text-[16px] font-bold text-[#4A3F6E] mb-3">
                {el.chartTitle}
              </div>
            )}
            <div className="flex-1 flex flex-col justify-around gap-3">
              {el.chartBars.map((bar, bIdx) => {
                const pct = Math.max(8, Math.round((bar.value / maxVal) * 100));
                return (
                  <div key={bIdx} className="flex items-center gap-4">
                    <div className="w-56 text-[14px] font-bold text-[#2A2244] truncate">
                      {bar.label}
                    </div>
                    <div className="flex-1 h-9 rounded-full bg-[#F1EEFF] p-1 flex items-center shadow-inner">
                      <div
                        style={{
                          width: `${pct}%`,
                          backgroundColor: bar.color,
                          boxShadow:
                            'inset 3px 3px 6px rgba(255,255,255,0.65), inset -3px -3px 6px rgba(0,0,0,0.12)',
                        }}
                        className="h-full rounded-full"
                      />
                    </div>
                    <div className="w-20 text-right font-mono-num font-bold text-[15px] text-[#2A2244]">
                      {bar.value}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        );
      }

      // Shape or Text element
      const isShapeWithFill = Boolean(el.fillColor);
      const isLine = el.shapeKind === 'line';

      if (isLine) {
        return (
          <div
            key={el.id}
            style={{
              ...commonStyle,
              height: `${Math.max(2, el.strokeWidth || 2)}px`,
              backgroundColor: el.strokeColor || el.fillColor || '#8C82AD',
            }}
          />
        );
      }

      const radius =
        el.shapeKind === 'ellipse' || el.shapeKind === 'pill'
          ? '9999px'
          : `${el.borderRadius ?? 0}px`;

      const resolvedPadding = el.padding
        ? `${el.padding.top}px ${el.padding.right}px ${el.padding.bottom}px ${el.padding.left}px`
        : isShapeWithFill && el.paragraphs && el.paragraphs.length > 0
          ? el.shapeKind === 'pill'
            ? '6px 16px'
            : '24px 28px'
          : '2px 4px';

      const justifyContent =
        el.verticalAlign === 'middle' || el.shapeKind === 'pill'
          ? 'center'
          : el.verticalAlign === 'bottom'
            ? 'flex-end'
            : 'flex-start';

      const clayShapeShadow =
        isSample && isShapeWithFill && !el.padding
          ? '10px 10px 22px rgba(110, 90, 170, 0.16), -6px -6px 16px rgba(255, 255, 255, 0.85), inset 5px 5px 10px rgba(255, 255, 255, 0.7), inset -5px -5px 10px rgba(140, 120, 200, 0.2)'
          : undefined;

      return (
        <div
          key={el.id}
          style={{
            ...commonStyle,
            backgroundColor: el.fillColor || 'transparent',
            borderRadius: radius,
            border: el.strokeColor
              ? `${el.strokeWidth || 1.5}px solid ${el.strokeColor}`
              : isSample && isShapeWithFill && !el.padding
                ? '1px solid rgba(255,255,255,0.6)'
                : undefined,
            boxShadow: clayShapeShadow,
            padding: resolvedPadding,
            display: 'flex',
            flexDirection: 'column',
            justifyContent,
            gap: el.padding ? '2px' : '10px',
            overflow: isShapeWithFill ? 'hidden' : 'visible',
          }}
        >
          {el.paragraphs?.map((para, pIdx) => (
            <div
              key={pIdx}
              style={{
                textAlign: para.align || 'left',
                lineHeight: para.lineHeight || 1.24,
                paddingLeft: para.level ? `${para.level * 18}px` : undefined,
                whiteSpace: el.wordWrap === false ? 'nowrap' : 'pre-wrap',
                wordBreak: 'break-word',
              }}
              className={para.bullet ? 'flex items-start gap-2' : ''}
            >
              {para.bullet && (
                <span
                  className="inline-block shrink-0 font-bold select-none"
                  style={{
                    color:
                      para.runs[0]?.color ||
                      slide.background.accentBlobColor ||
                      '#7C63E6',
                    fontSize: `${para.runs[0]?.fontSize || 15}px`,
                    lineHeight: 1.24,
                  }}
                >
                  {para.bulletChar || '•'}
                </span>
              )}
              <div className="flex-1 min-w-0">
                {para.runs.map((run, rIdx) => (
                  <span
                    key={rIdx}
                    style={{
                      fontSize: `${run.fontSize || 16}px`,
                      fontWeight: run.bold ? 700 : 400,
                      fontStyle: run.italic ? 'italic' : 'normal',
                      textDecoration: run.underline ? 'underline' : 'none',
                      color: run.color || '#251D3D',
                      fontFamily: isSample
                        ? (run.fontSize || 16) >= 24
                          ? "'Fredoka', 'Nunito', sans-serif"
                          : "'Nunito', sans-serif"
                        : run.fontFamily
                          ? `'${run.fontFamily}', Calibri, 'Nunito', sans-serif`
                          : "Calibri, 'Nunito', sans-serif",
                    }}
                  >
                    {run.text}
                  </span>
                ))}
              </div>
            </div>
          ))}
        </div>
      );
    };

    const scaleFactor =
      containerWidth > 0 ? containerWidth / activeWidth : 1;

    return (
      <div
        ref={outerContainerRef}
        className={`relative w-full overflow-hidden select-none ${className}`}
        style={{
          aspectRatio: `${computedRatio}`,
          backgroundColor: slide.background.color || '#FFFFFF',
        }}
      >
        {/* Vector Fallback / Built-in Sample Layer */}
        <svg
          viewBox={`0 0 ${baseWidth} ${baseHeight}`}
          className={`w-full h-full block ${
            nativeReady && !isSample ? 'opacity-0 pointer-events-none' : ''
          }`}
          preserveAspectRatio="xMidYMid meet"
        >
          <foreignObject x="0" y="0" width={baseWidth} height={baseHeight}>
            <div
              style={{
                width: `${baseWidth}px`,
                height: `${baseHeight}px`,
                backgroundColor: slide.background.color || '#FAF8FF',
                backgroundImage: slide.background.imageUrl
                  ? `url(${slide.background.imageUrl})`
                  : undefined,
                backgroundSize: '100% 100%',
                backgroundPosition: 'center',
              }}
              className="relative overflow-hidden"
            >
              {slide.background.accentBlobColor && isSample && (
                <div
                  aria-hidden="true"
                  style={{
                    position: 'absolute',
                    right: '-80px',
                    top: '-80px',
                    width: '280px',
                    height: '280px',
                    borderRadius: '9999px',
                    backgroundColor: slide.background.accentBlobColor,
                    opacity: 0.22,
                    filter: 'blur(45px)',
                    pointerEvents: 'none',
                  }}
                />
              )}

              {slide.elements.map(renderElement)}
            </div>
          </foreignObject>
        </svg>

        {/* Native 1:1 `pptx-preview` Layer (for all uploaded .pptx presentations) */}
        {!isSample && (
          <div
            ref={nativeHostRef}
            style={{
              position: 'absolute',
              left: 0,
              top: 0,
              width: `${activeWidth}px`,
              height: `${activeHeight}px`,
              transform: `scale(${scaleFactor})`,
              transformOrigin: '0 0',
              opacity: nativeReady ? 1 : 0,
              pointerEvents: isThumbnail ? 'none' : 'auto',
            }}
            className="overflow-hidden"
          />
        )}
      </div>
    );
  }
);

SlideCanvas.displayName = 'SlideCanvas';
