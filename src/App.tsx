import React, { useState, useEffect, useCallback, useRef, useMemo, memo } from 'react';
import { 
  Upload, FileText, Check, X, Trash2, Eye, EyeOff, 
  Layers, Plus, Scissors, Lock, Download, AlertCircle,
  History, Link, ChevronDown, ChevronUp, ChevronRight, Search, Filter,
  MoreVertical, Copy, CheckCircle2, User, MapPin, Phone, 
  CreditCard, Mail, Hash, Briefcase, Scale, Trash, RotateCcw, RotateCw,
  Shield, Save, FolderOpen, XCircle, Zap, Unlink, Type, List, Pencil, RefreshCw, Building2,
  ChevronLeft, Crop
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import * as pdfjs from 'pdfjs-dist';
// @ts-ignore
import pdfjsWorker from 'pdfjs-dist/build/pdf.worker.mjs?url';
import { PDFDocument, rgb, StandardFonts } from 'pdf-lib';
import mammoth from 'mammoth';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import JSZip from 'jszip';
import { 
  Document as DocxDocument, 
  Packer, 
  Paragraph, 
  TextRun, 
  Table, 
  TableRow, 
  TableCell, 
  AlignmentType, 
  HeadingLevel,
  FootnoteReferenceRun,
  UnderlineType,
  BorderStyle,
  WidthType,
  LineRuleType
} from 'docx';
import { jsPDF } from 'jspdf';
import html2canvas from 'html2canvas';
import 'jspdf-autotable';
import { 
  scanText, 
  groupSimilarEntities, 
  anonymizeText,
  splitEntity,
  getNextPseudonym,
  cleanName,
  superNormalize,
  generateId,
  deduplicateEntities,
  PIIEntity,
  PII_COLORS,
  Safelist
} from './lib/anonymizer';

const ORDERED_PII_TYPES = (() => {
  const types = Object.keys(PII_COLORS);
  const priority = ['NOME', 'ADVOGADO', 'JUIZ', 'AUTOR', 'COLETIVA', 'HEADER'];
  
  return types.sort((a, b) => {
    const aIdx = priority.indexOf(a);
    const bIdx = priority.indexOf(b);
    
    if (aIdx !== -1 && bIdx !== -1) return aIdx - bIdx;
    if (aIdx !== -1) return -1;
    if (bIdx !== -1) return 1;
    
    return a.localeCompare(b);
  });
})();

// Set up PDF.js worker
// @ts-ignore
pdfjs.GlobalWorkerOptions.workerSrc = pdfjsWorker;

interface FileData {
  id: string;
  name: string;
  type: string;
  content: string;
  htmlContent?: string;
  positions?: { start: number, end: number, page: number, y: number }[];
  rawFile: File;
  status: 'pending' | 'processing' | 'done' | 'error';
}

interface ExportSettings {
  format: 'pdf' | 'docx' | 'txt';
  removeHeadersFooters: boolean;
  footnotesAtEnd: boolean;
  preserveFormatting: boolean;
  sanitizeMetadata: boolean;
}

const PDFPage = memo(({ 
  file, 
  pageNum, 
  allEntities, 
  selectedEntityId, 
  selectedIds, 
  onVisible,
  selectionMode,
  onAreaSelected,
  headerZones,
  footerZones
}: { 
  file: File, 
  pageNum: number, 
  allEntities: PIIEntity[],
  selectedEntityId?: string | null,
  selectedIds?: Set<string>,
  onVisible?: () => void,
  selectionMode: 'none' | 'header' | 'footer',
  onAreaSelected: (rect: { top: number; bottom: number; left: number; right: number }, pageNum: number, image?: string) => void,
  headerZones: { top: number; bottom: number; left: number; right: number; image?: string }[],
  footerZones: { top: number; bottom: number; left: number; right: number; image?: string }[]
}) => {
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const containerRef = useRef<HTMLDivElement>(null);
  const [markerPos, setMarkerPos] = useState<{ x: number, y: number } | null>(null);
  const [isDrawing, setIsDrawing] = useState(false);
  const [startPos, setStartPos] = useState<{ x: number, y: number } | null>(null);
  const [currentRect, setCurrentRect] = useState<{ x: number, y: number, w: number, h: number } | null>(null);
  const [viewport, setViewport] = useState<any>(null);

  useEffect(() => {
    const observer = new IntersectionObserver((entries) => {
      if (entries[0].isIntersecting) onVisible?.();
    }, { threshold: 0.5 });
    if (containerRef.current) observer.observe(containerRef.current);
    return () => observer.disconnect();
  }, [onVisible]);

  useEffect(() => {
    let isMounted = true;
    const render = async () => {
      try {
        if (!file) return;
        
        // Robust file object detection
        let fileToUse: any = file;
        if (typeof fileToUse.arrayBuffer !== 'function' && fileToUse.rawFile) {
          fileToUse = fileToUse.rawFile;
        }

        if (!fileToUse || typeof fileToUse.arrayBuffer !== 'function') {
          console.error("Invalid file object in PDFPage:", fileToUse);
          return;
        }

        const arrayBuffer = await fileToUse.arrayBuffer();
        const pdf = await pdfjs.getDocument({ data: arrayBuffer }).promise;
        const page = await pdf.getPage(pageNum);
        const vp = page.getViewport({ scale: 1.5 });
        if (isMounted) setViewport(vp);
        
        const canvas = canvasRef.current;
        if (!canvas || !isMounted) return;
        const context = canvas.getContext('2d');
        if (!context) return;
        
        canvas.height = vp.height;
        canvas.width = vp.width;
        
        await (page as any).render({ canvasContext: context, viewport: vp }).promise;
        
        const textContent = await page.getTextContent();
        
        // 1. Highlight all detected entities (subtle)
        const activeEntities = allEntities.filter(e => (e.enabled && !e.ignored) || e.id === selectedEntityId || selectedIds?.has(e.id));
        
        let foundMarker = false;
        activeEntities.forEach(entity => {
          const isSelected = entity.id === selectedEntityId || selectedIds?.has(entity.id);
          const term = entity.original.toLowerCase();
          
          const colorInfo = PII_COLORS[entity.type] || { hex: '#6366F1' };
          const baseColor = colorInfo.hex;
          const r = parseInt(baseColor.substring(1, 3), 16) || 99;
          const g = parseInt(baseColor.substring(3, 5), 16) || 102;
          const b = parseInt(baseColor.substring(5, 7), 16) || 241;

          context.fillStyle = isSelected 
            ? 'rgba(254, 240, 138, 0.55)' 
            : entity.treated 
              ? 'rgba(16, 185, 129, 0.14)' 
              : `rgba(${r}, ${g}, ${b}, 0.12)`;
          
          textContent.items.forEach((item: any) => {
            const itemStrLower = item.str.toLowerCase();
            let startIndex = itemStrLower.indexOf(term);
            while (startIndex !== -1) {
              const [x, y] = vp.convertToViewportPoint(item.transform[4], item.transform[5]);
              const fontHeight = Math.sqrt(item.transform[2] * item.transform[2] + item.transform[3] * item.transform[3]);
              
              // Proportional estimation of match start and match width
              const totalLen = item.str.length || 1;
              const charWidthScale = (item.width * vp.scale) / totalLen;
              
              const matchX = x + (startIndex * charWidthScale);
              const matchWidth = term.length * charWidthScale;
              
              context.fillStyle = isSelected 
                ? 'rgba(254, 240, 138, 0.55)' 
                : entity.treated 
                  ? 'rgba(16, 185, 129, 0.14)' 
                  : `rgba(${r}, ${g}, ${b}, 0.12)`;

              context.fillRect(
                matchX, 
                y - (fontHeight * vp.scale), 
                matchWidth, 
                fontHeight * vp.scale * 1.2
              );

              // Draw a very subtle outer stroke for the bounding box
              context.strokeStyle = isSelected
                ? 'rgba(234, 179, 8, 0.8)'
                : entity.treated
                  ? 'rgba(16, 185, 129, 0.4)'
                  : `rgba(${r}, ${g}, ${b}, 0.35)`;
              context.lineWidth = 1;
              context.strokeRect(
                matchX, 
                y - (fontHeight * vp.scale), 
                matchWidth, 
                fontHeight * vp.scale * 1.2
              );
              
              if (isSelected) {
                if (!foundMarker && entity.id === selectedEntityId) {
                  setMarkerPos({ x: matchX, y: y - (fontHeight * vp.scale) });
                  foundMarker = true;
                }
              }

              // Check for subsequent occurrences in the same item line
              startIndex = itemStrLower.indexOf(term, startIndex + 1);
            }
          });
        });

        // Draw existing zones
        headerZones.forEach(zone => {
          context.fillStyle = 'rgba(239, 68, 68, 0.2)'; // Red for header
          context.strokeStyle = 'rgba(239, 68, 68, 0.5)';
          context.lineWidth = 1;
          context.fillRect(
            zone.left * vp.width,
            zone.top * vp.height,
            (zone.right - zone.left) * vp.width,
            (zone.bottom - zone.top) * vp.height
          );
          context.strokeRect(
            zone.left * vp.width,
            zone.top * vp.height,
            (zone.right - zone.left) * vp.width,
            (zone.bottom - zone.top) * vp.height
          );
        });

        footerZones.forEach(zone => {
          context.fillStyle = 'rgba(59, 130, 246, 0.2)'; // Blue for footer
          context.strokeStyle = 'rgba(59, 130, 246, 0.5)';
          context.lineWidth = 1;
          context.fillRect(
            zone.left * vp.width,
            zone.top * vp.height,
            (zone.right - zone.left) * vp.width,
            (zone.bottom - zone.top) * vp.height
          );
          context.strokeRect(
            zone.left * vp.width,
            zone.top * vp.height,
            (zone.right - zone.left) * vp.width,
            (zone.bottom - zone.top) * vp.height
          );
        });

        if (!foundMarker) setMarkerPos(null);
      } catch (err) {
        console.error("Error rendering PDF page:", err);
      }
    };
    render();
    return () => { isMounted = false; };
  }, [file, pageNum, allEntities, selectedEntityId, selectedIds, headerZones, footerZones]);

  const handleMouseDown = (e: React.MouseEvent) => {
    if (selectionMode === 'none') return;
    const rect = containerRef.current?.getBoundingClientRect();
    if (!rect) return;
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;
    setIsDrawing(true);
    setStartPos({ x, y });
    console.log(`MouseDown at ${x}, ${y} (Mode: ${selectionMode})`);
  };

  const handleMouseMove = (e: React.MouseEvent) => {
    if (!isDrawing || !startPos) return;
    const rect = containerRef.current?.getBoundingClientRect();
    if (!rect) return;
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;
    setCurrentRect({
      x: Math.min(x, startPos.x),
      y: Math.min(y, startPos.y),
      w: Math.abs(x - startPos.x),
      h: Math.abs(y - startPos.y)
    });
  };

  const handleMouseUp = () => {
    if (!isDrawing || !currentRect || !viewport) {
      setIsDrawing(false);
      setStartPos(null);
      setCurrentRect(null);
      return;
    }

    const rect = containerRef.current?.getBoundingClientRect();
    if (!rect) return;

    // Convert to percentage coordinates based on DISPLAYED size
    const zoneRect = {
      top: currentRect.y / rect.height,
      bottom: (currentRect.y + currentRect.h) / rect.height,
      left: currentRect.x / rect.width,
      right: (currentRect.x + currentRect.w) / rect.width
    };

    // Capture screenshot of the selected area
    let screenshot: string | undefined;
    const canvas = canvasRef.current;
    if (canvas) {
      try {
        const tempCanvas = document.createElement('canvas');
        const sx = currentRect.x * (canvas.width / rect.width);
        const sy = currentRect.y * (canvas.height / rect.height);
        const sw = currentRect.w * (canvas.width / rect.width);
        const sh = currentRect.h * (canvas.height / rect.height);
        
        tempCanvas.width = sw;
        tempCanvas.height = sh;
        const tempCtx = tempCanvas.getContext('2d');
        if (tempCtx) {
          tempCtx.drawImage(canvas, sx, sy, sw, sh, 0, 0, sw, sh);
          screenshot = tempCanvas.toDataURL('image/png');
        }
      } catch (err) {
        console.error("Error capturing zone screenshot:", err);
      }
    }

    console.log('Area Selected (Percentage):', zoneRect, 'Page:', pageNum);
    onAreaSelected(zoneRect, pageNum, screenshot);
    setIsDrawing(false);
    setStartPos(null);
    setCurrentRect(null);
  };

  return (
    <div 
      ref={containerRef} 
      className={`relative ${selectionMode !== 'none' ? 'cursor-crosshair' : ''}`}
      onMouseDown={handleMouseDown}
      onMouseMove={handleMouseMove}
      onMouseUp={handleMouseUp}
    >
      <canvas ref={canvasRef} className="max-w-full h-auto" />
      {currentRect && (
        <div 
          className="absolute border-2 pointer-events-none z-10"
          style={{
            left: currentRect.x,
            top: currentRect.y,
            width: currentRect.w,
            height: currentRect.h,
            borderColor: selectionMode === 'header' ? 'rgba(239, 68, 68, 0.8)' : 'rgba(59, 130, 246, 0.8)',
            backgroundColor: selectionMode === 'header' ? 'rgba(239, 68, 68, 0.1)' : 'rgba(59, 130, 246, 0.1)'
          }}
        />
      )}
      {markerPos && (
        <div 
          id="pdf-active-marker"
          className="absolute pointer-events-none"
          style={{ 
            left: markerPos.x, 
            top: markerPos.y,
            width: 10,
            height: 10
          }}
        />
      )}
    </div>
  );
});

const DocumentViewer = memo(({ 
  file, 
  entities, 
  selectedEntityId, 
  selectedIds, 
  setPendingManualTerm, 
  globalKnowledge, 
  safelist,
  selectionMode,
  onAreaSelected,
  headerZones,
  footerZones,
  onFileUpload
}: { 
  file: FileData | null, 
  entities: PIIEntity[],
  selectedEntityId?: string | null,
  selectedIds: Set<string>,
  setPendingManualTerm: (term: { text: string, x: number, y: number } | null) => void,
  globalKnowledge: Record<string, string>,
  safelist: Safelist,
  selectionMode: 'none' | 'header' | 'footer',
  onAreaSelected: (rect: { top: number; bottom: number; left: number; right: number }, pageNum: number, image?: string) => void,
  headerZones: { top: number; bottom: number; left: number; right: number; image?: string }[],
  footerZones: { top: number; bottom: number; left: number; right: number; image?: string }[],
  onFileUpload: (event: React.ChangeEvent<HTMLInputElement>) => void
}) => {
  const [pages, setPages] = useState<number[]>([]);
  const containerRef = useRef<HTMLDivElement>(null);
  const [currentPage, setCurrentPage] = useState(1);
  const [numPages, setNumPages] = useState(0);
  const [viewMode, setViewMode] = useState<'pdf' | 'text' | 'split'>('split');

  const isPDF = file?.type?.includes('pdf') || file?.name?.toLowerCase().endsWith('.pdf');

  useEffect(() => {
    if (!file || !isPDF) {
      setPages([]);
      setNumPages(0);
      return;
    }
    
    const loadPDF = async () => {
      try {
        if (!file.rawFile || typeof file.rawFile.arrayBuffer !== 'function') {
          return; // Silent return as it might be a restored state
        }
        const arrayBuffer = await file.rawFile.arrayBuffer();
        const pdf = await pdfjs.getDocument({ data: arrayBuffer }).promise;
        setNumPages(pdf.numPages);
        const p = [];
        for (let i = 1; i <= pdf.numPages; i++) p.push(i);
        setPages(p);
      } catch (err) {
        console.error("Error loading PDF for viewer:", err);
      }
    };
    loadPDF();
  }, [file, isPDF]);

  useEffect(() => {
    const selectedEntity = entities.find(e => e.id === selectedEntityId);
    if (!selectedEntity || !containerRef.current || !file) return;
    
    const findAndScroll = async () => {
      if (isPDF && viewMode === 'pdf') {
        try {
          if (!file.rawFile || typeof file.rawFile.arrayBuffer !== 'function') return;
          const arrayBuffer = await file.rawFile.arrayBuffer();
          const pdf = await pdfjs.getDocument({ data: arrayBuffer }).promise;
          const term = selectedEntity.original.toLowerCase();
          
          for (let i = 1; i <= pdf.numPages; i++) {
            const page = await pdf.getPage(i);
            const textContent = await page.getTextContent();
            const hasTerm = textContent.items.some((item: any) => item.str.toLowerCase().includes(term));
            
            if (hasTerm) {
              const pageEl = containerRef.current?.querySelector(`[data-page-number="${i}"]`);
              if (pageEl) {
                pageEl.scrollIntoView({ behavior: 'smooth', block: 'center' });
                setCurrentPage(i);
                
                // Try to find the specific marker after page scroll
                setTimeout(() => {
                  const marker = pageEl.querySelector('#pdf-active-marker');
                  if (marker) {
                    marker.scrollIntoView({ behavior: 'smooth', block: 'center' });
                  }
                }, 300);
                break;
              }
            }
          }
        } catch (err) {
          console.error("Error searching in PDF:", err);
        }
      } else {
        // Text mode scroll
        const scrollToActive = () => {
          const el = containerRef.current?.querySelector('#active-highlight');
          if (el) {
            el.scrollIntoView({ behavior: 'smooth', block: 'center' });
          }
        };
        
        // Try multiple times as rendering might be async
        setTimeout(scrollToActive, 100);
        setTimeout(scrollToActive, 300);
        setTimeout(scrollToActive, 600);
      }
    };
    findAndScroll();
  }, [selectedEntityId, file, isPDF, entities, viewMode]);

  if (!file) return (
    <div className="flex-1 flex flex-col items-center justify-center p-12 text-center bg-gray-50">
      <FileText className="w-12 h-12 text-gray-300 mb-4" />
      <p className="text-sm text-gray-500 italic">Selecione um ficheiro para visualizar o original</p>
    </div>
  );

  return (
    <div className="flex-1 flex flex-col bg-gray-100 overflow-hidden">
      <div className="bg-white border-b border-gray-200 px-4 py-2.5 flex items-center justify-between shadow-sm z-10 shrink-0">
        <div className="flex items-center gap-4">
          <div className="flex items-center gap-2">
            <FileText className="w-4 h-4 text-indigo-500" />
            <span className="text-xs font-bold truncate max-w-[150px]">{file.name}</span>
          </div>
          {isPDF && (
            <div className="flex bg-gray-100 p-0.5 rounded-lg border border-gray-200">
              <button 
                onClick={() => setViewMode('pdf')}
                className={`px-3 py-1 text-[9px] font-bold uppercase tracking-wider rounded-md transition-all ${viewMode === 'pdf' ? 'bg-white shadow-sm text-indigo-600' : 'text-gray-500 hover:text-gray-700'}`}
              >
                Forma (PDF Original)
              </button>
              <button 
                onClick={() => setViewMode('split')}
                className={`px-3 py-1 text-[9px] font-bold uppercase tracking-wider rounded-md transition-all ${viewMode === 'split' ? 'bg-white shadow-sm text-indigo-600' : 'text-gray-500 hover:text-gray-700'}`}
              >
                Consolidado (Lado a Lado)
              </button>
              <button 
                onClick={() => setViewMode('text')}
                className={`px-3 py-1 text-[9px] font-bold uppercase tracking-wider rounded-md transition-all ${viewMode === 'text' ? 'bg-white shadow-sm text-indigo-600' : 'text-gray-500 hover:text-gray-700'}`}
              >
                Conteúdo (Modo Texto)
              </button>
            </div>
          )}
        </div>
        {isPDF && viewMode === 'pdf' && (
          <div className="flex items-center gap-3 text-[10px] font-bold text-gray-500 uppercase tracking-wider">
            <span>Página {currentPage} / {numPages}</span>
          </div>
        )}
      </div>
      
      <div 
        ref={containerRef} 
        className={`flex-1 overflow-hidden flex ${viewMode === 'split' ? 'flex-col md:flex-row p-2 gap-4 bg-gray-200/50' : 'flex-col overflow-y-auto p-4 space-y-6 scroll-smooth bg-gray-200/50'}`}
        onMouseUp={(e) => {
          if (viewMode !== 'pdf' || !isPDF) {
            const selection = window.getSelection();
            const text = selection?.toString().trim();
            if (text && text.length > 1) {
              setPendingManualTerm({
                text,
                x: e.clientX,
                y: e.clientY
              });
            }
          }
        }}
      >
        {viewMode === 'split' && isPDF ? (
          <>
            {/* Split view: Left hand PDF Page original shape */}
            <div className="flex-1 h-full min-h-0 overflow-y-auto p-4 space-y-6 rounded-xl border border-gray-200/80 bg-gray-200/30 shadow-inner">
              <div className="p-2 bg-indigo-50/90 rounded-lg text-[10px] font-bold text-indigo-600 uppercase tracking-wider mb-2 text-center sticky top-0 backdrop-blur-sm z-10 border border-indigo-100 flex items-center justify-between px-3 shadow-xs">
                <span>Forma (PDF Original)</span>
                {pages.length > 0 && <span className="text-[9px] bg-indigo-200 text-indigo-800 px-1.5 py-0.5 rounded">Pág. {currentPage} / {numPages}</span>}
              </div>
              {pages.length > 0 ? (
                pages.map(pageNum => (
                  <div 
                    key={pageNum} 
                    data-page-number={pageNum}
                    className="bg-white shadow-lg mx-auto relative group rounded-sm overflow-hidden mb-4"
                    style={{ width: 'fit-content' }}
                  >
                    <PDFPage 
                      file={file.rawFile} 
                      pageNum={pageNum} 
                      allEntities={entities}
                      selectedEntityId={selectedEntityId}
                      selectedIds={selectedIds}
                      onVisible={() => setCurrentPage(pageNum)}
                      selectionMode={selectionMode}
                      onAreaSelected={onAreaSelected}
                      headerZones={headerZones}
                      footerZones={footerZones}
                    />
                    <div className="absolute top-2 left-2 bg-black/60 text-white text-[9px] font-bold px-2 py-1 rounded backdrop-blur-sm opacity-0 group-hover:opacity-100 transition-opacity">
                      PÁG. {pageNum}
                    </div>
                  </div>
                ))
              ) : (
                <div className="flex flex-col items-center justify-center py-10 text-center bg-white rounded-xl border border-dashed border-gray-300 max-w-sm mx-auto mt-4">
                  <AlertCircle className="w-6 h-6 text-amber-500 mb-2" />
                  <p className="text-xs text-gray-500 px-4">Layout original indisponível nesta restauração.</p>
                </div>
              )}
            </div>

            {/* Split view: Right hand clean document contents */}
            <div className="flex-1 h-full min-h-0 overflow-y-auto p-6 md:p-8 rounded-xl border border-gray-200 bg-white shadow-sm flex flex-col">
              <div className="p-2 bg-indigo-50/90 rounded-lg text-[10px] font-bold text-indigo-600 uppercase tracking-wider mb-4 text-center sticky top-0 backdrop-blur-sm z-10 border border-indigo-100 shadow-xs">
                Conteúdo (Modo Texto & Seleção Manual)
              </div>
              <div className="flex-1 min-h-0 prose prose-sm max-w-none font-mono text-sm whitespace-pre-wrap text-gray-700">
                <HighlightText 
                  text={file.content} 
                  entities={[
                    ...entities,
                    ...(selectedEntityId ? [{ 
                      id: 'selected-highlight', 
                      original: entities.find(e => e.id === selectedEntityId)?.original || '', 
                      type: 'ACTIVE_HIGHLIGHT', 
                      pseudonym: '', 
                      enabled: true 
                    }] : []),
                    ...Array.from(selectedIds).map(id => ({
                      id: `selected-id-${id}`,
                      original: entities.find(e => e.id === id)?.original || '',
                      type: 'HIGHLIGHT',
                      pseudonym: '',
                      enabled: true
                    }))
                  ]} 
                  mode="original" 
                  globalKnowledge={globalKnowledge}
                  safelist={safelist}
                />
              </div>
            </div>
          </>
        ) : (
          isPDF && viewMode === 'pdf' ? (
            pages.length > 0 ? (
              pages.map(pageNum => (
                <div 
                  key={pageNum} 
                  data-page-number={pageNum}
                  className="bg-white shadow-lg mx-auto relative group rounded-sm overflow-hidden"
                  style={{ width: 'fit-content' }}
                >
                  <PDFPage 
                    file={file.rawFile} 
                    pageNum={pageNum} 
                    allEntities={entities}
                    selectedEntityId={selectedEntityId}
                    selectedIds={selectedIds}
                    onVisible={() => setCurrentPage(pageNum)}
                    selectionMode={selectionMode}
                    onAreaSelected={onAreaSelected}
                    headerZones={headerZones}
                    footerZones={footerZones}
                  />
                  <div className="absolute top-2 left-2 bg-black/60 text-white text-[9px] font-bold px-2 py-1 rounded backdrop-blur-sm opacity-0 group-hover:opacity-100 transition-opacity">
                    PÁG. {pageNum}
                  </div>
                </div>
              ))
            ) : (
              <div className="flex flex-col items-center justify-center py-20 text-center bg-white rounded-xl border border-dashed border-gray-300 max-w-lg mx-auto mt-8">
                <div className="w-16 h-16 bg-amber-100 rounded-full flex items-center justify-center text-amber-600 mb-4">
                  <AlertCircle className="w-8 h-8" />
                </div>
                <h3 className="text-lg font-bold text-gray-900 leading-tight">Layout original indisponível</h3>
                <p className="text-sm text-gray-500 mt-2 px-8">
                  Para visualizar o layout original deste PDF e exportar com alta fidelidade após uma restauração de sessão, é necessário recarregar o ficheiro original.
                </p>
                <label className="mt-6 px-6 py-2 bg-indigo-600 text-white rounded-lg font-medium cursor-pointer hover:bg-indigo-700 transition-colors inline-block">
                  Recarregar Ficheiro
                  <input type="file" className="hidden" accept=".pdf" onChange={(e) => {
                    const f = e.target.files?.[0];
                    if (f) onFileUpload(e);
                  }} />
                </label>
              </div>
            )
          ) : (
            <div className="bg-white p-10 shadow-sm mx-auto max-w-4xl min-h-full rounded-xl border border-gray-200 relative">
               <div className="prose prose-sm max-w-none font-mono text-sm whitespace-pre-wrap text-gray-700">
                 <HighlightText 
                   text={file.content} 
                   entities={[
                     ...entities,
                     ...(selectedEntityId ? [{ 
                       id: 'selected-highlight', 
                       original: entities.find(e => e.id === selectedEntityId)?.original || '', 
                       type: 'ACTIVE_HIGHLIGHT', 
                       pseudonym: '', 
                       enabled: true 
                     }] : []),
                     ...Array.from(selectedIds).map(id => ({
                       id: `selected-id-${id}`,
                       original: entities.find(e => e.id === id)?.original || '',
                       type: 'HIGHLIGHT',
                       pseudonym: '',
                       enabled: true
                     }))
                   ]} 
                   mode="original" 
                   globalKnowledge={globalKnowledge}
                   safelist={safelist}
                 />
               </div>
            </div>
          )
        )}
      </div>
    </div>
  );
});

const HighlightText = memo(({ text, entities, mode, globalKnowledge, safelist }: { 
  text: string, 
  entities: PIIEntity[], 
  mode: 'original' | 'anonymized',
  globalKnowledge?: Record<string, string>,
  safelist?: Safelist
}) => {
  if (!text) return null;
  
  const result = useMemo(() => {
    // Filter entities that are actually being anonymized
    const activeEntities = [
      ...entities.filter(e => e.type === 'ACTIVE_HIGHLIGHT' || e.type === 'HIGHLIGHT'),
      ...entities.filter(e => (e.enabled && !e.ignored && e.type !== 'HIGHLIGHT' && e.type !== 'ACTIVE_HIGHLIGHT'))
    ];
    
    const patterns = activeEntities.map(e => ({
      pattern: mode === 'original' ? e.original : e.pseudonym,
      entity: e
    })).filter(p => p.pattern.length > 0);

    // Identify potential untreated PII in original mode
    const potentialUntreated: { text: string, type: string }[] = [];
    if (mode === 'original') {
      const untreatedPatterns: Record<string, RegExp> = {
        NIF: /\b\d{9}\b/g,
        CC: /\b\d{8}\s*\d\s*[A-Z]{2}\d\b/gi,
        MATRICULA: /\b[A-Z0-9]{2}-[A-Z0-9]{2}-[A-Z0-9]{2}\b/g,
        MORADA: /\b(?:Rua|Av\.|Avenida|Praça|Largo|Travessa)\s+[^,.;\n]{5,50}\b/gi
      };

      Object.entries(untreatedPatterns).forEach(([type, regex]) => {
        let match;
        while ((match = regex.exec(text)) !== null) {
          const matchText = match[0];
          const norm = matchText.toLowerCase().trim();
          
          const inEntities = entities.some(e => e.original.toLowerCase().trim() === norm);
          const isException = globalKnowledge?.[norm] === 'EXCECAO';
          const inSafelist = safelist?.words_ignore.some(w => w.toLowerCase() === norm) || 
                            safelist?.phrases_ignore.some(p => norm.includes(p.toLowerCase()));

          if (!inEntities && !isException && !inSafelist) {
            potentialUntreated.push({ text: matchText, type });
          }
        }
      });
    }

    const allPatterns = [
      ...patterns.map(p => ({ pattern: p.pattern, type: p.entity.type, entity: p.entity })),
      ...potentialUntreated.map(p => ({ pattern: p.text, type: p.type, entity: null }))
    ];

    if (allPatterns.length === 0) return { parts: [text], allPatterns: [] };

    allPatterns.sort((a, b) => b.pattern.length - a.pattern.length);
    const uniquePatterns = Array.from(new Set(allPatterns.map(p => p.pattern.toLowerCase())));
    const escapeRegExp = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    
    // Use word boundaries for short patterns to avoid matching inside other words (e.g. "os" in "instrumentos")
    // For longer patterns, word boundaries are also generally safer.
    const wordChars = 'a-zA-ZÀ-ÿ0-9_';
    const patternStrings = uniquePatterns.map(p => {
      const escaped = escapeRegExp(p);
      // If it's a simple word-like pattern, wrap in word boundaries that respect Unicode
      if (/^[a-z0-9À-ÿ]+$/i.test(p)) {
        return `(?<![${wordChars}])${escaped}(?![${wordChars}])`;
      }
      return escaped;
    });

    const combinedRegex = new RegExp(`(${patternStrings.join('|')})`, 'gi');
    
    return { parts: text.split(combinedRegex), allPatterns };
  }, [text, entities, mode, globalKnowledge, safelist]);

  return (
    <>
      {result.parts.map((part, i) => {
        const partLower = part.toLowerCase();
        const match = result.allPatterns.find(p => p.pattern.toLowerCase() === partLower);
        if (match) {
          if (match.entity) {
            const color = (match.entity.type === 'HIGHLIGHT' || match.entity.type === 'ACTIVE_HIGHLIGHT')
              ? { hex: '#FDE047', textHex: '#854D0E' } 
              : PII_COLORS[match.entity.type] || { hex: '#E5E7EB', textHex: '#374151' };
            const isTreated = match.entity.treated;
            return (
              <span 
                key={i} 
                id={match.entity.type === 'ACTIVE_HIGHLIGHT' ? 'active-highlight' : undefined}
                className={`transition-all duration-300 font-semibold text-xs ${match.entity.type === 'ACTIVE_HIGHLIGHT' ? 'ring-2 ring-yellow-400 animate-pulse' : ''}`}
                style={isTreated ? { 
                  backgroundColor: '#ecfdf5', 
                  color: '#047857',
                  border: '1px solid rgba(16, 185, 129, 0.4)',
                  borderRadius: '4px',
                  padding: '1px 4px',
                  margin: '0 1px'
                } : { 
                  backgroundColor: `${color.hex}18`, 
                  color: (color.textHex === '#FFFFFF' || color.textHex === '#ffffff') ? '#374151' : color.textHex,
                  border: `1px solid ${color.hex}80`,
                  borderRadius: '4px',
                  padding: '1px 4px',
                  margin: '0 1px'
                }}
                title={match.entity.type === 'ACTIVE_HIGHLIGHT' ? 'Elemento Selecionado' : `${match.entity.type}: ${match.entity.original} -> ${match.entity.pseudonym}${isTreated ? ' (Tratado)' : ''}`}
              >
                {part}
                {isTreated && (
                  <span className="ml-1 inline text-[10px] text-emerald-600 font-bold select-none align-middle">✓</span>
                )}
              </span>
            );
          } else {
            return (
              <span 
                key={i} 
                className="px-1 py-[1.5px] mx-0.5 rounded border border-red-200 bg-red-50/60 text-red-700 font-medium text-xs cursor-help untreated-pii transition-all duration-300 hover:bg-red-50"
                title={`Potencial ${match.type} não identificado. Selecione para adicionar.`}
              >
                {part}
              </span>
            );
          }
        }
        return part;
      })}
    </>
  );
});

const DEFAULT_JUDGES = [
  "Maria dos Prazeres Couceiro Pizarro Beleza", "Maria Clara Pereira de Sousa de Santiago Sottomayor", "Mário Belo Morgado",
  "Helena Isabel Gonçalves Moniz Falcão de Oliveira", "Júlio Manuel Vieira Gomes", "Maria da Graça Machado Trigo Franco Frazão",
  "Maria de Fátima Morais Gomes", "Graça Maria Lima de Figueiredo Amaral", "Maria Olinda da Silva Nunes Garcia",
  "Catarina Isabel da Silva Santos Serra", "António José dos Santos Oliveira Abreu", "Maria João Romão Carreiro Vaz Tomé",
  "Nuno António Gonçalves", "Nuno Manuel Pinto Oliveira", "Ricardo Alberto Santos Costa", "Paulo Jorge Fonseca Ferreira da Cunha",
  "José Maria Ferreira Lopes", "José António Pires Teles Pereira", "João Eduardo Cura Mariano Esteves",
  "António Fernando Barateiro Dias Martins", "Fernando Baptista de Oliveira", "Luís Filipe Castelo Branco do Espírito Santo",
  "António Francisco Martins", "Ana Paula da Fonseca Lobo", "Isabel Maria Manso Salgado", "Jorge Manuel Leitão Leal",
  "Luís Miguel Ferreira de Azevedo Mendes", "José Eduardo Miranda Santos Sapateiro", "Fernando Vaz Ventura",
  "Emídio Francisco Santos", "Jorge Manuel Baptista Gonçalves", "Nelson Paulo Martins de Borges Carneiro",
  "Heitor Bernardo Cardoso Vasques Osório", "Celso José das Neves Manata", "Antero Luís", "Maria do Rosário Pita Pegado Gonçalves",
  "Henrique Ataíde Rosa Antunes", "Maria de Deus Simão da Cruz Silva Damasceno Correia", "António Augusto Manso",
  "José Alberto Vaz Carreto", "Anabela Figueiredo Luna de Carvalho", "Orlando dos Santos Nascimento",
  "Cristina Maria Nunes Soares Tavares Coelho", "Carlos Alberto Gameiro de Campos Lobo", "Rui Manuel Duarte Amorim Machado e Moura",
  "Jorge Manuel Ortins de Simões Raposo", "Maria Margarida Costa Pereira Ramos de Almeida", "Carlos Jorge Ferreira Portela",
  "Jorge Manuel de Miranda Natividade Jacob", "Arlindo Martins de Oliveira", "António Domingos Pires Robalo",
  "José Joaquim Aniceto Piedade", "Ernesto de Jesus de Deus Nascimento", "Maria da Graça Martins Pontes dos Santos Silva",
  "Antero Dinis Ramos Veiga", "Leopoldo Miguel Peres Mansinho Soares", "Adelina da Conceição Cardoso Barradas de Oliveira",
  "Maria Isoleta de Almeida Costa", "Maria Eduarda de Mira Branquinho Canas Mendes", "Pedro Álvaro de Sousa Donas Botto Fernando",
  "Maria das Dores Eiró de Araújo", "João Carlos Proença de Oliveira Costa", "Carlos António Paula Moreira",
  "Olga Maria dos Santos Maurício", "Isabel Maria Brás da Fonseca", "Ana Maria Martins Teixeira",
  "Maria Luísa de Meireles Carvalho Franco Duarte Ramos", "Eduardo Manuel Baptista Martins Rodrigues Pires", "Mário João Canelas Brás",
  "Maria Teresa de Sequeira Mendes Pardal", "Joaquim Arménio Correia Gomes", "João Manuel Monteiro Amaro",
  "Maria José da Costa Machado", "António Paulo Esteves Aguiar de Vasconcelos", "Maria Rosa Papança Barroso",
  "Joaquim José Felizardo Paiva", "João Manuel Moreira do Carmo", "Carlos Manuel Gonçalves de Melo Marinho",
  "Márcia Portela", "Manuel Pinto dos Santos", "Nuno Maria Rosa da Silva Garcia", "Anabela Moreira de Sá Cesariny Calafate",
  "Paula Cristina Passos Barradas Guerreiro", "José Eusébio dos Santos Soeiro de Almeida", "João Carlos da Silva Abrunhosa de Carvalho",
  "Pedro Maria Godinho Vaz Pato", "Maria Adelaide de Jesus Domingos", "Rui Manuel Barata Penha", "António José Alves Duarte",
  "Maria Catarina Ramalho Gonçalves", "António Júlio Costa Sobrinho", "Carlos Pereira Gil", "Paulo Alexandre Pereira Guerra",
  "Luís Filipe Brites Lameiras", "Maria José Pais de Sousa da Costa Pinto", "José Manuel da Silva Castela Rio",
  "António Manuel Mendes Coelho", "António Carlos Falcão de Beça Pereira", "José da Fonte Ramos", "Francisco José Rodrigues de Matos",
  "Maria João Fontinha Areias Cardoso", "João Manuel Araújo Ramos Lopes", "Manuela Bento Fialho", "Edgar Gouveia Valente",
  "Paulo Duarte Barreto Ferreira", "Filipe Manuel Nunes Caroço", "António Manuel Fernandes dos Santos",
  "Paulo Jorge Tavares Fernandes da Silva", "António José Moreira Ramos", "Alberto Augusto Vicente Ruço",
  "Pedro Maria Martin Martins", "Ana Paula Pereira de Amorim", "Maria Deolinda Gaudêncio Gomes Dionísio",
  "Maria Luísa Senra Arantes", "António José da Ascensão Ramos", "Judite Lima de Oliveira Pires", "José Manuel Igreja Martins Matos",
  "Nuno Miguel Pereira Ribeiro Coelho", "Aristides Manuel da Silva Rodrigues de Almeida", "Manuel António do Carmo Bargado",
  "Ana Isabel de Azeredo Rodrigues Coelho Fernandes da Silva", "Jorge Manuel da Silva Loureiro", "Edgar Taborda Lopes",
  "Albertina Maria Gomes Pedroso", "Rui Manuel Correia Moreira", "Jorge Manuel Langweg", "Maria Inês Carvalho Brasil de Moura",
  "Manuel Domingos Alves Fernandes", "Renato Amorim Damas Barroso", "José Vítor dos Santos Amaral",
  "Miguel Fernando Baldaia Correia de Morais", "Luís Filipe Dias Cravo", "Paulo Eduardo Cristão Correia",
  "Alcina Maria Cleto Duarte da Costa Ribeiro", "Alda Maria de Oliveira Martins", "Sérgio Manuel da Silva de Almeida",
  "Jorge Miguel Pinto de Seabra", "Maria Amália Pereira dos Santos", "Ana Cristina Aparício de Oliveira Duarte",
  "Francisco João Machado da Cunha Xavier", "Francisca da Mata Mendes", "Luís Antunes Coimbra", "João Diogo de Frias Rodrigues",
  "Maria José Monteiro Guerra", "Anabela Andrade Miranda", "Francisca Micaela Fonseca da Mota Vieira",
  "Maria Dolores da Silva e Sousa", "Luís Filipe Pires de Sousa", "Carla Inês Brás Câmara", "José Manuel Costa Galo Tomé de Carvalho",
  "Manuel Henrique Ramos Soares", "Maria João Vasques de Sousa e Faro", "Helena Isabel Ribeiro Carmelo Dias Bolieiro",
  "Paulo Fernando Dias da Silva", "Maria José de Almeida Costeira", "Nelson Nunes Fernandes", "João Pedro Nunes Maldonado",
  "Manuel Alexandre Teixeira Advínculo Sequeira", "José Francisco Mota Ribeiro", "Helena Maria de Carvalho Gomes de Melo",
  "Mário Jorge dos Santos Branco Coelho", "Isabel Maria Socorro de Matos Peixoto Imaginário", "Higina Maria Almeida Orvalho da Silva Castelo",
  "João António Peres de Oliveira Coelho", "Vítor Manuel Leitão Ribeiro", "Maria Isabel Sousa Ribeiro Silva",
  "Raúl Eduardo Nunes Esteves", "Fernando Manuel Matos de Azevedo Correia Chaves", "José Júlio da Cunha Amorim Pinto",
  "Jorge Manuel Duarte Bispo", "Pedro Miguel Bengala Reis da Cunha Lopes", "Amélia Maria dos Reis Catarino Correia de Almeida",
  "Maria de Fátima Cardoso Bernardes", "Maria Fernanda Lopes Ventura", "Lina Aurora Ramada e Castro Bettencourt Baptista",
  "Pedro Alexandre Damião e Cunha", "Maria de Fátima Almeida Andrade", "Alexandra Maria Rolim Mendes",
  "Ausenda Gonçalves e Alexandre dos Reis", "Maria da Purificação Lopes de Carvalho", "Maria de Fátima Cerveira da Cunha Lopes Furtado",
  "Elsa de Jesus Coelho Paixão", "Maria dos Prazeres Rodrigues da Silva", "Vera Maria Guedes Barbosa de Sottomayor Bismark do Agro",
  "Maria João Marques Pinto de Matos", "Rita Maria Pereira Romeira", "Elisabete de Jesus Ribeiro Assunção", "Álvaro Monteiro",
  "Cláudia Sofia de Jesus Antunes Barata", "José Nuno Ramos Duarte", "Pedro José Esteves de Brito", "Isabel Maria Trocado Monteiro",
  "António José Barrocal Fialho", "Filipe João Aveiro de Sousa Marques", "Augusta Maria Pinto Ferreira Rodrigues Palma",
  "Teresa Manuela Pinto da Silva", "Nuno Marcelo de Nóbrega dos Santos de Freitas Araújo", "Maria de Fátima Silva Viegas",
  "Sónia Maria Fontes de Magalhães de Oliveira Pereira", "Ana Cristina de Jesus Batalha Cardoso", "Carla de Jesus da Costa Fraga Torres",
  "Fernando Alberto Caetano Besteiro", "Ana Cristina Rodrigues Clemente", "Susana Cristina Mendes Santos Martins da Silveira",
  "João Simões Presa Grilo de Amaral", "Alexandra Maria Bandeira Ferraz Lage", "João Filipe Pereira Bártolo",
  "Cristina Isabel Elias Henriques Esteves", "Ana Rute Alves da Costa Pereira", "Ana Lúcia dos Reis Gordinho",
  "Susana Pinto Santos Silva", "Ana Rita Varela Loja", "Filipe Amadeu César Osório Rodrigues Costa", "Sónia Alexandra Sousa de Moura",
  "Susana Isabel Santos Pinto de Oliveira Ferrão da Costa Cabral", "Ricardo Manuel Neto Miranda Peixoto", "Rui Miguel Pereira Poças",
  "Eduardo José Capela de Sousa Paiva", "Maria de Fátima da Rocha Marques Bessa", "Gabriela Lopes Feiteira",
  "Manuela Maria Marques Trocado", "Fernando Miguel Furtado André Alves", "Ana Paula Soares Ferreira Guedes",
  "Sara da Piedade Moreira das Neves de Pina Cabral", "Filipe Duarte Freitas Câmara", "Rosa Maria Cardoso Saraiva",
  "Diogo Coelho de Sousa Leitão", "Rosa dos Remédios Lima Teixeira", "Carlos Alberto Casas Azevedo", "Marlene Fortuna Rodrigues",
  "Pedro Miguel dos Reis Raposo de Figueiredo", "Estrela Aramita Dias Chambel Capelo de Sousa Chaby Rosa", "Maria Emília Guerreiro de Avillez Melo e Castro"
];

const DEFAULT_AUTHORS = [
  "Abrantes Geraldes", "Adriano Vaz Serra", "Alberto dos Reis", "Alberto Xavier", "Albertina Pedroso",
  "Almeida Costa", "Alexandre de Soveral Martins", "Alexandra Leitão", "Alexandra Rodrigues",
  "Ana Cristina Rangel", "Ana Neves", "Ana Paula Dourado", "Ana Rita Gil", "Ana Sofia Fonseca",
  "Anabela Miranda Rodrigues", "André Gonçalves Pereira", "André Lamas Leite", "Anselmo de Castro",
  "António Caeiro", "António Carlos dos Santos", "António Cândido de Oliveira", "António de Oliveira Ascensão",
  "António Henriques Gaspar", "António Joaquim Lobo Xavier", "António Menezes Cordeiro",
  "António Monteiro Fernandes", "António Nunes de Carvalho", "António Pinto Monteiro", "António Santos Justo",
  "António Vitorino", "Armindo Ribeiro Mendes", "Avelãs Nunes", "Benjamim Silva Rodrigues", "Brandão Proença",
  "Carlos Adérito Teixeira", "Carlos Alegre", "Carlos Alberto da Mota Pinto", "Carlos Blanco de Morais",
  "Carlos Ferreira de Almeida", "Carlos Lobo", "Carlos Lopes do Rego", "Carlos Trindade", "Catarina Frade",
  "Catarina Sarmento e Castro", "Cláudia Cruz Santos", "Cláudia Santos", "Cláudia Viana", "Cláudio de Oliveira",
  "Cristina Líbano Monteiro", "Cristina Queiroz", "Daniela Paiano", "David Duarte", "Diogo Freitas do Amaral",
  "Diogo Leite de Campos", "Dulce Lopes", "Dulce M. da Costa", "Eduardo Correia", "Eduardo Paz Ferreira",
  "Edgar Valles", "Eurico Lopes-Cardoso", "Fausto de Quadros", "Fernando Amâncio Ferreira", "Fernando Araújo",
  "Fernando Tomé", "Ferrer Correia", "Figueiredo Dias", "Francisco Amaral", "Francisco Lucas Pires",
  "Francisco Pereira Coutinho", "Francisco Rodrigues Pardal", "Francisco António de M. L. Ferreira de Almeida",
  "Frederico da Costa Pinto", "Freitas do Amaral", "Germano Marques da Silva", "Gonçalo S. de Melo Bandeira",
  "Gomes Canotilho", "Guilherme de Oliveira", "Guilherme Dray", "Guilherme Moreira", "Helena Moniz",
  "Helena Morão", "Helena Morais", "Helena Tomás", "Helena Mota", "Henrique Mesquita", "Isabel Alexandre",
  "Isabel Celeste Fonseca", "Isabel de Magalhães Colaço", "Isabel Ribeiro", "Isabel Alexandra Ponce de Leão",
  "João Calvão da Silva", "João Caupers", "João Conde Correia", "João de Matos Antunes Varela",
  "João Domingos Silva", "João Labareda", "João Leal Amado", "João Rato", "João Rego", "João Salazar",
  "João Tiago Silveira", "João Vaz Rodrigues", "João Zenha Martins", "Jorge Bacelar Gouveia", "Jorge de Brito",
  "Jorge de Figueiredo Dias", "Jorge de Seabra", "Jorge Duarte Pinheiro", "Jorge Leite", "Jorge Lobo Xavier",
  "Jorge Medeiros", "Jorge Miranda", "Jorge Reis Novais", "José Alberto dos Reis", "José António Barreiros",
  "José Carlos Vieira de Andrade", "José Casalta Nabais", "José Cândido de Pinho", "José Damião da Cunha",
  "José de Faria Costa", "José de Oliveira Ascensão", "José Eduardo Figueiredo Dias", "José Engrácia Antunes",
  "José Ferreira Gomes", "José Ferreira Oliveira", "José Gameiro", "José Guilherme Xavier de Basto",
  "José Joaquim de Sousa", "José José Joaquim Gomes Canotilho", "José Lebre de Freitas", "José Luís da Cruz Vilaça",
  "José Manuel Damião da Cunha", "José Manuel de Oliveira", "José Manuel Durão Barroso", "José Manuel Pureza",
  "José Manuel Sérvulo Correia", "José Maria Fernandes Pires", "José Souto de Moura", "José Tavares de Sousa",
  "Joaquim Freitas da Rocha", "Joaquim Sousa Ribeiro", "Jónatas Machado", "Júlio Gomes", "Laurinda Gemas",
  "Lebre de Freitas", "Lino Torgal", "Luís A. Carvalho Fernandes", "Luís Cabral de Moncada", "Luís Filipe Pires de Sousa",
  "Luís Gonçalves da Silva", "Luís Greco", "Luís Menezes Leitão", "Luís S. Cabral de Moncada", "Marcelo Rebelo de Sousa",
  "M. Maia Gonçalves", "Manuel A. Domingues de Andrade", "Manuel de Andrade", "Manuel da Costa Andrade",
  "Manuel Simas Santos", "Margarida Reis", "Margarida Silva Pereira", "Maria Benedita Urbano", "Maria Clara Sottomayor",
  "Maria de Fátima Ribeiro", "Maria do Carmo da Costa", "Maria do Carmo Teles", "Maria do Rosário Palma Ramalho",
  "Maria João Antunes", "Maria João Mimoso", "Maria João Vaz Tomé", "Maria Lúcia Amaral", "Maria Luísa Duarte",
  "Maria Luísa Portela de Sousa", "Maria Teresa de Melo Ribeiro", "Mariana Canotilho", "Mário Esteves de Oliveira",
  "Mário Ferreira Monte", "Mário Júlio de Almeida Costa", "Mário Soares", "Mário Tenreiro", "Miguel Gorjão-Henriques",
  "Miguel Lucas Pires", "Miguel Mesquita", "Miguel Nogueira de Brito", "Miguel Teixeira de Sousa",
  "Mónica G. N. Ferreira", "Nuno Brandão", "Nuno Cordeiro", "Nuno Cremona", "Nuno de Salter Cid",
  "Nuno Ferreira da Cunha", "Nuno Piçarra", "Nuno Sá Gomes", "Olga de Sousa", "Paula Costa e Silva",
  "Paula Quintas", "Paula Rosado Pereira", "Paulo Câmara", "Paulo da Mota Pinto", "Paulo de Sousa Mendes",
  "Paulo de Pitta e Cunha", "Paulo Morgado de Carvalho", "Paulo Mota Pinto", "Paulo Otero", "Paulo Olavo Cunha",
  "Paulo Pimenta", "Paulo Pinto de Albuquerque", "Paulo Pulido Adragão", "Paulo Saragoça da Matta",
  "Pedro Bacelar de Vasconcelos", "Pedro Caeiro", "Pedro Costa Gonçalves", "Pedro Furtado Martins",
  "Pedro Gonçalves", "Pedro Maia", "Pedro Madeira de Brito", "Pedro Martínez", "Pedro Pais de Vasconcelos",
  "Pedro Romano Martinez", "Rita Garcia Pereira", "Rita Lobo Xavier", "Rita Nóbrega", "Ricardo Jorge Bragança de Matos",
  "Rui Assis", "Rui de Alarcão", "Rui Medeiros", "Rui Morais", "Rui Pereira", "Rui Pinto", "Rui Pinto Duarte",
  "Rui M. de Medeiros", "Salvador da Costa", "Saldanha Sanches", "Sandra Barreira", "Sérgio Poças",
  "Sérvulo Correia", "Sinde Monteiro", "Susana Aires de Sousa", "Susana Tavares da Silva", "Suzana Tavares da Silva",
  "Teles de Menezes Leitão", "Teresa Arruda Alvim", "Teresa Coelho Moreira", "Teresa Pizarro Beleza",
  "Teresa Violante", "Tiago Caiado Milheiro", "Tiago Duarte", "Tiago Serrão", "Varela de Matos",
  "Vasco Costa", "Vasco Pereira da Silva", "Vieira de Andrade", "Vítor Gomes", "Vítor Ferreira", "Vital Moreira"
];

export default function App() {
  const [files, setFiles] = useState<FileData[]>([]);
  const [entities, setEntities] = useState<PIIEntity[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const [filterType, setFilterType] = useState<string | 'ALL'>('ALL');
  const [hideIgnored, setHideIgnored] = useState(false);
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());
  const [editingEntity, setEditingEntity] = useState<PIIEntity | null>(null);
  const [splitResults, setSplitResults] = useState<{ original: string, type: string, id: string }[] | null>(null);
  const [pendingManualTerm, setPendingManualTerm] = useState<{ text: string, x: number, y: number } | null>(null);
  const [copiedPseudonym, setCopiedPseudonym] = useState<string | null>(null);
  const [showMergeModal, setShowMergeModal] = useState(false);
  const [mergeSearch, setMergeSearch] = useState('');
  const [showAllMergeOptions, setShowAllMergeOptions] = useState(false);

  const mergeOptions = useMemo(() => {
    if (!editingEntity) return [];
    
    const baseEntities = entities.filter(e => !selectedIds.has(e.id));
    
    // Group by groupId or id
    const groups = baseEntities.reduce((acc, e) => {
      const key = e.groupId || e.id;
      if (!acc.find(item => (item.groupId || item.id) === key)) {
        acc.push(e);
      }
      return acc;
    }, [] as PIIEntity[]);

    if (showAllMergeOptions) return groups;

    // Filter by search or by similarity to editingEntity
    const searchLower = mergeSearch.toLowerCase().trim();
    const entityLower = editingEntity.original.toLowerCase().trim();
    
    return groups.filter(g => {
      const gLower = g.original.toLowerCase();
      if (searchLower) return gLower.includes(searchLower) || g.pseudonym.toLowerCase().includes(searchLower);
      
      // Default similarity: contains the text or shares words
      return gLower.includes(entityLower) || entityLower.includes(gLower);
    });
  }, [entities, selectedIds, editingEntity, mergeSearch, showAllMergeOptions]);
  const [showExceptionsModal, setShowExceptionsModal] = useState(false);
  const [showManualModal, setShowManualModal] = useState(false);
  const [showAmbiguityModal, setShowAmbiguityModal] = useState(false);
  const [ambiguousEntities, setAmbiguousEntities] = useState<PIIEntity[]>([]);
  const [exceptionsTab, setExceptionsTab] = useState<'EXCECAO' | 'JUIZ' | 'AUTOR' | 'COLETIVA' | 'SAFELIST' | 'HEADER'>('EXCECAO');
  const [knowledgeSearch, setKnowledgeSearch] = useState('');
  const [toast, setToast] = useState<{ message: string; type: 'success' | 'error' | 'info' } | null>(null);
  const [splitView, setSplitView] = useState(false);
  const [reviewMode, setReviewMode] = useState(false);
  const [selectedEntityId, setSelectedEntityId] = useState<string | null>(null);
  const [selectedFileId, setSelectedFileId] = useState<string | null>(null);
  const [headerZones, setHeaderZones] = useState<{ top: number; bottom: number; left: number; right: number; image?: string }[]>([]);
  const [footerZones, setFooterZones] = useState<{ top: number; bottom: number; left: number; right: number; image?: string }[]>([]);
  const [selectionMode, setSelectionMode] = useState<'none' | 'header' | 'footer'>('none');
  const [confirmingClear, setConfirmingClear] = useState(false);
  const [history, setHistory] = useState<{ entities: PIIEntity[], files: FileData[], ambiguousEntities: PIIEntity[] }[]>([]);
  const [historyIndex, setHistoryIndex] = useState(-1);
  const isHistoryAction = useRef(false);

  const leftScrollRef = useRef<HTMLDivElement>(null);
  const rightScrollRef = useRef<HTMLDivElement>(null);

  const handleSyncScroll = (source: 'left' | 'right') => {
    const sourceEl = source === 'left' ? leftScrollRef.current : rightScrollRef.current;
    const targetEl = source === 'left' ? rightScrollRef.current : leftScrollRef.current;
    
    if (sourceEl && targetEl) {
      const sourceMaxScroll = sourceEl.scrollHeight - sourceEl.clientHeight;
      const targetMaxScroll = targetEl.scrollHeight - targetEl.clientHeight;
      
      if (sourceMaxScroll > 0 && targetMaxScroll > 0) {
        const scrollPercentage = sourceEl.scrollTop / sourceMaxScroll;
        const targetScrollTop = scrollPercentage * targetMaxScroll;
        
        if (Math.abs(targetEl.scrollTop - targetScrollTop) > 2) {
          targetEl.scrollTop = targetScrollTop;
        }
      }
    }
  };

  // Scroll to selected entity in Split View
  useEffect(() => {
    if (!selectedEntityId || !splitView) return;
    
    const scrollToActive = () => {
      const leftEl = leftScrollRef.current?.querySelector('#active-highlight');
      const rightEl = rightScrollRef.current?.querySelector('#active-highlight');
      
      if (leftEl) {
        leftEl.scrollIntoView({ behavior: 'smooth', block: 'center' });
      }
      if (rightEl) {
        rightEl.scrollIntoView({ behavior: 'smooth', block: 'center' });
      }
    };

    // Small delay to ensure HighlightText has rendered
    const timer = setTimeout(scrollToActive, 300);
    return () => clearTimeout(timer);
  }, [selectedEntityId, splitView]);

  const navigateUntreated = (direction: 'next' | 'prev') => {
    const container = leftScrollRef.current;
    if (!container) return;

    const elements = Array.from(container.querySelectorAll('.untreated-pii'));
    if (elements.length === 0) {
      showToast("Não foram encontrados mais elementos potenciais.", "info");
      return;
    }

    // Find current element in view
    const viewTop = container.scrollTop;
    let targetIndex = -1;

    if (direction === 'next') {
      targetIndex = elements.findIndex(el => (el as HTMLElement).offsetTop > viewTop + 100);
      if (targetIndex === -1) targetIndex = 0; // Loop back
    } else {
      for (let i = elements.length - 1; i >= 0; i--) {
        if ((elements[i] as HTMLElement).offsetTop < viewTop - 10) {
          targetIndex = i;
          break;
        }
      }
      if (targetIndex === -1) targetIndex = elements.length - 1; // Loop back
    }

    const targetEl = elements[targetIndex] as HTMLElement;
    container.scrollTo({
      top: targetEl.offsetTop - 150,
      behavior: 'smooth'
    });
    
    // Highlight briefly
    targetEl.classList.add('ring-4', 'ring-red-400', 'ring-opacity-50');
    setTimeout(() => {
      targetEl.classList.remove('ring-4', 'ring-red-400', 'ring-opacity-50');
    }, 1500);
  };

  // Auto-hide toast
  useEffect(() => {
    if (toast) {
      const timer = setTimeout(() => setToast(null), 3000);
      return () => clearTimeout(timer);
    }
  }, [toast]);

  useEffect(() => {
    if (reviewMode && !selectedFileId && files.length > 0) {
      setSelectedFileId(files[0].id);
    }
    if (!reviewMode) {
      setSelectedEntityId(null);
    }
  }, [reviewMode, selectedFileId, files]);

  const showToast = (message: string, type: 'success' | 'error' | 'info' = 'info') => {
    setToast({ message, type });
  };

  const pushToHistory = useCallback((newEntities: PIIEntity[], newFiles: FileData[], newAmbiguous: PIIEntity[] = ambiguousEntities) => {
    if (isHistoryAction.current) {
      isHistoryAction.current = false;
      return;
    }
    const newState = { 
      entities: JSON.parse(JSON.stringify(newEntities)), 
      files: newFiles.map(f => ({ ...f })),
      ambiguousEntities: JSON.parse(JSON.stringify(newAmbiguous))
    };
    setHistory(prev => {
      const next = prev.slice(0, historyIndex + 1);
      next.push(newState);
      if (next.length > 30) next.shift();
      return next;
    });
    setHistoryIndex(prev => {
      const next = prev + 1;
      return next > 29 ? 29 : next;
    });
  }, [historyIndex, ambiguousEntities]);

  const undo = () => {
    if (historyIndex > 0) {
      isHistoryAction.current = true;
      const prevState = history[historyIndex - 1];
      setEntities(prevState.entities);
      setFiles(prevState.files);
      setAmbiguousEntities(prevState.ambiguousEntities);
      setHistoryIndex(historyIndex - 1);
      showToast("Ação anulada", "info");
    }
  };

  const redo = () => {
    if (historyIndex < history.length - 1) {
      isHistoryAction.current = true;
      const nextState = history[historyIndex + 1];
      setEntities(nextState.entities);
      setFiles(nextState.files);
      setAmbiguousEntities(nextState.ambiguousEntities);
      setHistoryIndex(historyIndex + 1);
      showToast("Ação refeita", "info");
    }
  };

  const handleExportReport = () => {
    if (entities.length === 0) return;
    
    const reportData = entities
      .filter(e => !e.ignored)
      .map(e => ({
        'Original': e.original,
        'Pseudónimo': e.pseudonym,
        'Tipo': e.type,
        'Estado': e.enabled ? 'Ativo' : 'Inativo',
        'Ficheiros': e.fileIds?.map(id => files.find(f => f.id === id)?.name).join(', ') || ''
      }));

    const ws = XLSX.utils.json_to_sheet(reportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Correspondências");
    
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    saveAs(new Blob([wbout], { type: "application/octet-stream" }), `relatorio_correspondencias_${new Date().toISOString().slice(0,10)}.xlsx`);
    showToast("Relatório exportado com sucesso", "success");
  };

  const [isRelated, setIsRelated] = useState(true);
  const [collapsedGroups, setCollapsedGroups] = useState<Set<string>>(new Set());
  const [globalKnowledge, setGlobalKnowledge] = useState<Record<string, string>>(() => {
    const saved = localStorage.getItem('pii_global_knowledge');
    return saved ? JSON.parse(saved) : {};
  });

  const [safelist, setSafelist] = useState<Safelist>(() => {
    const saved = localStorage.getItem('pii_safelist');
    return saved ? JSON.parse(saved) : { words_ignore: [], phrases_ignore: [] };
  });

  const [editingKnowledge, setEditingKnowledge] = useState<{
    original: string;
    current: string;
    type: 'KNOWLEDGE' | 'SAFELIST_WORD' | 'SAFELIST_PHRASE';
    category?: string;
  } | null>(null);

  const [deduplicationSuggestions, setDeduplicationSuggestions] = useState<Array<{
    item1: string;
    item2: string;
    score: number;
    type: string;
  }>>([]);
  const [showDeduplicationModal, setShowDeduplicationModal] = useState(false);

  // Migration script for old global exceptions
  useEffect(() => {
    const oldExceptions = localStorage.getItem('pii_global_exceptions');
    if (oldExceptions) {
      try {
        const parsed: string[] = JSON.parse(oldExceptions);
        if (Array.isArray(parsed) && parsed.length > 0) {
          setGlobalKnowledge(prev => {
            const next = { ...prev };
            parsed.forEach(ex => {
              const lower = ex.toLowerCase().trim();
              if (!next[lower]) {
                next[lower] = 'EXCECAO';
              }
            });
            return next;
          });
          // Remove old key after successful migration
          localStorage.removeItem('pii_global_exceptions');
        }
      } catch (err) {
        console.error("Migration error:", err);
      }
    }
  }, []);

  useEffect(() => {
    localStorage.setItem('pii_global_knowledge', JSON.stringify(globalKnowledge));
  }, [globalKnowledge]);

  useEffect(() => {
    localStorage.setItem('pii_safelist', JSON.stringify(safelist));
  }, [safelist]);

  // Auto-save project state
  useEffect(() => {
    const saved = localStorage.getItem('pii_project_state');
    if (saved) {
      try {
        const { files: savedFiles, entities: savedEntities } = JSON.parse(saved);
        if (savedFiles && savedEntities) {
          // Deduplicate entities by ID to prevent key collision errors
          const uniqueEntities = deduplicateEntities(savedEntities as PIIEntity[]);

          setFiles(savedFiles);
          setEntities(uniqueEntities);
        }
      } catch (e) {
        console.error("Error loading saved state", e);
      }
    }
  }, []);

  useEffect(() => {
    if (files.length > 0 || entities.length > 0) {
      localStorage.setItem('pii_project_state', JSON.stringify({ 
        files: files.map(({ rawFile, ...rest }) => rest), 
        entities 
      }));
    }
  }, [files, entities]);

  // Initial history state
  useEffect(() => {
    if (history.length === 0 && (files.length > 0 || entities.length > 0)) {
      setHistory([{ entities: JSON.parse(JSON.stringify(entities)), files: files.map(f => ({ ...f })) }]);
      setHistoryIndex(0);
    }
  }, [files, entities, history.length]);

  // --- File Handling ---

  const handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const uploadedFiles = event.target.files;
    if (!uploadedFiles) return;

    const newFiles: FileData[] = Array.from(uploadedFiles).map(file => ({
      id: generateId(),
      name: (file as File).name,
      type: (file as File).type,
      content: '',
      rawFile: file as File,
      status: 'pending'
    }));

    setFiles(prev => [...prev, ...newFiles]);
    processFiles(newFiles);
  };

  const processFiles = async (filesToProcess: FileData[]) => {
    setIsProcessing(true);
    let allNewEntities = [...entities];

    for (const fileData of filesToProcess) {
      // Yield to UI thread
      await new Promise(resolve => setTimeout(resolve, 50));
      
      try {
        setFiles(prev => prev.map(f => f.id === fileData.id ? { ...f, status: 'processing' } : f));
        
        let text = '';
        let htmlContent = '';
        let positions: { start: number, end: number, page: number, y: number }[] = [];
        if (fileData.type === 'application/pdf') {
          const pdfData = await extractTextAndHtmlFromPDF(fileData.rawFile);
          text = pdfData.text;
          htmlContent = pdfData.html;
          positions = pdfData.positions;
        } else if (fileData.type === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' || fileData.name.endsWith('.docx')) {
          const docxData = await extractTextAndHtmlFromDocx(fileData.rawFile);
          text = docxData.text;
          htmlContent = docxData.html;
        } else if (fileData.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || fileData.name.endsWith('.xlsx')) {
          text = await extractTextFromXlsx(fileData.rawFile);
        } else {
          // Fallback for other types or simple text
          text = await fileData.rawFile.text();
        }

        const fileEntities = scanText(text, fileData.id, allNewEntities, isRelated, globalKnowledge, safelist, positions);
        
        // Merge entities correctly to avoid duplicates
        const newOnes = fileEntities.filter(fe => !allNewEntities.some(ae => ae.id === fe.id));
        const updatedOnes = fileEntities.filter(fe => allNewEntities.some(ae => ae.id === fe.id));
        
        allNewEntities = allNewEntities.map(ae => {
          const updated = updatedOnes.find(uo => uo.id === ae.id);
          return updated || ae;
        });
        allNewEntities = deduplicateEntities([...allNewEntities, ...newOnes]);

        fileData.content = text;
        fileData.htmlContent = htmlContent;
        fileData.positions = positions;
        setFiles(prev => prev.map(f => f.id === fileData.id ? { ...f, content: text, htmlContent, positions, status: 'done' } : f));
      } catch (error) {
        console.error(`Error processing file ${fileData.name}:`, error);
        setFiles(prev => prev.map(f => f.id === fileData.id ? { ...f, status: 'error' } : f));
      }
    }

    const potentialAmbiguities = allNewEntities.filter(e => 
      (e.type === 'NOME' || e.type === 'AUTOR' || e.type === 'JUIZ') && 
      (e.original.includes(' e ') || e.original.includes(',') || e.original.includes('  '))
    );

    const updatedFiles = files.map(f => {
      const processed = filesToProcess.find(p => p.id === f.id);
      if (processed) return { 
        ...f, 
        content: (processed as any).content, 
        htmlContent: (processed as any).htmlContent,
        status: 'done' 
      };
      return f;
    });

    if (potentialAmbiguities.length > 0) {
      setAmbiguousEntities(allNewEntities);
      setShowAmbiguityModal(true);
      setIsProcessing(false);
      pushToHistory(entities, updatedFiles, allNewEntities);
      return;
    }

    const grouped = groupSimilarEntities(allNewEntities, isRelated);
    setEntities(grouped);
    pushToHistory(grouped, updatedFiles);
    setIsProcessing(false);
  };

  const extractTextAndHtmlFromDocx = async (file: File): Promise<{ text: string, html: string }> => {
    try {
      const arrayBuffer = await file.arrayBuffer();
      
      // Get HTML for formatting preservation with custom style map
      const options = {
        styleMap: [
          "u => u",
          "b => strong",
          "i => em",
          "strike => s",
          "p[style-name='Footnote Text'] => p.footnote-text",
          "r[style-name='Footnote Reference'] => span.footnote-reference"
        ]
      };
      
      const htmlResult = await mammoth.convertToHtml({ arrayBuffer }, options);
      const html = htmlResult.value;
      
      // Get raw text for scanning
      const textResult = await mammoth.extractRawText({ arrayBuffer });
      const text = textResult.value;
      
      return { text, html };
    } catch (error) {
      console.error('Error extracting from DOCX:', error);
      const text = await file.text();
      return { text, html: text };
    }
  };

  const extractTextFromXlsx = async (file: File): Promise<string> => {
    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    let fullText = '';
    workbook.SheetNames.forEach(sheetName => {
      const worksheet = workbook.Sheets[sheetName];
      fullText += XLSX.utils.sheet_to_txt(worksheet) + '\n\n';
    });
    return fullText;
  };

  const extractTextAndHtmlFromPDF = async (file: File): Promise<{ text: string, html: string, positions: { start: number, end: number, page: number, y: number }[] }> => {
    try {
      const arrayBuffer = await file.arrayBuffer();
      const pdf = await pdfjs.getDocument({ data: arrayBuffer }).promise;
      let fullText = '';
      let fullHtml = '';
      const positions: { start: number, end: number, page: number, y: number }[] = [];

      for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const textContent = await page.getTextContent();
        const styles = textContent.styles || {};
        const viewport = page.getViewport({ scale: 1.0 });
        const pageHeight = viewport.height;
        
        const items = (textContent.items as any[])
          .filter(item => typeof item.str === 'string')
          .sort((a, b) => {
            const yDiff = b.transform[5] - a.transform[5];
            if (Math.abs(yDiff) < 2) {
              return a.transform[4] - b.transform[4];
            }
            return yDiff;
          });

        if (items.length === 0) continue;

        // 1. Group items into lines
        const lines: { items: any[], y: number, fontSize: number, xStart: number, xEnd: number }[] = [];
        let currentLine: any[] = [];
        let lastY = -1;

        items.forEach(item => {
          const y = item.transform[5];
          if (lastY === -1 || Math.abs(y - lastY) < 2) {
            currentLine.push(item);
          } else {
            const fontSize = Math.abs(currentLine[0].transform[3]);
            const xStart = Math.min(...currentLine.map(it => it.transform[4]));
            const xEnd = Math.max(...currentLine.map(it => it.transform[4] + (it.width || 0)));
            lines.push({ items: currentLine, y: lastY, fontSize, xStart, xEnd });
            currentLine = [item];
          }
          lastY = y;
        });
        if (currentLine.length > 0) {
          const fontSize = Math.abs(currentLine[0].transform[3]);
          const xStart = Math.min(...currentLine.map(it => it.transform[4]));
          const xEnd = Math.max(...currentLine.map(it => it.transform[4] + (it.width || 0)));
          lines.push({ items: currentLine, y: lastY, fontSize, xStart, xEnd });
        }

        // 2. Identify Margins with safety
        const xStarts = lines.filter(l => l.items[0].str.trim().length > 3).map(l => l.xStart);
        const xEnds = lines.filter(l => l.items[0].str.trim().length > 3).map(l => l.xEnd);
        const leftMargin = xStarts.length > 0 ? Math.min(...xStarts) : 50;
        const rightMargin = xEnds.length > 0 ? Math.max(...xEnds) : 550;
        const lineFullWidth = Math.max(100, rightMargin - leftMargin);

        // 3. Group lines into paragraphs using Visual Logic
        const avgFontSize = lines.reduce((acc, l) => acc + l.fontSize, 0) / lines.length;
        const paragraphs: { lines: any[], isFootnote: boolean, y: number }[] = [];
        let currentParagraph: { lines: any[], isFootnote: boolean, y: number } | null = null;

        lines.forEach((line, idx) => {
          // Footnote detection: bottom 28% of page and smaller font or starts with digit
          const isBottom = (pageHeight - line.y) < pageHeight * 0.28;
          const isSmall = line.fontSize < avgFontSize * 0.9;
          const startsWithDigit = /^\d+[\s\.]/.test(line.items[0].str.trim());
          const looksLikeFootnote = (isSmall && isBottom) || (startsWithDigit && isBottom);

          const prevLine = idx > 0 ? lines[idx - 1] : null;
          const yGap = prevLine ? Math.abs(line.y - prevLine.y) : 0;
          
          // Visual Paragraph Logic:
          const isIndented = line.xStart > leftMargin + (lineFullWidth * 0.08);
          const prevWasShort = prevLine && prevLine.xEnd < rightMargin - (lineFullWidth * 0.15);
          
          const isNewPara = !prevLine || 
                           yGap > line.fontSize * 1.7 || 
                           isIndented || 
                           prevWasShort ||
                           Math.abs(line.fontSize - prevLine.fontSize) > 1.2 ||
                           looksLikeFootnote !== currentParagraph?.isFootnote;

          if (isNewPara) {
            currentParagraph = { lines: [line], isFootnote: looksLikeFootnote, y: line.y };
            paragraphs.push(currentParagraph);
          } else {
            currentParagraph?.lines.push(line);
          }
        });

        // 4. Generate HTML from paragraphs and track positions
        paragraphs.forEach(p => {
          const pClass = p.isFootnote ? 'pdf-footnote' : '';
          const pFontSize = p.lines[0].fontSize;
          const pFontName = p.lines[0].items[0].fontName;
          const pFontFamily = styles[pFontName]?.fontFamily || 'Times New Roman';
          
          const normalizedY = (pageHeight - p.y) / pageHeight;
          
          // Align detection
          let pAlign = 'justify';
          if (p.lines.length > 0) {
            const firstLine = p.lines[0];
            const firstLineCenter = (firstLine.xStart + firstLine.xEnd) / 2;
            const pageCenter = (leftMargin + rightMargin) / 2;
            const width = firstLine.xEnd - firstLine.xStart;
            const pageWidthEst = rightMargin - leftMargin;

            if (Math.abs(firstLineCenter - pageCenter) < 25 && width < pageWidthEst * 0.75) {
              pAlign = 'center';
            } else if (firstLine.xStart > leftMargin + (pageWidthEst * 0.25) && Math.abs(firstLine.xEnd - rightMargin) < 30) {
              pAlign = 'right';
            } else if (firstLine.xStart > leftMargin + (pageWidthEst * 0.15)) {
              pAlign = 'center';
            }
          }
          
          let pHtml = `<p class="${pClass}" data-y="${normalizedY}" style="text-align: ${pAlign}; margin-bottom: 10pt; font-family: '${pFontFamily}'; font-size: ${pFontSize}pt;">`;
          let pText = '';

          p.lines.forEach((line, lIdx) => {
            line.items.forEach((item: any, iIdx: number) => {
              const style = styles[item.fontName];
              const fontName = (item.fontName || '').toLowerCase();
              const fontFamily = (style?.fontFamily || '').toLowerCase();
              
              // Ultra-aggressive style detection
              const isBold = fontFamily.includes('bold') || fontName.includes('bold') || 
                             fontName.includes('bd') || fontName.includes('-b') || 
                             fontName.includes('_b') || fontName.includes('black') || 
                             fontName.includes('heavy') || fontName.includes('w7') || 
                             fontName.includes('w8') || fontName.includes('w9') || 
                             fontName.includes('semibold') || fontName.includes('medium') || 
                             fontName.includes('700') || fontName.includes('800') || fontName.includes('900');
              
              const isItalic = fontFamily.includes('italic') || fontFamily.includes('oblique') || 
                               fontName.includes('italic') || fontName.includes('oblique') || 
                               fontName.includes('it') || fontName.includes('-i') || 
                               fontName.includes('_i') || fontName.includes('slanted');
              
              const isUnderline = fontName.includes('underline') || fontName.includes('ul') || fontFamily.includes('underline');

              let itemText = item.str;
              pText += itemText;
              
              let styledText = itemText;
              if (isBold) styledText = `<strong>${styledText}</strong>`;
              if (isItalic) styledText = `<em>${styledText}</em>`;
              if (isUnderline) styledText = `<u>${styledText}</u>`;
              
              pHtml += styledText;
              if (iIdx < line.items.length - 1) {
                pHtml += ' ';
                pText += ' ';
              }
            });
            if (lIdx < p.lines.length - 1) {
              pHtml += ' '; 
              pText += ' ';
            }
          });

          pHtml += '</p>';
          
          // Track position for this paragraph
          const start = fullText.length;
          fullText += pText + '\n\n';
          const end = fullText.length;
          
          positions.push({
            start,
            end,
            page: i,
            y: (pageHeight - p.y) / pageHeight // Normalized y (0 at top, 1 at bottom)
          });
          
          fullHtml += pHtml;
        });
      }
      
      return { text: fullText, html: fullHtml, positions };
    } catch (err) {
      console.error("Error in extractTextAndHtmlFromPDF:", err);
      try {
        const text = await file.text();
        return { text, html: text.split('\n').map(l => `<p>${l}</p>`).join(''), positions: [] };
      } catch (e) {
        return { text: '', html: '', positions: [] };
      }
    }
  };

  const extractTextFromPDF = async (file: File): Promise<string> => {
    const data = await extractTextAndHtmlFromPDF(file);
    return data.text;
  };

  // --- Entity Management ---

  const toggleEntitySelection = (id: string) => {
    const newSelected = new Set(selectedIds);
    if (newSelected.has(id)) {
      newSelected.delete(id);
    } else {
      newSelected.add(id);
    }
    setSelectedIds(newSelected);
  };

  const handleValidateSelected = () => {
    const affectedGroupIds = new Set<string>();
    setEntities(prev => prev.map(e => {
      if (selectedIds.has(e.id)) {
        if (e.groupId) affectedGroupIds.add(e.groupId);
        return { ...e, treated: true, ignored: false };
      }
      return e;
    }));
    if (affectedGroupIds.size > 0) {
      setCollapsedGroups(prev => {
        const next = new Set(prev);
        affectedGroupIds.forEach(gid => next.add(gid));
        return next;
      });
    }
    setSelectedIds(new Set());
  };

  const handleBulkReclassAndValidateSelected = (newType: string) => {
    const affectedGroupIds = new Set<string>();
    setEntities(prev => {
      const selectedEntities = prev.filter(e => selectedIds.has(e.id));
      if (selectedEntities.length === 0) return prev;

      const originalToPseudonym: Record<string, string> = {};
      const groupToPseudonym: Record<string, string> = {};

      const updatedEntities = prev.map(e => {
        if (selectedIds.has(e.id)) {
          if (e.groupId) affectedGroupIds.add(e.groupId);
          const lowerOrig = e.original.toLowerCase().trim();
          let pseud = '';
          
          if (e.groupId && groupToPseudonym[e.groupId]) {
            pseud = groupToPseudonym[e.groupId];
          } else if (originalToPseudonym[lowerOrig]) {
            pseud = originalToPseudonym[lowerOrig];
          } else {
            const draftList = prev.filter(item => 
              (!selectedIds.has(item.id)) && 
              (item.groupId ? item.groupId !== e.groupId : true) && 
              (item.original.toLowerCase().trim() !== lowerOrig)
            );
            const temporaryList = [
              ...draftList,
              ...Object.values(originalToPseudonym).map(p => ({ pseudonym: p } as any))
            ];
            
            pseud = getNextPseudonym(newType, temporaryList);
            originalToPseudonym[lowerOrig] = pseud;
            if (e.groupId) {
              groupToPseudonym[e.groupId] = pseud;
            }
          }
          
          return {
            ...e,
            type: newType,
            pseudonym: pseud,
            treated: true,
            ignored: false,
            enabled: true
          };
        }
        return e;
      });

      const next = updatedEntities.map(e => {
        const lowerOrig = e.original.toLowerCase().trim();
        if (originalToPseudonym[lowerOrig]) {
          return {
            ...e,
            type: newType,
            pseudonym: originalToPseudonym[lowerOrig],
            treated: true,
            ignored: false,
            enabled: true
          };
        }
        if (e.groupId && groupToPseudonym[e.groupId]) {
          return {
            ...e,
            type: newType,
            pseudonym: groupToPseudonym[e.groupId],
            treated: true,
            ignored: false,
            enabled: true
          };
        }
        return e;
      });

      setTimeout(() => pushToHistory(next, files), 0);
      return next;
    });

    if (affectedGroupIds.size > 0) {
      setCollapsedGroups(prev => {
        const next = new Set(prev);
        affectedGroupIds.forEach(gid => next.add(gid));
        return next;
      });
    }

    setSelectedIds(new Set());
    showToast(`Elementos validados e reclassificados como ${newType}.`, "success");
  };

  const handleIgnoreSelected = () => {
    const affectedGroupIds = new Set<string>();
    setEntities(prev => prev.map(e => {
      if (selectedIds.has(e.id)) {
        if (e.groupId) affectedGroupIds.add(e.groupId);
        return { ...e, ignored: true, treated: false };
      }
      return e;
    }));
    if (affectedGroupIds.size > 0) {
      setCollapsedGroups(prev => {
        const next = new Set(prev);
        affectedGroupIds.forEach(gid => next.add(gid));
        return next;
      });
    }
    setSelectedIds(new Set());
  };

  const handleAddSelectedToExceptions = () => {
    const selectedEntities = entities.filter(e => selectedIds.has(e.id));
    if (selectedEntities.length === 0) return;

    const newKnowledge = { ...globalKnowledge };
    const textsToRemove = new Set<string>();
    
    selectedEntities.forEach(e => {
      const text = e.original.toLowerCase().trim();
      newKnowledge[text] = 'EXCECAO';
      textsToRemove.add(text);
    });
    
    setGlobalKnowledge(newKnowledge);
    
    setEntities(prev => {
      const next = prev.filter(e => !textsToRemove.has(e.original.toLowerCase().trim()));
      pushToHistory(next, files);
      return next;
    });
    
    setSelectedIds(new Set());
    showToast(`${selectedEntities.length} termos adicionados às EXCEÇÕES GLOBAIS.`, "success");
  };

  const handleAddSelectedToColetiva = () => {
    const selectedEntities = entities.filter(e => selectedIds.has(e.id));
    if (selectedEntities.length === 0) return;

    const newKnowledge = { ...globalKnowledge };
    const textsToRemove = new Set<string>();
    
    selectedEntities.forEach(e => {
      const text = e.original.toLowerCase().trim();
      newKnowledge[text] = 'COLETIVA';
      textsToRemove.add(text);
    });
    
    setGlobalKnowledge(newKnowledge);
    
    setEntities(prev => {
      const next = prev.filter(e => !textsToRemove.has(e.original.toLowerCase().trim()));
      pushToHistory(next, files);
      return next;
    });
    
    setSelectedIds(new Set());
    showToast(`${selectedEntities.length} termos adicionados como PESSOA COLETIVA.`, "success");
  };

  const handleGroupSelected = () => {
    if (selectedIds.size < 2) return;
    
    const selectedEntities = entities.filter(e => selectedIds.has(e.id));
    const groupId = `manual-group-${Date.now()}`;
    
    // Pick the best pseudonym: from a treated entity, or the longest name
    const sorted = [...selectedEntities].sort((a, b) => {
      if (a.treated && !b.treated) return -1;
      if (!a.treated && b.treated) return 1;
      return b.original.length - a.original.length;
    });
    const basePseudonym = sorted[0].pseudonym;

    setEntities(prev => {
      const next = prev.map(e => 
        selectedIds.has(e.id) ? { ...e, groupId, pseudonym: basePseudonym, treated: true } : e
      );
      setTimeout(() => pushToHistory(next, files), 0);
      return next;
    });
    setCollapsedGroups(prev => new Set(prev).add(groupId));
    setSelectedIds(new Set());
    showToast("Elementos agrupados e associados com sucesso.", "success");
  };

  const handleMergeToGroup = (targetGroupId: string) => {
    const targetGroup = entities.find(e => e.groupId === targetGroupId || e.id === targetGroupId);
    if (!targetGroup) return;

    const finalGroupId = targetGroup.groupId || `group-${targetGroup.id}`;
    const finalPseudonym = targetGroup.pseudonym;

    setEntities(prev => prev.map(e => {
      if (selectedIds.has(e.id)) {
        return { ...e, groupId: finalGroupId, pseudonym: finalPseudonym };
      }
      if (e.id === targetGroup.id && !e.groupId) {
        return { ...e, groupId: finalGroupId };
      }
      return e;
    }));
    
    setSelectedIds(new Set());
    setShowMergeModal(false);
  };

  const handleCopyPseudonym = (pseudonym: string) => {
    setCopiedPseudonym(pseudonym);
  };

  const handlePastePseudonym = () => {
    if (!copiedPseudonym || selectedIds.size === 0) return;
    
    const selectedGroupIds = new Set(
      entities
        .filter(e => selectedIds.has(e.id) && e.groupId)
        .map(e => e.groupId)
    );

    setEntities(prev => prev.map(e => {
      if (selectedIds.has(e.id) || (e.groupId && selectedGroupIds.has(e.groupId))) {
        return { ...e, pseudonym: copiedPseudonym };
      }
      return e;
    }));
    setSelectedIds(new Set());
  };

  const updatePseudonym = (id: string, newPseudonym: string) => {
    setEntities(prev => {
      const entity = prev.find(e => e.id === id);
      if (!entity) return prev;

      const next = prev.map(e => {
        if (e.id === id || (entity.groupId && e.groupId === entity.groupId)) {
          return { ...e, pseudonym: newPseudonym };
        }
        return e;
      });
      
      // We call pushToHistory outside or use a timeout to avoid React warnings about state updates during render
      // But since this is an event handler, it's usually fine. 
      // However, to be safe and ensure we use the latest 'files':
      setTimeout(() => pushToHistory(next, files), 0);
      return next;
    });
  };

  const handleSplitEntity = (entity: PIIEntity) => {
    setEditingEntity(entity);
  };

  const handleBulkUpdateType = (newType: string) => {
    if (selectedIds.size === 0) return;
    
    setEntities(prev => {
      const selectedEntities = prev.filter(e => selectedIds.has(e.id));
      const groupIds = new Set(selectedEntities.map(e => e.groupId).filter(Boolean) as string[]);
      
      const byGroup: Record<string, string> = {};
      const byOriginal: Record<string, string> = {};
      
      const next = prev.map(e => {
        const isSelected = selectedIds.has(e.id);
        const isInSelectedGroup = e.groupId && groupIds.has(e.groupId);
        
        if (isSelected || isInSelectedGroup) {
          if (e.groupId) {
            if (!byGroup[e.groupId]) {
              // Generate a pseudonym that isn't used in the new type (excluding this group's current members)
              byGroup[e.groupId] = getNextPseudonym(newType, prev.filter(ent => ent.groupId !== e.groupId));
            }
            return { ...e, type: newType, pseudonym: byGroup[e.groupId], treated: true };
          } else {
            const lowerOrig = e.original.toLowerCase();
            if (!byOriginal[lowerOrig]) {
              byOriginal[lowerOrig] = getNextPseudonym(newType, prev.filter(ent => ent.original.toLowerCase() !== lowerOrig));
            }
            return { ...e, type: newType, pseudonym: byOriginal[lowerOrig], treated: true };
          }
        }
        return e;
      });

      setTimeout(() => pushToHistory(next, files), 0);
      return next;
    });
    
    showToast(`${selectedIds.size} elementos alterados para ${newType}.`, "success");
    setSelectedIds(new Set());
  };

  const reclassifyEntities = () => {
    setEntities(prev => {
      let changed = false;
      const next = prev.map(entity => {
        if (entity.type === 'NOME') {
          const lower = entity.original.toLowerCase().trim();
          const knowledgeType = globalKnowledge[lower];
          
          if (knowledgeType && knowledgeType !== 'EXCECAO' && knowledgeType !== entity.type) {
            changed = true;
            return {
              ...entity,
              type: knowledgeType,
              pseudonym: getNextPseudonym(knowledgeType, prev),
              treated: true
            };
          }

          // Advanced check for partial matches in globalKnowledge
          const judges = Object.entries(globalKnowledge).filter(([_, t]) => t === 'JUIZ').map(([n]) => n.toLowerCase());
          const authors = Object.entries(globalKnowledge).filter(([_, t]) => t === 'AUTOR').map(([n]) => n.toLowerCase());
          const nameWords = lower.split(/\s+/).filter(w => w.length > 2);

          if (nameWords.length >= 2) {
            const isJudgeMatch = judges.some(judgeName => {
              const judgeWords = judgeName.split(/\s+/).filter(w => w.length > 2);
              if (judgeWords.length < 2) return false;
              const common = nameWords.filter(w => judgeWords.includes(w));
              return common.length >= 2;
            });

            if (isJudgeMatch) {
              changed = true;
              return {
                ...entity,
                type: 'JUIZ',
                pseudonym: getNextPseudonym('JUIZ', prev),
                treated: true
              };
            }

            const isAuthorMatch = authors.some(authorName => {
              const authorWords = authorName.split(/\s+/).filter(w => w.length > 2);
              if (authorWords.length < 2) return false;
              const common = nameWords.filter(w => authorWords.includes(w));
              return common.length >= 2;
            });

            if (isAuthorMatch) {
              changed = true;
              return {
                ...entity,
                type: 'AUTOR',
                pseudonym: getNextPseudonym('AUTOR', prev),
                treated: true
              };
            }
          }
        }
        return entity;
      });

      if (changed) {
        const grouped = groupSimilarEntities(next, isRelated);
        setTimeout(() => pushToHistory(grouped, files), 0);
        showToast("Elementos reclassificados com base no conhecimento global.", "success");
        return grouped;
      }
      return prev;
    });
  };

  const handleUpdateGroupType = (groupId: string, newType: string) => {
    setEntities(prev => {
      const groupEntities = prev.filter(e => e.groupId === groupId);
      if (groupEntities.length === 0) return prev;
      
      const newPseudonym = getNextPseudonym(newType, prev.filter(e => e.groupId !== groupId));
      
      const next = prev.map(e => {
        if (e.groupId === groupId) {
          return { ...e, type: newType, pseudonym: newPseudonym, treated: true };
        }
        return e;
      });

      setTimeout(() => pushToHistory(next, files), 0);
      return next;
    });
    showToast(`Categoria do grupo alterada para ${newType}.`, "success");
  };

  const handleUpdateEntity = (id: string, updates: Partial<PIIEntity>) => {
    const entityToUpdate = entities.find(e => e.id === id);
    if (!entityToUpdate) return;

    // If type is changing, update the pseudonym too
    if (updates.type && updates.type !== entityToUpdate.type) {
      updates.pseudonym = getNextPseudonym(updates.type, entities.filter(e => e.id !== id));
      updates.treated = true;
    }

    const updated = { ...entityToUpdate, ...updates };
    const newlySelectedIds: string[] = [];
    
    const nextEntities = entities.map(e => {
      if (e.id === id) return updated;

      // 1. Automatic correction for identical elements (User Request)
      // If the original text is updated, apply the same change to all identical elements
      if (updates.original !== undefined && e.original === entityToUpdate.original) {
        return { ...e, original: updates.original, treated: true };
      }
      
      // 2. Auto-expansion logic for group members when original text is updated
      if (updates.original && updated.groupId && e.groupId === updated.groupId) {
        const words = updated.original.split(/\s+/).filter(w => w.length > 0);
        let curOrig = e.original;
        let curBefore = e.contextBefore || "";
        let curAfter = e.contextAfter || "";
        let changed = false;

        // Expand Start
        while (true) {
          const wordsBefore = curBefore.trim().split(/\s+/);
          const lastWord = wordsBefore[wordsBefore.length - 1];
          if (lastWord && words.some(w => w.toLowerCase() === lastWord.toLowerCase()) && !curOrig.toLowerCase().includes(lastWord.toLowerCase())) {
            curOrig = `${lastWord} ${curOrig}`;
            const lastIdx = curBefore.lastIndexOf(lastWord);
            curBefore = curBefore.substring(0, lastIdx).trimEnd();
            changed = true;
          } else break;
        }

        // Expand End
        while (true) {
          const wordsAfter = curAfter.trim().split(/\s+/);
          const firstWord = wordsAfter[0];
          if (firstWord && words.some(w => w.toLowerCase() === firstWord.toLowerCase()) && !curOrig.toLowerCase().includes(firstWord.toLowerCase())) {
            curOrig = `${curOrig} ${firstWord}`;
            const firstIdx = curAfter.indexOf(firstWord);
            curAfter = curAfter.substring(firstIdx + firstWord.length).trimStart();
            changed = true;
          } else break;
        }

        if (changed) {
          newlySelectedIds.push(e.id);
          return { ...e, original: curOrig, contextBefore: curBefore, contextAfter: curAfter };
        }
      }
      return e;
    });

    setEntities(nextEntities);

    if (newlySelectedIds.length > 0) {
      setSelectedIds(prev => {
        const next = new Set(prev);
        newlySelectedIds.forEach(sid => next.add(sid));
        return next;
      });
    }

    // Side effects after state update
    if (updated.type === 'EXCECAO' || updated.type === 'JUIZ' || updated.type === 'AUTOR' || updated.type === 'COLETIVA') {
      addToGlobalKnowledge(updated.original, updated.type);
    }
    if (editingEntity && editingEntity.id === id) {
      setEditingEntity(updated);
    }
  };

  const handleSplitAllAndEntities = () => {
    setEntities(prev => {
      const next: PIIEntity[] = [];
      let splitCount = 0;
      const affectedGroupIds = new Set<string>();

      prev.forEach(entity => {
        if ((entity.type === 'NOME' || entity.type === 'AUTOR' || entity.type === 'JUIZ' || entity.type === 'COLETIVA') && 
            entity.original.includes(' e ') && 
            !entity.original.toLowerCase().startsWith('e ') && 
            !entity.original.toLowerCase().endsWith(' e')) {
          
          const parts = entity.original.split(/\s+e\s+/i).map(p => p.trim()).filter(p => p.length > 2);
          if (parts.length >= 2) {
            splitCount++;
            if (entity.groupId) affectedGroupIds.add(entity.groupId);

            parts.forEach(part => {
              const id = generateId();
              const pseudonym = getNextPseudonym(entity.type, [...prev, ...next]);
              next.push({
                ...entity,
                id,
                original: part,
                pseudonym,
                groupId: undefined,
                treated: false
              });
            });
            return;
          }
        }
        next.push(entity);
      });

      if (splitCount > 0) {
        // Release members of affected groups so they can find their correct homes
        let finalNext = next.map(e => (e.groupId && affectedGroupIds.has(e.groupId)) ? { ...e, groupId: undefined } : e);
        
        // Re-group everything
        finalNext = groupSimilarEntities(finalNext, isRelated);

        showToast(`${splitCount} nomes compostos foram divididos e grupos re-organizados.`, "success");
        setTimeout(() => pushToHistory(finalNext, files), 0);
        return finalNext;
      } else {
        showToast("Nenhum nome composto com 'e' foi encontrado.", "info");
        return prev;
      }
    });
  };
  const handleManualSplit = (entity: PIIEntity, separator: string | RegExp) => {
    const parts = entity.original.split(separator).map(p => p.trim()).filter(p => p.length > 0);
    if (parts.length < 2) return;

    setSplitResults(parts.map(p => ({
      original: p,
      type: entity.type,
      id: generateId()
    })));
  };

  const confirmSplit = () => {
    if (!editingEntity || !splitResults) return;

    const affectedGroupId = editingEntity.groupId;
    const newEntities: PIIEntity[] = [];
    let currentEntities = [...entities].filter(e => e.id !== editingEntity.id);

    // If the original entity was part of a group, we should probably break that group
    if (affectedGroupId) {
      currentEntities = currentEntities.map(e => e.groupId === affectedGroupId ? { ...e, groupId: undefined } : e);
    }

    splitResults.forEach(res => {
      const id = `split-${editingEntity.id}-${generateId()}`;
      const pseudonym = getNextPseudonym(res.type, currentEntities);
      const newEnt: PIIEntity = {
        id,
        original: res.original,
        type: res.type,
        pseudonym,
        enabled: true,
        ignored: false,
        treated: true,
        fileIds: editingEntity.fileIds,
        context: editingEntity.context,
        groupId: `manual-group-${id}`
      };
      newEntities.push(newEnt);
      currentEntities.push(newEnt);
    });

    // Update global knowledge
    const newKnowledge = { ...globalKnowledge };
    newEntities.forEach(e => {
      newKnowledge[e.original.toLowerCase().trim()] = e.type;
    });
    setGlobalKnowledge(newKnowledge);

    const finalEntities = groupSimilarEntities(currentEntities, isRelated);
    setEntities(finalEntities);
    pushToHistory(finalEntities, files);
    
    setSplitResults(null);
    setEditingEntity(null);
    showToast(`${splitResults.length} novos elementos criados e validados.`, "success");
  };

  const handleFinishAmbiguityReview = () => {
    const grouped = groupSimilarEntities(ambiguousEntities, isRelated);
    setEntities(grouped);
    pushToHistory(grouped, files);
    setShowAmbiguityModal(false);
    setAmbiguousEntities([]);
    showToast("Revisão concluída. Entidades agrupadas.", "success");
  };

  const handleSplitAmbiguous = (id: string, strategy: 'e' | 'comma' | 'space') => {
    setAmbiguousEntities(prev => {
      const next: PIIEntity[] = [];
      prev.forEach(entity => {
        if (entity.id === id) {
          let parts: string[] = [];
          if (strategy === 'e') {
            parts = entity.original.split(/\s+e\s+/i).map(p => p.trim()).filter(p => p.length > 2);
          } else if (strategy === 'comma') {
            parts = entity.original.split(/,/).map(p => p.trim()).filter(p => p.length > 2);
          } else if (strategy === 'space') {
            parts = entity.original.split(/\s{2,}/).map(p => p.trim()).filter(p => p.length > 2);
          }

          if (parts.length >= 2) {
            parts.forEach(part => {
              next.push({
                ...entity,
                id: generateId(),
                original: part,
                pseudonym: getNextPseudonym(entity.type, [...prev, ...next]),
                treated: false,
                reviewed: true
              });
            });
            return;
          }
        }
        next.push(entity);
      });
      pushToHistory(entities, files, next);
      return next;
    });
  };

  const handleMarkAsReviewed = (id: string) => {
    setAmbiguousEntities(prev => {
      const next = prev.map(e => e.id === id ? { ...e, reviewed: !e.reviewed } : e);
      pushToHistory(entities, files, next);
      return next;
    });
  };

  const handleAddToExceptionsFromAmbiguity = (text: string) => {
    addToGlobalKnowledge(text, 'EXCECAO');
    setAmbiguousEntities(prev => {
      const next = prev.filter(e => e.original.toLowerCase() !== text.toLowerCase());
      pushToHistory(entities, files, next);
      return next;
    });
    showToast(`"${text}" adicionado às exceções globais.`, "info");
  };

  const handleIgnoreSimilar = (entity: PIIEntity) => {
    const norm = entity.original.toLowerCase().trim();
    setEntities(prev => {
      const next = prev.map(e => 
        e.original.toLowerCase().trim() === norm 
          ? { ...e, ignored: true, treated: false } 
          : e
      );
      showToast(`Todas as ocorrências de "${entity.original}" foram marcadas para ignorar.`, "info");
      setTimeout(() => pushToHistory(next, files), 0);
      return next;
    });
  };

  const handleDissolveGroup = (groupId: string) => {
    setEntities(prev => {
      const next = prev.map(e => e.groupId === groupId ? { ...e, groupId: undefined, treated: false } : e);
      showToast("Grupo dissolvido. Os elementos podem agora ser re-agrupados ou editados individualmente.", "success");
      setTimeout(() => pushToHistory(next, files), 0);
      return next;
    });
  };

  const handleSplitGroup = (groupId: string, strategy: 'e' | 'space' = 'e') => {
    setEntities(prev => {
      const next: PIIEntity[] = [];
      let splitOccurred = false;
      
      prev.forEach(entity => {
        if (entity.groupId === groupId && 
            (entity.type === 'NOME' || entity.type === 'AUTOR' || entity.type === 'JUIZ')) {
          
          let parts: string[] = [];
          if (strategy === 'e' && entity.original.includes(' e ')) {
            parts = entity.original.split(/\s+e\s+/i).map(p => p.trim()).filter(p => p.length > 2);
          } else if (strategy === 'space' && entity.original.includes(' ')) {
            // Heuristic: split if we have multiple capitalized sequences that look like full names
            // For simplicity, we'll just split by double spaces or let the user decide.
            // Let's do a more aggressive split by space but filter out short words
            parts = entity.original.split(/\s{2,}/).map(p => p.trim()).filter(p => p.length > 2);
            if (parts.length < 2) {
              // Try to split by single space if it's a very long string (e.g. 4+ words)
              const words = entity.original.split(/\s+/);
              if (words.length >= 4) {
                // Split in half as a guess, or just split all words
                parts = words.map(w => w.trim()).filter(w => w.length > 2);
              }
            }
          }

          if (parts.length >= 2) {
            splitOccurred = true;
            parts.forEach(part => {
              const id = generateId();
              const pseudonym = getNextPseudonym(entity.type, [...prev, ...next]);
              next.push({
                ...entity,
                id,
                original: part,
                pseudonym,
                 groupId: `manual-group-${id}`,
                treated: true
              });
            });
            return;
          }
        }
        if (entity.groupId === groupId) {
          next.push({ ...entity, groupId: undefined });
          return;
        }
        next.push(entity);
      });

      if (splitOccurred) {
        const grouped = groupSimilarEntities(next, isRelated);
        showToast("Elementos divididos e re-organizados.", "success");
        setTimeout(() => pushToHistory(grouped, files), 0);
        return grouped;
      }
      return prev;
    });
  };

  const handleUnlockGroup = (groupId: string) => {
    setEntities(prev => {
      const next = prev.map(e => e.groupId === groupId ? { ...e, groupId: undefined } : e);
      const grouped = groupSimilarEntities(next, isRelated);
      showToast("Grupo libertado para re-agrupamento automático.", "success");
      setTimeout(() => pushToHistory(grouped, files), 0);
      return grouped;
    });
  };

  const addToGlobalKnowledge = (text: string, type: string = 'EXCECAO') => {
    setGlobalKnowledge(prev => ({ ...prev, [text.toLowerCase().trim()]: type }));
    // Only remove from entities if it's a generic exception, not a judge/author
    if (type === 'EXCECAO') {
      setEntities(prev => prev.filter(e => e.original.toLowerCase() !== text.toLowerCase()));
    } else {
      // For judges/authors, update their type in the current list
      setEntities(prev => prev.map(e => 
        e.original.toLowerCase() === text.toLowerCase() ? { ...e, type, treated: true } : e
      ));
    }
  };

  const handleImportJudges = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    
    setIsProcessing(true);
    try {
      let text = '';
      if (file.type === 'application/pdf') {
        text = await extractTextFromPDF(file);
      } else {
        text = await file.text();
      }
      
      if (!text || text.trim().length === 0) {
        showToast("Não foi possível extrair texto do ficheiro. Verifique se o PDF contém texto pesquisável.", "error");
        return;
      }

      // Split by lines and filter
      // Use a more flexible split to catch names separated by various delimiters
      const rawLines = text.split(/[\n\r,;|\t]+/).map(l => l.trim()).filter(l => l.length > 3);
      const newJudges: Record<string, string> = {};
      let count = 0;
      
      rawLines.forEach(line => {
        const cleaned = cleanName(line);
        // Basic validation: at least two words and minimum length
        if (cleaned.length >= 3 && cleaned.split(/\s+/).length >= 2) {
          newJudges[cleaned.toLowerCase()] = 'JUIZ';
          count++;
        }
      });
      
      if (count === 0) {
        showToast("Nenhum nome de juiz válido foi encontrado no ficheiro. Certifique-se de que os nomes estão completos.", "error");
      } else {
        setGlobalKnowledge(prev => {
          const next = { ...prev };
          // The user mentioned "limpa os nomes... que tenho na lista atual" for authors, 
          // but for judges they said "quero que a troques pela que está no documento".
          // So for both, we should probably clear the existing ones of that type.
          Object.keys(next).forEach(key => {
            if (next[key] === 'JUIZ') delete next[key];
          });
          return { ...next, ...newJudges };
        });
        showToast(`${count} juízes importados e lista atualizada.`, "success");
      }
    } catch (err) {
      console.error("Erro ao importar juízes:", err);
      showToast("Erro ao processar o ficheiro de juízes.", "error");
    } finally {
      setIsProcessing(false);
      event.target.value = '';
    }
  };

  const handleImportAuthors = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    
    setIsProcessing(true);
    try {
      let text = '';
      if (file.type === 'application/pdf') {
        text = await extractTextFromPDF(file);
      } else {
        text = await file.text();
      }
      
      if (!text || text.trim().length === 0) {
        showToast("Não foi possível extrair texto do ficheiro. Verifique se o PDF contém texto pesquisável.", "error");
        return;
      }

      const rawLines = text.split(/[\n\r,;|\t]+/).map(l => l.trim()).filter(l => l.length > 3);
      const newAuthors: Record<string, string> = {};
      let count = 0;
      
      rawLines.forEach(line => {
        const cleaned = cleanName(line);
        if (cleaned.length >= 3 && cleaned.split(/\s+/).length >= 2) {
          newAuthors[cleaned.toLowerCase()] = 'AUTOR';
          count++;
        }
      });
      
      if (count === 0) {
        showToast("Nenhum nome de autor válido foi encontrado no ficheiro.", "error");
      } else {
        setGlobalKnowledge(prev => {
          const next = { ...prev };
          // User specifically asked to clear existing authors: "limpa os nomes de autores que tenho na lista atual"
          Object.keys(next).forEach(key => {
            if (next[key] === 'AUTOR') delete next[key];
          });
          return { ...next, ...newAuthors };
        });
        showToast(`${count} autores importados e lista atualizada.`, "success");
      }
    } catch (err) {
      console.error("Erro ao importar autores:", err);
      showToast("Erro ao processar o ficheiro de autores.", "error");
    } finally {
      setIsProcessing(false);
      event.target.value = '';
    }
  };

  const handleImportColetivas = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    
    setIsProcessing(true);
    try {
      let text = '';
      if (file.type === 'application/pdf') {
        text = await extractTextFromPDF(file);
      } else {
        text = await file.text();
      }
      
      if (!text || text.trim().length === 0) {
        showToast("Não foi possível extrair texto do ficheiro.", "error");
        return;
      }

      const rawLines = text.split(/[\n\r,;|\t]+/).map(l => l.trim()).filter(l => l.length > 3);
      const newColetivas: Record<string, string> = {};
      let count = 0;
      
      rawLines.forEach(line => {
        const cleaned = line.replace(/\s+/g, ' ').trim();
        if (cleaned.length >= 3) {
          newColetivas[cleaned.toLowerCase()] = 'COLETIVA';
          count++;
        }
      });
      
      if (count === 0) {
        showToast("Nenhuma pessoa coletiva válida foi encontrada.", "error");
      } else {
        setGlobalKnowledge(prev => {
          const next = { ...prev };
          Object.keys(next).forEach(key => {
            if (next[key] === 'COLETIVA') delete next[key];
          });
          return { ...next, ...newColetivas };
        });
        showToast(`${count} pessoas coletivas importadas e lista atualizada.`, "success");
      }
    } catch (err) {
      console.error("Erro ao importar pessoas coletivas:", err);
      showToast("Erro ao processar o ficheiro.", "error");
    } finally {
      setIsProcessing(false);
      event.target.value = '';
    }
  };

  const handleClearGlobalKnowledge = (type: string) => {
    setGlobalKnowledge(prev => {
      const next = { ...prev };
      Object.keys(next).forEach(key => {
        if (next[key] === type) delete next[key];
      });
      return next;
    });
    
    if (type === 'HEADER') {
      setHeaderZones([]);
      setFooterZones([]);
    }

    const typeName = type === 'EXCECAO' ? 'exceções' : type === 'JUIZ' ? 'juízes' : type === 'AUTOR' ? 'autores' : type === 'HEADER' ? 'cabeçalhos' : 'coletivas';
    showToast(`Lista de ${typeName} limpa com sucesso.`, "info");
  };

  const handleClearAllGlobalKnowledge = () => {
    setGlobalKnowledge({});
    setHeaderZones([]);
    setFooterZones([]);
    showToast("Todo o conhecimento global foi limpo.", "info");
  };

  const handleStartEditKnowledge = (text: string, type: 'KNOWLEDGE' | 'SAFELIST_WORD' | 'SAFELIST_PHRASE', category?: string) => {
    setEditingKnowledge({ original: text, current: text, type, category });
  };

  const handleSaveEditKnowledge = () => {
    if (!editingKnowledge) return;
    const { original, current, type, category } = editingKnowledge;
    const trimmed = current.trim();
    if (!trimmed) return;

    if (type === 'KNOWLEDGE') {
      setGlobalKnowledge(prev => {
        const next = { ...prev };
        const oldType = next[original];
        delete next[original];
        next[trimmed.toLowerCase()] = category || oldType || 'EXCECAO';
        return next;
      });
    } else if (type === 'SAFELIST_WORD') {
      setSafelist(prev => ({
        ...prev,
        words_ignore: prev.words_ignore.map(w => w === original ? trimmed : w)
      }));
    } else if (type === 'SAFELIST_PHRASE') {
      setSafelist(prev => ({
        ...prev,
        phrases_ignore: prev.phrases_ignore.map(p => p === original ? trimmed : p)
      }));
    }

    setEditingKnowledge(null);
    showToast("Alteração guardada com sucesso.", "success");
  };

  const handleIdentifyDuplicates = (type: string) => {
    const items = Object.entries(globalKnowledge)
      .filter(([_, t]) => t === type)
      .map(([text]) => text);
    
    const suggestions: Array<{ item1: string, item2: string, score: number, type: string }> = [];
    
    for (let i = 0; i < items.length; i++) {
      for (let j = i + 1; j < items.length; j++) {
        const s1 = items[i];
        const s2 = items[j];
        
        // Simple similarity: check if one is contained in another or share many words
        const w1 = s1.split(/\s+/);
        const w2 = s2.split(/\s+/);
        const common = w1.filter(w => w2.includes(w));
        const score = common.length / Math.max(w1.length, w2.length);
        
        if (score >= 0.6 || s1.includes(s2) || s2.includes(s1)) {
          suggestions.push({ item1: s1, item2: s2, score, type });
        }
      }
    }
    
    if (suggestions.length === 0) {
      showToast("Não foram encontradas duplicações óbvias.", "info");
    } else {
      setDeduplicationSuggestions(suggestions);
      setShowDeduplicationModal(true);
    }
  };

  const handleResolveDuplicate = (keep: string, remove: string) => {
    setGlobalKnowledge(prev => {
      const next = { ...prev };
      delete next[remove];
      return next;
    });
    setDeduplicationSuggestions(prev => prev.filter(s => !(s.item1 === remove || s.item2 === remove)));
    showToast(`Mantido: "${keep}". Removido: "${remove}".`, "success");
  };

  const handleResolveAllDuplicates = () => {
    const toRemove = new Set<string>();
    deduplicationSuggestions.forEach(s => {
      // Keep item1, remove item2 by default
      toRemove.add(s.item2);
    });

    setGlobalKnowledge(prev => {
      const next = { ...prev };
      toRemove.forEach(name => delete next[name]);
      return next;
    });

    setDeduplicationSuggestions([]);
    showToast(`${toRemove.size} sugestões de duplicação resolvidas automaticamente.`, "success");
  };

  const handleValidateAllEntities = () => {
    setEntities(prev => {
      const next = prev.map(e => {
        if (e.enabled && !e.ignored && !e.treated) {
          return { ...e, treated: true, enabled: true };
        }
        return e;
      });
      return next;
    });
    showToast("Todos os elementos foram validados.", "success");
  };

  const handleDiscardSuggestion = (index: number) => {
    setDeduplicationSuggestions(prev => prev.filter((_, i) => i !== index));
    showToast("Sugestão descartada.", "info");
  };

  const handleSuggestGroups = () => {
    // More aggressive grouping for names
    let next: PIIEntity[] = [];
    setEntities(prev => {
      next = [...prev];
      const nameEntities = next.filter(e => e.type === 'NOME' || e.type === 'JUIZ' || e.type === 'AUTOR');
      
      nameEntities.forEach(entity => {
        if (entity.groupId) return;
        
        const words = entity.original.toLowerCase().split(/\s+/).filter(w => w.length > 2);
        if (words.length < 2) return;

        const match = nameEntities.find(other => {
          if (other.id === entity.id) return false;
          const otherWords = other.original.toLowerCase().split(/\s+/).filter(w => w.length > 2);
          if (otherWords.length < 2) return false;
          
          // Match if first name + any other name matches
          if (words[0] === otherWords[0]) {
            const otherMatch = words.slice(1).some(w => otherWords.slice(1).includes(w));
            if (otherMatch) return true;
          }
          
          // Or if they share at least 2 significant words
          const common = words.filter(w => otherWords.includes(w));
          return common.length >= 2;
        });

        if (match) {
          const groupId = match.groupId || `suggested-${entity.id}`;
          entity.groupId = groupId;
          match.groupId = groupId;
          
          // Sync type if one is JUIZ or AUTOR
          if (match.type === 'JUIZ' || entity.type === 'JUIZ') {
            entity.type = 'JUIZ';
            match.type = 'JUIZ';
          } else if (match.type === 'AUTOR' || entity.type === 'AUTOR') {
            entity.type = 'AUTOR';
            match.type = 'AUTOR';
          }
          
          // Sync pseudonyms
          entity.pseudonym = match.pseudonym;
        }
      });
      
      return next;
    });
    pushToHistory(next, files);
    showToast("Sugestões de grupos aplicadas", "success");
  };

  const handleRescan = async () => {
    if (files.length === 0) return;
    setIsProcessing(true);
    try {
      const preRescanIds = new Set(entities.map(e => e.id));
      const preRescanOriginals = new Set(entities.map(e => e.original.toLowerCase().trim()));

      // Keep track of original treated or ignored entities to restore them untouched
      const originalTreatedAndIgnoredMap = new Map<string, PIIEntity>();
      entities.forEach(e => {
        if (e.treated || e.ignored) {
          originalTreatedAndIgnoredMap.set(e.id, { ...e });
        }
      });

      // 1. Collect all currently identified entities to ensure consistency
      const sessionKnowledge: Record<string, string> = { ...globalKnowledge };
      entities.forEach(e => {
        if (!e.ignored) {
          sessionKnowledge[e.original.toLowerCase().trim()] = e.type;
        }
      });

      const foundInScan: PIIEntity[] = [];
      const foundIdsMap = new Map<string, PIIEntity>();
      
      for (const file of files) {
        const newEntities = scanText(file.content, file.id, entities, isRelated, sessionKnowledge, safelist, file.positions || []);
        
        // Filter out new entities that are already in the entities list and treated/ignored
        const filteredNew = newEntities.filter(ne => {
          const existing = entities.find(e => e.original === ne.original && e.type === ne.type);
          return !(existing && (existing.treated || existing.ignored));
        });

        filteredNew.forEach(ne => {
          if (foundIdsMap.has(ne.id)) {
            const existing = foundIdsMap.get(ne.id)!;
            ne.fileIds?.forEach(fid => {
              if (!existing.fileIds?.includes(fid)) {
                existing.fileIds = [...(existing.fileIds || []), fid];
              }
            });
          } else {
            foundIdsMap.set(ne.id, ne);
            foundInScan.push(ne);
          }
        });
      }

      // 2. Merge logic:
      // Keep all entities from foundInScan
      // AND any entities from the original list that were NOT found OR were already treated/ignored
      const finalEntities = [...foundInScan];
      
      entities.forEach(existing => {
        const isFound = foundIdsMap.has(existing.id);
        const isTreatedOrIgnored = existing.treated || existing.ignored;
        const isManual = existing.groupId?.startsWith('manual-group-');

        if (isTreatedOrIgnored || isManual) {
          // Always keep treated/ignored/manual work
          // If it was also found in scan, we prefer the existing one because it has user validation state
          if (isFound) {
            const idx = finalEntities.findIndex(f => f.id === existing.id);
            if (idx !== -1) finalEntities[idx] = existing;
          } else {
            finalEntities.push(existing);
          }
        } else if (!isFound) {
          // If it wasn't found in scan and wasn't treated, we might still want to keep it if it was from a previous scan
          // but the user said rescan is to show what hasn't been treated.
          // Let's keep it to be safe, but mark as not treated.
          finalEntities.push(existing);
        }
      });

      // 3. Group and update
      let grouped = groupSimilarEntities(finalEntities, isRelated);

      // Restore treated and ignored elements EXACTLY as they were to keep them untouched
      // AND flag any newly identified entities
      grouped = grouped.map(e => {
        if (originalTreatedAndIgnoredMap.has(e.id)) {
          const original = originalTreatedAndIgnoredMap.get(e.id)!;
          const combinedFileIds = Array.from(new Set([...(original.fileIds || []), ...(e.fileIds || [])]));
          return {
            ...original,
            fileIds: combinedFileIds,
            isNewlyIdentified: false
          };
        }
        const isNew = !preRescanIds.has(e.id) && !preRescanOriginals.has(e.original.toLowerCase().trim());
        return {
          ...e,
          isNewlyIdentified: isNew
        };
      });

      const hasNew = grouped.some(e => e.isNewlyIdentified);

      setEntities(grouped);
      pushToHistory(grouped, files);

      if (hasNew) {
        showToast("Re-análise concluída. Novos elementos identificados.", "success");
      } else {
        showToast("Re-análise concluída. Nenhum novo elemento foi identificado.", "info");
      }
    } catch (err) {
      console.error("Erro durante a re-análise:", err);
      showToast("Erro durante a re-análise.", "error");
    } finally {
      setIsProcessing(false);
    }
  };

  const handleExpandStart = () => {
    if (!editingEntity) return;
    const words = editingEntity.contextBefore.trim().split(/\s+/);
    if (words.length === 0) return;
    const lastWord = words[words.length - 1];
    const newOriginal = `${lastWord} ${editingEntity.original}`;
    const newContextBefore = editingEntity.contextBefore.slice(0, editingEntity.contextBefore.lastIndexOf(lastWord)).trim();
    
    const updated = { ...editingEntity, original: newOriginal, contextBefore: newContextBefore };
    setEditingEntity(updated);
    handleUpdateEntity(editingEntity.id, updated);
  };

  const handleShrinkStart = () => {
    if (!editingEntity) return;
    const words = editingEntity.original.trim().split(/\s+/);
    if (words.length < 2) return;
    const firstWord = words[0];
    const newOriginal = words.slice(1).join(' ');
    const newContextBefore = `${editingEntity.contextBefore} ${firstWord}`.trim();
    
    const updated = { ...editingEntity, original: newOriginal, contextBefore: newContextBefore };
    setEditingEntity(updated);
    handleUpdateEntity(editingEntity.id, updated);
  };

  const handleExpandEnd = () => {
    if (!editingEntity) return;
    const words = editingEntity.contextAfter.trim().split(/\s+/);
    if (words.length === 0) return;
    const firstWord = words[0];
    const newOriginal = `${editingEntity.original} ${firstWord}`;
    const newContextAfter = editingEntity.contextAfter.slice(editingEntity.contextAfter.indexOf(firstWord) + firstWord.length).trim();
    
    const updated = { ...editingEntity, original: newOriginal, contextAfter: newContextAfter };
    setEditingEntity(updated);
    handleUpdateEntity(editingEntity.id, updated);
  };

  const handleShrinkEnd = () => {
    if (!editingEntity) return;
    const words = editingEntity.original.trim().split(/\s+/);
    if (words.length < 2) return;
    const lastWord = words[words.length - 1];
    const newOriginal = words.slice(0, -1).join(' ');
    const newContextAfter = `${lastWord} ${editingEntity.contextAfter}`.trim();
    
    const updated = { ...editingEntity, original: newOriginal, contextAfter: newContextAfter };
    setEditingEntity(updated);
    handleUpdateEntity(editingEntity.id, updated);
  };

  const toggleSelectAllFiltered = () => {
    const allFilteredIds = filteredEntities.map(e => e.id);
    const allSelected = allFilteredIds.length > 0 && allFilteredIds.every(id => selectedIds.has(id));
    
    const newSelected = new Set(selectedIds);
    if (allSelected) {
      allFilteredIds.forEach(id => newSelected.delete(id));
    } else {
      allFilteredIds.forEach(id => newSelected.add(id));
    }
    setSelectedIds(newSelected);
  };

  const toggleSelectGroup = (groupEntities: PIIEntity[]) => {
    const groupIds = groupEntities.map(e => e.id);
    const allSelected = groupIds.length > 0 && groupIds.every(id => selectedIds.has(id));
    
    const newSelected = new Set(selectedIds);
    if (allSelected) {
      groupIds.forEach(id => newSelected.delete(id));
    } else {
      groupIds.forEach(id => newSelected.add(id));
    }
    setSelectedIds(newSelected);
  };

  const clearAll = () => {
    if (!confirmingClear) {
      setConfirmingClear(true);
      showToast("Clique novamente para confirmar a limpeza total do projeto.", "info");
      setTimeout(() => setConfirmingClear(false), 4000);
      return;
    }
    
    setFiles([]);
    setEntities([]);
    setSelectedIds(new Set());
    setSelectedFileId(null);
    setHistory([]);
    setHistoryIndex(-1);
    setConfirmingClear(false);
    localStorage.removeItem('pii_project_state');
    showToast("Projeto limpo. Iniciado novo projeto.", "success");
  };

  const handleReGroup = () => {
    const grouped = groupSimilarEntities(entities, isRelated);
    setEntities(grouped);
    pushToHistory(grouped, files);
    showToast("Agrupamentos re-analisados", "success");
  };

  const handleFixAllConflicts = () => {
    setEntities(prev => {
      // 1. Re-calculate conflicts based on the absolute latest state 'prev'
      const pToGroups = new Map<string, Set<string>>();
      prev.forEach(e => {
        if (e.ignored || !e.enabled) return;
        const p = e.pseudonym.trim();
        if (!p) return;
        const groupKey = e.groupId || `single-${e.id}`;
        if (!pToGroups.has(p)) pToGroups.set(p, new Set());
        pToGroups.get(p)!.add(groupKey);
      });
      
      const conflicts = new Set<string>();
      pToGroups.forEach((groups, pseudonym) => {
        if (groups.size > 1) conflicts.add(pseudonym);
      });

      if (conflicts.size === 0) {
        // Use a timeout for side effects like toast/history to avoid issues inside setEntities
        setTimeout(() => showToast("Não foram encontrados conflitos", "info"), 0);
        return prev;
      }

      const next = [...prev];
      // Keep track of how many "new" unique names we've introduced globally in this batch
      // to ensure getNextPseudonym always moves forward
      let extraNamesCount = 0;

      conflicts.forEach(pseudonym => {
        const groups = Array.from(pToGroups.get(pseudonym) || []);
        // Keep the first group with this pseudonym, change the others
        for (let i = 1; i < groups.length; i++) {
          const groupToChange = groups[i];
          
          // Find an entity in this group to get the type
          const sampleEntity = next.find(e => (e.groupId && e.groupId === groupToChange) || (!e.groupId && `single-${e.id}` === groupToChange));
          if (!sampleEntity) continue;
          
          const type = sampleEntity.type;
          
          // Find a new unique pseudonym
          let newP = "";
          let foundUnique = false;
          let safety = 0;
          
          while (!foundUnique && safety < 200) {
            extraNamesCount++;
            safety++;
            
            // Create dummies to force getNextPseudonym to give us a higher count
            const dummies = Array.from({ length: extraNamesCount }, (_, k) => ({
              id: `temp-${k}`,
              original: `TEMP_UNIQUE_NAME_${type}_${k}`,
              type,
              pseudonym: '',
              enabled: true
            } as PIIEntity));
            
            newP = getNextPseudonym(type, [...next, ...dummies]);
            
            // Ensure it doesn't conflict with existing ones in 'next'
            if (!next.some(e => e.pseudonym === newP)) {
              foundUnique = true;
            }
          }

          // Update all entities in this group
          next.forEach((e, idx) => {
            const currentGroupKey = e.groupId || `single-${e.id}`;
            if (currentGroupKey === groupToChange) {
              next[idx] = { ...e, pseudonym: newP, treated: true };
            }
          });
        }
      });

      setTimeout(() => {
        pushToHistory(next, files);
        showToast("Conflitos de pseudónimos resolvidos", "success");
      }, 0);
      
      return next;
    });
  };

  const handleFixSingleConflict = (entity: PIIEntity) => {
    setEntities(prev => {
      const type = entity.type;
      const groupKey = entity.groupId || `single-${entity.id}`;
      
      let newP = "";
      let extra = 0;
      let found = false;
      
      while (!found && extra < 100) {
        extra++;
        const dummies = Array.from({ length: extra }, (_, k) => ({
          id: `temp-${k}`,
          original: `TEMP_SINGLE_FIX_${k}`,
          type,
          pseudonym: '',
          enabled: true
        } as PIIEntity));
        
        newP = getNextPseudonym(type, [...prev, ...dummies]);
        if (!prev.some(e => e.pseudonym === newP)) {
          found = true;
        }
      }
      
      const next = prev.map(e => {
        const currentGroupKey = e.groupId || `single-${e.id}`;
        if (currentGroupKey === groupKey) {
          return { ...e, pseudonym: newP, treated: true };
        }
        return e;
      });
      
      setTimeout(() => {
        pushToHistory(next, files);
        showToast(`Novo pseudónimo atribuído a ${entity.original}`, "success");
      }, 0);
      
      return next;
    });
  };

  const handleImportSafelist = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      let newWords: string[] = [];
      let newPhrases: string[] = [];

      if (file.name.endsWith('.json')) {
        const text = await file.text();
        const data = JSON.parse(text);
        // Suporta tanto o formato simples quanto o formato estruturado do user
        if (data.words_ignore || data.phrases_ignore) {
          newWords = data.words_ignore || [];
          newPhrases = data.phrases_ignore || [];
        } else if (Array.isArray(data)) {
          // Fallback para array simples
          data.forEach(item => {
            if (typeof item === 'string') {
              if (item.includes(' ')) newPhrases.push(item);
              else newWords.push(item);
            }
          });
        }
      } else if (file.name.endsWith('.txt')) {
        const text = await file.text();
        const lines = text.split('\n').map(l => l.trim()).filter(l => l.length > 0);
        lines.forEach(line => {
          if (line.includes(' ')) {
            newPhrases.push(line);
          } else {
            newWords.push(line);
          }
        });
      }

      setSafelist(prev => {
        const words = new Set([...prev.words_ignore, ...newWords]);
        const phrases = new Set([...prev.phrases_ignore, ...newPhrases]);
        return {
          words_ignore: Array.from(words),
          phrases_ignore: Array.from(phrases)
        };
      });

      showToast(`Safelist importada: ${newWords.length} palavras, ${newPhrases.length} frases.`, "success");
    } catch (err) {
      console.error("Erro ao importar safelist:", err);
      showToast("Erro ao importar ficheiro de safelist.", "error");
    }
    e.target.value = '';
  };

  const handleExportAllKnowledge = () => {
    const data = {
      globalKnowledge,
      safelist,
      version: "1.0",
      exportDate: new Date().toISOString()
    };
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
    saveAs(blob, `conhecimento_anonymiza_${new Date().toISOString().slice(0,10)}.json`);
    showToast("Todo o conhecimento exportado com sucesso", "success");
  };

  const handleImportAllKnowledge = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      const text = await file.text();
      const data = JSON.parse(text);
      
      if (data.globalKnowledge) {
        setGlobalKnowledge(prev => ({ ...prev, ...data.globalKnowledge }));
      }
      
      if (data.safelist) {
        setSafelist(prev => {
          const words = new Set([...prev.words_ignore, ...(data.safelist.words_ignore || [])]);
          const phrases = new Set([...prev.phrases_ignore, ...(data.safelist.phrases_ignore || [])]);
          return {
            words_ignore: Array.from(words),
            phrases_ignore: Array.from(phrases)
          };
        });
      }
      
      showToast("Conhecimento importado e mesclado com sucesso", "success");
    } catch (err) {
      console.error("Erro ao importar conhecimento:", err);
      showToast("Erro ao importar ficheiro de conhecimento.", "error");
    }
    e.target.value = '';
  };

  const handleExportExceptions = () => {
    const blob = new Blob([JSON.stringify(globalKnowledge, null, 2)], { type: 'application/json' });
    saveAs(blob, 'conhecimento_global.json');
  };

  const handleTransferExceptionsToSafelist = () => {
    const exceptionsToMove = Object.entries(globalKnowledge).filter(([_, type]) => type === 'EXCECAO');
    if (exceptionsToMove.length === 0) {
      showToast("Nenhuma exceção para transferir.", "info");
      return;
    }

    const newWords: string[] = [];
    const newPhrases: string[] = [];

    exceptionsToMove.forEach(([text]) => {
      if (text.includes(' ')) {
        newPhrases.push(text);
      } else {
        newWords.push(text);
      }
    });

    setSafelist(prev => {
      const words = new Set([...prev.words_ignore, ...newWords]);
      const phrases = new Set([...prev.phrases_ignore, ...newPhrases]);
      return {
        words_ignore: Array.from(words),
        phrases_ignore: Array.from(phrases)
      };
    });

    setGlobalKnowledge(prev => {
      const next = { ...prev };
      exceptionsToMove.forEach(([text]) => {
        delete next[text];
      });
      return next;
    });

    showToast(`${exceptionsToMove.length} exceções transferidas para a Safelist.`, "success");
  };

  const [showExportModal, setShowExportModal] = useState(false);
  const [exportSettings, setExportSettings] = useState<ExportSettings>({
    format: 'pdf',
    removeHeadersFooters: false,
    footnotesAtEnd: true,
    preserveFormatting: true,
    sanitizeMetadata: true
  });

  const extractTextFromZone = async (rect: { top: number; bottom: number; left: number; right: number }, type: string, pageNum: number) => {
    const fileData = files.find(f => f.id === selectedFileId);
    if (!fileData || !fileData.rawFile) return;

    try {
      const arrayBuffer = await fileData.rawFile.arrayBuffer();
      const pdf = await pdfjs.getDocument({ data: arrayBuffer }).promise;
      
      const page = await pdf.getPage(pageNum);
      const viewport = page.getViewport({ scale: 1.0 });
      const textContent = await page.getTextContent();
      
      const extractedTexts: string[] = [];
      
      textContent.items.forEach((item: any) => {
        const [x, y] = viewport.convertToViewportPoint(item.transform[4], item.transform[5]);
        
        // Convert to percentage
        const itemTop = y / viewport.height;
        const itemLeft = x / viewport.width;
        
        // Check if item is inside the rect
        if (itemTop >= rect.top && itemTop <= rect.bottom &&
            itemLeft >= rect.left && itemLeft <= rect.right) {
          if (item.str.trim()) {
            extractedTexts.push(item.str.trim());
          }
        }
      });

      console.log('Extracted texts from zone:', extractedTexts);

      if (extractedTexts.length > 0) {
        // Add each line as a separate header pattern
        extractedTexts.forEach(text => {
          if (text.length > 2) {
            setGlobalKnowledge(prev => ({
              ...prev,
              [text]: type
            }));
          }
        });

        // Also add the combined text just in case
        const fullText = extractedTexts.join(' ');
        if (fullText.length > 3) {
          setGlobalKnowledge(prev => ({
            ...prev,
            [fullText]: type
          }));
          
          // If it's a signature or process line, also add the prefix
          if (fullText.toLowerCase().includes('processo:')) {
            setGlobalKnowledge(prev => ({ ...prev, ['Processo:']: type }));
          }
          if (fullText.toLowerCase().includes('referência:')) {
            setGlobalKnowledge(prev => ({ ...prev, ['Referência:']: type }));
          }
          if (fullText.toLowerCase().includes('assinado')) {
            setGlobalKnowledge(prev => ({ ...prev, ['Assinado']: type }));
          }
        }
        
        // Visual feedback
        const msg = type === 'HEADER' ? 'Zona de Cabeçalho/Rodapé adicionada' : 'Zona adicionada';
        const toast = document.createElement('div');
        toast.className = 'fixed bottom-4 right-4 bg-gray-800 text-white px-4 py-2 rounded-lg shadow-lg z-[100] text-sm animate-in fade-in slide-in-from-bottom-4';
        toast.textContent = msg;
        document.body.appendChild(toast);
        setTimeout(() => toast.remove(), 3000);
      }
    } catch (err) {
      console.error("Error extracting text from zone:", err);
    }
  };

  const handleExport = async () => {
    setIsProcessing(true);
    setShowExportModal(false);
    const zip = new JSZip();
    const processedFiles: { name: string, data: Uint8Array | string | Blob, isPdf: boolean, format: string }[] = [];
    
    try {
      for (const fileData of files) {
        if (fileData.status !== 'done') continue;

        const baseName = fileData.name.replace(/\.[^/.]+$/, "");
        
        if (exportSettings.format === 'pdf') {
          if (fileData.type === 'application/pdf') {
            try {
              // Try the highly robust method for original PDFs first
              const pdfBytes = await exportAnonymizedPDFBytes(fileData);
              processedFiles.push({ name: `${baseName}_anon.pdf`, data: pdfBytes, isPdf: true, format: 'pdf' });
              zip.file(`${baseName}_anon.pdf`, pdfBytes);
            } catch (err) {
              console.warn(`Original PDF ${fileData.name} not available or failed layout export, using fallback:`, err);
              // Fallback 1: Reconstructed layout from HTML if available
              if (fileData.htmlContent) {
                const pdfBlob = await createPdfFromHtml(fileData.htmlContent, fileData.name, fileData.id);
                processedFiles.push({ name: `${baseName}_anon.pdf`, data: pdfBlob, isPdf: true, format: 'pdf' });
                zip.file(`${baseName}_anon.pdf`, pdfBlob);
              } else {
                // Fallback 2: Simple text-to-pdf
                const pdfBlob = await generateAnonymizedPdfFromText(fileData);
                processedFiles.push({ name: `${baseName}_anon.pdf`, data: pdfBlob, isPdf: true, format: 'pdf' });
                zip.file(`${baseName}_anon.pdf`, pdfBlob);
              }
            }
          } else if (exportSettings.preserveFormatting && fileData.htmlContent) {
            // Reconstructed layout for other formats (HTML, DOCX)
            const pdfBlob = await createPdfFromHtml(fileData.htmlContent, fileData.name, fileData.id);
            processedFiles.push({ name: `${baseName}_anon.pdf`, data: pdfBlob, isPdf: true, format: 'pdf' });
            zip.file(`${baseName}_anon.pdf`, pdfBlob);
          } else {
            // Simple text fallback
            const pdfBlob = await generateAnonymizedPdfFromText(fileData);
            processedFiles.push({ name: `${baseName}_anon.pdf`, data: pdfBlob, isPdf: true, format: 'pdf' });
            zip.file(`${baseName}_anon.pdf`, pdfBlob);
          }
        } else if (exportSettings.format === 'docx') {
          const docxBlob = await generateAnonymizedDocx(fileData);
          processedFiles.push({ name: `${baseName}_anon.docx`, data: docxBlob, isPdf: false, format: 'docx' });
          zip.file(`${baseName}_anon.docx`, docxBlob);
        } else {
          const anonymizedText = anonymizeText(fileData.content, entities);
          processedFiles.push({ name: `${baseName}_anon.txt`, data: anonymizedText, isPdf: false, format: 'txt' });
          zip.file(`${baseName}_anon.txt`, anonymizedText);
        }
      }

      if (processedFiles.length === 1) {
        const file = processedFiles[0];
        const blob = file.data instanceof Blob ? file.data : new Blob([file.data], { 
          type: file.format === 'pdf' ? 'application/pdf' : 
                file.format === 'docx' ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' : 
                'text/plain' 
        });
        saveAs(blob, file.name);
        showToast("Documento exportado com sucesso.", "success");
      } else if (processedFiles.length > 1) {
        const content = await zip.generateAsync({ type: 'blob' });
        saveAs(content, 'documentos_anonimizados.zip');
        showToast(`${processedFiles.length} documentos exportados em ZIP.`, "success");
      } else {
        showToast("Nenhum documento pronto para exportar.", "error");
      }
    } catch (error) {
      console.error("Erro ao exportar:", error);
      showToast("Erro ao exportar documentos.", "error");
    } finally {
      setIsProcessing(false);
    }
  };

  const generateAnonymizedPdfFromText = async (fileData: FileData): Promise<Blob> => {
    const doc = new jsPDF();
    if (exportSettings.sanitizeMetadata) {
      doc.setProperties({
        title: 'Documento Anonimizado',
        author: 'Anonimiza',
        subject: 'Documento Processado em Conformidade com RGPD',
        keywords: 'anonymized, rgpd, anonimizado',
        creator: 'Anonimiza Platform'
      });
    }
    const anonymizedText = anonymizeText(fileData.content, entities);
    
    const margin = 15;
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    const usableWidth = pageWidth - (margin * 2);
    
    const lines = doc.splitTextToSize(anonymizedText, usableWidth);
    
    let cursorY = 20;
    const lineHeight = 7;
    
    lines.forEach((line: string) => {
      if (cursorY + lineHeight > pageHeight - margin) {
        doc.addPage();
        cursorY = 20;
      }
      doc.text(line, margin, cursorY);
      cursorY += lineHeight;
    });
    
    return doc.output('blob');
  };

  const generateAnonymizedDocx = async (fileData: FileData): Promise<Blob> => {
    // Ensure we have HTML content for consistent processing (PDFs, TXT, etc.)
    let htmlToUse = fileData.htmlContent;
    if (!htmlToUse && fileData.content) {
      htmlToUse = fileData.content.split('\n')
        .map(line => line.trim() ? `<p>${line}</p>` : '')
        .filter(p => p !== '')
        .join('');
    }

    if (exportSettings.preserveFormatting) {
      if (fileData.name.toLowerCase().endsWith('.docx') && fileData.rawFile) {
        try {
          const zip = await JSZip.loadAsync(fileData.rawFile);
          
          const parser = new DOMParser();
          const serializer = new XMLSerializer();

          if (exportSettings.sanitizeMetadata) {
            // 1. Sanitize docProps/core.xml if it exists
            const coreFile = Object.keys(zip.files).find(name => name.toLowerCase() === 'docprops/core.xml');
            if (coreFile) {
              try {
                const content = await zip.file(coreFile).async('text');
                const doc = parser.parseFromString(content, 'application/xml');
                
                const dcCreator = doc.getElementsByTagName('dc:creator')[0];
                if (dcCreator) dcCreator.textContent = 'Anonimiza';
                
                const cpLastModifiedBy = doc.getElementsByTagName('cp:lastModifiedBy')[0];
                if (cpLastModifiedBy) cpLastModifiedBy.textContent = 'Anonimiza';
                
                const dcTitle = doc.getElementsByTagName('dc:title')[0];
                if (dcTitle) dcTitle.textContent = 'Documento Anonimizado';

                const dcSubject = doc.getElementsByTagName('dc:subject')[0];
                if (dcSubject) dcSubject.textContent = 'Documento Processado';
                
                const cpRevision = doc.getElementsByTagName('cp:revision')[0];
                if (cpRevision) cpRevision.textContent = '1';

                const dctermsCreated = doc.getElementsByTagName('dcterms:created')[0];
                if (dctermsCreated) dctermsCreated.textContent = '2026-06-08T12:00:00Z';

                const dctermsModified = doc.getElementsByTagName('dcterms:modified')[0];
                if (dctermsModified) dctermsModified.textContent = '2026-06-08T12:00:00Z';

                const serialized = serializer.serializeToString(doc);
                zip.file(coreFile, serialized);
              } catch (e) {
                console.error("Error sanitizing docProps/core.xml:", e);
              }
            }

            // 2. Sanitize docProps/app.xml if it exists
            const appFile = Object.keys(zip.files).find(name => name.toLowerCase() === 'docprops/app.xml');
            if (appFile) {
              try {
                const content = await zip.file(appFile).async('text');
                const doc = parser.parseFromString(content, 'application/xml');
                
                const company = doc.getElementsByTagName('Company')[0];
                if (company) company.textContent = 'Anonimiza';

                const manager = doc.getElementsByTagName('Manager')[0];
                if (manager) manager.textContent = '';

                const serialized = serializer.serializeToString(doc);
                zip.file(appFile, serialized);
              } catch (e) {
                console.error("Error sanitizing docProps/app.xml:", e);
              }
            }

            // 3. Remove custom file if it exists
            const customFile = Object.keys(zip.files).find(name => name.toLowerCase() === 'docprops/custom.xml');
            if (customFile) {
              zip.remove(customFile);
            }
          }

          const xmlFiles = Object.keys(zip.files).filter(name => 
            name.endsWith('.xml') && 
            (name.startsWith('word/') || name === '[Content_Types].xml')
          );
          
          const findRedactions = (fullText: string, entitiesList: typeof entities) => {
            const sortedEntities = [...entitiesList]
              .filter(e => (e.enabled || e.treated) && !e.ignored)
              .sort((a, b) => b.original.length - a.original.length);

            const redactionsList: { start: number, end: number, pseudonym: string }[] = [];
            const isRedacted = new Array(fullText.length).fill(false);
            const handled = new Set<string>();

            sortedEntities.forEach(entity => {
              const original = entity.original.trim();
              if (!original || handled.has(original.toLowerCase())) return;
              handled.add(original.toLowerCase());

              const escaped = original
                .replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
                .replace(/\s+/g, '\\s+');
              
              const wordChars = 'a-zA-ZÀ-ÿ0-9_';
              const regex = new RegExp(`(?<![${wordChars}])${escaped}(?![${wordChars}])`, 'gi');
              
              let match;
              while ((match = regex.exec(fullText)) !== null) {
                const start = match.index;
                const end = start + match[0].length;
                
                let overlap = false;
                for (let i = start; i < end; i++) {
                  if (isRedacted[i]) {
                    overlap = true;
                    break;
                  }
                }
                
                if (!overlap) {
                  for (let i = start; i < end; i++) {
                    isRedacted[i] = true;
                  }
                  redactionsList.push({ start, end, pseudonym: entity.pseudonym });
                }
              }
            });

            return redactionsList.sort((a, b) => a.start - b.start);
          };

          for (const xmlFile of xmlFiles) {
            const content = await zip.file(xmlFile).async('text');
            const xmlDoc = parser.parseFromString(content, 'application/xml');
            
            const paragraphs = xmlDoc.getElementsByTagName('w:p');
            const processedWTNodes = new Set<Element>();
            let modified = false;

            for (let pIdx = 0; pIdx < paragraphs.length; pIdx++) {
              const p = paragraphs[pIdx];
              const tNodes = Array.from(p.getElementsByTagName('w:t'));
              if (tNodes.length === 0) continue;

              const textSegments = tNodes.map(node => node.textContent || '');
              const fullParagraphText = textSegments.join('');
              
              if (!fullParagraphText.trim()) continue;

              const redactions = findRedactions(fullParagraphText, entities);
              if (redactions.length === 0) continue;

              // Map each character index to its node index and local offset
              interface CharPos {
                nodeIndex: number;
                localOffset: number;
              }
              const charPositions: CharPos[] = [];
              tNodes.forEach((node, nodeIdx) => {
                processedWTNodes.add(node);
                const text = node.textContent || '';
                for (let i = 0; i < text.length; i++) {
                  charPositions.push({
                    nodeIndex: nodeIdx,
                    localOffset: i
                  });
                }
              });

              // Apply redactions in reverse order to preserve indices for preceding edits
              for (let rIdx = redactions.length - 1; rIdx >= 0; rIdx--) {
                const { start, end, pseudonym } = redactions[rIdx];
                
                const startPos = charPositions[start];
                const endPos = charPositions[end - 1];
                if (!startPos || !endPos) continue;

                const nodeAIdx = startPos.nodeIndex;
                const offsetA = startPos.localOffset;
                const nodeBIdx = endPos.nodeIndex;
                const offsetB = endPos.localOffset;

                const nodeA = tNodes[nodeAIdx];
                const nodeB = tNodes[nodeBIdx];
                const textA = nodeA.textContent || '';
                const textB = nodeB.textContent || '';

                if (nodeAIdx === nodeBIdx) {
                  nodeA.textContent = textA.substring(0, offsetA) + pseudonym + textA.substring(offsetB + 1);
                } else {
                  nodeA.textContent = textA.substring(0, offsetA) + pseudonym;
                  for (let i = nodeAIdx + 1; i < nodeBIdx; i++) {
                    tNodes[i].textContent = '';
                  }
                  nodeB.textContent = textB.substring(offsetB + 1);
                }
                modified = true;
              }
            }

            // Fallback for any leftover w:t elements that weren't inside w:p paragraphs
            const allTNodes = xmlDoc.getElementsByTagName('w:t');
            for (let i = 0; i < allTNodes.length; i++) {
              const node = allTNodes[i] as Element;
              if (processedWTNodes.has(node)) continue;
              const originalText = node.textContent || '';
              if (originalText.trim()) {
                const anonymizedText = anonymizeText(originalText, entities);
                if (anonymizedText !== originalText) {
                  node.textContent = anonymizedText;
                  modified = true;
                }
              }
            }
            
            if (exportSettings.sanitizeMetadata) {
              const allElements = xmlDoc.getElementsByTagName('*');
              for (let i = 0; i < allElements.length; i++) {
                const el = allElements[i];
                if (el.hasAttribute('w:author')) {
                  el.setAttribute('w:author', 'Anonimiza');
                  modified = true;
                }
                if (el.hasAttribute('w:date')) {
                  el.setAttribute('w:date', '2026-06-08T12:00:00Z');
                  modified = true;
                }
              }

              if (xmlFile.endsWith('settings.xml')) {
                const trackRevisionsElements = xmlDoc.getElementsByTagName('w:trackRevisions');
                while (trackRevisionsElements.length > 0) {
                  trackRevisionsElements[0].parentNode?.removeChild(trackRevisionsElements[0]);
                  modified = true;
                }
              }
            }

            if (modified) {
              const serialized = serializer.serializeToString(xmlDoc);
              zip.file(xmlFile, serialized);
            }
          }
          
          return await zip.generateAsync({ type: 'blob' });
        } catch (err) {
          console.error("Error doing XML overlay for DOCX:", err);
          // Fallback to HTML representation on error
        }
      }

      if (htmlToUse) {
        return await createDocxFromHtml(htmlToUse, fileData.name, fileData.id);
      }
    }

    // Fallback to simple text-based DOCX (should rarely be hit now)
    const splitTextIntoRuns = (text: string, styles: { bold?: boolean, italics?: boolean, underline?: boolean, fontSize?: number, fontFamily?: string }) => {
      const activeEntities = entities.filter(e => e.enabled && !e.ignored);
      const sortedEntities = [...activeEntities].sort((a, b) => b.original.length - a.original.length);

      let segments: { text: string; isPseudonym: boolean; entityType?: string }[] = [{ text, isPseudonym: false }];

      for (const entity of sortedEntities) {
        const originalTerm = entity.original;
        if (!originalTerm || originalTerm.trim().length === 0) continue;

        const newSegments: typeof segments = [];
        for (const seg of segments) {
          if (seg.isPseudonym) {
            newSegments.push(seg);
            continue;
          }

          let currentText = seg.text;
          let index = currentText.toLowerCase().indexOf(originalTerm.toLowerCase());

          while (index !== -1) {
            if (index > 0) {
              newSegments.push({ text: currentText.substring(0, index), isPseudonym: false });
            }
            
            newSegments.push({ 
              text: entity.pseudonym, 
              isPseudonym: true,
              entityType: entity.type
            });

            currentText = currentText.substring(index + originalTerm.length);
            index = currentText.toLowerCase().indexOf(originalTerm.toLowerCase());
          }

          if (currentText) {
            newSegments.push({ text: currentText, isPseudonym: false });
          }
        }
        segments = newSegments;
      }

      return segments.map(seg => {
        if (seg.isPseudonym) {
          let highlightColor = "yellow";
          if (seg.entityType === 'NOME') highlightColor = "yellow";
          else if (seg.entityType === 'ADVOGADO') highlightColor = "cyan";
          else if (seg.entityType === 'JUIZ') highlightColor = "green";
          else if (seg.entityType === 'AUTOR') highlightColor = "magenta";
          else if (seg.entityType === 'COLETIVA') highlightColor = "lightGray";

          return new TextRun({
            text: seg.text,
            bold: true,
            highlight: highlightColor as any,
            italics: styles.italics,
            underline: styles.underline ? { type: UnderlineType.SINGLE } : undefined,
            size: styles.fontSize ? styles.fontSize * 2 : 24,
            font: styles.fontFamily || "Times New Roman"
          });
        } else {
          return new TextRun({
            text: seg.text,
            bold: styles.bold,
            italics: styles.italics,
            underline: styles.underline ? { type: UnderlineType.SINGLE } : undefined,
            size: styles.fontSize ? styles.fontSize * 2 : 24,
            font: styles.fontFamily || "Times New Roman"
          });
        }
      });
    };

    const docMeta = exportSettings.sanitizeMetadata ? {
      title: "Documento Anonimizado",
      subject: "Documento Processado em Conformidade com RGPD",
      creator: "Anonimiza Platform",
      description: "Documento processado para remoção de dados pessoais (PII)",
      lastModifiedBy: "Anonimiza Platform",
      revision: 1,
    } : {};

    const doc = new DocxDocument({
      ...docMeta,
      sections: [{
        properties: {},
        children: fileData.content.split('\n').map(line => {
          const trimmed = line.trim();
          if (!trimmed) return null;
          
          // Basic header check even in fallback
          if (exportSettings.removeHeadersFooters) {
            const lower = trimmed.toLowerCase();
            if (lower.includes('processo:') || lower.includes('referência:') || lower.includes('página') || lower.includes('tribunal')) {
               if (trimmed.length < 100) return null;
            }
          }

          const lineRuns = splitTextIntoRuns(line, { fontSize: 12, fontFamily: "Times New Roman" });

          return new Paragraph({
            children: lineRuns,
            alignment: line.length > 50 ? AlignmentType.JUSTIFIED : undefined,
          });
        }).filter(p => p !== null) as Paragraph[],
      }],
    });

    return await Packer.toBlob(doc);
  };

  const prepareAnonymizedHtml = (html: string, fileId: string, mergeParagraphs: boolean = false): { domDoc: Document, footnotes: Record<string, string> } => {
    const parser = new DOMParser();
    const domDoc = parser.parseFromString(html, 'text/html');
    const footnotes: Record<string, string> = {};

    // Filter entities for this file
    const fileEntities = entities.filter(e => e.fileIds?.includes(fileId) && e.enabled && !e.ignored);
    const headerEntities = fileEntities.filter(e => e.type === 'HEADER');

    // Extract footnotes if they exist
    const footnoteElements = domDoc.querySelectorAll('li[id*="footnote"], div[id*="footnote"], p[id*="footnote"], [class*="footnote"], .pdf-footnote');
    footnoteElements.forEach(el => {
      // Skip reference links and container divs
      const elementId = el.id.toLowerCase();
      const elementClass = el.className.toLowerCase();
      if (elementId.includes('ref') || elementClass.includes('ref')) return;
      if (el.tagName.toLowerCase() === 'div' && el.querySelector('ol, ul')) return;
      if (el.tagName.toLowerCase() === 'ol' || el.tagName.toLowerCase() === 'ul') return;

      const idMatch = el.id.match(/footnote-(\d+)/) || el.className.match(/footnote-(\d+)/) || el.textContent?.match(/^(\d+)[\.\s]/);
      const id = idMatch ? idMatch[1] : (el.id && !el.id.includes('footnotes') ? el.id : null);
      
      if (!id) return;

      // Remove back-reference arrows and links
      el.querySelectorAll('a').forEach(a => {
        const aText = a.textContent || '';
        if (aText.includes('↑') || a.getAttribute('href')?.includes('ref')) a.remove();
      });

      const cleanText = el.textContent?.replace(/^\d+[\.\s]*/, '').trim() || '';
      if (cleanText) {
        if (!footnotes[id]) {
          footnotes[id] = cleanText;
        } else if (footnotes[id] !== cleanText && !footnotes[id].includes(cleanText)) {
          footnotes[id] += ' ' + cleanText;
        }
      }
      
      if (exportSettings.footnotesAtEnd) {
        el.remove();
      }
    });

    const dynamicHeaders = Object.entries(globalKnowledge)
      .filter(([_, type]) => type === 'HEADER')
      .map(([text]) => text.toLowerCase().trim().replace(/\s+/g, ' '));

    const headerPatterns = [
      ...dynamicHeaders.map(h => {
        if (h.includes('processo') || h.includes('referência')) {
          const prefix = h.split(':')[0].trim();
          return new RegExp('^' + prefix + ':?', 'i');
        }
        const escaped = h.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\\ /g, '\\s+');
        return new RegExp(escaped, 'i');
      }),
      /^assinado em \d{2}-\d{2}-\d{4}, por/i,
      /^assinado por:?/i,
      /^processo:?\s+/i,
      /^referência:?\s+/i,
      /tribunal da relação/i,
      /secção social/i,
      /campo mártires da pátria/i,
      /4099-012 porto/i,
      /^telef:?\s+\d+/i,
      /^fax:?\s+\d+/i,
      /^mail:?\s+.*@.*/i,
      /^página\s+\d+/i,
      /^pág\.\s+\d+/i,
      /^folha\s+\d+/i,
      /tribunal judicial/i,
      /juízo do trabalho/i,
      /comarca de/i,
      /instância central/i,
      /instância local/i,
      /secção cível/i,
      /secção criminal/i,
      /procuradoria-geral/i,
      /ministério público/i,
      /^\d+$/
    ];

    const normalizeForMatch = (s: string) => s.toLowerCase().replace(/\s+/g, ' ').trim();

    const processElement = (element: Element) => {
      const tagName = element.tagName.toLowerCase();
      
      if (tagName === 'div' || tagName === 'section' || tagName === 'article') {
        Array.from(element.children).forEach(child => processElement(child));
        return;
      }

      if (tagName === 'p') {
        const originalText = element.textContent || '';
        const normalizedText = normalizeForMatch(originalText);
        
        const isMarkedHeader = dynamicHeaders.some(h => 
          normalizedText.includes(h) || h.includes(normalizedText)
        );
        if (isMarkedHeader) {
          element.remove();
          return;
        }

        if (exportSettings.removeHeadersFooters) {
          const yPos = parseFloat(element.getAttribute('data-y') || '0.5');
          const isAtTop = yPos < 0.20;
          const isAtBottom = yPos > 0.80;
          
          const isHeader = isAtTop && (
            headerPatterns.some(pattern => pattern.test(normalizedText)) || 
            headerEntities.some(e => normalizedText.includes(normalizeForMatch(e.original))) ||
            (normalizedText.length < 80 && (
              normalizedText.includes('porto') || 
              normalizedText.includes('tribunal') || 
              normalizedText.includes('relação') ||
              normalizedText.includes('assinado em') ||
              normalizedText.includes('referência') ||
              normalizedText.includes('processo')
            ))
          );

          const isFooter = isAtBottom && (
            headerEntities.some(e => normalizedText.includes(normalizeForMatch(e.original))) ||
            /^\d+$/.test(normalizedText) ||
            normalizedText.includes('página') ||
            normalizedText.includes('pág.') ||
            normalizedText.includes('folha')
          );
          
          if (isHeader || isFooter) {
            element.remove();
            return;
          }
        }
      }
    };

    Array.from(domDoc.body.children).forEach(child => processElement(child));
    
    if (mergeParagraphs) {
      // Post-process: Merge paragraphs that don't end with punctuation
      // Following the rule: "o word só pode fazer paragrafo depois de um ponto, um ponto de exclamação, um ponto de interrogação ou dois pontos."
      const paragraphs = Array.from(domDoc.querySelectorAll('p'));
      for (let i = 0; i < paragraphs.length - 1; i++) {
        const p = paragraphs[i];
        const nextP = paragraphs[i + 1];
        
        // Skip if either is a footnote reference or special element
        if (p.className.includes('footnote') || nextP.className.includes('footnote')) continue;
        
        const text = p.textContent?.trim() || '';
        // Check if it ends with . ! ? or : (allowing for closing quotes/parens)
        if (text && !/[.!?:][\s"”'’)]*$/.test(text)) {
          // Move all children from nextP to p, adding a space if needed
          p.appendChild(domDoc.createTextNode(' '));
          while (nextP.firstChild) {
            p.appendChild(nextP.firstChild);
          }
          nextP.remove();
          paragraphs.splice(i + 1, 1);
          i--; // Check the merged paragraph again
        }
      }
    }
    
    return { domDoc, footnotes };
  };

  const createPdfFromHtml = async (html: string, originalName: string, fileId: string): Promise<Blob> => {
    const { domDoc: anonymizedDoc, footnotes } = prepareAnonymizedHtml(html, fileId, false);
    
    // Create a temporary container for jsPDF to render
    const container = document.createElement('div');
    container.id = 'pdf-export-container';
    container.style.width = '800px';
    container.style.minWidth = '800px';
    container.style.maxWidth = '800px';
    container.style.padding = '60px';
    container.style.boxSizing = 'border-box';
    container.style.fontFamily = 'Times New Roman, serif';
    container.style.fontSize = '11.5pt';
    container.style.lineHeight = '1.5';
    container.style.backgroundColor = 'white';
    container.style.color = 'black';
    container.style.textAlign = 'justify';

    const styleSheet = document.createElement('style');
    styleSheet.textContent = `
      #pdf-export-container p, #pdf-export-container div, #pdf-export-container section {
        text-align: justify !important;
        margin-bottom: 12pt !important;
        line-height: 1.5 !important;
        font-family: "Times New Roman", Times, serif !important;
      }
      #pdf-export-container {
        font-family: "Times New Roman", Times, serif !important;
        background-color: white !important;
        color: black !important;
      }
    `;
    container.appendChild(styleSheet);
    
    // Clone the body content
    const bodyClone = anonymizedDoc.body.cloneNode(true) as HTMLElement;
    
    // Robust paragraph/block-level anonymization that handles names/terms split across multiple text nodes (spans, formattings)
    const findRedactions = (fullText: string, entitiesList: typeof entities) => {
      const sortedEntities = [...entitiesList]
        .filter(e => (e.enabled || e.treated) && !e.ignored)
        .sort((a, b) => b.original.length - a.original.length);

      const redactionsList: { start: number, end: number, pseudonym: string }[] = [];
      const isRedacted = new Array(fullText.length).fill(false);
      const handled = new Set<string>();

      sortedEntities.forEach(entity => {
        const original = entity.original.trim();
        if (!original || handled.has(original.toLowerCase())) return;
        handled.add(original.toLowerCase());

        const escaped = original
          .replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
          .replace(/\s+/g, '\\s+');
        
        const wordChars = 'a-zA-ZÀ-ÿ0-9_';
        const regex = new RegExp(`(?<![${wordChars}])${escaped}(?![${wordChars}])`, 'gi');
        
        let match;
        while ((match = regex.exec(fullText)) !== null) {
          const start = match.index;
          const end = start + match[0].length;
          
          let overlap = false;
          for (let i = start; i < end; i++) {
            if (isRedacted[i]) {
              overlap = true;
              break;
            }
          }
          
          if (!overlap) {
            for (let i = start; i < end; i++) {
              isRedacted[i] = true;
            }
            redactionsList.push({ start, end, pseudonym: entity.pseudonym });
          }
        }
      });

      return redactionsList.sort((a, b) => a.start - b.start);
    };

    const anonymizeDomElementContent = (element: Element) => {
      const textNodes: Text[] = [];
      const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT, null);
      let tNode;
      while (tNode = walker.nextNode()) {
        textNodes.push(tNode as Text);
      }
      
      if (textNodes.length === 0) return;
      
      const textSegments = textNodes.map(node => node.textContent || '');
      const fullText = textSegments.join('');
      if (!fullText.trim()) return;
      
      const redactions = findRedactions(fullText, entities);
      if (redactions.length === 0) {
        // Fallback: anonymize each node individually
        textNodes.forEach(node => {
          node.textContent = anonymizeText(node.textContent || '', entities);
        });
        return;
      }
      
      // Map character indexes to textNode and offset
      interface CharPos {
        nodeIndex: number;
        localOffset: number;
      }
      const charPositions: CharPos[] = [];
      textNodes.forEach((node, nodeIdx) => {
        const text = node.textContent || '';
        for (let i = 0; i < text.length; i++) {
          charPositions.push({
            nodeIndex: nodeIdx,
            localOffset: i
          });
        }
      });
      
      // Apply redactions in reverse order
      for (let rIdx = redactions.length - 1; rIdx >= 0; rIdx--) {
        const { start, end, pseudonym } = redactions[rIdx];
        
        const startPos = charPositions[start];
        const endPos = charPositions[end - 1];
        if (!startPos || !endPos) continue;
        
        const nodeAIdx = startPos.nodeIndex;
        const offsetA = startPos.localOffset;
        const nodeBIdx = endPos.nodeIndex;
        const offsetB = endPos.localOffset;
        
        const nodeA = textNodes[nodeAIdx];
        const nodeB = textNodes[nodeBIdx];
        const textA = nodeA.textContent || '';
        const textB = nodeB.textContent || '';
        
        if (nodeAIdx === nodeBIdx) {
          nodeA.textContent = textA.substring(0, offsetA) + pseudonym + textA.substring(offsetB + 1);
        } else {
          nodeA.textContent = textA.substring(0, offsetA) + pseudonym;
          for (let i = nodeAIdx + 1; i < nodeBIdx; i++) {
            textNodes[i].textContent = '';
          }
          nodeB.textContent = textB.substring(offsetB + 1);
        }
      }

      // Finally, run individual anonymizeText on any textNodes that were modified but might have leftover text with other entities
      textNodes.forEach(node => {
        if (node.textContent && node.textContent.trim()) {
          node.textContent = anonymizeText(node.textContent, entities);
        }
      });
    };

    const processElementForAnonymization = (element: Element) => {
      const tagName = element.tagName.toLowerCase();
      const isContainer = ['body', 'div', 'section', 'article', 'table', 'tbody', 'tr', 'ul', 'ol', 'blockquote'].includes(tagName);
      
      let hasContainerChildren = false;
      for (let i = 0; i < element.children.length; i++) {
        const childTag = element.children[i].tagName.toLowerCase();
        if (['p', 'li', 'td', 'th', 'div', 'section', 'article', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6'].includes(childTag)) {
          hasContainerChildren = true;
          break;
        }
      }
      
      if (isContainer && hasContainerChildren) {
        Array.from(element.children).forEach(child => processElementForAnonymization(child));
      } else {
        anonymizeDomElementContent(element);
      }
    };

    // Run robust block-level anonymization across bodyClone elements
    processElementForAnonymization(bodyClone);
    
    // Append children to avoid nested bodies
    while (bodyClone.firstChild) {
      container.appendChild(bodyClone.firstChild);
    }

    // Add footnotes at the end
    if (exportSettings.footnotesAtEnd && Object.keys(footnotes).length > 0) {
      const hr = document.createElement('hr');
      hr.style.marginTop = '20pt';
      container.appendChild(hr);
      
      const title = document.createElement('h2');
      title.textContent = 'NOTAS DE RODAPÉ';
      title.style.textAlign = 'center';
      title.style.fontSize = '14pt';
      container.appendChild(title);
      
      Object.entries(footnotes).forEach(([id, text]) => {
        const p = document.createElement('p');
        p.style.fontSize = '10pt';
        p.style.textAlign = 'justify';
        p.innerHTML = `<strong>${id}.</strong> ${anonymizeText(text, entities)}`;
        container.appendChild(p);
      });
    }

    // Append to body temporarily but keep it invisible to the user using absolute/opacity so html2canvas renders coordinates correctly
    container.style.position = 'absolute';
    container.style.left = '0';
    container.style.top = '0';
    container.style.opacity = '0.01';
    container.style.pointerEvents = 'none';
    container.style.zIndex = '-9999';
    container.style.background = 'white';
    document.body.appendChild(container);

    const pdf = new jsPDF('p', 'mm', 'a4');
    if (exportSettings.sanitizeMetadata) {
      pdf.setProperties({
        title: 'Documento Anonimizado',
        author: 'Anonimiza',
        subject: 'Documento Processado em Conformidade com RGPD',
        keywords: 'anonymized, rgpd, anonimizado',
        creator: 'Anonimiza Platform'
      });
    }
    
    try {
      // Small delay for layout to stabilize
      await new Promise(r => setTimeout(r, 100));

      await new Promise<void>((resolve, reject) => {
        pdf.html(container, {
          callback: function (doc) {
            resolve();
          },
          x: 0,
          y: 0,
          width: 210, // A4 width
          windowWidth: 800,
          autoPaging: 'text',
          margin: [10, 10, 10, 10],
          html2canvas: {
            scale: 1, 
            useCORS: true,
            logging: false,
            allowTaint: true,
            backgroundColor: '#ffffff'
          }
        }).catch(err => {
          console.error("jsPDF.html rendering error:", err);
          reject(err);
        });
      });
      
      if (container.parentNode) document.body.removeChild(container);
      return pdf.output('blob');
    } catch (err) {
      console.error("Error generating PDF from HTML:", err);
      if (container.parentNode) document.body.removeChild(container);
      
      // Fallback: try simple text-based output if HTML rendering failed
      const textForFallback = anonymizeText(html.replace(/<[^>]*>?/gm, ' '), entities);
      const fallbackPdf = new jsPDF();
      if (exportSettings.sanitizeMetadata) {
        fallbackPdf.setProperties({
          title: 'Documento Anonimizado',
          author: 'Anonimiza',
          subject: 'Documento Processado em Conformidade com RGPD',
          keywords: 'anonymized, rgpd, anonimizado',
          creator: 'Anonimiza Platform'
        });
      }
      const splitText = fallbackPdf.splitTextToSize(textForFallback, 180);
      fallbackPdf.text(splitText, 15, 20);
      return fallbackPdf.output('blob');
    }
  };

  const createDocxFromHtml = async (html: string, originalName: string, fileId: string): Promise<Blob> => {
    const { domDoc: doc, footnotes } = prepareAnonymizedHtml(html, fileId, true);
    const children: any[] = [];

    const dynamicHeaders = Object.entries(globalKnowledge)
      .filter(([_, type]) => type === 'HEADER')
      .map(([text]) => text.toLowerCase().trim().replace(/\s+/g, ' '));

    const headerPatterns = [
      ...dynamicHeaders.map(h => {
        if (h.includes('processo') || h.includes('referência')) {
          const prefix = h.split(':')[0].trim();
          return new RegExp('^' + prefix + ':?', 'i');
        }
        const escaped = h.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\\ /g, '\\s+');
        return new RegExp(escaped, 'i');
      }),
      /^assinado em \d{2}-\d{2}-\d{4}, por/i,
      /^assinado por:?/i,
      /^processo:?\s+/i,
      /^referência:?\s+/i,
      /tribunal da relação/i,
      /secção social/i,
      /campo mártires da pátria/i,
      /4099-012 porto/i,
      /^telef:?\s+\d+/i,
      /^fax:?\s+\d+/i,
      /^mail:?\s+.*@.*/i,
      /^página\s+\d+/i,
      /^pág\.\s+\d+/i,
      /^folha\s+\d+/i,
      /tribunal judicial/i,
      /juízo do trabalho/i,
      /comarca de/i,
      /instância central/i,
      /instância local/i,
      /secção cível/i,
      /secção criminal/i,
      /procuradoria-geral/i,
      /ministério público/i,
      /^\d+$/
    ];

    const normalizeForMatch = (s: string) => s.toLowerCase().replace(/\s+/g, ' ').trim();

    // Process body elements
    let currentParagraphRuns: any[] = [];
    
    const isTerminalPunctuation = (text: string) => {
      const trimmed = text.trim();
      if (!trimmed) return true;
      // Also check for common Portuguese abbreviations or numbers that shouldn't trigger a merge
      if (trimmed.endsWith('Art.') || trimmed.endsWith('n.º') || trimmed.endsWith('Pág.')) return false;
      const lastChar = trimmed[trimmed.length - 1];
      return ['.', '!', '?', ':'].includes(lastChar);
    };

    const flushParagraph = (alignment = AlignmentType.BOTH, isSpecial: boolean = false) => {
      if (currentParagraphRuns.length > 0) {
        children.push(new Paragraph({ 
          children: [...currentParagraphRuns],
          alignment: alignment,
          spacing: { 
            before: 120, 
            after: 120, 
            line: 360, 
            lineRule: LineRuleType.AUTO 
          },
          indent: isSpecial ? undefined : {
            firstLine: 567, // 1 cm indent (approx 567 twips)
          }
        }));
        currentParagraphRuns = [];
      }
    };

    const processElement = (element: Element) => {
      const tagName = element.tagName.toLowerCase();
      
      if (tagName === 'div' || tagName === 'section' || tagName === 'article') {
        Array.from(element.children).forEach(child => processElement(child));
        return;
      }

      if (tagName === 'p') {
        const originalText = element.textContent || '';
        const normalizedText = normalizeForMatch(originalText);
        
        // Handle standalone asterisks or separators
        if (normalizedText === '*' || normalizedText === '***' || normalizedText === '---') {
          flushParagraph(); // Flush accumulated text
          children.push(new Paragraph({
            children: [new TextRun({ text: originalText, bold: true })],
            alignment: AlignmentType.CENTER,
            spacing: { before: 240, after: 240 }
          }));
          return;
        }

        // Detect citation style (often has large left margin or specific class)
        const style = element.getAttribute('style') || '';
        const isCitation = element.className?.toLowerCase().includes('citation') || 
                          style.includes('margin-left') || 
                          style.includes('padding-left');

        // Header removal logic
        const isMarkedHeader = dynamicHeaders.some(h => 
          normalizedText.includes(h) || h.includes(normalizedText)
        );
        if (isMarkedHeader) return;

        if (exportSettings.removeHeadersFooters) {
          const isHeader = headerPatterns.some(pattern => pattern.test(normalizedText)) || 
                          (normalizedText.length < 80 && (
                            normalizedText.includes('porto') || 
                            normalizedText.includes('tribunal') || 
                            normalizedText.includes('relação') ||
                            normalizedText.includes('assinado em')
                          ));
          if (isHeader) return;
        }

        const runs: any[] = [];
        
        // Get paragraph-level styles
        const pStyle = element.getAttribute('style') || '';
        const pFontSizeMatch = pStyle.match(/font-size:\s*([\d.]+)pt/);
        const pFontSize = pFontSizeMatch ? parseFloat(pFontSizeMatch[1]) : undefined;
        const pFontFamilyMatch = pStyle.match(/font-family:\s*'([^']+)'/);
        const pFontFamily = pFontFamilyMatch ? pFontFamilyMatch[1] : undefined;

        // Helper function inside processNode to split text and highlight entities
        const splitTextIntoRuns = (text: string, currentStyles: any) => {
          const activeEntities = entities.filter(e => e.enabled && !e.ignored);
          const sortedEntities = [...activeEntities].sort((a, b) => b.original.length - a.original.length);

          let segments: { text: string; isPseudonym: boolean; entityType?: string }[] = [{ text, isPseudonym: false }];

          for (const entity of sortedEntities) {
            const originalTerm = entity.original;
            if (!originalTerm || originalTerm.trim().length === 0) continue;

            const newSegments: typeof segments = [];
            for (const seg of segments) {
              if (seg.isPseudonym) {
                newSegments.push(seg);
                continue;
              }

              let currentText = seg.text;
              let index = currentText.toLowerCase().indexOf(originalTerm.toLowerCase());

              while (index !== -1) {
                if (index > 0) {
                  newSegments.push({ text: currentText.substring(0, index), isPseudonym: false });
                }
                
                newSegments.push({ 
                  text: entity.pseudonym, 
                  isPseudonym: true,
                  entityType: entity.type
                });

                currentText = currentText.substring(index + originalTerm.length);
                index = currentText.toLowerCase().indexOf(originalTerm.toLowerCase());
              }

              if (currentText) {
                newSegments.push({ text: currentText, isPseudonym: false });
              }
            }
            segments = newSegments;
          }

          return segments.map(seg => {
            if (seg.isPseudonym) {
              let highlightColor = "yellow";
              if (seg.entityType === 'NOME') highlightColor = "yellow";
              else if (seg.entityType === 'ADVOGADO') highlightColor = "cyan";
              else if (seg.entityType === 'JUIZ') highlightColor = "green";
              else if (seg.entityType === 'AUTOR') highlightColor = "magenta";
              else if (seg.entityType === 'COLETIVA') highlightColor = "lightGray";

              return new TextRun({
                text: seg.text,
                bold: true,
                highlight: highlightColor as any,
                italics: currentStyles.italics,
                underline: currentStyles.underline ? { type: UnderlineType.SINGLE } : undefined,
                size: currentStyles.fontSize ? currentStyles.fontSize * 2 : pFontSize ? pFontSize * 2 : 24,
                font: currentStyles.fontFamily || pFontFamily || "Times New Roman"
              });
            } else {
              return new TextRun({
                text: seg.text,
                bold: currentStyles.bold,
                italics: currentStyles.italics,
                underline: currentStyles.underline ? { type: UnderlineType.SINGLE } : undefined,
                size: currentStyles.fontSize ? currentStyles.fontSize * 2 : pFontSize ? pFontSize * 2 : 24,
                font: currentStyles.fontFamily || pFontFamily || "Times New Roman"
              });
            }
          });
        };

        // Recursive function to collect styles and text
        const processNode = (node: Node, styles: { bold?: boolean, italics?: boolean, underline?: boolean, fontSize?: number, fontFamily?: string }) => {
          if (node.nodeType === Node.TEXT_NODE) {
            const content = node.textContent || '';
            if (content) {
              const nodeRuns = splitTextIntoRuns(content, styles);
              runs.push(...nodeRuns);
            }
          } else if (node instanceof Element) {
            const subTagName = node.tagName.toLowerCase();
            const newStyles = { ...styles };
            
            if (subTagName === 'strong' || subTagName === 'b' || 
                (node as HTMLElement).style?.fontWeight === 'bold' || 
                (node as HTMLElement).style?.fontWeight === '700' ||
                (node as HTMLElement).style?.fontWeight === '800' ||
                (node as HTMLElement).style?.fontWeight === '900') newStyles.bold = true;
            
            if (subTagName === 'em' || subTagName === 'i' || 
                (node as HTMLElement).style?.fontStyle === 'italic') newStyles.italics = true;
            
            if (subTagName === 'u' || subTagName === 'ins' ||
                (node as HTMLElement).style?.textDecoration === 'underline' || 
                (node as HTMLElement).style?.textDecorationLine === 'underline' ||
                (node as HTMLElement).style?.borderBottom?.includes('solid')) newStyles.underline = true;
            
            const nodeStyle = node.getAttribute('style') || '';
            const nodeFontSizeMatch = nodeStyle.match(/font-size:\s*([\d.]+)pt/);
            if (nodeFontSizeMatch) newStyles.fontSize = parseFloat(nodeFontSizeMatch[1]);
            
            const nodeFontFamilyMatch = nodeStyle.match(/font-family:\s*'([^']+)'/);
            if (nodeFontFamilyMatch) newStyles.fontFamily = nodeFontFamilyMatch[1];

            if ((subTagName === 'a' || subTagName === 'sup' || subTagName === 'span') && 
                (node.getAttribute('id')?.includes('footnote') || 
                 node.getAttribute('href')?.includes('footnote') || 
                 (node as HTMLElement).className?.includes('footnote'))) {
              const idMatch = (node.getAttribute('id') || node.getAttribute('href') || (node as HTMLElement).className || '').match(/footnote-(\d+)/);
              const refId = idMatch ? idMatch[1] : null;
              if (refId) {
                if (exportSettings.footnotesAtEnd) {
                  runs.push(new TextRun({
                    text: `[${refId}]`,
                    superScript: true,
                    bold: true
                  }));
                } else {
                  runs.push(new FootnoteReferenceRun(parseInt(refId)));
                }
              }
              return;
            }

            node.childNodes.forEach(child => processNode(child, newStyles));
          }
        };

        element.childNodes.forEach(node => processNode(node, {}));

        let pAlignType: any = AlignmentType.BOTH;
        if (pStyle.includes('text-align: center') || pStyle.includes('text-align:center')) {
          pAlignType = AlignmentType.CENTER;
        } else if (pStyle.includes('text-align: right') || pStyle.includes('text-align:right')) {
          pAlignType = AlignmentType.RIGHT;
        } else if (pStyle.includes('text-align: left') || pStyle.includes('text-align:left')) {
          pAlignType = AlignmentType.LEFT;
        }

        if (runs.length > 0) {
          if (isCitation) {
            flushParagraph(); // Flush previous
            children.push(new Paragraph({
              children: runs,
              alignment: AlignmentType.BOTH,
              indent: { 
                left: 720,
                firstLine: 567 // 1 cm first-line indent
              },
              spacing: { before: 240, after: 240, line: 360 }
            }));
          } else {
            if (currentParagraphRuns.length > 0) {
              currentParagraphRuns.push(new TextRun(" "));
            }
            currentParagraphRuns.push(...runs);
            
            // Flush immediately with specified alignment to preserve centering, tables, headers, and right alignments.
            flushParagraph(pAlignType);
          }
        }
      } else if (tagName === 'table') {
        flushParagraph(); // Ensure previous paragraph is flushed before table
        const rows: TableRow[] = [];
        element.querySelectorAll('tr').forEach(tr => {
          const cells: TableCell[] = [];
          tr.querySelectorAll('td, th').forEach(td => {
            cells.push(new TableCell({
              children: [new Paragraph({
                children: [new TextRun(anonymizeText(td.textContent || '', entities))],
                alignment: AlignmentType.CENTER
              })],
              width: { size: 100 / tr.querySelectorAll('td, th').length, type: WidthType.PERCENTAGE }
            }));
          });
          rows.push(new TableRow({ children: cells }));
        });
        children.push(new Table({ rows, width: { size: 100, type: WidthType.PERCENTAGE } }));
      } else if (tagName.startsWith('h') && tagName.length === 2) {
        flushParagraph(); // Ensure previous paragraph is flushed before heading
        const level = parseInt(tagName[1]);
        children.push(new Paragraph({
          text: anonymizeText(element.textContent || '', entities),
          heading: level === 1 ? HeadingLevel.HEADING_1 : level === 2 ? HeadingLevel.HEADING_2 : HeadingLevel.HEADING_3,
          alignment: AlignmentType.BOTH,
          indent: {
            firstLine: 567, // 1 cm indent
          },
          spacing: { before: 240, after: 120 }
        }));
      }
    };

    Array.from(doc.body.children).forEach(child => {
      processElement(child);
    });
    flushParagraph(); // Final flush

    // Add footnotes at the end if requested
    if (exportSettings.footnotesAtEnd && Object.keys(footnotes).length > 0) {
      children.push(new Paragraph({ text: '', spacing: { before: 400 } }));
      children.push(new Paragraph({ 
        text: 'NOTAS DE RODAPÉ', 
        heading: HeadingLevel.HEADING_2,
        alignment: AlignmentType.CENTER
      }));
      
      Object.entries(footnotes).forEach(([id, text]) => {
        children.push(new Paragraph({
          children: [
            new TextRun({ text: `${id}. `, bold: true }),
            new TextRun(anonymizeText(text, entities))
          ],
          alignment: AlignmentType.BOTH,
          spacing: { before: 60, after: 60 }
        }));
      });
    }

    const docMeta = exportSettings.sanitizeMetadata ? {
      title: "Documento Anonimizado",
      subject: "Documento Processado em Conformidade com RGPD",
      creator: "Anonimiza Platform",
      description: "Documento processado para remoção de dados pessoais (PII)",
      lastModifiedBy: "Anonimiza Platform",
      revision: 1,
    } : {};

    const docx = new DocxDocument({
      ...docMeta,
      styles: {
        default: {
          document: {
            run: {
              font: "Times New Roman",
              size: 24, // 12pt
            },
            paragraph: {
              alignment: AlignmentType.BOTH,
              spacing: { line: 360, before: 120, after: 120 }
            }
          }
        }
      },
      sections: [{
        properties: {
          page: {
            margin: {
              top: 1440, // 1 inch
              right: 1440,
              bottom: 1440,
              left: 1440,
            }
          }
        },
        children: children,
      }],
    });

    return await Packer.toBlob(docx);
  };

  const dataURLToUint8Array = (dataURL: string) => {
    const base64 = dataURL.split(',')[1];
    const binaryString = window.atob(base64);
    const len = binaryString.length;
    const bytes = new Uint8Array(len);
    for (let i = 0; i < len; i++) {
      bytes[i] = binaryString.charCodeAt(i);
    }
    return bytes;
  };

  const exportAnonymizedPDFBytes = async (fileData: FileData): Promise<Uint8Array> => {
    if (!fileData.rawFile || typeof fileData.rawFile.arrayBuffer !== 'function') {
      throw new Error(`Original binary file data is missing for ${fileData.name}. This typically happens after a page refresh because File objects are not persisted in localStorage.`);
    }
    const arrayBuffer = await fileData.rawFile.arrayBuffer();
    
    // Load with pdfjs for rendering
    const loadingTask = pdfjs.getDocument({ 
      data: arrayBuffer,
      standardFontDataUrl: `https://unpkg.com/pdfjs-dist@${pdfjs.version}/standard_fonts/`
    });
    const pdf = await loadingTask.promise;
    
    // Create new PDF with pdf-lib
    const outPdfDoc = await PDFDocument.create();

    if (exportSettings.sanitizeMetadata) {
      outPdfDoc.setTitle("Documento Anonimizado");
      outPdfDoc.setAuthor("Anonimiza");
      outPdfDoc.setSubject("Documento Processado em Conformidade com RGPD");
      outPdfDoc.setCreator("Anonimiza Platform");
      outPdfDoc.setProducer("Anonimiza Platform");
      outPdfDoc.setCreationDate(new Date('2026-06-08T12:00:00Z'));
      outPdfDoc.setModificationDate(new Date('2026-06-08T12:00:00Z'));
    }

    // Filter entities for this file (include all types of entities)
    const fileEntities = entities.filter(e => 
      e.fileIds?.includes(fileData.id) && 
      e.enabled && 
      !e.ignored
    );

    // Sort entities by length descending
    const sortedEntities = [...fileEntities].sort((a, b) => b.original.length - a.original.length);

    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i);
      const scale = 2.0; // High resolution for quality
      const viewport = page.getViewport({ scale });
      
      const canvas = document.createElement('canvas');
      canvas.style.position = 'fixed';
      canvas.style.left = '-10000px';
      canvas.style.top = '0';
      canvas.style.opacity = '0';
      document.body.appendChild(canvas);

      const context = canvas.getContext('2d')!;
      canvas.height = viewport.height;
      canvas.width = viewport.width;

      // Fill with white background explicitly
      context.fillStyle = 'white';
      context.fillRect(0, 0, canvas.width, canvas.height);

      // Render original page to canvas (matching the working viewer layout)
      await (page as any).render({ 
        canvasContext: context, 
        viewport: viewport
      }).promise;

      // Get text content to find coordinates for redaction
      const textContent = await page.getTextContent();
      
      // Draw redactions directly on the canvas
      for (const item of textContent.items as any[]) {
        const text = item.str;
        if (!text.trim() || text.trim().length < 2) continue;

        // Check if the center of this text item falls inside any of the marked header/footer zones of page i (or global)
        const [, , , scaleY, x, y] = item.transform;
        const pHeight = viewport.viewBox ? viewport.viewBox[3] : (viewport.height / viewport.scale);
        const pWidth = viewport.viewBox ? viewport.viewBox[2] : (viewport.width / viewport.scale);
        const itemYPercent = (pHeight - y) / pHeight;
        const itemXPercent = x / pWidth;

        const isInSkippedZone = [...headerZones, ...footerZones].some(zone => {
          if (zone.pageNum === undefined || zone.pageNum === i) {
            return itemYPercent >= zone.top && itemYPercent <= zone.bottom &&
                   itemXPercent >= zone.left && itemXPercent <= zone.right;
          }
          return false;
        });

        if (isInSkippedZone) continue;

        const matchesInItem = sortedEntities
          .filter(entity => text.toLowerCase().includes(entity.original.toLowerCase()))
          .sort((a, b) => text.toLowerCase().indexOf(a.original.toLowerCase()) - text.toLowerCase().indexOf(b.original.toLowerCase()));

        if (matchesInItem.length > 0) {
          const [viewX, viewY] = viewport.convertToViewportPoint(item.transform[4], item.transform[5]);
          const fontHeight = Math.sqrt(item.transform[2] * item.transform[2] + item.transform[3] * item.transform[3]);
          const itemWidthInCanvas = item.width * viewport.scale;
          const fontHeightInCanvas = fontHeight * viewport.scale;

          matchesInItem.forEach(matchingEntity => {
            const original = matchingEntity.original;
            const startIndex = text.toLowerCase().indexOf(original.toLowerCase());
            if (startIndex === -1) return;

            const charWidth = itemWidthInCanvas / text.length;

            const cX = viewX + (startIndex * charWidth);
            const cY = viewY - fontHeightInCanvas;
            const cWidth = original.length * charWidth;
            const cHeight = fontHeightInCanvas * 1.2;

            const colorInfo = PII_COLORS[matchingEntity.type] || { hex: '#FFD700', textHex: '#000000' };

            // Fill redaction rectangle
            context.fillStyle = colorInfo.hex;
            context.fillRect(cX, cY, cWidth, cHeight);

            // Draw pseudonym overlay text
            context.fillStyle = colorInfo.textHex;
            const drawFontSize = Math.max(9, fontHeightInCanvas * 0.7);
            context.font = `bold ${drawFontSize}px Arial, sans-serif`;
            context.textBaseline = 'middle';
            context.textAlign = 'center';
            context.fillText(matchingEntity.pseudonym, cX + (cWidth / 2), cY + (cHeight / 2));
          });
        }
      }

      // 4. Paint solid white over marked header and footer zones so they are physically removed from visual output
      headerZones.forEach(zone => {
        if (zone.pageNum === undefined || zone.pageNum === i) {
          context.fillStyle = 'white';
          context.fillRect(
            zone.left * viewport.width,
            zone.top * viewport.height,
            (zone.right - zone.left) * viewport.width,
            (zone.bottom - zone.top) * viewport.height
          );
        }
      });

      footerZones.forEach(zone => {
        if (zone.pageNum === undefined || zone.pageNum === i) {
          context.fillStyle = 'white';
          context.fillRect(
            zone.left * viewport.width,
            zone.top * viewport.height,
            (zone.right - zone.left) * viewport.width,
            (zone.bottom - zone.top) * viewport.height
          );
        }
      });

      // Convert canvas to image and add to new PDF
      const imageData = canvas.toDataURL('image/png');
      const imageBytes = dataURLToUint8Array(imageData);
      const image = await outPdfDoc.embedPng(imageBytes);
      
      const outPage = outPdfDoc.addPage([viewport.width / scale, viewport.height / scale]);
      outPage.drawImage(image, {
        x: 0,
        y: 0,
        width: viewport.width / scale,
        height: viewport.height / scale,
      });

      // Remove canvas from DOM
      document.body.removeChild(canvas);
    }

    await pdf.destroy();
    return await outPdfDoc.save();
  };

  // --- UI Helpers ---

  const filteredEntities = entities.filter(e => {
    const matchesSearch = e.original.toLowerCase().includes(searchTerm.toLowerCase()) || 
                         e.pseudonym.toLowerCase().includes(searchTerm.toLowerCase());
    const matchesType = filterType === 'ALL' || e.type === filterType;
    const matchesIgnored = hideIgnored ? !e.ignored : true;
    
    // Filter by selected file if in review mode
    const matchesFile = (reviewMode && selectedFileId) ? e.fileIds?.includes(selectedFileId) : true;
    
    return matchesSearch && matchesType && matchesIgnored && matchesFile;
  });

  const pseudonymAnalysis = useMemo(() => {
    const pToO = new Map<string, Set<string>>();
    const pToGroups = new Map<string, Set<string>>();
    
    entities.forEach(e => {
      if (e.ignored || !e.enabled) return;
      const p = e.pseudonym.trim();
      if (!p) return;
      
      const normalizedOriginal = e.original.toLowerCase().trim();
      if (!pToO.has(p)) pToO.set(p, new Set());
      pToO.get(p)!.add(normalizedOriginal);
      
      const groupKey = e.groupId || `single-${e.id}`;
      if (!pToGroups.has(p)) pToGroups.set(p, new Set());
      pToGroups.get(p)!.add(groupKey);
    });
    
    const conflicts = new Set<string>();
    pToGroups.forEach((groups, pseudonym) => {
      if (groups.size > 1) conflicts.add(pseudonym);
    });
    
    return { conflicts, pToO, pToGroups };
  }, [entities]);

  const groupedEntities = filteredEntities.reduce((acc, entity) => {
    const key = entity.groupId || `single-${entity.id}`;
    if (!acc[key]) acc[key] = [];
    acc[key].push(entity);
    return acc;
  }, {} as Record<string, PIIEntity[]>);

  const sortedGroupedEntries = useMemo(() => {
    return Object.entries(groupedEntities).sort((a, b) => {
      const aVals = a[1] as PIIEntity[];
      const bVals = b[1] as PIIEntity[];
      const aHasNew = aVals.some(e => e.isNewlyIdentified);
      const bHasNew = bVals.some(e => e.isNewlyIdentified);
      if (aHasNew && !bHasNew) return -1;
      if (!aHasNew && bHasNew) return 1;
      return 0;
    });
  }, [groupedEntities]);

  const handleExportProject = () => {
    // We omit rawFile as it's not JSON serializable easily
    const projectData = {
      files: files.map(({ rawFile, ...rest }) => rest),
      entities,
      globalKnowledge,
      isRelated,
      version: '1.0'
    };
    const blob = new Blob([JSON.stringify(projectData, null, 2)], { type: 'application/json' });
    saveAs(blob, `projeto_anonimizacao_${new Date().toISOString().split('T')[0]}.json`);
  };

  const handleImportProject = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const imported = JSON.parse(e.target?.result as string);
        if (imported.files && imported.entities) {
          // Note: rawFile will be missing, so PDF re-export won't work without re-upload
          setFiles(imported.files.map((f: any) => ({ ...f, status: 'done' })));
          setEntities(imported.entities);
          if (imported.globalKnowledge) setGlobalKnowledge(imported.globalKnowledge);
          if (imported.isRelated !== undefined) setIsRelated(imported.isRelated);
        }
      } catch (err) {
        console.error("Erro ao importar projeto:", err);
      }
    };
    reader.readAsText(file);
  };

  return (
    <div className="min-h-screen bg-gray-50 text-gray-900 font-sans">
      {/* Entity Detail Modal */}
      <AnimatePresence>
        {editingEntity && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/50 backdrop-blur-sm">
            <motion.div 
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              className="bg-white rounded-2xl shadow-xl w-full max-w-2xl overflow-hidden"
            >
              <div className="p-6 border-b border-gray-100 flex items-center justify-between">
                <h3 className="text-lg font-bold">Tratar Elemento</h3>
                <button 
                  onClick={() => {
                    setEditingEntity(null);
                    setSplitResults(null);
                  }} 
                  className="p-2 hover:bg-gray-100 rounded-full"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>
              
              <div className="p-6 space-y-6 max-h-[70vh] overflow-y-auto">
                {!splitResults ? (
                  <>
                    <div>
                      <label className="block text-sm font-medium text-gray-500 mb-2">Contexto no Documento</label>
                      <div className="p-4 bg-gray-50 rounded-xl text-sm leading-relaxed whitespace-pre-wrap text-center">
                        <div className="text-gray-400 opacity-60 text-[10px] mb-2 italic">
                          {editingEntity.contextBefore?.split(/\s+/).slice(-30, -15).join(' ') || '...'}
                        </div>
                        <div className="bg-white p-4 rounded-xl border border-gray-100 shadow-sm inline-block max-w-full text-left">
                          <span className="text-gray-400">{editingEntity.contextBefore?.split(/\s+/).slice(-15).join(' ')}</span>
                          <span className="bg-yellow-200 px-1.5 py-0.5 rounded font-bold mx-1 text-gray-900 shadow-sm ring-2 ring-yellow-400/20">{editingEntity.original}</span>
                          <span className="text-gray-400">{editingEntity.contextAfter?.split(/\s+/).slice(0, 15).join(' ')}</span>
                        </div>
                        <div className="text-gray-400 opacity-60 text-[10px] mt-2 italic">
                          {editingEntity.contextAfter?.split(/\s+/).slice(15, 30).join(' ') || '...'}
                        </div>
                      </div>
                    </div>

                    <div className="grid grid-cols-3 gap-4">
                      <div>
                        <label className="block text-sm font-medium text-gray-500 mb-2">Texto Original</label>
                        <input 
                          type="text" 
                          value={editingEntity.original}
                          onChange={(e) => handleUpdateEntity(editingEntity.id, { original: e.target.value })}
                          className="w-full px-4 py-2 border border-gray-200 rounded-lg focus:ring-2 focus:ring-indigo-500 outline-none"
                        />
                      </div>
                      <div>
                        <label className="block text-sm font-medium text-gray-500 mb-2">Tipo</label>
                        <select 
                          value={editingEntity.type}
                          onChange={(e) => handleUpdateEntity(editingEntity.id, { type: e.target.value })}
                          className="w-full px-4 py-2 border border-gray-200 rounded-lg focus:ring-2 focus:ring-indigo-500 outline-none bg-white"
                        >
                          {ORDERED_PII_TYPES.map(type => (
                            <option key={type} value={type}>{type}</option>
                          ))}
                        </select>
                      </div>
                      <div>
                        <label className="block text-sm font-medium text-gray-500 mb-2">Pseudónimo (Sigla)</label>
                        <input 
                          type="text" 
                          value={editingEntity.pseudonym}
                          onChange={(e) => handleUpdateEntity(editingEntity.id, { pseudonym: e.target.value })}
                          className="w-full px-4 py-2 border border-gray-200 rounded-lg focus:ring-2 focus:ring-indigo-500 outline-none"
                        />
                      </div>
                    </div>

                    <div className="space-y-3">
                      <label className="block text-sm font-medium text-gray-500">Ajustar Limites</label>
                      <div className="flex flex-wrap gap-2">
                        <button 
                          onClick={handleExpandStart}
                          className="flex items-center gap-1 px-3 py-1.5 bg-indigo-50 text-indigo-700 rounded-lg text-xs hover:bg-indigo-100 transition-colors"
                        >
                          <Plus className="w-3 h-3" />
                          <span>Expandir Início</span>
                        </button>
                        <button 
                          onClick={handleShrinkStart}
                          className="flex items-center gap-1 px-3 py-1.5 bg-gray-50 text-gray-700 rounded-lg text-xs hover:bg-gray-100 transition-colors"
                        >
                          <Scissors className="w-3 h-3" />
                          <span>Reduzir Início</span>
                        </button>
                        <div className="w-px h-6 bg-gray-200 mx-1" />
                        <button 
                          onClick={handleShrinkEnd}
                          className="flex items-center gap-1 px-3 py-1.5 bg-gray-50 text-gray-700 rounded-lg text-xs hover:bg-gray-100 transition-colors"
                        >
                          <Scissors className="w-3 h-3" />
                          <span>Reduzir Fim</span>
                        </button>
                        <button 
                          onClick={handleExpandEnd}
                          className="flex items-center gap-1 px-3 py-1.5 bg-indigo-50 text-indigo-700 rounded-lg text-xs hover:bg-indigo-100 transition-colors"
                        >
                          <Plus className="w-3 h-3" />
                          <span>Expandir Fim</span>
                        </button>
                      </div>
                    </div>

                    <div className="space-y-3">
                      <label className="block text-sm font-medium text-gray-500">Ações de Limpeza e Associação</label>
                      <div className="flex flex-wrap gap-2">
                        <button 
                          onClick={() => {
                            if (!editingEntity) return;
                            // Remove single spaces between letters, e.g., "S i m õ e s" -> "Simões"
                            let cleaned = editingEntity.original;
                            for (let i = 0; i < 5; i++) {
                              cleaned = cleaned.replace(/(\b\w)\s+(?=\w\b)/g, '$1');
                            }
                            cleaned = cleaned.replace(/\s+/g, ' ').trim();
                            handleUpdateEntity(editingEntity.id, { original: cleaned });
                            showToast("Espaços de OCR limpos.", "info");
                          }}
                          className="flex items-center gap-1 px-3 py-1.5 bg-amber-50 text-amber-700 rounded-lg text-xs hover:bg-amber-100 transition-colors border border-amber-200"
                        >
                          <Zap className="w-3 h-3" />
                          <span>Limpar Espaços OCR</span>
                        </button>
                        <button 
                          onClick={() => {
                            if (!editingEntity) return;
                            setShowMergeModal(true);
                            setSelectedIds(new Set([editingEntity.id]));
                          }}
                          className="flex items-center gap-1 px-3 py-1.5 bg-indigo-50 text-indigo-700 rounded-lg text-xs hover:bg-indigo-100 transition-colors border border-indigo-200"
                        >
                          <Link className="w-3 h-3" />
                          <span>Mesclar com outro...</span>
                        </button>
                      </div>
                    </div>

                    <div className="space-y-3">
                      <div className="flex items-center justify-between">
                        <label className="block text-sm font-medium text-gray-500">Ações de Divisão (Múltiplos Elementos)</label>
                        <button 
                          onClick={() => {
                            if (!editingEntity) return;
                            const fullContext = editingEntity.context?.text || editingEntity.original;
                            handleUpdateEntity(editingEntity.id, { original: fullContext });
                            showToast("Texto expandido para o contexto completo.", "info");
                          }}
                          className="text-[10px] font-bold text-indigo-600 hover:text-indigo-700 uppercase tracking-wider"
                        >
                          Usar Contexto Completo
                        </button>
                      </div>
                      <div className="flex flex-wrap gap-2">
                        <button 
                          onClick={() => handleManualSplit(editingEntity, /\s+/)}
                          className="px-3 py-1.5 bg-indigo-50 text-indigo-700 rounded-lg text-sm hover:bg-indigo-100 transition-colors border border-indigo-100"
                        >
                          Dividir por Espaço
                        </button>
                        <button 
                          onClick={() => handleManualSplit(editingEntity, /\s+e\s+/i)}
                          className="px-3 py-1.5 bg-indigo-50 text-indigo-700 rounded-lg text-sm hover:bg-indigo-100 transition-colors border border-indigo-100"
                        >
                          Dividir por " e "
                        </button>
                        <button 
                          onClick={() => handleManualSplit(editingEntity, ',')}
                          className="px-3 py-1.5 bg-indigo-50 text-indigo-700 rounded-lg text-sm hover:bg-indigo-100 transition-colors border border-indigo-100"
                        >
                          Dividir por Vírgula
                        </button>
                        <button 
                          onClick={() => {
                            if (!editingEntity) return;
                            setSplitResults([
                              { original: editingEntity.original, type: editingEntity.type, id: generateId() }
                            ]);
                          }}
                          className="px-3 py-1.5 bg-gray-50 text-gray-700 rounded-lg text-sm hover:bg-gray-100 transition-colors border border-gray-200"
                        >
                          Dividir Manualmente
                        </button>
                      </div>
                      <p className="text-[10px] text-gray-400 italic">
                        Dica: Se o elemento contiver vários nomes (ex: "Nome1 e Nome2"), use "Dividir por ' e '".
                      </p>
                    </div>
                  </>
                ) : (
                  <div className="space-y-4">
                    <div className="p-3 bg-indigo-50 rounded-xl border border-indigo-100 mb-4">
                      <p className="text-xs text-indigo-700 font-medium">
                        O elemento foi dividido em {splitResults.length} partes. Classifique cada uma abaixo antes de confirmar.
                      </p>
                    </div>
                    {splitResults.map((part, idx) => (
                      <div key={part.id} className="flex gap-3 items-end p-3 bg-gray-50 rounded-xl border border-gray-100 group relative">
                        <div className="flex-1">
                          <label className="block text-[10px] font-bold text-gray-400 uppercase mb-1">Texto {idx + 1}</label>
                          <input 
                            type="text" 
                            value={part.original}
                            onChange={(e) => {
                              const next = [...splitResults];
                              next[idx].original = e.target.value;
                              setSplitResults(next);
                            }}
                            className="w-full px-3 py-1.5 text-sm border border-gray-200 rounded-lg bg-white focus:ring-2 focus:ring-indigo-500 outline-none"
                          />
                        </div>
                        <div className="w-32">
                          <label className="block text-[10px] font-bold text-gray-400 uppercase mb-1">Tipo</label>
                          <select 
                            value={part.type}
                            onChange={(e) => {
                              const next = [...splitResults];
                              next[idx].type = e.target.value;
                              setSplitResults(next);
                            }}
                            className="w-full px-2 py-1.5 text-sm border border-gray-200 rounded-lg bg-white focus:ring-2 focus:ring-indigo-500 outline-none"
                          >
                            {ORDERED_PII_TYPES.map(type => (
                              <option key={type} value={type}>{type}</option>
                            ))}
                          </select>
                        </div>
                        {splitResults.length > 1 && (
                          <button 
                            onClick={() => {
                              setSplitResults(splitResults.filter((_, i) => i !== idx));
                            }}
                            className="p-2 text-red-400 hover:text-red-600 transition-colors"
                            title="Remover esta parte"
                          >
                            <X className="w-4 h-4" />
                          </button>
                        )}
                      </div>
                    ))}
                    <div className="flex items-center justify-between pt-2">
                      <button 
                        onClick={() => setSplitResults(null)}
                        className="text-xs text-indigo-600 font-bold hover:underline"
                      >
                        ← Cancelar divisão e voltar
                      </button>
                      <button 
                        onClick={() => {
                          if (!editingEntity) return;
                          setSplitResults(prev => [
                            ...(prev || []),
                            { original: "", type: editingEntity.type, id: generateId() }
                          ]);
                        }}
                        className="flex items-center gap-2 px-4 py-2 bg-white border border-indigo-200 text-indigo-600 rounded-lg text-xs font-bold hover:bg-indigo-50 transition-colors shadow-sm uppercase tracking-wider"
                      >
                        <Plus className="w-3 h-3" />
                        <span>Adicionar Parte</span>
                      </button>
                    </div>
                  </div>
                )}
              </div>

              <div className="p-6 bg-gray-50 flex justify-end gap-3">
                <button 
                  onClick={() => {
                    setEditingEntity(null);
                    setSplitResults(null);
                  }}
                  className="px-6 py-2 text-gray-600 font-medium hover:bg-gray-200 rounded-lg transition-colors"
                >
                  Fechar
                </button>
                {splitResults ? (
                  <button 
                    onClick={confirmSplit}
                    disabled={splitResults.length === 0}
                    className="px-6 py-2 bg-indigo-600 text-white font-medium hover:bg-indigo-700 rounded-lg transition-colors disabled:opacity-50"
                  >
                    Confirmar Divisão e Validar
                  </button>
                ) : (
                  <button 
                    onClick={() => {
                      handleUpdateEntity(editingEntity.id, { treated: true });
                      setEditingEntity(null);
                    }}
                    className="px-6 py-2 bg-indigo-600 text-white font-medium hover:bg-indigo-700 rounded-lg transition-colors"
                  >
                    Validar e Sair
                  </button>
                )}
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Header */}
      <header className="bg-white border-b border-gray-200 sticky top-0 z-30">
        <div className="max-w-full mx-auto px-6 h-16 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="bg-indigo-600 p-2 rounded-lg">
              <Lock className="w-5 h-5 text-white" />
            </div>
            <h1 className="text-xl font-bold tracking-tight">SEATRI</h1>
          </div>
          
          <div className="flex items-center gap-4">
            <div className="flex items-center gap-1 border-r border-gray-200 pr-4 mr-2">
              <button 
                onClick={undo}
                disabled={historyIndex <= 0}
                className="p-2 hover:bg-gray-100 rounded-lg transition-colors text-gray-600 disabled:opacity-30"
                title="Anular (Undo)"
              >
                <RotateCcw className="w-4 h-4" />
              </button>
              <button 
                onClick={redo}
                disabled={historyIndex >= history.length - 1}
                className="p-2 hover:bg-gray-100 rounded-lg transition-colors text-gray-600 disabled:opacity-30"
                title="Refazer (Redo)"
              >
                <RotateCw className="w-4 h-4" />
              </button>
            </div>

            <div className="flex items-center gap-1 border-r border-gray-200 pr-4 mr-2">
              <button 
                onClick={handleRescan}
                disabled={files.length === 0 || isProcessing}
                className={`p-2 hover:bg-gray-100 rounded-lg transition-colors text-gray-600 disabled:opacity-30 ${isProcessing ? 'animate-spin' : ''}`}
                title="Re-analisar Documentos (Consistência Global)"
              >
                <RefreshCw className="w-4 h-4" />
              </button>
              <button 
                onClick={() => setSplitView(!splitView)}
                className={`p-2 rounded-lg transition-colors ${splitView ? 'bg-indigo-100 text-indigo-600' : 'hover:bg-gray-100 text-gray-600'}`}
                title="Visualização Lado-a-Lado (Split View)"
              >
                <Layers className="w-4 h-4" />
              </button>
              <button 
                onClick={() => setReviewMode(!reviewMode)}
                className={`p-2 rounded-lg transition-colors ${reviewMode ? 'bg-indigo-100 text-indigo-600' : 'hover:bg-gray-100 text-gray-600'}`}
                title="Modo de Revisão (Documento Original)"
              >
                <Eye className="w-4 h-4" />
              </button>
              <button 
                onClick={handleExportReport}
                disabled={entities.length === 0}
                className="p-2 hover:bg-gray-100 rounded-lg transition-colors text-gray-600 disabled:opacity-30"
                title="Exportar Relatório de Correspondências (Excel)"
              >
                <FileText className="w-4 h-4" />
              </button>
            </div>

            <div className="flex items-center gap-1 border-r border-gray-200 pr-4 mr-2">
              <label className="cursor-pointer p-2 hover:bg-gray-100 rounded-lg transition-colors text-gray-600" title="Importar Projeto (JSON)">
                <FolderOpen className="w-4 h-4" />
                <input type="file" accept=".json" className="hidden" onChange={handleImportProject} />
              </label>
              <button 
                onClick={handleExportProject}
                disabled={files.length === 0}
                className="p-2 hover:bg-gray-100 rounded-lg transition-colors text-gray-600 disabled:opacity-30"
                title="Exportar Projeto (JSON)"
              >
                <Save className="w-4 h-4" />
              </button>
            </div>
            <button 
              onClick={() => setShowExceptionsModal(true)}
              className="flex items-center gap-2 text-gray-600 hover:text-indigo-600 px-3 py-2 rounded-lg font-medium transition-colors"
            >
              <Shield className="w-4 h-4" />
              <span>Exceções</span>
            </button>

            <button 
              onClick={clearAll}
              className={`flex items-center gap-2 px-3 py-2 rounded-lg font-medium transition-all ${
                confirmingClear 
                  ? 'bg-red-600 text-white animate-pulse' 
                  : 'text-red-600 hover:bg-red-50'
              }`}
              title={confirmingClear ? "Confirmar Limpeza" : "Limpar todo o projeto e começar um novo"}
            >
              <Trash2 className="w-4 h-4" />
              <span className="hidden sm:inline">
                {confirmingClear ? "Confirmar?" : "Novo Projeto"}
              </span>
            </button>

            <button 
              onClick={() => setShowExportModal(true)}
              disabled={files.length === 0 || isProcessing}
              title="Gera versões anonimizadas dos ficheiros carregados. PDFs são reconstruídos com pseudónimos."
              className="flex items-center gap-2 bg-indigo-600 hover:bg-indigo-700 disabled:bg-gray-300 text-white px-4 py-2 rounded-lg font-medium transition-colors"
            >
              <Download className="w-4 h-4" />
              <span>Exportar</span>
            </button>
          </div>
        </div>
      </header>

      <main className="max-w-full mx-auto px-4 py-8">
        <div className={`grid grid-cols-1 ${splitView ? 'lg:grid-cols-1' : 'lg:grid-cols-12'} gap-8`}>
          {/* Left Column: Files & Controls or Document Viewer */}
          {!splitView && (
            <div className={`${reviewMode ? 'lg:col-span-8 sticky top-8' : 'lg:col-span-4'} space-y-6 flex flex-col h-[calc(100vh-120px)]`}>
              {/* Search Box (Always visible at top) */}
              <div className="bg-white p-6 rounded-2xl border border-gray-200 shadow-sm shrink-0">
                <div className="flex items-center justify-between mb-4">
                  <h2 className="text-sm font-semibold text-gray-500 uppercase tracking-wider">Procurar</h2>
                  {searchTerm && (
                    <button 
                      onClick={() => setSearchTerm('')}
                      className="text-xs text-indigo-600 hover:text-indigo-800 font-medium"
                    >
                      Limpar
                    </button>
                  )}
                </div>
                <div className="relative">
                  <Search className="w-5 h-5 absolute left-4 top-1/2 -translate-y-1/2 text-gray-400" />
                  <input 
                    type="text" 
                    placeholder="Procurar elementos ou pseudónimos..."
                    value={searchTerm}
                    onChange={(e) => setSearchTerm(e.target.value)}
                    autoComplete="off"
                    className="w-full pl-12 pr-4 py-3 bg-gray-50 border border-gray-200 rounded-xl text-base focus:outline-none focus:ring-2 focus:ring-indigo-500/20 transition-all font-medium"
                  />
                </div>
              </div>

              {reviewMode ? (
                <div className="flex-1 flex flex-col bg-white rounded-2xl border border-gray-200 shadow-sm overflow-hidden min-h-0">
                  <div className="p-4 border-b border-gray-100 flex items-center justify-between bg-gray-50/50 shrink-0">
                    <div className="flex items-center gap-4">
                      <div className="flex items-center gap-2">
                        <Eye className="w-4 h-4 text-indigo-600" />
                        <h2 className="text-sm font-bold text-gray-700 uppercase tracking-wider">Visualizador Original</h2>
                      </div>
                      <div className="flex items-center gap-2 border-l pl-4">
                        <button
                          onClick={() => setSelectionMode(selectionMode === 'header' ? 'none' : 'header')}
                          className={`flex items-center gap-2 px-3 py-1.5 rounded-lg text-xs font-medium transition-colors ${
                            selectionMode === 'header' 
                              ? 'bg-red-100 text-red-700 border border-red-200' 
                              : 'bg-white text-gray-600 hover:bg-gray-100 border border-gray-200'
                          }`}
                        >
                          <Crop className="w-3.5 h-3.5" />
                          Marcar Cabeçalho
                        </button>
                        <button
                          onClick={() => setSelectionMode(selectionMode === 'footer' ? 'none' : 'footer')}
                          className={`flex items-center gap-2 px-3 py-1.5 rounded-lg text-xs font-medium transition-colors ${
                            selectionMode === 'footer' 
                              ? 'bg-blue-100 text-blue-700 border border-blue-200' 
                              : 'bg-white text-gray-600 hover:bg-gray-100 border border-gray-200'
                          }`}
                        >
                          <Crop className="w-3.5 h-3.5" />
                          Marcar Rodapé
                        </button>
                        {(headerZones.length > 0 || footerZones.length > 0) && (
                          <button
                            onClick={() => {
                              setHeaderZones([]);
                              setFooterZones([]);
                              setSelectionMode('none');
                            }}
                            className="flex items-center gap-1 px-2 py-1.5 text-xs text-red-500 hover:bg-red-50 rounded-lg transition-colors"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                            Limpar Zonas
                          </button>
                        )}
                      </div>
                    </div>
                    <button 
                      onClick={() => {
                        setReviewMode(false);
                        setSelectionMode('none');
                      }} 
                      className="text-[10px] font-bold text-gray-400 hover:text-gray-600 uppercase tracking-widest flex items-center gap-1"
                    >
                      <X className="w-3 h-3" />
                      Fechar
                    </button>
                  </div>
                  <DocumentViewer 
                    file={files.find(f => f.id === selectedFileId) || null} 
                    entities={entities}
                    selectedEntityId={selectedEntityId}
                    selectedIds={selectedIds}
                    setPendingManualTerm={setPendingManualTerm}
                    globalKnowledge={globalKnowledge}
                    safelist={safelist}
                    selectionMode={selectionMode}
                    onAreaSelected={(rect, pageNum, image) => {
                      if (selectionMode === 'header') {
                        setHeaderZones(prev => [...prev, { ...rect, image }]);
                        extractTextFromZone(rect, 'HEADER', pageNum);
                      }
                      if (selectionMode === 'footer') {
                        setFooterZones(prev => [...prev, { ...rect, image }]);
                        extractTextFromZone(rect, 'HEADER', pageNum); // Both go to HEADER for removal logic
                      }
                    }}
                    headerZones={headerZones}
                    footerZones={footerZones}
                    onFileUpload={handleFileUpload}
                  />
                </div>
              ) : (
                <>
                  {/* Upload Box */}
                  <div className="bg-white p-6 rounded-2xl border border-gray-200 shadow-sm">
                    <div className="flex items-center justify-between mb-4">
                      <h2 className="text-sm font-semibold text-gray-500 uppercase tracking-wider">Documentos</h2>
                      {files.length > 0 && (
                        <button 
                          onClick={clearAll}
                          className="text-xs text-red-500 hover:text-red-700 flex items-center gap-1 font-medium"
                        >
                          <RotateCcw className="w-3 h-3" />
                          Limpar Tudo
                        </button>
                      )}
                    </div>
                    <label className="flex flex-col items-center justify-center w-full h-32 border-2 border-dashed border-gray-300 rounded-xl cursor-pointer hover:bg-gray-50 transition-colors">
                      <div className="flex flex-col items-center justify-center pt-5 pb-6">
                        <Upload className="w-8 h-8 text-gray-400 mb-2" />
                        <p className="text-sm text-gray-500 text-center px-4">Clique ou arraste ficheiros (incluindo relacionados)</p>
                        <p className="text-xs text-gray-400 mt-1">PDF, DOCX, XLSX, TXT</p>
                      </div>
                      <input type="file" className="hidden" multiple accept=".pdf,.docx,.xlsx,.txt" onChange={handleFileUpload} />
                    </label>

                    <div className="mt-4 space-y-2">
                      {files.map(file => (
                        <div 
                          key={file.id} 
                          onClick={() => setSelectedFileId(file.id)}
                          className={`flex items-center justify-between p-3 rounded-lg border cursor-pointer transition-all ${selectedFileId === file.id ? 'bg-indigo-50 border-indigo-200 shadow-sm' : 'bg-gray-50 border-gray-100 hover:border-gray-300'}`}
                        >
                          <div className="flex items-center gap-3 overflow-hidden">
                            <FileText className="w-4 h-4 text-indigo-500 flex-shrink-0" />
                            <span className="text-sm font-medium truncate">{file.name}</span>
                          </div>
                          {file.status === 'processing' ? (
                            <div className="w-4 h-4 border-2 border-indigo-500 border-t-transparent rounded-full animate-spin" />
                          ) : file.status === 'done' ? (
                            <CheckCircle2 className="w-4 h-4 text-green-500" />
                          ) : file.status === 'error' ? (
                            <AlertCircle className="w-4 h-4 text-red-500" />
                          ) : null}
                        </div>
                      ))}
                    </div>
                  </div>

                  {/* Manual Button */}
                  <button 
                    onClick={() => setShowManualModal(true)}
                    className="w-full flex items-center justify-between p-4 bg-white rounded-2xl border border-gray-200 shadow-sm hover:border-indigo-300 hover:bg-indigo-50/30 transition-all group"
                  >
                    <div className="flex items-center gap-3">
                      <div className="w-10 h-10 bg-indigo-100 rounded-xl flex items-center justify-center text-indigo-600 group-hover:bg-indigo-600 group-hover:text-white transition-all">
                        <FileText className="w-5 h-5" />
                      </div>
                      <div className="text-left">
                        <h3 className="text-sm font-bold text-gray-900">Manual do Utilizador</h3>
                        <p className="text-xs text-gray-500">Guia completo e ajuda</p>
                      </div>
                    </div>
                    <ChevronRight className="w-5 h-5 text-gray-400 group-hover:text-indigo-600 transition-all" />
                  </button>

                  <div className="p-4 bg-amber-50 rounded-2xl border border-amber-100 flex gap-3">
                    <AlertCircle className="w-5 h-5 text-amber-600 shrink-0" />
                    <div className="space-y-1">
                      <h4 className="text-xs font-bold text-amber-900 uppercase tracking-wider">Algo escapou?</h4>
                      <p className="text-xs text-amber-800 leading-relaxed">
                        Se a aplicação não detetou um nome, <strong>selecione o texto diretamente no documento original</strong> à esquerda ou use o botão <strong>"Re-analisar"</strong> no topo para uma varredura de consistência.
                      </p>
                    </div>
                  </div>
                </>
              )}
            </div>
          )}

          {/* Right Column: Entities List or Split View */}
          <div className={`${splitView ? 'lg:col-span-12' : reviewMode ? 'lg:col-span-4' : 'lg:col-span-8'} space-y-6`}>
            {splitView ? (
              <div className="bg-white rounded-2xl border border-gray-200 shadow-sm overflow-hidden flex flex-col h-[calc(100vh-200px)]">
                <div className="p-4 border-b border-gray-100 flex items-center justify-between bg-gray-50/50">
                  <div className="flex items-center gap-4">
                    <h2 className="text-sm font-bold text-gray-700">Comparação Lado-a-Lado</h2>
                    <select 
                      value={selectedFileId || ''} 
                      onChange={(e) => setSelectedFileId(e.target.value)}
                      className="text-xs border border-gray-200 rounded px-2 py-1 bg-white"
                    >
                      <option value="">Selecionar Ficheiro...</option>
                      {files.map(f => <option key={f.id} value={f.id}>{f.name}</option>)}
                    </select>
                  </div>
                  <button onClick={() => setSplitView(false)} className="text-xs text-gray-500 hover:text-gray-700">Fechar</button>
                </div>
                
                <div className="flex-1 flex overflow-hidden">
                  {/* Original Text */}
                  <div className="flex-1 border-r border-gray-100 flex flex-col">
                    <div className="p-2 bg-gray-100/50 text-[10px] font-bold text-gray-500 uppercase tracking-wider text-center">Original</div>
                    <div 
                      ref={leftScrollRef}
                      onScroll={() => handleSyncScroll('left')}
                      onMouseUp={(e) => {
                        const selection = window.getSelection();
                        const text = selection?.toString().trim();
                        if (text && text.length > 1) {
                          setPendingManualTerm({
                            text,
                            x: e.clientX,
                            y: e.clientY
                          });
                        }
                      }}
                      className="flex-1 p-6 overflow-y-auto font-mono text-sm whitespace-pre-wrap text-gray-600 bg-white relative"
                    >
                      {selectedFileId ? (
                        <HighlightText 
                          text={files.find(f => f.id === selectedFileId)?.content || ""} 
                          entities={[
                            ...entities,
                            ...(selectedEntityId ? [{ 
                              id: 'selected-highlight', 
                              original: entities.find(e => e.id === selectedEntityId)?.original || '', 
                              type: 'ACTIVE_HIGHLIGHT', 
                              pseudonym: '', 
                              enabled: true 
                            }] : []),
                            ...Array.from(selectedIds).map(id => ({
                              id: `selected-id-${id}`,
                              original: entities.find(e => e.id === id)?.original || '',
                              type: 'HIGHLIGHT',
                              pseudonym: '',
                              enabled: true
                            }))
                          ]} 
                          mode="original" 
                          globalKnowledge={globalKnowledge}
                          safelist={safelist}
                        />
                      ) : "Selecione um ficheiro para visualizar"}

              {/* Navigation Arrows for Untreated Elements */}
              {selectedFileId && (
                <div className="fixed bottom-8 right-1/2 translate-x-[-20px] flex flex-col gap-2 z-50">
                  <button 
                    onClick={() => navigateUntreated('prev')}
                    className="p-3 bg-white shadow-lg border border-gray-200 rounded-full text-gray-600 hover:bg-gray-50 hover:text-indigo-600 transition-all"
                    title="Anterior elemento não tratado"
                  >
                    <ChevronUp className="w-5 h-5" />
                  </button>
                  <button 
                    onClick={() => navigateUntreated('next')}
                    className="p-3 bg-white shadow-lg border border-gray-200 rounded-full text-gray-600 hover:bg-gray-50 hover:text-indigo-600 transition-all"
                    title="Próximo elemento não tratado"
                  >
                    <ChevronDown className="w-5 h-5" />
                  </button>
                </div>
              )}

                    </div>
                  </div>
                  
                  {/* Anonymized Text */}
                  <div className="flex-1 flex flex-col">
                    <div className="p-2 bg-indigo-50 text-[10px] font-bold text-indigo-500 uppercase tracking-wider text-center">Anonimizado</div>
                    <div 
                      ref={rightScrollRef}
                      onScroll={() => handleSyncScroll('right')}
                      className="flex-1 p-6 overflow-y-auto font-mono text-sm whitespace-pre-wrap text-gray-900 bg-white"
                    >
                      {selectedFileId ? (
                        <HighlightText 
                          text={anonymizeText(files.find(f => f.id === selectedFileId)?.content || "", entities)} 
                          entities={entities} 
                          mode="anonymized" 
                        />
                      ) : "Selecione um ficheiro para visualizar"}
                    </div>
                  </div>
                </div>
              </div>
            ) : (
              <>
                {/* Search & Bulk Actions Toolbar */}
                <div className="bg-white p-4 rounded-2xl border border-gray-200 shadow-sm space-y-4">
              <div className="flex flex-wrap items-center justify-center sm:justify-start gap-2 max-h-[92px] overflow-y-auto w-full">
                <div className="flex items-center gap-1 bg-gray-100 p-1 rounded-xl mr-2">
                  <button 
                    onClick={() => setIsRelated(true)}
                    className={`px-3 py-1.5 text-[10px] font-bold uppercase tracking-wider rounded-lg transition-all ${isRelated ? 'bg-white shadow-sm text-indigo-600' : 'text-gray-500 hover:text-gray-700'}`}
                    title="Agrupar elementos entre todos os documentos"
                  >
                    Relacionados
                  </button>
                  <button 
                    onClick={() => setIsRelated(false)}
                    className={`px-3 py-1.5 text-[10px] font-bold uppercase tracking-wider rounded-lg transition-all ${!isRelated ? 'bg-white shadow-sm text-indigo-600' : 'text-gray-500 hover:text-gray-700'}`}
                    title="Agrupar elementos apenas dentro do mesmo documento"
                  >
                    Independentes
                  </button>
                </div>
                <button 
                  onClick={handleValidateAllEntities}
                  className="flex items-center gap-2 px-4 py-2 bg-green-600 text-white rounded-xl text-sm font-bold hover:bg-green-700 transition-all shadow-sm"
                  title="Validar todos os elementos detetados de uma vez"
                >
                  <CheckCircle2 className="w-4 h-4" />
                  <span>Validar Tudo</span>
                </button>
                <button 
                  onClick={handleSplitAllAndEntities}
                  className="flex items-center gap-2 px-4 py-2 bg-blue-50 text-blue-700 rounded-xl text-sm font-bold hover:bg-blue-100 transition-colors"
                  title="Dividir nomes que contêm ' e ' (ex: Nome1 e Nome2)"
                >
                  <Scissors className="w-4 h-4" />
                  <span>Dividir 'e'</span>
                </button>
                <button 
                  onClick={handleRescan}
                  className="flex items-center gap-2 px-4 py-2 bg-emerald-50 text-emerald-700 rounded-xl text-sm font-bold hover:bg-emerald-100 transition-colors"
                  title="Re-analisar todos os documentos com base no conhecimento global atual"
                >
                  <RotateCw className="w-4 h-4" />
                  <span>Re-analisar Documentos</span>
                </button>
                <button 
                  onClick={handleSuggestGroups}
                  className="flex items-center gap-2 px-4 py-2 bg-amber-50 text-amber-700 rounded-xl text-sm font-bold hover:bg-amber-100 transition-colors"
                  title="Sugerir agrupamentos automáticos para nomes semelhantes"
                >
                  <Plus className="w-4 h-4" />
                  <span>Sugerir Grupos</span>
                </button>
                <button 
                  onClick={handleReGroup}
                  className="flex items-center gap-2 px-4 py-2 bg-indigo-50 text-indigo-700 rounded-xl text-sm font-bold hover:bg-indigo-100 transition-colors"
                  title="Re-analisar agrupamentos com base no tratamento atual"
                >
                  <RotateCw className="w-4 h-4" />
                  <span>Re-agrupar</span>
                </button>
                <button
                  onClick={() => setHideIgnored(!hideIgnored)}
                  className={`flex items-center gap-2 px-3 py-2 rounded-lg text-xs font-bold transition-colors border ${
                    hideIgnored 
                      ? 'bg-amber-50 text-amber-700 border-amber-200' 
                      : 'bg-gray-50 text-gray-600 border-gray-200 hover:bg-gray-100'
                  }`}
                  title={hideIgnored ? "Mostrar todos os elementos" : "Ocultar elementos ignorados"}
                >
                  {hideIgnored ? <Eye className="w-3.5 h-3.5" /> : <EyeOff className="w-3.5 h-3.5" />}
                  <span>{hideIgnored ? "Ocultar Ignorados" : "Ignorados Visíveis"}</span>
                </button>
              </div>

            {selectedIds.size > 0 && (
              <div className="flex items-center justify-between pt-4 border-t border-gray-100">
                <div className="flex items-center gap-2">
                  <span className="bg-indigo-100 text-indigo-700 px-2 py-1 rounded text-xs font-bold">
                    {selectedIds.size} selecionados
                  </span>
                  <button onClick={() => setSelectedIds(new Set())} className="text-xs text-gray-400 hover:text-gray-600">
                    Limpar
                  </button>
                </div>
                <div className="flex items-center gap-2">
                  <div className="flex items-center gap-1 bg-indigo-50 text-indigo-700 px-3 py-1.5 rounded-lg text-xs font-bold transition-colors border border-indigo-200">
                    <span className="mr-1">Alterar para:</span>
                    <select 
                      className="bg-transparent border-none focus:outline-none cursor-pointer"
                      value=""
                      onChange={(e) => {
                        if (e.target.value) handleBulkUpdateType(e.target.value);
                      }}
                    >
                      <option value="" disabled>Selecionar...</option>
                      {ORDERED_PII_TYPES.map(type => (
                        <option key={type} value={type}>{type}</option>
                      ))}
                    </select>
                  </div>
                  <button 
                    onClick={handleValidateSelected}
                    className="flex items-center gap-1 bg-green-50 text-green-700 hover:bg-green-100 px-3 py-1.5 rounded-lg text-xs font-bold transition-colors border border-green-200"
                  >
                    <Check className="w-3.5 h-3.5" />
                    Validar
                  </button>
                  <button 
                    onClick={handleGroupSelected}
                    className="flex items-center gap-1 bg-indigo-50 text-indigo-700 hover:bg-indigo-100 px-3 py-1.5 rounded-lg text-xs font-bold transition-colors border border-indigo-200"
                  >
                    <Layers className="w-3.5 h-3.5" />
                    Agrupar / Associar
                  </button>
                  {copiedPseudonym && (
                    <button 
                      onClick={handlePastePseudonym}
                      className="flex items-center gap-1 bg-amber-50 text-amber-700 hover:bg-amber-100 px-3 py-1.5 rounded-lg text-xs font-bold transition-colors border border-amber-200"
                    >
                      <Link className="w-3.5 h-3.5" />
                      Colar
                    </button>
                  )}
                  <button 
                    onClick={handleIgnoreSelected}
                    className="flex items-center gap-1 bg-gray-50 text-gray-700 hover:bg-gray-100 px-3 py-1.5 rounded-lg text-xs font-bold transition-colors border border-gray-200"
                  >
                    <EyeOff className="w-3.5 h-3.5" />
                    Ignorar
                  </button>
                </div>
              </div>
            )}
          </div>

          <div className="flex items-center justify-between">
            <div className="flex flex-col gap-1">
              <div className="flex items-center gap-4">
                <h2 className="text-lg font-bold">Elementos Detetados ({filteredEntities.length})</h2>
                {filteredEntities.length > 0 && (
                  <button 
                    onClick={toggleSelectAllFiltered}
                    className="text-xs font-medium text-indigo-600 hover:text-indigo-800 flex items-center gap-1 bg-indigo-50 px-2 py-1 rounded-md transition-colors"
                  >
                    {filteredEntities.every(e => selectedIds.has(e.id)) ? 'Desmarcar Todos' : 'Selecionar Todos'}
                  </button>
                )}
              </div>
              {pseudonymAnalysis.conflicts.size > 0 && (
                <div className="flex flex-col gap-1">
                  <div className="flex items-center gap-2 text-red-600 animate-pulse">
                    <AlertCircle className="w-4 h-4" />
                    <span className="text-xs font-bold">Atenção: Existem {pseudonymAnalysis.conflicts.size} pseudónimos duplicados para nomes diferentes!</span>
                  </div>
                  <button 
                    onClick={handleFixAllConflicts}
                    className="text-[10px] font-bold text-red-700 bg-red-50 hover:bg-red-100 px-2 py-0.5 rounded border border-red-200 w-fit transition-colors"
                  >
                    Corrigir Todos Automaticamente
                  </button>
                </div>
              )}
            </div>
            <div className="flex items-center gap-2">
              {copiedPseudonym && (
                <div className="flex items-center gap-2 bg-indigo-50 text-indigo-700 px-3 py-1 rounded-full text-[10px] font-bold border border-indigo-100">
                  <Copy className="w-3 h-3" />
                  <span>Copiado: {copiedPseudonym}</span>
                  <button onClick={() => setCopiedPseudonym(null)} className="hover:text-indigo-900">
                    <X className="w-3 h-3" />
                  </button>
                </div>
              )}
            </div>
          </div>

          <div className="space-y-4">
            {sortedGroupedEntries.map(([groupId, group]: [string, PIIEntity[]]) => {
              const isCollapsed = collapsedGroups.has(groupId);
              const isManualGroup = groupId.startsWith('manual-group-') || groupId.startsWith('group-');
              const isProcessed = group.every(e => e.treated || e.ignored);

              return (
                <motion.div 
                  key={groupId}
                  layout
                  initial={{ opacity: 0, y: 20 }}
                  animate={{ opacity: 1, y: 0 }}
                  className={`bg-white rounded-2xl border shadow-sm overflow-hidden transition-all duration-300 ${
                    group.some(e => e.isNewlyIdentified) 
                      ? 'border-indigo-400 ring-2 ring-indigo-50/50 shadow-indigo-100/30' 
                      : 'border-gray-200'
                  }`}
                >
                  {group.length > 1 || isManualGroup ? (
                    <div 
                      className={`p-4 flex items-center justify-between cursor-pointer transition-colors ${
                        isCollapsed ? 'bg-gray-50/80' : 'bg-gray-50/50 border-b border-gray-100'
                      }`}
                      onClick={() => {
                        setCollapsedGroups(prev => {
                          const next = new Set(prev);
                          if (next.has(groupId)) next.delete(groupId);
                          else next.add(groupId);
                          return next;
                        });
                      }}
                    >
                      <div className="flex items-center gap-3">
                        <div 
                          onClick={(e) => {
                            e.stopPropagation();
                            toggleSelectGroup(group);
                          }}
                        >
                          <input 
                            type="checkbox"
                            checked={group.every(e => selectedIds.has(e.id))}
                            onChange={() => {}} // Handled by parent div click or stopPropagation
                            className="w-4 h-4 rounded border-gray-300 text-indigo-600 focus:ring-indigo-500"
                          />
                        </div>
                        <div className="bg-indigo-100 text-indigo-700 p-1.5 rounded-lg">
                          <Layers className="w-4 h-4" />
                        </div>
                        <div>
                          <h3 className="text-sm font-bold flex items-center gap-2">
                            Grupo: {group[0].original}
                            {group.some(e => e.isNewlyIdentified) && (
                              <span className="inline-flex items-center gap-1 px-1.5 py-0.5 rounded-full text-[9px] font-bold bg-indigo-100 text-indigo-700 animate-pulse border border-indigo-200 uppercase tracking-widest leading-none">
                                Novo
                              </span>
                            )}
                            {isProcessed && <CheckCircle2 className="w-3.5 h-3.5 text-green-500" />}
                            {groupId.startsWith('manual-group-') && (
                              <div className="flex items-center gap-1">
                                <span className="text-[8px] bg-amber-50 text-amber-600 px-1 py-0.5 rounded border border-amber-200 font-bold uppercase tracking-tighter">Manual</span>
                                <button
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    handleUnlockGroup(groupId);
                                  }}
                                  className="p-1 hover:bg-white rounded text-amber-400 hover:text-amber-600 transition-colors"
                                  title="Desbloquear grupo (permitir re-agrupamento automático)"
                                >
                                  <RotateCw className="w-3 h-3" />
                                </button>
                              </div>
                            )}
                            <div className="flex items-center gap-1">
                              <button
                                onClick={(e) => {
                                  e.stopPropagation();
                                  handleDissolveGroup(groupId);
                                }}
                                className="p-1 hover:bg-white rounded text-red-400 hover:text-red-600 transition-colors"
                                title="Dissolver grupo (separar todos os elementos)"
                              >
                                <Unlink className="w-3 h-3" />
                              </button>
                              {group.some(e => (e.type === 'NOME' || e.type === 'AUTOR' || e.type === 'JUIZ') && (e.original.includes(' e ') || e.original.includes('  ') || e.original.split(' ').length >= 4)) && (
                                <div className="flex items-center bg-gray-100 rounded p-0.5">
                                  <button
                                    onClick={(e) => {
                                      e.stopPropagation();
                                      handleSplitGroup(groupId, 'e');
                                    }}
                                    className="p-1 hover:bg-white rounded text-indigo-400 hover:text-indigo-600 transition-colors"
                                    title="Dividir por 'e'"
                                  >
                                    <Scissors className="w-3 h-3" />
                                  </button>
                                  <button
                                    onClick={(e) => {
                                      e.stopPropagation();
                                      handleSplitGroup(groupId, 'space');
                                    }}
                                    className="p-1 hover:bg-white rounded text-indigo-400 hover:text-indigo-600 transition-colors"
                                    title="Dividir por espaços/nomes longos"
                                  >
                                    <Type className="w-3 h-3" />
                                  </button>
                                </div>
                              )}
                            </div>
                          </h3>
                          <div className="text-xs text-gray-500 flex items-center gap-1">
                            {group.length} ocorrências • {group[0].pseudonym}
                            {pseudonymAnalysis.conflicts.has(group[0].pseudonym.trim()) && (
                              <div className="flex items-center gap-1">
                                <AlertCircle 
                                  className="w-3 h-3 text-red-500" 
                                  title={`Aviso: Este pseudónimo também está a ser usado para: ${Array.from(pseudonymAnalysis.pToO.get(group[0].pseudonym.trim()) || []).filter(o => o !== group[0].original.toLowerCase().trim()).join(', ')}`} 
                                />
                                <button 
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    handleFixSingleConflict(group[0]);
                                  }}
                                  className="text-[8px] bg-red-50 text-red-600 hover:bg-red-100 px-1 py-0.5 rounded border border-red-200 font-bold"
                                >
                                  Corrigir
                                </button>
                              </div>
                            )}
                          </div>
                        </div>
                      </div>
                      <div className="flex items-center gap-2">
                        <select 
                          className={`text-[10px] font-bold bg-white border border-gray-200 rounded px-1 py-0.5 focus:outline-none focus:ring-1 focus:ring-indigo-500 ${group.some(e => e.type !== group[0].type) ? 'text-amber-600 border-amber-200' : ''}`}
                          value={group.every(e => e.type === group[0].type) ? group[0].type : 'MIXED'}
                          onClick={(e) => e.stopPropagation()}
                          onChange={(e) => {
                            e.stopPropagation();
                            handleUpdateGroupType(groupId, e.target.value);
                          }}
                          title="Alterar categoria de todo o grupo"
                        >
                          {group.some(e => e.type !== group[0].type) && (
                            <option value="MIXED" disabled>VÁRIOS</option>
                          )}
                          {ORDERED_PII_TYPES.map(type => (
                            <option key={type} value={type}>{type}</option>
                          ))}
                        </select>
                        <button 
                          onClick={(e) => {
                            e.stopPropagation();
                            handleCopyPseudonym(group[0].pseudonym);
                          }}
                          className="p-2 hover:bg-white rounded-lg text-gray-400 hover:text-indigo-600 transition-colors"
                          title="Copiar Pseudónimo"
                        >
                          <History className="w-4 h-4" />
                        </button>
                        {isCollapsed ? <ChevronDown className="w-4 h-4 text-gray-400" /> : <ChevronRight className="w-4 h-4 text-gray-400" />}
                      </div>
                    </div>
                  ) : null}

                  {!isCollapsed && (
                    <div className="divide-y divide-gray-100">
                      {group.map((entity: PIIEntity) => (
                        <div 
                          key={entity.id} 
                          onClick={() => {
                            if (reviewMode) {
                              setSelectedEntityId(entity.id);
                              if (entity.fileIds && entity.fileIds.length > 0) {
                                setSelectedFileId(entity.fileIds[0]);
                              }
                            }
                          }}
                          className={`p-4 flex items-center gap-4 transition-all duration-300 cursor-pointer ${
                            selectedIds.has(entity.id) ? 'bg-indigo-50/45 ring-1 ring-indigo-200' : 
                            selectedEntityId === entity.id ? 'bg-indigo-100 border-l-4 border-indigo-600' : 
                            entity.treated ? 'bg-emerald-50 border-l-4 border-emerald-500 hover:bg-emerald-100/60' :
                            entity.ignored ? 'bg-amber-50/70 border-l-4 border-amber-500 hover:bg-amber-100/60' :
                            entity.isNewlyIdentified ? 'bg-indigo-50/35 border-l-4 border-indigo-500 hover:bg-indigo-50/50' :
                            'hover:bg-gray-50/50 border-l-4 border-transparent'
                          }`}
                        >
                          <input 
                            type="checkbox"
                            checked={selectedIds.has(entity.id)}
                            onChange={() => toggleEntitySelection(entity.id)}
                            className="w-4 h-4 rounded border-gray-300 text-indigo-600 focus:ring-indigo-500"
                          />
                          
                          <div className="flex-1 min-w-0">
                            <div className="flex items-center gap-2 mb-1">
                              <select 
                                value={entity.type}
                                onClick={(e) => e.stopPropagation()}
                                onChange={(e) => {
                                  const newType = e.target.value;
                                  setEntities(prev => {
                                    const next = prev.map(item => item.id === entity.id ? { ...item, type: newType } : item);
                                    showToast(`Tipo de "${entity.original}" alterado para ${newType}`, "success");
                                    setTimeout(() => pushToHistory(next, files), 0);
                                    return next;
                                  });
                                }}
                                className="px-1.5 py-0.5 rounded text-[10px] font-bold uppercase tracking-wider focus:outline-none cursor-pointer border-none shadow-sm font-sans"
                                style={{ 
                                  backgroundColor: PII_COLORS[entity.type]?.hex,
                                  color: PII_COLORS[entity.type]?.textHex || '#000000',
                                }}
                                title="Clique para recategorizar com um clique direto"
                              >
                                {ORDERED_PII_TYPES.map(type => (
                                  <option key={type} value={type} className="bg-white text-gray-800 font-sans font-bold text-[10px] uppercase">
                                    {type}
                                  </option>
                                ))}
                              </select>
                              <span className="text-sm font-semibold truncate">
                                {['NIF', 'CC', 'PASSPORT', 'IBAN'].includes(entity.type) && entity.contextSnippet ? (
                                  <span className="text-xs italic text-gray-500 font-normal">
                                    ...{entity.contextSnippet}...
                                  </span>
                                ) : (
                                  entity.original
                                )}
                              </span>
                              {entity.isNewlyIdentified && (
                                <span className="inline-flex items-center gap-1 px-1.5 py-0.5 rounded-full text-[9px] font-bold bg-indigo-100 text-indigo-700 animate-pulse border border-indigo-200 uppercase tracking-widest leading-none">
                                  Novo
                                </span>
                              )}
                            </div>
                              <div className="flex items-center gap-2">
                                <input 
                                  type="text"
                                  value={entity.pseudonym}
                                  onChange={(e) => updatePseudonym(entity.id, e.target.value)}
                                  className={`text-xs font-mono bg-gray-100 px-2 py-1 rounded border outline-none w-32 ${
                                    pseudonymAnalysis.conflicts.has(entity.pseudonym.trim()) 
                                      ? 'border-red-300 bg-red-50 focus:border-red-500' 
                                      : 'border-transparent focus:border-indigo-300 focus:bg-white'
                                  }`}
                                />
                                {pseudonymAnalysis.conflicts.has(entity.pseudonym.trim()) && (
                                  <div className="flex items-center gap-1">
                                    <AlertCircle 
                                      className="w-3.5 h-3.5 text-red-500" 
                                      title={`Aviso: Este pseudónimo também está a ser usado para: ${Array.from(pseudonymAnalysis.pToO.get(entity.pseudonym.trim()) || []).filter(o => o !== entity.original.toLowerCase().trim()).join(', ')}`} 
                                    />
                                    <button 
                                      onClick={() => handleFixSingleConflict(entity)}
                                      className="text-[9px] bg-red-50 text-red-600 hover:bg-red-100 px-1.5 py-0.5 rounded border border-red-200 font-bold"
                                    >
                                      Corrigir
                                    </button>
                                  </div>
                                )}
                                {entity.treated && <CheckCircle2 className="w-3 h-3 text-green-500" />}
                                {entity.ignored && <EyeOff className="w-3 h-3 text-gray-400" />}
                              </div>
                          </div>

                          <div className="flex items-center gap-1">
                            <button 
                              onClick={() => addToGlobalKnowledge(entity.original)}
                              className="p-2 hover:bg-white rounded-lg text-gray-400 hover:text-red-600 transition-colors"
                              title="Adicionar às exceções globais"
                            >
                              <Shield className="w-4 h-4" />
                            </button>
                            <button 
                              onClick={() => setEditingEntity(entity)}
                              className="p-2 hover:bg-white rounded-lg text-gray-400 hover:text-indigo-600 transition-colors"
                              title="Ver contexto e editar"
                            >
                              <Eye className="w-4 h-4" />
                            </button>
                            <button 
                              onClick={() => setEntities(prev => prev.map(e => e.id === entity.id ? { ...e, treated: !e.treated, ignored: false, enabled: true } : e))}
                              className={`p-2 rounded-lg transition-colors ${entity.treated ? 'text-green-600 bg-green-50' : 'text-gray-400 hover:bg-white hover:text-green-600'}`}
                              title="Validar"
                            >
                              <Check className="w-4 h-4" />
                            </button>
                            <button 
                              onClick={() => setEntities(prev => prev.map(e => e.id === entity.id ? { ...e, ignored: !e.ignored, treated: false } : e))}
                              className={`p-2 rounded-lg transition-colors ${entity.ignored ? 'text-red-600 bg-red-50' : 'text-gray-400 hover:bg-white hover:text-red-600'}`}
                              title="Ignorar"
                            >
                              <EyeOff className="w-4 h-4" />
                            </button>
                            <button 
                              onClick={(e) => {
                                e.stopPropagation();
                                handleIgnoreSimilar(entity);
                              }}
                              className="p-2 hover:bg-amber-100 rounded-lg text-amber-500 hover:text-amber-800 transition-colors"
                              title="Ignorar todos os semelhantes (Ação Coletiva)"
                            >
                              <XCircle className="w-4 h-4" />
                            </button>
                            {entity.original.includes(' ') && (
                              <button 
                                onClick={() => handleSplitEntity(entity)}
                                className="p-2 hover:bg-white rounded-lg text-gray-400 hover:text-indigo-600 transition-colors"
                                title="Dividir Elemento"
                              >
                                <Scissors className="w-4 h-4" />
                              </button>
                            )}
                            <button 
                              onClick={() => handleCopyPseudonym(entity.pseudonym)}
                              className="p-2 hover:bg-white rounded-lg text-gray-400 hover:text-indigo-600 transition-colors"
                              title="Copiar"
                            >
                              <History className="w-4 h-4" />
                            </button>
                            <button 
                              onClick={() => {
                                setSelectedIds(new Set([entity.id]));
                                setShowMergeModal(true);
                              }}
                              className="p-2 hover:bg-white rounded-lg text-gray-400 hover:text-indigo-600 transition-colors"
                              title="Unir a Grupo"
                            >
                              <Link className="w-4 h-4" />
                            </button>
                          </div>
                        </div>
                      ))}
                    </div>
                  )}
                </motion.div>
              );
            })}
          </div>
        </>
      )}
    </div>
  </div>
</main>

      {/* Export Modal */}
      <AnimatePresence>
        {showExportModal && (
          <div className="fixed inset-0 z-[100] flex items-center justify-center p-4">
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setShowExportModal(false)}
              className="absolute inset-0 bg-black/40 backdrop-blur-sm"
            />
            <motion.div 
              initial={{ scale: 0.95, opacity: 0, y: 20 }}
              animate={{ scale: 1, opacity: 1, y: 0 }}
              exit={{ scale: 0.95, opacity: 0, y: 20 }}
              className="bg-white rounded-2xl shadow-2xl w-full max-w-5xl overflow-hidden relative z-10 border border-gray-100 flex flex-col lg:flex-row h-[85vh] md:h-[650px]"
            >
              {/* Left Column: Configuration settings */}
              <div className="w-full lg:w-[420px] flex flex-col h-full bg-white shrink-0 border-r border-gray-100 justify-between">
                <div>
                  <div className="p-6 border-b border-gray-100 flex items-center justify-between bg-gray-50/50">
                    <div className="flex items-center gap-3">
                      <div className="w-10 h-10 rounded-xl bg-indigo-100 flex items-center justify-center text-indigo-600">
                        <Download className="w-5 h-5" />
                      </div>
                      <div>
                        <h3 className="text-lg font-bold text-gray-900">Configurar Exportação</h3>
                        <p className="text-xs text-gray-500 font-sans">Escolha o formato e opções de saída</p>
                      </div>
                    </div>
                    <button onClick={() => setShowExportModal(false)} className="p-2 hover:bg-gray-200 rounded-full transition-colors lg:hidden">
                      <X className="w-5 h-5 text-gray-400" />
                    </button>
                  </div>

                  <div className="p-6 space-y-6 overflow-y-auto max-h-[380px] lg:max-h-[480px]">
                    <div>
                      <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest mb-3 block">Formato de Saída</label>
                      <div className="grid grid-cols-3 gap-3">
                        {[
                          { id: 'pdf', label: 'PDF', icon: FileText, desc: 'Preserva layout' },
                          { id: 'docx', label: 'Word', icon: FileText, desc: 'Editável' },
                          { id: 'txt', label: 'Texto', icon: Type, desc: 'Simples' }
                        ].map(format => (
                          <button
                            key={format.id}
                            onClick={() => setExportSettings(prev => ({ ...prev, format: format.id as any }))}
                            className={`flex flex-col items-center gap-2 p-3 rounded-xl border-2 transition-all ${
                              exportSettings.format === format.id 
                                ? 'border-indigo-500 bg-indigo-50/50 text-indigo-700' 
                                : 'border-gray-100 hover:border-gray-200 text-gray-500'
                            }`}
                          >
                            <format.icon className={`w-6 h-6 ${exportSettings.format === format.id ? 'text-indigo-600' : 'text-gray-400'}`} />
                            <span className="text-xs font-bold font-sans">{format.label}</span>
                          </button>
                        ))}
                      </div>
                    </div>

                    <div className="space-y-4 font-sans">
                      <label className="text-[10px] font-bold text-gray-400 uppercase tracking-widest block">Opções Adicionais</label>
                      
                      <div className="space-y-3">
                        <label className="flex items-center justify-between p-3 rounded-xl border border-gray-100 hover:bg-gray-50 cursor-pointer transition-colors">
                          <div className="flex items-center gap-3">
                            <div className="w-8 h-8 rounded-lg bg-blue-50 flex items-center justify-center text-blue-600">
                              <Layers className="w-4 h-4" />
                            </div>
                            <div>
                              <p className="text-xs font-bold text-gray-700">Preservar Formatação</p>
                              <p className="text-[10px] text-gray-500">Negritos, itálicos e tabelas</p>
                            </div>
                          </div>
                          <input 
                            type="checkbox" 
                            checked={exportSettings.preserveFormatting}
                            onChange={(e) => setExportSettings(prev => ({ ...prev, preserveFormatting: e.target.checked }))}
                            className="w-4 h-4 text-indigo-600 rounded border-gray-300 focus:ring-indigo-500"
                          />
                        </label>

                        <label className="flex items-center justify-between p-3 rounded-xl border border-gray-100 hover:bg-gray-50 cursor-pointer transition-colors">
                          <div className="flex items-center gap-3">
                            <div className="w-8 h-8 rounded-lg bg-orange-50 flex items-center justify-center text-orange-600">
                              <Scissors className="w-4 h-4" />
                            </div>
                            <div>
                              <p className="text-xs font-bold text-gray-700">Remover Cabeçalhos/Rodapés</p>
                              <p className="text-[10px] text-gray-500">Tenta limpar elementos repetitivos</p>
                            </div>
                          </div>
                          <input 
                            type="checkbox" 
                            checked={exportSettings.removeHeadersFooters}
                            onChange={(e) => setExportSettings(prev => ({ ...prev, removeHeadersFooters: e.target.checked }))}
                            className="w-4 h-4 text-indigo-600 rounded border-gray-300 focus:ring-indigo-500"
                          />
                        </label>

                        <label className="flex items-center justify-between p-3 rounded-xl border border-gray-100 hover:bg-gray-50 cursor-pointer transition-colors">
                          <div className="flex items-center gap-3">
                            <div className="w-8 h-8 rounded-lg bg-purple-50 flex items-center justify-center text-purple-600">
                              <List className="w-4 h-4" />
                            </div>
                            <div>
                              <p className="text-xs font-bold text-gray-700">Notas de Rodapé no Fim</p>
                              <p className="text-[10px] text-gray-500">Agrupa todas as notas no final</p>
                            </div>
                          </div>
                          <input 
                            type="checkbox" 
                            checked={exportSettings.footnotesAtEnd}
                            onChange={(e) => setExportSettings(prev => ({ ...prev, footnotesAtEnd: e.target.checked }))}
                            className="w-4 h-4 text-indigo-600 rounded border-gray-300 focus:ring-indigo-500"
                          />
                        </label>

                        <label className="flex items-center justify-between p-3 rounded-xl border border-gray-100 hover:bg-gray-50 cursor-pointer transition-colors">
                          <div className="flex items-center gap-3">
                            <div className="w-8 h-8 rounded-lg bg-emerald-50 flex items-center justify-center text-emerald-600">
                              <Shield className="w-4 h-4" />
                            </div>
                            <div>
                              <p className="text-xs font-bold text-gray-700">Limpeza de Metadados Ocultos</p>
                              <p className="text-[10px] text-gray-500 font-sans">Remove propriedades de gravação (RGPD)</p>
                            </div>
                          </div>
                          <input 
                            type="checkbox" 
                            checked={exportSettings.sanitizeMetadata}
                            onChange={(e) => setExportSettings(prev => ({ ...prev, sanitizeMetadata: e.target.checked }))}
                            className="w-4 h-4 text-indigo-600 rounded border-gray-300 focus:ring-indigo-500"
                          />
                        </label>
                      </div>
                    </div>
                  </div>
                </div>

                <div className="p-6 bg-gray-50 border-t border-gray-100 flex gap-3 shrink-0">
                  <button 
                    onClick={() => setShowExportModal(false)}
                    className="flex-1 px-4 py-2.5 rounded-xl border border-gray-200 text-gray-600 font-bold text-xs hover:bg-white transition-all font-sans"
                  >
                    Cancelar
                  </button>
                  <button 
                    onClick={handleExport}
                    className="flex-[2] px-4 py-2.5 rounded-xl bg-indigo-600 text-white font-bold text-xs hover:bg-indigo-700 shadow-lg shadow-indigo-200 transition-all flex items-center justify-center gap-2 font-sans"
                  >
                    <Download className="w-4 h-4" />
                    <span>Exportar Agora</span>
                  </button>
                </div>
              </div>

              {/* Right Column: Real-time side-by-side anonymization preview */}
              <div className="flex-1 bg-slate-50 flex flex-col h-full min-w-0 overflow-hidden relative border-t lg:border-t-0 border-gray-100">
                <div className="absolute top-4 right-4 z-50">
                  <button onClick={() => setShowExportModal(false)} className="p-2 bg-white/85 backdrop-blur-xs border border-gray-200 hover:bg-gray-100 rounded-full transition-colors shadow-sm hidden lg:block">
                    <X className="w-5 h-5 text-gray-500" />
                  </button>
                </div>

                <div className="p-6 border-b border-gray-200/60 bg-white flex items-center justify-between shrink-0 font-sans">
                  <div className="flex items-center gap-2">
                    <span className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse" />
                    <span className="text-xs font-black text-gray-700 uppercase tracking-widest">Pré-visualização em Tempo Real</span>
                  </div>
                  {files.length > 1 ? (
                    <select 
                      value={selectedFileId || ""}
                      onChange={(e) => setSelectedFileId(e.target.value)}
                      className="text-[10px] font-bold bg-gray-50 border border-gray-200 focus:border-indigo-400 rounded-lg px-2.5 py-1.5 focus:outline-none focus:ring-1 focus:ring-indigo-500/20 text-gray-600"
                    >
                      {files.map(f => (
                        <option key={f.id} value={f.id}>{f.name}</option>
                      ))}
                    </select>
                  ) : (
                    <span className="text-[10px] font-mono text-gray-400 lg:pr-10 truncate max-w-[200px]">{files.find(f => f.id === selectedFileId)?.name || files[0]?.name}</span>
                  )}
                </div>

                <div className="flex-1 p-6 overflow-y-auto font-mono text-xs text-gray-800 whitespace-pre-wrap leading-relaxed select-text bg-white m-4 rounded-xl border border-gray-200/60 shadow-xs relative">
                  {(files.find(f => f.id === selectedFileId) || files[0]) ? (
                    <HighlightText 
                      text={anonymizeText((files.find(f => f.id === selectedFileId) || files[0])?.content || "", entities)} 
                      entities={entities} 
                      mode="anonymized" 
                    />
                  ) : (
                    <div className="flex flex-col items-center justify-center h-full text-center text-gray-400">
                      <FileText className="w-8 h-8 mb-2 text-gray-300" />
                      <p className="text-xs font-semibold">Nenhum ficheiro para pré-visualizar.</p>
                    </div>
                  )}
                </div>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {pendingManualTerm && (
        <div 
          className="fixed z-[100] bg-white shadow-2xl border border-gray-200 rounded-xl p-3 flex flex-col gap-2 animate-in fade-in zoom-in duration-200 min-w-[220px]"
          style={{ left: Math.min(window.innerWidth - 240, pendingManualTerm.x), top: pendingManualTerm.y + 15 }}
        >
          <div className="flex items-center justify-between mb-1">
            <div className="flex flex-col">
              <span className="text-[10px] font-bold text-gray-400 uppercase tracking-wider">Adicionar como:</span>
              <span className="text-[9px] text-indigo-500 font-medium truncate max-w-[150px]">"{pendingManualTerm.text}"</span>
            </div>
            <button onClick={() => setPendingManualTerm(null)} className="p-1 hover:bg-gray-100 rounded-full">
              <X className="w-3 h-3 text-gray-400" />
            </button>
          </div>
          <div className="grid grid-cols-2 gap-1.5 max-h-[200px] overflow-y-auto pr-1">
            {ORDERED_PII_TYPES.map(type => {
              const color = PII_COLORS[type];
              return (
                <button 
                  key={type}
                  onClick={() => {
                    setGlobalKnowledge(prev => ({ ...prev, [pendingManualTerm.text.toLowerCase().trim()]: type }));
                    
                    // Also add to current entities as a treated entity
                    const newEntity: PIIEntity = {
                      id: `manual-${Date.now()}-${generateId()}`,
                      original: pendingManualTerm.text,
                      type: type,
                      pseudonym: getNextPseudonym(type, entities),
                      enabled: true,
                      treated: true,
                      fileIds: [selectedFileId || 'manual']
                    };
                    setEntities(prev => [...prev, newEntity]);
                    
                    setPendingManualTerm(null);
                    showToast(`"${pendingManualTerm.text}" adicionado como ${type} e VALIDADO.`, "success");
                  }}
                  className="px-2 py-1.5 text-[10px] font-bold rounded transition-all hover:scale-105 text-center shadow-sm border border-black/5"
                  style={{ backgroundColor: color.hex, color: color.textHex }}
                >
                  {type}
                </button>
              );
            })}
            <button 
              onClick={() => {
                setGlobalKnowledge(prev => ({ ...prev, [pendingManualTerm.text.toLowerCase().trim()]: 'HEADER' }));
                setPendingManualTerm(null);
                showToast(`"${pendingManualTerm.text}" marcado como CABEÇALHO/RODAPÉ.`, "success");
              }}
              className="col-span-2 px-2 py-1.5 bg-indigo-50 text-indigo-700 text-[10px] font-bold rounded hover:bg-indigo-100 transition-all text-center border border-indigo-200 mt-1"
            >
              CABEÇALHO / RODAPÉ
            </button>
            <button 
              onClick={() => {
                setGlobalKnowledge(prev => ({ ...prev, [pendingManualTerm.text.toLowerCase().trim()]: 'EXCECAO' }));
                setPendingManualTerm(null);
                showToast(`"${pendingManualTerm.text}" adicionado às EXCEÇÕES.`, "success");
              }}
              className="col-span-2 px-2 py-1.5 bg-gray-100 text-gray-600 text-[10px] font-bold rounded hover:bg-gray-200 transition-all text-center border border-gray-200"
            >
              EXCEÇÃO
            </button>
          </div>
        </div>
      )}

      {/* Floating Action Bar */}
      <AnimatePresence>
        {selectedIds.size > 0 && (
          <motion.div 
            initial={{ y: 100, opacity: 0 }}
            animate={{ y: 0, opacity: 1 }}
            exit={{ y: 100, opacity: 0 }}
            className="fixed bottom-8 left-1/2 -translate-x-1/2 z-40 bg-white border border-gray-200 shadow-2xl rounded-2xl px-6 py-4 flex items-center gap-6"
          >
            <div className="flex items-center gap-3 border-r border-gray-100 pr-6">
              <span className="bg-indigo-600 text-white w-6 h-6 rounded-full flex items-center justify-center text-xs font-bold">
                {selectedIds.size}
              </span>
              <span className="text-sm font-medium text-gray-600">Selecionados</span>
              <button onClick={() => setSelectedIds(new Set())} className="text-xs text-gray-400 hover:text-indigo-600 font-medium">
                Limpar
              </button>
            </div>

            <div className="flex items-center gap-2">
              <button 
                onClick={handleValidateSelected}
                className="flex items-center gap-2 bg-green-600 hover:bg-green-700 text-white px-4 py-2 rounded-xl text-sm font-bold transition-all shadow-lg shadow-green-200"
              >
                <Check className="w-4 h-4" />
                <span>Validar</span>
              </button>
              <button 
                onClick={() => handleBulkReclassAndValidateSelected('NOME')}
                className="flex items-center gap-2 bg-yellow-500 hover:bg-yellow-600 text-white px-4 py-2 rounded-xl text-sm font-bold transition-all shadow-lg shadow-yellow-100"
                title="Tratar todos os selecionados como Nome"
              >
                <User className="w-4 h-4" />
                <span>Nome</span>
              </button>
              <button 
                onClick={() => handleBulkReclassAndValidateSelected('JUIZ')}
                className="flex items-center gap-2 bg-purple-600 hover:bg-purple-700 text-white px-4 py-2 rounded-xl text-sm font-bold transition-all shadow-lg shadow-purple-100"
                title="Tratar todos os selecionados como Juiz"
              >
                <Scale className="w-4 h-4" />
                <span>Juiz</span>
              </button>
              <button 
                onClick={() => handleBulkReclassAndValidateSelected('AUTOR')}
                className="flex items-center gap-2 bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded-xl text-sm font-bold transition-all shadow-lg shadow-blue-100"
                title="Tratar todos os selecionados como Autor"
              >
                <Briefcase className="w-4 h-4" />
                <span>Autor</span>
              </button>
              <button 
                onClick={handleGroupSelected}
                className="flex items-center gap-2 bg-indigo-600 hover:bg-indigo-700 text-white px-4 py-2 rounded-xl text-sm font-bold transition-all shadow-lg shadow-indigo-200"
              >
                <Layers className="w-4 h-4" />
                <span>Agrupar</span>
              </button>
              <button 
                onClick={handleAddSelectedToExceptions}
                className="flex items-center gap-2 bg-amber-600 hover:bg-amber-700 text-white px-4 py-2 rounded-xl text-sm font-bold transition-all shadow-lg shadow-amber-200"
                title="Adicionar à Safelist e ignorar em todos os documentos"
              >
                <Shield className="w-4 h-4" />
                <span>Exceção</span>
              </button>
              <button 
                onClick={handleAddSelectedToColetiva}
                className="flex items-center gap-2 bg-stone-600 hover:bg-stone-700 text-white px-4 py-2 rounded-xl text-sm font-bold transition-all shadow-lg shadow-stone-200"
                title="Marcar como Pessoa Coletiva e ignorar em todos os documentos"
              >
                <Building2 className="w-4 h-4" />
                <span>Coletiva</span>
              </button>
              {copiedPseudonym && (
                <button 
                  onClick={handlePastePseudonym}
                  className="flex items-center gap-2 bg-amber-500 hover:bg-amber-600 text-white px-4 py-2 rounded-xl text-sm font-bold transition-all shadow-lg shadow-amber-200"
                >
                  <Link className="w-4 h-4" />
                  <span>Colar</span>
                </button>
              )}
              <button 
                onClick={handleIgnoreSelected}
                className="flex items-center gap-2 bg-gray-600 hover:bg-gray-700 text-white px-4 py-2 rounded-xl text-sm font-bold transition-all shadow-lg shadow-gray-200"
              >
                <EyeOff className="w-4 h-4" />
                <span>Ignorar</span>
              </button>
            </div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Toast Notification */}
      <AnimatePresence>
        {toast && (
          <motion.div
            initial={{ opacity: 0, y: 50, x: '-50%' }}
            animate={{ opacity: 1, y: 0, x: '-50%' }}
            exit={{ opacity: 0, y: 50, x: '-50%' }}
            className={`fixed bottom-24 left-1/2 z-[100] px-6 py-3 rounded-2xl shadow-2xl flex items-center gap-3 border ${
              toast.type === 'success' ? 'bg-green-50 border-green-100 text-green-800' :
              toast.type === 'error' ? 'bg-red-50 border-red-100 text-red-800' :
              'bg-indigo-50 border-indigo-100 text-indigo-800'
            }`}
          >
            {toast.type === 'success' ? <CheckCircle2 className="w-5 h-5 text-green-600" /> :
             toast.type === 'error' ? <AlertCircle className="w-5 h-5 text-red-600" /> :
             <Shield className="w-5 h-5 text-indigo-600" />}
            <span className="font-bold text-sm">{toast.message}</span>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Merge Modal */}
      <AnimatePresence>
        {showMergeModal && (
          <div className="fixed inset-0 z-[60] flex items-center justify-center p-4">
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setShowMergeModal(false)}
              className="absolute inset-0 bg-black/60 backdrop-blur-sm"
            />
            <motion.div 
              initial={{ scale: 0.9, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.9, opacity: 0 }}
              className="relative bg-white w-full max-w-md rounded-3xl shadow-2xl overflow-hidden"
            >
              <div className="p-6 border-b border-gray-100 flex items-center justify-between">
                <h3 className="text-xl font-bold">Unir a Grupo Existente</h3>
                <button onClick={() => {
                  setShowMergeModal(false);
                  setMergeSearch('');
                  setShowAllMergeOptions(false);
                }} className="p-2 hover:bg-gray-100 rounded-full transition-colors">
                  <X className="w-5 h-5" />
                </button>
              </div>
              
              <div className="p-4 border-b border-gray-50 bg-gray-50/50">
                <div className="relative">
                  <Search className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
                  <input 
                    type="text"
                    placeholder="Procurar grupo..."
                    value={mergeSearch}
                    onChange={(e) => setMergeSearch(e.target.value)}
                    className="w-full pl-9 pr-4 py-2 bg-white border border-gray-200 rounded-xl text-sm focus:ring-2 focus:ring-indigo-500 outline-none"
                  />
                </div>
                {!showAllMergeOptions && !mergeSearch && (
                  <div className="mt-2 flex items-center justify-between">
                    <p className="text-[10px] text-gray-500 font-medium uppercase tracking-wider">Sugestões baseadas em "{editingEntity?.original}"</p>
                    <button 
                      onClick={() => setShowAllMergeOptions(true)}
                      className="text-[10px] text-indigo-600 font-bold hover:underline"
                    >
                      Mostrar Todos
                    </button>
                  </div>
                )}
              </div>

              <div className="p-4 max-h-[50vh] overflow-y-auto space-y-2">
                {mergeOptions.length > 0 ? (
                  mergeOptions.map(groupHead => (
                    <button
                      key={groupHead.id}
                      onClick={() => {
                        handleMergeToGroup(groupHead.groupId || groupHead.id);
                        setMergeSearch('');
                        setShowAllMergeOptions(false);
                      }}
                      className="w-full p-4 flex items-center justify-between hover:bg-indigo-50 rounded-2xl border border-gray-100 transition-all hover:border-indigo-200 text-left group"
                    >
                      <div className="flex items-center gap-3">
                        <div className="bg-gray-100 group-hover:bg-white p-2 rounded-xl transition-colors">
                          <Layers className="w-4 h-4 text-gray-500 group-hover:text-indigo-600" />
                        </div>
                        <div>
                          <div className="font-bold text-gray-900">{groupHead.original}</div>
                          <div className="text-xs text-gray-500 font-mono">{groupHead.pseudonym}</div>
                        </div>
                      </div>
                      <ChevronRight className="w-5 h-5 text-gray-300 group-hover:text-indigo-400" />
                    </button>
                  ))
                ) : (
                  <div className="py-8 text-center">
                    <p className="text-sm text-gray-400">Nenhum grupo encontrado.</p>
                    {!showAllMergeOptions && (
                      <button 
                        onClick={() => setShowAllMergeOptions(true)}
                        className="mt-2 text-xs text-indigo-600 font-bold hover:underline"
                      >
                        Ver todas as opções
                      </button>
                    )}
                  </div>
                )}
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Exceptions Modal */}
      <AnimatePresence>
        {showManualModal && (
          <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm">
            <motion.div 
              initial={{ opacity: 0, scale: 0.95, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              className="bg-white rounded-3xl shadow-2xl w-full max-w-5xl max-h-[90vh] flex flex-col overflow-hidden border border-gray-100"
            >
              <div className="p-6 border-b border-gray-100 flex justify-between items-center bg-indigo-600 text-white">
                <div className="flex items-center gap-3">
                  <FileText className="w-6 h-6" />
                  <h2 className="text-xl font-bold">Manual do Utilizador - SEATRI</h2>
                </div>
                <div className="flex items-center gap-3">
                  <button 
                    onClick={() => {
                      const printContent = document.getElementById('manual-content');
                      if (printContent) {
                        const win = window.open('', '_blank');
                        win?.document.write(`
                          <html>
                            <head>
                              <title>Manual do Utilizador - SEATRI</title>
                              <link href="https://cdn.jsdelivr.net/npm/tailwindcss@2.2.19/dist/tailwind.min.css" rel="stylesheet">
                              <style>
                                @media print {
                                  .no-print { display: none; }
                                  body { padding: 20px; }
                                }
                              </style>
                            </head>
                            <body>
                              ${printContent.innerHTML}
                            </body>
                          </html>
                        `);
                        win?.document.close();
                        setTimeout(() => win?.print(), 500);
                      }
                    }}
                    className="flex items-center gap-2 px-4 py-2 bg-white/20 hover:bg-white/30 rounded-xl text-sm font-bold transition-all"
                  >
                    <Download className="w-4 h-4" />
                    Descarregar PDF
                  </button>
                  <button onClick={() => setShowManualModal(false)} className="p-2 hover:bg-white/20 rounded-full transition-colors">
                    <X className="w-6 h-6" />
                  </button>
                </div>
              </div>

              <div id="manual-content" className="p-8 overflow-y-auto flex-1 space-y-16 text-gray-800">
                {/* Table of Contents */}
                <section className="p-8 bg-indigo-50 rounded-3xl border border-indigo-100">
                  <h3 className="text-xl font-bold text-indigo-900 mb-6 flex items-center gap-2">
                    <List className="w-5 h-5" />
                    Índice do Manual
                  </h3>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-x-12 gap-y-3 text-sm font-medium text-indigo-700">
                    <a href="#privacidade" className="hover:underline flex items-center gap-2"><span>1.</span> Privacidade e Segurança</a>
                    <a href="#carregamento" className="hover:underline flex items-center gap-2"><span>2.</span> Carregamento e Configuração</a>
                    <a href="#ambiguidade" className="hover:underline flex items-center gap-2"><span>3.</span> Revisão de Ambiguidade</a>
                    <a href="#entidades" className="hover:underline flex items-center gap-2"><span>4.</span> Gestão de Entidades</a>
                    <a href="#editor" className="hover:underline flex items-center gap-2"><span>5.</span> Edição de Detalhe (O Bisturi)</a>
                    <a href="#conhecimento" className="hover:underline flex items-center gap-2"><span>6.</span> Conhecimento Global</a>
                    <a href="#exportacao" className="hover:underline flex items-center gap-2"><span>7.</span> Exportação e Resultados</a>
                    <a href="#botoes" className="hover:underline flex items-center gap-2"><span>8.</span> Guia de Botões e Funções</a>
                    <a href="#dicas" className="hover:underline flex items-center gap-2"><span>9.</span> Dicas de Especialista</a>
                  </div>
                </section>

                {/* Intro Section */}
                <section id="privacidade" className="space-y-6">
                  <div className="flex items-center gap-4 text-indigo-600 border-b-2 border-indigo-100 pb-4">
                    <Shield className="w-10 h-10" />
                    <h3 className="text-3xl font-extrabold tracking-tight">Privacidade e Segurança Local</h3>
                  </div>
                  <div className="grid grid-cols-1 lg:grid-cols-2 gap-12 items-start">
                    <div className="space-y-4">
                      <p className="text-lg leading-relaxed text-gray-600">
                        O <strong>SEATRI</strong> foi desenvolvido para responder às necessidades rigorosas de proteção de dados no setor jurídico. A sua arquitetura é baseada no princípio de <strong>"Privacidade por Design"</strong>.
                      </p>
                      <div className="space-y-4">
                        <div className="flex gap-4 p-4 bg-green-50 rounded-2xl border border-green-100">
                          <Lock className="w-6 h-6 text-green-600 flex-shrink-0" />
                          <div>
                            <h5 className="font-bold text-green-900">Processamento 100% Local</h5>
                            <p className="text-sm text-green-800">Os seus documentos nunca saem do seu computador. Todo o processamento ocorre na memória do seu navegador.</p>
                          </div>
                        </div>
                        <div className="flex gap-4 p-4 bg-blue-50 rounded-2xl border border-blue-100">
                          <Zap className="w-6 h-6 text-blue-600 flex-shrink-0" />
                          <div>
                            <h5 className="font-bold text-blue-900">Sem Inteligência Artificial Externa</h5>
                            <p className="text-sm text-blue-800">Utilizamos algoritmos de Processamento de Linguagem Natural (NLP) locais. Não enviamos dados para APIs de terceiros (como OpenAI ou Google Cloud) para treino ou análise.</p>
                          </div>
                        </div>
                      </div>
                    </div>
                    <div className="bg-gray-50 p-8 rounded-3xl border border-gray-200 shadow-inner">
                      <h5 className="text-xs font-bold text-gray-400 uppercase tracking-widest mb-6 text-center">Arquitetura de Segurança</h5>
                      <div className="relative h-40 flex items-center justify-center">
                        <div className="w-24 h-24 bg-white rounded-2xl shadow-lg border border-gray-100 flex items-center justify-center z-10">
                          <User className="w-10 h-10 text-indigo-600" />
                        </div>
                        <div className="absolute inset-0 flex items-center justify-center">
                          <div className="w-48 h-48 border-4 border-dashed border-indigo-100 rounded-full animate-[spin_20s_linear_infinite]" />
                        </div>
                        <div className="absolute -top-4 left-1/2 -translate-x-1/2 bg-indigo-600 text-white px-3 py-1 rounded-full text-[10px] font-bold">BROWSER</div>
                      </div>
                      <p className="text-center text-xs text-gray-400 mt-4 italic">Os dados permanecem no círculo de confiança do seu navegador.</p>
                    </div>
                  </div>
                </section>

                {/* Step 1: Upload */}
                <section id="carregamento" className="space-y-8">
                  <div className="flex items-center gap-4 text-indigo-600">
                    <div className="w-12 h-12 bg-indigo-600 text-white rounded-2xl flex items-center justify-center text-2xl font-black shadow-lg shadow-indigo-200">1</div>
                    <h3 className="text-2xl font-bold">Início: Carregamento e Configuração</h3>
                  </div>
                  
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
                    <div className="space-y-4">
                      <h4 className="text-lg font-bold text-gray-700">Formatos Suportados</h4>
                      <div className="flex flex-wrap gap-2">
                        {['PDF', 'DOCX', 'XLSX', 'TXT'].map(ext => (
                          <span key={ext} className="px-3 py-1 bg-gray-100 text-gray-600 rounded-lg font-mono font-bold text-xs border border-gray-200">{ext}</span>
                        ))}
                      </div>
                      <p className="text-gray-600">Pode arrastar vários ficheiros em simultâneo para a zona de documentos.</p>
                      
                      <div className="p-5 bg-indigo-50 rounded-2xl border border-indigo-100 space-y-3">
                        <div className="flex items-center gap-2 text-indigo-700 font-bold">
                          <Link className="w-4 h-4" />
                          <span>Documentos Relacionados</span>
                        </div>
                        <p className="text-sm text-indigo-800">
                          Esta opção é vital para processos com múltiplos volumes. Quando ativa, a app garante que o mesmo nome (ex: "João Silva") receba o mesmo pseudónimo (ex: "NOME_1") em todos os ficheiros carregados.
                        </p>
                      </div>
                    </div>
                    <div className="bg-white p-6 rounded-2xl border border-gray-200 shadow-sm space-y-4">
                      <div className="flex items-center justify-between border-b pb-2">
                        <span className="text-xs font-bold text-gray-400 uppercase">Interface de Upload</span>
                      </div>
                      <div className="border-2 border-dashed border-gray-200 rounded-xl p-8 flex flex-col items-center justify-center text-gray-400">
                        <Upload className="w-8 h-8 mb-2" />
                        <span className="text-xs font-medium">Arraste os seus ficheiros aqui</span>
                      </div>
                      <div className="flex items-center gap-2 p-2 bg-gray-50 rounded-lg border border-gray-100">
                        <div className="w-4 h-4 bg-indigo-600 rounded flex items-center justify-center">
                          <Check className="w-3 h-3 text-white" />
                        </div>
                        <span className="text-xs font-bold text-gray-700">Documentos Relacionados</span>
                      </div>
                    </div>
                  </div>
                </section>

                {/* Step 2: Ambiguity */}
                <section id="ambiguidade" className="space-y-8">
                  <div className="flex items-center gap-4 text-indigo-600">
                    <div className="w-12 h-12 bg-indigo-600 text-white rounded-2xl flex items-center justify-center text-2xl font-black shadow-lg shadow-indigo-200">2</div>
                    <h3 className="text-2xl font-bold">Revisão de Ambiguidade: O Filtro Inicial</h3>
                  </div>
                  
                  <div className="space-y-6">
                    <p className="text-gray-600 leading-relaxed">
                      Após o scan, a app deteta elementos que podem conter várias pessoas "coladas". Esta etapa impede que o algoritmo de agrupamento cometa erros graves.
                    </p>
                    
                    <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                      <div className="p-6 bg-amber-50 rounded-2xl border border-amber-100 space-y-3">
                        <div className="w-10 h-10 bg-amber-100 rounded-xl flex items-center justify-center text-amber-600">
                          <Scissors className="w-5 h-5" />
                        </div>
                        <h5 className="font-bold text-amber-900">Dividir por "e"</h5>
                        <p className="text-xs text-amber-800">Separa nomes como "João Silva e Maria Santos" em dois registos independentes.</p>
                      </div>
                      <div className="p-6 bg-amber-50 rounded-2xl border border-amber-100 space-y-3">
                        <div className="w-10 h-10 bg-amber-100 rounded-xl flex items-center justify-center text-amber-600">
                          <Unlink className="w-5 h-5" />
                        </div>
                        <h5 className="font-bold text-amber-900">Dividir por Vírgula</h5>
                        <p className="text-xs text-amber-800">Ideal para listas de testemunhas ou partes separadas por vírgulas.</p>
                      </div>
                      <div className="p-6 bg-amber-50 rounded-2xl border border-amber-100 space-y-3">
                        <div className="w-10 h-10 bg-amber-100 rounded-xl flex items-center justify-center text-amber-600">
                          <Shield className="w-5 h-5" />
                        </div>
                        <h5 className="font-bold text-amber-900">Exceção Global</h5>
                        <p className="text-xs text-amber-800">Se a app detetar um termo que não deve ser anonimizado, use o escudo para o ignorar para sempre.</p>
                      </div>
                    </div>

                    <div className="bg-gray-900 p-6 rounded-3xl overflow-hidden shadow-2xl">
                      <div className="flex items-center gap-2 mb-4">
                        <div className="w-3 h-3 bg-red-500 rounded-full" />
                        <div className="w-3 h-3 bg-yellow-500 rounded-full" />
                        <div className="w-3 h-3 bg-green-500 rounded-full" />
                        <span className="text-[10px] text-gray-500 font-mono ml-2">AMBIGUITY_REVIEW_MODAL.EXE</span>
                      </div>
                      <div className="space-y-3">
                        <div className="flex items-center justify-between p-3 bg-gray-800 rounded-xl border border-gray-700">
                          <span className="text-white font-medium">Armindo Pereira e Liliana Ferreira</span>
                          <div className="flex gap-2">
                            <div className="px-2 py-1 bg-indigo-600 text-white text-[10px] font-bold rounded">DIVIDIR POR "E"</div>
                            <div className="px-2 py-1 bg-gray-700 text-gray-400 text-[10px] font-bold rounded">MANTER ASSIM</div>
                          </div>
                        </div>
                        <div className="flex items-center justify-between p-3 bg-gray-800 rounded-xl border border-gray-700 opacity-50">
                          <span className="text-white font-medium">Tribunal da Relação de Lisboa</span>
                          <div className="flex gap-2">
                            <div className="p-1 bg-gray-700 text-white rounded"><Shield className="w-3 h-3" /></div>
                            <div className="px-2 py-1 bg-green-600 text-white text-[10px] font-bold rounded">REVISADO</div>
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>
                </section>

                {/* Step 3: Entity List */}
                <section id="entidades" className="space-y-8">
                  <div className="flex items-center gap-4 text-indigo-600">
                    <div className="w-12 h-12 bg-indigo-600 text-white rounded-2xl flex items-center justify-center text-2xl font-black shadow-lg shadow-indigo-200">3</div>
                    <h3 className="text-2xl font-bold">Gestão de Entidades: A Lista Mestra</h3>
                  </div>
                  
                  <div className="grid grid-cols-1 lg:grid-cols-2 gap-12">
                    <div className="space-y-6">
                      <p className="text-gray-600">Aqui é onde passa a maior parte do tempo. Os nomes são agrupados automaticamente, mas tem controlo total sobre cada um.</p>
                      
                      <div className="space-y-4">
                        <div className="flex items-start gap-4">
                          <div className="w-8 h-8 bg-indigo-100 text-indigo-600 rounded-lg flex items-center justify-center flex-shrink-0"><Plus className="w-4 h-4" /></div>
                          <div>
                            <h6 className="font-bold text-gray-900">Agrupamento Manual</h6>
                            <p className="text-sm text-gray-500">Selecione dois ou mais elementos e clique no botão de grupo para os unir sob o mesmo pseudónimo.</p>
                          </div>
                        </div>
                        <div className="flex items-start gap-4">
                          <div className="w-8 h-8 bg-red-100 text-red-600 rounded-lg flex items-center justify-center flex-shrink-0"><Unlink className="w-4 h-4" /></div>
                          <div>
                            <h6 className="font-bold text-gray-900">Dissolver Grupo</h6>
                            <p className="text-sm text-gray-500">Se o sistema agrupou pessoas erradas, use este botão para separar todos os elementos do grupo.</p>
                          </div>
                        </div>
                        <div className="flex items-start gap-4">
                          <div className="w-8 h-8 bg-gray-100 text-gray-600 rounded-lg flex items-center justify-center flex-shrink-0"><Search className="w-4 h-4" /></div>
                          <div>
                            <h6 className="font-bold text-gray-900">Filtros Inteligentes</h6>
                            <p className="text-sm text-gray-500">Filtre por tipo (NOME, LOCAL, PHONE, etc.) ou use a barra de pesquisa para encontrar rapidamente um termo.</p>
                          </div>
                        </div>
                      </div>
                    </div>
                    
                    <div className="space-y-4">
                      <div className="bg-white border border-gray-200 rounded-2xl shadow-sm overflow-hidden">
                        <div className="p-3 bg-gray-50 border-b border-gray-100 flex justify-between items-center">
                          <span className="text-[10px] font-bold text-gray-400 uppercase">Exemplo de Grupo</span>
                          <Unlink className="w-3 h-3 text-red-400" />
                        </div>
                        <div className="p-4 space-y-3">
                          <div className="flex items-center justify-between">
                            <div className="flex items-center gap-2">
                              <div className="w-2 h-2 bg-indigo-600 rounded-full" />
                              <span className="font-bold text-sm">Maria Silva</span>
                            </div>
                            <span className="text-[10px] font-mono bg-gray-100 px-2 py-0.5 rounded">NOME_1</span>
                          </div>
                          <div className="pl-4 space-y-2 border-l-2 border-gray-100">
                            <div className="flex items-center justify-between text-xs text-gray-500">
                              <span>Maria S. Silva</span>
                              <Scissors className="w-3 h-3" />
                            </div>
                            <div className="flex items-center justify-between text-xs text-gray-500">
                              <span>M. Silva</span>
                              <Trash2 className="w-3 h-3" />
                            </div>
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>
                </section>

                {/* Step 4: Detail Editor */}
                <section id="editor" className="space-y-8">
                  <div className="flex items-center gap-4 text-indigo-600">
                    <div className="w-12 h-12 bg-indigo-600 text-white rounded-2xl flex items-center justify-center text-2xl font-black shadow-lg shadow-indigo-200">4</div>
                    <h3 className="text-2xl font-bold">Edição de Detalhe: O Bisturi</h3>
                  </div>
                  
                  <div className="grid grid-cols-1 lg:grid-cols-2 gap-12 items-center">
                    <div className="bg-white border border-gray-200 rounded-3xl shadow-xl overflow-hidden">
                      <div className="p-4 bg-indigo-600 text-white flex justify-between items-center">
                        <span className="text-xs font-bold uppercase">Editor de Elemento</span>
                        <X className="w-4 h-4" />
                      </div>
                      <div className="p-6 space-y-6">
                        <div className="p-4 bg-gray-50 rounded-xl text-xs leading-relaxed border border-gray-100">
                          <span className="text-gray-400">...o depoimento de </span>
                          <span className="bg-yellow-200 px-1 rounded font-bold mx-1">João Manuel Silva</span>
                          <span className="text-gray-400"> foi decisivo para...</span>
                        </div>
                        <div className="grid grid-cols-2 gap-4">
                          <div className="space-y-1">
                            <span className="text-[10px] font-bold text-gray-400 uppercase">Original</span>
                            <div className="p-2 bg-gray-50 rounded border text-xs">João Manuel Silva</div>
                          </div>
                          <div className="space-y-1">
                            <span className="text-[10px] font-bold text-gray-400 uppercase">Pseudónimo</span>
                            <div className="p-2 bg-indigo-50 rounded border border-indigo-100 text-xs font-bold text-indigo-600">NOME_4</div>
                          </div>
                        </div>
                        <div className="flex gap-2">
                          <div className="flex-1 p-2 bg-gray-100 rounded text-[10px] font-bold text-center">EXPANDIR INÍCIO</div>
                          <div className="flex-1 p-2 bg-gray-100 rounded text-[10px] font-bold text-center">EXPANDIR FIM</div>
                        </div>
                      </div>
                    </div>
                    
                    <div className="space-y-6">
                      <h4 className="text-xl font-bold text-gray-900">Funcionalidades do Editor</h4>
                      <ul className="space-y-4">
                        <li className="flex gap-4">
                          <div className="w-10 h-10 bg-indigo-100 rounded-xl flex items-center justify-center text-indigo-600 flex-shrink-0"><Eye className="w-5 h-5" /></div>
                          <div>
                            <h6 className="font-bold text-gray-900">Visualização de Contexto</h6>
                            <p className="text-sm text-gray-500">Veja exatamente onde o nome aparece no documento original para não ter dúvidas sobre quem se trata.</p>
                          </div>
                        </li>
                        <li className="flex gap-4">
                          <div className="w-10 h-10 bg-indigo-100 rounded-xl flex items-center justify-center text-indigo-600 flex-shrink-0"><Plus className="w-5 h-5" /></div>
                          <div>
                            <h6 className="font-bold text-gray-900">Ajuste de Limites</h6>
                            <p className="text-sm text-gray-500">Se a app "cortou" um apelido ou apanhou uma palavra a mais, use os botões de Expandir/Contrair para ajustar a seleção com precisão de caractere.</p>
                          </div>
                        </li>
                        <li className="flex gap-4">
                          <div className="w-10 h-10 bg-indigo-100 rounded-xl flex items-center justify-center text-indigo-600 flex-shrink-0"><Link className="w-5 h-5" /></div>
                          <div>
                            <h6 className="font-bold text-gray-900">Mesclagem Direta</h6>
                            <p className="text-sm text-gray-500">Pode unir o elemento atual a qualquer outro grupo existente diretamente do editor.</p>
                          </div>
                        </li>
                      </ul>
                    </div>
                  </div>
                </section>

                {/* Step 5: Knowledge Base */}
                <section id="conhecimento" className="space-y-8">
                  <div className="flex items-center gap-4 text-indigo-600">
                    <div className="w-12 h-12 bg-indigo-600 text-white rounded-2xl flex items-center justify-center text-2xl font-black shadow-lg shadow-indigo-200">5</div>
                    <h3 className="text-2xl font-bold">Conhecimento Global: Exceções, Juízes, Autores e Safelist</h3>
                  </div>
                  
                  <div className="space-y-6">
                    <p className="text-gray-600 leading-relaxed">
                      O sistema mantém uma base de dados de termos que devem ser tratados de forma especial em todos os documentos. Esta base é <strong>persistente</strong>: o que adicionar hoje estará disponível amanhã, mesmo em novos projetos.
                    </p>
                    
                    <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
                      <div className="p-6 bg-white border border-gray-200 rounded-3xl shadow-sm space-y-4">
                        <div className="w-12 h-12 bg-indigo-100 text-indigo-600 rounded-2xl flex items-center justify-center"><Shield className="w-6 h-6" /></div>
                        <h5 className="font-bold text-gray-900">Exceções</h5>
                        <p className="text-sm text-gray-500">Nomes de instituições, cidades ou termos técnicos que a app pode confundir com nomes de pessoas. Estes termos são removidos da anonimização.</p>
                      </div>
                      <div className="p-6 bg-white border border-gray-200 rounded-3xl shadow-sm space-y-4">
                        <div className="w-12 h-12 bg-blue-100 text-blue-600 rounded-2xl flex items-center justify-center"><Scale className="w-6 h-6" /></div>
                        <h5 className="font-bold text-gray-900">Juízes</h5>
                        <p className="text-sm text-gray-500">Pode carregar uma lista oficial de juízes. Quando detetados, são marcados como tal e pode decidir se os mantém ou anonimiza.</p>
                      </div>
                      <div className="p-6 bg-white border border-gray-200 rounded-3xl shadow-sm space-y-4">
                        <div className="w-12 h-12 bg-orange-100 text-orange-600 rounded-2xl flex items-center justify-center"><User className="w-6 h-6" /></div>
                        <h5 className="font-bold text-gray-900">Autores</h5>
                        <p className="text-sm text-gray-500">Semelhante aos juízes, para identificar magistrados do Ministério Público ou advogados específicos.</p>
                      </div>
                      <div className="p-6 bg-white border border-gray-200 rounded-3xl shadow-sm space-y-4">
                        <div className="w-12 h-12 bg-green-100 text-green-600 rounded-2xl flex items-center justify-center"><List className="w-6 h-6" /></div>
                        <h5 className="font-bold text-gray-900">Safelist</h5>
                        <p className="text-sm text-gray-500">Uma lista técnica de palavras e frases comuns que devem ser ignoradas pelo motor de deteção para evitar falsos positivos.</p>
                      </div>
                    </div>

                    <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
                      <div className="p-6 bg-indigo-50 rounded-3xl border border-indigo-100 space-y-4">
                        <h6 className="font-bold text-indigo-900 flex items-center gap-2">
                          <Upload className="w-5 h-5" />
                          Importação de Safelist (JSON/TXT)
                        </h6>
                        <p className="text-sm text-indigo-800 leading-relaxed">
                          A Safelist permite importar listas massivas de termos jurídicos ou institucionais. 
                          Pode importar ficheiros <strong>.json</strong> estruturados ou simples ficheiros <strong>.txt</strong> (onde cada linha é um termo). 
                          A app separa automaticamente palavras de frases para otimizar a precisão.
                        </p>
                      </div>
                      <div className="p-6 bg-indigo-600 text-white rounded-3xl shadow-xl space-y-4">
                        <h6 className="font-bold flex items-center gap-2">
                          <Download className="w-5 h-5" />
                          Exportação e Partilha de Conhecimento
                        </h6>
                        <p className="text-sm text-indigo-100 leading-relaxed">
                          Pode exportar <strong>todo o seu conhecimento</strong> (Exceções + Juízes + Autores + Safelist) num único ficheiro JSON. 
                          Isto permite que equipas partilhem as suas bases de dados de termos a ignorar, garantindo consistência entre diferentes utilizadores.
                        </p>
                      </div>
                    </div>
                  </div>
                </section>

                {/* Step 6: Export */}
                <section id="exportacao" className="space-y-8">
                  <div className="flex items-center gap-4 text-indigo-600">
                    <div className="w-12 h-12 bg-indigo-600 text-white rounded-2xl flex items-center justify-center text-2xl font-black shadow-lg shadow-indigo-200">6</div>
                    <h3 className="text-2xl font-bold">Exportação e Resultados</h3>
                  </div>
                  
                  <div className="grid grid-cols-1 lg:grid-cols-2 gap-12 items-center">
                    <div className="space-y-6">
                      <p className="text-gray-600">O trabalho termina com a geração dos novos documentos. O sistema garante que a formatação original é preservada tanto quanto possível.</p>
                      
                      <div className="space-y-4">
                        <div className="flex gap-4 items-center p-4 bg-gray-50 rounded-2xl border border-gray-100">
                          <div className="w-10 h-10 bg-white rounded-xl shadow-sm flex items-center justify-center text-indigo-600"><FileText className="w-5 h-5" /></div>
                          <div>
                            <h6 className="font-bold text-gray-900">Documentos Anonimizados</h6>
                            <p className="text-xs text-gray-500">PDFs, DOCX e TXT com os pseudónimos aplicados nos locais corretos.</p>
                          </div>
                        </div>
                        <div className="flex gap-4 items-center p-4 bg-gray-50 rounded-2xl border border-gray-100">
                          <div className="w-10 h-10 bg-white rounded-xl shadow-sm flex items-center justify-center text-green-600"><CheckCircle2 className="w-5 h-5" /></div>
                          <div>
                            <h6 className="font-bold text-gray-900">Relatório Excel</h6>
                            <p className="text-xs text-gray-500">Uma tabela detalhada com a correspondência "Original &rarr; Pseudónimo" para arquivo interno.</p>
                          </div>
                        </div>
                        <div className="flex gap-4 items-center p-4 bg-gray-50 rounded-2xl border border-gray-100">
                          <div className="w-10 h-10 bg-white rounded-xl shadow-sm flex items-center justify-center text-amber-600"><Save className="w-5 h-5" /></div>
                          <div>
                            <h6 className="font-bold text-gray-900">Ficheiro de Projeto (JSON)</h6>
                            <p className="text-xs text-gray-500">Guarde o estado atual para continuar o trabalho noutro dia ou noutro computador.</p>
                          </div>
                        </div>
                      </div>
                    </div>
                    
                    <div className="bg-indigo-600 p-8 rounded-[40px] text-white space-y-6 shadow-2xl relative overflow-hidden">
                      <div className="absolute top-0 right-0 w-32 h-32 bg-white/10 rounded-full -mr-16 -mt-16 blur-2xl" />
                      <div className="relative z-10 space-y-4">
                        <h5 className="text-xl font-bold">Pronto para Exportar?</h5>
                        <p className="text-indigo-100 text-sm">Verifique se todos os elementos importantes estão marcados como "Tratados" (verde) para garantir a máxima qualidade.</p>
                        <div className="pt-4">
                          <div className="w-full py-4 bg-white text-indigo-600 rounded-2xl font-black text-center shadow-xl flex items-center justify-center gap-3">
                            <Download className="w-6 h-6" />
                            EXPORTAR TUDO (.ZIP)
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>
                </section>

                {/* Step 8: Management Buttons */}
                <section id="botoes" className="space-y-8">
                  <div className="flex items-center gap-4 text-indigo-600 border-b-2 border-indigo-100 pb-4">
                    <Zap className="w-10 h-10" />
                    <h3 className="text-3xl font-extrabold tracking-tight">Guia de Funcionalidades e Botões</h3>
                  </div>
                  
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                    <div className="p-6 bg-white border border-gray-200 rounded-3xl shadow-sm space-y-4">
                      <div className="flex items-center gap-3 text-green-600">
                        <CheckCircle2 className="w-6 h-6" />
                        <h4 className="font-bold text-lg">Validar Tudo</h4>
                      </div>
                      <p className="text-sm text-gray-600"><strong>O que faz:</strong> Marca todos os elementos da lista atual como "Validados" de uma só vez. Útil quando já revisou visualmente a lista e concorda com as deteções.</p>
                      <p className="text-xs text-gray-500 bg-gray-50 p-3 rounded-xl italic"><strong>Efeito no trabalho:</strong> Altera o estado de "Tratado" para verdadeiro em todos os itens. Não altera nomes, tipos ou pseudónimos.</p>
                    </div>

                    <div className="p-6 bg-white border border-gray-200 rounded-3xl shadow-sm space-y-4">
                      <div className="flex items-center gap-3 text-blue-600">
                        <Scissors className="w-6 h-6" />
                        <h4 className="font-bold text-lg">Dividir 'e'</h4>
                      </div>
                      <p className="text-sm text-gray-600"><strong>O que faz:</strong> Procura em toda a lista nomes que contenham a conjunção " e " (ex: "João e Maria") e separa-os em dois registos diferentes.</p>
                      <p className="text-xs text-gray-500 bg-gray-50 p-3 rounded-xl italic"><strong>Efeito no trabalho:</strong> Cria novos elementos na lista. Pode alterar o agrupamento se os novos nomes forem semelhantes a outros já existentes.</p>
                    </div>

                    <div className="p-6 bg-white border border-gray-200 rounded-3xl shadow-sm space-y-4">
                      <div className="flex items-center gap-3 text-emerald-600">
                        <RotateCw className="w-6 h-6" />
                        <h4 className="font-bold text-lg">Re-analisar Documentos</h4>
                      </div>
                      <p className="text-sm text-gray-600"><strong>O que faz:</strong> Executa uma nova varredura em todos os ficheiros abertos. Utiliza o "Conhecimento Global" (Juízes, Autores e Exceções) para encontrar novos termos ou corrigir tipos.</p>
                      <p className="text-xs text-gray-500 bg-gray-50 p-3 rounded-xl italic"><strong>Efeito no trabalho:</strong> <u>Não altera elementos que já foram validados ou ignorados</u> pelo utilizador. Apenas adiciona novas deteções ou atualiza elementos que ainda não foram mexidos.</p>
                    </div>

                    <div className="p-6 bg-white border border-gray-200 rounded-3xl shadow-sm space-y-4">
                      <div className="flex items-center gap-3 text-amber-600">
                        <Plus className="w-6 h-6" />
                        <h4 className="font-bold text-lg">Sugerir Grupos</h4>
                      </div>
                      <p className="text-sm text-gray-600"><strong>O que faz:</strong> Abre uma janela com sugestões de nomes muito parecidos que o sistema acha que são a mesma pessoa, mas que ainda estão com pseudónimos diferentes (ex: "Maria Silva" e "Maria S. Silva").</p>
                      <p className="text-xs text-gray-500 bg-gray-50 p-3 rounded-xl italic"><strong>Efeito no trabalho:</strong> Não altera nada automaticamente. O utilizador decide, caso a caso, se aceita ou rejeita a sugestão apresentada.</p>
                    </div>

                    <div className="p-6 bg-white border border-gray-200 rounded-3xl shadow-sm space-y-4">
                      <div className="flex items-center gap-3 text-indigo-600">
                        <RotateCw className="w-6 h-6" />
                        <h4 className="font-bold text-lg">Re-agrupar</h4>
                      </div>
                      <p className="text-sm text-gray-600"><strong>O que faz:</strong> Re-executa o algoritmo de similaridade para garantir que todos os nomes que devem estar juntos partilham o mesmo pseudónimo.</p>
                      <p className="text-xs text-gray-500 bg-gray-50 p-3 rounded-xl italic"><strong>Efeito no trabalho:</strong> Pode alterar Pseudónimos e Grupos de entidades <u>não validadas</u> se o algoritmo detetar que elas pertencem a um grupo diferente. Mantém grupos manuais intactos.</p>
                    </div>

                    <div className="p-6 bg-white border border-gray-200 rounded-3xl shadow-sm space-y-4">
                      <div className="flex items-center gap-3 text-red-600">
                        <Trash2 className="w-6 h-6" />
                        <h4 className="font-bold text-lg">Limpar Tudo</h4>
                      </div>
                      <p className="text-sm text-gray-600"><strong>O que faz:</strong> Apaga todos os documentos e todas as entidades detetadas, reiniciando a sessão.</p>
                      <p className="text-xs text-gray-500 bg-gray-50 p-3 rounded-xl italic"><strong>Efeito no trabalho:</strong> Perda total do trabalho da sessão atual (exceto o Conhecimento Global).</p>
                    </div>
                  </div>
                </section>

                {/* Final Tips */}
                <section id="dicas" className="p-10 bg-gray-900 rounded-[40px] text-white">
                  <div className="flex items-center gap-3 mb-8">
                    <Zap className="w-8 h-8 text-amber-400" />
                    <h3 className="text-2xl font-bold">Dicas de Especialista</h3>
                  </div>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-10">
                    <div className="space-y-2">
                      <h5 className="font-bold text-amber-400">Split View (Lado-a-Lado)</h5>
                      <p className="text-sm text-gray-400 leading-relaxed">
                        Ative o ícone de camadas no topo para abrir o visualizador duplo. Pode ler o documento original à esquerda e ver como fica anonimizado à direita, com sincronização de scroll.
                      </p>
                    </div>
                    <div className="space-y-2">
                      <h5 className="font-bold text-amber-400">Histórico de Ações</h5>
                      <p className="text-sm text-gray-400 leading-relaxed">
                        A app guarda as suas últimas 30 ações. Se apagar um grupo por engano ou fizer uma divisão errada, use as setas de Undo/Redo no cabeçalho para voltar atrás no tempo.
                      </p>
                    </div>
                    <div className="space-y-2">
                      <h5 className="font-bold text-amber-400">Atalhos de Teclado</h5>
                      <p className="text-sm text-gray-400 leading-relaxed">
                        No campo de adição de exceções, prima <strong>Enter</strong> para adicionar rapidamente um termo sem precisar de clicar no botão.
                      </p>
                    </div>
                    <div className="space-y-2">
                      <h5 className="font-bold text-amber-400">Limpeza de Projeto</h5>
                      <p className="text-sm text-gray-400 leading-relaxed">
                        O botão "Novo Projeto" limpa tudo da memória. Use-o sempre que terminar um caso para garantir que não mistura dados de processos diferentes.
                      </p>
                    </div>
                  </div>
                </section>
              </div>
            </motion.div>
          </div>
        )}

        {showAmbiguityModal && (
          <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm">
            <motion.div 
              initial={{ opacity: 0, scale: 0.95, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              className="bg-white rounded-3xl shadow-2xl w-full max-w-4xl max-h-[85vh] flex flex-col overflow-hidden border border-gray-100"
            >
              <div className="p-6 border-b border-gray-100 flex justify-between items-center bg-amber-50/50">
                <div className="flex items-center gap-3">
                  <div className="w-10 h-10 bg-amber-100 rounded-xl flex items-center justify-center text-amber-600">
                    <AlertCircle className="w-6 h-6" />
                  </div>
                  <div>
                    <h2 className="text-xl font-bold text-gray-900">Revisão de Ambiguidade</h2>
                    <p className="text-sm text-amber-700">Detetámos nomes que podem ser listas de várias pessoas. Por favor, verifique antes de prosseguir.</p>
                  </div>
                </div>
                <div className="flex items-center gap-2">
                  <button 
                    onClick={undo}
                    disabled={historyIndex <= 0}
                    className="p-2.5 bg-white border border-gray-200 rounded-xl text-gray-600 hover:bg-gray-50 disabled:opacity-30 disabled:cursor-not-allowed transition-all shadow-sm"
                    title="Anular última ação (Undo)"
                  >
                    <RotateCcw className="w-5 h-5" />
                  </button>
                  <button 
                    onClick={handleFinishAmbiguityReview}
                    className="bg-indigo-600 text-white px-6 py-2.5 rounded-xl font-bold hover:bg-indigo-700 transition-all shadow-lg shadow-indigo-200 flex items-center gap-2"
                  >
                    <Check className="w-4 h-4" />
                    Concluir e Agrupar
                  </button>
                </div>
              </div>

              <div className="p-6 overflow-y-auto flex-1 space-y-4">
                {ambiguousEntities.filter(e => 
                  (e.type === 'NOME' || e.type === 'AUTOR' || e.type === 'JUIZ') && 
                  (e.original.includes(' e ') || e.original.includes(','))
                ).length === 0 ? (
                  <div className="text-center py-12">
                    <CheckCircle2 className="w-16 h-16 text-green-500 mx-auto mb-4" />
                    <p className="text-gray-600 font-medium">Todas as ambiguidades foram resolvidas!</p>
                    <button 
                      onClick={handleFinishAmbiguityReview}
                      className="mt-4 text-indigo-600 font-bold hover:underline"
                    >
                      Clique aqui para avançar para o agrupamento
                    </button>
                  </div>
                ) : (
                  <div className="grid grid-cols-1 gap-3">
                    {ambiguousEntities
                      .filter(e => 
                        (e.type === 'NOME' || e.type === 'AUTOR' || e.type === 'JUIZ') && 
                        (e.original.includes(' e ') || e.original.includes(','))
                      )
                      .map((entity) => (
                        <div key={entity.id} className="flex items-center justify-between p-4 bg-gray-50 rounded-2xl border border-gray-200 group hover:border-indigo-200 transition-all">
                          <div className="flex-1">
                            <div className="flex items-center gap-2 mb-1">
                              <span className="text-[10px] font-bold px-1.5 py-0.5 rounded bg-gray-200 text-gray-600 uppercase tracking-wider">
                                {entity.type}
                              </span>
                              <span className="text-xs text-gray-400 italic">Original:</span>
                            </div>
                            <span className="text-lg font-medium text-gray-900">{entity.original}</span>
                          </div>
                          
                          <div className="flex items-center gap-2">
                            {entity.original.includes(' e ') && (
                              <button 
                                onClick={() => handleSplitAmbiguous(entity.id, 'e')}
                                className="flex items-center gap-2 px-3 py-2 bg-white border border-indigo-100 text-indigo-600 rounded-xl text-sm font-bold hover:bg-indigo-50 transition-all"
                              >
                                <Scissors className="w-4 h-4" />
                                Dividir por "e"
                              </button>
                            )}
                            {entity.original.includes(',') && (
                              <button 
                                onClick={() => handleSplitAmbiguous(entity.id, 'comma')}
                                className="flex items-center gap-2 px-3 py-2 bg-white border border-indigo-100 text-indigo-600 rounded-xl text-sm font-bold hover:bg-indigo-50 transition-all"
                              >
                                <Unlink className="w-4 h-4" />
                                Dividir por ","
                              </button>
                            )}
                            
                            <button 
                              onClick={() => handleAddToExceptionsFromAmbiguity(entity.original)}
                              className="p-2 text-gray-400 hover:text-indigo-600 hover:bg-indigo-50 rounded-xl transition-all"
                              title="Adicionar às exceções globais"
                            >
                              <Shield className="w-5 h-5" />
                            </button>

                            <button 
                              onClick={() => handleMarkAsReviewed(entity.id)}
                              className={`px-4 py-2 rounded-xl text-sm font-bold transition-all ${
                                entity.reviewed 
                                  ? 'bg-green-100 text-green-700 border border-green-200' 
                                  : 'bg-white border border-gray-200 text-gray-600 hover:bg-gray-50'
                              }`}
                            >
                              {entity.reviewed ? (
                                <span className="flex items-center gap-2">
                                  <Check className="w-4 h-4" />
                                  Revisado
                                </span>
                              ) : 'Manter assim'}
                            </button>
                          </div>
                        </div>
                      ))}
                  </div>
                )}
              </div>
            </motion.div>
          </div>
        )}

        {showDeduplicationModal && (
          <div className="fixed inset-0 z-[110] flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm">
            <motion.div 
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              className="bg-white rounded-3xl shadow-2xl w-full max-w-2xl max-h-[80vh] flex flex-col overflow-hidden"
            >
              <div className="p-6 border-b border-gray-100 flex justify-between items-center bg-amber-50">
                <div className="flex items-center gap-3">
                  <Copy className="w-6 h-6 text-amber-600" />
                  <h2 className="text-xl font-bold text-gray-900">Sugestões de Duplicação</h2>
                </div>
                <button onClick={() => setShowDeduplicationModal(false)} className="p-2 hover:bg-amber-100 rounded-full transition-colors">
                  <X className="w-5 h-5" />
                </button>
              </div>
              <div className="p-6 overflow-y-auto flex-1 space-y-4">
                {deduplicationSuggestions.length === 0 ? (
                  <div className="text-center py-12 text-gray-500">
                    <CheckCircle2 className="w-12 h-12 mx-auto mb-4 text-green-500" />
                    <p>Não foram encontradas mais duplicações.</p>
                  </div>
                ) : (
                  deduplicationSuggestions.map((s, idx) => (
                    <div key={idx} className="p-4 bg-gray-50 rounded-2xl border border-gray-200 space-y-3">
                      <div className="flex justify-between items-start">
                        <span className="text-xs font-bold text-amber-600 uppercase tracking-wider">Possível Duplicado ({Math.round(s.score * 100)}%)</span>
                        <span className="text-[10px] bg-gray-200 px-1.5 py-0.5 rounded text-gray-600 font-bold uppercase">{s.type}</span>
                      </div>
                      <div className="grid grid-cols-2 gap-4">
                        <div className="space-y-2">
                          <p className="text-sm font-medium text-gray-900 break-words">{s.item1}</p>
                          <button 
                            onClick={() => handleResolveDuplicate(s.item1, s.item2)}
                            className="w-full py-2 bg-indigo-50 text-indigo-700 rounded-xl text-xs font-bold hover:bg-indigo-100 transition-all"
                          >
                            Manter este
                          </button>
                        </div>
                        <div className="space-y-2 border-l border-gray-200 pl-4">
                          <p className="text-sm font-medium text-gray-900 break-words">{s.item2}</p>
                          <button 
                            onClick={() => handleResolveDuplicate(s.item2, s.item1)}
                            className="w-full py-2 bg-indigo-50 text-indigo-700 rounded-xl text-xs font-bold hover:bg-indigo-100 transition-all"
                          >
                            Manter este
                          </button>
                        </div>
                      </div>
                      <div className="pt-2 border-t border-gray-100 flex justify-center">
                        <button 
                          onClick={() => handleDiscardSuggestion(idx)}
                          className="text-xs font-bold text-gray-500 hover:text-indigo-600 transition-colors flex items-center gap-1.5 py-1 px-3 rounded-lg hover:bg-white"
                        >
                          <CheckCircle2 className="w-3.5 h-3.5" />
                          Manter Ambos (Não são duplicados)
                        </button>
                      </div>
                    </div>
                  ))
                )}
              </div>
              <div className="p-4 border-t border-gray-100 bg-gray-50 flex justify-between items-center">
                <div className="flex items-center gap-2">
                  {deduplicationSuggestions.length > 0 && (
                    <button 
                      onClick={handleResolveAllDuplicates}
                      className="px-4 py-2 bg-amber-100 text-amber-700 rounded-xl text-sm font-bold hover:bg-amber-200 transition-all flex items-center gap-2"
                    >
                      <CheckCircle2 className="w-4 h-4" />
                      Validar Todas as Sugestões
                    </button>
                  )}
                </div>
                <button 
                  onClick={() => setShowDeduplicationModal(false)}
                  className="px-6 py-2 bg-white border border-gray-200 rounded-xl font-bold hover:bg-gray-100 transition-all"
                >
                  Fechar
                </button>
              </div>
            </motion.div>
          </div>
        )}

        {editingKnowledge && (
          <div className="fixed inset-0 z-[120] flex items-center justify-center p-4 bg-black/40 backdrop-blur-sm">
            <motion.div 
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              className="bg-white rounded-2xl shadow-2xl w-full max-w-md overflow-hidden border border-gray-100"
            >
              <div className="p-4 border-b border-gray-100 flex justify-between items-center">
                <h3 className="font-bold text-gray-900 flex items-center gap-2">
                  <Pencil className="w-4 h-4 text-indigo-600" />
                  Editar Elemento
                </h3>
                <button onClick={() => setEditingKnowledge(null)} className="p-1 hover:bg-gray-100 rounded-full transition-colors">
                  <X className="w-4 h-4" />
                </button>
              </div>
              <div className="p-6 space-y-4">
                <div className="space-y-1.5">
                  <label className="text-xs font-bold text-gray-400 uppercase tracking-wider ml-1">Texto do Elemento</label>
                  <input 
                    autoFocus
                    type="text" 
                    value={editingKnowledge.current}
                    onChange={(e) => setEditingKnowledge({ ...editingKnowledge, current: e.target.value })}
                    onKeyDown={(e) => e.key === 'Enter' && handleSaveEditKnowledge()}
                    className="w-full px-4 py-3 bg-gray-50 border border-gray-200 rounded-xl focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 outline-none transition-all font-medium"
                  />
                </div>
                <div className="flex gap-3 pt-2">
                  <button 
                    onClick={() => setEditingKnowledge(null)}
                    className="flex-1 py-2.5 bg-gray-100 text-gray-600 rounded-xl font-bold hover:bg-gray-200 transition-all"
                  >
                    Cancelar
                  </button>
                  <button 
                    onClick={handleSaveEditKnowledge}
                    className="flex-1 py-2.5 bg-indigo-600 text-white rounded-xl font-bold hover:bg-indigo-700 transition-all shadow-lg shadow-indigo-200"
                  >
                    Guardar
                  </button>
                </div>
              </div>
            </motion.div>
          </div>
        )}

        {showExceptionsModal && (
          <div className="fixed inset-0 z-[60] flex items-center justify-center p-4">
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setShowExceptionsModal(false)}
              className="absolute inset-0 bg-black/60 backdrop-blur-sm"
            />
            <motion.div 
              initial={{ scale: 0.9, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.9, opacity: 0 }}
              className="relative bg-white w-full max-w-2xl rounded-3xl shadow-2xl overflow-hidden flex flex-col max-h-[90vh]"
            >
              <div className="p-6 border-b border-gray-100 flex items-center justify-between">
                <div>
                  <h3 className="text-xl font-bold">Gestão de Conhecimento</h3>
                  <p className="text-xs text-gray-500">Elementos que têm tratamento especial em todos os projetos</p>
                </div>
                <div className="flex items-center gap-2">
                  <label className="cursor-pointer flex items-center gap-1.5 px-3 py-1.5 bg-indigo-50 text-indigo-700 rounded-lg text-xs font-bold hover:bg-indigo-100 transition-colors" title="Importar Base de Conhecimento Completa (JSON)">
                    <Upload className="w-4 h-4" />
                    <span>Importar Tudo</span>
                    <input type="file" accept=".json" className="hidden" onChange={handleImportAllKnowledge} />
                  </label>
                  <button 
                    onClick={handleExportAllKnowledge} 
                    className="flex items-center gap-1.5 px-3 py-1.5 bg-indigo-600 text-white rounded-lg text-xs font-bold hover:bg-indigo-700 transition-colors" 
                    title="Exportar Base de Conhecimento Completa (JSON)"
                  >
                    <Download className="w-4 h-4" />
                    <span>Exportar Tudo</span>
                  </button>
                  <div className="w-px h-6 bg-gray-200 mx-1" />
                  <button onClick={() => setShowExceptionsModal(false)} className="p-2 hover:bg-gray-100 rounded-full transition-colors">
                    <X className="w-5 h-5" />
                  </button>
                </div>
              </div>

              <div className="flex border-b border-gray-100">
                <button 
                  onClick={() => setExceptionsTab('EXCECAO')}
                  className={`flex-1 py-3 text-sm font-bold transition-colors ${exceptionsTab === 'EXCECAO' ? 'text-indigo-600 border-b-2 border-indigo-600 bg-indigo-50/30' : 'text-gray-500 hover:bg-gray-50'}`}
                >
                  Exceções
                </button>
                <button 
                  onClick={() => setExceptionsTab('JUIZ')}
                  className={`flex-1 py-3 text-sm font-bold transition-colors ${exceptionsTab === 'JUIZ' ? 'text-indigo-600 border-b-2 border-indigo-600 bg-indigo-50/30' : 'text-gray-500 hover:bg-gray-50'}`}
                >
                  Juízes
                </button>
                <button 
                  onClick={() => setExceptionsTab('AUTOR')}
                  className={`flex-1 py-3 text-sm font-bold transition-colors ${exceptionsTab === 'AUTOR' ? 'text-indigo-600 border-b-2 border-indigo-600 bg-indigo-50/30' : 'text-gray-500 hover:bg-gray-50'}`}
                >
                  Autores
                </button>
                <button 
                  onClick={() => setExceptionsTab('COLETIVA')}
                  className={`flex-1 py-3 text-sm font-bold transition-colors ${exceptionsTab === 'COLETIVA' ? 'text-indigo-600 border-b-2 border-indigo-600 bg-indigo-50/30' : 'text-gray-500 hover:bg-gray-50'}`}
                >
                  Coletivas
                </button>
                <button 
                  onClick={() => setExceptionsTab('HEADER')}
                  className={`flex-1 py-3 text-sm font-bold transition-colors ${exceptionsTab === 'HEADER' ? 'text-indigo-600 border-b-2 border-indigo-600 bg-indigo-50/30' : 'text-gray-500 hover:bg-gray-50'}`}
                >
                  Cabeçalhos
                </button>
                <button 
                  onClick={() => setExceptionsTab('SAFELIST')}
                  className={`flex-1 py-3 text-sm font-bold transition-colors ${exceptionsTab === 'SAFELIST' ? 'text-indigo-600 border-b-2 border-indigo-600 bg-indigo-50/30' : 'text-gray-500 hover:bg-gray-50'}`}
                >
                  Safelist
                </button>
              </div>
              
              <div className="p-4 border-b border-gray-100 bg-gray-50/50 space-y-3">
                <div className="relative">
                  <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-gray-400" />
                  <input 
                    type="text" 
                    placeholder={`Pesquisar em ${exceptionsTab === 'EXCECAO' ? 'exceções' : exceptionsTab === 'JUIZ' ? 'juízes' : exceptionsTab === 'AUTOR' ? 'autores' : 'safelist'}...`}
                    value={knowledgeSearch}
                    onChange={(e) => setKnowledgeSearch(e.target.value)}
                    className="w-full pl-10 pr-4 py-2 bg-white border border-gray-200 rounded-xl text-sm focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 outline-none transition-all"
                  />
                </div>
                <div className="flex gap-2">
                  <input 
                    type="text" 
                    placeholder={`Adicionar novo(a) ${exceptionsTab === 'EXCECAO' ? 'exceção' : exceptionsTab === 'JUIZ' ? 'juiz' : exceptionsTab === 'AUTOR' ? 'autor' : exceptionsTab === 'HEADER' ? 'cabeçalho' : 'termo na safelist'}...`}
                    className="flex-1 px-4 py-2 bg-white border border-gray-200 rounded-xl text-sm focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 outline-none transition-all"
                    onKeyDown={(e) => {
                      if (e.key === 'Enter') {
                        const val = e.currentTarget.value.trim();
                        if (val) {
                          if (exceptionsTab === 'SAFELIST') {
                            setSafelist(prev => {
                              const isPhrase = val.includes(' ');
                              const words = isPhrase ? prev.words_ignore : Array.from(new Set([...prev.words_ignore, val]));
                              const phrases = isPhrase ? Array.from(new Set([...prev.phrases_ignore, val])) : prev.phrases_ignore;
                              return { words_ignore: words, phrases_ignore: phrases };
                            });
                          } else {
                            setGlobalKnowledge(prev => ({ ...prev, [val]: exceptionsTab }));
                          }
                          e.currentTarget.value = '';
                          showToast(`"${val}" adicionado.`, "success");
                        }
                      }
                    }}
                  />
                  <button 
                    onClick={(e) => {
                      const input = e.currentTarget.previousElementSibling as HTMLInputElement;
                      const val = input.value.trim();
                      if (val) {
                        if (exceptionsTab === 'SAFELIST') {
                          setSafelist(prev => {
                            const isPhrase = val.includes(' ');
                            const words = isPhrase ? prev.words_ignore : Array.from(new Set([...prev.words_ignore, val]));
                            const phrases = isPhrase ? Array.from(new Set([...prev.phrases_ignore, val])) : prev.phrases_ignore;
                            return { words_ignore: words, phrases_ignore: phrases };
                          });
                        } else {
                          setGlobalKnowledge(prev => ({ ...prev, [val]: exceptionsTab }));
                        }
                        input.value = '';
                        showToast(`"${val}" adicionado.`, "success");
                      }
                    }}
                    className="bg-indigo-600 text-white px-4 py-2 rounded-xl text-sm font-bold hover:bg-indigo-700 transition-all"
                  >
                    Adicionar
                  </button>
                </div>
              </div>

              <div className="p-6 overflow-y-auto flex-1">
                <div className="flex justify-between items-center mb-4">
                  <div className="flex items-center gap-4">
                    <span className="text-xs font-bold text-gray-400 uppercase tracking-wider">
                      {exceptionsTab === 'SAFELIST' 
                        ? `${safelist.words_ignore.length + safelist.phrases_ignore.length} Termos`
                        : `${Object.entries(globalKnowledge).filter(([text, type]) => type === exceptionsTab && text.toLowerCase().includes(knowledgeSearch.toLowerCase())).length} Elementos`
                      }
                    </span>
                    {exceptionsTab === 'EXCECAO' && (
                      <div className="flex items-center gap-2">
                        <button 
                          onClick={handleTransferExceptionsToSafelist}
                          className="text-xs font-bold text-indigo-600 hover:text-indigo-800 flex items-center gap-1 transition-colors bg-indigo-50 px-2 py-1 rounded"
                          title="Mover todas as exceções para a Safelist e limpar esta lista"
                        >
                          <RotateCw className="w-3 h-3" />
                          Transferir para Safelist
                        </button>
                      </div>
                    )}
                    {exceptionsTab === 'JUIZ' && (
                      <div className="flex items-center gap-2">
                        <button 
                          onClick={() => handleIdentifyDuplicates('JUIZ')}
                          className="text-xs font-bold text-amber-600 hover:text-amber-800 flex items-center gap-1 transition-colors bg-amber-50 px-2 py-1 rounded"
                          title="Identificar possíveis nomes duplicados"
                        >
                          <Copy className="w-3 h-3" />
                          Identificar Duplicados
                        </button>
                        <label className="cursor-pointer text-xs font-bold text-gray-500 hover:text-gray-700 flex items-center gap-1 transition-colors bg-gray-50 px-2 py-1 rounded" title="Importar novo ficheiro PDF/TXT de juízes">
                          <Upload className="w-3 h-3" />
                          <span>Importar Novo (PDF/TXT)</span>
                          <input type="file" accept=".pdf,.txt" className="hidden" onChange={handleImportJudges} />
                        </label>
                      </div>
                    )}
                    {exceptionsTab === 'AUTOR' && (
                      <div className="flex items-center gap-2">
                        <button 
                          onClick={() => handleIdentifyDuplicates('AUTOR')}
                          className="text-xs font-bold text-amber-600 hover:text-amber-800 flex items-center gap-1 transition-colors bg-amber-50 px-2 py-1 rounded"
                          title="Identificar possíveis nomes duplicados"
                        >
                          <Copy className="w-3 h-3" />
                          Identificar Duplicados
                        </button>
                        <label className="cursor-pointer text-xs font-bold text-gray-500 hover:text-gray-700 flex items-center gap-1 transition-colors bg-gray-50 px-2 py-1 rounded" title="Importar novo ficheiro PDF/TXT de autores">
                          <Upload className="w-3 h-3" />
                          <span>Importar Novo (PDF/TXT)</span>
                          <input type="file" accept=".pdf,.txt" className="hidden" onChange={handleImportAuthors} />
                        </label>
                      </div>
                    )}
                    {exceptionsTab === 'COLETIVA' && (
                      <div className="flex items-center gap-2">
                        <button 
                          onClick={() => handleIdentifyDuplicates('COLETIVA')}
                          className="text-xs font-bold text-amber-600 hover:text-amber-800 flex items-center gap-1 transition-colors bg-amber-50 px-2 py-1 rounded"
                          title="Identificar possíveis nomes duplicados"
                        >
                          <Copy className="w-3 h-3" />
                          Identificar Duplicados
                        </button>
                        <label className="cursor-pointer text-xs font-bold text-gray-500 hover:text-gray-700 flex items-center gap-1 transition-colors bg-gray-50 px-2 py-1 rounded" title="Importar novo ficheiro PDF/TXT de pessoas coletivas">
                          <Upload className="w-3 h-3" />
                          <span>Importar Novo (PDF/TXT)</span>
                          <input type="file" accept=".pdf,.txt" className="hidden" onChange={handleImportColetivas} />
                        </label>
                      </div>
                    )}
                    {exceptionsTab === 'SAFELIST' && (
                      <div className="flex items-center gap-2">
                        <label className="cursor-pointer text-xs font-bold text-indigo-600 hover:text-indigo-800 flex items-center gap-1 transition-colors bg-indigo-50 px-2 py-1 rounded" title="Importar Safelist (JSON ou TXT)">
                          <Upload className="w-3 h-3" />
                          <span>Importar Safelist (JSON/TXT)</span>
                          <input type="file" accept=".json,.txt" className="hidden" onChange={handleImportSafelist} />
                        </label>
                      </div>
                    )}
                  </div>
                  <div className="flex items-center gap-4">
                    <button 
                      onClick={() => {
                        if (exceptionsTab === 'SAFELIST') {
                          setSafelist({ words_ignore: [], phrases_ignore: [] });
                        } else {
                          handleClearGlobalKnowledge(exceptionsTab);
                        }
                      }}
                      className="text-xs font-bold text-red-500 hover:text-red-700 flex items-center gap-1 transition-colors"
                    >
                      <Trash2 className="w-3 h-3" />
                      Limpar {exceptionsTab === 'EXCECAO' ? 'Exceções' : exceptionsTab === 'JUIZ' ? 'Juízes' : exceptionsTab === 'AUTOR' ? 'Autores' : exceptionsTab === 'COLETIVA' ? 'Coletivas' : exceptionsTab === 'HEADER' ? 'Cabeçalhos' : 'Safelist'}
                    </button>
                    <button 
                      onClick={handleClearAllGlobalKnowledge}
                      className="text-xs font-bold text-red-700 hover:text-red-900 flex items-center gap-1 transition-colors border-l pl-4 border-gray-200"
                    >
                      <XCircle className="w-3 h-3" />
                      Limpar Tudo
                    </button>
                  </div>
                </div>

                {exceptionsTab === 'SAFELIST' ? (
                  <div className="space-y-6">
                    {safelist.phrases_ignore.length > 0 && (
                      <div>
                        <h4 className="text-xs font-bold text-gray-400 uppercase mb-2">Expressões (Phrases)</h4>
                        <div className="space-y-2">
                          {safelist.phrases_ignore
                            .filter(p => superNormalize(p).includes(superNormalize(knowledgeSearch)))
                            .sort()
                            .map(phrase => (
                              <div key={phrase} className="flex items-center justify-between bg-indigo-50/50 p-3 rounded-xl border border-indigo-100 group">
                                <span className="font-medium text-indigo-900">{phrase}</span>
                                <div className="flex items-center gap-1 opacity-0 group-hover:opacity-100 transition-opacity">
                                  <button 
                                    onClick={() => handleStartEditKnowledge(phrase, 'SAFELIST_PHRASE')}
                                    className="p-1.5 text-gray-400 hover:text-indigo-600 transition-colors"
                                    title="Editar"
                                  >
                                    <Pencil className="w-4 h-4" />
                                  </button>
                                  <button 
                                    onClick={() => setSafelist(prev => ({
                                      ...prev,
                                      phrases_ignore: prev.phrases_ignore.filter(p => p !== phrase)
                                    }))}
                                    className="p-1.5 text-gray-400 hover:text-red-500 transition-colors"
                                  >
                                    <Trash2 className="w-4 h-4" />
                                  </button>
                                </div>
                              </div>
                            ))}
                        </div>
                      </div>
                    )}
                    
                    {safelist.words_ignore.length > 0 && (
                      <div>
                        <h4 className="text-xs font-bold text-gray-400 uppercase mb-2">Palavras (Words)</h4>
                        <div className="flex flex-wrap gap-2">
                          {safelist.words_ignore
                            .filter(w => superNormalize(w).includes(superNormalize(knowledgeSearch)))
                            .sort()
                            .map(word => (
                              <div key={word} className="flex items-center gap-2 bg-gray-50 px-3 py-1.5 rounded-full border border-gray-200 group">
                                <span className="text-sm">{word}</span>
                                <button 
                                  onClick={() => handleStartEditKnowledge(word, 'SAFELIST_WORD')}
                                  className="text-gray-400 hover:text-indigo-600 transition-colors opacity-0 group-hover:opacity-100"
                                  title="Editar"
                                >
                                  <Pencil className="w-3 h-3" />
                                </button>
                                <button 
                                  onClick={() => setSafelist(prev => ({
                                    ...prev,
                                    words_ignore: prev.words_ignore.filter(w => w !== word)
                                  }))}
                                  className="text-gray-400 hover:text-red-500 transition-colors"
                                >
                                  <X className="w-3 h-3" />
                                </button>
                              </div>
                            ))}
                        </div>
                      </div>
                    )}

                    {safelist.words_ignore.length === 0 && safelist.phrases_ignore.length === 0 && (
                      <div className="text-center py-12 text-gray-400">
                        <Shield className="w-12 h-12 mx-auto mb-4 opacity-20" />
                        <p>A Safelist está vazia.</p>
                      </div>
                    )}
                  </div>
                ) : (
                  <>
                    {Object.entries(globalKnowledge).filter(([_, type]) => type === exceptionsTab).length === 0 ? (
                      <div className="text-center py-12 text-gray-400">
                        <Shield className="w-12 h-12 mx-auto mb-4 opacity-20" />
                        <p>Nenhum elemento nesta categoria.</p>
                      </div>
                    ) : (
                      <div className="space-y-2">
                        {Object.entries(globalKnowledge)
                          .filter(([text, type]) => type === exceptionsTab && superNormalize(text).includes(superNormalize(knowledgeSearch)))
                          .sort((a, b) => a[0].localeCompare(b[0]))
                          .map(([text, type]) => (
                            <div key={text} className="flex items-center justify-between bg-gray-50 p-3 rounded-xl border border-gray-100 group">
                              <span className="font-medium">{text}</span>
                              <div className="flex items-center gap-2 opacity-0 group-hover:opacity-100 transition-opacity">
                                <button 
                                  onClick={() => handleStartEditKnowledge(text, 'KNOWLEDGE', type as string)}
                                  className="p-1.5 text-gray-400 hover:text-indigo-600 transition-colors"
                                  title="Editar"
                                >
                                  <Pencil className="w-4 h-4" />
                                </button>
                                  <select 
                                    value={type as string}
                                    onChange={(e) => setGlobalKnowledge(prev => ({ ...prev, [text]: e.target.value }))}
                                    className="text-xs bg-white border border-gray-200 rounded px-2 py-1 outline-none focus:border-indigo-300"
                                  >
                                    <option value="EXCECAO">Exceção</option>
                                    <option value="JUIZ">Juiz</option>
                                    <option value="AUTOR">Autor</option>
                                    <option value="COLETIVA">Coletiva</option>
                                    <option value="HEADER">Cabeçalho</option>
                                  </select>
                                <button 
                                  onClick={() => setGlobalKnowledge(prev => {
                                    const next = { ...prev };
                                    delete next[text];
                                    return next;
                                  })}
                                  className="p-1.5 text-gray-400 hover:text-red-500 transition-colors"
                                  title="Remover"
                                >
                                  <Trash2 className="w-4 h-4" />
                                </button>
                              </div>
                            </div>
                          ))}
                      </div>
                    )}
                    {exceptionsTab === 'HEADER' && (headerZones.length > 0 || footerZones.length > 0) && (
                      <div className="mt-8 pt-8 border-t border-gray-100">
                        <h4 className="text-xs font-bold text-gray-400 uppercase mb-4">Zonas Visuais Selecionadas</h4>
                        <div className="grid grid-cols-2 gap-4">
                          {[...headerZones.map((z, i) => ({ ...z, type: 'Cabeçalho', originalIndex: i })), ...footerZones.map((z, i) => ({ ...z, type: 'Rodapé', originalIndex: i }))].map((zone, idx) => (
                            <div key={idx} className="bg-white border border-gray-200 rounded-xl overflow-hidden group">
                              <div className="p-2 bg-gray-50 border-b border-gray-100 flex items-center justify-between">
                                <span className="text-[10px] font-bold text-gray-500 uppercase">{zone.type}</span>
                                <button 
                                  onClick={() => {
                                    if (zone.type === 'Cabeçalho') setHeaderZones(prev => prev.filter((_, i) => i !== zone.originalIndex));
                                    else setFooterZones(prev => prev.filter((_, i) => i !== zone.originalIndex));
                                  }}
                                  className="p-1 text-gray-400 hover:text-red-500 transition-colors"
                                >
                                  <Trash2 className="w-3 h-3" />
                                </button>
                              </div>
                              {zone.image ? (
                                <img src={zone.image} alt="Zone" className="w-full h-24 object-contain bg-gray-100" referrerPolicy="no-referrer" />
                              ) : (
                                <div className="w-full h-24 bg-gray-100 flex items-center justify-center text-[10px] text-gray-400 italic">Sem imagem</div>
                              )}
                            </div>
                          ))}
                        </div>
                      </div>
                    )}
                  </>
                )}
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}
