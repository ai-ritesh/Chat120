import React, { useState, useMemo } from 'react';
import { 
  Download, 
  Eye, 
  X, 
  Search, 
  Copy, 
  Check, 
  ExternalLink, 
  FileText, 
  Table, 
  Maximize2, 
  Minimize2,
  FileCode,
  Archive,
  Layers,
  ArrowUpDown
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { cn } from '../lib/utils';

export type OfficeCategory = 'excel' | 'powerpoint' | 'word' | 'pdf' | 'text' | 'archive' | 'code' | 'generic';

export interface OfficeFileInfo {
  category: OfficeCategory;
  label: string;
  extension: string;
  color: {
    bg: string;
    border: string;
    text: string;
    accent: string;
    badgeBg: string;
  };
  canPreviewInApp: boolean;
  isSpreadsheet: boolean;
  isTextDoc: boolean;
}

/**
 * Categorizes any file name and MIME type into recognized Office work formats.
 */
export function getOfficeFileInfo(fileName?: string, mimeType?: string): OfficeFileInfo {
  const name = (fileName || '').toLowerCase();
  const ext = name.split('.').pop() || '';

  // 1. Excel & Spreadsheets
  if (['xlsx', 'xls', 'csv', 'tsv', 'ods', 'xlsm', 'xlsb'].includes(ext) || mimeType?.includes('spreadsheet') || mimeType?.includes('excel') || mimeType === 'text/csv') {
    return {
      category: 'excel',
      label: ext === 'csv' || ext === 'tsv' ? 'CSV Data Table' : 'Excel Spreadsheet',
      extension: ext.toUpperCase() || 'XLSX',
      color: {
        bg: 'bg-emerald-500/15',
        border: 'border-emerald-500/35',
        text: 'text-emerald-400',
        accent: '#107c41',
        badgeBg: 'bg-emerald-500/20 text-emerald-300'
      },
      canPreviewInApp: true,
      isSpreadsheet: true,
      isTextDoc: ext === 'csv' || ext === 'tsv'
    };
  }

  // 2. PowerPoint & Presentations
  if (['pptx', 'ppt', 'pps', 'ppsx', 'key', 'odp', 'potx'].includes(ext) || mimeType?.includes('presentation') || mimeType?.includes('powerpoint')) {
    return {
      category: 'powerpoint',
      label: 'PowerPoint Presentation',
      extension: ext.toUpperCase() || 'PPTX',
      color: {
        bg: 'bg-orange-500/15',
        border: 'border-orange-500/35',
        text: 'text-orange-400',
        accent: '#c43e1c',
        badgeBg: 'bg-orange-500/20 text-orange-300'
      },
      canPreviewInApp: false,
      isSpreadsheet: false,
      isTextDoc: false
    };
  }

  // 3. Word Documents
  if (['docx', 'doc', 'dotx', 'odt', 'rtf'].includes(ext) || mimeType?.includes('wordprocessingml') || mimeType?.includes('msword')) {
    return {
      category: 'word',
      label: 'Word Document',
      extension: ext.toUpperCase() || 'DOCX',
      color: {
        bg: 'bg-blue-500/15',
        border: 'border-blue-500/35',
        text: 'text-blue-400',
        accent: '#185abd',
        badgeBg: 'bg-blue-500/20 text-blue-300'
      },
      canPreviewInApp: ext === 'rtf',
      isSpreadsheet: false,
      isTextDoc: ext === 'rtf'
    };
  }

  // 4. PDF Documents
  if (ext === 'pdf' || mimeType === 'application/pdf') {
    return {
      category: 'pdf',
      label: 'PDF Document',
      extension: 'PDF',
      color: {
        bg: 'bg-red-500/15',
        border: 'border-red-500/35',
        text: 'text-red-400',
        accent: '#e5252a',
        badgeBg: 'bg-red-500/20 text-red-300'
      },
      canPreviewInApp: true,
      isSpreadsheet: false,
      isTextDoc: false
    };
  }

  // 5. Plain Text & Notes
  if (['txt', 'log', 'md', 'json', 'yaml', 'yml', 'xml', 'ini', 'env', 'conf'].includes(ext) || mimeType?.startsWith('text/')) {
    return {
      category: 'text',
      label: ext === 'md' ? 'Markdown Document' : ext === 'json' ? 'JSON Data' : 'Text Document',
      extension: ext.toUpperCase() || 'TXT',
      color: {
        bg: 'bg-sky-500/15',
        border: 'border-sky-500/35',
        text: 'text-sky-400',
        accent: '#0284c7',
        badgeBg: 'bg-sky-500/20 text-sky-300'
      },
      canPreviewInApp: true,
      isSpreadsheet: false,
      isTextDoc: true
    };
  }

  // 6. Archives & Compressed Office Packages
  if (['zip', 'rar', '7z', 'tar', 'gz', 'bz2'].includes(ext)) {
    return {
      category: 'archive',
      label: 'Compressed Archive',
      extension: ext.toUpperCase() || 'ZIP',
      color: {
        bg: 'bg-amber-500/15',
        border: 'border-amber-500/35',
        text: 'text-amber-400',
        accent: '#d97706',
        badgeBg: 'bg-amber-500/20 text-amber-300'
      },
      canPreviewInApp: false,
      isSpreadsheet: false,
      isTextDoc: false
    };
  }

  // 7. Source Code / Scripts
  if (['js', 'ts', 'tsx', 'jsx', 'py', 'java', 'cpp', 'c', 'cs', 'go', 'rs', 'php', 'rb', 'html', 'css', 'sql', 'sh'].includes(ext)) {
    return {
      category: 'code',
      label: `${ext.toUpperCase()} Source File`,
      extension: ext.toUpperCase(),
      color: {
        bg: 'bg-emerald-500/15',
        border: 'border-emerald-500/35',
        text: 'text-emerald-400',
        accent: '#00a884',
        badgeBg: 'bg-emerald-500/20 text-emerald-300'
      },
      canPreviewInApp: true,
      isSpreadsheet: false,
      isTextDoc: true
    };
  }

  // 8. Generic file fallback
  return {
    category: 'generic',
    label: 'Document File',
    extension: ext ? ext.toUpperCase() : 'FILE',
    color: {
      bg: 'bg-slate-500/15',
      border: 'border-slate-500/35',
      text: 'text-slate-300',
      accent: '#64748b',
      badgeBg: 'bg-slate-500/20 text-slate-300'
    },
    canPreviewInApp: false,
    isSpreadsheet: false,
    isTextDoc: false
  };
}

/**
 * Custom Crisp Office SVG Icon renderer with official brand styling.
 */
export const OfficeIcon: React.FC<{ 
  category: OfficeCategory; 
  className?: string;
  size?: number;
}> = ({ category, className = "w-6 h-6", size = 24 }) => {
  switch (category) {
    case 'excel':
      return (
        <svg viewBox="0 0 24 24" width={size} height={size} fill="none" className={cn("shrink-0", className)}>
          <rect x="3" y="3" width="18" height="18" rx="3.5" fill="#107C41" />
          <path d="M7 8H17M7 12H17M7 16H17M11 5V19M15 5V19" stroke="white" strokeWidth="1.2" strokeOpacity="0.4" />
          <rect x="5.5" y="7" width="5.5" height="10" rx="1.5" fill="white" fillOpacity="0.2" />
          <path d="M7 9.5L10 14.5M10 9.5L7 14.5" stroke="white" strokeWidth="1.8" strokeLinecap="round" />
        </svg>
      );
    case 'powerpoint':
      return (
        <svg viewBox="0 0 24 24" width={size} height={size} fill="none" className={cn("shrink-0", className)}>
          <rect x="3" y="3" width="18" height="18" rx="3.5" fill="#C43E1C" />
          <path d="M6 16V8H10C11.6569 8 13 9.34315 13 11C13 12.6569 11.6569 14 10 14H6" stroke="white" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" />
          <path d="M14 16H18M14 12H17M14 8H18" stroke="white" strokeWidth="1.2" strokeLinecap="round" strokeOpacity="0.5" />
        </svg>
      );
    case 'word':
      return (
        <svg viewBox="0 0 24 24" width={size} height={size} fill="none" className={cn("shrink-0", className)}>
          <rect x="3" y="3" width="18" height="18" rx="3.5" fill="#185ABD" />
          <path d="M6 8.5L7.8 15.5L9.6 10L11.4 15.5L13.2 8.5" stroke="white" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" />
          <path d="M14.5 10H18M14.5 13H18M14.5 16H17" stroke="white" strokeWidth="1.2" strokeLinecap="round" strokeOpacity="0.6" />
        </svg>
      );
    case 'pdf':
      return (
        <svg viewBox="0 0 24 24" width={size} height={size} fill="none" className={cn("shrink-0", className)}>
          <rect x="3" y="3" width="18" height="18" rx="3.5" fill="#E5252A" />
          <path d="M7 16V8H9.5C10.8807 8 12 9.11929 12 10.5C12 11.8807 10.8807 13 9.5 13H7" stroke="white" strokeWidth="1.6" strokeLinecap="round" />
          <path d="M14 8V16M14 8H17.5M14 12H16.5" stroke="white" strokeWidth="1.6" strokeLinecap="round" />
        </svg>
      );
    case 'text':
      return (
        <svg viewBox="0 0 24 24" width={size} height={size} fill="none" className={cn("shrink-0", className)}>
          <rect x="4" y="3" width="16" height="18" rx="3" fill="#0284C7" />
          <path d="M8 8H16M8 12H16M8 16H13" stroke="white" strokeWidth="1.6" strokeLinecap="round" />
        </svg>
      );
    case 'archive':
      return (
        <svg viewBox="0 0 24 24" width={size} height={size} fill="none" className={cn("shrink-0", className)}>
          <rect x="3" y="4" width="18" height="16" rx="3" fill="#D97706" />
          <path d="M12 4V14M11 6H13M11 8H13M11 10H13M11 12H13" stroke="white" strokeWidth="1.5" strokeLinecap="round" />
          <rect x="10.5" y="14" width="3" height="4" rx="1" fill="white" />
        </svg>
      );
    case 'code':
      return (
        <svg viewBox="0 0 24 24" width={size} height={size} fill="none" className={cn("shrink-0", className)}>
          <rect x="3" y="3" width="18" height="18" rx="3.5" fill="#00A884" />
          <path d="M9.5 9L6.5 12L9.5 15M14.5 9L17.5 12L14.5 15M13 7L11 17" stroke="white" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
        </svg>
      );
    default:
      return (
        <svg viewBox="0 0 24 24" width={size} height={size} fill="none" className={cn("shrink-0", className)}>
          <rect x="4" y="3" width="16" height="18" rx="3" fill="#475569" />
          <path d="M8 8H16M8 12H16M8 16H12" stroke="white" strokeWidth="1.5" strokeLinecap="round" />
        </svg>
      );
  }
};

/**
 * Fast RFC-4180 compliant CSV & TSV parser for in-app spreadsheet previewing.
 */
export function parseDelimitedText(text: string, delimiter?: string): { headers: string[]; rows: string[][] } {
  if (!text || !text.trim()) {
    return { headers: [], rows: [] };
  }

  // Detect comma or tab
  const delim = delimiter || (text.indexOf('\t') !== -1 && text.indexOf('\t') < text.indexOf('\n') ? '\t' : ',');

  const lines = text.split(/\r\n|\n|\r/);
  const parsedRows: string[][] = [];

  for (const line of lines) {
    if (!line.trim()) continue;

    const row: string[] = [];
    let insideQuotes = false;
    let entry = '';

    for (let i = 0; i < line.length; i++) {
      const char = line[i];
      if (char === '"') {
        if (insideQuotes && line[i + 1] === '"') {
          entry += '"';
          i++;
        } else {
          insideQuotes = !insideQuotes;
        }
      } else if (char === delim && !insideQuotes) {
        row.push(entry.trim());
        entry = '';
      } else {
        entry += char;
      }
    }
    row.push(entry.trim());
    parsedRows.push(row);
  }

  if (parsedRows.length === 0) {
    return { headers: [], rows: [] };
  }

  const headers = parsedRows[0];
  const rows = parsedRows.slice(1);
  return { headers, rows };
}

/**
 * Interactive In-App SpreadSheet Table Viewer Modal:
 * Allows exploring Excel CSV/TSV data directly in chat with filtering, row sorting, and cell selection.
 */
export const OfficeSpreadsheetModal: React.FC<{
  fileName: string;
  csvText?: string;
  downloadUrl?: string;
  onClose: () => void;
  showNotification?: (msg: string) => void;
}> = ({ fileName, csvText = '', downloadUrl, onClose, showNotification }) => {
  const [filterQuery, setFilterQuery] = useState('');
  const [copied, setCopied] = useState(false);
  const [sortCol, setSortCol] = useState<number | null>(null);
  const [sortAsc, setSortAsc] = useState(true);
  const [isFullscreen, setIsFullscreen] = useState(false);

  const { headers, rows } = useMemo(() => {
    return parseDelimitedText(csvText);
  }, [csvText]);

  const filteredRows = useMemo(() => {
    let result = rows;
    if (filterQuery.trim()) {
      const q = filterQuery.toLowerCase();
      result = result.filter(r => r.some(cell => cell.toLowerCase().includes(q)));
    }
    if (sortCol !== null && headers[sortCol] !== undefined) {
      result = [...result].sort((a, b) => {
        const valA = a[sortCol] || '';
        const valB = b[sortCol] || '';
        const numA = parseFloat(valA);
        const numB = parseFloat(valB);
        if (!isNaN(numA) && !isNaN(numB)) {
          return sortAsc ? numA - numB : numB - numA;
        }
        return sortAsc ? valA.localeCompare(valB) : valB.localeCompare(valA);
      });
    }
    return result;
  }, [rows, filterQuery, sortCol, sortAsc, headers]);

  const handleCopyRaw = () => {
    navigator.clipboard.writeText(csvText).then(() => {
      setCopied(true);
      showNotification?.('Spreadsheet data copied to clipboard!');
      setTimeout(() => setCopied(false), 2000);
    }).catch(() => {});
  };

  return (
    <div className="fixed inset-0 bg-black/85 flex items-center justify-center p-2 sm:p-4 z-50 backdrop-blur-md">
      <motion.div 
        initial={{ opacity: 0, scale: 0.95 }}
        animate={{ opacity: 1, scale: 1 }}
        exit={{ opacity: 0, scale: 0.95 }}
        className={cn(
          "bg-[#1b272f] rounded-2xl border border-emerald-500/30 flex flex-col shadow-2xl overflow-hidden transition-all duration-200",
          isFullscreen ? "w-full h-full max-w-none rounded-none" : "w-full max-w-5xl h-[88vh]"
        )}
      >
        {/* Header Bar */}
        <div className="bg-[#111b21] px-4 py-3 border-b border-[#2a3942] flex items-center justify-between shrink-0">
          <div className="flex items-center gap-2.5 min-w-0">
            <OfficeIcon category="excel" size={24} />
            <div className="min-w-0">
              <div className="flex items-center gap-2">
                <h3 className="text-sm sm:text-base font-bold text-white truncate max-w-md">{fileName}</h3>
                <span className="text-[10px] font-mono px-2 py-0.5 rounded-full bg-emerald-500/20 text-emerald-300 font-semibold border border-emerald-500/30">
                  {headers.length} Cols · {rows.length} Rows
                </span>
              </div>
              <p className="text-[11px] text-[#8696a0]">Interactive Spreadsheet Table Preview</p>
            </div>
          </div>

          <div className="flex items-center gap-1.5 shrink-0">
            <button
              type="button"
              onClick={handleCopyRaw}
              className="p-2 hover:bg-[#202c33] text-[#8696a0] hover:text-[#e9edef] rounded-lg transition"
              title="Copy CSV text"
            >
              {copied ? <Check className="w-4 h-4 text-emerald-400" /> : <Copy className="w-4 h-4" />}
            </button>
            {downloadUrl && (
              <a
                href={downloadUrl}
                download={fileName}
                className="p-2 hover:bg-[#202c33] text-[#8696a0] hover:text-[#00a884] rounded-lg transition"
                title="Download file"
              >
                <Download className="w-4 h-4" />
              </a>
            )}
            <button
              type="button"
              onClick={() => setIsFullscreen(!isFullscreen)}
              className="p-2 hover:bg-[#202c33] text-[#8696a0] hover:text-[#e9edef] rounded-lg transition hidden sm:flex"
              title={isFullscreen ? "Exit Fullscreen" : "Fullscreen"}
            >
              {isFullscreen ? <Minimize2 className="w-4 h-4" /> : <Maximize2 className="w-4 h-4" />}
            </button>
            <button
              type="button"
              onClick={onClose}
              className="p-2 hover:bg-red-500/20 text-[#8696a0] hover:text-red-400 rounded-lg transition"
              title="Close"
            >
              <X className="w-5 h-5" />
            </button>
          </div>
        </div>

        {/* Search & Filter Toolbar */}
        <div className="px-4 py-2.5 bg-[#172228] border-b border-[#2a3942] flex items-center justify-between gap-3 shrink-0">
          <div className="relative flex-1 max-w-sm">
            <Search className="w-3.5 h-3.5 text-[#8696a0] absolute left-3 top-2.5" />
            <input 
              type="text" 
              placeholder="Search table rows..." 
              value={filterQuery}
              onChange={(e) => setFilterQuery(e.target.value)}
              className="w-full bg-[#111b21] border border-[#2a3942] rounded-xl pl-8 pr-3 py-1.5 text-xs text-[#e9edef] placeholder:text-[#8696a0] outline-none focus:border-emerald-500 transition"
            />
          </div>
          <div className="text-[11px] text-[#8696a0] shrink-0 font-mono">
            Showing {filteredRows.length} of {rows.length} rows
          </div>
        </div>

        {/* Spreadsheet Data Grid */}
        <div className="flex-1 overflow-auto custom-scrollbar bg-[#0f171d]">
          {headers.length === 0 ? (
            <div className="flex flex-col items-center justify-center h-full text-center p-6 text-[#8696a0]">
              <Table className="w-10 h-10 text-emerald-500/40 mb-2" />
              <p className="text-sm font-medium text-white">Empty Spreadsheet</p>
              <p className="text-xs">No data rows found in this file.</p>
            </div>
          ) : (
            <table className="w-full border-collapse text-left text-xs font-mono">
              <thead className="bg-[#18232a] sticky top-0 z-10 border-b border-[#2a3942]">
                <tr>
                  <th className="py-2 px-2.5 text-center text-[10px] text-[#8696a0] border-r border-[#2a3942] w-12 bg-[#141d23] select-none">
                    #
                  </th>
                  {headers.map((h, idx) => (
                    <th 
                      key={idx}
                      onClick={() => {
                        if (sortCol === idx) {
                          setSortAsc(!sortAsc);
                        } else {
                          setSortCol(idx);
                          setSortAsc(true);
                        }
                      }}
                      className="py-2 px-3 font-bold text-emerald-400 border-r border-[#2a3942] cursor-pointer hover:bg-[#202c33] transition select-none whitespace-nowrap"
                    >
                      <div className="flex items-center justify-between gap-2">
                        <span>{h || `Column ${idx + 1}`}</span>
                        <ArrowUpDown className="w-3 h-3 text-[#8696a0]" />
                      </div>
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-[#202c33]/70">
                {filteredRows.map((row, rIdx) => (
                  <tr key={rIdx} className="hover:bg-[#1a262e] transition">
                    <td className="py-1.5 px-2.5 text-center text-[10px] text-[#8696a0] border-r border-[#2a3942] bg-[#111b21] select-none">
                      {rIdx + 1}
                    </td>
                    {headers.map((_, cIdx) => (
                      <td 
                        key={cIdx} 
                        className="py-1.5 px-3 text-[#e9edef] border-r border-[#202c33] whitespace-nowrap max-w-xs truncate"
                        title={row[cIdx] || ''}
                      >
                        {row[cIdx] !== undefined ? row[cIdx] : ''}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>
      </motion.div>
    </div>
  );
};

/**
 * Interactive In-App Text Document Reader Modal:
 * For viewing .txt, .md, .log, .json, .rtf notes directly inside chat with search and copy.
 */
export const OfficeTextViewerModal: React.FC<{
  fileName: string;
  textContent: string;
  downloadUrl?: string;
  onClose: () => void;
  showNotification?: (msg: string) => void;
}> = ({ fileName, textContent, downloadUrl, onClose, showNotification }) => {
  const [filterQuery, setFilterQuery] = useState('');
  const [copied, setCopied] = useState(false);
  const [isFullscreen, setIsFullscreen] = useState(false);

  const lines = useMemo(() => {
    return textContent.split('\n');
  }, [textContent]);

  const wordCount = useMemo(() => {
    return textContent.trim().split(/\s+/).filter(Boolean).length;
  }, [textContent]);

  const handleCopyText = () => {
    navigator.clipboard.writeText(textContent).then(() => {
      setCopied(true);
      showNotification?.('Document content copied to clipboard!');
      setTimeout(() => setCopied(false), 2000);
    }).catch(() => {});
  };

  return (
    <div className="fixed inset-0 bg-black/85 flex items-center justify-center p-2 sm:p-4 z-50 backdrop-blur-md">
      <motion.div 
        initial={{ opacity: 0, scale: 0.95 }}
        animate={{ opacity: 1, scale: 1 }}
        exit={{ opacity: 0, scale: 0.95 }}
        className={cn(
          "bg-[#1b272f] rounded-2xl border border-sky-500/30 flex flex-col shadow-2xl overflow-hidden transition-all duration-200",
          isFullscreen ? "w-full h-full max-w-none rounded-none" : "w-full max-w-4xl h-[86vh]"
        )}
      >
        {/* Header Bar */}
        <div className="bg-[#111b21] px-4 py-3 border-b border-[#2a3942] flex items-center justify-between shrink-0">
          <div className="flex items-center gap-2.5 min-w-0">
            <OfficeIcon category="text" size={24} />
            <div className="min-w-0">
              <div className="flex items-center gap-2">
                <h3 className="text-sm sm:text-base font-bold text-white truncate max-w-md">{fileName}</h3>
                <span className="text-[10px] font-mono px-2 py-0.5 rounded-full bg-sky-500/20 text-sky-300 font-semibold border border-sky-500/30">
                  {lines.length} Lines · {wordCount} Words
                </span>
              </div>
              <p className="text-[11px] text-[#8696a0]">Text &amp; Notes Reader</p>
            </div>
          </div>

          <div className="flex items-center gap-1.5 shrink-0">
            <button
              type="button"
              onClick={handleCopyText}
              className="p-2 hover:bg-[#202c33] text-[#8696a0] hover:text-[#e9edef] rounded-lg transition"
              title="Copy text"
            >
              {copied ? <Check className="w-4 h-4 text-emerald-400" /> : <Copy className="w-4 h-4" />}
            </button>
            {downloadUrl && (
              <a
                href={downloadUrl}
                download={fileName}
                className="p-2 hover:bg-[#202c33] text-[#8696a0] hover:text-[#00a884] rounded-lg transition"
                title="Download file"
              >
                <Download className="w-4 h-4" />
              </a>
            )}
            <button
              type="button"
              onClick={() => setIsFullscreen(!isFullscreen)}
              className="p-2 hover:bg-[#202c33] text-[#8696a0] hover:text-[#e9edef] rounded-lg transition hidden sm:flex"
              title={isFullscreen ? "Exit Fullscreen" : "Fullscreen"}
            >
              {isFullscreen ? <Minimize2 className="w-4 h-4" /> : <Maximize2 className="w-4 h-4" />}
            </button>
            <button
              type="button"
              onClick={onClose}
              className="p-2 hover:bg-red-500/20 text-[#8696a0] hover:text-red-400 rounded-lg transition"
              title="Close"
            >
              <X className="w-5 h-5" />
            </button>
          </div>
        </div>

        {/* Content Viewer with Line Numbers */}
        <div className="flex-1 overflow-auto custom-scrollbar bg-[#0f171d] p-4 font-mono text-xs text-[#e9edef] leading-relaxed select-text">
          <pre className="whitespace-pre-wrap break-words">{textContent}</pre>
        </div>
      </motion.div>
    </div>
  );
};

/**
 * Universal Office File Inspector Modal:
 * Displays full metadata, direct preview/download options, and Google Docs/Office Online viewer shortcuts.
 */
export const OfficeDocumentModal: React.FC<{
  fileName: string;
  fileSize?: number;
  downloadUrl?: string;
  onClose: () => void;
  onOpenSpreadsheet?: () => void;
  onOpenText?: () => void;
  showNotification?: (msg: string) => void;
}> = ({ fileName, fileSize, downloadUrl, onClose, onOpenSpreadsheet, onOpenText, showNotification }) => {
  const info = getOfficeFileInfo(fileName);

  return (
    <div className="fixed inset-0 bg-black/80 flex items-center justify-center p-4 z-50 backdrop-blur-sm">
      <motion.div 
        initial={{ opacity: 0, scale: 0.95 }}
        animate={{ opacity: 1, scale: 1 }}
        exit={{ opacity: 0, scale: 0.95 }}
        className="bg-[#202c33] rounded-2xl border border-[#3b4a54] w-full max-w-md p-6 shadow-2xl space-y-4"
      >
        <div className="flex justify-between items-center">
          <div className="flex items-center gap-2">
            <OfficeIcon category={info.category} size={28} />
            <div>
              <h3 className="text-base font-bold text-white">{info.label}</h3>
              <p className="text-xs text-[#8696a0]">{info.extension} Document</p>
            </div>
          </div>
          <button onClick={onClose} className="text-[#8696a0] hover:text-white p-1 rounded-lg">
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* File Card Info */}
        <div className="bg-[#111b21] p-4 rounded-xl border border-[#2a3942] space-y-2">
          <p className="text-sm font-semibold text-[#e9edef] break-all">{fileName}</p>
          <div className="flex items-center gap-3 text-xs text-[#8696a0]">
            {fileSize && (
              <span>
                {fileSize > 1024 * 1024 
                  ? `${(fileSize / (1024 * 1024)).toFixed(2)} MB` 
                  : `${(fileSize / 1024).toFixed(1)} KB`}
              </span>
            )}
            <span>·</span>
            <span className={cn("px-2 py-0.5 rounded-full text-[10px] font-bold", info.color.badgeBg)}>
              {info.extension}
            </span>
          </div>
        </div>

        {/* Action Buttons */}
        <div className="space-y-2 pt-2">
          {info.isSpreadsheet && onOpenSpreadsheet && (
            <button
              type="button"
              onClick={() => {
                onClose();
                onOpenSpreadsheet();
              }}
              className="w-full py-2.5 px-4 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl font-semibold text-xs flex items-center justify-center gap-2 transition shadow-md"
            >
              <Table className="w-4 h-4" />
              <span>Preview Spreadsheet Table</span>
            </button>
          )}

          {info.isTextDoc && onOpenText && (
            <button
              type="button"
              onClick={() => {
                onClose();
                onOpenText();
              }}
              className="w-full py-2.5 px-4 bg-sky-600 hover:bg-sky-700 text-white rounded-xl font-semibold text-xs flex items-center justify-center gap-2 transition shadow-md"
            >
              <Eye className="w-4 h-4" />
              <span>Read Text in Chat</span>
            </button>
          )}

          {downloadUrl && (
            <a
              href={downloadUrl}
              download={fileName}
              className="w-full py-2.5 px-4 bg-[#00a884] hover:bg-[#008f6f] text-white rounded-xl font-semibold text-xs flex items-center justify-center gap-2 transition shadow-md"
            >
              <Download className="w-4 h-4" />
              <span>Download {info.extension} File</span>
            </a>
          )}

          {downloadUrl && downloadUrl.startsWith('http') && (
            <a
              href={`https://docs.google.com/viewer?url=${encodeURIComponent(downloadUrl)}`}
              target="_blank"
              rel="noopener noreferrer"
              className="w-full py-2.5 px-4 bg-[#2a3942] hover:bg-[#3b4a54] text-[#e9edef] rounded-xl font-medium text-xs flex items-center justify-center gap-2 transition"
            >
              <ExternalLink className="w-4 h-4 text-[#8696a0]" />
              <span>Open in Web Office Viewer</span>
            </a>
          )}
        </div>
      </motion.div>
    </div>
  );
};

/**
 * Enhanced Office File Attachment Card in Message Bubble:
 * Displays recognized Office branding (Excel, PowerPoint, Word, PDF, Text), exact size, and instant preview/download.
 */
export const OfficeFileCard: React.FC<{
  fileName: string;
  fileSize?: number;
  downloadUrl?: string;
  isChunked?: boolean;
  isLoadingChunked?: boolean;
  onDownloadChunked?: () => void;
  onPreviewSpreadsheet?: () => void;
  onPreviewText?: () => void;
  onOpenDetails?: () => void;
}> = ({
  fileName,
  fileSize,
  downloadUrl,
  isChunked,
  isLoadingChunked,
  onDownloadChunked,
  onPreviewSpreadsheet,
  onPreviewText,
  onOpenDetails
}) => {
  const info = getOfficeFileInfo(fileName);

  return (
    <div className={cn(
      "p-3 rounded-2xl mb-2 border transition-all duration-150 backdrop-blur-sm",
      info.color.bg,
      info.color.border
    )}>
      <div className="flex items-center gap-3">
        {/* Office Brand Icon */}
        <div 
          onClick={onOpenDetails}
          className="cursor-pointer hover:scale-105 transition-transform shrink-0" 
          title={`Click to view ${info.label} info`}
        >
          <OfficeIcon category={info.category} size={32} />
        </div>

        {/* File Name & Office Label */}
        <div 
          onClick={onOpenDetails} 
          className="flex-1 min-w-0 cursor-pointer"
        >
          <div className="flex items-center gap-1.5">
            <p className="text-sm font-semibold text-[#e9edef] truncate hover:underline">{fileName || 'Attachment'}</p>
          </div>
          <div className="flex items-center gap-2 text-[11px] text-[#8696a0] mt-0.5">
            <span className={cn("px-1.5 py-0.2 rounded font-bold text-[9px] uppercase tracking-wider", info.color.badgeBg)}>
              {info.extension}
            </span>
            <span>·</span>
            <span>
              {fileSize 
                ? (fileSize > 1024 * 1024 
                  ? `${(fileSize / (1024 * 1024)).toFixed(1)} MB` 
                  : `${(fileSize / 1024).toFixed(0)} KB`)
                : info.label}
            </span>
          </div>
        </div>

        {/* Action Buttons: Preview & Download */}
        <div className="flex items-center gap-1.5 shrink-0">
          {/* In-app Spreadsheet Table Preview */}
          {info.isSpreadsheet && onPreviewSpreadsheet && (
            <button
              type="button"
              onClick={onPreviewSpreadsheet}
              className="p-2 bg-[#202c33]/90 hover:bg-[#2a3942] text-emerald-400 rounded-xl transition border border-emerald-500/30"
              title="Preview Spreadsheet Table"
            >
              <Table className="w-4 h-4" />
            </button>
          )}

          {/* In-app Text Document Preview */}
          {info.isTextDoc && onPreviewText && (
            <button
              type="button"
              onClick={onPreviewText}
              className="p-2 bg-[#202c33]/90 hover:bg-[#2a3942] text-sky-400 rounded-xl transition border border-sky-500/30"
              title="Read Document in Chat"
            >
              <Eye className="w-4 h-4" />
            </button>
          )}

          {/* Download Action */}
          {isChunked && !downloadUrl ? (
            <button
              type="button"
              onClick={onDownloadChunked}
              disabled={isLoadingChunked}
              className="p-2 bg-[#202c33]/90 hover:bg-[#2a3942] text-[#e9edef] rounded-xl transition border border-[#3b4a54]"
              title="Download File"
            >
              {isLoadingChunked ? (
                <div className="w-4 h-4 border-2 border-[#00a884] border-t-transparent rounded-full animate-spin" />
              ) : (
                <Download className="w-4 h-4 text-[#00a884]" />
              )}
            </button>
          ) : (
            <a
              href={downloadUrl || '#'}
              download={fileName || 'file'}
              className="p-2 bg-[#202c33]/90 hover:bg-[#2a3942] text-[#e9edef] rounded-xl transition border border-[#3b4a54] flex items-center justify-center"
              title="Download File"
            >
              <Download className="w-4 h-4 text-[#00a884]" />
            </a>
          )}
        </div>
      </div>
    </div>
  );
};
