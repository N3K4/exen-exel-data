import { useState, useCallback, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx';

interface ExcelData {
  headers: string[];
  rows: string[][];
}

// Форматирование даты в читаемый вид (ДД.ММ.ГГГГ)
function formatDate(value: any): string {
  if (value instanceof Date) {
    // Используем UTC-методы, чтобы избежать сдвига из-за часовых зон
    const day = String(value.getUTCDate()).padStart(2, '0');
    const month = String(value.getUTCMonth() + 1).padStart(2, '0');
    const year = value.getUTCFullYear();
    return `${day}.${month}.${year}`;
  }
  return String(value);
}

// Проверка, является ли формат ячейки форматом даты
function isDateFormat(fmt: string): boolean {
  if (!fmt) return false;
  return /[dmyDММY]/i.test(fmt) && !/^[#0,.]+$/.test(fmt);
}

// Форматирование даты из Excel serial number
function formatExcelDate(dateCode: { y: number; m: number; d: number }): string {
  const day = String(dateCode.d).padStart(2, '0');
  const month = String(dateCode.m).padStart(2, '0');
  const year = dateCode.y;
  return `${day}.${month}.${year}`;
}

// Fallback: конвертация серийного номера Excel в дату
function serialToDate(serial: number): { y: number; m: number; d: number } | null {
  if (serial < 1 || serial > 200000) return null;
  
  // Excel base date: 31 декабря 1899
  // serial 1 = 1 января 1900
  // Excel считает 1900 високосным годом (баг), поэтому для serial > 60 нужно вычитать 1
  let adjustedSerial = serial;
  if (serial > 60) {
    adjustedSerial = serial - 1;
  }
  
  const excelEpoch = new Date(Date.UTC(1899, 11, 31));
  const date = new Date(excelEpoch.getTime() + adjustedSerial * 86400000);
  
  const year = date.getUTCFullYear();
  if (year < 1900 || year > 2100) return null;
  
  return {
    y: year,
    m: date.getUTCMonth() + 1,
    d: date.getUTCDate()
  };
}

// Конвертация ячейки-числа в дату
function convertNumericDate(cellValue: number): string {
  try {
    const dateCode = XLSX.SSF.parse_date_code(cellValue);
    if (dateCode) {
      return formatExcelDate(dateCode);
    }
  } catch {
    // Fallback
  }
  
  const dateCode = serialToDate(cellValue);
  if (dateCode) {
    return formatExcelDate(dateCode);
  }
  
  return String(cellValue);
}

function App() {
  const [data, setData] = useState<ExcelData | null>(null);
  const [currentIndex, setCurrentIndex] = useState(0);
  const [copiedCell, setCopiedCell] = useState<string | null>(null);
  const [fileName, setFileName] = useState<string>('');
  const fileInputRef = useRef<HTMLInputElement>(null);
  const copyTimeoutRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  const handleFileUpload = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setFileName(file.name);
    const reader = new FileReader();

    reader.onload = (event) => {
      const arrayBuffer = event.target?.result;
      // cellDates: false — не конвертируем автоматически, делаем это вручную
      const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: false, cellNF: true });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];

      if (!sheet['!ref']) return;
      const range = XLSX.utils.decode_range(sheet['!ref']);

      const allRows: string[][] = [];

      for (let R = range.s.r; R <= range.e.r; R++) {
        const row: string[] = [];
        for (let C = range.s.c; C <= range.e.c; C++) {
          const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
          const cell = sheet[cellRef];

          if (!cell) {
            row.push('');
            continue;
          }

          const cellFormat = cell.z || '';
          const isDate = isDateFormat(cellFormat);

          // Все числовые значения с форматом даты конвертируем вручную
          if (cell.t === 'n' && isDate) {
            row.push(convertNumericDate(cell.v));
          } else {
            row.push(cell.w !== undefined ? String(cell.w) : String(cell.v));
          }
        }
        allRows.push(row);
      }

      if (allRows.length > 0) {
        const headers = allRows[0].map((h) => String(h));
        const rows = allRows.slice(1).map((row) => {
          const paddedRow = [...row];
          while (paddedRow.length < headers.length) {
            paddedRow.push('');
          }
          return paddedRow;
        });

        setData({ headers, rows });
        setCurrentIndex(0);
      }
    };

    reader.readAsArrayBuffer(file);
  }, []);

  const goToPrevious = useCallback(() => {
    setCurrentIndex((prev) => Math.max(0, prev - 1));
  }, []);

  const goToNext = useCallback(() => {
    if (data) {
      setCurrentIndex((prev) => Math.min(data.rows.length - 1, prev + 1));
    }
  }, [data]);

  const goToFirst = useCallback(() => setCurrentIndex(0), []);

  const goToLast = useCallback(() => {
    if (data) setCurrentIndex(data.rows.length - 1);
  }, [data]);

  const copyToClipboard = useCallback(async (text: string, cellId: string) => {
    try {
      await navigator.clipboard.writeText(text);
    } catch {
      const textArea = document.createElement('textarea');
      textArea.value = text;
      document.body.appendChild(textArea);
      textArea.select();
      document.execCommand('copy');
      document.body.removeChild(textArea);
    }
    setCopiedCell(cellId);
    if (copyTimeoutRef.current) clearTimeout(copyTimeoutRef.current);
    copyTimeoutRef.current = setTimeout(() => setCopiedCell(null), 1200);
  }, []);

  // Keyboard navigation
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      const target = e.target as HTMLElement;
      if (target.tagName === 'INPUT' || target.tagName === 'TEXTAREA') return;

      if (e.key === 'ArrowUp') {
        e.preventDefault();
        goToPrevious();
      } else if (e.key === 'ArrowDown') {
        e.preventDefault();
        goToNext();
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [goToPrevious, goToNext]);

  // Mouse wheel navigation
  useEffect(() => {
    if (!data) return;

    let lastWheelTime = 0;
    const THROTTLE_MS = 120;

    const handleWheel = (e: WheelEvent) => {
      const target = e.target as HTMLElement;
      if (target.tagName === 'INPUT' || target.tagName === 'TEXTAREA') return;

      const now = Date.now();
      if (now - lastWheelTime < THROTTLE_MS) return;

      if (e.deltaY > 0) {
        e.preventDefault();
        goToNext();
        lastWheelTime = now;
      } else if (e.deltaY < 0) {
        e.preventDefault();
        goToPrevious();
        lastWheelTime = now;
      }
    };

    window.addEventListener('wheel', handleWheel, { passive: false });
    return () => window.removeEventListener('wheel', handleWheel);
  }, [data, goToPrevious, goToNext]);

  const getRow = (index: number): string[] | null => {
    if (!data || index < 0 || index >= data.rows.length) return null;
    return data.rows[index];
  };

  // Upload screen
  if (!data) {
    return (
      <div className="min-h-screen flex items-center justify-center p-4">
        <div className="w-full max-w-sm">
          <div className="mb-8">
            <div className="flex items-center gap-2.5 mb-3">
              <div className="w-8 h-8 rounded-lg bg-gradient-to-br from-blue-500 to-indigo-600 flex items-center justify-center shadow-sm">
                <svg className="w-4 h-4 text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth={2}>
                  <path strokeLinecap="round" strokeLinejoin="round" d="M3 10h18M3 14h18m-9-4v8m-7 0h14a2 2 0 002-2V8a2 2 0 00-2-2H5a2 2 0 00-2 2v8a2 2 0 002 2z" />
                </svg>
              </div>
              <h1 className="text-xl font-semibold text-gray-900 tracking-tight">Excel Навигатор</h1>
            </div>
            <p className="text-sm text-gray-500 leading-relaxed">
              Загрузите файл для построчного просмотра и быстрого копирования данных
            </p>
          </div>

          <label className="flex items-center justify-center gap-2 w-full py-2.5 px-4 bg-blue-600 hover:bg-blue-700 text-white text-sm font-medium rounded-lg cursor-pointer shadow-sm hover:shadow-md">
            <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth={2}>
              <path strokeLinecap="round" strokeLinejoin="round" d="M4 16v2a2 2 0 002 2h12a2 2 0 002-2v-2M7 10l5-5 5 5M12 5v12" />
            </svg>
            Выбрать файл
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls,.csv"
              onChange={handleFileUpload}
              className="hidden"
            />
          </label>

          <p className="text-xs text-gray-400 mt-3 text-center">
            .xlsx · .xls · .csv
          </p>
        </div>
      </div>
    );
  }

  const currentRow = getRow(currentIndex);

  const renderRow = (row: string[] | null, label: string) => {
    if (!row) return null;

    return (
      <div className="rounded-xl overflow-hidden card-shadow ring-2 ring-blue-200 bg-white">
        <div className="px-4 py-2 border-b bg-blue-50 text-blue-900 border-blue-100 text-xs font-medium">
          {label}
        </div>

        <table className="w-full">
          <tbody>
            {data.headers.map((header, colIdx) => {
              const cellId = `curr-${colIdx}`;
              const isCopied = copiedCell === cellId;
              const value = row[colIdx];

              return (
                <tr
                  key={cellId}
                  className={`cell-hover ${isCopied ? 'copied-flash' : ''} border-b border-gray-100 last:border-b-0`}
                >
                  <td className="px-4 py-2.5 text-xs text-gray-500 font-medium w-36 align-top">
                    {header}
                  </td>
                  <td
                    className="px-4 py-2.5 text-sm text-gray-900 cursor-pointer"
                    onClick={() => copyToClipboard(value, cellId)}
                    title="Клик — скопировать"
                  >
                    <div className="flex items-center justify-between gap-3">
                      <span className="break-all leading-relaxed">{value || <span className="text-gray-300 italic">—</span>}</span>
                      {isCopied && (
                        <span className="text-xs text-green-700 font-medium whitespace-nowrap flex-shrink-0 bg-green-50 px-2 py-0.5 rounded-full">
                          ✓ скопировано
                        </span>
                      )}
                    </div>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    );
  };

  return (
    <div className="min-h-screen">
      <div className="sticky top-0 z-10 bg-white/90 backdrop-blur-md border-b border-gray-200/80">
        <div className="max-w-3xl mx-auto px-4 py-3">
          <div className="flex items-center justify-between gap-4 flex-wrap">
            <div className="flex items-center gap-3 min-w-0">
              <div className="min-w-0">
                <div className="text-sm font-semibold text-gray-900 truncate">{fileName}</div>
                <div className="text-xs text-gray-400">
                  {data.rows.length} строк · {data.headers.length} столбцов
                </div>
              </div>
              <label className="flex-shrink-0 text-xs text-blue-600 hover:text-blue-700 cursor-pointer font-medium">
                Заменить
                <input
                  type="file"
                  accept=".xlsx,.xls,.csv"
                  onChange={handleFileUpload}
                  className="hidden"
                />
              </label>
            </div>

            <div className="flex items-center gap-1">
              <button
                onClick={goToFirst}
                disabled={currentIndex === 0}
                className="w-7 h-7 flex items-center justify-center rounded-md text-gray-400 hover:bg-gray-100 hover:text-gray-700 disabled:opacity-30 disabled:cursor-not-allowed"
                title="Первая"
              >
                <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth={2.5}>
                  <path strokeLinecap="round" strokeLinejoin="round" d="M11 19l-7-7 7-7m8 14l-7-7 7-7" />
                </svg>
              </button>
              <button
                onClick={goToPrevious}
                disabled={currentIndex === 0}
                className="w-7 h-7 flex items-center justify-center rounded-md text-gray-400 hover:bg-gray-100 hover:text-gray-700 disabled:opacity-30 disabled:cursor-not-allowed"
                title="Предыдущая (↑)"
              >
                <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth={2.5}>
                  <path strokeLinecap="round" strokeLinejoin="round" d="M15 19l-7-7 7-7" />
                </svg>
              </button>

              <div className="flex items-center gap-1.5 px-2.5 py-1.5 bg-gray-100/80 rounded-md text-sm">
                <input
                  type="number"
                  min={1}
                  max={data.rows.length}
                  value={currentIndex + 1}
                  onChange={(e) => {
                    const val = parseInt(e.target.value);
                    if (val >= 1 && val <= data.rows.length) {
                      setCurrentIndex(val - 1);
                    }
                  }}
                  className="w-9 text-center bg-transparent text-gray-900 font-semibold text-sm focus:outline-none"
                />
                <span className="text-gray-400 text-xs font-medium">/ {data.rows.length}</span>
              </div>

              <button
                onClick={goToNext}
                disabled={currentIndex >= data.rows.length - 1}
                className="w-7 h-7 flex items-center justify-center rounded-md text-gray-400 hover:bg-gray-100 hover:text-gray-700 disabled:opacity-30 disabled:cursor-not-allowed"
                title="Следующая (↓)"
              >
                <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth={2.5}>
                  <path strokeLinecap="round" strokeLinejoin="round" d="M9 5l7 7-7 7" />
                </svg>
              </button>
              <button
                onClick={goToLast}
                disabled={currentIndex >= data.rows.length - 1}
                className="w-7 h-7 flex items-center justify-center rounded-md text-gray-400 hover:bg-gray-100 hover:text-gray-700 disabled:opacity-30 disabled:cursor-not-allowed"
                title="Последняя"
              >
                <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth={2.5}>
                  <path strokeLinecap="round" strokeLinejoin="round" d="M13 5l7 7-7 7M5 5l7 7-7 7" />
                </svg>
              </button>
            </div>
          </div>
        </div>
      </div>

      <div className="max-w-3xl mx-auto px-4 py-5 fade-in">
        {renderRow(currentRow, `Строка ${currentIndex + 1}`)}
      </div>

      <div className="max-w-3xl mx-auto px-4 pb-6">
        <div className="text-center">
          <span className="inline-flex items-center gap-1.5 text-xs text-gray-400 bg-gray-100/60 px-3 py-1.5 rounded-full flex-wrap justify-center">
            <span>Клик по ячейке — копировать</span>
            <span className="text-gray-300">·</span>
            <span>
              <kbd className="px-1 py-0.5 bg-white rounded text-[10px] font-mono border border-gray-200 text-gray-500">↑</kbd>
              <kbd className="px-1 py-0.5 bg-white rounded text-[10px] font-mono border border-gray-200 text-gray-500 ml-0.5">↓</kbd>
              <span className="ml-1">или колесо мыши</span>
            </span>
          </span>
        </div>
      </div>
    </div>
  );
}

export default App;
