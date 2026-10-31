import React, { useMemo, useRef, useState } from 'react';
import { ArrowDown, ArrowUp, ArrowUpDown, CheckCircle2, XCircle } from 'lucide-react';
import type { DbRowItem, DbSortOption } from '../lib/excelParser';
import { Badge } from './ui/badge';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from './ui/table';

export interface VirtualWindow {
  startIndex: number;
  endIndex: number;
  topSpacerHeight: number;
  bottomSpacerHeight: number;
}

export function computeVirtualWindow(
  scrollTop: number,
  viewportHeight: number,
  rowHeight: number,
  totalRows: number,
  overscan = 15,
): VirtualWindow {
  if (totalRows <= 0 || rowHeight <= 0) {
    return {
      startIndex: 0,
      endIndex: 0,
      topSpacerHeight: 0,
      bottomSpacerHeight: 0,
    };
  }

  const safeViewport = Math.max(rowHeight, viewportHeight);
  const maxScrollTop = Math.max(0, totalRows * rowHeight - safeViewport);
  const safeScrollTop = Math.max(0, Math.min(scrollTop, maxScrollTop));
  const visibleCount = Math.ceil(safeViewport / rowHeight);
  const rawStart = Math.floor(safeScrollTop / rowHeight);

  const startIndex = Math.max(0, Math.min(totalRows, rawStart - overscan));
  const endIndex = Math.max(
    startIndex,
    Math.min(totalRows, rawStart + visibleCount + overscan),
  );

  const topSpacerHeight = startIndex * rowHeight;
  const bottomSpacerHeight = Math.max(0, (totalRows - endIndex) * rowHeight);

  return {
    startIndex,
    endIndex,
    topSpacerHeight,
    bottomSpacerHeight,
  };
}

export function findJumpTargetIndex(
  rows: ReadonlyArray<{
    rowNumber: number;
    kind: 'section' | 'record';
    sectionCode?: string;
    values: Record<string, string | undefined>;
  }>,
  rawTarget: string,
): number {
  const cleaned = rawTarget.trim().replace(/^#/, '');
  if (!cleaned) {
    return -1;
  }

  if (/^\d+$/.test(cleaned)) {
    const targetRowNum = Number(cleaned);
    const exactIdx = rows.findIndex((r) => r.rowNumber === targetRowNum);
    if (exactIdx !== -1) {
      return exactIdx;
    }
    return rows.findIndex((r) => r.rowNumber >= targetRowNum);
  }

  const lower = cleaned.toLowerCase();
  return rows.findIndex(
    (r) =>
      r.sectionCode?.toLowerCase() === lower ||
      r.values['組別']?.toLowerCase() === lower ||
      r.sectionCode?.toLowerCase().includes(lower) ||
      r.values['組別']?.toLowerCase().includes(lower),
  );
}

export interface VirtualizedDbTableProps {
  rows: DbRowItem[];
  allRows: DbRowItem[];
  visibleHeaders: string[];
  selectedRowNumbers: ReadonlySet<number>;
  sort: DbSortOption;
  onToggleSort: (column: string) => void;
  onToggleRecordRow: (rowNumber: number) => void;
  onToggleSectionRows: (sectionRowNumber: number, targetRowNumbers: number[]) => void;
  onToggleAllVisibleRecords: () => void;
}

const ROW_HEIGHT = 40;
const VIEWPORT_HEIGHT = 600;

export function VirtualizedDbTable({
  rows,
  allRows,
  visibleHeaders,
  selectedRowNumbers,
  sort,
  onToggleSort,
  onToggleRecordRow,
  onToggleSectionRows,
  onToggleAllVisibleRecords,
}: VirtualizedDbTableProps): React.ReactElement {
  const containerRef = useRef<HTMLDivElement | null>(null);
  const [scrollTop, setScrollTop] = useState(0);
  const [jumpInput, setJumpInput] = useState('');

  const visibleRecordRows = useMemo(
    () => rows.filter((r) => r.kind === 'record'),
    [rows],
  );

  const allVisibleSelected =
    visibleRecordRows.length > 0 &&
    visibleRecordRows.every((r) => selectedRowNumbers.has(r.rowNumber));

  const someVisibleSelected =
    !allVisibleSelected &&
    visibleRecordRows.some((r) => selectedRowNumbers.has(r.rowNumber));

  // Map section rowNumber -> visible descendant record rowNumbers (or all descendant records if none filtered)
  const sectionDescendantMap = useMemo(() => {
    const map = new Map<number, number[]>();
    const sourceRecords = visibleRecordRows.length > 0 ? visibleRecordRows : allRows.filter((r) => r.kind === 'record');
    for (const rec of sourceRecords) {
      for (const anc of rec.ancestorSectionRows ?? []) {
        const list = map.get(anc);
        if (list) {
          list.push(rec.rowNumber);
        } else {
          map.set(anc, [rec.rowNumber]);
        }
      }
    }
    return map;
  }, [visibleRecordRows, allRows]);

  const { startIndex, endIndex, topSpacerHeight, bottomSpacerHeight } = useMemo(
    () => computeVirtualWindow(scrollTop, VIEWPORT_HEIGHT, ROW_HEIGHT, rows.length, 15),
    [scrollTop, rows.length],
  );

  const visibleSlice = useMemo(
    () => rows.slice(startIndex, endIndex),
    [rows, startIndex, endIndex],
  );

  const handleScroll = (e: React.UIEvent<HTMLDivElement>): void => {
    setScrollTop(e.currentTarget.scrollTop);
  };

  const handleQuickJump = (e: React.FormEvent): void => {
    e.preventDefault();
    const idx = findJumpTargetIndex(rows, jumpInput);
    if (idx !== -1 && containerRef.current) {
      const nextTop = idx * ROW_HEIGHT;
      containerRef.current.scrollTop = nextTop;
      setScrollTop(nextTop);
    }
  };

  return (
    <div className="border-t border-slate-200">
      <div className="flex flex-wrap items-center justify-between gap-2 bg-slate-50/70 px-4 py-1.5 border-b border-slate-200 text-xs text-slate-600">
        <span>
          檢視範圍：第 {rows.length === 0 ? 0 : startIndex + 1} - {endIndex} 列（共 {rows.length.toLocaleString()} 列）
        </span>
        <form onSubmit={handleQuickJump} className="flex items-center gap-1.5">
          <input
            type="text"
            value={jumpInput}
            onChange={(e) => setJumpInput(e.target.value)}
            placeholder="跳轉列號 (#50) 或組別 (MZ001)..."
            className="h-6 w-48 rounded border border-slate-300 bg-white px-2 text-xs text-slate-800 focus:outline-none focus:ring-1 focus:ring-slate-400"
          />
          <button
            type="submit"
            className="h-6 rounded border border-slate-300 bg-white px-2 text-xs font-medium text-slate-700 hover:bg-slate-100 cursor-pointer"
          >
            跳轉
          </button>
        </form>
      </div>
      <div
        ref={containerRef}
        onScroll={handleScroll}
        className="max-h-[600px] overflow-auto relative"
      >
      <Table>
        <TableHeader className="sticky top-0 z-10 bg-slate-50 shadow-2xs">
          <TableRow>
            <TableHead className="w-10 text-center px-2">
              <input
                type="checkbox"
                checked={allVisibleSelected}
                ref={(el) => {
                  if (el) el.indeterminate = someVisibleSelected;
                }}
                onChange={onToggleAllVisibleRecords}
                disabled={visibleRecordRows.length === 0}
                title="全選 / 取消全選目前篩選結果"
                className="h-4 w-4 rounded border-slate-300 text-slate-900 focus:ring-slate-500 cursor-pointer"
              />
            </TableHead>
            <TableHead className="w-16 text-center">列號</TableHead>
            {visibleHeaders.map((header) => {
              const isSorted = sort.column === header;
              return (
                <TableHead
                  key={header}
                  onClick={() => onToggleSort(header)}
                  className="cursor-pointer select-none hover:bg-slate-200/70 transition-colors"
                  title={`點擊依「${header}」排序檢視`}
                >
                  <div className="flex items-center gap-1">
                    <span>{header}</span>
                    {isSorted ? (
                      sort.direction === 'asc' ? (
                        <ArrowUp className="h-3.5 w-3.5 text-blue-600 shrink-0" />
                      ) : (
                        <ArrowDown className="h-3.5 w-3.5 text-blue-600 shrink-0" />
                      )
                    ) : (
                      <ArrowUpDown className="h-3 w-3 text-slate-300 shrink-0" />
                    )}
                  </div>
                </TableHead>
              );
            })}
          </TableRow>
        </TableHeader>
        <TableBody>
          {rows.length === 0 ? (
            <TableRow>
              <TableCell
                colSpan={visibleHeaders.length + 2}
                className="h-24 text-center text-slate-400"
              >
                沒有符合篩選條件的資料列
              </TableCell>
            </TableRow>
          ) : (
            <>
              {topSpacerHeight > 0 && (
                <TableRow aria-hidden="true" className="border-0 hover:bg-transparent">
                  <TableCell
                    colSpan={visibleHeaders.length + 2}
                    style={{ height: `${topSpacerHeight}px`, padding: 0 }}
                  />
                </TableRow>
              )}

              {visibleSlice.map((row) => {
                if (row.kind === 'section') {
                  const isTopSector = row.sectionCode?.endsWith('000') ?? false;
                  const sectionRecords = sectionDescendantMap.get(row.rowNumber) ?? [];
                  const allSecSelected =
                    sectionRecords.length > 0 &&
                    sectionRecords.every((rNum) => selectedRowNumbers.has(rNum));
                  const someSecSelected =
                    !allSecSelected &&
                    sectionRecords.some((rNum) => selectedRowNumbers.has(rNum));

                  return (
                    <TableRow
                      key={`section-${row.rowNumber}`}
                      style={{ height: `${ROW_HEIGHT}px` }}
                      className={
                        isTopSector
                          ? 'bg-slate-200/85 hover:bg-slate-200 font-semibold'
                          : 'bg-slate-100/95 hover:bg-slate-200/60 font-medium'
                      }
                    >
                      <TableCell className="text-center px-2">
                        {sectionRecords.length > 0 && (
                          <input
                            type="checkbox"
                            checked={allSecSelected}
                            ref={(el) => {
                              if (el) el.indeterminate = someSecSelected;
                            }}
                            onChange={() => onToggleSectionRows(row.rowNumber, sectionRecords)}
                            title={`勾選 / 取消此分組下 ${sectionRecords.length} 筆名單`}
                            className="h-4 w-4 rounded border-slate-300 text-slate-900 cursor-pointer"
                          />
                        )}
                      </TableCell>
                      <TableCell className="text-center font-mono text-[11px] text-slate-500">
                        #{row.rowNumber}
                      </TableCell>
                      <TableCell colSpan={visibleHeaders.length} className="py-1.5">
                        <div className="flex flex-wrap items-center gap-2">
                          <Badge variant={isTopSector ? 'default' : 'outline'}>
                            {row.sectionCode}
                          </Badge>
                          <span className="text-slate-900">
                            {row.sectionTitle || row.values['姓名']}
                          </span>
                          {sectionRecords.length > 0 && (
                            <span className="text-[11px] text-slate-500 font-normal">
                              ({sectionRecords.length} 筆)
                            </span>
                          )}
                          {row.values['Remark'] && (
                            <span className="ml-2 rounded bg-amber-50 border border-amber-200 px-2 py-0.5 text-[11px] font-normal text-amber-800 truncate max-w-xl" title={row.values['Remark']}>
                              備註：{row.values['Remark']}
                            </span>
                          )}
                        </div>
                      </TableCell>
                    </TableRow>
                  );
                }

                const isSelected = selectedRowNumbers.has(row.rowNumber);
                return (
                  <TableRow
                    key={`row-${row.rowNumber}`}
                    style={{ height: `${ROW_HEIGHT}px` }}
                    className={isSelected ? 'bg-blue-50/70 hover:bg-blue-100/60' : undefined}
                  >
                    <TableCell className="text-center px-2">
                      <input
                        type="checkbox"
                        checked={isSelected}
                        onChange={() => onToggleRecordRow(row.rowNumber)}
                        className="h-4 w-4 rounded border-slate-300 text-slate-900 cursor-pointer"
                      />
                    </TableCell>
                    <TableCell className="text-center font-mono text-[11px] text-slate-400">
                      {row.rowNumber}
                    </TableCell>
                    {visibleHeaders.map((header) => {
                      const val = row.values[header] ?? '';

                      if (header === '出席活動') {
                        return (
                          <TableCell key={header}>
                            {val === 'Y' ? (
                              <Badge variant="success" className="gap-1">
                                <CheckCircle2 className="h-3 w-3" /> 出席 (Y)
                              </Badge>
                            ) : val === '不出席活動' ? (
                              <Badge variant="destructive" className="gap-1">
                                <XCircle className="h-3 w-3" /> 不出席活動
                              </Badge>
                            ) : (
                              val || <span className="text-slate-300">-</span>
                            )}
                          </TableCell>
                        );
                      }

                      if (header === '組別' && val) {
                        return (
                          <TableCell key={header}>
                            <Badge variant="outline" className="font-mono">
                              {val}
                            </Badge>
                          </TableCell>
                        );
                      }

                      if (header === '計算') {
                        return (
                          <TableCell key={header} className="text-center font-mono">
                            {val === '1' ? (
                              <Badge variant="secondary">1</Badge>
                            ) : (
                              <span className="text-slate-400">{val || '0'}</span>
                            )}
                          </TableCell>
                        );
                      }

                      if (header === '酒會' || header === '升旗') {
                        return (
                          <TableCell key={header} className="text-center font-mono">
                            {val ? (
                              <Badge variant="warning">{val}</Badge>
                            ) : (
                              <span className="text-slate-300">-</span>
                            )}
                          </TableCell>
                        );
                      }

                      return (
                        <TableCell key={header} className="max-w-xs truncate" title={val}>
                          {val || <span className="text-slate-300">-</span>}
                        </TableCell>
                      );
                    })}
                  </TableRow>
                );
              })}

              {bottomSpacerHeight > 0 && (
                <TableRow aria-hidden="true" className="border-0 hover:bg-transparent">
                  <TableCell
                    colSpan={visibleHeaders.length + 2}
                    style={{ height: `${bottomSpacerHeight}px`, padding: 0 }}
                  />
                </TableRow>
              )}
            </>
          )}
        </TableBody>
      </Table>
      </div>
    </div>
  );
}
