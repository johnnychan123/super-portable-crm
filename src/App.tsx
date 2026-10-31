import React, { useEffect, useMemo, useRef, useState } from 'react';
import pkg from '../package.json';
import {
  ArrowUpDown,
  Building2,
  CheckCircle2,
  CheckSquare,
  Download,
  FileSpreadsheet,
  Filter,
  FolderKanban,
  Layers,
  PanelLeftClose,
  PanelLeftOpen,
  Plus,
  RotateCcw,
  Search,
  SlidersHorizontal,
  Trash2,
  Upload,
  Users,
} from 'lucide-react';
import type { AppRuntimeInfo } from './vite-env';
import {
  type CustomFilterOperator,
  type CustomFilterRule,
  type DbColumnPreset,
  type DbSortOption,
  type ParsedCrmWorkbook,
  computeDbFacetCounts,
  filterDbRows,
  filterGroups,
  getVisibleDbHeaders,
  parseCrmWorkbook,
} from './lib/excelParser';
import {
  buildDefaultExportFileName,
  buildExportWorkbookBuffer,
  resolveExportDbRows,
} from './lib/excelExporter';
import {
  CUSTOM_MATCH_KEY_PREFIX,
  type ExternalFilterMatchConfig,
  type ExternalFilterViewMode,
  type GroupMatchOverride,
  type ImportedFilterFile,
  applyPrefixGroupSelection,
  computeExternalFilterRowSet,
  encodeCustomMatchKey,
  filterGroupsByPrefix,
  getAvailableMatchKeys,
  getGlobalMatchKeyOptions,
  getGroupOptionsForOverride,
  isCustomMatchKey,
  matchExternalFilterToDb,
  mergeExtraValuesFromFiles,
  parseCustomMatchKeyColumns,
  parseExternalFilterWorkbook,
  prioritizeMatchColumns,
  toggleColumnInCustomMatchKey,
} from './lib/externalFilterMatcher';
import { VirtualizedDbTable } from './components/VirtualizedDbTable';
import { Badge } from './components/ui/badge';
import { Button } from './components/ui/button';
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from './components/ui/card';
import { Input } from './components/ui/input';
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from './components/ui/table';
import { Tabs, TabsContent, TabsList, TabsTrigger } from './components/ui/tabs';

export const APP_VERSION = pkg.version;

export function formatPlatformLabel(platform: string, arch?: string): string {
  const names: Record<string, string> = {
    win32: 'Windows',
    darwin: 'macOS',
    linux: 'Linux',
  };
  const label = names[platform] ?? platform;
  return arch ? `${label} (${arch})` : label;
}

export function getGreetingMessage(appName = 'Super Portable CRM'): string {
  return `Hello World! Welcome to ${appName}`;
}

export function applyBatchSelection(
  current: ReadonlySet<number>,
  targetRowNumbers: number[],
  mode: 'select' | 'deselect' | 'invert' | 'clear',
): Set<number> {
  if (mode === 'clear') {
    return new Set<number>();
  }
  const next = new Set<number>(current);
  for (const rowNum of targetRowNumbers) {
    if (mode === 'select') {
      next.add(rowNum);
    } else if (mode === 'deselect') {
      next.delete(rowNum);
    } else if (mode === 'invert') {
      if (next.has(rowNum)) {
        next.delete(rowNum);
      } else {
        next.add(rowNum);
      }
    }
  }
  return next;
}

export function getNextSortState(
  current: DbSortOption,
  clickedColumn: string,
): DbSortOption {
  if (current.column !== clickedColumn) {
    return { column: clickedColumn, direction: 'asc' };
  }
  if (current.direction === 'asc') {
    return { column: clickedColumn, direction: 'desc' };
  }
  return { column: null, direction: 'asc' };
}

export type ExportColumnMode = 'db_all' | 'from_file' | 'custom';

export function reconcileExternalFilterState(
  fileIds: string[],
  current: {
    externalFilterMode: ExternalFilterViewMode;
    exportTemplateFileId: string;
    exportColumnMode: ExportColumnMode;
  },
): {
  externalFilterMode: ExternalFilterViewMode;
  exportTemplateFileId: string;
  exportColumnMode: ExportColumnMode;
} {
  if (fileIds.length === 0) {
    return {
      externalFilterMode: 'all',
      exportTemplateFileId: '',
      exportColumnMode:
        current.exportColumnMode === 'from_file' ? 'db_all' : current.exportColumnMode,
    };
  }

  const fileIdSet = new Set(fileIds);
  let nextFilterMode = current.externalFilterMode;
  if (nextFilterMode === 'intersection' && fileIds.length < 2) {
    nextFilterMode = 'all';
  } else if (
    nextFilterMode !== 'all' &&
    nextFilterMode !== 'union' &&
    nextFilterMode !== 'intersection' &&
    !fileIdSet.has(nextFilterMode)
  ) {
    nextFilterMode = 'all';
  }

  const nextTemplateId = fileIdSet.has(current.exportTemplateFileId)
    ? current.exportTemplateFileId
    : (fileIds[0] ?? '');

  return {
    externalFilterMode: nextFilterMode,
    exportTemplateFileId: nextTemplateId,
    exportColumnMode: current.exportColumnMode,
  };
}

export function resolveInitialCustomMatchKey(
  currentKey: string,
  availableColumns: ReadonlyArray<string>,
): string {
  const existingCols = parseCustomMatchKeyColumns(currentKey);
  if (existingCols.length > 0) {
    return encodeCustomMatchKey(existingCols);
  }
  const colSet = new Set(availableColumns);
  if (colSet.has('組別') && colSet.has('姓名')) {
    return encodeCustomMatchKey(['組別', '姓名']);
  }
  const prioritized = prioritizeMatchColumns(availableColumns);
  return encodeCustomMatchKey(prioritized.slice(0, 2));
}

export function toggleGroupCodeInList(
  currentCodes: ReadonlyArray<string>,
  targetCode: string,
): string[] {
  const targetUpper = targetCode.trim().toUpperCase();
  if (!targetUpper) return [...currentCodes];
  const exists = currentCodes.some((c) => c.trim().toUpperCase() === targetUpper);
  if (exists) {
    return currentCodes.filter((c) => c.trim().toUpperCase() !== targetUpper);
  }
  return [...currentCodes, targetCode.trim()];
}

export function recomputeFilesWithGlobalRules(
  files: ImportedFilterFile[],
  dbRows: ParsedCrmWorkbook['dbRows'],
  config: ExternalFilterMatchConfig,
): ImportedFilterFile[] {
  return files.map((file) => {
    if (file.selectedMatchKey !== 'global') {
      return file;
    }
    return {
      ...file,
      matchResult: matchExternalFilterToDb(file.sheet, dbRows, 'global', config),
    };
  });
}

export function resolveActiveOutputHeaders(
  dbHeaders: string[],
  mode: ExportColumnMode,
  templateFileHeaders: string[],
  customSelectedDbHeaders: string[],
  customExtraHeaders: string[],
): string[] {
  if (mode === 'from_file' && templateFileHeaders.length > 0) {
    return templateFileHeaders;
  }
  if (mode === 'custom') {
    const seen = new Set<string>();
    const result: string[] = [];
    for (const h of [...customSelectedDbHeaders, ...customExtraHeaders]) {
      const trimmed = h.trim();
      if (trimmed && !seen.has(trimmed)) {
        seen.add(trimmed);
        result.push(trimmed);
      }
    }
    return result.length > 0 ? result : dbHeaders;
  }
  return dbHeaders;
}

const OPERATOR_LABELS: Record<CustomFilterOperator, string> = {
  contains: '包含',
  not_contains: '不包含',
  equals: '等於',
  not_equals: '不等於',
  is_empty: '為空白',
  is_not_empty: '不為空白',
};

export default function App(): React.ReactElement {
  const [info, setInfo] = useState<AppRuntimeInfo | null>(null);
  const [workbook, setWorkbook] = useState<ParsedCrmWorkbook | null>(null);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [exportNotice, setExportNotice] = useState<string | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [isExporting, setIsExporting] = useState(false);

  // DB table filters & view states
  const [dbSearch, setDbSearch] = useState('');
  const [attendanceFilter, setAttendanceFilter] = useState<'all' | 'Y' | '不出席活動' | 'blank'>('all');
  const [receptionFilter, setReceptionFilter] = useState<'all' | '%' | 'blank'>('all');
  const [flagRaisingFilter, setFlagRaisingFilter] = useState<'all' | '*' | 'blank'>('all');
  const [calcFilter, setCalcFilter] = useState<'all' | '1' | '0'>('all');
  const [dbSectorFilter, setDbSectorFilter] = useState('all');
  const [dbCategoryFilter, setDbCategoryFilter] = useState('all');
  const [dbGroupFilter, setDbGroupFilter] = useState('all');
  const [selectionFilter, setSelectionFilter] = useState<'all' | 'selected' | 'unselected'>('all');
  const [showSectionHeaders, setShowSectionHeaders] = useState(true);
  const [columnPreset, setColumnPreset] = useState<DbColumnPreset>('core');
  const [sort, setSort] = useState<DbSortOption>({ column: null, direction: 'asc' });

  // Custom multi-column compound filter builder
  const [showCustomFilterPanel, setShowCustomFilterPanel] = useState(false);
  const [customRules, setCustomRules] = useState<CustomFilterRule[]>([]);
  const [customRulesMode, setCustomRulesMode] = useState<'AND' | 'OR'>('AND');

  // External filter Excel files (multi-import & match)
  const [showExternalFilterPanel, setShowExternalFilterPanel] = useState(false);
  const [importedFilterFiles, setImportedFilterFiles] = useState<ImportedFilterFile[]>([]);
  const [externalFilterMode, setExternalFilterMode] = useState<ExternalFilterViewMode>('all');
  const [defaultExternalMatchKey, setDefaultExternalMatchKey] = useState<string>('auto');
  const [groupMatchOverrides, setGroupMatchOverrides] = useState<GroupMatchOverride[]>([]);
  const [rulePrefixInputs, setRulePrefixInputs] = useState<Record<string, string>>({});

  // Export column configuration ("按需選擇匯出欄位")
  const [showExportColumnPanel, setShowExportColumnPanel] = useState(false);
  const [exportColumnMode, setExportColumnMode] = useState<ExportColumnMode>('db_all');
  const [exportTemplateFileId, setExportTemplateFileId] = useState<string>('');
  const [customExportDbHeaders, setCustomExportDbHeaders] = useState<string[]>([]);
  const [customExtraHeaders, setCustomExtraHeaders] = useState<string[]>([]);
  const [newExtraHeaderInput, setNewExtraHeaderInput] = useState('');
  const [populateExtraColumnValues, setPopulateExtraColumnValues] = useState(false);

  // Left navigation sidebar state
  const [showSidebar, setShowSidebar] = useState(true);
  const [sidebarTab, setSidebarTab] = useState<'sectors' | 'categories' | 'groups'>('sectors');
  const [sidebarGroupSearch, setSidebarGroupSearch] = useState('');

  // Selected record rowNumbers
  const [selectedRowNumbers, setSelectedRowNumbers] = useState<Set<number>>(new Set());

  // 組別 tab filters
  const [groupSearch, setGroupSearch] = useState('');
  const [sectorFilter, setSectorFilter] = useState('all');
  const [categoryFilter, setCategoryFilter] = useState('all');

  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const externalFilterInputRef = useRef<HTMLInputElement | null>(null);

  useEffect(() => {
    if (window.electronAPI?.getAppInfo) {
      window.electronAPI
        .getAppInfo()
        .then(setInfo)
        .catch(() => {
          // Fallback to static preload values if IPC is unavailable
        });
    }
  }, []);

  const platform = info?.platform ?? window.electronAPI?.platform ?? 'web';
  const arch = info?.arch ?? window.electronAPI?.arch ?? '';

  const resetAllDbFilters = (): void => {
    setDbSearch('');
    setAttendanceFilter('all');
    setReceptionFilter('all');
    setFlagRaisingFilter('all');
    setCalcFilter('all');
    setDbSectorFilter('all');
    setDbCategoryFilter('all');
    setDbGroupFilter('all');
    setSelectionFilter('all');
    setExternalFilterMode('all');
    setCustomRules([]);
    setSort({ column: null, direction: 'asc' });
  };

  const handleLoadBuffer = (
    buffer: Uint8Array | ArrayBuffer,
    fileName: string,
    filePath?: string,
  ): void => {
    try {
      const parsed = parseCrmWorkbook(buffer, fileName, filePath);
      setWorkbook(parsed);
      setErrorMessage(null);
      setExportNotice(null);
      resetAllDbFilters();
      setSelectedRowNumbers(new Set());
      setImportedFilterFiles([]);
      setDefaultExternalMatchKey('auto');
      setGroupMatchOverrides([]);
      setRulePrefixInputs({});
      setExportColumnMode('db_all');
      setExportTemplateFileId('');
      setCustomExportDbHeaders(parsed.dbHeaders);
      setCustomExtraHeaders([]);
      setPopulateExtraColumnValues(false);
      setGroupSearch('');
      setSectorFilter('all');
      setCategoryFilter('all');
    } catch (err) {
      setErrorMessage(err instanceof Error ? err.message : '無法解析此 Excel 檔案');
    }
  };

  const handleAddExternalFilterBuffers = (
    items: Array<{ fileName: string; buffer: Uint8Array | ArrayBuffer }>,
  ): void => {
    if (!workbook || items.length === 0) return;
    try {
      const matchConfig: ExternalFilterMatchConfig = {
        defaultMatchKey: defaultExternalMatchKey,
        groupOverrides: groupMatchOverrides,
      };
      const added: ImportedFilterFile[] = items.map((item, idx) => {
        const sheet = parseExternalFilterWorkbook(item.buffer, item.fileName, workbook.dbHeaders);
        const availableMatchKeys = getAvailableMatchKeys(sheet, workbook.dbRows);
        const matchResult = matchExternalFilterToDb(
          sheet,
          workbook.dbRows,
          'global',
          matchConfig,
        );
        return {
          id: `ext-${Date.now()}-${idx}-${Math.random().toString(36).slice(2, 6)}`,
          fileName: item.fileName,
          sheet,
          availableMatchKeys,
          selectedMatchKey: 'global',
          matchResult,
        };
      });

      const nextFiles = [...importedFilterFiles, ...added];
      const reconciled = reconcileExternalFilterState(
        nextFiles.map((f) => f.id),
        {
          externalFilterMode,
          exportTemplateFileId,
          exportColumnMode,
        },
      );
      setImportedFilterFiles(nextFiles);
      setExternalFilterMode(reconciled.externalFilterMode);
      setExportTemplateFileId(reconciled.exportTemplateFileId);
      setExportColumnMode(reconciled.exportColumnMode);
      setShowExternalFilterPanel(true);
      setErrorMessage(null);
    } catch (err) {
      setErrorMessage(err instanceof Error ? err.message : '無法解析篩選 Excel 檔案');
    }
  };

  const handleImportExternalFilterClick = async (): Promise<void> => {
    if (!workbook) return;
    setErrorMessage(null);
    if (window.electronAPI?.openMultipleExcelFiles) {
      setIsLoading(true);
      try {
        const res = await window.electronAPI.openMultipleExcelFiles();
        if (!res.canceled && res.files && res.files.length > 0) {
          handleAddExternalFilterBuffers(res.files);
        }
      } catch (err) {
        setErrorMessage(err instanceof Error ? err.message : '開啟篩選 Excel 檔案失敗');
      } finally {
        setIsLoading(false);
      }
      return;
    }

    if (window.electronAPI?.openExcelFile) {
      setIsLoading(true);
      try {
        const res = await window.electronAPI.openExcelFile();
        if (!res.canceled && res.buffer && res.fileName) {
          handleAddExternalFilterBuffers([{ fileName: res.fileName, buffer: res.buffer }]);
        }
      } catch (err) {
        setErrorMessage(err instanceof Error ? err.message : '開啟篩選 Excel 檔案失敗');
      } finally {
        setIsLoading(false);
      }
      return;
    }

    externalFilterInputRef.current?.click();
  };

  const handleExternalFilterInputChange = async (
    event: React.ChangeEvent<HTMLInputElement>,
  ): Promise<void> => {
    const fileList = event.target.files;
    if (!fileList || fileList.length === 0) return;
    setIsLoading(true);
    try {
      const items = await Promise.all(
        Array.from(fileList).map(async (file) => ({
          fileName: file.name,
          buffer: await file.arrayBuffer(),
        })),
      );
      handleAddExternalFilterBuffers(items);
    } catch (err) {
      setErrorMessage(err instanceof Error ? err.message : '讀取篩選 Excel 檔案失敗');
    } finally {
      setIsLoading(false);
      event.target.value = '';
    }
  };

  const handleChangeExternalMatchKey = (fileId: string, nextMatchKey: string): void => {
    if (!workbook) return;
    const matchConfig: ExternalFilterMatchConfig = {
      defaultMatchKey: defaultExternalMatchKey,
      groupOverrides: groupMatchOverrides,
    };
    setImportedFilterFiles((prev) =>
      prev.map((f) => {
        if (f.id !== fileId) return f;
        return {
          ...f,
          selectedMatchKey: nextMatchKey,
          matchResult: matchExternalFilterToDb(
            f.sheet,
            workbook.dbRows,
            nextMatchKey,
            matchConfig,
          ),
        };
      }),
    );
  };

  const handleChangeDefaultExternalMatchKey = (nextDefaultKey: string): void => {
    setDefaultExternalMatchKey(nextDefaultKey);
    if (!workbook) return;
    const nextConfig: ExternalFilterMatchConfig = {
      defaultMatchKey: nextDefaultKey,
      groupOverrides: groupMatchOverrides,
    };
    setImportedFilterFiles((prev) =>
      recomputeFilesWithGlobalRules(prev, workbook.dbRows, nextConfig),
    );
  };

  const handleAddGroupMatchOverride = (): void => {
    if (!workbook) return;
    const nextOverrides: GroupMatchOverride[] = [
      ...groupMatchOverrides,
      {
        id: `grp-match-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
        groupCodes: [],
        matchKey: '組別+姓名',
      },
    ];
    setGroupMatchOverrides(nextOverrides);
    setImportedFilterFiles((prev) =>
      recomputeFilesWithGlobalRules(prev, workbook.dbRows, {
        defaultMatchKey: defaultExternalMatchKey,
        groupOverrides: nextOverrides,
      }),
    );
  };

  const handleUpdateGroupMatchOverride = (
    id: string,
    patch: Partial<Pick<GroupMatchOverride, 'groupCodes' | 'groupCode' | 'matchKey'>>,
  ): void => {
    if (!workbook) return;
    const nextOverrides = groupMatchOverrides.map((item) =>
      item.id === id ? { ...item, ...patch } : item,
    );
    setGroupMatchOverrides(nextOverrides);
    setImportedFilterFiles((prev) =>
      recomputeFilesWithGlobalRules(prev, workbook.dbRows, {
        defaultMatchKey: defaultExternalMatchKey,
        groupOverrides: nextOverrides,
      }),
    );
  };

  const handleToggleGroupInOverride = (ruleId: string, groupCode: string): void => {
    const rule = groupMatchOverrides.find((r) => r.id === ruleId);
    if (!rule) return;
    const currentCodes = [
      ...(rule.groupCodes ?? []),
      ...(rule.groupCode ? [rule.groupCode] : []),
    ];
    const nextCodes = toggleGroupCodeInList(currentCodes, groupCode);
    handleUpdateGroupMatchOverride(ruleId, {
      groupCodes: nextCodes,
      groupCode: undefined,
    });
  };

  const handlePrefixBatchSelectInOverride = (
    ruleId: string,
    prefix: string,
    mode: 'select' | 'deselect',
  ): void => {
    const rule = groupMatchOverrides.find((r) => r.id === ruleId);
    if (!rule) return;
    const currentCodes = [
      ...(rule.groupCodes ?? []),
      ...(rule.groupCode ? [rule.groupCode] : []),
    ];
    const nextCodes = applyPrefixGroupSelection(
      currentCodes,
      allGroupOverrideOptions,
      prefix,
      mode,
    );
    handleUpdateGroupMatchOverride(ruleId, {
      groupCodes: nextCodes,
      groupCode: undefined,
    });
  };

  const handleRemoveGroupMatchOverride = (id: string): void => {
    if (!workbook) return;
    const nextOverrides = groupMatchOverrides.filter((item) => item.id !== id);
    setGroupMatchOverrides(nextOverrides);
    setRulePrefixInputs((prev) => {
      const next = { ...prev };
      delete next[id];
      return next;
    });
    setImportedFilterFiles((prev) =>
      recomputeFilesWithGlobalRules(prev, workbook.dbRows, {
        defaultMatchKey: defaultExternalMatchKey,
        groupOverrides: nextOverrides,
      }),
    );
  };

  const renderCustomCompositeColumnPicker = (
    currentMatchKey: string,
    availableColumns: ReadonlyArray<string>,
    onChangeKey: (nextKey: string) => void,
  ): React.JSX.Element | null => {
    if (!isCustomMatchKey(currentMatchKey)) return null;
    const selectedCols = parseCustomMatchKeyColumns(currentMatchKey);
    const selectedSet = new Set(selectedCols);
    const orderedCols = prioritizeMatchColumns(availableColumns);

    return (
      <div className="rounded-md border border-indigo-200 bg-indigo-50/40 p-2.5 space-y-2 text-xs">
        <div className="flex flex-wrap items-center justify-between gap-2">
          <div className="flex flex-wrap items-center gap-1.5">
            <span className="font-semibold text-indigo-950">
              自訂複合欄組合（已選 {selectedCols.length} 欄）：
            </span>
            {selectedCols.length === 0 ? (
              <span className="text-rose-600 font-medium">
                請點擊下方欄位勾選至少 1 個比對欄位
              </span>
            ) : (
              selectedCols.map((col, idx) => (
                <React.Fragment key={col}>
                  {idx > 0 && <span className="text-indigo-500 font-bold">+</span>}
                  <Badge
                    variant="secondary"
                    className="gap-1 bg-indigo-600 text-white hover:bg-indigo-700 text-[11px]"
                  >
                    <span>{col}</span>
                    <button
                      type="button"
                      onClick={() =>
                        onChangeKey(toggleColumnInCustomMatchKey(currentMatchKey, col))
                      }
                      className="ml-0.5 text-indigo-100 hover:text-white"
                      title={`移除「${col}」`}
                    >
                      ×
                    </button>
                  </Badge>
                </React.Fragment>
              ))
            )}
          </div>
          {selectedCols.length > 0 && (
            <Button
              variant="ghost"
              size="sm"
              className="h-6 px-2 text-[11px] text-slate-500 hover:text-rose-600"
              onClick={() => onChangeKey(CUSTOM_MATCH_KEY_PREFIX)}
            >
              清空已選欄位
            </Button>
          )}
        </div>

        <div className="flex flex-wrap items-center gap-1 max-h-28 overflow-y-auto rounded border border-indigo-100 bg-white p-1.5">
          {orderedCols.map((col) => {
            const checked = selectedSet.has(col);
            return (
              <button
                key={col}
                type="button"
                onClick={() =>
                  onChangeKey(toggleColumnInCustomMatchKey(currentMatchKey, col))
                }
                className={`inline-flex items-center gap-1 rounded px-2 py-0.5 text-[11px] border transition-colors ${
                  checked
                    ? 'border-indigo-500 bg-indigo-50 text-indigo-950 font-semibold'
                    : 'border-slate-200 bg-slate-50/60 text-slate-600 hover:bg-slate-100'
                }`}
              >
                <input
                  type="checkbox"
                  checked={checked}
                  readOnly
                  className="h-3 w-3 rounded border-slate-300 pointer-events-none"
                />
                <span>{col}</span>
              </button>
            );
          })}
        </div>
      </div>
    );
  };

  const handleRemoveExternalFilterFile = (fileId: string): void => {
    const nextFiles = importedFilterFiles.filter((f) => f.id !== fileId);
    const reconciled = reconcileExternalFilterState(
      nextFiles.map((f) => f.id),
      {
        externalFilterMode,
        exportTemplateFileId,
        exportColumnMode,
      },
    );
    setImportedFilterFiles(nextFiles);
    setExternalFilterMode(reconciled.externalFilterMode);
    setExportTemplateFileId(reconciled.exportTemplateFileId);
    setExportColumnMode(reconciled.exportColumnMode);
  };

  const handleApplyFileColumnsToExport = (file: ImportedFilterFile): void => {
    setExportTemplateFileId(file.id);
    setExportColumnMode('from_file');
    setCustomExportDbHeaders(file.sheet.sharedHeaders);
    setCustomExtraHeaders(file.sheet.extraHeaders);
    setShowExportColumnPanel(true);
  };

  const handleToggleCustomExportDbHeader = (header: string): void => {
    if (!workbook) return;
    setCustomExportDbHeaders((prev) => {
      const set = new Set(prev);
      if (set.has(header)) {
        set.delete(header);
      } else {
        set.add(header);
      }
      // Keep original DB header order
      return workbook.dbHeaders.filter((h) => set.has(h));
    });
  };

  const handleAddCustomExtraHeader = (): void => {
    const trimmed = newExtraHeaderInput.trim();
    if (!trimmed) return;
    setCustomExtraHeaders((prev) => (prev.includes(trimmed) ? prev : [...prev, trimmed]));
    setNewExtraHeaderInput('');
  };

  const handleRemoveCustomExtraHeader = (header: string): void => {
    setCustomExtraHeaders((prev) => prev.filter((h) => h !== header));
  };

  const handleImportClick = async (): Promise<void> => {
    setErrorMessage(null);
    if (window.electronAPI?.openExcelFile) {
      setIsLoading(true);
      try {
        const res = await window.electronAPI.openExcelFile();
        if (!res.canceled && res.buffer && res.fileName) {
          handleLoadBuffer(res.buffer, res.fileName, res.filePath);
        }
      } catch (err) {
        setErrorMessage(err instanceof Error ? err.message : '開啟檔案失敗');
      } finally {
        setIsLoading(false);
      }
      return;
    }

    fileInputRef.current?.click();
  };

  const handleFileInputChange = async (
    event: React.ChangeEvent<HTMLInputElement>,
  ): Promise<void> => {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }
    setIsLoading(true);
    try {
      const arrayBuffer = await file.arrayBuffer();
      handleLoadBuffer(arrayBuffer, file.name);
    } catch (err) {
      setErrorMessage(err instanceof Error ? err.message : '讀取檔案失敗');
    } finally {
      setIsLoading(false);
      event.target.value = '';
    }
  };

  const visibleDbHeaders = useMemo(
    () => (workbook ? getVisibleDbHeaders(workbook.dbHeaders, columnPreset) : []),
    [workbook, columnPreset],
  );

  const externalMatchedRowNumbers = useMemo(
    () => computeExternalFilterRowSet(importedFilterFiles, externalFilterMode),
    [importedFilterFiles, externalFilterMode],
  );

  const externalUnionRowNumbers = useMemo(
    () => computeExternalFilterRowSet(importedFilterFiles, 'union') ?? new Set<number>(),
    [importedFilterFiles],
  );

  const mergedExtraValuesByRowNumber = useMemo(
    () => mergeExtraValuesFromFiles(importedFilterFiles),
    [importedFilterFiles],
  );

  const allGroupOverrideOptions = useMemo(
    () =>
      workbook
        ? getGroupOptionsForOverride(workbook.groups, workbook.dbRows, importedFilterFiles)
        : [],
    [workbook, importedFilterFiles],
  );

  const activeOutputHeaders = useMemo(() => {
    if (!workbook) return [];
    const templateFile =
      importedFilterFiles.find((f) => f.id === exportTemplateFileId) ??
      importedFilterFiles[0];
    return resolveActiveOutputHeaders(
      workbook.dbHeaders,
      exportColumnMode,
      templateFile?.sheet.headers ?? [],
      customExportDbHeaders,
      customExtraHeaders,
    );
  }, [
    workbook,
    importedFilterFiles,
    exportTemplateFileId,
    exportColumnMode,
    customExportDbHeaders,
    customExtraHeaders,
  ]);

  const activeExtraHeaders = useMemo(() => {
    if (!workbook) return [];
    const dbSet = new Set(workbook.dbHeaders);
    return activeOutputHeaders.filter((h) => !dbSet.has(h));
  }, [workbook, activeOutputHeaders]);

  const filteredDbRows = useMemo(
    () =>
      workbook
        ? filterDbRows(workbook.dbRows, {
            searchQuery: dbSearch,
            attendanceFilter,
            receptionFilter,
            flagRaisingFilter,
            calcFilter,
            sectorFilter: dbSectorFilter,
            categoryFilter: dbCategoryFilter,
            groupFilter: dbGroupFilter,
            selectionFilter,
            selectedRowNumbers,
            externalMatchedRowNumbers,
            customRules,
            customRulesMode,
            showSectionHeaders,
            sort,
          })
        : [],
    [
      workbook,
      dbSearch,
      attendanceFilter,
      receptionFilter,
      flagRaisingFilter,
      calcFilter,
      dbSectorFilter,
      dbCategoryFilter,
      dbGroupFilter,
      selectionFilter,
      selectedRowNumbers,
      externalMatchedRowNumbers,
      customRules,
      customRulesMode,
      showSectionHeaders,
      sort,
    ],
  );

  const filteredRecordRows = useMemo(
    () => filteredDbRows.filter((r) => r.kind === 'record'),
    [filteredDbRows],
  );

  const totalFacetCounts = useMemo(
    () => (workbook ? computeDbFacetCounts(workbook.dbRows) : null),
    [workbook],
  );

  const filteredFacetCounts = useMemo(
    () => computeDbFacetCounts(filteredRecordRows),
    [filteredRecordRows],
  );

  // Precompute record rowNumbers by sector, category, and group for 1-click sidebar batch selection
  const recordRowsByFacet = useMemo(() => {
    const bySector = new Map<string, number[]>();
    const byCategory = new Map<string, number[]>();
    const byGroup = new Map<string, number[]>();
    if (!workbook) return { bySector, byCategory, byGroup };

    for (const row of workbook.dbRows) {
      if (row.kind !== 'record') continue;
      if (row.sectorCode) {
        const list = bySector.get(row.sectorCode) ?? [];
        list.push(row.rowNumber);
        bySector.set(row.sectorCode, list);
      }
      if (row.category) {
        const list = byCategory.get(row.category) ?? [];
        list.push(row.rowNumber);
        byCategory.set(row.category, list);
      }
      const gCode = row.values['組別'] ?? '';
      if (gCode) {
        const list = byGroup.get(gCode) ?? [];
        list.push(row.rowNumber);
        byGroup.set(gCode, list);
      }
    }
    return { bySector, byCategory, byGroup };
  }, [workbook]);

  const filteredExportSummary = useMemo(() => {
    if (!workbook) return { rowsToExport: [], recordCount: 0, sectionCount: 0 };
    const targetSet = new Set(filteredRecordRows.map((r) => r.rowNumber));
    return resolveExportDbRows(workbook, { targetRecordRowNumbers: targetSet });
  }, [workbook, filteredRecordRows]);

  const selectedExportSummary = useMemo(() => {
    if (!workbook) return { rowsToExport: [], recordCount: 0, sectionCount: 0 };
    return resolveExportDbRows(workbook, { targetRecordRowNumbers: selectedRowNumbers });
  }, [workbook, selectedRowNumbers]);

  const handleExportExcel = async (mode: 'filtered' | 'selected'): Promise<void> => {
    if (!workbook) return;
    const targetSet =
      mode === 'selected'
        ? selectedRowNumbers
        : new Set(filteredRecordRows.map((r) => r.rowNumber));

    if (targetSet.size === 0) {
      setErrorMessage(
        mode === 'selected'
          ? '目前尚未勾選任何名單，無法匯出'
          : '目前篩選結果為 0 筆名單，無法匯出',
      );
      return;
    }

    setIsExporting(true);
    setErrorMessage(null);
    setExportNotice(null);

    try {
      const summary = mode === 'selected' ? selectedExportSummary : filteredExportSummary;
      const suffix = mode === 'selected' ? '已勾選匯出' : '篩選匯出';
      const defaultFileName = buildDefaultExportFileName(workbook.fileName, suffix);
      const buffer = buildExportWorkbookBuffer(workbook, {
        targetRecordRowNumbers: targetSet,
        outputHeaders: activeOutputHeaders,
        populateExtraColumnValues,
        extraValuesByRowNumber: mergedExtraValuesByRowNumber,
      });
      const colDesc =
        exportColumnMode === 'db_all'
          ? `完整 ${activeOutputHeaders.length} 欄`
          : `自訂/沿用 ${activeOutputHeaders.length} 欄`;

      if (window.electronAPI?.saveExcelFile) {
        const res = await window.electronAPI.saveExcelFile({
          defaultFileName,
          buffer,
        });
        if (!res.canceled && res.filePath) {
          setExportNotice(
            `已成功匯出 ${summary.recordCount} 筆名單（含 ${summary.sectionCount} 列上層分組標題列、${colDesc}及組別/界別工作表）至：${res.filePath}`,
          );
        }
        return;
      }

      // Browser fallback download
      const blob = new Blob([buffer as unknown as BlobPart], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = defaultFileName;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
      setExportNotice(
        `已下載匯出檔 ${defaultFileName}（共 ${summary.recordCount} 筆名單、保留 ${summary.sectionCount} 列分組標題列、${colDesc}）`,
      );
    } catch (err) {
      setErrorMessage(err instanceof Error ? err.message : '匯出 Excel 失敗');
    } finally {
      setIsExporting(false);
    }
  };

  const handleToggleRecordRow = (rowNumber: number): void => {
    setSelectedRowNumbers((prev) => applyBatchSelection(prev, [rowNumber], 'invert'));
  };

  const handleToggleGroupOfRows = (targetRowNumbers: number[]): void => {
    if (targetRowNumbers.length === 0) return;
    setSelectedRowNumbers((prev) => {
      const allSelected = targetRowNumbers.every((r) => prev.has(r));
      return applyBatchSelection(prev, targetRowNumbers, allSelected ? 'deselect' : 'select');
    });
  };

  const handleToggleAllVisibleRecords = (): void => {
    const target = filteredRecordRows.map((r) => r.rowNumber);
    handleToggleGroupOfRows(target);
  };

  const handleAddCustomRule = (): void => {
    if (!workbook) return;
    const defaultCol = workbook.dbHeaders.includes('姓名')
      ? '姓名'
      : (workbook.dbHeaders[0] ?? '組別');
    setCustomRules((prev) => [
      ...prev,
      {
        id: `rule-${Date.now()}-${prev.length}`,
        column: defaultCol,
        operator: 'contains',
        value: '',
      },
    ]);
    setShowCustomFilterPanel(true);
  };

  const handleUpdateCustomRule = (
    id: string,
    patch: Partial<CustomFilterRule>,
  ): void => {
    setCustomRules((prev) =>
      prev.map((r) => (r.id === id ? { ...r, ...patch } : r)),
    );
  };

  const handleRemoveCustomRule = (id: string): void => {
    setCustomRules((prev) => prev.filter((r) => r.id !== id));
  };

  const filteredGroups = useMemo(
    () =>
      workbook
        ? filterGroups(workbook.groups, {
            searchQuery: groupSearch,
            sectorFilter,
            categoryFilter,
          })
        : [],
    [workbook, groupSearch, sectorFilter, categoryFilter],
  );

  const sidebarGroups = useMemo(() => {
    if (!workbook) return [];
    const q = sidebarGroupSearch.trim().toLowerCase();
    return workbook.groups.filter((g) => {
      if (g.category === '--') return false;
      if (dbSectorFilter !== 'all' && g.sectorCode !== dbSectorFilter) return false;
      if (dbCategoryFilter !== 'all' && g.category !== dbCategoryFilter) return false;
      if (!q) return true;
      return (
        g.code.toLowerCase().includes(q) ||
        g.nameZh.toLowerCase().includes(q) ||
        g.category.toLowerCase().includes(q)
      );
    });
  }, [workbook, sidebarGroupSearch, dbSectorFilter, dbCategoryFilter]);

  const uniqueCategories = useMemo(() => {
    if (!workbook) return [];
    return Array.from(
      new Set(workbook.groups.map((g) => g.category).filter((c) => c && c !== '--')),
    );
  }, [workbook]);

  const groupCountBySector = useMemo(() => {
    const counts = new Map<string, number>();
    if (!workbook) return counts;
    for (const g of workbook.groups) {
      counts.set(g.sectorCode, (counts.get(g.sectorCode) ?? 0) + 1);
    }
    return counts;
  }, [workbook]);

  const activeFilterCount = useMemo(() => {
    let count = 0;
    if (dbSearch.trim() !== '') count++;
    if (attendanceFilter !== 'all') count++;
    if (receptionFilter !== 'all') count++;
    if (flagRaisingFilter !== 'all') count++;
    if (calcFilter !== 'all') count++;
    if (dbSectorFilter !== 'all') count++;
    if (dbCategoryFilter !== 'all') count++;
    if (dbGroupFilter !== 'all') count++;
    if (selectionFilter !== 'all') count++;
    if (externalFilterMode !== 'all') count++;
    count += customRules.filter(
      (r) =>
        r.operator === 'is_empty' ||
        r.operator === 'is_not_empty' ||
        r.value.trim() !== '',
    ).length;
    return count;
  }, [
    dbSearch,
    attendanceFilter,
    receptionFilter,
    flagRaisingFilter,
    calcFilter,
    dbSectorFilter,
    dbCategoryFilter,
    dbGroupFilter,
    selectionFilter,
    externalFilterMode,
    customRules,
  ]);

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 flex flex-col">
      {/* Hidden file input fallback for browser/dev mode */}
      <input
        ref={fileInputRef}
        type="file"
        accept=".xlsx,.xls"
        onChange={handleFileInputChange}
        className="hidden"
        data-testid="excel-file-input"
      />
      <input
        ref={externalFilterInputRef}
        type="file"
        accept=".xlsx,.xls"
        multiple
        onChange={handleExternalFilterInputChange}
        className="hidden"
        data-testid="external-filter-file-input"
      />

      {/* Top Navigation Bar */}
      <header className="sticky top-0 z-20 border-b border-slate-200 bg-white/95 backdrop-blur px-6 py-3">
        <div className="mx-auto flex max-w-[1720px] flex-wrap items-center justify-between gap-3">
          <div className="flex items-center gap-3">
            <div className="flex h-9 w-9 items-center justify-center rounded-lg bg-slate-900 text-white shadow-xs">
              <FileSpreadsheet className="h-5 w-5" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h1 className="text-base font-bold tracking-tight text-slate-900">
                  Super Portable CRM
                </h1>
                <Badge variant="secondary" className="text-[11px]">
                  v{info?.version ?? APP_VERSION} · {formatPlatformLabel(platform, arch)}
                </Badge>
              </div>
              <p className="text-xs text-slate-500">
                大型活動名單多維度篩選、批次選取與原結構 Excel 匯出系統
              </p>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-2.5">
            {workbook && (
              <>
                <div className="hidden xl:flex items-center gap-2 rounded-lg border border-slate-200 bg-slate-50 px-3 py-1.5 text-xs">
                  <FileSpreadsheet className="h-3.5 w-3.5 text-emerald-600" />
                  <span className="font-medium text-slate-800">{workbook.fileName}</span>
                  <span className="text-slate-300">|</span>
                  <span className="text-slate-500">
                    總名單 {workbook.summary.totalRecords} 筆 · 篩選命中{' '}
                    <strong className="text-slate-900">{filteredRecordRows.length}</strong> 筆 · 已勾選{' '}
                    <strong className="text-blue-700">{selectedRowNumbers.size}</strong> 筆
                  </span>
                </div>

                <Button
                  variant="default"
                  size="default"
                  disabled={isExporting || filteredRecordRows.length === 0}
                  onClick={() => handleExportExcel('filtered')}
                  title={`將目前篩選命中的 ${filteredExportSummary.recordCount} 筆名單與 ${filteredExportSummary.sectionCount} 列分組標題匯出為原結構 Excel`}
                >
                  <Download className="h-4 w-4" />
                  匯出篩選結果 ({filteredExportSummary.recordCount} 筆)
                </Button>

                <Button
                  variant="secondary"
                  size="default"
                  disabled={isExporting || selectedRowNumbers.size === 0}
                  onClick={() => handleExportExcel('selected')}
                  title={`僅匯出目前已勾選的 ${selectedExportSummary.recordCount} 筆名單與 ${selectedExportSummary.sectionCount} 列分組標題`}
                >
                  <CheckSquare className="h-4 w-4 text-blue-700" />
                  匯出已勾選 ({selectedExportSummary.recordCount} 筆)
                </Button>

                <Button
                  variant="outline"
                  size="default"
                  disabled={isLoading}
                  onClick={handleImportExternalFilterClick}
                  title="匯入過去匯出的 Excel 作為本次篩選與選取條件（可多次匯入）"
                >
                  <Filter className="h-4 w-4 text-indigo-600" />
                  匯入篩選 Excel
                  {importedFilterFiles.length > 0 ? ` (${importedFilterFiles.length})` : ''}
                </Button>
              </>
            )}

            <Button
              variant={workbook ? 'outline' : 'default'}
              onClick={handleImportClick}
              disabled={isLoading}
              size="default"
            >
              <Upload className="h-4 w-4" />
              {isLoading ? '正在讀取...' : workbook ? '重新匯入 Excel' : '匯入 Excel 檔案'}
            </Button>
          </div>
        </div>
      </header>

      {/* Main Content Area */}
      <main className="mx-auto w-full max-w-[1720px] flex-1 p-5 space-y-5">
        {errorMessage && (
          <div className="rounded-lg border border-rose-200 bg-rose-50 px-4 py-2.5 text-sm text-rose-700 flex items-center justify-between">
            <span>錯誤：{errorMessage}</span>
            <Button variant="ghost" size="sm" onClick={() => setErrorMessage(null)}>
              關閉
            </Button>
          </div>
        )}

        {exportNotice && (
          <div className="rounded-lg border border-emerald-200 bg-emerald-50 px-4 py-2.5 text-sm text-emerald-800 flex items-center justify-between">
            <span className="flex items-center gap-2">
              <CheckCircle2 className="h-4 w-4 text-emerald-600 shrink-0" />
              {exportNotice}
            </span>
            <Button variant="ghost" size="sm" onClick={() => setExportNotice(null)}>
              關閉
            </Button>
          </div>
        )}

        {!workbook ? (
          /* Empty State before importing */
          <div className="flex min-h-[70vh] items-center justify-center">
            <Card className="w-full max-w-xl border-dashed border-2 border-slate-300 bg-white/90 text-center py-6">
              <CardHeader className="items-center space-y-3">
                <div className="flex h-14 w-14 items-center justify-center rounded-2xl bg-slate-100 text-slate-700">
                  <FileSpreadsheet className="h-7 w-7" />
                </div>
                <CardTitle className="text-xl">匯入 Excel 資料庫以開始篩選與匯出</CardTitle>
                <CardDescription className="max-w-md text-sm leading-relaxed">
                  支援匯入包含 <code className="rounded bg-slate-100 px-1.5 py-0.5 font-mono text-slate-800">DB</code>、
                  <code className="rounded bg-slate-100 px-1.5 py-0.5 font-mono text-slate-800">組別</code>、
                  <code className="rounded bg-slate-100 px-1.5 py-0.5 font-mono text-slate-800">界別</code> 三層結構之萬筆名單 Excel 活頁簿（如 <code className="rounded bg-slate-100 px-1.5 py-0.5 font-mono text-slate-800">test.xlsx</code>）。
                </CardDescription>
              </CardHeader>
              <CardContent className="flex flex-col items-center gap-4">
                <Button onClick={handleImportClick} disabled={isLoading} size="lg">
                  <Upload className="h-4 w-4" />
                  選取並匯入 Excel 檔案 (.xlsx)
                </Button>
                <div className="grid grid-cols-3 gap-3 w-full max-w-md pt-2 text-left">
                  <div className="rounded-lg border border-slate-200 bg-slate-50/70 p-3">
                    <div className="text-xs font-semibold text-slate-800">萬筆高速虛擬表</div>
                    <div className="text-[11px] text-slate-500 mt-0.5">
                      50 欄位、多維度篩選、欄位檢視排序與批次勾選
                    </div>
                  </div>
                  <div className="rounded-lg border border-slate-200 bg-slate-50/70 p-3">
                    <div className="text-xs font-semibold text-slate-800">界別與組別導覽</div>
                    <div className="text-[11px] text-slate-500 mt-0.5">
                      26 個界別、14 種職務分類與 558 個組別樹狀過濾
                    </div>
                  </div>
                  <div className="rounded-lg border border-slate-200 bg-slate-50/70 p-3">
                    <div className="text-xs font-semibold text-slate-800">原結構 Excel 匯出</div>
                    <div className="text-[11px] text-slate-500 mt-0.5">
                      自動保留命中名單所屬分組標題列與完整三工作表
                    </div>
                  </div>
                </div>
              </CardContent>
            </Card>
          </div>
        ) : (
          /* Loaded Workbook Dashboard */
          <>
            {/* Summary KPI Cards */}
            <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 lg:grid-cols-4">
              <Card>
                <CardHeader className="flex flex-row items-center justify-between pb-2">
                  <CardTitle className="text-xs font-medium text-slate-500">
                    名單總筆數 / 計算人數
                  </CardTitle>
                  <Users className="h-4 w-4 text-slate-400" />
                </CardHeader>
                <CardContent>
                  <div className="flex items-baseline gap-2">
                    <span className="text-2xl font-bold">{workbook.summary.totalRecords}</span>
                    <span className="text-xs text-slate-500">
                      筆（去重計算：<strong className="text-slate-800">{workbook.summary.countedHeadcount}</strong> 人）
                    </span>
                  </div>
                  <p className="mt-1 text-xs text-slate-400">
                    目前篩選命中 {filteredRecordRows.length} 筆 · 已勾選 {selectedRowNumbers.size} 筆
                  </p>
                </CardContent>
              </Card>

              <Card>
                <CardHeader className="flex flex-row items-center justify-between pb-2">
                  <CardTitle className="text-xs font-medium text-slate-500">
                    出席活動狀態
                  </CardTitle>
                  <CheckCircle2 className="h-4 w-4 text-emerald-600" />
                </CardHeader>
                <CardContent>
                  <div className="flex items-center gap-3">
                    <div>
                      <span className="text-2xl font-bold text-emerald-700">
                        {workbook.summary.attendingCount}
                      </span>
                      <span className="ml-1 text-xs text-slate-500">出席 (Y)</span>
                    </div>
                    <span className="text-slate-300">/</span>
                    <div>
                      <span className="text-xl font-semibold text-rose-600">
                        {workbook.summary.notAttendingCount}
                      </span>
                      <span className="ml-1 text-xs text-slate-500">不出席</span>
                    </div>
                  </div>
                  <p className="mt-1 text-xs text-slate-400">
                    出席率{' '}
                    {workbook.summary.totalRecords > 0
                      ? Math.round(
                          (workbook.summary.attendingCount / workbook.summary.totalRecords) * 100,
                        )
                      : 0}
                    %
                  </p>
                </CardContent>
              </Card>

              <Card>
                <CardHeader className="flex flex-row items-center justify-between pb-2">
                  <CardTitle className="text-xs font-medium text-slate-500">
                    大型活動標記（酒會 / 升旗）
                  </CardTitle>
                  <Layers className="h-4 w-4 text-blue-600" />
                </CardHeader>
                <CardContent>
                  <div className="flex items-center gap-4">
                    <div>
                      <span className="text-2xl font-bold">{workbook.summary.receptionCount}</span>
                      <span className="ml-1 text-xs text-slate-500">酒會 (%)</span>
                    </div>
                    <span className="text-slate-300">|</span>
                    <div>
                      <span className="text-2xl font-bold">{workbook.summary.flagRaisingCount}</span>
                      <span className="ml-1 text-xs text-slate-500">升旗 (*)</span>
                    </div>
                  </div>
                  <p className="mt-1 text-xs text-slate-400">
                    依 DB 工作表 A、B 欄符號統計
                  </p>
                </CardContent>
              </Card>

              <Card>
                <CardHeader className="flex flex-row items-center justify-between pb-2">
                  <CardTitle className="text-xs font-medium text-slate-500">
                    預計匯出結構預覽
                  </CardTitle>
                  <Building2 className="h-4 w-4 text-slate-400" />
                </CardHeader>
                <CardContent>
                  <div className="flex items-center gap-2 text-xs">
                    <span className="font-semibold text-slate-800">
                      篩選匯出：{filteredExportSummary.recordCount} 筆名單
                    </span>
                    <span className="text-slate-400">
                      (+{filteredExportSummary.sectionCount} 標題列)
                    </span>
                  </div>
                  <p className="mt-1.5 text-xs text-slate-500">
                    勾選匯出：<strong className="text-blue-700">{selectedExportSummary.recordCount}</strong> 筆名單 (+
                    {selectedExportSummary.sectionCount} 標題列) · 保留 50 欄與組別/界別表
                  </p>
                </CardContent>
              </Card>
            </div>

            {/* Worksheet Tabs */}
            <Tabs defaultValue="DB" className="w-full">
              <div className="flex flex-wrap items-center justify-between gap-3">
                <TabsList>
                  <TabsTrigger value="DB">
                    DB 主名單 ({filteredRecordRows.length} / {workbook.summary.totalRecords})
                  </TabsTrigger>
                  <TabsTrigger value="groups">
                    組別 ({workbook.groups.length})
                  </TabsTrigger>
                  <TabsTrigger value="sectors">
                    界別 ({workbook.sectors.length})
                  </TabsTrigger>
                </TabsList>

                <div className="flex items-center gap-2">
                  {workbook.filePath && (
                    <span className="text-xs text-slate-400 truncate max-w-md" title={workbook.filePath}>
                      來源：{workbook.filePath}
                    </span>
                  )}
                </div>
              </div>

              {/* Tab 1: DB Sheet with Split-Pane Sidebar + Filter Builder + Virtualized Table */}
              <TabsContent value="DB">
                <div className="flex flex-col lg:flex-row gap-4 items-start">
                  {/* Left Collapsible Facet Sidebar */}
                  {showSidebar && (
                    <Card className="w-full lg:w-72 shrink-0">
                      <CardHeader className="pb-2.5 px-3.5 pt-3.5">
                        <div className="flex items-center justify-between">
                          <CardTitle className="text-sm font-semibold flex items-center gap-1.5">
                            <Building2 className="h-4 w-4 text-slate-600" />
                            階層篩選與整組勾選
                          </CardTitle>
                          <Button
                            variant="ghost"
                            size="sm"
                            className="h-7 px-2 text-xs text-slate-500"
                            onClick={() => setShowSidebar(false)}
                            title="收合側邊導覽面板"
                          >
                            <PanelLeftClose className="h-4 w-4" />
                          </Button>
                        </div>

                        {/* Sidebar mode switch */}
                        <div className="grid grid-cols-3 gap-1 rounded-lg bg-slate-100 p-1 mt-2">
                          <button
                            type="button"
                            onClick={() => setSidebarTab('sectors')}
                            className={`rounded-md py-1 text-xs font-medium transition-colors ${
                              sidebarTab === 'sectors'
                                ? 'bg-white text-slate-900 shadow-2xs'
                                : 'text-slate-600 hover:text-slate-900'
                            }`}
                          >
                            界別 ({workbook.sectors.length})
                          </button>
                          <button
                            type="button"
                            onClick={() => setSidebarTab('categories')}
                            className={`rounded-md py-1 text-xs font-medium transition-colors ${
                              sidebarTab === 'categories'
                                ? 'bg-white text-slate-900 shadow-2xs'
                                : 'text-slate-600 hover:text-slate-900'
                            }`}
                          >
                            分類 ({uniqueCategories.length})
                          </button>
                          <button
                            type="button"
                            onClick={() => setSidebarTab('groups')}
                            className={`rounded-md py-1 text-xs font-medium transition-colors ${
                              sidebarTab === 'groups'
                                ? 'bg-white text-slate-900 shadow-2xs'
                                : 'text-slate-600 hover:text-slate-900'
                            }`}
                          >
                            組別
                          </button>
                        </div>
                      </CardHeader>

                      <CardContent className="px-3 pb-3 pt-1">
                        {sidebarTab === 'sectors' && (
                          <div className="max-h-[560px] overflow-y-auto space-y-1 pr-1">
                            <button
                              type="button"
                              onClick={() => setDbSectorFilter('all')}
                              className={`w-full flex items-center justify-between rounded-md px-2.5 py-1.5 text-xs transition-colors ${
                                dbSectorFilter === 'all'
                                  ? 'bg-slate-900 text-white font-medium'
                                  : 'hover:bg-slate-100 text-slate-700'
                              }`}
                            >
                              <span>全部界別</span>
                              <span className="font-mono text-[11px] opacity-80">
                                {workbook.summary.totalRecords} 筆
                              </span>
                            </button>

                            {workbook.sectors.map((sec) => {
                              const totalInSec = totalFacetCounts?.bySector.get(sec.code) ?? 0;
                              const filteredInSec = filteredFacetCounts.bySector.get(sec.code) ?? 0;
                              const secRowNums = recordRowsByFacet.bySector.get(sec.code) ?? [];
                              const allSecChecked =
                                secRowNums.length > 0 &&
                                secRowNums.every((r) => selectedRowNumbers.has(r));
                              const someSecChecked =
                                !allSecChecked &&
                                secRowNums.some((r) => selectedRowNumbers.has(r));
                              const isActive = dbSectorFilter === sec.code;

                              return (
                                <div
                                  key={sec.code}
                                  className={`flex items-center gap-1.5 rounded-md px-2 py-1.5 text-xs transition-colors ${
                                    isActive
                                      ? 'bg-blue-50 border border-blue-200 text-blue-950 font-medium'
                                      : 'hover:bg-slate-100 text-slate-700'
                                  }`}
                                >
                                  <input
                                    type="checkbox"
                                    checked={allSecChecked}
                                    disabled={secRowNums.length === 0}
                                    ref={(el) => {
                                      if (el) el.indeterminate = someSecChecked;
                                    }}
                                    onChange={() => handleToggleGroupOfRows(secRowNums)}
                                    title={`勾選 / 取消 ${sec.code} 全界別 (${secRowNums.length} 筆)`}
                                    className="h-3.5 w-3.5 rounded border-slate-300 cursor-pointer shrink-0"
                                  />
                                  <button
                                    type="button"
                                    onClick={() =>
                                      setDbSectorFilter(isActive ? 'all' : sec.code)
                                    }
                                    className="flex-1 flex items-center justify-between text-left truncate gap-1"
                                  >
                                    <span className="truncate">
                                      <span className="font-mono font-semibold mr-1">{sec.code}</span>
                                      <span className="text-slate-600">{sec.nameZh || ''}</span>
                                    </span>
                                    <span className="font-mono text-[11px] text-slate-500 shrink-0">
                                      {filteredInSec}/{totalInSec}
                                    </span>
                                  </button>
                                </div>
                              );
                            })}
                          </div>
                        )}

                        {sidebarTab === 'categories' && (
                          <div className="max-h-[560px] overflow-y-auto space-y-1 pr-1">
                            <button
                              type="button"
                              onClick={() => setDbCategoryFilter('all')}
                              className={`w-full flex items-center justify-between rounded-md px-2.5 py-1.5 text-xs transition-colors ${
                                dbCategoryFilter === 'all'
                                  ? 'bg-slate-900 text-white font-medium'
                                  : 'hover:bg-slate-100 text-slate-700'
                              }`}
                            >
                              <span>全部職務分類</span>
                              <span className="font-mono text-[11px] opacity-80">
                                {workbook.summary.totalRecords} 筆
                              </span>
                            </button>

                            {uniqueCategories.map((cat) => {
                              const totalInCat = totalFacetCounts?.byCategory.get(cat) ?? 0;
                              const filteredInCat = filteredFacetCounts.byCategory.get(cat) ?? 0;
                              const catRowNums = recordRowsByFacet.byCategory.get(cat) ?? [];
                              const allCatChecked =
                                catRowNums.length > 0 &&
                                catRowNums.every((r) => selectedRowNumbers.has(r));
                              const someCatChecked =
                                !allCatChecked &&
                                catRowNums.some((r) => selectedRowNumbers.has(r));
                              const isActive = dbCategoryFilter === cat;

                              return (
                                <div
                                  key={cat}
                                  className={`flex items-center gap-1.5 rounded-md px-2 py-1.5 text-xs transition-colors ${
                                    isActive
                                      ? 'bg-blue-50 border border-blue-200 text-blue-950 font-medium'
                                      : 'hover:bg-slate-100 text-slate-700'
                                  }`}
                                >
                                  <input
                                    type="checkbox"
                                    checked={allCatChecked}
                                    disabled={catRowNums.length === 0}
                                    ref={(el) => {
                                      if (el) el.indeterminate = someCatChecked;
                                    }}
                                    onChange={() => handleToggleGroupOfRows(catRowNums)}
                                    title={`勾選 / 取消「${cat}」分類 (${catRowNums.length} 筆)`}
                                    className="h-3.5 w-3.5 rounded border-slate-300 cursor-pointer shrink-0"
                                  />
                                  <button
                                    type="button"
                                    onClick={() =>
                                      setDbCategoryFilter(isActive ? 'all' : cat)
                                    }
                                    className="flex-1 flex items-center justify-between text-left truncate gap-1"
                                  >
                                    <span className="truncate">{cat}</span>
                                    <span className="font-mono text-[11px] text-slate-500 shrink-0">
                                      {filteredInCat}/{totalInCat}
                                    </span>
                                  </button>
                                </div>
                              );
                            })}
                          </div>
                        )}

                        {sidebarTab === 'groups' && (
                          <div className="space-y-2">
                            <div className="relative">
                              <Search className="absolute left-2 top-2 h-3.5 w-3.5 text-slate-400" />
                              <Input
                                placeholder="搜尋組別代碼或名稱..."
                                value={sidebarGroupSearch}
                                onChange={(e) => setSidebarGroupSearch(e.target.value)}
                                className="h-7 pl-7 text-xs"
                              />
                            </div>
                            <div className="max-h-[520px] overflow-y-auto space-y-1 pr-1">
                              <button
                                type="button"
                                onClick={() => setDbGroupFilter('all')}
                                className={`w-full flex items-center justify-between rounded-md px-2.5 py-1.5 text-xs transition-colors ${
                                  dbGroupFilter === 'all'
                                    ? 'bg-slate-900 text-white font-medium'
                                    : 'hover:bg-slate-100 text-slate-700'
                                }`}
                              >
                                <span>全部組別</span>
                                <span className="font-mono text-[11px] opacity-80">
                                  {sidebarGroups.length} 組
                                </span>
                              </button>

                              {sidebarGroups.map((grp) => {
                                const totalInGrp = totalFacetCounts?.byGroup.get(grp.code) ?? 0;
                                const filteredInGrp = filteredFacetCounts.byGroup.get(grp.code) ?? 0;
                                const grpRowNums = recordRowsByFacet.byGroup.get(grp.code) ?? [];
                                const allGrpChecked =
                                  grpRowNums.length > 0 &&
                                  grpRowNums.every((r) => selectedRowNumbers.has(r));
                                const someGrpChecked =
                                  !allGrpChecked &&
                                  grpRowNums.some((r) => selectedRowNumbers.has(r));
                                const isActive = dbGroupFilter === grp.code;

                                return (
                                  <div
                                    key={`${grp.id}-${grp.code}`}
                                    className={`flex items-center gap-1.5 rounded-md px-2 py-1.5 text-xs transition-colors ${
                                      isActive
                                        ? 'bg-blue-50 border border-blue-200 text-blue-950 font-medium'
                                        : 'hover:bg-slate-100 text-slate-700'
                                    }`}
                                  >
                                    <input
                                      type="checkbox"
                                      checked={allGrpChecked}
                                      disabled={grpRowNums.length === 0}
                                      ref={(el) => {
                                        if (el) el.indeterminate = someGrpChecked;
                                      }}
                                      onChange={() => handleToggleGroupOfRows(grpRowNums)}
                                      title={`勾選 / 取消組別 ${grp.code} (${grpRowNums.length} 筆)`}
                                      className="h-3.5 w-3.5 rounded border-slate-300 cursor-pointer shrink-0"
                                    />
                                    <button
                                      type="button"
                                      onClick={() =>
                                        setDbGroupFilter(isActive ? 'all' : grp.code)
                                      }
                                      className="flex-1 flex items-center justify-between text-left truncate gap-1"
                                    >
                                      <span className="truncate">
                                        <span className="font-mono font-semibold mr-1">
                                          {grp.code}
                                        </span>
                                        <span className="text-slate-500">
                                          {grp.nameZh || grp.category}
                                        </span>
                                      </span>
                                      <span className="font-mono text-[11px] text-slate-500 shrink-0">
                                        {filteredInGrp}/{totalInGrp}
                                      </span>
                                    </button>
                                  </div>
                                );
                              })}
                            </div>
                          </div>
                        )}
                      </CardContent>
                    </Card>
                  )}

                  {/* Main Right Table & Filter Card */}
                  <Card className="flex-1 min-w-0 w-full">
                    <CardHeader className="pb-3 space-y-3">
                      {/* Row 1: Title, Sidebar toggle, Column Presets */}
                      <div className="flex flex-wrap items-center justify-between gap-3">
                        <div className="flex items-center gap-2">
                          {!showSidebar && (
                            <Button
                              variant="outline"
                              size="sm"
                              onClick={() => setShowSidebar(true)}
                              title="展開界別/組別導覽面板"
                            >
                              <PanelLeftOpen className="h-4 w-4" />
                              展開界別導覽
                            </Button>
                          )}
                          <div>
                            <CardTitle className="text-base">DB 主名單資料表</CardTitle>
                            <CardDescription>
                              符合條件共 <strong>{filteredRecordRows.length}</strong> 筆人員記錄（表格顯示{' '}
                              {filteredDbRows.length} 列，檢視 {visibleDbHeaders.length} /{' '}
                              {workbook.dbHeaders.length} 欄；匯出時完整保留 50 欄）
                            </CardDescription>
                          </div>
                        </div>

                        {/* Column Group Selector */}
                        <div className="flex flex-wrap items-center gap-1.5">
                          <span className="text-xs text-slate-500 mr-1 flex items-center gap-1">
                            <FolderKanban className="h-3.5 w-3.5" />
                            欄位檢視：
                          </span>
                          {(
                            [
                              ['core', '核心與活動 (15欄)'],
                              ['contact', '聯絡與地址 (19欄)'],
                              ['personal', '個人與證件 (19欄)'],
                              ['all', `完整欄位 (${workbook.dbHeaders.length}欄)`],
                            ] as const
                          ).map(([presetKey, label]) => (
                            <Button
                              key={presetKey}
                              variant={columnPreset === presetKey ? 'default' : 'outline'}
                              size="sm"
                              onClick={() => setColumnPreset(presetKey)}
                            >
                              {label}
                            </Button>
                          ))}
                        </div>
                      </div>

                      {/* Row 2: Search + Quick Dimension Filters */}
                      <div className="flex flex-wrap items-center gap-2.5 pt-1">
                        <div className="relative flex-1 min-w-[260px] max-w-md">
                          <Search className="absolute left-2.5 top-2.5 h-4 w-4 text-slate-400" />
                          <Input
                            placeholder="多關鍵字搜尋 (空白=AND，|=OR，如：MA002 Y | MZ001)..."
                            value={dbSearch}
                            onChange={(e) => setDbSearch(e.target.value)}
                            className="pl-8"
                          />
                        </div>

                        {/* Attendance Filter */}
                        <div className="flex items-center gap-1.5">
                          <label className="text-xs text-slate-500">出席：</label>
                          <select
                            value={attendanceFilter}
                            onChange={(e) =>
                              setAttendanceFilter(
                                e.target.value as 'all' | 'Y' | '不出席活動' | 'blank',
                              )
                            }
                            className="h-8 rounded-md border border-slate-200 bg-white px-2 text-xs text-slate-800"
                          >
                            <option value="all">全部 ({workbook.summary.totalRecords})</option>
                            <option value="Y">出席 Y ({workbook.summary.attendingCount})</option>
                            <option value="不出席活動">
                              不出席 ({workbook.summary.notAttendingCount})
                            </option>
                            <option value="blank">未填空白</option>
                          </select>
                        </div>

                        {/* Reception % Filter */}
                        <div className="flex items-center gap-1.5">
                          <label className="text-xs text-slate-500">酒會：</label>
                          <select
                            value={receptionFilter}
                            onChange={(e) =>
                              setReceptionFilter(e.target.value as 'all' | '%' | 'blank')
                            }
                            className="h-8 rounded-md border border-slate-200 bg-white px-2 text-xs text-slate-800"
                          >
                            <option value="all">全部</option>
                            <option value="%">酒會 % ({workbook.summary.receptionCount})</option>
                            <option value="blank">無標記</option>
                          </select>
                        </div>

                        {/* Flag Raising * Filter */}
                        <div className="flex items-center gap-1.5">
                          <label className="text-xs text-slate-500">升旗：</label>
                          <select
                            value={flagRaisingFilter}
                            onChange={(e) =>
                              setFlagRaisingFilter(e.target.value as 'all' | '*' | 'blank')
                            }
                            className="h-8 rounded-md border border-slate-200 bg-white px-2 text-xs text-slate-800"
                          >
                            <option value="all">全部</option>
                            <option value="*">升旗 * ({workbook.summary.flagRaisingCount})</option>
                            <option value="blank">無標記</option>
                          </select>
                        </div>

                        {/* Headcount Calc Filter */}
                        <div className="flex items-center gap-1.5">
                          <label className="text-xs text-slate-500">計算：</label>
                          <select
                            value={calcFilter}
                            onChange={(e) =>
                              setCalcFilter(e.target.value as 'all' | '1' | '0')
                            }
                            className="h-8 rounded-md border border-slate-200 bg-white px-2 text-xs text-slate-800"
                          >
                            <option value="all">全部</option>
                            <option value="1">計算=1 ({workbook.summary.countedHeadcount})</option>
                            <option value="0">不計算=0</option>
                          </select>
                        </div>

                        {/* Selection State Filter */}
                        <div className="flex items-center gap-1.5">
                          <label className="text-xs text-slate-500">勾選狀態：</label>
                          <select
                            value={selectionFilter}
                            onChange={(e) =>
                              setSelectionFilter(
                                e.target.value as 'all' | 'selected' | 'unselected',
                              )
                            }
                            className="h-8 rounded-md border border-slate-200 bg-white px-2 text-xs text-slate-800"
                          >
                            <option value="all">全部顯示</option>
                            <option value="selected">僅看已勾選 ({selectedRowNumbers.size})</option>
                            <option value="unselected">僅看未勾選</option>
                          </select>
                        </div>

                        <Button
                          variant={showCustomFilterPanel || customRules.length > 0 ? 'secondary' : 'outline'}
                          size="sm"
                          onClick={() => setShowCustomFilterPanel((prev) => !prev)}
                        >
                          <SlidersHorizontal className="h-3.5 w-3.5" />
                          自訂欄位條件 {customRules.length > 0 ? `(${customRules.length})` : ''}
                        </Button>

                        <Button
                          variant={
                            showExternalFilterPanel || importedFilterFiles.length > 0
                              ? 'secondary'
                              : 'outline'
                          }
                          size="sm"
                          onClick={() => setShowExternalFilterPanel((prev) => !prev)}
                        >
                          <Filter className="h-3.5 w-3.5 text-indigo-600" />
                          外部篩選 Excel{' '}
                          {importedFilterFiles.length > 0 ? `(${importedFilterFiles.length})` : ''}
                        </Button>

                        <Button
                          variant={
                            showExportColumnPanel || exportColumnMode !== 'db_all'
                              ? 'secondary'
                              : 'outline'
                          }
                          size="sm"
                          onClick={() => setShowExportColumnPanel((prev) => !prev)}
                        >
                          <FolderKanban className="h-3.5 w-3.5 text-emerald-600" />
                          匯出欄位設定 ({activeOutputHeaders.length} 欄)
                        </Button>

                        <Button
                          variant={showSectionHeaders ? 'secondary' : 'outline'}
                          size="sm"
                          onClick={() => setShowSectionHeaders((prev) => !prev)}
                          disabled={sort.column !== null}
                        >
                          {showSectionHeaders ? '✓ 顯示分組標題列' : '隱藏分組標題列'}
                        </Button>

                        {activeFilterCount > 0 && (
                          <Button
                            variant="ghost"
                            size="sm"
                            onClick={resetAllDbFilters}
                            className="text-rose-600 hover:text-rose-700 hover:bg-rose-50"
                          >
                            <RotateCcw className="h-3.5 w-3.5" />
                            清除篩選 ({activeFilterCount})
                          </Button>
                        )}
                      </div>

                      {/* Active Facet Pills */}
                      {(dbSectorFilter !== 'all' ||
                        dbCategoryFilter !== 'all' ||
                        dbGroupFilter !== 'all' ||
                        externalFilterMode !== 'all') && (
                        <div className="flex flex-wrap items-center gap-2 pt-1">
                          <span className="text-xs text-slate-500 flex items-center gap-1">
                            <Filter className="h-3 w-3" />
                            目前鎖定階層：
                          </span>
                          {dbSectorFilter !== 'all' && (
                            <Badge variant="default" className="gap-1.5">
                              界別：{dbSectorFilter}
                              <button
                                type="button"
                                onClick={() => setDbSectorFilter('all')}
                                className="ml-1 hover:opacity-75"
                              >
                                ×
                              </button>
                            </Badge>
                          )}
                          {dbCategoryFilter !== 'all' && (
                            <Badge variant="default" className="gap-1.5">
                              職務分類：{dbCategoryFilter}
                              <button
                                type="button"
                                onClick={() => setDbCategoryFilter('all')}
                                className="ml-1 hover:opacity-75"
                              >
                                ×
                              </button>
                            </Badge>
                          )}
                          {dbGroupFilter !== 'all' && (
                            <Badge variant="default" className="gap-1.5">
                              組別：{dbGroupFilter}
                              <button
                                type="button"
                                onClick={() => setDbGroupFilter('all')}
                                className="ml-1 hover:opacity-75"
                              >
                                ×
                              </button>
                            </Badge>
                          )}
                          {externalFilterMode !== 'all' && (
                            <Badge variant="default" className="gap-1.5 bg-indigo-700 text-white">
                              外部篩選檔：
                              {externalFilterMode === 'union'
                                ? '任一篩選檔命中 (聯集)'
                                : externalFilterMode === 'intersection'
                                  ? '全部篩選檔皆命中 (交集)'
                                  : (importedFilterFiles.find((f) => f.id === externalFilterMode)
                                      ?.fileName ?? externalFilterMode)}
                              <button
                                type="button"
                                onClick={() => setExternalFilterMode('all')}
                                className="ml-1 hover:opacity-75"
                              >
                                ×
                              </button>
                            </Badge>
                          )}
                        </div>
                      )}

                      {/* Custom Multi-Column Filter Builder Panel */}
                      {showCustomFilterPanel && (
                        <div className="rounded-lg border border-slate-200 bg-slate-50/90 p-3 space-y-2.5">
                          <div className="flex flex-wrap items-center justify-between gap-2">
                            <div className="flex items-center gap-2 text-xs font-medium text-slate-700">
                              <SlidersHorizontal className="h-3.5 w-3.5 text-slate-500" />
                              <span>自訂多欄位複合條件（支援全部 50 個欄位）</span>
                              <div className="flex items-center rounded-md border border-slate-200 bg-white p-0.5 ml-2">
                                <button
                                  type="button"
                                  onClick={() => setCustomRulesMode('AND')}
                                  className={`px-2 py-0.5 text-[11px] rounded ${
                                    customRulesMode === 'AND'
                                      ? 'bg-slate-900 text-white'
                                      : 'text-slate-600'
                                  }`}
                                >
                                  AND (滿足全部)
                                </button>
                                <button
                                  type="button"
                                  onClick={() => setCustomRulesMode('OR')}
                                  className={`px-2 py-0.5 text-[11px] rounded ${
                                    customRulesMode === 'OR'
                                      ? 'bg-slate-900 text-white'
                                      : 'text-slate-600'
                                  }`}
                                >
                                  OR (滿足任一)
                                </button>
                              </div>
                            </div>

                            <div className="flex items-center gap-2">
                              <Button variant="outline" size="sm" onClick={handleAddCustomRule}>
                                <Plus className="h-3.5 w-3.5" />
                                新增欄位條件
                              </Button>
                              {customRules.length > 0 && (
                                <Button
                                  variant="ghost"
                                  size="sm"
                                  onClick={() => setCustomRules([])}
                                  className="text-slate-500"
                                >
                                  清空條件
                                </Button>
                              )}
                            </div>
                          </div>

                          {customRules.length === 0 ? (
                            <p className="text-xs text-slate-400">
                              尚未設定自訂欄位條件。點擊「新增欄位條件」可針對職位、部門/機構、手機、地址、不出席原因等任意欄位進行進階篩選。
                            </p>
                          ) : (
                            <div className="space-y-2">
                              {customRules.map((rule, idx) => {
                                const needsValue =
                                  rule.operator !== 'is_empty' && rule.operator !== 'is_not_empty';
                                return (
                                  <div
                                    key={rule.id}
                                    className="flex flex-wrap items-center gap-2 text-xs"
                                  >
                                    <span className="w-10 text-right font-mono text-slate-400">
                                      {idx === 0 ? '條件' : customRulesMode}
                                    </span>
                                    <select
                                      value={rule.column}
                                      onChange={(e) =>
                                        handleUpdateCustomRule(rule.id, { column: e.target.value })
                                      }
                                      className="h-8 rounded-md border border-slate-200 bg-white px-2 text-xs text-slate-800"
                                    >
                                      {workbook.dbHeaders.map((h) => (
                                        <option key={h} value={h}>
                                          {h}
                                        </option>
                                      ))}
                                    </select>

                                    <select
                                      value={rule.operator}
                                      onChange={(e) =>
                                        handleUpdateCustomRule(rule.id, {
                                          operator: e.target.value as CustomFilterOperator,
                                        })
                                      }
                                      className="h-8 rounded-md border border-slate-200 bg-white px-2 text-xs text-slate-800"
                                    >
                                      {(
                                        Object.entries(OPERATOR_LABELS) as [
                                          CustomFilterOperator,
                                          string,
                                        ][]
                                      ).map(([opKey, opLabel]) => (
                                        <option key={opKey} value={opKey}>
                                          {opLabel}
                                        </option>
                                      ))}
                                    </select>

                                    {needsValue && (
                                      <Input
                                        placeholder="輸入比對內容..."
                                        value={rule.value}
                                        onChange={(e) =>
                                          handleUpdateCustomRule(rule.id, { value: e.target.value })
                                        }
                                        className="h-8 w-56 text-xs bg-white"
                                      />
                                    )}

                                    <Button
                                      variant="ghost"
                                      size="sm"
                                      onClick={() => handleRemoveCustomRule(rule.id)}
                                      className="h-8 px-2 text-slate-400 hover:text-rose-600"
                                      title="移除此條件"
                                    >
                                      <Trash2 className="h-3.5 w-3.5" />
                                    </Button>
                                  </div>
                                );
                              })}
                            </div>
                          )}
                        </div>
                      )}

                      {/* External Filter Excel Multi-Import Panel */}
                      {showExternalFilterPanel && (
                        <div className="rounded-lg border border-indigo-200 bg-indigo-50/40 p-3.5 space-y-3">
                          <div className="flex flex-wrap items-center justify-between gap-2">
                            <div className="flex flex-wrap items-center gap-2 text-xs font-medium text-slate-800">
                              <Filter className="h-3.5 w-3.5 text-indigo-600" />
                              <span>
                                外部匯出 Excel 篩選與選取（支援多次匯入，自動容錯增刪欄位與活動專屬欄位）
                              </span>
                              {importedFilterFiles.length > 0 && (
                                <Badge variant="secondary" className="text-[11px]">
                                  已匯入 {importedFilterFiles.length} 份檔案 · 聯集命中{' '}
                                  {externalUnionRowNumbers.size} 筆主 DB 名單
                                </Badge>
                              )}
                            </div>

                            <div className="flex flex-wrap items-center gap-2">
                              <Button
                                variant="default"
                                size="sm"
                                onClick={handleImportExternalFilterClick}
                                disabled={isLoading}
                              >
                                <Plus className="h-3.5 w-3.5" />
                                匯入篩選 Excel（可複選/多次匯入）
                              </Button>

                              {importedFilterFiles.length > 0 && (
                                <>
                                  <Button
                                    variant="outline"
                                    size="sm"
                                    disabled={externalUnionRowNumbers.size === 0}
                                    onClick={() =>
                                      setSelectedRowNumbers((prev) =>
                                        applyBatchSelection(
                                          prev,
                                          Array.from(externalUnionRowNumbers),
                                          'select',
                                        ),
                                      )
                                    }
                                  >
                                    勾選全部篩選檔命中 ({externalUnionRowNumbers.size})
                                  </Button>
                                  <Button
                                    variant="outline"
                                    size="sm"
                                    disabled={externalUnionRowNumbers.size === 0}
                                    onClick={() =>
                                      setSelectedRowNumbers((prev) =>
                                        applyBatchSelection(
                                          prev,
                                          Array.from(externalUnionRowNumbers),
                                          'deselect',
                                        ),
                                      )
                                    }
                                  >
                                    取消全部篩選檔勾選
                                  </Button>
                                  <Button
                                    variant="ghost"
                                    size="sm"
                                    onClick={() => {
                                      const reconciled = reconcileExternalFilterState([], {
                                        externalFilterMode,
                                        exportTemplateFileId,
                                        exportColumnMode,
                                      });
                                      setImportedFilterFiles([]);
                                      setExternalFilterMode(reconciled.externalFilterMode);
                                      setExportTemplateFileId(reconciled.exportTemplateFileId);
                                      setExportColumnMode(reconciled.exportColumnMode);
                                    }}
                                    className="text-slate-500 hover:text-rose-600"
                                  >
                                    清空篩選檔
                                  </Button>
                                </>
                              )}
                            </div>
                          </div>

                          {/* Global Default Match Key & Per-Group Match Overrides */}
                          <div className="rounded-md border border-indigo-100 bg-white/90 p-3 space-y-2.5 text-xs">
                            <div className="flex flex-wrap items-center justify-between gap-2">
                              <div className="flex flex-wrap items-center gap-2">
                                <span className="font-semibold text-slate-800">
                                  預設對比依據：
                                </span>
                                <select
                                  value={
                                    isCustomMatchKey(defaultExternalMatchKey)
                                      ? CUSTOM_MATCH_KEY_PREFIX
                                      : defaultExternalMatchKey
                                  }
                                  onChange={(e) => {
                                    const val = e.target.value;
                                    const nextKey = isCustomMatchKey(val)
                                      ? resolveInitialCustomMatchKey(
                                          defaultExternalMatchKey,
                                          workbook.dbHeaders,
                                        )
                                      : val;
                                    handleChangeDefaultExternalMatchKey(nextKey);
                                  }}
                                  className="h-7 rounded border border-slate-200 bg-white px-2 text-xs text-slate-800"
                                >
                                  {getGlobalMatchKeyOptions(workbook.dbHeaders).map((opt) => (
                                    <option key={opt.key} value={opt.key}>
                                      {opt.label}
                                    </option>
                                  ))}
                                </select>
                                <span className="text-slate-400">
                                  （未單獨指定規則的組別皆以此預設依據比對）
                                </span>
                              </div>

                              <Button
                                variant="outline"
                                size="sm"
                                className="h-7 text-xs"
                                onClick={handleAddGroupMatchOverride}
                              >
                                <Plus className="h-3.5 w-3.5" />
                                新增特定組別對比依據
                              </Button>
                            </div>

                            {renderCustomCompositeColumnPicker(
                              defaultExternalMatchKey,
                              workbook.dbHeaders,
                              handleChangeDefaultExternalMatchKey,
                            )}

                            {groupMatchOverrides.length > 0 && (
                              <div className="space-y-2.5 pt-1 border-t border-slate-100">
                                <div className="flex flex-wrap items-center justify-between gap-2">
                                  <span className="font-medium text-slate-700">
                                    各組別專屬對比依據（共 {groupMatchOverrides.length} 條規則 · 支援多選組別與輸入 Prefix 快速批次選取）：
                                  </span>
                                </div>

                                <div className="space-y-2.5">
                                  {groupMatchOverrides.map((rule, ruleIdx) => {
                                    const selectedCodes = [
                                      ...(rule.groupCodes ?? []),
                                      ...(rule.groupCode ? [rule.groupCode] : []),
                                    ];
                                    const selectedUpperSet = new Set(
                                      selectedCodes.map((c) => c.trim().toUpperCase()),
                                    );
                                    const prefixVal = rulePrefixInputs[rule.id] ?? '';
                                    const filteredGroupOptions = filterGroupsByPrefix(
                                      allGroupOverrideOptions,
                                      prefixVal,
                                    );
                                    const groupMetaByUpper = new Map(
                                      allGroupOverrideOptions.map((g) => [
                                        g.code.toUpperCase(),
                                        g,
                                      ]),
                                    );

                                    return (
                                      <div
                                        key={rule.id}
                                        className="rounded-lg border border-slate-200 bg-slate-50/70 p-3 space-y-2"
                                      >
                                        {/* Row 1: Rule Match Key & Actions */}
                                        <div className="flex flex-wrap items-center justify-between gap-2">
                                          <div className="flex flex-wrap items-center gap-2">
                                            <Badge variant="outline" className="bg-white">
                                              規則 #{ruleIdx + 1}
                                            </Badge>
                                            <span className="text-slate-700 font-medium">
                                              使用對比依據：
                                            </span>
                                            <select
                                              value={
                                                isCustomMatchKey(rule.matchKey)
                                                  ? CUSTOM_MATCH_KEY_PREFIX
                                                  : rule.matchKey
                                              }
                                              onChange={(e) => {
                                                const val = e.target.value;
                                                const nextKey = isCustomMatchKey(val)
                                                  ? resolveInitialCustomMatchKey(
                                                      rule.matchKey,
                                                      workbook.dbHeaders,
                                                    )
                                                  : val;
                                                handleUpdateGroupMatchOverride(rule.id, {
                                                  matchKey: nextKey,
                                                });
                                              }}
                                              className="h-7 rounded border border-slate-200 bg-white px-2 text-xs text-slate-800"
                                            >
                                              {getGlobalMatchKeyOptions(workbook.dbHeaders).map(
                                                (opt) => (
                                                  <option key={opt.key} value={opt.key}>
                                                    {opt.label}
                                                  </option>
                                                ),
                                              )}
                                            </select>
                                            <Badge variant="secondary" className="text-[11px]">
                                              已選 {selectedCodes.length} 個組別
                                            </Badge>
                                          </div>

                                          <div className="flex items-center gap-1.5">
                                            {selectedCodes.length > 0 && (
                                              <Button
                                                variant="ghost"
                                                size="sm"
                                                onClick={() =>
                                                  handleUpdateGroupMatchOverride(rule.id, {
                                                    groupCodes: [],
                                                    groupCode: undefined,
                                                  })
                                                }
                                                className="h-7 text-xs text-slate-500 hover:text-rose-600"
                                              >
                                                清空已選組別
                                              </Button>
                                            )}
                                            <Button
                                              variant="ghost"
                                              size="sm"
                                              onClick={() =>
                                                handleRemoveGroupMatchOverride(rule.id)
                                              }
                                              className="h-7 px-2 text-slate-400 hover:text-rose-600"
                                              title="移除此組別規則"
                                            >
                                              <Trash2 className="h-3.5 w-3.5" />
                                            </Button>
                                          </div>
                                        </div>

                                        {renderCustomCompositeColumnPicker(
                                          rule.matchKey,
                                          workbook.dbHeaders,
                                          (nextKey) =>
                                            handleUpdateGroupMatchOverride(rule.id, {
                                              matchKey: nextKey,
                                            }),
                                        )}

                                        {/* Row 2: Prefix Quick Filter & Batch Select Bar */}
                                        <div className="flex flex-wrap items-center gap-2 bg-white rounded border border-slate-200 px-2.5 py-1.5">
                                          <span className="text-slate-600 font-medium">
                                            Prefix 快速選取組別：
                                          </span>
                                          <Input
                                            placeholder="輸入組別前綴或名稱 (如 MZ、MA01，按 Enter 直接全選)..."
                                            value={prefixVal}
                                            onChange={(e) =>
                                              setRulePrefixInputs((prev) => ({
                                                ...prev,
                                                [rule.id]: e.target.value,
                                              }))
                                            }
                                            onKeyDown={(e) => {
                                              if (e.key === 'Enter' && prefixVal.trim()) {
                                                e.preventDefault();
                                                handlePrefixBatchSelectInOverride(
                                                  rule.id,
                                                  prefixVal,
                                                  'select',
                                                );
                                              }
                                            }}
                                            className="h-7 w-64 text-xs"
                                          />
                                          <Button
                                            variant="default"
                                            size="sm"
                                            className="h-7 text-xs"
                                            disabled={filteredGroupOptions.length === 0}
                                            onClick={() =>
                                              handlePrefixBatchSelectInOverride(
                                                rule.id,
                                                prefixVal,
                                                'select',
                                              )
                                            }
                                          >
                                            <CheckSquare className="h-3.5 w-3.5" />
                                            {prefixVal.trim()
                                              ? `全選前綴「${prefixVal.trim().toUpperCase()}」(${filteredGroupOptions.length})`
                                              : `全選全部組別 (${filteredGroupOptions.length})`}
                                          </Button>
                                          <Button
                                            variant="outline"
                                            size="sm"
                                            className="h-7 text-xs"
                                            disabled={filteredGroupOptions.length === 0}
                                            onClick={() =>
                                              handlePrefixBatchSelectInOverride(
                                                rule.id,
                                                prefixVal,
                                                'deselect',
                                              )
                                            }
                                          >
                                            {prefixVal.trim()
                                              ? `取消前綴「${prefixVal.trim().toUpperCase()}」`
                                              : '取消全部'}
                                          </Button>
                                          {prefixVal.trim() && (
                                            <Button
                                              variant="ghost"
                                              size="sm"
                                              className="h-7 text-xs text-slate-500"
                                              onClick={() =>
                                                setRulePrefixInputs((prev) => ({
                                                  ...prev,
                                                  [rule.id]: '',
                                                }))
                                              }
                                            >
                                              清除前綴過濾
                                            </Button>
                                          )}
                                        </div>

                                        {/* Row 3: Selected Group Chips */}
                                        {selectedCodes.length > 0 && (
                                          <div className="flex flex-wrap items-center gap-1.5 pt-0.5">
                                            <span className="text-[11px] text-slate-500">
                                              已套用組別：
                                            </span>
                                            {selectedCodes.map((code) => {
                                              const meta = groupMetaByUpper.get(
                                                code.trim().toUpperCase(),
                                              );
                                              return (
                                                <Badge
                                                  key={code}
                                                  variant="secondary"
                                                  className="gap-1 bg-indigo-100/80 text-indigo-950 text-[11px]"
                                                >
                                                  <span>
                                                    {code}
                                                    {meta?.nameZh ? ` ${meta.nameZh}` : ''}
                                                  </span>
                                                  <button
                                                    type="button"
                                                    onClick={() =>
                                                      handleToggleGroupInOverride(rule.id, code)
                                                    }
                                                    className="ml-0.5 text-slate-500 hover:text-rose-600"
                                                    title={`移除 ${code}`}
                                                  >
                                                    ×
                                                  </button>
                                                </Badge>
                                              );
                                            })}
                                          </div>
                                        )}

                                        {/* Row 4: Multi-Select Checkbox Grid */}
                                        <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-5 gap-1 max-h-36 overflow-y-auto rounded border border-slate-200 bg-white p-1.5">
                                          {filteredGroupOptions.length === 0 ? (
                                            <span className="col-span-full text-slate-400 py-2 text-center">
                                              找不到符合前綴「{prefixVal}」的組別
                                            </span>
                                          ) : (
                                            filteredGroupOptions.map((g) => {
                                              const checked = selectedUpperSet.has(
                                                g.code.toUpperCase(),
                                              );
                                              return (
                                                <label
                                                  key={g.code}
                                                  className={`flex items-center gap-1.5 rounded px-2 py-1 cursor-pointer truncate ${
                                                    checked
                                                      ? 'bg-indigo-50 text-indigo-950 font-medium'
                                                      : 'text-slate-600 hover:bg-slate-50'
                                                  }`}
                                                  title={`${g.code} ${g.nameZh}${
                                                    g.externalCount > 0
                                                      ? ` (篩選檔內 ${g.externalCount} 筆)`
                                                      : ''
                                                  }`}
                                                >
                                                  <input
                                                    type="checkbox"
                                                    checked={checked}
                                                    onChange={() =>
                                                      handleToggleGroupInOverride(
                                                        rule.id,
                                                        g.code,
                                                      )
                                                    }
                                                    className="h-3.5 w-3.5 rounded border-slate-300 shrink-0"
                                                  />
                                                  <span className="truncate">
                                                    {g.code} {g.nameZh}
                                                    {g.externalCount > 0
                                                      ? ` (${g.externalCount})`
                                                      : ''}
                                                  </span>
                                                </label>
                                              );
                                            })
                                          )}
                                        </div>
                                      </div>
                                    );
                                  })}
                                </div>
                              </div>
                            )}
                          </div>

                          {importedFilterFiles.length > 0 && (
                            <div className="flex flex-wrap items-center gap-2 text-xs bg-white/80 rounded-md border border-indigo-100 px-3 py-2">
                              <span className="text-slate-600 font-medium">
                                表格過濾模式：
                              </span>
                              <select
                                value={externalFilterMode}
                                onChange={(e) => setExternalFilterMode(e.target.value)}
                                className="h-7 rounded border border-slate-200 bg-white px-2 text-xs text-slate-800"
                              >
                                <option value="all">不限（顯示全部主 DB 名單）</option>
                                <option value="union">
                                  任一篩選檔命中（聯集 {externalUnionRowNumbers.size} 筆）
                                </option>
                                {importedFilterFiles.length > 1 && (
                                  <option value="intersection">
                                    全部篩選檔皆命中（交集{' '}
                                    {computeExternalFilterRowSet(
                                      importedFilterFiles,
                                      'intersection',
                                    )?.size ?? 0}{' '}
                                    筆）
                                  </option>
                                )}
                                {importedFilterFiles.map((f) => (
                                  <option key={f.id} value={f.id}>
                                    僅看檔案：{f.fileName}（命中{' '}
                                    {f.matchResult.matchedDbRowNumbers.size} 筆）
                                  </option>
                                ))}
                              </select>
                              <span className="text-slate-400">
                                提示：匯入時預設不會自動改動勾選狀態，請點擊下方檔案卡片的「勾選此檔命中名單」進行選取。
                              </span>
                            </div>
                          )}

                          {importedFilterFiles.length === 0 ? (
                            <p className="text-xs text-slate-500">
                              尚未匯入任何外部篩選 Excel。點擊右上角「匯入篩選 Excel」可選擇過去活動匯出的 Excel（即使欄位已被刪減或新增了活動專屬欄位皆可自動比對）。
                            </p>
                          ) : (
                            <div className="grid grid-cols-1 gap-2.5">
                              {importedFilterFiles.map((file) => {
                                const matchedNums = Array.from(
                                  file.matchResult.matchedDbRowNumbers,
                                );
                                const selectedInFileCount = matchedNums.filter((r) =>
                                  selectedRowNumbers.has(r),
                                ).length;
                                const isFilteringThisFile = externalFilterMode === file.id;

                                return (
                                  <div
                                    key={file.id}
                                    className="rounded-lg border border-slate-200 bg-white p-3 flex flex-wrap items-center justify-between gap-3 text-xs shadow-2xs"
                                  >
                                    <div className="space-y-1 min-w-[240px] flex-1">
                                      <div className="flex flex-wrap items-center gap-2">
                                        <FileSpreadsheet className="h-4 w-4 text-indigo-600 shrink-0" />
                                        <span className="font-semibold text-slate-900">
                                          {file.fileName}
                                        </span>
                                        <Badge variant="secondary" className="text-[11px]">
                                          命中主 DB {matchedNums.length} 筆（目前已勾選{' '}
                                          {selectedInFileCount} / {matchedNums.length}）
                                        </Badge>
                                        <span className="text-slate-400">
                                          外部檔共 {file.matchResult.totalExternalRowCount} 筆記錄（對應成功{' '}
                                          {file.matchResult.matchedExternalRowCount} 筆 · 未對應{' '}
                                          {file.matchResult.unmatchedExternalRowCount} 筆）
                                        </span>
                                      </div>

                                      <div className="flex flex-wrap items-center gap-2 text-[11px] text-slate-500">
                                        <span>
                                          共同 DB 欄位：{file.sheet.sharedHeaders.length} 欄
                                        </span>
                                        {file.sheet.extraHeaders.length > 0 && (
                                          <span className="text-amber-700 bg-amber-50 border border-amber-200 rounded px-1.5 py-0.5">
                                            活動專屬額外欄位 ({file.sheet.extraHeaders.length})：
                                            {file.sheet.extraHeaders.join('、')}
                                          </span>
                                        )}
                                      </div>

                                      {renderCustomCompositeColumnPicker(
                                        file.selectedMatchKey,
                                        file.sheet.sharedHeaders,
                                        (nextKey) =>
                                          handleChangeExternalMatchKey(file.id, nextKey),
                                      )}
                                    </div>

                                    <div className="flex flex-wrap items-center gap-2">
                                      <div className="flex items-center gap-1">
                                        <span className="text-slate-500">比對依據：</span>
                                        <select
                                          value={
                                            isCustomMatchKey(file.selectedMatchKey)
                                              ? CUSTOM_MATCH_KEY_PREFIX
                                              : file.selectedMatchKey
                                          }
                                          onChange={(e) => {
                                            const val = e.target.value;
                                            const nextKey = isCustomMatchKey(val)
                                              ? resolveInitialCustomMatchKey(
                                                  file.selectedMatchKey,
                                                  file.sheet.sharedHeaders,
                                                )
                                              : val;
                                            handleChangeExternalMatchKey(file.id, nextKey);
                                          }}
                                          className="h-7 rounded border border-slate-200 bg-slate-50 px-2 text-xs text-slate-800"
                                        >
                                          {file.availableMatchKeys.map((opt) => (
                                            <option key={opt.key} value={opt.key}>
                                              {opt.label}
                                            </option>
                                          ))}
                                        </select>
                                      </div>

                                      <Button
                                        variant="default"
                                        size="sm"
                                        className="h-7 text-xs"
                                        disabled={matchedNums.length === 0}
                                        onClick={() =>
                                          setSelectedRowNumbers((prev) =>
                                            applyBatchSelection(prev, matchedNums, 'select'),
                                          )
                                        }
                                      >
                                        <CheckSquare className="h-3.5 w-3.5" />
                                        勾選此檔命中 ({matchedNums.length})
                                      </Button>

                                      <Button
                                        variant="outline"
                                        size="sm"
                                        className="h-7 text-xs"
                                        disabled={selectedInFileCount === 0}
                                        onClick={() =>
                                          setSelectedRowNumbers((prev) =>
                                            applyBatchSelection(prev, matchedNums, 'deselect'),
                                          )
                                        }
                                      >
                                        取消勾選此檔
                                      </Button>

                                      <Button
                                        variant={isFilteringThisFile ? 'secondary' : 'outline'}
                                        size="sm"
                                        className="h-7 text-xs"
                                        onClick={() =>
                                          setExternalFilterMode(
                                            isFilteringThisFile ? 'all' : file.id,
                                          )
                                        }
                                      >
                                        {isFilteringThisFile ? '✓ 僅檢視此檔中' : '僅檢視此檔名單'}
                                      </Button>

                                      <Button
                                        variant="outline"
                                        size="sm"
                                        className="h-7 text-xs"
                                        onClick={() => handleApplyFileColumnsToExport(file)}
                                        title={`將匯出欄位設為此檔案的 ${file.sheet.headers.length} 個欄位`}
                                      >
                                        套用此檔欄位結構 ({file.sheet.headers.length}欄)
                                      </Button>

                                      <Button
                                        variant="ghost"
                                        size="sm"
                                        className="h-7 px-2 text-slate-400 hover:text-rose-600"
                                        onClick={() => handleRemoveExternalFilterFile(file.id)}
                                        title="移除此篩選 Excel"
                                      >
                                        <Trash2 className="h-3.5 w-3.5" />
                                      </Button>
                                    </div>
                                  </div>
                                );
                              })}
                            </div>
                          )}
                        </div>
                      )}

                      {/* Flexible Export Column Configuration Panel */}
                      {showExportColumnPanel && (
                        <div className="rounded-lg border border-emerald-200 bg-emerald-50/40 p-3.5 space-y-3 text-xs">
                          <div className="flex flex-wrap items-center justify-between gap-2">
                            <div className="flex flex-wrap items-center gap-2 font-medium text-slate-800">
                              <FolderKanban className="h-3.5 w-3.5 text-emerald-600" />
                              <span>匯出 Excel 欄位設定（按需選擇要匯出的欄位與活動專屬額外欄位）</span>
                              <Badge variant="secondary" className="text-[11px]">
                                目前將匯出 {activeOutputHeaders.length} 欄
                                {activeExtraHeaders.length > 0
                                  ? `（含 ${activeExtraHeaders.length} 個活動額外欄位）`
                                  : ''}
                              </Badge>
                            </div>

                            <div className="flex flex-wrap items-center gap-1.5">
                              <Button
                                variant={exportColumnMode === 'db_all' ? 'default' : 'outline'}
                                size="sm"
                                className="h-7 text-xs"
                                onClick={() => setExportColumnMode('db_all')}
                              >
                                主 DB 完整欄位 ({workbook.dbHeaders.length} 欄)
                              </Button>
                              <Button
                                variant={exportColumnMode === 'from_file' ? 'default' : 'outline'}
                                size="sm"
                                className="h-7 text-xs"
                                disabled={importedFilterFiles.length === 0}
                                onClick={() => {
                                  if (!exportTemplateFileId && importedFilterFiles[0]) {
                                    setExportTemplateFileId(importedFilterFiles[0].id);
                                  }
                                  setExportColumnMode('from_file');
                                }}
                              >
                                沿用篩選 Excel 欄位結構
                              </Button>
                              <Button
                                variant={exportColumnMode === 'custom' ? 'default' : 'outline'}
                                size="sm"
                                className="h-7 text-xs"
                                onClick={() => setExportColumnMode('custom')}
                              >
                                自訂勾選與增刪欄位
                              </Button>
                            </div>
                          </div>

                          {exportColumnMode === 'from_file' && (
                            <div className="flex flex-wrap items-center gap-2 bg-white rounded-md border border-emerald-100 p-2.5">
                              <span className="text-slate-600">選擇要沿用欄位結構的篩選 Excel：</span>
                              <select
                                value={exportTemplateFileId}
                                onChange={(e) => setExportTemplateFileId(e.target.value)}
                                className="h-7 rounded border border-slate-200 bg-white px-2 text-xs text-slate-800"
                              >
                                {importedFilterFiles.map((f) => (
                                  <option key={f.id} value={f.id}>
                                    {f.fileName}（共 {f.sheet.headers.length} 欄：含{' '}
                                    {f.sheet.sharedHeaders.length} 個 DB 欄 +{' '}
                                    {f.sheet.extraHeaders.length} 個額外欄）
                                  </option>
                                ))}
                              </select>
                              <Button
                                variant="outline"
                                size="sm"
                                className="h-7 text-xs"
                                onClick={() => {
                                  const f = importedFilterFiles.find(
                                    (item) => item.id === exportTemplateFileId,
                                  );
                                  if (f) {
                                    setCustomExportDbHeaders(f.sheet.sharedHeaders);
                                    setCustomExtraHeaders(f.sheet.extraHeaders);
                                    setExportColumnMode('custom');
                                  }
                                }}
                              >
                                以此檔欄位為基礎進入自訂微調
                              </Button>
                            </div>
                          )}

                          {exportColumnMode === 'custom' && (
                            <div className="space-y-2.5 bg-white rounded-md border border-emerald-100 p-3">
                              <div className="flex flex-wrap items-center justify-between gap-2">
                                <span className="font-medium text-slate-700">
                                  1. 勾選要保留的主 DB 欄位（已選 {customExportDbHeaders.length} /{' '}
                                  {workbook.dbHeaders.length} 欄）：
                                </span>
                                <div className="flex flex-wrap items-center gap-1.5">
                                  <Button
                                    variant="outline"
                                    size="sm"
                                    className="h-6 text-[11px]"
                                    onClick={() => setCustomExportDbHeaders(workbook.dbHeaders)}
                                  >
                                    全選 50 欄
                                  </Button>
                                  <Button
                                    variant="outline"
                                    size="sm"
                                    className="h-6 text-[11px]"
                                    onClick={() =>
                                      setCustomExportDbHeaders(
                                        getVisibleDbHeaders(workbook.dbHeaders, 'core'),
                                      )
                                    }
                                  >
                                    僅選核心 15 欄
                                  </Button>
                                  <Button
                                    variant="outline"
                                    size="sm"
                                    className="h-6 text-[11px]"
                                    onClick={() => setCustomExportDbHeaders(visibleDbHeaders)}
                                  >
                                    套用目前表格檢視欄位 ({visibleDbHeaders.length})
                                  </Button>
                                </div>
                              </div>

                              <div className="grid grid-cols-2 sm:grid-cols-4 md:grid-cols-6 lg:grid-cols-8 gap-1.5 max-h-40 overflow-y-auto p-1 border border-slate-100 rounded">
                                {workbook.dbHeaders.map((h) => {
                                  const checked = customExportDbHeaders.includes(h);
                                  return (
                                    <label
                                      key={h}
                                      className={`flex items-center gap-1.5 rounded px-2 py-1 cursor-pointer truncate ${
                                        checked
                                          ? 'bg-emerald-50 text-emerald-950 font-medium'
                                          : 'text-slate-500 hover:bg-slate-50'
                                      }`}
                                    >
                                      <input
                                        type="checkbox"
                                        checked={checked}
                                        onChange={() => handleToggleCustomExportDbHeader(h)}
                                        className="h-3.5 w-3.5 rounded border-slate-300"
                                      />
                                      <span className="truncate" title={h}>
                                        {h}
                                      </span>
                                    </label>
                                  );
                                })}
                              </div>

                              <div className="flex flex-wrap items-center gap-2 pt-1">
                                <span className="font-medium text-slate-700">
                                  2. 本次活動專屬額外欄位（如：桌號、車牌、簽到）：
                                </span>
                                <Input
                                  placeholder="輸入新欄位名稱..."
                                  value={newExtraHeaderInput}
                                  onChange={(e) => setNewExtraHeaderInput(e.target.value)}
                                  onKeyDown={(e) => {
                                    if (e.key === 'Enter') {
                                      e.preventDefault();
                                      handleAddCustomExtraHeader();
                                    }
                                  }}
                                  className="h-7 w-44 text-xs"
                                />
                                <Button
                                  variant="outline"
                                  size="sm"
                                  className="h-7 text-xs"
                                  onClick={handleAddCustomExtraHeader}
                                >
                                  <Plus className="h-3.5 w-3.5" />
                                  新增欄位
                                </Button>

                                {customExtraHeaders.map((eh) => (
                                  <Badge key={eh} variant="secondary" className="gap-1">
                                    + {eh}
                                    <button
                                      type="button"
                                      onClick={() => handleRemoveCustomExtraHeader(eh)}
                                      className="ml-1 hover:text-rose-600"
                                    >
                                      ×
                                    </button>
                                  </Badge>
                                ))}
                              </div>
                            </div>
                          )}

                          {activeExtraHeaders.length > 0 && (
                            <label className="flex items-center gap-2 bg-white rounded-md border border-amber-200 px-3 py-2 cursor-pointer text-slate-800">
                              <input
                                type="checkbox"
                                checked={populateExtraColumnValues}
                                onChange={(e) => setPopulateExtraColumnValues(e.target.checked)}
                                className="h-4 w-4 rounded border-slate-300"
                              />
                              <span>
                                <strong>帶入舊篩選檔中對應人員的額外欄位資料</strong>
                                （目前額外欄位：{activeExtraHeaders.join('、')}；預設不勾選時僅保留空白欄位標題供本次活動重新填寫）
                              </span>
                            </label>
                          )}
                        </div>
                      )}

                      {/* Row 3: Batch Selection Toolbar & Sort Status Indicator */}
                      <div className="flex flex-wrap items-center justify-between gap-2 rounded-lg border border-slate-200 bg-slate-50 px-3 py-2 text-xs">
                        <div className="flex flex-wrap items-center gap-2">
                          <span className="font-medium text-slate-700">
                            批次選取（已勾選 <strong className="text-blue-700">{selectedRowNumbers.size}</strong> 筆）：
                          </span>
                          <Button
                            variant="outline"
                            size="sm"
                            className="h-7 text-xs"
                            disabled={filteredRecordRows.length === 0}
                            onClick={() =>
                              setSelectedRowNumbers((prev) =>
                                applyBatchSelection(
                                  prev,
                                  filteredRecordRows.map((r) => r.rowNumber),
                                  'select',
                                ),
                              )
                            }
                          >
                            全選目前篩選 ({filteredRecordRows.length})
                          </Button>
                          <Button
                            variant="outline"
                            size="sm"
                            className="h-7 text-xs"
                            disabled={filteredRecordRows.length === 0}
                            onClick={() =>
                              setSelectedRowNumbers((prev) =>
                                applyBatchSelection(
                                  prev,
                                  filteredRecordRows.map((r) => r.rowNumber),
                                  'deselect',
                                ),
                              )
                            }
                          >
                            取消目前篩選
                          </Button>
                          <Button
                            variant="outline"
                            size="sm"
                            className="h-7 text-xs"
                            disabled={filteredRecordRows.length === 0}
                            onClick={() =>
                              setSelectedRowNumbers((prev) =>
                                applyBatchSelection(
                                  prev,
                                  filteredRecordRows.map((r) => r.rowNumber),
                                  'invert',
                                ),
                              )
                            }
                          >
                            反選目前篩選
                          </Button>
                          {selectedRowNumbers.size > 0 && (
                            <Button
                              variant="ghost"
                              size="sm"
                              className="h-7 text-xs text-rose-600 hover:text-rose-700"
                              onClick={() => setSelectedRowNumbers(new Set())}
                            >
                              清空全部已選
                            </Button>
                          )}
                        </div>

                        <div className="flex items-center gap-2">
                          {sort.column ? (
                            <div className="flex items-center gap-2 rounded bg-blue-50 border border-blue-200 px-2.5 py-1 text-blue-900">
                              <ArrowUpDown className="h-3.5 w-3.5 text-blue-600" />
                              <span>
                                檢視排序：<strong>{sort.column}</strong> (
                                {sort.direction === 'asc' ? '升冪 ▲' : '降冪 ▼'}
                                ) · 匯出仍維持原檔組別與列順序
                              </span>
                              <button
                                type="button"
                                onClick={() => setSort({ column: null, direction: 'asc' })}
                                className="ml-1 underline font-medium hover:text-blue-700"
                              >
                                恢復原檔順序
                              </button>
                            </div>
                          ) : (
                            <span className="text-slate-400">
                              提示：點擊任意欄位標題可切換升冪/降冪排序檢視
                            </span>
                          )}
                        </div>
                      </div>
                    </CardHeader>

                    <CardContent className="p-0">
                      <VirtualizedDbTable
                        rows={filteredDbRows}
                        allRows={workbook.dbRows}
                        visibleHeaders={visibleDbHeaders}
                        selectedRowNumbers={selectedRowNumbers}
                        sort={sort}
                        onToggleSort={(col) => setSort((prev) => getNextSortState(prev, col))}
                        onToggleRecordRow={handleToggleRecordRow}
                        onToggleSectionRows={(_secRow, targetRows) =>
                          handleToggleGroupOfRows(targetRows)
                        }
                        onToggleAllVisibleRecords={handleToggleAllVisibleRecords}
                      />
                    </CardContent>
                  </Card>
                </div>
              </TabsContent>

              {/* Tab 2: 組別 Sheet */}
              <TabsContent value="groups">
                <Card>
                  <CardHeader className="pb-4 space-y-3">
                    <div>
                      <CardTitle className="text-base">組別定義表 (組別)</CardTitle>
                      <CardDescription>
                        顯示 {filteredGroups.length} / {workbook.groups.length} 個組別（匯出時將完整保留此工作表）
                      </CardDescription>
                    </div>

                    <div className="flex flex-wrap items-center gap-3">
                      <div className="relative flex-1 min-w-[220px] max-w-xs">
                        <Search className="absolute left-2.5 top-2.5 h-4 w-4 text-slate-400" />
                        <Input
                          placeholder="搜尋組別簡稱、ID、分類或中文名..."
                          value={groupSearch}
                          onChange={(e) => setGroupSearch(e.target.value)}
                          className="pl-8"
                        />
                      </div>

                      <div className="flex items-center gap-2">
                        <label className="text-xs text-slate-500">界別：</label>
                        <select
                          value={sectorFilter}
                          onChange={(e) => setSectorFilter(e.target.value)}
                          className="h-9 rounded-md border border-slate-200 bg-white px-2.5 text-xs text-slate-800"
                        >
                          <option value="all">全部界別 ({workbook.sectors.length})</option>
                          {workbook.sectors.map((s) => (
                            <option key={s.code} value={s.code}>
                              {s.code} {s.nameZh ? `(${s.nameZh})` : ''}
                            </option>
                          ))}
                        </select>
                      </div>

                      <div className="flex items-center gap-2">
                        <label className="text-xs text-slate-500">職務分類：</label>
                        <select
                          value={categoryFilter}
                          onChange={(e) => setCategoryFilter(e.target.value)}
                          className="h-9 rounded-md border border-slate-200 bg-white px-2.5 text-xs text-slate-800"
                        >
                          <option value="all">全部分類 ({uniqueCategories.length})</option>
                          {uniqueCategories.map((cat) => (
                            <option key={cat} value={cat}>
                              {cat}
                            </option>
                          ))}
                        </select>
                      </div>
                    </div>
                  </CardHeader>

                  <CardContent className="p-0">
                    <div className="max-h-[600px] overflow-auto border-t border-slate-200">
                      <Table>
                        <TableHeader>
                          <TableRow>
                            <TableHead className="w-20">ID</TableHead>
                            <TableHead className="w-28">界別</TableHead>
                            <TableHead className="w-48">分類</TableHead>
                            <TableHead className="w-36">簡稱</TableHead>
                            <TableHead>中文名稱</TableHead>
                          </TableRow>
                        </TableHeader>
                        <TableBody>
                          {filteredGroups.map((group) => (
                            <TableRow key={`${group.id}-${group.code}`}>
                              <TableCell className="font-mono text-slate-500">{group.id}</TableCell>
                              <TableCell>
                                <Badge variant="secondary" className="font-mono">
                                  {group.sectorCode}
                                </Badge>
                              </TableCell>
                              <TableCell>
                                {group.category === '--' ? (
                                  <span className="text-slate-400">-- (大組標題)</span>
                                ) : (
                                  <Badge variant="outline">{group.category}</Badge>
                                )}
                              </TableCell>
                              <TableCell className="font-mono font-semibold">
                                {group.code}
                              </TableCell>
                              <TableCell>
                                {group.nameZh || <span className="text-slate-300">-</span>}
                              </TableCell>
                            </TableRow>
                          ))}
                        </TableBody>
                      </Table>
                    </div>
                  </CardContent>
                </Card>
              </TabsContent>

              {/* Tab 3: 界別 Sheet */}
              <TabsContent value="sectors">
                <Card>
                  <CardHeader className="pb-4">
                    <CardTitle className="text-base">界別總覽表 (界別)</CardTitle>
                    <CardDescription>
                      共 {workbook.sectors.length} 個頂層界別（MA000 – MZ000，匯出時將完整保留此工作表）
                    </CardDescription>
                  </CardHeader>

                  <CardContent className="p-0">
                    <div className="max-h-[600px] overflow-auto border-t border-slate-200">
                      <Table>
                        <TableHeader>
                          <TableRow>
                            <TableHead className="w-20">ID</TableHead>
                            <TableHead className="w-32">簡稱</TableHead>
                            <TableHead>中文名稱</TableHead>
                            <TableHead className="w-36">下轄組別數</TableHead>
                            <TableHead>備註</TableHead>
                          </TableRow>
                        </TableHeader>
                        <TableBody>
                          {workbook.sectors.map((sector) => (
                            <TableRow key={`${sector.id}-${sector.code}`}>
                              <TableCell className="font-mono text-slate-500">
                                {sector.id}
                              </TableCell>
                              <TableCell>
                                <Badge variant="default" className="font-mono">
                                  {sector.code}
                                </Badge>
                              </TableCell>
                              <TableCell className="font-medium">
                                {sector.nameZh || <span className="text-slate-300">-</span>}
                              </TableCell>
                              <TableCell>
                                <Badge variant="secondary">
                                  {groupCountBySector.get(sector.code) ?? 0} 個組別
                                </Badge>
                              </TableCell>
                              <TableCell>
                                {sector.remark ? (
                                  <Badge variant="warning">{sector.remark}</Badge>
                                ) : (
                                  <span className="text-slate-300">-</span>
                                )}
                              </TableCell>
                            </TableRow>
                          ))}
                        </TableBody>
                      </Table>
                    </div>
                  </CardContent>
                </Card>
              </TabsContent>
            </Tabs>
          </>
        )}
      </main>
    </div>
  );
}

