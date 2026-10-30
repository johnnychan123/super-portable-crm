import React, { useEffect, useMemo, useRef, useState } from 'react';
import pkg from '../package.json';
import {
  Building2,
  CheckCircle2,
  FileSpreadsheet,
  Filter,
  FolderKanban,
  Layers,
  Search,
  Upload,
  Users,
  XCircle,
} from 'lucide-react';
import type { AppRuntimeInfo } from './vite-env';
import {
  type DbColumnPreset,
  type ParsedCrmWorkbook,
  filterDbRows,
  filterGroups,
  getVisibleDbHeaders,
  parseCrmWorkbook,
} from './lib/excelParser';
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

export default function App(): React.ReactElement {
  const [info, setInfo] = useState<AppRuntimeInfo | null>(null);
  const [workbook, setWorkbook] = useState<ParsedCrmWorkbook | null>(null);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [isLoading, setIsLoading] = useState(false);

  // DB table filters
  const [dbSearch, setDbSearch] = useState('');
  const [attendanceFilter, setAttendanceFilter] = useState<'all' | 'Y' | '不出席活動'>('all');
  const [showSectionHeaders, setShowSectionHeaders] = useState(true);
  const [columnPreset, setColumnPreset] = useState<DbColumnPreset>('core');

  // 組別 table filters
  const [groupSearch, setGroupSearch] = useState('');
  const [sectorFilter, setSectorFilter] = useState('all');
  const [categoryFilter, setCategoryFilter] = useState('all');

  const fileInputRef = useRef<HTMLInputElement | null>(null);

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

  const handleLoadBuffer = (
    buffer: Uint8Array | ArrayBuffer,
    fileName: string,
    filePath?: string,
  ): void => {
    try {
      const parsed = parseCrmWorkbook(buffer, fileName, filePath);
      setWorkbook(parsed);
      setErrorMessage(null);
      setDbSearch('');
      setAttendanceFilter('all');
      setGroupSearch('');
      setSectorFilter('all');
      setCategoryFilter('all');
    } catch (err) {
      setErrorMessage(err instanceof Error ? err.message : '無法解析此 Excel 檔案');
    }
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

  const filteredDbRows = useMemo(
    () =>
      workbook
        ? filterDbRows(workbook.dbRows, {
            searchQuery: dbSearch,
            attendanceFilter,
            showSectionHeaders,
          })
        : [],
    [workbook, dbSearch, attendanceFilter, showSectionHeaders],
  );

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

  const uniqueCategories = useMemo(() => {
    if (!workbook) return [];
    return Array.from(new Set(workbook.groups.map((g) => g.category).filter(Boolean)));
  }, [workbook]);

  const groupCountBySector = useMemo(() => {
    const counts = new Map<string, number>();
    if (!workbook) return counts;
    for (const g of workbook.groups) {
      counts.set(g.sectorCode, (counts.get(g.sectorCode) ?? 0) + 1);
    }
    return counts;
  }, [workbook]);

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

      {/* Top Navigation Bar */}
      <header className="sticky top-0 z-20 border-b border-slate-200 bg-white/95 backdrop-blur px-6 py-3.5">
        <div className="mx-auto flex max-w-[1600px] items-center justify-between gap-4">
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
                Excel 活動名單與界別/組別資料庫管理系統
              </p>
            </div>
          </div>

          <div className="flex items-center gap-3">
            {workbook && (
              <div className="hidden sm:flex items-center gap-2 rounded-lg border border-slate-200 bg-slate-50 px-3 py-1.5 text-xs">
                <FileSpreadsheet className="h-3.5 w-3.5 text-emerald-600" />
                <span className="font-medium text-slate-800">{workbook.fileName}</span>
                <span className="text-slate-400">|</span>
                <span className="text-slate-500">
                  {workbook.sheetNames.join(' / ')} ({workbook.summary.totalRecords} 筆名單)
                </span>
              </div>
            )}

            <Button onClick={handleImportClick} disabled={isLoading} size="default">
              <Upload className="h-4 w-4" />
              {isLoading ? '正在讀取...' : workbook ? '重新匯入 Excel' : '匯入 Excel 檔案'}
            </Button>
          </div>
        </div>
      </header>

      {/* Main Content Area */}
      <main className="mx-auto w-full max-w-[1600px] flex-1 p-6 space-y-6">
        {errorMessage && (
          <div className="rounded-lg border border-rose-200 bg-rose-50 px-4 py-3 text-sm text-rose-700 flex items-center justify-between">
            <span>匯入錯誤：{errorMessage}</span>
            <Button variant="ghost" size="sm" onClick={() => setErrorMessage(null)}>
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
                <CardTitle className="text-xl">匯入 Excel 資料庫以開始檢視</CardTitle>
                <CardDescription className="max-w-md text-sm leading-relaxed">
                  支援匯入包含 <code className="rounded bg-slate-100 px-1.5 py-0.5 font-mono text-slate-800">DB</code>、
                  <code className="rounded bg-slate-100 px-1.5 py-0.5 font-mono text-slate-800">組別</code>、
                  <code className="rounded bg-slate-100 px-1.5 py-0.5 font-mono text-slate-800">界別</code> 三層結構之 Excel 活頁簿（如 <code className="rounded bg-slate-100 px-1.5 py-0.5 font-mono text-slate-800">test.xlsx</code>）。
                </CardDescription>
              </CardHeader>
              <CardContent className="flex flex-col items-center gap-4">
                <Button onClick={handleImportClick} disabled={isLoading} size="lg">
                  <Upload className="h-4 w-4" />
                  選取並匯入 Excel 檔案 (.xlsx)
                </Button>
                <div className="grid grid-cols-3 gap-3 w-full max-w-md pt-2 text-left">
                  <div className="rounded-lg border border-slate-200 bg-slate-50/70 p-3">
                    <div className="text-xs font-semibold text-slate-800">DB 主名單</div>
                    <div className="text-[11px] text-slate-500 mt-0.5">
                      50 欄位、分組標題列與出席標記解析
                    </div>
                  </div>
                  <div className="rounded-lg border border-slate-200 bg-slate-50/70 p-3">
                    <div className="text-xs font-semibold text-slate-800">組別清單</div>
                    <div className="text-[11px] text-slate-500 mt-0.5">
                      558 個組別代碼與 14 種職務分類篩選
                    </div>
                  </div>
                  <div className="rounded-lg border border-slate-200 bg-slate-50/70 p-3">
                    <div className="text-xs font-semibold text-slate-800">界別總覽</div>
                    <div className="text-[11px] text-slate-500 mt-0.5">
                      MA000–MZ000 共 26 個界別統計
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
                    含 {workbook.summary.sectionCount} 列分組標題分隔列
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
                    組別與界別規模
                  </CardTitle>
                  <Building2 className="h-4 w-4 text-slate-400" />
                </CardHeader>
                <CardContent>
                  <div className="flex items-center gap-3">
                    <div>
                      <span className="text-2xl font-bold">{workbook.summary.groupCount}</span>
                      <span className="ml-1 text-xs text-slate-500">個組別</span>
                    </div>
                    <span className="text-slate-300">/</span>
                    <div>
                      <span className="text-2xl font-bold">{workbook.summary.sectorCount}</span>
                      <span className="ml-1 text-xs text-slate-500">個界別</span>
                    </div>
                  </div>
                  <p className="mt-1 text-xs text-slate-400">
                    涵蓋 {uniqueCategories.length} 種職務分類
                  </p>
                </CardContent>
              </Card>
            </div>

            {/* Worksheet Tabs */}
            <Tabs defaultValue="DB" className="w-full">
              <div className="flex flex-wrap items-center justify-between gap-3">
                <TabsList>
                  <TabsTrigger value="DB">
                    DB 主名單 ({workbook.summary.totalRecords})
                  </TabsTrigger>
                  <TabsTrigger value="groups">
                    組別 ({workbook.groups.length})
                  </TabsTrigger>
                  <TabsTrigger value="sectors">
                    界別 ({workbook.sectors.length})
                  </TabsTrigger>
                </TabsList>

                {workbook.filePath && (
                  <span className="text-xs text-slate-400 truncate max-w-xl" title={workbook.filePath}>
                    檔案路徑：{workbook.filePath}
                  </span>
                )}
              </div>

              {/* Tab 1: DB Sheet */}
              <TabsContent value="DB">
                <Card>
                  <CardHeader className="pb-4 space-y-3">
                    <div className="flex flex-wrap items-center justify-between gap-3">
                      <div>
                        <CardTitle className="text-base">DB 主名單資料表</CardTitle>
                        <CardDescription>
                          目前顯示 {filteredDbRows.length} 列（共 {workbook.dbHeaders.length} 個欄位，已選取顯示{' '}
                          {visibleDbHeaders.length} 欄）
                        </CardDescription>
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

                    {/* Filter & Search Bar */}
                    <div className="flex flex-wrap items-center gap-3 pt-1">
                      <div className="relative flex-1 min-w-[240px] max-w-sm">
                        <Search className="absolute left-2.5 top-2.5 h-4 w-4 text-slate-400" />
                        <Input
                          placeholder="搜尋組別代碼 (如 MA002、MZ001)、姓名、備註..."
                          value={dbSearch}
                          onChange={(e) => setDbSearch(e.target.value)}
                          className="pl-8"
                        />
                      </div>

                      <div className="flex items-center gap-1.5">
                        <span className="text-xs text-slate-500 flex items-center gap-1">
                          <Filter className="h-3.5 w-3.5" />
                          出席篩選：
                        </span>
                        <Button
                          variant={attendanceFilter === 'all' ? 'secondary' : 'outline'}
                          size="sm"
                          onClick={() => setAttendanceFilter('all')}
                        >
                          全部 ({workbook.summary.totalRecords})
                        </Button>
                        <Button
                          variant={attendanceFilter === 'Y' ? 'secondary' : 'outline'}
                          size="sm"
                          onClick={() => setAttendanceFilter('Y')}
                        >
                          出席 Y ({workbook.summary.attendingCount})
                        </Button>
                        <Button
                          variant={attendanceFilter === '不出席活動' ? 'secondary' : 'outline'}
                          size="sm"
                          onClick={() => setAttendanceFilter('不出席活動')}
                        >
                          不出席 ({workbook.summary.notAttendingCount})
                        </Button>
                      </div>

                      <Button
                        variant={showSectionHeaders ? 'secondary' : 'outline'}
                        size="sm"
                        onClick={() => setShowSectionHeaders((prev) => !prev)}
                      >
                        {showSectionHeaders ? '✓ 顯示分組標題列' : '隱藏分組標題列'}
                      </Button>
                    </div>
                  </CardHeader>

                  <CardContent className="p-0">
                    <div className="max-h-[600px] overflow-auto border-t border-slate-200">
                      <Table>
                        <TableHeader>
                          <TableRow>
                            <TableHead className="w-14 text-center">列號</TableHead>
                            {visibleDbHeaders.map((header) => (
                              <TableHead key={header}>{header}</TableHead>
                            ))}
                          </TableRow>
                        </TableHeader>
                        <TableBody>
                          {filteredDbRows.length === 0 ? (
                            <TableRow>
                              <TableCell
                                colSpan={visibleDbHeaders.length + 1}
                                className="h-24 text-center text-slate-400"
                              >
                                沒有符合篩選條件的資料列
                              </TableCell>
                            </TableRow>
                          ) : (
                            filteredDbRows.map((row) => {
                              if (row.kind === 'section') {
                                const isTopSector = row.sectionCode?.endsWith('000') ?? false;
                                return (
                                  <TableRow
                                    key={`section-${row.rowNumber}`}
                                    className={
                                      isTopSector
                                        ? 'bg-slate-200/80 hover:bg-slate-200 font-semibold'
                                        : 'bg-slate-100/90 hover:bg-slate-100 font-medium'
                                    }
                                  >
                                    <TableCell className="text-center font-mono text-[11px] text-slate-500">
                                      #{row.rowNumber}
                                    </TableCell>
                                    <TableCell colSpan={visibleDbHeaders.length} className="py-2">
                                      <div className="flex flex-wrap items-center gap-2">
                                        <Badge variant={isTopSector ? 'default' : 'outline'}>
                                          {row.sectionCode}
                                        </Badge>
                                        <span className="text-slate-900">
                                          {row.sectionTitle || row.values['姓名']}
                                        </span>
                                        {row.values['Remark'] && (
                                          <span className="ml-2 rounded bg-amber-50 border border-amber-200 px-2 py-0.5 text-[11px] font-normal text-amber-800 whitespace-pre-wrap">
                                            備註：{row.values['Remark']}
                                          </span>
                                        )}
                                      </div>
                                    </TableCell>
                                  </TableRow>
                                );
                              }

                              return (
                                <TableRow key={`row-${row.rowNumber}`}>
                                  <TableCell className="text-center font-mono text-[11px] text-slate-400">
                                    {row.rowNumber}
                                  </TableCell>
                                  {visibleDbHeaders.map((header) => {
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
                                            val
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
                                      <TableCell
                                        key={header}
                                        className="max-w-xs truncate"
                                        title={val}
                                      >
                                        {val || <span className="text-slate-300">-</span>}
                                      </TableCell>
                                    );
                                  })}
                                </TableRow>
                              );
                            })
                          )}
                        </TableBody>
                      </Table>
                    </div>
                  </CardContent>
                </Card>
              </TabsContent>

              {/* Tab 2: 組別 Sheet */}
              <TabsContent value="groups">
                <Card>
                  <CardHeader className="pb-4 space-y-3">
                    <div>
                      <CardTitle className="text-base">組別定義表 (組別)</CardTitle>
                      <CardDescription>
                        顯示 {filteredGroups.length} / {workbook.groups.length} 個組別
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
                      共 {workbook.sectors.length} 個頂層界別（MA000 – MZ000）
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

