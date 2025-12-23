import { useEffect, useMemo, useRef, useState } from "react";
import { Button } from "../ui/Button";
import { Card } from "../ui/Card";
import { Badge } from "../ui/Badge";
import type { ConvertMode, ConvertResponse, TemplateInfo } from "../types";
import { convertApi } from "../api/convertApi";
import { templatesApi } from "../api/templatesApi";

type FileItem = {
  file: File;
  status: "idle" | "processing" | "success" | "error";
  result?: ConvertResponse | null;
  error?: string | null;
};

function formatNumber(n: number | undefined) {
  if (typeof n !== "number" || Number.isNaN(n)) return "-";
  return new Intl.NumberFormat("zh-CN").format(n);
}

function StatusBadge({ status }: { status: FileItem["status"] }) {
  if (status === "success") return <Badge variant="success">成功</Badge>;
  if (status === "error") return <Badge variant="error">失败</Badge>;
  if (status === "processing") return <Badge variant="info">处理中</Badge>;
  return <Badge variant="neutral">等待</Badge>;
}

function MetricBadge({ label, value }: { label: string; value: number | undefined }) {
  return (
    <span className="inline-flex items-center gap-1 rounded-md border border-slate-200 bg-white px-2 py-0.5 text-xs text-slate-700">
      <span className="text-slate-500">{label}</span>
      <span className="font-semibold text-slate-900">{formatNumber(value)}</span>
    </span>
  );
}

function classifyLog(line: string): { variant: "neutral" | "success" | "warning" | "error" | "info"; text: string; time: string } {
  const timeMatch = line.match(/^\[(.*?)\]\s*/);
  const time = timeMatch?.[1] ?? "";
  const text = line.replace(/^\[.*?\]\s*/, "");

  const lower = text.toLowerCase();
  const isError = text.includes("错误") || lower.includes("error") || lower.includes("failed");
  const isWarning = text.includes("警告") || lower.includes("warning");
  const isSuccess = text.includes("成功") || lower.includes("success");
  const isSummary = text.includes("汇总:");

  if (isError) return { variant: "error", text, time };
  if (isWarning) return { variant: "warning", text, time };
  if (isSummary) return { variant: "info", text, time };
  if (isSuccess) return { variant: "success", text, time };
  return { variant: "neutral", text, time };
}

function ModeTab({
  mode,
  current,
  onClick,
  label,
  className = "bg-white text-slate-600 border border-slate-200 hover:bg-slate-50 hover:text-slate-900"
}: {
  mode: ConvertMode;
  current: ConvertMode;
  onClick: (m: ConvertMode) => void;
  label: string;
  className?: string;
}) {
  return (
    <button
      onClick={() => onClick(mode)}
      className={[
        "px-2 py-2 rounded-md text-xs font-medium transition-colors w-full",
        current === mode 
          ? "bg-slate-900 text-white shadow-md" 
          : className
      ].join(" ")}
    >
      {label}
    </button>
  );
}

export function DashboardPage() {
  const [mode, setMode] = useState<ConvertMode>("allocation");
  const [fileList, setFileList] = useState<FileItem[]>([]);
  const [detailFile, setDetailFile] = useState<File | null>(null);
  const [templateFile, setTemplateFile] = useState<File | null>(null);
  const [weekNum, setWeekNum] = useState<string>("");
  const [templates, setTemplates] = useState<TemplateInfo[]>([]);
  const [selectedLibTemplate, setSelectedLibTemplate] = useState<string>("");
  const [logs, setLogs] = useState<string[]>([]);
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const detailInputRef = useRef<HTMLInputElement | null>(null);
  const templateInputRef = useRef<HTMLInputElement | null>(null);

  useEffect(() => {
    templatesApi
      .list()
      .then(setTemplates)
      .catch(() => setTemplates([]));
  }, []);

  useEffect(() => {
    setFileList([]);
    setDetailFile(null);
    setTemplateFile(null);
    setSelectedLibTemplate("");
    setLogs([]);
    setWeekNum("");
    if (fileInputRef.current) fileInputRef.current.value = "";
    if (detailInputRef.current) detailInputRef.current.value = "";
    if (templateInputRef.current) templateInputRef.current.value = "";
  }, [mode]);

  const okCount = useMemo(() => fileList.filter((f) => f.status === "success").length, [fileList]);
  const processingCount = useMemo(() => fileList.filter((f) => f.status === "processing").length, [fileList]);
  const successStats = useMemo(() => {
    const stats = fileList
      .filter((f) => f.status === "success" && f.result?.stats)
      .map((f) => f.result!.stats!);
    const sum = <K extends keyof NonNullable<ConvertResponse["stats"]>>(k: K) =>
      stats.reduce((acc, s) => acc + (typeof s[k] === "number" ? (s[k] as number) : 0), 0);

    const last = stats.length ? stats[stats.length - 1] : undefined;

    return {
      count: stats.length,
      store_count: sum("store_count"),
      box_count: sum("box_count"),
      sku_count: last?.sku_count,
      pt_count: last?.pt_count,
      total_qty: sum("total_qty"),
      jan_map_count: last?.jan_map_count,
      jan_match_success: last?.jan_match_success,
      jan_match_fail: last?.jan_match_fail
    };
  }, [fileList]);

  function addLog(message: string) {
    const t = new Date().toLocaleTimeString();
    setLogs((prev) => [...prev, `[${t}] ${message}`]);
  }

  function onPickFiles(files: FileList) {
    const newFiles: FileItem[] = Array.from(files).map((f) => ({ file: f, status: "idle", result: null, error: null }));
    setFileList((prev) => [...prev, ...newFiles]);
    addLog(`添加了 ${newFiles.length} 个文件`);
  }

  async function onPickFactoryFile(e: React.ChangeEvent<HTMLInputElement>) {
    // Deprecated
  }

  async function processOne(item: FileItem, index: number) {
    setFileList((prev) => {
      const next = [...prev];
      next[index] = { ...next[index], status: "processing", error: null };
      return next;
    });
    addLog(`[${item.file.name}] 开始处理...`);
    try {
      let res: ConvertResponse;
      
      if (mode === "box_label") {
        // Special handling for box label generation
        res = await convertApi.generateLabelsFromFile(item.file);
      } else {
        // Standard conversion
        res = await convertApi.convert({
          file: item.file,
          mode,
          templateFile,
          templateName: selectedLibTemplate || null,
          detailFile,
          weekNum: mode === "assortment" ? weekNum : null
        });
      }

      if (res.status !== "success") throw new Error(res.message || "Convert failed");
      
      addLog(`[${item.file.name}] 处理成功`);
      if (res.message) addLog(`> ${res.message}`);
      if (res.logs?.length) res.logs.forEach((l) => addLog(`> ${l}`));
      
      if (res.stats?.store_count || res.stats?.sku_count || res.stats?.box_count) {
        addLog(
          `[${item.file.name}] 汇总: 店铺 ${formatNumber(res.stats.store_count)}, 箱数 ${formatNumber(
            res.stats.box_count
          )}, SKU ${formatNumber(res.stats.sku_count)}, PT ${formatNumber(res.stats.pt_count)}, 总枚数 ${formatNumber(res.stats.total_qty)}`
        );
      }

      // Auto download for box labels if URL provided
      // Explicitly disable auto-download for box_label mode
      if (res.download_url && mode !== "box_label") {
        addLog(`[${item.file.name}] 准备下载...`);
        window.open(res.download_url, "_blank");
      }

      setFileList((prev) => {
        const next = [...prev];
        next[index] = { ...next[index], status: "success", result: res, error: null };
        return next;
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      addLog(`[${item.file.name}] 错误: ${msg}`);
      setFileList((prev) => {
        const next = [...prev];
        next[index] = { ...next[index], status: "error", error: msg };
        return next;
      });
    }
  }

  async function onConvertAll() {
    if (!fileList.length) return;
    if (mode === "allocation" && !detailFile) {
      addLog("allocation 模式需要先选择明细表");
      return;
    }
    for (let i = 0; i < fileList.length; i += 1) {
      if (fileList[i].status !== "success") {
        await processOne(fileList[i], i);
      }
    }
    addLog("所有任务处理完成");
  }

  return (
    <div className="space-y-6">
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 items-start">
        <div className="lg:col-span-2">
          <div className="text-2xl font-bold text-slate-900">
            {mode === "allocation" && "箱设定明细作成工具"}
            {mode === "delivery_note" && "受渡伝票作成工具"}
            {mode === "assortment" && "アソート明細作成工具"}
            {mode === "box_label" && "箱贴作成工具"}
          </div>
          <div className="text-slate-500 mt-2 text-sm leading-relaxed">
            {mode === "allocation" && "上传WD系统下载的【配分表】和【明细表】，自动转化为【箱设定明细】和【各店铺明细表】"}
            {mode === "delivery_note" && "根据仓库返回真实分箱明细自动生成【受渡伝票】用于上传WD系统"}
            {mode === "assortment" && "根据仓库返回真实分箱明细自动生成【アソート明細】"}
            {mode === "box_label" && "根据仓库返回真实分箱明细和WD系统的【物流管理表】自动生成【箱贴明细】"}
          </div>
        </div>
        
        <div className="bg-white p-3 rounded-xl border border-slate-200 shadow-sm">
          <div className="flex items-center gap-2 mb-3 px-1 border-b border-slate-100 pb-2">
            <span className="text-lg">🔁</span>
            <span className="font-semibold text-slate-700 text-sm">Switch</span>
          </div>
          <div className="grid grid-cols-2 gap-2">
            <ModeTab 
              mode="allocation" 
              current={mode} 
              onClick={setMode} 
              label="箱设定明细" 
              className="bg-purple-50 text-purple-700 border border-purple-200 hover:bg-purple-100"
            />
            <ModeTab 
              mode="delivery_note" 
              current={mode} 
              onClick={setMode} 
              label="受渡伝票" 
              className="bg-orange-50 text-orange-700 border border-orange-200 hover:bg-orange-100"
            />
            <ModeTab 
              mode="assortment" 
              current={mode} 
              onClick={setMode} 
              label="アソート明細" 
              className="bg-cyan-50 text-cyan-700 border border-cyan-200 hover:bg-cyan-100"
            />
            <ModeTab 
              mode="box_label" 
              current={mode} 
              onClick={setMode} 
              label="箱贴作成" 
              className="bg-slate-50 text-slate-700 border border-slate-200 hover:bg-slate-100"
            />
          </div>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        <div className="lg:col-span-2 space-y-6">
          <Card>
            <div className="flex items-center justify-between gap-4 flex-wrap">
              <div className="text-sm font-semibold text-slate-900">处理摘要</div>
              <div className="flex items-center gap-2 flex-wrap">
                <Badge variant="info">完成 {formatNumber(successStats.count)}</Badge>
                <MetricBadge label="店铺" value={successStats.store_count} />
                <MetricBadge label="箱数" value={successStats.box_count} />
                <MetricBadge label="总枚数" value={successStats.total_qty} />
                {mode === "allocation" || mode === "box_label" ? (
                  <>
                    <MetricBadge label="SKU" value={successStats.sku_count} />
                    <MetricBadge label="PT" value={successStats.pt_count} />
                    {mode === "allocation" && (
                      <span className="inline-flex items-center gap-2">
                        <Badge variant="neutral">JAN</Badge>
                        <span className="text-xs text-slate-600">
                          明细 {formatNumber(successStats.jan_map_count)} / 匹配 {formatNumber(successStats.jan_match_success)} /
                          失败 {formatNumber(successStats.jan_match_fail)}
                        </span>
                      </span>
                    )}
                  </>
                ) : null}
              </div>
            </div>
            <div className="mt-3 text-xs text-slate-500">
              SKU/PT/JAN 显示为“最后一个成功任务”的统计；店铺/箱数/总枚数为“所有成功任务累加”。
            </div>
          </Card>

          <Card>
            <div className="flex flex-col gap-4">
              <div className="flex items-center justify-between gap-4 flex-wrap">
                <div className="text-sm text-slate-600">
                  文件列表：{fileList.length}，完成：{okCount}，处理中：{processingCount}
                </div>
                <div className="flex gap-2">
                  <Button variant="outline" size="sm" onClick={() => setFileList([])} disabled={processingCount > 0}>
                    清空列表
                  </Button>
                  <Button size="sm" onClick={onConvertAll} disabled={processingCount > 0 || fileList.length === 0}>
                    开始批量转换
                  </Button>
                </div>
              </div>

              <div className="border border-dashed border-slate-300 rounded-lg p-4 bg-slate-50 flex flex-col items-center justify-center min-h-[120px] gap-3">
                <input
                  ref={fileInputRef}
                  type="file"
                  className="hidden"
                  accept=".xls,.xlsx"
                  multiple
                  onChange={(e) => {
                    if (e.target.files) onPickFiles(e.target.files);
                    e.target.value = "";
                  }}
                />
                
                <div className="text-center">
                  {mode === "box_label" ? (
                    <>
                      <div className="text-slate-900 font-medium mb-1">请上传工厂返还的箱设定明细表</div>
                      <div className="text-slate-500 text-xs">支持 .xlsx 格式，需包含 PT-xx 页</div>
                    </>
                  ) : (
                    <>
                      <div className="text-slate-900 font-medium mb-1">请选择源文件</div>
                      <div className="text-slate-500 text-xs">支持 .xls, .xlsx 格式</div>
                    </>
                  )}
                </div>

                <Button variant="outline" size="sm" onClick={() => fileInputRef.current?.click()}>
                  选择文件（支持多选）
                </Button>
              </div>

              {mode !== "box_label" && (
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div className="space-y-2">
                    <div className="text-sm font-medium">模板选择</div>
                    <select
                      className="w-full border border-slate-300 rounded-md px-2 py-1 text-xs bg-white h-8"
                      value={selectedLibTemplate}
                      onChange={(e) => {
                        setSelectedLibTemplate(e.target.value);
                        setTemplateFile(null);
                        if (templateInputRef.current) templateInputRef.current.value = "";
                      }}
                    >
                      <option value="">使用系统默认模板</option>
                      {templates.map((t) => (
                        <option key={t.name} value={t.name}>
                          {t.name}
                        </option>
                      ))}
                    </select>
                    <div className="flex items-center gap-2">
                      <input
                        ref={templateInputRef}
                        type="file"
                        className="hidden"
                        accept=".xls,.xlsx"
                        onChange={(e) => {
                          const f = e.target.files?.[0] ?? null;
                          setTemplateFile(f);
                          if (f) setSelectedLibTemplate("");
                        }}
                      />
                      <Button variant="outline" size="sm" onClick={() => templateInputRef.current?.click()}>
                        上传自定义模板
                      </Button>
                      <div className="text-xs text-slate-500 truncate">{templateFile ? templateFile.name : "未选择"}</div>
                    </div>
                  </div>

                  {mode === "allocation" ? (
                    <div className="space-y-2">
                      <div className="text-sm font-medium">
                        明细表 <span className="text-red-600 text-xs">必填</span>
                      </div>
                      <input
                        ref={detailInputRef}
                        type="file"
                        className="hidden"
                        accept=".xls,.xlsx"
                        onChange={(e) => {
                          const f = e.target.files?.[0] ?? null;
                          setDetailFile(f);
                        }}
                      />
                      <div className="flex items-center gap-2">
                        <Button variant="outline" size="sm" onClick={() => detailInputRef.current?.click()}>
                          选择明细表
                        </Button>
                        <div className="text-xs text-slate-500 truncate flex-1">{detailFile ? detailFile.name : "未选择"}</div>
                        {detailFile ? (
                          <Button variant="danger" size="xs" onClick={() => setDetailFile(null)}>
                            清除
                          </Button>
                        ) : null}
                      </div>
                    </div>
                  ) : mode === "assortment" ? (
                    <div className="space-y-2">
                      <div className="text-sm font-medium">本年度周数（Week Number）</div>
                      <input
                        className="w-full border border-slate-300 rounded-md px-2 py-1 text-xs bg-white h-8"
                        value={weekNum}
                        onChange={(e) => setWeekNum(e.target.value)}
                        placeholder="例如：42 或 42W"
                      />
                    </div>
                  ) : (
                    <div />
                  )}
                </div>
              )}

              <div className="space-y-2">
                <div className="text-sm font-medium">待处理文件</div>
                <div className="space-y-2">
                  {fileList.length === 0 ? (
                    <div className="text-sm text-slate-500">暂无文件</div>
                  ) : (
                    fileList.map((item, idx) => (
                      <div key={`${item.file.name}-${idx}`} className="flex items-center justify-between gap-3 bg-white border border-slate-200 rounded-md px-3 py-2">
                        <div className="min-w-0">
                          <div className="flex items-center gap-2 min-w-0">
                            <div className="text-sm font-medium truncate">{item.file.name}</div>
                            <StatusBadge status={item.status} />
                          </div>
                          <div className="text-xs text-slate-500">
                            {item.status === "idle"
                              ? "等待中"
                              : item.status === "processing"
                                ? "处理中"
                                : item.status === "success"
                                  ? "完成"
                                  : `失败：${item.error ?? ""}`}
                          </div>
                          {item.status === "success" && item.result?.stats ? (
                            <div className="mt-2 flex items-center gap-2 flex-wrap">
                              <MetricBadge label="店铺" value={item.result.stats.store_count} />
                              <MetricBadge label="箱数" value={item.result.stats.box_count} />
                              <MetricBadge label="SKU" value={item.result.stats.sku_count} />
                              <MetricBadge label="PT" value={item.result.stats.pt_count} />
                              <MetricBadge label="总枚数" value={item.result.stats.total_qty} />
                            </div>
                          ) : null}
                        </div>
                        <div className="flex items-center gap-2">
                          {item.status === "success" && item.result?.download_url ? (
                            <a className="no-underline" href={item.result.download_url} download>
                              <Badge variant="warning">下载结果</Badge>
                            </a>
                          ) : null}
                          {item.status === "idle" ? (
                            <Button
                              variant="danger"
                              size="xs"
                              onClick={() => setFileList((prev) => prev.filter((_, i) => i !== idx))}
                            >
                              移除
                            </Button>
                          ) : null}
                        </div>
                      </div>
                    ))
                  )}
                </div>
              </div>
            </div>
          </Card>
        </div>

        <div className="space-y-6">
          <Card>
            <div className="text-sm font-semibold uppercase tracking-wider">处理日志</div>
            <div className="mt-3 h-[520px] overflow-y-auto rounded-lg border border-slate-200 bg-slate-50 p-3 text-sm space-y-2">
              {logs.length === 0 ? <div className="text-slate-400">等待任务开始...</div> : null}
              {logs.map((l, i) => {
                const c = classifyLog(l);
                const color =
                  c.variant === "error"
                    ? "text-rose-700"
                    : c.variant === "warning"
                      ? "text-amber-800"
                      : c.variant === "success"
                        ? "text-emerald-700"
                        : c.variant === "info"
                          ? "text-blue-700"
                          : "text-slate-700";
                const badgeVariant =
                  c.variant === "error"
                    ? "error"
                    : c.variant === "warning"
                      ? "warning"
                      : c.variant === "success"
                        ? "success"
                        : c.variant === "info"
                          ? "info"
                          : "neutral";
                return (
                  <div key={`${l}-${i}`} className="break-words flex items-start gap-2">
                    <Badge variant={badgeVariant}>{c.variant.toUpperCase()}</Badge>
                    <div className="min-w-0">
                      <div className={[color, "leading-5"].join(" ")}>{c.text}</div>
                      {c.time ? <div className="text-[10px] text-slate-400 mt-0.5">{c.time}</div> : null}
                    </div>
                  </div>
                );
              })}
            </div>
          </Card>
        </div>
      </div>
    </div>
  );
}
