import { useEffect, useRef, useState } from "react";
import { Button } from "../ui/Button";
import { Card } from "../ui/Card";
import { Badge } from "../ui/Badge";
import type { TemplateInfo } from "../types";
import { templatesApi } from "../api/templatesApi";
import { useToast } from "../ui/Toast";

export function TemplatesPage() {
  const [items, setItems] = useState<TemplateInfo[]>([]);
  const [status, setStatus] = useState<"idle" | "loading" | "error">("idle");
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const { toast } = useToast();

  const systemTemplates = [
    "①箱设定_模板（配分表用）.xlsx",
    "②アソート明細_模板.xlsx",
    "③受渡伝票_模板（上传系统资料）.xls",
    "③受渡伝票_模板（上传系统资料）.xlsx",
    "③受渡伝票_模板（上传系统资料） .xlsx",
    "④各店铺明细_模板.xlsx",
    "template.xlsx",
    "template.xls"
  ];

  async function refresh() {
    setStatus("loading");
    try {
      const list = await templatesApi.list();
      setItems(list);
      setStatus("idle");
    } catch {
      setStatus("error");
      toast({ type: "error", title: "加载失败", message: "无法获取模板列表" });
    }
  }

  useEffect(() => {
    refresh();
  }, []);

  const isSystemTemplate = (name: string) => {
    // Normalize both to NFC to ensure matching works
    const normalizedName = name.normalize('NFC');
    return systemTemplates.some(t => t.normalize('NFC') === normalizedName);
  };

  return (
    <div className="space-y-6">
      <div className="flex items-center justify-between gap-4 flex-wrap">
        <div>
          <div className="text-2xl font-bold">模板管理</div>
          <div className="text-sm text-slate-500 mt-1">
            <span className="inline-flex items-center gap-2">
              <Badge variant="info">API</Badge>
              <span>GET/POST/DELETE /api/templates</span>
            </span>
          </div>
        </div>
        <div className="flex items-center gap-2">
          <input
            ref={fileInputRef}
            type="file"
            className="hidden"
            accept=".xls,.xlsx"
            onChange={async (e) => {
              const f = e.target.files?.[0];
              e.target.value = "";
              if (!f) return;
              try {
                await templatesApi.upload(f);
                toast({ type: "success", title: "上传成功", message: `模板 ${f.name} 已上传` });
                await refresh();
              } catch (e) {
                toast({ type: "error", title: "上传失败", message: String(e) });
              }
            }}
          />
          <Button onClick={() => fileInputRef.current?.click()}>上传新模板</Button>
          <Button variant="outline" onClick={refresh}>
            刷新
          </Button>
        </div>
      </div>

      <Card>
        {status === "loading" ? <div className="text-slate-500">加载中...</div> : null}
        {status === "error" ? <div className="text-red-600">加载失败</div> : null}
        {status === "idle" && items.length === 0 ? <div className="text-slate-500">暂无模板</div> : null}

        {items.length > 0 ? (
          <div className="overflow-x-auto">
            <table className="min-w-full text-sm">
              <thead className="text-left text-slate-600">
                <tr className="border-b border-slate-200">
                  <th className="py-2 pr-4">文件名</th>
                  <th className="py-2 pr-4">大小</th>
                  <th className="py-2 pr-4">修改时间</th>
                  <th className="py-2 text-right">操作</th>
                </tr>
              </thead>
              <tbody>
                {items.map((t) => {
                  const isSystem = isSystemTemplate(t.name);
                  return (
                    <tr key={t.name} className="border-b border-slate-100">
                      <td className="py-2 pr-4 font-medium text-slate-800">
                        <div className="flex items-center gap-2">
                          {t.name}
                          {isSystem && (
                            <span className="inline-flex items-center rounded-md bg-slate-100 px-2 py-1 text-xs font-medium text-slate-600 ring-1 ring-inset ring-slate-500/10">
                              Default
                            </span>
                          )}
                        </div>
                      </td>
                      <td className="py-2 pr-4 text-slate-600">{(t.size / 1024).toFixed(1)} KB</td>
                      <td className="py-2 pr-4 text-slate-600">{t.modified}</td>
                      <td className="py-2 text-right">
                        <button
                          className={`inline-flex items-center rounded-full border px-2 py-0.5 text-xs font-semibold transition-colors ${isSystem
                            ? "border-slate-200 bg-slate-50 text-slate-400 cursor-not-allowed"
                            : "border-rose-200 bg-rose-50 text-rose-800 hover:bg-rose-100"
                            }`}
                          onClick={async () => {
                            if (isSystem) {
                              toast({ type: "warning", title: "操作禁止", message: "系统默认模板不可删除" });
                              return;
                            }
                            if (!window.confirm(`确定删除模板：${t.name} ？`)) return;
                            try {
                              await templatesApi.remove(t.name);
                              toast({ type: "success", title: "删除成功", message: `模板 ${t.name} 已删除` });
                              await refresh();
                            } catch (e) {
                              toast({ type: "error", title: "删除失败", message: String(e) });
                            }
                          }}
                          type="button"
                          disabled={isSystem}
                        >
                          删除
                        </button>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        ) : null}
      </Card>
    </div>
  );
}
