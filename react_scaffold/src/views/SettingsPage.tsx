import { useEffect, useState } from "react";
import { Button } from "../ui/Button";
import { Card } from "../ui/Card";
import { useToast } from "../ui/Toast";
import { settingsApi, SettingsMap } from "../api/settingsApi";

export function SettingsPage() {
  const [settings, setSettings] = useState<SettingsMap>({});
  const [loading, setLoading] = useState(false);
  const [saving, setSaving] = useState(false);
  const { toast } = useToast();

  // Local state for edits
  const [prefix, setPrefix] = useState("42");

  useEffect(() => {
    loadSettings();
  }, []);

  async function loadSettings() {
    setLoading(true);
    try {
      const data = await settingsApi.getAll();
      setSettings(data);
      if (data["delivery_note_prefix"]) {
        setPrefix(data["delivery_note_prefix"].value);
      }
    } catch (e) {
      toast({ type: "error", title: "加载失败", message: String(e) });
    } finally {
      setLoading(false);
    }
  }

  async function onSave() {
    setSaving(true);
    try {
      // Update delivery_note_prefix
      await settingsApi.update("delivery_note_prefix", prefix, "受渡伝票NO前缀");
      
      toast({ type: "success", title: "保存成功", message: "系统设置已更新" });
      loadSettings(); // Reload to be sure
    } catch (e) {
      toast({ type: "error", title: "保存失败", message: String(e) });
    } finally {
      setSaving(false);
    }
  }

  if (loading && Object.keys(settings).length === 0) {
    return <div className="p-8 text-slate-500">加载中...</div>;
  }

  return (
    <div className="space-y-6 max-w-4xl">
      <div>
        <h1 className="text-2xl font-bold text-slate-900">系统设置</h1>
        <p className="text-slate-500 mt-2 text-sm">
          管理系统的全局参数配置。
        </p>
      </div>

      <Card>
        <div className="space-y-6">
          <div>
            <h3 className="text-lg font-medium text-slate-900 mb-4">受渡伝票设置</h3>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              <div className="space-y-2">
                <label className="block text-sm font-medium text-slate-700">
                  受渡伝票 NO 前缀
                </label>
                <div className="flex items-center gap-2">
                  <input
                    type="text"
                    className="flex-1 border border-slate-300 rounded-md px-3 py-2 text-sm focus:ring-purple-500 focus:border-purple-500"
                    value={prefix}
                    onChange={(e) => setPrefix(e.target.value)}
                    placeholder="例如: 42"
                    maxLength={4}
                  />
                  <div className="text-sm text-slate-500">
                    + 4位序号 (例: {prefix}0001)
                  </div>
                </div>
                <p className="text-xs text-slate-500">
                  用于生成受渡伝票的编号前缀。原本为 "81"，现默认为 "42"。
                </p>
              </div>
            </div>
          </div>

          <div className="pt-4 border-t border-slate-100 flex justify-end">
            <Button onClick={onSave} disabled={saving}>
              {saving ? "保存中..." : "保存设置"}
            </Button>
          </div>
        </div>
      </Card>
    </div>
  );
}
