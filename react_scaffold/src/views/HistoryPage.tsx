import React, { useEffect, useState } from "react";
import { historyApi } from "../api/historyApi";
import { HistoryRecord, PreviewData } from "../types";
import { Card } from "../ui/Card";
import { Badge } from "../ui/Badge";
import { Button } from "../ui/Button";
import { useToast } from "../ui/Toast";

// Simple Modal Component
const Modal = ({ isOpen, onClose, title, children, maxWidth = "max-w-md" }: any) => {
    if (!isOpen) return null;
    return (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black bg-opacity-50 p-4">
            <div className={`bg-white rounded-lg shadow-xl w-full ${maxWidth} p-6 max-h-[90vh] flex flex-col`}>
                <div className="flex justify-between items-center mb-4 flex-shrink-0">
                    <h3 className="text-lg font-bold">{title}</h3>
                    <button onClick={onClose} className="text-gray-500 hover:text-gray-700">✕</button>
                </div>
                <div className="overflow-auto flex-grow">
                    {children}
                </div>
            </div>
        </div>
    );
};

export const HistoryPage: React.FC = () => {
    const [history, setHistory] = useState<HistoryRecord[]>([]);
    const [loading, setLoading] = useState(true);
    const [editingNoteId, setEditingNoteId] = useState<number | null>(null);
    const [noteInput, setNoteInput] = useState("");
    const { toast } = useToast();

    // Pagination state
    const [currentPage, setCurrentPage] = useState(1);
    const [totalCount, setTotalCount] = useState(0);
    const pageSize = 20;

    // Batch selection state
    const [selectedIds, setSelectedIds] = useState<number[]>([]);

    // Rerun state
    const [rerunModalOpen, setRerunModalOpen] = useState(false);
    const [selectedRecord, setSelectedRecord] = useState<HistoryRecord | null>(null);
    const [rerunMode, setRerunMode] = useState("allocation");
    const [rerunWeek, setRerunWeek] = useState("");
    const [rerunLoading, setRerunLoading] = useState(false);

    // Preview state
    const [previewModalOpen, setPreviewModalOpen] = useState(false);
    const [previewData, setPreviewData] = useState<PreviewData | null>(null);
    const [previewLoading, setPreviewLoading] = useState(false);

    const fetchHistory = async () => {
        try {
            setLoading(true);
            const offset = (currentPage - 1) * pageSize;
            const data = await historyApi.getHistory(pageSize, offset);
            setHistory(data.items);
            setTotalCount(data.total);
        } catch (error) {
            console.error("Failed to fetch history", error);
            toast({ type: "error", title: "加载失败", message: "无法获取历史记录" });
        } finally {
            setLoading(false);
        }
    };

    useEffect(() => {
        fetchHistory();
    }, [currentPage]);

    const handleDelete = async (id: number) => {
        if (!confirm("确定要删除这条记录吗？")) return;
        try {
            await historyApi.deleteHistory(id);
            toast({ type: "success", title: "删除成功", message: "记录已删除" });
            fetchHistory();
        } catch (error) {
            console.error("Failed to delete history", error);
            toast({ type: "error", title: "删除失败", message: String(error) });
        }
    };

    const handleBatchDelete = async () => {
        if (selectedIds.length === 0) return;
        if (!confirm(`确定要删除选中的 ${selectedIds.length} 条记录吗？`)) return;
        try {
            await historyApi.deleteBatch(selectedIds);
            toast({ type: "success", title: "批量删除成功", message: `已删除 ${selectedIds.length} 条记录` });
            setSelectedIds([]);
            fetchHistory();
        } catch (error) {
            console.error("Batch delete failed", error);
            toast({ type: "error", title: "批量删除失败", message: String(error) });
        }
    };

    const toggleSelect = (id: number) => {
        setSelectedIds(prev =>
            prev.includes(id) ? prev.filter(i => i !== id) : [...prev, id]
        );
    };

    const toggleSelectAll = () => {
        if (selectedIds.length === history.length) {
            setSelectedIds([]);
        } else {
            setSelectedIds(history.map(h => h.id));
        }
    };

    const handleDownload = (filename: string) => {
        window.open(`/api/download/${filename}`, "_blank");
    };

    const handlePreview = async (id: number) => {
        try {
            setPreviewLoading(true);
            setPreviewModalOpen(true);
            const data = await historyApi.preview(id);
            setPreviewData(data);
        } catch (error) {
            console.error("Preview failed", error);
            toast({ type: "error", title: "预览失败", message: "无法读取文件内容" });
            setPreviewModalOpen(false);
        } finally {
            setPreviewLoading(false);
        }
    };

    const startEditNote = (record: HistoryRecord) => {
        setEditingNoteId(record.id);
        setNoteInput(record.note || "");
    };

    const saveNote = async (id: number) => {
        try {
            await historyApi.updateNote(id, noteInput);
            setEditingNoteId(null);
            fetchHistory();
            toast({ type: "success", title: "备注已更新", message: "备注保存成功" });
        } catch (error) {
            console.error("Failed to update note", error);
            toast({ type: "error", title: "更新失败", message: "无法保存备注" });
        }
    };

    const openRerunModal = (record: HistoryRecord) => {
        setSelectedRecord(record);
        setRerunMode(record.mode);
        setRerunWeek("");
        setRerunModalOpen(true);
    };

    const handleRerun = async () => {
        if (!selectedRecord) return;
        try {
            setRerunLoading(true);
            await historyApi.rerun(selectedRecord.id, rerunMode, undefined, rerunWeek || undefined);
            setRerunModalOpen(false);
            fetchHistory();
            toast({ type: "success", title: "任务已提交", message: "转换任务已在后台开始" });
        } catch (error) {
            console.error("Rerun failed", error);
            toast({ type: "error", title: "重跑失败", message: String(error) });
        } finally {
            setRerunLoading(false);
        }
    };

    const getModeBadge = (mode: string) => {
        switch (mode) {
            case "allocation":
                return <Badge variant="purple">配分表</Badge>;
            case "delivery_note":
                return <Badge variant="orange">受渡伝票</Badge>;
            case "assortment":
                return <Badge variant="cyan">Assortment</Badge>;
            case "box_label":
                return <Badge variant="neutral">箱贴作成</Badge>;
            default:
                return <Badge>{mode}</Badge>;
        }
    };

    const getStatusBadge = (status: string) => {
        switch (status) {
            case "success":
                return <Badge variant="success">成功</Badge>;
            case "failed":
                return <Badge variant="error">失败</Badge>;
            case "processing":
                return <Badge variant="warning">处理中</Badge>;
            default:
                return <Badge>{status}</Badge>;
        }
    };

    const totalPages = Math.ceil(totalCount / pageSize);

    return (
        <div className="p-4 w-full">
            <div className="flex justify-between items-center mb-6">
                <h1 className="text-2xl font-bold text-gray-800">转换历史记录</h1>
                {selectedIds.length > 0 && (
                    <Button variant="danger" size="sm" onClick={handleBatchDelete}>
                        批量删除 ({selectedIds.length})
                    </Button>
                )}
            </div>

            <Card className="overflow-hidden">
                <div className="overflow-x-auto">
                    <table className="min-w-full divide-y divide-gray-200 table-fixed">
                        <thead className="bg-gray-50">
                            <tr>
                                <th className="px-3 py-2 text-left w-10">
                                    <input
                                        type="checkbox"
                                        checked={history.length > 0 && selectedIds.length === history.length}
                                        onChange={toggleSelectAll}
                                        className="rounded border-gray-300 text-blue-600 focus:ring-blue-500"
                                    />
                                </th>
                                <th className="px-3 py-2 text-left text-xs font-medium text-gray-500 uppercase tracking-wider w-12">ID</th>
                                <th className="px-3 py-2 text-left text-xs font-medium text-gray-500 uppercase tracking-wider w-24">时间</th>
                                <th className="px-3 py-2 text-left text-xs font-medium text-gray-500 uppercase tracking-wider w-auto">文件名 / 备注</th>
                                <th className="px-3 py-2 text-left text-xs font-medium text-gray-500 uppercase tracking-wider w-24">模式</th>
                                <th className="px-3 py-2 text-left text-xs font-medium text-gray-500 uppercase tracking-wider w-20">状态</th>
                                <th className="px-3 py-2 text-left text-xs font-medium text-gray-500 uppercase tracking-wider w-32">统计</th>
                                <th className="px-3 py-2 text-right text-xs font-medium text-gray-500 uppercase tracking-wider w-48">操作</th>
                            </tr>
                        </thead>
                        <tbody className="bg-white divide-y divide-gray-200">
                            {loading ? (
                                <tr>
                                    <td colSpan={8} className="px-3 py-4 text-center text-gray-500">
                                        加载中...
                                    </td>
                                </tr>
                            ) : history.length === 0 ? (
                                <tr>
                                    <td colSpan={8} className="px-3 py-4 text-center text-gray-500">
                                        暂无记录
                                    </td>
                                </tr>
                            ) : (
                                history.map((record) => (
                                    <tr key={record.id} className={`hover:bg-gray-50 ${selectedIds.includes(record.id) ? 'bg-blue-50' : ''}`}>
                                        <td className="px-3 py-2">
                                            <input
                                                type="checkbox"
                                                checked={selectedIds.includes(record.id)}
                                                onChange={() => toggleSelect(record.id)}
                                                className="rounded border-gray-300 text-blue-600 focus:ring-blue-500"
                                            />
                                        </td>
                                        <td className="px-3 py-2 whitespace-nowrap text-xs text-gray-500">
                                            {record.id}
                                        </td>
                                        <td className="px-3 py-2 whitespace-nowrap text-xs text-gray-900">
                                            <div className="flex flex-col">
                                                <span>{new Date(record.created_at).toLocaleDateString()}</span>
                                                <span className="text-gray-500">{new Date(record.created_at).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}</span>
                                            </div>
                                        </td>
                                        <td className="px-3 py-2 text-sm text-gray-900 overflow-hidden">
                                            <div className="font-medium truncate" title={record.original_filename}>{record.original_filename}</div>
                                            {record.output_filename && (
                                                <div className="text-xs text-gray-500 truncate" title={record.output_filename}>
                                                    → {record.output_filename}
                                                </div>
                                            )}

                                            <div className="mt-1">
                                                {editingNoteId === record.id ? (
                                                    <div className="flex gap-1">
                                                        <input
                                                            type="text"
                                                            className="text-xs border rounded px-1 py-0.5 w-full"
                                                            value={noteInput}
                                                            onChange={(e) => setNoteInput(e.target.value)}
                                                            onKeyDown={(e) => e.key === 'Enter' && saveNote(record.id)}
                                                            autoFocus
                                                        />
                                                        <button onClick={() => saveNote(record.id)} className="text-xs text-green-600">✓</button>
                                                        <button onClick={() => setEditingNoteId(null)} className="text-xs text-gray-500">✕</button>
                                                    </div>
                                                ) : (
                                                    <div
                                                        className="text-xs text-gray-500 hover:text-gray-700 cursor-pointer flex items-center gap-1 group truncate"
                                                        onClick={() => startEditNote(record)}
                                                        title={record.note || "点击添加备注"}
                                                    >
                                                        <span className={`italic ${!record.note && "opacity-50"}`}>{record.note || "添加备注..."}</span>
                                                        <span className="hidden group-hover:inline opacity-50 ml-1">✎</span>
                                                    </div>
                                                )}
                                            </div>
                                        </td>
                                        <td className="px-3 py-2 whitespace-nowrap text-xs text-gray-500">
                                            {getModeBadge(record.mode)}
                                        </td>
                                        <td className="px-3 py-2 whitespace-nowrap">
                                            {getStatusBadge(record.status)}
                                            {record.error_message && (
                                                <div className="text-xs text-red-500 mt-1 w-20 truncate" title={record.error_message}>
                                                    {record.error_message}
                                                </div>
                                            )}
                                        </td>
                                        <td className="px-3 py-2 whitespace-nowrap text-xs text-gray-500">
                                            {record.stats ? (
                                                <div className="grid grid-cols-2 gap-x-2 gap-y-0.5">
                                                    {record.stats.store_count !== undefined && <span>店: {record.stats.store_count}</span>}
                                                    {record.stats.box_count !== undefined && <span>箱: {record.stats.box_count}</span>}
                                                    {record.stats.total_qty !== undefined && <span className="col-span-2">总: {record.stats.total_qty}</span>}
                                                </div>
                                            ) : (
                                                "-"
                                            )}
                                        </td>
                                        <td className="px-3 py-2 whitespace-nowrap text-right text-xs font-medium">
                                            <div className="flex justify-end gap-1.5 items-center">
                                                {record.status === "success" && record.output_filename && (
                                                    <>
                                                        {!record.output_filename.toLowerCase().endsWith('.zip') && (
                                                            <Button
                                                                size="xs"
                                                                variant="outline"
                                                                onClick={() => handlePreview(record.id)}
                                                                title="预览"
                                                            >
                                                                预览
                                                            </Button>
                                                        )}
                                                        <Button
                                                            size="xs"
                                                            variant="warning"
                                                            onClick={() => handleDownload(record.output_filename!)}
                                                            title="下载"
                                                        >
                                                            下载
                                                        </Button>
                                                    </>
                                                )}

                                                <Button
                                                    size="xs"
                                                    variant="secondary"
                                                    onClick={() => openRerunModal(record)}
                                                    title="再次转换"
                                                >
                                                    重跑
                                                </Button>
                                                <Button
                                                    size="xs"
                                                    variant="danger"
                                                    onClick={() => handleDelete(record.id)}
                                                    title="删除"
                                                >
                                                    删除
                                                </Button>
                                            </div>
                                        </td>
                                    </tr>
                                ))
                            )}
                        </tbody>
                    </table>
                </div>

                {/* Pagination */}
                {!loading && totalCount > 0 && (
                    <div className="bg-white px-3 py-2 flex items-center justify-between border-t border-gray-200 sm:px-4">
                        <div className="flex-1 flex justify-between sm:hidden">
                            <Button
                                onClick={() => setCurrentPage(prev => Math.max(prev - 1, 1))}
                                disabled={currentPage === 1}
                                variant="outline"
                                size="sm"
                            >
                                上一页
                            </Button>
                            <Button
                                onClick={() => setCurrentPage(prev => Math.min(prev + 1, totalPages))}
                                disabled={currentPage === totalPages}
                                variant="outline"
                                size="sm"
                            >
                                下一页
                            </Button>
                        </div>
                        <div className="hidden sm:flex-1 sm:flex sm:items-center sm:justify-between">
                            <div>
                                <p className="text-xs text-gray-600">
                                    显示第 <span className="font-medium">{(currentPage - 1) * pageSize + 1}</span> 到 <span className="font-medium">{Math.min(currentPage * pageSize, totalCount)}</span> 条，共 <span className="font-medium">{totalCount}</span> 条记录
                                </p>
                            </div>
                            <div>
                                <nav className="relative z-0 inline-flex rounded shadow-sm -space-x-px" aria-label="Pagination">
                                    <button
                                        onClick={() => setCurrentPage(prev => Math.max(prev - 1, 1))}
                                        disabled={currentPage === 1}
                                        className="relative inline-flex items-center px-2 py-1 rounded-l border border-gray-300 bg-white text-xs font-medium text-gray-500 hover:bg-gray-50 disabled:opacity-50"
                                    >
                                        <span className="sr-only">Previous</span>
                                        &lt;
                                    </button>
                                    {[...Array(totalPages)].map((_, i) => (
                                        <button
                                            key={i + 1}
                                            onClick={() => setCurrentPage(i + 1)}
                                            className={`relative inline-flex items-center px-3 py-1 border text-xs font-medium ${currentPage === i + 1
                                                ? "z-10 bg-blue-50 border-blue-500 text-blue-600"
                                                : "bg-white border-gray-300 text-gray-500 hover:bg-gray-50"
                                                }`}
                                        >
                                            {i + 1}
                                        </button>
                                    ))}
                                    <button
                                        onClick={() => setCurrentPage(prev => Math.min(prev + 1, totalPages))}
                                        disabled={currentPage === totalPages}
                                        className="relative inline-flex items-center px-2 py-1 rounded-r border border-gray-300 bg-white text-xs font-medium text-gray-500 hover:bg-gray-50 disabled:opacity-50"
                                    >
                                        <span className="sr-only">Next</span>
                                        &gt;
                                    </button>
                                </nav>
                            </div>
                        </div>
                    </div>
                )}
            </Card>

            {/* Rerun Modal */}
            <Modal
                isOpen={rerunModalOpen}
                onClose={() => setRerunModalOpen(false)}
                title="重新转换任务"
            >
                <div className="space-y-4">
                    <div>
                        <label className="block text-sm font-medium text-gray-700 mb-1">源文件</label>
                        <div className="text-sm text-gray-900 bg-gray-50 p-2 rounded truncate">
                            {selectedRecord?.original_filename}
                        </div>
                    </div>

                    <div>
                        <label className="block text-sm font-medium text-gray-700 mb-1">转换模式</label>
                        <select
                            className="w-full border-gray-300 rounded-md shadow-sm focus:border-indigo-500 focus:ring-indigo-500 sm:text-sm p-2 border"
                            value={rerunMode}
                            onChange={(e) => setRerunMode(e.target.value)}
                        >
                            <option value="allocation">配分表转换 (Allocation)</option>
                            <option value="delivery_note">受渡伝票生成 (Delivery Note)</option>
                            <option value="assortment">Assortment 生成</option>
                        </select>
                    </div>

                    {rerunMode === "assortment" && (
                        <div>
                            <label className="block text-sm font-medium text-gray-700 mb-1">周数 (Week Num)</label>
                            <input
                                type="text"
                                className="w-full border-gray-300 rounded-md shadow-sm focus:border-indigo-500 focus:ring-indigo-500 sm:text-sm p-2 border"
                                placeholder="例如: 2501"
                                value={rerunWeek}
                                onChange={(e) => setRerunWeek(e.target.value)}
                            />
                        </div>
                    )}

                    <div className="pt-4 flex justify-end gap-3">
                        <Button variant="outline" onClick={() => setRerunModalOpen(false)}>取消</Button>
                        <Button onClick={handleRerun} disabled={rerunLoading}>
                            {rerunLoading ? "提交中..." : "开始转换"}
                        </Button>
                    </div>
                </div>
            </Modal>

            {/* Preview Modal */}
            <Modal
                isOpen={previewModalOpen}
                onClose={() => {
                    setPreviewModalOpen(false);
                    setPreviewData(null);
                }}
                title="文件预览 (前20行)"
                maxWidth="max-w-5xl"
            >
                {previewLoading ? (
                    <div className="py-12 text-center text-gray-500">正在读取文件内容...</div>
                ) : previewData ? (
                    <div className="overflow-x-auto border rounded">
                        <table className="min-w-full divide-y divide-gray-200">
                            <thead className="bg-gray-50">
                                <tr>
                                    {/* Use first row as header if available, otherwise use columns */}
                                    {(previewData.data.length > 0 ? previewData.data[0] : previewData.columns).map((col: any, i: number) => (
                                        <th key={i} className="px-3 py-2 text-left text-xs font-semibold text-gray-600 border-r last:border-r-0 whitespace-nowrap">
                                            {String(col || "")}
                                        </th>
                                    ))}
                                </tr>
                            </thead>
                            <tbody className="bg-white divide-y divide-gray-200">
                                {/* Skip first row if we used it as header */}
                                {previewData.data.slice(previewData.data.length > 0 ? 1 : 0).map((row, i) => (
                                    <tr key={i} className="hover:bg-gray-50">
                                        {row.map((cell, j) => (
                                            <td key={j} className="px-3 py-1.5 text-xs text-gray-700 border-r last:border-r-0 whitespace-nowrap">
                                                {cell === null ? "" : String(cell)}
                                            </td>
                                        ))}
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>
                ) : (
                    <div className="py-12 text-center text-red-500">预览加载失败</div>
                )}
                <div className="mt-4 flex justify-end">
                    <Button variant="outline" onClick={() => setPreviewModalOpen(false)}>关闭</Button>
                </div>
            </Modal>
        </div>
    );
};

