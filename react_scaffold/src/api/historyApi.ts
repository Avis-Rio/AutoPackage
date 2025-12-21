import { httpClient } from "./httpClient";
import { HistoryRecord, HistoryResponse, PreviewData } from "../types";

export const historyApi = {
    getHistory: async (limit: number = 20, offset: number = 0) => {
        const params = new URLSearchParams({
            limit: limit.toString(),
            offset: offset.toString(),
        });
        // The backend now returns { items: [], total: number }
        const data = await httpClient.getJson<HistoryResponse>(`/api/history?${params.toString()}`);
        return data;
    },

    deleteHistory: async (id: number) => {
        await httpClient.delete(`/api/history/${id}`);
        return { status: "success", message: "Record deleted" };
    },

    deleteBatch: async (ids: number[]) => {
        await httpClient.postJson(`/api/history/delete_batch`, { ids });
        return { status: "success", message: "Records deleted" };
    },

    updateNote: async (id: number, note: string) => {
        await httpClient.patchJson(`/api/history/${id}`, { note });
    },

    rerun: async (id: number, mode: string, templateName?: string, weekNum?: string) => {
        const formData = new FormData();
        formData.append("mode", mode);
        if (templateName) formData.append("template_name", templateName);
        if (weekNum) formData.append("week_num", weekNum);

        const res = await httpClient.postForm<{
            status: string;
            message: string;
            download_url?: string;
            stats?: any;
        }>(`/api/history/${id}/rerun`, formData);
        return res;
    },

    preview: async (id: number) => {
        const data = await httpClient.getJson<PreviewData>(`/api/history/${id}/preview`);
        return data;
    }
};
