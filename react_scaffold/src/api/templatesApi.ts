import { httpClient } from "./httpClient";
import type { TemplateInfo } from "../types";

export const templatesApi = {
  async list(): Promise<TemplateInfo[]> {
    const res = await httpClient.getJson<{ templates?: TemplateInfo[] }>("/api/templates");
    return res.templates ?? [];
  },
  async upload(file: File): Promise<void> {
    const formData = new FormData();
    formData.append("file", file);
    await httpClient.postForm("/api/templates", formData);
  },
  async remove(filename: string): Promise<void> {
    await httpClient.delete(`/api/templates/${encodeURIComponent(filename)}`);
  }
};

