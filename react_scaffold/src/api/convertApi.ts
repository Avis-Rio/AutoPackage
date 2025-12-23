import { httpClient } from "./httpClient";
import type { ConvertMode, ConvertResponse } from "../types";

export type ConvertRequest = {
  file: File;
  mode: ConvertMode;
  templateFile?: File | null;
  templateName?: string | null;
  detailFile?: File | null;
  weekNum?: string | null;
  startNo?: string | null;
  isHanger?: boolean;
};

export const convertApi = {
  async convert(req: ConvertRequest): Promise<ConvertResponse> {
    const formData = new FormData();
    formData.append("file", req.file);
    formData.append("mode", req.mode);
    if (req.templateFile) formData.append("template", req.templateFile);
    if (req.templateName) formData.append("template_name", req.templateName);
    if (req.detailFile) formData.append("detail_file", req.detailFile);
    if (req.weekNum) formData.append("week_num", req.weekNum);
    if (req.startNo) formData.append("start_no", req.startNo);
    if (req.isHanger !== undefined) formData.append("is_hanger", String(req.isHanger));
    return await httpClient.postForm("/api/convert", formData);
  },

  async generateLabelsFromFile(file: File): Promise<ConvertResponse> {
    const formData = new FormData();
    formData.append("file", file);
    return await httpClient.postForm("/api/generate-labels-from-file", formData);
  }
};

