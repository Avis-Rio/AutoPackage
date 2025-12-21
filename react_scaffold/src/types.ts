export type ConvertMode = "allocation" | "delivery_note" | "assortment";

export type ConvertResponse = {
  status: "success" | "error";
  message: string;
  download_url?: string;
  stats?: {
    items_processed?: number;
    generated_file?: string;
    sku_count?: number;
    pt_count?: number;
    store_count?: number;
    box_count?: number;
    total_qty?: number;
    jan_map_count?: number;
    jan_match_success?: number;
    jan_match_fail?: number;
  };
  logs?: string[];
};

export type TemplateInfo = {
  name: string;
  size: number;
  modified: string;
};

export type HistoryRecord = {
  id: number;
  created_at: string;
  original_filename: string;
  mode: ConvertMode;
  status: "processing" | "success" | "failed";
  output_filename?: string;
  file_path?: string;
  stats?: ConvertResponse["stats"];
  error_message?: string;
  note?: string;
};

export type HistoryResponse = {
  items: HistoryRecord[];
  total: number;
};

export type PreviewData = {
  columns: string[];
  data: any[][];
};
