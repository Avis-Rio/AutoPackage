import { httpClient } from "./httpClient";

export type SystemSetting = {
  value: string;
  description?: string;
};

export type SettingsMap = Record<string, SystemSetting>;

export const settingsApi = {
  getAll: async (): Promise<SettingsMap> => {
    return await httpClient.getJson<SettingsMap>("/api/settings");
  },

  update: async (key: string, value: string, description?: string): Promise<{ status: string; setting: { key: string; value: string } }> => {
    return await httpClient.postJson<{ status: string; setting: { key: string; value: string } }>(
      `/api/settings/${key}`,
      { value, description }
    );
  },
};
