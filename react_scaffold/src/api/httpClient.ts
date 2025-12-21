export type HttpClientOptions = {
  baseUrl?: string;
};

const defaultBaseUrl = (import.meta.env.VITE_API_BASE_URL as string | undefined) ?? "";

export class HttpClient {
  private readonly baseUrl: string;

  constructor(options?: HttpClientOptions) {
    this.baseUrl = options?.baseUrl ?? defaultBaseUrl;
  }

  async getJson<T>(path: string): Promise<T> {
    const res = await fetch(this.baseUrl + path, { method: "GET" });
    if (!res.ok) {
      throw new Error(await res.text());
    }
    return (await res.json()) as T;
  }

  async delete(path: string): Promise<void> {
    const res = await fetch(this.baseUrl + path, { method: "DELETE" });
    if (!res.ok) {
      throw new Error(await res.text());
    }
  }

  async postForm<T>(path: string, formData: FormData): Promise<T> {
    const res = await fetch(this.baseUrl + path, { method: "POST", body: formData });
    if (!res.ok) {
      throw new Error(await res.text());
    }
    return (await res.json()) as T;
  }

  async patchJson<T>(path: string, data: any): Promise<T> {
    const res = await fetch(this.baseUrl + path, {
      method: "PATCH",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(data),
    });
    if (!res.ok) {
      throw new Error(await res.text());
    }
    return (await res.json()) as T;
  }

  async postJson<T>(path: string, data: any): Promise<T> {
    const res = await fetch(this.baseUrl + path, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(data),
    });
    if (!res.ok) {
      throw new Error(await res.text());
    }
    return (await res.json()) as T;
  }
}

export const httpClient = new HttpClient();

