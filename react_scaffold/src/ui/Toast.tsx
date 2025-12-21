import React, { createContext, useContext, useState, useCallback } from "react";

export type ToastType = "success" | "error" | "info" | "warning";

export interface Toast {
    id: string;
    type: ToastType;
    title?: string;
    message: string;
    duration?: number;
}

interface ToastContextValue {
    toast: (props: Omit<Toast, "id">) => void;
}

const ToastContext = createContext<ToastContextValue | null>(null);

export function useToast() {
    const context = useContext(ToastContext);
    if (!context) {
        throw new Error("useToast must be used within a ToastProvider");
    }
    return context;
}

export const ToastProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
    const [toasts, setToasts] = useState<Toast[]>([]);

    const removeToast = useCallback((id: string) => {
        setToasts((prev) => prev.filter((t) => t.id !== id));
    }, []);

    const toast = useCallback(({ type, title, message, duration = 3000 }: Omit<Toast, "id">) => {
        const id = Math.random().toString(36).substring(2, 9);
        setToasts((prev) => [...prev, { id, type, title, message, duration }]);
        if (duration > 0) {
            setTimeout(() => {
                removeToast(id);
            }, duration);
        }
    }, [removeToast]);

    return (
        <ToastContext.Provider value={{ toast }}>
            {children}
            <div className="fixed top-4 right-4 z-50 flex flex-col gap-2 w-full max-w-sm pointer-events-none">
                {toasts.map((t) => (
                    <div
                        key={t.id}
                        className={`pointer-events-auto flex w-full flex-col gap-1 rounded-lg border bg-white p-4 shadow-lg transition-all animate-in slide-in-from-right-full ${t.type === "error"
                                ? "border-red-200 bg-red-50 text-red-900"
                                : t.type === "success"
                                    ? "border-green-200 bg-green-50 text-green-900"
                                    : t.type === "warning"
                                        ? "border-amber-200 bg-amber-50 text-amber-900"
                                        : "border-slate-200 text-slate-900"
                            }`}
                    >
                        {t.title && <div className="text-sm font-semibold">{t.title}</div>}
                        <div className="text-sm opacity-90">{t.message}</div>
                    </div>
                ))}
            </div>
        </ToastContext.Provider>
    );
};
