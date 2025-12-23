import type { PropsWithChildren } from "react";

export type BadgeVariant = "neutral" | "info" | "success" | "warning" | "error" | "purple" | "orange" | "cyan";

export function Badge({ children, variant = "neutral" }: PropsWithChildren<{ variant?: BadgeVariant }>) {
  const styles: Record<BadgeVariant, string> = {
    neutral: "bg-slate-100 text-slate-700 border-slate-200",
    info: "bg-blue-50 text-blue-800 border-blue-200",
    success: "bg-emerald-50 text-emerald-800 border-emerald-200",
    warning: "bg-amber-50 text-amber-900 border-amber-200",
    error: "bg-rose-50 text-rose-800 border-rose-200",
    purple: "bg-purple-50 text-purple-800 border-purple-200",
    orange: "bg-orange-50 text-orange-800 border-orange-200",
    cyan: "bg-cyan-50 text-cyan-800 border-cyan-200"
  };

  return (
    <span className={["inline-flex items-center rounded-full border px-2 py-0.5 text-xs font-semibold", styles[variant]].join(" ")}>
      {children}
    </span>
  );
}

