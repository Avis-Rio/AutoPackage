import type { PropsWithChildren } from "react";

export function Card({ children, className }: PropsWithChildren<{ className?: string }>) {
  return (
    <div className={["rounded-xl bg-white border border-slate-200 shadow-sm p-6", className].filter(Boolean).join(" ")}>
      {children}
    </div>
  );
}

