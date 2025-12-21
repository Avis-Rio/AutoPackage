import type { ButtonHTMLAttributes, PropsWithChildren } from "react";

export type ButtonProps = PropsWithChildren<
  ButtonHTMLAttributes<HTMLButtonElement> & {
    variant?: "primary" | "outline" | "ghost" | "danger" | "warning" | "secondary";
    size?: "xs" | "sm" | "md";
  }
>;

export function Button({ children, className, variant = "primary", size = "md", ...props }: ButtonProps) {
  const base =
    "rounded-lg font-medium transition-colors duration-150 inline-flex items-center justify-center gap-2 disabled:opacity-50 disabled:cursor-not-allowed";

  const sizes = {
    xs: "h-6 px-2 text-xs",
    sm: "h-8 px-3 text-xs",
    md: "h-10 px-4 text-sm"
  };

  const variants: Record<NonNullable<ButtonProps["variant"]>, string> = {
    primary: "bg-blue-600 text-white hover:bg-blue-700",
    outline: "border border-slate-300 bg-white text-slate-700 hover:bg-slate-50",
    ghost: "text-slate-700 hover:bg-slate-100",
    danger: "bg-red-50 text-red-600 border border-red-200 hover:bg-red-100",
    warning: "bg-amber-50 text-amber-700 border border-amber-200 hover:bg-amber-100",
    secondary: "bg-slate-50 text-slate-700 border border-slate-300 hover:bg-slate-100"
  };

  return (
    <button className={[base, sizes[size], variants[variant], className].filter(Boolean).join(" ")} {...props}>
      {children}
    </button>
  );
}
