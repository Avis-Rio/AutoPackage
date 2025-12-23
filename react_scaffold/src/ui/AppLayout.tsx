import { NavLink, Outlet } from "react-router-dom";
import type { ReactNode } from "react";

function NavItem({ to, label, icon }: { to: string; label: string; icon: ReactNode }) {
  return (
    <NavLink
      to={to}
      className={({ isActive }) =>
        [
          "w-full flex items-center gap-2.5 px-3 py-2 rounded-md text-sm font-medium transition-colors",
          isActive ? "bg-slate-100 text-slate-900" : "text-slate-600 hover:bg-slate-50"
        ].join(" ")
      }
      end={to === "/"}
    >
      <span className="inline-flex h-5 w-5 items-center justify-center text-slate-500">{icon}</span>
      {label}
    </NavLink>
  );
}

export function AppLayout() {
  return (
    <div className="min-h-screen flex">
      <aside className="w-48 bg-white border-r border-slate-200 hidden md:flex flex-col flex-shrink-0">
        <div className="p-4 border-b border-slate-100">
          <div className="flex items-center gap-2">
            <span className="inline-flex h-7 w-7 items-center justify-center rounded-lg bg-slate-900 text-white">
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" aria-hidden="true">
                <path
                  d="M21 8.5V16c0 .7-.4 1.4-1 1.7l-7 4c-.6.3-1.4.3-2 0l-7-4C3.4 17.4 3 16.7 3 16V8.5"
                  stroke="currentColor"
                  strokeWidth="2"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                />
                <path
                  d="M12 22V12"
                  stroke="currentColor"
                  strokeWidth="2"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                />
                <path
                  d="M21 8.5 12 13 3 8.5 12 4 21 8.5Z"
                  stroke="currentColor"
                  strokeWidth="2"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                />
              </svg>
            </span>
            <div className="text-lg font-bold tracking-tight">AutoPackage</div>
          </div>
          <div className="text-xs text-slate-400 mt-1">V2 Web</div>
        </div>
        <nav className="flex-1 p-3 space-y-1">
          <NavItem
            to="/"
            label="工作台"
            icon={
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" aria-hidden="true">
                <path d="M3 13h8V3H3v10Z" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                <path d="M13 21h8V11h-8v10Z" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                <path d="M3 21h8v-6H3v6Z" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                <path d="M13 9h8V3h-8v6Z" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
              </svg>
            }
          />
          <NavItem
            to="/templates"
            label="模板管理"
            icon={
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" aria-hidden="true">
                <path
                  d="M14 2H6a2 2 0 0 0-2 2v16c0 1.1.9 2 2 2h12a2 2 0 0 0 2-2V8l-6-6Z"
                  stroke="currentColor"
                  strokeWidth="2"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                />
                <path
                  d="M14 2v6h6"
                  stroke="currentColor"
                  strokeWidth="2"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                />
                <path
                  d="M8 13h8"
                  stroke="currentColor"
                  strokeWidth="2"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                />
                <path
                  d="M8 17h8"
                  stroke="currentColor"
                  strokeWidth="2"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                />
              </svg>
            }
          />
          <NavItem
            to="/history"
            label="历史记录"
            icon={
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" aria-hidden="true">
                <path
                  d="M12 22c5.523 0 10-4.477 10-10S17.523 2 12 2 2 6.477 2 12s4.477 10 10 10Z"
                  stroke="currentColor"
                  strokeWidth="2"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                />
                <path
                  d="M12 6v6l4 2"
                  stroke="currentColor"
                  strokeWidth="2"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                />
              </svg>
            }
          />
        </nav>
      </aside>
      <main className="flex-1 p-4 md:p-6 overflow-y-auto">
        <Outlet />
      </main>
    </div>
  );
}
