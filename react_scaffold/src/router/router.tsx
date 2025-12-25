import { createBrowserRouter } from "react-router-dom";
import { AppLayout } from "../ui/AppLayout";
import { DashboardPage } from "../views/DashboardPage";
import { TemplatesPage } from "../views/TemplatesPage";
import { HistoryPage } from "../views/HistoryPage";
import { SettingsPage } from "../views/SettingsPage";

export const router = createBrowserRouter([
  {
    element: <AppLayout />,
    children: [
      { path: "/", element: <DashboardPage /> },
      { path: "/templates", element: <TemplatesPage /> },
      { path: "/history", element: <HistoryPage /> },
      { path: "/settings", element: <SettingsPage /> }
    ]
  }
]);

