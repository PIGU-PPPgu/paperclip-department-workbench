import type { PaperclipPluginManifestV1 } from "@paperclipai/plugin-sdk";

const manifest: PaperclipPluginManifestV1 = {
  id: "paperclip.department-workbench",
  apiVersion: 1,
  version: "1.0.5",
  displayName: "部门工作台",
  description: "查看部门项目、任务、目标与工作记录的入口。",
  author: "Wisdom Campus Office",
  categories: ["ui"],
  capabilities: [
    "ui.sidebar.register",
    "ui.page.register",
    "ui.dashboardWidget.register",
    "ui.detailTab.register",
    "projects.read",
    "issues.read",
    "goals.read",
    "goals.create",
    "access.members.read",
    "database.namespace.migrate",
    "database.namespace.read",
    "database.namespace.write"
  ],
  database: { namespaceSlug: "department_workbench", migrationsDir: "migrations", coreReadTables: ["companies", "projects", "goals"] },
  entrypoints: { worker: "./dist/worker.js", ui: "./dist/ui" },
  ui: {
    slots: [
      { id: "office-sidebar", type: "sidebar", displayName: "部门工作台", exportName: "SidebarLink", order: 32 },
      { id: "office-dashboard", type: "dashboardWidget", displayName: "部门工作台", exportName: "DashboardWidget", order: 10 },
      { id: "office-project-detail", type: "detailTab", entityTypes: ["project"], displayName: "部门工作", exportName: "ProjectDetail", order: 40 },
      { id: "office-page", type: "page", routePath: "office", displayName: "部门工作台", exportName: "OfficePage" }
    ]
  }
};

export default manifest;
