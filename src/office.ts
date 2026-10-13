import { createHash } from "node:crypto";

export type EntryKind = "meeting" | "development" | "training" | "milestone" | "kr" | "document" | "report";
export type OfficeEntry = {
  id: string;
  companyId: string;
  projectId: string | null;
  goalId: string | null;
  kind: EntryKind;
  status: string;
  date: string | null;
  title: string;
  detail: string;
  data: Record<string, unknown>;
  revision: number;
  createdByType: string;
  createdById: string | null;
  updatedByType: string;
  updatedById: string | null;
  createdAt: string;
  updatedAt: string;
};
export type OfficeProject = { id: string; name: string; status: string };
export type OfficeGoal = { id: string; title: string; status: string };
export type OfficeIssue = { id: string; identifier: string | null; title: string; status: string; projectId: string | null; assigneeAgentId: string | null; updatedAt: string };
export type OfficeAgent = { id: string; urlKey: string; name: string; status: string; scheduled: boolean; lastHeartbeatAt: string | null };
export type OfficeSnapshot = { projects: OfficeProject[]; goals: OfficeGoal[]; issues: OfficeIssue[]; agents: OfficeAgent[]; entries: OfficeEntry[] };
export type DocumentVersion = { number: number; date: string; note: string; content: string; authorWithAi: boolean; authorId: string | null };
export type KrCheckin = { date: string; value: number; authorId: string | null };

export const text = (value: unknown, label: string, max: number, required = true): string => {
  if (typeof value !== "string") throw new Error(`${label}无效`);
  const result = value.trim();
  if ((required && !result) || result.length > max) throw new Error(`${label}须为 ${required ? "1" : "0"}–${max} 字`);
  return result;
};
export const date = (value: unknown, label = "日期"): string => {
  if (typeof value !== "string" || !/^\d{4}-\d{2}-\d{2}$/.test(value)) throw new Error(`${label}无效`);
  const parsed = Date.parse(`${value}T00:00:00Z`);
  if (!Number.isFinite(parsed) || new Date(parsed).toISOString().slice(0, 10) !== value) throw new Error(`${label}无效`);
  return value;
};
export const month = (value: unknown): string => {
  if (typeof value !== "string" || !/^\d{4}-(0[1-9]|1[0-2])$/.test(value)) throw new Error("月份无效");
  return value;
};
export const object = (value: unknown): Record<string, unknown> => {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new Error("内容无效");
  return value as Record<string, unknown>;
};
export const stringList = (value: unknown, label: string, maxCount: number, maxLength: number): string[] => {
  if (value === undefined || value === null) return [];
  if (!Array.isArray(value) || value.length > maxCount) throw new Error(`${label}数量无效`);
  return value.map((item) => text(item, label, maxLength));
};
export const githubUrls = (value: unknown): string[] => stringList(value, "GitHub 地址", 8, 300).map((raw) => {
  let url: URL;
  try { url = new URL(raw); } catch { throw new Error("GitHub 地址无效"); }
  if (url.protocol !== "https:" || url.hostname !== "github.com") throw new Error("只接受 GitHub HTTPS 地址");
  return url.toString();
});
export const finiteNumber = (value: unknown, label: string, min: number, max: number): number => {
  if (typeof value !== "number" || !Number.isFinite(value) || value < min || value > max) throw new Error(`${label}须在 ${min}–${max} 之间`);
  return value;
};

export function scopeEntries(entries: OfficeEntry[], period: string, projectId = ""): OfficeEntry[] {
  return entries.filter((entry) => entry.status !== "archived" && entry.date?.startsWith(period) && (!projectId || entry.projectId === projectId));
}
export function krValueAsOf(entry: OfficeEntry, period: string): number | null {
  const checkins = (Array.isArray(entry.data.checkins) ? entry.data.checkins : []) as KrCheckin[];
  const relevant = checkins.filter((item) => item.date.slice(0, 7) <= period).sort((a, b) => a.date.localeCompare(b.date));
  return relevant.length ? relevant.at(-1)!.value : null;
}
export function reportSources(entries: OfficeEntry[], period: string, projectId = ""): OfficeEntry[] {
  return entries.filter((entry) => {
    if (entry.status === "archived" || entry.kind === "report" || (projectId && entry.projectId !== projectId)) return false;
    if (["meeting", "development", "training", "document"].includes(entry.kind) && entry.status !== "confirmed") return false;
    if (entry.kind === "kr") return entry.status === "active" && krValueAsOf(entry, period) !== null;
    if (entry.kind === "document") return (Array.isArray(entry.data.versions) ? entry.data.versions : []).some((version: DocumentVersion) => version.date.startsWith(period));
    if (entry.kind === "milestone") return entry.status === "done" && String(entry.data.completedAt ?? "").startsWith(period);
    return !!entry.date?.startsWith(period);
  });
}
export function reportFingerprint(entries: OfficeEntry[], period: string, projectId = ""): string {
  const sources = reportSources(entries, period, projectId).map((entry) => {
    if (entry.kind === "document") return [entry.id, entry.status, entry.title, (entry.data.versions as DocumentVersion[]).filter((version) => version.date.startsWith(period))];
    if (entry.kind === "kr") return [entry.id, entry.status, entry.title, entry.data.target, entry.data.unit, krValueAsOf(entry, period)];
    return [entry.id, entry.revision, entry.status, entry.data];
  });
  sources.sort((a, b) => String(a[0]).localeCompare(String(b[0])));
  return createHash("sha256").update(JSON.stringify(sources)).digest("hex");
}
export function reportText(entries: OfficeEntry[], period: string, projectId = "", projectName = ""): string {
  const scoped = scopeEntries(entries, period, projectId);
  const journal = scoped.filter((entry) => ["meeting", "development", "training"].includes(entry.kind) && entry.status === "confirmed");
  const meetings = journal.filter((entry) => entry.kind === "meeting").length;
  const development = journal.filter((entry) => entry.kind === "development").length;
  const trainings = journal.filter((entry) => entry.kind === "training").reduce((sum, entry) => sum + Number(entry.data.sessions ?? 1), 0);
  const done = entries.filter((entry) => entry.kind === "milestone" && entry.status === "done" && String(entry.data.completedAt ?? "").startsWith(period) && (!projectId || entry.projectId === projectId));
  const milestones = done.length;
  const documents = entries.filter((entry) => entry.kind === "document" && entry.status === "confirmed" && (!projectId || entry.projectId === projectId));
  const versions = documents.flatMap((entry) => (Array.isArray(entry.data.versions) ? entry.data.versions : []) as DocumentVersion[]).filter((version) => version.date.startsWith(period)).length;
  const krs = entries.filter((entry) => entry.kind === "kr" && entry.status === "active" && (!projectId || entry.projectId === projectId) && krValueAsOf(entry, period) !== null);
  const krCount = krs.length;
  const heading = projectId ? `${projectName || "项目"} · ${period}` : period;
  const lines = [`${heading} 月报`, `会议 ${meetings} 场，开发 ${development} 项，培训 ${trainings} 场。`, `完成里程碑 ${milestones} 个，文件新增版本 ${versions} 个，跟踪 KR ${krCount} 项。`];
  const dated = (kind: EntryKind) => journal.filter((entry) => entry.kind === kind).sort((a, b) => (a.date ?? "").localeCompare(b.date ?? ""));
  for (const [kind, label] of [["meeting", "会议"], ["development", "开发"], ["training", "培训"]] as const) {
    const items = dated(kind);
    if (items.length) lines.push(`${label}：${items.map((entry) => `${entry.date} ${entry.title}`).join("；")}`);
  }
  if (done.length) lines.push(`里程碑：${done.map((entry) => `${entry.data.completedAt} ${entry.title}`).join("；")}`);
  const changedDocs = documents.flatMap((entry) => (Array.isArray(entry.data.versions) ? entry.data.versions as DocumentVersion[] : []).filter((version) => version.date.startsWith(period)).map((version) => `${entry.title} v${version.number}`));
  if (changedDocs.length) lines.push(`知识库：${changedDocs.join("；")}`);
  if (krs.length) lines.push(`KR：${krs.map((entry) => `${entry.title} ${krValueAsOf(entry, period)}/${entry.data.target}${entry.data.unit ?? ""}`).join("；")}`);
  return lines.join("\n");
}
