import { useHostLocation, useHostNavigation, usePluginAction, usePluginData, type PluginDetailTabProps, type PluginPageProps, type PluginSidebarProps, type PluginWidgetProps } from "@paperclipai/plugin-sdk/ui";
import { useEffect, useMemo, useState, type FormEvent } from "react";
import { createPortal } from "react-dom";
import type { DocumentVersion, EntryKind, OfficeEntry, OfficeSnapshot } from "../office.js";
import { officePath } from "../routes.js";

type View = "overview" | "calendar" | "journal" | "review" | "okr" | "milestones" | "knowledge";
type CreateKind = Exclude<EntryKind, "report"> | "goal";
const tabs: { key: View; label: string }[] = [
  { key: "overview", label: "总览" }, { key: "calendar", label: "行事历" }, { key: "journal", label: "记事" },
  { key: "review", label: "待确认" }, { key: "okr", label: "OKR" }, { key: "milestones", label: "里程碑" }, { key: "knowledge", label: "知识库" },
];
const kindNames: Record<CreateKind, string> = { meeting: "会议", development: "开发", training: "培训", milestone: "里程碑", kr: "KR", document: "文件", goal: "目标" };
const empty: OfficeSnapshot = { projects: [], issues: [], goals: [], agents: [], entries: [] };
const currentMonth = () => new Date().toLocaleDateString("en-CA", { timeZone: "Asia/Shanghai", year: "numeric", month: "2-digit" }).slice(0, 7);
const today = () => new Date().toLocaleDateString("en-CA", { timeZone: "Asia/Shanghai", year: "numeric", month: "2-digit", day: "2-digit" });
const shortDate = (value: string | null) => value ? `${Number(value.slice(5, 7))}月${Number(value.slice(8, 10))}日` : "—";
const versions = (entry: OfficeEntry): DocumentVersion[] => Array.isArray(entry.data.versions) ? entry.data.versions as DocumentVersion[] : [];
const projectName = (data: OfficeSnapshot, id: string | null) => data.projects.find((project) => project.id === id)?.name ?? "未归属项目";
const goalName = (data: OfficeSnapshot, id: string | null) => data.goals.find((goal) => goal.id === id)?.title ?? "未关联目标";
const textList = (value: unknown): string[] => Array.isArray(value) ? value.filter((item): item is string => typeof item === "string") : [];
const agentStatus = (status: string) => ({ idle: "待命", running: "运行中", paused: "已暂停", error: "错误", pending_approval: "待批准", terminated: "已停用" } as Record<string, string>)[status] ?? status;

const css = `
.dw {
  --ink: var(--foreground, #171717);
  --muted-ink: var(--muted-foreground, #737373);
  --line: var(--border, #e5e5e5);
  --paper: var(--card, #fff);
  --surface: var(--muted, #f7f7f7);
  --accent: var(--primary, #262626);
  --accent-fg: var(--primary-foreground, #fff);
  --radius: var(--radius, .5rem);
  color: var(--ink);
  font: inherit;
  font-size: 14px;
  line-height: 1.5;
  padding: 8px 0 40px;
  min-height: 0;
  container-type: inline-size;
}
.dw *, .dw-modal * { box-sizing: border-box; }
.dw button, .dw input, .dw select, .dw textarea, .dw-modal button, .dw-modal input, .dw-modal select, .dw-modal textarea { font: inherit; }
.dw button, .dw-modal button { cursor: pointer; }
.dw :is(button,a,input,select,textarea):focus-visible, .dw-modal :is(button,a,input,select,textarea):focus-visible { outline: 2px solid var(--ring, #a3a3a3); outline-offset: 2px; }
.dw h1, .dw h2, .dw-modal h2 { margin: 0; color: var(--ink); }
.dw h2 { font-size: 14px; font-weight: 600; }
.dw p { line-height: 1.55; }
.dw a { color: var(--ink); text-decoration: none; }
.dw a:hover { text-decoration: underline; }
.dw-top { display: flex; align-items: center; justify-content: space-between; gap: 12px; flex-wrap: wrap; margin-bottom: 16px; }
.dw-filters { display: flex; align-items: center; flex-wrap: wrap; gap: 8px; }
.dw-filters label { color: var(--muted-ink); font-size: 12px; font-weight: 500; }
.dw-filters select, .dw-filters input { min-height: 36px; border: 1px solid var(--line); border-radius: 6px; padding: 0 10px; background: var(--paper); color: var(--ink); font-size: 14px; }
.dw-tabs { display: flex; gap: 2px; overflow-x: auto; border-bottom: 1px solid var(--line); margin-bottom: 18px; }
.dw-tabs button { flex: none; min-height: 40px; padding: 0 12px; border: 0; border-bottom: 2px solid transparent; background: transparent; color: var(--muted-ink); font-size: 13px; font-weight: 500; }
.dw-tabs button:hover { color: var(--ink); }
.dw-tabs button[aria-selected=true] { border-bottom-color: var(--ink); color: var(--ink); font-weight: 600; }
.dw-head { display: flex; align-items: center; justify-content: space-between; gap: 10px; margin-bottom: 14px; }
.dw-actions { display: flex; gap: 8px; flex-wrap: wrap; align-items: center; }
.dw-btn { display: inline-flex; align-items: center; justify-content: center; gap: 6px; min-height: 36px; padding: 0 12px; border: 1px solid var(--line); border-radius: 6px; background: var(--paper); color: var(--ink); box-shadow: 0 1px 2px rgb(0 0 0 / .05); font-size: 14px; font-weight: 500; white-space: nowrap; }
.dw-btn:hover { background: var(--surface); text-decoration: none; }
.dw-btn:disabled { opacity: .5; cursor: not-allowed; }
.dw-btn.primary { background: var(--accent); border-color: var(--accent); color: var(--accent-fg); }
.dw-btn.primary:hover { opacity: .88; }
.dw-btn.quiet { box-shadow: none; }
.dw-btn.danger { color: var(--destructive, #a43e38); }
.dw-error { padding: 10px 12px; margin: 10px 0; border: 1px solid var(--destructive, #dba6a2); border-radius: 6px; color: var(--destructive, #a43e38); background: var(--paper); font-size: 13px; }
.dw-empty { padding: 28px 16px; color: var(--muted-ink); font-size: 13px; }
.dw-muted { color: var(--muted-ink); font-size: 12px; }
.dw-stats { display: grid; grid-template-columns: repeat(4,minmax(0,1fr)); gap: 10px; margin-bottom: 16px; }
.dw-stat { padding: 14px 16px; border: 1px solid var(--line); border-radius: 8px; background: var(--paper); }
.dw-stat span { display: block; color: var(--muted-ink); font-size: 12px; }
.dw-stat strong { display: block; margin-top: 6px; color: var(--ink); font-size: 24px; line-height: 1.2; font-weight: 600; }
.dw-two { display: grid; grid-template-columns: minmax(0,1.5fr) minmax(260px,.9fr); gap: 14px; align-items: start; }
.dw-stack { display: grid; gap: 14px; }
.dw-panel { overflow: hidden; border: 1px solid var(--line); border-radius: 8px; background: var(--paper); }
.dw-panel > header { display: flex; align-items: center; justify-content: space-between; gap: 8px; min-height: 52px; padding: 10px 16px; border-bottom: 1px solid var(--line); }
.dw-panel > .body { padding: 16px; }
.dw-row { padding: 12px 0; border-bottom: 1px solid var(--line); }
.dw-row:first-child { padding-top: 0; }
.dw-row:last-child { padding-bottom: 0; border-bottom: 0; }
.dw-row strong { display: block; color: var(--ink); font-size: 14px; font-weight: 500; }
.dw-row small { display: block; margin-top: 4px; color: var(--muted-ink); font-size: 12px; }
.dw-agent-grid { display: grid; grid-template-columns: repeat(2,minmax(0,1fr)); }
.dw-agent-item { min-width: 0; padding: 13px 16px; border-bottom: 1px solid var(--line); }
.dw-agent-item:nth-child(odd) { border-right: 1px solid var(--line); }
.dw-agent-item strong { display: block; overflow: hidden; white-space: nowrap; text-overflow: ellipsis; font-size: 14px; font-weight: 500; }
.dw-agent-item small { display: block; margin-top: 4px; color: var(--muted-ink); font-size: 12px; }
.dw-agent-item .dw-pill { margin-top: 7px; }
.dw-agent-item .dw-agent-issue { display: block; margin-top: 6px; font-size: 12px; text-decoration: underline; text-underline-offset: 2px; }
.dw-agent-summary { display: flex; align-items: center; gap: 18px; flex-wrap: wrap; min-height: 42px; padding: 8px 16px; margin: -2px 0 16px; border: 1px solid var(--line); border-radius: 8px; background: var(--paper); font-size: 13px; }
.dw-agent-summary strong { font-weight: 600; }
.dw-agent-summary span { color: var(--muted-ink); }
.dw-agent-summary a { margin-left: auto; text-decoration: underline; text-underline-offset: 2px; }
.dw-pill { display: inline-block; padding: 3px 8px; border-radius: 5px; background: var(--surface); color: var(--muted-ink); font-size: 12px; font-weight: 500; }
.dw-pill.warm { background: color-mix(in oklch,var(--surface) 75%,#e7b367); color: var(--ink); }
.dw-pill.gray { background: var(--surface); color: var(--muted-ink); }
.dw-form { display: grid; grid-template-columns: repeat(3,minmax(0,1fr)); gap: 12px; padding: 16px; border: 1px solid var(--line); border-radius: 8px; background: var(--paper); margin-bottom: 14px; }
.dw-form h2 { grid-column: 1/-1; font-size: 16px; font-weight: 600; }
.dw-form label { display: grid; gap: 6px; color: var(--ink); font-size: 13px; font-weight: 500; }
.dw-form .wide { grid-column: 1/-1; }
.dw-form input, .dw-form select, .dw-form textarea { min-width: 0; min-height: 40px; border: 1px solid var(--line); border-radius: 6px; padding: 8px 10px; color: var(--ink); background: var(--paper); font-size: 14px; }
.dw-form textarea { min-height: 88px; resize: vertical; }
.dw-form .check { display: flex; align-items: center; gap: 8px; }
.dw-form .check input { min-height: auto; width: auto; }
.dw-form footer { grid-column: 1/-1; display: flex; justify-content: flex-end; gap: 8px; margin-top: 4px; }
.dw-modal { position: fixed; inset: 0; z-index: 100; display: flex; align-items: center; justify-content: center; padding: 24px; background: rgb(0 0 0 / .5); --ink:var(--foreground,#171717); --muted-ink:var(--muted-foreground,#737373); --line:var(--border,#e5e5e5); --paper:var(--card,#fff); --surface:var(--muted,#f7f7f7); --accent:var(--primary,#262626); --accent-fg:var(--primary-foreground,#fff); }
.dw-modal .dw-form { width: min(620px,100%); max-height: calc(100dvh - 48px); overflow-y: auto; grid-template-columns: repeat(2,minmax(0,1fr)); margin: 0; padding: 24px; box-shadow: 0 16px 48px rgb(0 0 0 / .18); }
.dw-modal .dw-form h2 { padding-right: 32px; margin-bottom: 2px; }
.dw-modal .dw-form .dw-close { position: absolute; top: 18px; right: 18px; width: 32px; height: 32px; min-height: 0; padding: 0; border: 0; background: transparent; color: var(--muted-ink); box-shadow: none; font-size: 20px; }
.dw-modal .dw-form { position: relative; }
.dw-kinds, .dw-cards { display: grid; grid-template-columns: repeat(3,minmax(0,1fr)); gap: 12px; }
.dw-kinds { margin-bottom: 14px; }
.dw-kind, .dw-record, .dw-doc { padding: 16px; border: 1px solid var(--line); border-radius: 8px; background: var(--paper); }
.dw-kind strong { font-size: 14px; font-weight: 500; }
.dw-kind span { display: block; margin-top: 12px; font-size: 24px; font-weight: 600; }
.dw-cards { grid-template-columns: repeat(2,minmax(0,1fr)); }
.dw-record h2, .dw-doc h2 { margin: 10px 0 5px; }
.dw-record p, .dw-doc p { margin: 8px 0; font-size: 13px; }
.dw-card-actions { display: flex; gap: 7px; flex-wrap: wrap; margin-top: 12px; }
.dw-card-actions .dw-btn { min-height: 32px; padding: 0 10px; font-size: 12px; }
.dw-calendar { display: grid; grid-template-columns: repeat(7,minmax(0,1fr)); border-top: 1px solid var(--line); border-left: 1px solid var(--line); }
.dw-calendar > * { min-height: 76px; padding: 7px; border-right: 1px solid var(--line); border-bottom: 1px solid var(--line); background: var(--paper); }
.dw-calendar .weekday { min-height: 28px; text-align: center; color: var(--muted-ink); font-size: 12px; background: var(--surface); }
.dw-calendar button { width: 100%; border-top: 0; border-left: 0; color: var(--ink); text-align: left; font-size: 12px; }
.dw-calendar button[data-active=true] { background: var(--surface); box-shadow: inset 0 0 0 1px var(--ink); }
.dw-calendar em { display: block; margin-top: 4px; overflow: hidden; white-space: nowrap; text-overflow: ellipsis; border-radius: 4px; padding: 3px 5px; background: var(--surface); color: var(--ink); font-size: 11px; font-style: normal; }
.dw-calendar em.milestone { border: 1px solid var(--line); }
.dw-kr { padding: 14px 0; border-bottom: 1px solid var(--line); }
.dw-kr:last-child { border: 0; }
.dw-kr > div:first-child { display: flex; align-items: center; justify-content: space-between; gap: 8px; }
.dw-kr strong { font-size: 14px; font-weight: 500; }
.dw-progress { height: 7px; overflow: hidden; border-radius: 8px; background: var(--surface); margin: 10px 0; }
.dw-progress i { display: block; height: 100%; border-radius: 8px; background: var(--ink); }
.dw-mini { display: flex; gap: 8px; align-items: center; flex-wrap: wrap; }
.dw-mini input { width: 90px; min-height: 36px; padding: 6px 8px; border: 1px solid var(--line); border-radius: 6px; background: var(--paper); color: var(--ink); font-size: 13px; }
.dw-timeline { padding: 8px 0; }
.dw-milestone { position: relative; display: grid; grid-template-columns: 72px 14px minmax(0,1fr) auto; gap: 10px; align-items: center; padding: 17px 16px; }
.dw-milestone:not(:last-child)::after { content: ""; position: absolute; left: 104px; top: 32px; bottom: -32px; width: 2px; background: var(--line); }
.dw-milestone .dot { z-index: 1; width: 12px; height: 12px; border: 2px solid var(--muted-ink); border-radius: 50%; background: var(--paper); }
.dw-milestone.done .dot { border-color: var(--ink); background: var(--ink); box-shadow: inset 0 0 0 2px var(--paper); }
.dw-milestone time { color: var(--muted-ink); font-size: 12px; font-weight: 500; }
.dw-milestone strong { display: block; font-size: 14px; font-weight: 500; }
.dw-milestone small { display: block; margin-top: 4px; color: var(--muted-ink); font-size: 12px; }
.dw-version { padding: 10px 0; border-top: 1px solid var(--line); }
.dw-version strong { font-size: 13px; font-weight: 600; }
.dw-version small { margin-left: 8px; color: var(--muted-ink); font-size: 12px; }
.dw-version details { margin-top: 7px; font-size: 13px; white-space: pre-wrap; }
.dw-summary { min-height: 0; padding: 16px; border: 1px solid var(--line); border-radius: 8px; background: var(--paper); }
.dw-summary h2 { margin-bottom: 12px; }
.dw-summary .dw-stats { grid-template-columns: repeat(2,minmax(0,1fr)); }
.dw-summary .dw-stat { padding: 10px; }
.dw-summary .dw-stat strong { font-size: 20px; }
.dw-project { min-height: 0; }
@container (max-width:700px) { .dw-two { grid-template-columns: 1fr; }.dw-stats { grid-template-columns: repeat(2,minmax(0,1fr)); }.dw-cards { grid-template-columns: 1fr; }.dw-agent-grid { grid-template-columns: 1fr; }.dw-agent-item:nth-child(odd) { border-right: 0; } }
@container (max-width:490px) { .dw-milestone { grid-template-columns: 57px 14px minmax(0,1fr); }.dw-milestone .dw-btn { grid-column: 3; justify-self: start; }.dw-milestone:not(:last-child)::after { left: 89px; }.dw-calendar > * { min-height: 58px; } }
@media (max-width:600px) { .dw-modal { align-items: stretch; padding: 12px; }.dw-modal .dw-form { max-height: calc(100dvh - 24px); grid-template-columns: 1fr; padding: 18px; }.dw-modal .dw-form .wide { grid-column: 1; } }
`;

export function SidebarLink({ context }: PluginSidebarProps) {
  const navigation = useHostNavigation();
  const { pathname } = useHostLocation();
  return <a {...navigation.linkProps(officePath(context.companyPrefix, pathname))} className="flex items-center gap-2.5 px-3 py-2 text-[13px] font-medium text-foreground/80 transition-colors hover:bg-accent/50 hover:text-foreground"><span aria-hidden="true">▦</span><span>部门工作台</span></a>;
}

function AgentStatusPanel({ data, projectId, prefix }: { data: OfficeSnapshot; projectId: string; prefix: string }) {
  const navigation = useHostNavigation();
  const path = (suffix: string) => `${prefix ? `/${prefix}` : ""}${suffix}`;
  return <section className="dw-panel" style={{ marginTop: 14 }}><header><h2>智能体</h2><a className="dw-btn" {...navigation.linkProps(path("/agents/all"))}>查看全部</a></header><div className="dw-agent-grid">{data.agents.map((agent) => {
    const blocked = data.issues.filter((issue) => issue.assigneeAgentId === agent.id && issue.status === "blocked" && (!projectId || issue.projectId === projectId));
    return <div className="dw-agent-item" key={agent.id}><a {...navigation.linkProps(path(`/agents/${agent.urlKey}/overview`))}><strong>{agent.name}</strong></a><span className={`dw-pill ${agent.status === "paused" ? "warm" : ""}`}>{agentStatus(agent.status)}</span><small>{agent.scheduled ? "定时运行" : "手动运行"}{blocked.length ? ` · ${blocked.length} 项受阻` : ""}</small>{blocked[0] ? <a className="dw-agent-issue" {...navigation.linkProps(path(`/issues/${blocked[0].id}`))}>{blocked[0].identifier ?? "任务"} · {blocked[0].title} ↗</a> : null}</div>;
  })}{!data.agents.length ? <div className="dw-empty">暂无智能体</div> : null}</div></section>;
}

function EntryForm({ kind, data, projectId, initial, onCancel, onSave, busy }: { kind: CreateKind; data: OfficeSnapshot; projectId: string; initial?: OfficeEntry | null; onCancel: () => void; onSave: (payload: Record<string, unknown>) => Promise<boolean>; busy: boolean }) {
  const [title, setTitle] = useState(initial?.title ?? "");
  const [detail, setDetail] = useState(initial?.detail ?? "");
  const [project, setProject] = useState(initial?.projectId ?? projectId ?? data.projects[0]?.id ?? "");
  const [goal, setGoal] = useState(initial?.goalId ?? data.goals[0]?.id ?? "");
  const [when, setWhen] = useState(initial?.date ?? today());
  const [pending, setPending] = useState(initial?.status === "pending");
  const [level, setLevel] = useState(String(initial?.data.level ?? "项目组"));
  const [sessions, setSessions] = useState(String(initial?.data.sessions ?? 1));
  const [contributors, setContributors] = useState(textList(initial?.data.contributors).join("、"));
  const [repoUrls, setRepoUrls] = useState(textList(initial?.data.repoUrls).join("\n"));
  const [target, setTarget] = useState(String(initial?.data.target ?? ""));
  const [current, setCurrent] = useState(String(initial?.data.current ?? "0"));
  const [unit, setUnit] = useState(String(initial?.data.unit ?? "项"));
  const [content, setContent] = useState("");
  const [note, setNote] = useState("初版");
  const [authorWithAi, setAuthorWithAi] = useState(false);
  const isGoal = kind === "goal";
  const isJournal = ["meeting", "development", "training"].includes(kind);
  useEffect(() => {
    const previous = document.body.style.overflow;
    document.body.style.overflow = "hidden";
    return () => { document.body.style.overflow = previous; };
  }, []);
  const submit = async (event: FormEvent) => {
    event.preventDefault();
    const payload: Record<string, unknown> = { kind, title, detail, projectId: project, goalId: kind === "kr" || kind === "milestone" ? goal : "", date: when, pending };
    if (kind === "meeting") payload.data = { level, contributors: contributors.split(/[、,，\n]+/).map((item) => item.trim()).filter(Boolean) };
    if (kind === "development") payload.data = { contributors: contributors.split(/[、,，\n]+/).map((item) => item.trim()).filter(Boolean), repoUrls: repoUrls.split(/[\n,，]+/).map((item) => item.trim()).filter(Boolean) };
    if (kind === "training") payload.data = { sessions: Number(sessions), contributors: contributors.split(/[、,，\n]+/).map((item) => item.trim()).filter(Boolean) };
    if (kind === "kr") payload.data = { target: Number(target), current: Number(current), unit };
    if (kind === "document") payload.data = { content, note, authorWithAi };
    if (initial) payload.revision = initial.revision;
    if (await onSave(payload)) onCancel();
  };
  if (typeof document === "undefined") return null;
  return createPortal(<div className="dw-modal" onMouseDown={(event) => { if (event.target === event.currentTarget) onCancel(); }}><form className="dw-form" role="dialog" aria-modal="true" aria-label={initial ? `编辑${kindNames[kind]}` : `新建${kindNames[kind]}`} onKeyDown={(event) => { if (event.key === "Escape") onCancel(); }} onSubmit={submit}><h2>{initial ? `编辑${kindNames[kind]}` : `新建${kindNames[kind]}`}</h2><button className="dw-close" type="button" aria-label="关闭" onClick={onCancel}>×</button>
    <label className="wide">标题<input autoFocus required maxLength={120} value={title} onChange={(event) => setTitle(event.target.value)} /></label>
    {!isGoal ? <label>项目<select required={kind !== "kr"} value={project} onChange={(event) => setProject(event.target.value)}><option value="">选择项目</option>{data.projects.map((item) => <option value={item.id} key={item.id}>{item.name}</option>)}</select></label> : null}
    {(kind === "kr" || kind === "milestone") ? <label>目标<select required={kind === "kr"} value={goal} onChange={(event) => setGoal(event.target.value)}><option value="">{kind === "kr" ? "选择目标" : "不关联"}</option>{data.goals.map((item) => <option value={item.id} key={item.id}>{item.title}</option>)}</select></label> : null}
    {!isGoal && kind !== "kr" ? <label>日期<input type="date" required value={when} onChange={(event) => setWhen(event.target.value)} /></label> : null}
    <label className="wide">{isGoal ? "说明" : kind === "document" ? "说明" : "内容"}<textarea value={detail} onChange={(event) => setDetail(event.target.value)} maxLength={20000} /></label>
    {kind === "meeting" ? <label>会议级别<input value={level} onChange={(event) => setLevel(event.target.value)} maxLength={40} /></label> : null}
    {kind === "training" ? <label>培训场次<input type="number" min="1" max="1000" value={sessions} onChange={(event) => setSessions(event.target.value)} /></label> : null}
    {isJournal ? <label className="wide">参与人<input value={contributors} onChange={(event) => setContributors(event.target.value)} placeholder="用逗号分隔" /></label> : null}
    {kind === "development" ? <label className="wide">GitHub 地址<textarea value={repoUrls} onChange={(event) => setRepoUrls(event.target.value)} placeholder="每行一个 HTTPS 链接" /></label> : null}
    {kind === "kr" ? <><label>目标值<input type="number" min="0.000001" step="any" required value={target} onChange={(event) => setTarget(event.target.value)} /></label><label>当前值<input type="number" min="0" step="any" required value={current} onChange={(event) => setCurrent(event.target.value)} /></label><label>单位<input required maxLength={20} value={unit} onChange={(event) => setUnit(event.target.value)} /></label></> : null}
    {kind === "document" && !initial ? <><label className="wide">正文<textarea value={content} onChange={(event) => setContent(event.target.value)} maxLength={50000} /></label><label>版本说明<input value={note} onChange={(event) => setNote(event.target.value)} maxLength={120} /></label></> : null}
    {kind === "document" && !initial ? <label className="check"><input type="checkbox" checked={authorWithAi} onChange={(event) => setAuthorWithAi(event.target.checked)} />AI 协作</label> : null}
    {(isJournal || kind === "document") && !initial ? <label className="check"><input type="checkbox" checked={pending} onChange={(event) => setPending(event.target.checked)} />提交待确认</label> : null}
    <footer><button className="dw-btn primary" type="submit" disabled={busy}>{initial ? "保存修改" : "创建"}</button><button className="dw-btn quiet" type="button" onClick={onCancel}>取消</button></footer>
  </form></div>, document.body);
}

function Workbench({ context }: PluginPageProps) {
  const navigation = useHostNavigation();
  const { search, pathname } = useHostLocation();
  const params = new URLSearchParams(search);
  useEffect(() => {
    if (!params.has("mode")) return;
    const next = new URLSearchParams(search);
    next.delete("mode");
    navigation.navigate(`/office${next.size ? `?${next}` : ""}`);
  }, [search, navigation]);
  const view = tabs.some((tab) => tab.key === params.get("view")) ? params.get("view") as View : "overview";
  const month = /^\d{4}-(0[1-9]|1[0-2])$/.test(params.get("month") ?? "") ? params.get("month")! : currentMonth();
  const projectId = params.get("project") ?? "";
  const request = useMemo(() => ({ companyId: context.companyId ?? "" }), [context.companyId]);
  const query = usePluginData<OfficeSnapshot>("office", request);
  const create = usePluginAction("office-create");
  const createGoal = usePluginAction("office-goal");
  const changeEntry = usePluginAction("office-change");
  const generateReport = usePluginAction("office-report");
  const data = query.data ?? empty;
  const [createKind, setCreateKind] = useState<CreateKind | null>(null);
  const [editing, setEditing] = useState<OfficeEntry | null>(null);
  const [selectedDay, setSelectedDay] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);
  const [progressValues, setProgressValues] = useState<Record<string, string>>({});
  const [versionDoc, setVersionDoc] = useState("");
  const [versionContent, setVersionContent] = useState("");
  const [versionNote, setVersionNote] = useState("");
  const [versionAi, setVersionAi] = useState(false);
  const [reportEdit, setReportEdit] = useState(false);
  const [reportBody, setReportBody] = useState("");
  const prefix = context.companyPrefix || pathname.split("/").filter(Boolean)[0] || "";
  const hostPath = (suffix: string) => prefix ? `/${prefix}${suffix}` : suffix;
  const changeRoute = (updates: Record<string, string>) => { const next = new URLSearchParams(search); next.delete("mode"); for (const [key, value] of Object.entries(updates)) value ? next.set(key, value) : next.delete(key); navigation.navigate(`/office${next.size ? `?${next}` : ""}`); };
  const run = async (action: () => Promise<unknown>): Promise<boolean> => {
    if (busy) return false;
    setBusy(true); setError("");
    try { await action(); query.refresh(); return true; }
    catch (cause) { setError(cause instanceof Error ? cause.message : String(cause)); query.refresh(); return false; }
    finally { setBusy(false); }
  };
  const save = async (payload: Record<string, unknown>) => {
    if (editing) return run(() => changeEntry({ companyId: context.companyId, id: editing.id, revision: editing.revision, operation: "edit", ...payload }));
    if (createKind === "goal") return run(() => createGoal({ companyId: context.companyId, title: payload.title, description: payload.detail }));
    return run(() => create({ companyId: context.companyId, entry: payload }));
  };
  const mutate = (entry: OfficeEntry, operation: string, extra: Record<string, unknown> = {}) => run(() => changeEntry({ companyId: context.companyId, id: entry.id, revision: entry.revision, operation, ...extra }));
  const showCreate = (kind: CreateKind) => { setEditing(null); setCreateKind(kind); setError(""); };
  const showEdit = (entry: OfficeEntry) => { setCreateKind(entry.kind as CreateKind); setEditing(entry); setError(""); };
  const closeForm = () => { setCreateKind(null); setEditing(null); };
  const visible = data.entries.filter((entry) => entry.status !== "archived" && (!projectId || entry.projectId === projectId));
  const journals = visible.filter((entry) => ["meeting", "development", "training"].includes(entry.kind) && entry.status === "confirmed" && entry.date?.startsWith(month));
  const milestones = visible.filter((entry) => entry.kind === "milestone" && entry.date?.startsWith(month));
  const documents = visible.filter((entry) => entry.kind === "document" && entry.status === "confirmed");
  const pending = visible.filter((entry) => entry.status === "pending");
  const reviewIssues = data.issues.filter((issue) => issue.status === "in_review" && (!projectId || issue.projectId === projectId));
  const reports = visible.filter((entry) => entry.kind === "report");
  const report = reports.find((entry) => entry.date === `${month}-01`) ?? null;
  const goals = data.goals;
  const count = (kind: string) => journals.filter((entry) => entry.kind === kind).reduce((sum, entry) => sum + (kind === "training" ? Number(entry.data.sessions ?? 1) : 1), 0);
  const calendarEvents = (day: string) => [...journals.filter((entry) => entry.date === day), ...milestones.filter((entry) => entry.date === day)];
  const topAction = ({ journal: "meeting", calendar: "milestone", okr: "goal", milestones: "milestone", knowledge: "document" } as Partial<Record<View, CreateKind>>)[view];

  return <><style>{css}</style><div className="dw"><div className="dw-top"><div className="dw-filters"><label htmlFor="dw-project">项目</label><select id="dw-project" value={projectId} onChange={(event) => changeRoute({ project: event.target.value })}><option value="">全部项目</option>{data.projects.map((project) => <option value={project.id} key={project.id}>{project.name}</option>)}</select><label htmlFor="dw-month">月份</label><input id="dw-month" type="month" value={month} onChange={(event) => { setSelectedDay(""); changeRoute({ month: event.target.value }); }} /></div><div className="dw-actions"><a className="dw-btn" {...navigation.linkProps(hostPath("/projects"))}>项目</a><button className="dw-btn primary" onClick={() => showCreate(topAction ?? "meeting")}>＋ 新建</button></div></div>
    <nav className="dw-tabs" aria-label="部门工作台视图">{tabs.map((tab) => <button type="button" key={tab.key} aria-selected={view === tab.key} onClick={() => changeRoute({ view: tab.key })}>{tab.label}</button>)}</nav>
    {query.loading && !query.data ? <div className="dw-empty">正在加载</div> : null}{query.error ? <div className="dw-error">读取失败：{query.error.message}</div> : null}{error ? <div className="dw-error" role="alert">{error}</div> : null}
    {createKind ? <EntryForm key={`${createKind}-${editing?.id ?? "new"}`} kind={createKind} data={data} projectId={projectId} initial={editing} onCancel={closeForm} onSave={save} busy={busy} /> : null}

    {view === "overview" ? <><div className="dw-stats"><div className="dw-stat"><span>会议</span><strong>{count("meeting")}</strong></div><div className="dw-stat"><span>开发</span><strong>{count("development")}</strong></div><div className="dw-stat"><span>培训场次</span><strong>{count("training")}</strong></div><div className="dw-stat"><span>待确认</span><strong>{pending.length + reviewIssues.length + (report?.status === "draft" ? 1 : 0)}</strong></div></div><div className="dw-agent-summary"><strong>智能体</strong><span>{data.agents.filter((agent) => agent.status === "idle" || agent.status === "running").length} 可运行</span><span>{data.agents.filter((agent) => agent.status === "paused").length} 已暂停</span><span>{data.issues.filter((issue) => issue.status === "blocked" && issue.assigneeAgentId).length} 受阻任务</span><a {...navigation.linkProps(hostPath("/agents/all"))}>查看</a></div><div className="dw-two"><div className="dw-stack"><section className="dw-panel"><header><h2>本月节点</h2><button className="dw-btn" onClick={() => changeRoute({ view: "milestones" })}>里程碑</button></header><div className="body">{milestones.length ? milestones.slice().sort((a, b) => (a.date ?? "").localeCompare(b.date ?? "")).map((entry) => <div className="dw-row" key={entry.id}><strong>{shortDate(entry.date)} · {entry.title}</strong><small>{projectName(data, entry.projectId)} · {entry.status === "done" ? "已完成" : "计划中"}</small></div>) : <div className="dw-muted">暂无节点</div>}</div></section><section className="dw-panel"><header><h2>本月记录</h2><button className="dw-btn" onClick={() => changeRoute({ view: "journal" })}>记事</button></header><div className="body">{journals.length ? journals.slice().sort((a, b) => (b.date ?? "").localeCompare(a.date ?? "")).slice(0, 5).map((entry) => <div className="dw-row" key={entry.id}><strong>{entry.title}</strong><small>{shortDate(entry.date)} · {kindNames[entry.kind as CreateKind]} · {projectName(data, entry.projectId)}</small></div>) : <div className="dw-muted">暂无记录</div>}</div></section></div><div className="dw-stack"><section className="dw-panel"><header><h2>月报</h2></header><div className="body">{report ? <><span className={`dw-pill ${report.status === "draft" ? "warm" : ""}`}>{report.status === "draft" ? "待确认" : "已确认"}</span><p style={{ whiteSpace: "pre-wrap", fontSize: 11 }}>{report.detail}</p><div className="dw-card-actions">{report.status === "draft" ? <><button className="dw-btn" onClick={() => { setReportBody(report.detail); setReportEdit(!reportEdit); }}>编辑</button><button className="dw-btn primary" disabled={busy} onClick={() => mutate(report, "confirm-report")}>确认</button></> : null}<button className="dw-btn" disabled={busy} onClick={() => run(() => generateReport({ companyId: context.companyId, month, projectId }))}>重新生成</button></div>{reportEdit && report.status === "draft" ? <div className="dw-form" style={{ marginTop: 10 }}><label className="wide">正文<textarea value={reportBody} onChange={(event) => setReportBody(event.target.value)} /></label><footer><button className="dw-btn primary" disabled={busy} onClick={() => void mutate(report, "edit", { detail: reportBody }).then((ok) => ok && setReportEdit(false))}>保存</button></footer></div> : null}</> : <><p className="dw-muted">暂无月报</p><button className="dw-btn primary" disabled={busy} onClick={() => run(() => generateReport({ companyId: context.companyId, month, projectId }))}>生成月报</button></>}</div></section><section className="dw-panel"><header><h2>目标</h2><button className="dw-btn" onClick={() => showCreate("goal")}>新建目标</button></header><div className="body">{goals.length ? goals.slice(0, 4).map((goal) => <div className="dw-row" key={goal.id}><strong>{goal.title}</strong><small>{goal.status}</small></div>) : <div className="dw-muted">暂无目标</div>}</div></section></div></div></> : null}

    {view === "overview" ? <AgentStatusPanel data={data} projectId={projectId} prefix={prefix} /> : null}
    {view === "calendar" ? <div className="dw-two"><section className="dw-panel"><header><h2>月历</h2><button className="dw-btn" onClick={() => showCreate("milestone")}>新建节点</button></header><div className="dw-calendar">{["一", "二", "三", "四", "五", "六", "日"].map((day) => <div className="weekday" key={day}>{day}</div>)}{(() => { const [year, m] = month.split("-").map(Number); const offset = (new Date(year, m - 1, 1).getDay() + 6) % 7; const count = new Date(year, m, 0).getDate(); return [...Array(offset).fill(null), ...Array.from({ length: count }, (_, index) => index + 1)].map((day, index) => { if (!day) return <div key={`e-${index}`} />; const value = `${month}-${String(day).padStart(2, "0")}`; return <button type="button" data-active={selectedDay === value} key={value} onClick={() => setSelectedDay(selectedDay === value ? "" : value)}><b>{day}</b>{calendarEvents(value).slice(0, 2).map((entry) => <em key={entry.id} className={entry.kind === "milestone" ? "milestone" : ""}>{entry.title}</em>)}</button>; }); })()}</div></section><section className="dw-panel"><header><h2>{selectedDay ? shortDate(selectedDay) : "本月事项"}</h2></header><div className="body">{(() => { const events = [...journals, ...milestones].filter((entry) => !selectedDay || entry.date === selectedDay).sort((a, b) => (a.date ?? "").localeCompare(b.date ?? "")); return events.length ? events.map((entry) => <div className="dw-row" key={entry.id}><strong>{entry.title}</strong><small>{shortDate(entry.date)} · {entry.kind === "milestone" ? "里程碑" : kindNames[entry.kind as CreateKind]} · {projectName(data, entry.projectId)}</small></div>) : <div className="dw-muted">暂无事项</div>; })()}</div></section></div> : null}

    {view === "journal" ? <><div className="dw-head"><h2>记事</h2><div className="dw-actions">{(["meeting", "development", "training"] as CreateKind[]).map((kind) => <button className="dw-btn" key={kind} onClick={() => showCreate(kind)}>新建{kindNames[kind]}</button>)}</div></div><div className="dw-kinds">{(["meeting", "development", "training"] as const).map((kind) => <div className="dw-kind" key={kind}><strong>{kindNames[kind]}</strong><span>{count(kind)}</span></div>)}</div><div className="dw-cards">{journals.slice().sort((a, b) => (b.date ?? "").localeCompare(a.date ?? "")).map((entry) => <article className="dw-record" key={entry.id}><span className="dw-pill">{kindNames[entry.kind as CreateKind]}{entry.kind === "meeting" ? ` · ${String(entry.data.level ?? "")}` : ""}</span><h2>{entry.title}</h2><small className="dw-muted">{shortDate(entry.date)} · {projectName(data, entry.projectId)}{entry.kind === "training" ? ` · ${entry.data.sessions ?? 1} 场` : ""}</small><p>{entry.detail}</p>{textList(entry.data.contributors).length ? <small className="dw-muted">参与：{textList(entry.data.contributors).join("、")}</small> : null}{textList(entry.data.repoUrls).map((url) => <div key={url}><a href={url} target="_blank" rel="noreferrer">{url} ↗</a></div>)}<div className="dw-card-actions"><button className="dw-btn" onClick={() => showEdit(entry)}>编辑</button></div></article>)}{!journals.length ? <div className="dw-empty">本月暂无记事。点击上方按钮新建。</div> : null}</div></> : null}

    {view === "review" ? <div className="dw-two"><section className="dw-panel"><header><h2>待确认</h2><span className="dw-pill warm">{pending.length + reviewIssues.length + (report?.status === "draft" ? 1 : 0)}</span></header><div className="body">{pending.map((entry) => <div className="dw-row" key={entry.id}><strong>{entry.title}</strong><small>{kindNames[entry.kind as CreateKind]} · {projectName(data, entry.projectId)} · {shortDate(entry.date)}</small><p style={{ fontSize: 11 }}>{entry.detail}</p><div className="dw-card-actions"><button className="dw-btn primary" disabled={busy} onClick={() => mutate(entry, "confirm")}>确认入账</button><button className="dw-btn danger" disabled={busy} onClick={() => mutate(entry, "reject")}>驳回</button><button className="dw-btn" onClick={() => showEdit(entry)}>编辑</button></div></div>)}{report?.status === "draft" ? <div className="dw-row"><strong>{report.title}</strong><p style={{ whiteSpace: "pre-wrap", fontSize: 11 }}>{report.detail}</p><button className="dw-btn primary" disabled={busy} onClick={() => mutate(report, "confirm-report")}>确认月报</button></div> : null}{reviewIssues.map((issue) => <a className="dw-row" key={issue.id} {...navigation.linkProps(hostPath(`/issues/${issue.id}`))}><strong>{issue.title}</strong><small>{issue.identifier ?? "任务"} · Paperclip 待审阅</small></a>)}{!pending.length && !reviewIssues.length && report?.status !== "draft" ? <div className="dw-muted">暂无待确认内容</div> : null}</div></section><section className="dw-panel"><header><h2>已处理</h2></header><div className="body">{visible.filter((entry) => entry.status === "rejected").slice(0, 20).map((entry) => <div className="dw-row" key={entry.id}><strong>{entry.title}</strong><small>已驳回 · {shortDate(entry.date)}</small></div>)}{!visible.some((entry) => entry.status === "rejected") ? <div className="dw-muted">暂无记录</div> : null}</div></section></div> : null}

    {view === "okr" ? <><div className="dw-head"><h2>OKR</h2><div className="dw-actions"><button className="dw-btn" onClick={() => showCreate("goal")}>新建目标</button><button className="dw-btn primary" onClick={() => showCreate("kr")}>新增 KR</button></div></div><div className="dw-stack">{goals.map((goal) => { const krs = visible.filter((entry) => entry.kind === "kr" && entry.goalId === goal.id && entry.status === "active"); return <section className="dw-panel" key={goal.id}><header><h2>{goal.title}</h2><span className="dw-pill">{goal.status}</span></header><div className="body">{krs.map((entry) => { const target = Number(entry.data.target ?? 0); const current = Number(entry.data.current ?? 0); return <div className="dw-kr" key={entry.id}><div><strong>{entry.title}</strong><small className="dw-muted">{projectName(data, entry.projectId)}</small></div><div className="dw-progress"><i style={{ width: `${target ? Math.min(100, current / target * 100) : 0}%` }} /></div><div className="dw-mini"><input aria-label={`${entry.title}进度`} type="number" step="any" min="0" max={target} value={progressValues[entry.id] ?? String(current)} onChange={(event) => setProgressValues({ ...progressValues, [entry.id]: event.target.value })} /><span className="dw-muted">/ {target} {String(entry.data.unit ?? "")}</span><button className="dw-btn" disabled={busy} onClick={() => mutate(entry, "check-in", { value: Number(progressValues[entry.id] ?? current) })}>更新</button><button className="dw-btn quiet" onClick={() => showEdit(entry)}>编辑</button></div></div>; })}{!krs.length ? <div className="dw-muted">暂无 KR</div> : null}<div style={{ marginTop: 11 }}><button className="dw-btn" onClick={() => showCreate("kr")}>新增 KR</button></div></div></section>; })}{!goals.length ? <section className="dw-panel"><div className="dw-empty">暂无目标。点击“新建目标”开始。</div></section> : null}</div></> : null}

    {view === "milestones" ? <section className="dw-panel"><header><h2>里程碑</h2><div className="dw-actions"><span className="dw-pill">{milestones.filter((entry) => entry.status === "done").length} / {milestones.length} 完成</span><button className="dw-btn primary" onClick={() => showCreate("milestone")}>新建里程碑</button></div></header><div className="dw-timeline">{milestones.slice().sort((a, b) => (a.date ?? "").localeCompare(b.date ?? "")).map((entry) => <div className={`dw-milestone ${entry.status === "done" ? "done" : ""}`} key={entry.id}><time>{shortDate(entry.date)}</time><i className="dot" aria-hidden="true" /><div><strong>{entry.title}</strong><small>{projectName(data, entry.projectId)} · {goalName(data, entry.goalId)}</small></div><div className="dw-actions"><button className="dw-btn" disabled={busy} onClick={() => mutate(entry, "milestone")}>{entry.status === "done" ? "重新打开" : "标记完成"}</button><button className="dw-btn quiet" onClick={() => showEdit(entry)}>编辑</button></div></div>)}{!milestones.length ? <div className="dw-empty">本月暂无里程碑。点击“新建里程碑”添加。</div> : null}</div></section> : null}

    {view === "knowledge" ? <><div className="dw-head"><h2>知识库</h2><button className="dw-btn primary" onClick={() => showCreate("document")}>新建文件</button></div><div className="dw-stats"><div className="dw-stat"><span>文件</span><strong>{documents.length}</strong></div><div className="dw-stat"><span>版本</span><strong>{documents.reduce((sum, entry) => sum + versions(entry).length, 0)}</strong></div><div className="dw-stat"><span>AI 协作文件</span><strong>{documents.filter((entry) => versions(entry).some((version) => version.authorWithAi)).length}</strong></div><div className="dw-stat"><span>本月版本</span><strong>{documents.flatMap(versions).filter((version) => version.date.startsWith(month)).length}</strong></div></div><div className="dw-cards">{documents.map((entry) => <article className="dw-doc" key={entry.id}><span className="dw-pill">{projectName(data, entry.projectId)}</span><h2>{entry.title}</h2><p>{entry.detail}</p>{versions(entry).slice().reverse().map((version) => <div className="dw-version" key={version.number}><strong>v{version.number}</strong><small>{shortDate(version.date)} · {version.authorWithAi ? "AI+人" : "人"} · {version.note}</small><details><summary>查看正文</summary><p>{version.content}</p></details></div>)}{!versions(entry).length ? <p className="dw-muted">暂无版本</p> : null}<div className="dw-card-actions"><button className="dw-btn" onClick={() => { setVersionDoc(versionDoc === entry.id ? "" : entry.id); setVersionContent(versions(entry).at(-1)?.content ?? ""); setVersionNote(""); setVersionAi(false); }}>新增版本</button><button className="dw-btn quiet" onClick={() => showEdit(entry)}>编辑</button></div>{versionDoc === entry.id ? <form className="dw-form" style={{ marginTop: 11 }} onSubmit={(event) => { event.preventDefault(); void mutate(entry, "version", { date: today(), content: versionContent, note: versionNote, authorWithAi: versionAi }).then((ok) => ok && setVersionDoc("")); }}><label className="wide">版本说明<input required value={versionNote} onChange={(event) => setVersionNote(event.target.value)} maxLength={120} /></label><label className="wide">正文<textarea required value={versionContent} onChange={(event) => setVersionContent(event.target.value)} maxLength={50000} /></label><label className="check"><input type="checkbox" checked={versionAi} onChange={(event) => setVersionAi(event.target.checked)} />AI 协作</label><footer><button className="dw-btn primary" type="submit" disabled={busy}>保存版本</button><button className="dw-btn quiet" type="button" onClick={() => setVersionDoc("")}>取消</button></footer></form> : null}</article>)}{!documents.length ? <div className="dw-empty">暂无文件。点击“新建文件”添加。</div> : null}</div></> : null}
  </div></>;
}

export function OfficePage(props: PluginPageProps) { return <Workbench {...props} />; }

export function DashboardWidget({ context }: PluginWidgetProps) {
  const navigation = useHostNavigation();
  const { pathname } = useHostLocation();
  const request = useMemo(() => ({ companyId: context.companyId ?? "" }), [context.companyId]);
  const query = usePluginData<OfficeSnapshot>("office", request);
  const data = query.data ?? empty;
  const entries = data.entries.filter((entry) => entry.status !== "archived");
  const pending = entries.filter((entry) => entry.status === "pending" || entry.status === "draft").length;
  return <><style>{css}</style><div className="dw dw-summary"><h2>部门工作台</h2><div className="dw-stats"><div className="dw-stat"><span>项目</span><strong>{data.projects.length}</strong></div><div className="dw-stat"><span>记事</span><strong>{entries.filter((entry) => ["meeting", "development", "training"].includes(entry.kind) && entry.status === "confirmed").length}</strong></div><div className="dw-stat"><span>里程碑</span><strong>{entries.filter((entry) => entry.kind === "milestone").length}</strong></div><div className="dw-stat"><span>待确认</span><strong>{pending}</strong></div></div><a className="dw-btn" {...navigation.linkProps(officePath(context.companyPrefix, pathname))}>打开工作台</a>{query.error ? <div className="dw-error">读取失败：{query.error.message}</div> : null}</div></>;
}

export function ProjectDetail({ context }: PluginDetailTabProps) {
  const navigation = useHostNavigation();
  const { pathname } = useHostLocation();
  const request = useMemo(() => ({ companyId: context.companyId ?? "" }), [context.companyId]);
  const query = usePluginData<OfficeSnapshot>("office", request);
  const data = query.data ?? empty;
  const entries = data.entries.filter((entry) => entry.projectId === context.entityId && entry.status !== "archived");
  const issues = data.issues.filter((issue) => issue.projectId === context.entityId);
  return <><style>{css}</style><div className="dw dw-project"><div className="dw-stats"><div className="dw-stat"><span>任务</span><strong>{issues.length}</strong></div><div className="dw-stat"><span>记事</span><strong>{entries.filter((entry) => ["meeting", "development", "training"].includes(entry.kind) && entry.status === "confirmed").length}</strong></div><div className="dw-stat"><span>里程碑</span><strong>{entries.filter((entry) => entry.kind === "milestone").length}</strong></div><div className="dw-stat"><span>文件</span><strong>{entries.filter((entry) => entry.kind === "document" && entry.status === "confirmed").length}</strong></div></div><div className="dw-two"><section className="dw-panel"><header><h2>任务</h2></header><div className="body">{issues.map((issue) => <div className="dw-row" key={issue.id}><strong>{issue.title}</strong><small>{issue.identifier ?? "任务"} · {issue.status}</small></div>)}{!issues.length ? <div className="dw-muted">暂无任务</div> : null}</div></section><section className="dw-panel"><header><h2>里程碑</h2></header><div className="body">{entries.filter((entry) => entry.kind === "milestone").map((entry) => <div className="dw-row" key={entry.id}><strong>{entry.title}</strong><small>{shortDate(entry.date)} · {entry.status === "done" ? "已完成" : "计划中"}</small></div>)}{!entries.some((entry) => entry.kind === "milestone") ? <div className="dw-muted">暂无里程碑</div> : null}<div style={{ marginTop: 12 }}><a className="dw-btn" {...navigation.linkProps(officePath(context.companyPrefix, pathname, `?project=${encodeURIComponent(context.entityId)}&view=milestones`))}>打开工作台</a></div></div></section></div>{query.error ? <div className="dw-error">读取失败：{query.error.message}</div> : null}</div></>;
}
