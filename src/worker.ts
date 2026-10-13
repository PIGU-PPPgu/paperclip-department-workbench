import { randomUUID } from "node:crypto";
import { definePlugin, runWorker } from "@paperclipai/plugin-sdk";
import { date, finiteNumber, githubUrls, month, object, reportFingerprint, reportText, stringList, text, type DocumentVersion, type EntryKind, type OfficeEntry, type OfficeSnapshot } from "./office.js";

export type { OfficeSnapshot } from "./office.js";

type Row = Record<string, unknown>;
const uuid = (value: unknown, label = "ID") => {
  if (typeof value !== "string" || !/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(value)) throw new Error(`${label}无效`);
  return value;
};
const optionalUuid = (value: unknown, label: string): string | null => value === "" || value === null || value === undefined ? null : uuid(value, label);
const localDate = () => new Intl.DateTimeFormat("en-CA", { timeZone: "Asia/Shanghai", year: "numeric", month: "2-digit", day: "2-digit" }).format(new Date());
const locks = new Map<string, Promise<void>>();
async function locked<T>(companyId: string, run: () => Promise<T>): Promise<T> {
  const previous = locks.get(companyId) ?? Promise.resolve();
  let release!: () => void;
  const gate = new Promise<void>((resolve) => { release = resolve; });
  const tail = previous.then(() => gate);
  locks.set(companyId, tail);
  await previous;
  try { return await run(); }
  finally { release(); if (locks.get(companyId) === tail) locks.delete(companyId); }
}
const columns = `id::text AS id, company_id::text AS "companyId", project_id::text AS "projectId", goal_id::text AS "goalId", kind, status, occurred_on::text AS date, title, detail, data, revision, created_by_type AS "createdByType", created_by_id AS "createdById", updated_by_type AS "updatedByType", updated_by_id AS "updatedById", created_at::text AS "createdAt", updated_at::text AS "updatedAt"`;

const plugin = definePlugin({
  async setup(ctx) {
    const table = `${ctx.db.namespace}.entries`;
    const listEntries = async (companyId: string): Promise<OfficeEntry[]> => {
      const rows = await ctx.db.query<Row>(`SELECT ${columns} FROM ${table} WHERE company_id = $1::uuid ORDER BY created_at DESC LIMIT 5001`, [companyId]);
      if (rows.length > 5000) throw new Error("工作台记录超过当前读取上限，请联系管理员扩容");
      return rows as OfficeEntry[];
    };
    const one = async (companyId: string, id: string): Promise<OfficeEntry> => {
      const rows = await ctx.db.query<Row>(`SELECT ${columns} FROM ${table} WHERE company_id = $1::uuid AND id = $2::uuid LIMIT 1`, [companyId, uuid(id)]);
      if (!rows[0]) throw new Error("记录不存在");
      return rows[0] as OfficeEntry;
    };
    const native = async (companyId: string): Promise<Omit<OfficeSnapshot, "entries">> => {
      const [projects, issues, goals, agents] = await Promise.all([
        ctx.projects.list({ companyId, limit: 200 }),
        ctx.issues.list({ companyId, includePluginOperations: false, limit: 200 }),
        ctx.goals.list({ companyId, limit: 200 }),
        ctx.agents.list({ companyId, limit: 200 }),
      ]);
      return {
        projects: projects.filter((project) => !project.archivedAt && !project.managedByPlugin).map((project) => ({ id: project.id, name: project.name, status: project.status })),
        issues: issues.map((issue) => ({ id: issue.id, identifier: issue.identifier ?? null, title: issue.title, status: issue.status, projectId: issue.projectId ?? null, assigneeAgentId: issue.assigneeAgentId ?? null, updatedAt: new Date(issue.updatedAt).toISOString() })),
        goals: goals.map((goal) => ({ id: goal.id, title: goal.title, status: goal.status })),
        agents: agents.map((agent) => { const heartbeat = agent.runtimeConfig.heartbeat; return { id: agent.id, urlKey: agent.urlKey, name: agent.name, status: agent.status, scheduled: !!heartbeat && typeof heartbeat === "object" && (heartbeat as Record<string, unknown>).enabled === true, lastHeartbeatAt: agent.lastHeartbeatAt ? new Date(agent.lastHeartbeatAt).toISOString() : null }; }),
      };
    };
    const checkProject = async (companyId: string, projectId: string | null) => {
      if (!projectId) return;
      const project = await ctx.projects.get(projectId, companyId);
      if (!project || project.archivedAt || project.managedByPlugin) throw new Error("项目不存在或不可用");
    };
    const checkGoal = async (companyId: string, goalId: string | null) => {
      if (!goalId) return;
      if (!await ctx.goals.get(goalId, companyId)) throw new Error("目标不存在");
    };
    const scope = async (params: Record<string, unknown>, actionContext: { companyId: string | null; actor: { type: string; userId: string | null; agentId: string | null } }) => {
      const companyId = actionContext.companyId;
      if (!companyId || (params.companyId !== undefined && params.companyId !== companyId)) throw new Error("公司范围无效");
      if (actionContext.actor.type === "agent") throw new Error("工作台编辑只允许人工操作");
      if (actionContext.actor.type === "user") {
        const members = await ctx.access.members.list({ companyId });
        const member = members.find((item) => item.principalType === "user" && item.principalId === actionContext.actor.userId && item.status === "active");
        if (!member || member.membershipRole === "viewer") throw new Error("当前账号没有工作台编辑权限");
      }
      return { companyId, actorType: actionContext.actor.type, actorId: actionContext.actor.userId };
    };
    const insert = async (input: { companyId: string; projectId: string | null; goalId: string | null; kind: EntryKind; status: string; date: string | null; title: string; detail: string; data: Record<string, unknown>; actorType: string; actorId: string | null }) => {
      const id = randomUUID();
      await ctx.db.execute(`INSERT INTO ${table} (id, company_id, project_id, goal_id, kind, status, occurred_on, title, detail, data, created_by_type, created_by_id, updated_by_type, updated_by_id) VALUES ($1::uuid, $2::uuid, $3::uuid, $4::uuid, $5, $6, $7::date, $8, $9, $10::jsonb, $11, $12, $11, $12)`, [id, input.companyId, input.projectId, input.goalId, input.kind, input.status, input.date, input.title, input.detail, JSON.stringify(input.data), input.actorType, input.actorId]);
      return one(input.companyId, id);
    };
    const update = async (entry: OfficeEntry, input: { projectId?: string | null; goalId?: string | null; date?: string | null; title?: string; detail?: string; status?: string; data?: Record<string, unknown> }, actorType: string, actorId: string | null) => {
      const result = await ctx.db.execute(`UPDATE ${table} SET project_id = $1::uuid, goal_id = $2::uuid, occurred_on = $3::date, title = $4, detail = $5, status = $6, data = $7::jsonb, revision = revision + 1, updated_by_type = $8, updated_by_id = $9, updated_at = now() WHERE company_id = $10::uuid AND id = $11::uuid AND revision = $12`, [Object.hasOwn(input, "projectId") ? input.projectId : entry.projectId, Object.hasOwn(input, "goalId") ? input.goalId : entry.goalId, Object.hasOwn(input, "date") ? input.date : entry.date, input.title ?? entry.title, input.detail ?? entry.detail, input.status ?? entry.status, JSON.stringify(input.data ?? entry.data), actorType, actorId, entry.companyId, entry.id, entry.revision]);
      if (result.rowCount !== 1) throw new Error("记录已被修改，请刷新后重试");
      return one(entry.companyId, entry.id);
    };

    ctx.data.register("office", async (params): Promise<OfficeSnapshot> => {
      const companyId = typeof params.companyId === "string" ? params.companyId : "";
      if (!companyId) return { projects: [], issues: [], goals: [], agents: [], entries: [] };
      const [base, entries] = await Promise.all([native(companyId), listEntries(companyId)]);
      return { ...base, entries };
    });

    ctx.actions.register("office-create", async (params, actionContext) => {
      const { companyId, actorType, actorId } = await scope(params, actionContext);
      const input = object(params.entry);
      const kind = input.kind;
      if (!["meeting", "development", "training", "milestone", "kr", "document"].includes(String(kind))) throw new Error("类型无效");
      const projectId = optionalUuid(input.projectId, "项目 ID");
      const goalId = optionalUuid(input.goalId, "目标 ID");
      await checkProject(companyId, projectId);
      await checkGoal(companyId, goalId);
      if (kind === "kr" && !goalId) throw new Error("请选择目标");
      if (["meeting", "development", "training", "milestone", "document"].includes(String(kind)) && !projectId) throw new Error("请选择项目");
      const title = text(input.title, "标题", 120);
      const detail = text(input.detail ?? "", "内容", 20000, false);
      const when = kind === "kr" ? null : date(input.date ?? localDate());
      const raw = object(input.data ?? {});
      let data: Record<string, unknown> = {};
      let status = "confirmed";
      if (["meeting", "development", "training"].includes(String(kind))) {
        status = input.pending === true ? "pending" : "confirmed";
        data = {
          level: kind === "meeting" ? text(raw.level ?? "项目组", "会议级别", 40) : "",
          sessions: kind === "training" ? finiteNumber(raw.sessions ?? 1, "培训场次", 1, 1000) : 0,
          contributors: stringList(raw.contributors, "参与人", 30, 80),
          repoUrls: kind === "development" ? githubUrls(raw.repoUrls) : [],
        };
      } else if (kind === "milestone") {
        status = "planned";
      } else if (kind === "kr") {
        status = "active";
        const target = finiteNumber(raw.target, "目标值", 0.000001, 1_000_000_000);
        const current = finiteNumber(raw.current ?? 0, "当前值", 0, target);
        data = { target, current, unit: text(raw.unit ?? "项", "单位", 20), checkins: [{ date: localDate(), value: current, authorId: actorId }] };
      } else if (kind === "document") {
        status = input.pending === true ? "pending" : "confirmed";
        const content = text(raw.content ?? "", "正文", 50000, false);
        const note = text(raw.note ?? "初版", "版本说明", 120);
        data = { versions: content ? [{ number: 1, date: when, note, content, authorWithAi: raw.authorWithAi === true, authorId: actorId }] : [] };
      }
      return insert({ companyId, projectId, goalId, kind: kind as EntryKind, status, date: when, title, detail, data, actorType, actorId });
    });

    ctx.actions.register("office-goal", async (params, actionContext) => {
      const { companyId } = await scope(params, actionContext);
      const title = text(params.title, "目标", 160);
      const description = text(params.description ?? "", "说明", 2000, false);
      return ctx.goals.create({ companyId, title, description, level: "company", status: "active" });
    });

    ctx.actions.register("office-change", async (params, actionContext) => {
      const { companyId, actorType, actorId } = await scope(params, actionContext);
      const id = uuid(params.id);
      const revision = finiteNumber(params.revision, "版本", 1, 1_000_000);
      return locked(companyId, async () => {
        const entry = await one(companyId, id);
        if (entry.revision !== revision) throw new Error("记录已被修改，请刷新后重试");
        const operation = params.operation;
        if (operation === "confirm" || operation === "reject") {
          if (!["meeting", "development", "training", "document"].includes(entry.kind) || entry.status !== "pending") throw new Error("这条记录不在待确认状态");
          return update(entry, { status: operation === "confirm" ? "confirmed" : "rejected" }, actorType, actorId);
        }
        if (operation === "milestone") {
          if (entry.kind !== "milestone") throw new Error("不是里程碑");
          const done = entry.status !== "done";
          const data = { ...entry.data };
          if (done) data.completedAt = localDate(); else delete data.completedAt;
          return update(entry, { status: done ? "done" : "planned", data }, actorType, actorId);
        }
        if (operation === "check-in") {
          if (entry.kind !== "kr") throw new Error("不是 KR");
          const target = Number(entry.data.target);
          const current = finiteNumber(params.value, "当前值", 0, target);
          const checkins = Array.isArray(entry.data.checkins) ? entry.data.checkins : [];
          return update(entry, { data: { ...entry.data, current, checkins: [...checkins, { date: localDate(), value: current, authorId: actorId }] } }, actorType, actorId);
        }
        if (operation === "version") {
          if (entry.kind !== "document" || entry.status !== "confirmed") throw new Error("文件尚未确认");
          const versions = Array.isArray(entry.data.versions) ? entry.data.versions as DocumentVersion[] : [];
          const next: DocumentVersion = { number: versions.length + 1, date: date(params.date ?? localDate()), note: text(params.note, "版本说明", 120), content: text(params.content, "正文", 50000), authorWithAi: params.authorWithAi === true, authorId: actorId };
          return update(entry, { data: { ...entry.data, versions: [...versions, next] } }, actorType, actorId);
        }
        if (operation === "edit") {
          if (entry.kind === "report") {
            if (entry.status !== "draft") throw new Error("已确认月报不可直接修改");
            return update(entry, { detail: text(params.detail, "月报正文", 20000) }, actorType, actorId);
          }
          const title = text(params.title, "标题", 120);
          const detail = text(params.detail ?? "", "内容", 20000, false);
          const when = entry.kind === "kr" ? null : date(params.date ?? entry.date);
          const projectId = optionalUuid(params.projectId ?? entry.projectId, "项目 ID");
          await checkProject(companyId, projectId);
          const goalId = optionalUuid(params.goalId ?? entry.goalId, "目标 ID");
          await checkGoal(companyId, goalId);
          const raw = object(params.data ?? entry.data);
          let data = entry.data;
          if (["meeting", "development", "training"].includes(entry.kind)) data = {
            level: entry.kind === "meeting" ? text(raw.level ?? "项目组", "会议级别", 40) : "",
            sessions: entry.kind === "training" ? finiteNumber(raw.sessions ?? 1, "培训场次", 1, 1000) : 0,
            contributors: stringList(raw.contributors, "参与人", 30, 80),
            repoUrls: entry.kind === "development" ? githubUrls(raw.repoUrls) : [],
          };
          if (entry.kind === "kr") {
            const target = finiteNumber(raw.target, "目标值", 0.000001, 1_000_000_000);
            const current = finiteNumber(raw.current ?? 0, "当前值", 0, target);
            const checkins = Array.isArray(entry.data.checkins) ? entry.data.checkins : [];
            data = { target, current, unit: text(raw.unit ?? "项", "单位", 20), checkins: current === entry.data.current ? checkins : [...checkins, { date: localDate(), value: current, authorId: actorId }] };
          }
          return update(entry, { title, detail, date: when, projectId, goalId, data }, actorType, actorId);
        }
        if (operation === "archive") {
          if (entry.status === "archived") throw new Error("记录已归档");
          return update(entry, { status: "archived" }, actorType, actorId);
        }
        if (operation === "confirm-report") {
          if (entry.kind !== "report" || entry.status !== "draft") throw new Error("月报不在草稿状态");
          const period = month(entry.date?.slice(0, 7));
          const entries = await listEntries(companyId);
          if (entry.data.fingerprint !== reportFingerprint(entries, period, entry.projectId ?? "")) throw new Error("来源数据已变化，请重新生成月报");
          return update(entry, { status: "confirmed" }, actorType, actorId);
        }
        throw new Error("不支持的操作");
      });
    });

    ctx.actions.register("office-report", async (params, actionContext) => {
      const { companyId, actorType, actorId } = await scope(params, actionContext);
      const period = month(params.month);
      const projectId = optionalUuid(params.projectId, "项目 ID");
      await checkProject(companyId, projectId);
      return locked(companyId, async () => {
        const entries = await listEntries(companyId);
        const projectName = projectId ? (await ctx.projects.get(projectId, companyId))?.name ?? "" : "";
        const body = reportText(entries, period, projectId ?? "", projectName);
        const fingerprint = reportFingerprint(entries, period, projectId ?? "");
        const existing = entries.find((entry) => entry.kind === "report" && entry.date === `${period}-01` && entry.projectId === projectId && entry.status !== "archived");
        if (existing) {
          const history = Array.isArray(existing.data.history) ? existing.data.history : [];
          return update(existing, { status: "draft", detail: body, data: { ...existing.data, fingerprint, history: [...history, { revision: existing.revision, content: existing.detail, status: existing.status, at: existing.updatedAt }] } }, actorType, actorId);
        }
        return insert({ companyId, projectId, goalId: null, kind: "report", status: "draft", date: `${period}-01`, title: `${period} 月报`, detail: body, data: { fingerprint, history: [] }, actorType, actorId });
      });
    });

  },
  async onHealth() { return { status: "ok" }; },
});

export default plugin;
runWorker(plugin, import.meta.url);
