import assert from "node:assert/strict";
import test from "node:test";
import { date, githubUrls, krValueAsOf, reportFingerprint, reportText, type OfficeEntry } from "../src/office.ts";

const entry = (patch: Partial<OfficeEntry>): OfficeEntry => ({
  id: "00000000-0000-4000-8000-000000000001", companyId: "company", projectId: "project-a", goalId: null,
  kind: "meeting", status: "confirmed", date: "2026-10-13", title: "部门例会", detail: "完成评审", data: { level: "部门级" },
  revision: 1, createdByType: "user", createdById: "user", updatedByType: "user", updatedById: "user",
  createdAt: "2026-10-13T01:00:00Z", updatedAt: "2026-10-13T01:00:00Z", ...patch,
});

test("real entry validation rejects invalid dates and non GitHub URLs", () => {
  assert.equal(date("2026-10-13"), "2026-10-13");
  assert.throws(() => date("2026-02-31"), /日期无效/);
  assert.deepEqual(githubUrls(["https://github.com/PIGU-PPPgu/repo"]), ["https://github.com/PIGU-PPPgu/repo"]);
  assert.throws(() => githubUrls(["https://github.com.evil.test/repo"]), /GitHub/);
});

test("report lists confirmed monthly work and excludes pending proposals", () => {
  const rows = [
    entry({}),
    entry({ id: "00000000-0000-4000-8000-000000000002", status: "pending", title: "待确认会议" }),
    entry({ id: "00000000-0000-4000-8000-000000000003", kind: "training", title: "应用培训", data: { sessions: 2 } }),
    entry({ id: "00000000-0000-4000-8000-000000000004", kind: "milestone", status: "done", title: "一期验收", data: { completedAt: "2026-10-13" } }),
    entry({ id: "00000000-0000-4000-8000-000000000005", kind: "document", status: "confirmed", data: { versions: [{ number: 1, date: "2026-10-13", note: "初版", content: "正文", authorWithAi: false, authorId: "user" }] } }),
    entry({ id: "00000000-0000-4000-8000-000000000006", kind: "kr", status: "active", date: null, data: { current: 2, target: 4, unit: "项", checkins: [{ date: "2026-10-13", value: 2, authorId: "user" }] }, title: "交付项" }),
  ];
  const report = reportText(rows, "2026-10", "project-a", "部门项目");
  assert.match(report, /会议 1 场，开发 0 项，培训 2 场/);
  assert.match(report, /里程碑：2026-10-13 一期验收/);
  assert.match(report, /KR：交付项 2\/4项/);
  assert.doesNotMatch(report, /待确认会议/);
});

test("report fingerprint changes for relevant work only", () => {
  const base = [entry({})];
  const original = reportFingerprint(base, "2026-10", "project-a");
  assert.notEqual(original, reportFingerprint([entry({ revision: 2, title: "修订会议" })], "2026-10", "project-a"));
  assert.equal(original, reportFingerprint([...base, entry({ id: "00000000-0000-4000-8000-000000000002", projectId: "project-b" })], "2026-10", "project-a"));
  assert.equal(original, reportFingerprint([...base, entry({ id: "00000000-0000-4000-8000-000000000003", date: "2026-11-01" })], "2026-10", "project-a"));
});

test("historical KR value is stable after a later check-in", () => {
  const base = entry({ kind: "kr", status: "active", date: null, data: { current: 2, target: 4, unit: "项", checkins: [{ date: "2026-10-13", value: 2, authorId: "user" }] } });
  const later = entry({ ...base, revision: 2, data: { ...base.data, current: 3, checkins: [...base.data.checkins as unknown[], { date: "2026-11-01", value: 3, authorId: "user" }] } });
  assert.equal(krValueAsOf(later, "2026-10"), 2);
  assert.equal(reportFingerprint([base], "2026-10"), reportFingerprint([later], "2026-10"));
});
