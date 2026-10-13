# Paperclip 部门工作台

面向部门的 Paperclip 原生插件。它在宿主侧栏、仪表盘、项目详情和独立页面中显示工作记录，并通过官方 Plugin SDK 读取项目、任务和 Goal。业务记录保存在插件自己的 PostgreSQL 命名空间，不修改 Paperclip 核心页面或核心表。

## 功能

- **行事历**：按月和项目查看已确认的会议、开发、培训及计划里程碑。
- **记事**：新建会议、开发和培训记录；保存会议级别、参与人、GitHub 链接和培训场次。可提交待确认，确认后才计入正式统计。
- **OKR**：新建 Paperclip Goal，在插件中创建关联 KR、设置目标值并记录每次进度检查。
- **里程碑**：关联项目和可选 Goal，创建节点并记录实际完成日期。
- **知识库**：创建带正文的文件，逐版保存内容、版本说明和 AI 协作标记。
- **月报**：按月份和项目汇总已确认的业务记录，支持人工编辑、来源变更校验及确认。
- **智能体状态**：总览显示待命、运行中、暂停及手动/定时运行方式，列出关联的受阻任务，并跳转到 Paperclip 原生智能体与任务页面。

进入 `/:companyPrefix/office`。顶部的项目、月份是筛选条件；各视图有自己的“新建”入口。项目本身仍由 Paperclip 的原生项目页管理。页面没有演示数据，也不会自动创建业务记录。

## 权限与数据

写操作从 Paperclip 宿主获得公司与操作者身份。Agent 和仅查看成员不能通过插件动作写入；记录保存创建者、更新者、时间与修订号，更新时检查修订冲突。GitHub 地址只接受 `https://github.com/`，插件不会向 GitHub 发送 API 请求。知识库正文存放在插件表中，目前不自动同步 Paperclip Wiki 或 NAS。

迁移在宿主派生的插件命名空间中运行。`migrations/001_office_entries.sql` 创建插件表和索引；已经应用的迁移不得修改。新增迁移时请添加下一个编号的 SQL 文件。安装前备份 Paperclip 数据库。上游更新后重跑构建、插件健康和页面验收；插件不依赖 Paperclip `ui/src` 私有代码。

## 开发

要求 Node.js 24+，依赖锁定 `@paperclipai/plugin-sdk@2026.916.1`。

```sh
npm ci
npm run typecheck
npm test
npm run build
npm audit --omit=dev --audit-level=high
```

本地安装时先确认目标实例，再通过官方 CLI 安装：

```sh
paperclipai plugin target --api-base http://127.0.0.1:3100
paperclipai plugin install . --local --api-base http://127.0.0.1:3100
```

本机宿主在增加插件 capability 时拒绝直接热升级；用**不带 `--force`** 的软卸载和本地重装可保留同一插件 ID 与插件表。能力集不变时可用 `paperclipai plugin upgrade paperclip.department-workbench`。不要使用 `--force`，它会清除插件状态。当前宿主版本的安装器会误判 `minimumHostVersion`，因此 manifest 暂不设置这个字段；请以 `plugin target` 和兼容性验收确认目标版本。

## 当前边界

这是可录入的第一版。文件正文在插件表中，尚无附件上传、Wiki 双向同步、GitHub 自动统计、提醒通知和细粒度项目级编辑授权；这些能力需要单独设计与验收。插件只为已授权的公司成员提供操作入口，不把 AI 草稿自动算作正式成果。

MIT License.
