CREATE TABLE plugin_department_workbench_d1974d3b8a.entries (
  id uuid PRIMARY KEY,
  company_id uuid NOT NULL REFERENCES public.companies(id) ON DELETE CASCADE,
  project_id uuid REFERENCES public.projects(id) ON DELETE SET NULL,
  goal_id uuid REFERENCES public.goals(id) ON DELETE SET NULL,
  kind text NOT NULL CHECK (kind IN ('meeting', 'development', 'training', 'milestone', 'kr', 'document', 'report')),
  status text NOT NULL,
  occurred_on date,
  title text NOT NULL,
  detail text NOT NULL DEFAULT '',
  data jsonb NOT NULL DEFAULT '{}'::jsonb,
  revision integer NOT NULL DEFAULT 1,
  created_by_type text NOT NULL,
  created_by_id text,
  updated_by_type text NOT NULL,
  updated_by_id text,
  created_at timestamptz NOT NULL DEFAULT now(),
  updated_at timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX office_entries_company_kind_date ON plugin_department_workbench_d1974d3b8a.entries (company_id, kind, occurred_on);
CREATE INDEX office_entries_company_project ON plugin_department_workbench_d1974d3b8a.entries (company_id, project_id);
CREATE INDEX office_entries_company_goal ON plugin_department_workbench_d1974d3b8a.entries (company_id, goal_id);
CREATE UNIQUE INDEX office_entries_report_scope ON plugin_department_workbench_d1974d3b8a.entries (company_id, occurred_on, COALESCE(project_id, '00000000-0000-0000-0000-000000000000'::uuid)) WHERE kind = 'report' AND status <> 'archived';
