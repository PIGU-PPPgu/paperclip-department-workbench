import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { chmodSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import { test } from "node:test";

test("Codex wrapper fixes managed MCP auth and approves only project reads", (t) => {
  const home = mkdtempSync(join(tmpdir(), "paperclip-codex-test-"));
  t.after(() => rmSync(home, { recursive: true, force: true }));
  const configPath = join(home, "config.toml");
  writeFileSync(configPath, `
[mcp_servers."paperclip-projects"]
url = "http://127.0.0.1:3102/api/mcp/project-tools"
headers = { Authorization = "Bearer test-token" }

[mcp_servers."paperclip-connections"]
url = "http://127.0.0.1:3102/mcp/runtime-tools"
headers = { Authorization = "Bearer test-token" }

[mcp_servers.other]
headers = { Example = "unchanged" }
`, { mode: 0o600 });
  const fakeCodex = join(home, "codex.mjs");
  writeFileSync(fakeCodex, `#!/usr/bin/env node\nconsole.log(JSON.stringify(process.argv.slice(2)));\n`);
  chmodSync(fakeCodex, 0o700);
  const wrapper = resolve("scripts/paperclip-codex-cli.mjs");
  for (let i = 0; i < 2; i += 1) {
    const run = spawnSync(wrapper, ["exec", "--json"], {
      encoding: "utf8",
      env: { ...process.env, CODEX_HOME: home, PAPERCLIP_CODEX_BIN: fakeCodex }
    });
    assert.equal(run.status, 0, run.stderr);
    assert.deepEqual(JSON.parse(run.stdout.trim()), ["exec", "--json"]);
  }
  const config = readFileSync(configPath, "utf8");
  assert.equal((config.match(/http_headers = \{ Authorization = "Bearer test-token" \}/g) ?? []).length, 2);
  assert.match(config, /\[mcp_servers\.other\]\nheaders = \{ Example = "unchanged" \}/);
  assert.equal((config.match(/approval_mode = "approve"/g) ?? []).length, 1);
  assert.match(config, /\[mcp_servers\."paperclip-projects"\.tools\.list_projects\]\napproval_mode = "approve"/);
  assert.doesNotMatch(config, /tools\.(?:create_project|create_task)/);
});
