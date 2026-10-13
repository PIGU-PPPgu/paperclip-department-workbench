#!/usr/bin/env node

/**
 * Paperclip v2026.916.1 writes `headers` for its managed HTTP MCP gateways.
 * Codex CLI 0.158 reads `http_headers`. Translate only Paperclip's managed
 * gateways after Paperclip has prepared CODEX_HOME and before Codex starts.
 * This lives outside the managed Paperclip installation so upgrades preserve it.
 */

import { spawn } from "node:child_process";
import { readFileSync, renameSync, unlinkSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { randomUUID } from "node:crypto";

const home = process.env.CODEX_HOME;
const codex = process.env.PAPERCLIP_CODEX_BIN;
if (!home || !codex) {
  console.error("Paperclip Codex wrapper requires CODEX_HOME and PAPERCLIP_CODEX_BIN");
  process.exit(2);
}

const configPath = join(home, "config.toml");
const before = readFileSync(configPath, "utf8");
let section = "";
let changed = false;
let after = before.split("\n").map((line) => {
  const heading = line.match(/^\s*\[([^\]]+)\]\s*$/);
  if (heading) section = heading[1].replaceAll('"', "").replaceAll("'", "");
  if (!/^mcp_servers\.paperclip(?:-|$)/.test(section)) return line;
  if (!/^\s*headers\s*=\s*\{\s*Authorization\s*=/.test(line)) return line;
  changed = true;
  return line.replace(/^(\s*)headers(\s*=)/, "$1http_headers$2");
}).join("\n");

// Paperclip's list_projects is a GET-only operation. Its MCP descriptor in
// v2026.916.1 has no readOnlyHint, so headless Codex otherwise asks for an
// approval that cannot be answered. Approve only this verified read tool.
if (/^\s*\[mcp_servers\.(?:"paperclip-projects"|paperclip-projects)\]/m.test(after) &&
    !/^\s*\[mcp_servers\.(?:"paperclip-projects"|paperclip-projects)\.tools\.list_projects\]/m.test(after)) {
  after += '\n[mcp_servers."paperclip-projects".tools.list_projects]\napproval_mode = "approve"\n';
  changed = true;
}

if (changed) {
  const tempPath = join(dirname(configPath), `.config.toml.${randomUUID()}.tmp`);
  try {
    writeFileSync(tempPath, after, { mode: 0o600, flag: "wx" });
    renameSync(tempPath, configPath);
  } catch (error) {
    try { unlinkSync(tempPath); } catch { /* The temporary file may not exist. */ }
    console.error(`Failed to prepare Codex MCP configuration: ${error.message}`);
    process.exit(2);
  }
}

const child = spawn(codex, process.argv.slice(2), { stdio: "inherit", env: process.env });
for (const signal of ["SIGINT", "SIGTERM"]) {
  process.on(signal, () => child.kill(signal));
}
child.on("error", (error) => {
  console.error(`Failed to start Codex CLI: ${error.message}`);
  process.exitCode = 127;
});
child.on("exit", (code, signal) => {
  process.exitCode = code ?? (signal ? 128 + (signal === "SIGINT" ? 2 : 15) : 1);
});
