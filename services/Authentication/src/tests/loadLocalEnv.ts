// src/tests/loadLocalEnv.ts
//
// Shared helper for the manual test scripts in this folder, which run as
// plain node processes (not through the Azure Functions host), so they
// don't get TENANT_ID/CLIENT_ID/CLIENT_SECRET for free the way `func start`
// does.

import * as fs from "fs";
import * as path from "path";

// Falls back to local.settings.json (the file `func start` normally reads)
// or a .env file if TENANT_ID/CLIENT_ID/CLIENT_SECRET aren't already in the
// environment.
export function loadEnvFileIfNeeded(): void {
  if (process.env.TENANT_ID && process.env.CLIENT_ID && process.env.CLIENT_SECRET) {
    return;
  }

  const root = path.join(__dirname, "..", "..");
  loadFromLocalSettingsJson(path.join(root, "local.settings.json"));
  loadFromDotEnv(path.join(root, ".env"));
}

function loadFromLocalSettingsJson(settingsPath: string): void {
  if (!fs.existsSync(settingsPath)) return;

  const parsed = JSON.parse(fs.readFileSync(settingsPath, "utf-8")) as {
    Values?: Record<string, string>;
  };

  for (const [key, value] of Object.entries(parsed.Values ?? {})) {
    if (key && !process.env[key]) process.env[key] = value;
  }
}

function loadFromDotEnv(envPath: string): void {
  if (!fs.existsSync(envPath)) return;

  const lines = fs.readFileSync(envPath, "utf-8").split(/\r?\n/);
  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) continue;
    const eq = trimmed.indexOf("=");
    if (eq === -1) continue;
    const key = trimmed.slice(0, eq).trim();
    const value = trimmed.slice(eq + 1).trim();
    if (key && !process.env[key]) process.env[key] = value;
  }
}
