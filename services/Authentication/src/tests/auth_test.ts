// src/tests/auth_test.ts
//
// Manual Entra ID lookup tool. Type a phone number in the terminal and see
// everything Microsoft Graph returns for the matching employee — this does
// NOT go through VAPI or the webhook, it calls Graph directly.
//
// Usage:
//   npm run build
//   npm run test:caller
// (or: node dist/tests/auth_test.js)

import * as fs from "fs";
import * as path from "path";
import * as readline from "readline";
import { getGraphAccessToken } from "../lib/graphAuth";
import { buildPhoneVariants, escapeODataValue } from "../lib/callerLookup";

// Falls back to local.settings.json (the file `func start` normally reads)
// or a .env file if TENANT_ID/CLIENT_ID/CLIENT_SECRET aren't already in the
// environment — needed because this script runs as a plain node process,
// not through the Azure Functions host.
function loadEnvFileIfNeeded(): void {
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

// A wide $select — every standard Graph user field useful for confirming
// who a caller is, not just the id/displayName the production lookup uses.
const SELECT_FIELDS = [
  "id",
  "displayName",
  "givenName",
  "surname",
  "userPrincipalName",
  "mail",
  "mobilePhone",
  "businessPhones",
  "jobTitle",
  "department",
  "companyName",
  "officeLocation",
  "employeeId",
  "city",
  "country",
  "accountEnabled",
].join(",");

async function fetchFullUserRecords(
  rawNumber: string,
  token: string
): Promise<Record<string, unknown>[]> {
  const variants = buildPhoneVariants(rawNumber);

  const filterParts: string[] = [];
  for (const variant of variants) {
    const escaped = escapeODataValue(variant);
    filterParts.push(`mobilePhone eq '${escaped}'`);
    filterParts.push(`businessPhones/any(p:p eq '${escaped}')`);
  }
  const filter = filterParts.join(" or ");

  const url = `https://graph.microsoft.com/v1.0/users?$filter=${encodeURIComponent(
    filter
  )}&$select=${SELECT_FIELDS}&$count=true`;

  const response = await fetch(url, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${token}`,
      ConsistencyLevel: "eventual",
    },
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Graph API error (${response.status}): ${errorText}`);
  }

  const data = (await response.json()) as { value: Record<string, unknown>[] };
  return data.value ?? [];
}

function printUser(user: Record<string, unknown>, index: number): void {
  console.log(`\n--- Match ${index + 1} ---`);
  for (const [key, value] of Object.entries(user)) {
    if (key.startsWith("@")) continue; // odata metadata, not a real field
    const printedValue = Array.isArray(value)
      ? value.join(", ") || "(none)"
      : value ?? "(none)";
    console.log(`  ${key.padEnd(18)}: ${printedValue}`);
  }
}

async function runLookup(rawNumber: string): Promise<void> {
  const token = await getGraphAccessToken();
  if (!token) {
    console.error(
      "Could not acquire a Graph access token — check TENANT_ID/CLIENT_ID/CLIENT_SECRET."
    );
    return;
  }

  try {
    const users = await fetchFullUserRecords(rawNumber, token);

    if (users.length === 0) {
      console.log(`\nNo Entra ID user matches "${rawNumber}".`);
      return;
    }

    console.log(`\nFound ${users.length} match(es) for "${rawNumber}":`);
    users.forEach(printUser);
  } catch (error) {
    console.error(`Lookup failed: ${(error as Error).message}`);
  }
}

function main(): void {
  loadEnvFileIfNeeded();

  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  console.log("Entra ID caller lookup — type a phone number, or 'exit' to quit.\n");

  const ask = (): void => {
    rl.question("Phone number: ", async (answer) => {
      const trimmed = answer.trim();
      if (trimmed.toLowerCase() === "exit" || trimmed.toLowerCase() === "quit") {
        rl.close();
        return;
      }
      if (trimmed) {
        await runLookup(trimmed);
      }
      ask();
    });
  };

  ask();
}

main();
