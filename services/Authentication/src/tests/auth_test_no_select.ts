// src/tests/auth_test_no_select.ts
//
// Same manual Entra ID lookup as auth_test.ts, but the Graph query has NO
// $select at all — this exists to show, side by side, that dropping $select
// does NOT return "all" user fields. Graph falls back to its own small
// default set of properties instead.
//
// Usage:
//   npm run build
//   node dist/tests/auth_test_no_select.js

import * as readline from "readline";
import { getGraphAccessToken } from "../lib/graphAuth";
import { buildPhoneVariants, escapeODataValue } from "../lib/callerLookup";
import { loadEnvFileIfNeeded } from "./loadLocalEnv";

async function fetchUserRecords(
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

  // No $select here on purpose.
  const url = `https://graph.microsoft.com/v1.0/users?$filter=${encodeURIComponent(
    filter
  )}&$count=true`;

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
    const users = await fetchUserRecords(rawNumber, token);

    if (users.length === 0) {
      console.log(`\nNo Entra ID user matches "${rawNumber}".`);
      return;
    }

    console.log(`\nFound ${users.length} match(es) for "${rawNumber}" (no $select applied):`);
    users.forEach(printUser);
  } catch (error) {
    console.error(`Lookup failed: ${(error as Error).message}`);
  }
}

function main(): void {
  loadEnvFileIfNeeded();

  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  console.log("Entra ID caller lookup (NO $select) — type a phone number, or 'exit' to quit.\n");

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
