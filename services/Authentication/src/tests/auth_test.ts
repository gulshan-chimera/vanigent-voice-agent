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

import * as readline from "readline";
import { getGraphAccessToken } from "../lib/graphAuth";
import { buildPhoneVariants, escapeODataValue } from "../lib/callerLookup";
import { loadEnvFileIfNeeded } from "./loadLocalEnv";

// A wide $select — every standard Graph user field useful for confirming
// who a caller is, not just the id/displayName the production lookup uses.
const SELECT_FIELDS = [
  "id",
  "displayName",
  "givenName",
  "surname",
  "userPrincipalName",
  "mail",
  "otherMails",
  "mobilePhone",
  "businessPhones",
  "jobTitle",
  "department",
  "companyName",
  "officeLocation",
  "employeeId",
  "employeeType",
  "employeeHireDate",
  "employeeOrgData",
  "streetAddress",
  "city",
  "state",
  "postalCode",
  "country",
  "userType",
  "creationType",
  "createdDateTime",
  "preferredLanguage",
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
/**
 * Fetches the user's manager as a separate request. Graph exposes manager
 * as a relationship, not a property, and $expand is incompatible with the
 * advanced-query filter we use on phone numbers — so it needs its own call.
 */
async function fetchManager(
  userId: string,
  token: string
): Promise<Record<string, unknown> | null> {
  const url = `https://graph.microsoft.com/v1.0/users/${userId}/manager?$select=id,displayName,mail,jobTitle,userPrincipalName,department`;

  const response = await fetch(url, {
    method: "GET",
    headers: { Authorization: `Bearer ${token}` },
  });

  // 404 is the normal answer for "no manager assigned" — not an error.
  if (response.status === 404) {
    console.log("  (no manager assigned to this user)");
    return null;
  }

  if (!response.ok) {
    const errorText = await response.text();
    console.error(`  Manager fetch failed (${response.status}): ${errorText}`);
    return null;
  }

  return (await response.json()) as Record<string, unknown>;
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

    for (let i = 0; i < users.length; i++) {
      printUser(users[i], i);

      const userId = users[i].id;
      if (typeof userId === "string") {
        const manager = await fetchManager(userId, token);
        if (manager) {
          console.log(`\n  --- Manager ---`);
          for (const [key, value] of Object.entries(manager)) {
            if (key.startsWith("@")) continue;
            console.log(`  ${key.padEnd(18)}: ${value ?? "(none)"}`);
          }
        }
      }
    }
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
