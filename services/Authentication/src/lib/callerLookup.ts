// src/lib/callerLookup.ts
// Given a caller's phone number (as sent by VAPI), checks whether it belongs
// to an employee in Microsoft Entra ID by querying the Graph /users endpoint.

import { getGraphAccessToken } from "./graphAuth";
import { EntraUser, CallerLookupResult, EntraManager } from "../types/vapi";
/**
 * Builds a set of plausible phone number variants to check against Entra ID,
 * since numbers may be stored in different formats by different admins
 * (with/without +, with/without country code, spaced, dashed, etc.)
 */
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
  "preferredLanguage",
  "accountEnabled",
].join(",");
export function buildPhoneVariants(rawNumber: string): string[] {
  const variants = new Set<string>();

  // Original, as received
  variants.add(rawNumber);

  // Strip all non-digit characters except leading +
  const digitsOnly = rawNumber.replace(/[^\d+]/g, "");
  variants.add(digitsOnly);

  // Without leading +
  if (digitsOnly.startsWith("+")) {
    variants.add(digitsOnly.slice(1));
  } else {
    variants.add(`+${digitsOnly}`);
  }

  // Without country code, assuming India (+91) or similar 2-digit code,
  // keeping the last 10 digits — common local storage format
  const last10 = digitsOnly.replace(/\D/g, "").slice(-10);
  if (last10.length === 10) {
    variants.add(last10);
    variants.add(`+91${last10}`);
    variants.add(`91${last10}`);
  }

  return Array.from(variants).filter((v) => v.length > 0);
}

/**
 * Fetches the caller's manager as a separate Graph request.
 * Graph doesn't support $expand alongside advanced-query phone filters,
 * so a second call is required. Returns null on 404 (no manager assigned)
 * or any error — the caller's authentication decision is never affected.
 */
async function fetchManager(
  userId: string,
  token: string
): Promise<EntraManager | null> {
  const url = `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(userId)}/manager?$select=id,displayName,mail,jobTitle,userPrincipalName,department`;

  try {
    const response = await fetch(url, {
      method: "GET",
      headers: {
        Authorization: `Bearer ${token}`,
        ConsistencyLevel: "eventual",
      },
    });

    // 404 is the normal answer for "no manager assigned" — not an error.
    if (response.status === 404) {
      console.log("[CALLER-LOOKUP] No manager assigned to this user.");
      return null;
    }

    if (!response.ok) {
      const errorText = await response.text();
      console.warn(
        `[CALLER-LOOKUP] Manager fetch failed (${response.status}): ${errorText} — continuing without manager.`
      );
      return null;
    }

    const data = (await response.json()) as EntraManager & { "@odata.type"?: string };

    return {
      id: data.id,
      displayName: data.displayName,
      mail: data.mail ?? null,
      jobTitle: data.jobTitle ?? null,
      userPrincipalName: data.userPrincipalName ?? null,
      department: data.department ?? null,
    };
  } catch (error) {
    console.warn(
      `[CALLER-LOOKUP] Manager fetch network error: ${(error as Error).message} — continuing without manager.`
    );
    return null;
  }
}
/**
 * Escapes single quotes in a string for safe use inside an OData filter.
 */
export function escapeODataValue(value: string): string {
  return value.replace(/'/g, "''");
}

export async function lookupCaller(rawNumber: string): Promise<CallerLookupResult> {
  const notAuthenticated: CallerLookupResult = { isAuthenticated: false, user: null, manager: null };
  if (!rawNumber) {
    console.warn("[CALLER-LOOKUP] No phone number provided — failing closed.");
    return notAuthenticated;
  }

  const token = await getGraphAccessToken();
  if (!token) {
    console.error("[CALLER-LOOKUP] Could not acquire Graph token — failing closed.");
    return notAuthenticated;
  }

  const variants = buildPhoneVariants(rawNumber);

  // Build an OData filter checking mobilePhone and businessPhones against
  // every variant of the number.
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

  try {
    const response = await fetch(url, {
      method: "GET",
      headers: {
        Authorization: `Bearer ${token}`,
        ConsistencyLevel: "eventual", // required for advanced OData queries like $count and certain filters
      },
    });

    if (!response.ok) {
      const errorText = await response.text();
      console.error(
        `[CALLER-LOOKUP] Graph API error (${response.status}): ${errorText}`
      );
      return notAuthenticated;
    }

    const data = (await response.json()) as { value: EntraUser[] };

    if (!data.value || data.value.length === 0) {
      console.log(`[CALLER-LOOKUP] No match found for number: ${rawNumber}`);
      return notAuthenticated;
    }

    if (data.value.length > 1) {
      console.warn(
        `[CALLER-LOOKUP] Multiple matches (${data.value.length}) found for number: ${rawNumber}. Using first match.`
      );
    }

    const matchedUser = data.value[0];
    console.log(
      `[CALLER-LOOKUP] Match found for ${rawNumber}: ${matchedUser.displayName} (${matchedUser.id})`
    );

    // Fetch manager as a separate call — Graph can't $expand this
    // alongside the advanced-query phone filter.
    // A failed manager lookup never blocks authentication.
    const manager = await fetchManager(matchedUser.id, token);

    return { isAuthenticated: true, user: matchedUser, manager };
  } catch (error) {
    console.error(
      `[CALLER-LOOKUP] Network error during Graph query: ${(error as Error).message}`
    );
    return notAuthenticated;
  }
}

export function formatUserDetails(user: EntraUser, manager: EntraManager | null): string {
  const fields: [string, string | undefined][] = [
    ["Job title", user.jobTitle ?? undefined],
    ["Department", user.department ?? undefined],
    ["Company", user.companyName ?? undefined],
    ["Office location", user.officeLocation ?? undefined],
    ["Employee ID", user.employeeId ?? undefined],
    ["Employee type", user.employeeType ?? undefined],
    ["Hire date", user.employeeHireDate ?? undefined],
    ["User principal name", user.userPrincipalName ?? undefined],
    ["Email", user.mail ?? undefined],
    ["Other emails", user.otherMails?.join(", ")],
    ["Mobile phone", user.mobilePhone ?? undefined],
    ["Business phone", user.businessPhones?.join(", ")],
    ["Preferred language", user.preferredLanguage ?? undefined],
    [
      "Account enabled",
      typeof user.accountEnabled === "boolean" ? String(user.accountEnabled) : undefined,
    ],
  ];

  const lines = [`Caller name: ${user.displayName}`];

  for (const [label, value] of fields) {
    if (value) lines.push(`${label}: ${value}`);
  }

  // Manager details — only included when a manager is assigned.
  // A missing manager is normal and expected for some accounts.
  if (manager) {
    lines.push("--- Manager ---");
    if (manager.displayName) lines.push(`Manager name: ${manager.displayName}`);
    if (manager.mail) lines.push(`Manager email: ${manager.mail}`);
    if (manager.userPrincipalName && manager.userPrincipalName !== manager.mail) {
      lines.push(`Manager UPN: ${manager.userPrincipalName}`);
    }
  }

  return lines.join("\n");
}