import { getGraphAccessToken } from "./graphAuth";
import { EntraUser } from "../types/entraUser";
import { buildPhoneVariants } from "../../../Authentication/src/lib/callerLookup";

const USERS_URL = "https://graph.microsoft.com/v1.0/users";
const SELECT_FIELDS = "id,displayName,givenName,surname,mail,userPrincipalName,department,jobTitle,accountEnabled";
const RESULT_LIMIT = 25;
const MAX_PAGES = 5;

export class EntraUserSearchError extends Error {
  constructor(
    public readonly code: "INVALID_INPUT" | "AUTH_FAILED" | "GRAPH_REQUEST_FAILED" | "THROTTLED" | "INVALID_RESPONSE",
    message: string,
    public readonly status?: number,
    public readonly retryAfterSeconds?: number,
  ) {
    super(message);
    this.name = "EntraUserSearchError";
  }
}

function text(value: unknown): string | null {
  return typeof value === "string" && value.trim() ? value.trim() : null;
}

function emailIdentity(value: string | null): string | null {
  // Basic identity screening, not mailbox validation. Guest UPNs are sign-in
  // identifiers, not deliverable addresses; a guest's actual mail is still usable.
  return value && !value.toUpperCase().includes("#EXT#") && /^[^\s@<>]+@[^\s@<>]+\.[^\s@<>]+$/u.test(value)
    ? value
    : null;
}

function normalizeUser(value: unknown): EntraUser | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const user = value as Record<string, unknown>;
  const id = text(user.id);
  const displayName = text(user.displayName);
  const mail = text(user.mail);
  const userPrincipalName = text(user.userPrincipalName);
  const email = emailIdentity(mail) ?? emailIdentity(userPrincipalName);
  // Unknown account state is also excluded: do not assume an account is enabled.
  if (user.accountEnabled !== true || !id || !displayName || !email) return null;
  return {
    id, displayName, email, mail, userPrincipalName,
    givenName: text(user.givenName), surname: text(user.surname),
    department: text(user.department), jobTitle: text(user.jobTitle),
    accountEnabled: true,
  };
}

function retryAfterSeconds(value: string | null): number | undefined {
  if (!value) return undefined;
  if (/^\d+$/.test(value)) {
    const seconds = Number(value);
    return Number.isSafeInteger(seconds) ? seconds : undefined;
  }
  const date = Date.parse(value);
  return Number.isFinite(date) ? Math.max(0, Math.ceil((date - Date.now()) / 1000)) : undefined;
}

function invalidResponse(): EntraUserSearchError {
  return new EntraUserSearchError("INVALID_RESPONSE", "Microsoft Graph returned an invalid directory response.");
}

/**
 * Prefix search only: no fuzzy matching, ranking, or automatic attendee choice.
 * Returns up to 25 enabled, email-addressable candidates, scanning at most five
 * pages. An empty array means no usable candidates within that bounded search.
 * Operational failures throw sanitized errors, never masquerading as no matches.
 */
export async function searchEntraUsers(name: string): Promise<EntraUser[]> {
  if (typeof name !== "string" || !name.trim() || name.trim().length > 256 || /[\u0000-\u001f\u007f]/u.test(name)) {
    throw new EntraUserSearchError("INVALID_INPUT", "Provide an attendee name of 1 to 256 characters without control characters.");
  }
  const term = name.trim().replace(/\s+/gu, " ").replace(/'/g, "''");
  const match = term.includes("@")
    ? `(mail eq '${term}' or userPrincipalName eq '${term}')`
    : `(startswith(displayName,'${term}') or startswith(givenName,'${term}') or startswith(surname,'${term}'))`;
  return queryUsers(`accountEnabled eq true and ${match}`);
}

/** Resolve only the phone supplied in authenticated Vapi call metadata. */
export async function lookupMeetingCaller(phone: string): Promise<EntraUser | null> {
  if (!/^\+[1-9]\d{7,14}$/.test(phone)) return null;
  // Reuse phone formatting, but exclude local-number/country-code substitutions.
  const digits = phone.slice(1);
  const variants = buildPhoneVariants(phone).filter(value => value.replace(/\D/g, "") === digits);
  const filter = variants.map(value => {
    const escaped = value.replace(/'/g, "''");
    return `mobilePhone eq '${escaped}' or businessPhones/any(p:p eq '${escaped}')`;
  }).join(" or ");
  const users = await queryUsers(`accountEnabled eq true and (${filter})`, true);
  return users.length === 1 ? users[0] : null;
}

async function queryUsers(filter: string, requireUnique = false, limit = RESULT_LIMIT): Promise<EntraUser[]> {
  let token: string | null;
  try {
    token = await getGraphAccessToken();
  } catch {
    throw new EntraUserSearchError("AUTH_FAILED", "Microsoft Graph authentication failed.");
  }
  if (!token) throw new EntraUserSearchError("AUTH_FAILED", "Microsoft Graph authentication failed.");

  const url = new URL(USERS_URL);
  url.searchParams.set("$select", SELECT_FIELDS);
  url.searchParams.set("$filter", filter);
  url.searchParams.set("$orderby", "displayName");
  url.searchParams.set("$count", "true");
  url.searchParams.set("$top", String(Math.min(limit, 200)));

  let nextUrl: string | undefined = url.toString();
  const users: EntraUser[] = [];
  const seenIds = new Set<string>();
  const seenPages = new Set<string>();
  for (let page = 0; nextUrl && page < MAX_PAGES && users.length < limit; page++) {
    // Never forward the bearer token to a foreign nextLink or HTTP redirect.
    let target: URL;
    try { target = new URL(nextUrl); } catch { throw invalidResponse(); }
    if (target.origin !== "https://graph.microsoft.com" || target.pathname !== "/v1.0/users" || target.username || target.password || target.hash || seenPages.has(target.href)) {
      throw invalidResponse();
    }
    seenPages.add(target.href);

    let response: Response;
    try {
      response = await fetch(target.href, {
        headers: { Authorization: `Bearer ${token}`, ConsistencyLevel: "eventual" },
        signal: AbortSignal.timeout(10_000),
        redirect: "error",
      });
    } catch {
      throw new EntraUserSearchError("GRAPH_REQUEST_FAILED", "Microsoft Graph directory request failed or timed out.");
    }
    if (!response.ok) {
      const throttled = response.status === 429;
      throw new EntraUserSearchError(
        throttled ? "THROTTLED" : "GRAPH_REQUEST_FAILED",
        throttled ? "Microsoft Graph is throttling directory requests. Retry later." : "Microsoft Graph directory request failed.",
        response.status,
        retryAfterSeconds(response.headers.get("retry-after")),
      );
    }
    let body: unknown;
    try { body = await response.json(); } catch { throw invalidResponse(); }
    if (!body || typeof body !== "object" || !Array.isArray((body as Record<string, unknown>).value)) throw invalidResponse();
    const data = body as { value: unknown[]; "@odata.nextLink"?: unknown };
    if (requireUnique) {
      const count = (body as Record<string, unknown>)["@odata.count"];
      if (count !== 1 || data.value.length !== 1 || data["@odata.nextLink"]) return [];
    }
    for (const item of data.value) {
      const user = normalizeUser(item);
      if (user && !seenIds.has(user.id)) {
        seenIds.add(user.id);
        users.push(user);
        if (users.length === limit) break;
      }
    }
    const next = data["@odata.nextLink"];
    if (next !== undefined && next !== null && (typeof next !== "string" || !next)) throw invalidResponse();
    nextUrl = typeof next === "string" ? next : undefined;
  }
  return users;
}

function nameKey(value: string): string {
  return value.normalize("NFD").replace(/\p{M}/gu, "").toLowerCase().replace(/[^\p{L}\p{N}]+/gu, " ").trim();
}

function editDistance(a: string, b: string): number {
  const rows = Array.from({ length: a.length + 1 }, () => Array<number>(b.length + 1).fill(0));
  for (let i = 0; i <= a.length; i++) rows[i][0] = i;
  for (let j = 0; j <= b.length; j++) rows[0][j] = j;
  for (let i = 1; i <= a.length; i++) for (let j = 1; j <= b.length; j++) {
    rows[i][j] = Math.min(rows[i - 1][j] + 1, rows[i][j - 1] + 1, rows[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    if (i > 1 && j > 1 && a[i - 1] === b[j - 2] && a[i - 2] === b[j - 1]) rows[i][j] = Math.min(rows[i][j], rows[i - 2][j - 2] + 1);
  }
  return rows[a.length][b.length];
}

/** Similarity only proposes candidates. It must never authorize an attendee choice. */
export function rankAttendeeCandidates(name: string, users: EntraUser[]): EntraUser[] {
  const key = nameKey(name);
  if (key.length < 3 || key.length > 256) return [];
  const scored = users.map(user => {
    const names = [user.displayName, user.givenName, user.surname].filter((value): value is string => !!value).map(nameKey);
    const score = Math.max(...names.map(candidate => {
      if (candidate.startsWith(key)) return 0.95;
      if (Math.abs(candidate.length - key.length) > 2) return 0;
      const distance = editDistance(key, candidate);
      return distance <= 2 ? 1 - distance / Math.max(key.length, candidate.length) : 0;
    }));
    return { user, score };
  });
  return scored.filter(item => item.score >= 0.65)
    .sort((a, b) => b.score - a.score || a.user.displayName.localeCompare(b.user.displayName))
    .slice(0, 10).map(item => item.user);
}

export async function findAttendeeCandidates(name: string): Promise<{ users: EntraUser[]; exact: boolean }> {
  if (!name.trim() || name.length > 256 || /[\u0000-\u001f\u007f]/.test(name)) {
    throw new EntraUserSearchError("INVALID_INPUT", "Provide a valid attendee name or email.");
  }
  if (name.includes("@")) return { users: await searchEntraUsers(name), exact: true };
  const term = name.trim().replace(/\s+/g, " ").replace(/'/g, "''");
  const exact = await queryUsers(`accountEnabled eq true and (displayName eq '${term}' or givenName eq '${term}' or surname eq '${term}')`);
  if (exact.length) return { users: exact, exact: true };
  const prefixes = await searchEntraUsers(name);
  // Bounded fallback for speech/spelling variants such as Neha/Sneha and Kai/Cai.
  // This is deliberately not an exhaustive search of very large directories.
  const directory = await queryUsers("accountEnabled eq true", false, 1000);
  const unique = new Map([...prefixes, ...directory].map(user => [user.id, user]));
  return { users: rankAttendeeCandidates(name, [...unique.values()]), exact: false };
}
