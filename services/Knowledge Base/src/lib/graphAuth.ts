// src/lib/graphAuth.ts
//
// Acquires an app-only (client_credentials) access token for Microsoft
// Graph, and CACHES it in module scope.
//
// Caching matters more than it looks: every Graph call previously
// triggered a fresh POST to login.microsoftonline.com, and
// downloadDriveFile makes two Graph calls per file. With several files
// indexing concurrently that produced a burst of identical token
// requests — which is what caused the "fetch failed" errors that killed
// three files mid-sync. Tokens are valid for roughly an hour, so one
// fetch per hour is all that's needed.

interface TokenResponse {
  access_token?: string;
  expires_in?: number; // seconds until expiry
  error?: string;
  error_description?: string;
}

interface CachedToken {
  token: string;
  expiresAt: number; // epoch ms
}

let cachedToken: CachedToken | null = null;

// Refresh this long before actual expiry, so a token can't lapse
// mid-request during a long-running index job.
const EXPIRY_BUFFER_MS = 5 * 60 * 1000; // 5 minutes

// Shared in-flight request. Without this, several concurrent callers on
// a cold cache would each fire their own token request — exactly the
// burst we're trying to eliminate. They now all await the same promise.
let inFlightRequest: Promise<string | null> | null = null;

async function fetchNewToken(): Promise<string | null> {
  const tenantId = process.env.TENANT_ID;
  const clientId = process.env.CLIENT_ID;
  const clientSecret = process.env.CLIENT_SECRET;

  if (!tenantId || !clientId || !clientSecret) {
    console.error("[GRAPH-AUTH] Missing TENANT_ID, CLIENT_ID, or CLIENT_SECRET.");
    return null;
  }

  const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    grant_type: "client_credentials",
    client_id: clientId,
    client_secret: clientSecret,
    scope: "https://graph.microsoft.com/.default",
  });

  // Transient network failures happen — retry rather than failing the
  // whole file, since a token failure aborts everything downstream.
  const MAX_ATTEMPTS = 3;

  for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
    try {
      const response = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: body.toString(),
      });

      const data = (await response.json()) as TokenResponse;

      if (!response.ok || !data.access_token) {
        // A bad secret is permanent — retrying won't help, so fail fast.
        console.error(
          `[GRAPH-AUTH] Token request rejected: ${data.error} - ${data.error_description}`
        );
        return null;
      }

      const expiresInMs = (data.expires_in ?? 3600) * 1000;
      cachedToken = {
        token: data.access_token,
        expiresAt: Date.now() + expiresInMs,
      };

      console.log(
        `[GRAPH-AUTH] New token acquired, valid for ${Math.round(expiresInMs / 60000)} minute(s).`
      );
      return data.access_token;
    } catch (error) {
      const message = (error as Error).message;

      if (attempt === MAX_ATTEMPTS) {
        console.error(
          `[GRAPH-AUTH] Network error after ${MAX_ATTEMPTS} attempts: ${message}`
        );
        return null;
      }

      const waitMs = 1000 * Math.pow(2, attempt - 1); // 1s, 2s
      console.warn(
        `[GRAPH-AUTH] Network error (attempt ${attempt}/${MAX_ATTEMPTS}): ${message} — retrying in ${waitMs / 1000}s.`
      );
      await new Promise((resolve) => setTimeout(resolve, waitMs));
    }
  }

  return null;
}

export async function getGraphAccessToken(): Promise<string | null> {
  // Cached and comfortably valid — the overwhelmingly common path.
  if (cachedToken && Date.now() < cachedToken.expiresAt - EXPIRY_BUFFER_MS) {
    return cachedToken.token;
  }

  // A fetch is already running; join it rather than starting another.
  if (inFlightRequest) {
    return inFlightRequest;
  }

  inFlightRequest = fetchNewToken().finally(() => {
    inFlightRequest = null;
  });

  return inFlightRequest;
}

/** Clears the cached token. Exposed for diagnostics/testing. */
export function clearTokenCache(): void {
  cachedToken = null;
  console.log("[GRAPH-AUTH] Token cache cleared.");
}