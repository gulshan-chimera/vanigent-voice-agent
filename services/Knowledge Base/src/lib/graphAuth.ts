// src/lib/graphAuth.ts
//
// Acquires an app-only (client_credentials) access token for Microsoft Graph.
// Reuses the same App Registration as the Authentication service.

interface TokenResponse {
  access_token?: string;
  error?: string;
  error_description?: string;
}

export async function getGraphAccessToken(): Promise<string | null> {
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

  try {
    const response = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: body.toString(),
    });

    const data = (await response.json()) as TokenResponse;

    if (!response.ok || !data.access_token) {
      console.error(`[GRAPH-AUTH] Token request failed: ${data.error} - ${data.error_description}`);
      return null;
    }

    return data.access_token;
  } catch (error) {
    console.error(`[GRAPH-AUTH] Network error: ${(error as Error).message}`);
    return null;
  }
}