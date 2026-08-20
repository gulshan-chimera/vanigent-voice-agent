// src/lib/sharepointExplore.ts
//
// TEMPORARY exploration helper. Resolves the site ID for VanigentPortal
// and lists its document libraries ("drives"), so we can confirm whether
// "Knowledge Base" is its own drive or a folder inside the default one,
// before building the real list/download logic.

import { getGraphAccessToken } from "./graphAuth";

export async function exploreSite(): Promise<unknown | null> {
  const token = await getGraphAccessToken();
  if (!token) {
    console.error("[SHAREPOINT-EXPLORE] Could not acquire Graph token.");
    return null;
  }

  const hostname = process.env.SHAREPOINT_HOSTNAME;
  const sitePath = process.env.SHAREPOINT_SITE_PATH;

  if (!hostname || !sitePath) {
    console.error("[SHAREPOINT-EXPLORE] Missing SHAREPOINT_HOSTNAME or SHAREPOINT_SITE_PATH.");
    return null;
  }

  // Step 1: resolve the site ID from its path
  const siteUrl = `https://graph.microsoft.com/v1.0/sites/${hostname}:/sites/${sitePath}`;

  const siteResponse = await fetch(siteUrl, {
    headers: { Authorization: `Bearer ${token}` },
  });

  if (!siteResponse.ok) {
    const errorText = await siteResponse.text();
    console.error(`[SHAREPOINT-EXPLORE] Site lookup failed (${siteResponse.status}): ${errorText}`);
    return null;
  }

  const site = await siteResponse.json();
  const siteId = site.id;

  // Step 2: list the document libraries ("drives") on that site
  const drivesUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`;

  const drivesResponse = await fetch(drivesUrl, {
    headers: { Authorization: `Bearer ${token}` },
  });

  if (!drivesResponse.ok) {
    const errorText = await drivesResponse.text();
    console.error(`[SHAREPOINT-EXPLORE] Drives lookup failed (${drivesResponse.status}): ${errorText}`);
    return { site, drives: null, error: errorText };
  }

  const drives = await drivesResponse.json();

  return { site, drives };
}