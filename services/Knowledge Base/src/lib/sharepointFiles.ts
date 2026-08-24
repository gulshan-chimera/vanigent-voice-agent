// src/lib/sharepointFiles.ts
//
// Lists and downloads files from SharePoint document libraries via
// Microsoft Graph. Works against any library ("drive") on the site —
// the drive ID is a parameter, not a fixed setting.

import { getGraphAccessToken } from "./graphAuth";
import {
  DriveItemListResponse,
  DriveItem,
  KbFileContent,
  SiteDrive,
} from "../types/knowledgeBase";

// Explicitly request cTag — we rely on it as the change marker, so we
// don't want to depend on Graph's default field selection including it.
const ITEM_SELECT =
  "id,name,size,webUrl,cTag,eTag,createdDateTime,lastModifiedDateTime,file,folder";

/**
 * Resolves the site ID, then lists every document library on it.
 */
export async function listSiteDrives(): Promise<SiteDrive[] | null> {
  const token = await getGraphAccessToken();
  if (!token) {
    console.error("[SHAREPOINT-FILES] Could not acquire Graph token.");
    return null;
  }

  const hostname = process.env.SHAREPOINT_HOSTNAME;
  const sitePath = process.env.SHAREPOINT_SITE_PATH;

  if (!hostname || !sitePath) {
    console.error("[SHAREPOINT-FILES] Missing SHAREPOINT_HOSTNAME or SHAREPOINT_SITE_PATH.");
    return null;
  }

  const authHeader = { Authorization: `Bearer ${token}` };

  try {
    const siteUrl = `https://graph.microsoft.com/v1.0/sites/${hostname}:/sites/${sitePath}`;
    const siteResponse = await fetch(siteUrl, { headers: authHeader });

    if (!siteResponse.ok) {
      const errorText = await siteResponse.text();
      console.error(`[SHAREPOINT-FILES] Site lookup failed (${siteResponse.status}): ${errorText}`);
      return null;
    }

    const site = (await siteResponse.json()) as { id: string };

    const drivesUrl = `https://graph.microsoft.com/v1.0/sites/${site.id}/drives`;
    const drivesResponse = await fetch(drivesUrl, { headers: authHeader });

    if (!drivesResponse.ok) {
      const errorText = await drivesResponse.text();
      console.error(`[SHAREPOINT-FILES] Drives lookup failed (${drivesResponse.status}): ${errorText}`);
      return null;
    }

    const data = (await drivesResponse.json()) as { value: SiteDrive[] };
    console.log(`[SHAREPOINT-FILES] Found ${data.value.length} document librar(ies) on the site.`);
    return data.value;
  } catch (error) {
    console.error(`[SHAREPOINT-FILES] Network error listing drives: ${(error as Error).message}`);
    return null;
  }
}

/**
 * Applies KB_LIBRARY_ALLOWLIST. Empty/unset means "all libraries".
 * Names are matched case-insensitively, trimmed.
 */
/**
 * Converts a simple glob (`*` = any characters) into an anchored regex.
 * Everything else is escaped, so a pattern like "Sales * KB" behaves
 * literally apart from the wildcard.
 */
function globToRegex(pattern: string): RegExp {
  const escaped = pattern
    .trim()
    .replace(/[.+?^${}()|[\]\\]/g, "\\$&") // escape regex metacharacters
    .replace(/\*/g, ".*");                  // then turn * into .*
  return new RegExp(`^${escaped}$`, "i");   // anchored, case-insensitive
}

export interface DriveSelection {
  included: SiteDrive[];
  excluded: { name: string; reason: string }[];
  mode: "allowlist" | "pattern" | "all";
}

/**
 * Decides which libraries to index.
 *
 * Precedence:
 *   1. KB_LIBRARY_ALLOWLIST — explicit names, wins when set. Intended
 *      for testing or re-syncing a single library.
 *   2. KB_LIBRARY_PATTERN — glob matched against the library name.
 *      This is the normal production path: a library opts in by being
 *      named to match (e.g. "* KB"), so libraries added after deployment
 *      are picked up with no config change.
 *   3. Neither set — everything. Logged loudly, since that includes
 *      SharePoint's default "Documents" library.
 *
 * Excluded libraries are returned rather than silently dropped, so a
 * mistyped name shows up in the sync summary instead of a library
 * quietly never being indexed.
 */
export function selectDrives(drives: SiteDrive[]): DriveSelection {
  const allowlistRaw = process.env.KB_LIBRARY_ALLOWLIST?.trim();
  const patternRaw = process.env.KB_LIBRARY_PATTERN?.trim();

  if (allowlistRaw) {
    const allowed = allowlistRaw
      .split(",")
      .map((n) => n.trim().toLowerCase())
      .filter((n) => n.length > 0);

    const included = drives.filter((d) => allowed.includes(d.name.trim().toLowerCase()));
    const excluded = drives
      .filter((d) => !allowed.includes(d.name.trim().toLowerCase()))
      .map((d) => ({ name: d.name, reason: "not in KB_LIBRARY_ALLOWLIST" }));

    // Names in the allowlist that don't exist on the site — almost
    // always a typo, and invisible unless we say so.
    const missing = allowed.filter(
      (name) => !drives.some((d) => d.name.trim().toLowerCase() === name)
    );
    for (const name of missing) {
      console.warn(`[SHAREPOINT-FILES] Allowlist name not found on the site: "${name}"`);
    }

    console.log(
      `[SHAREPOINT-FILES] ALLOWLIST mode — ${included.length} librar(ies): ${included
        .map((d) => d.name)
        .join(", ")}`
    );
    return { included, excluded, mode: "allowlist" };
  }

  if (patternRaw) {
    const regex = globToRegex(patternRaw);

    const included = drives.filter((d) => regex.test(d.name.trim()));
    const excluded = drives
      .filter((d) => !regex.test(d.name.trim()))
      .map((d) => ({ name: d.name, reason: `does not match pattern "${patternRaw}"` }));

    console.log(
      `[SHAREPOINT-FILES] PATTERN mode "${patternRaw}" — ${included.length} of ${drives.length} librar(ies) matched: ${included
        .map((d) => d.name)
        .join(", ")}`
    );

    if (excluded.length > 0) {
      console.log(
        `[SHAREPOINT-FILES] Excluded: ${excluded.map((e) => e.name).join(", ")}`
      );
    }

    return { included, excluded, mode: "pattern" };
  }

  console.warn(
    "[SHAREPOINT-FILES] Neither KB_LIBRARY_ALLOWLIST nor KB_LIBRARY_PATTERN is set — including ALL libraries, including SharePoint's default 'Documents'."
  );
  return { included: drives, excluded: [], mode: "all" };
}

/**
 * Lists the immediate children of a folder (root by default), following
 * Graph's paging links so we don't silently stop at the first ~200 items.
 */
export async function listDriveChildren(
  driveId: string,
  itemId?: string
): Promise<DriveItem[] | null> {
  const token = await getGraphAccessToken();
  if (!token) {
    console.error("[SHAREPOINT-FILES] Could not acquire Graph token.");
    return null;
  }

  const path = itemId
    ? `/v1.0/drives/${driveId}/items/${itemId}/children`
    : `/v1.0/drives/${driveId}/root/children`;

  let url: string | undefined = `https://graph.microsoft.com${path}?$select=${ITEM_SELECT}&$top=200`;
  const items: DriveItem[] = [];

  try {
    while (url) {
      const response: Response = await fetch(url, {
        headers: { Authorization: `Bearer ${token}` },
      });

      if (!response.ok) {
        const errorText = await response.text();
        console.error(`[SHAREPOINT-FILES] List failed (${response.status}): ${errorText}`);
        return null;
      }

      const data = (await response.json()) as DriveItemListResponse;
      items.push(...data.value);
      url = data["@odata.nextLink"];
    }

    return items;
  } catch (error) {
    console.error(`[SHAREPOINT-FILES] Network error during list: ${(error as Error).message}`);
    return null;
  }
}

/**
 * Walks a whole library recursively and returns every FILE in it,
 * including files nested inside subfolders.
 */
export async function listAllFilesInDrive(driveId: string): Promise<DriveItem[] | null> {
  const allFiles: DriveItem[] = [];
  const foldersToVisit: (string | undefined)[] = [undefined]; // undefined = root

  while (foldersToVisit.length > 0) {
    const folderId = foldersToVisit.shift();
    const children = await listDriveChildren(driveId, folderId);

    if (children === null) {
      console.error(`[SHAREPOINT-FILES] Aborting walk of drive ${driveId} — a listing failed.`);
      return null;
    }

    for (const child of children) {
      if (child.folder) {
        foldersToVisit.push(child.id);
      } else if (child.file) {
        allFiles.push(child);
      }
    }
  }

  return allFiles;
}

/**
 * Downloads a file's content by drive + item ID.
 */
export async function downloadDriveFile(
  driveId: string,
  itemId: string
): Promise<KbFileContent | null> {
  const token = await getGraphAccessToken();
  if (!token) {
    console.error("[SHAREPOINT-FILES] Could not acquire Graph token.");
    return null;
  }

  const authHeader = { Authorization: `Bearer ${token}` };

  try {
    const metaUrl = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}`;
    const metaResponse = await fetch(metaUrl, { headers: authHeader });

    if (!metaResponse.ok) {
      console.error(
        `[SHAREPOINT-FILES] Metadata fetch failed (${metaResponse.status}) for item ${itemId}`
      );
      return null;
    }

    const meta = (await metaResponse.json()) as DriveItem;

    const contentUrl = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/content`;
    const contentResponse = await fetch(contentUrl, { headers: authHeader });

    if (!contentResponse.ok) {
      console.error(
        `[SHAREPOINT-FILES] Content fetch failed (${contentResponse.status}) for item ${itemId}`
      );
      return null;
    }

    const arrayBuffer = await contentResponse.arrayBuffer();

    return {
      fileName: meta.name,
      contentType: meta.file?.mimeType ?? "application/octet-stream",
      base64Content: Buffer.from(arrayBuffer).toString("base64"),
    };
  } catch (error) {
    console.error(`[SHAREPOINT-FILES] Network error during download: ${(error as Error).message}`);
    return null;
  }
}