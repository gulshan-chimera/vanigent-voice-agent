// src/lib/sharepointFiles.ts
//
// Lists and downloads files from the Human Resources KB SharePoint
// document library, using Microsoft Graph's drive API. Scoped to a single
// fixed drive ID for now (SHAREPOINT_HR_KB_DRIVE_ID).

import { getGraphAccessToken } from "./graphAuth";
import { DriveItemListResponse, DriveItem, KbFileContent } from "../types/knowledgeBase";

function getDriveId(): string | null {
  const driveId = process.env.SHAREPOINT_HR_KB_DRIVE_ID;
  if (!driveId) {
    console.error("[SHAREPOINT-FILES] Missing SHAREPOINT_HR_KB_DRIVE_ID.");
    return null;
  }
  return driveId;
}

/**
 * Lists items (files and folders) at the root of the HR KB library, or
 * inside a specific subfolder if folderPath is given (e.g. "Policies").
 */
export async function listHrKbFiles(folderPath?: string): Promise<DriveItemListResponse | null> {
  const token = await getGraphAccessToken();
  const driveId = getDriveId();

  if (!token || !driveId) {
    console.error("[SHAREPOINT-FILES] Missing Graph token or drive ID.");
    return null;
  }

const path = folderPath
    ? `/v1.0/drives/${driveId}/root:/${encodeURIComponent(folderPath)}:/children`
    : `/v1.0/drives/${driveId}/root/children`;

  const url = `https://graph.microsoft.com${path}`;

  try {
    const response = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` },
    });

    if (!response.ok) {
      const errorText = await response.text();
      console.error(`[SHAREPOINT-FILES] List failed (${response.status}): ${errorText}`);
      return null;
    }

    const data = (await response.json()) as DriveItemListResponse;
    return data;
  } catch (error) {
    console.error(`[SHAREPOINT-FILES] Network error during list: ${(error as Error).message}`);
    return null;
  }
}

/**
 * Downloads a specific file's content by its item ID. Fetches metadata
 * first (for the correct file name and MIME type), then the raw content.
 */
export async function downloadHrKbFile(itemId: string): Promise<KbFileContent | null> {
  const token = await getGraphAccessToken();
  const driveId = getDriveId();

  if (!token || !driveId) {
    console.error("[SHAREPOINT-FILES] Missing Graph token or drive ID.");
    return null;
  }

  const authHeader = { Authorization: `Bearer ${token}` };

  try {
    // Step 1: get metadata (name, mime type)
    const metaUrl = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}`;
    const metaResponse = await fetch(metaUrl, { headers: authHeader });

    if (!metaResponse.ok) {
      console.error(`[SHAREPOINT-FILES] Metadata fetch failed (${metaResponse.status}) for item ${itemId}`);
      return null;
    }

    const meta = (await metaResponse.json()) as DriveItem;

    // Step 2: get the raw file content
    const contentUrl = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/content`;
    const contentResponse = await fetch(contentUrl, { headers: authHeader });

    if (!contentResponse.ok) {
      console.error(`[SHAREPOINT-FILES] Content fetch failed (${contentResponse.status}) for item ${itemId}`);
      return null;
    }

    const arrayBuffer = await contentResponse.arrayBuffer();
    const base64Content = Buffer.from(arrayBuffer).toString("base64");

    return {
      fileName: meta.name,
      contentType: meta.file?.mimeType ?? "application/octet-stream",
      base64Content,
    };
  } catch (error) {
    console.error(`[SHAREPOINT-FILES] Network error during download: ${(error as Error).message}`);
    return null;
  }
}