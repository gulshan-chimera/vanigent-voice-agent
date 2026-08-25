// src/types/knowledgeBase.ts
//
// Shared types for the Knowledge Base service (SharePoint file access
// across all document libraries on the site).

export interface DriveItem {
  id: string;
  name: string;
  size?: number;
  webUrl?: string;
  cTag?: string; // changes on CONTENT edits only — our change marker
  eTag?: string; // changes on content OR metadata edits — not used for diffing
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  file?: { mimeType: string };
  folder?: { childCount: number };
}

export interface DriveItemListResponse {
  value: DriveItem[];
  "@odata.nextLink"?: string;
}

/** A SharePoint document library. */
export interface SiteDrive {
  id: string;
  name: string;
  webUrl?: string;
}

export interface KbFileListRequestBody {
  driveId?: string;   // omit to use the first allowlisted library
  folderPath?: string;
}

export interface KbFileDownloadRequestBody {
  itemId: string;
  driveId?: string;
}

export interface KbFileContent {
  fileName: string;
  contentType: string;
  base64Content: string;
}