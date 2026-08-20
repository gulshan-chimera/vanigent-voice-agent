// src/types/knowledgeBase.ts
//
// Shared types for the Knowledge Base service (SharePoint file access,
// scoped to the Human Resources KB document library for now).

export interface DriveItem {
  id: string;
  name: string;
  size?: number;
  webUrl?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  file?: { mimeType: string };
  folder?: { childCount: number };
}

export interface DriveItemListResponse {
  value: DriveItem[];
}

export interface KbFileListRequestBody {
  folderPath?: string; // omit for root of the HR KB library
}

export interface KbFileDownloadRequestBody {
  itemId: string;
}

export interface KbFileContent {
  fileName: string;
  contentType: string;
  base64Content: string;
}