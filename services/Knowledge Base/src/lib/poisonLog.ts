// src/lib/poisonLog.ts
//
// Persists permanently-failed files to Azure Table Storage so they can
// be reported by /api/queueStatus. Logs alone aren't enough — they
// scroll away, and in production nobody is watching a terminal.
//
// Note this uses the same storage account as the queue, and is the same
// mechanism automatic sync will later use for subscription IDs and
// delta tokens.

import { TableClient, TableEntity } from "@azure/data-tables";

const TABLE_NAME = "kbPoisonedFiles";

export interface PoisonedFile {
  itemName: string;
  driveName: string;
  driveId: string;
  itemId: string;
  failedAt: string;
}

function getTableClient(): TableClient | null {
  const connectionString = process.env.AzureWebJobsStorage;

  if (!connectionString) {
    console.error("[POISON-LOG] AzureWebJobsStorage is not set.");
    return null;
  }

  return TableClient.fromConnectionString(connectionString, TABLE_NAME, {
    allowInsecureConnection: true, // required for Azurite over http locally
  });
}

export async function recordPoisonedFile(file: PoisonedFile): Promise<void> {
  const client = getTableClient();
  if (!client) return;

  try {
    await client.createTable();

    // Table keys can't contain / \ # ?, and itemIds are opaque strings —
    // encode to be safe.
    const rowKey = Buffer.from(`${file.driveId}__${file.itemId}`)
      .toString("base64")
      .replace(/[+/=]/g, "_");

    const entity: TableEntity<Record<string, unknown>> = {
      partitionKey: "poisoned",
      rowKey,
      ...file,
    };

    // upsert: the same file failing twice should update, not duplicate.
    await client.upsertEntity(entity, "Replace");
  } catch (error) {
    console.error(`[POISON-LOG] Failed to record "${file.itemName}": ${(error as Error).message}`);
  }
}

export async function listPoisonedFiles(): Promise<PoisonedFile[]> {
  const client = getTableClient();
  if (!client) return [];

  try {
    await client.createTable();

    const files: PoisonedFile[] = [];
    const entities = client.listEntities<TableEntity<Record<string, unknown>>>();

    for await (const entity of entities) {
      files.push({
        itemName: String(entity.itemName ?? ""),
        driveName: String(entity.driveName ?? ""),
        driveId: String(entity.driveId ?? ""),
        itemId: String(entity.itemId ?? ""),
        failedAt: String(entity.failedAt ?? ""),
      });
    }

    return files;
  } catch (error) {
    console.error(`[POISON-LOG] Failed to list: ${(error as Error).message}`);
    return [];
  }
}

/** Clears the record for one file — call after a successful re-index. */
export async function clearPoisonedFile(driveId: string, itemId: string): Promise<void> {
  const client = getTableClient();
  if (!client) return;

  try {
    const rowKey = Buffer.from(`${driveId}__${itemId}`)
      .toString("base64")
      .replace(/[+/=]/g, "_");

    await client.deleteEntity("poisoned", rowKey);
  } catch {
    // Not being there is fine — nothing to clear.
  }
}