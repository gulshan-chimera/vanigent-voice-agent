// src/lib/callDepartmentTracker.ts
//
// Tracks which departments (libraries) were queried during each VAPI
// call. The KB webhook records departments as answers are generated,
// and the Authentication service fetches them at end-of-call to
// populate "Systems Affected" in the follow-up email.
//
// Uses Azure Table Storage so data persists across Function App
// instances and survives cold starts / instance recycling.

import { TableClient, TableServiceClient } from "@azure/data-tables";

const TABLE_NAME = "CallDepartments";

let tableClient: TableClient | null = null;

function getTableClient(): TableClient | null {
  if (tableClient) return tableClient;

  const connectionString = process.env.AzureWebJobsStorage;
  if (!connectionString) {
    console.error("[DEPT-TRACKER] AzureWebJobsStorage is not set — cannot track departments.");
    return null;
  }

  const serviceClient = TableServiceClient.fromConnectionString(connectionString);
  // Ensure the table exists (no-op if it already does).
  serviceClient.createTable(TABLE_NAME).catch(() => {
    // Table already exists — ignore.
  });

  tableClient = TableClient.fromConnectionString(connectionString, TABLE_NAME);
  return tableClient;
}

/** Records one or more departments for a given call ID. */
export async function trackDepartments(callId: string, departments: string[]): Promise<void> {
  const client = getTableClient();
  if (!client || departments.length === 0) return;

  try {
    // Try to fetch existing entry first
    let existing: string[] = [];
    try {
      const entity = await client.getEntity<{ departments: string }>("call", callId);
      existing = JSON.parse(entity.departments);
    } catch {
      // Entity doesn't exist yet — that's fine.
    }

    // Merge with new departments (deduplicated)
    const merged = [...new Set([...existing, ...departments])];

    await client.upsertEntity(
      {
        partitionKey: "call",
        rowKey: callId,
        departments: JSON.stringify(merged),
        timestamp: new Date().toISOString(),
      },
      "Replace"
    );

    console.log(`[DEPT-TRACKER] Stored departments for call ${callId}: ${merged.join(", ")}`);
  } catch (error) {
    console.warn(
      `[DEPT-TRACKER] Failed to store departments for call ${callId}: ${(error as Error).message}`
    );
  }
}

/** Returns the accumulated departments for a call, or an empty array. */
export async function getDepartments(callId: string): Promise<string[]> {
  const client = getTableClient();
  if (!client) return [];

  try {
    const entity = await client.getEntity<{ departments: string }>("call", callId);
    return JSON.parse(entity.departments);
  } catch {
    // Entity not found
    return [];
  }
}
