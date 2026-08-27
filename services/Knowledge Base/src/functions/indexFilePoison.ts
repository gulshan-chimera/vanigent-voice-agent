// src/functions/indexFilePoison.ts
//
// Azure moves a message to "<queue>-poison" after maxDequeueCount failed
// attempts. Without this handler those messages vanish silently — the
// file never gets indexed, nothing logs it, and a caller asking about
// that document simply hears "I don't have that information."
//
// This catches them, logs loudly with the file's identity, and records
// the failure so /api/queueStatus can report it.

import { app, InvocationContext } from "@azure/functions";
import { IndexFileMessage } from "../lib/indexQueue";
import { recordPoisonedFile } from "../lib/poisonLog";

export async function indexFilePoison(
  queueItem: unknown,
  context: InvocationContext
): Promise<void> {
  let job: Partial<IndexFileMessage> = {};

  try {
    job = (typeof queueItem === "string" ? JSON.parse(queueItem) : queueItem) as IndexFileMessage;
  } catch {
    context.error(
      `[POISON] Unparseable message reached the poison queue: ${String(queueItem)}`
    );
    return;
  }

  context.error(
    `[POISON] PERMANENTLY FAILED after all retries: "${job.itemName}" from "${job.driveName}". ` +
      `This file is NOT in the index. driveId=${job.driveId} itemId=${job.itemId}`
  );

  await recordPoisonedFile({
    itemName: job.itemName ?? "(unknown)",
    driveName: job.driveName ?? "(unknown)",
    driveId: job.driveId ?? "",
    itemId: job.itemId ?? "",
    failedAt: new Date().toISOString(),
  });
}

app.storageQueue("indexFilePoison", {
  queueName: "kb-index-queue-poison",
  connection: "AzureWebJobsStorage",
  handler: indexFilePoison,
});