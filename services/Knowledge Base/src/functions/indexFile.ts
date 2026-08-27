// src/functions/indexFile.ts
//
// Queue-triggered worker. Fires once per message written by syncIndex,
// with its own execution budget — so a 59-page handbook no longer has
// to fit inside a shared HTTP request timeout.
//
// Throwing on failure is deliberate: it tells the Functions runtime to
// retry, and after maxDequeueCount attempts the message lands in
// "kb-index-queue-poison", where indexFilePoison records it permanently.

import { app, InvocationContext } from "@azure/functions";
import { IndexFileMessage } from "../lib/indexQueue";
import { indexOneFile } from "../lib/fileIndexer";
import { clearPoisonedFile } from "../lib/poisonLog";

export async function indexFile(
  queueItem: unknown,
  context: InvocationContext
): Promise<void> {
  // The runtime hands us a parsed object when the message is JSON, but
  // a string if it couldn't parse it — handle both rather than assume.
  let job: IndexFileMessage;

  try {
    job = (typeof queueItem === "string" ? JSON.parse(queueItem) : queueItem) as IndexFileMessage;
  } catch (error) {
    context.error(`[INDEX-FILE] Unparseable queue message — discarding: ${String(queueItem)}`);
    return; // Don't throw: retrying malformed JSON will never succeed.
  }

  if (!job?.driveId || !job?.itemId) {
    context.error(`[INDEX-FILE] Message missing driveId/itemId — discarding.`);
    return;
  }

  context.log(`[INDEX-FILE] Processing "${job.itemName}" from "${job.driveName}".`);

  const result = await indexOneFile(job);

  if (!result.ok) {
    // Throw so the runtime retries, then poison-queues if it keeps failing.
    throw new Error(
      `[INDEX-FILE] Failed to index "${job.itemName}": ${result.error ?? "no pages indexed"}`
    );
  }

  // This file previously failed permanently but has now succeeded —
  // clear its record. Without this, queueStatus reports healthy: false
  // forever, and an alert that never clears is one people stop reading.
  await clearPoisonedFile(job.driveId, job.itemId);

  context.log(
    `[INDEX-FILE] Done "${job.itemName}": ${result.pagesIndexed} page(s) indexed, ${result.pagesCaptioned} captioned, ${result.pagesFailed} failed.`
  );
}

app.storageQueue("indexFile", {
  queueName: "kb-index-queue",
  connection: "AzureWebJobsStorage",
  handler: indexFile,
});