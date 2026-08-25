// src/functions/queueStatus.ts
//
// Reports indexing health: how many files are waiting, how many are
// stuck in the poison queue, and which files permanently failed.
//
// This is the answer to "did the sync actually finish, and did anything
// get lost?" — which previously required watching the terminal.

import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { QueueClient } from "@azure/storage-queue";
import { isRequestAuthorized } from "../lib/verifyWebhookAuth";
import { listPoisonedFiles } from "../lib/poisonLog";

async function getApproximateCount(queueName: string): Promise<number | null> {
  const connectionString = process.env.AzureWebJobsStorage;
  if (!connectionString) return null;

  try {
    const client = new QueueClient(connectionString, queueName);
    const exists = await client.exists();
    if (!exists) return 0;

    const properties = await client.getProperties();
    return properties.approximateMessagesCount ?? 0;
  } catch (error) {
    console.error(`[QUEUE-STATUS] Failed to read "${queueName}": ${(error as Error).message}`);
    return null;
  }
}

export async function queueStatus(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  if (!isRequestAuthorized(request)) {
    context.warn("[QUEUE-STATUS] Rejected request — invalid or missing bearer token.");
    return { status: 401, jsonBody: { error: "Unauthorized" } };
  }

  const queueName = process.env.KB_INDEX_QUEUE_NAME ?? "kb-index-queue";

  const [pending, poisoned, poisonedFiles] = await Promise.all([
    getApproximateCount(queueName),
    getApproximateCount(`${queueName}-poison`),
    listPoisonedFiles(),
  ]);

  const healthy = pending === 0 && poisoned === 0 && poisonedFiles.length === 0;

  return {
    status: 200,
    jsonBody: {
      healthy,
      pendingMessages: pending,
      poisonMessages: poisoned,
      permanentlyFailedFiles: poisonedFiles,
      note: healthy
        ? "Nothing pending, nothing failed."
        : "Indexing is either still in progress or some files failed permanently. Re-run syncIndex to retry failed files.",
    },
  };
}

app.http("queueStatus", {
  methods: ["GET"],
  authLevel: "anonymous",
  route: "queueStatus",
  handler: queueStatus,
});