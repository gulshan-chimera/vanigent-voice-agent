// src/lib/indexQueue.ts
//
// Writes per-file indexing jobs onto an Azure Storage Queue. syncIndex
// decides WHAT needs indexing and enqueues one message per file; the
// queue-triggered indexFile worker does the actual (slow) work.
//
// This exists because a full sync takes ~9 minutes, and Azure Functions
// on the Consumption plan hard-kills an HTTP request at 230 seconds.
// Splitting the work gives each file its own execution budget, plus
// automatic retries and a poison queue for files that genuinely fail.

import { QueueClient } from "@azure/storage-queue";

/** The job payload. Everything the worker needs to index one file. */
export interface IndexFileMessage {
  driveId: string;
  driveName: string;
  itemId: string;
  itemName: string;
  cTag: string;
  webUrl: string;
  lastModifiedDateTime: string;
  /** True when this file was already indexed and is being replaced. */
  isUpdate: boolean;
}

function getQueueClient(): QueueClient | null {
  const connectionString = process.env.AzureWebJobsStorage;
  const queueName = process.env.KB_INDEX_QUEUE_NAME;

  if (!connectionString) {
    console.error("[INDEX-QUEUE] AzureWebJobsStorage is not set — cannot enqueue.");
    return null;
  }

  if (!queueName) {
    console.error("[INDEX-QUEUE] KB_INDEX_QUEUE_NAME is not set — cannot enqueue.");
    return null;
  }

  return new QueueClient(connectionString, queueName);
}

/**
 * Ensures the queue exists. Safe to call repeatedly — createIfNotExists
 * is idempotent.
 */
export async function ensureQueueExists(): Promise<boolean> {
  const client = getQueueClient();
  if (!client) return false;

  try {
    await client.createIfNotExists();
    return true;
  } catch (error) {
    console.error(`[INDEX-QUEUE] Failed to create queue: ${(error as Error).message}`);
    return false;
  }
}

/**
 * Enqueues one indexing job.
 *
 * The Functions queue trigger expects base64-encoded message content by
 * default, while the SDK sends raw text — so we encode explicitly here
 * to keep the two sides in agreement.
 */
export async function enqueueIndexJob(message: IndexFileMessage): Promise<boolean> {
  const client = getQueueClient();
  if (!client) return false;

  try {
    const encoded = Buffer.from(JSON.stringify(message)).toString("base64");
    await client.sendMessage(encoded);
    return true;
  } catch (error) {
    console.error(
      `[INDEX-QUEUE] Failed to enqueue "${message.itemName}": ${(error as Error).message}`
    );
    return false;
  }
}