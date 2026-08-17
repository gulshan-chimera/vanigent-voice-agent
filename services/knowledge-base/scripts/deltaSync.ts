/**
 * scripts/deltaSync.ts
 * CLI wrapper for SyncService
 */
import { SyncService } from "../src/sync/syncService.js";

async function main() {
  const service = new SyncService();
  await service.runDeltaSync();
}

main().catch((error) => {
  console.error("[FATAL]", error);
  process.exit(1);
});
