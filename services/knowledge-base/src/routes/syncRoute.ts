import { Router, Request, Response } from "express";
import { SyncService } from "../sync/syncService.js";
import { authMiddleware } from "../middleware/authMiddleware.js";

export function createSyncRouter(syncService: SyncService): Router {
  const router = Router();

  router.post("/sync", authMiddleware, async (req: Request, res: Response): Promise<void> => {
    try {
      console.log("[API] Manual sync triggered via API");
      // Fire and forget so we don't hold the HTTP request open for minutes
      syncService.runDeltaSync().catch(err => {
        console.error("[API] Background sync failed:", err);
      });
      
      res.status(202).json({ 
        message: "Sync started in the background. Check server logs for progress." 
      });
    } catch (error) {
      console.error("[API] Failed to start sync:", error);
      res.status(500).json({ error: "Failed to start synchronization process." });
    }
  });

  return router;
}
