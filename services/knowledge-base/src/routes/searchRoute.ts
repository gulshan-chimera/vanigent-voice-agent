import { Router, Request, Response } from "express";
import { SearchService } from "../search/searchService.js";
import { KnowledgeBaseError, toErrorResponse } from "../errors.js";
import type { SearchRequest } from "../search/types.js";
import { authMiddleware } from "../middleware/authMiddleware.js";

// ---------------------------------------------------------------------------
// POST /api/search
// ---------------------------------------------------------------------------

export function createSearchRouter(searchService: SearchService): Router {
  const router = Router();

  router.post("/search", authMiddleware, async (req: Request, res: Response): Promise<void> => {
    try {
      const body = req.body as Partial<SearchRequest>;

      const request: SearchRequest = {
        query: body.query ?? "",
        groupIds: body.groupIds,
        top: body.top,
      };

      const result = await searchService.search(request);

      res.status(200).json(result);
    } catch (error) {
      const errorResponse = toErrorResponse(error);
      const statusCode =
        error instanceof KnowledgeBaseError ? error.statusCode : 500;

      res.status(statusCode).json(errorResponse);
    }
  });

  return router;
}

