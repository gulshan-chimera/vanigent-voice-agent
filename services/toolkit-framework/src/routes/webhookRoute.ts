import { Router, Request, Response } from "express";
import { AgentToolkit } from "../core/agentToolkit.js";
import { ToolCallContext } from "../core/types.js";

// ---------------------------------------------------------------------------
// VAPI Webhook Route — POST /api/v1/webhook
// ---------------------------------------------------------------------------

/**
 * VAPI sends different event types via webhook:
 *
 * 1. "assistant-request"  — VAPI asks which assistant config to use
 * 2. "function-call"      — VAPI asks us to execute a tool/function
 * 3. "end-of-call-report" — VAPI sends a summary after the call ends
 * 4. "status-update"      — VAPI sends call status changes
 * 5. "hang"               — Call was hung up
 *
 * This route handles all of them and dispatches function-call
 * events to the AgentToolkit for execution.
 */

interface VapiWebhookBody {
  message: {
    type: string;
    call?: {
      id: string;
      customer?: {
        number?: string;
      };
    };
    functionCall?: {
      name: string;
      parameters: Record<string, unknown>;
    };
    [key: string]: unknown;
  };
}

export function createWebhookRouter(toolkit: AgentToolkit): Router {
  const router = Router();

  router.post("/webhook", async (req: Request, res: Response): Promise<void> => {
    const body = req.body as VapiWebhookBody;

    if (!body?.message?.type) {
      res.status(400).json({
        ok: false,
        error: { code: "INVALID_WEBHOOK", message: "Missing message.type" },
      });
      return;
    }

    const messageType = body.message.type;
    const callId = body.message.call?.id ?? "unknown";

    console.log(`[WEBHOOK] Event: "${messageType}" | Call: ${callId}`);

    switch (messageType) {
      // ── VAPI asks which assistant to use ──
      case "assistant-request": {
        const callerNumber = body.message.call?.customer?.number ?? "unknown";
        console.log(`[WEBHOOK] Assistant request from: ${callerNumber}`);

        // Return assistant configuration
        // In production, you'd look up the caller in Entra ID here
        // and return a personalized assistant config
        res.status(200).json({
          assistant: {
            firstMessage: "Hello! How can I help you today?",
            model: {
              provider: "openai",
              model: "gpt-4o",
              messages: [
                {
                  role: "system",
                  content:
                    "You are a helpful voice assistant. Be concise and friendly.",
                },
              ],
            },
            voice: {
              provider: "11labs",
              voiceId: "21m00Tcm4TlvDq8ikWAM",
            },
          },
        });
        return;
      }

      // ── VAPI asks us to execute a function/tool ──
      case "function-call": {
        const functionCall = body.message.functionCall;

        if (!functionCall?.name) {
          res.status(400).json({
            ok: false,
            error: {
              code: "INVALID_FUNCTION_CALL",
              message: "Missing functionCall.name",
            },
          });
          return;
        }

        console.log(
          `[WEBHOOK] Function call: "${functionCall.name}" | Params: ${JSON.stringify(functionCall.parameters)}`
        );

        // Build the tool call context from the VAPI call metadata
        const context: ToolCallContext = {
          agentId: "vapi-voice-agent",
          requestId: `vapi_${callId}_${Date.now()}`,
          traceId: callId,
          metadata: {
            callId,
            callerNumber: body.message.call?.customer?.number,
          },
        };

        // Execute the tool via the AgentToolkit
        const result = await toolkit.call(
          functionCall.name,
          functionCall.parameters,
          context
        );

        if (result.ok) {
          // VAPI expects the result in a specific format
          res.status(200).json({ result: JSON.stringify(result.data) });
        } else {
          console.error(
            `[WEBHOOK] Tool failed: ${result.error.code} — ${result.error.message}`
          );
          res.status(200).json({
            result: JSON.stringify({
              error: true,
              message: result.error.message,
            }),
          });
        }
        return;
      }

      // ── Call ended — log the summary ──
      case "end-of-call-report": {
        console.log(`[WEBHOOK] Call ended: ${callId}`);
        console.log(
          `[WEBHOOK] Report: ${JSON.stringify(body.message, null, 2)}`
        );
        res.status(200).json({ ok: true });
        return;
      }

      // ── Status updates (ringing, in-progress, ended) ──
      case "status-update": {
        const status = (body.message as any).status ?? "unknown";
        console.log(`[WEBHOOK] Status update: ${status} | Call: ${callId}`);
        res.status(200).json({ ok: true });
        return;
      }

      // ── Call hung up ──
      case "hang": {
        console.log(`[WEBHOOK] Call hung up: ${callId}`);
        res.status(200).json({ ok: true });
        return;
      }

      // ── Unknown event type ──
      default: {
        console.warn(`[WEBHOOK] Unhandled event type: "${messageType}"`);
        res.status(200).json({ ok: true });
        return;
      }
    }
  });

  return router;
}
