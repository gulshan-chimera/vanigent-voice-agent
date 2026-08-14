// src/types/vapi.ts
//
// Shared TypeScript interfaces for this project.

export interface EntraUser {
  id: string;
  displayName: string;
}

export interface CallerLookupResult {
  isAuthenticated: boolean;
  user: EntraUser | null;
}

// --- VAPI webhook payload shapes (tool-calls event) ---

export interface VapiCustomer {
  number?: string;
}

export interface VapiCall {
  id?: string;
  customer?: VapiCustomer;
}

export interface VapiToolCallItem {
  id: string;
  name?: string;
  arguments?: Record<string, unknown>;
  // Some VAPI versions nest name/arguments under a "function" object instead
  // of putting them at the top level — we handle both shapes defensively.
  function?: {
    name?: string;
    arguments?: Record<string, unknown> | string;
  };
}

export interface VapiToolCallsMessage {
  type: string; // "tool-calls"
  call?: VapiCall;
  toolCallList?: VapiToolCallItem[];
}

export interface VapiWebhookBody {
  message: VapiToolCallsMessage;
}

// --- Our responses back to VAPI ---

export interface VapiToolResult {
  toolCallId: string;
  result: string;
}

export interface VapiToolCallsResponse {
  results: VapiToolResult[];
}