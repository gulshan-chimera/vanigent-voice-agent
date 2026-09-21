// src/lib/callDepartmentTracker.ts
//
// Tracks which departments (libraries) were queried during each VAPI
// call. The KB webhook records departments as answers are generated,
// and the Authentication service fetches them at end-of-call to
// populate "Systems Affected" in the follow-up email.
//
// Uses a simple in-memory Map with TTL cleanup. Calls rarely last
// longer than 30 minutes, so a 1-hour TTL is more than sufficient.

const TTL_MS = 60 * 60 * 1000; // 1 hour

interface CallEntry {
  departments: Set<string>;
  createdAt: number;
}

const callMap = new Map<string, CallEntry>();

/** Removes expired entries to prevent unbounded memory growth. */
function cleanup(): void {
  const now = Date.now();
  for (const [callId, entry] of callMap) {
    if (now - entry.createdAt > TTL_MS) {
      callMap.delete(callId);
    }
  }
}

/** Records one or more departments for a given call ID. */
export function trackDepartments(callId: string, departments: string[]): void {
  cleanup();

  let entry = callMap.get(callId);
  if (!entry) {
    entry = { departments: new Set(), createdAt: Date.now() };
    callMap.set(callId, entry);
  }

  for (const dept of departments) {
    if (dept) entry.departments.add(dept);
  }
}

/** Returns the accumulated departments for a call, or an empty array. */
export function getDepartments(callId: string): string[] {
  const entry = callMap.get(callId);
  return entry ? [...entry.departments] : [];
}
