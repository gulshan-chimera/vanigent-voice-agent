// src/lib/helpTopics.ts
//
// Maps help topic NAMES (as shown on osTicket's "Open a New Ticket" page,
// e.g. "HR - Benefits") to osTicket help topic IDs.
//
// osTicket decides the department from the help topic, so the topic is
// what routes a ticket to HR, IT, etc. Its API only accepts the numeric
// topicId, and has no endpoint to list topics, so the map is configured:
//
//   OSTICKET_HELP_TOPICS = {"HR - Benefits": 12, "IT - Hardware": 3, ...}
//
// Find each ID in osTicket: Admin Panel → Manage → Help Topics → open the
// topic; the URL ends in "?id=<topicId>".
//
// Names are matched case-insensitively with surrounding whitespace ignored.

let cachedRaw: string | undefined;
let cachedMap: Map<string, { name: string; id: number }> | null = null;

function normalize(name: string): string {
  return name.trim().replace(/\s+/g, " ").toLowerCase();
}

function loadTopics(): Map<string, { name: string; id: number }> | null {
  const raw = process.env.OSTICKET_HELP_TOPICS;
  if (!raw) return null;
  if (raw === cachedRaw) return cachedMap;

  cachedRaw = raw;
  cachedMap = null;
  try {
    const parsed = JSON.parse(raw) as Record<string, unknown>;
    const map = new Map<string, { name: string; id: number }>();
    for (const [name, id] of Object.entries(parsed)) {
      if (typeof id === "number" && Number.isInteger(id)) {
        map.set(normalize(name), { name, id });
      } else {
        console.error(`[OSTICKET] OSTICKET_HELP_TOPICS: "${name}" has a non-integer ID — skipped.`);
      }
    }
    cachedMap = map;
  } catch {
    console.error("[OSTICKET] OSTICKET_HELP_TOPICS is not valid JSON.");
  }
  return cachedMap;
}

export type TopicLookup =
  | { ok: true; topicId: number }
  | { ok: false; reason: "not_configured" }
  | { ok: false; reason: "unknown"; knownTopics: string[] };

export function resolveTopicId(topicName: string): TopicLookup {
  const topics = loadTopics();
  if (!topics) return { ok: false, reason: "not_configured" };

  const match = topics.get(normalize(topicName));
  if (match) return { ok: true, topicId: match.id };
  return { ok: false, reason: "unknown", knownTopics: [...topics.values()].map((t) => t.name) };
}
