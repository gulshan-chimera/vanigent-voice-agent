// src/lib/confidentiality.ts
//
// Detects documents marked confidential so the answer layer knows to
// withhold figures from them.
//
// Detection is automatic rather than a configured file list: the client
// uploads documents continuously, and anything requiring someone to
// remember to add a filename to a setting will eventually miss one.
// Documents that carry a confidentiality marking are self-identifying.

/**
 * Markers that indicate a document is not for general disclosure.
 * Matched case-insensitively against the document's full text.
 *
 * Deliberately conservative: "confidential" alone would match any
 * document that merely discusses confidentiality — the employee
 * handbook's confidentiality policy, for instance — and wrongly
 * restrict it. These phrases are ones used to LABEL a document.
 */
const CONFIDENTIALITY_MARKERS = [
  "company confidential",
  "confidential and proprietary",
  "internal use only",
  "for internal distribution only",
  "restricted distribution",
  "do not distribute",
  "not for external distribution",
];

export interface ConfidentialityResult {
  restricted: boolean;
  marker: string | null;
  /** How many times the marker appeared — a page footer repeats, a
   *  passing mention in body text does not. */
  occurrences: number;
}

/**
 * Minimum occurrences before a marker counts as a classification label.
 * The employee handbook discusses confidentiality obligations in prose
 * and would otherwise be restricted, which would needlessly filter the
 * most frequently asked-about document in the knowledge base.
 */
const MIN_OCCURRENCES = 2;

function countOccurrences(haystack: string, needle: string): number {
  let count = 0;
  let index = haystack.indexOf(needle);

  while (index !== -1) {
    count++;
    index = haystack.indexOf(needle, index + needle.length);
  }

  return count;
}

export function detectConfidentiality(text: string): ConfidentialityResult {
  if (!text) return { restricted: false, marker: null, occurrences: 0 };

  const haystack = text.toLowerCase();

  let bestMarker: string | null = null;
  let bestCount = 0;

  for (const marker of CONFIDENTIALITY_MARKERS) {
    const count = countOccurrences(haystack, marker);
    if (count > bestCount) {
      bestCount = count;
      bestMarker = marker;
    }
  }

  return {
    restricted: bestCount >= MIN_OCCURRENCES,
    marker: bestMarker,
    occurrences: bestCount,
  };
}

/**
 * Turns a library name into the team a caller should be referred to.
 * "Sales Operations KB" becomes "Sales Operations". The libraries are
 * already named by owning department, so no mapping table is needed.
 */
export function teamNameForLibrary(library: string): string {
  const cleaned = library
    .replace(/\s*\bKB\b\s*$/i, "")
    .replace(/\s*knowledge base\s*$/i, "")
    .trim();

  return cleaned.length > 0 ? cleaned : library;
}

/**
 * Filename patterns that restrict a document regardless of markers.
 *
 * Marker detection requires the label to appear at least twice, which
 * separates a footer from a passing mention — but a document marked
 * only on its cover page would score one and be missed silently.
 * This list is the deliberate override for document types known to be
 * sensitive by their nature.
 *
 * Comma-separated in KB_RESTRICTED_PATTERNS, matched case-insensitively
 * as substrings of the file name.
 */
function getFilenamePatterns(): string[] {
  const raw = process.env.KB_RESTRICTED_PATTERNS?.trim();
  if (!raw) return [];

  return raw
    .split(",")
    .map((s) => s.trim().toLowerCase())
    .filter((s) => s.length > 0);
}

export function matchesRestrictedFilename(fileName: string): string | null {
  const lower = fileName.toLowerCase();

  for (const pattern of getFilenamePatterns()) {
    if (lower.includes(pattern)) return pattern;
  }

  return null;
}

/**
 * Combines both checks. A document is restricted if its name matches a
 * configured pattern OR it carries a repeated confidentiality marking.
 */
export function isDocumentRestricted(
  fileName: string,
  text: string
): { restricted: boolean; reason: string | null } {
  const patternMatch = matchesRestrictedFilename(fileName);
  if (patternMatch) {
    return { restricted: true, reason: `filename pattern "${patternMatch}"` };
  }

  const detected = detectConfidentiality(text);
  if (detected.restricted) {
    return {
      restricted: true,
      reason: `marker "${detected.marker}" x${detected.occurrences}`,
    };
  }

  return { restricted: false, reason: null };
}