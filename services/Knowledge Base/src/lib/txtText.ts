// src/lib/txtText.ts
//
// Reads a plain text file. Simplest of the formats: no parsing library,
// no structure to infer, no page or slide boundaries.
//
// The only real concern is encoding. Files that land in SharePoint come
// from Windows editors, Unix tooling, and exports — so UTF-8 with a BOM,
// and CRLF line endings, both turn up routinely. Neither breaks anything
// visibly, but a stray BOM character at the start of the first chunk
// would end up in the embedded text.

export function extractTxtText(base64Content: string): string | null {
  try {
    const buffer = Buffer.from(base64Content, "base64");

    let text = buffer.toString("utf-8");

    // Strip a UTF-8 byte order mark. Windows editors add one by default
    // and it would otherwise become an invisible character at the very
    // start of chunk 0.
    if (text.charCodeAt(0) === 0xfeff) {
      text = text.slice(1);
    }

    // Normalise Windows line endings so textChunker's paragraph split
    // (which looks for blank lines) behaves the same for every file.
    text = text.replace(/\r\n/g, "\n");

    const trimmed = text.trim();

    if (trimmed.length === 0) {
      console.warn("[TXT-TEXT] File is empty or whitespace only.");
      return null;
    }

    console.log(`[TXT-TEXT] Extracted ${trimmed.length} characters.`);
    return trimmed;
  } catch (error) {
    console.error(`[TXT-TEXT] Extraction failed: ${(error as Error).message}`);
    return null;
  }
}