// src/lib/docxText.ts
//
// Extracts plain text from a .docx file.
//
// Unlike PDF, DOCX has no page concept — Word reflows text based on
// printer and display settings, so there are no fixed page boundaries
// to align text and images against. That means:
//   - chunking is by paragraph, not page (see textChunker.ts)
//   - no page rendering, so no vision captioning
//
// Embedded images are ignored for now: extracting them is possible, but
// they'd have no reliable positional relationship to the surrounding
// text, which is precisely what made the PDF page captions useful.

import mammoth from "mammoth";

export async function extractDocxText(base64Content: string): Promise<string | null> {
  try {
    const buffer = Buffer.from(base64Content, "base64");

    // extractRawText gives plain text without HTML markup — we want the
    // words for embedding, not the formatting.
    const result = await mammoth.extractRawText({ buffer });

    // mammoth reports non-fatal issues (unsupported styles etc.) rather
    // than throwing. Surface them, since they can indicate lost content.
    if (result.messages.length > 0) {
      const warnings = result.messages
        .filter((m) => m.type === "warning" || m.type === "error")
        .map((m) => m.message);

      if (warnings.length > 0) {
        console.warn(`[DOCX-TEXT] Extraction warnings: ${warnings.slice(0, 3).join(" | ")}`);
      }
    }

    const text = result.value?.trim() ?? "";

    if (text.length === 0) {
      console.warn("[DOCX-TEXT] No text extracted — document may be empty or image-only.");
      return null;
    }

    console.log(`[DOCX-TEXT] Extracted ${text.length} characters.`);
    return text;
  } catch (error) {
    console.error(`[DOCX-TEXT] Extraction failed: ${(error as Error).message}`);
    return null;
  }
}