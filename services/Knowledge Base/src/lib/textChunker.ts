// src/lib/textChunker.ts
//
// Splits long text into overlapping chunks, breaking on paragraph
// boundaries where possible.
//
// Used by the DOCX path. PDFs don't need this — they chunk by page,
// which keeps a step's text together with its own screenshot caption.
// DOCX has no page concept, so paragraphs are the next best boundary.
//
// The overlap matters: if an answer spans a chunk boundary, neither
// chunk loses the context needed to make sense of it.

const TARGET_CHUNK_SIZE = 3000; // characters, not tokens — deliberately conservative
const OVERLAP_SIZE = 300;

export function chunkText(text: string): string[] {
  const paragraphs = text
    .split(/\n\s*\n/)
    .map((p) => p.trim())
    .filter((p) => p.length > 0);

  const chunks: string[] = [];
  let current = "";

  for (const paragraph of paragraphs) {
    // A single paragraph longer than the target has to be hard-split —
    // there's no natural boundary inside it to use.
    if (paragraph.length > TARGET_CHUNK_SIZE) {
      if (current.length > 0) {
        chunks.push(current);
        current = "";
      }
      for (let i = 0; i < paragraph.length; i += TARGET_CHUNK_SIZE) {
        chunks.push(paragraph.slice(i, i + TARGET_CHUNK_SIZE));
      }
      continue;
    }

    // Would adding this paragraph push us over the target?
    if (current.length + paragraph.length + 2 > TARGET_CHUNK_SIZE && current.length > 0) {
      chunks.push(current);
      // Carry the tail of the previous chunk into the next one.
      const overlapText = current.slice(-OVERLAP_SIZE);
      current = overlapText + "\n\n" + paragraph;
    } else {
      current = current.length > 0 ? current + "\n\n" + paragraph : paragraph;
    }
  }

  if (current.length > 0) {
    chunks.push(current);
  }

  return chunks;
}