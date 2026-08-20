// src/lib/textChunker.ts
//
// Splits long text into overlapping chunks, breaking on paragraph
// boundaries where possible. Keeps each chunk comfortably under
// Azure OpenAI's 8192-token embedding limit, and improves retrieval
// precision by keeping chunks focused rather than whole-document.

const TARGET_CHUNK_SIZE = 3000; // characters, not tokens — conservative
const OVERLAP_SIZE = 300;

export function chunkText(text: string): string[] {
  const paragraphs = text
    .split(/\n\s*\n/)
    .map((p) => p.trim())
    .filter((p) => p.length > 0);

  const chunks: string[] = [];
  let current = "";

  for (const paragraph of paragraphs) {
    // If a single paragraph alone exceeds the target size, hard-split it.
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
      // Start the next chunk with a small overlap from the end of the previous one.
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