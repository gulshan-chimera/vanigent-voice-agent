/**
 * src/utils/chunker.ts
 * Custom text chunking utility for Semantic Search (Vector Embeddings)
 */

export interface ChunkerOptions {
  chunkSize?: number;
  chunkOverlap?: number;
}

/**
 * Splits text into chunks of approximately `chunkSize` characters,
 * with an overlap of `chunkOverlap` characters between chunks.
 * Attempts to split on natural boundaries (paragraphs, sentences, words).
 */
export function chunkText(
  text: string,
  options?: ChunkerOptions
): string[] {
  const chunkSize = options?.chunkSize ?? 1000;
  const chunkOverlap = options?.chunkOverlap ?? 200;

  if (!text || text.trim().length === 0) return [];

  // Normalize newlines
  const normalizedText = text.replace(/\r\n/g, "\n");

  // Define split separators in order of preference
  const separators = ["\n\n", ". ", "\n", " ", ""];

  function splitRecursively(str: string, maxLen: number): string[] {
    if (str.length <= maxLen) return [str];

    for (const separator of separators) {
      if (separator === "") {
        // Fallback: forcefully split by maxLen
        const chunks: string[] = [];
        let i = 0;
        while (i < str.length) {
          chunks.push(str.substring(i, i + maxLen));
          i += maxLen;
        }
        return chunks;
      }

      const parts = str.split(separator);
      if (parts.length > 1) {
        const finalChunks: string[] = [];
        let currentChunk = "";

        for (const part of parts) {
          const nextLength = currentChunk ? currentChunk.length + separator.length + part.length : part.length;
          
          if (nextLength > maxLen && currentChunk.length > 0) {
            finalChunks.push(currentChunk);
            currentChunk = part;
          } else {
            currentChunk += (currentChunk ? separator : "") + part;
          }
        }
        
        if (currentChunk.length > 0) {
          finalChunks.push(currentChunk);
        }

        // Check if any resulting chunk is still too big, if so, we recurse on that chunk
        const needsMoreSplitting = finalChunks.some(c => c.length > maxLen);
        if (!needsMoreSplitting) {
           return finalChunks;
        }
      }
    }
    return [str];
  }

  const initialChunks = splitRecursively(normalizedText, chunkSize);
  if (chunkOverlap === 0) return initialChunks;

  // Apply overlap
  const result: string[] = [];
  for (let i = 0; i < initialChunks.length; i++) {
    if (i === 0) {
      result.push(initialChunks[i]);
    } else {
      // Prepend the end of the previous chunk to the current chunk
      const prevChunk = result[result.length - 1];
      const overlapText = prevChunk.slice(-chunkOverlap);
      const merged = overlapText + initialChunks[i];
      // Note: we might exceed chunkSize slightly here if the split wasn't perfect,
      // but this is standard practice for overlapping chunkers.
      result.push(merged);
    }
  }

  return result.map(c => c.trim()).filter(c => c.length > 0);
}
