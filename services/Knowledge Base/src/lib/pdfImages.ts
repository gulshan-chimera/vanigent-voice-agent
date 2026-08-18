// src/lib/pdfImages.ts
//
// Extracts embedded images (not whole rendered pages) from a PDF buffer.
// Pure JS (pdfjs-dist + pngjs under the hood) — no native dependencies,
// safe for Azure Functions. Returns ready-to-use PNG buffers.

import { extractImagesFromPdf } from "pdf-extract-image";

export async function extractPdfImages(base64Content: string): Promise<Buffer[]> {
  try {
    const pdfBuffer = Buffer.from(base64Content, "base64");
    // pdfjs-dist specifically rejects Node's Buffer type, even though
    // Buffer is technically a Uint8Array subclass — needs an explicit
    // conversion to a plain Uint8Array.
    const pdfUint8Array = new Uint8Array(pdfBuffer);

    const images = await extractImagesFromPdf(pdfUint8Array);

    console.log(`[PDF-IMAGES] Extracted ${images.length} image(s).`);
    return images;
  } catch (error) {
    console.error(`[PDF-IMAGES] Extraction failed: ${(error as Error).message}`);
    return [];
  }
}