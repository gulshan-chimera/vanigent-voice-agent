// src/lib/pdfImages.ts
//
// Renders each page of a PDF as a full PNG image, so nothing visible on
// the page — screenshots, charts, logos, whatever — can be missed,
// regardless of how it's technically embedded internally. Pure JS, no
// native compilation required.

import { pdf } from "pdf-to-img";

export async function renderPdfPagesAsImages(base64Content: string): Promise<Buffer[]> {
  try {
    const pdfBuffer = Buffer.from(base64Content, "base64");
    const pdfUint8Array = new Uint8Array(pdfBuffer);
    const document = await pdf(pdfUint8Array, { scale: 2 });

    const pageImages: Buffer[] = [];
    for await (const pageImage of document) {
      pageImages.push(pageImage);
    }

    console.log(`[PDF-IMAGES] Rendered ${pageImages.length} page(s) as images.`);
    return pageImages;
  } catch (error) {
    console.error(`[PDF-IMAGES] Page rendering failed: ${(error as Error).message}`);
    return [];
  }
}