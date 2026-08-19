// src/lib/pdfPageText.ts
//
// Extracts text PER PAGE (not the whole document as one blob), using
// pdfjs-dist directly — the same engine pdf-to-img uses for rendering,
// guaranteeing page N's text lines up with page N's rendered image.

import * as pdfjsLib from "pdfjs-dist/legacy/build/pdf.mjs";

export async function extractPageTexts(base64Content: string): Promise<string[]> {
  try {
    const pdfBuffer = Buffer.from(base64Content, "base64");
    const pdfUint8Array = new Uint8Array(pdfBuffer);

    const loadingTask = pdfjsLib.getDocument({ data: pdfUint8Array });
    const document = await loadingTask.promise;

    const pageTexts: string[] = [];

    for (let pageNum = 1; pageNum <= document.numPages; pageNum++) {
      const page = await document.getPage(pageNum);
      const textContent = await page.getTextContent();

      const pageText = textContent.items
        .map((item) => ("str" in item ? item.str : ""))
        .join(" ");

      pageTexts.push(pageText.trim());
    }

    console.log(`[PDF-PAGE-TEXT] Extracted text from ${pageTexts.length} page(s).`);
    return pageTexts;
  } catch (error) {
    console.error(`[PDF-PAGE-TEXT] Extraction failed: ${(error as Error).message}`);
    return [];
  }
}