// src/lib/pptxText.ts
//
// Extracts text PER SLIDE directly from a .pptx file's own XML — no PDF
// conversion involved. PPTX is a zip archive of per-slide XML parts
// (ppt/slides/slideN.xml); every visible text run is wrapped in a
// DrawingML <a:t> tag, so a slide's script is just the runs inside its
// own slideN.xml, joined in document order.
//
// This exists for the same reason docxText.ts reads DOCX directly
// instead of going through Graph's PDF conversion: pdfjs's text
// extraction from a converted PDF reconstructs reading order from
// on-page text positions, which scrambles multi-column slide layouts
// (e.g. a step list on the left, a screenshot on the right) badly enough
// that some slides' content became unretrievable. Reading the XML
// directly keeps each slide's authored text in its actual order.

import JSZip from "jszip";

function decodeXmlEntities(text: string): string {
  return text
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

function extractTextFromSlideXml(xml: string): string {
  const matches = xml.match(/<a:t>([^<]*)<\/a:t>/g) ?? [];
  return matches
    .map((tag) => decodeXmlEntities(tag.replace(/^<a:t>/, "").replace(/<\/a:t>$/, "")))
    .join(" ")
    .trim();
}

/** slide10.xml sorts before slide2.xml lexicographically — sort by the numeric part instead. */
function slideNumber(path: string): number {
  const match = path.match(/slide(\d+)\.xml$/);
  return match ? parseInt(match[1], 10) : Number.MAX_SAFE_INTEGER;
}

export async function extractPptxSlideTexts(base64Content: string): Promise<string[]> {
  try {
    const buffer = Buffer.from(base64Content, "base64");
    const zip = await JSZip.loadAsync(buffer);

    const slidePaths = Object.keys(zip.files)
      .filter((path) => /^ppt\/slides\/slide\d+\.xml$/.test(path))
      .sort((a, b) => slideNumber(a) - slideNumber(b));

    const slideTexts: string[] = [];
    for (const path of slidePaths) {
      const xml = await zip.files[path].async("string");
      slideTexts.push(extractTextFromSlideXml(xml));
    }

    console.log(`[PPTX-TEXT] Extracted text from ${slideTexts.length} slide(s).`);
    return slideTexts;
  } catch (error) {
    console.error(`[PPTX-TEXT] Extraction failed: ${(error as Error).message}`);
    return [];
  }
}
