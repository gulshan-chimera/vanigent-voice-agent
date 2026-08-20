// src/lib/pdfText.ts
//
// Extracts plain text from PDF file content. Reuses the PII redaction
// approach from the earlier reference sync script.

import { PDFParse } from "pdf-parse";

const PII_PATTERNS: { name: string; pattern: RegExp }[] = [
  { name: "Email", pattern: /[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/g },
  { name: "Phone", pattern: /(\+?(\d[\s\-.]?){10,14}\d)/g },
  { name: "Aadhaar", pattern: /\b\d{4}\s?\d{4}\s?\d{4}\b/g },
  { name: "PAN", pattern: /\b[A-Z]{5}[0-9]{4}[A-Z]{1}\b/g },
  { name: "SSN", pattern: /\b\d{3}[- ]?\d{2}[- ]?\d{4}\b/g },
  { name: "CreditCard", pattern: /\b(?:\d[ \-]?){15,16}\b/g },
];

function redactPII(text: string): { redacted: string; found: string[] } {
  let redacted = text;
  const found: string[] = [];

  for (const { name, pattern } of PII_PATTERNS) {
    const matches = redacted.match(pattern);
    if (matches && matches.length > 0) {
      found.push(`${name} (${matches.length})`);
      redacted = redacted.replace(pattern, "[REDACTED]");
    }
  }

  return { redacted, found };
}

export async function extractPdfText(base64Content: string): Promise<string | null> {
  try {
    const buffer = Buffer.from(base64Content, "base64");
    const parser = new PDFParse({ data: buffer });

    let rawText: string;
    try {
      const data = await parser.getText();
      rawText = data.text;
    } finally {
      await parser.destroy();
    }

    if (!rawText || rawText.trim().length === 0) {
      console.log("[PDF-TEXT] No extractable text — PDF may be image-based (needs vision model later).");
      return null;
    }

    const { redacted, found } = redactPII(rawText);
    if (found.length > 0) {
      console.log(`[PDF-TEXT] PII redacted: ${found.join(" | ")}`);
    }

    return redacted;
  } catch (error) {
    console.error(`[PDF-TEXT] Extraction failed: ${(error as Error).message}`);
    return null;
  }
}
