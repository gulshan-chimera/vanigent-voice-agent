// src/lib/pdfImageDetect.ts
//
// Determines which pages contain images LARGE ENOUGH to carry meaning,
// by walking each page's operator list and tracking the current
// transformation matrix — which gives each image's rendered size in PDF
// points. This distinguishes a header logo (~60pt) from a screenshot or
// chart (~300-500pt), which a simple "has an image" check cannot.
//
// Detects RASTER images only. Vector charts are path operations, not
// images — pathOpCount is reported so that blind spot stays visible.

import * as pdfjsLib from "pdfjs-dist/legacy/build/pdf.mjs";

export interface PageVisualInfo {
  page: number;              // 1-based
  imageOpCount: number;
  largestImageWidth: number;  // PDF points
  largestImageHeight: number;
  largestImageArea: number;
  pathOpCount: number;
  charCount: number;
}

type Matrix = [number, number, number, number, number, number];

const IDENTITY: Matrix = [1, 0, 0, 1, 0, 0];

/** Standard PDF matrix multiply: applies `m` on top of `base`. */
function multiply(m: Matrix, base: Matrix): Matrix {
  return [
    m[0] * base[0] + m[1] * base[2],
    m[0] * base[1] + m[1] * base[3],
    m[2] * base[0] + m[3] * base[2],
    m[2] * base[1] + m[3] * base[3],
    m[4] * base[0] + m[5] * base[2] + base[4],
    m[4] * base[1] + m[5] * base[3] + base[5],
  ];
}

function opCode(name: string): number | undefined {
  return (pdfjsLib.OPS as Record<string, number | undefined>)[name];
}

function codeSet(names: string[]): Set<number> {
  const codes = new Set<number>();
  for (const name of names) {
    const code = opCode(name);
    if (typeof code === "number") codes.add(code);
  }
  return codes;
}

export async function detectPageVisuals(base64Content: string): Promise<PageVisualInfo[]> {
  try {
    const pdfUint8Array = new Uint8Array(Buffer.from(base64Content, "base64"));
    const document = await pdfjsLib.getDocument({ data: pdfUint8Array }).promise;

    const imageOps = codeSet([
      "paintImageXObject",
      "paintImageXObjectRepeat",
      "paintInlineImageXObject",
      "paintInlineImageXObjectGroup",
      "paintImageMaskXObject",
      "paintImageMaskXObjectRepeat",
      "paintImageMaskXObjectGroup",
      "paintJpegXObject",
    ]);
    const pathOps = codeSet([
      "constructPath",
      "fill",
      "stroke",
      "eoFill",
      "fillStroke",
      "shadingFill",
    ]);

    const saveOp = opCode("save");
    const restoreOp = opCode("restore");
    const transformOp = opCode("transform");

    const info: PageVisualInfo[] = [];

    for (let pageNum = 1; pageNum <= document.numPages; pageNum++) {
      const page = await document.getPage(pageNum);
      const opList = await page.getOperatorList();

      let ctm: Matrix = [...IDENTITY] as Matrix;
      const stack: Matrix[] = [];

      let imageOpCount = 0;
      let pathOpCount = 0;
      let largestW = 0;
      let largestH = 0;
      let largestArea = 0;

      for (let i = 0; i < opList.fnArray.length; i++) {
        const fn = opList.fnArray[i];

        if (fn === saveOp) {
          stack.push([...ctm] as Matrix);
        } else if (fn === restoreOp) {
          ctm = stack.pop() ?? ([...IDENTITY] as Matrix);
        } else if (fn === transformOp) {
          const args = opList.argsArray[i] as number[];
          if (args && args.length >= 6) {
            ctm = multiply(args.slice(0, 6) as Matrix, ctm);
          }
        } else if (imageOps.has(fn)) {
          imageOpCount++;
          // An image is painted into the unit square, so the CTM's scale
          // components give its rendered size in points.
          const width = Math.hypot(ctm[0], ctm[1]);
          const height = Math.hypot(ctm[2], ctm[3]);
          const area = width * height;
          if (area > largestArea) {
            largestArea = area;
            largestW = width;
            largestH = height;
          }
        } else if (pathOps.has(fn)) {
          pathOpCount++;
        }
      }

      const textContent = await page.getTextContent();
      const charCount = textContent.items
        .map((item) => ("str" in item ? item.str : ""))
        .join(" ")
        .trim().length;

      info.push({
        page: pageNum,
        imageOpCount,
        largestImageWidth: Math.round(largestW),
        largestImageHeight: Math.round(largestH),
        largestImageArea: Math.round(largestArea),
        pathOpCount,
        charCount,
      });
    }

    console.log(`[PDF-VISUAL-DETECT] Analysed ${info.length} page(s).`);
    return info;
  } catch (error) {
    console.error(`[PDF-VISUAL-DETECT] Detection failed: ${(error as Error).message}`);
    return [];
  }
}