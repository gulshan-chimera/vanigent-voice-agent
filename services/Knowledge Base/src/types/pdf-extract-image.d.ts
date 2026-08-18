declare module "pdf-extract-image" {
  export function extractImagesFromPdf(input: Buffer | Uint8Array | string): Promise<Buffer[]>;
}