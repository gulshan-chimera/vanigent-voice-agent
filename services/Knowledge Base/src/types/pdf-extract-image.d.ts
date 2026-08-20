declare module "pdf-extract-image" {
  export function extractImagesFromPdf(input: Buffer | string): Promise<Buffer[]>;
}