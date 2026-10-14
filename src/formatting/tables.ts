/**
 * tables.ts — batch image operations (no native one-click equivalent).
 * Table-to-text and repeat-headers removed as native-duplicate
 * (Table Tools Layout covers both).
 */
import { safeWordRun } from "../shared/wordUtils";

export async function resizeImages(percent: number): Promise<{ data?: number; error?: string }> {
  if (!(percent >= 10 && percent <= 400)) {
    return { error: "Percent must be between 10 and 400." };
  }
  return safeWordRun(async (context) => {
    const pictures = context.document.body.inlinePictures;
    pictures.load("items, width, height");
    await context.sync();
    for (const pic of pictures.items) {
      pic.width = (pic.width * percent) / 100;
      pic.height = (pic.height * percent) / 100;
    }
    await context.sync();
    return pictures.items.length;
  });
}
