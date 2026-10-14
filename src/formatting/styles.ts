/**
 * styles.ts — style application engine (used by profiles; style browsing and
 * creation UI removed as native-duplicate — Word's Styles gallery covers them).
 */
import { getSelectionParagraphsOrThrow, safeWordRun } from "../shared/wordUtils";

export async function applyStyle(styleName: string): Promise<{ error?: string }> {
  return safeWordRun(async (context) => {
    const paragraphs = await getSelectionParagraphsOrThrow(context);
    for (const para of paragraphs.items) {
      para.style = styleName;
    }
    await context.sync();
  });
}
