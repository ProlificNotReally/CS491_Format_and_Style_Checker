import rules from "./config/rules.json";

/**
 * Check margins across all sections
 * Note: Microsoft Word API does not support programmatically setting margins,
 * so this checker only identifies issues and provides manual instructions.
 */
export async function checkMargins() {
  return Word.run(async (context) => {
    const results = [];
    const { margins } = rules;

    if (!margins.enforceMargins) {
      return results;
    }

    const sections = context.document.sections;
    sections.load("items/pageSetup");
    await context.sync();

    for (let i = 0; i < sections.items.length; i++) {
      const section = sections.items[i];
      const s = section.pageSetup;
      const orientation = s.orientation;
      let isValid = true;

      if (orientation === Word.PageOrientation.portrait) {
        isValid =
          s.topMargin === margins.topMarginPortrait &&
          s.bottomMargin === margins.bottomMarginPortrait &&
          s.leftMargin === margins.leftMarginPortrait &&
          s.rightMargin === margins.rightMarginPortrait;
      } else if (orientation === Word.PageOrientation.landscape) {
        isValid =
          s.topMargin === margins.topMarginLandscape &&
          s.bottomMargin === margins.bottomMarginLandscape &&
          s.leftMargin === margins.leftMarginLandscape &&
          s.rightMargin === margins.rightMarginLandscape;
      }

      if (!isValid) {
        // Create bookmark for navigation
        const range = section.body.getRange();
        const bk = `margin_issue_${i + 1}`;
        range.insertBookmark(bk);
        await context.sync();

        const expectedMargins = orientation === Word.PageOrientation.portrait
          ? {
              top: margins.topMarginPortrait / 72,
              bottom: margins.bottomMarginPortrait / 72,
              left: margins.leftMarginPortrait / 72,
              right: margins.rightMarginPortrait / 72
            }
          : {
              top: margins.topMarginLandscape / 72,
              bottom: margins.bottomMarginLandscape / 72,
              left: margins.leftMarginLandscape / 72,
              right: margins.rightMarginLandscape / 72
            };

        results.push({
          id: `margins-${i + 1}`,
          type: "Margins",
          message: `Section ${i + 1} (${orientation}): Incorrect margins. Expected - Top: ${expectedMargins.top}", Bottom: ${expectedMargins.bottom}", Left: ${expectedMargins.left}", Right: ${expectedMargins.right}"`,
          location: bk,
          canLocate: true,
        });
      }
    }

    return results;
  });
}
