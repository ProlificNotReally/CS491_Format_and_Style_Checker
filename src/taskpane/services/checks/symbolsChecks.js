import rules from "../../config/rules.json";

/**
 * Checks for invalid use of Symbol font in document.
 * Symbol font should not be used for regular text content.
 * Returns an array of issue objects with location bookmarks.
 */
export async function checkSymbols() {
  const { symbols } = rules;

  return Word.run(async (context) => {
    const results = [];
    let name = '';

    // Only run check if symbol font is not allowed
    if (!symbols.allowSymbolFont) {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items/font/name");
      await context.sync();

      // Collect all paragraphs using Symbol font
      const symbolParagraphIndexes = [];
      for (let i = 0; i < paragraphs.items.length; i++) {
        const p = paragraphs.items[i];
        if (p.font && p.font.name && p.font.name.toLowerCase() === "symbol") {
          symbolParagraphIndexes.push(i);
        }
      }

      // Load text for paragraphs with Symbol font
      const rangesToLoad = [];
      for (const idx of symbolParagraphIndexes) {
        const para = paragraphs.items[idx];
        const range = para.getRange();
        range.load("text");
        rangesToLoad.push({ range, index: idx });
      }

      // If Symbol font paragraphs found, create bookmarks and flag errors
      if (rangesToLoad.length > 0) {
        await context.sync();
        
        for (let j = 0; j < rangesToLoad.length; j++) {
          const { range, index } = rangesToLoad[j];
          name = `symbolfont_${index + 1}`;
          range.insertBookmark(name);

          results.push({
            id: `symbol-${index + 1}`,
            type: "Symbols",
            message: `Paragraph ${index + 1}: Styled with Symbol font (not allowed).`,
            text: range.text || "(empty)",
            location: name,
            canLocate: true,
          });
        }
      }
    }

    // Return results (empty array if no issues)
    return results;
  });
}
