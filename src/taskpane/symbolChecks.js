import rules from "./config/rules.json";

/**
 * Runs symbols check to detect valid symbols that are NOT styled with Symbol font
 * Valid symbols should be in Symbol font - if they're not, it's an error
 * Returns an array of issue objects
 */
export async function checkSymbols() {
  return Word.run(async (context) => {
    const results = [];

    // Common mathematical and special symbols that should be in Symbol font
    const validSymbols = [
      // Greek letters
      'α', 'β', 'γ', 'δ', 'ε', 'ζ', 'η', 'θ', 'ι', 'κ', 'λ', 'μ', 'ν', 'ξ', 'ο', 'π', 'ρ', 'σ', 'τ', 'υ', 'φ', 'χ', 'ψ', 'ω',
      'Α', 'Β', 'Γ', 'Δ', 'Ε', 'Ζ', 'Η', 'Θ', 'Ι', 'Κ', 'Λ', 'Μ', 'Ν', 'Ξ', 'Ο', 'Π', 'Ρ', 'Σ', 'Τ', 'Υ', 'Φ', 'Χ', 'Ψ', 'Ω',
      // Mathematical operators
      '±', '×', '÷', '≠', '≈', '≡', '≤', '≥', '∞', '∑', '∏', '∫', '∂', '√', '∆',
      // Arrows
      '←', '→', '↑', '↓', '↔', '⇐', '⇒', '⇔',
      // Other symbols
      '°', '′', '″', '∠', '⊥', '∥', '⊂', '⊃', '⊆', '⊇', '∈', '∉', '∅', '∩', '∪',
      '¬', '∧', '∨', '∀', '∃', '⊕', '⊗', '℃', '℉', '†', '‡', '§', '¶', '•', '◦'
    ];

    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items");
    await context.sync();

    let issueCount = 0;

    for (let i = 0; i < paragraphs.items.length; i++) {
      const para = paragraphs.items[i];
      para.load("text");
      await context.sync();

      const text = para.text || "";
      
      // Check if paragraph contains any valid symbols
      for (const symbol of validSymbols) {
        if (text.includes(symbol)) {
          // Found a valid symbol, now check if it's using Symbol font
          const range = para.getRange();
          range.load("font/name");
          await context.sync();

          // If the paragraph contains the symbol but is NOT in Symbol font, flag it
          if (range.font.name.toLowerCase() !== "symbol") {
            const name = `symbol_not_styled_${issueCount + 1}`;
            range.insertBookmark(name);
            await context.sync();

            results.push({
              id: `symbol-missing-font-${issueCount + 1}`,
              type: "Symbols",
              message: `Symbol "${symbol}" found but not styled with Symbol font (currently using ${range.font.name}).`,
              text: text.substring(0, 100) + (text.length > 100 ? "..." : ""),
              location: name,
              canLocate: true,
            });

            issueCount++;
            break; // Only report once per paragraph
          }
        }
      }
    }

    return results;
  });
}

/**
 * Fixes a single symbol issue by changing the font to Symbol
 */
export async function fixSymbolIssue(issue) {
  return Word.run(async (context) => {
    try {
      // Parse the issue ID to get the bookmark name
      const idMatch = issue.id.match(/symbol-missing-font-(\d+)/);
      if (!idMatch) {
        console.error("Could not parse symbol issue ID:", issue.id);
        return { success: false, error: "Invalid issue ID format" };
      }

      const bookmarkName = `symbol_not_styled_${idMatch[1]}`;

      // Get the bookmarked range
      const bookmarks = context.document.body.bookmarks;
      bookmarks.load("items");
      await context.sync();

      const bookmark = bookmarks.items.find(b => b.name === bookmarkName);
      if (!bookmark) {
        console.error("Bookmark not found:", bookmarkName);
        return { success: false, error: "Bookmark not found" };
      }

      const range = bookmark.getRange();
      range.load("font");
      await context.sync();

      // Change font to Symbol
      range.font.name = "Symbol";
      await context.sync();

      return { success: true };
    } catch (error) {
      console.error("Error fixing symbol issue:", error);
      return { success: false, error: error.message };
    }
  });
}

/**
 * Fixes all symbol issues by changing fonts to Symbol
 */
export async function fixAllSymbolIssues(issues) {
  const results = [];
  
  for (const issue of issues) {
    const result = await fixSymbolIssue(issue);
    results.push({
      issue,
      success: result.success,
      error: result.error
    });
  }

  return results;
}
