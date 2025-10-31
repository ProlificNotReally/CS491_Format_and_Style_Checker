import rules from "./config/rules.json";

/**
 * Runs all formatting checks based on rules.json.
 * Returns an array of issue objects with optional location ranges
 * that can be used to jump to the offending text in Word.
 */
export async function analyzeFormatting() {
  return Word.run(async (context) => {
    const { formatting } = rules;
    const results = [];

    //Load paragraphs safely
    const paragraphs = context.document.body.paragraphs;
    if (paragraphs && typeof paragraphs.load === "function") {
      paragraphs.load("items/font,items/paragraphFormat,items/text");
      await context.sync();
    }

    //Handle empty document
    if (!paragraphs?.items || paragraphs.items.length === 0) {
      results.push({
        id: "info-empty",
        type: "Info",
        message: "No paragraphs found in the document.",
        text: "",
        canLocate: false
      });
      return results;
    }

    //Loop through paragraphs
    paragraphs.items.forEach((p, i) => {
      const text = p.text?.trim() || "";
      const font = p.font || {};
      const alignment = p.paragraphFormat
        ? p.paragraphFormat.alignment
        : "Unknown";

      //Create a reusable paragraph range reference
      const range = p.getRange();

      //Highlighting check
      if (
        !formatting.allowHighlighting &&
        font.highlightColor &&
        font.highlightColor !== "None"
      ) {
        results.push({
          id: `p${i + 1}-highlight`,
          type: "Highlighting",
          message: `Paragraph ${i + 1}: Highlighting detected.`,
          text,
          location: range,
          canLocate: true
        });
      }

      //Hidden text check
      if (!formatting.allowHiddenText && font.hidden) {
        results.push({
          id: `p${i + 1}-hidden`,
          type: "Hidden Text",
          message: `Paragraph ${i + 1}: Hidden text detected.`,
          text,
          location: range,
          canLocate: true
        });
      }

      //Font color check
      const allowedColors = formatting.allowedFontColors.map((c) =>
        c.toLowerCase()
      );
      if (
        font.color &&
        !allowedColors.includes(font.color.toLowerCase())
      ) {
        results.push({
          id: `p${i + 1}-color`,
          type: "Font Color",
          message: `Paragraph ${i + 1}: Font color "${font.color}" is not allowed.`,
          text,
          location: range,
          canLocate: true
        });
      }

      //Font size range check
      const [minSize, maxSize] = formatting.allowedFontSizeRange;
      if (font.size < minSize || font.size > maxSize) {
        results.push({
          id: `p${i + 1}-size`,
          type: "Font Size",
          message: `Paragraph ${i + 1}: Font size ${font.size}pt is outside the allowed range (${minSize}–${maxSize}).`,
          text,
          location: range,
          canLocate: true
        });
      }

      //Font family check
      if (font.name && font.name !== formatting.allowedFont) {
        results.push({
          id: `p${i + 1}-font`,
          type: "Font",
          message: `Paragraph ${i + 1}: Font "${font.name}" should be "${formatting.allowedFont}".`,
          text,
          location: range,
          canLocate: true
        });
      }

      //Text justification consistency
      if (
        formatting.enforceTextJustification &&
        alignment !== "Justified" &&
        alignment !== "Left" &&
        alignment !== "Unknown"
      ) {
        results.push({
          id: `p${i + 1}-alignment`,
          type: "Justification",
          message: `Paragraph ${i + 1}: Inconsistent text justification (${alignment}).`,
          text,
          location: range,
          canLocate: true
        });
      }
    });

    await context.sync();

    //Check for tables safely
    let tables;
    try {
      tables = context.document.body.tables;
    } catch {
      tables = null;
    }

    if (tables && typeof tables.load === "function") {
      tables.load("items/rows/items/rowFormat/repeatAsHeaderRow");
      await context.sync();

      if (tables.items?.length > 0) {
        tables.items.forEach((table, i) => {
          const headerRow = table.rows?.items?.[0];
          if (
            formatting.enforceRepeatingHeaderRowsForContinuousTables &&
            headerRow?.rowFormat &&
            !headerRow.rowFormat.repeatAsHeaderRow
          ) {
            results.push({
              id: `table-${i + 1}-header`,
              type: "Table Header",
              message: `Table ${i + 1}: Header row is not set to repeat on each page.`,
              canLocate: false
            });
          }
        });
      }
    }

    //Check hyperlinks safely
    let hyperlinks;
    try {
      hyperlinks = context.document.body.hyperlinks;
    } catch {
      hyperlinks = null;
    }

    if (hyperlinks && typeof hyperlinks.load === "function") {
      hyperlinks.load("items/address");
      await context.sync();

      if (hyperlinks.items?.length > 0) {
        hyperlinks.items.forEach((link, i) => {
          if (!formatting.allowWebHyperlinks && link.address) {
            results.push({
              id: `link-${i + 1}`,
              type: "Hyperlink",
              message: `Hyperlink detected: ${link.address}`,
              canLocate: false
            });
          }
        });
      }
    }

    //If still no issues, add success message
    if (results.length === 0) {
      results.push({
        id: "info-clean",
        type: "Info",
        message: "✅ No issues found or document is empty.",
        text: "",
        canLocate: false
      });
    }

    return results;
  });
}

/**
 * Navigates to a given error location inside the Word document.
 * @param {Word.Range} location - The Word range for the error.
 */
export async function goToError(location) {
  if (!location) return;

  await Word.run(async (context) => {
    // Rehydrate the range from the original document
    const range = location.getRange
      ? location.getRange()
      : context.document.getSelection();

    range.select();
    context.sync();
  });
}
