import { checkHeaderFooterFormatting } from "./checkHeaderFooterFormatting";
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

    // Load all sections
    const sections = context.document.sections;
    sections.load("items/body");
    await context.sync();

    // Iterate through sections
    for (let s = 0; s < sections.items.length; s++) {
      const section = sections.items[s];
      const body = section.body;
      const paragraphs = body.paragraphs;

      if (paragraphs && typeof paragraphs.load === "function") {
        paragraphs.load("items/font,items/paragraphFormat,items/text");
        await context.sync();
      }

      // Handle empty section
      if (!paragraphs?.items || paragraphs.items.length === 0) {
        results.push({
          id: `s${s + 1}-empty`,
          section: s + 1,
          type: "Info",
          message: `Section ${s + 1}: No paragraphs found.`,
          text: "",
          canLocate: false,
        });
        continue;
      }

      // Batch-load OOXML for all paragraphs (1 sync total per section)
      const ooxmlObjs = paragraphs.items.map((p) => p.getOoxml());
      await context.sync();

      // Pre-lowercase allowed colors for faster checks
      const allowedColors = formatting.allowedFontColors.map((c) =>
        c.toLowerCase()
      );

      // Loop through paragraphs with pre-fetched XML
      for (let i = 0; i < paragraphs.items.length; i++) {
        const p = paragraphs.items[i];
        const ooxml = ooxmlObjs[i].value || "";
        const text = p.text?.trim() || "";
        const font = p.font || {};
        const alignment = p.paragraphFormat
          ? p.paragraphFormat.alignment
          : "Unknown";
        const range = p.getRange();

        // --- Highlighting check ---
        if (
          !formatting.allowHighlighting &&
          (font.highlightColor && font.highlightColor !== "None")
        ) {
          results.push({
            id: `s${s + 1}-p${i + 1}-highlight`,
            section: s + 1,
            type: "Highlighting",
            message: `Section ${s + 1}, Paragraph ${
              i + 1
            }: Highlighting detected.`,
            text,
            location: range,
            canLocate: true,
          });
        }

        // --- Hidden text check ---
        if (!formatting.allowHiddenText && font.hidden) {
          results.push({
            id: `s${s + 1}-p${i + 1}-hidden`,
            section: s + 1,
            type: "Hidden Text",
            message: `Section ${s + 1}, Paragraph ${
              i + 1
            }: Hidden text detected.`,
            text,
            location: range,
            canLocate: true,
          });
        }

        // --- Font color check (OOXML regex-based) ---
        const colorMatches = [
          ...ooxml.matchAll(/<w:color[^>]*w:val="([^"]+)"/g),
        ];
        const uniqueColors = [
          ...new Set(colorMatches.map((m) => m[1].toLowerCase())),
        ];

        if (uniqueColors.length > 0) {
          const invalidColors = uniqueColors.filter(
            (c) => !allowedColors.includes("#" + c)
          );

          invalidColors.forEach((badColor) => {
            results.push({
              id: `s${s + 1}-p${i + 1}-color-${badColor.replace("#", "")}`,
              section: s + 1,
              type: "Font Color",
              message: `Section ${s + 1}, Paragraph ${
                i + 1
              }: Contains disallowed font color "${badColor}".`,
              text,
              location: range,
              canLocate: true,
            });
          });
        }

        // --- Font size check ---
        const [minSize, maxSize] = formatting.allowedFontSizeRange;
        if (font.size == null || font.size === "") {
          results.push({
            id: `s${s + 1}-p${i + 1}-multisize`,
            section: s + 1,
            type: "Font Size",
            message: `Section ${s + 1}, Paragraph ${
              i + 1
            }: Contains multiple or undefined font sizes.`,
            text,
            location: range,
            canLocate: true,
          });
        } else if (font.size < minSize || font.size > maxSize) {
          results.push({
            id: `s${s + 1}-p${i + 1}-size`,
            section: s + 1,
            type: "Font Size",
            message: `Section ${s + 1}, Paragraph ${
              i + 1
            }: Font size ${font.size}pt is outside allowed range (${minSize}–${maxSize}).`,
            text,
            location: range,
            canLocate: true,
          });
        }

        // --- Font family check ---
        if (font.name && font.name !== formatting.allowedFont) {
          results.push({
            id: `s${s + 1}-p${i + 1}-font`,
            section: s + 1,
            type: "Font",
            message: `Section ${s + 1}, Paragraph ${
              i + 1
            }: Font "${font.name}" should be "${formatting.allowedFont}".`,
            text,
            location: range,
            canLocate: true,
          });
        }

        // --- Justification check ---
        if (
          formatting.enforceTextJustification &&
          alignment !== "Justified" &&
          alignment !== "Left" &&
          alignment !== "Unknown"
        ) {
          results.push({
            id: `s${s + 1}-p${i + 1}-alignment`,
            section: s + 1,
            type: "Justification",
            message: `Section ${s + 1}, Paragraph ${
              i + 1
            }: Inconsistent text justification (${alignment}).`,
            text,
            location: range,
            canLocate: true,
          });
        }
      }

      // --- Table header check ---
      let tables;
      try {
        tables = body.tables;
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
                id: `s${s + 1}-table-${i + 1}-header`,
                section: s + 1,
                type: "Table Header",
                message: `Section ${s + 1}, Table ${
                  i + 1
                }: Header row is not set to repeat on each page.`,
                canLocate: false,
              });
            }
          });
        }
      }
    }

    // If still no issues, add success message
    if (results.length === 0) {
      results.push({
        id: "info-clean",
        type: "Info",
        message: "✅ No issues found or document is empty.",
        text: "",
        canLocate: false,
      });
    }

    
    const headerFooterIssues = await checkHeaderFooterFormatting(context);
    results.push(...headerFooterIssues);
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
    const range = location.getRange()
      ? location.getRange()
      : context.document.getSelection();
    console.log(range);

    range.select();
    context.sync();
  });
}