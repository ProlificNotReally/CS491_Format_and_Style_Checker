import rules from "./config/rules.json";

/**
 * Runs all formatting checks across all sections based on rules.json.
 * Returns an array of issue objects, each tagged with the section number.
 */
export async function analyzeFormatting() {
  return Word.run(async (context) => {
    const { formatting } = rules;
    const results = [];

    //Load all sections
    const sections = context.document.sections;
    sections.load("items");
    await context.sync();

    if (!sections.items || sections.items.length === 0) {
      results.push({
        id: "no-sections",
        type: "Info",
        message: "No sections found in the document.",
        canLocate: false,
      });
      return results;
    }

    //Iterate through each section
    for (let s = 0; s < sections.items.length; s++) {
      const section = sections.items[s];
      const sectionNum = s + 1;

      //Load paragraphs for this section
      const paragraphs = section.body.paragraphs;
      if (paragraphs && typeof paragraphs.load === "function") {
        paragraphs.load("items/font,items/paragraphFormat,items/text");
        await context.sync();
      }

      //Handle empty sections
      if (!paragraphs?.items || paragraphs.items.length === 0) {
        results.push({
          id: `section-${sectionNum}-empty`,
          section: sectionNum,
          type: "Info",
          message: `Section ${sectionNum} is empty.`,
          text: "",
          canLocate: false,
        });
        continue;
      }

      //Loop through paragraphs safely
      for (let i = 0; i < paragraphs.items.length; i++) {
        const p = paragraphs.items[i];
        const text = p.text?.trim() || "";
        const font = p.font || {};
        const alignment = p.paragraphFormat
          ? p.paragraphFormat.alignment
          : "Unknown";
        const range = p.getRange();

        //Highlighting check
        if (
          !formatting.allowHighlighting &&
          (p.font.highlightColor ||
            p.font.highlightColor === undefined) &&
          p.font.highlightColor !== "None" &&
          p.font.highlightColor !== null
        ) {
          results.push({
            id: `s${sectionNum}-p${i + 1}-highlight`,
            section: sectionNum,
            type: "Highlighting",
            message: `Section ${sectionNum}, Paragraph ${
              i + 1
            }: Highlighting detected.`,
            text,
            location: range,
            canLocate: true,
          });
        }

        //Hidden text check
        if (!formatting.allowHiddenText && font.hidden) {
          results.push({
            id: `s${sectionNum}-p${i + 1}-hidden`,
            section: sectionNum,
            type: "Hidden Text",
            message: `Section ${sectionNum}, Paragraph ${
              i + 1
            }: Hidden text detected.`,
            text,
            location: range,
            canLocate: true,
          });
        }

        //Font color check
        const allowedColors = formatting.allowedFontColors.map((c) =>
          c.toLowerCase()
        );

        //console.log("FONT COLOR >> ", p.font.color);

        if (p.font.color && p.font.color !== "None") {
          //Case A: uniform color
          const colorLower = p.font.color.toLowerCase();

          if (!allowedColors.includes(colorLower)) {
            results.push({
              id: `s${sectionNum}-p${i + 1}-color`,
              section: sectionNum,
              type: "Font Color",
              message: `Section ${sectionNum}, Paragraph ${
                i + 1
              }: Font color "${p.font.color}" is not allowed.`,
              text,
              location: range,
              canLocate: true,
            });
          }
        } else if (!p.font.color || p.font.color.trim() === "") {
          
          //Case B: mixed colors → check runs
          const runs = p.getRange().getTextRanges(["Any"], true);
          runs.load("items/font/color,items/text");
          await context.sync();

          const invalidColors = new Set();
          const seenColors = new Set();

          runs.items.forEach((run) => {
            const color = run.font.color?.toLowerCase();
            //console.log("color - ", color);
            if (!color || color === "none") return;
            seenColors.add(color);
            if (!allowedColors.includes(color)) invalidColors.add(color);
          });

          if (invalidColors.size > 0) {
            invalidColors.forEach((badColor) => {
              results.push({
                id: `s${sectionNum}-p${i + 1}-color-${badColor.replace("#", "")}`,
                section: sectionNum,
                type: "Font Color",
                message: `Section ${sectionNum}, Paragraph ${
                  i + 1
                }: Contains disallowed font color "${badColor}".`,
                text,
                location: range,
                canLocate: true,
              });
            });
          }
        }

        //Font size range and consistency check
        const [minSize, maxSize] = formatting.allowedFontSizeRange;
        if (p.font.size === null) {
          results.push({
            id: `s${sectionNum}-p${i + 1}-multisize`,
            section: sectionNum,
            type: "Font Size",
            message: `Section ${sectionNum}, Paragraph ${
              i + 1
            }: Multiple font sizes detected within the same paragraph.`,
            text,
            location: range,
            canLocate: true,
          });
        } else if (p.font.size < minSize || p.font.size > maxSize) {
          results.push({
            id: `s${sectionNum}-p${i + 1}-size`,
            section: sectionNum,
            type: "Font Size",
            message: `Section ${sectionNum}, Paragraph ${
              i + 1
            }: Font size ${p.font.size}pt is outside the allowed range (${minSize}–${maxSize}).`,
            text,
            location: range,
            canLocate: true,
          });
        }

        //Font family check
        if (p.font.name === null) {
          results.push({
            id: `s${sectionNum}-p${i + 1}-multifont`,
            section: sectionNum,
            type: "Font",
            message: `Section ${sectionNum}, Paragraph ${
              i + 1
            }: Multiple font types detected within the same paragraph.`,
            text,
            location: range,
            canLocate: true,
          });
        } else if (
          p.font.name &&
          p.font.name !== formatting.allowedFont
        ) {
          results.push({
            id: `s${sectionNum}-p${i + 1}-font`,
            section: sectionNum,
            type: "Font",
            message: `Section ${sectionNum}, Paragraph ${
              i + 1
            }: Font "${p.font.name}" should be "${formatting.allowedFont}".`,
            text,
            location: range,
            canLocate: true,
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
            id: `s${sectionNum}-p${i + 1}-alignment`,
            section: sectionNum,
            type: "Justification",
            message: `Section ${sectionNum}, Paragraph ${
              i + 1
            }: Inconsistent text justification (${alignment}).`,
            text,
            location: range,
            canLocate: true,
          });
        }
      } //end paragraph loop

      //Table checks for this section
      let tables;
      try {
        tables = section.body.tables;
      } catch {
        tables = null;
      }

      if (tables && typeof tables.load === "function") {
        tables.load("items/rows/items/rowFormat/repeatAsHeaderRow");
        await context.sync();

        if (tables.items?.length > 0) {
          tables.items.forEach((table, t) => {
            const headerRow = table.rows?.items?.[0];
            if (
              formatting.enforceRepeatingHeaderRowsForContinuousTables &&
              headerRow?.rowFormat &&
              !headerRow.rowFormat.repeatAsHeaderRow
            ) {
              results.push({
                id: `s${sectionNum}-t${t + 1}-header`,
                section: sectionNum,
                type: "Table Header",
                message: `Section ${sectionNum}, Table ${
                  t + 1
                }: Header row is not set to repeat on each page.`,
                canLocate: false,
              });
            }
          });
        }
      }

      //Hyperlink checks for this section
      if (!formatting.allowWebHyperlinks) {
        for (let i = 0; i < paragraphs.items.length; i++) {
          const p = paragraphs.items[i];
          if (!p.text || !p.text.toLowerCase().includes("http")) continue;

          const runs = p.getRange().getTextRanges(["Any"], true);
          runs.load("items/hyperlink");
          await context.sync();

          if (runs.items && runs.items.length > 0) {
            runs.items.forEach((run, j) => {
              if (run.hyperlink && run.hyperlink.address) {
                results.push({
                  id: `s${sectionNum}-p${i + 1}-link${j + 1}`,
                  section: sectionNum,
                  type: "Hyperlink",
                  message: `Section ${sectionNum}, Paragraph ${
                    i + 1
                  }: Hyperlink detected - ${run.hyperlink.address}`,
                  text: p.text,
                  location: run.getRange(),
                  canLocate: true,
                });
              }
            });
          }
        }
      }
    } //end section loop

    //If no issues found at all
    if (results.length === 0) {
      results.push({
        id: "info-clean",
        type: "Info",
        message: "✅ No issues found across all sections.",
        canLocate: false,
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
    const range = location.getRange
      ? location.getRange()
      : context.document.getSelection();

    range.select();
    context.sync();
  });
}
