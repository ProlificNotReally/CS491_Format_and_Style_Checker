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
        paragraphs.load("items/font,items/text,items/style,items/isInsideTable,items/paragraphFormat");
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

      let name = '';

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
        if (!formatting.allowHighlighting) {
          const hl = p.font.highlightColor;

          // FULL HIGHLIGHT
          const isFull = typeof hl === "string" && hl.startsWith("#");

          // MIXED HIGHLIGHT
          const isMixed = hl === "" || hl === undefined;

          // NO HIGHLIGHT
          const isNone = hl === null;

          const hasHighlight = isFull || isMixed;

          console.log("HighlightColor =", hl, "| full:", isFull, "| mixed:", isMixed);

          if (hasHighlight) {
            const name = `highlighting_${i + 1}`;
            range.insertBookmark(name);

            results.push({
              id: `s${s + 1}-p${i + 1}-highlight`,
              section: s + 1,
              type: "Highlighting",
              message: `Section ${s + 1}, Paragraph ${i + 1}: Contains highlighting.`,
              text,
              location: name,
              canLocate: true,
            });
          }
        }


        // --- Hidden text check (OOXML-based, unified message) ---
        if (!formatting.allowHiddenText) {
          const runRegex = /<w:r\b[^>]*>([\s\S]*?)<\/w:r>/g; // each run
          let hiddenTextRuns = 0;
          let m;

          while ((m = runRegex.exec(ooxml)) !== null) {
            const runXml = m[1];
            // Skip runs with no actual text
            if (!/<w:t\b[^>]*>[\s\S]*?<\/w:t>/.test(runXml)) continue;
            // Hidden text appears as <w:vanish/>
            const isHidden =
              /<w:vanish\b[^>]*\/>|<w:vanish\b[^>]*><\/w:vanish>/.test(runXml);
            if (isHidden) hiddenTextRuns++;
          }

          if (hiddenTextRuns > 0) {
            const name = `hiddentext_${i + 1}`;
            range.insertBookmark(name);

            results.push({
              id: `s${s + 1}-p${i + 1}-hidden`,
              section: s + 1,
              type: "Hidden Text",
              message: `Section ${s + 1}, Paragraph ${i + 1}: Hidden text detected.`,
              text,
              location: name,
              canLocate: true,
            });
          }
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

          name = `fontcolor_${i + 1}`;
          range.insertBookmark(name);

          invalidColors.forEach((badColor) => {
            results.push({
              id: `s${s + 1}-p${i + 1}-color-${badColor.replace("#", "")}`,
              section: s + 1,
              type: "Font Color",
              message: `Section ${s + 1}, Paragraph ${
                i + 1
              }: Contains disallowed font color "${badColor}".`,
              text,
              location: name,
              canLocate: true,
            });
          });
        }

        // --- Font size check ---
        const [minSize, maxSize] = formatting.allowedFontSizeRange;
        name = `fontsize_${i + 1}`;
        if (font.size == null || font.size === "") {
          range.insertBookmark(name);
          results.push({
            id: `s${s + 1}-p${i + 1}-multisize`,
            section: s + 1,
            type: "Font Size",
            message: `Section ${s + 1}, Paragraph ${
              i + 1
            }: Contains multiple font sizes.`,
            text,
            location: name,
            canLocate: true,
          });
        } else if (font.size < minSize || font.size > maxSize) {
          range.insertBookmark(name);
          results.push({
            id: `s${s + 1}-p${i + 1}-size`,
            section: s + 1,
            type: "Font Size",
            message: `Section ${s + 1}, Paragraph ${
              i + 1
            }: Font size ${font.size}pt is outside allowed range (${minSize}–${maxSize}).`,
            text,
            location: name,
            canLocate: true,
          });
        }

        // --- Font family check (simple + reliable) ---
        {
          const allowed = formatting.allowedFont.toLowerCase();

          // If Word cannot give a summary font, it means mixed formatting
          const paraFontRaw = p.font?.name || "";
          const paraFont = paraFontRaw.trim().toLowerCase();

          

          // CASE 1 — paragraph has mixed fonts (Word cannot summarize)
          if (!paraFont) {

            console.log("HERE");

            const name = `fontfamily_${i + 1}`;
            range.insertBookmark(name);

            results.push({
              id: `s${s + 1}-p${i + 1}-mixedfont`,
              section: s + 1,
              type: "Font",
              message: `Section ${s + 1}, Paragraph ${i + 1}: Contains multiple fonts. Expected "${formatting.allowedFont}".`,
              text,
              location: name,
              canLocate: true,
            });

            continue;
          }

          // CASE 2 — single font detected → compare to allowed font
          if (paraFont !== allowed) {
            const name = `fontfamily_${i + 1}`;
            range.insertBookmark(name);

            results.push({
              id: `s${s + 1}-p${i + 1}-wrongfont`,
              section: s + 1,
              type: "Font",
              message: `Section ${s + 1}, Paragraph ${i + 1}: Font "${paraFontRaw}" should be "${formatting.allowedFont}".`,
              text,
              location: name,
              canLocate: true,
            });
          }
        }


      // --- Justification check (XML-based, reliable) ---
      {
        // Skip tables (captions have style = "Caption")
        const text = p.text?.trim() || "";
        const styleName = (p.style || "").toString().toLowerCase();
        const isCaption = styleName === "caption";

        if (!text || p.isInsideTable || isCaption) {
          // allowed cases → skip
        } else {
          // XML analysis for alignment
          let align = "left"; // default if <w:jc> missing

          const jcMatch = ooxml.match(/<w:jc[^>]*w:val="([^"]+)"/);
          if (jcMatch) {
            const raw = jcMatch[1].toLowerCase();

            // Translate Word XML terms into human-friendly labels
            if (raw === "both") align = "justified";
            else if (raw === "center") align = "centered";
            else align = raw; // left, right
          }

          // Only LEFT is acceptable
          if (align !== "left") {
            const name = `justification_${i + 1}`;
            range.insertBookmark(name);

            results.push({
              id: `s${s + 1}-p${i + 1}-alignment`,
              section: s + 1,
              type: "Justification",
              message: `Section ${s + 1}, Paragraph ${i + 1}: Expected LEFT alignment, found "${align}".`,
              text,
              location: name,
              canLocate: true,
            });
          }
        }
      }

      /*
      // --- HYPERLINK DETECTION (XML <w:hyperlink> ONLY) ---
      if (!formatting.allowWebHyperlinks) {
        // Detect explicit Word hyperlink elements
        const hyperlinkRegex = /<w:hyperlink\b[^>]*r:id="([^"]+)"/g;
        const hyperlinkTags = [...ooxml.matchAll(hyperlinkRegex)];

        if (hyperlinkTags.length > 0) {
          const name = `hyper_${s + 1}_${i + 1}`;
          range.insertBookmark(name);

          const messages = hyperlinkTags.map(
            (m, idx) => `Hyperlink ${idx + 1} (id=${m[1]}) detected.`
          );

          results.push({
            id: `s${s + 1}-p${i + 1}-hyperlinks`,
            section: s + 1,
            type: "Hyperlink",
            message: `Section ${s + 1}, Paragraph ${i + 1}: Contains hyperlinks.\n` + messages.join("\n"),
            text,
            location: name,
            canLocate: true,
          });
        }
      }
      */




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
    // const range = location.getRange()
    //   ? location.getRange()
    //   : context.document.getSelection();

    const range = context.document.getBookmarkRange(location);
    range.select();
    await context.sync();    
  });
}