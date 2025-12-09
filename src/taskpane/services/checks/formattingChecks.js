import rules from "../../config/rules.json";

/**
 * Runs all formatting checks based on rules.json.
 * Returns an array of issue objects with optional location ranges
 * that can be used to jump to the offending text in Word.
 */
export async function analyzeFormatting() {
  return Word.run(async (context) => {
    const { formatting } = rules;
    const results = [];
    const blankPagesSeen = new Set();
    const bookmarksToCreate = []; // Batch bookmark creation

    // Pre-compile regexes outside loops (performance optimization)
    const runRegex = /<w:r\b[^>]*>([\s\S]*?)<\/w:r>/g;
    const textRegex = /<w:t\b[^>]*>[\s\S]*?<\/w:t>/;
    const vanishRegex = /<w:vanish\b[^>]*\/>|<w:vanish\b[^>]*><\/w:vanish>/;
    const runColorRegex = /<w:r\b[^>]*>[\s\S]*?<w:color[^>]*w:val="([^"]+)"[^>]*\/?[\s\S]*?<\/w:r>/gi;
    const jcRegex = /<w:jc[^>]*w:val="([^"]+)"/i;

    // Pre-lowercase allowed colors once
    const allowedColors = formatting.allowedFontColors.map((c) => c.toLowerCase());
    const allowedFont = formatting.allowedFont.toLowerCase();
    const [minSize, maxSize] = formatting.allowedFontSizeRange;

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

      let name = '';

      // Loop through paragraphs with pre-fetched XML
      for (let i = 0; i < paragraphs.items.length; i++) {
        const p = paragraphs.items[i];
        const ooxml = ooxmlObjs[i].value || "";
        const text = p.text?.trim() || "";
        const font = p.font || {};
        const range = p.getRange();

        // --- Highlighting check ---
        if (!formatting.allowHighlighting) {
          const hl = p.font.highlightColor;
          const hasHighlight = (typeof hl === "string" && hl.startsWith("#")) || (hl === "" || hl === undefined);

          if (hasHighlight) {
            name = `highlighting_${s + 1}_${i + 1}`;
            bookmarksToCreate.push({ range, name });

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

        // --- Hidden text check (OOXML-based, optimized) ---
        if (!formatting.allowHiddenText && ooxml.includes("w:vanish")) {
          let hiddenTextRuns = 0;
          runRegex.lastIndex = 0; // Reset regex
          let m;

          while ((m = runRegex.exec(ooxml)) !== null) {
            const runXml = m[1];
            if (!textRegex.test(runXml)) continue;
            if (vanishRegex.test(runXml)) hiddenTextRuns++;
          }

          if (hiddenTextRuns > 0) {
            name = `hiddentext_${s + 1}_${i + 1}`;
            bookmarksToCreate.push({ range, name });

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

        // --- Font color check (optimized with early exit) ---
        if (ooxml.includes("w:color")) {
          runColorRegex.lastIndex = 0; // Reset regex
          const colorMatches = [...ooxml.matchAll(runColorRegex)];
          
          if (colorMatches.length > 0) {
            const uniqueColors = [...new Set(colorMatches.map((m) => m[1].toLowerCase()))];
            const invalidColors = uniqueColors.filter((c) => {
              if (c === "auto") return false;
              return !allowedColors.includes("#" + c);
            });

            if (invalidColors.length > 0) {
              name = `fontcolor_${s + 1}_${i + 1}`;
              bookmarksToCreate.push({ range, name });

              invalidColors.forEach((badColor) => {
                results.push({
                  id: `s${s + 1}-p${i + 1}-color-${badColor.replace("#", "")}`,
                  section: s + 1,
                  type: "Font Color",
                  message: `Section ${s + 1}, Paragraph ${i + 1}: Contains disallowed font color "${badColor}".`,
                  text,
                  location: name,
                  canLocate: true,
                });
              });
            }
          }
        }

        // --- Font size check ---
        name = `fontsize_${s + 1}_${i + 1}`;
        if (font.size == null || font.size === "") {
          bookmarksToCreate.push({ range, name });
          results.push({
            id: `s${s + 1}-p${i + 1}-multisize`,
            section: s + 1,
            type: "Font Size",
            message: `Section ${s + 1}, Paragraph ${i + 1}: Contains multiple font sizes.`,
            text,
            location: name,
            canLocate: true,
          });
        } else if (font.size < minSize || font.size > maxSize) {
          bookmarksToCreate.push({ range, name });
          results.push({
            id: `s${s + 1}-p${i + 1}-size`,
            section: s + 1,
            type: "Font Size",
            message: `Section ${s + 1}, Paragraph ${i + 1}: Font size ${font.size}pt is outside allowed range (${minSize}–${maxSize}).`,
            text,
            location: name,
            canLocate: true,
          });
        }

        // --- Font family check (optimized) ---
        const paraFontRaw = p.font?.name || "";
        const paraFont = paraFontRaw.trim().toLowerCase();

        if (!paraFont) {
          name = `fontfamily_${s + 1}_${i + 1}`;
          bookmarksToCreate.push({ range, name });

          results.push({
            id: `s${s + 1}-p${i + 1}-mixedfont`,
            section: s + 1,
            type: "Font",
            message: `Section ${s + 1}, Paragraph ${i + 1}: Contains multiple fonts. Expected "${formatting.allowedFont}".`,
            text,
            location: name,
            canLocate: true,
          });
        } else if (paraFont !== allowedFont) {
          name = `fontfamily_${s + 1}_${i + 1}`;
          bookmarksToCreate.push({ range, name });

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


      // --- Justification check (XML-based, optimized) ---
      if (text) {
        const styleName = (p.style || "").toString().toLowerCase();
        const isCaption = styleName === "caption";

        if (!isCaption) {
          // Default alignment is "left" when no <w:jc/> is present
          let align = "left";
          jcRegex.lastIndex = 0; // Reset regex
          const jcMatch = jcRegex.exec(ooxml);
          if (jcMatch) {
            const raw = jcMatch[1].toLowerCase();
            if (raw === "both") align = "justified";
            else if (raw === "center") align = "centered";
            else align = raw;
          }

          const insideTable = p.isInsideTable === true;
          const expectedAlign = insideTable ? "centered" : "justified";

          if (align !== expectedAlign) {
            name = `justification_${s + 1}_${i + 1}`;
            bookmarksToCreate.push({ range, name });

            results.push({
              id: `s${s + 1}-p${i + 1}-alignment`,
              section: s + 1,
              type: "Justification",
              message: insideTable
                ? `Section ${s + 1}, Paragraph ${i + 1}: Expected CENTER alignment for table text, found "${align}".`
                : `Section ${s + 1}, Paragraph ${i + 1}: Expected JUSTIFIED alignment, found "${align}".`,
              text,
              location: name,
              canLocate: true,
            });
          }
        }
      }

      // --- HYPERLINK DETECTION (XML - optimized with early exit) ---
      if (!formatting.allowWebHyperlinks && text && (ooxml.includes("w:hyperlink") || ooxml.includes("HYPERLINK"))) {
        const styleName = (p.style || "").toString().toLowerCase();

        const isTOCStyle =
          styleName.startsWith("toc ") ||
          styleName.includes("table of contents") ||
          styleName.includes("list of tables") ||
          styleName.includes("list of figures");

        const hasExplicitTag = ooxml.includes("<w:hyperlink");
        const hasSimpleField = ooxml.includes("w:fldSimple") && ooxml.includes("HYPERLINK");
        const hasInstrText = ooxml.includes("w:instrText") && ooxml.includes("HYPERLINK");
        const hasComplexField = ooxml.includes('w:fldCharType="begin"') && ooxml.includes("HYPERLINK");

        const isInternalAnchor = /HYPERLINK\s+\\l\s+"(_Ref|_Toc)[^"]*"/i.test(ooxml);
        const isCrossRef = /\b(REF|PAGEREF)\s+_(Ref|Toc)\d+/i.test(ooxml);

        const shouldSkip = isTOCStyle || isInternalAnchor || isCrossRef;
        const isHyperlinked = hasExplicitTag || hasSimpleField || hasInstrText || hasComplexField;

        if (isHyperlinked && !shouldSkip) {
          name = `hyper_${s + 1}_${i + 1}`;
          bookmarksToCreate.push({ range, name });

          results.push({
            id: `s${s + 1}-p${i + 1}-hyperlinks`,
            section: s + 1,
            type: "Hyperlink",
            message: `Section ${s + 1}, Paragraph ${i + 1}: Contains at least one web hyperlink.`,
            text,
            location: name,
            canLocate: true,
          });
        }
      }
    }

    // Batch create all bookmarks (single sync instead of many)
    if (bookmarksToCreate.length > 0) {
      for (const bookmark of bookmarksToCreate) {
        bookmark.range.insertBookmark(bookmark.name);
      }
      await context.sync();
    }

    // --- BLANK PAGE DETECTION (OOXML PAGE SPLIT - optimized) ---
    const sectionXmlObj = body.getOoxml();
    await context.sync();
    const xml = sectionXmlObj.value || "";

    // TRUE page breaks only
    const pages = xml.split(/<w:br[^>]*w:type="page"[^>]*>/g);
    const pageHasContent = (xml) => /<w:t\b[^>]*>[ \t\r\n]*\S[\s\S]*?<\/w:t>/.test(xml);

    for (let p = 0; p < pages.length; p++) {
      const pageXml = pages[p];

      // Skip trailing non-body content (rels, styles, theme, settings)
      if (!/<w:p\b/.test(pageXml) && !/<w:body\b/.test(pageXml)) {
        continue;
      }

      // Detect actual blank page
      if (!pageHasContent(pageXml)) {
        const key = `${s + 1}-${p + 1}`;
        if (blankPagesSeen.has(key)) continue;
        blankPagesSeen.add(key);

        const firstPara = paragraphs.items[0];
        name = `blankpage_${s + 1}_${p + 1}`;
        firstPara.getRange().insertBookmark(name);

        results.push({
          id: `s${s + 1}-blankpage-${p + 1}`,
          section: s + 1,
          type: "Blank Page",
          message: `Section ${s + 1}: Page ${p + 1} is blank.`,
          location: name,
          text: "",
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
    
  return results;
  });
}
