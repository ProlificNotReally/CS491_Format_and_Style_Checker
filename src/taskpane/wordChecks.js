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
    const blankPagesSeen = new Set();


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
        const runColorRegex =
          /<w:r\b[^>]*>[\s\S]*?<w:color[^>]*w:val="([^"]+)"[^>]*\/?[\s\S]*?<\/w:r>/gi;

        const colorMatches = [...ooxml.matchAll(runColorRegex)];
        const uniqueColors = [...new Set(colorMatches.map((m) => m[1].toLowerCase()))];

        if (uniqueColors.length > 0) {
          const invalidColors = uniqueColors.filter((c) => {
          if (c === "auto") return false; // auto = default = OK
          return !allowedColors.includes("#" + c);
        });


          if (invalidColors.length > 0) {
            const name = `fontcolor_${i + 1}`;
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
        const text = p.text?.trim() || "";
        if (!text) {
          // Ignore empty paragraphs
        } else {
          const styleName = (p.style || "").toString().toLowerCase();

          // Skip captions if you don't want to enforce justification on them
          const isCaption = styleName === "caption";

          if (!isCaption) {
            // Default alignment is "left" when no <w:jc/> is present
            let align = "left";
            const jcMatch = ooxml.match(/<w:jc[^>]*w:val="([^"]+)"/i);
            if (jcMatch) {
              const raw = jcMatch[1].toLowerCase();
              if (raw === "both") align = "justified";
              else if (raw === "center") align = "centered";
              else align = raw; // left, right, etc.
            }

            // Paragraph is inside a table?
            const insideTable = p.isInsideTable === true;

            // EXPECTED RULES:
            //  - inside table → CENTERED
            //  - outside table → JUSTIFIED
            const expectedAlign = insideTable ? "centered" : "justified";

            if (align !== expectedAlign) {
              const name = `justification_${s + 1}_${i + 1}`;
              range.insertBookmark(name);

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
      }




      
      // --- HYPERLINK DETECTION (XML) ---
      if (!formatting.allowWebHyperlinks) {
        const styleName = (p.style || "").toString().toLowerCase();
        const text = p.text?.trim() || "";

        // 1) Skip obvious TOC / list styles
        const isTOCStyle =
          styleName.startsWith("toc ") ||
          styleName.includes("table of contents") ||
          styleName.includes("list of tables") ||
          styleName.includes("list of figures");

        // 2) Basic hyperlink markers
        const hasExplicitTag = /<w:hyperlink\b/i.test(ooxml);
        const hasSimpleField = /<w:fldSimple[^>]*HYPERLINK\b/i.test(ooxml);
        const hasInstrText = /<w:instrText[^>]*>[^<]*HYPERLINK\b[^<]*/i.test(ooxml);

        // Complex field: begin + "HYPERLINK" somewhere in the field code
        const hasComplexField =
          /<w:fldChar[^>]*w:fldCharType="begin"[^>]*>/i.test(ooxml) &&
          /HYPERLINK\b/i.test(ooxml);

        // 3) INTERNAL cross-references we want to ALLOW:
        //    - HYPERLINK \l "_Ref..." or "_Toc..." (anchors)
        //    - REF _Ref..., REF _Toc..., PAGEREF _Toc...
        const isInternalAnchor =
          /HYPERLINK\s+\\l\s+"(_Ref|_Toc)[^"]*"/i.test(ooxml);

        const isCrossRef =
          /\bREF\s+_Ref\d+/i.test(ooxml) ||
          /\bREF\s+_Toc\d+/i.test(ooxml) ||
          /\bPAGEREF\s+_Toc\d+/i.test(ooxml);

        const shouldSkip = isTOCStyle || isInternalAnchor || isCrossRef;

        const isHyperlinked =
          hasExplicitTag || hasSimpleField || hasInstrText || hasComplexField;

        if (isHyperlinked && !shouldSkip) {
          const name = `hyper_${s + 1}_${i + 1}`;
          range.insertBookmark(name);

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


    {
    console.log("===== SECTION OOXML START =====");
    const oox = body.getOoxml();
    await context.sync();
    console.log(oox.value);
    console.log("===== SECTION OOXML END =====");
}

      // --- BLANK PAGE DETECTION (OOXML PAGE SPLIT) ---
      // --- BLANK PAGE DETECTION (OOXML PAGE SPLIT) ---
{
    const sectionXmlObj = body.getOoxml();
    await context.sync();
    const xml = sectionXmlObj.value || "";

    // TRUE page breaks only
    let pages = xml.split(/<w:br[^>]*w:type="page"[^>]*>/g);

    const pageHasContent = (xml) =>
        /<w:t\b[^>]*>[ \t\r\n]*\S[\s\S]*?<\/w:t>/.test(xml);

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

            let firstPara = paragraphs.items[0];
            const name = `blankpage_${s + 1}_${p + 1}`;
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
    // const range = location.getRange()
    //   ? location.getRange()
    //   : context.document.getSelection();

    const range = context.document.getBookmarkRange(location);
    range.select();
    await context.sync();    
  });
}