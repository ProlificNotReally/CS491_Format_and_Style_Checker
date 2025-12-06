import rules from "./config/rules.json";

export async function checkHeaderFooterFormatting() {
  return Word.run(async (context) => {
    const results = [];
    const { formatting, header, footer } = rules;

    const sections = context.document.sections;
    sections.load("items");
    await context.sync();

    let headerTexts = [];
    let footerTexts = [];

    for (let i = 0; i < sections.items.length; i++) {
      const section = sections.items[i];

      // Check for Landscape orientation
      section.pageSetup.load("orientation");
      section.pageSetup.load(["headerDistance", "footerDistance"]);
      await context.sync();

      const expectedDistance = 36; // 0.5 inches in points

      if (Math.abs(section.pageSetup.headerDistance - expectedDistance) > 0.1) {
        const range = section.body.getRange();
        const bk = `header_margin_${i + 1}`;
        range.insertBookmark(bk);
        await context.sync();

        results.push({
          id: `header-margin-${i + 1}`,
          type: "Header",
          message: `Section ${i + 1} header margin is not set to 0.5 inches.`,
          location: bk,
          canLocate: true
        });
      }

      if (Math.abs(section.pageSetup.footerDistance - expectedDistance) > 0.1) {
        const range = section.body.getRange();
        const bk = `footer_margin_${i + 1}`;
        range.insertBookmark(bk);
        await context.sync();

        results.push({
          id: `footer-margin-${i + 1}`,
          type: "Footer",
          message: `Section ${i + 1} footer margin is not set to 0.5 inches.`,
          location: bk,
          canLocate: true
        });
      }

      await context.sync();

      const headerObj = section.getHeader("Primary");
      const footerObj = section.getFooter("Primary");

      const headerParas = headerObj.paragraphs;
      const footerParas = footerObj.paragraphs;

      headerParas.load("items");
      footerParas.load("items");
      await context.sync();

      const headerCount = headerParas.items.length;

      if (headerCount !== header.requiredLineCount) {
        for (let j = 0; j < headerParas.items.length; j++) {
          const p = headerParas.items[j];
          const range = p.getRange();
          const bkmark = `headerlinecount_${i + 1}_${j + 1}`;
          range.insertBookmark(bkmark);
          await context.sync();

          results.push({
            id: `header-lines-${i}-${j}`,
            type: "Header",
            message: `Section ${i + 1} header: Expected ${header.requiredLineCount} lines, found ${headerCount}`,
            location: bkmark,
            canLocate: true,
          });
        }
      }

      for (let j = 0; j < Math.min(headerCount, header.requiredLineCount); j++) {
        const para = headerParas.items[j];
        const range = para.getRange();
        para.load(["alignment", "font/underline", "font/name", "font/size", "font/color", "text"]);
        await context.sync();

        if ((para.text || "").toLowerCase().includes("draft")) {
          const range = para.getRange();
          const bk = `header_draft_${i + 1}_${j + 1}`;
          range.insertBookmark(bk);
          await context.sync();

          results.push({
            id: `header-draft-${i + 1}-${j + 1}`,
            type: "Header",
            message: `Section ${i + 1} header line ${j + 1} contains the word "Draft".`,
            location: bk,
            canLocate: true
          });
        }

        const { name, size, color, underline } = para.font;

        if (para.alignment !== header.alignment) {
          const bkmark = `headeralignment_${i + 1}_${j + 1}`;
          range.insertBookmark(bkmark);
          await context.sync();
          results.push({
            id: `header-align-${i}-${j}`,
            type: "Header",
            message: `Section ${i + 1} header line ${j + 1}: Not ${header.alignment}-aligned`,
            location: bkmark,
            canLocate: true,
          });
        }

        if (j === 1 && header.secondLineUnderline && underline === "None") {
          const bkmark = `headerunderline_${i + 1}_${j + 1}`;
          range.insertBookmark(bkmark);
          await context.sync();
          results.push({
            id: `header-underline-${i}`,
            type: "Header",
            message: `Section ${i + 1} header line 2: Not underlined`,
            location: bkmark,
            canLocate: true,
          });
        }

        if (name !== formatting.allowedFont) {
          const bkmark = `headerfont_${i + 1}_${j + 1}`;
          range.insertBookmark(bkmark);
          await context.sync();
          results.push({
            id: `header-font-${i}-${j}`,
            type: "Header",
            message: `Section ${i + 1} header line ${j + 1}: Font not ${formatting.allowedFont}`,
            location: bkmark,
            canLocate: true,
          });
        }

        if (size < formatting.allowedFontSizeRange[0] || size > formatting.allowedFontSizeRange[1]) {
          const bkmark = `headerfontsize_${i + 1}_${j + 1}`;
          range.insertBookmark(bkmark);
          await context.sync();
          results.push({
            id: `header-size-${i}-${j}`,
            type: "Header",
            message: `Section ${i + 1} header line ${j + 1}: Font size ${size} not in allowed range`,
            location: bkmark,
            canLocate: true,
          });
        }

        const allowedColors = formatting.allowedFontColors.map(c => c.toLowerCase());
        if (!allowedColors.includes((color || "").toLowerCase())) {
          const bkmark = `headerfontcolor_${i + 1}_${j + 1}`;
          range.insertBookmark(bkmark);
          await context.sync();
          results.push({
            id: `header-color-${i}-${j}`,
            type: "Header",
            message: `Section ${i + 1} header line ${j + 1}: Font color "${color}" not allowed`,
            location: bkmark,
            canLocate: true,
          });
        }

        const text = para.text || "";
        if (header.pageNumberPreTab && /\d+$/.test(text)) {
          const numIndex = text.search(/\d+$/);
          const before = text.slice(0, numIndex);
          if (!before.includes("\t") && /  +$/.test(before)) {
            const bkmark = `headertabs_${i + 1}_${j + 1}`;
            range.insertBookmark(bkmark);
            await context.sync();
            results.push({
              id: `header-tab-${i}-${j}`,
              type: "Header",
              message: `Section ${i + 1} header line ${j + 1}: Page number not preceded by TAB`,
              location: bkmark,
              canLocate: true,
            });
          }
        }
      }

      for (let k = 0; k < footerParas.items.length; k++) {
        const para = footerParas.items[k];
        const range = para.getRange();
        para.load(["inlinePictures", "text", "alignment"]);
        await context.sync();

        const pics = para.inlinePictures;
        pics.load("items");
        await context.sync();

        if (pics.items.length > 0) {
          const text = para.text || "";
          const isCentered = para.alignment === "Centered" || text.includes("\t");
          if (footer.imageMustBeCentered && !isCentered) {
            const bkmark = `footeralignment_${i + 1}_${k + 1}`;
            range.insertBookmark(bkmark);
            await context.sync();
            results.push({
              id: `footer-image-${i}-${k}`,
              type: "Footer",
              message: `Section ${i + 1} footer line ${k + 1}: Inline image not centered`,
              location: bkmark,
              canLocate: true,
            });
          }
        }
      }

      // Gather cleaned content for consistency check
      const headerText = headerParas.items.map(p => (p.text || "").replace(/\\d+/g, "").trim().toLowerCase()).join("||");
      const footerText = footerParas.items.map(p => (p.text || "").replace(/\\d+/g, "").trim().toLowerCase()).join("||");

      headerTexts.push({ text: headerText, para: headerParas.items[0], index: i });

      const hasFooterText = footerParas.items.some(p => (p.text || "").trim().length > 0);
      if (hasFooterText) {
        footerTexts.push({ text: footerText, para: footerParas.items[0], index: i });
        // --- Append (Landscape) note if section is landscape and has errors ---
        const isLandscape = section.pageSetup.orientation === "Landscape";
        if (isLandscape) {
          results.forEach(entry => {
            if (entry.message.includes(`Section ${i + 1}`)) {
              entry.message += " (Landscape)";
            }
          });
        }
      }
    }

  // --- Split header/footer groups by orientation ---
  let portraitHeaders = [], landscapeHeaders = [];
  let portraitFooters = [], landscapeFooters = [];

  for (let i = 0; i < sections.items.length; i++) {
    const section = sections.items[i];
    const headerObj = section.getHeader("Primary");
    const footerObj = section.getFooter("Primary");

    const headerParas = headerObj.paragraphs;
    const footerParas = footerObj.paragraphs;

    headerParas.load("items/text");
    footerParas.load("items/text");
    await context.sync();

    const headerText = headerParas.items.map(p => (p.text || "").replace(/\\d+/g, "").trim().toLowerCase()).join("||");
    const footerText = footerParas.items.map(p => (p.text || "").replace(/\\d+/g, "").trim().toLowerCase()).join("||");

    const hasFooterText = footerParas.items.some(p => (p.text || "").trim().length > 0);
    const group = section.pageSetup.orientation === "Landscape" ? "landscape" : "portrait";

    const headerEntry = { text: headerText, para: headerParas.items[0], index: i };
    const footerEntry = { text: footerText, para: footerParas.items[0], index: i };

    if (group === "landscape") {
      landscapeHeaders.push(headerEntry);
      if (hasFooterText) landscapeFooters.push(footerEntry);
    } else {
      portraitHeaders.push(headerEntry);
      if (hasFooterText) portraitFooters.push(footerEntry);
    }
  }

  // --- Function to check consistency within group ---
  async function checkGroupConsistency(group, label) {
    const base = group[0]?.text;
    for (let i = 1; i < group.length; i++) {
      if (group[i].text !== base) {
        const range = group[i].para.getRange();
        const bk = `inconsistent_${label}_${group[i].index + 1}`;
        range.insertBookmark(bk);
        await context.sync();
        results.push({
          id: `inconsistent-${label}-${group[i].index + 1}`,
          type: label.charAt(0).toUpperCase() + label.slice(1),
          message: `Section ${group[i].index + 1} ${label} is inconsistent with other ${label}s of same orientation.`,
          location: bk,
          canLocate: true
        });
      }
    }
  }

  // Apply to both portrait and landscape separately
  await checkGroupConsistency(portraitHeaders, "header");
  await checkGroupConsistency(landscapeHeaders, "header");
  await checkGroupConsistency(portraitFooters, "footer");
  await checkGroupConsistency(landscapeFooters, "footer");


    return results;
  });
}