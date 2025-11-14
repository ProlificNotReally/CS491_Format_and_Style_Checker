import rules from "./config/rules.json";

export async function checkHeaderFooterFormatting(context) {
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

    if (section.pageSetup.orientation === "Landscape") {
      const range = section.body.getRange();
      const bkmark = `landscape_section_${i + 1}`;
      range.insertBookmark(bkmark);
      await context.sync();

      results.push({
        id: `landscape-page-${i + 1}`,
        type: "Layout",
        message: `Section ${i + 1} is in Landscape orientation.`,
        location: bkmark,
        canLocate: true,
      });
    }

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

  // Header consistency check
  const baseHeader = headerTexts[0]?.text;
  for (let i = 1; i < headerTexts.length; i++) {
    if (headerTexts[i].text !== baseHeader) {
      const range = headerTexts[i].para.getRange();
      const bk = `inconsistent_header_${i + 1}`;
      range.insertBookmark(bk);
      await context.sync();
      results.push({
        id: `inconsistent-header-${i}`,
        type: "Header",
        message: `Section ${headerTexts[i].index + 1} header is inconsistent with others.`,
        location: bk,
        canLocate: true,
      });
    }
  }

  // Footer consistency check
  if (footerTexts.length > 1) {
    const baseFooter = footerTexts[0]?.text;
    for (let i = 1; i < footerTexts.length; i++) {
      if (footerTexts[i].text !== baseFooter) {
        const range = footerTexts[i].para.getRange();
        const bk = `inconsistent_footer_${i + 1}`;
        range.insertBookmark(bk);
        await context.sync();
        results.push({
          id: `inconsistent-footer-${i}`,
          type: "Footer",
          message: `Section ${footerTexts[i].index + 1} footer is inconsistent with others.`,
          location: bk,
          canLocate: true,
        });
      }
    }
  }

  return results;
}
