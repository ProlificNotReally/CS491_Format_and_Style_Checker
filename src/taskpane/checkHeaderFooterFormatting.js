import rules from "./config/rules.json";

export async function checkHeaderFooterFormatting() {
  return Word.run(async (context) => {
    const results = [];
    const { formatting, header, footer } = rules;

    const sections = context.document.sections;
    sections.load("items");
    await context.sync();

    // --- First pass: per-section checks ---
    for (let i = 0; i < sections.items.length; i++) {
      const section = sections.items[i];

      // Load page setup (orientation + header/footer distances)
      section.pageSetup.load(["orientation", "headerDistance", "footerDistance"]);
      await context.sync();

      const expectedDistance = 36; // 0.5 inches in points

      // Header margin distance
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
          canLocate: true,
        });
      }

      // Footer margin distance
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
          canLocate: true,
        });
      }

      const headerObj = section.getHeader("Primary");
      const footerObj = section.getFooter("Primary");

      const headerParas = headerObj.paragraphs;
      const footerParas = footerObj.paragraphs;

      // We need header text loaded so we can ignore empty paragraphs
      headerParas.load("items/text");
      footerParas.load("items");
      await context.sync();

      // Only count HEADER lines that actually have text
      const nonEmptyHeaderParas = headerParas.items.filter(
        (p) => (p.text || "").trim().length > 0
      );
      const headerCount = nonEmptyHeaderParas.length;

      // If line count is wrong, create ONE issue anchored to the first non-empty line
      if (
        typeof header.requiredLineCount === "number" &&
        headerCount !== header.requiredLineCount &&
        nonEmptyHeaderParas.length > 0
      ) {
        const firstPara = nonEmptyHeaderParas[0];
        const range = firstPara.getRange();
        const bkmark = `headerlinecount_${i + 1}_1`;
        range.insertBookmark(bkmark);
        await context.sync();

        results.push({
          id: `header-lines-${i}`, // one issue per section now
          type: "Header",
          message: `Section ${i + 1} header: Expected ${header.requiredLineCount} lines, found ${headerCount}`,
          location: bkmark,
          canLocate: true,
        });
      }


      // --- Per-header-line formatting checks ---
      for (let j = 0; j < Math.min(headerCount, header.requiredLineCount); j++) {
        const para = nonEmptyHeaderParas[j];
        const range = para.getRange();
        para.load([
          "alignment",
          "font/underline",
          "font/name",
          "font/size",
          "font/color",
          "text",
        ]);
        await context.sync();

        const text = para.text || "";

        // "Draft" in header text
        if (text.toLowerCase().includes("draft")) {
          const bk = `header_draft_${i + 1}_${j + 1}`;
          range.insertBookmark(bk);
          await context.sync();

          results.push({
            id: `header-draft-${i + 1}-${j + 1}`,
            type: "Header",
            message: `Section ${i + 1} header line ${j + 1} contains the word "Draft".`,
            location: bk,
            canLocate: true,
          });
        }

        const { name, size, color, underline } = para.font;

        // Alignment
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

        // Second line underline
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

        // Font family
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

        // Font size
        if (
          size < formatting.allowedFontSizeRange[0] ||
          size > formatting.allowedFontSizeRange[1]
        ) {
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

        // Font color
        const allowedColors = formatting.allowedFontColors.map((c) => c.toLowerCase());
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

        // Page number preceding TAB
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

      // --- Footer checks (inline image centered) ---
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
    }

    // --- Second pass: consistency by orientation ---
    const portraitHeaders = [];
    const landscapeHeaders = [];
    const portraitFooters = [];
    const landscapeFooters = [];

    for (let i = 0; i < sections.items.length; i++) {
      const section = sections.items[i];
      const headerObj = section.getHeader("Primary");
      const footerObj = section.getFooter("Primary");

      const headerParas = headerObj.paragraphs;
      const footerParas = footerObj.paragraphs;

      headerParas.load("items/text");
      footerParas.load("items/text");
      await context.sync();

      const headerText = headerParas.items
        .map((p) => (p.text || "").replace(/\d+/g, "").trim().toLowerCase())
        .join("||");
      const footerText = footerParas.items
        .map((p) => (p.text || "").replace(/\d+/g, "").trim().toLowerCase())
        .join("||");

      const hasFooterText = footerParas.items.some((p) => (p.text || "").trim().length > 0);
      const isLandscape = section.pageSetup.orientation === "Landscape";

      const headerEntry = { text: headerText, sectionIndex: i };
      const footerEntry = { text: footerText, sectionIndex: i };

      if (isLandscape) {
        landscapeHeaders.push(headerEntry);
        if (hasFooterText) landscapeFooters.push(footerEntry);
      } else {
        portraitHeaders.push(headerEntry);
        if (hasFooterText) portraitFooters.push(footerEntry);
      }
    }

    async function checkGroupConsistency(group, label) {
      if (!group.length) return;
      const base = group[0].text;
      for (let i = 1; i < group.length; i++) {
        if (group[i].text !== base) {
          const idx = group[i].sectionIndex;
          const section = sections.items[idx];
          const obj = label === "header" ? section.getHeader("Primary") : section.getFooter("Primary");
          const firstPara = obj.paragraphs.getFirst();
          const range = firstPara.getRange();
          const bk = `inconsistent_${label}_${idx + 1}`;
          range.insertBookmark(bk);
          await context.sync();
          results.push({
            id: `inconsistent-${label}-${idx + 1}`,
            type: label.charAt(0).toUpperCase() + label.slice(1),
            message: `Section ${idx + 1} ${label} is inconsistent with other ${label}s of same orientation.`,
            location: bk,
            canLocate: true,
          });
        }
      }
    }

    await checkGroupConsistency(portraitHeaders, "header");
    await checkGroupConsistency(landscapeHeaders, "header");
    await checkGroupConsistency(portraitFooters, "footer");
    await checkGroupConsistency(landscapeFooters, "footer");

    return results;
  });
}

export async function fixHeaderFooterIssue(issue) {
  return Word.run(async (context) => {
    if (!issue || !issue.location || !issue.id) return;

    const { formatting, header, footer } = rules;
    const doc = context.document;
    const id = issue.id;

    // ❗ FIX: load must use a single string or an array, not multiple args
    const range = doc.getBookmarkRange(issue.location);
    range.load(["text", "paragraphs", "sections"]);
    await context.sync();

    const sections = range.sections;
    sections.load("items/pageSetup");
    await context.sync();

    if (!sections.items.length) return;

    const section = sections.items[0];

    // Some bookmarks (like margin issues) may not be inside a paragraph;
    // guard against that.
    let para = null;
    try {
      para = range.paragraphs.getFirst();
      para.load(["alignment", "font", "text"]);
      await context.sync();
    } catch (e) {
      // If there is no paragraph, only margin/inconsistent fixes will run.
    }

    const ensureSectionMargins = () => {
      try {
        section.pageSetup.headerDistance = 36; // 0.5"
        section.pageSetup.footerDistance = 36; // 0.5"
      } catch (e) {}
    };

    // 1) Margin issues
    if (id.startsWith("header-margin-")) {
      try {
        section.pageSetup.headerDistance = 36;
      } catch (e) {}
      ensureSectionMargins();
      await context.sync();
      return;
    }

    if (id.startsWith("footer-margin-")) {
      try {
        section.pageSetup.footerDistance = 36;
      } catch (e) {}
      ensureSectionMargins();
      await context.sync();
      return;
    }

    // If we need a paragraph for the remaining fixes and we don't have one, bail.
    if (!para) return;

    // 2) Header alignment
    if (id.startsWith("header-align-")) {
      if (header && header.alignment) {
        para.alignment = header.alignment; // e.g. "Left"
      }
      ensureSectionMargins();
      await context.sync();
      return;
    }

    // 3) Footer inline image alignment
    if (id.startsWith("footer-image-")) {
      if (footer && footer.imageMustBeCentered) {
        para.alignment = "Centered"; // must match detection
      }
      ensureSectionMargins();
      await context.sync();
      return;
    }

    // 4) Header underline (second line)
    if (id.startsWith("header-underline-")) {
      para.font.underline = "Single";
      ensureSectionMargins();
      await context.sync();
      return;
    }

    // 5) Header font family
    if (id.startsWith("header-font-")) {
      if (formatting && formatting.allowedFont) {
        para.font.name = formatting.allowedFont;
      }
      ensureSectionMargins();
      await context.sync();
      return;
    }

    // 6) Header font size
    if (id.startsWith("header-size-")) {
      if (formatting && Array.isArray(formatting.allowedFontSizeRange)) {
        const [minSize, maxSize] = formatting.allowedFontSizeRange;
        let target = para.font.size || maxSize || 12;
        if (target < minSize) target = minSize;
        if (target > maxSize) target = maxSize;
        para.font.size = target;
      }
      ensureSectionMargins();
      await context.sync();
      return;
    }

    // 7) Header font color
    if (id.startsWith("header-color-")) {
      if (
        formatting &&
        Array.isArray(formatting.allowedFontColors) &&
        formatting.allowedFontColors.length > 0
      ) {
        para.font.color = formatting.allowedFontColors[0];
      }
      ensureSectionMargins();
      await context.sync();
      return;
    }

    // 8) Header tab before page number
    if (id.startsWith("header-tab-")) {
      const original = para.text || "";
      const fixed = original.replace(/\s*(\d+)\s*$/, "\t$1");
      para.insertText(fixed, "Replace");
      ensureSectionMargins();
      await context.sync();
      return;
    }

    // 9) Header contains "Draft"
    if (id.startsWith("header-draft-")) {
      const original = para.text || "";
      const fixed = original.replace(/draft/gi, "").replace(/\s+/g, " ").trim();
      para.insertText(fixed, "Replace");
      ensureSectionMargins();
      await context.sync();
      return;
    }

    // 10) Header line-count issues
    if (id.startsWith("header-lines-")) {
      const headerObj = section.getHeader("Primary");
      const headerParas = headerObj.paragraphs;
      headerParas.load("items");
      await context.sync();

      if (header && typeof header.requiredLineCount === "number") {
        const required = header.requiredLineCount;
        if (headerParas.items.length > required) {
          for (let idx = required; idx < headerParas.items.length; idx++) {
            headerParas.items[idx].getRange().clear();
          }
        }
      }

      ensureSectionMargins();
      await context.sync();
      return;
    }

    // 11) Inconsistent header/footer → Link to Previous (same orientation only)
    if (
      id.startsWith("inconsistent-header-") ||
      id.startsWith("inconsistent-footer-")
    ) {
      const isHeader = id.startsWith("inconsistent-header-");
      const parts = id.split("-");
      const sectionNumber = parseInt(parts[2], 10); // 1-based
      if (!Number.isFinite(sectionNumber)) return;

      const allSections = doc.sections;
      allSections.load("items/pageSetup");
      await context.sync();

      const idx = sectionNumber - 1;
      if (idx < 0 || idx >= allSections.items.length) return;

      const currentSection = allSections.items[idx];
      const currentOrientation = currentSection.pageSetup.orientation;

      // Only link/copy from the immediately previous section if it has the SAME orientation.
      if (idx > 0) {
        const prevSection = allSections.items[idx - 1];
        const prevOrientation = prevSection.pageSetup.orientation;

        if (prevOrientation === currentOrientation) {
          if (isHeader) {
            const currentHeader = currentSection.getHeader("Primary");
            try {
              currentHeader.linkToPrevious = true;
            } catch (e) {
              const prevHeader = prevSection.getHeader("Primary");
              const prevRange = prevHeader.getRange();
              prevRange.load("text");
              await context.sync();

              const thisRange = currentHeader.getRange();
              thisRange.clear();
              thisRange.insertText((prevRange.text || "").trim(), "Start");
            }
          } else {
            const currentFooter = currentSection.getFooter("Primary");
            try {
              currentFooter.linkToPrevious = true;
            } catch (e) {
              const prevFooter = prevSection.getFooter("Primary");
              const prevRange = prevFooter.getRange();
              prevRange.load("text");
              await context.sync();

              const thisRange = currentFooter.getRange();
              thisRange.clear();
              thisRange.insertText((prevRange.text || "").trim(), "Start");
            }
          }

          ensureSectionMargins();
          await context.sync();
          return;
        }
      }

      // If there is no previous section of the same orientation,
      // do NOT link/copy anything here. The final orientation sync will handle it.
      return;
    }
  });
}


export async function fixAllHeaderFooterIssues() {
  return Word.run(async (context) => {
    const { formatting, header, footer } = rules;
    const doc = context.document;

    const sections = doc.sections;
    sections.load("items/pageSetup");
    await context.sync();

    if (!sections.items.length) return;

    const groups = {};

    for (let i = 0; i < sections.items.length; i++) {
      const section = sections.items[i];
      const orientation = section.pageSetup.orientation || "Unknown";
      const headerBody = section.getHeader("Primary");
      const footerBody = section.getFooter("Primary");

      if (!groups[orientation]) groups[orientation] = [];
      groups[orientation].push({ section, headerBody, footerBody });

      headerBody.paragraphs.load("items");
      footerBody.paragraphs.load("items");
    }

    await context.sync();

    const applyFormattingToHeader = (section, headerBody) => {
      const paras = headerBody.paragraphs;
      for (let i = 0; i < paras.items.length; i++) {
        const p = paras.items[i];

        if (header && header.alignment) p.alignment = header.alignment;
        if (formatting && formatting.allowedFont) p.font.name = formatting.allowedFont;

        if (formatting && Array.isArray(formatting.allowedFontSizeRange)) {
          const [minSize, maxSize] = formatting.allowedFontSizeRange;
          let target = p.font.size || maxSize || 12;
          if (target < minSize) target = minSize;
          if (target > maxSize) target = maxSize;
          p.font.size = target;
        }

        if (
          formatting &&
          Array.isArray(formatting.allowedFontColors) &&
          formatting.allowedFontColors.length > 0
        ) {
          p.font.color = formatting.allowedFontColors[0];
        }

        if (header && header.secondLineUnderline && i === 1) {
          p.font.underline = "Single";
        }

        if (header && header.pageNumberPreTab) {
          const original = p.text || "";
          const fixed = original.replace(/\s*(\d+)\s*$/, "\t$1");
          p.insertText(fixed, "Replace");
        }
      }

      try {
        section.pageSetup.headerDistance = 36;
      } catch (e) {}
    };

    const applyFormattingToFooter = (section, footerBody) => {
      const paras = footerBody.paragraphs;
      for (let i = 0; i < paras.items.length; i++) {
        const p = paras.items[i];

        if (footer && footer.imageMustBeCentered) {
          p.alignment = "Centered";
        }

        if (formatting && formatting.allowedFont) {
          p.font.name = formatting.allowedFont;
        }

        if (formatting && Array.isArray(formatting.allowedFontSizeRange)) {
          const [minSize, maxSize] = formatting.allowedFontSizeRange;
          let target = p.font.size || maxSize || 12;
          if (target < minSize) target = minSize;
          if (target > maxSize) target = maxSize;
          p.font.size = target;
        }

        if (
          formatting &&
          Array.isArray(formatting.allowedFontColors) &&
          formatting.allowedFontColors.length > 0
        ) {
          p.font.color = formatting.allowedFontColors[0];
        }
      }

      try {
        section.pageSetup.footerDistance = 36;
      } catch (e) {}
    };

    for (const orientation of Object.keys(groups)) {
      const group = groups[orientation];

      for (const { section, headerBody, footerBody } of group) {
        applyFormattingToHeader(section, headerBody);
        applyFormattingToFooter(section, footerBody);
      }

      await context.sync();

      const headerRanges = group.map((g) => g.headerBody.getRange());
      const footerRanges = group.map((g) => g.footerBody.getRange());

      headerRanges.forEach((r) => r.load("text"));
      footerRanges.forEach((r) => r.load("text"));
      await context.sync();

      let canonicalHeader = "";
      let canonicalFooter = "";

      for (let i = 0; i < headerRanges.length; i++) {
        const hText = (headerRanges[i].text || "").trim();
        const fText = (footerRanges[i].text || "").trim();

        if (!canonicalHeader && hText) canonicalHeader = hText;
        if (!canonicalFooter && fText) canonicalFooter = fText;
        if (canonicalHeader && canonicalFooter) break;
      }

      for (const { section, headerBody, footerBody } of group) {
        const hRange = headerBody.getRange();
        const fRange = footerBody.getRange();

        if (canonicalHeader) {
          hRange.clear();
          hRange.insertText(canonicalHeader, "Start");
        }

        if (canonicalFooter) {
          fRange.clear();
          fRange.insertText(canonicalFooter, "Start");
        }

        try {
          section.pageSetup.headerDistance = 36;
          section.pageSetup.footerDistance = 36;
        } catch (e) {}
      }

      await context.sync();
    }
  });
}

// Final pass: for each orientation (portrait / landscape), find the most
// common header + footer pattern and copy those (text + images + PAGE fields)
// to sections that don't match, without mixing portrait and landscape.
export async function syncHeaderFooterByOrientation() {
  return Word.run(async (context) => {
    const doc = context.document;
    const sections = doc.sections;

    sections.load("items/pageSetup");
    await context.sync();

    if (!sections.items.length) return;

    // Separate sections by orientation
    const portrait = [];   // { section, headerBody, footerBody, index }
    const landscape = [];

    for (let i = 0; i < sections.items.length; i++) {
      const s = sections.items[i];
      const entry = {
        section: s,
        headerBody: s.getHeader("Primary"),
        footerBody: s.getFooter("Primary"),
        index: i, // 0-based section index
      };
      if (s.pageSetup.orientation === "Landscape") {
        landscape.push(entry);
      } else {
        portrait.push(entry);
      }
    }

    // Helper: normalize text so we can compare headers/footers ignoring page numbers
    const normalizeText = (t) =>
      (t || "")
        .replace(/\d+/g, "")       // strip digits (page numbers)
        .replace(/\s+/g, " ")      // collapse whitespace
        .trim()
        .toLowerCase();

    async function syncGroup(group, label) {
      if (!group.length) return;

      // Load text + OOXML for each header/footer in the group
      const headerRanges = group.map((g) => g.headerBody.getRange());
      const footerRanges = group.map((g) => g.footerBody.getRange());

      // Load plain text for comparison
      headerRanges.forEach((r) => r.load("text"));
      footerRanges.forEach((r) => r.load("text"));

      // Request OOXML (this carries fields, images, formatting, etc.)
      const headerOOXML = headerRanges.map((r) => r.getOoxml());
      const footerOOXML = footerRanges.map((r) => r.getOoxml());

      await context.sync();

      // Build frequency maps of normalized header/footer text
      const headerFreq = new Map(); // key -> { count, index }
      const footerFreq = new Map();

      for (let i = 0; i < group.length; i++) {
        const hNorm = normalizeText(headerRanges[i].text);
        const fNorm = normalizeText(footerRanges[i].text);

        if (hNorm) {
          const prev = headerFreq.get(hNorm);
          if (prev) {
            prev.count++;
          } else {
            headerFreq.set(hNorm, { count: 1, index: i });
          }
        }

        if (fNorm) {
          const prev = footerFreq.get(fNorm);
          if (prev) {
            prev.count++;
          } else {
            footerFreq.set(fNorm, { count: 1, index: i });
          }
        }
      }

      // Pick canonical header/footer = most frequent pattern (majority wins)
      let canonicalHeaderNorm = "";
      let canonicalFooterNorm = "";
      let canonicalHeaderXml = "";
      let canonicalFooterXml = "";
      let canonicalHeaderSection = -1;
      let canonicalFooterSection = -1;

      // Header
      let maxHeaderCount = 0;
      for (const [norm, info] of headerFreq.entries()) {
        if (info.count > maxHeaderCount) {
          maxHeaderCount = info.count;
          canonicalHeaderNorm = norm;
          canonicalHeaderXml = headerOOXML[info.index].value || "";
          canonicalHeaderSection = group[info.index].index;
        }
      }

      // Footer
      let maxFooterCount = 0;
      for (const [norm, info] of footerFreq.entries()) {
        if (info.count > maxFooterCount) {
          maxFooterCount = info.count;
          canonicalFooterNorm = norm;
          canonicalFooterXml = footerOOXML[info.index].value || "";
          canonicalFooterSection = group[info.index].index;
        }
      }

      // Debug logs
      if (canonicalHeaderXml) {
        console.log(
          `[Header/Footer Sync] ${label} canonical HEADER (majority pattern) from Section ${
            canonicalHeaderSection + 1
          }`
        );
      } else {
        console.log(
          `[Header/Footer Sync] ${label} canonical HEADER (majority pattern): none found`
        );
      }

      if (canonicalFooterXml) {
        console.log(
          `[Header/Footer Sync] ${label} canonical FOOTER (majority pattern) from Section ${
            canonicalFooterSection + 1
          }`
        );
      } else {
        console.log(
          `[Header/Footer Sync] ${label} canonical FOOTER (majority pattern): none found`
        );
      }

      // If there's nothing to copy, bail out
      if (!canonicalHeaderXml && !canonicalFooterXml) return;

      // Now copy canonicals only where the normalized pattern differs
      for (let i = 0; i < group.length; i++) {
        const { section, headerBody, footerBody } = group[i];

        const currentHeaderNorm = normalizeText(headerRanges[i].text);
        const currentFooterNorm = normalizeText(footerRanges[i].text);

        if (canonicalHeaderXml && currentHeaderNorm && currentHeaderNorm !== canonicalHeaderNorm) {
          const hr = headerBody.getRange();
          hr.clear();
          hr.insertOoxml(canonicalHeaderXml, "Start");
        }

        if (canonicalFooterXml && currentFooterNorm && currentFooterNorm !== canonicalFooterNorm) {
          const fr = footerBody.getRange();
          fr.clear();
          fr.insertOoxml(canonicalFooterXml, "Start");
        }

        // Always enforce 0.5" margins
        try {
          section.pageSetup.headerDistance = 36;
          section.pageSetup.footerDistance = 36;
        } catch (e) {
          // ignore if not supported
        }
      }

      await context.sync();
    }

    // Portrait and landscape are handled completely separately
    await syncGroup(portrait, "Portrait");
    await syncGroup(landscape, "Landscape");
  });
}
