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


      // Check if "DRAFT" exists in any line of the header
      let hasDraft = false;
      for (let j = 0; j < nonEmptyHeaderParas.length; j++) {
        const para = nonEmptyHeaderParas[j];
        const text = para.text || "";
        if (text.toLowerCase().includes("draft")) {
          hasDraft = true;
          break;
        }
      }

      // Flag if DRAFT is missing from header
      if (!hasDraft && headerCount >= 1) {
        const range = nonEmptyHeaderParas[0].getRange();
        const bk = `header_draft_${i + 1}`;
        range.insertBookmark(bk);
        await context.sync();

        results.push({
          id: `header-draft-${i + 1}`,
          type: "Header",
          message: `Section ${i + 1}: "DRAFT" is missing from header.`,
          location: bk,
          canLocate: true
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

// Fix individual header/footer issue
export async function fixHeaderFooterIssue(issue) {
  return Word.run(async (context) => {
    const { formatting, header, footer } = rules;
    const sections = context.document.sections;
    sections.load("items");
    await context.sync();

    // Extract section index from issue ID
    // ID formats: header-margin-1, header-align-0-1, footer-margin-1, etc.
    let sectionIndex = 0;
    
    if (issue.id.includes("header-margin") || issue.id.includes("footer-margin")) {
      const match = issue.id.match(/margin-(\d+)/);
      sectionIndex = match ? parseInt(match[1]) - 1 : 0;
    } else if (issue.id.includes("header-") || issue.id.includes("footer-")) {
      const match = issue.id.match(/-(\d+)-/);
      sectionIndex = match ? parseInt(match[1]) : 0;
    }
    
    if (sectionIndex < 0 || sectionIndex >= sections.items.length) {
      return { success: false, message: `Invalid section index: ${sectionIndex}` };
    }

    const section = sections.items[sectionIndex];

    try {
      // Fix header margin
      if (issue.id.includes("header-margin")) {
        section.pageSetup.load("headerDistance");
        await context.sync();
        section.pageSetup.headerDistance = 36; // 0.5 inches
        await context.sync();
        return { success: true, message: "Header margin fixed to 0.5 inches" };
      }

      // Fix footer margin
      if (issue.id.includes("footer-margin")) {
        section.pageSetup.load("footerDistance");
        await context.sync();
        section.pageSetup.footerDistance = 36; // 0.5 inches
        await context.sync();
        return { success: true, message: "Footer margin fixed to 0.5 inches" };
      }

      // Fix header alignment
      if (issue.id.includes("header-align")) {
        const headerObj = section.getHeader("Primary");
        const headerParas = headerObj.paragraphs;
        headerParas.load("items");
        await context.sync();
        
        const paraMatch = issue.id.match(/-(\d+)$/);
        const paraIndex = paraMatch ? parseInt(paraMatch[1]) : 0;
        if (headerParas.items[paraIndex]) {
          headerParas.items[paraIndex].alignment = header.alignment;
          await context.sync();
          return { success: true, message: "Header alignment fixed" };
        }
      }

      // Fix header underline
      if (issue.id.includes("header-underline")) {
        const headerObj = section.getHeader("Primary");
        const headerParas = headerObj.paragraphs;
        headerParas.load("items");
        await context.sync();
        
        if (headerParas.items[1]) {
          headerParas.items[1].font.underline = "Single";
          await context.sync();
          return { success: true, message: "Header underline applied" };
        }
      }

      // Fix header font
      if (issue.id.includes("header-font") && !issue.id.includes("size") && !issue.id.includes("color")) {
        const headerObj = section.getHeader("Primary");
        const headerParas = headerObj.paragraphs;
        headerParas.load("items");
        await context.sync();
        
        const paraMatch = issue.id.match(/-(\d+)$/);
        const paraIndex = paraMatch ? parseInt(paraMatch[1]) : 0;
        if (headerParas.items[paraIndex]) {
          headerParas.items[paraIndex].font.name = formatting.allowedFont;
          await context.sync();
          return { success: true, message: "Header font fixed" };
        }
      }

      // Fix header font size
      if (issue.id.includes("header-size") || issue.id.includes("headerfontsize")) {
        const headerObj = section.getHeader("Primary");
        const headerParas = headerObj.paragraphs;
        headerParas.load("items");
        await context.sync();
        
        const paraMatch = issue.id.match(/-(\d+)$/);
        const paraIndex = paraMatch ? parseInt(paraMatch[1]) : 0;
        if (headerParas.items[paraIndex]) {
          headerParas.items[paraIndex].font.size = formatting.allowedFontSizeRange[0];
          await context.sync();
          return { success: true, message: "Header font size fixed" };
        }
      }

      // Fix header font color
      if (issue.id.includes("header-color") || issue.id.includes("headerfontcolor")) {
        const headerObj = section.getHeader("Primary");
        const headerParas = headerObj.paragraphs;
        headerParas.load("items");
        await context.sync();
        
        const paraMatch = issue.id.match(/-(\d+)$/);
        const paraIndex = paraMatch ? parseInt(paraMatch[1]) : 0;
        if (headerParas.items[paraIndex]) {
          headerParas.items[paraIndex].font.color = "#000000";
          await context.sync();
          return { success: true, message: "Header font color fixed" };
        }
      }

      // Fix footer image alignment
      if (issue.id.includes("footer-image") || issue.id.includes("footeralignment")) {
        const footerObj = section.getFooter("Primary");
        const footerParas = footerObj.paragraphs;
        footerParas.load("items");
        await context.sync();
        
        const paraMatch = issue.id.match(/-(\d+)$/);
        const paraIndex = paraMatch ? parseInt(paraMatch[1]) : 0;
        if (footerParas.items[paraIndex]) {
          footerParas.items[paraIndex].alignment = "Centered";
          await context.sync();
          return { success: true, message: "Footer image centered" };
        }
      }

      return { success: false, message: `Fix not implemented for issue type: ${issue.id}` };
    } catch (err) {
      console.error("Error in fixHeaderFooterIssue:", err);
      return { success: false, message: `Error: ${err.message}` };
    }
  });
}

// Fix all header/footer issues at once
export async function fixAllHeaderFooterIssues(issues) {
  const results = [];
  
  for (const issue of issues) {
    // Skip info messages and non-fixable issues
    if (issue.type === "Info" || issue.id.includes("draft") || issue.id.includes("inconsistent")) {
      results.push({ issue: issue.id, success: false, message: "Not auto-fixable" });
      continue;
    }

    try {
      const result = await fixHeaderFooterIssue(issue);
      results.push({ issue: issue.id, ...result });
    } catch (err) {
      results.push({ issue: issue.id, success: false, message: err.message });
    }
  }

  return results;
}
