
import rules from "./config/rules.json";

/**
 * Header/Footer Formatting Checker Module with rule.json validation
 */
export async function checkHeaderFooterFormatting(context) {
  const results = [];
  const { formatting, header, footer } = rules;

  const sections = context.document.sections;
  sections.load("items");
  await context.sync();

  for (let i = 0; i < sections.items.length; i++) {
    const section = sections.items[i];
    const headerObj = section.getHeader("Primary");
    const footerObj = section.getFooter("Primary");

    const headerParas = headerObj.paragraphs;
    const footerParas = footerObj.paragraphs;

    headerParas.load("items");
    footerParas.load("items");
    await context.sync();

    const headerCount = headerParas.items.length;

    let bkmark_name = '';
    if (headerCount !== header.requiredLineCount) {
      
      let i = 0;
      headerParas.items.forEach(p => {
        bkmark_name= `headerlinecount_${i + 1}`;
        const range = p.getRange();
        range.insertBookmark(bkmark_name);

        p.font.highlightColor = "Yellow";
        results.push({ id: `header-lines-${i}`, type: "Header", message: `Section ${i + 1} header: Expected ${header.requiredLineCount} lines, found ${headerCount}`,
          location: bkmark_name,
          canLocate: true
        });

        i++;
      });
    }

    for (let j = 0; j < Math.min(headerCount, header.requiredLineCount); j++) {
      const para = headerParas.items[j];
      const range = para.getRange();
      
      para.load(["alignment", "font/underline", "font/name", "font/size", "font/color", "text"]);
      await context.sync();

      const { name, size, color, underline } = para.font;

      if (para.alignment !== header.alignment) {
        bkmark_name = `headeralignment_${j + 1}`;
        range.insertBookmark(bkmark_name);

        para.font.highlightColor = "Yellow";
        results.push({ id: `header-align-${i}-${j}`, type: "Header", message: `Section ${i + 1} header line ${j + 1}: Not ${header.alignment}-aligned`,
          location: bkmark_name,
          canLocate: true
        });
      }

      if (j === 1 && header.secondLineUnderline && underline === "None") {
        bkmark_name = `headerunderline_${j + 1}`;
        range.insertBookmark(bkmark_name);

        para.font.highlightColor = "Yellow";
        results.push({ id: `header-underline-${i}`, type: "Header", message: `Section ${i + 1} header line 2: Not underlined`,
          location: bkmark_name,
          canLocate: true
        });
      }

      if (name !== formatting.allowedFont) {
        bkmark_name = `headerfont_${j + 1}`;
        range.insertBookmark(bkmark_name);

        para.font.highlightColor = "Yellow";
        results.push({ id: `header-font-${i}-${j}`, type: "Header", message: `Section ${i + 1} header line ${j + 1}: Font not ${formatting.allowedFont}`,
          location: bkmark_name,
          canLocate: true
        });
      }

      if (size < formatting.allowedFontSizeRange[0] || size > formatting.allowedFontSizeRange[1]) {
        bkmark_name = `headerfontsize_${j + 1}`;
        range.insertBookmark(bkmark_name);

        para.font.highlightColor = "Yellow";
        results.push({ id: `header-size-${i}-${j}`, type: "Header", message: `Section ${i + 1} header line ${j + 1}: Font size ${size} not in allowed range`,
          location: bkmark_name,
          canLocate: true
        });
      }

      const allowedColors = formatting.allowedFontColors.map(c => c.toLowerCase());
      if (!allowedColors.includes((color || "").toLowerCase())) {
        bkmark_name = `headerfontcolor_${j + 1}`;
        range.insertBookmark(bkmark_name);

        para.font.highlightColor = "Yellow";
        results.push({ id: `header-color-${i}-${j}`, type: "Header", message: `Section ${i + 1} header line ${j + 1}: Font color "${color}" not allowed`,
          location: bkmark_name,
          canLocate: true
        });
      }

      const text = para.text || "";
      if (header.pageNumberPreTab && /\d+$/.test(text)) {
        bkmark_name = `headertabs_${j + 1}`;

        const numIndex = text.search(/\d+$/);
        const before = text.slice(0, numIndex);
        if (!before.includes("\t") && /  +$/.test(before)) {
          range.insertBookmark(bkmark_name);

          para.font.highlightColor = "Yellow";
          results.push({ id: `header-tab-${i}-${j}`, type: "Header", message: `Section ${i + 1} header line ${j + 1}: Page number not preceded by TAB`,
            location: bkmark_name,
            canLocate: true
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
          bkmark_name = `footeralignment_${k + 1}`;
          range.insertBookmark(bkmark_name);

          para.font.highlightColor = "Yellow";
          results.push({
            id: `footer-image-${i}-${k}`, type: "Footer", message: `Section ${i + 1} footer line ${k + 1}: Inline image not centered`,
            location: bkmark_name,
            canLocate: true
          });
        }
      }
    }
  }

  return results;
}
