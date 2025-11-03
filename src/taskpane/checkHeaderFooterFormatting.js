
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

    if (headerCount !== header.requiredLineCount) {
      headerParas.items.forEach(p => {
        p.font.highlightColor = "Yellow";
        results.push({ id: `header-lines-${i}`, type: "Header", message: `Section ${i + 1} header: Expected ${header.requiredLineCount} lines, found ${headerCount}`,
          location: p,
          canLocate: true
        });
      });
    }

    for (let j = 0; j < Math.min(headerCount, header.requiredLineCount); j++) {
      const para = headerParas.items[j];
      para.load(["alignment", "font/underline", "font/name", "font/size", "font/color", "text"]);
      await context.sync();

      const { name, size, color, underline } = para.font;

      if (para.alignment !== header.alignment) {
        para.font.highlightColor = "Yellow";
        results.push({ id: `header-align-${i}-${j}`, type: "Header", message: `Section ${i + 1} header line ${j + 1}: Not ${header.alignment}-aligned`,
          location: para,
          canLocate: true
        });
      }

      if (j === 1 && header.secondLineUnderline && underline === "None") {
        para.font.highlightColor = "Yellow";
        results.push({ id: `header-underline-${i}`, type: "Header", message: `Section ${i + 1} header line 2: Not underlined`,
          location: para,
          canLocate: true
        });
      }

      if (name !== formatting.allowedFont) {
        para.font.highlightColor = "Yellow";
        results.push({ id: `header-font-${i}-${j}`, type: "Header", message: `Section ${i + 1} header line ${j + 1}: Font not ${formatting.allowedFont}`,
          location: para,
          canLocate: true
        });
      }

      if (size < formatting.allowedFontSizeRange[0] || size > formatting.allowedFontSizeRange[1]) {
        para.font.highlightColor = "Yellow";
        results.push({ id: `header-size-${i}-${j}`, type: "Header", message: `Section ${i + 1} header line ${j + 1}: Font size ${size} not in allowed range`,
          location: para,
          canLocate: true
        });
      }

      const allowedColors = formatting.allowedFontColors.map(c => c.toLowerCase());
      if (!allowedColors.includes((color || "").toLowerCase())) {
        para.font.highlightColor = "Yellow";
        results.push({ id: `header-color-${i}-${j}`, type: "Header", message: `Section ${i + 1} header line ${j + 1}: Font color "${color}" not allowed`,
          location: para,
          canLocate: true
        });
      }

      const text = para.text || "";
      if (header.pageNumberPreTab && /\d+$/.test(text)) {
        const numIndex = text.search(/\d+$/);
        const before = text.slice(0, numIndex);
        if (!before.includes("\t") && /  +$/.test(before)) {
          para.font.highlightColor = "Yellow";
          results.push({ id: `header-tab-${i}-${j}`, type: "Header", message: `Section ${i + 1} header line ${j + 1}: Page number not preceded by TAB`,
            location: para,
            canLocate: true
          });
        }
      }
    }

    for (let k = 0; k < footerParas.items.length; k++) {
      const para = footerParas.items[k];
      para.load(["inlinePictures", "text", "alignment"]);
      await context.sync();

      const pics = para.inlinePictures;
      pics.load("items");
      await context.sync();

      if (pics.items.length > 0) {
        const text = para.text || "";
        const isCentered = para.alignment === "Centered" || text.includes("\t");

        if (footer.imageMustBeCentered && !isCentered) {
          para.font.highlightColor = "Yellow";
          results.push({
            id: `footer-image-${i}-${k}`, type: "Footer", message: `Section ${i + 1} footer line ${k + 1}: Inline image not centered`,
            location: para,
            canLocate: true
          });
        }
      }
    }
  }

  return results;
}
