import rules from "../../config/rules.json";

/**
 * Runs document-level checks (comments, text boxes, watermarks, references, margins, etc.)
 * Returns an array of issue objects, compatible with the UI mapping.
 */
export async function checkDocument() {
  const { document_checking, margins, page_size } = rules;

  return Word.run(async (context) => {
    const results = [];
    let name = '';

    const comments = context.document.comments;
    const revisions = context.document.revisions;
    const shapes = context.document.body.shapes;
    const sections = context.document.sections;
    const refFields = context.document.body.fields.getByTypes([Word.FieldType.ref]);
    const paragraphs = context.document.body.paragraphs;

    if(!document_checking.allowComments){
      comments.load("items");
    }

    revisions.load("items,type,range");

    if (!document_checking.allowTextBoxes) {
      shapes.load("items/type");
    }

    if (!document_checking.allowWaterMarks || margins.enforceMargins || page_size.enforcePageSize) {
        sections.load("items, items/pageSetup");
    }

    if (document_checking.enforceValidReferenceSources) {
        refFields.load("items,result");
    }

    await context.sync();

    try {
      // --- COMMENTS ---
      if (!document_checking.allowComments) {
        // const comments = context.document.comments;
        // comments.load("items");
        // await context.sync();

        if (comments.items.length > 0) {
          comments.items.forEach((comment, i) => {
            const range = comment.getRange();
            name = `comment_${i + 1}`;
            range.insertBookmark(name);

            results.push({
              id: `comment-${i + 1}`,
              type: "Comment",
              message: `Comment found: "${comment.content}"`,
              text: comment.content,
              location: name,
              canLocate: true,
            });
          });

          // await context.sync();
        }
      }

      // --- REVISIONS ---
      // const revisions = context.document.revisions;
      // revisions.load("items,type,range");
      // await context.sync();

      revisions.items.forEach((rev, i) => {
        name = `revision_${i + 1}`;
        rev.range.insertBookmark(name);
        results.push({
          id: `revision-${i + 1}`,
          type: "Revision",
          message: `Revision detected (${rev.type}).`,
          text: rev.range.text || "",
          location: name,
          canLocate: true,
        });
      });

      // --- TEXT BOXES ---
      if (!document_checking.allowTextBoxes) {
        // const shapes = context.document.body.shapes;
        // shapes.load("items/type");
        // await context.sync();

        const textBoxes = shapes.items.filter((shape) => shape.type === "TextBox");
        textBoxes.forEach((box, i) => {
          results.push({
            id: `textbox-${i + 1}`,
            type: "Text Box",
            message: `Text box found in document.`,
            canLocate: false,
          });
        });
      }

      // --- WATERMARKS ---
      if (!document_checking.allowWaterMarks) {
        const headerTypes = ["primary", "firstPage", "evenPages"];
        const section = sections.items[0];
         

        for (const type of headerTypes) {
          const header = section.getHeader(type);
          const ooxml = header.getOoxml();
          await context.sync();

          const xml = ooxml.value

          if (xml.includes("PowerPlusWaterMarkObject") || xml.includes("WordPictureWatermark")) {
            const match = xml.match(/<v:textpath[^>]*string="([^"]+)"/i);
            const watermarkText = match ? match[1] : "(image watermark)";
            results.push({
              id: `watermark-${type}`,
              type: "Watermark",
              message: `Watermark detected in document, header type: ${type}. Watermark data: ${watermarkText}`,
              text: watermarkText,
              canLocate: false,
            });
            break
          }
        }
        
      }

      // --- INVALID REFERENCES ---
      if (document_checking.enforceValidReferenceSources) {
        // const refFields = context.document.body.fields.getByTypes([Word.FieldType.ref]);
        // refFields.load("items,result");
        // await context.sync();

        refFields.items.forEach((field, i) => {
          const text = field.result.text;
          name = `reference_${i + 1}`;
        
          if (text.includes("Error! Reference source not found")) {
            field.result.insertBookmark(name);
            results.push({
              id: `ref-${i + 1}`,
              type: "Invalid Reference",
              message: `Invalid cross-reference found: "${text}"`,
              text,
              location: name,
              canLocate: true,
            });
          }
        });

        // await context.sync();
      }

      // --- MARGINS ---
      if (margins.enforceMargins) {
        // const sections = context.document.sections;
        // sections.load("items/pageSetup");
        // await context.sync();

        sections.items.forEach((section, index) => {
          const s = section.pageSetup;
          const orientation = s.orientation;
          let isValid = true;

          if (orientation === Word.PageOrientation.portrait) {
            isValid =
              s.topMargin === margins.topMarginPortrait &&
              s.bottomMargin === margins.bottomMarginPortrait &&
              s.leftMargin === margins.leftMarginPortrait &&
              s.rightMargin === margins.rightMarginPortrait;
          } else if (orientation === Word.PageOrientation.landscape) {
            isValid =
              s.topMargin === margins.topMarginLandscape &&
              s.bottomMargin === margins.bottomMarginLandscape &&
              s.leftMargin === margins.leftMarginLandscape &&
              s.rightMargin === margins.rightMarginLandscape;
          }

          if (!isValid) {
            results.push({
              id: `margins-${index + 1}`,
              type: "Margins",
              message: `Section ${index + 1} (${orientation}) has incorrect margins. Correct ${orientation} margins: top=${orientation === Word.PageOrientation.portrait ? margins.topMarginPortrait/72 : margins.topMarginLandscape/72}, bottom=${orientation === Word.PageOrientation.portrait ? margins.bottomMarginPortrait/72 : margins.bottomMarginLandscape/72}, left=${orientation === Word.PageOrientation.portrait ? margins.leftMarginPortrait/72 : margins.leftMarginLandscape/72}, right=${orientation === Word.PageOrientation.portrait ? margins.rightMarginPortrait/72 : margins.rightMarginLandscape/72} in inches`,
              canLocate: false,
            });
          }
        });
      }

      // --- PAGE SIZE ---
      if (page_size.enforcePageSize) {
        // const sections = context.document.sections;
        // sections.load("items/pageSetup");
        // await context.sync();

        sections.items.forEach((section, i) => {
          const setup = section.pageSetup;
          if (page_size.type === "letter" && setup.paperSize !== Word.PaperSize.letter) {
            results.push({
              id: `pagesize-${i + 1}`,
              type: "Page Size",
              message: `Section ${i + 1} page size is not Letter.`,
              canLocate: false,
            });
          }
        });
      }

      // --- SUMMARY ---
      if (results.length === 0) {
        results.push({
          id: "info-clean",
          type: "Info",
          message: "✅ No document-level issues found.",
          text: "",
          canLocate: false,
        });
      }

      await context.sync();
      return results;
    } catch (err) {
      console.error("Error in checkDocument:", err);
      return [
        {
          id: "error",
          type: "Error",
          message: `An error occurred: ${err.message}`,
          canLocate: false,
        },
      ];
    }
  });
}