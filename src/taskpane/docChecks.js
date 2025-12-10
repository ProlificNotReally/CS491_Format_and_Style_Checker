import rules from "./config/rules.json";

/**
 * Runs document-level checks (comments, text boxes, watermarks, references, page size, etc.)
 * Returns an array of issue objects, compatible with the UI mapping.
 */
export async function checkDocument() {
  const { document_checking, page_size, symbols } = rules;

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

    if (!document_checking.allowWaterMarks || page_size.enforcePageSize) {
        sections.load("items, items/pageSetup");
    }

    if (document_checking.enforceValidReferenceSources) {
        refFields.load("items,result");
    }

    if (!symbols.allowSymbolFont) {
        //console.log("Running symbols check...");
        paragraphs.load("items/font/name");
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
              message: `Watermark detected in document, header type: ${type}. Watermark data: ${watermarkText}.\nRemove it manually: Design -> Watermark -> Remove Watermark`,
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

      // -------- SYMBOL FONT (efficient) --------
      if (!symbols.allowSymbolFont) {
        // paragraphs.items contains only font.name (we loaded that earlier)
        // collect indices that match Symbol, then load text only for those specific paragraphs:
        const symbolParagraphIndexes = [];
        for (let i = 0; i < paragraphs.items.length; i++) {
          const p = paragraphs.items[i];
          if (p.font && p.font.name && p.font.name.toLowerCase() === "symbol") {
            symbolParagraphIndexes.push(i);
          }
        }

        // If matches exist, fetch their text ranges in one batch:
        const rangesToLoad = [];
        for (const idx of symbolParagraphIndexes) {
          const para = paragraphs.items[idx];
          const range = para.getRange();
          range.load("text");
          rangesToLoad.push(range);
        }

        if (rangesToLoad.length > 0) {
          await context.sync(); 
          for (let j = 0; j < rangesToLoad.length; j++) {
            const range = rangesToLoad[j];
            name = `symbolfont_${j + 1}`;
            range.insertBookmark(name);

            results.push({
              id: `symbol-${j + 1}`,
              type: "Symbols",
              message: `Invalid use of Symbol font in a paragraph.`,
              text: range.text,
              location: name,
              canLocate: true,
            });
          }
        }
      }

      // // --- SUMMARY ---
      // if (results.length === 0) {
      //   results.push({
      //     id: "info-clean",
      //     type: "Info",
      //     message: "✅ No document-level issues found.",
      //     text: "",
      //     canLocate: false,
      //   });
      // }

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