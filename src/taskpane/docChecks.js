import rules from "./config/rules.json";

/**
 * Runs document-level checks (comments, text boxes, watermarks, references, margins, etc.)
 * Returns an array of issue objects, compatible with the UI mapping.
 */
export async function checkDocument() {
  const { document_checking, margins, page_size, symbols } = rules;

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


// import rules from "./config/rules.json";

// /**
//  * Fast, optimized document-level checker for Word add-ins.
//  * Key optimizations:
//  *  - All loads grouped before a single initial context.sync()
//  *  - Avoids getOoxml(); detects watermarks via header shapes
//  *  - Loads paragraph text only for Symbol-font matches
//  *  - Inserts bookmarks in-memory and does a single final sync
//  */
// export async function checkDocument() {
//   const { document_checking, margins, page_size, symbols } = rules;

//   return Word.run(async (context) => {
//     const results = [];
//     try {
//       //
//       // 1) DECLARE handles up-front (so they are always defined)
//       //
//       const comments = context.document.comments;
//       const revisions = context.document.revisions;
//       const shapes = context.document.body.shapes;
//       const sections = context.document.sections;
//       const paragraphs = context.document.body.paragraphs;
//       const refFields = context.document.body.fields.getByTypes([Word.FieldType.ref]);

//       //
//       // 2) CONDITIONALLY load minimal properties (batch loads)
//       //

//       if (!document_checking.allowComments) {
//         // only need the items; we will read comment.content later
//         comments.load("items/content");
//       }

//       // revisions are always relevant for this ruleset
//       revisions.load("items,type,range/text");

//       if (!document_checking.allowTextBoxes) {
//         // only need shape types to find TextBoxes
//         shapes.load("items/type");
//       }

//       if (!document_checking.allowWaterMarks) {
//         // we'll inspect header shapes to detect watermarks (avoid OOXML)
//         // load sections so we can get headers and their shapes later
//         sections.load("items");
//       }

//       if (document_checking.enforceValidReferenceSources) {
//         refFields.load("items,result");
//       }

//       // pageSetup is needed for margins and page size checks
//       if (margins.enforceMargins || page_size.enforcePageSize) {
//         sections.load("items/pageSetup");
//       }

//       // symbol font: load only font names (avoid loading text for every paragraph)
//       if (!symbols.allowSymbolFont) {
//         paragraphs.load("items/font/name");
//       }

//       // Single initial sync for all loads
//       await context.sync();

//       //
//       // 3) PROCESS LOADED DATA (no syncs inside loops)
//       //
//       let bookmarkCounter = 0;
//       const makeBookmarkName = () => `doc_check_loc_${++bookmarkCounter}`;

//       // -------- COMMENTS --------
//       if (!document_checking.allowComments) {
//         for (let i = 0; i < comments.items.length; i++) {
//           const comment = comments.items[i];
//           // queue bookmark insertion (no sync yet)
//           const bname = makeBookmarkName();
//           comment.getRange().insertBookmark(bname);

//           results.push({
//             id: `comment-${i + 1}`,
//             type: "Comment",
//             message: `Comment found: "${comment.content}"`,
//             text: comment.content,
//             location: bname,
//             canLocate: true,
//           });
//         }
//       }

//       // -------- REVISIONS --------
//       // revisions were loaded earlier; their ranges are available
//       for (let i = 0; i < revisions.items.length; i++) {
//         const rev = revisions.items[i];
//         const bname = makeBookmarkName();
//         rev.range.insertBookmark(bname);

//         results.push({
//           id: `revision-${i + 1}`,
//           type: "Revision",
//           message: `Revision detected (${rev.type}).`,
//           text: rev.range.text || "",
//           location: bname,
//           canLocate: true,
//         });
//       }

//       // -------- TEXT BOXES --------
//       if (!document_checking.allowTextBoxes) {
//         const textBoxes = shapes.items.filter((s) => s.type === "TextBox");
//         for (let i = 0; i < textBoxes.length; i++) {
//           results.push({
//             id: `textbox-${i + 1}`,
//             type: "Text Box",
//             message: "Text box found in document.",
//             canLocate: false,
//           });
//         }
//       }

//       // --- WATERMARKS ---
//       if (!document_checking.allowWaterMarks) {
//         // const sections = context.document.sections;
//         // sections.load("items");
//         // await context.sync();
//         const headerTypes = ["primary", "firstPage", "evenPages"];
//         const section = sections.items[0];
         

//         for (const type of headerTypes) {
//           const header = section.getHeader(type);
//           const ooxml = header.getOoxml();
//           await context.sync();

//           const xml = ooxml.value

//           if (xml.includes("PowerPlusWaterMarkObject") || xml.includes("WordPictureWatermark")) {
//             const match = xml.match(/<v:textpath[^>]*string="([^"]+)"/i);
//             const watermarkText = match ? match[1] : "(image watermark)";
//             results.push({
//               id: `watermark-${type}`,
//               type: "Watermark",
//               message: `Watermark detected in document, header type: ${type}. Watermark data: ${watermarkText}`,
//               text: watermarkText,
//               canLocate: false,
//             });
//             break
//           }
//         }
        
//       }

//       // -------- INVALID REFERENCES --------
//       if (document_checking.enforceValidReferenceSources) {
//         for (let i = 0; i < refFields.items.length; i++) {
//           const field = refFields.items[i];
//           const text = field.result.text;
//           if (text && text.includes("Error! Reference source not found")) {
//             const b = makeBookmarkName();
//             field.result.insertBookmark(b);

//             results.push({
//               id: `ref-${i + 1}`,
//               type: "Invalid Reference",
//               message: `Invalid cross-reference found: "${text}"`,
//               text,
//               location: b,
//               canLocate: true,
//             });
//           }
//         }
//       }

//       // -------- MARGINS --------
//       if (margins.enforceMargins) {
//         // We loaded sections.pageSetup earlier if needed
//         for (let i = 0; i < sections.items.length; i++) {
//           const s = sections.items[i].pageSetup;
//           const isPortrait = s.orientation === Word.PageOrientation.portrait;

//           const valid =
//             (isPortrait &&
//               s.topMargin === margins.topMarginPortrait &&
//               s.bottomMargin === margins.bottomMarginPortrait &&
//               s.leftMargin === margins.leftMarginPortrait &&
//               s.rightMargin === margins.rightMarginPortrait) ||
//             (!isPortrait &&
//               s.topMargin === margins.topMarginLandscape &&
//               s.bottomMargin === margins.bottomMarginLandscape &&
//               s.leftMargin === margins.leftMarginLandscape &&
//               s.rightMargin === margins.rightMarginLandscape);

//           if (!valid) {
//             const correct = isPortrait
//               ? {
//                   top: margins.topMarginPortrait / 72,
//                   bottom: margins.bottomMarginPortrait / 72,
//                   left: margins.leftMarginPortrait / 72,
//                   right: margins.rightMarginPortrait / 72,
//                 }
//               : {
//                   top: margins.topMarginLandscape / 72,
//                   bottom: margins.bottomMarginLandscape / 72,
//                   left: margins.leftMarginLandscape / 72,
//                   right: margins.rightMarginLandscape / 72,
//                 };

//             results.push({
//               id: `margins-${i + 1}`,
//               type: "Margins",
//               message: `Section ${i + 1} has incorrect margins. Expected (inches): top=${correct.top}", bottom=${correct.bottom}", left=${correct.left}", right=${correct.right}".`,
//               canLocate: false,
//             });
//           }
//         }
//       }

//       // -------- PAGE SIZE --------
//       if (page_size.enforcePageSize) {
//         for (let i = 0; i < sections.items.length; i++) {
//           const setup = sections.items[i].pageSetup;
//           if (page_size.type === "letter" && setup.paperSize !== Word.PaperSize.letter) {
//             results.push({
//               id: `pagesize-${i + 1}`,
//               type: "Page Size",
//               message: `Section ${i + 1} page size is not Letter.`,
//               canLocate: false,
//             });
//           }
//         }
//       }


//       // -------- SYMBOL FONT (efficient) --------
//       if (!symbols.allowSymbolFont) {
//         // paragraphs.items contains only font.name (we loaded that earlier)
//         // collect indices that match Symbol, then load text only for those specific paragraphs:
//         const symbolParagraphIndexes = [];
//         for (let i = 0; i < paragraphs.items.length; i++) {
//           const p = paragraphs.items[i];
//           if (p.font && p.font.name && p.font.name.toLowerCase() === "symbol") {
//             symbolParagraphIndexes.push(i);
//           }
//         }

//         // If matches exist, fetch their text ranges in one batch:
//         const rangesToLoad = [];
//         for (const idx of symbolParagraphIndexes) {
//           const para = paragraphs.items[idx];
//           const range = para.getRange();
//           range.load("text");
//           rangesToLoad.push(range);
//         }

//         if (rangesToLoad.length > 0) {
//           await context.sync(); // one sync to get text of only the matched paragraphs
//           // now process them and insert bookmarks (queued)
//           for (let j = 0; j < rangesToLoad.length; j++) {
//             const rng = rangesToLoad[j];
//             const b = makeBookmarkName();
//             rng.insertBookmark(b);

//             results.push({
//               id: `symbol-${j + 1}`,
//               type: "Symbols",
//               message: `Invalid use of Symbol font in a paragraph.`,
//               text: rng.text,
//               location: b,
//               canLocate: true,
//             });
//           }
//         }
//       }
      
//       // 4) FINAL SUMMARY if no issues
//       //
//       if (results.length === 0) {
//         results.push({
//           id: "info-clean",
//           type: "Info",
//           message: "✅ No document-level issues found.",
//           text: "",
//           canLocate: false,
//         });
//       }

//       // //
//       // // 5) Final sync: apply all insertBookmark() operations queued earlier
//       // //
//       // // Note: some insertBookmark calls were already queued earlier (comments, revisions, invalid refs, symbol paragraphs).
//       // // We need a final sync to apply them to the document.
//       await context.sync();

//       return results;
//     } catch (err) {
//       console.error("Error in optimized checkDocument:", err);
//       return [
//         {
//           id: "error",
//           type: "Error",
//           message: `An error occurred: ${err.message}`,
//           canLocate: false,
//         },
//       ];
//     }
//   });
// }

