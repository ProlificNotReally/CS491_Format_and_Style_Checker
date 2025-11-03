// import rules from "./config/rules.json";

// export async function checkDocument() {
// //Check the document for comments, text boxes, watermarks, and invalid reference sources based on toggled rules given by rules.json
//   const { document_checking, margins, page_size } = rules;
//   try{
//     await Word.run(async (context) => {
//       if(!document_checking.allowComments){
//         const comments = context.document.comments;
//         console.log("Attempting to load comments");
//         comments.load("items");
//         console.log("Loaded comments");
//         await context.sync();

//         if(comments.items.length > 0){
//           comments.items.forEach(comment => {
//             console.log("Comment text: ", comment.content);
//             const range = comment.getRange();
//             const paragraph = range.paragraphs.getFirst();
//             paragraph.font.highlightColor = "#FF0000";
//             console.log("Got comment and highlighted paragraph containing it");
//           });
//           await context.sync();
//           console.log(`Found and highlighted ${comments.items.length} comments.`);
//         } else {
//           console.log("No comments found.")
//         }

//         const revisions = context.document.revisions;
//         revisions.load("items");
//         await context.sync();

//         for (const rev of revisions.items) {
//           console.log(`Revision type: ${rev.type}`);
//         }

//         await context.sync();

//         for (const rev of revisions.items) {
//           console.log(`Revision type: ${rev.type}`);
//           if (rev.type === "Insert") {
//             rev.range.font.highlightColor = "#00FF00"; 
//           } else if (rev.type === "Delete") {
//             rev.range.font.highlightColor = "#FF0000"; 
//           }
//         }

//         await context.sync();
//       } 

//       if(!document_checking.allowTextBoxes){
//         const shapes = context.document.body.shapes;
//         shapes.load("items/type");
//         await context.sync();

//         const textBoxes = shapes.items.filter(shape => shape.type === "TextBox");
//         console.log(`Found ${textBoxes.length} text boxes.`);

//         for (const box of textBoxes) {
//             const body = box.body;
//             body.font.highlightColor = "#0FF000"; 
//         }

//         await context.sync();
//       }

//       if(!document_checking.allowWaterMarks){
//         // const templates = context.application.templates;
//         // templates.load("items/name,items/buildingBlockEntries");
//         // await context.sync();

//         // let watermarkBlocks = []

//         // for (const template of templates.items){
//         //   const blockEntries = template.buildingBlockEntries;
//         //   const count = blockEntries.getCount();
//         //   await context.sync();

//         //   console.log(`There are ${count.value} blocks in the template named ${template.name}`);

//         //   for (let i = 0; i < count.value; i++) {
//         //     const entry = blockEntries.getItemAt(i);
//         //     context.load(entry, "name,type,value,description,insertType");
//         //     await context.sync();

//         //     if (
//         //         entry.type.name === "Watermarks"
//         //     ) {
//         //         watermarkBlocks.push({
//         //             template: template.name,
//         //             type: entry.type.name,
//         //             name: entry.name,
//         //             value: entry.value,
//         //             description: entry.description,
//         //             insertType: entry.insertType
//         //         });
//         //     }
//         //   }
         
//         // }

//         // console.log("Collected building block information for built-in watermarks:", watermarkBlocks);

//         const sections = context.document.sections;
//         sections.load("items");
//         await context.sync();

//         for (let i = 0; i < sections.items.length; i++) {
//           const section = sections.items[i];
//           const headerTypes = ["primary", "firstPage", "evenPages"];
//           for (const type of headerTypes) {
//             const header = section.getHeader(type);
//             const ooxml = header.getOoxml();
//             await context.sync();
//             // console.log(ooxml.value);
//             if (ooxml.value.includes("docPartGallery w:val=\"Watermarks\"")) {
//               console.log(`Watermark detected in section ${i + 1}, header: ${type}`);
//               let updatedXml = ooxml.value.replace(
//                 /(<v:shape[^>]*PowerPlusWaterMarkObject[^>]*\b)fillcolor="[^"]*"/i,
//                 '$1fillcolor="red"'
//               );

//               updatedXml = updatedXml.replace(
//                 /(<v:fill[^>]*\b)opacity="[^"]*"/gi,
//                 `$1opacity="1"`
//               );
//               header.insertOoxml(updatedXml, "Replace");

//               const match = ooxml.value.match(/<v:textpath[^>]*string="([^"]+)"/i);
              
//               await context.sync();

//               if (match && match[1]) {
//                 console.log(`Watermark text: "${match[1]}"`);
//               } else {
//                 console.log("Could not extract watermark text. May be an image watermark.");
//               }
//             }
//           }
//         }    

//       }

//       if(document_checking.enforceValidReferenceSources){
//         console.log("loading reference fields");
//         const refFields = context.document.body.fields.getByTypes([Word.FieldType.ref]);
//         refFields.load("items"); 
//         await context.sync();
//         console.log("loaded reference fields");
        

//         let invalidCount = 0;

//         for (const field of refFields.items) {
//           field.result.load("text");
//         }

//         await context.sync();

//         for (const field of refFields.items) {
//           const text = field.result.text;
//           if (text.includes("Error! Reference source not found")) {
//             console.log("Invalid cross reference found:", text);
//             field.result.font.highlightColor = "#FFA500"; 
//             invalidCount++;
//           }
//         }

//         await context.sync();
//         console.log(`Found and highlighted ${invalidCount} invalid cross references.`);
//       }
      
      
//       const sections = context.document.sections;
//       sections.load("items/pageSetup");
//       await context.sync();

//       if(margins.enforceMargins){
        
//         sections.items.forEach((section, index) => {
//           const setup = section.pageSetup;
//           const orientation = setup.orientation;
        
//           if (orientation === Word.PageOrientation.portrait) {
//             if(setup.topMargin !== margins.topMarginPortrait || setup.bottomMargin !== margins.bottomMarginPortrait ||setup.leftMargin !== margins.leftMarginPortrait || setup.rightMargin !== margins.rightMarginPortrait){
//               console.warn(
//                 `Section ${index + 1} (${orientation}) margins are incorrect: ` +
//                 `Top:${setup.topMargin}, Bottom:${setup.bottomMargin}, Left:${setup.leftMargin}, Right:${setup.rightMargin}`
//               );
//             } else {
//               console.log(`Section ${index + 1} (${orientation}) margins are correct.`);
//             }
//           } else if (orientation === Word.PageOrientation.landscape) {
//             if(setup.topMargin !== margins.topMarginLandscape || setup.bottomMargin !== margins.bottomMarginLandscape ||setup.leftMargin !== margins.leftMarginLandscape || setup.rightMargin !== margins.rightMarginLandscape){
//               console.warn(
//                 `Section ${index + 1} (${orientation}) margins are incorrect: ` +
//                 `Top:${setup.topMargin}, Bottom:${setup.bottomMargin}, Left:${setup.left}, Right:${setup.right}`
//               );
//             } else {
//               console.log(`Section ${index + 1} (${orientation}) margins are correct.`);
//             }
//           }
//         }); 

//       }

      
//       if(page_size.enforcePageSize){
//         sections.items.forEach((section, index) => {
//           const setup = section.pageSetup;
//           const orientation = setup.orientation; 

//           if (page_size.type == "letter" && setup.paperSize !== Word.PaperSize.letter) {
//             console.warn(`Section ${index + 1} paper size is not Letter.`);
//           }
//           else {
//             console.log(`Section ${index + 1}: ${orientation} page size is correct.`)
//           }
//         });
//       }
      
      
//     });
//   } catch (error) {
//     console.log("Error: " + error);
//   }
// }


import rules from "./config/rules.json";

/**
 * Runs document-level checks (comments, text boxes, watermarks, references, margins, etc.)
 * Returns an array of issue objects, compatible with the UI mapping.
 */
export async function checkDocument() {
  const { document_checking, margins, page_size } = rules;

  return Word.run(async (context) => {
    const results = [];

    try {
      // --- COMMENTS ---
      if (!document_checking.allowComments) {
        const comments = context.document.comments;
        comments.load("items");
        await context.sync();

        if (comments.items.length > 0) {
          comments.items.forEach((comment, i) => {
            const range = comment.getRange();
            results.push({
              id: `comment-${i + 1}`,
              type: "Comment",
              message: `Comment found: "${comment.content}"`,
              text: comment.content,
              location: range,
              canLocate: true,
            });
          });
        }
      }

      // --- REVISIONS ---
      const revisions = context.document.revisions;
      revisions.load("items,type,range");
      await context.sync();

      revisions.items.forEach((rev, i) => {
        results.push({
          id: `revision-${i + 1}`,
          type: "Revision",
          message: `Revision detected (${rev.type}).`,
          text: rev.range.text || "",
          location: rev.range,
          canLocate: true,
        });
      });

      // --- TEXT BOXES ---
      if (!document_checking.allowTextBoxes) {
        const shapes = context.document.body.shapes;
        shapes.load("items/type");
        await context.sync();

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
        const sections = context.document.sections;
        sections.load("items");
        await context.sync();

        for (let i = 0; i < sections.items.length; i++) {
          const section = sections.items[i];
          const headerTypes = ["primary", "firstPage", "evenPages"];

          for (const type of headerTypes) {
            const header = section.getHeader(type);
            const ooxml = header.getOoxml();
            await context.sync();

            if (ooxml.value.includes('docPartGallery w:val="Watermarks"')) {
              const match = ooxml.value.match(/<v:textpath[^>]*string="([^"]+)"/i);
              const watermarkText = match ? match[1] : "(image watermark)";
              results.push({
                id: `watermark-${i + 1}-${type}`,
                type: "Watermark",
                message: `Watermark detected in section ${i + 1}, header: ${type}.`,
                text: watermarkText,
                canLocate: false,
              });
            }
          }
        }
      }

      // --- INVALID REFERENCES ---
      if (document_checking.enforceValidReferenceSources) {
        const refFields = context.document.body.fields.getByTypes([Word.FieldType.ref]);
        refFields.load("items,result");
        await context.sync();

        refFields.items.forEach((field, i) => {
          const text = field.result.text;
          if (text.includes("Error! Reference source not found")) {
            results.push({
              id: `ref-${i + 1}`,
              type: "Invalid Reference",
              message: `Invalid cross-reference found: "${text}"`,
              text,
              location: field.result,
              canLocate: true,
            });
          }
        });
      }

      // --- MARGINS ---
      if (margins.enforceMargins) {
        const sections = context.document.sections;
        sections.load("items/pageSetup");
        await context.sync();

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
              message: `Section ${index + 1} (${orientation}) has incorrect margins.`,
              canLocate: false,
            });
          }
        });
      }

      // --- PAGE SIZE ---
      if (page_size.enforcePageSize) {
        const sections = context.document.sections;
        sections.load("items/pageSetup");
        await context.sync();

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
