/* global Word console */
import rules from "./config/rules.json"
export async function insertText(text) {
  // Write text to the document.
  try {
    await Word.run(async (context) => {
      let body = context.document.body;
      body.insertParagraph(text, Word.InsertLocation.end);
      await context.sync();
      console.log("Inserted text");
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

export async function checkDocument() {
//Check the document for comments, text boxes, watermarks, and invalid reference sources based on toggled rules given by rules.json
  const { document_checking, margins, page_size } = rules;
  try{
    await Word.run(async (context) => {
      if (rules.header_footer) {
        await checkHeaderFooterFormatting(context, rules.header_footer);
      }
      if(!document_checking.allowComments){
        const comments = context.document.comments;
        console.log("Attempting to load comments");
        comments.load("items");
        console.log("Loaded comments");
        await context.sync();

        if(comments.items.length > 0){
          comments.items.forEach(comment => {
            console.log("Comment text: ", comment.content);
            const range = comment.getRange();
            const paragraph = range.paragraphs.getFirst();
            paragraph.font.highlightColor = "#FF0000";
            console.log("Got comment and highlighted paragraph containing it");
          });
          await context.sync();
          console.log(`Found and highlighted ${comments.items.length} comments.`);
        } else {
          console.log("No comments found.")
        }

        const revisions = context.document.revisions;
        revisions.load("items");
        await context.sync();

        for (const rev of revisions.items) {
          console.log(`Revision type: ${rev.type}`);
        }

        await context.sync();

        for (const rev of revisions.items) {
          console.log(`Revision type: ${rev.type}`);
          if (rev.type === "Insert") {
            rev.range.font.highlightColor = "#00FF00"; 
          } else if (rev.type === "Delete") {
            rev.range.font.highlightColor = "#FF0000"; 
          }
        }

        await context.sync();
      } 

      if(!document_checking.allowTextBoxes){
        const shapes = context.document.body.shapes;
        shapes.load("items/type");
        await context.sync();

        const textBoxes = shapes.items.filter(shape => shape.type === "TextBox");
        console.log(`Found ${textBoxes.length} text boxes.`);

        for (const box of textBoxes) {
            const body = box.body;
            body.font.highlightColor = "#0FF000"; 
        }

        await context.sync();
      }

      if(!document_checking.allowWaterMarks){
        const sections = context.document.sections;
        sections.load("items");
        await context.sync();

        let watermarkCount = 0;

        for (const [i, section] of sections.items.entries()) {
          const header = section.getHeader("Primary");
          const shapes = header.shapes;

          shapes.load("items/type,textFrame/textRange,textFrame/hasText,fill/transparency");
          await context.sync();

          shapes.items.forEach((shape) => {
            const isTextWatermark =
              shape.textFrame?.hasText &&
              shape.fill?.transparency >= 0.3 && 
              shape.textFrame.textRange?.text?.length > 0;

            const isImageWatermark = shape.type === "Picture";

            if (isTextWatermark || isImageWatermark) {
              watermarkCount++;
              console.log(
                `Watermark detected in Section ${i + 1}: ${
                  isTextWatermark ? "Text" : "Image"
                }`
              );

              if (isTextWatermark) {
                shape.textFrame.textRange.font.highlightColor = "#FFFF00";
              }
            }
          });
        }

        await context.sync();

        console.log(`Found and highlighted ${watermarkCount} watermark(s).`);
      }

      if(document_checking.enforceValidReferenceSources){
        console.log("loading reference fields");
        const refFields = context.document.body.fields.getByTypes([Word.FieldType.ref]);
        console.log("loading reference fields");
        refFields.load("items"); 
        await context.sync();
        console.log("loaded reference fields");
        

        let invalidCount = 0;

        for (const field of refFields.items) {
          field.result.load("text");
        }

        await context.sync();

        for (const field of refFields.items) {
          const text = field.result.text;
          if (text.includes("Error! Reference source not found")) {
            console.log("Invalid cross reference found:", text);
            field.result.font.highlightColor = "#FFA500"; 
            invalidCount++;
          }
        }

        await context.sync();
        console.log(`Found and highlighted ${invalidCount} invalid cross references.`);
      }
      
      
      const sections = context.document.sections;
      sections.load("items/pageSetup");
      await context.sync();

      if(margins.enforceMargins){
        
        sections.items.forEach((section, index) => {
          const setup = section.pageSetup;
          const orientation = setup.orientation;
          
          console.log(`Section ${index + 1}: ${orientation}`);

          if (orientation === Word.PageOrientation.portrait) {
            if(setup.topMargin !== margins.topMarginPortrait || setup.bottomMargin !== margins.bottomMarginPortrait ||setup.leftMargin !== margins.leftMarginPortrait || setup.rightMargin !== margins.rightMarginPortrait){
              console.warn(
                `Section ${index + 1} (${orientation}) margins are incorrect: ` +
                `Top:${setup.topMargin}, Bottom:${setup.bottomMargin}, Left:${setup.leftMargin}, Right:${setup.rightMargin}`
              );
            } else {
              console.log(`Section ${index + 1} (${orientation}) margins are correct.`);
            }
          } else if (orientation === Word.PageOrientation.landscape) {
            if(setup.topMargin !== margins.topMarginLandscape || setup.bottomMargin !== margins.bottomMarginLandscape ||setup.leftMargin !== margins.leftMarginLandscape || setup.rightMargin !== margins.rightMarginLandscape){
              console.warn(
                `Section ${index + 1} (${orientation}) margins are incorrect: ` +
                `Top:${setup.topMargin}, Bottom:${setup.bottomMargin}, Left:${setup.left}, Right:${setup.right}`
              );
            } else {
              console.log(`Section ${index + 1} (${orientation}) margins are correct.`);
            }
          }
        }); 

      }

      
      if(page_size.enforcePageSize){
        sections.items.forEach((section, index) => {
          const setup = section.pageSetup;
          const orientation = setup.orientation; 
          console.log(`Section ${index + 1}: ${orientation}`);
          if (page_size.type == "letter" && setup.paperSize !== Word.PaperSize.letter) {
            console.warn(`Section ${index + 1} paper size is not Letter.`);
          }
        });
      }

    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

async function checkHeaderFooterFormatting() {
  await Word.run(async (context) => {
    const sections = context.document.sections;
    sections.load("items");
    await context.sync();

    for (let i = 0; i < sections.items.length; i++) {
      const section = sections.items[i];
      const header = section.getHeader("Primary");
      const footer = section.getFooter("Primary");

      const headerParas = header.paragraphs;
      const footerParas = footer.paragraphs;

      headerParas.load("items");
      footerParas.load("items");
      await context.sync();

      // --- HEADER RULES ---
      const headerCount = headerParas.items.length;

      // Rule 1: Must be exactly 2 lines
      if (headerCount !== 2) {
        headerParas.items.forEach(p => p.font.highlightColor = "Yellow");
        console.log(`❌ Section ${i + 1} header: Expected 2 lines, found ${headerCount}`);
      }

      for (let j = 0; j < Math.min(headerCount, 2); j++) {
        const para = headerParas.items[j];
        para.load(["alignment", "font/underline", "text"]);
        await context.sync();

        // Rule 2: Each header line must be left-aligned
        if (para.alignment !== "Left") {
          para.font.highlightColor = "Yellow";
          console.log(`❌ Section ${i + 1} header line ${j + 1}: Not left-aligned`);
        }

        // Rule 3: Second line must be underlined
        if (j === 1 && para.font.underline === "None") {
          para.font.highlightColor = "Yellow";
          console.log(`❌ Section ${i + 1} header line 2: Not underlined`);
        }

        // Rule 4: Page number must be preceded by a TAB, not spaces
        const text = para.text || "";
        if (/\d+$/.test(text)) {
          const numIndex = text.search(/\d+$/);
          const before = text.slice(0, numIndex);
          if (!before.includes("\t") && / {2,}$/.test(before)) {
            para.font.highlightColor = "Yellow";
            console.log(`❌ Section ${i + 1} header line ${j + 1}: Page number not preceded by a TAB`);
          }
        }
      }

      // --- FOOTER RULES ---
      for (let k = 0; k < footerParas.items.length; k++) {
        const para = footerParas.items[k];
        para.load(["inlinePictures", "text"]);
        await context.sync();

        const pics = para.inlinePictures;
        pics.load("items");
        await context.sync();

        if (pics.items.length > 0) {
          const paraText = para.text || "";

          const imageCentered = para.alignment === "Centered" || paraText.includes("\t");

          if (!imageCentered) {
            para.font.highlightColor = "Yellow";
            console.log(`❌ Section ${i + 1} footer line ${k + 1}: Inline image not visually centered`);
          }
        }
      }
    }

    await context.sync();
    console.log("✅ Format check complete");
  });
}
