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
        // const templates = context.application.templates;
        // templates.load("items/name,items/buildingBlockEntries");
        // await context.sync();

        // let watermarkBlocks = []

        // for (const template of templates.items){
        //   const blockEntries = template.buildingBlockEntries;
        //   const count = blockEntries.getCount();
        //   await context.sync();

        //   console.log(`There are ${count.value} blocks in the template named ${template.name}`);

        //   for (let i = 0; i < count.value; i++) {
        //     const entry = blockEntries.getItemAt(i);
        //     context.load(entry, "name,type,value,description,insertType");
        //     await context.sync();

        //     if (
        //         entry.type.name === "Watermarks"
        //     ) {
        //         watermarkBlocks.push({
        //             template: template.name,
        //             type: entry.type.name,
        //             name: entry.name,
        //             value: entry.value,
        //             description: entry.description,
        //             insertType: entry.insertType
        //         });
        //     }
        //   }
         
        // }

        // console.log("Collected building block information for built-in watermarks:", watermarkBlocks);

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
            // console.log(ooxml.value);
            if (ooxml.value.includes("docPartGallery w:val=\"Watermarks\"")) {
              console.log(`Watermark detected in section ${i + 1}, header: ${type}`);
              let updatedXml = ooxml.value.replace(
                /(<v:shape[^>]*PowerPlusWaterMarkObject[^>]*\b)fillcolor="[^"]*"/i,
                '$1fillcolor="red"'
              );

              updatedXml = updatedXml.replace(
                /(<v:fill[^>]*\b)opacity="[^"]*"/gi,
                `$1opacity="1"`
              );
              header.insertOoxml(updatedXml, "Replace");

              const match = ooxml.value.match(/<v:textpath[^>]*string="([^"]+)"/i);
              
              await context.sync();

              if (match && match[1]) {
                console.log(`Watermark text: "${match[1]}"`);
              } else {
                console.log("Could not extract watermark text. May be an image watermark.");
              }
            }
          }
        }    

      }

      if(document_checking.enforceValidReferenceSources){
        console.log("loading reference fields");
        const refFields = context.document.body.fields.getByTypes([Word.FieldType.ref]);
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

          if (page_size.type == "letter" && setup.paperSize !== Word.PaperSize.letter) {
            console.warn(`Section ${index + 1} paper size is not Letter.`);
          }
          else {
            console.log(`Section ${index + 1}: ${orientation} page size is correct.`)
          }
        });
      }
      
      const sym = rules.symbols;
      if (!sym.allowSymbolFont) {
        console.log("Running symbols check...");
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load("items/font/name");
        await context.sync();

        let symbolCount = 0;
        for (const para of paragraphs.items) {
          if (para.font.name === "Symbol") {
            para.font.highlightColor = "#00CED1";  // Turquoise
            symbolCount++;
          }
        }
        await context.sync();
        console.log(`Symbols check: ${symbolCount} Symbol font instance(s) highlighted.`);
      }

    });
  } catch (error) {
    console.log("Error: " + error);
  }
}
