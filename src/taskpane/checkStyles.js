import rules from "./config/rules.json";

export async function checkStyles(){
    const { styles } = rules;
    return Word.run(async (context)=>{
        const results = [];
        let name = '';
        

        const paras = context.document.body.paragraphs;
        paras.load("items/style,items/text");
        await context.sync();

        try {
            if(!styles.allowInconsistentCaptions){
                let tabCount = 0;
                let colonCount = 0;
                let noDelimCount = 0;
                let captionCount = 0;

                paras.items.forEach((p, i) => {
                    if(p.style !== "Caption")
                        return;

                    const m = p.text.match(/^\s*([A-Za-z]+)\s+(\d+)[ ]*([\t:])?\s*(.*)$/);

                    if (m) {
                        captionCount++;

                        const delim = m[3]; // may be "\t", ":", or undefined

                        if (delim === "\t") tabCount++;
                        else if (delim === ":") colonCount++;
                        else noDelimCount++;
                    }
                });
            
                if((tabCount > 0 && colonCount > 0) || (noDelimCount > 0 && (tabCount > 0 || colonCount > 0))){
                    results.push({
                        id: `caption_delims`,
                        type: "Caption",
                        message: `Inconsistent tab and colon delimiters across captions.`,
                        canLocate: false
                    })
                }
            }


            if(!styles.allowBlankParaBkmarkStyles){
                const bookmarkStyles = ["Heading 1", "Heading 2", "Heading 3", "Heading 4", "Heading 5", "Heading 6", "Heading 7", "Heading 8", "Heading 9", "Caption"]; // add any other relevant styles

                paras.items.forEach((p, i) => {
                    const text = p.text || "";
                    const style = p.style || "";

                    // Check if paragraph has a bookmark style
                    const isBookmarkStyle = bookmarkStyles.includes(style);

                    // Check if paragraph is blank
                    const isBlank = isBlankParagraph(text);

                    if (isBookmarkStyle && isBlank) {
                        name = `parabkmarkstyle_${i+1}`;
                        p.getRange().insertBookmark(name);
                        results.push({
                            id: `parabkmarkstyle_${i+1}`,
                            type: "Blank Paragraph Mark",
                            style,
                            text,
                            message: `Blank paragraph with bookmark style '${style}' at paragraph ${i+1}`,
                            location: name,
                            canLocate: true, // you can select this range if needed
                        });
                    }
                });
            }

            if (!styles.allowBreakBkmarkStyles){
                const bookmarkStyles = ["Heading 1", "Heading 2", "Heading 3", "Heading 4", "Heading 5", "Heading 6", "Heading 7", "Heading 8", "Heading 9", "Caption"];
                const pages = context.document.body.getRange().pages;
                pages.load("items/breaks/items/range/style,index");
                await context.sync();

                pages.items.forEach((page, pageIdx) => {
                    page.breaks.items.forEach((br, brIdx) => {
                        const brStyle = br.range.style;
                        if (bookmarkStyles.includes(brStyle)) {
                            name = `break_bkmark_style${pageIdx + 1}_${brIdx + 1}`;
                            br.range.insertBookmark(name);
                            results.push({
                                id: name,
                                type: "Section/Page Breaks",
                                style: brStyle,
                                message: `Break styled with bookmark style '${brStyle} on page ${page.index}'`,
                                location: name,
                                canLocate: true,
                            });
                        }
                    });
                });
            }

            if (styles.requireConsecutiveNumbering) {
                let captions = { Figure: [], Table: [] };
                let appendixHeadings = [];

                paras.items.forEach(p => {
                    const text = p.text;

                    // --- Check captions ---
                    if (p.style === "Caption") {
                        // Match Figure/Table captions: e.g., "Figure 1: Something" or "Table 3\tSomething"
                        const m = text.match(/^([A-Za-z]+)\s+(\d+)/);
                        if (m) {
                            const label = m[1];      // "Figure" or "Table"
                            const num = parseInt(m[2], 10); // Convert caption number to integer

                            if (captions[label]) captions[label].push(num);
                        }
                    }

                    // --- Check Appendix headings ---
                    // Assuming style is Heading 1 for appendices, text like "Appendix A – Title"
                    if (p.style === "Heading 1") {
                        const mApp = text.match(/^Appendix\s+([A-Z])/i);
                        if (mApp) {
                            appendixHeadings.push(mApp[1].toUpperCase()); // Store letter
                        }
                    }
                });

                // --- Check consecutive numbering for Figure/Table captions ---
                Object.entries(captions).forEach(([label, nums]) => {
                    nums.sort((a, b) => a - b); // Ensure ascending order
                    for (let i = 0; i < nums.length; i++) {
                        if (nums[i] !== i + 1) {
                            results.push({
                                id: `nonconsecutive_${label.toLowerCase()}`,
                                type: "Caption",
                                message: `${label} captions are not consecutively numbered.`,
                                canLocate: false
                            });
                            break; // Only report once per label type
                        }
                    }
                });

                // --- Check consecutive lettering for Appendix headings ---
                const letterToNumber = (letter) => letter.charCodeAt(0) - "A".charCodeAt(0) + 1;
                for (let i = 0; i < appendixHeadings.length; i++) {
                    if (letterToNumber(appendixHeadings[i]) !== i + 1) {
                        results.push({
                            id: `nonconsecutive_appendix`,
                            type: "Heading",
                            message: `Appendix headings are not consecutively lettered.`,
                            canLocate: false
                        });
                        break; // Only report once
                    }
                }
            }


            if (results.length === 0) {
                results.push({
                    id: "info-clean",
                    type: "Info",
                    message: "✅ No style-level issues found.",
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

// Utility function to detect blank paragraphs
function isBlankParagraph(text) {
    if (!text) return true;
    // Remove Word paragraph marks (\r) and trim spaces/tabs
    const cleaned = text.replace(/\r/g, "").trim();
    return cleaned.length === 0;
}