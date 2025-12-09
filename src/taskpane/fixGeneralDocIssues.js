import rules from "./config/rules.json";

// export async function fixGeneralDocumentIssue(issue) {
//     return Word.run(async context => {
//         if (!issue || !issue.location || !issue.id) return;


//     });
// }

export async function fixGeneralDocumentIssue(issue) {
    return Word.run(async (context) => {
        if (!issue || !issue.id || !issue.type) return;

        const { document_checking, margins, page_size, symbols } = rules;

        try {
            if (issue.type === "Comment") {
                // Comments cannot be located by bookmark, so remove all comments.
                const comments = context.document.comments;
                comments.load("items");
                await context.sync();

                comments.items.forEach(c => c.delete());
                await context.sync();
                return;
            }

            if (issue.type === "Revision") {
                const revisions = context.document.revisions;
                revisions.acceptAll();
                await context.sync();
                return;
            }

            if (issue.type === "Text Box") {
                // Word API: Shapes cannot be located individually through bookmark. SO, we will remove all textboxes
                const shapes = context.document.body.shapes;
                shapes.load("items/type");
                await context.sync();

                shapes.items
                    .filter(s => s.type === "TextBox")
                    .forEach(s => s.delete());

                await context.sync();
                return;
            }

            if (issue.type === "Invalid Reference" && issue.location) {
                const bookmarkRange = context.document.getBookmarkRange(issue.location);
                bookmarkRange.load("text");
                await context.sync();

                bookmarkRange.insertText(
                    "[Reference removed - source not found]",
                    Word.InsertLocation.replace
                );

                await context.sync();
                return;
            }

            if (issue.type === "Margins") {
                const sections = context.document.sections;
                sections.load("items/pageSetup");
                await context.sync();

                sections.items.forEach(sec => {
                    const ps = sec.pageSetup;
                    const orientation = ps.orientation;

                    if (orientation === Word.PageOrientation.portrait) {
                        ps.topMargin = margins.topMarginPortrait;
                        ps.bottomMargin = margins.bottomMarginPortrait;
                        ps.leftMargin = margins.leftMarginPortrait;
                        ps.rightMargin = margins.rightMarginPortrait;
                    } else {
                        ps.topMargin = margins.topMarginLandscape;
                        ps.bottomMargin = margins.bottomMarginLandscape;
                        ps.leftMargin = margins.leftMarginLandscape;
                        ps.rightMargin = margins.rightMarginLandscape;
                    }
                });

                await context.sync();
                return;
            }

            if (issue.type === "Page Size") {
                const sections = context.document.sections;
                sections.load("items/pageSetup");
                await context.sync();

                if (page_size.type === "letter") {
                    sections.items.forEach(sec => {
                        sec.pageSetup.paperSize = Word.PaperSize.letter;
                    });
                }

                await context.sync();
                return;
            }

            if (issue.type === "Symbols") {
                try {
                    // Locate the bookmark created earlier
                    const bookmarkRange = context.document.getBookmarkRange(issue.location);

                    if (!bookmarkRange) {
                        return { success: false, message: "Unable to locate symbol font text." };
                    }

                    // Load text + font for this range
                    bookmarkRange.load("text,font/name");
                    await context.sync();

                    const originalFont = bookmarkRange.font.name;

                    // If it is indeed Symbol font, fix it
                    if (originalFont && originalFont.toLowerCase() === "symbol") {
                        bookmarkRange.font.name = "Times New Roman";  
                    } else {
                        return {
                            success: false,
                            message: "Bookmark located but font is no longer Symbol — already fixed?"
                        };
                    }

                    await context.sync();

                    return {
                        success: true,
                        message: "Symbol font converted to Times New Roman."
                    };
                } catch (err) {
                    console.error("Symbol font fix failed:", err);
                    return {
                        success: false,
                        message: `Error fixing symbol font: ${err.message}`
                    };
                }
            }

            console.warn("No fix implemented for issue:", issue.type);
            return;

        } catch (err) {
            console.error("Error in fixGeneralDocumentIssue:", err);
            throw err;
        }
    });
}
