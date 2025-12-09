import rules from "../../config/rules.json";

export async function autoFixIssues(issues) {
  return Word.run(async (context) => {
    const { formatting } = rules;
    const fixedIssues = [];
    const unfixedIssues = [];

    for (const issue of issues) {
      try {
        let fixed = false;

        switch (issue.type) {
          case "Font":
            fixed = await fixFont(context, issue, formatting.allowedFont);
            break;
          case "Font Size":
            fixed = await fixFontSize(context, issue, formatting.allowedFontSizeRange);
            break;
          case "Font Color":
            fixed = await fixFontColor(context, issue);
            break;
          case "Highlighting":
            fixed = await fixHighlighting(context, issue);
            break;
          case "Hidden Text":
            fixed = await fixHiddenText(context, issue);
            break;
          case "Justification":
            fixed = await fixJustification(context, issue);
            break;
          case "Hyperlink":
            fixed = await fixHyperlink(context, issue);
            break;
          case "Comment":
            fixed = await fixComment(context, issue);
            break;
          case "Revision":
            fixed = await fixRevision(context, issue);
            break;
          case "Blank Paragraph Mark":
            fixed = await fixBlankParagraphMark(context, issue);
            break;
          case "Header":
            fixed = await fixHeader(context, issue);
            break;
          case "Footer":
            fixed = await fixFooter(context, issue);
            break;
          default:
            unfixedIssues.push({
              ...issue,
              reason: "Auto-fix not available for this issue type"
            });
            break;
        }

        if (fixed) {
          fixedIssues.push(issue);
        } else if (issue.type !== "default") {
          unfixedIssues.push({
            ...issue,
            reason: "Could not automatically fix"
          });
        }
      } catch (error) {
        console.error(`Error fixing issue ${issue.id}:`, error);
        unfixedIssues.push({
          ...issue,
          reason: error.message
        });
      }
    }

    await context.sync();

    return {
      fixed: fixedIssues,
      unfixed: unfixedIssues,
      summary: `Fixed ${fixedIssues.length} of ${issues.length} issues`
    };
  });
}

async function fixFont(context, issue, allowedFont) {
  if (!issue.location) return false;
  try {
    const bookmarks = context.document.body.bookmarks;
    bookmarks.load("items");
    await context.sync();
    const bookmark = bookmarks.items.find(b => b.name === issue.location);
    if (!bookmark) return false;
    const range = bookmark.getRange();
    range.load("font");
    await context.sync();
    range.font.name = allowedFont;
    await context.sync();
    return true;
  } catch (error) {
    console.error("Error fixing font:", error);
    return false;
  }
}

async function fixFontSize(context, issue, allowedSizeRange) {
  if (!issue.location) return false;
  try {
    const bookmarks = context.document.body.bookmarks;
    bookmarks.load("items");
    await context.sync();
    const bookmark = bookmarks.items.find(b => b.name === issue.location);
    if (!bookmark) return false;
    const range = bookmark.getRange();
    range.load("font");
    await context.sync();
    const [minSize, maxSize] = allowedSizeRange;
    const currentSize = range.font.size;
    if (currentSize === null || currentSize === "") {
      range.font.size = 12;
    } else if (currentSize < minSize) {
      range.font.size = minSize;
    } else if (currentSize > maxSize) {
      range.font.size = maxSize;
    }
    await context.sync();
    return true;
  } catch (error) {
    console.error("Error fixing font size:", error);
    return false;
  }
}

async function fixFontColor(context, issue) {
  if (!issue.location) return false;
  try {
    const bookmarks = context.document.body.bookmarks;
    bookmarks.load("items");
    await context.sync();
    const bookmark = bookmarks.items.find(b => b.name === issue.location);
    if (!bookmark) return false;
    const range = bookmark.getRange();
    range.load("font");
    await context.sync();
    range.font.color = "#000000";
    await context.sync();
    return true;
  } catch (error) {
    console.error("Error fixing font color:", error);
    return false;
  }
}

async function fixHighlighting(context, issue) {
  if (!issue.location) return false;
  try {
    const bookmarks = context.document.body.bookmarks;
    bookmarks.load("items");
    await context.sync();
    const bookmark = bookmarks.items.find(b => b.name === issue.location);
    if (!bookmark) return false;
    const range = bookmark.getRange();
    range.load("font");
    await context.sync();
    range.font.highlightColor = null;
    await context.sync();
    return true;
  } catch (error) {
    console.error("Error fixing highlighting:", error);
    return false;
  }
}

async function fixHiddenText(context, issue) {
  if (!issue.location) return false;
  try {
    const bookmarks = context.document.body.bookmarks;
    bookmarks.load("items");
    await context.sync();
    const bookmark = bookmarks.items.find(b => b.name === issue.location);
    if (!bookmark) return false;
    const range = bookmark.getRange();
    range.load("font");
    await context.sync();
    range.font.hidden = false;
    await context.sync();
    return true;
  } catch (error) {
    console.error("Error fixing hidden text:", error);
    return false;
  }
}

async function fixJustification(context, issue) {
  if (!issue.location) return false;
  try {
    const bookmarks = context.document.body.bookmarks;
    bookmarks.load("items");
    await context.sync();
    const bookmark = bookmarks.items.find(b => b.name === issue.location);
    if (!bookmark) return false;
    const range = bookmark.getRange();
    const paragraphs = range.paragraphs;
    paragraphs.load("items/isInsideTable");
    await context.sync();
    for (const para of paragraphs.items) {
      if (para.isInsideTable) {
        para.alignment = Word.Alignment.centered;
      } else {
        para.alignment = Word.Alignment.justified;
      }
    }
    await context.sync();
    return true;
  } catch (error) {
    console.error("Error fixing justification:", error);
    return false;
  }
}

async function fixHyperlink(context, issue) {
  if (!issue.location) return false;
  try {
    const bookmarks = context.document.body.bookmarks;
    bookmarks.load("items");
    await context.sync();
    const bookmark = bookmarks.items.find(b => b.name === issue.location);
    if (!bookmark) return false;
    const range = bookmark.getRange();
    const hyperlinks = range.hyperlinks;
    hyperlinks.load("items");
    await context.sync();
    for (const hyperlink of hyperlinks.items) {
      hyperlink.delete();
    }
    await context.sync();
    return true;
  } catch (error) {
    console.error("Error fixing hyperlink:", error);
    return false;
  }
}

async function fixComment(context, issue) {
  if (!issue.location) return false;
  try {
    const bookmarks = context.document.body.bookmarks;
    bookmarks.load("items");
    await context.sync();
    const bookmark = bookmarks.items.find(b => b.name === issue.location);
    if (!bookmark) return false;
    const range = bookmark.getRange();
    const comments = context.document.comments;
    comments.load("items");
    await context.sync();
    for (const comment of comments.items) {
      const commentRange = comment.getRange();
      commentRange.load("text");
      await context.sync();
      if (commentRange.text === range.text) {
        comment.delete();
      }
    }
    await context.sync();
    return true;
  } catch (error) {
    console.error("Error fixing comment:", error);
    return false;
  }
}

async function fixRevision(context, issue) {
  if (!issue.location) return false;
  try {
    const bookmarks = context.document.body.bookmarks;
    bookmarks.load("items");
    await context.sync();
    const bookmark = bookmarks.items.find(b => b.name === issue.location);
    if (!bookmark) return false;
    const range = bookmark.getRange();
    const revisions = context.document.revisions;
    revisions.load("items");
    await context.sync();
    for (const revision of revisions.items) {
      const revRange = revision.range;
      revRange.load("text");
      await context.sync();
      if (revRange.text === range.text) {
        revision.accept();
      }
    }
    await context.sync();
    return true;
  } catch (error) {
    console.error("Error fixing revision:", error);
    return false;
  }
}

async function fixBlankParagraphMark(context, issue) {
  if (!issue.location) return false;
  try {
    const bookmarks = context.document.body.bookmarks;
    bookmarks.load("items");
    await context.sync();
    const bookmark = bookmarks.items.find(b => b.name === issue.location);
    if (!bookmark) return false;
    const range = bookmark.getRange();
    const paragraphs = range.paragraphs;
    paragraphs.load("items");
    await context.sync();
    for (const para of paragraphs.items) {
      const text = para.text?.trim() || "";
      if (!text) {
        para.delete();
      }
    }
    await context.sync();
    return true;
  } catch (error) {
    console.error("Error fixing blank paragraph mark:", error);
    return false;
  }
}

async function fixHeader(context, issue) {
  if (!issue.location) return false;
  try {
    const bookmarks = context.document.body.bookmarks;
    bookmarks.load("items");
    await context.sync();
    
    const bookmark = bookmarks.items.find(b => b.name === issue.location);
    if (!bookmark) return false;
    
    const range = bookmark.getRange();
    
    // Determine what type of fix is needed based on the issue message
    if (issue.message.includes("style not applied")) {
      // Extract the expected style from the message
      const styleMatch = issue.message.match(/(Portrait Header|Landscape Header)/);
      if (styleMatch) {
        const expectedStyle = styleMatch[1];
        range.load("paragraphs");
        await context.sync();
        
        for (const para of range.paragraphs.items) {
          para.style = expectedStyle;
        }
        await context.sync();
        return true;
      }
    } else if (issue.message.includes("Font not")) {
      // Fix font
      range.load("font");
      await context.sync();
      range.font.name = rules.formatting.allowedFont;
      await context.sync();
      return true;
    } else if (issue.message.includes("Font size")) {
      // Fix font size to middle of allowed range
      range.load("font");
      await context.sync();
      const midSize = (rules.formatting.allowedFontSizeRange[0] + rules.formatting.allowedFontSizeRange[1]) / 2;
      range.font.size = midSize;
      await context.sync();
      return true;
    } else if (issue.message.includes("Font color")) {
      // Fix font color to black
      range.load("font");
      await context.sync();
      range.font.color = "#000000";
      await context.sync();
      return true;
    } else if (issue.message.includes("Not Left-aligned") || issue.message.includes("Not") && issue.message.includes("-aligned")) {
      // Fix alignment
      range.load("paragraphs");
      await context.sync();
      for (const para of range.paragraphs.items) {
        para.alignment = "Left";
      }
      await context.sync();
      return true;
    } else if (issue.message.includes("Not underlined")) {
      // Fix underline
      range.load("font");
      await context.sync();
      range.font.underline = "Single";
      await context.sync();
      return true;
    } else if (issue.message.includes("should appear in header")) {
      // Add "Draft" text at the beginning of the header line
      range.load("text");
      await context.sync();
      const text = range.text || "";
      // Only add Draft if it's not already there
      if (!text.toLowerCase().includes("draft")) {
        const newText = "Draft - " + text;
        range.insertText(newText, "Replace");
        await context.sync();
        return true;
      }
      return true; // Already has Draft
    } else if (issue.message.includes("margin")) {
      // Margin fixes would need section-level access
      return false; // Cannot auto-fix margins from range
    }
    
    return false;
  } catch (error) {
    console.error("Error fixing header:", error);
    return false;
  }
}

async function fixFooter(context, issue) {
  if (!issue.location) return false;
  try {
    const bookmarks = context.document.body.bookmarks;
    bookmarks.load("items");
    await context.sync();
    
    const bookmark = bookmarks.items.find(b => b.name === issue.location);
    if (!bookmark) return false;
    
    const range = bookmark.getRange();
    
    // Determine what type of fix is needed based on the issue message
    if (issue.message.includes("style not applied")) {
      // Extract the expected style from the message
      const styleMatch = issue.message.match(/(Portrait Footer|Landscape Footer)/);
      if (styleMatch) {
        const expectedStyle = styleMatch[1];
        range.load("paragraphs");
        await context.sync();
        
        for (const para of range.paragraphs.items) {
          para.style = expectedStyle;
        }
        await context.sync();
        return true;
      }
    } else if (issue.message.includes("not centered")) {
      // Fix centering
      range.load("paragraphs");
      await context.sync();
      for (const para of range.paragraphs.items) {
        para.alignment = "Centered";
      }
      await context.sync();
      return true;
    } else if (issue.message.includes("Confidential")) {
      // Add "Confidential" text at the beginning of the footer line
      range.load("text");
      await context.sync();
      const text = range.text || "";
      // Only add Confidential if it's not already there
      if (!text.toLowerCase().includes("confidential")) {
        const newText = "Confidential - " + text;
        range.insertText(newText, "Replace");
        await context.sync();
        return true;
      }
      return true; // Already has Confidential
    } else if (issue.message.includes("margin")) {
      // Margin fixes would need section-level access
      return false; // Cannot auto-fix margins from range
    }
    
    return false;
  } catch (error) {
    console.error("Error fixing footer:", error);
    return false;
  }
}
