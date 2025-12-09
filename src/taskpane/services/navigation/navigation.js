/**
 * Navigation utilities for Word document
 */

/**
 * Navigates to a given error location inside the Word document using bookmarks
 * @param {string} location - The bookmark name for the error location
 */
export async function goToError(location) {
  if (!location) return;

  await Word.run(async (context) => {
    const range = context.document.getBookmarkRange(location);
    range.select();
    await context.sync();    
  });
}
