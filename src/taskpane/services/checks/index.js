/**
 * Barrel export for all check modules
 * Provides a single import point for all document checking functionality
 */

export { analyzeFormatting } from './formattingChecks';
export { checkDocument } from './documentChecks';
export { checkStyles } from './styleChecks';
export { checkHeaderFooterFormatting } from './headerFooterChecks';
export { checkSymbols } from './symbolsChecks';
