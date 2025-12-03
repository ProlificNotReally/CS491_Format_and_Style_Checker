# CS491_Format_and_Style_Checker
A comprehensive Microsoft Word add-in that automatically detects and corrects formatting inconsistencies, style violations, and document structure issues in professional documents.
Overview
This project is designed to help maintain consistent document formatting standards by automatically scanning Word documents for common formatting errors and style violations. The tool identifies issues across multiple categories including formatting, general document structure, headers/footers, margins, styles, and symbols.
Features
🎨 Formatting Check

Text Highlighting Detection - Identifies inappropriate use of text highlighting
Hidden Text Detection - Locates hidden text content that may cause issues
Font Color Validation - Flags non-standard font colors (accepts only black or blue)
Blank Page Detection - Identifies unnecessary blank pages
Font Size Compliance - Ensures text falls within 8-12pt range
Font Family Enforcement - Validates Times New Roman usage throughout document
Text Justification Consistency - Checks for uniform text alignment
Orphaned Elements Detection - Identifies headings/captions separated from content
Table Header Validation - Ensures continuous tables maintain repeating headers
Hyperlink Verification - Confirms websites are properly hyperlinked

📄 General Document Check

Field Code Error Detection - Identifies broken reference source field codes
Comment & Track Changes Detection - Locates remaining comments or tracked changes
Text Box Identification - Flags use of text boxes in documents
Watermark Detection - Identifies document watermarks

📋 Headers and Footers

Style Application Validation - Ensures proper header/footer styles are applied
Landscape Orientation Support - Validates landscape-specific header/footer styles
Draft Text Detection - Identifies "Draft" text in headers
Information Consistency - Checks header/footer information across all pages
Margin Compliance - Validates 0.5" header/footer margin requirements

📏 Margins Check

Paper Size Validation - Ensures 8.5" x 11" Letter size compliance
Margin Standards Enforcement

Portrait: Left 1.25", Right 1.0", Top 1.0", Bottom 1.0"
Landscape: Left 1.0", Right 1.0", Top 1.25", Bottom 1.0"



🎯 Styles Check

Caption Consistency - Validates consistent use of tabs and colons in captions
Bookmark Style Validation - Identifies misused bookmark styles on blank paragraphs
Page Break Style Check - Ensures proper styling of page/section breaks
Sequential Numbering - Validates consecutive numbering in headings and captions

Installation

Clone the repository
Install dependencies: npm install
Build the project: npm run build
Load the add-in into Microsoft Word via Developer mode

Usage

Open a Microsoft Word document
Launch the Document Format & Style Checker add-in
Run comprehensive scan or select specific check categories
Review identified issues in the results panel
Apply suggested corrections automatically or manually

🔣 Symbols Check

Symbol Font Validation - Ensures symbols use appropriate fonts (not Symbol font)
