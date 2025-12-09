# CS491 Format and Style Checker

A comprehensive Microsoft Word add-in that automatically detects and corrects formatting inconsistencies, style violations, and document structure issues in professional documents.

## 📋 Overview

This project helps maintain consistent document formatting standards by automatically scanning Word documents for common formatting errors and style violations. The tool identifies issues across multiple categories including formatting, general document structure, headers/footers, margins, styles, and symbols.

## 🏗️ Project Architecture

```
CS491_Format_and_Style_Checker/
├── src/                          # Main source code (Word Add-in)
│   ├── commands/                 # Office command handlers
│   │   ├── commands.html
│   │   └── commands.js
│   └── taskpane/                 # Task pane UI and logic
│       ├── components/           # React components
│       │   └── App.jsx          # Main UI component with all checkers
│       ├── config/              # Configuration files
│       │   └── rules.json       # Formatting rules and validation criteria
│       ├── checkHeaderFooterFormatting.js  # Header/footer validation & fixes
│       ├── checkStyles.js                   # Style consistency checks
│       ├── docChecks.js                     # Document structure validation
│       ├── symbolChecks.js                  # Symbol font validation & fixes
│       ├── wordChecks.js                    # Word formatting checks
│       ├── index.jsx                        # React entry point
│       ├── taskpane.html                    # Task pane HTML template
│       └── taskpane.js                      # Task pane initialization
├── assets/                       # Image assets (icons, logos)
├── dist/                         # Compiled output (generated)
├── app.py                        # Flask web interface (separate tool)
├── static/                       # Flask static files
├── templates/                    # Flask HTML templates
├── manifest.xml                  # Office Add-in manifest
├── webpack.config.js             # Build configuration
├── package.json                  # Node dependencies
└── README.md                     # This file
```

---

## ✨ Features

### 🎨 Formatting Check
- **Text Highlighting Detection** - Identifies inappropriate use of text highlighting
- **Hidden Text Detection** - Locates hidden text content that may cause issues
- **Font Color Validation** - Flags non-standard font colors (accepts only black or blue)
- **Blank Page Detection** - Identifies unnecessary blank pages
- **Font Size Compliance** - Ensures text falls within 8-12pt range
- **Font Family Enforcement** - Validates Times New Roman usage throughout document
- **Text Justification Consistency** - Checks for uniform text alignment
- **Orphaned Elements Detection** - Identifies headings/captions separated from content
- **Table Header Validation** - Ensures continuous tables maintain repeating headers
- **Hyperlink Verification** - Confirms websites are properly hyperlinked

### 📄 General Document Check
- **Field Code Error Detection** - Identifies broken reference source field codes
- **Comment & Track Changes Detection** - Locates remaining comments or tracked changes
- **Text Box Identification** - Flags use of text boxes in documents
- **Watermark Detection** - Identifies document watermarks

### 📋 Headers and Footers
- **Style Application Validation** - Ensures proper header/footer styles are applied
- **Landscape Orientation Support** - Validates landscape-specific header/footer styles
- **Draft Text Detection** - Identifies missing "DRAFT" text in headers (lines 1-2)
- **Information Consistency** - Checks header/footer information across all pages
- **Margin Compliance** - Validates 0.5" header/footer margin requirements
- **Auto-Fix Support** - Individual and bulk fixes for margins, alignment, fonts, colors, and images

### 📏 Margins Check
- **Paper Size Validation** - Ensures 8.5" x 11" Letter size compliance
- **Margin Standards Enforcement**
  - **Portrait:** Left 1.25", Right 1.0", Top 1.0", Bottom 1.0"
  - **Landscape:** Left 1.0", Right 1.0", Top 1.25", Bottom 1.0"

### 🎯 Styles Check
- **Caption Consistency** - Validates consistent use of tabs and colons in captions
- **Bookmark Style Validation** - Identifies misused bookmark styles on blank paragraphs
- **Page Break Style Check** - Ensures proper styling of page/section breaks
- **Sequential Numbering** - Validates consecutive numbering in headings and captions

### 🔣 Symbols Check
- **Symbol Font Validation** - Ensures symbols use appropriate fonts (not Symbol font)

---

## 🚀 Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/yourusername/CS491_Format_and_Style_Checker.git
   cd CS491_Format_and_Style_Checker
   ```

2. **Install dependencies**
   ```bash
   npm install
   ```

3. **Build the project**
   ```bash
   npm run build
   ```

4. **Load the add-in into Microsoft Word**
   - Enable Developer mode in Word
   - Load the add-in through the Developer tab

---

## 📖 Usage

1. **Open Microsoft Word** and load your document
2. **Launch the Document Format & Style Checker** add-in
3. **Run comprehensive scan** or select specific check categories
4. **Review identified issues** in the results panel
5. **Apply corrections** automatically or manually as needed

---

## 🛠️ Requirements

- Microsoft Word (Office 365 or 2019+)
- Node.js (v14 or higher)
- npm or yarn package manager

---

## 📝 License

This project is part of CS491 coursework and follows university guidelines for academic projects.

---

## 🤝 Contributing

This is an academic project. For questions or suggestions, please contact the development team through the course platform.

---

## 📞 Support

For technical support or feature requests, please create an issue in the project repository or contact the development team.
