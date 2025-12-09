# CS491 Format and Style Checker

A comprehensive Microsoft Word add-in that automatically detects and corrects formatting inconsistencies, style violations, and document structure issues in professional documents.

## 📋 Overview

This project helps maintain consistent document formatting standards by automatically scanning Word documents for common formatting errors and style violations. The tool identifies issues across multiple categories including formatting, general document structure, headers/footers, margins, styles, and symbols.

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
- **Draft Text Detection** - Identifies "Draft" text in headers
- **Information Consistency** - Checks header/footer information across all pages
- **Margin Compliance** - Validates 0.5" header/footer margin requirements

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

### 🔧 Auto-Fix Functionality

- **Automatic Corrections** - Toggle to enable automatic fixing of issues
- **Manual Fix Button** - On-demand "Fix Issues Now" button for each checker
- **Fix Results Display** - Shows count of successfully fixed and unfixed issues
- **Supported Fixes**:
  - Font corrections (Times New Roman)
  - Font size adjustments (8-12pt range)
  - Font color normalization (black)
  - Highlighting removal
  - Hidden text removal
  - Text justification (justified alignment)
  - Hyperlink removal
  - Comment deletion
  - Revision acceptance
  - Blank paragraph mark removal

### 🎨 Visual Features

- **Loading Indicators** - Animated spinner with status messages during checks
- **Color-Coded Categories** - Quick visual identification of issue types:
  - 🔵 **Light Blue** - Formatting issues
  - 🟠 **Light Orange** - General document issues
  - 🟣 **Light Purple** - Header/Footer issues
  - 🔴 **Light Pink** - Margin issues
  - 🟢 **Light Green** - Style issues
  - 🟡 **Light Yellow** - Symbol issues
- **Color Legend** - Reference guide displayed at the top of the Comprehensive Checker

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
