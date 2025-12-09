# Architecture Documentation

## Project Structure

```
CS491_Format_and_Style_Checker/
│
├── src/
│   ├── commands/                          # Office Add-in Commands
│   │   ├── commands.html
│   │   └── commands.js
│   │
│   └── taskpane/                          # Main Application
│       ├── index.jsx                      # Entry point
│       ├── taskpane.html                  # HTML template
│       ├── taskpane.js                    # Initialization
│       │
│       ├── components/                    # React UI Components
│       │   └── App.jsx                    # Main app component with all checkers
│       │
│       ├── services/                      # Business Logic Layer
│       │   │
│       │   ├── checks/                    # Document Checking Services
│       │   │   ├── index.js              # Barrel export (all checks)
│       │   │   ├── formattingChecks.js   # Font, size, color, highlighting, etc.
│       │   │   ├── documentChecks.js     # Comments, revisions, text boxes, etc.
│       │   │   ├── styleChecks.js        # Caption styles, blank paragraphs, etc.
│       │   │   └── headerFooterChecks.js # Header/footer formatting
│       │   │
│       │   ├── fixes/                     # Auto-Fix Services
│       │   │   ├── index.js              # Barrel export
│       │   │   └── autoFix.js            # Automatic issue fixing
│       │   │
│       │   └── navigation/                # Navigation Utilities
│       │       ├── index.js              # Barrel export
│       │       └── navigation.js         # Document navigation (goToError)
│       │
│       └── config/                        # Configuration
│           └── rules.json                 # Validation rules
│
├── assets/                                # Static assets (icons, images)
├── manifest.xml                           # Office Add-in manifest
├── package.json                           # Dependencies and scripts
└── webpack.config.js                      # Build configuration
```

## Architecture Diagram

```
┌─────────────────────────────────────────────────────────────────┐
│                          USER INTERFACE                          │
│  ┌────────────────────────────────────────────────────────────┐ │
│  │                      App.jsx (React)                        │ │
│  │                                                              │ │
│  │  • Auto-fix Toggle                                          │ │
│  │  • Comprehensive Checker                                    │ │
│  │  • Formatting Checker                                       │ │
│  │  • Document Checker                                         │ │
│  │  • Styles Checker                                           │ │
│  │  • Header/Footer Checker                                    │ │
│  └────────────────────────────────────────────────────────────┘ │
└──────────────────────┬──────────────────────────────────────────┘
                       │
                       ▼
┌─────────────────────────────────────────────────────────────────┐
│                       SERVICES LAYER                             │
│                                                                  │
│  ┌──────────────────┐  ┌──────────────────┐  ┌──────────────┐ │
│  │  CHECKS          │  │  FIXES           │  │  NAVIGATION  │ │
│  │                  │  │                  │  │              │ │
│  │ • Formatting ────┼──┼──► Auto-fix ◄───┼──┼─ goToError() │ │
│  │ • Document       │  │                  │  │              │ │
│  │ • Styles         │  │  Issues → Fixed  │  │  Bookmark    │ │
│  │ • HeaderFooter   │  │                  │  │  Navigation  │ │
│  └────────┬─────────┘  └──────────────────┘  └──────────────┘ │
│           │                                                      │
└───────────┼──────────────────────────────────────────────────────┘
            │
            ▼
┌─────────────────────────────────────────────────────────────────┐
│                      CONFIGURATION                               │
│  ┌────────────────────────────────────────────────────────────┐ │
│  │                      rules.json                             │ │
│  │                                                              │ │
│  │  • Formatting rules (fonts, sizes, colors)                 │ │
│  │  • Document rules (comments, revisions)                    │ │
│  │  • Style rules (captions, paragraphs)                      │ │
│  │  • Margins, page size, symbols                             │ │
│  └────────────────────────────────────────────────────────────┘ │
└─────────────────────────────────────────────────────────────────┘
            │
            ▼
┌─────────────────────────────────────────────────────────────────┐
│                    MICROSOFT WORD API                            │
│                     (Office.js SDK)                              │
│                                                                  │
│  • Word.run() context                                           │
│  • Document.paragraphs, sections, tables                        │
│  • Font properties, styles, formatting                          │
│  • OOXML manipulation                                           │
│  • Bookmarks for navigation                                     │
└─────────────────────────────────────────────────────────────────┘
```

## Data Flow

### Check Flow
```
User clicks "Run Check"
      ↓
App.jsx calls check function
      ↓
services/checks/* analyzes document
      ↓
Returns array of issues [{type, message, location}]
      ↓
App.jsx displays issues in UI
      ↓
User clicks issue → goToError() navigates to location
```

### Auto-Fix Flow
```
User enables "Auto-fix" toggle
      ↓
User clicks "Run Check"
      ↓
services/checks/* finds issues
      ↓
services/fixes/autoFix processes fixable issues
      ↓
Returns {fixed: [], unfixed: [], summary}
      ↓
App.jsx shows only unfixed issues
```

## Module Responsibilities

### Components Layer
- **App.jsx**: UI orchestration, state management, user interactions

### Services Layer

#### checks/
- **formattingChecks.js**: Font, size, color, highlighting, hidden text, justification, hyperlinks, blank pages
- **documentChecks.js**: Comments, revisions, text boxes, watermarks, margins, page size, symbols
- **styleChecks.js**: Caption consistency, blank paragraph marks, section breaks
- **headerFooterChecks.js**: Header/footer formatting validation

#### fixes/
- **autoFix.js**: Automated corrections for all fixable issues

#### navigation/
- **navigation.js**: Document navigation using bookmarks

### Configuration
- **rules.json**: Centralized validation rules and settings

## Performance Optimizations

1. **Batch Bookmark Creation**: All bookmarks created in single sync
2. **Regex Pre-compilation**: Compiled once, reused for all paragraphs
3. **Early Exit Checks**: Skip processing when content doesn't match
4. **OOXML Batch Loading**: Load all paragraph OOXML at once
5. **Variable Hoisting**: Constants extracted outside loops

## Import Pattern

```javascript
// Barrel exports enable clean imports
import { analyzeFormatting, checkDocument, checkStyles } from '../services/checks';
import { autoFixIssues } from '../services/fixes';
import { goToError } from '../services/navigation';
```

## UI Features

### Auto-Fix Functionality

- **Toggle**: Enable/disable automatic fixing of issues
- **Manual Fix**: "Fix Issues Now" button for on-demand corrections
- **Results Display**: Shows count of fixed and unfixed issues
- **Supported Fixes**: Font, Font Size, Font Color, Highlighting, Hidden Text, Justification, Hyperlinks, Comments, Revisions, Blank Paragraph Marks

### Loading Indicators

- **Overlay Message**: "Check is currently running, please hold..."
- **Animated Spinner**: Visual feedback during operations
- **Fixing Message**: "Fixing issues, please hold..." during auto-fix

### Color-Coded Categories

Issues are color-coded by category for quick visual identification:

| Category | Color | Issue Types |
|----------|-------|-------------|
| 🔵 **Formatting Check** | Light Blue (`#e3f2fd`) | Font, Font Size, Font Color, Highlighting, Hidden Text, Justification, Hyperlinks, Blank Pages, Table Headers |
| 🟠 **General Document** | Light Orange (`#fff3e0`) | Comments, Revisions, Text Boxes, Watermarks, Invalid References, Page Size |
| 🟣 **Headers/Footers** | Light Purple (`#f3e5f5`) | Header, Footer |
| 🔴 **Margins** | Light Pink (`#fce4ec`) | Margins |
| 🟢 **Styles** | Light Green (`#e8f5e9`) | Captions, Blank Paragraph Marks, Section/Page Breaks, Headings |
| 🟡 **Symbols** | Light Yellow (`#fff9c4`) | Symbols |

A color legend is displayed at the top of the Comprehensive Checker for easy reference.

## Benefits of New Architecture

✅ **Separation of Concerns**: UI, logic, and config are separate
✅ **Scalability**: Easy to add new checkers or fixers
✅ **Maintainability**: Clear module boundaries
✅ **Testability**: Services can be tested independently
✅ **Reusability**: Barrel exports make imports clean
✅ **Performance**: Optimized checks with minimal context syncs
✅ **User Experience**: Visual feedback with loading states and color coding
