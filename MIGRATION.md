# Migration Guide

## Code Reorganization Summary

The project has been reorganized into a cleaner, more maintainable architecture. Here's what changed:

### 📁 Old Structure
```
src/taskpane/
├── wordChecks.js
├── docChecks.js
├── checkStyles.js
├── checkHeaderFooterFormatting.js
├── autoFix.js
├── components/App.jsx
└── config/rules.json
```

### 📁 New Structure
```
src/taskpane/
├── components/App.jsx
├── services/
│   ├── checks/
│   │   ├── index.js (barrel export)
│   │   ├── formattingChecks.js (was wordChecks.js)
│   │   ├── documentChecks.js (was docChecks.js)
│   │   ├── styleChecks.js (was checkStyles.js)
│   │   └── headerFooterChecks.js (was checkHeaderFooterFormatting.js)
│   ├── fixes/
│   │   ├── index.js (barrel export)
│   │   └── autoFix.js
│   └── navigation/
│       ├── index.js (barrel export)
│       └── navigation.js (extracted from wordChecks.js)
└── config/rules.json
```

## Import Changes

### ❌ Old Imports
```javascript
import { analyzeFormatting, goToError } from "../wordChecks";
import { checkDocument } from "../docChecks";
import { checkStyles } from "../checkStyles";
import { checkHeaderFooterFormatting } from "../checkHeaderFooterFormatting";
import { autoFixIssues } from "../autoFix";
```

### ✅ New Imports
```javascript
import { analyzeFormatting, checkDocument, checkStyles, checkHeaderFooterFormatting } from "../services/checks";
import { autoFixIssues } from "../services/fixes";
import { goToError } from "../services/navigation";
```

## Benefits

1. **Better Organization**: Related functionality grouped together
2. **Cleaner Imports**: Barrel exports reduce import statements
3. **Separation of Concerns**: Clear boundaries between modules
4. **Easier Testing**: Services can be tested independently
5. **Scalability**: Easy to add new checks or fixes

## Breaking Changes

⚠️ **None** - All functionality remains the same, only file locations changed.

## What to Do

If you have custom code that imports the old files:

1. Update import paths to use new structure
2. Use barrel exports from `services/checks`, `services/fixes`, `services/navigation`
3. Check `ARCHITECTURE.md` for complete documentation

## Questions?

See `ARCHITECTURE.md` for detailed architecture documentation.
