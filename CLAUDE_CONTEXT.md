# Arioneo Horse Training Data Application - Claude Context

## Project Overview
- **App URL**: https://arioneo-redis.vercel.app
- **GitHub**: https://github.com/madisonscott157/arioneo-redis.git
- **Stack**: Node.js/Express backend, vanilla JS frontend, Upstash Redis for persistence, hosted on Vercel

## What the App Does
- Displays horse training data from uploaded Excel/CSV files
- Main table shows horses sorted by most recent training date (default)
- Click a horse to see full training history (detail view)
- Supports Active/Historic horse views
- Filters by owner, country, age
- Upload race chart PDFs to add race results to horse training profiles

## Project Structure
```
arioneo-redis-main/
├── server.js              # Express backend (all API routes)
├── package.json           # Dependencies: pdf-parse, fuse.js, express, etc.
├── public/
│   ├── index.html         # Main HTML (loads xlsx.js, exceljs.js, CSS, JS)
│   ├── css/
│   │   └── styles.css     # All CSS styles (~1900 lines)
│   └── js/
│       ├── app.js         # Main application JS (~3500 lines)
│       └── race-upload.js # Race chart upload module (~700 lines)
└── CLAUDE_CONTEXT.md      # This file
```

## External Libraries (loaded via CDN in index.html)
- **SheetJS (xlsx.full.min.js)** - For reading Excel/CSV uploads
- **ExcelJS (exceljs.min.js)** - For exporting Excel files with styling (just added)

## Recent Changes (This Session)

### UI Improvements
1. **CSV Upload Modal** - Drag-and-drop modal for uploading CSV/Excel files
   - `openCsvUploadModal()`, `closeCsvUploadModal()`, `initCsvDropzone()` in app.js
   - CSS in styles.css under "CSV UPLOAD MODAL STYLES"

2. **Manage Horses Modal** - Added filter and search
   - Status filter dropdown (Active/Historic/All) - defaults to Active
   - Search box to filter horses by name, owner, or aliases
   - `filterHorseMappings()` function in app.js

3. **Default Sort** - Main table always defaults to most recent training date
   - `currentSort = { column: 'lastTrainingDate', order: 'desc' }` set on:
     - Page load
     - After any upload
     - When returning from horse detail view (`showMainView()`)
   - Removed sort preference restoration from `loadUserPreferences()`

### Export Functionality
1. **Main Page Export** - Uses SheetJS, exports as .xlsx
   - `exportToCsv()` function in app.js
   - Columns: Horse Name, Owner, Country, Last Training, Age, 1F, 5F, Fast, 15 min
   - Time columns forced to text format to prevent Excel conversion

2. **Horse Detail Export** - Uses ExcelJS for styling support
   - `exportHorseDataToCsv()` function in app.js (async)
   - Exports all training columns
   - **ISSUE**: Color coding not working - attempted to add background colors for:
     - Best 5F column (blue/green/cream/yellow/red based on time)
     - Fast Recovery column (based on numeric value)
     - 15 Recovery column (based on numeric value)

### Color Coding Logic (defined in app.js)
```javascript
// Best 5F - getBest5FColor(timeStr)
<= 60 sec: '#d1ecf1' (light blue - fastest)
<= 65 sec: '#d4edda' (light green)
<= 70 sec: '#f9f7e3' (light cream)
<= 75 sec: '#fff3cd' (light yellow)
> 75 sec: '#fdeaea' (light red - slowest)

// Fast Recovery - getFastRecoveryColor(value)
>= 140: '#fdeaea' (light red)
>= 125: '#fff3cd' (light yellow)
>= 119: '#f9f7e3' (light cream)
>= 101: '#d4edda' (light green)
< 101: '#d1ecf1' (light blue)

// 15 Recovery - getRecovery15Color(value)
>= 116: '#fdeaea' (light red)
>= 102: '#fff3cd' (light yellow)
>= 81: '#d4edda' (light green)
< 81: '#d1ecf1' (light blue)
```

### Bug Fixes
1. **Export button not working** - Added null check for removed `sortBy` element (line ~507)
2. **sortBy dropdown removed** - Was causing JS error that blocked other listeners

## Race Chart Position Parsing (RESOLVED)

### Issue Summary
Race chart PDF position parsing had multiple bugs affecting different PDF formats.

### Bug 1: Code Not Deployed
**Problem**: Changes weren't committed to git, so Vercel was running old code.
**Fix**: Always commit and push before expecting changes on live site.

### Bug 2: Double-Digit Positions (Florida Derby format)
**Problem**: Position string "109765" was parsed as [1, 0, 9, 7, 6, 5] instead of [10, 9, 7, 6, 5].
**Fix**: Added `parsePositions()` method that detects when "1" is followed by "0" (since position 0 doesn't exist, "10" must be 10th place).

### Bug 3: Wrong Final Time Selected (Saratoga format)
**Problem**: Horse boundary detection included next horse's data, so code found two M:SS.SS times and used the LAST one (wrong horse's time).
**Fix**: Use FIRST M:SS.SS time as final time, not last.

### Bug 4: Ambiguous "11" Parsing
**Problem**: "1111" was parsed as [11, 11] instead of [1, 1, 1, 1] (four 1st places).
**Root Cause**: Only "10" is unambiguous (position 0 invalid). "11" could be position 11 OR positions 1,1.
**Fix**: Added `expectedCount` parameter to `parsePositions()`:
- If digit count matches expected positions → parse all as singles
- If digit count exceeds expected → apply double-digit logic

### Final parsePositions Logic (server.js ~lines 3306-3347)
```javascript
parsePositions(digitsStr, expectedCount = 0)
// - "1111" with expected 4 → [1, 1, 1, 1] ✓
// - "109765" with expected 5 → [10, 9, 7, 6, 5] ✓
// - "119765" with expected 5 → [11, 9, 7, 6, 5] ✓
// - "12865" with expected 5 → [1, 2, 8, 6, 5] ✓ (not 12)
```

### Two PDF Formats Supported
1. **Format A (Florida Derby style)**: Positions concatenated as digits after final time
   - Example: "109765" → [10, 9, 7, 6, 5]
2. **Format B (Saratoga style)**: Space-separated tokens with margins attached
   - Example: "11 1hd 1½ 11" → positions 1, 1, 1, 1 (margins on separate line in text extraction)

## Key Functions Reference

### app.js
- `exportToCsv()` - Main page export (SheetJS)
- `exportHorseDataToCsv()` - Horse detail export (ExcelJS) - NEEDS FIX
- `getBest5FColor()`, `getFastRecoveryColor()`, `getRecovery15Color()` - Color logic
- `filterHorseMappings()` - Manage horses filter
- `showMainView()` - Return to main table (resets sort)
- `handleArioneoUpload()` - CSV upload handler
- `openCsvUploadModal()` - Opens drag-drop modal

### server.js
- `RaceChartParser` class - PDF parsing
- `/api/upload/race-charts` - Race PDF upload
- `/api/race-charts/save` - Save races
- `/api/upload/arioneo` - CSV/Excel upload

## Testing Notes
- Local: `npm install` then `node server.js`
- Deploy: `git add -A && git commit -m "message" && git push`
- Vercel auto-deploys on push to main
