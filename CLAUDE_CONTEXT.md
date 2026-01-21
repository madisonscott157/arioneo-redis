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

## CURRENT ISSUE TO FIX

**Excel export color coding not working for horse detail page.**

The `exportHorseDataToCsv()` function uses ExcelJS to export with cell background colors, but the colors are not appearing in the exported Excel file.

Current implementation (app.js lines ~2813-2917):
- Uses `ExcelJS.Workbook()` and `worksheet.addRow()`
- Attempts to set `cell.fill` with `type: 'pattern'`, `pattern: 'solid'`, `fgColor: { argb: '...' }`
- Colors are being calculated correctly (using getBest5FColor, etc.)
- But the exported file has no colors

Column indices for colored columns (1-based for ExcelJS):
- BEST5F_COL = 13
- FAST_RECOVERY_COL = 17
- RECOVERY15_COL = 20

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
