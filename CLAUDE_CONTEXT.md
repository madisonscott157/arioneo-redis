# Arioneo Horse Training Data Application - Claude Context

## Project Overview
- **App URL**: https://arioneo-redis.vercel.app
- **GitHub**: https://github.com/madisonscott157/arioneo-redis.git
- **Stack**: Node.js/Express backend, vanilla JS frontend, Upstash Redis for persistence, hosted on Vercel

## What the App Does
- Displays horse training data from uploaded Excel/CSV files
- Main table shows horses with their most recent training data (1F, 5F times, Fast Recovery, 15 min Recovery)
- Click a horse to see full training history
- Supports Active/Historic horse views
- Filters by owner, country, age
- **NEW**: Upload race chart PDFs to add race results to horse training profiles

## Project Structure (Recently Reorganized)
```
arioneo-redis-main/
├── server.js              # Express backend (all API routes)
├── package.json           # Dependencies include pdf-parse, fuse.js
├── public/
│   ├── index.html         # Main HTML (237 lines, references external CSS/JS)
│   ├── css/
│   │   └── styles.css     # All CSS styles (~1700 lines)
│   └── js/
│       ├── app.js         # Main application JS (~3400 lines)
│       └── race-upload.js # Race chart upload module (~700 lines)
└── CLAUDE_CONTEXT.md      # This file
```

## CURRENT WORK IN PROGRESS: Race Chart Upload Feature

### What It Does
- Upload Equibase race chart PDFs (bulk upload, up to 50 at once)
- Parse race data: date, track, surface, distance, race type, final time, positions, comments
- Fuzzy match horses in chart to existing horses in system
- Review/edit parsed data before saving
- Save race results to horse training profiles

### Implementation Status
- **Working**: PDF parsing, horse detection, data extraction, modal UI, horse dropdown selection
- **Just Fixed**: Server 500 error on save (null check), modal height issues
- **Needs Testing**: Full end-to-end flow after latest fixes

### Key Files Modified
1. **server.js** (lines ~2350-3400):
   - `RaceChartParser` class - parses Equibase PDF format
   - `POST /api/upload/race-charts` - handles bulk PDF upload and parsing
   - `POST /api/race-charts/save` - saves reviewed race data
   - `buildHorseSummaryFromDetailData()` - rebuilds horse summary
   - `fuzzyMatchHorse()` - uses fuse.js for name matching
   - `checkDuplicateRace()` - prevents duplicate entries

2. **public/js/race-upload.js** (new file):
   - `RaceChartUploader` class - modal UI for upload workflow
   - `buildHorseSelector()` - dropdown with horses found in chart
   - `handleHorseChange()` - updates form when different horse selected
   - `updateFormFields()` - populates form with selected horse's data
   - `collectRaceData()` / `saveAllRaces()` - save workflow

3. **public/css/styles.css** (lines ~1260-1750):
   - Race upload modal styles
   - Review card styles
   - Made modal 90vh tall with scrollable review list

### Data Mapping (Race to Training Entry)
When a race is saved, it creates a training entry with these mappings:
- `type`: "Race"
- `maxSpeed`: Race type (MSW, AOC, G1, etc.)
- `best1f`: 1/4 position
- `best2f`: 1/4 time
- `best3f`: 1/2 position
- `best4f`: 1/2 time
- `best5f`: 5F reduction time (calculated)
- `best6f`: 3/4 position
- `best7f`: 3/4 time
- `maxHR`: Final time
- `fastRecovery`: Finish position
- `notes`: Chart comments

### Recent Bug Fixes (This Session)
1. Surface detection - Fixed to detect "Inner Turf"/"Turf" correctly
2. Race type - Fixed "Allowance Optional Claiming" -> AOC
3. Horse name extraction - Now searches after "H Wt" header
4. Final time regex - Only matches `[12]:\d{2}\.\d{2}` format
5. Position formatting - Fixed "31th" -> "3rd"
6. `odds is not defined` error - Removed from return statement
7. Horse dropdown - Now shows all horses found in chart
8. Data switching - Form updates when different horse selected
9. Server 500 error - Added null checks for `horse.horse` property
10. Modal height - Made 90vh tall with flex layout

### Known Issues / TODO
- Modal may still need height adjustment on smaller screens
- Need to verify main page keeps all horses after race save
- Slowness during save (many Redis calls, expected without Redis configured)

## Key Features (Previously Implemented)

### Horse Mapping System
- Manage Horses modal for owner/country assignments
- **Merge Horses**: Combine multiple names for same horse
- **Rename Horse**: Change display name while keeping training data linked via alias
- **Auto-add**: New horses from uploads automatically added to mappings

### Training Data
- Edit individual training entries (type, track, surface, notes)
- Notes display as truncated text with full tooltip
- CSV export includes notes

### Data Flow
1. Upload Excel/CSV -> processed on server
2. Data stored in Redis (session data + horse mapping + training edits)
3. On load: fetch session -> apply alias merging -> apply training edits -> display

## Important Code Locations

### Server (server.js)
- `generateHorseSummary()` - Creates main table data from training entries
- `resolveHorseAlias()` - Maps alias names to primary names
- `mergeAliasedHorseData()` - Combines training data for aliased horses
- `RaceChartParser` class - PDF parsing (lines ~2350-2970)
- `/api/upload/race-charts` - Race PDF upload endpoint
- `/api/race-charts/save` - Save races endpoint

### Client
- `public/js/app.js` - Main application logic
- `public/js/race-upload.js` - Race chart upload module
- `public/css/styles.css` - All styles

## Testing Notes
- Local testing requires `npm install` first
- Redis errors expected locally (no credentials) - falls back to memory
- Test PDFs in `/Users/madisonscott/Desktop/Claude/Streamline/`: paradise.pdf, ontario.pdf, blanco.pdf
- Deploy via git push to main - Vercel auto-deploys
