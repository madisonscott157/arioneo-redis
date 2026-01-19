const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const cors = require('cors');
const { v4: uuidv4 } = require('uuid');
const path = require('path');
const fs = require('fs');
const compression = require('compression');
const helmet = require('helmet');

// Import Upstash Redis for persistent storage
let redis = null;
try {
  const { Redis } = require('@upstash/redis');
  redis = new Redis({
    url: process.env.KV_REST_API_URL,
    token: process.env.KV_REST_API_TOKEN,
  });
  console.log('Upstash Redis loaded successfully');
} catch (error) {
  console.log('Upstash Redis not available, using memory storage:', error.message);
}

const app = express();
const port = process.env.PORT || 3000;

// Middleware
app.use(helmet({
  contentSecurityPolicy: false, // Temporarily disable CSP to allow onclick handlers
}));
app.use(compression());
app.use(cors({
  origin: ['http://localhost:3000', 'http://127.0.0.1:3000'],
  credentials: true
}));
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ limit: '50mb', extended: true }));

// Serve static files
app.use(express.static('public'));

// Create uploads directory if it doesn't exist
const uploadsDir = process.env.NODE_ENV === 'production' ? '/tmp/uploads' : path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadsDir)) {
  fs.mkdirSync(uploadsDir, { recursive: true });
}

// Configure multer for file uploads
const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, uploadsDir);
  },
  filename: function (req, file, cb) {
    const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
    cb(null, file.fieldname + '-' + uniqueSuffix + path.extname(file.originalname));
  }
});

const upload = multer({ 
  storage: storage,
  limits: {
    fileSize: 50 * 1024 * 1024, // 50MB limit
  },
  fileFilter: function (req, file, cb) {
    const allowedTypes = ['.xlsx', '.xls', '.xlsm', '.csv'];
    const ext = path.extname(file.originalname).toLowerCase();
    if (allowedTypes.includes(ext)) {
      cb(null, true);
    } else {
      cb(new Error('Only Excel and CSV files are allowed'));
    }
  }
});

// Hybrid storage: KV for persistence, memory for speed

// Initialize global storage as backup
if (typeof global.sessionStorage === 'undefined') {
  global.sessionStorage = new Map();
  global.latestSession = null;
}

// Storage functions with Redis + memory backup
async function saveSession(sessionId, fileName, horseData, allHorseDetailData, sheetData = null) {
  const sessionData = {
    id: sessionId,
    fileName: fileName,
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    horseData: horseData,
    allHorseDetailData: allHorseDetailData,
    // Include sheet data if provided
    allSheets: sheetData?.allSheets || null,
    sheetNames: sheetData?.sheetNames || null,
    currentSheetName: sheetData?.currentSheetName || null
  };
  
  // Try to save to Redis first (persistent storage)
  try {
    if (redis) {
      await redis.set(`session:${sessionId}`, JSON.stringify(sessionData));
      await redis.set('latest_session', JSON.stringify({ sessionId: sessionId, updatedAt: sessionData.updatedAt }));
      console.log('Session saved to Redis:', sessionId);
    }
  } catch (error) {
    console.error('Error saving to Redis:', error);
  }
  
  // Always store in memory as backup
  global.sessionStorage.set(sessionId, sessionData);
  global.latestSession = { sessionId: sessionId, updatedAt: sessionData.updatedAt };
  
  console.log('Session saved to memory:', sessionId);
}

// Function to save/update sheet data for an existing session
async function updateSessionSheets(sessionId, sheetData) {
  try {
    // Get existing session data first
    const existingSession = await getSession(sessionId);
    if (!existingSession) {
      throw new Error('Session not found');
    }

    // Update session with new sheet data
    const updatedSessionData = {
      ...existingSession,
      allSheets: sheetData.allSheets,
      sheetNames: sheetData.sheetNames,
      currentSheetName: sheetData.currentSheetName,
      updatedAt: new Date().toISOString()
    };

    // Save to Redis
    if (redis) {
      await redis.set(`session:${sessionId}`, JSON.stringify(updatedSessionData));
      console.log('Sheet data updated in Redis for session:', sessionId);
    }

    // Update memory cache
    global.sessionStorage.set(sessionId, updatedSessionData);
    console.log('Sheet data updated in memory for session:', sessionId);

    return updatedSessionData;
  } catch (error) {
    console.error('Error updating session sheets:', error);
    throw error;
  }
}

async function getSession(sessionId) {
  try {
    // Try Redis first (persistent storage)
    if (redis) {
      console.log('Trying to get session from Redis:', sessionId);
      const redisResult = await redis.get(`session:${sessionId}`);
      console.log('Redis result:', redisResult ? 'found' : 'not found');
      if (redisResult) {
        // Upstash Redis auto-parses JSON, so check if it's already an object
        const result = typeof redisResult === 'string' ? JSON.parse(redisResult) : redisResult;
        console.log('Session retrieved from Redis:', sessionId);
        // Update memory cache
        global.sessionStorage.set(sessionId, result);
        return result;
      }
    }
    
    // Fallback to memory
    const result = global.sessionStorage.get(sessionId);
    if (result) {
      console.log('Session retrieved from memory:', sessionId);
      return result;
    }
    
    console.log('Session not found:', sessionId);
    return null;
  } catch (error) {
    console.error('Error reading session:', error);
    return global.sessionStorage.get(sessionId) || null;
  }
}

async function getLatestSession() {
  try {
    // Try Redis first (persistent storage)
    if (redis) {
      const redisResult = await redis.get('latest_session');
      if (redisResult) {
        // Upstash Redis auto-parses JSON, so check if it's already an object
        const result = typeof redisResult === 'string' ? JSON.parse(redisResult) : redisResult;
        console.log('Latest session retrieved from Redis');
        // Update memory cache
        global.latestSession = result;
        return result;
      }
    }
    
    // Fallback to memory
    const result = global.latestSession;
    if (result) {
      console.log('Latest session retrieved from memory');
    }
    return result;
  } catch (error) {
    console.error('Error reading latest session:', error);
    return global.latestSession;
  }
}

// Helper function to process Excel data (extracted from your original code)
function processExcelData(filePath) {
  try {
    const workbook = XLSX.readFile(filePath, {
      cellStyles: true,
      cellHTML: true,
      cellFormula: true,
      bookSST: true,
      cellNF: true,
      cellDates: true,
      raw: false
    });

    const horseData = [];
    const allHorseDetailData = {};

    workbook.SheetNames.forEach(sheetName => {
      const sheetNameLower = sheetName.toLowerCase();
      if (sheetNameLower === 'sheet1' || sheetNameLower === 'sheet2' || sheetNameLower === 'buttons') {
        return;
      }

      const worksheet = workbook.Sheets[sheetName];
      if (!worksheet) return;

      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      const horseStats = processSheetData(jsonData, sheetName);
      
      if (horseStats && horseStats.horseData) {
        horseData.push(horseStats.horseData);
        allHorseDetailData[sheetName] = horseStats.detailData;
      }
    });

    return {
      horseData,
      allHorseDetailData
    };
  } catch (error) {
    console.error('Error processing Excel file:', error);
    throw error;
  }
}

function processSheetData(jsonData, horseName) {
  if (jsonData.length === 0) return null;

  const headers = jsonData[0].map(h => h ? h.toString().toLowerCase().trim() : '');
  
  let age = null;
  let best1f = null;
  let best5f = null;
  let dateOfBest5f = null;
  let maxSpeed = null;
  let fastRecovery = null;
  let recovery15min = null;
  
  const times1f = [];
  const times5f = [];
  const speeds = [];
  const detailData = [];

  for (let i = 1; i < jsonData.length; i++) {
    const row = jsonData[i];
    if (!row || row.length === 0) continue;

    const rowData = {};
    headers.forEach((header, index) => {
      rowData[header] = row[index] ? row[index].toString().trim() : '';
    });

    let dateValue = row[0] || rowData.date || '';
    if (typeof dateValue === 'number' && dateValue > 40000) {
      const excelEpoch = new Date(1899, 11, 30);
      const jsDate = new Date(excelEpoch.getTime() + dateValue * 24 * 60 * 60 * 1000);
      dateValue = jsDate.toLocaleDateString();
    } else if (dateValue instanceof Date) {
      dateValue = dateValue.toLocaleDateString();
    } else if (dateValue.toString().includes('T')) {
      const date = new Date(dateValue);
      dateValue = date.toLocaleDateString();
    }

    const typeValue = row[2] || rowData.type || '';
    const isRace = typeValue.toString().toLowerCase().includes('race');
    const isWork = typeValue.toString().toLowerCase().includes('work');

    const detailRow = {
      date: dateValue,
      horse: rowData.horse || horseName,
      type: typeValue,
      track: rowData.track || '',
      surface: rowData.surface || '',
      distance: rowData.distance || '',
      avgSpeed: rowData['avg speed'] || '',
      maxSpeed: rowData['max speed'] || '',
      best1f: rowData['best 1f'] || '',
      best2f: rowData['best 2f'] || '',
      best3f: rowData['best 3f'] || '',
      best4f: rowData['best 4f'] || '',
      best5f: rowData['best 5f'] || '',
      best6f: rowData['best 6f'] || '',
      best7f: rowData['best 7f'] || '',
      maxHR: rowData['max hr'] || '',
      fastRecovery: rowData['fast recovery'] || '',
      fastQuality: rowData['fast quality'] || '',
      fastPercent: rowData['fast %'] || '',
      recovery15: rowData['15 recovery'] || '',
      quality15: rowData['15 quality'] || '',
      hr15Percent: rowData['hr 15%'] || '',
      maxSL: rowData['max sl'] || '',
      slGallop: rowData['sl gallop'] || '',
      sfGallop: rowData['sf gallop'] || '',
      slWork: rowData['sl work'] || '',
      sfWork: rowData['sf work'] || '',
      hr2min: rowData['hr 2 min'] || '',
      hr5min: rowData['hr 5 min'] || '',
      symmetry: rowData.symmetry || '',
      regularity: rowData.regularity || '',
      bpm120: rowData['120bpm'] || '',
      zone5: rowData['zone 5'] || '',
      age: rowData.age || '',
      sex: rowData.sex || '',
      temp: rowData.temp || '',
      distanceCol: rowData.distance || '',
      trotHR: rowData['trot hr'] || '',
      walkHR: rowData['walk hr'] || '',
      isRace: isRace,
      isWork: isWork
    };

    if (detailRow.date || detailRow.horse || detailRow.type) {
      detailData.push(detailRow);
    }

    if (!isRace) {
      if (!age && detailRow.age) {
        const parsedAge = parseInt(detailRow.age);
        if (!isNaN(parsedAge)) {
          age = parsedAge;
        }
      }

      if (detailRow.best1f && isValidTime(detailRow.best1f)) {
        times1f.push({ time: detailRow.best1f, seconds: timeToSeconds(detailRow.best1f) });
      }

      if (detailRow.best5f && isValidTime(detailRow.best5f)) {
        times5f.push({
          time: detailRow.best5f,
          seconds: timeToSeconds(detailRow.best5f),
          date: dateValue.toString(),
          fastRecovery: detailRow.fastRecovery,
          recovery15min: detailRow.recovery15
        });
      }

      if (detailRow.maxSpeed) {
        const speed = parseFloat(detailRow.maxSpeed);
        if (!isNaN(speed)) {
          speeds.push(speed);
        }
      }
    }
  }

  if (times1f.length > 0) {
    const fastest1f = times1f.reduce((min, current) =>
      current.seconds < min.seconds ? current : min
    );
    best1f = fastest1f.time;
  }

  if (times5f.length > 0) {
    const fastest5f = times5f.reduce((min, current) =>
      current.seconds < min.seconds ? current : min
    );
    best5f = fastest5f.time;
    dateOfBest5f = fastest5f.date;
    fastRecovery = fastest5f.fastRecovery;
    recovery15min = fastest5f.recovery15min;
  }

  if (speeds.length > 0) {
    maxSpeed = Math.max(...speeds);
  }

  return {
    horseData: {
      name: horseName,
      age: age,
      best1f: best1f,
      best5f: best5f,
      dateOfBest5f: dateOfBest5f,
      fastRecovery: fastRecovery,
      recovery15min: recovery15min,
      maxSpeed: maxSpeed,
      best5fColor: getBest5FColor(best5f),
      fastRecoveryColor: getFastRecoveryColor(fastRecovery),
      recovery15Color: getRecovery15Color(recovery15min)
    },
    detailData: detailData
  };
}

function isValidTime(timeStr) {
  if (!timeStr) return false;
  const str = timeStr.toString().trim();
  if (str === '-' || str === '' || str === 'NaN') return false;
  return /^\d{2}:\d{2}\.\d{2}$/.test(str);
}

function timeToSeconds(timeStr) {
  const parts = timeStr.toString().trim().split(':');
  const minutes = parseInt(parts[0]);
  const secondsParts = parts[1].split('.');
  const seconds = parseInt(secondsParts[0]);
  const hundredths = parseInt(secondsParts[1]);
  
  return minutes * 60 + seconds + hundredths / 100;
}

function getFastRecoveryColor(value) {
  if (!value || value === '-') return null;
  
  const numValue = parseFloat(value);
  if (isNaN(numValue)) return null;
  
  if (numValue >= 140) return '#fdeaea';
  if (numValue >= 125) return '#fff3cd';
  if (numValue >= 119) return '#f9f7e3';
  if (numValue >= 101) return '#d4edda';
  return '#d1ecf1';
}

function getRecovery15Color(value) {
  if (!value || value === '-') return null;
  
  const numValue = parseFloat(value);
  if (isNaN(numValue)) return null;
  
  if (numValue >= 116) return '#fdeaea';
  if (numValue >= 102) return '#fff3cd';
  if (numValue >= 81) return '#d4edda';
  return '#d1ecf1';
}

function getBest5FColor(timeStr) {
  if (!timeStr || timeStr === '-' || !isValidTime(timeStr)) return null;

  const seconds = timeToSeconds(timeStr);

  if (seconds <= 60) return '#d1ecf1';
  if (seconds <= 65) return '#d4edda';
  if (seconds <= 70) return '#f9f7e3';
  if (seconds <= 75) return '#fff3cd';
  return '#fdeaea';
}

// ============================================
// ARIONEO CSV TRANSFORMATION FUNCTIONS
// (Replaces Excel macros)
// ============================================

// Track name abbreviations
function transformTrackName(trackName) {
  if (!trackName) return '';
  const track = trackName.toString().toLowerCase();

  // Saratoga
  if (track.includes('saratoga - turf') || track.includes('saratoga - grass')) return 'SAR T';
  if (track.includes('saratoga - main') || track.includes('saratoga')) return 'SAR';

  // Oklahoma (Saratoga training)
  if (track.includes('oklahoma grass') || track.includes('oklahoma turf')) return 'SARtr T';
  if (track.includes('oklahoma dirt') || track.includes('oklahoma')) return 'SARtr';

  // Belmont
  if (track.includes('belmont training')) return 'BELtr';
  if (track.includes('belmont main') || track.includes('belmont')) return 'BEL';

  // Churchill
  if (track.includes('churchill')) return 'CD';

  // Keeneland
  if (track.includes('keeneland')) return 'KEE';

  // Payson
  if (track.includes('payson turf') || track.includes('payson grass')) return 'PAY T';
  if (track.includes('payson dirt') || track.includes('payson')) return 'PAY';

  // Palm Beach
  if (track.includes('palm beach turf') || track.includes('palm beach grass')) return 'PBD T';
  if (track.includes('palm beach dirt') || track.includes('palm beach downs')) return 'PBD';

  // Turfway
  if (track.includes('turfway')) return 'TP';

  // Palm Meadows
  if (track.includes('palm meadows')) return 'PMM';

  // Winstar
  if (track.includes('winstar')) return 'WS';

  // Woodbine
  if (track.includes('woodbine')) return 'WO';

  // Gulfstream
  if (track.includes('gulfstream')) return 'GP';

  // Aqueduct
  if (track.includes('aqueduct')) return 'AQU';

  return trackName; // Return original if no match
}

// Surface abbreviations
function transformSurface(surface) {
  if (!surface) return '';
  const s = surface.toString().toLowerCase();

  if (s.includes('weather')) return 'AWT';
  if (s.includes('dirt')) return 'D';
  if (s.includes('turf') || s.includes('grass')) return 'T';

  return surface; // Return original if no match
}

// Training type normalization
function transformTrainingType(type) {
  if (!type) return '';
  const t = type.toString().toLowerCase().trim();

  // Clear N/A
  if (t === 'n/a') return '';

  // Gate work variations
  if (t.includes('work g') || t.includes('work gate') || t.includes('work from the gate') ||
      t.includes('breeze gate') || t.includes('breeze g')) {
    return 'Work - G';
  }

  // Breeze to Work
  if (t.includes('breeze')) {
    return type.toString().replace(/breeze/gi, 'Work');
  }

  return type;
}

// Date formatting - remove time, format as MM/DD/YYYY
function transformDate(dateValue) {
  if (!dateValue) return '';

  let date;

  // Handle Excel serial date
  if (typeof dateValue === 'number' && dateValue > 40000) {
    const excelEpoch = new Date(1899, 11, 30);
    date = new Date(excelEpoch.getTime() + dateValue * 24 * 60 * 60 * 1000);
  }
  // Handle Date object
  else if (dateValue instanceof Date) {
    date = dateValue;
  }
  // Handle ISO string with time
  else if (typeof dateValue === 'string' && dateValue.includes('T')) {
    date = new Date(dateValue);
  }
  // Handle other date strings
  else if (typeof dateValue === 'string') {
    // Try parsing as date string (handles "01/19/2026 11:40" format)
    const parsed = new Date(dateValue);
    if (!isNaN(parsed.getTime())) {
      date = parsed;
    } else {
      return dateValue; // Return as-is if can't parse
    }
  }
  else {
    return dateValue.toString();
  }

  // Format as MM/DD/YYYY
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const year = date.getFullYear();

  return `${month}/${day}/${year}`;
}

// Process raw Arioneo CSV data
function processArioneoCSV(csvData) {
  const rows = [];
  const headers = csvData[0].map(h => h ? h.toString().trim() : '');

  console.log('CSV Headers found (all):', headers);

  // Map Arioneo column names to our internal names (with many variations)
  const columnMap = {
    // Date variations
    'date': 'date',
    'training date': 'date',
    'session date': 'date',

    // Horse name variations - CRITICAL for grouping
    'horse': 'horse',
    'horse name': 'horse',
    'horsename': 'horse',
    'name': 'horse',
    'animal': 'horse',
    'animal name': 'horse',
    'equine': 'horse',
    'equine name': 'horse',
    'cheval': 'horse',  // French

    // Training type variations
    'training type': 'type',
    'type': 'type',
    'session type': 'type',
    'workout type': 'type',
    'activity': 'type',
    'activity type': 'type',

    // Track variations
    'track name': 'track',
    'track': 'track',
    'location': 'track',
    'venue': 'track',
    'training location': 'track',

    // Surface variations
    'track surface': 'surface',
    'surface': 'surface',
    'ground': 'surface',
    'footing': 'surface',

    // Distance variations
    'working distance': 'distance',
    'distance': 'distance',
    'work distance': 'distance',

    // Speed variations
    'main working average speed': 'avgSpeed',
    'avg speed': 'avgSpeed',
    'average speed': 'avgSpeed',
    'avg. speed': 'avgSpeed',
    'max speed': 'maxSpeed',
    'maximum speed': 'maxSpeed',
    'top speed': 'maxSpeed',

    // Furlong time variations
    'time best 1f': 'best1f',
    'best 1f': 'best1f',
    '1f': 'best1f',
    'time best 2f': 'best2f',
    'best 2f': 'best2f',
    '2f': 'best2f',
    'time best 3f': 'best3f',
    'best 3f': 'best3f',
    '3f': 'best3f',
    'time best 4f': 'best4f',
    'best 4f': 'best4f',
    '4f': 'best4f',
    'time best 5f': 'best5f',
    'best 5f': 'best5f',
    '5f': 'best5f',
    'time best 6f': 'best6f',
    'best 6f': 'best6f',
    '6f': 'best6f',
    'time best 7f': 'best7f',
    'best 7f': 'best7f',
    '7f': 'best7f',

    // Heart rate variations
    'max heart rate reached during training': 'maxHR',
    'max hr': 'maxHR',
    'max heart rate': 'maxHR',
    'maximum heart rate': 'maxHR',
    'fast recovery': 'fastRecovery',
    'fast recovery quality': 'fastQuality',
    'fast recovery in % of max hr': 'fastPercent',
    'heart rate after 15 min': 'recovery15',
    '15 recovery': 'recovery15',
    '15min recovery': 'recovery15',
    '15min recovery quality': 'quality15',
    '15 quality': 'quality15',
    'hr after 15 min in % of max hr': 'hr15Percent',

    // Stride variations
    'max stride length': 'maxSL',
    'max sl': 'maxSL',
    'stride length at 20.5 mph': 'slGallop',
    'sl gallop': 'slGallop',
    'stride frequency at 20.5 mph': 'sfGallop',
    'sf gallop': 'sfGallop',
    'stride length at 37.3 mph': 'slWork',
    'sl work': 'slWork',
    'stride frequency at 37.3 mph': 'sfWork',
    'sf work': 'sfWork',

    // Recovery variations
    'heart rate after 2 min': 'hr2min',
    'hr 2 min': 'hr2min',
    'hr 2min': 'hr2min',
    'heart rate after 5 min': 'hr5min',
    'hr 5 min': 'hr5min',
    'hr 5min': 'hr5min',

    // Other variations
    'mean symmetry first trot': 'symmetry',
    'symmetry': 'symmetry',
    'mean regularity first trot': 'regularity',
    'regularity': 'regularity',
    'time to 120 bpm': 'bpm120',
    '120bpm': 'bpm120',
    '120 bpm': 'bpm120',
    'duration effort zone 5': 'zone5',
    'zone 5': 'zone5',
    'zone5': 'zone5',
    'age': 'age',
    'sex': 'sex',
    'gender': 'sex',
    'temperature': 'temp',
    'temp': 'temp',
    'trotting average heart rate': 'trotHR',
    'trot hr': 'trotHR',
    'walking average hr': 'walkHR',
    'walk hr': 'walkHR'
  };

  // Find column indices
  const colIndices = {};
  headers.forEach((header, index) => {
    const normalizedHeader = header.toLowerCase().trim();
    if (columnMap[normalizedHeader]) {
      colIndices[columnMap[normalizedHeader]] = index;
    }
  });

  console.log('Column indices found:', colIndices);
  console.log('Horse column index:', colIndices['horse']);

  // FALLBACK: If no horse column found, try to detect it
  if (colIndices['horse'] === undefined) {
    console.log('WARNING: Horse column not found by header name. Trying fallback detection...');

    // Look for a column that contains text values that look like horse names
    // (not dates, not numbers, has multiple unique values)
    for (let colIdx = 0; colIdx < headers.length; colIdx++) {
      // Skip columns we already identified
      if (Object.values(colIndices).includes(colIdx)) continue;

      // Sample some values from this column
      const sampleValues = [];
      for (let rowIdx = 1; rowIdx < Math.min(10, csvData.length); rowIdx++) {
        const val = csvData[rowIdx] && csvData[rowIdx][colIdx];
        if (val) sampleValues.push(val.toString().trim());
      }

      // Check if values look like horse names (text, not dates/numbers, varied)
      const uniqueValues = [...new Set(sampleValues)];
      const allText = sampleValues.every(v => {
        // Not a date (no slashes, dashes in date format)
        if (/^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}/.test(v)) return false;
        // Not a pure number
        if (/^[\d.,]+$/.test(v)) return false;
        // Not a time format
        if (/^\d{2}:\d{2}/.test(v)) return false;
        return true;
      });

      if (allText && uniqueValues.length > 1 && sampleValues.length > 0) {
        console.log(`FALLBACK: Using column ${colIdx} ("${headers[colIdx]}") as horse column. Sample values:`, uniqueValues.slice(0, 5));
        colIndices['horse'] = colIdx;
        break;
      }
    }
  }

  if (colIndices['horse'] === undefined) {
    console.error('CRITICAL: Could not find horse column in CSV! Headers were:', headers);
  }

  // Process each row
  for (let i = 1; i < csvData.length; i++) {
    const row = csvData[i];
    if (!row || row.length === 0) continue;

    const getValue = (key) => {
      const idx = colIndices[key];
      return idx !== undefined && row[idx] ? row[idx].toString().trim() : '';
    };

    // Apply transformations
    const type = transformTrainingType(getValue('type'));
    const isRace = type.toLowerCase().includes('race');
    const isWork = type.toLowerCase().includes('work');

    const processedRow = {
      date: transformDate(getValue('date')),
      horse: getValue('horse'),
      type: type,
      track: transformTrackName(getValue('track')),
      surface: transformSurface(getValue('surface')),
      distance: getValue('distance'),
      avgSpeed: getValue('avgSpeed'),
      maxSpeed: getValue('maxSpeed'),
      best1f: getValue('best1f'),
      best2f: getValue('best2f'),
      best3f: getValue('best3f'),
      best4f: getValue('best4f'),
      best5f: getValue('best5f'),
      best6f: getValue('best6f'),
      best7f: getValue('best7f'),
      maxHR: getValue('maxHR'),
      fastRecovery: getValue('fastRecovery'),
      fastQuality: getValue('fastQuality'),
      fastPercent: getValue('fastPercent'),
      recovery15: getValue('recovery15'),
      quality15: getValue('quality15'),
      hr15Percent: getValue('hr15Percent'),
      maxSL: getValue('maxSL'),
      slGallop: getValue('slGallop'),
      sfGallop: getValue('sfGallop'),
      slWork: getValue('slWork'),
      sfWork: getValue('sfWork'),
      hr2min: getValue('hr2min'),
      hr5min: getValue('hr5min'),
      symmetry: getValue('symmetry'),
      regularity: getValue('regularity'),
      bpm120: getValue('bpm120'),
      zone5: getValue('zone5'),
      age: getValue('age'),
      sex: getValue('sex'),
      temp: getValue('temp'),
      distanceCol: getValue('distanceCol'),
      trotHR: getValue('trotHR'),
      walkHR: getValue('walkHR'),
      isRace: isRace,
      isWork: isWork,
      notes: '' // New field for user notes
    };

    if (processedRow.date || processedRow.horse) {
      rows.push(processedRow);
    }
  }

  return rows;
}

// Merge new data with existing, avoiding duplicates
function mergeTrainingData(existingData, newData) {
  const merged = { ...existingData };

  newData.forEach(row => {
    const horseName = row.horse;
    if (!horseName) return;

    if (!merged[horseName]) {
      merged[horseName] = [];
    }

    // Check for duplicate (same horse + same date)
    const isDuplicate = merged[horseName].some(existing =>
      existing.date === row.date && existing.horse === row.horse
    );

    if (!isDuplicate) {
      merged[horseName].push(row);
    }
  });

  // Sort each horse's data by date (most recent first)
  Object.keys(merged).forEach(horseName => {
    merged[horseName].sort((a, b) => {
      const dateA = new Date(a.date);
      const dateB = new Date(b.date);
      return dateB - dateA; // Descending
    });
  });

  return merged;
}

// Generate summary data for main view
function generateHorseSummary(allHorseDetailData, horseMapping = {}) {
  const horseData = [];

  Object.keys(allHorseDetailData).forEach(horseName => {
    const entries = allHorseDetailData[horseName];
    if (!entries || entries.length === 0) return;

    // Get non-race entries for stats
    const workEntries = entries.filter(e => !e.isRace);

    let age = null;
    let best1f = null;
    let best5f = null;
    let dateOfBest5f = null;
    let maxSpeed = null;
    let fastRecovery = null;
    let recovery15min = null;
    let lastWork = null;

    // Find last work date
    if (workEntries.length > 0) {
      lastWork = workEntries[0].date;
    }

    // Collect times and speeds
    const times1f = [];
    const times5f = [];
    const speeds = [];

    workEntries.forEach(entry => {
      if (!age && entry.age) {
        const parsedAge = parseInt(entry.age);
        if (!isNaN(parsedAge)) age = parsedAge;
      }

      if (entry.best1f && isValidTime(entry.best1f)) {
        times1f.push({ time: entry.best1f, seconds: timeToSeconds(entry.best1f) });
      }

      if (entry.best5f && isValidTime(entry.best5f)) {
        times5f.push({
          time: entry.best5f,
          seconds: timeToSeconds(entry.best5f),
          date: entry.date,
          fastRecovery: entry.fastRecovery,
          recovery15min: entry.recovery15
        });
      }

      if (entry.maxSpeed) {
        const speed = parseFloat(entry.maxSpeed);
        if (!isNaN(speed)) speeds.push(speed);
      }
    });

    // Find best times
    if (times1f.length > 0) {
      const fastest1f = times1f.reduce((min, curr) => curr.seconds < min.seconds ? curr : min);
      best1f = fastest1f.time;
    }

    if (times5f.length > 0) {
      const fastest5f = times5f.reduce((min, curr) => curr.seconds < min.seconds ? curr : min);
      best5f = fastest5f.time;
      dateOfBest5f = fastest5f.date;
      fastRecovery = fastest5f.fastRecovery;
      recovery15min = fastest5f.recovery15min;
    }

    if (speeds.length > 0) {
      maxSpeed = Math.max(...speeds);
    }

    // Get owner/country from mapping (case-insensitive lookup)
    const horseNameLower = horseName.toLowerCase();
    const mappingKey = Object.keys(horseMapping).find(k => k.toLowerCase() === horseNameLower);
    const mapping = mappingKey ? horseMapping[mappingKey] : {};

    horseData.push({
      name: horseName,
      age: age,
      lastWork: lastWork,
      best1f: best1f,
      best5f: best5f,
      dateOfBest5f: dateOfBest5f,
      fastRecovery: fastRecovery,
      recovery15min: recovery15min,
      maxSpeed: maxSpeed,
      best5fColor: getBest5FColor(best5f),
      fastRecoveryColor: getFastRecoveryColor(fastRecovery),
      recovery15Color: getRecovery15Color(recovery15min),
      owner: mapping.owner || '',
      country: mapping.country || ''
    });
  });

  // Sort by name
  horseData.sort((a, b) => a.name.localeCompare(b.name));

  return horseData;
}

// ============================================
// HORSE-OWNER-COUNTRY MAPPING STORAGE
// ============================================

async function getHorseMapping() {
  try {
    if (redis) {
      const data = await redis.get('horse-mapping');
      if (data) {
        return typeof data === 'string' ? JSON.parse(data) : data;
      }
    }
    return global.horseMapping || {};
  } catch (error) {
    console.error('Error getting horse mapping:', error);
    return global.horseMapping || {};
  }
}

async function saveHorseMapping(mapping) {
  try {
    if (redis) {
      await redis.set('horse-mapping', JSON.stringify(mapping));
    }
    global.horseMapping = mapping;
    console.log('Horse mapping saved');
  } catch (error) {
    console.error('Error saving horse mapping:', error);
  }
}

// ============================================
// TRAINING ENTRY EDITS STORAGE
// ============================================

async function getTrainingEdits() {
  try {
    if (redis) {
      const data = await redis.get('training-edits');
      if (data) {
        return typeof data === 'string' ? JSON.parse(data) : data;
      }
    }
    return global.trainingEdits || {};
  } catch (error) {
    console.error('Error getting training edits:', error);
    return global.trainingEdits || {};
  }
}

async function saveTrainingEdits(edits) {
  try {
    if (redis) {
      await redis.set('training-edits', JSON.stringify(edits));
    }
    global.trainingEdits = edits;
    console.log('Training edits saved');
  } catch (error) {
    console.error('Error saving training edits:', error);
  }
}

// Apply edits to training data
function applyTrainingEdits(allHorseDetailData, edits) {
  if (!edits || Object.keys(edits).length === 0) return allHorseDetailData;

  const result = { ...allHorseDetailData };

  Object.keys(result).forEach(horseName => {
    result[horseName] = result[horseName].map(entry => {
      // Create unique key for this entry
      const editKey = `${entry.horse}|${entry.date}`;
      const edit = edits[editKey];

      if (edit) {
        return {
          ...entry,
          type: edit.type !== undefined ? edit.type : entry.type,
          track: edit.track !== undefined ? edit.track : entry.track,
          surface: edit.surface !== undefined ? edit.surface : entry.surface,
          notes: edit.notes !== undefined ? edit.notes : (entry.notes || ''),
          isRace: edit.type ? edit.type.toLowerCase().includes('race') : entry.isRace,
          isWork: edit.type ? edit.type.toLowerCase().includes('work') : entry.isWork
        };
      }
      return entry;
    });
  });

  return result;
}

// Routes
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Serve shared view
app.get('/share/:sessionId', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'share.html'));
});

// Upload and process Excel file
app.post('/api/upload', upload.single('excel'), async (req, res) => {
  try {
    console.log('Upload request received');
    if (!req.file) {
      console.log('No file uploaded');
      return res.status(400).json({ error: 'No file uploaded' });
    }

    console.log('Processing file:', req.file.originalname);
    const processedData = processExcelData(req.file.path);
    // Use a fixed session ID so the same link always works
    const sessionId = 'arioneo-main-session';
    console.log('Using fixed session ID:', sessionId);

    // Save using KV storage
    try {
      await saveSession(sessionId, req.file.originalname, processedData.horseData, processedData.allHorseDetailData);
      
      // Clean up uploaded file
      fs.unlinkSync(req.file.path);

      res.json({
        success: true,
        sessionId: sessionId,
        shareUrl: `/share/${sessionId}`,
        data: processedData
      });
    } catch (error) {
      console.error('Error saving session:', error);
      res.status(500).json({ error: 'Failed to save session' });
    }
  } catch (error) {
    console.error('Error processing file:', error);
    res.status(500).json({ error: 'Failed to process file' });
  }
});

// Get shared session data
app.get('/api/session/:sessionId', async (req, res) => {
  const sessionId = req.params.sessionId;
  console.log('Session data requested for:', sessionId);
  
  try {
    const sessionData = await getSession(sessionId);
    
    if (!sessionData) {
      console.log('Session not found:', sessionId);
      return res.status(404).json({ error: 'Session not found' });
    }
    
    console.log('Session found:', sessionId, 'with', sessionData.horseData.length, 'horses');
    
    res.json({
      sessionId: sessionData.id,
      fileName: sessionData.fileName,
      uploadedAt: sessionData.createdAt,
      updatedAt: sessionData.updatedAt,
      horseData: sessionData.horseData,
      allHorseDetailData: sessionData.allHorseDetailData,
      // Include sheet data if available
      allSheets: sessionData.allSheets,
      sheetNames: sessionData.sheetNames,
      currentSheetName: sessionData.currentSheetName
    });
  } catch (error) {
    console.error('Error retrieving session:', error);
    return res.status(500).json({ error: 'Failed to retrieve session' });
  }
});

// Save sheet data for an existing session
app.post('/api/session/:sessionId/sheets', async (req, res) => {
  const sessionId = req.params.sessionId;
  console.log('Sheet data update requested for session:', sessionId);

  try {
    const sheetData = req.body;

    // Validate sheet data
    if (!sheetData || !sheetData.allSheets || !sheetData.sheetNames) {
      return res.status(400).json({ error: 'Invalid sheet data provided' });
    }

    console.log('Updating sheet data with', sheetData.sheetNames.length, 'sheets');

    // Update the session with sheet data
    await updateSessionSheets(sessionId, sheetData);

    res.json({
      success: true,
      message: 'Sheet data updated successfully',
      sheetNames: sheetData.sheetNames,
      currentSheetName: sheetData.currentSheetName
    });
  } catch (error) {
    console.error('Error updating sheet data:', error);
    return res.status(500).json({ error: 'Failed to update sheet data' });
  }
});

// Get latest session (for the main page)
app.get('/api/latest', async (req, res) => {
  console.log('Latest session requested');
  
  try {
    const latestInfo = await getLatestSession();
    
    if (!latestInfo) {
      console.log('No sessions found in storage');
      return res.json({ sessionId: null });
    }
    
    console.log('Latest session found:', latestInfo.sessionId, 'updated at:', latestInfo.updatedAt);
    res.json({ sessionId: latestInfo.sessionId });
  } catch (error) {
    console.error('Error retrieving latest session:', error);
    return res.status(500).json({ error: 'Failed to retrieve latest session' });
  }
});

// Basic health check removed - using enhanced version below

// ============================================
// ARIONEO CSV UPLOAD (Direct from Arioneo - no Excel needed)
// ============================================
app.post('/api/upload/arioneo', upload.single('csv'), async (req, res) => {
  try {
    console.log('Arioneo CSV upload received');
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    console.log('Processing Arioneo CSV:', req.file.originalname);

    // Read and parse CSV/Excel
    const workbook = XLSX.readFile(req.file.path);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const csvData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    // Process the raw Arioneo data (applies all transformations)
    const processedRows = processArioneoCSV(csvData);
    console.log(`Processed ${processedRows.length} training entries`);

    // Debug: Check horse names in processed rows
    const horseNamesInRows = [...new Set(processedRows.map(r => r.horse).filter(Boolean))];
    console.log(`Found ${horseNamesInRows.length} unique horses in CSV:`, horseNamesInRows.slice(0, 10));

    // If no horses found, something went wrong
    if (horseNamesInRows.length === 0) {
      console.error('ERROR: No horse names were extracted from CSV!');
      console.log('First 3 processed rows:', JSON.stringify(processedRows.slice(0, 3), null, 2));
    }

    // Get existing session data
    const sessionId = 'arioneo-main-session';
    let existingSession = await getSession(sessionId);
    let existingDetailData = existingSession?.allHorseDetailData || {};

    // Debug: Show existing data keys
    const existingKeys = Object.keys(existingDetailData);
    console.log('Existing data keys:', existingKeys);

    // Get training edits to apply
    const trainingEdits = await getTrainingEdits();

    // Merge new data with existing
    const mergedDetailData = mergeTrainingData(existingDetailData, processedRows);

    // Debug: Show merged data keys
    const mergedKeys = Object.keys(mergedDetailData);
    console.log('Merged data keys:', mergedKeys);

    // Apply any manual edits
    const finalDetailData = applyTrainingEdits(mergedDetailData, trainingEdits);

    // Get horse mapping for owner/country info
    const horseMapping = await getHorseMapping();

    // Generate summary data
    const horseData = generateHorseSummary(finalDetailData, horseMapping);

    // Save to session
    await saveSession(sessionId, req.file.originalname, horseData, finalDetailData);

    // Clean up uploaded file
    fs.unlinkSync(req.file.path);

    // Count new entries
    const totalEntries = Object.values(finalDetailData).reduce((sum, arr) => sum + arr.length, 0);
    const previousEntries = Object.values(existingDetailData).reduce((sum, arr) => sum + arr.length, 0);
    const newEntries = totalEntries - previousEntries;

    res.json({
      success: true,
      message: `Processed ${processedRows.length} entries, added ${newEntries} new entries`,
      sessionId: sessionId,
      totalHorses: horseData.length,
      totalEntries: totalEntries,
      newEntries: newEntries,
      debug: {
        csvRowCount: csvData.length,
        processedRowCount: processedRows.length,
        uniqueHorsesInCSV: horseNamesInRows.length,
        horseNames: horseNamesInRows.slice(0, 20),
        existingDataKeys: existingKeys.slice(0, 20),
        finalDataKeys: Object.keys(finalDetailData).slice(0, 20)
      },
      data: {
        horseData,
        allHorseDetailData: finalDetailData
      }
    });

  } catch (error) {
    console.error('Error processing Arioneo CSV:', error);
    res.status(500).json({ error: 'Failed to process file: ' + error.message });
  }
});

// ============================================
// TRAINING ENTRY EDIT ENDPOINTS
// ============================================

// Edit a specific training entry
app.put('/api/training/edit', async (req, res) => {
  try {
    const { horse, date, type, track, surface, notes } = req.body;

    if (!horse || !date) {
      return res.status(400).json({ error: 'Horse and date are required' });
    }

    console.log(`Editing training entry: ${horse} on ${date}`);

    // Get existing edits
    const edits = await getTrainingEdits();

    // Create edit key
    const editKey = `${horse}|${date}`;

    // Save the edit
    edits[editKey] = {
      type: type,
      track: track,
      surface: surface,
      notes: notes,
      editedAt: new Date().toISOString()
    };

    await saveTrainingEdits(edits);

    // Update the session data with the edit
    const sessionId = 'arioneo-main-session';
    const session = await getSession(sessionId);

    if (session && session.allHorseDetailData) {
      const updatedDetailData = applyTrainingEdits(session.allHorseDetailData, edits);
      const horseMapping = await getHorseMapping();
      const horseData = generateHorseSummary(updatedDetailData, horseMapping);
      await saveSession(sessionId, session.fileName, horseData, updatedDetailData);
    }

    res.json({
      success: true,
      message: 'Training entry updated',
      editKey: editKey
    });

  } catch (error) {
    console.error('Error editing training entry:', error);
    res.status(500).json({ error: 'Failed to edit training entry' });
  }
});

// Get all edits
app.get('/api/training/edits', async (req, res) => {
  try {
    const edits = await getTrainingEdits();
    res.json({ edits });
  } catch (error) {
    console.error('Error getting training edits:', error);
    res.status(500).json({ error: 'Failed to get training edits' });
  }
});

// ============================================
// HORSE-OWNER-COUNTRY MAPPING ENDPOINTS
// ============================================

// Get all horse mappings
app.get('/api/horses', async (req, res) => {
  try {
    const mapping = await getHorseMapping();
    const horses = Object.entries(mapping).map(([name, data]) => ({
      name,
      owner: data.owner || '',
      country: data.country || '',
      addedAt: data.addedAt || ''
    }));

    // Get unique owners and countries for filter dropdowns
    const owners = [...new Set(horses.map(h => h.owner).filter(Boolean))].sort();
    const countries = [...new Set(horses.map(h => h.country).filter(Boolean))].sort();

    res.json({
      horses,
      owners,
      countries
    });
  } catch (error) {
    console.error('Error getting horse mapping:', error);
    res.status(500).json({ error: 'Failed to get horse mapping' });
  }
});

// Add or update a single horse mapping
app.post('/api/horses', async (req, res) => {
  try {
    const { name, owner, country } = req.body;

    if (!name) {
      return res.status(400).json({ error: 'Horse name is required' });
    }

    const mapping = await getHorseMapping();

    mapping[name] = {
      owner: owner || '',
      country: country || '',
      addedAt: mapping[name]?.addedAt || new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };

    await saveHorseMapping(mapping);

    // Update session data with new mapping
    const sessionId = 'arioneo-main-session';
    const session = await getSession(sessionId);
    if (session && session.allHorseDetailData) {
      const horseData = generateHorseSummary(session.allHorseDetailData, mapping);
      await saveSession(sessionId, session.fileName, horseData, session.allHorseDetailData);
    }

    res.json({
      success: true,
      message: `Horse "${name}" mapping saved`,
      horse: { name, owner, country }
    });

  } catch (error) {
    console.error('Error saving horse mapping:', error);
    res.status(500).json({ error: 'Failed to save horse mapping' });
  }
});

// Bulk import horse mappings from CSV
app.post('/api/horses/import', upload.single('csv'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    console.log('Importing horse mappings from:', req.file.originalname);

    // Read CSV/Excel
    const workbook = XLSX.readFile(req.file.path);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    // Get existing mapping
    const mapping = await getHorseMapping();

    // Parse headers (first row)
    const headers = data[0].map(h => h ? h.toString().toLowerCase().trim() : '');
    const nameIdx = headers.findIndex(h => h.includes('horse') || h.includes('name'));
    const ownerIdx = headers.findIndex(h => h.includes('owner'));
    const countryIdx = headers.findIndex(h => h.includes('country'));

    if (nameIdx === -1) {
      fs.unlinkSync(req.file.path);
      return res.status(400).json({ error: 'Could not find "Horse Name" column in file' });
    }

    let imported = 0;
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const name = row[nameIdx] ? row[nameIdx].toString().trim() : '';

      if (name) {
        mapping[name] = {
          owner: ownerIdx !== -1 && row[ownerIdx] ? row[ownerIdx].toString().trim() : (mapping[name]?.owner || ''),
          country: countryIdx !== -1 && row[countryIdx] ? row[countryIdx].toString().trim() : (mapping[name]?.country || ''),
          addedAt: mapping[name]?.addedAt || new Date().toISOString(),
          updatedAt: new Date().toISOString()
        };
        imported++;
      }
    }

    await saveHorseMapping(mapping);

    // Clean up
    fs.unlinkSync(req.file.path);

    // Update session data
    const sessionId = 'arioneo-main-session';
    const session = await getSession(sessionId);
    if (session && session.allHorseDetailData) {
      const horseData = generateHorseSummary(session.allHorseDetailData, mapping);
      await saveSession(sessionId, session.fileName, horseData, session.allHorseDetailData);
    }

    res.json({
      success: true,
      message: `Imported ${imported} horse mappings`,
      totalHorses: Object.keys(mapping).length
    });

  } catch (error) {
    console.error('Error importing horse mappings:', error);
    res.status(500).json({ error: 'Failed to import horse mappings' });
  }
});

// Delete a horse mapping
app.delete('/api/horses/:name', async (req, res) => {
  try {
    const horseName = decodeURIComponent(req.params.name);

    const mapping = await getHorseMapping();

    if (!mapping[horseName]) {
      return res.status(404).json({ error: 'Horse not found' });
    }

    delete mapping[horseName];
    await saveHorseMapping(mapping);

    res.json({
      success: true,
      message: `Horse "${horseName}" mapping deleted`
    });

  } catch (error) {
    console.error('Error deleting horse mapping:', error);
    res.status(500).json({ error: 'Failed to delete horse mapping' });
  }
});

// Get list of owners
app.get('/api/owners', async (req, res) => {
  try {
    const mapping = await getHorseMapping();
    const owners = [...new Set(Object.values(mapping).map(h => h.owner).filter(Boolean))].sort();
    res.json({ owners });
  } catch (error) {
    console.error('Error getting owners:', error);
    res.status(500).json({ error: 'Failed to get owners' });
  }
});

// Get list of countries
app.get('/api/countries', async (req, res) => {
  try {
    const mapping = await getHorseMapping();
    const countries = [...new Set(Object.values(mapping).map(h => h.country).filter(Boolean))].sort();
    res.json({ countries });
  } catch (error) {
    console.error('Error getting countries:', error);
    res.status(500).json({ error: 'Failed to get countries' });
  }
});

// Clear all session data (to remove stale data)
app.delete('/api/session/clear', async (req, res) => {
  try {
    const sessionId = 'arioneo-main-session';

    // Clear from Redis
    if (redis) {
      await redis.del(`session:${sessionId}`);
      await redis.del('latest_session');
      console.log('Session cleared from Redis');
    }

    // Clear from memory
    global.sessionStorage.delete(sessionId);
    global.latestSession = null;

    res.json({
      success: true,
      message: 'All session data cleared. You can now upload fresh data.'
    });
  } catch (error) {
    console.error('Error clearing session:', error);
    res.status(500).json({ error: 'Failed to clear session data' });
  }
});

// Debug endpoint to see current data structure
app.get('/api/debug/data', async (req, res) => {
  try {
    const sessionId = 'arioneo-main-session';
    const session = await getSession(sessionId);

    if (!session) {
      return res.json({ hasData: false, message: 'No session data found' });
    }

    // Get keys from allHorseDetailData to see what the grouping looks like
    const horseKeys = Object.keys(session.allHorseDetailData || {});
    const sampleData = {};

    // Get first entry from each horse for debugging
    horseKeys.slice(0, 5).forEach(key => {
      const entries = session.allHorseDetailData[key];
      sampleData[key] = {
        entryCount: entries ? entries.length : 0,
        firstEntry: entries && entries[0] ? {
          date: entries[0].date,
          horse: entries[0].horse,
          type: entries[0].type,
          track: entries[0].track
        } : null
      };
    });

    res.json({
      hasData: true,
      sessionId: session.id,
      fileName: session.fileName,
      horseDataCount: session.horseData ? session.horseData.length : 0,
      horseDetailKeys: horseKeys,
      horseDetailKeyCount: horseKeys.length,
      sampleData: sampleData
    });
  } catch (error) {
    console.error('Error in debug endpoint:', error);
    res.status(500).json({ error: error.message });
  }
});

// Debug route
app.get('/api/debug', (req, res) => {
  res.json({ 
    message: 'Server is working', 
    headers: req.headers,
    protocol: req.protocol,
    secure: req.secure,
    host: req.get('host'),
    allEnvVars: Object.keys(process.env).filter(key => key.includes('KV') || key.includes('REDIS') || key.includes('UPSTASH')),
    specificEnvs: {
      KV_REST_API_URL: process.env.KV_REST_API_URL ? 'exists' : 'missing',
      KV_REST_API_TOKEN: process.env.KV_REST_API_TOKEN ? 'exists' : 'missing'
    }
  });
});

// Error handling middleware
app.use((error, req, res, next) => {
  console.error('Error:', error);
  res.status(500).json({ error: 'Internal server error' });
});

// Start server (only in development)
if (process.env.NODE_ENV !== 'production') {
  app.listen(port, () => {
    console.log(`Arioneo US server running on port ${port}`);
    console.log(`Upload interface: http://localhost:${port}`);
  });

  // Graceful shutdown
  process.on('SIGTERM', () => {
    console.log('Received SIGTERM, shutting down gracefully');
    process.exit(0);
  });
}

// Initialize global storage
console.log('Initializing global session storage...');

// Auto-restore from Redis on startup
async function autoRestoreFromRedis() {
  try {
    if (redis) {
      // Try to restore main session from Redis
      const sessionData = await redis.get('session:arioneo-main-session');
      if (sessionData) {
        const parsed = JSON.parse(sessionData);
        global.sessionStorage.set(parsed.id, parsed);
        console.log('Auto-restored session from Redis');
      }
      
      // Try to restore latest session info
      const latestSession = await redis.get('latest_session');
      if (latestSession) {
        global.latestSession = JSON.parse(latestSession);
        console.log('Auto-restored latest session info from Redis');
      }
    }
  } catch (error) {
    console.error('Error auto-restoring from Redis:', error);
  }
}

// Run auto-restore on startup
autoRestoreFromRedis();

// Add an API endpoint to manually restore data
app.get('/api/restore', async (req, res) => {
  try {
    await autoRestoreFromRedis();
    const sessionData = await getSession('arioneo-main-session');
    const hasData = sessionData !== null;
    
    res.json({ 
      restored: hasData, 
      sessionId: hasData ? sessionData.id : null,
      message: hasData ? 'Data restored successfully' : 'No data found to restore'
    });
  } catch (error) {
    console.error('Error in restore endpoint:', error);
    res.status(500).json({ error: 'Failed to restore data' });
  }
});

// Add a health check that also checks data
app.get('/api/health', async (req, res) => {
  try {
    const sessionData = await getSession('arioneo-main-session');
    const hasData = sessionData !== null;
    
    res.json({ 
      status: 'OK', 
      timestamp: new Date().toISOString(),
      dataLoaded: hasData,
      sessionId: hasData ? sessionData.id : null,
      redisConnected: redis !== null,
      envVars: {
        hasKvUrl: !!process.env.KV_REST_API_URL,
        hasKvToken: !!process.env.KV_REST_API_TOKEN
      }
    });
  } catch (error) {
    console.error('Health check error:', error);
    res.status(500).json({ error: 'Health check failed', details: error.message });
  }
});

// Debug endpoint to test Redis directly
app.get('/api/redis-test', async (req, res) => {
  try {
    if (!redis) {
      return res.json({ error: 'Redis not connected' });
    }
    
    const testResult = await redis.get('session:arioneo-main-session');
    
    // Test what getSession returns
    const sessionResult = await getSession('arioneo-main-session');
    
    res.json({ 
      keyExists: !!testResult,
      dataType: typeof testResult,
      rawData: testResult,
      getSessionResult: sessionResult ? 'success' : 'failed',
      getSessionData: sessionResult ? { id: sessionResult.id, fileName: sessionResult.fileName } : null
    });
  } catch (error) {
    res.json({ error: error.message });
  }
});

// Export for Vercel
module.exports = app;
