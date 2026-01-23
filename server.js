const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const cors = require('cors');
const { v4: uuidv4 } = require('uuid');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');
const compression = require('compression');
const helmet = require('helmet');
const cookieParser = require('cookie-parser');
const pdfParse = require('pdf-parse');
const Fuse = require('fuse.js');

// Auth credentials from environment variables
const AUTH_USERNAME = process.env.AUTH_USERNAME || 'admin';
const AUTH_PASSWORD = process.env.AUTH_PASSWORD || 'password';
const AUTH_SECRET = process.env.SESSION_SECRET || 'horse-training-secret-key-2024';

// Generate auth token from credentials
function generateAuthToken() {
  return crypto.createHmac('sha256', AUTH_SECRET)
    .update(AUTH_USERNAME + AUTH_PASSWORD)
    .digest('hex');
}

const VALID_AUTH_TOKEN = generateAuthToken();

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
app.use(cookieParser());

// Auth middleware - protect all routes except login
function requireAuth(req, res, next) {
  // Allow access to login page and login POST
  if (req.path === '/login' || req.path === '/login.html') {
    return next();
  }

  // Check if user is authenticated via cookie
  const authToken = req.cookies.auth_token;
  if (authToken && authToken === VALID_AUTH_TOKEN) {
    return next();
  }

  // Redirect to login for HTML requests, return 401 for API requests
  if (req.path.startsWith('/api/')) {
    return res.status(401).json({ error: 'Unauthorized' });
  }

  return res.redirect('/login.html');
}

// Login route
app.post('/login', (req, res) => {
  const { username, password } = req.body;

  if (username === AUTH_USERNAME && password === AUTH_PASSWORD) {
    // Set auth cookie (7 days)
    res.cookie('auth_token', VALID_AUTH_TOKEN, {
      httpOnly: true,
      secure: process.env.NODE_ENV === 'production',
      sameSite: 'lax',
      maxAge: 3 * 24 * 60 * 60 * 1000
    });
    return res.redirect('/');
  }

  return res.redirect('/login.html?error=1');
});

// Logout route
app.get('/logout', (req, res) => {
  res.clearCookie('auth_token');
  res.redirect('/login.html');
});

// Auth status endpoint
app.get('/api/auth/status', (req, res) => {
  const authToken = req.cookies.auth_token;
  res.json({ authenticated: authToken === VALID_AUTH_TOKEN });
});

// Apply auth middleware to all routes
app.use(requireAuth);

// Serve static files (use absolute path for Vercel compatibility)
app.use(express.static(path.join(__dirname, 'public')));

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

// Configure multer for PDF uploads (race charts)
const pdfUpload = multer({
  storage: storage,
  limits: {
    fileSize: 50 * 1024 * 1024, // 50MB limit
  },
  fileFilter: function (req, file, cb) {
    const ext = path.extname(file.originalname).toLowerCase();
    if (ext === '.pdf') {
      cb(null, true);
    } else {
      cb(new Error('Only PDF files are allowed'));
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
      // Skip default Excel sheet names and other non-horse sheets
      const invalidSheetNames = ['sheet1', 'sheet2', 'sheet3', 'buttons', 'worksheet', 'data', 'default'];
      if (invalidSheetNames.includes(sheetNameLower) || sheetNameLower.startsWith('sheet')) {
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

// Format horse name for display: "HORSENAME YY"
// Handles: "2022 Stoweshoe" -> "STOWESHOE 22", "23' GINGER PUNCH" -> "GINGER PUNCH 23"
function formatHorseNameForDisplay(name) {
  if (!name) return '';

  let horseName = name.toString().trim();
  let year = null;

  // Strip any non-alphanumeric characters from start (quotes, spaces, etc)
  horseName = horseName.replace(/^[^a-zA-Z0-9]+/, '');

  // Pattern 1: 4-digit year at start (2021-2099)
  const fourDigitYearStart = horseName.match(/^(20[2-9]\d)\s*(.+)$/);
  if (fourDigitYearStart) {
    year = fourDigitYearStart[1].slice(-2);
    horseName = fourDigitYearStart[2];
  }

  // Pattern 2: 2-digit year at start (21-99)
  if (!year) {
    const twoDigitYearStart = horseName.match(/^([2-9]\d)[^a-zA-Z0-9]*(.+)$/);
    if (twoDigitYearStart) {
      year = twoDigitYearStart[1];
      horseName = twoDigitYearStart[2];
    }
  }

  // Pattern 3: Year at end
  if (!year) {
    const yearAtEnd = horseName.match(/^(.+?)[^a-zA-Z0-9]*([2-9]\d)[^a-zA-Z0-9]*$/);
    if (yearAtEnd && yearAtEnd[1].length > 2) {
      horseName = yearAtEnd[1];
      year = yearAtEnd[2];
    }
  }

  // Final cleanup - remove any non-letter characters from start/end, then uppercase
  horseName = horseName.replace(/^[^a-zA-Z]+|[^a-zA-Z]+$/g, '').trim().toUpperCase();

  if (year) {
    return `${horseName} ${year}`;
  }

  return horseName;
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

  // Invalid horse names to filter out
  const invalidNames = ['worksheet', 'sheet', 'sheet1', 'sheet2', 'sheet3', 'data', 'horse', 'horse name', 'name'];

  newData.forEach(row => {
    const horseName = row.horse;
    if (!horseName) return;

    // Skip invalid/placeholder horse names
    if (invalidNames.includes(horseName.toLowerCase().trim())) return;

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

  // Invalid horse names to filter out
  const invalidNames = ['worksheet', 'sheet', 'sheet1', 'sheet2', 'sheet3', 'data', 'horse', 'horse name', 'name'];

  Object.keys(allHorseDetailData).forEach(horseName => {
    const entries = allHorseDetailData[horseName];
    if (!entries || entries.length === 0) return;

    // Skip invalid/placeholder horse names
    if (invalidNames.includes(horseName.toLowerCase().trim())) return;

    // Sort entries by date (most recent first) - entries should already be sorted but ensure it
    const sortedEntries = [...entries].sort((a, b) => {
      const dateA = new Date(a.date);
      const dateB = new Date(b.date);
      return dateB - dateA;
    });

    // Get the most recent TRAINING entry (exclude races and notes for main table display)
    // Races have isRace: true flag, notes have isNote: true flag
    const trainingEntries = sortedEntries.filter(e => !e.isRace && !e.isNote);
    const lastTraining = trainingEntries.length > 0 ? trainingEntries[0] : null;

    // Debug: log first horse's date info
    if (sortedEntries.length > 0 && Object.keys(allHorseDetailData).indexOf(horseName) < 3) {
      console.log(`Horse ${horseName}: first entry date = "${lastTraining?.date}", entries count = ${sortedEntries.length}`);
    }

    let age = null;
    let lastTrainingDate = null;
    let best1f = null;
    let best5f = null;
    let fastRecovery = null;
    let recovery15min = null;

    // Get data from most recent training
    if (lastTraining) {
      lastTrainingDate = lastTraining.date;

      if (lastTraining.age) {
        const parsedAge = parseInt(lastTraining.age);
        if (!isNaN(parsedAge)) age = parsedAge;
      }

      if (lastTraining.best1f && isValidTime(lastTraining.best1f)) {
        best1f = lastTraining.best1f;
      }

      if (lastTraining.best5f && isValidTime(lastTraining.best5f)) {
        best5f = lastTraining.best5f;
      }

      fastRecovery = lastTraining.fastRecovery || null;
      recovery15min = lastTraining.recovery15 || null;
    }

    // If no age from last training, search through other entries
    if (!age) {
      for (const entry of sortedEntries) {
        if (entry.age) {
          const parsedAge = parseInt(entry.age);
          if (!isNaN(parsedAge)) {
            age = parsedAge;
            break;
          }
        }
      }
    }

    // Get owner/country/historic from mapping (case-insensitive + fuzzy lookup)
    const horseNameLower = horseName.toLowerCase();
    const horseNameStripped = horseName.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
    let mappingKey = Object.keys(horseMapping).find(k => k.toLowerCase() === horseNameLower);
    // If no exact match, try matching after stripping special characters
    if (!mappingKey) {
      mappingKey = Object.keys(horseMapping).find(k =>
        k.replace(/[^a-zA-Z0-9]/g, '').toLowerCase() === horseNameStripped
      );
    }
    const mapping = mappingKey ? horseMapping[mappingKey] : {};

    horseData.push({
      name: horseName,
      displayName: formatHorseNameForDisplay(horseName),
      age: age,
      lastTrainingDate: lastTrainingDate,
      best1f: best1f,
      best5f: best5f,
      fastRecovery: fastRecovery,
      recovery15min: recovery15min,
      best5fColor: getBest5FColor(best5f),
      fastRecoveryColor: getFastRecoveryColor(fastRecovery),
      recovery15Color: getRecovery15Color(recovery15min),
      owner: mapping.owner || '',
      country: mapping.country || '',
      isHistoric: mapping.isHistoric || false
    });
  });

  // Sort by display name
  horseData.sort((a, b) => a.displayName.localeCompare(b.displayName));

  return horseData;
}

// Ensure all horses from training data exist in the mapping
async function ensureHorsesInMapping(allHorseDetailData) {
  const mapping = await getHorseMapping();
  let addedCount = 0;

  // Invalid horse names to skip
  const invalidNames = ['worksheet', 'sheet', 'sheet1', 'sheet2', 'sheet3', 'data', 'horse', 'horse name', 'name'];

  for (const horseName of Object.keys(allHorseDetailData)) {
    // Skip invalid names
    if (invalidNames.includes(horseName.toLowerCase().trim())) continue;

    // Check if horse exists in mapping (exact or fuzzy match)
    const nameLower = horseName.toLowerCase();
    const nameStripped = horseName.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
    const nameNormalized = normalizeHorseNameForMatch(horseName);

    const existingKey = Object.keys(mapping).find(k =>
      k.toLowerCase() === nameLower ||
      k.replace(/[^a-zA-Z0-9]/g, '').toLowerCase() === nameStripped ||
      normalizeHorseNameForMatch(k) === nameNormalized
    );

    // Also check if it's an alias
    let isAlias = false;
    for (const [primaryName, data] of Object.entries(mapping)) {
      if (data.aliases && data.aliases.some(alias =>
        alias.toLowerCase() === nameLower ||
        alias.replace(/[^a-zA-Z0-9]/g, '').toLowerCase() === nameStripped ||
        normalizeHorseNameForMatch(alias) === nameNormalized
      )) {
        isAlias = true;
        break;
      }
    }

    // If not found, add to mapping
    if (!existingKey && !isAlias) {
      mapping[horseName] = {
        owner: '',
        country: '',
        isHistoric: false,
        aliases: [],
        addedAt: new Date().toISOString(),
        autoAdded: true
      };
      addedCount++;
    }
  }

  if (addedCount > 0) {
    await saveHorseMapping(mapping);
    console.log(`Auto-added ${addedCount} new horses to mapping`);
  }

  return mapping;
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

// Normalize horse name for comparison (handles year format differences)
// "2022 Ginger Punch" and "GINGER PUNCH 22" should both normalize similarly
function normalizeHorseNameForMatch(name) {
  if (!name) return '';

  // Extract letters and numbers separately
  const letters = name.replace(/[^a-zA-Z]/g, '').toLowerCase();

  // Extract year - could be 4-digit (2022) or 2-digit (22)
  const fourDigitYear = name.match(/\b(19|20)\d{2}\b/);
  const twoDigitYear = name.match(/\b(\d{2})\b/);

  let year = '';
  if (fourDigitYear) {
    year = fourDigitYear[0].slice(-2); // Get last 2 digits
  } else if (twoDigitYear) {
    year = twoDigitYear[1];
  }

  return letters + year;
}

// Resolve a horse name to its primary name (if it's an alias)
function resolveHorseAlias(horseName, horseMapping) {
  if (!horseName || !horseMapping) return horseName;

  const nameLower = horseName.toLowerCase().trim();
  const nameStripped = horseName.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
  const nameNormalized = normalizeHorseNameForMatch(horseName);

  // Check if this name is a primary name (exact or fuzzy match)
  const directMatch = Object.keys(horseMapping).find(k =>
    k.toLowerCase() === nameLower ||
    k.replace(/[^a-zA-Z0-9]/g, '').toLowerCase() === nameStripped ||
    normalizeHorseNameForMatch(k) === nameNormalized
  );
  if (directMatch) return directMatch;

  // Check if this name is an alias of another horse
  for (const [primaryName, data] of Object.entries(horseMapping)) {
    if (data.aliases && Array.isArray(data.aliases)) {
      const aliasMatch = data.aliases.find(alias =>
        alias.toLowerCase() === nameLower ||
        alias.replace(/[^a-zA-Z0-9]/g, '').toLowerCase() === nameStripped ||
        normalizeHorseNameForMatch(alias) === nameNormalized
      );
      if (aliasMatch) {
        console.log(`Resolved alias: "${horseName}" -> "${primaryName}" (matched: "${aliasMatch}")`);
        return primaryName;
      }
    }
  }

  return horseName;
}

// Merge training data for horses with aliases
function mergeAliasedHorseData(allHorseDetailData, horseMapping) {
  const merged = {};

  // Log what aliases we have
  const allAliases = [];
  Object.entries(horseMapping).forEach(([name, data]) => {
    if (data.aliases && data.aliases.length > 0) {
      allAliases.push({ primary: name, aliases: data.aliases });
    }
  });
  if (allAliases.length > 0) {
    console.log('Horse aliases configured:', JSON.stringify(allAliases));
  }

  const dataKeys = Object.keys(allHorseDetailData);
  console.log('Training data horse names:', dataKeys.slice(0, 10).join(', ') + (dataKeys.length > 10 ? '...' : ''));

  Object.keys(allHorseDetailData).forEach(horseName => {
    const primaryName = resolveHorseAlias(horseName, horseMapping);

    if (primaryName !== horseName) {
      console.log(`Merging: "${horseName}" -> "${primaryName}"`);
    }

    if (!merged[primaryName]) {
      merged[primaryName] = [];
    }

    // Add all entries, updating the horse name to primary
    const entries = allHorseDetailData[horseName].map(entry => ({
      ...entry,
      horse: primaryName,
      originalName: horseName !== primaryName ? horseName : undefined
    }));

    merged[primaryName].push(...entries);
  });

  // Sort each horse's entries by date (most recent first)
  Object.keys(merged).forEach(horseName => {
    merged[horseName].sort((a, b) => {
      const dateA = new Date(a.date);
      const dateB = new Date(b.date);
      return dateB - dateA;
    });
  });

  console.log(`Merge complete: ${dataKeys.length} raw names -> ${Object.keys(merged).length} merged names`);

  return merged;
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

// ============================================
// HORSE NOTES STORAGE
// ============================================

async function getHorseNotes() {
  try {
    if (redis) {
      const data = await redis.get('horse-notes');
      if (data) {
        return typeof data === 'string' ? JSON.parse(data) : data;
      }
    }
    return global.horseNotes || {};
  } catch (error) {
    console.error('Error getting horse notes:', error);
    return global.horseNotes || {};
  }
}

async function saveHorseNotes(notes) {
  try {
    if (redis) {
      await redis.set('horse-notes', JSON.stringify(notes));
    }
    global.horseNotes = notes;
    console.log('Horse notes saved');
  } catch (error) {
    console.error('Error saving horse notes:', error);
  }
}

// Apply notes to training data (adds note entries to each horse's data)
function applyHorseNotes(allHorseDetailData, notes) {
  // First, remove any existing notes from the data (they may have been incorrectly saved before)
  const result = {};
  Object.keys(allHorseDetailData).forEach(horseName => {
    result[horseName] = (allHorseDetailData[horseName] || []).filter(entry => !entry.isNote);
  });

  if (!notes || Object.keys(notes).length === 0) return result;

  Object.keys(notes).forEach(horseName => {
    const horseNotes = notes[horseName] || [];
    if (horseNotes.length > 0) {
      if (!result[horseName]) {
        result[horseName] = [];
      }
      // Add each note as a separate entry
      horseNotes.forEach(note => {
        result[horseName].push({
          date: note.date,
          horse: horseName,
          type: 'Note',
          notes: note.note,
          isNote: true,
          track: '-', surface: '-', distance: '-', avgSpeed: '-', maxSpeed: '-',
          best1f: '-', best2f: '-', best3f: '-', best4f: '-', best5f: '-',
          best6f: '-', best7f: '-', maxHR: '-', fastRecovery: '-', fastQuality: '-',
          fastPercent: '-', recovery15: '-', quality15: '-', hr15Percent: '-',
          maxSL: '-', slGallop: '-', sfGallop: '-', slWork: '-', sfWork: '-',
          hr2min: '-', hr5min: '-', symmetry: '-', regularity: '-', bpm120: '-',
          zone5: '-', age: '-', sex: '-', temp: '-', distanceCol: '-',
          trotHR: '-', walkHR: '-'
        });
      });
    }
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

    // IMPORTANT: Preserve existing race data before overwriting
    // Get existing session to extract race entries
    const existingSession = await getSession(sessionId);
    if (existingSession && existingSession.allHorseDetailData) {
      console.log('Preserving existing race data during CSV upload');

      // Extract race entries from existing data and merge into new data
      for (const [existingHorseName, entries] of Object.entries(existingSession.allHorseDetailData)) {
        if (!entries || !Array.isArray(entries)) continue;

        // Filter for race entries only
        const raceEntries = entries.filter(e => e.isRace === true);
        if (raceEntries.length === 0) continue;

        console.log(`Found ${raceEntries.length} race(s) for horse: ${existingHorseName}`);

        // Find matching horse in new data (case-insensitive)
        const newHorseKeys = Object.keys(processedData.allHorseDetailData);
        const matchingKey = newHorseKeys.find(k => k.toLowerCase() === existingHorseName.toLowerCase());

        if (matchingKey) {
          // Merge race entries into existing horse data
          processedData.allHorseDetailData[matchingKey].push(...raceEntries);
          // Sort by date (newest first)
          processedData.allHorseDetailData[matchingKey].sort((a, b) => {
            const dateA = new Date(a.date);
            const dateB = new Date(b.date);
            return dateB - dateA;
          });
          console.log(`Merged races into existing horse: ${matchingKey}`);
        } else {
          // Horse not in new CSV - preserve their race data under original name
          processedData.allHorseDetailData[existingHorseName] = raceEntries;
          console.log(`Preserved races for horse not in CSV: ${existingHorseName}`);
        }
      }
    }

    // Ensure all horses from the upload are in the mapping
    await ensureHorsesInMapping(processedData.allHorseDetailData);

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

    // Apply alias merging to combine data for horses with aliases
    const horseMapping = await getHorseMapping();
    const mergedDetailData = mergeAliasedHorseData(sessionData.allHorseDetailData, horseMapping);

    // Apply horse notes to the merged data
    const horseNotes = await getHorseNotes();
    const dataWithNotes = applyHorseNotes(mergedDetailData, horseNotes);

    const mergedHorseData = generateHorseSummary(dataWithNotes, horseMapping);

    console.log('Session found:', sessionId, 'with', mergedHorseData.length, 'horses (after alias merge and notes)');

    res.json({
      sessionId: sessionData.id,
      fileName: sessionData.fileName,
      uploadedAt: sessionData.createdAt,
      updatedAt: sessionData.updatedAt,
      horseData: mergedHorseData,
      allHorseDetailData: dataWithNotes,
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
    const editedDetailData = applyTrainingEdits(mergedDetailData, trainingEdits);

    // Ensure all horses from the data are in the mapping
    await ensureHorsesInMapping(editedDetailData);

    // Get horse mapping for owner/country info (refresh after ensuring horses)
    const horseMapping = await getHorseMapping();

    // Apply alias merging to combine data for horses with aliases (for display)
    const aliasedDetailData = mergeAliasedHorseData(editedDetailData, horseMapping);

    // Apply horse notes for display (but don't save notes to session - they're stored separately)
    const horseNotes = await getHorseNotes();
    const dataWithNotes = applyHorseNotes(aliasedDetailData, horseNotes);

    // Generate summary data (with notes for display)
    const horseData = generateHorseSummary(dataWithNotes, horseMapping);

    // Save to session WITHOUT notes (notes are stored separately in Redis)
    await saveSession(sessionId, req.file.originalname, horseData, editedDetailData);

    // Clean up uploaded file
    fs.unlinkSync(req.file.path);

    // Count new entries (use editedDetailData to not count notes)
    const totalEntries = Object.values(editedDetailData).reduce((sum, arr) => sum + arr.length, 0);
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
        finalDataKeys: Object.keys(aliasedDetailData).slice(0, 20)
      },
      data: {
        horseData,
        allHorseDetailData: dataWithNotes
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
      // Preserve existing sheet data when saving
      await saveSession(sessionId, session.fileName, horseData, updatedDetailData, {
        allSheets: session.allSheets,
        sheetNames: session.sheetNames,
        currentSheetName: session.currentSheetName
      });
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

// Delete a training entry
app.delete('/api/training/delete', async (req, res) => {
  try {
    const { horse, date } = req.body;

    if (!horse || !date) {
      return res.status(400).json({ error: 'Horse and date are required' });
    }

    console.log(`Deleting training entry: ${horse} on ${date}`);

    const sessionId = 'arioneo-main-session';
    const session = await getSession(sessionId);

    if (!session || !session.allHorseDetailData) {
      return res.status(404).json({ error: 'Session not found' });
    }

    // Find the horse in detail data (case-insensitive)
    const horseKeys = Object.keys(session.allHorseDetailData);
    const matchingKey = horseKeys.find(k => k.toLowerCase() === horse.toLowerCase());

    if (!matchingKey || !session.allHorseDetailData[matchingKey]) {
      return res.status(404).json({ error: 'Horse not found' });
    }

    // Find and remove the entry with the matching date
    const entries = session.allHorseDetailData[matchingKey];
    const originalLength = entries.length;
    session.allHorseDetailData[matchingKey] = entries.filter(entry => entry.date !== date);

    if (session.allHorseDetailData[matchingKey].length === originalLength) {
      return res.status(404).json({ error: 'Training entry not found' });
    }

    // If horse has no more entries, remove the horse key
    if (session.allHorseDetailData[matchingKey].length === 0) {
      delete session.allHorseDetailData[matchingKey];
    }

    // Regenerate summary and save
    const horseMapping = await getHorseMapping();
    const horseData = generateHorseSummary(session.allHorseDetailData, horseMapping);
    await saveSession(sessionId, session.fileName, horseData, session.allHorseDetailData, {
      allSheets: session.allSheets,
      sheetNames: session.sheetNames,
      currentSheetName: session.currentSheetName
    });

    res.json({
      success: true,
      message: 'Training entry deleted'
    });

  } catch (error) {
    console.error('Error deleting training entry:', error);
    res.status(500).json({ error: 'Failed to delete training entry' });
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
      displayName: formatHorseNameForDisplay(name),
      owner: data.owner || '',
      country: data.country || '',
      isHistoric: data.isHistoric || false,
      aliases: data.aliases || [],
      addedAt: data.addedAt || ''
    }));

    // Sort horses alphabetically by display name
    horses.sort((a, b) => (a.displayName || a.name).localeCompare(b.displayName || b.name));

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
    const { name, owner, country, isHistoric } = req.body;

    if (!name) {
      return res.status(400).json({ error: 'Horse name is required' });
    }

    const mapping = await getHorseMapping();

    // Check if there's an existing mapping with similar name (handles special char variations)
    const nameLower = name.toLowerCase();
    const nameStripped = name.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
    let existingKey = Object.keys(mapping).find(k => k.toLowerCase() === nameLower);
    if (!existingKey) {
      existingKey = Object.keys(mapping).find(k =>
        k.replace(/[^a-zA-Z0-9]/g, '').toLowerCase() === nameStripped
      );
    }

    // Use existing key if found, otherwise use the provided name
    const keyToUse = existingKey || name;

    mapping[keyToUse] = {
      owner: owner || '',
      country: country || '',
      isHistoric: isHistoric || false,
      addedAt: mapping[keyToUse]?.addedAt || new Date().toISOString(),
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
      horse: { name, owner, country, isHistoric }
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

// Merge horses - combine multiple horse names into one with aliases
app.post('/api/horses/merge', async (req, res) => {
  try {
    const { primaryName, aliasNames } = req.body;

    if (!primaryName || !aliasNames || !Array.isArray(aliasNames) || aliasNames.length === 0) {
      return res.status(400).json({ error: 'Primary name and alias names are required' });
    }

    const mapping = await getHorseMapping();

    // Ensure primary horse exists in mapping
    if (!mapping[primaryName]) {
      mapping[primaryName] = {
        owner: '',
        country: '',
        isHistoric: false,
        aliases: [],
        addedAt: new Date().toISOString()
      };
    }

    // Initialize aliases array if not exists
    if (!mapping[primaryName].aliases) {
      mapping[primaryName].aliases = [];
    }

    // Add each alias
    const addedAliases = [];
    for (const aliasName of aliasNames) {
      if (aliasName === primaryName) continue; // Skip if same as primary

      // Check if alias is already used
      const existingPrimary = Object.entries(mapping).find(([name, data]) =>
        name !== primaryName && data.aliases && data.aliases.includes(aliasName)
      );
      if (existingPrimary) {
        console.log(`Alias "${aliasName}" already belongs to "${existingPrimary[0]}"`);
        continue;
      }

      // If alias was a primary horse, copy its data and remove it
      if (mapping[aliasName]) {
        // Copy owner/country if primary doesn't have them
        if (!mapping[primaryName].owner && mapping[aliasName].owner) {
          mapping[primaryName].owner = mapping[aliasName].owner;
        }
        if (!mapping[primaryName].country && mapping[aliasName].country) {
          mapping[primaryName].country = mapping[aliasName].country;
        }
        // Copy any existing aliases from the merged horse
        if (mapping[aliasName].aliases) {
          mapping[primaryName].aliases.push(...mapping[aliasName].aliases);
        }
        delete mapping[aliasName];
      }

      // Add to aliases if not already there
      if (!mapping[primaryName].aliases.includes(aliasName)) {
        mapping[primaryName].aliases.push(aliasName);
        addedAliases.push(aliasName);
      }
    }

    mapping[primaryName].updatedAt = new Date().toISOString();
    await saveHorseMapping(mapping);

    console.log(`Merged horses: "${primaryName}" now includes aliases: ${addedAliases.join(', ')}`);

    res.json({
      success: true,
      message: `Merged ${addedAliases.length} horse(s) into "${primaryName}"`,
      primaryName,
      aliases: mapping[primaryName].aliases
    });

  } catch (error) {
    console.error('Error merging horses:', error);
    res.status(500).json({ error: 'Failed to merge horses' });
  }
});

// Remove an alias from a horse
app.post('/api/horses/unmerge', async (req, res) => {
  try {
    const { primaryName, aliasName } = req.body;

    if (!primaryName || !aliasName) {
      return res.status(400).json({ error: 'Primary name and alias name are required' });
    }

    const mapping = await getHorseMapping();

    if (!mapping[primaryName]) {
      return res.status(404).json({ error: 'Primary horse not found' });
    }

    if (!mapping[primaryName].aliases || !mapping[primaryName].aliases.includes(aliasName)) {
      return res.status(404).json({ error: 'Alias not found for this horse' });
    }

    // Remove the alias
    mapping[primaryName].aliases = mapping[primaryName].aliases.filter(a => a !== aliasName);
    mapping[primaryName].updatedAt = new Date().toISOString();

    // Optionally: create a new mapping entry for the unmerged horse
    mapping[aliasName] = {
      owner: '',
      country: '',
      isHistoric: false,
      aliases: [],
      addedAt: new Date().toISOString()
    };

    await saveHorseMapping(mapping);

    console.log(`Unmerged: removed "${aliasName}" from "${primaryName}"`);

    res.json({
      success: true,
      message: `Removed "${aliasName}" from "${primaryName}"`,
      primaryName,
      aliases: mapping[primaryName].aliases
    });

  } catch (error) {
    console.error('Error unmerging horses:', error);
    res.status(500).json({ error: 'Failed to unmerge horses' });
  }
});

// Rename a horse (change display name, keep old name as alias for training data)
app.post('/api/horses/rename', async (req, res) => {
  try {
    const { oldName, newName } = req.body;

    if (!oldName || !newName) {
      return res.status(400).json({ error: 'Old name and new name are required' });
    }

    if (oldName.toLowerCase() === newName.toLowerCase()) {
      return res.status(400).json({ error: 'New name must be different from the current name' });
    }

    const mapping = await getHorseMapping();

    // Find the old horse entry (might be exact or fuzzy match)
    const oldNameLower = oldName.toLowerCase();
    const oldNameStripped = oldName.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
    let oldKey = Object.keys(mapping).find(k =>
      k.toLowerCase() === oldNameLower ||
      k.replace(/[^a-zA-Z0-9]/g, '').toLowerCase() === oldNameStripped
    );

    // Get old horse data (or create empty if doesn't exist)
    const oldData = oldKey ? mapping[oldKey] : {
      owner: '',
      country: '',
      isHistoric: false,
      aliases: []
    };

    // Check if new name already exists
    const newNameLower = newName.toLowerCase();
    const existingNew = Object.keys(mapping).find(k => k.toLowerCase() === newNameLower);
    if (existingNew) {
      return res.status(400).json({ error: `A horse named "${existingNew}" already exists. Use Merge instead.` });
    }

    // Create new entry with the new name
    mapping[newName] = {
      owner: oldData.owner || '',
      country: oldData.country || '',
      isHistoric: oldData.isHistoric || false,
      aliases: oldData.aliases || [],
      addedAt: oldData.addedAt || new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      renamedFrom: oldKey || oldName
    };

    // Add old name(s) as aliases so training data resolves
    if (oldKey && !mapping[newName].aliases.includes(oldKey)) {
      mapping[newName].aliases.push(oldKey);
    }
    if (oldName !== oldKey && !mapping[newName].aliases.includes(oldName)) {
      mapping[newName].aliases.push(oldName);
    }

    // Delete old entry if it existed
    if (oldKey) {
      delete mapping[oldKey];
    }

    await saveHorseMapping(mapping);

    console.log(`Renamed horse: "${oldKey || oldName}" -> "${newName}" (aliases: ${mapping[newName].aliases.join(', ')})`);

    res.json({
      success: true,
      message: `Renamed "${oldKey || oldName}" to "${newName}"`,
      oldName: oldKey || oldName,
      newName,
      aliases: mapping[newName].aliases
    });

  } catch (error) {
    console.error('Error renaming horse:', error);
    res.status(500).json({ error: 'Failed to rename horse' });
  }
});

// Add a note for a horse
app.post('/api/notes', async (req, res) => {
  try {
    const { horseName, date, note } = req.body;

    if (!horseName || !date || !note) {
      return res.status(400).json({ error: 'Horse name, date, and note are required' });
    }

    const notes = await getHorseNotes();

    // Initialize array for this horse if needed
    if (!notes[horseName]) {
      notes[horseName] = [];
    }

    // Add the new note
    notes[horseName].push({
      date: date,
      note: note,
      createdAt: new Date().toISOString()
    });

    await saveHorseNotes(notes);

    console.log(`Added note for ${horseName} on ${date}`);

    res.json({
      success: true,
      message: `Note added for ${horseName}`,
      note: { horseName, date, note }
    });

  } catch (error) {
    console.error('Error adding note:', error);
    res.status(500).json({ error: 'Failed to add note' });
  }
});

// Get notes for a horse
app.get('/api/notes/:horseName', async (req, res) => {
  try {
    const { horseName } = req.params;
    const notes = await getHorseNotes();
    const horseNotes = notes[horseName] || [];

    res.json({
      success: true,
      horseName,
      notes: horseNotes
    });

  } catch (error) {
    console.error('Error getting notes:', error);
    res.status(500).json({ error: 'Failed to get notes' });
  }
});

// Delete a note for a horse
app.delete('/api/notes', async (req, res) => {
  try {
    const { horseName, date } = req.body;

    if (!horseName || !date) {
      return res.status(400).json({ error: 'Horse name and date are required' });
    }

    const notes = await getHorseNotes();

    if (!notes[horseName] || notes[horseName].length === 0) {
      return res.status(404).json({ error: 'No notes found for this horse' });
    }

    // Filter out the note with the matching date
    const originalLength = notes[horseName].length;
    notes[horseName] = notes[horseName].filter(n => n.date !== date);

    if (notes[horseName].length === originalLength) {
      return res.status(404).json({ error: 'Note not found for the specified date' });
    }

    await saveHorseNotes(notes);

    console.log(`Deleted note for ${horseName} on ${date}`);

    res.json({
      success: true,
      message: `Note deleted for ${horseName} on ${date}`
    });

  } catch (error) {
    console.error('Error deleting note:', error);
    res.status(500).json({ error: 'Failed to delete note' });
  }
});

// Regenerate all display names (apply formatting fixes without re-uploading)
app.post('/api/regenerate', async (req, res) => {
  try {
    const sessionId = 'arioneo-main-session';
    const session = await getSession(sessionId);

    if (!session || !session.allHorseDetailData) {
      return res.status(404).json({ error: 'No session data found' });
    }

    const horseMapping = await getHorseMapping();
    const horseData = generateHorseSummary(session.allHorseDetailData, horseMapping);

    await saveSession(sessionId, session.fileName, horseData, session.allHorseDetailData);

    res.json({
      success: true,
      message: `Regenerated data for ${horseData.length} horses`,
      horses: horseData.map(h => ({ name: h.name, displayName: h.displayName }))
    });
  } catch (error) {
    console.error('Error regenerating data:', error);
    res.status(500).json({ error: 'Failed to regenerate data' });
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

// Debug route - disabled in production for security
app.get('/api/debug', (req, res) => {
  if (process.env.NODE_ENV === 'production') {
    return res.status(404).json({ error: 'Not found' });
  }
  res.json({
    message: 'Server is working (dev only)',
    redisConnected: redis !== null
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

// ============================================
// RACE CHART UPLOAD AND PARSING
// ============================================

// Track name to abbreviation mapping
const trackAbbreviations = {
  'churchill downs': 'CD',
  'churchill': 'CD',
  'woodbine': 'WO',
  'santa anita': 'SA',
  'belmont': 'BEL',
  'belmont park': 'BEL',
  'saratoga': 'SAR',
  'keeneland': 'KEE',
  'gulfstream': 'GP',
  'gulfstream park': 'GP',
  'del mar': 'DMR',
  'arlington': 'AP',
  'oaklawn': 'OP',
  'oaklawn park': 'OP',
  'aqueduct': 'AQU',
  'laurel': 'LRL',
  'laurel park': 'LRL',
  'pimlico': 'PIM',
  'monmouth': 'MTH',
  'monmouth park': 'MTH',
  'tampa bay': 'TAM',
  'tampa bay downs': 'TAM',
  'fair grounds': 'FG',
  'los alamitos': 'LA',
  'golden gate': 'GG',
  'golden gate fields': 'GG',
  'turfway': 'TP',
  'turfway park': 'TP',
  'parx': 'PRX',
  'penn national': 'PEN',
  'charles town': 'CT',
  'remington': 'RP',
  'remington park': 'RP',
  'lone star': 'LS',
  'lone star park': 'LS',
  'indiana grand': 'IND',
  'canterbury': 'CBY',
  'ellis park': 'ELP',
  'colonial downs': 'CNL',
  'horseshoe indianapolis': 'IND'
};

// Get track abbreviation from full name
function getTrackAbbreviation(trackName) {
  if (!trackName) return 'UNK';
  const lower = trackName.toLowerCase().trim();

  // Check exact matches first
  if (trackAbbreviations[lower]) {
    return trackAbbreviations[lower];
  }

  // Check partial matches
  for (const [name, abbrev] of Object.entries(trackAbbreviations)) {
    if (lower.includes(name) || name.includes(lower)) {
      return abbrev;
    }
  }

  // If no match, create abbreviation from first letters
  const words = trackName.split(' ').filter(w => w.length > 0);
  if (words.length >= 2) {
    return (words[0][0] + words[1][0]).toUpperCase();
  }
  return trackName.substring(0, 3).toUpperCase();
}

// Race chart parsing class - rewritten for Equibase PDF format
class RaceChartParser {

  extractRaceMetadata(text) {
    let distanceInFurlongs = 0;
    let distanceText = 'Unknown';
    let raceDate = 'Unknown Date';
    let track = 'UNK';
    let surface = 'D'; // Default to Dirt
    let raceType = 'UNK';

    // Extract surface FIRST - check the distance/surface line at the top
    // Format: "6 Furlongs Dirt" or "1 1/16 Miles Inner Turf" or "7 Furlongs Turf"
    if (text.match(/inner\s+turf/i) || text.match(/turf/i)) {
      surface = 'T';
    } else if (text.match(/synthetic/i) || text.match(/all[\s-]?weather/i)) {
      surface = 'AWT';
    } else if (text.match(/dirt/i)) {
      surface = 'D';
    }

    // Extract distance - Miles format (e.g., "1 1/16 Miles")
    let milesMatch = text.match(/(\d+)\s+(\d+)\/(\d+)\s*Miles?/i);
    if (milesMatch) {
      const wholeMiles = parseInt(milesMatch[1]);
      const numerator = parseInt(milesMatch[2]);
      const denominator = parseInt(milesMatch[3]);
      const miles = wholeMiles + (numerator / denominator);
      distanceInFurlongs = Math.round(miles * 8 * 10) / 10; // Round to 1 decimal
      distanceText = `${distanceInFurlongs}F`;
    } else {
      // Try simple miles format (e.g., "1 Mile")
      milesMatch = text.match(/(\d+)\s*Miles?(?:\s|$)/i);
      if (milesMatch) {
        const miles = parseInt(milesMatch[1]);
        distanceInFurlongs = miles * 8;
        distanceText = `${distanceInFurlongs}F`;
      } else {
        // Furlongs format - fractional (e.g., "5 1/2 Furlongs")
        let furlongMatch = text.match(/(\d+)\s+(\d+)\/(\d+)\s*Furlong/i);

        if (furlongMatch) {
          const wholeFurlongs = parseInt(furlongMatch[1]);
          const numerator = parseInt(furlongMatch[2]);
          const denominator = parseInt(furlongMatch[3]);
          distanceInFurlongs = wholeFurlongs + (numerator / denominator);
          distanceText = `${distanceInFurlongs}F`;
        } else {
          // Decimal format
          furlongMatch = text.match(/(\d+(?:\.\d+)?)\s*Furlong/i);
          if (furlongMatch) {
            distanceInFurlongs = parseFloat(furlongMatch[1]);
            distanceText = `${distanceInFurlongs}F`;
          } else {
            // Word numbers
            furlongMatch = text.match(/(\d+|Seven|Six|Five|Eight|Nine)\s*Furlong/i);
            if (furlongMatch) {
              const wordToNumber = { 'five': 5, 'six': 6, 'seven': 7, 'eight': 8, 'nine': 9 };
              const furlongStr = furlongMatch[1].toLowerCase();
              distanceInFurlongs = wordToNumber[furlongStr] || parseInt(furlongMatch[1]);
              if (distanceInFurlongs && !isNaN(distanceInFurlongs)) {
                distanceText = `${distanceInFurlongs}F`;
              }
            }
          }
        }
      }
    }

    // Extract track name
    const trackNames = [
      'Churchill Downs', 'Woodbine', 'Santa Anita', 'Belmont', 'Saratoga',
      'Keeneland', 'Gulfstream', 'Del Mar', 'Arlington', 'Oaklawn',
      'Aqueduct', 'Laurel', 'Pimlico', 'Monmouth', 'Tampa Bay', 'Fair Grounds',
      'Los Alamitos', 'Golden Gate', 'Turfway', 'Parx', 'Penn National',
      'Charles Town', 'Remington', 'Lone Star', 'Indiana Grand', 'Canterbury',
      'Ellis Park', 'Colonial Downs', 'Horseshoe Indianapolis'
    ];

    for (const trackName of trackNames) {
      if (text.match(new RegExp(trackName.replace(' ', '\\s+'), 'i'))) {
        track = getTrackAbbreviation(trackName);
        break;
      }
    }

    // Turfway Park has a synthetic surface - always AWT
    if (track === 'TP') {
      surface = 'AWT';
    }

    // Extract race date
    let dateMatch = text.match(/(\w+,\s+\w+\s+\d{1,2},\s+\d{4})\s+(?:Saratoga|Churchill|Woodbine|Belmont|Keeneland)/i);
    if (!dateMatch) {
      dateMatch = text.match(/(\w+,\s+\w+\s+\d{1,2},\s+\d{4})\s+[A-Z][a-zA-Z]+/);
    }
    if (!dateMatch) {
      const allDates = text.match(/(\w+,\s+\w+\s+\d{1,2},\s+\d{4})/g);
      if (allDates && allDates.length > 0) {
        dateMatch = [null, allDates[allDates.length - 1]];
      }
    }
    raceDate = dateMatch ? dateMatch[1] : 'Unknown Date';

    // Extract race type - ORDER MATTERS! Check specific patterns first
    // Also extract race name for stakes races
    let raceName = '';

    // Check for stakes race name (e.g., "Seeking the Ante S." after track name)
    const stakesNameMatch = text.match(/(?:Saratoga|Churchill|Belmont|Keeneland|Gulfstream|Del Mar|Santa Anita|Aqueduct|Woodbine|Oaklawn|Tampa Bay|Fair Grounds|Monmouth|Laurel|Parx|Pimlico)\s*\n\s*([A-Z][A-Za-z'\s\-\.]+(?:S\.|Stakes|H\.|Handicap|Derby|Oaks|Cup|Futurity|Classic|Mile|Sprint|Distaff|Breeders'))/i);
    if (stakesNameMatch) {
      raceName = stakesNameMatch[1].trim();
    }

    if (text.match(/Allowance\s+Optional\s+Claiming/i) || text.match(/Optional\s+Claiming/i)) {
      raceType = 'AOC';
    } else if (text.match(/Maiden\s+Special\s+Weight/i)) {
      raceType = 'MSW';
    } else if (text.match(/Grade\s*([I1]{1,3}|\d+)/i)) {
      const gradeMatch = text.match(/Grade\s*([I1]{1,3}|\d+)/i);
      if (gradeMatch) {
        let grade = gradeMatch[1];
        if (grade === 'I') grade = '1';
        else if (grade === 'II') grade = '2';
        else if (grade === 'III') grade = '3';
        raceType = `G${grade}`;
      }
    } else if (raceName || text.match(/\bStakes\s+\d+yo/i) || text.match(/^\s*Stakes\b/m)) {
      // Stakes race - detected by race name or "Stakes 2yo" pattern
      raceType = 'STK';
    } else if (text.match(/Maiden\s+Claiming/i)) {
      raceType = 'MCL';
    } else if (text.match(/Starter\s+Allowance/i)) {
      raceType = 'STR';
    } else if (text.match(/\bAllowance\b/i) && !text.match(/Claiming\s+Price\s*\$/i)) {
      raceType = 'ALW';
    } else if (text.match(/Claiming\s+Price\s*\$/i) || text.match(/\bClaiming\s+\$\d/i)) {
      // Only CLM if there's an actual claiming price with $
      raceType = 'CLM';
    } else if (text.match(/\bHandicap\b/i)) {
      raceType = 'HCP';
    }

    return {
      raceDate,
      track,
      surface,
      raceType,
      raceName,
      distance: distanceText,
      distanceInFurlongs
    };
  }

  parseTime(timeString) {
    if (!timeString) return null;

    const timeRegex = /(?:(\d{1,2}):)?(\d{1,2})\.(\d{2})/;
    const match = timeString.match(timeRegex);

    if (!match) return null;

    const minutes = match[1] ? parseInt(match[1]) : 0;
    const seconds = parseInt(match[2]);
    const hundredths = parseInt(match[3]);

    const totalSeconds = minutes * 60 + seconds + hundredths / 100;

    const display = minutes > 0 ?
      `${minutes}:${seconds.toString().padStart(2, '0')}.${hundredths.toString().padStart(2, '0')}` :
      `${seconds}.${hundredths.toString().padStart(2, '0')}`;

    return { display, seconds: totalSeconds };
  }

  parsePosition(posString) {
    if (!posString) return '-';

    const posMatch = posString.match(/(\d+)([½¼¾])?/);
    if (!posMatch) return posString;

    const position = parseInt(posMatch[1]);
    const fraction = posMatch[2];

    let suffix;
    if (position % 100 >= 11 && position % 100 <= 13) {
      suffix = 'th';
    } else {
      switch (position % 10) {
        case 1: suffix = 'st'; break;
        case 2: suffix = 'nd'; break;
        case 3: suffix = 'rd'; break;
        default: suffix = 'th'; break;
      }
    }

    return `${position}${suffix}`;
  }

  calculateAvgSpeed(finalTimeSeconds, distanceInFurlongs) {
    const distanceInMiles = distanceInFurlongs / 8;
    const timeInHours = finalTimeSeconds / 3600;
    return distanceInMiles / timeInHours;
  }

  calculateFiveFReduction(finalTimeSeconds, distanceInFurlongs) {
    // Calculate time for 5 furlongs at same pace
    const pacePerFurlong = finalTimeSeconds / distanceInFurlongs;
    const fiveFTime = pacePerFurlong * 5;

    const mins = Math.floor(fiveFTime / 60);
    const secs = Math.floor(fiveFTime % 60);
    const hundr = Math.round((fiveFTime % 1) * 100);

    return mins > 0 ?
      `${mins}:${secs.toString().padStart(2, '0')}.${hundr.toString().padStart(2, '0')}` :
      `${secs}.${hundr.toString().padStart(2, '0')}`;
  }

  extractFinishPositionFromRow(horseLine, finalTime) {
    if (!finalTime) {
      return { pos1_4: '-', pos1_2: '-', pos3_4: '-', posFin: '-' };
    }

    const finalTimeStr = finalTime.display;
    const finalTimeIndex = horseLine.indexOf(finalTimeStr);
    if (finalTimeIndex === -1) {
      return { pos1_4: '-', pos1_2: '-', pos3_4: '-', posFin: '-' };
    }

    const afterFinalTime = horseLine.substring(finalTimeIndex + finalTimeStr.length);
    const positionMatch = afterFinalTime.match(/^\s*(\d+)\s+(\d+)\s+(\d+)\s+(\d+)/);

    if (positionMatch) {
      return {
        pos1_4: this.parsePosition(positionMatch[1]),
        pos1_2: this.parsePosition(positionMatch[2]),
        pos3_4: this.parsePosition(positionMatch[3]),
        posFin: this.parsePosition(positionMatch[4])
      };
    }

    return { pos1_4: '-', pos1_2: '-', pos3_4: '-', posFin: '-' };
  }

  parseHorseData(text, horseName, format, metadata) {
    if (format === 'column_format') {
      return this.parseColumnFormat(text, horseName, metadata);
    } else if (format === 'earnings_split_format') {
      return this.parseEarningsSplitFormat(text, horseName, metadata);
    } else {
      let result = this.parseColumnFormat(text, horseName, metadata);
      if (!result) {
        result = this.parseEarningsSplitFormat(text, horseName, metadata);
      }
      return result;
    }
  }

  parseColumnFormat(text, horseName, metadata) {
    const raceResultsSection = text.split('H Wt')[1];
    if (!raceResultsSection) return null;

    const horseIndex = raceResultsSection.toLowerCase().indexOf(horseName.toLowerCase());
    if (horseIndex === -1) return null;

    let endIndex = raceResultsSection.length;
    const afterHorse = raceResultsSection.substring(horseIndex + horseName.length);
    const nextHorseMatch = afterHorse.match(/\s+\d+\s+[A-Z][a-zA-Z\s()]+\s+\d+\.\d+/);
    if (nextHorseMatch) {
      endIndex = horseIndex + horseName.length + nextHorseMatch.index;
    }

    const horseFullLine = raceResultsSection.substring(horseIndex, endIndex);
    return this.parseHorseLineFromResults(horseFullLine, metadata);
  }

  parseEarningsSplitFormat(text, horseName, metadata) {
    const raceDataSection = text.split('H Wt')[1];
    if (!raceDataSection) return null;

    const horseIndex = raceDataSection.toLowerCase().indexOf(horseName.toLowerCase());
    if (horseIndex === -1) return null;

    let endIndex = raceDataSection.length;
    const afterHorse = raceDataSection.substring(horseIndex + horseName.length);
    const nextHorseMatch = afterHorse.match(/\s+\d+\s+[A-Z][a-zA-Z\s()]+\s+\d+\.\d+/);
    if (nextHorseMatch) {
      endIndex = horseIndex + horseName.length + nextHorseMatch.index;
    }

    const horseFullLine = raceDataSection.substring(horseIndex, endIndex);
    return this.parseWoodbineHorseLine(horseFullLine, metadata);
  }

  parseHorseLineFromResults(horseLine, metadata) {
    const timeRegex = /(\d{1,2}:\d{2}\.\d{2}|\d{2}\.\d{2})/g;
    const allTimes = [];
    let match;
    while ((match = timeRegex.exec(horseLine)) !== null) {
      allTimes.push(match[1]);
    }

    const raceTimes = allTimes.filter((time) => {
      const parsed = this.parseTime(time);
      if (!parsed) return false;
      return parsed.seconds >= 20;
    });

    let f1Time = null, f2Time = null, f3Time = null, finalTime = null;

    if (raceTimes.length >= 3) {
      f1Time = this.parseTime(raceTimes[0]);
      f2Time = this.parseTime(raceTimes[1]);

      const finalTimes = raceTimes.filter(time => time.includes(':'));
      if (finalTimes.length >= 1) {
        finalTime = this.parseTime(finalTimes[0]);
        if (finalTimes.length >= 2) {
          f3Time = this.parseTime(finalTimes[1]);
        }
      }
    }

    // Try to find F3 if not found
    if (!f3Time && raceTimes.length > 3) {
      for (let i = 2; i < raceTimes.length; i++) {
        const time = this.parseTime(raceTimes[i]);
        if (time && time.seconds >= 45 && time.seconds <= 90 && !raceTimes[i].includes(':')) {
          f3Time = time;
          break;
        }
      }
    }

    const positions = this.extractFinishPositionFromRow(horseLine, finalTime);

    // Extract comment
    const commentMatch = horseLine.match(/(\d+\.\d+)\s+([a-zA-Z0-9][a-zA-Z0-9\s,\-\'\(\)\/\\]+?)\s+(\d{2}\.\d{2})/);
    let comment = '';
    if (commentMatch) {
      comment = commentMatch[2].trim();
    }

    return this.buildHorseDataResult(f1Time, f2Time, f3Time, finalTime, positions, comment, metadata.distanceInFurlongs);
  }

  parseWoodbineHorseLine(horseLine, metadata) {
    const earningsIndex = horseLine.indexOf('$');
    if (earningsIndex === -1) return null;

    const beforeEarnings = horseLine.substring(0, earningsIndex);
    const afterEarnings = horseLine.substring(earningsIndex);

    const timeRegex = /(\d{1,2}:\d{2}\.\d{2}|\d{2}\.\d{2})/g;

    const beforeTimes = [];
    let match;
    while ((match = timeRegex.exec(beforeEarnings)) !== null) {
      beforeTimes.push(match[1]);
    }

    timeRegex.lastIndex = 0;
    const afterTimes = [];
    while ((match = timeRegex.exec(afterEarnings)) !== null) {
      afterTimes.push(match[1]);
    }

    let f1Time = null, f2Time = null;
    if (beforeTimes.length >= 2) {
      const raceTimes = beforeTimes.filter(time => {
        const seconds = parseFloat(time);
        return seconds >= 20 && seconds <= 60;
      });

      if (raceTimes.length >= 2) {
        f1Time = this.parseTime(raceTimes[0]);
        f2Time = this.parseTime(raceTimes[1]);
      }
    }

    let finalTime = null, f3Time = null;
    if (afterTimes.length >= 1) {
      finalTime = this.parseTime(afterTimes[0]);

      for (let i = 1; i < afterTimes.length; i++) {
        const timeStr = afterTimes[i];
        const timeSeconds = this.parseTime(timeStr)?.seconds;

        if (timeStr && !timeStr.includes(':') && timeSeconds &&
            timeSeconds >= 45 && timeSeconds <= 90) {
          f3Time = this.parseTime(timeStr);
          break;
        }
      }
    }

    const positions = this.extractFinishPositionFromRow(horseLine, finalTime);

    const commentMatch = horseLine.match(/\d+\.\d+\s+([a-zA-Z].*?)\s+\d+\.\d+/);
    const comment = commentMatch ? commentMatch[1].trim() : '';

    return this.buildHorseDataResult(f1Time, f2Time, f3Time, finalTime, positions, comment, metadata.distanceInFurlongs);
  }

  buildHorseDataResult(f1Time, f2Time, f3Time, finalTime, positions, comment, distanceInFurlongs) {
    const avgSpeedMph = finalTime && finalTime.seconds > 0 && distanceInFurlongs > 0 ?
      this.calculateAvgSpeed(finalTime.seconds, distanceInFurlongs) : 0;
    const fiveFReductionTime = finalTime && finalTime.seconds > 0 && distanceInFurlongs > 0 ?
      this.calculateFiveFReduction(finalTime.seconds, distanceInFurlongs) : '';

    return {
      finalTime: finalTime ? finalTime.display : '-',
      avgSpeedMph,
      fiveFReductionTime,
      f1Time: f1Time ? f1Time.display : '-',
      f2Time: f2Time ? f2Time.display : '-',
      f3Time: f3Time ? f3Time.display : '-',
      pos1_4: positions.pos1_4,
      pos1_2: positions.pos1_2,
      pos3_4: positions.pos3_4,
      posFin: positions.posFin,
      comment
    };
  }

  parseRaceChart(text, horseName) {
    const metadata = this.extractRaceMetadata(text);

    // Use the new parser
    const horseData = this.parseHorseFromChart(text, horseName);

    if (!horseData) return null;

    // Calculate avg speed and 5F reduction
    let avgSpeedMph = 0;
    let fiveFReductionTime = '-';

    if (horseData.finalTime && horseData.finalTime !== '-' && metadata.distanceInFurlongs > 0) {
      const finalSeconds = this.timeToSeconds(horseData.finalTime);
      if (finalSeconds > 0) {
        avgSpeedMph = this.calculateAvgSpeed(finalSeconds, metadata.distanceInFurlongs);
        fiveFReductionTime = this.calculateFiveFReduction(finalSeconds, metadata.distanceInFurlongs);
      }
    }

    // For 6F races, the final time IS the 6F time (no separate 3/4 time)
    // For longer races, use the extracted f3Time
    let f3Time = horseData.f3Time;
    if (metadata.distanceInFurlongs === 6) {
      f3Time = horseData.finalTime;
    }

    return {
      ...metadata,
      finalTime: horseData.finalTime,
      avgSpeedMph,
      fiveFReductionTime,
      f1Time: horseData.f1Time,
      f2Time: horseData.f2Time,
      f3Time: f3Time,
      pos1_4: horseData.pos1_4,
      pos1_2: horseData.pos1_2,
      pos3_4: horseData.pos3_4,
      posFin: horseData.posFin,
      comment: horseData.comment,
      horseName
    };
  }

  timeToSeconds(timeStr) {
    if (!timeStr || timeStr === '-') return 0;
    if (timeStr.includes(':')) {
      const parts = timeStr.split(':');
      return parseInt(parts[0]) * 60 + parseFloat(parts[1]);
    }
    return parseFloat(timeStr);
  }

  // Find all horses mentioned in the race chart
  findAllHorsesInChart(text) {
    const horses = [];

    // Split by 'H Wt' which appears before the results table
    const parts = text.split(/H\s*Wt/i);
    if (parts.length < 2) return horses;

    const raceDataSection = parts[1];

    // Pattern: Horse name followed directly by odds (e.g., "Paradise3.80" or "Ontario6.15")
    // Horse names are capitalized words, may include spaces, may have country code like (IRE)
    // The pattern looks for: CapitalizedName + optional(MoreWords) + optional(CountryCode) + Odds
    const horsePattern = /\n([A-Z][a-zA-Z']+(?:\s+[A-Z][a-zA-Z']+)*(?:\s*\([A-Z]{2,3}\))?)\s*(\d+\.\d+)/g;

    let match;
    while ((match = horsePattern.exec(raceDataSection)) !== null) {
      let name = match[1].trim();
      // Skip common non-horse-name patterns
      if (name.length > 2 &&
          !name.match(/^(Scratches|Horse|Jockey|Trainer|Owner|Pool|Exacta|Trifecta|Super|Daily|Pick|WPS|Omni)/i) &&
          !horses.some(h => h.toLowerCase() === name.toLowerCase())) {
        horses.push(name);
      }
    }

    // Also try to find horses in the jockey/trainer section format:
    // "HorseName\nJockeyName\n" pattern
    const altPattern = /\n([A-Z][a-zA-Z']+(?:\s+[A-Z][a-zA-Z']+)*(?:\s*\([A-Z]{2,3}\))?)\n[A-Z][a-z]+\s+[A-Z]/g;
    while ((match = altPattern.exec(text)) !== null) {
      let name = match[1].trim();
      if (name.length > 2 &&
          !name.match(/^(Scratches|Horse|Jockey|Trainer|Owner|Last|Raced|Fin)/i) &&
          !horses.some(h => h.toLowerCase() === name.toLowerCase())) {
        horses.push(name);
      }
    }

    return horses;
  }

  // Parse a specific horse's data from the race chart
  parseHorseFromChart(text, horseName) {
    // Split by 'H Wt' to get results section only
    const parts = text.split(/H\s*Wt/i);
    if (parts.length < 2) return null;

    const resultsSection = parts[1];

    // Find the line with this horse's data in the results section
    // Format: HorseName + Odds + Comment + F1Time + F2Time + Earnings + FinalTime + Positions
    // Example: "Paradise3.80brkout,3-5p,kpt at bay23.1246.19$24,0001011:11.077422"
    const lines = resultsSection.split('\n');
    let horseLine = '';
    let horseLineIndex = -1;

    for (let i = 0; i < lines.length; i++) {
      // Match horse name followed immediately by odds (e.g., "Paradise3.80")
      // Use case-insensitive matching since DB has uppercase names but PDFs may have mixed case
      const pattern = new RegExp(horseName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\d+\\.\\d+', 'i');
      if (pattern.test(lines[i])) {
        horseLine = lines[i];
        horseLineIndex = i;
        break;
      }

      // Also check if horse name is on its own line and odds are on the next line
      // This handles PDFs where text extraction splits them across lines
      const horseOnlyPattern = new RegExp('^' + horseName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '$', 'i');
      if (horseOnlyPattern.test(lines[i].trim()) && i + 1 < lines.length) {
        // Check if next line starts with odds
        const nextLine = lines[i + 1];
        if (/^\d+\.\d+/.test(nextLine.trim())) {
          // Combine this line with subsequent lines to form the full horse data
          // But STOP before the margins line or 6F time line
          horseLine = lines[i] + nextLine;
          for (let j = 2; j <= 5 && i + j < lines.length; j++) {
            const checkLine = lines[i + j];
            // Stop if we hit the 6F time line (contains /m or /f followed by time and state)
            if (/\d+L\d+\/[a-z][12]:\d{2}\.\d{2}[A-Z]{2}/i.test(checkLine)) {
              break;
            }
            // Stop if we hit the margins line (contains ½, ¼, ¾ or starts with position margins like "hd", "nk")
            if (/[½¼¾]/.test(checkLine) || /^(hd|nk|nse|\d+[½¼¾])/i.test(checkLine.trim())) {
              break;
            }
            horseLine += checkLine;
          }
          horseLineIndex = i;
          break;
        }
      }
    }

    if (!horseLine) return null;

    // Look at subsequent lines (up to 6) for the 6F time
    // Format: {weight}L{num}/{letter}{time}{state} e.g., "117L4/f1:13.91ON" or "115L3/m1:11.83KY"
    // The letter can be 'f', 'm', or others - so match any letter before the time
    let sixFurlongTime = null;
    for (let j = 1; j <= 6 && horseLineIndex + j < lines.length; j++) {
      const nextLine = lines[horseLineIndex + j];
      // Look for /{letter} followed by M:SS.SS time (6F time pattern)
      const sixFMatch = nextLine.match(/\/[a-z]([12]:\d{2}\.\d{2})/i);
      if (sixFMatch) {
        sixFurlongTime = sixFMatch[1];
        break;
      }
    }

    // Extract after horse name (case-insensitive search)
    const horseIndex = horseLine.toLowerCase().indexOf(horseName.toLowerCase());
    const afterHorse = horseLine.substring(horseIndex + horseName.length);

    // Extract odds and comment
    // Pattern: odds + comment + first time
    const oddsCommentMatch = afterHorse.match(/^(\d+\.\d+)([a-zA-Z][a-zA-Z0-9,\s\-'\/\\]+?)(\d{2}\.\d{2})/);
    const comment = oddsCommentMatch ? oddsCommentMatch[2].trim() : '';

    // Extract all times from the line
    // Fractional times: SS.SS format (20-95 seconds range)
    const allFractionalMatches = afterHorse.match(/\d{2}\.\d{2}/g) || [];
    const fractionalTimes = allFractionalMatches.filter(t => {
      const secs = parseFloat(t);
      return secs >= 20 && secs <= 95;
    });

    // M:SS.SS format times (6F time and final time are in this format)
    // 6F time is typically 1:07.00 - 1:20.00
    // Final time is typically 1:07.00 - 2:30.00
    const allMinutesTimes = afterHorse.match(/[12]:\d{2}\.\d{2}/g) || [];

    // FIRST M:SS.SS time is the final time (not last - horse boundary detection may include next horse's data)
    const finalTime = allMinutesTimes.length > 0 ? allMinutesTimes[0] : '-';

    // 6F time (3/4 time): Priority order:
    // 1. sixFurlongTime from subsequent line (/f pattern)
    // 2. M:SS.SS time in 1:07-1:20 range from horse line
    // 3. SS.SS fractional time in 55-95 second range
    let f3Time = '-';

    // First priority: sixFurlongTime from /f pattern in subsequent lines
    if (sixFurlongTime) {
      f3Time = sixFurlongTime;
    }

    // Second priority: M:SS.SS time from horse line in 6F range
    if (f3Time === '-') {
      for (const time of allMinutesTimes) {
        // Skip the final time
        if (time === finalTime && allMinutesTimes.length > 1) continue;
        // Check if it's in the 6F range (1:07.00 - 1:20.00)
        const [mins, secs] = time.split(':');
        const totalSecs = parseInt(mins) * 60 + parseFloat(secs);
        if (totalSecs >= 67 && totalSecs <= 80) { // 1:07 to 1:20
          f3Time = time;
          break;
        }
      }
    }

    // Third priority: fractional times in SS.SS format
    if (f3Time === '-') {
      for (let i = 2; i < fractionalTimes.length; i++) {
        const secs = parseFloat(fractionalTimes[i]);
        if (secs >= 55 && secs <= 95) {
          f3Time = fractionalTimes[i];
          break;
        }
      }
    }

    // Parse times: F1 (1/4 time), F2 (1/2 time)
    let f1Time = fractionalTimes[0] || '-';
    let f2Time = fractionalTimes[1] || '-';

    // Extract positions from end of line after final time
    // Positions are typically single digits (1-9) or occasionally 10-12
    // Format varies: could be "7422" (4 positions) or "31112" (5 positions with start)
    let pos1_4 = '-', pos1_2 = '-', pos3_4 = '-', posFin = '-';

    if (finalTime !== '-') {
      const afterFinalTime = afterHorse.substring(afterHorse.lastIndexOf(finalTime) + finalTime.length);

      // Try Format A first: concatenated digits like "109765" or "7422"
      const positionMatch = afterFinalTime.match(/^(\d+)/);
      const digitsOnly = positionMatch ? positionMatch[1] : '';
      // Expected positions: 4 for sprints, 5 for routes
      // Use digit length as heuristic: 4 digits = 4 positions, 5+ digits may have double-digit positions
      const expectedPositions = digitsOnly.length <= 5 ? digitsOnly.length : 5;
      let positions = this.parsePositions(digitsOnly, expectedPositions);

      console.log(`[DEBUG] Horse: ${horseName}, afterFinalTime: "${afterFinalTime.substring(0, 50)}", digitsOnly: "${digitsOnly}"`);

      // If Format A didn't work, try Format B: space-separated tokens with margins
      // Format B example: " 4 3 11 1hd 1½ 11 in hand" (PP, St, 1/4, 1/2, Str, Fin)
      if (positions.length < 4) {
        const tokens = afterFinalTime.trim().split(/\s+/);
        // Look for position tokens: start with digit, may have margin attached (e.g., "11", "1hd", "32½")
        const positionTokens = [];
        for (const token of tokens) {
          if (/^\d/.test(token)) {
            positionTokens.push(token);
          } else {
            // Stop when we hit non-position tokens like "in hand" or comments
            if (positionTokens.length >= 4) break;
          }
        }

        console.log(`[DEBUG] Format B tokens: [${positionTokens.join(', ')}]`);

        // Format B has: PP, St, then 4 position columns (1/4, 1/2, Str, Fin for sprints)
        // or: PP, St, then 5 position columns (1/4, 1/2, 3/4, Str, Fin for routes)
        if (positionTokens.length >= 6) {
          // Skip PP (index 0) and St (index 1), take positions from index 2 onwards
          const posTokensOnly = positionTokens.slice(2);
          positions = posTokensOnly.map(t => this.extractPositionFromToken(t));
          console.log(`[DEBUG] Format B positions (skipped PP/St): [${positions.join(', ')}]`);
        }
      }

      if (positions.length >= 5) {
        // 5 positions: 1/4, 1/2, 3/4, Str, Fin - take indices 0,1,2,4 (skip Str)
        pos1_4 = this.formatPosition(positions[0]);
        pos1_2 = this.formatPosition(positions[1]);
        pos3_4 = this.formatPosition(positions[2]);
        posFin = this.formatPosition(positions[4]);
      } else if (positions.length === 4) {
        // 4 positions: 1/4, 1/2, Str, Fin (for sprints without 3/4 call)
        pos1_4 = this.formatPosition(positions[0]);
        pos1_2 = this.formatPosition(positions[1]);
        pos3_4 = this.formatPosition(positions[2]); // Actually Str position for sprints
        posFin = this.formatPosition(positions[3]);
      } else if (positions.length >= 2) {
        // If fewer positions, just get finish
        posFin = this.formatPosition(positions[positions.length - 1]);
      }

      console.log(`[DEBUG] Final positions - 1/4: ${pos1_4}, 1/2: ${pos1_2}, 3/4: ${pos3_4}, Fin: ${posFin}`);
    }

    return {
      f1Time,
      f2Time,
      f3Time,
      finalTime,
      pos1_4,
      pos1_2,
      pos3_4,
      posFin,
      comment
    };
  }

  formatPosition(pos) {
    const p = parseInt(pos);
    if (isNaN(p)) return pos;
    if (p === 1) return '1st';
    if (p === 2) return '2nd';
    if (p === 3) return '3rd';
    return `${p}th`;
  }

  // Extract position number from a token that may have margin attached
  // Examples: "11" -> 1 (position 1, margin 1), "1hd" -> 1, "32" -> 3, "10" -> 10
  // Position 0 doesn't exist, margin 0 doesn't exist, so "10" must be position 10
  extractPositionFromToken(token) {
    if (!token) return 0;

    // Extract leading digits
    const digitMatch = token.match(/^(\d+)/);
    if (!digitMatch) return 0;

    const digits = digitMatch[1];

    // Special case: "10", "11", "12" with no additional margin chars could be position 10/11/12
    // But "10" as position 1 margin 0 is invalid (margin 0 doesn't exist)
    if (digits === '10') return 10;
    if (digits === '11' && token.length === 2) return 11; // "11" alone could be position 11
    if (digits === '12' && token.length === 2) return 12;

    // For tokens with non-digit margins (like "1hd", "3nk", "2½"), first digit is position
    if (token.length > digits.length) {
      // Has margin chars after digits - first digit is position
      return parseInt(digits[0]);
    }

    // For pure digit tokens like "32", "64", "11" with nothing after:
    // These are position+margin where first digit is position
    // (Since races rarely have >12 horses, "32" can't be position 32)
    if (parseInt(digits) > 12) {
      return parseInt(digits[0]);
    }

    // Ambiguous case: "11" could be position 11 or position 1 margin 1
    // Default to first digit as position (safer for typical races)
    return parseInt(digits[0]);
  }

  // Parse position string into array of positions, handling double-digit positions
  // Only "10" is unambiguous (position 0 doesn't exist, so "10" must be position 10)
  // "11" and "12" are ambiguous: could be position 11/12 OR positions 1,1 / 1,2
  // We use heuristic: if digit count matches expected positions (4-5), parse as singles
  parsePositions(digitsStr, expectedCount = 0) {
    const positions = [];
    let i = 0;

    // If digit count matches expected, all are single digits (e.g., "1111" → [1,1,1,1])
    if (expectedCount > 0 && digitsStr.length === expectedCount) {
      for (const d of digitsStr) {
        positions.push(parseInt(d));
      }
      return positions;
    }

    while (i < digitsStr.length) {
      // Only "10" is unambiguously double-digit (position 0 is invalid)
      if (digitsStr[i] === '1' && i + 1 < digitsStr.length && digitsStr[i + 1] === '0') {
        positions.push(10);
        i += 2;
        continue;
      }

      // For "11" and "12": only treat as double-digit if we have MORE digits than expected
      // This handles cases like "109765" (6 digits for 5 positions) but not "1111" (4 digits for 4 positions)
      if (digitsStr[i] === '1' && i + 1 < digitsStr.length) {
        const nextDigit = digitsStr[i + 1];
        if ((nextDigit === '1' || nextDigit === '2') &&
            (expectedCount === 0 || digitsStr.length > expectedCount)) {
          // Likely a double-digit position (11 or 12)
          positions.push(parseInt(digitsStr.substring(i, i + 2)));
          i += 2;
          continue;
        }
      }
      // Single digit position
      positions.push(parseInt(digitsStr[i]));
      i++;
    }
    return positions;
  }
}

const raceChartParser = new RaceChartParser();

// Fuzzy matching for horse names
function fuzzyMatchHorse(chartHorseName, existingHorses) {
  if (!existingHorses || existingHorses.length === 0) {
    return { match: null, confidence: 0 };
  }

  // Normalize the chart horse name
  const normalizedChartName = chartHorseName
    .replace(/\s*\([A-Z]{2,3}\)$/g, '') // Remove country codes like (GB), (IRE)
    .trim()
    .toUpperCase();

  // First check for exact match (case-insensitive)
  const exactMatch = existingHorses.find(h =>
    h.name.toUpperCase() === normalizedChartName ||
    h.displayName?.toUpperCase() === normalizedChartName
  );

  if (exactMatch) {
    return { match: exactMatch, confidence: 1.0 };
  }

  // Check aliases
  for (const horse of existingHorses) {
    if (horse.aliases && horse.aliases.length > 0) {
      const aliasMatch = horse.aliases.find(alias =>
        alias.toUpperCase() === normalizedChartName
      );
      if (aliasMatch) {
        return { match: horse, confidence: 1.0 };
      }
    }
  }

  // Use Fuse.js for fuzzy matching
  const fuse = new Fuse(existingHorses, {
    keys: ['name', 'displayName', 'aliases'],
    threshold: 0.4, // Lower = stricter matching
    includeScore: true
  });

  const results = fuse.search(normalizedChartName);

  if (results.length > 0) {
    const bestMatch = results[0];
    const confidence = 1 - bestMatch.score; // Fuse score is 0 (perfect) to 1 (no match)
    return { match: bestMatch.item, confidence };
  }

  return { match: null, confidence: 0 };
}

// Check for duplicate race entries
async function checkDuplicateRace(horseName, raceDate, track) {
  try {
    const sessionData = await getSession('arioneo-main-session');
    if (!sessionData || !sessionData.allHorseDetailData) {
      return false;
    }

    // Normalize horse name for lookup
    const normalizedName = horseName.toUpperCase().replace(/\s*\([A-Z]{2,3}\)$/g, '').trim();

    // Check all variations of the horse name
    for (const [key, entries] of Object.entries(sessionData.allHorseDetailData)) {
      const keyNormalized = key.toUpperCase();
      if (keyNormalized === normalizedName || keyNormalized.includes(normalizedName) || normalizedName.includes(keyNormalized)) {
        // Found horse, check for duplicate race
        for (const entry of entries) {
          if (entry.isRace && entry.date === raceDate && entry.track === track) {
            return true; // Duplicate found
          }
        }
      }
    }

    return false;
  } catch (error) {
    console.error('Error checking duplicate race:', error);
    return false;
  }
}

// Upload and parse multiple race chart PDFs
app.post('/api/upload/race-charts', pdfUpload.array('pdfs', 50), async (req, res) => {
  try {
    if (!req.files || req.files.length === 0) {
      return res.status(400).json({ error: 'No PDF files uploaded' });
    }

    console.log(`Processing ${req.files.length} race chart PDFs`);

    // Get existing horses for fuzzy matching from BOTH horse mapping AND session training data
    const horseMapping = await getHorseMapping();
    const sessionData = await getSession('arioneo-main-session');

    // Combine horses from mapping and training data
    const existingHorsesMap = new Map();

    // Add horses from mapping
    for (const horse of Object.values(horseMapping)) {
      if (horse && horse.name) {
        existingHorsesMap.set(horse.name.toLowerCase(), horse);
      }
    }

    // Add horses from training data (session) - these may not be in the mapping yet
    if (sessionData && sessionData.allHorseDetailData) {
      for (const horseName of Object.keys(sessionData.allHorseDetailData)) {
        if (!existingHorsesMap.has(horseName.toLowerCase())) {
          existingHorsesMap.set(horseName.toLowerCase(), {
            name: horseName,
            displayName: horseName,
            owner: '-',
            country: '-'
          });
        }
      }
    }

    const existingHorses = Array.from(existingHorsesMap.values());
    console.log(`Found ${existingHorses.length} existing horses for fuzzy matching`);
    if (existingHorses.length > 0) {
      console.log('Training data horse names:', existingHorses.slice(0, 10).map(h => h.name).join(', ') + '...');
    }

    const results = [];

    for (const file of req.files) {
      const result = {
        fileName: file.originalname,
        success: false,
        data: null,
        error: null,
        horsesFound: [],
        needsVerification: false,
        matchedHorse: null,
        matchConfidence: 0,
        isDuplicate: false
      };

      try {
        // Read and parse PDF
        const pdfBuffer = fs.readFileSync(file.path);
        const pdfData = await pdfParse(pdfBuffer);
        const text = pdfData.text;

        // Find all horses in the chart
        const horsesInChart = raceChartParser.findAllHorsesInChart(text);
        // Filter to clean horse names (no newlines, reasonable length)
        // Filter to clean horse names (no newlines, reasonable length, no jockey/trainer concatenations)
        const cleanHorses = horsesInChart.filter(h => {
          if (!h || h.includes('\n') || h.length < 2 || h.length > 25) return false;
          // Filter out jockey/trainer name concatenations (lowercase followed immediately by uppercase)
          if (/[a-z][A-Z]/.test(h)) return false;
          // Filter out common non-horse patterns
          if (/^(Scratches|Horse|Jockey|Trainer|Owner|Calumet|Douglas)/i.test(h)) return false;
          return true;
        });
        result.horsesFound = cleanHorses;

        // Parse data for ALL horses in the chart
        const allHorseData = {};
        for (const chartHorse of cleanHorses) {
          const horseRaceData = raceChartParser.parseRaceChart(text, chartHorse);
          if (horseRaceData) {
            allHorseData[chartHorse] = horseRaceData;
          }
        }
        result.allHorseData = allHorseData;

        // Try to match each horse found with existing horses
        let bestMatch = null;
        let bestConfidence = 0;
        let bestChartHorse = null;

        for (const chartHorse of cleanHorses) {
          const { match, confidence } = fuzzyMatchHorse(chartHorse, existingHorses);
          if (confidence > bestConfidence) {
            bestMatch = match;
            bestConfidence = confidence;
            bestChartHorse = chartHorse;
          }
        }

        if (bestMatch && bestConfidence >= 0.6) {
          // Use the matched horse's data
          const raceData = allHorseData[bestChartHorse];

          if (raceData) {
            // Check for duplicate
            const isDuplicate = await checkDuplicateRace(
              bestMatch.name,
              formatRaceDate(raceData.raceDate),
              raceData.track
            );

            result.success = true;
            result.data = raceData;
            result.matchedHorse = bestMatch;
            result.matchConfidence = bestConfidence;
            result.needsVerification = bestConfidence < 0.9;
            result.isDuplicate = isDuplicate;
            result.selectedHorse = bestChartHorse;
          } else {
            result.error = `Could not parse race data for ${bestChartHorse}`;
          }
        } else if (cleanHorses.length > 0) {
          // Found horses but no good match - use first horse's data as default
          result.error = 'No matching horse found in your system';
          result.needsVerification = true;

          const firstHorse = cleanHorses[0];
          result.data = allHorseData[firstHorse] || null;
          result.selectedHorse = firstHorse;
        } else {
          result.error = 'No horses found in race chart';
        }

      } catch (parseError) {
        console.error(`Error parsing ${file.originalname}:`, parseError);
        result.error = `Parse error: ${parseError.message}`;
        result.needsVerification = true;
      } finally {
        // Clean up uploaded file
        try {
          fs.unlinkSync(file.path);
        } catch (e) {
          console.error('Error deleting temp file:', e);
        }
      }

      results.push(result);
    }

    // Return existing horses for manual selection if needed
    const horsesForSelection = existingHorses.map(h => ({
      name: h.name,
      displayName: h.displayName || h.name,
      owner: h.owner,
      country: h.country
    }));

    res.json({
      success: true,
      results,
      existingHorses: horsesForSelection,
      totalProcessed: results.length,
      successfulMatches: results.filter(r => r.success && !r.needsVerification).length,
      needsVerification: results.filter(r => r.needsVerification).length,
      duplicates: results.filter(r => r.isDuplicate).length
    });

  } catch (error) {
    console.error('Error processing race chart uploads:', error);
    res.status(500).json({ error: 'Failed to process race charts: ' + error.message });
  }
});

// Helper to format race date to MM/DD/YYYY
function formatRaceDate(dateStr) {
  if (!dateStr || dateStr === 'Unknown Date') return dateStr;

  try {
    const date = new Date(dateStr);
    if (isNaN(date.getTime())) return dateStr;

    const month = date.getMonth() + 1;
    const day = date.getDate();
    const year = date.getFullYear();

    return `${month}/${day}/${year}`;
  } catch (e) {
    return dateStr;
  }
}

// Save reviewed race data
app.post('/api/race-charts/save', async (req, res) => {
  try {
    const { races } = req.body;

    if (!races || !Array.isArray(races) || races.length === 0) {
      return res.status(400).json({ error: 'No race data provided' });
    }

    // Get current session data
    let sessionData = await getSession('arioneo-main-session');
    if (!sessionData) {
      sessionData = {
        id: 'arioneo-main-session',
        fileName: 'race-charts',
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
        horseData: [],
        allHorseDetailData: {}
      };
    }

    const savedRaces = [];
    const skippedRaces = [];

    for (const race of races) {
      // Find existing horse key (case-insensitive) FIRST before creating entry
      let horseKey = race.horseName;
      const existingKeys = Object.keys(sessionData.allHorseDetailData);
      const matchingKey = existingKeys.find(k => k.toLowerCase() === race.horseName.toLowerCase());
      if (matchingKey) {
        horseKey = matchingKey; // Use existing case (e.g., "Blanco" instead of "BLANCO")
      }

      // Skip duplicates (check with the normalized horse key)
      const isDuplicate = await checkDuplicateRace(horseKey, race.date, race.track);
      if (isDuplicate) {
        skippedRaces.push({ horseName: horseKey, date: race.date, reason: 'Duplicate entry' });
        continue;
      }

      // Create training entry from race data (use horseKey, not race.horseName)
      const raceEntry = {
        date: race.date,
        horse: horseKey,
        type: 'Race',
        track: race.track,
        surface: race.surface,
        distance: race.distance,
        avgSpeed: race.avgSpeed ? parseFloat(race.avgSpeed).toFixed(1) : '-',
        maxSpeed: race.raceType, // Race type in max speed column
        best1f: race.pos1_4 || '-', // 1/4 position
        best2f: race.f1Time || '-', // 1/4 time
        best3f: race.pos1_2 || '-', // 1/2 position
        best4f: race.f2Time || '-', // 1/2 time
        best5f: race.fiveFReduction || '-', // 5F reduction time
        best6f: race.pos3_4 || '-', // 3/4 position
        best7f: race.f3Time || '-', // 3/4 time
        maxHR: race.finalTime || '-', // Final time
        fastRecovery: race.finishPosition || '-', // Finish position
        fastQuality: '-',
        fastPercent: '-',
        recovery15: '-',
        quality15: '-',
        hr15Percent: '-',
        maxSL: '-',
        slGallop: '-',
        sfGallop: '-',
        slWork: '-',
        sfWork: '-',
        hr2min: '-',
        hr5min: '-',
        symmetry: '-',
        regularity: '-',
        bpm120: '-',
        zone5: '-',
        age: '-',
        sex: '-',
        temp: '-',
        distanceCol: '-',
        trotHR: '-',
        walkHR: '-',
        isRace: true,
        isWork: false,
        notes: race.comments || ''
      };

      // Create array for horse if doesn't exist
      if (!sessionData.allHorseDetailData[horseKey]) {
        sessionData.allHorseDetailData[horseKey] = [];
      }

      // Add race entry
      sessionData.allHorseDetailData[horseKey].push(raceEntry);

      // Sort by date (newest first)
      sessionData.allHorseDetailData[horseKey].sort((a, b) => {
        const dateA = new Date(a.date);
        const dateB = new Date(b.date);
        return dateB - dateA;
      });

      savedRaces.push({ horseName: horseKey, date: race.date });

      // Ensure horse is in horse mapping
      const horseMapping = await getHorseMapping();
      if (!horseMapping[horseKey.toLowerCase()]) {
        // Add to mapping if new horse was manually created
        if (race.isNewHorse) {
          horseMapping[horseKey.toLowerCase()] = {
            name: horseKey,
            displayName: horseKey,
            owner: race.owner || '-',
            country: race.country || '-',
            isHistoric: false
          };
          await saveHorseMapping(horseMapping);
        }
      }
    }

    // Update summary horse data - only for horses that had races added
    // Preserve all existing horse data, just update the ones we touched
    const horsesWithNewRaces = new Set(savedRaces.map(r => r.horseName.toLowerCase()));
    const updatedSummaries = buildHorseSummaryFromDetailData(sessionData.allHorseDetailData);

    // Create a map of existing horses by name (case-insensitive)
    // Note: CSV data uses 'name' property, race summaries use 'horse' property
    const existingHorseMap = new Map();
    for (const horse of sessionData.horseData || []) {
      const horseName = horse?.horse || horse?.name;
      if (horseName) {
        existingHorseMap.set(horseName.toLowerCase(), horse);
      }
    }

    // Update only the horses that had races added
    for (const summary of updatedSummaries) {
      if (summary && summary.horse) {
        const horseKey = summary.horse.toLowerCase();
        if (horsesWithNewRaces.has(horseKey)) {
          existingHorseMap.set(horseKey, summary);
        }
      }
    }

    // Also add any new horses that weren't in the original data
    for (const summary of updatedSummaries) {
      if (summary && summary.horse) {
        const horseKey = summary.horse.toLowerCase();
        if (!existingHorseMap.has(horseKey)) {
          existingHorseMap.set(horseKey, summary);
        }
      }
    }

    sessionData.horseData = Array.from(existingHorseMap.values());
    sessionData.updatedAt = new Date().toISOString();

    // Save session
    await saveSession(
      sessionData.id,
      sessionData.fileName,
      sessionData.horseData,
      sessionData.allHorseDetailData
    );

    res.json({
      success: true,
      savedCount: savedRaces.length,
      skippedCount: skippedRaces.length,
      savedRaces,
      skippedRaces
    });

  } catch (error) {
    console.error('Error saving race data:', error);
    res.status(500).json({ error: 'Failed to save race data: ' + error.message });
  }
});

// Helper to rebuild horse summary from detail data
function buildHorseSummaryFromDetailData(allHorseDetailData) {
  const summaryData = [];

  for (const [horseName, entries] of Object.entries(allHorseDetailData)) {
    if (!entries || entries.length === 0) continue;

    // Get the most recent entry
    const sortedEntries = [...entries].sort((a, b) => {
      const dateA = new Date(a.date);
      const dateB = new Date(b.date);
      return dateB - dateA;
    });

    const latestEntry = sortedEntries[0];

    // Find best times (excluding races for some metrics)
    const workEntries = entries.filter(e => !e.isRace);

    let best1f = '-', best5f = '-', fastRecovery = '-', recovery15 = '-';

    for (const entry of workEntries) {
      if (entry.best1f && entry.best1f !== '-') {
        if (best1f === '-' || parseFloat(entry.best1f) < parseFloat(best1f)) {
          best1f = entry.best1f;
        }
      }
      if (entry.best5f && entry.best5f !== '-') {
        if (best5f === '-' || parseFloat(entry.best5f) < parseFloat(best5f)) {
          best5f = entry.best5f;
        }
      }
    }

    summaryData.push({
      horse: horseName,
      displayName: horseName,
      owner: '-',
      country: '-',
      lastTrainingDate: latestEntry.date,
      age: latestEntry.age || '-',
      best1f,
      best5f,
      fastRecovery,
      recovery15min: recovery15
    });
  }

  return summaryData;
}

// Export for Vercel
module.exports = app;
