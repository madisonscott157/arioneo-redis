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
async function saveSession(sessionId, fileName, horseData, allHorseDetailData) {
  const sessionData = {
    id: sessionId,
    fileName: fileName,
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    horseData: horseData,
    allHorseDetailData: allHorseDetailData
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
      allHorseDetailData: sessionData.allHorseDetailData
    });
  } catch (error) {
    console.error('Error retrieving session:', error);
    return res.status(500).json({ error: 'Failed to retrieve session' });
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
