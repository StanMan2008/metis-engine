/**
 * ═══════════════════════════════════════════════════════════════════════════
 * METIS — VS CODE QUICK START
 * 
 * Complete setup file for running METIS with Express.js backend
 * Ready to go in VS Code tomorrow morning
 * 
 * Instructions:
 * 1. Copy this file to your project root as 'server.js'
 * 2. Copy METIS_ENGINE_v4.0_COMPLETE.js to ./src/
 * 3. Copy Excel file to ./data/metis_database.xlsx
 * 4. Run: npm install && node server.js
 * ═══════════════════════════════════════════════════════════════════════════
 */

// ═══════════════════════════════════════════════════════════════════════════
// SETUP DEPENDENCIES
// ═══════════════════════════════════════════════════════════════════════════

// npm install express exceljs cors dotenv body-parser

import express from 'express';
import cors from 'cors';
import bodyParser from 'body-parser';
import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

// Import METIS engine
import { 
  METIS, 
  METIS_CONFIG, 
  METIS_SCHEMAS 
} from './src/METIS_ENGINE_v4.0_COMPLETE.js';

// ═══════════════════════════════════════════════════════════════════════════
// CONSTANTS
// ═══════════════════════════════════════════════════════════════════════════

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const PORT = process.env.PORT || 3000;
const EXCEL_PATH = path.join(__dirname, 'data', 'metis_database.xlsx');
const FLUTTER_API_PREFIX = '/api/v1';

// ═══════════════════════════════════════════════════════════════════════════
// EXPRESS APP SETUP
// ═══════════════════════════════════════════════════════════════════════════

const app = express();

// Middleware
app.use(cors());
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));

// Logging middleware
app.use((req, res, next) => {
  console.log(`[${new Date().toISOString()}] ${req.method} ${req.path}`);
  next();
});

// ═══════════════════════════════════════════════════════════════════════════
// INITIALIZE METIS ENGINE
// ═══════════════════════════════════════════════════════════════════════════

let metisEngine = null;
let isReady = false;

async function initializeMetis() {
  try {
    console.log('🔄 Loading Excel workbook...');
    
    if (!fs.existsSync(EXCEL_PATH)) {
      throw new Error(`Excel file not found at ${EXCEL_PATH}`);
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(EXCEL_PATH);

    console.log('✅ Excel workbook loaded');
    console.log('🚀 Initializing METIS engine...');

    metisEngine = new METIS(workbook);
    isReady = true;

    console.log('✅ METIS Engine ready!');
    console.log(`📊 Configuration:
      - Version: ${METIS_CONFIG.version}
      - Excel sheets: ${Object.values(METIS_CONFIG.EXCEL_SHEETS).join(', ')}
      - Windows: ${JSON.stringify(METIS_CONFIG.WINDOWS)}
    `);
  } catch (error) {
    console.error('❌ Error initializing METIS:', error);
    process.exit(1);
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// HEALTH CHECK
// ═══════════════════════════════════════════════════════════════════════════

app.get('/', (req, res) => {
  if (!isReady) {
    return res.status(503).json({
      status: 'initializing',
      message: 'METIS engine is initializing...'
    });
  }

  res.json({
    status: 'ok',
    service: 'METIS Statistical Engine',
    version: METIS_CONFIG.version,
    buildDate: METIS_CONFIG.buildDate,
    timestamp: new Date().toISOString(),
    docs: `${FLUTTER_API_PREFIX}/docs`
  });
});

app.get('/health', (req, res) => {
  res.json({
    status: isReady ? 'healthy' : 'initializing',
    metis_ready: isReady
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// API ENDPOINTS — TEAMS
// ═══════════════════════════════════════════════════════════════════════════

/**
 * GET /api/v1/team/:teamId/stats
 * Get team statistics (matches, goals, form, streaks, etc)
 */
app.get(`${FLUTTER_API_PREFIX}/team/:teamId/stats`, (req, res) => {
  try {
    const { teamId } = req.params;
    const { competition_id = 'COMP_italian-serie-a', season = '2024-2025' } = req.query;

    const result = metisEngine.getTeamStats(teamId, competition_id, season);
    res.json(result);
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

/**
 * GET /api/v1/team/:teamId/metis-index
 * Get team Metis Index (0-100) with component breakdown
 */
app.get(`${FLUTTER_API_PREFIX}/team/:teamId/metis-index`, (req, res) => {
  try {
    const { teamId } = req.params;
    const { competition_id = 'COMP_italian-serie-a', season = '2024-2025' } = req.query;

    const result = metisEngine.getTeamMetisIndex(teamId, competition_id, season);
    res.json(result);
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

/**
 * GET /api/v1/team/:teamId/insights
 * Get automatic insights about the team
 */
app.get(`${FLUTTER_API_PREFIX}/team/:teamId/insights`, (req, res) => {
  try {
    const { teamId } = req.params;
    const { competition_id = 'COMP_italian-serie-a', season = '2024-2025' } = req.query;

    const result = metisEngine.getTeamInsights(teamId, competition_id, season);
    res.json({ status: 'success', data: result });
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

/**
 * GET /api/v1/team/:teamId/full
 * Get complete team profile (stats + index + insights)
 */
app.get(`${FLUTTER_API_PREFIX}/team/:teamId/full`, (req, res) => {
  try {
    const { teamId } = req.params;
    const { competition_id = 'COMP_italian-serie-a', season = '2024-2025' } = req.query;

    const stats = metisEngine.getTeamStats(teamId, competition_id, season);
    const index = metisEngine.getTeamMetisIndex(teamId, competition_id, season);
    const insights = metisEngine.getTeamInsights(teamId, competition_id, season);

    res.json({
      status: 'success',
      data: {
        stats: stats.data,
        index: index.data,
        insights: insights.data
      }
    });
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

// ═══════════════════════════════════════════════════════════════════════════
// API ENDPOINTS — MATCHES
// ═══════════════════════════════════════════════════════════════════════════

/**
 * GET /api/v1/match/:homeTeamId/:awayTeamId/prediction
 * Get pre-match analysis and win probability
 */
app.get(`${FLUTTER_API_PREFIX}/match/:homeTeamId/:awayTeamId/prediction`, (req, res) => {
  try {
    const { homeTeamId, awayTeamId } = req.params;
    const { competition_id = 'COMP_italian-serie-a', season = '2024-2025' } = req.query;

    const result = metisEngine.getMatchPrediction(homeTeamId, awayTeamId, competition_id, season);
    res.json(result);
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

/**
 * GET /api/v1/match/:homeTeamId/:awayTeamId/insights
 * Get match-specific insights
 */
app.get(`${FLUTTER_API_PREFIX}/match/:homeTeamId/:awayTeamId/insights`, (req, res) => {
  try {
    const { homeTeamId, awayTeamId } = req.params;
    const { competition_id = 'COMP_italian-serie-a', season = '2024-2025' } = req.query;

    const result = metisEngine.getMatchInsights(homeTeamId, awayTeamId, competition_id, season);
    res.json({ status: 'success', data: result });
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

/**
 * GET /api/v1/match/:homeTeamId/:awayTeamId/full
 * Get complete match analysis (prediction + insights)
 */
app.get(`${FLUTTER_API_PREFIX}/match/:homeTeamId/:awayTeamId/full`, (req, res) => {
  try {
    const { homeTeamId, awayTeamId } = req.params;
    const { competition_id = 'COMP_italian-serie-a', season = '2024-2025' } = req.query;

    const prediction = metisEngine.getMatchPrediction(homeTeamId, awayTeamId, competition_id, season);
    const insights = metisEngine.getMatchInsights(homeTeamId, awayTeamId, competition_id, season);

    res.json({
      status: 'success',
      data: {
        prediction: prediction.data,
        insights: insights.data
      }
    });
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

// ═══════════════════════════════════════════════════════════════════════════
// API ENDPOINTS — PLAYERS
// ═══════════════════════════════════════════════════════════════════════════

/**
 * GET /api/v1/player/:playerId/rating
 * Get player Metis Rating (0-100)
 */
app.get(`${FLUTTER_API_PREFIX}/player/:playerId/rating`, (req, res) => {
  try {
    const { playerId } = req.params;
    const { competition_id = 'COMP_italian-serie-a', season = '2024-2025' } = req.query;

    const result = metisEngine.getPlayerRating(playerId, competition_id, season);
    res.json(result);
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

// ═══════════════════════════════════════════════════════════════════════════
// API ENDPOINTS — LEAGUE
// ═══════════════════════════════════════════════════════════════════════════

/**
 * GET /api/v1/league/:competitionId/standings
 * Get league standings
 */
app.get(`${FLUTTER_API_PREFIX}/league/:competitionId/standings`, (req, res) => {
  try {
    const { competitionId } = req.params;
    const { season = '2024-2025' } = req.query;

    const result = metisEngine.getLeagueStandings(competitionId, season);
    res.json(result);
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

/**
 * GET /api/v1/league/:competitionId/stats
 * Get league-wide statistics
 */
app.get(`${FLUTTER_API_PREFIX}/league/:competitionId/stats`, (req, res) => {
  try {
    const { competitionId } = req.params;
    const { season = '2024-2025' } = req.query;

    const result = metisEngine.api.getLeagueStats(competitionId, season);
    res.json(result);
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

/**
 * GET /api/v1/league/:competitionId/dashboard
 * Get league dashboard (standings + top scorers + stats)
 */
app.get(`${FLUTTER_API_PREFIX}/league/:competitionId/dashboard`, (req, res) => {
  try {
    const { competitionId } = req.params;
    const { season = '2024-2025' } = req.query;

    const result = metisEngine.getDashboard(competitionId, season);
    res.json(result);
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

// ═══════════════════════════════════════════════════════════════════════════
// API ENDPOINTS — EXPORT
// ═══════════════════════════════════════════════════════════════════════════

/**
 * GET /api/v1/export/json
 * Export all data as JSON (for offline use or Flutter local sync)
 */
app.get(`${FLUTTER_API_PREFIX}/export/json`, (req, res) => {
  try {
    const result = metisEngine.exportAsJSON();
    res.json({ status: 'success', data: result });
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

/**
 * GET /api/v1/config
 * Get METIS configuration
 */
app.get(`${FLUTTER_API_PREFIX}/config`, (req, res) => {
  res.json({
    status: 'success',
    data: {
      version: METIS_CONFIG.version,
      sheets: METIS_CONFIG.EXCEL_SHEETS,
      windows: METIS_CONFIG.WINDOWS,
      metis_weights: METIS_CONFIG.METIS_WEIGHTS,
      match_status: METIS_CONFIG.MATCH_STATUS,
      thresholds: METIS_CONFIG.THRESHOLDS
    }
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// API DOCUMENTATION
// ═══════════════════════════════════════════════════════════════════════════

app.get(`${FLUTTER_API_PREFIX}/docs`, (req, res) => {
  const docs = `
    ╔═══════════════════════════════════════════════════════════════╗
    ║           METIS STATISTICAL ENGINE — API DOCUMENTATION       ║
    ╚═══════════════════════════════════════════════════════════════╝

    📊 TEAM ENDPOINTS
    ─────────────────────────────────────────────────────────────

    GET /api/v1/team/:teamId/stats
      → Team statistics (matches, goals, form, streaks)
      Query: ?competition_id=COMP_italian-serie-a&season=2024-2025

    GET /api/v1/team/:teamId/metis-index
      → Team Metis Index (0-100) with component breakdown
      Query: ?competition_id=COMP_italian-serie-a&season=2024-2025

    GET /api/v1/team/:teamId/insights
      → Automatic insights (attack, defense, form)
      Query: ?competition_id=COMP_italian-serie-a&season=2024-2025

    GET /api/v1/team/:teamId/full
      → Complete team profile (stats + index + insights)
      Query: ?competition_id=COMP_italian-serie-a&season=2024-2025

    ⚽ MATCH ENDPOINTS
    ─────────────────────────────────────────────────────────────

    GET /api/v1/match/:homeTeamId/:awayTeamId/prediction
      → Pre-match analysis (win probability, xG, risk factors)
      Query: ?competition_id=COMP_italian-serie-a&season=2024-2025

    GET /api/v1/match/:homeTeamId/:awayTeamId/insights
      → Match-specific insights
      Query: ?competition_id=COMP_italian-serie-a&season=2024-2025

    GET /api/v1/match/:homeTeamId/:awayTeamId/full
      → Complete match analysis
      Query: ?competition_id=COMP_italian-serie-a&season=2024-2025

    👤 PLAYER ENDPOINTS
    ─────────────────────────────────────────────────────────────

    GET /api/v1/player/:playerId/rating
      → Player Metis Rating (0-100)
      Query: ?competition_id=COMP_italian-serie-a&season=2024-2025

    🏆 LEAGUE ENDPOINTS
    ─────────────────────────────────────────────────────────────

    GET /api/v1/league/:competitionId/standings
      → League table standings
      Query: ?season=2024-2025

    GET /api/v1/league/:competitionId/stats
      → League-wide statistics (goals/match, BTTS%, over/under)
      Query: ?season=2024-2025

    GET /api/v1/league/:competitionId/dashboard
      → Dashboard (standings + top scorers + stats)
      Query: ?season=2024-2025

    📤 EXPORT ENDPOINTS
    ─────────────────────────────────────────────────────────────

    GET /api/v1/export/json
      → Export all data as JSON

    GET /api/v1/config
      → Get METIS configuration

    ═════════════════════════════════════════════════════════════════

    EXAMPLE REQUESTS:

    curl http://localhost:3000/api/v1/team/TEAM_ac-milan/stats
    curl "http://localhost:3000/api/v1/match/TEAM_ac-milan/TEAM_inter/prediction"
    curl http://localhost:3000/api/v1/league/COMP_italian-serie-a/dashboard
    curl http://localhost:3000/api/v1/player/PLAYER_rafael-leao/rating

    ═════════════════════════════════════════════════════════════════
  `;

  res.type('text/plain').send(docs);
});

// ═══════════════════════════════════════════════════════════════════════════
// ERROR HANDLING
// ═══════════════════════════════════════════════════════════════════════════

app.use((err, req, res, next) => {
  console.error('❌ Error:', err);
  res.status(500).json({
    error: 'Internal server error',
    message: err.message
  });
});

app.use((req, res) => {
  res.status(404).json({
    error: 'Not found',
    path: req.path,
    docs: `${FLUTTER_API_PREFIX}/docs`
  });
});

// ═══════════════════════════════════════════════════════════════════════════
// SERVER START
// ═══════════════════════════════════════════════════════════════════════════

async function start() {
  console.log(`
╔════════════════════════════════════════════════════════════════╗
║                                                                ║
║     🚀 METIS STATISTICAL ENGINE v4.0                          ║
║     The Perfect Machine for Football Statistics               ║
║                                                                ║
╚════════════════════════════════════════════════════════════════╝
  `);

  // Initialize METIS
  await initializeMetis();

  // Start server
  app.listen(PORT, () => {
    console.log(`
✅ Server running on http://localhost:${PORT}

📚 Documentation:     http://localhost:${PORT}/api/v1/docs
🏥 Health check:     http://localhost:${PORT}/health
📊 Dashboard:        http://localhost:${PORT}/api/v1/league/COMP_italian-serie-a/dashboard

Press Ctrl+C to stop

    `);
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// GO!
// ═══════════════════════════════════════════════════════════════════════════

start().catch(error => {
  console.error('❌ Failed to start server:', error);
  process.exit(1);
});

export default app;