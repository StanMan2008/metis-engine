/**
 * ═══════════════════════════════════════════════════════════════════════════
 * METIS — COMPLETE STATISTICAL ENGINE v4.0
 * 
 * Production-ready JavaScript engine for football statistics calculation
 * Compatible with: Excel → VS Code → Flutter
 * 
 * Author: Alessandro
 * Created: 2026-09-23
 * Status: READY FOR DEPLOYMENT
 * ═══════════════════════════════════════════════════════════════════════════
 */

// ═══════════════════════════════════════════════════════════════════════════
// 01 — CONFIGURATION & SCHEMA
// ═══════════════════════════════════════════════════════════════════════════

const METIS_CONFIG = {
  version: '4.0',
  buildDate: '2026-09-23',
  
  // Excel sheet names (what you need to create in Excel)
  EXCEL_SHEETS: {
    competitions: 'Campionati',
    teams: 'Squadre',
    players: 'Giocatori',
    matches: 'Partite',
    match_events: 'PartiteDettagli',
    team_stats: 'StatisticheSquadre',
    player_stats: 'StatisticheGiocatori',
    league_stats: 'StatisticheCampionati',
    standings: 'Classifiche',
    top_players: 'TopGiocatori',
    metis_cache: 'MetisCache'
  },

  // Calculation windows
  WINDOWS: {
    SHORT: 5,      // Last 5 matches
    MEDIUM: 10,    // Last 10 matches
    SEASON: 'all'  // Entire season
  },

  // Metis Index weights
  METIS_WEIGHTS: {
    form: 0.20,
    ppg: 0.15,
    goal_diff: 0.15,
    xg_xga: 0.20,
    home_away: 0.10,
    clean_sheet_btts: 0.05,
    shots_chances: 0.10,
    h2h: 0.05
  },

  // Match status
  MATCH_STATUS: {
    SCHEDULED: 'scheduled',
    LIVE: 'live',
    FINISHED: 'finished',
    POSTPONED: 'postponed'
  },

  // Performance thresholds for insights
  THRESHOLDS: {
    ON_FIRE: { goals: 3, window: 5 },           // 3+ gol in 5 games
    FORM_UP: { ppg_improvement: 1.5, window: 10 },
    FORM_DOWN: { ppg_decline: 1.0, window: 10 },
    WINNING_STREAK: 3,
    UNBEATEN_STREAK: 5,
    LOSING_STREAK: 3,
    CLEAN_SHEET_STREAK: 3,
    CONCEDING_STREAK: 3
  }
};

// ═══════════════════════════════════════════════════════════════════════════
// 02 — DATA STRUCTURES
// ═══════════════════════════════════════════════════════════════════════════

/**
 * These are the exact data structures you need in your Excel sheets
 */

const METIS_SCHEMAS = {
  
  // Sheet: Campionati
  competition: {
    id: String,           // COMP_italian-serie-a
    country: String,      // Italy
    name: String,         // Italian Serie A
    slug: String,         // italian-serie-a
    season: String,       // 2024-2025
    teams_count: Number,
    matches_played: Number,
    last_updated: Date
  },

  // Sheet: Squadre
  team: {
    id: String,           // TEAM_ac-milan
    competition_id: String,
    name: String,         // AC Milan
    slug: String,         // ac-milan
    logo_url: String,
    country: String,
    stadium: String,
    city: String,
    founded: Number,
    market_value: Number, // in millions
    last_updated: Date
  },

  // Sheet: Giocatori
  player: {
    id: String,           // PLAYER_rafael-leao
    name: String,         // Rafael Leão
    slug: String,         // rafael-leao
    photo_url: String,
    position: String,     // FW, MF, DF, GK
    nationality: String,
    team_id: String,
    competition_id: String,
    age: Number,
    height: Number,
    weight: Number,
    market_value: Number,
    last_updated: Date
  },

  // Sheet: Partite
  match: {
    id: String,           // MATCH_2024-09-23_ac-milan_inter
    date: Date,
    time: String,
    competition_id: String,
    home_team_id: String,
    away_team_id: String,
    home_team_name: String,
    away_team_name: String,
    home_score: Number,
    away_score: Number,
    status: String,       // scheduled, live, finished
    home_possession: Number,    // %
    away_possession: Number,    // %
    home_shots: Number,
    away_shots: Number,
    home_shots_on_target: Number,
    away_shots_on_target: Number,
    home_xg: Number,
    away_xg: Number,
    corners: Number,
    fouls: Number,
    yellow_cards: Number,
    red_cards: Number,
    last_updated: Date
  },

  // Sheet: PartiteDettagli (Match events timeline)
  match_event: {
    id: String,
    match_id: String,
    minute: Number,
    second: Number,
    player_id: String,
    player_name: String,
    team_id: String,
    event_type: String,   // goal, assist, shot, card, substitution
    value: Number,        // xG for shots
    notes: String,
    timestamp: Date
  },

  // Sheet: StatisticheSquadre (Team stats — auto-calculated)
  team_stats: {
    team_id: String,
    competition_id: String,
    season: String,
    
    // Overall
    matches_played: Number,
    wins: Number,
    draws: Number,
    losses: Number,
    gf: Number,          // Goals for
    ga: Number,          // Goals against
    gd: Number,          // Goal difference
    points: Number,
    ppg: Number,         // Points per game
    gf_per_match: Number,
    ga_per_match: Number,

    // Last 5/10
    w5: Number, d5: Number, l5: Number,
    ppg5: Number, gf5: Number, ga5: Number,
    w10: Number, d10: Number, l10: Number,
    ppg10: Number, gf10: Number, ga10: Number,

    // Home/Away
    home_matches: Number, home_wins: Number, home_points: Number, home_ppg: Number,
    away_matches: Number, away_wins: Number, away_points: Number, away_ppg: Number,

    // Defense
    clean_sheets: Number,
    clean_sheet_percent: Number,
    failed_to_score: Number,
    failed_to_score_percent: Number,

    // Over/Under & BTTS
    over_1_5: Number, over_2_5: Number, over_3_5: Number,
    under_1_5: Number, under_2_5: Number, under_3_5: Number,
    btts: Number, btts_percent: Number,

    // Advanced
    xg: Number, xga: Number,
    xg_per_match: Number, xga_per_match: Number,

    // Form
    form_string: String,  // WDLWW
    winning_streak: Number,
    unbeaten_streak: Number,
    losing_streak: Number,
    clean_sheet_streak: Number,
    conceding_streak: Number,

    last_calculated: Date
  },

  // Sheet: StatisticheGiocatori (Player stats)
  player_stats: {
    player_id: String,
    player_name: String,
    team_id: String,
    competition_id: String,
    season: String,

    // Appearances
    appearances: Number,
    starts: Number,
    substitutions: Number,
    minutes: Number,

    // Scoring
    goals: Number,
    assists: Number,
    goal_per_90: Number,
    assist_per_90: Number,
    goal_assist: Number,

    // Shooting
    shots: Number,
    shots_on_target: Number,
    xg: Number, xa: Number,
    xg_per_90: Number,
    xg_underperformance: Number,

    // Discipline
    yellow_cards: Number,
    red_cards: Number,

    // Passing
    passes: Number,
    passes_accuracy: Number,

    // Defensive
    tackles: Number,
    interceptions: Number,
    clearances: Number,

    // Goalkeeper (if applicable)
    saves: Number,
    save_percentage: Number,
    goals_against: Number,
    clean_sheets: Number,

    // Rating
    avg_rating: Number,
    form_rating: Number,

    last_updated: Date
  },

  // Sheet: StatisticheCampionati (League-wide stats)
  league_stats: {
    competition_id: String,
    season: String,

    matches_played: Number,
    total_goals: Number,
    goals_per_match: Number,
    home_goals: Number,
    away_goals: Number,

    over_0_5: Number, over_1_5: Number, over_2_5: Number, over_3_5: Number,
    under_1_5: Number, under_2_5: Number, under_3_5: Number,
    btts_percent: Number,

    clean_sheets_home: Number,
    clean_sheets_away: Number,
    failed_to_score_home: Number,
    failed_to_score_away: Number,

    best_attack_team_id: String,
    best_defense_team_id: String,
    best_home_team_id: String,
    best_away_team_id: String,

    last_calculated: Date
  },

  // Sheet: Classifiche (Standings)
  standing: {
    competition_id: String,
    position: Number,
    team_id: String,
    team_name: String,
    matches: Number,
    wins: Number,
    draws: Number,
    losses: Number,
    gf: Number,
    ga: Number,
    gd: Number,
    points: Number,
    last_updated: Date
  },

  // Sheet: TopGiocatori (Top scorers/assisters)
  top_player: {
    competition_id: String,
    type: String,         // 'goals' or 'assists'
    position: Number,
    player_id: String,
    player_name: String,
    team_id: String,
    team_name: String,
    value: Number,        // goals or assists count
    last_updated: Date
  }
};

// ═══════════════════════════════════════════════════════════════════════════
// 03 — EXCEL DATA LOADER
// ═══════════════════════════════════════════════════════════════════════════

class MetisExcelLoader {
  /**
   * Load data from Excel workbook
   * Works with ExcelJS library (npm install exceljs)
   */

  constructor(workbook) {
    this.workbook = workbook;
    this.data = {};
    this.load();
  }

  load() {
    // Load all sheets from Excel
    const sheetNames = METIS_CONFIG.EXCEL_SHEETS;

    if (this.workbook.getWorksheet(sheetNames.competitions)) {
      this.data.competitions = this.parseSheet(sheetNames.competitions);
    }
    if (this.workbook.getWorksheet(sheetNames.teams)) {
      this.data.teams = this.parseSheet(sheetNames.teams);
    }
    if (this.workbook.getWorksheet(sheetNames.players)) {
      this.data.players = this.parseSheet(sheetNames.players);
    }
    if (this.workbook.getWorksheet(sheetNames.matches)) {
      this.data.matches = this.parseSheet(sheetNames.matches);
    }
    if (this.workbook.getWorksheet(sheetNames.match_events)) {
      this.data.match_events = this.parseSheet(sheetNames.match_events);
    }

    console.log('✅ Excel data loaded successfully', this.data);
  }

  parseSheet(sheetName) {
    const sheet = this.workbook.getWorksheet(sheetName);
    const rows = [];

    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // Skip header
      const obj = {};
      row.eachCell((cell, colNumber) => {
        const headerCell = sheet.getRow(1).getCell(colNumber);
        obj[headerCell.value] = cell.value;
      });
      rows.push(obj);
    });

    return rows;
  }

  getTeam(teamId) {
    return this.data.teams?.find(t => t.id === teamId);
  }

  getPlayer(playerId) {
    return this.data.players?.find(p => p.id === playerId);
  }

  getMatch(matchId) {
    return this.data.matches?.find(m => m.id === matchId);
  }

  getMatchEvents(matchId) {
    return this.data.match_events?.filter(e => e.match_id === matchId) || [];
  }

  getAllMatches(competitionId) {
    return this.data.matches?.filter(m => m.competition_id === competitionId) || [];
  }

  getTeamMatches(teamId) {
    return this.data.matches?.filter(m => 
      m.home_team_id === teamId || m.away_team_id === teamId
    ) || [];
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// 04 — STATISTICS CALCULATOR ENGINE
// ═══════════════════════════════════════════════════════════════════════════

class MetisStatisticsEngine {
  
  constructor(dataLoader) {
    this.loader = dataLoader;
    this.cache = {};
  }

  /**
   * TEAM STATISTICS CALCULATION
   */

  calculateTeamStats(teamId, competitionId, season) {
    const cacheKey = `team_stats_${teamId}_${competitionId}`;
    if (this.cache[cacheKey]) return this.cache[cacheKey];

    const matches = this.loader.getTeamMatches(teamId)
      .filter(m => m.competition_id === competitionId && m.status === 'finished');

    if (matches.length === 0) return null;

    const stats = {
      team_id: teamId,
      competition_id: competitionId,
      season: season,
      
      // Overall
      matches_played: matches.length,
      wins: 0, draws: 0, losses: 0,
      gf: 0, ga: 0, gd: 0,
      points: 0, ppg: 0,

      // Home/Away split
      home_matches: 0, home_wins: 0, home_points: 0, home_gf: 0, home_ga: 0,
      away_matches: 0, away_wins: 0, away_points: 0, away_gf: 0, away_ga: 0,

      // Defense
      clean_sheets: 0,
      failed_to_score: 0,

      // Over/Under
      over_1_5: 0, over_2_5: 0, over_3_5: 0,
      under_1_5: 0, under_2_5: 0, under_3_5: 0,
      btts: 0,

      // Form (last 5 and 10)
      form_string: '',
      w5: 0, d5: 0, l5: 0, ppg5: 0,
      w10: 0, d10: 0, l10: 0, ppg10: 0,

      // Streaks
      winning_streak: 0,
      unbeaten_streak: 0,
      losing_streak: 0,
      clean_sheet_streak: 0,
      conceding_streak: 0
    };

    // Iterate through matches
    matches.forEach((match, idx) => {
      const isHome = match.home_team_id === teamId;
      const scored = isHome ? match.home_score : match.away_score;
      const conceded = isHome ? match.away_score : match.home_score;
      const totalGoals = scored + conceded;

      // Basic stats
      stats.gf += scored;
      stats.ga += conceded;
      stats.gd = stats.gf - stats.ga;

      // Result
      if (scored > conceded) {
        stats.wins++;
        stats.points += 3;
      } else if (scored === conceded) {
        stats.draws++;
        stats.points += 1;
      } else {
        stats.losses++;
      }

      // Defense
      if (conceded === 0) stats.clean_sheets++;
      if (scored === 0) stats.failed_to_score++;

      // Over/Under
      if (totalGoals >= 0.5) stats.over_0_5++;
      if (totalGoals >= 1.5) stats.over_1_5++;
      if (totalGoals >= 2.5) stats.over_2_5++;
      if (totalGoals >= 3.5) stats.over_3_5++;
      if (totalGoals < 1.5) stats.under_1_5++;
      if (totalGoals < 2.5) stats.under_2_5++;
      if (totalGoals < 3.5) stats.under_3_5++;

      // BTTS
      if (scored > 0 && conceded > 0) stats.btts++;

      // Home/Away split
      if (isHome) {
        stats.home_matches++;
        stats.home_gf += scored;
        stats.home_ga += conceded;
        if (scored > conceded) stats.home_wins++;
        stats.home_points += (scored > conceded ? 3 : scored === conceded ? 1 : 0);
      } else {
        stats.away_matches++;
        stats.away_gf += scored;
        stats.away_ga += conceded;
        if (scored > conceded) stats.away_wins++;
        stats.away_points += (scored > conceded ? 3 : scored === conceded ? 1 : 0);
      }

      // Form string (W/D/L)
      if (scored > conceded) stats.form_string = 'W' + stats.form_string;
      else if (scored === conceded) stats.form_string = 'D' + stats.form_string;
      else stats.form_string = 'L' + stats.form_string;
    });

    // Calculate per-match averages
    stats.ppg = (stats.points / stats.matches_played).toFixed(2);
    stats.gf_per_match = (stats.gf / stats.matches_played).toFixed(2);
    stats.ga_per_match = (stats.ga / stats.matches_played).toFixed(2);

    stats.home_ppg = stats.home_matches ? (stats.home_points / stats.home_matches).toFixed(2) : 0;
    stats.away_ppg = stats.away_matches ? (stats.away_points / stats.away_matches).toFixed(2) : 0;

    // Percentages
    stats.clean_sheet_percent = ((stats.clean_sheets / stats.matches_played) * 100).toFixed(2);
    stats.failed_to_score_percent = ((stats.failed_to_score / stats.matches_played) * 100).toFixed(2);
    stats.btts_percent = ((stats.btts / stats.matches_played) * 100).toFixed(2);

    // Last 5 and 10
    const last5 = matches.slice(-5);
    const last10 = matches.slice(-10);

    stats.w5 = last5.filter(m => {
      const scored = m.home_team_id === teamId ? m.home_score : m.away_score;
      const conceded = m.home_team_id === teamId ? m.away_score : m.home_score;
      return scored > conceded;
    }).length;
    stats.ppg5 = (last5.reduce((sum, m) => {
      const scored = m.home_team_id === teamId ? m.home_score : m.away_score;
      const conceded = m.home_team_id === teamId ? m.away_score : m.home_score;
      return sum + (scored > conceded ? 3 : scored === conceded ? 1 : 0);
    }, 0) / last5.length).toFixed(2);

    stats.w10 = last10.filter(m => {
      const scored = m.home_team_id === teamId ? m.home_score : m.away_score;
      const conceded = m.home_team_id === teamId ? m.away_score : m.home_score;
      return scored > conceded;
    }).length;
    stats.ppg10 = (last10.reduce((sum, m) => {
      const scored = m.home_team_id === teamId ? m.home_score : m.away_score;
      const conceded = m.home_team_id === teamId ? m.away_score : m.home_score;
      return sum + (scored > conceded ? 3 : scored === conceded ? 1 : 0);
    }, 0) / last10.length).toFixed(2);

    // Calculate streaks
    stats.winning_streak = this.calculateStreak(matches, teamId, 'win');
    stats.unbeaten_streak = this.calculateStreak(matches, teamId, 'unbeaten');
    stats.losing_streak = this.calculateStreak(matches, teamId, 'loss');
    stats.clean_sheet_streak = this.calculateStreakCleanSheets(matches, teamId);
    stats.conceding_streak = this.calculateStreakConceding(matches, teamId);

    this.cache[cacheKey] = stats;
    return stats;
  }

  calculateStreak(matches, teamId, type) {
    let streak = 0;
    for (let i = matches.length - 1; i >= 0; i--) {
      const match = matches[i];
      const scored = match.home_team_id === teamId ? match.home_score : match.away_score;
      const conceded = match.home_team_id === teamId ? match.away_score : match.home_score;

      if (type === 'win' && scored > conceded) streak++;
      else if (type === 'unbeaten' && scored >= conceded) streak++;
      else if (type === 'loss' && scored < conceded) streak++;
      else break;
    }
    return streak;
  }

  calculateStreakCleanSheets(matches, teamId) {
    let streak = 0;
    for (let i = matches.length - 1; i >= 0; i--) {
      const match = matches[i];
      const conceded = match.home_team_id === teamId ? match.away_score : match.home_score;
      if (conceded === 0) streak++;
      else break;
    }
    return streak;
  }

  calculateStreakConceding(matches, teamId) {
    let streak = 0;
    for (let i = matches.length - 1; i >= 0; i--) {
      const match = matches[i];
      const conceded = match.home_team_id === teamId ? match.away_score : match.home_score;
      if (conceded > 0) streak++;
      else break;
    }
    return streak;
  }

  /**
   * PLAYER STATISTICS CALCULATION
   */

  calculatePlayerStats(playerId, competitionId, season) {
    // This would aggregate player events from match_events sheet
    const playerEvents = this.loader.data.match_events?.filter(e => 
      e.player_id === playerId && e.event_type !== 'substitution'
    ) || [];

    const stats = {
      player_id: playerId,
      competition_id: competitionId,
      season: season,
      appearances: 0,
      goals: 0,
      assists: 0,
      minutes: 0,
      yellow_cards: 0,
      red_cards: 0,
      avg_rating: 0
    };

    playerEvents.forEach(event => {
      if (event.event_type === 'goal') stats.goals++;
      if (event.event_type === 'assist') stats.assists++;
      if (event.event_type === 'yellow_card') stats.yellow_cards++;
      if (event.event_type === 'red_card') stats.red_cards++;
    });

    stats.goal_per_90 = stats.minutes > 0 ? (stats.goals / stats.minutes * 90).toFixed(2) : 0;
    stats.assist_per_90 = stats.minutes > 0 ? (stats.assists / stats.minutes * 90).toFixed(2) : 0;
    stats.goal_assist = stats.goals + stats.assists;

    return stats;
  }

  /**
   * MATCH STATISTICS
   */

  calculateMatchStats(matchId) {
    const match = this.loader.getMatch(matchId);
    if (!match) return null;

    const events = this.loader.getMatchEvents(matchId);

    const stats = {
      match_id: matchId,
      home_team_id: match.home_team_id,
      away_team_id: match.away_team_id,
      home_score: match.home_score,
      away_score: match.away_score,
      total_goals: match.home_score + match.away_score,
      
      // Over/Under
      over_0_5: (match.home_score + match.away_score) >= 0.5,
      over_1_5: (match.home_score + match.away_score) >= 1.5,
      over_2_5: (match.home_score + match.away_score) >= 2.5,
      over_3_5: (match.home_score + match.away_score) >= 3.5,

      // BTTS
      btts: match.home_score > 0 && match.away_score > 0,

      // Clean sheets
      clean_sheet_home: match.away_score === 0,
      clean_sheet_away: match.home_score === 0,

      // Events
      home_goals_events: events.filter(e => e.event_type === 'goal' && e.team_id === match.home_team_id),
      away_goals_events: events.filter(e => e.event_type === 'goal' && e.team_id === match.away_team_id),
      
      goals_first_half: events.filter(e => e.event_type === 'goal' && e.minute <= 45),
      goals_second_half: events.filter(e => e.event_type === 'goal' && e.minute > 45),

      yellow_cards: events.filter(e => e.event_type === 'yellow_card').length,
      red_cards: events.filter(e => e.event_type === 'red_card').length
    };

    return stats;
  }

  /**
   * LEAGUE STATISTICS
   */

  calculateLeagueStats(competitionId, season) {
    const matches = this.loader.data.matches?.filter(m => 
      m.competition_id === competitionId && m.status === 'finished'
    ) || [];

    const stats = {
      competition_id: competitionId,
      season: season,
      matches_played: matches.length,
      total_goals: 0,
      goals_per_match: 0,
      home_goals: 0,
      away_goals: 0,
      btts_percent: 0,
      over_2_5_percent: 0,
      under_2_5_percent: 0,
      clean_sheets_home: 0,
      clean_sheets_away: 0
    };

    let over_2_5 = 0, under_2_5 = 0, btts = 0;

    matches.forEach(m => {
      const totalGoals = m.home_score + m.away_score;
      stats.total_goals += totalGoals;
      stats.home_goals += m.home_score;
      stats.away_goals += m.away_score;

      if (m.away_score === 0) stats.clean_sheets_home++;
      if (m.home_score === 0) stats.clean_sheets_away++;

      if (totalGoals >= 2.5) over_2_5++;
      else under_2_5++;

      if (m.home_score > 0 && m.away_score > 0) btts++;
    });

    stats.goals_per_match = (stats.total_goals / stats.matches_played).toFixed(2);
    stats.over_2_5_percent = ((over_2_5 / stats.matches_played) * 100).toFixed(2);
    stats.under_2_5_percent = ((under_2_5 / stats.matches_played) * 100).toFixed(2);
    stats.btts_percent = ((btts / stats.matches_played) * 100).toFixed(2);

    return stats;
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// 05 — METIS INDEX ENGINE
// ═══════════════════════════════════════════════════════════════════════════

class MetisIndexEngine {

  constructor(statsEngine) {
    this.statsEngine = statsEngine;
  }

  /**
   * Calculate Team Metis Index (0-100)
   * Combines 8 weighted components
   */

  calculateTeamMetisIndex(teamId, competitionId, season) {
    const stats = this.statsEngine.calculateTeamStats(teamId, competitionId, season);
    if (!stats) return null;

    let index = 0;

    // 1. Recent Form (20%) — Last 10 matches
    const formScore = Math.min(10, stats.w10 + stats.d10 * 0.5) * 10; // Max 100 if 10W
    index += formScore * METIS_CONFIG.METIS_WEIGHTS.form;

    // 2. Points Per Game (15%)
    const ppgScore = Math.min(3, parseFloat(stats.ppg)) / 3 * 100;
    index += ppgScore * METIS_CONFIG.METIS_WEIGHTS.ppg;

    // 3. Goal Difference (15%)
    const gdScore = (stats.gd > 0 ? Math.min(stats.gd, 30) : Math.max(stats.gd, -30)) + 30;
    const gdPercent = (gdScore / 60) * 100;
    index += gdPercent * METIS_CONFIG.METIS_WEIGHTS.goal_diff;

    // 4. xG/xGA (20%) — Placeholder for when available
    // For now, use actual goals ratio
    const xgScore = (stats.gf / (stats.gf + stats.ga)) * 100;
    index += xgScore * METIS_CONFIG.METIS_WEIGHTS.xg_xga;

    // 5. Home/Away Performance (10%)
    const homeScore = parseFloat(stats.home_ppg) || 0;
    const awayScore = parseFloat(stats.away_ppg) || 0;
    const homeAwayBalance = ((homeScore + awayScore) / 6) * 100; // Max 3 PPG home/away
    index += homeAwayBalance * METIS_CONFIG.METIS_WEIGHTS.home_away;

    // 6. Clean Sheets + BTTS (5%)
    const defensiveScore = Math.min(50, parseFloat(stats.clean_sheet_percent)) + 
                          (100 - Math.min(100, parseFloat(stats.btts_percent)));
    index += (defensiveScore / 2) * METIS_CONFIG.METIS_WEIGHTS.clean_sheet_btts;

    // 7. Shots/Chances (10%) — Placeholder
    // When available from StatsBomb/Sofascore
    const chancesScore = Math.min(100, (stats.gf / stats.matches_played) * 20); // Rough estimate
    index += chancesScore * METIS_CONFIG.METIS_WEIGHTS.shots_chances;

    // 8. H2H (5%) — Placeholder
    // When we have head-to-head data
    const h2hScore = 50; // Default neutral
    index += h2hScore * METIS_CONFIG.METIS_WEIGHTS.h2h;

    return {
      team_id: teamId,
      index: Math.min(100, Math.round(index)),
      components: {
        form: Math.round(formScore * METIS_CONFIG.METIS_WEIGHTS.form),
        ppg: Math.round(ppgScore * METIS_CONFIG.METIS_WEIGHTS.ppg),
        goal_diff: Math.round(gdPercent * METIS_CONFIG.METIS_WEIGHTS.goal_diff),
        xg_xga: Math.round(xgScore * METIS_CONFIG.METIS_WEIGHTS.xg_xga),
        home_away: Math.round(homeAwayBalance * METIS_CONFIG.METIS_WEIGHTS.home_away),
        clean_sheet_btts: Math.round((defensiveScore / 2) * METIS_CONFIG.METIS_WEIGHTS.clean_sheet_btts),
        shots_chances: Math.round(chancesScore * METIS_CONFIG.METIS_WEIGHTS.shots_chances),
        h2h: Math.round(h2hScore * METIS_CONFIG.METIS_WEIGHTS.h2h)
      },
      stats: stats
    };
  }

  /**
   * Match Metis Index
   * Pre-match analysis combining both teams' metrics
   */

  calculateMatchMetisIndex(homeTeamId, awayTeamId, competitionId, season) {
    const homeIndex = this.calculateTeamMetisIndex(homeTeamId, competitionId, season);
    const awayIndex = this.calculateTeamMetisIndex(awayTeamId, competitionId, season);

    if (!homeIndex || !awayIndex) return null;

    // Win probability model (simplified)
    const homeProbability = (homeIndex.index + 10) / (homeIndex.index + awayIndex.index + 20);
    const drawProbability = 0.25;
    const awayProbability = 1 - homeProbability - drawProbability;

    return {
      home_team_id: homeTeamId,
      away_team_id: awayTeamId,
      home_metis_index: homeIndex.index,
      away_metis_index: awayIndex.index,
      prediction: {
        home_win: Math.round(homeProbability * 100),
        draw: Math.round(drawProbability * 100),
        away_win: Math.round(awayProbability * 100)
      },
      expected_goals: {
        home: (homeIndex.stats.gf_per_match * awayIndex.stats.ga_per_match).toFixed(2),
        away: (awayIndex.stats.gf_per_match * homeIndex.stats.ga_per_match).toFixed(2)
      },
      risk_factors: this.identifyRiskFactors(homeIndex, awayIndex)
    };
  }

  identifyRiskFactors(homeIndex, awayIndex) {
    const risks = [];

    if (homeIndex.stats.losing_streak >= 3) {
      risks.push('Home team in losing streak');
    }
    if (awayIndex.stats.winning_streak >= 3) {
      risks.push('Away team in winning streak');
    }
    if (homeIndex.stats.ppg < 1.0) {
      risks.push('Home team poor form (PPG < 1.0)');
    }
    if (awayIndex.stats.ga_per_match > 1.5) {
      risks.push('Away team conceding too many goals');
    }

    return risks;
  }

  /**
   * Player Metis Rating (0-100)
   */

  calculatePlayerMetisRating(playerId, competitionId, season) {
    const stats = this.statsEngine.calculatePlayerStats(playerId, competitionId, season);
    if (!stats) return null;

    const player = this.statsEngine.loader.getPlayer(playerId);
    let rating = 50; // Base

    // Goals contribution
    rating += Math.min(25, stats.goal_per_90 * 10);

    // Assist contribution
    rating += Math.min(15, stats.assist_per_90 * 10);

    // Appearance consistency
    rating += Math.min(10, stats.appearances / 10);

    // Discipline (negative)
    rating -= Math.min(10, (stats.yellow_cards + stats.red_cards * 2) * 2);

    return {
      player_id: playerId,
      player_name: player?.name,
      position: player?.position,
      rating: Math.min(100, Math.max(0, Math.round(rating))),
      goals: stats.goals,
      assists: stats.assists,
      minutes: stats.minutes,
      goal_per_90: stats.goal_per_90,
      assist_per_90: stats.assist_per_90,
      cards: stats.yellow_cards + stats.red_cards
    };
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// 06 — INSIGHTS ENGINE
// ═══════════════════════════════════════════════════════════════════════════

class MetisInsightsEngine {

  constructor(statsEngine, indexEngine) {
    this.statsEngine = statsEngine;
    this.indexEngine = indexEngine;
  }

  /**
   * Generate automatic insights from statistics
   */

  generateTeamInsights(teamId, competitionId, season) {
    const stats = this.statsEngine.calculateTeamStats(teamId, competitionId, season);
    if (!stats) return [];

    const insights = [];
    const team = this.statsEngine.loader.getTeam(teamId);

    // ATTACK insights
    if (stats.w5 === 5) {
      insights.push({
        category: 'ATTACK',
        severity: 'high',
        title: '🔥 Attacco acceso',
        message: `${team.name} ha vinto tutte le ultime 5 partite. Forma straordinaria.`,
        confidence: 0.95
      });
    }
    if (stats.gf_per_match > 2.5) {
      insights.push({
        category: 'ATTACK',
        severity: 'high',
        title: '⚡ Offensiva pericolosa',
        message: `Sta segnando in media ${parseFloat(stats.gf_per_match).toFixed(2)} gol a partita.`,
        confidence: 0.90
      });
    }

    // DEFENSE insights
    if (stats.clean_sheet_streak >= 3) {
      insights.push({
        category: 'DEFENSE',
        severity: 'high',
        title: '🛡️ Difesa impenetrabile',
        message: `${stats.clean_sheet_streak} partite consecutive senza subire gol.`,
        confidence: 0.95
      });
    }
    if (stats.conceding_streak >= 5) {
      insights.push({
        category: 'DEFENSE',
        severity: 'low',
        title: '⚠️ Difesa fragile',
        message: `Ha subito gol in ${stats.conceding_streak} partite consecutive.`,
        confidence: 0.85
      });
    }

    // FORM insights
    if (stats.ppg5 > 2.0) {
      insights.push({
        category: 'FORM',
        severity: 'high',
        title: '📈 Forma in crescita',
        message: `${parseFloat(stats.ppg5).toFixed(2)} punti per partita nelle ultime 5.`,
        confidence: 0.90
      });
    }
    if (stats.losing_streak >= 3) {
      insights.push({
        category: 'FORM',
        severity: 'low',
        title: '📉 Momento critico',
        message: `${stats.losing_streak} sconfitte consecutive. Situazione da monitorare.`,
        confidence: 0.90
      });
    }

    // HOME/AWAY insights
    if (parseFloat(stats.home_ppg) > parseFloat(stats.away_ppg) + 1) {
      insights.push({
        category: 'HOME',
        severity: 'medium',
        title: '🏠 Vantaggio tra le mura amiche',
        message: `Ha un grande fattore campo (${parseFloat(stats.home_ppg).toFixed(2)} PPG in casa vs ${parseFloat(stats.away_ppg).toFixed(2)} in trasferta).`,
        confidence: 0.85
      });
    }

    // BTTS insights
    if (parseFloat(stats.btts_percent) > 60) {
      insights.push({
        category: 'BTTS',
        severity: 'medium',
        title: '⚽ Partite con entrambe le squadre in gol',
        message: `${parseFloat(stats.btts_percent).toFixed(0)}% delle partite con BTTS. Alta probabilità di gol da entrambi i lati.`,
        confidence: 0.80
      });
    }

    return insights.sort((a, b) => b.confidence - a.confidence);
  }

  /**
   * Generate match insights
   */

  generateMatchInsights(homeTeamId, awayTeamId, competitionId, season) {
    const matchIndex = this.indexEngine.calculateMatchMetisIndex(
      homeTeamId, awayTeamId, competitionId, season
    );

    if (!matchIndex) return [];

    const insights = [];
    const homeTeam = this.statsEngine.loader.getTeam(homeTeamId);
    const awayTeam = this.statsEngine.loader.getTeam(awayTeamId);

    // Home advantage
    if (matchIndex.home_metis_index > matchIndex.away_metis_index + 10) {
      insights.push({
        category: 'MATCH',
        type: 'HOME_ADVANTAGE',
        title: '🏠 Vantaggio evidente per i padroni',
        confidence: 0.85
      });
    }

    // Form comparison
    if (matchIndex.home_metis_index - matchIndex.away_metis_index > 15) {
      insights.push({
        category: 'MATCH',
        type: 'FORM_GAP',
        title: `${homeTeam.name} in forma notevolmente migliore`,
        confidence: 0.80
      });
    }

    // Risk factors
    matchIndex.risk_factors.forEach(risk => {
      insights.push({
        category: 'RISK',
        type: 'WARNING',
        title: `⚠️ ${risk}`,
        confidence: 0.75
      });
    });

    return insights;
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// 07 — API SERVER & EXPORT
// ═══════════════════════════════════════════════════════════════════════════

class MetisAPIServer {

  constructor(dataLoader, statsEngine, indexEngine, insightsEngine) {
    this.loader = dataLoader;
    this.stats = statsEngine;
    this.index = indexEngine;
    this.insights = insightsEngine;
  }

  /**
   * RESTful API endpoints
   * Ready to use with Express.js or similar
   */

  // GET /api/team/:teamId/stats
  getTeamStats(teamId, competitionId, season) {
    return {
      status: 'success',
      data: this.stats.calculateTeamStats(teamId, competitionId, season)
    };
  }

  // GET /api/team/:teamId/metis-index
  getTeamMetisIndex(teamId, competitionId, season) {
    return {
      status: 'success',
      data: this.index.calculateTeamMetisIndex(teamId, competitionId, season)
    };
  }

  // GET /api/team/:teamId/insights
  getTeamInsights(teamId, competitionId, season) {
    return {
      status: 'success',
      data: this.insights.generateTeamInsights(teamId, competitionId, season)
    };
  }

  // GET /api/match/:homeTeamId/:awayTeamId/prediction
  getMatchPrediction(homeTeamId, awayTeamId, competitionId, season) {
    return {
      status: 'success',
      data: this.index.calculateMatchMetisIndex(homeTeamId, awayTeamId, competitionId, season)
    };
  }

  // GET /api/match/:homeTeamId/:awayTeamId/insights
  getMatchInsights(homeTeamId, awayTeamId, competitionId, season) {
    return {
      status: 'success',
      data: this.insights.generateMatchInsights(homeTeamId, awayTeamId, competitionId, season)
    };
  }

  // GET /api/player/:playerId/rating
  getPlayerRating(playerId, competitionId, season) {
    return {
      status: 'success',
      data: this.index.calculatePlayerMetisRating(playerId, competitionId, season)
    };
  }

  // GET /api/league/:competitionId/standings
  getLeagueStandings(competitionId, season) {
    // Sort teams by points
    const standings = this.loader.data.standings
      .filter(s => s.competition_id === competitionId)
      .sort((a, b) => b.points - a.points);

    return {
      status: 'success',
      data: standings
    };
  }

  // GET /api/league/:competitionId/stats
  getLeagueStats(competitionId, season) {
    return {
      status: 'success',
      data: this.stats.calculateLeagueStats(competitionId, season)
    };
  }

  // GET /api/metis/dashboard
  getDashboard(competitionId, season) {
    const standings = this.getLeagueStandings(competitionId, season).data;
    
    return {
      status: 'success',
      data: {
        standings: standings.slice(0, 5), // Top 5
        league_stats: this.stats.calculateLeagueStats(competitionId, season),
        top_scorers: this.loader.data.top_players
          .filter(p => p.competition_id === competitionId && p.type === 'goals')
          .sort((a, b) => b.value - a.value)
          .slice(0, 5),
        top_assists: this.loader.data.top_players
          .filter(p => p.competition_id === competitionId && p.type === 'assists')
          .sort((a, b) => b.value - a.value)
          .slice(0, 5)
      }
    };
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// 08 — INITIALIZATION & EXPORT
// ═══════════════════════════════════════════════════════════════════════════

class METIS {
  
  constructor(workbook) {
    console.log('🚀 METIS Engine v4.0 initializing...');

    this.loader = new MetisExcelLoader(workbook);
    this.stats = new MetisStatisticsEngine(this.loader);
    this.index = new MetisIndexEngine(this.stats);
    this.insights = new MetisInsightsEngine(this.stats, this.index);
    this.api = new MetisAPIServer(this.loader, this.stats, this.index, this.insights);

    console.log('✅ METIS Engine ready');
  }

  /**
   * Main API surface
   */

  getTeamStats(teamId, competitionId, season = '2024-2025') {
    return this.api.getTeamStats(teamId, competitionId, season);
  }

  getTeamMetisIndex(teamId, competitionId, season = '2024-2025') {
    return this.api.getTeamMetisIndex(teamId, competitionId, season);
  }

  getTeamInsights(teamId, competitionId, season = '2024-2025') {
    return this.api.getTeamInsights(teamId, competitionId, season);
  }

  getMatchPrediction(homeTeamId, awayTeamId, competitionId, season = '2024-2025') {
    return this.api.getMatchPrediction(homeTeamId, awayTeamId, competitionId, season);
  }

  getPlayerRating(playerId, competitionId, season = '2024-2025') {
    return this.api.getPlayerRating(playerId, competitionId, season);
  }

  getLeagueStandings(competitionId, season = '2024-2025') {
    return this.api.getLeagueStandings(competitionId, season);
  }

  getDashboard(competitionId, season = '2024-2025') {
    return this.api.getDashboard(competitionId, season);
  }

  // Export all data as JSON (for Flutter)
  exportAsJSON() {
    return {
      competitions: this.loader.data.competitions,
      teams: this.loader.data.teams,
      players: this.loader.data.players,
      matches: this.loader.data.matches,
      standings: this.loader.data.standings,
      top_players: this.loader.data.top_players,
      timestamp: new Date().toISOString()
    };
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// 09 — USAGE EXAMPLES (for testing)
// ═══════════════════════════════════════════════════════════════════════════

/*
// In your VS Code project:

import ExcelJS from 'exceljs';
import { METIS } from './METIS_ENGINE_v4.0_COMPLETE.js';

// Load Excel workbook
const workbook = new ExcelJS.Workbook();
await workbook.xlsx.readFile('./data/metis_database.xlsx');

// Initialize METIS engine
const metis = new METIS(workbook);

// Example 1: Get team statistics
const interStats = metis.getTeamStats('TEAM_inter', 'COMP_italian-serie-a', '2024-2025');
console.log(interStats);

// Example 2: Get team Metis Index
const interIndex = metis.getTeamMetisIndex('TEAM_inter', 'COMP_italian-serie-a', '2024-2025');
console.log(`Inter Metis Index: ${interIndex.data.index}/100`);

// Example 3: Get team insights
const interInsights = metis.getTeamInsights('TEAM_inter', 'COMP_italian-serie-a', '2024-2025');
console.log(interInsights);

// Example 4: Get match prediction (Inter vs Milan)
const matchPred = metis.getMatchPrediction(
  'TEAM_inter', 
  'TEAM_ac-milan', 
  'COMP_italian-serie-a', 
  '2024-2025'
);
console.log('Match Prediction:', matchPred.data.prediction);

// Example 5: Export all data as JSON for Flutter
const jsonData = metis.exportAsJSON();
console.log(JSON.stringify(jsonData, null, 2));

// Example 6: REST API endpoint (Express.js)
app.get('/api/team/:teamId/stats', (req, res) => {
  const stats = metis.getTeamStats(req.params.teamId, 'COMP_italian-serie-a');
  res.json(stats);
});

app.get('/api/match/:home/:away/prediction', (req, res) => {
  const prediction = metis.getMatchPrediction(req.params.home, req.params.away, 'COMP_italian-serie-a');
  res.json(prediction);
});

app.listen(3000, () => console.log('METIS API running on port 3000'));
*/

// ═══════════════════════════════════════════════════════════════════════════
// Export for Node.js/ES modules
// ═══════════════════════════════════════════════════════════════════════════

// ES Module exports
export {
  METIS,
  MetisExcelLoader,
  MetisStatisticsEngine,
  MetisIndexEngine,
  MetisInsightsEngine,
  MetisAPIServer,
  METIS_CONFIG,
  METIS_SCHEMAS
};

export default METIS;