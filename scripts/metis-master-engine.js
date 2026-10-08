import fs from 'fs';
import path from 'path';
import ExcelJS from 'exceljs';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// ============================================================
// METIS — CONFIGURAZIONE
// ============================================================

const EXCEL_PATH = path.join(__dirname, '../data/metis_database.xlsx');
const PUBLIC_EXCEL_PATH = path.join(__dirname, '../public/metis_database.xlsx');
const HISTORICAL_DIR = path.join(__dirname, '../data/historical');
const MAIN_DIR = path.join(HISTORICAL_DIR, 'main_leagues');
const EXTRA_DIR = path.join(HISTORICAL_DIR, 'extra_leagues');
const FIXTURES_DIR = path.join(__dirname, '../data/fixtures');

const SEASON_START = new Date(2026, 6, 1);
const SEASON_END = new Date(2027, 6, 2);

// UNICO elenco di campionati ammessi (Nomi Ufficiali)
const DIV_NAMES = {
    E0: 'Premier League', E1: 'Championship', E2: 'League One', E3: 'League Two', SC0: 'Premiership Scozia',
    D1: 'Bundesliga Germania', D2: '2. Bundesliga Germania', I1: 'Serie A Italia', I2: 'Serie B Italia',
    SP1: 'La Liga Spagna', SP2: 'Segunda División Spagna', F1: 'Ligue 1 Francia', F2: 'Ligue 2 Francia',
    N1: 'Eredivisie Paesi Bassi', B1: 'Jupiler League Belgio', P1: 'Liga Portugal', T1: 'Süper Lig Turchia',
    G1: 'Super League Grecia', AUT: 'Austria Bundesliga', DNK: 'Danimarca Superliga', NOR: 'Norvegia Eliteserien',
    SWE: 'Svezia Allsvenskan', POL: 'Polonia Ekstraklasa', ROU: 'Romania SuperLiga', SWZ: 'Svizzera Super League',
    ARG: 'Argentina Primera', BRA: 'Brasile Serie A', CHN: 'Cina Super League', JPN: 'Giappone J1 League',
    MEX: 'Messico Liga MX', USA: 'USA MLS', HRV: 'HNL Croazia', CZE: 'Fortuna Liga Rep. Ceca', 
    HUN: 'Nemzeti Bajnokság I Ungheria', BGR: 'First League Bulgaria', AZE: 'Premyer Liqa Azerbaigian', 
    ISR: 'Israeli Premier League', CYP: 'First Division Cipro', SRB: 'Super League Serbia', 
    SVN: 'Prva Liga Slovenia', BIH: 'Premijer Liga Bosnia', SVK: 'Super League Slovacchia', 
    MDA: 'Liga I Moldavia', EST: 'Meistriliiga Estonia', LVA: 'Virsliga Lettonia', LTU: 'A Lyga Lituania'
};

// Aliases per accorpare i doppioni dei campionati sotto l'unico nome ufficiale
const LEAGUE_ALIASES = {
    bundesligaaustria: 'Austria Bundesliga', superligadanimarca: 'Danimarca Superliga',
    eliteserien: 'Norvegia Eliteserien', eliteservienorvegia: 'Norvegia Eliteserien',
    ekstraklasapolonia: 'Polonia Ekstraklasa', superligaromania: 'Romania SuperLiga',
    superleaguesvizzera: 'Svizzera Super League', bundesliga: 'Bundesliga Germania',
    seriea: 'Serie A Italia', serieb: 'Serie B Italia', laliga: 'La Liga Spagna', ligue1: 'Ligue 1 Francia',
    eredivisie: 'Eredivisie Paesi Bassi', premierleague: 'Premier League',
    superlig: 'Süper Lig Turchia', superleaguegrecia: 'Super League Grecia',
    ligaportugal: 'Liga Portugal', jupilerleague: 'Jupiler League Belgio'
};

const TEAM_ALIASES = {
    'Manchester United': 'Man United', 'Manchester City': 'Man City', 'Newcastle United': 'Newcastle',
    'Nottingham Forest': "Nott'm Forest", 'AC Milan': 'Milan', 'Bayern München': 'Bayern Munich',
    'Borussia Dortmund': 'Dortmund', 'Paris Saint Germain': 'Paris SG', 'Paris Saint-Germain': 'Paris SG',
    'Atletico Madrid': 'Ath Madrid', 'Athletic Club': 'Ath Bilbao', 'AFC Bournemouth': 'Bournemouth',
    'Brighton & Hove Albion': 'Brighton', 'Brighton and Hove Albion': 'Brighton', 'Coventry City': 'Coventry',
    'Hull City': 'Hull', 'Ipswich Town': 'Ipswich', 'Leeds United': 'Leeds', 'Tottenham Hotspur': 'Tottenham'
};

const normalizeTeam = name => TEAM_ALIASES[name] || name;

// ============================================================
// UTILITÀ E SISTEMA ANTI-DOPPIONI
// ============================================================

const clean = value => String(value ?? '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9]/g, '').trim();
const clamp = (v, min = 0, max = 100) => Math.min(max, Math.max(min, v));
const CANON = new Map(Object.values(DIV_NAMES).map(n => [clean(n), n]));

function titleCase(str) {
    return str.toLowerCase().split(/[\s_-]+/).map(w => w.charAt(0).toUpperCase() + w.slice(1)).join(' ');
}

// Funzione blindata anti-clone per i campionati
function resolveLeague(division, fallback) {
    const code = String(division || '').trim().toUpperCase();
    if (DIV_NAMES[code]) return DIV_NAMES[code];
    for (const raw of [division, fallback]) {
        if (!raw) continue;
        const key = clean(String(raw).replace(/\s*\d{2,4}([\s\/_-]?\d{2,4})?$/, ''));
        if (!key) continue;
        if (CANON.has(key)) return CANON.get(key);
        if (LEAGUE_ALIASES[key]) return LEAGUE_ALIASES[key];
    }
    const finalName = fallback || division;
    return finalName ? titleCase(finalName.replace(/\s*\d{2,4}([\s\/_-]?\d{2,4})?$/, '')) : null;
}

function round(value, decimals = 2) {
    if (!Number.isFinite(value)) return null;
    const f = 10 ** decimals;
    return Math.round(value * f) / f;
}
function safeAverage(total, count) { return count > 0 ? total / count : null; }
function splitCSVLine(line) {
    const out = []; let cur = ''; let quoted = false;
    for (let i = 0; i < line.length; i++) {
        const ch = line[i];
        if (ch === '"') { if (quoted && line[i + 1] === '"') { cur += '"'; i++; } else quoted = !quoted; } 
        else if (ch === ',' && !quoted) { out.push(cur); cur = ''; } 
        else cur += ch;
    }
    out.push(cur); return out.map(v => v.trim());
}
function parseCSV(content) {
    const lines = String(content ?? '').replace(/^\uFEFF/, '').split(/\r?\n/).filter(l => l.trim().length > 0);
    if (lines.length < 2) return [];
    const headers = splitCSVLine(lines[0]).map(h => h.replace(/^\./, ''));
    return lines.slice(1).filter(line => splitCSVLine(line)[0] !== headers[0]).map(line => {
        const cells = splitCSVLine(line); const row = {};
        headers.forEach((h, i) => row[h] = cells[i] ?? ''); return row;
    });
}
function pick(row, ...keys) {
    for (const key of keys) { if (row[key] !== undefined && String(row[key]).trim() !== '') return String(row[key]).trim(); }
    return '';
}
function parseMatchDate(value) {
    if (!value) return null;
    const input = String(value).trim().split(/[ T]/)[0];
    const parts = input.split(/[\/\-.]/).map(Number);
    if (parts.length !== 3 || parts.some(Number.isNaN)) return null;
    let year, month, day;
    if (parts[0] > 31) [year, month, day] = parts; else [day, month, year] = parts;
    if (year < 100) year += 2000;
    const date = new Date(year, month - 1, day);
    if (Number.isNaN(date.getTime()) || date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) return null;
    return date;
}

const isoDate = d => [d.getFullYear(), String(d.getMonth() + 1).padStart(2, '0'), String(d.getDate()).padStart(2, '0')].join('-');
function parseNum(value) {
    if (value === null || value === undefined || value === '') return null;
    const n = Number(value); return Number.isFinite(n) ? n : null;
}
const ensureDirectory = filePath => fs.mkdirSync(path.dirname(filePath), { recursive: true });
function getOrCreate(map, key, factory) { if (!map.has(key)) map.set(key, factory()); return map.get(key); }

// Funzioni ripristinate per la lettura file
const listCSV = dir => fs.existsSync(dir) ? fs.readdirSync(dir).filter(f => f.toLowerCase().endsWith('.csv')) : [];
const fileToLeague = file => file.replace(/\.csv$/i, '').replace(/[-_]/g, ' ').trim();

// ============================================================
// MODELLI DATI
// ============================================================

function buildIndexResult(c, trendValue) {
    return {
        'Metis Index': round(c.form * 0.2 + c.attack * 0.15 + c.defence * 0.15 + c.creation * 0.15 + c.performance * 0.1 + c.context * 0.1 + c.efficiency * 0.05 + c.trend * 0.1, 1),
        'Forma (20%)': Math.round(c.form), 'Attacco (15%)': Math.round(c.attack), 'Difesa (15%)': Math.round(c.defence),
        'Creazione (15%)': Math.round(c.creation), 'Performance (10%)': Math.round(c.performance),
        'Casa/Trasferta (10%)': Math.round(c.context), 'Efficienza (5%)': Math.round(c.efficiency),
        'Trend (10%)': trendValue >= 0 ? `+${trendValue}` : String(trendValue)
    };
}
const createHistoricalTeamStats = (name, league) => ({
    name, league, matches: 0, homeMatches: 0, homeGF: 0, homeGA: 0, awayMatches: 0, awayGF: 0, awayGA: 0,
    wins: 0, draws: 0, losses: 0, points: 0, cleanSheets: 0, over25: 0, btts: 0
});
const createSeasonStats = (name, league) => ({
    name, league, pH: 0, wH: 0, dH: 0, lH: 0, gfH: 0, gaH: 0, pA: 0, wA: 0, dA: 0, lA: 0, gfA: 0, gaA: 0, over25: 0, gg: 0, cs: 0
});

// ============================================================
// MOTORE PRINCIPALE
// ============================================================

async function runMetisMasterEngine() {
    console.log('🦉 METIS — Avvio Motore Dati (Ottimizzazione HTML Perfetta)');

    ensureDirectory(EXCEL_PATH); ensureDirectory(PUBLIC_EXCEL_PATH);
    const workbook = new ExcelJS.Workbook();
    const todayStart = new Date(new Date().getFullYear(), new Date().getMonth(), new Date().getDate());

    const activeLeaguesMap = new Map();
    const uniqueTeams = new Map();
    const h2hMap = new Map();
    const rawTeamMatches = new Map();
    const teamSeasonStatsMap = new Map();
    const realFixtures = new Map();
    const historicalTeamStats = new Map();
    const leagueGoalStats = new Map();
    const globalGoalStats = { matches: 0, goals: 0, homeGoals: 0, awayGoals: 0 };
    const leagueNamesSet = new Set();
    const skipped = { storico: 0, calendario: 0 };

    // FUNZIONE CHIAVE: Registra la squadra ovunque, garantendo la compatibilità HTML al 100%
    const registerTeamPresence = (team, league) => {
        if (!team || !league) return;
        uniqueTeams.set(team, league); // Per il tab Squadre e ricerca
        const currentLeague = getOrCreate(activeLeaguesMap, league, () => ({ nome: league, matches: 0, teams: new Map() }));
        getOrCreate(currentLeague.teams, team, () => ({ name: team, p: 0, w: 0, d: 0, l: 0, gf: 0, ga: 0, pts: 0 })); // Per Classifiche (previene undefined)
        getOrCreate(teamSeasonStatsMap, team, () => createSeasonStats(team, league)); // Per StatisticheSquadra
        getOrCreate(rawTeamMatches, team, () => []); // Per MetisIndex
    };

    const registerHistoricalStats = (team, league, gf, ga, loc) => {
        const s = getOrCreate(historicalTeamStats, team, () => createHistoricalTeamStats(team, league));
        s.matches++;
        if (loc === 'H') { s.homeMatches++; s.homeGF += gf; s.homeGA += ga; } else { s.awayMatches++; s.awayGF += gf; s.awayGA += ga; }
        if (gf > ga) { s.wins++; s.points += 3; } else if (gf === ga) { s.draws++; s.points++; } else s.losses++;
        if (ga === 0) s.cleanSheets++;
        if (gf + ga > 2.5) s.over25++;
        if (gf > 0 && ga > 0) s.btts++;
    };

    const getLeagueAverage = (league, type) => {
        const src = leagueGoalStats.get(league)?.matches > 0 ? leagueGoalStats.get(league) : (globalGoalStats.matches > 0 ? globalGoalStats : null);
        if (src) return type === 'homeGF' ? src.homeGoals / src.matches : src.awayGoals / src.matches;
        return type === 'homeGF' ? 1.35 : 1.15;
    };

    const getFallbackMetisIndex = (team, league) => {
        const s = historicalTeamStats.get(team);
        if (!s || !s.matches) {
            const simGF = getLeagueAverage(league, 'homeGF') * 0.85; const simGA = getLeagueAverage(league, 'awayGA') * 1.15;
            return {
                'Metis Index': 35.0, 'Forma (20%)': 35, 'Attacco (15%)': Math.round((simGF/2.5)*100), 'Difesa (15%)': Math.round(((2 - Math.min(2, simGA))/2)*50 + 11),
                'Creazione (15%)': Math.round(((simGF + 0.45)/3)*100), 'Performance (10%)': 35, 'Casa/Trasferta (10%)': 35, 'Efficienza (5%)': 50, 'Trend (10%)': '+0'
            };
        }
        const ppg = s.points / s.matches;
        const gf = s.homeGF + s.awayGF; const ga = s.homeGA + s.awayGA;
        return buildIndexResult({
            form: clamp((ppg / 3) * 100), attack: clamp(((gf/s.matches) / 2.5) * 100), defence: clamp(((2 - Math.min(2, ga/s.matches)) / 2) * 50 + (s.cleanSheets/s.matches) * 50),
            creation: clamp((((gf/s.matches) + (s.over25/s.matches)) / 3) * 100), performance: clamp((ppg / 3) * 100), context: clamp((ppg / 3) * 100), 
            efficiency: clamp((gf / Math.max(1, gf + ga)) * 100), trend: 50
        }, 0);
    };

    // 1. LETTURA STORICO
    for (const dir of [MAIN_DIR, EXTRA_DIR]) {
        for (const file of listCSV(dir)) {
            const fileLeague = fileToLeague(file);
            let rows; try { rows = parseCSV(fs.readFileSync(path.join(dir, file), 'utf8')); } catch (e) { continue; }
            for (const match of rows) {
                const home = normalizeTeam(pick(match, 'HomeTeam', 'Home', 'home', 'Casa'));
                const away = normalizeTeam(pick(match, 'AwayTeam', 'Away', 'away', 'Trasferta'));
                const hg = parseNum(pick(match, 'FTHG', 'HG', 'home_score'));
                const ag = parseNum(pick(match, 'FTAG', 'AG', 'away_score'));
                const date = parseMatchDate(pick(match, 'Date', 'date', 'Data'));
                if (!home || !away) continue;
                
                const league = resolveLeague(pick(match, 'Div'), fileLeague);
                if (!league) { skipped.storico++; continue; }
                
                leagueNamesSet.add(league);
                registerTeamPresence(home, league); registerTeamPresence(away, league);

                if (hg === null || ag === null) continue;

                registerHistoricalStats(home, league, hg, ag, 'H');
                registerHistoricalStats(away, league, ag, hg, 'A');
                
                const lgs = getOrCreate(leagueGoalStats, league, () => ({ matches: 0, goals: 0, homeGoals: 0, awayGoals: 0 }));
                lgs.matches++; lgs.homeGoals += hg; lgs.awayGoals += ag;
                globalGoalStats.matches++; globalGoalStats.homeGoals += hg; globalGoalStats.awayGoals += ag;

                const [tA, tB] = [home, away].sort();
                const h2h = getOrCreate(h2hMap, `${tA} vs ${tB}`, () => ({ teamA: tA, teamB: tB, matches: 0, winA: 0, winB: 0, draws: 0, goalsA: 0, goalsB: 0 }));
                h2h.matches++; if (hg === ag) h2h.draws++; else if ((hg > ag ? home : away) === h2h.teamA) h2h.winA++; else h2h.winB++;
                if (home === h2h.teamA) { h2h.goalsA += hg; h2h.goalsB += ag; } else { h2h.goalsA += ag; h2h.goalsB += hg; }

                if (!date || date < SEASON_START || date > todayStart || date >= SEASON_END) continue;

                const curLg = activeLeaguesMap.get(league); curLg.matches++;
                rawTeamMatches.get(home).push({ date, gf: hg, ga: ag, pts: hg > ag ? 3 : hg === ag ? 1 : 0, loc: 'H' });
                rawTeamMatches.get(away).push({ date, gf: ag, ga: hg, pts: ag > hg ? 3 : ag === hg ? 1 : 0, loc: 'A' });

                const sH = teamSeasonStatsMap.get(home), sA = teamSeasonStatsMap.get(away);
                sH.pH++; sH.gfH += hg; sH.gaH += ag; if (hg > ag) sH.wH++; else if (hg === ag) sH.dH++; else sH.lH++;
                sA.pA++; sA.gfA += ag; sA.gaA += hg; if (ag > hg) sA.wA++; else if (ag === hg) sA.dA++; else sA.lA++;
                if (hg + ag > 2.5) { sH.over25++; sA.over25++; }
                if (hg > 0 && ag > 0) { sH.gg++; sA.gg++; }
                if (ag === 0) sH.cs++; if (hg === 0) sA.cs++;

                const ht = curLg.teams.get(home); ht.p++; ht.gf += hg; ht.ga += ag; if (hg > ag) { ht.w++; ht.pts += 3; } else if (hg === ag) { ht.d++; ht.pts++; } else ht.l++;
                const at = curLg.teams.get(away); at.p++; at.gf += ag; at.ga += hg; if (ag > hg) { at.w++; at.pts += 3; } else if (ag === hg) { at.d++; at.pts++; } else at.l++;
            }
        }
    }

    // 2. LETTURA CALENDARIO
    if (fs.existsSync(FIXTURES_DIR)) {
        for (const file of listCSV(FIXTURES_DIR)) {
            const fileLeague = fileToLeague(file);
            let rows; try { rows = parseCSV(fs.readFileSync(path.join(FIXTURES_DIR, file), 'utf8')); } catch (e) { continue; }
            for (const match of rows) {
                const home = normalizeTeam(pick(match, 'HomeTeam', 'Home', 'home', 'Casa', 'S1'));
                const away = normalizeTeam(pick(match, 'AwayTeam', 'Away', 'away', 'Trasferta', 'S2'));
                const date = parseMatchDate(pick(match, 'Date', 'date', 'Data'));
                const time = pick(match, 'Time', 'time', 'Ora', 'Orario');
                const division = pick(match, 'Div', 'Campionato', 'League');
                const status = pick(match, 'Status').toUpperCase();

                if (!home || !away || !date || parseNum(pick(match, 'FTHG')) !== null || ['FINISHED', 'CANCELLED', 'POSTPONED'].includes(status)) continue;
                
                const league = resolveLeague(division, fileLeague);
                if (!league) { skipped.calendario++; continue; }
                
                leagueNamesSet.add(league);
                registerTeamPresence(home, league); registerTeamPresence(away, league);

                if (date >= SEASON_START && date < SEASON_END) {
                    const dateKey = isoDate(date);
                    const key = `${dateKey}|${league}|${home}|${away}`;
                    if (!realFixtures.has(key)) realFixtures.set(key, { data: dateKey, ora: String(time ?? '').trim().substring(0, 5), campionato: league, casa: home, trasferta: away });
                }
            }
        }
    }

    const metisIndexResults = new Map();
    for (const [teamName, matches] of rawTeamMatches.entries()) {
        if (!matches.length) continue;
        matches.sort((a, b) => a.date - b.date);
        const n = matches.length;
        let pts = 0, gf = 0, ga = 0, cs = 0, o25 = 0, hGames = 0, hPts = 0, aGames = 0, aPts = 0;
        for (const m of matches) {
            pts += m.pts; gf += m.gf; ga += m.ga; if (m.ga === 0) cs++; if (m.gf + m.ga > 2.5) o25++;
            if (m.loc === 'H') { hGames++; hPts += m.pts; } else { aGames++; aPts += m.pts; }
        }
        let weighted = 0; matches.slice(-5).forEach((m, i) => { weighted += m.pts * (i + 1); });
        const maxRecent = matches.slice(-5).length ? ((matches.slice(-5).length * (matches.slice(-5).length + 1)) / 2) * 3 : 1;
        let trendValue = 0;
        if (n >= 4) {
            const mid = Math.floor(n / 2);
            trendValue = round((matches.slice(mid).reduce((s, m) => s + m.pts, 0) / matches.slice(mid).length) - (matches.slice(0, mid).reduce((s, m) => s + m.pts, 0) / matches.slice(0, mid).length), 2);
        }
        metisIndexResults.set(teamName, buildIndexResult({
            form: clamp((weighted / maxRecent) * 100), attack: clamp(((gf / n) / 2.5) * 100),
            defence: clamp(((2 - Math.min(2, ga / n)) / 2) * 50 + (cs / n) * 50), creation: clamp((((gf / n) + (o25 / n)) / 3) * 100),
            performance: clamp((pts / n / 3) * 100), context: clamp((((hGames ? hPts / hGames : 0) + (aGames ? aPts / aGames : 0)) / 6) * 100),
            efficiency: clamp((gf / Math.max(1, gf + ga)) * 100), trend: clamp(50 + trendValue * 25)
        }, trendValue));
    }

    const generatedFixtures = Array.from(realFixtures.values()).sort((a, b) => a.data.localeCompare(b.data) || a.ora.localeCompare(b.ora));
    const reportRows = [];
    for (const fx of generatedFixtures) {
        const { casa: home, trasferta: away, campionato } = fx;
        const hIdx = metisIndexResults.get(home) || getFallbackMetisIndex(home, campionato);
        const aIdx = metisIndexResults.get(away) || getFallbackMetisIndex(away, campionato);
        const hVal = hIdx['Metis Index'], aVal = aIdx['Metis Index'];
        const eH = clamp((((teamSeasonStatsMap.get(home).pH > 0 ? teamSeasonStatsMap.get(home).gfH / teamSeasonStatsMap.get(home).pH : getLeagueAverage(campionato, 'homeGF')) + (teamSeasonStatsMap.get(away).pA > 0 ? teamSeasonStatsMap.get(away).gaA / teamSeasonStatsMap.get(away).pA : getLeagueAverage(campionato, 'awayGF'))) / 2), 0.15, 5);
        const eA = clamp((((teamSeasonStatsMap.get(away).pA > 0 ? teamSeasonStatsMap.get(away).gfA / teamSeasonStatsMap.get(away).pA : getLeagueAverage(campionato, 'awayGF')) + (teamSeasonStatsMap.get(home).pH > 0 ? teamSeasonStatsMap.get(home).gaH / teamSeasonStatsMap.get(home).pH : getLeagueAverage(campionato, 'homeGF'))) / 2), 0.15, 5);
        const eT = eH + eA;

        reportRows.push({
            Campionato: campionato, Match: `${home} vs ${away}`, 'Metis Index Casa': hVal, 'Metis Index Trasf': aVal,
            'Esito / Chi Vince': Math.round((hVal / (hVal + aVal || 1)) * 100) > 50 ? `Vittoria Casa (${home})` : `Vittoria Trasferta (${away})`,
            'Over/Under 1.5': eT > 1.6 ? 'Over 1.5' : 'Under 1.5', 'Over/Under 2.5': eT > 2.4 ? 'Over 2.5' : 'Under 2.5',
            'Over/Under 3.5': eT > 3.3 ? 'Over 3.5' : 'Under 3.5', 'Over/Under 4.5': eT > 4.2 ? 'Over 4.5' : 'Under 4.5',
            'Entrambe Segnano (BTTS)': (eH >= 0.8 && eA >= 0.8) ? 'Sì' : 'No',
            'Clean Sheet (Porta Inviolata)': hVal > aVal + 10 ? `${home} — indicatore statistico` : (aVal > hVal + 10 ? `${away} — indicatore statistico` : 'Incerto')
        });
    }

    // 6. SCRITTURA ESATTA SECONDO STRUTTURA HTML ORIGINALE
    const readme = workbook.addWorksheet('README'); readme.addRow(['METIS ENGINE', SEASON_START.getFullYear() + '/' + SEASON_END.getFullYear()]);
    const fixturesSheet = workbook.addWorksheet('Fixtures'); fixturesSheet.addRow(['Data', 'Campionato', 'Casa', 'Trasferta', 'Ora']);
    for (const f of generatedFixtures) fixturesSheet.addRow([f.data, f.campionato, f.casa, f.trasferta, f.ora]);

    const reportSheet = workbook.addWorksheet('Report'); reportSheet.addRow(['Campionato', 'Match', 'Metis Index Casa', 'Metis Index Trasf', 'Esito / Chi Vince', 'Over/Under 1.5', 'Over/Under 2.5', 'Over/Under 3.5', 'Over/Under 4.5', 'Entrambe Segnano (BTTS)', 'Clean Sheet (Porta Inviolata)']);
    for (const r of reportRows) reportSheet.addRow([r.Campionato, r.Match, r['Metis Index Casa'], r['Metis Index Trasf'], r['Esito / Chi Vince'], r['Over/Under 1.5'], r['Over/Under 2.5'], r['Over/Under 3.5'], r['Over/Under 4.5'], r['Entrambe Segnano (BTTS)'], r['Clean Sheet (Porta Inviolata)']]);

    const leaguesSheet = workbook.addWorksheet('Campionati'); leaguesSheet.addRow(['Campionato Ufficiale']);
    Array.from(leagueNamesSet).sort().forEach(l => leaguesSheet.addRow([l]));
    if (leagueNamesSet.size === 0) Object.values(DIV_NAMES).forEach(l => leaguesSheet.addRow([l]));

    const teamsSheet = workbook.addWorksheet('Squadre'); teamsSheet.addRow(['Squadra', 'Campionato']);
    for (const [tName, src] of uniqueTeams.entries()) teamsSheet.addRow([tName, src]);

    const indexSheet = workbook.addWorksheet('MetisIndex'); indexSheet.addRow(['Squadra', 'Metis Index', 'Forma (20%)', 'Attacco (15%)', 'Difesa (15%)', 'Creazione (15%)', 'Performance (10%)', 'Casa/Trasferta (10%)', 'Efficienza (5%)', 'Trend (10%)']);
    for (const teamName of uniqueTeams.keys()) {
        const i = metisIndexResults.get(teamName) || getFallbackMetisIndex(teamName, uniqueTeams.get(teamName));
        indexSheet.addRow([teamName, i['Metis Index'], i['Forma (20%)'], i['Attacco (15%)'], i['Difesa (15%)'], i['Creazione (15%)'], i['Performance (10%)'], i['Casa/Trasferta (10%)'], i['Efficienza (5%)'], i['Trend (10%)']]);
    }

    const h2hSheet = workbook.addWorksheet('H2H'); h2hSheet.addRow(['Confronto', 'Match Totali', 'Vittorie A', 'Pareggi', 'Vittorie B', 'Media Gol Totali']);
    for (const h of h2hMap.values()) if (uniqueTeams.has(h.teamA) && uniqueTeams.has(h.teamB)) h2hSheet.addRow([`${h.teamA} vs ${h.teamB}`, h.matches, h.winA, h.draws, h.winB, ((h.goalsA + h.goalsB) / h.matches).toFixed(2)]);

    const statsSheet = workbook.addWorksheet('StatisticheSquadra'); statsSheet.addRow(['Squadra', 'Campionato', 'Media GF Casa', 'Media GS Casa', 'Media GF Fuori', 'Media GS Fuori', '% Over 2.5', '% Gol/Gol', '% Clean sheet']);
    for (const s of teamSeasonStatsMap.values()) {
        const total = s.pH + s.pA;
        statsSheet.addRow([s.name, s.league, round(safeAverage(s.gfH, s.pH))||0, round(safeAverage(s.gaH, s.pH))||0, round(safeAverage(s.gfA, s.pA))||0, round(safeAverage(s.gaA, s.pA))||0, total ? `${((s.over25 / total) * 100).toFixed(1)}%` : '0.0%', total ? `${((s.gg / total) * 100).toFixed(1)}%` : '0.0%', total ? `${((s.cs / total) * 100).toFixed(1)}%` : '0.0%']);
    }

    const standingsSheet = workbook.addWorksheet('Classifiche'); standingsSheet.addRow(['Campionato', 'Posizione', 'Squadra', 'Punti', 'Giocate', 'Gol Fatti', 'Gol Subiti', 'Diff Reti']);
    for (const l of activeLeaguesMap.values()) {
        Array.from(l.teams.values()).sort((a, b) => b.pts - a.pts || b.gf - b.ga - (a.gf - a.ga) || b.gf - a.gf || a.name.localeCompare(b.name)).forEach((t, i) => {
            standingsSheet.addRow([l.nome, i + 1, t.name, t.pts, t.p, t.gf, t.ga, t.gf - t.ga]);
        });
    }

    await workbook.xlsx.writeFile(EXCEL_PATH); fs.copyFileSync(EXCEL_PATH, PUBLIC_EXCEL_PATH);
    console.log('----------------------------------------\nMETIS — Elaborazione Perfetta\n----------------------------------------');
}
runMetisMasterEngine().catch(e => { console.error('Errore:', e); process.exitCode = 1; });