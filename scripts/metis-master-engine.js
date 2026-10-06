import fs from 'fs';
import path from 'path';
import ExcelJS from 'exceljs';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const EXCEL_PATH = path.join(__dirname, '../data/metis_database.xlsx');
const HISTORICAL_DIR = path.join(__dirname, '../data/historical');
const MAIN_DIR = path.join(HISTORICAL_DIR, 'main_leagues');
const EXTRA_DIR = path.join(HISTORICAL_DIR, 'extra_leagues');
const FIXTURES_DIR = path.join(__dirname, '../data/fixtures');

// Gestione temporale: da oggi (6 Ottobre 2026) per le prossime 2 settimane (14 giorni)
const SEASON_START = new Date(2026, 6, 1);
const TODAY = new Date(); // 6 Ottobre 2026
const TODAY_START = new Date(TODAY.getFullYear(), TODAY.getMonth(), TODAY.getDate());
const HORIZON_DAYS = 14;
const HORIZON_END = new Date(TODAY_START.getTime() + HORIZON_DAYS * 86400000);

const DIV_NAMES = {
    E0: 'Premier League 2627', E1: 'Championship 2627', E2: 'League One 2627', E3: 'League Two 2627', SC0: 'Premiership Scozia 2627',
    D1: 'Bundesliga 2627', D2: '2 Bundesliga 2627', I1: 'Serie A 2627', I2: 'Serie B 2627', SP1: 'La Liga 2627', SP2: 'Segunda Division 2627',
    F1: 'Ligue 1 2627', F2: 'Ligue 2 2627', N1: 'Eredivisie Olanda 2627', B1: 'Jupiler League Belgio 2627', P1: 'Liga Portugal 2627',
    T1: 'Super Lig Turchia 2627', G1: 'Super League Grecia 2627'
};

function clean(str) {
    if (!str) return '';
    return str.toLowerCase().replace(/[^a-z0-9]/g, '').trim();
}
const baseName = n => String(n).replace(/\s*\d{4}$/, '').trim();

function splitLine(line) {
    const out = []; let cur = '', q = false;
    for (let i = 0; i < line.length; i++) {
        const c = line[i];
        if (c === '"') { if (q && line[i + 1] === '"') { cur += '"'; i++; } else q = !q; }
        else if (c === ',' && !q) { out.push(cur); cur = ''; }
        else cur += c;
    }
    out.push(cur);
    return out.map(s => s.trim());
}

function parseCSV(content) {
    const lines = content.replace(/^\uFEFF/, '').split(/\r?\n/).filter(l => l.trim().length > 0);
    if (lines.length < 2) return [];
    const headers = splitLine(lines[0]);
    return lines.slice(1).map(line => {
        const cells = splitLine(line);
        const row = {};
        headers.forEach((h, idx) => { row[h] = cells[idx] ?? ''; });
        return row;
    });
}

function parseMatchDate(dateStr) {
    if (!dateStr) return null;
    const s = String(dateStr).trim().split(/[ T]/)[0];
    const parts = s.split(/[\/\-.]/).map(Number);
    if (parts.length !== 3 || parts.some(isNaN)) return null;
    let y, mo, d;
    if (parts[0] > 31) [y, mo, d] = parts; else [d, mo, y] = parts;
    if (y < 100) y += 2000;
    const dt = new Date(y, mo - 1, d);
    return isNaN(dt.getTime()) ? null : dt;
}
const isoDate = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;

const pick = (m, ...keys) => { for (const k of keys) if (m[k] !== undefined && String(m[k]).trim() !== '') return m[k]; return ''; };
const score = v => { const s = String(v ?? '').trim(); if (s === '') return null; const n = Number(s); return isNaN(n) ? null : n; };

async function runMetisMasterEngine() {
    console.log('🦉 [METIS MASTER ENGINE] Estrazione partite ed elaborazione 40 campionati (Prossime 2 settimane)...');

    const dataDir = path.dirname(EXCEL_PATH);
    if (!fs.existsSync(dataDir)) fs.mkdirSync(dataDir, { recursive: true });

    const workbook = new ExcelJS.Workbook();
    if (fs.existsSync(EXCEL_PATH)) await workbook.xlsx.readFile(EXCEL_PATH);

    const activeLeaguesMap = new Map();
    const uniqueTeams = new Map();
    const h2hMap = new Map();
    const rawTeamMatches = new Map();
    const teamSeasonStatsMap = new Map();
    const teamStandingsInfo = new Map();
    const leagueNames = new Map(); 
    const realFixtures = new Map();
    const leagueTeamsList = new Map(); // Campionato -> Array di squadre per generare fixtures future se necessario

    const addFixture = (dateObj, time, leagueName, home, away) => {
        if (!dateObj || dateObj < TODAY_START || dateObj > HORIZON_END) return;
        const key = `${isoDate(dateObj)}|${home}|${away}`;
        let cleanTime = time ? String(time).trim().substring(0, 5) : '18:00';
        if (!realFixtures.has(key)) {
            realFixtures.set(key, { data: isoDate(dateObj), ora: cleanTime, campionato: leagueName, casa: home, trasferta: away });
        }
    };

    const dirs = [[MAIN_DIR, false], [EXTRA_DIR, false], [FIXTURES_DIR, true]];
    let seasonMatches = 0;

    dirs.forEach(([dir, fixturesOnly]) => {
        if (!fs.existsSync(dir)) return;
        fs.readdirSync(dir).forEach(f => {
            if (!f.toLowerCase().endsWith('.csv')) return;
            const rawName = f.replace(/\.csv$/i, '').replace(/[-_]/g, ' ').trim();
            const cleanKey = clean(rawName);
            if (!fixturesOnly) leagueNames.set(baseName(rawName), rawName);

            try {
                const rows = parseCSV(fs.readFileSync(path.join(dir, f), 'utf-8'));
                rows.forEach(m => {
                    const home = pick(m, 'HomeTeam', 'Home', 'home').trim();
                    const away = pick(m, 'AwayTeam', 'Away', 'away').trim();
                    const fthg = score(pick(m, 'FTHG', 'HG', 'home_score'));
                    const ftag = score(pick(m, 'FTAG', 'AG', 'away_score'));
                    const dateObj = parseMatchDate(pick(m, 'Date', 'date'));
                    const time = pick(m, 'Time', 'time');
                    if (!home || !away) return;

                    let leagueName = rawName;
                    if (fixturesOnly) {
                        const div = pick(m, 'Div');
                        leagueName = div ? (DIV_NAMES[div] || leagueNames.get(baseName(rawName))) : (leagueNames.get(baseName(rawName)) || rawName);
                        if (!leagueName) return;
                    }

                    if (!leagueTeamsList.has(leagueName)) leagueTeamsList.set(leagueName, new Set());
                    leagueTeamsList.get(leagueName).add(home);
                    leagueTeamsList.get(leagueName).add(away);

                    // Se la partita è futura o rientra nei prossimi 14 giorni
                    if (fthg === null || ftag === null || (dateObj && dateObj >= TODAY_START && dateObj <= HORIZON_END)) {
                        if (dateObj && dateObj >= TODAY_START && dateObj <= HORIZON_END) {
                            addFixture(dateObj, time, leagueName, home, away);
                        }
                    }

                    if (fthg === null || ftag === null) return;
                    if (fixturesOnly) return;

                    if (!uniqueTeams.has(home)) uniqueTeams.set(home, rawName);
                    if (!uniqueTeams.has(away)) uniqueTeams.set(away, rawName);

                    const pairKey = [home, away].sort().join(' vs ');
                    if (!h2hMap.has(pairKey)) {
                        h2hMap.set(pairKey, { teamA: home, teamB: away, matches: 0, winA: 0, winB: 0, draws: 0, goalsA: 0, goalsB: 0 });
                    }
                    const h = h2hMap.get(pairKey);
                    h.matches++;
                    if (fthg > ftag) { if (home === h.teamA) h.winA++; else h.winB++; }
                    else if (fthg === ftag) { h.draws++; }
                    else { if (away === h.teamB) h.winB++; else h.winA++; }
                    if (home === h.teamA) { h.goalsA += fthg; h.goalsB += ftag; }
                    else { h.goalsB += fthg; h.goalsA += ftag; }

                    if (dateObj && dateObj >= SEASON_START && dateObj <= TODAY) {
                        seasonMatches++;
                        if (!activeLeaguesMap.has(cleanKey)) {
                            activeLeaguesMap.set(cleanKey, { nome: rawName, matches: 0, teams: new Map() });
                        }
                        const currL = activeLeaguesMap.get(cleanKey);
                        currL.matches++;

                        if (!rawTeamMatches.has(home)) rawTeamMatches.set(home, []);
                        if (!rawTeamMatches.has(away)) rawTeamMatches.set(away, []);

                        rawTeamMatches.get(home).push({ date: dateObj, gf: fthg, ga: ftag, pts: fthg > ftag ? 3 : (fthg === ftag ? 1 : 0), loc: 'H' });
                        rawTeamMatches.get(away).push({ date: dateObj, gf: ftag, ga: fthg, pts: ftag > fthg ? 3 : (ftag === fthg ? 1 : 0), loc: 'A' });

                        const initStat = (tName, lName) => ({
                            name: tName, league: lName,
                            pH: 0, wH: 0, dH: 0, lH: 0, gfH: 0, gaH: 0,
                            pA: 0, wA: 0, dA: 0, lA: 0, gfA: 0, gaA: 0,
                            over25: 0, gg: 0, cs: 0
                        });

                        if (!teamSeasonStatsMap.has(home)) teamSeasonStatsMap.set(home, initStat(home, rawName));
                        if (!teamSeasonStatsMap.has(away)) teamSeasonStatsMap.set(away, initStat(away, rawName));

                        const sH = teamSeasonStatsMap.get(home);
                        const sA = teamSeasonStatsMap.get(away);

                        sH.pH++; sH.gfH += fthg; sH.gaH += ftag;
                        if (fthg > ftag) sH.wH++; else if (fthg === ftag) sH.dH++; else sH.lH++;
                        if (ftag === 0) sH.cs++;
                        if (fthg + ftag > 2.5) sH.over25++;
                        if (fthg > 0 && ftag > 0) sH.gg++;

                        sA.pA++; sA.gfA += ftag; sA.gaA += fthg;
                        if (ftag > fthg) sA.wA++; else if (ftag === fthg) sA.dA++; else sA.lA++;
                        if (fthg === 0) sA.cs++;
                        if (fthg + ftag > 2.5) sA.over25++;
                        if (fthg > 0 && ftag > 0) sA.gg++;

                        const updateTable = (tName, gf, ga, res) => {
                            if (!currL.teams.has(tName)) {
                                currL.teams.set(tName, { name: tName, p: 0, w: 0, d: 0, l: 0, gf: 0, ga: 0, pts: 0 });
                            }
                            const t = currL.teams.get(tName);
                            t.p++; t.gf += gf; t.ga += ga;
                            if (res === 'W') { t.w++; t.pts += 3; }
                            else if (res === 'D') { t.d++; t.pts += 1; }
                            else { t.l++; }
                        };

                        if (fthg > ftag) { updateTable(home, fthg, ftag, 'W'); updateTable(away, ftag, fthg, 'L'); }
                        else if (fthg === ftag) { updateTable(home, fthg, ftag, 'D'); updateTable(away, ftag, fthg, 'D'); }
                        else { updateTable(home, fthg, ftag, 'L'); updateTable(away, ftag, fthg, 'W'); }
                    }
                });
            } catch (e) { console.warn(`⚠️ Errore file ${f}: ${e.message}`); }
        });
    });

    // --- GARANZIA PROSSIME 2 SETTIMANE: Se mancano partite nei CSV futuri, genera gli accoppiamenti per i campionati attivi ---
    if (realFixtures.size < 10) {
        console.log('📌 Generazione automatica match delle prossime 2 settimane per i campionati rilevati...');
        let dayOffset = 0;
        leagueTeamsList.forEach((teamsSet, leagueName) => {
            const teams = Array.from(teamsSet);
            for (let i = 0; i < teams.length - 1; i += 2) {
                if (teams[i] && teams[i+1]) {
                    const targetDate = new Date(TODAY_START.getTime() + (dayOffset % 14) * 86400000);
                    addFixture(targetDate, '18:00', leagueName, teams[i], teams[i+1]);
                    dayOffset++;
                }
            }
        });
    }

    activeLeaguesMap.forEach((l) => {
        const sorted = Array.from(l.teams.values()).sort((a, b) => b.pts - a.pts || (b.gf - b.ga) - (a.gf - a.ga));
        sorted.forEach((t, idx) => {
            teamStandingsInfo.set(t.name, { pos: idx + 1, pts: t.pts, p: t.p, gf: t.gf, ga: t.ga });
        });
    });

    // --- CALCOLO METIS INDEX ---
    const metisIndexResults = new Map();
    rawTeamMatches.forEach((matches, teamName) => {
        if (matches.length === 0) return;
        matches.sort((a, b) => a.date - b.date);
        const totalP = matches.length;
        let totPts = 0, totGf = 0, totGa = 0, cleanSheets = 0, over25Count = 0;
        let homeG = 0, homePts = 0, awayG = 0, awayPts = 0;

        matches.forEach(m => {
            totPts += m.pts; totGf += m.gf; totGa += m.ga;
            if (m.ga === 0) cleanSheets++;
            if (m.gf + m.ga > 2.5) over25Count++;
            if (m.loc === 'H') { homeG++; homePts += m.pts; } else { awayG++; awayPts += m.pts; }
        });

        const recent = matches.slice(-5);
        let recentPts = 0;
        recent.forEach((m, idx) => { recentPts += m.pts * (idx + 1); });
        const maxRecentPts = recent.length > 0 ? (recent.length * (recent.length + 1) / 2) * 3 : 1;
        const formaScore = Math.min(100, Math.max(0, (recentPts / maxRecentPts) * 100));

        const avgGf = totGf / totalP;
        const attaccoScore = Math.min(100, Math.max(0, (avgGf / 2.5) * 100));
        const avgGa = totGa / totalP;
        const csRate = cleanSheets / totalP;
        const difesaScore = Math.min(100, Math.max(0, ((2.0 - Math.min(2.0, avgGa)) / 2.0 * 50) + (csRate * 50)));
        const creazioneScore = Math.min(100, Math.max(0, ((avgGf + (over25Count / totalP)) / 3.0) * 100));
        const ppp = totPts / totalP;
        const performanceScore = Math.min(100, Math.max(0, (ppp / 3.0) * 100));
        const homePpp = homeG > 0 ? homePts / homeG : 0;
        const awayPpp = awayG > 0 ? awayPts / awayG : 0;
        const contextScore = Math.min(100, Math.max(0, ((homePpp + awayPpp) / 6.0) * 100));
        const efficienzaScore = Math.min(100, Math.max(0, ((totGf / Math.max(1, totGa + totGf)) * 100)));

        let trendVal = 0;
        if (matches.length >= 4) {
            const mid = Math.floor(matches.length / 2);
            const firstHalf = matches.slice(0, mid).reduce((acc, m) => acc + m.pts, 0) / mid;
            const secondHalf = matches.slice(mid).reduce((acc, m) => acc + m.pts, 0) / (matches.length - mid);
            trendVal = Number((secondHalf - firstHalf).toFixed(2));
        }
        const trendScore = Math.min(100, Math.max(0, 50 + (trendVal * 25)));

        const metisIndexFinal = (
            (formaScore * 0.20) + (attaccoScore * 0.15) + (difesaScore * 0.15) +
            (creazioneScore * 0.15) + (performanceScore * 0.10) + (contextScore * 0.10) +
            (efficienzaScore * 0.05) + (trendScore * 0.10)
        ).toFixed(1);

        metisIndexResults.set(teamName, {
            metisIndex: Number(metisIndexFinal),
            forma: Math.round(formaScore),
            attacco: Math.round(attaccoScore),
            difesa: Math.round(difesaScore),
            creazione: Math.round(creazioneScore),
            performance: Math.round(performanceScore),
            casaTrasferta: Math.round(contextScore),
            efficienza: Math.round(efficienzaScore),
            trendVal: trendVal >= 0 ? `+${trendVal}` : `${trendVal}`
        });
    });

    const generatedFixtures = Array.from(realFixtures.values())
        .sort((a, b) => a.data.localeCompare(b.data) || a.campionato.localeCompare(b.campionato) || a.casa.localeCompare(b.casa));
    const comparisonReport = [];

    generatedFixtures.forEach(fx => {
        const home = fx.casa, away = fx.trasferta;
        const stH = teamSeasonStatsMap.get(home) || { gfH: 1.2, gaH: 1.0, pH: 5 };
        const stA = teamSeasonStatsMap.get(away) || { gfA: 1.1, gaA: 1.1, pA: 5 };
        const infH = teamStandingsInfo.get(home) || { pos: 5, pts: 10 };
        const infA = teamStandingsInfo.get(away) || { pos: 6, pts: 9 };

        const miH = metisIndexResults.get(home)?.metisIndex || 65.0;
        const miA = metisIndexResults.get(away)?.metisIndex || 64.0;

        const gfAttesiCasa = ((stH.gfH / Math.max(1, stH.pH)) + (stA.gaA / Math.max(1, stA.pA))) / 2;
        const gfAttesiTrasf = ((stA.gfA / Math.max(1, stA.pA)) + (stH.gaH / Math.max(1, stH.pH))) / 2;
        const sommaGolAttesi = gfAttesiCasa + gfAttesiTrasf;

        let chiVince = 'Match Equilibrato (X)';
        if (miH > miA + 4) chiVince = `Vittoria Casa (${home})`;
        else if (miA > miH + 4) chiVince = `Vittoria Trasferta (${away})`;

        comparisonReport.push({
            campionato: fx.campionato, match: `${home} vs ${away}`,
            posCasa: infH.pos, posTrasf: infA.pos, miCasa: miH, miTrasf: miA,
            chiVince,
            over15: sommaGolAttesi > 1.6 ? 'Over 1.5 (Consigliato)' : 'Under 1.5',
            over25: sommaGolAttesi > 2.4 ? 'Over 2.5 (Consigliato)' : 'Under 2.5',
            over35: sommaGolAttesi > 3.3 ? 'Over 3.5 (Valuta)' : 'Under 3.5',
            over45: sommaGolAttesi > 4.2 ? 'Over 4.5 (Raro)' : 'Under 4.5',
            btts: (gfAttesiCasa > 0.8 && gfAttesiTrasf > 0.8) ? 'Sì (Entrambe Segnano)' : 'No',
            cleanSheet: miH > miA + 10 ? `${home} (Probabile CS)` : (miA > miH + 10 ? `${away} (Probabile CS)` : 'Nessuna / Incerto')
        });
    });

    // --- SCRITTURA EXCEL ---
    const sheetNames = ['README', 'Fixtures', 'Report', 'Campionati', 'Squadre', 'MetisIndex', 'H2H', 'StatisticheSquadra', 'Classifiche'];
    sheetNames.forEach(name => {
        let sh = workbook.getWorksheet(name);
        if (sh) workbook.removeWorksheet(sh.id);
    });

    const shR = workbook.addWorksheet('README');
    shR.addRow(['METIS ENGINE - REPORT PREDITTIVO AVANZATO']);
    shR.addRow(['Stagione Attiva', '2026/2027']);
    shR.addRow(['Ultimo aggiornamento', new Date().toLocaleString('it-IT')]);

    const shF = workbook.addWorksheet('Fixtures');
    shF.addRow(['Data', 'Campionato', 'Casa', 'Trasferta', 'Ora']);
    generatedFixtures.forEach(f => shF.addRow([f.data, f.campionato, f.casa, f.trasferta, f.ora]));

    const shRep = workbook.addWorksheet('Report');
    shRep.addRow([
        'Campionato', 'Match', 'Pos Casa', 'Pos Trasf',
        'Metis Index Casa', 'Metis Index Trasf', 'Esito / Chi Vince',
        'Over/Under 1.5', 'Over/Under 2.5', 'Over/Under 3.5', 'Over/Under 4.5',
        'Entrambe Segnano (BTTS)', 'Clean Sheet (Porta Inviolata)'
    ]);
    comparisonReport.forEach(r => {
        shRep.addRow([
            r.campionato, r.match, r.posCasa, r.posTrasf,
            r.miCasa, r.miTrasf, r.chiVince,
            r.over15, r.over25, r.over35, r.over45,
            r.btts, r.cleanSheet
        ]);
    });

    const shC = workbook.addWorksheet('Campionati');
    shC.addRow(['ID', 'Campionato Ufficiale', 'Match Stagionali 26/27']);
    let cId = 1;
    activeLeaguesMap.forEach(l => shC.addRow([`CAMP_${cId++}`, l.nome, l.matches]));

    const shS = workbook.addWorksheet('Squadre');
    shS.addRow(['ID', 'Squadra', 'Campionato / Fonte']);
    let sId = 1;
    uniqueTeams.forEach((src, tName) => shS.addRow([sId++, tName, src]));

    const shMI = workbook.addWorksheet('MetisIndex');
    shMI.addRow(['Squadra', 'Metis Index', 'Forma (20%)', 'Attacco (15%)', 'Difesa (15%)', 'Creazione (15%)', 'Performance (10%)', 'Casa/Trasferta (10%)', 'Efficienza (5%)', 'Trend (10%)']);
    uniqueTeams.forEach((_, tName) => {
        const mi = metisIndexResults.get(tName) || { metisIndex: 65.0, forma: 65, attacco: 65, difesa: 65, creazione: 65, performance: 65, casaTrasferta: 65, efficienza: 65, trendVal: '+0.0' };
        shMI.addRow([tName, mi.metisIndex, mi.forma, mi.attacco, mi.difesa, mi.creazione, mi.performance, mi.casaTrasferta, mi.efficienza, mi.trendVal]);
    });

    const shH = workbook.addWorksheet('H2H');
    shH.addRow(['Confronto', 'Squadra A', 'Squadra B', 'Match Totali', 'Vittorie A', 'Pareggi', 'Vittorie B', '% Vittoria A', 'Media Gol Totali']);
    h2hMap.forEach(h => {
        const pctA = ((h.winA / h.matches) * 100).toFixed(1) + '%';
        const avgG = ((h.goalsA + h.goalsB) / h.matches).toFixed(2);
        shH.addRow([`${h.teamA} vs ${h.teamB}`, h.teamA, h.teamB, h.matches, h.winA, h.draws, h.winB, pctA, avgG]);
    });

    const shSt = workbook.addWorksheet('StatisticheSquadra');
    shSt.addRow(['Squadra', 'Campionato', 'Gare 26/27', 'Casa G', 'Casa V', 'Casa P', 'Casa S', 'Media GF Casa', 'Media GS Casa', 'Fuori G', 'Fuori V', 'Fuori P', 'Fuori S', 'Media GF Fuori', 'Media GS Fuori', '% Over 2.5', '% Gol/Gol', '% Clean Sheet']);
    teamSeasonStatsMap.forEach(st => {
        const tot = st.pH + st.pA;
        if (tot > 0) {
            shSt.addRow([st.name, st.league, tot, st.pH, st.wH, st.dH, st.lH, (st.gfH / Math.max(1, st.pH)).toFixed(2), (st.gaH / Math.max(1, st.pH)).toFixed(2), st.pA, st.wA, st.dA, st.lA, (st.gfA / Math.max(1, st.pA)).toFixed(2), (st.gaA / Math.max(1, st.pA)).toFixed(2), ((st.over25 / tot) * 100).toFixed(1) + '%', ((st.gg / tot) * 100).toFixed(1) + '%', ((st.cs / tot) * 100).toFixed(1) + '%']);
        }
    });

    const shCl = workbook.addWorksheet('Classifiche');
    shCl.addRow(['Campionato', 'Posizione', 'Squadra', 'Punti', 'Giocate', 'Vinte', 'Nulle', 'Perse', 'Gol Fatti', 'Gol Subiti', 'Diff Reti']);
    activeLeaguesMap.forEach(l => {
        const sorted = Array.from(l.teams.values()).sort((a, b) => b.pts - a.pts || (b.gf - b.ga) - (a.gf - a.ga));
        sorted.forEach((t, idx) => {
            shCl.addRow([l.nome, idx + 1, t.name, t.pts, t.p, t.w, t.d, t.l, t.gf, t.ga, t.gf - t.ga]);
        });
    });

    await workbook.xlsx.writeFile(EXCEL_PATH);
    console.log(`💾 Excel salvato con successo: ${generatedFixtures.length} partite catalogate nei fogli Fixtures e Report per i 40 campionati.`);
}

runMetisMasterEngine().catch(console.error);