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

const SEASON_START = new Date('2026-07-01');
const TODAY = new Date('2026-10-06T23:59:59');

function clean(str) {
    if (!str) return '';
    return str.toLowerCase().replace(/[^a-z0-9]/g, '').trim();
}

function parseCSV(content) {
    const lines = content.split('\n').filter(l => l.trim().length > 0);
    if (lines.length < 2) return [];

    const headers = lines[0].split(',').map(h => h.trim().replace(/^"|"$/g, ''));
    const rows = [];

    for (let i = 1; i < lines.length; i++) {
        const cells = lines[i].match(/(".*?"|[^",\s]+)(?=\s*,|\s*$)/g) || lines[i].split(',');
        const row = {};
        headers.forEach((h, idx) => {
            row[h] = cells[idx] ? cells[idx].trim().replace(/^"|"$/g, '') : '';
        });
        rows.push(row);
    }
    return rows;
}

function parseMatchDate(dateStr) {
    if (!dateStr) return null;
    const cleanStr = dateStr.trim();
    if (cleanStr.includes('-')) {
        const d = new Date(cleanStr);
        return isNaN(d.getTime()) ? null : d;
    }
    const parts = cleanStr.split('/');
    if (parts.length === 3) {
        let year = parseInt(parts[2], 10);
        if (year < 100) year += 2000;
        const month = parseInt(parts[1], 10) - 1;
        const day = parseInt(parts[0], 10);
        const d = new Date(year, month, day);
        return isNaN(d.getTime()) ? null : d;
    }
    return null;
}

async function runMetisMasterEngine() {
    console.log('🚀 [METIS MASTER ENGINE] Avvio analisi predittiva e generazione report comparativo avanzato...');

    if (!fs.existsSync(EXCEL_PATH)) {
        console.error(`❌ File non trovato: ${EXCEL_PATH}`);
        return;
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(EXCEL_PATH);

    const activeLeaguesMap = new Map();
    const uniqueTeams = new Map();
    const h2hMap = new Map();
    const rawTeamMatches = new Map();
    const teamSeasonStatsMap = new Map();
    const teamStandingsInfo = new Map();

    const dirs = [MAIN_DIR, EXTRA_DIR];
    let totalFiles = 0;
    let seasonMatches = 0;

    dirs.forEach(dir => {
        if (!fs.existsSync(dir)) return;
        const files = fs.readdirSync(dir);

        files.forEach(f => {
            if (!f.endsWith('.csv')) return;
            totalFiles++;
            const rawName = f.replace('.csv', '').replace(/[-_]/g, ' ').trim();
            const cleanKey = clean(rawName);

            try {
                const content = fs.readFileSync(path.join(dir, f), 'utf-8');
                const rows = parseCSV(content);

                rows.forEach(m => {
                    const home = (m.HomeTeam || m.Home || m.home || '').trim();
                    const away = (m.AwayTeam || m.Away || m.away || '').trim();
                    const fthg = Number(m.FTHG ?? m.HG ?? m.home_score);
                    const ftag = Number(m.FTAG ?? m.AG ?? m.away_score);
                    const dateObj = parseMatchDate(m.Date || m.date);

                    if (home && away) {
                        if (home && !uniqueTeams.has(home)) uniqueTeams.set(home, rawName);
                        if (away && !uniqueTeams.has(away)) uniqueTeams.set(away, rawName);

                        if (!isNaN(fthg) && !isNaN(ftag)) {
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
                                rawTeamMatches.get(away).push({ date: dateObj, gf: ftag, ga: fthg, pts: ftag > fthg ? 3 : (ftag === ftag ? 1 : 0), loc: 'A' });

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
                                if (ftag > fthg) sA.wA++; else if (ftag === ftag) sA.dA++; else sA.lA++;
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
                        }
                    }
                });
            } catch {}
        });
    });

    // Calcolo classifiche e posizioni
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

    // --- GENERAZIONE FIXTURES & REPORT PREDITTIVO COMPLETO ---
    const generatedFixtures = [];
    const comparisonReport = [];

    activeLeaguesMap.forEach((l) => {
        const teamsArr = Array.from(l.teams.keys());
        for (let i = 0; i < teamsArr.length; i += 2) {
            if (i + 1 < teamsArr.length) {
                const home = teamsArr[i];
                const away = teamsArr[i + 1];
                const matchName = `${home} vs ${away}`;
                
                generatedFixtures.push({
                    data: '2026-10-12',
                    campionato: l.nome,
                    casa: home,
                    trasferta: away
                });

                const stH = teamSeasonStatsMap.get(home) || { gfH: 0, gaH: 0, pH: 1 };
                const stA = teamSeasonStatsMap.get(away) || { gfA: 0, gaA: 0, pA: 1 };
                const infH = teamStandingsInfo.get(home) || { pos: '-', pts: 0 };
                const infA = teamStandingsInfo.get(away) || { pos: '-', pts: 0 };

                const miH = metisIndexResults.get(home)?.metisIndex || 50.0;
                const miA = metisIndexResults.get(away)?.metisIndex || 50.0;

                const gfAttesiCasa = ((stH.gfH / Math.max(1, stH.pH)) + (stA.gaA / Math.max(1, stA.pA))) / 2;
                const gfAttesiTrasf = ((stA.gfA / Math.max(1, stA.pA)) + (stH.gaH / Math.max(1, stH.pH))) / 2;
                const sommaGolAttesi = gfAttesiCasa + gfAttesiTrasf;

                // Esito e Squadra Favorita
                let chiVince = 'Match Equilibrato (X)';
                if (miH > miA + 4) chiVince = `Vittoria Casa (${home})`;
                else if (miA > miH + 4) chiVince = `Vittoria Trasferta (${away})`;

                // Over / Under
                const over15 = sommaGolAttesi > 1.6 ? 'Over 1.5 (Consigliato)' : 'Under 1.5';
                const over25 = sommaGolAttesi > 2.4 ? 'Over 2.5 (Consigliato)' : 'Under 2.5';
                const over35 = sommaGolAttesi > 3.3 ? 'Over 3.5 (Valuta)' : 'Under 3.5';
                const over45 = sommaGolAttesi > 4.2 ? 'Over 4.5 (Raro)' : 'Under 4.5';

                // Entrambe le squadre segnano (Goal / No Goal)
                const btts = (gfAttesiCasa > 0.8 && gfAttesiTrasf > 0.8) ? 'Sì (Entrambe Segnano)' : 'No';

                // Clean Sheet
                let cleanSheetTeam = 'Nessuna / Incerto';
                if (miH > miA + 10) cleanSheetTeam = `${home} (Probabile CS)`;
                else if (miA > miH + 10) cleanSheetTeam = `${away} (Probabile CS)`;

                comparisonReport.push({
                    campionato: l.nome,
                    match: matchName,
                    posCasa: infH.pos,
                    posTrasf: infA.pos,
                    miCasa: miH,
                    miTrasf: miA,
                    chiVince: chiVince,
                    over15: over15,
                    over25: over25,
                    over35: over35,
                    over45: over45,
                    btts: btts,
                    cleanSheet: cleanSheetTeam
                });
            }
        }
    });

    // --- SCRITTURA EXCEL ---
    const sheetNames = ['README', 'Fixtures', 'Report', 'Campionati', 'Squadre', 'MetisIndex', 'H2H', 'StatisticheSquadra', 'Classifiche'];
    sheetNames.forEach(name => {
        let sh = workbook.getWorksheet(name);
        if (sh) workbook.removeWorksheet(sh.id);
    });

    workbook.addWorksheet('README');
    const shR = workbook.getWorksheet('README');
    shR.addRow(['METIS ENGINE - REPORT PREDITTIVO AVANZATO']);
    shR.addRow(['Stagione Attiva', '2026/2027']);

    // Fixtures
    const shF = workbook.addWorksheet('Fixtures');
    shF.addRow(['Data', 'Campionato', 'Casa', 'Trasferta']);
    generatedFixtures.forEach(f => shF.addRow([f.data, f.campionato, f.casa, f.trasferta]));

    // Report Predittivo Dettagliato
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

    // Campionati
    const shC = workbook.addWorksheet('Campionati');
    shC.addRow(['ID', 'Campionato Ufficiale', 'Match Stagionali 26/27']);
    let cId = 1;
    activeLeaguesMap.forEach(l => shC.addRow([`CAMP_${cId++}`, l.nome, l.matches]));

    // Squadre
    const shS = workbook.addWorksheet('Squadre');
    shS.addRow(['ID', 'Squadra', 'Campionato / Fonte']);
    let sId = 1;
    uniqueTeams.forEach((src, tName) => shS.addRow([sId++, tName, src]));

    // MetisIndex
    const shMI = workbook.addWorksheet('MetisIndex');
    shMI.addRow(['Squadra', 'Metis Index', 'Forma (20%)', 'Attacco (15%)', 'Difesa (15%)', 'Creazione (15%)', 'Performance (10%)', 'Casa/Trasferta (10%)', 'Efficienza (5%)', 'Trend (10%)']);
    uniqueTeams.forEach((_, tName) => {
        const mi = metisIndexResults.get(tName) || { metisIndex: 50.0, forma: 50, attacco: 50, difesa: 50, creazione: 50, performance: 50, casaTrasferta: 50, efficienza: 50, trendVal: '0.0' };
        shMI.addRow([tName, mi.metisIndex, mi.forma, mi.attacco, mi.difesa, mi.creazione, mi.performance, mi.casaTrasferta, mi.efficienza, mi.trendVal]);
    });

    // H2H
    const shH = workbook.addWorksheet('H2H');
    shH.addRow(['Confronto', 'Squadra A', 'Squadra B', 'Match Totali', 'Vittorie A', 'Pareggi', 'Vittorie B', '% Vittoria A', 'Media Gol Totali']);
    h2hMap.forEach(h => {
        const pctA = ((h.winA / h.matches) * 100).toFixed(1) + '%';
        const avgG = ((h.goalsA + h.goalsB) / h.matches).toFixed(2);
        shH.addRow([`${h.teamA} vs ${h.teamB}`, h.teamA, h.teamB, h.matches, h.winA, h.draws, h.winB, pctA, avgG]);
    });

    // StatisticheSquadra
    const shSt = workbook.addWorksheet('StatisticheSquadra');
    shSt.addRow(['Squadra', 'Campionato', 'Gare 26/27', 'Casa G', 'Casa V', 'Casa P', 'Casa S', 'Media GF Casa', 'Media GS Casa', 'Fuori G', 'Fuori V', 'Fuori P', 'Fuori S', 'Media GF Fuori', 'Media GS Fuori', '% Over 2.5', '% Gol/Gol', '% Clean Sheet']);
    teamSeasonStatsMap.forEach(st => {
        const tot = st.pH + st.pA;
        if (tot > 0) {
            shSt.addRow([st.name, st.league, tot, st.pH, st.wH, st.dH, st.lH, (st.gfH / Math.max(1, st.pH)).toFixed(2), (st.gaH / Math.max(1, st.pH)).toFixed(2), st.pA, st.wA, st.dA, st.lA, (st.gfA / Math.max(1, st.pA)).toFixed(2), (st.gaA / Math.max(1, st.pA)).toFixed(2), ((st.over25 / tot) * 100).toFixed(1) + '%', ((st.gg / tot) * 100).toFixed(1) + '%', ((st.cs / tot) * 100).toFixed(1) + '%']);
        }
    });

    // Classifiche
    const shCl = workbook.addWorksheet('Classifiche');
    shCl.addRow(['Campionato', 'Posizione', 'Squadra', 'Punti', 'Giocate', 'Vinte', 'Nulle', 'Perse', 'Gol Fatti', 'Gol Subiti', 'Diff Reti']);
    activeLeaguesMap.forEach(l => {
        const sorted = Array.from(l.teams.values()).sort((a, b) => b.pts - a.pts || (b.gf - b.ga) - (a.gf - a.ga));
        sorted.forEach((t, idx) => {
            shCl.addRow([l.nome, idx + 1, t.name, t.pts, t.p, t.w, t.d, t.l, t.gf, t.ga, t.gf - t.ga]);
        });
    });

    await workbook.xlsx.writeFile(EXCEL_PATH);
    console.log(`💾 File Excel salvato con successo con Report Predittivo Completo!`);
}

runMetisMasterEngine().catch(console.error);