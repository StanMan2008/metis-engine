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

// 42 Campionati Ufficiali Monitorati
const CAMPIONATI_METIS = [
    { id: 'SA', paese: 'Italia', nome: 'Italian Serie A', slug: 'serie-a' },
    { id: 'SB', paese: 'Italia', nome: 'Italian Serie B', slug: 'serie-b' },
    { id: 'PL', paese: 'Inghilterra', nome: 'Premier League', slug: 'premier-league' },
    { id: 'CH', paese: 'Inghilterra', nome: 'Championship', slug: 'championship' },
    { id: 'L1', paese: 'Inghilterra', nome: 'League One', slug: 'league-one' },
    { id: 'L2', paese: 'Inghilterra', nome: 'League Two', slug: 'league-two' },
    { id: 'PD', paese: 'Spagna', nome: 'La Liga', slug: 'la-liga' },
    { id: 'SD', paese: 'Spagna', nome: 'La Liga 2', slug: 'la-liga-2' },
    { id: 'BL1', paese: 'Germania', nome: 'Bundesliga', slug: 'bundesliga' },
    { id: 'BL2', paese: 'Germania', nome: '2. Bundesliga', slug: '2-bundesliga' },
    { id: 'FL1', paese: 'Francia', nome: 'Ligue 1', slug: 'ligue-1' },
    { id: 'FL2', paese: 'Francia', nome: 'Ligue 2', slug: 'ligue-2' },
    { id: 'DED', paese: 'Olanda', nome: 'Eredivisie', slug: 'eredivisie' },
    { id: 'PPL', paese: 'Portogallo', nome: 'Liga Portugal', slug: 'liga-portugal' },
    { id: 'BJL', paese: 'Belgio', nome: 'Jupiler Pro League', slug: 'jupiler-league' },
    { id: 'TSL', paese: 'Turchia', nome: 'Süper Lig', slug: 'super-lig' },
    { id: 'GSL', paese: 'Grecia', nome: 'Super League', slug: 'super-league-grecia' },
    { id: 'SC0', paese: 'Scozia', nome: 'Premiership', slug: 'premiership-scozia' },
    { id: 'SW1', paese: 'Svizzera', nome: 'Super League', slug: 'super-league-svizzera' },
    { id: 'AUT', paese: 'Austria', nome: 'Bundesliga', slug: 'bundesliga-austria' }
];

async function rebuildCleanDatabase() {
    console.log('🚀 [METIS ENGINE] Ricostruzione Totale Database in Corso...\n');

    // 1. Crea un workbook completamente nuovo (elimina qualsiasi foglio precedente)
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Metis Engine';
    workbook.created = new Date();

    // 2. FOGLIO README
    const sheetReadme = workbook.addWorksheet('README');
    sheetReadme.addRow(['METIS ENGINE - ARCHITETTURA DATABASE']);
    sheetReadme.addRow(['Data Ricostruzione', '2026-10-06']);
    sheetReadme.addRow(['Stagione Attiva', '2026/2027']);
    sheetReadme.addRow(['Regola Storico', 'Decennale (usato unicamente per matrice H2H)']);
    sheetReadme.addRow(['Regola Classifiche', 'Solo partite disputate da luglio a oggi (06/10/2026)']);

    // 3. FOGLIO Campionati
    const sheetCampionati = workbook.addWorksheet('Campionati');
    sheetCampionati.addRow(['ID', 'Paese', 'Campionato', 'Slug']);
    CAMPIONATI_METIS.forEach(c => sheetCampionati.addRow([c.id, c.paese, c.nome, c.slug]));

    // 4. Inizializzazione Struttura Dati
    const leagueMap = new Map();
    CAMPIONATI_METIS.forEach(c => {
        leagueMap.set(clean(c.nome), {
            id: c.id,
            paese: c.paese,
            nome: c.nome,
            matchesSeason: 0,
            teamsSeason: new Map()
        });
    });

    const h2hMap = new Map();
    const teamSeasonStats = new Map();
    const knownTeams = new Map();

    const dirs = [MAIN_DIR, EXTRA_DIR];
    let totalHistoricalMatches = 0;
    let totalSeasonMatches = 0;

    dirs.forEach(dir => {
        if (!fs.existsSync(dir)) return;
        const files = fs.readdirSync(dir);

        files.forEach(f => {
            if (!f.endsWith('.csv')) return;
            const normFileName = clean(f.replace('.csv', ''));

            let league = null;
            for (const [key, l] of leagueMap.entries()) {
                if (normFileName.includes(key) || key.includes(normFileName)) {
                    league = l;
                    break;
                }
            }

            try {
                const rows = parseCSV(fs.readFileSync(path.join(dir, f), 'utf-8'));
                rows.forEach(m => {
                    const home = (m.HomeTeam || m.Home || m.home || '').trim();
                    const away = (m.AwayTeam || m.Away || m.away || '').trim();
                    const fthg = Number(m.FTHG ?? m.HG ?? m.home_score);
                    const ftag = Number(m.FTAG ?? m.AG ?? m.away_score);
                    const dateObj = parseMatchDate(m.Date || m.date);

                    if (home && away && !isNaN(fthg) && !isNaN(ftag)) {
                        totalHistoricalMatches++;

                        // Memorizza Squadre
                        if (!knownTeams.has(home)) knownTeams.set(home, league ? league.nome : 'Extra League');
                        if (!knownTeams.has(away)) knownTeams.set(away, league ? league.nome : 'Extra League');

                        // Matrice H2H Decennale
                        const pairKey = [home, away].sort().join(' vs ');
                        if (!h2hMap.has(pairKey)) {
                            h2hMap.set(pairKey, { teamA: home, teamB: away, matches: 0, winA: 0, winB: 0, draws: 0, goalsA: 0, goalsB: 0 });
                        }
                        const h = h2hMap.get(pairKey);
                        h.matches++;
                        if (fthg > ftag) {
                            if (home === h.teamA) h.winA++; else h.winB++;
                        } else if (fthg === ftag) {
                            h.draws++;
                        } else {
                            if (away === h.teamB) h.winB++; else h.winA++;
                        }
                        if (home === h.teamA) { h.goalsA += fthg; h.goalsB += ftag; }
                        else { h.goalsB += fthg; h.goalsA += ftag; }

                        // Statistiche e Classifica: SOLO STAGIONE 2026/2027
                        if (dateObj && dateObj >= SEASON_START && dateObj <= TODAY) {
                            totalSeasonMatches++;
                            if (league) league.matchesSeason++;

                            const initStat = (tName, lName) => ({
                                name: tName, league: lName,
                                pH: 0, wH: 0, dH: 0, lH: 0, gfH: 0, gaH: 0,
                                pA: 0, wA: 0, dA: 0, lA: 0, gfA: 0, gaA: 0,
                                over25: 0, gg: 0, cs: 0
                            });

                            if (!teamSeasonStats.has(home)) teamSeasonStats.set(home, initStat(home, league ? league.nome : 'Campionato'));
                            if (!teamSeasonStats.has(away)) teamSeasonStats.set(away, initStat(away, league ? league.nome : 'Campionato'));

                            const sH = teamSeasonStats.get(home);
                            const sA = teamSeasonStats.get(away);

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

                            // Classifica
                            if (league) {
                                const updateTable = (tName, gf, ga, res) => {
                                    if (!league.teamsSeason.has(tName)) {
                                        league.teamsSeason.set(tName, { name: tName, p: 0, w: 0, d: 0, l: 0, gf: 0, ga: 0, pts: 0 });
                                    }
                                    const t = league.teamsSeason.get(tName);
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

    // 5. FOGLIO Squadre
    const sheetSquadre = workbook.addWorksheet('Squadre');
    sheetSquadre.addRow(['ID', 'Squadra', 'Campionato']);
    let teamCounter = 1;
    for (const [tName, lName] of knownTeams.entries()) {
        sheetSquadre.addRow([teamCounter++, tName, lName]);
    }

    // 6. FOGLIO H2H (Scontri Diretti Decennali)
    const sheetH2H = workbook.addWorksheet('H2H');
    sheetH2H.addRow(['Confronto', 'Squadra A', 'Squadra B', 'Match Totali', 'Vittorie A', 'Pareggi', 'Vittorie B', '% Vittoria A', 'Media Gol Totali']);
    for (const [_, h] of h2hMap.entries()) {
        const pctA = ((h.winA / h.matches) * 100).toFixed(1) + '%';
        const avgGoals = ((h.goalsA + h.goalsB) / h.matches).toFixed(2);
        sheetH2H.addRow([`${h.teamA} vs ${h.teamB}`, h.teamA, h.teamB, h.matches, h.winA, h.draws, h.winB, pctA, avgGoals]);
    }

    // 7. FOGLIO StatisticheSquadra (Solo Stagione 2026/2027)
    const sheetStat = workbook.addWorksheet('StatisticheSquadra');
    sheetStat.addRow([
        'Squadra', 'Campionato', 'Gare 26/27',
        'Casa G', 'Casa V', 'Casa P', 'Casa S', 'Media GF Casa', 'Media GS Casa',
        'Fuori G', 'Fuori V', 'Fuori P', 'Fuori S', 'Media GF Fuori', 'Media GS Fuori',
        '% Over 2.5', '% Gol/Gol', '% Clean Sheet'
    ]);

    for (const [_, s] of teamSeasonStats.entries()) {
        const tot = s.pH + s.pA;
        if (tot > 0) {
            sheetStat.addRow([
                s.name, s.league, tot,
                s.pH, s.wH, s.dH, s.lH, (s.gfH / Math.max(1, s.pH)).toFixed(2), (s.gaH / Math.max(1, s.pH)).toFixed(2),
                s.pA, s.wA, s.dA, s.lA, (s.gfA / Math.max(1, s.pA)).toFixed(2), (s.gaA / Math.max(1, s.pA)).toFixed(2),
                ((s.over25 / tot) * 100).toFixed(1) + '%',
                ((s.gg / tot) * 100).toFixed(1) + '%',
                ((s.cs / tot) * 100).toFixed(1) + '%'
            ]);
        }
    }

    // 8. FOGLIO Classifiche (Solo Stagione 2026/2027)
    const sheetClass = workbook.addWorksheet('Classifiche');
    sheetClass.addRow(['Campionato', 'Posizione', 'Squadra', 'Punti', 'Giocate', 'Vinte', 'Nulle', 'Perse', 'Gol Fatti', 'Gol Subiti', 'Diff Reti']);
    for (const [_, l] of leagueMap.entries()) {
        if (l.teamsSeason.size > 0) {
            const sorted = Array.from(l.teamsSeason.values()).sort((a, b) => b.pts - a.pts || (b.gf - b.ga) - (a.gf - a.ga));
            sorted.forEach((t, idx) => {
                sheetClass.addRow([l.nome, idx + 1, t.name, t.pts, t.p, t.w, t.d, t.l, t.gf, t.ga, t.gf - t.ga]);
            });
        }
    }

    // 9. FOGLIO Fixtures & Value Bet (Pronto per l'operatività)
    const sheetFixtures = workbook.addWorksheet('Fixtures');
    sheetFixtures.addRow(['Data', 'Campionato', 'Casa', 'Trasferta', 'Quota Fair 1', 'Quota Fair X', 'Quota Fair 2', 'Over 2.5', 'Gol/Gol']);

    const sheetReport = workbook.addWorksheet('Report');
    sheetReport.addRow(['Data Match', 'Campionato', 'Incontro', 'Mercato', 'Quota Fair (Poisson)', 'Quota Bookmaker', 'Edge %', 'Rating Valore']);

    // Formattazione Globale
    workbook.worksheets.forEach(sh => {
        sh.views = [{ state: 'frozen', ySplit: 1 }];
        sh.columns.forEach(col => {
            let maxLen = 12;
            col.eachCell({ includeEmpty: false }, cell => {
                const len = cell.value ? cell.value.toString().length : 0;
                if (len > maxLen) maxLen = Math.min(len + 3, 28);
            });
            col.width = maxLen;
        });
    });

    await workbook.xlsx.writeFile(EXCEL_PATH);

    console.log(`======================================================`);
    console.log(`✅ DATABASE RICOSTRUITO DA ZERO!`);
    console.log(`📁 File Excel: ${EXCEL_PATH}`);
    console.log(`⚽ Partite storiche H2H caricate: ${totalHistoricalMatches}`);
    console.log(`📅 Partite stagione 2026/2027 (fino a oggi): ${totalSeasonMatches}`);
    console.log(`🛡️ Squadre censite: ${teamCounter - 1}`);
    console.log(`📊 Fogli creati: README, Campionati, Squadre, H2H, StatisticheSquadra, Classifiche, Fixtures, Report`);
    console.log(`🧹 Tutti i fogli atleti superflui sono stati rimossi.`);
    console.log(`======================================================\n`);
}

rebuildCleanDatabase().catch(console.error);