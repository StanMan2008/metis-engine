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

async function rebuildFromSourceFiles() {
    console.log('🔄 [METIS ENGINE] Scansione file per file (Main & Extra Leagues)...\n');

    if (!fs.existsSync(EXCEL_PATH)) {
        console.error(`❌ File non trovato: ${EXCEL_PATH}`);
        return;
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(EXCEL_PATH);

    const leaguesMap = new Map();
    const teamSeasonStats = new Map();
    const uniqueTeams = new Map();
    const h2hMap = new Map();

    const targets = [
        { dir: MAIN_DIR, type: 'Campionato Principale (Top)' },
        { dir: EXTRA_DIR, type: 'Campionato Extra / Internazionale' }
    ];

    let filesProcessed = 0;
    let totalMatches = 0;
    let seasonMatches = 0;

    targets.forEach(t => {
        if (!fs.existsSync(t.dir)) return;
        const files = fs.readdirSync(t.dir);

        files.forEach(f => {
            if (!f.endsWith('.csv')) return;
            filesProcessed++;
            const leagueName = f.replace('.csv', '').replace(/[-_]/g, ' ').trim();
            const lKey = clean(leagueName);

            if (!leaguesMap.has(lKey)) {
                leaguesMap.set(lKey, {
                    nome: leagueName,
                    tipo: t.type,
                    matchesSeason: 0,
                    teamsSeason: new Map()
                });
            }
            const currentLeague = leaguesMap.get(lKey);

            try {
                const content = fs.readFileSync(path.join(t.dir, f), 'utf-8');
                const rows = parseCSV(content);

                rows.forEach(m => {
                    const home = (m.HomeTeam || m.Home || m.home || '').trim();
                    const away = (m.AwayTeam || m.Away || m.away || '').trim();
                    const fthg = Number(m.FTHG ?? m.HG ?? m.home_score);
                    const ftag = Number(m.FTAG ?? m.AG ?? m.away_score);
                    const dateObj = parseMatchDate(m.Date || m.date);

                    if (home && away && !isNaN(fthg) && !isNaN(ftag)) {
                        totalMatches++;

                        // Registra squadra reale associata al campionato del file
                        if (home && !uniqueTeams.has(home)) uniqueTeams.set(home, currentLeague.nome);
                        if (away && !uniqueTeams.has(uniqueTeams.has(away) ? away : away)) uniqueTeams.set(away, currentLeague.nome);

                        // H2H Decennale
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

                        // Stagione 2026/2027
                        if (dateObj && dateObj >= SEASON_START && dateObj <= TODAY) {
                            seasonMatches++;
                            currentLeague.matchesSeason++;

                            const initStat = (tName, lName) => ({
                                name: tName, league: lName,
                                pH: 0, wH: 0, dH: 0, lH: 0, gfH: 0, gaH: 0,
                                pA: 0, wA: 0, dA: 0, lA: 0, gfA: 0, gaA: 0,
                                over25: 0, gg: 0, cs: 0
                            });

                            if (!teamSeasonStats.has(home)) teamSeasonStats.set(home, initStat(home, currentLeague.nome));
                            if (!teamSeasonStats.has(away)) teamSeasonStats.set(away, initStat(away, currentLeague.nome));

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

                            const updateTable = (tName, gf, ga, res) => {
                                if (!currentLeague.teamsSeason.has(tName)) {
                                    currentLeague.teamsSeason.set(tName, { name: tName, p: 0, w: 0, d: 0, l: 0, gf: 0, ga: 0, pts: 0 });
                                }
                                const t = currentLeague.teamsSeason.get(tName);
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
                });
            } catch {}
        });
    });

    console.log(`📂 File CSV elaborati: ${filesProcessed}`);
    console.log(`⚽ Partite storiche totali: ${totalMatches}`);
    console.log(`📅 Partite stagione 2026/2027: ${seasonMatches}`);
    console.log(`🏆 Campionati unici rilevati: ${leaguesMap.size}`);
    console.log(`🛡️ Squadre reali uniche censite: ${uniqueTeams.size}\n`);

    // Scrittura Fogli Puliti
    // 1. Campionati
    let shC = workbook.getWorksheet('Campionati');
    if (shC) workbook.removeWorksheet(shC.id);
    shC = workbook.addWorksheet('Campionati');
    shC.addRow(['ID', 'Tipo', 'Campionato', 'Match Stagionali 26/27']);
    let cIdx = 1;
    for (const [_, l] of leaguesMap.entries()) {
        shC.addRow([`CAMP_${cIdx++}`, l.tipo, l.nome, l.matchesSeason]);
    }

    // 2. Squadre (100% reali dai file)
    let shS = workbook.getWorksheet('Squadre');
    if (shS) workbook.removeWorksheet(shS.id);
    shS = workbook.addWorksheet('Squadre');
    shS.addRow(['ID', 'Squadra Reale', 'Campionato di Appartenza']);
    let sIdx = 1;
    for (const [tName, lName] of uniqueTeams.entries()) {
        shS.addRow([sIdx++, tName, lName]);
    }

    // 3. H2H
    let shH = workbook.getWorksheet('H2H');
    if (shH) workbook.removeWorksheet(shH.id);
    shH = workbook.addWorksheet('H2H');
    shH.addRow(['Confronto', 'Squadra A', 'Squadra B', 'Match Totali', 'Vittorie A', 'Pareggi', 'Vittorie B', '% Vittoria A', 'Media Gol Totali']);
    for (const [_, h] of h2hMap.entries()) {
        const pctA = ((h.winA / h.matches) * 100).toFixed(1) + '%';
        const avgGoals = ((h.goalsA + h.goalsB) / h.matches).toFixed(2);
        shH.addRow([`${h.teamA} vs ${h.teamB}`, h.teamA, h.teamB, h.matches, h.winA, h.draws, h.winB, pctA, avgGoals]);
    }

    // 4. StatisticheSquadra
    let shSt = workbook.getWorksheet('StatisticheSquadra');
    if (shSt) workbook.removeWorksheet(shSt.id);
    shSt = workbook.addWorksheet('StatisticheSquadra');
    shSt.addRow([
        'Squadra', 'Campionato', 'Gare 26/27',
        'Casa G', 'Casa V', 'Casa P', 'Casa S', 'Media GF Casa', 'Media GS Casa',
        'Fuori G', 'Fuori V', 'Fuori P', 'Fuori S', 'Media GF Fuori', 'Media GS Fuori',
        '% Over 2.5', '% Gol/Gol', '% Clean Sheet'
    ]);
    for (const [_, s] of teamSeasonStats.entries()) {
        const tot = s.pH + s.pA;
        if (tot > 0) {
            shSt.addRow([
                s.name, s.league, tot,
                s.pH, s.wH, s.dH, s.lH, (s.gfH / Math.max(1, s.pH)).toFixed(2), (s.gaH / Math.max(1, s.pH)).toFixed(2),
                s.pA, s.wA, s.dA, s.lA, (s.gfA / Math.max(1, s.pA)).toFixed(2), (s.gaA / Math.max(1, s.pA)).toFixed(2),
                ((s.over25 / tot) * 100).toFixed(1) + '%',
                ((s.gg / tot) * 100).toFixed(1) + '%',
                ((s.cs / tot) * 100).toFixed(1) + '%'
            ]);
        }
    }

    // 5. Classifiche
    let shCl = workbook.getWorksheet('Classifiche');
    if (shCl) workbook.removeWorksheet(shCl.id);
    shCl = workbook.addWorksheet('Classifiche');
    shCl.addRow(['Campionato', 'Posizione', 'Squadra', 'Punti', 'Giocate', 'Vinte', 'Nulle', 'Perse', 'Gol Fatti', 'Gol Subiti', 'Diff Reti']);
    for (const [_, l] of leaguesMap.entries()) {
        if (l.teamsSeason.size > 0) {
            const sorted = Array.from(l.teamsSeason.values()).sort((a, b) => b.pts - a.pts || (b.gf - b.ga) - (a.gf - a.ga));
            sorted.forEach((t, idx) => {
                shCl.addRow([l.nome, idx + 1, t.name, t.pts, t.p, t.w, t.d, t.l, t.gf, t.ga, t.gf - t.ga]);
            });
        }
    }

    // Formattazione colonne
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
    console.log(`💾 File Excel salvato: ${EXCEL_PATH}`);
    console.log(`✅ RIGENERAZIONE COMPLETATA CON SUCCESSO DA SORGENTE!`);
}

rebuildFromSourceFiles().catch(console.error);