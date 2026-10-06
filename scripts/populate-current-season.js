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

// Limite temporale stagione 2026/2027
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

// Converte date formato DD/MM/YYYY, DD/MM/YY o YYYY-MM-DD
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

async function safeWriteWorkbook(workbook, filePath, retries = 3) {
    for (let i = 0; i < retries; i++) {
        try {
            await workbook.xlsx.writeFile(filePath);
            console.log(`💾 File Excel salvato: ${filePath}`);
            return;
        } catch (err) {
            if (err.code === 'EBUSY') {
                console.warn(`⚠️ File Excel occupato. Riprovo (${i + 1}/${retries})...`);
                await new Promise(r => setTimeout(r, 2000));
            } else throw err;
        }
    }
}

async function aggiornaStagione2026() {
    console.log('🚀 [METIS ENGINE] Ricalcolo Statistiche e Classifiche: STAGIONE 2026/2027 (Fino al 06/10/2026)...\n');

    if (!fs.existsSync(EXCEL_PATH)) {
        console.error(`❌ File non trovato: ${EXCEL_PATH}`);
        return;
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(EXCEL_PATH);

    const sheetCampionati = workbook.getWorksheet('Campionati');
    if (!sheetCampionati) {
        console.error('❌ Foglio "Campionati" non trovato.');
        return;
    }

    const leagueData = new Map();
    sheetCampionati.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;
        const id = row.getCell(1).value;
        const paese = row.getCell(2).value;
        const nome = row.getCell(3).value;
        const slug = row.getCell(4).value;

        if (nome) {
            leagueData.set(clean(nome), {
                id,
                paese,
                nome,
                slug,
                matches: 0,
                homeWins: 0,
                draws: 0,
                awayWins: 0,
                totalGoals: 0,
                over25: 0,
                gg: 0,
                teams: new Map()
            });
        }
    });

    const teamSeasonStats = new Map();
    const dirs = [MAIN_DIR, EXTRA_DIR];
    let validMatchesCount = 0;

    dirs.forEach(dir => {
        if (!fs.existsSync(dir)) return;
        const files = fs.readdirSync(dir);

        files.forEach(f => {
            if (!f.endsWith('.csv')) return;
            const normFileName = clean(f.replace('.csv', ''));

            let targetLeague = null;
            for (const [key, l] of leagueData.entries()) {
                if (normFileName.includes(key) || key.includes(normFileName) || (l.slug && normFileName.includes(clean(l.slug)))) {
                    targetLeague = l;
                    break;
                }
            }

            if (!targetLeague) {
                for (const [_, l] of leagueData.entries()) {
                    const parts = clean(l.nome).split(/\s+/).filter(p => p.length >= 5);
                    if (parts.some(p => normFileName.includes(p))) {
                        targetLeague = l;
                        break;
                    }
                }
            }

            if (!targetLeague) return;

            try {
                const content = fs.readFileSync(path.join(dir, f), 'utf-8');
                const rows = parseCSV(content);

                rows.forEach(m => {
                    const dateObj = parseMatchDate(m.Date || m.date);
                    
                    // FILTRO TEMPORALE RIGIDO: Solo gare disputate nella season 2026/27 fino a oggi
                    if (!dateObj || dateObj < SEASON_START || dateObj > TODAY) {
                        return;
                    }

                    const home = (m.HomeTeam || m.Home || m.home || '').trim();
                    const away = (m.AwayTeam || m.Away || m.away || '').trim();
                    const fthg = Number(m.FTHG ?? m.HG ?? m.home_score);
                    const ftag = Number(m.FTAG ?? m.AG ?? m.away_score);

                    if (home && away && !isNaN(fthg) && !isNaN(ftag)) {
                        validMatchesCount++;
                        targetLeague.matches++;
                        targetLeague.totalGoals += (fthg + ftag);

                        if (fthg > ftag) targetLeague.homeWins++;
                        else if (fthg === ftag) targetLeague.draws++;
                        else targetLeague.awayWins++;

                        if (fthg + ftag > 2.5) targetLeague.over25++;
                        if (fthg > 0 && ftag > 0) targetLeague.gg++;

                        // Inizializzazione statistiche club 26/27
                        const initStat = (name, league) => ({
                            name, league,
                            playedH: 0, winH: 0, drawH: 0, lossH: 0, scoredH: 0, concededH: 0,
                            playedA: 0, winA: 0, drawA: 0, lossA: 0, scoredA: 0, concededA: 0,
                            over25: 0, gg: 0, cleanSheet: 0
                        });

                        const kHome = clean(home);
                        const kAway = clean(away);

                        if (!teamSeasonStats.has(kHome)) teamSeasonStats.set(kHome, initStat(home, targetLeague.nome));
                        if (!teamSeasonStats.has(kAway)) teamSeasonStats.set(kAway, initStat(away, targetLeague.nome));

                        const sH = teamSeasonStats.get(kHome);
                        const sA = teamSeasonStats.get(kAway);

                        sH.playedH++;
                        sH.scoredH += fthg;
                        sH.concededH += ftag;
                        if (fthg > ftag) sH.winH++;
                        else if (fthg === ftag) sH.drawH++;
                        else sH.lossH++;
                        if (ftag === 0) sH.cleanSheet++;
                        if (fthg + ftag > 2.5) sH.over25++;
                        if (fthg > 0 && ftag > 0) sH.gg++;

                        sA.playedA++;
                        sA.scoredA += ftag;
                        sA.concededA += fthg;
                        if (ftag > fthg) sA.winA++;
                        else if (ftag === ftag) sA.drawA++;
                        else sA.lossA++;
                        if (fthg === 0) sA.cleanSheet++;
                        if (fthg + ftag > 2.5) sA.over25++;
                        if (fthg > 0 && ftag > 0) sA.gg++;

                        // Aggiornamento Classifica 2026/2027
                        const updateTable = (teamName, gf, gs, res) => {
                            if (!targetLeague.teams.has(teamName)) {
                                targetLeague.teams.set(teamName, {
                                    name: teamName,
                                    p: 0, w: 0, d: 0, l: 0, gf: 0, ga: 0, pts: 0
                                });
                            }
                            const t = targetLeague.teams.get(teamName);
                            t.p++;
                            t.gf += gf;
                            t.ga += gs;
                            if (res === 'W') { t.w++; t.pts += 3; }
                            else if (res === 'D') { t.d++; t.pts += 1; }
                            else { t.l++; }
                        };

                        if (fthg > ftag) {
                            updateTable(home, fthg, ftag, 'W');
                            updateTable(away, ftag, fthg, 'L');
                        } else if (fthg === ftag) {
                            updateTable(home, fthg, ftag, 'D');
                            updateTable(away, ftag, fthg, 'D');
                        } else {
                            updateTable(home, fthg, ftag, 'L');
                            updateTable(away, ftag, fthg, 'W');
                        }
                    }
                });
            } catch {}
        });
    });

    console.log(`⚽ Partite effettive stagione 2026/2027 (fino a oggi): ${validMatchesCount}\n`);

    // 1. Riscrivi StatisticheSquadra
    let sheetStat = workbook.getWorksheet('StatisticheSquadra');
    if (sheetStat) workbook.removeWorksheet(sheetStat.id);
    sheetStat = workbook.addWorksheet('StatisticheSquadra');

    sheetStat.addRow([
        'Squadra', 'Campionato', 'Gare 26/27',
        'Casa Giocate', 'Vittorie C', 'Pari C', 'Sconfitte C', 'Media Gol Fatti C', 'Media Gol Subiti C',
        'Fuori Giocate', 'Vittorie F', 'Pari F', 'Sconfitte F', 'Media Gol Fatti F', 'Media Gol Subiti F',
        '% Over 2.5', '% Gol/Gol', '% Clean Sheet'
    ]);

    let countSquadre = 0;
    for (const [_, s] of teamSeasonStats.entries()) {
        const totPlayed = s.playedH + s.playedA;
        if (totPlayed > 0) {
            const avgScoredH = s.playedH > 0 ? (s.scoredH / s.playedH).toFixed(2) : '0.00';
            const avgConcededH = s.playedH > 0 ? (s.concededH / s.playedH).toFixed(2) : '0.00';
            const avgScoredA = s.playedA > 0 ? (s.scoredA / s.playedA).toFixed(2) : '0.00';
            const avgConcededA = s.playedA > 0 ? (s.concededA / s.playedA).toFixed(2) : '0.00';

            const pctOver = ((s.over25 / totPlayed) * 100).toFixed(1) + '%';
            const pctGG = ((s.gg / totPlayed) * 100).toFixed(1) + '%';
            const pctCS = ((s.cleanSheet / totPlayed) * 100).toFixed(1) + '%';

            sheetStat.addRow([
                s.name, s.league, totPlayed,
                s.playedH, s.winH, s.drawH, s.lossH, avgScoredH, avgConcededH,
                s.playedA, s.winA, s.drawA, s.lossA, avgScoredA, avgConcededA,
                pctOver, pctGG, pctCS
            ]);
            countSquadre++;
        }
    }

    // 2. Riscrivi Classifiche
    let sheetClass = workbook.getWorksheet('Classifiche');
    if (sheetClass) workbook.removeWorksheet(sheetClass.id);
    sheetClass = workbook.addWorksheet('Classifiche');

    sheetClass.addRow([
        'Campionato', 'Paese', 'Posizione', 'Squadra', 'Punti',
        'Giocate', 'Vinte', 'Nulle', 'Perse', 'Gol Fatti', 'Gol Subiti', 'Diff Reti'
    ]);

    let countClass = 0;
    for (const [_, l] of leagueData.entries()) {
        if (l.teams.size > 0) {
            const sorted = Array.from(l.teams.values()).sort((a, b) => {
                if (b.pts !== a.pts) return b.pts - a.pts;
                const diffB = b.gf - b.ga;
                const diffA = a.gf - a.ga;
                if (diffB !== diffA) return diffB - diffA;
                return b.gf - a.gf;
            });

            sorted.forEach((t, idx) => {
                sheetClass.addRow([
                    l.nome, l.paese, idx + 1, t.name, t.pts,
                    t.p, t.w, t.d, t.l, t.gf, t.ga, (t.gf - t.ga)
                ]);
                countClass++;
            });
        }
    }

    // 3. Svuota fogli giocatori sintetici non certificati
    ['Giocatori', 'StatisticheGiocatori', 'StatistichePortieri', 'TopGiocatori'].forEach(name => {
        let sh = workbook.getWorksheet(name);
        if (sh) {
            workbook.removeWorksheet(sh.id);
            const emptySh = workbook.addWorksheet(name);
            emptySh.addRow(['ID', 'Nome', 'Squadra', 'Stato Dati']);
            emptySh.addRow(['-', '-', '-', 'In attesa di feed ufficiale marcatori 26/27']);
        }
    });

    [sheetStat, sheetClass].forEach(sh => {
        sh.views = [{ state: 'frozen', ySplit: 1 }];
        sh.columns.forEach(col => {
            let maxLen = 12;
            col.eachCell({ includeEmpty: false }, cell => {
                const len = cell.value ? cell.value.toString().length : 0;
                if (len > maxLen) maxLen = Math.min(len + 3, 26);
            });
            col.width = maxLen;
        });
    });

    await safeWriteWorkbook(workbook, EXCEL_PATH);
    console.log(`======================================================`);
    console.log(`✅ FOGLI AGGIORNATI CON I SOLI DATI REALI DELLA STAGIONE 2026/2027!`);
    console.log(`📊 Squadre con statistiche 2026/2027: ${countSquadre}`);
    console.log(`🏆 Squadre in classifica 2026/2027: ${countClass}`);
    console.log(`🧹 Fogli atleti sintetici azzerati in attesa di dati certificati`);
    console.log(`======================================================\n`);
}

aggiornaStagione2026().catch(console.error);
