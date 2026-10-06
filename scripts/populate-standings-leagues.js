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

async function safeWriteWorkbook(workbook, filePath, retries = 3) {
    for (let i = 0; i < retries; i++) {
        try {
            await workbook.xlsx.writeFile(filePath);
            console.log(`💾 File Excel salvato: ${filePath}`);
            return;
        } catch (err) {
            if (err.code === 'EBUSY') {
                console.warn(`⚠️ File Excel occupato. Tentativo ${i + 1}/${retries}...`);
                await new Promise(r => setTimeout(r, 2000));
            } else throw err;
        }
    }
}

async function generaClassificheEStatistiche() {
    console.log('🚀 [METIS ENGINE] Calcolo Classifiche e Statistiche Campionati...\n');

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

    // Struttura dati per aggregazione per campionato
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
                teams: new Map() // Per la classifica
            });
        }
    });

    const dirs = [MAIN_DIR, EXTRA_DIR];
    dirs.forEach(dir => {
        if (!fs.existsSync(dir)) return;
        const files = fs.readdirSync(dir);
        files.forEach(f => {
            if (!f.endsWith('.csv')) return;
            const normFileName = clean(f.replace('.csv', ''));

            // Trova campionato corrispondente
            let targetLeague = null;
            for (const [key, l] of leagueData.entries()) {
                if (normFileName.includes(key) || key.includes(normFileName) || (l.slug && normFileName.includes(clean(l.slug)))) {
                    targetLeague = l;
                    break;
                }
            }

            if (!targetLeague) {
                // Fallback su parole chiave
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
                    const home = (m.HomeTeam || m.Home || m.home || '').trim();
                    const away = (m.AwayTeam || m.Away || m.away || '').trim();
                    const fthg = Number(m.FTHG ?? m.HG ?? m.home_score);
                    const ftag = Number(m.FTAG ?? m.AG ?? m.away_score);

                    if (home && away && !isNaN(fthg) && !isNaN(ftag)) {
                        targetLeague.matches++;
                        targetLeague.totalGoals += (fthg + ftag);

                        if (fthg > ftag) targetLeague.homeWins++;
                        else if (fthg === ftag) targetLeague.draws++;
                        else targetLeague.awayWins++;

                        if (fthg + ftag > 2.5) targetLeague.over25++;
                        if (fthg > 0 && ftag > 0) targetLeague.gg++;

                        // Aggiorna Classifica Squadra
                        const updateTeam = (teamName, gf, gs, res) => {
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
                            updateTeam(home, fthg, ftag, 'W');
                            updateTeam(away, ftag, fthg, 'L');
                        } else if (fthg === ftag) {
                            updateTeam(home, fthg, ftag, 'D');
                            updateTeam(away, ftag, fthg, 'D');
                        } else {
                            updateTeam(home, fthg, ftag, 'L');
                            updateTeam(away, ftag, fthg, 'W');
                        }
                    }
                });
            } catch {}
        });
    });

    // 1. Popolamento Foglio "StatisticheCampionati"
    let sheetStatCamp = workbook.getWorksheet('StatisticheCampionati');
    if (sheetStatCamp) workbook.removeWorksheet(sheetStatCamp.id);
    sheetStatCamp = workbook.addWorksheet('StatisticheCampionati');

    sheetStatCamp.addRow([
        'ID Campionato', 'Campionato', 'Paese', 'Gare Monitorate',
        'Fattore Casa (% 1)', 'Pareggio (% X)', 'Trasferta (% 2)',
        'Media Gol Partita', '% Over 2.5', '% Gol/Gol'
    ]);

    let leagCount = 0;
    for (const [_, l] of leagueData.entries()) {
        if (l.matches >= 20) {
            const pct1 = ((l.homeWins / l.matches) * 100).toFixed(1) + '%';
            const pctX = ((l.draws / l.matches) * 100).toFixed(1) + '%';
            const pct2 = ((l.awayWins / l.matches) * 100).toFixed(1) + '%';
            const avgG = (l.totalGoals / l.matches).toFixed(2);
            const pctOver = ((l.over25 / l.matches) * 100).toFixed(1) + '%';
            const pctGG = ((l.gg / l.matches) * 100).toFixed(1) + '%';

            sheetStatCamp.addRow([
                l.id, l.nome, l.paese, l.matches,
                pct1, pctX, pct2, avgG, pctOver, pctGG
            ]);
            leagCount++;
        }
    }

    sheetStatCamp.views = [{ state: 'frozen', ySplit: 1 }];
    sheetStatCamp.columns.forEach(col => {
        let maxLen = 14;
        col.eachCell({ includeEmpty: false }, cell => {
            const len = cell.value ? cell.value.toString().length : 0;
            if (len > maxLen) maxLen = Math.min(len + 3, 28);
        });
        col.width = maxLen;
    });

    // 2. Popolamento Foglio "Classifiche"
    let sheetClass = workbook.getWorksheet('Classifiche');
    if (sheetClass) workbook.removeWorksheet(sheetClass.id);
    sheetClass = workbook.addWorksheet('Classifiche');

    sheetClass.addRow([
        'Campionato', 'Paese', 'Posizione', 'Squadra', 'Punti',
        'Giocate', 'Vinte', 'Nulle', 'Perse', 'Gol Fatti', 'Gol Subiti', 'Diff Reti'
    ]);

    let classRows = 0;
    for (const [_, l] of leagueData.entries()) {
        if (l.teams.size > 0) {
            // Ordina squadre per punti, poi diff reti, poi gol fatti
            const sortedTeams = Array.from(l.teams.values()).sort((a, b) => {
                if (b.pts !== a.pts) return b.pts - a.pts;
                const diffB = b.gf - b.ga;
                const diffA = a.gf - a.ga;
                if (diffB !== diffA) return diffB - diffA;
                return b.gf - a.gf;
            });

            sortedTeams.forEach((t, idx) => {
                sheetClass.addRow([
                    l.nome, l.paese, idx + 1, t.name, t.pts,
                    t.p, t.w, t.d, t.l, t.gf, t.ga, (t.gf - t.ga)
                ]);
                classRows++;
            });
        }
    }

    sheetClass.views = [{ state: 'frozen', ySplit: 1 }];
    sheetClass.columns.forEach(col => {
        let maxLen = 12;
        col.eachCell({ includeEmpty: false }, cell => {
            const len = cell.value ? cell.value.toString().length : 0;
            if (len > maxLen) maxLen = Math.min(len + 3, 26);
        });
        col.width = maxLen;
    });

    await safeWriteWorkbook(workbook, EXCEL_PATH);
    console.log(`\n======================================================`);
    console.log(`✅ Foglio StatisticheCampionati popolato: ${leagCount} leghe analizzate`);
    console.log(`✅ Foglio Classifiche popolato: ${classRows} posizioni generate`);
    console.log(`======================================================\n`);
}

generaClassificheEStatistiche().catch(console.error);