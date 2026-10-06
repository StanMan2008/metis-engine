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
            } else {
                throw err;
            }
        }
    }
}

async function calcolaStatisticheSquadre() {
    console.log('🚀 [METIS ENGINE] Calcolo statistiche aggregate per squadra...\n');

    if (!fs.existsSync(EXCEL_PATH)) {
        console.error(`❌ File non trovato: ${EXCEL_PATH}`);
        return;
    }

    const teamStats = new Map();
    const dirs = [MAIN_DIR, EXTRA_DIR];

    dirs.forEach(dir => {
        if (!fs.existsSync(dir)) return;
        const files = fs.readdirSync(dir);
        files.forEach(f => {
            if (!f.endsWith('.csv')) return;
            try {
                const content = fs.readFileSync(path.join(dir, f), 'utf-8');
                const rows = parseCSV(content);

                rows.forEach(m => {
                    const home = (m.HomeTeam || m.Home || m.home || '').trim();
                    const away = (m.AwayTeam || m.Away || m.away || '').trim();
                    const fthg = Number(m.FTHG ?? m.HG ?? m.home_score);
                    const ftag = Number(m.FTAG ?? m.AG ?? m.away_score);

                    if (home && away && !isNaN(fthg) && !isNaN(ftag)) {
                        const initStat = (name) => ({
                            name,
                            playedH: 0, winH: 0, drawH: 0, lossH: 0, scoredH: 0, concededH: 0,
                            playedA: 0, winA: 0, drawA: 0, lossA: 0, scoredA: 0, concededA: 0,
                            over25: 0, gg: 0, cleanSheet: 0
                        });

                        const kHome = clean(home);
                        const kAway = clean(away);

                        if (!teamStats.has(kHome)) teamStats.set(kHome, initStat(home));
                        if (!teamStats.has(kAway)) teamStats.set(kAway, initStat(away));

                        const sH = teamStats.get(kHome);
                        const sA = teamStats.get(kAway);

                        // Dati Home
                        sH.playedH++;
                        sH.scoredH += fthg;
                        sH.concededH += ftag;
                        if (fthg > ftag) sH.winH++;
                        else if (fthg === ftag) sH.drawH++;
                        else sH.lossH++;
                        if (ftag === 0) sH.cleanSheet++;
                        if (fthg + ftag > 2.5) sH.over25++;
                        if (fthg > 0 && ftag > 0) sH.gg++;

                        // Dati Away
                        sA.playedA++;
                        sA.scoredA += ftag;
                        sA.concededA += fthg;
                        if (ftag > fthg) sA.winA++;
                        else if (ftag === fthg) sA.drawA++;
                        else sA.lossA++;
                        if (fthg === 0) sA.cleanSheet++;
                        if (fthg + ftag > 2.5) sA.over25++;
                        if (fthg > 0 && ftag > 0) sA.gg++;
                    }
                });
            } catch {
                // Skip file illeggibili
            }
        });
    });

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(EXCEL_PATH);

    let sheetStat = workbook.getWorksheet('StatisticheSquadra');
    if (sheetStat) {
        workbook.removeWorksheet(sheetStat.id);
    }
    sheetStat = workbook.addWorksheet('StatisticheSquadra');

    sheetStat.addRow([
        'Squadra', 'Gare Totali',
        'Gare Casa', 'Vittorie Casa', 'Pareggi Casa', 'Sconfitte Casa', 'Media Gol Fatti Casa', 'Media Gol Subiti Casa',
        'Gare Trasferta', 'Vittorie Trasf', 'Pareggi Trasf', 'Sconfitte Trasf', 'Media Gol Fatti Trasf', 'Media Gol Subiti Trasf',
        '% Over 2.5', '% Gol/Gol', '% Clean Sheet'
    ]);

    let rowCount = 0;
    for (const [_, s] of teamStats.entries()) {
        const totPlayed = s.playedH + s.playedA;
        if (totPlayed >= 10) { // Solo squadre con almeno 10 gare registrate
            const avgScoredH = s.playedH > 0 ? (s.scoredH / s.playedH).toFixed(2) : '0.00';
            const avgConcededH = s.playedH > 0 ? (s.concededH / s.playedH).toFixed(2) : '0.00';

            const avgScoredA = s.playedA > 0 ? (s.scoredA / s.playedA).toFixed(2) : '0.00';
            const avgConcededA = s.playedA > 0 ? (s.concededA / s.playedA).toFixed(2) : '0.00';

            const pctOver = ((s.over25 / totPlayed) * 100).toFixed(1) + '%';
            const pctGG = ((s.gg / totPlayed) * 100).toFixed(1) + '%';
            const pctCS = ((s.cleanSheet / totPlayed) * 100).toFixed(1) + '%';

            sheetStat.addRow([
                s.name, totPlayed,
                s.playedH, s.winH, s.drawH, s.lossH, avgScoredH, avgConcededH,
                s.playedA, s.winA, s.drawA, s.lossA, avgScoredA, avgConcededA,
                pctOver, pctGG, pctCS
            ]);
            rowCount++;
        }
    }

    sheetStat.views = [{ state: 'frozen', ySplit: 1 }];
    sheetStat.columns.forEach(col => {
        let maxLen = 12;
        col.eachCell({ includeEmpty: false }, cell => {
            const len = cell.value ? cell.value.toString().length : 0;
            if (len > maxLen) maxLen = Math.min(len + 3, 26);
        });
        col.width = maxLen;
    });

    await safeWriteWorkbook(workbook, EXCEL_PATH);
    console.log(`\n✅ Foglio StatisticheSquadra popolato con successo per ${rowCount} club!\n`);
}

calcolaStatisticheSquadre().catch(console.error);