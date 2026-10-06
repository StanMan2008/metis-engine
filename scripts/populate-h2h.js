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

async function generaMatriceH2H() {
    console.log('🚀 [METIS ENGINE] Generazione archivio H2H da storico decennale...\n');

    if (!fs.existsSync(EXCEL_PATH)) {
        console.error(`❌ File non trovato: ${EXCEL_PATH}`);
        return;
    }

    const pairMap = new Map();
    const dirs = [MAIN_DIR, EXTRA_DIR];
    let totalLoaded = 0;

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
                    const date = m.Date || m.date || '';

                    if (home && away && !isNaN(fthg) && !isNaN(ftag)) {
                        totalLoaded++;
                        // Chiave alfabetica univoca per la coppia di club
                        const isOrder = home.localeCompare(away) <= 0;
                        const teamA = isOrder ? home : away;
                        const teamB = isOrder ? away : home;
                        const scoreA = isOrder ? fthg : ftag;
                        const scoreB = isOrder ? ftag : fthg;

                        const key = `${clean(teamA)}__${clean(teamB)}`;

                        if (!pairMap.has(key)) {
                            pairMap.set(key, {
                                nameA: teamA,
                                nameB: teamB,
                                played: 0,
                                winsA: 0,
                                draws: 0,
                                winsB: 0,
                                goalsA: 0,
                                goalsB: 0,
                                lastScore: '',
                                lastDate: ''
                            });
                        }

                        const stat = pairMap.get(key);
                        stat.played++;
                        stat.goalsA += scoreA;
                        stat.goalsB += scoreB;

                        if (scoreA > scoreB) stat.winsA++;
                        else if (scoreA === scoreB) stat.draws++;
                        else stat.winsB++;

                        stat.lastScore = `${scoreA}-${scoreB}`;
                        stat.lastDate = date;
                    }
                });
            } catch {
                // Skip file illeggibili
            }
        });
    });

    console.log(`⚡ Partite elaborate: ${totalLoaded}`);
    console.log(`📊 Coppie di club identificate: ${pairMap.size}`);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(EXCEL_PATH);

    let sheetH2H = workbook.getWorksheet('H2H');
    if (sheetH2H) {
        workbook.removeWorksheet(sheetH2H.id);
    }
    sheetH2H = workbook.addWorksheet('H2H');

    sheetH2H.addRow([
        'Coppia ID', 'Squadra A', 'Squadra B', 'Match Totali',
        'Vittorie A', 'Pareggi', 'Vittorie B',
        '% Vittoria A', '% Pareggio', '% Vittoria B',
        'Media Gol A', 'Media Gol B', 'Media Gol Totali', 'Ultimo Incontro'
    ]);

    // Filtra per coppie con almeno 3 precedenti storici per evitare dispersioni
    let exportedRows = 0;
    for (const [key, s] of pairMap.entries()) {
        if (s.played >= 3) {
            const pctA = ((s.winsA / s.played) * 100).toFixed(1) + '%';
            const pctX = ((s.draws / s.played) * 100).toFixed(1) + '%';
            const pctB = ((s.winsB / s.played) * 100).toFixed(1) + '%';

            const avgA = (s.goalsA / s.played).toFixed(2);
            const avgB = (s.goalsB / s.played).toFixed(2);
            const avgTot = ((s.goalsA + s.goalsB) / s.played).toFixed(2);

            sheetH2H.addRow([
                key, s.nameA, s.nameB, s.played,
                s.winsA, s.draws, s.winsB,
                pctA, pctX, pctB,
                avgA, avgB, avgTot,
                `${s.lastScore} (${s.lastDate})`
            ]);
            exportedRows++;
        }
    }

    sheetH2H.views = [{ state: 'frozen', ySplit: 1 }];
    sheetH2H.columns.forEach(col => {
        let maxLen = 12;
        col.eachCell({ includeEmpty: false }, cell => {
            const len = cell.value ? cell.value.toString().length : 0;
            if (len > maxLen) maxLen = Math.min(len + 3, 28);
        });
        col.width = maxLen;
    });

    await safeWriteWorkbook(workbook, EXCEL_PATH);
    console.log(`\n✅ Foglio H2H popolato con ${exportedRows} matrici di scontro diretto!\n`);
}

generaMatriceH2H().catch(console.error);