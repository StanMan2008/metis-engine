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

async function popolaPartiteEDettagli() {
    console.log('🚀 [METIS ENGINE] Popolamento fogli Partite e PartiteDettagli...\n');

    if (!fs.existsSync(EXCEL_PATH)) {
        console.error(`❌ File non trovato: ${EXCEL_PATH}`);
        return;
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(EXCEL_PATH);

    // Reset foglio Partite
    let sheetPartite = workbook.getWorksheet('Partite');
    if (sheetPartite) workbook.removeWorksheet(sheetPartite.id);
    sheetPartite = workbook.addWorksheet('Partite');
    sheetPartite.addRow([
        'Match ID', 'Data', 'Campionato', 'Squadra Casa', 'Squadra Trasferta',
        'Gol Casa FT', 'Gol Trasferta FT', 'Esito FT', 'Gol Casa HT', 'Gol Trasferta HT', 'Esito HT'
    ]);

    // Reset foglio PartiteDettagli
    let sheetDettagli = workbook.getWorksheet('PartiteDettagli');
    if (sheetDettagli) workbook.removeWorksheet(sheetDettagli.id);
    sheetDettagli = workbook.addWorksheet('PartiteDettagli');
    sheetDettagli.addRow([
        'Match ID', 'Tiri Casa', 'Tiri Trasferta', 'Tiri Porta Casa', 'Tiri Porta Trasferta',
        'Falli Casa', 'Falli Trasferta', 'Corner Casa', 'Corner Trasferta',
        'Gialli Casa', 'Gialli Trasferta', 'Rossi Casa', 'Rossi Trasferta'
    ]);

    let matchCount = 0;
    const dirs = [MAIN_DIR, EXTRA_DIR];

    // Seleziona i file delle stagioni recenti per mantenere leggero il foglio
    dirs.forEach(dir => {
        if (!fs.existsSync(dir)) return;
        const files = fs.readdirSync(dir);
        files.forEach(f => {
            if (!f.endsWith('.csv')) return;
            const isRecent = f.includes('2024') || f.includes('2025') || f.includes('2026') || f.includes('2425') || f.includes('2526');
            if (!isRecent) return;

            const campName = f.replace('.csv', '').split('_').slice(0, 2).join(' ');

            try {
                const content = fs.readFileSync(path.join(dir, f), 'utf-8');
                const rows = parseCSV(content);

                rows.forEach((m, idx) => {
                    const home = (m.HomeTeam || m.Home || m.home || '').trim();
                    const away = (m.AwayTeam || m.Away || m.away || '').trim();
                    const date = m.Date || m.date || '';
                    const fthg = m.FTHG ?? m.HG ?? m.home_score;
                    const ftag = m.FTAG ?? m.AG ?? m.away_score;
                    const ftr = m.FTR ?? m.Res ?? (fthg > ftag ? 'H' : (fthg === ftag ? 'D' : 'A'));

                    if (home && away && fthg !== '' && ftag !== '') {
                        matchCount++;
                        const matchId = `HIST_${matchCount}`;

                        const hthg = m.HTHG ?? '';
                        const htag = m.HTAG ?? '';
                        const htr = m.HTR ?? '';

                        sheetPartite.addRow([
                            matchId, date, campName, home, away,
                            Number(fthg), Number(ftag), ftr,
                            hthg !== '' ? Number(hthg) : '',
                            htag !== '' ? Number(htag) : '',
                            htr
                        ]);

                        // Dettagli tecnici (se presenti nel CSV)
                        const hs = m.HS ?? '';
                        const as = m.AS ?? '';
                        const hst = m.HST ?? '';
                        const ast = m.AST ?? '';
                        const hf = m.HF ?? '';
                        const af = m.AF ?? '';
                        const hc = m.HC ?? '';
                        const ac = m.AC ?? '';
                        const hy = m.HY ?? '';
                        const ay = m.AY ?? '';
                        const hr = m.HR ?? '';
                        const ar = m.AR ?? '';

                        if (hs !== '' || hc !== '' || hy !== '') {
                            sheetDettagli.addRow([
                                matchId,
                                hs !== '' ? Number(hs) : '',
                                as !== '' ? Number(as) : '',
                                hst !== '' ? Number(hst) : '',
                                ast !== '' ? Number(ast) : '',
                                hf !== '' ? Number(hf) : '',
                                af !== '' ? Number(af) : '',
                                hc !== '' ? Number(hc) : '',
                                ac !== '' ? Number(ac) : '',
                                hy !== '' ? Number(hy) : '',
                                ay !== '' ? Number(ay) : '',
                                hr !== '' ? Number(hr) : '',
                                ar !== '' ? Number(ar) : ''
                            ]);
                        }
                    }
                });
            } catch {}
        });
    });

    sheetPartite.views = [{ state: 'frozen', ySplit: 1 }];
    sheetDettagli.views = [{ state: 'frozen', ySplit: 1 }];

    await safeWriteWorkbook(workbook, EXCEL_PATH);
    console.log(`\n======================================================`);
    console.log(`✅ Fogli Partite e PartiteDettagli popolati con ${matchCount} gare recenti!`);
    console.log(`======================================================\n`);
}

popolaPartiteEDettagli().catch(console.error);