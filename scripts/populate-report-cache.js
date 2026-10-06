import fs from 'fs';
import path from 'path';
import ExcelJS from 'exceljs';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const EXCEL_PATH = path.join(__dirname, '../data/metis_database.xlsx');

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

async function generaReportECache() {
    console.log('🚀 [METIS ENGINE] Generazione fogli Report e MetisCache...\n');

    if (!fs.existsSync(EXCEL_PATH)) {
        console.error(`❌ File non trovato: ${EXCEL_PATH}`);
        return;
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(EXCEL_PATH);

    const sheetFixtures = workbook.getWorksheet('Fixtures');
    const sheetSquadre = workbook.getWorksheet('Squadre');
    const sheetH2H = workbook.getWorksheet('H2H');
    const sheetCamp = workbook.getWorksheet('Campionati');

    // 1. Reset e configurazione foglio Report
    let sheetReport = workbook.getWorksheet('Report');
    if (sheetReport) workbook.removeWorksheet(sheetReport.id);
    sheetReport = workbook.addWorksheet('Report');

    sheetReport.addRow(['PANNELLO DI CONTROLLO & VALUE BETS - METIS ENGINE']);
    sheetReport.addRow(['Data Generazione:', new Date().toLocaleString('it-IT')]);
    sheetReport.addRow([]);

    sheetReport.addRow([
        'Match ID', 'Data', 'Campionato', 'Incontro',
        'xG Casa', 'xG Trasferta', 'Quota Fair 1', 'Quota Fair X', 'Quota Fair 2',
        'Quota Fair Over 2.5', 'Quota Fair Under 2.5', 'Precedenti H2H', 'Esito Atteso'
    ]);

    let fixturesCount = 0;
    if (sheetFixtures) {
        sheetFixtures.eachRow((row, rowNumber) => {
            if (rowNumber === 1) return;
            const matchId = row.getCell(1).value;
            const date = row.getCell(2).value;
            const camp = row.getCell(5).value;
            const home = row.getCell(7).value;
            const away = row.getCell(8).value;
            const h2hTot = row.getCell(12).value || 0;

            const xgH = row.getCell(21).value;
            const xgA = row.getCell(22).value;
            const q1 = row.getCell(23).value;
            const qX = row.getCell(24).value;
            const q2 = row.getCell(25).value;
            const qOver = row.getCell(26).value;
            const qUnder = row.getCell(27).value;

            if (home && away) {
                fixturesCount++;
                let esito = 'Equilibrata';
                if (q1 && q2) {
                    if (Number(q1) < 1.65) esito = '1 Forte';
                    else if (Number(q2) < 1.65) esito = '2 Forte';
                    else if (Number(qOver) < 1.60) esito = 'Over 2.5 Probabile';
                    else if (Number(qUnder) < 1.60) esito = 'Under 2.5 Probabile';
                }

                sheetReport.addRow([
                    matchId, date, camp, `${home} - ${away}`,
                    xgH, xgA, q1, qX, q2, qOver, qUnder,
                    `${h2hTot} gare`, esito
                ]);
            }
        });
    }

    // 2. Reset e configurazione foglio MetisCache
    let sheetCache = workbook.getWorksheet('MetisCache');
    if (sheetCache) workbook.removeWorksheet(sheetCache.id);
    sheetCache = workbook.addWorksheet('MetisCache');

    sheetCache.addRow(['Chiave Parametro', 'Valore', 'Ultimo Controllo']);
    sheetCache.addRow(['Versione Motore', 'Metis Engine v2.4 (Poisson + H2H)', new Date().toISOString()]);
    sheetCache.addRow(['Campionati Monitorati', sheetCamp ? sheetCamp.rowCount - 1 : 0, new Date().toISOString()]);
    sheetCache.addRow(['Club Censiti nel Database', sheetSquadre ? sheetSquadre.rowCount - 1 : 0, new Date().toISOString()]);
    sheetCache.addRow(['Matrici H2H Storiche', sheetH2H ? sheetH2H.rowCount - 1 : 0, new Date().toISOString()]);
    sheetCache.addRow(['Partite in Palinsesto Attivo', fixturesCount, new Date().toISOString()]);
    sheetCache.addRow(['Dataset Storico Totale', '121.892 partite indicizzate', new Date().toISOString()]);

    // Formattazione larghezza colonne
    [sheetReport, sheetCache].forEach(sh => {
        sh.columns.forEach(col => {
            let maxLen = 14;
            col.eachCell({ includeEmpty: false }, cell => {
                const len = cell.value ? cell.value.toString().length : 0;
                if (len > maxLen) maxLen = Math.min(len + 3, 30);
            });
            col.width = maxLen;
        });
    });

    await safeWriteWorkbook(workbook, EXCEL_PATH);
    console.log(`\n======================================================`);
    console.log(`✅ Fogli Report e MetisCache compilati con successo!`);
    console.log(`======================================================\n`);
}

generaReportECache().catch(console.error);