import axios from 'axios';
import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const EXCEL_PATH = path.join(__dirname, '../data/metis_database.xlsx');

const MAX_PAGES = 4; // Scansiona fino a 400 eventi per coprire mattina, pomeriggio e sera

const http = axios.create({
    headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'application/json'
    },
    timeout: 10000
});

function getTargetDate() {
    const argDate = process.argv[2];
    if (argDate && /^\d{4}-\d{2}-\d{2}$/.test(argDate)) return argDate;
    return new Date().toISOString().split('T')[0];
}

function cleanString(str) {
    if (!str) return '';
    return str.toLowerCase().replace(/[^a-z0-9]/g, '').trim();
}

async function safeWriteWorkbook(workbook, filePath, retries = 3) {
    for (let i = 0; i < retries; i++) {
        try {
            await workbook.xlsx.writeFile(filePath);
            console.log(`💾 File Excel salvato: ${filePath}`);
            return;
        } catch (err) {
            if (err.code === 'EBUSY') {
                console.warn(`⚠️ File Excel occupato. Riprovo tra 2 secondi (${i + 1}/${retries})...`);
                await new Promise(r => setTimeout(r, 2000));
            } else {
                throw err;
            }
        }
    }
}

async function scaricaPartiteGiorno() {
    const targetDate = getTargetDate();
    console.log(`🚀 [METIS ENGINE] Download programmato partite per: ${targetDate}\n`);

    if (!fs.existsSync(EXCEL_PATH)) {
        console.error(`❌ Errore: File non trovato in ${EXCEL_PATH}`);
        return;
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(EXCEL_PATH);

    const sheetCampionati = workbook.getWorksheet('Campionati');
    if (!sheetCampionati) {
        console.error('❌ Errore: Foglio "Campionati" non trovato.');
        return;
    }

    const monitoredCompetitions = [];
    sheetCampionati.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;
        const metisId = (row.getCell(1).value || '').toString().trim();
        const paese = (row.getCell(2).value || '').toString().trim();
        const nome = (row.getCell(3).value || '').toString().trim();
        const slug = (row.getCell(4).value || '').toString().trim();

        if (nome) {
            monitoredCompetitions.push({
                metisId,
                paese,
                nome,
                slug,
                cleanNome: cleanString(nome),
                cleanSlug: cleanString(slug),
                cleanPaese: cleanString(paese)
            });
        }
    });

    let sheetFixtures = workbook.getWorksheet('Fixtures');
    if (sheetFixtures) {
        workbook.removeWorksheet(sheetFixtures.id);
    }
    sheetFixtures = workbook.addWorksheet('Fixtures');
    sheetFixtures.addRow([
        'Match ID', 'Data Partita', 'Orario UTC', 'Metis League ID', 'Campionato', 'Paese',
        'Squadra Casa', 'Squadra Trasferta', 'Stato', 'Gol Casa', 'Gol Trasferta', 'Slug Match', 'Data Rilevazione'
    ]);

    const oraRilevazione = new Date().toISOString();
    let countImported = 0;
    const addedSlugs = new Set();

    for (let page = 1; page <= MAX_PAGES; page++) {
        try {
            console.log(`📡 Richiesta pagina ${page}/${MAX_PAGES}...`);
            const url = `https://sportscore.com/api/v1/fixtures/?sport=football&date=${targetDate}&page=${page}`;
            const response = await http.get(url);
            const rawMatches = response.data?.events || response.data?.fixtures || response.data?.matches || [];

            if (!Array.isArray(rawMatches) || rawMatches.length === 0) break;

            rawMatches.forEach(item => {
                const rawComp = (item.competition || '').trim();
                const cleanComp = cleanString(rawComp);
                const homeTeam = (item.home || '').trim();
                const awayTeam = (item.away || '').trim();

                if (!cleanComp || !homeTeam || !awayTeam) return;

                // Esclusione categorie minori
                const fullStr = `${homeTeam} ${awayTeam} ${rawComp}`.toLowerCase();
                const blacklist = ['(w)', 'women', 'u18', 'u19', 'u20', 'u21', 'u23', 'youth', 'reserve', 'reserves', 'academy', 'amateur', 'police', 'young mens'];
                if (blacklist.some(term => fullStr.includes(term))) return;

                // Match rigoroso sul torneo
                const matchedLeague = monitoredCompetitions.find(c => {
                    const isGeneric = ['premierleague', 'championship', 'superleague', 'superliga', 'seriea', 'serieb', 'leagueone', 'leaguetwo'].includes(c.cleanNome);
                    if (isGeneric) {
                        return (cleanComp.includes(c.cleanPaese) && cleanComp.includes(c.cleanNome)) || cleanComp === c.cleanSlug;
                    }
                    if (cleanComp === c.cleanNome || cleanComp === c.cleanSlug) return true;
                    if (c.cleanNome.length >= 8 && cleanComp.includes(c.cleanNome)) return true;
                    return false;
                });

                if (matchedLeague) {
                    let matchSlug = '';
                    if (item.url) {
                        const parts = item.url.split('/').filter(Boolean);
                        matchSlug = parts[2] || parts[parts.length - 1] || '';
                    }
                    const matchId = matchSlug ? `M_${matchSlug}` : `M_${cleanString(homeTeam)}_${cleanString(awayTeam)}`;
                    
                    if (addedSlugs.has(matchId)) return;
                    addedSlugs.add(matchId);

                    const status = item.status_text || item.status || 'SCHEDULED';
                    const homeScore = item.home_score !== undefined && item.home_score !== null ? item.home_score : '';
                    const awayScore = item.away_score !== undefined && item.away_score !== null ? item.away_score : '';

                    let matchTime = '';
                    if (item.time && item.time.includes('T')) {
                        matchTime = item.time.split('T')[1].slice(0, 5);
                    }

                    sheetFixtures.addRow([
                        matchId, targetDate, matchTime, matchedLeague.metisId, matchedLeague.nome,
                        matchedLeague.paese, homeTeam, awayTeam, status, homeScore, awayScore,
                        matchSlug, oraRilevazione
                    ]);

                    countImported++;
                    console.log(`   ⚽ [${matchedLeague.nome} - ${matchedLeague.paese}] ${homeTeam} vs ${awayTeam}`);
                }
            });

            if (rawMatches.length < 50) break;
            await new Promise(r => setTimeout(r, 150));
        } catch (err) {
            console.warn(`Fine scansione a pagina ${page}:`, err.message);
            break;
        }
    }

    console.log(`\n======================================================`);
    console.log(`✅ Partite certificate importate: ${countImported}`);

    sheetFixtures.views = [{ state: 'frozen', ySplit: 1 }];
    sheetFixtures.columns.forEach(col => {
        let maxLen = 14;
        col.eachCell({ includeEmpty: false }, cell => {
            const len = cell.value ? cell.value.toString().length : 0;
            if (len > maxLen) maxLen = Math.min(len + 3, 35);
        });
        col.width = maxLen;
    });

    await safeWriteWorkbook(workbook, EXCEL_PATH);
    console.log(`======================================================\n`);
}

scaricaPartiteGiorno();