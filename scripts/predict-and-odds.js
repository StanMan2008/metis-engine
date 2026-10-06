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

// Parole generiche che non identificano in modo univoco una squadra
const STOP_WORDS = new Set([
    'fc', 'cf', 'ac', 'sc', 'cd', 'as', 'ssc', 'afc', 'bsc', 'fk', 'sk', 'bk', 'ifk', 'ff',
    'club', 'de', 'do', 'da', 'del', 'la', 'le', 'the', 'sport', 'calcio', 'football',
    'women', 'w', 'youth', 'u19', 'u20', 'u21', 'u23', 'ii', 'b', 'rj', 'sp',
    'city', 'united', 'town', 'rovers', 'wanderers', 'athletic', 'atletico', 'real',
    'inter', 'sporting', 'racing', 'deportivo', 'union', 'olympic', 'olympique', 'al', 'el'
]);

function tokenize(name) {
    if (!name) return [];
    return name
        .toLowerCase()
        .replace(/[^a-z0-9\s]/g, ' ')
        .split(/\s+/)
        .filter(w => w.length > 2 && !STOP_WORDS.has(w));
}

function isMatch(nameA, nameB) {
    if (!nameA || !nameB) return false;
    const cleanA = nameA.toLowerCase().replace(/[^a-z0-9]/g, '');
    const cleanB = nameB.toLowerCase().replace(/[^a-z0-9]/g, '');

    if (cleanA === cleanB) return true;
    if (cleanA.includes(cleanB) || cleanB.includes(cleanA)) {
        // Se uno include l'altro e la parte corta ha almeno 4 lettere (es: "puebla", "toluca", "tokyo", "urawa")
        const shorter = cleanA.length < cleanB.length ? cleanA : cleanB;
        if (shorter.length >= 4) return true;
    }

    const tokensA = tokenize(nameA);
    const tokensB = tokenize(nameB);

    if (tokensA.length === 0 || tokensB.length === 0) return false;

    // Se hanno in comune almeno un token caratterizzante di almeno 4 lettere (es: "leon", "tigres", "tokyo")
    const common = tokensA.filter(t => tokensB.includes(t));
    if (common.some(t => t.length >= 4)) return true;

    return false;
}

function poisson(k, lambda) {
    if (lambda <= 0) return k === 0 ? 1 : 0;
    let fact = 1;
    for (let i = 2; i <= k; i++) fact *= i;
    return (Math.pow(lambda, k) * Math.exp(-lambda)) / fact;
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

function loadHistoricalData() {
    console.log('📂 Lettura e indicizzazione dataset decennale in memoria...');
    const allMatches = [];
    const dirs = [MAIN_DIR, EXTRA_DIR];

    dirs.forEach(dir => {
        if (!fs.existsSync(dir)) return;
        const files = fs.readdirSync(dir);
        files.forEach(f => {
            if (!f.endsWith('.csv')) return;
            try {
                const content = fs.readFileSync(path.join(dir, f), 'utf-8');
                const parsed = parseCSV(content);
                parsed.forEach(m => {
                    const h = (m.HomeTeam || m.Home || m.home || '').trim();
                    const a = (m.AwayTeam || m.Away || m.away || '').trim();
                    const fthg = Number(m.FTHG ?? m.HG ?? m.home_score);
                    const ftag = Number(m.FTAG ?? m.AG ?? m.away_score);
                    const date = m.Date || m.date || '';

                    if (h && a && !isNaN(fthg) && !isNaN(ftag)) {
                        allMatches.push({
                            home: h,
                            away: a,
                            homeScore: fthg,
                            awayScore: ftag,
                            date
                        });
                    }
                });
            } catch {
                // Salta eventuali file o righe con encoding non valido
            }
        });
    });

    console.log(`⚡ ${allMatches.length} partite storiche caricate con successo!\n`);
    return allMatches;
}

function getH2H(historicalMatches, homeTeam, awayTeam) {
    let homeWins = 0;
    let draws = 0;
    let awayWins = 0;
    let totalGoals = 0;
    let lastMatchStr = 'N/D';

    const directClashes = historicalMatches.filter(m => {
        return (isMatch(m.home, homeTeam) && isMatch(m.away, awayTeam)) ||
               (isMatch(m.home, awayTeam) && isMatch(m.away, homeTeam));
    });

    directClashes.forEach((m, idx) => {
        totalGoals += (m.homeScore + m.awayScore);
        const isCurrentHomePlayingHome = isMatch(m.home, homeTeam);

        if (m.homeScore > m.awayScore) {
            if (isCurrentHomePlayingHome) homeWins++;
            else awayWins++;
        } else if (m.homeScore === m.awayScore) {
            draws++;
        } else {
            if (isCurrentHomePlayingHome) awayWins++;
            else homeWins++;
        }

        if (idx === directClashes.length - 1) {
            lastMatchStr = `${m.homeScore}-${m.awayScore} (${m.date || 'rec'})`;
        }
    });

    const total = directClashes.length;
    const avgGoals = total > 0 ? (totalGoals / total).toFixed(2) : '-';

    return {
        total,
        homeWins,
        draws,
        awayWins,
        avgGoals,
        lastMatch: lastMatchStr,
        pct1: total > 0 ? ((homeWins / total) * 100).toFixed(0) + '%' : '-',
        pctX: total > 0 ? ((draws / total) * 100).toFixed(0) + '%' : '-',
        pct2: total > 0 ? ((awayWins / total) * 100).toFixed(0) + '%' : '-'
    };
}

function computeOdds(historicalMatches, homeTeam, awayTeam) {
    let hScored = 0, hCount = 0;
    let aScored = 0, aCount = 0;

    historicalMatches.forEach(m => {
        if (isMatch(m.home, homeTeam)) {
            hScored += m.homeScore;
            hCount++;
        }
        if (isMatch(m.away, awayTeam)) {
            aScored += m.awayScore;
            aCount++;
        }
    });

    let lambda = hCount >= 3 ? (hScored / hCount) : 1.35;
    let mu = aCount >= 3 ? (aScored / aCount) : 1.10;

    lambda = Math.max(0.4, Math.min(3.5, lambda));
    mu = Math.max(0.3, Math.min(3.0, mu));

    let p1 = 0, pX = 0, p2 = 0;
    let over25 = 0;

    for (let x = 0; x <= 6; x++) {
        for (let y = 0; y <= 6; y++) {
            const prob = poisson(x, lambda) * poisson(y, mu);
            if (x > y) p1 += prob;
            else if (x === y) pX += prob;
            else p2 += prob;

            if (x + y > 2.5) over25 += prob;
        }
    }

    const tot = p1 + pX + p2;
    p1 /= tot; pX /= tot; p2 /= tot;

    return {
        lambda: lambda.toFixed(2),
        mu: mu.toFixed(2),
        fair1: (1 / p1).toFixed(2),
        fairX: (1 / pX).toFixed(2),
        fair2: (1 / p2).toFixed(2),
        fairOver25: (1 / over25).toFixed(2),
        fairUnder25: (1 / (1 - over25)).toFixed(2)
    };
}

async function safeWriteWorkbook(workbook, filePath, retries = 3) {
    for (let i = 0; i < retries; i++) {
        try {
            await workbook.xlsx.writeFile(filePath);
            console.log(`💾 File Excel salvato: ${filePath}`);
            return;
        } catch (err) {
            if (err.code === 'EBUSY') {
                console.warn(`⚠️ File Excel occupato (chiudi Excel se aperto!). Tentativo ${i + 1}/${retries}...`);
                await new Promise(r => setTimeout(r, 2000));
            } else {
                throw err;
            }
        }
    }
    console.error(`❌ Impossibile salvare su ${filePath}: file ancora bloccato da Excel o OneDrive.`);
}

async function elaboraConStorico() {
    console.log('🚀 [METIS ENGINE] Calcolo H2H Rigoroso e Quote su Fixtures...\n');

    if (!fs.existsSync(EXCEL_PATH)) {
        console.error(`❌ Errore: File non trovato in ${EXCEL_PATH}`);
        return;
    }

    const matchesHistory = loadHistoricalData();

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(EXCEL_PATH);

    const sheetFixtures = workbook.getWorksheet('Fixtures');
    if (!sheetFixtures) {
        console.error('❌ Errore: Foglio "Fixtures" non trovato.');
        return;
    }

    const headers = [
        'Match ID', 'Data Partita', 'Orario UTC', 'Metis League ID', 'Campionato', 'Paese',
        'Squadra Casa', 'Squadra Trasferta', 'Stato', 'Gol Casa', 'Gol Trasferta',
        'H2H Match Tot', 'H2H Vinte Casa', 'H2H Pareggi', 'H2H Vinte Ospite',
        'H2H % 1', 'H2H % X', 'H2H % 2', 'H2H Media Gol', 'Ultimo Scontro',
        'xG Casa', 'xG Ospite', 'Quota 1', 'Quota X', 'Quota 2', 'Quota Over 2.5', 'Quota Under 2.5'
    ];

    sheetFixtures.getRow(1).values = headers;

    let processed = 0;

    sheetFixtures.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;

        const home = row.getCell(7).value ? String(row.getCell(7).value).trim() : '';
        const away = row.getCell(8).value ? String(row.getCell(8).value).trim() : '';

        if (home && away) {
            processed++;
            const h2h = getH2H(matchesHistory, home, away);
            const odds = computeOdds(matchesHistory, home, away);

            row.getCell(12).value = h2h.total;
            row.getCell(13).value = h2h.homeWins;
            row.getCell(14).value = h2h.draws;
            row.getCell(15).value = h2h.awayWins;
            row.getCell(16).value = h2h.pct1;
            row.getCell(17).value = h2h.pctX;
            row.getCell(18).value = h2h.pct2;
            row.getCell(19).value = h2h.avgGoals;
            row.getCell(20).value = h2h.lastMatch;

            row.getCell(21).value = Number(odds.lambda);
            row.getCell(22).value = Number(odds.mu);
            row.getCell(23).value = Number(odds.fair1);
            row.getCell(24).value = Number(odds.fairX);
            row.getCell(25).value = Number(odds.fair2);
            row.getCell(26).value = Number(odds.fairOver25);
            row.getCell(27).value = Number(odds.fairUnder25);

            if (h2h.total > 0) {
                console.log(`🔥 H2H: ${home} vs ${away} -> ${h2h.total} gare (V:${h2h.homeWins} P:${h2h.draws} S:${h2h.awayWins}) | Ultimo: ${h2h.lastMatch}`);
            } else {
                console.log(`⚪ ${home} vs ${away} -> Nessun precedente decennale`);
            }
        }
    });

    sheetFixtures.views = [{ state: 'frozen', ySplit: 1 }];
    sheetFixtures.columns.forEach(col => {
        let maxLen = 12;
        col.eachCell({ includeEmpty: false }, cell => {
            const len = cell.value ? cell.value.toString().length : 0;
            if (len > maxLen) maxLen = Math.min(len + 3, 26);
        });
        col.width = maxLen;
    });

    await safeWriteWorkbook(workbook, EXCEL_PATH);

    console.log(`\n======================================================`);
    console.log(`🎯 H2H E QUOTE RICALCOLATI CON PRECISIONE!`);
    console.log(`📊 Partite elaborate: ${processed}`);
    console.log(`======================================================\n`);
}

elaboraConStorico().catch(err => {
    console.error('❌ Errore durante l\'elaborazione:', err);
});