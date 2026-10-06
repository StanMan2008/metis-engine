import axios from 'axios';
import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const EXCEL_PATH = path.join(__dirname, '../data/metis_database.xlsx');
const BASE_STANDINGS_URL = 'https://sportscore.com/api/v1/standings/';
const BASE_SEARCH_URL = 'https://sportscore.com/api/v1/search/';

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

const http = axios.create({
    headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'application/json'
    },
    timeout: 15000
});

// Mappa esatta e definitiva per scavalcare i falsi positivi del motore di ricerca
const KNOWN_SLUG_MAP = {
    // Inghilterra
    'premier-league': 'english-premier-league',
    'championship': 'english-football-league-championship',
    'league-one': 'english-football-league-one',
    'league-two': 'english-football-league-two',
    
    // Spagna
    'la-liga': 'spanish-la-liga',
    'la-liga-2': 'spanish-segunda-division',
    'spanish-segunda': 'spanish-segunda-division',
    
    // Francia
    'ligue-1': 'french-ligue-1',
    'ligue-2': 'french-ligue-2',
    
    // Germania
    '2-bundesliga': 'german-bundesliga-2',
    '3-liga': 'german-3liga',
    
    // Portogallo
    'liga-portugal': 'portuguese-primera-liga',
    
    // Americhe
    'mls': 'united-states-major-league-soccer',
    'argentine-primera-division': 'copa-de-la-liga-profesional',
    'liga-mx': 'mexico-liga-mx',
    'chilean-primera-division': 'chilean-cup',
    
    // Correzioni mirate anti-falsi positivi
    'super-lig': 'turkish-super-league',
    'super-league': 'greek-super-league',
    'superliga': 'romanian-super-liga',
    
    // Altri campionati internazionali
    'saudi-pro-league': 'saudi-professional-league',
    'eliteserien': 'norwegian-eliteserien',
    'allsvenskan': 'sweden-allsvenskan',
    'ekstraklasa': 'pko-bank-polski-ekstraklasa',
    'swiss-super-league': 'switzerland-super-league',
    'j1-league': 'japanese-j1-league',
    'k-league-1': 'korean-k-league-1',
    'chinese-super-league': 'china-super-league'
};

async function fetchStandings(slug) {
    if (!slug) return null;
    try {
        const response = await http.get(`${BASE_STANDINGS_URL}?sport=football&slug=${encodeURIComponent(slug)}`);
        if (response.data && Array.isArray(response.data.tables) && response.data.tables.length > 0) {
            return response.data;
        }
        return null;
    } catch {
        return null;
    }
}

async function findSlugBySearch(nomeCampionato, paese) {
    const queries = [];
    if (paese && nomeCampionato) queries.push(`${paese} ${nomeCampionato}`);
    if (nomeCampionato) queries.push(nomeCampionato);

    for (const q of queries) {
        try {
            const res = await http.get(`${BASE_SEARCH_URL}?q=${encodeURIComponent(q)}&sport=football&limit=5`);
            const comps = res.data?.competitions || res.data?.data?.competitions || [];
            if (Array.isArray(comps) && comps.length > 0 && comps[0]?.slug) {
                return comps[0].slug;
            }
        } catch {
            // Prosegui con il prossimo tentativo
        }
        await sleep(150);
    }
    return null;
}

function parseStat(val) {
    const n = Number(val);
    return Number.isFinite(n) ? n : 0;
}

async function avviaRaccoltaDefinitiva() {
    console.log('🚀 [METIS ENGINE] Avvio sincronizzazione finale del database...\n');

    if (!fs.existsSync(EXCEL_PATH)) {
        console.error(`❌ ERRORE: File non trovato in: ${EXCEL_PATH}`);
        return;
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(EXCEL_PATH);

    const sheetCampionati = workbook.getWorksheet('Campionati');
    if (!sheetCampionati) {
        console.error('❌ ERRORE: Foglio "Campionati" non trovato nel file Excel.');
        return;
    }

    let sheetSquadre = workbook.getWorksheet('Squadre');
    if (!sheetSquadre) {
        sheetSquadre = workbook.addWorksheet('Squadre');
        sheetSquadre.addRow([
            'Team ID', 'Metis League ID', 'Campionato', 'Nome Squadra', 'Abbreviazione',
            'Slug Squadra', 'Logo URL', 'Paese', 'Stadio', 'Città', 'Capienza',
            'Anno Fondazione', 'Allenatore', 'Link', 'Sorgente', 'Data Aggiornamento'
        ]);
    } else if (sheetSquadre.rowCount > 1) {
        sheetSquadre.spliceRows(2, sheetSquadre.rowCount - 1);
    }

    let sheetClassifiche = workbook.getWorksheet('Classifiche');
    if (!sheetClassifiche) {
        sheetClassifiche = workbook.addWorksheet('Classifiche');
        sheetClassifiche.addRow([
            'Metis League ID', 'Campionato', 'Girone', 'Posizione', 'Nome Squadra',
            'Team ID', 'Slug Squadra', 'Giocate', 'Vinte', 'Pareggiate', 'Perse',
            'Gol Fatti', 'Gol Subiti', 'Differenza Reti', 'Punti', 'Data Aggiornamento'
        ]);
    } else if (sheetClassifiche.rowCount > 1) {
        sheetClassifiche.spliceRows(2, sheetClassifiche.rowCount - 1);
    }

    let sheetReport = workbook.getWorksheet('Report');
    if (!sheetReport) {
        sheetReport = workbook.addWorksheet('Report');
        sheetReport.addRow([
            'Data', 'Operazione', 'Esito', 'Paese', 'Campionato', 'Stato', 'Elementi', 'Dettagli'
        ]);
    }

    const dataDiOggi = new Date().toISOString();

    const campionatiList = [];
    sheetCampionati.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;
        campionatiList.push({
            rowNumber,
            metisId: row.getCell(1).value ? String(row.getCell(1).value).trim() : `COMP_${rowNumber}`,
            paese: row.getCell(2).value ? String(row.getCell(2).value).trim() : '',
            nomeCampionato: row.getCell(3).value ? String(row.getCell(3).value).trim() : '',
            slug: row.getCell(4).value ? String(row.getCell(4).value).trim() : ''
        });
    });

    let successCount = 0;
    let skipCount = 0;

    for (const camp of campionatiList) {
        if (!camp.nomeCampionato) continue;

        console.log(`📡 Elaborazione: ${camp.nomeCampionato} (${camp.paese || 'N/D'})...`);

        // Priorità assoluta alla mappa fissa per evitare associazioni errate
        let targetSlug = KNOWN_SLUG_MAP[camp.slug] || KNOWN_SLUG_MAP[camp.slug.toLowerCase()] || camp.slug;
        let data = targetSlug ? await fetchStandings(targetSlug) : null;

        // Se lo slug non è valido o assente, usa il motore di ricerca di fallback
        if (!data) {
            const foundSlug = await findSlugBySearch(camp.nomeCampionato, camp.paese);
            if (foundSlug && foundSlug !== targetSlug) {
                targetSlug = foundSlug;
                data = await fetchStandings(targetSlug);
            }
        }

        if (data && Array.isArray(data.tables) && data.tables.length > 0) {
            let totalTeams = 0;

            // Aggiorna lo slug validato nel foglio Campionati
            const rowInSheet = sheetCampionati.getRow(camp.rowNumber);
            rowInSheet.getCell(4).value = targetSlug;

            data.tables.forEach((tableGroup) => {
                const groupName = tableGroup.group || 'Regular Season';
                const rows = Array.isArray(tableGroup.rows) ? tableGroup.rows : [];

                rows.forEach((item) => {
                    totalTeams++;
                    const teamName = item.team || 'Sconosciuta';
                    const rawSlug = item.team_slug || teamName.toLowerCase().replace(/[^a-z0-9]/g, '-').replace(/-+/g, '-');
                    const teamSlug = rawSlug.replace(/^-|-$/g, '') || `team-${totalTeams}`;
                    const teamIdMetis = `TEAM_${teamSlug}`;

                    const pos = parseStat(item.pos ?? item.position);
                    const played = parseStat(item.p ?? item.played ?? item.matches);
                    const won = parseStat(item.w ?? item.won);
                    const draws = parseStat(item.d ?? item.draws);
                    const lost = parseStat(item.l ?? item.lost);
                    const goalsFor = parseStat(item.f ?? item.goals_for);
                    const goalsAgainst = parseStat(item.a ?? item.goals_against);
                    const gd = item.gd !== undefined ? parseStat(item.gd) : (goalsFor - goalsAgainst);
                    const points = parseStat(item.pts ?? item.points);

                    sheetSquadre.addRow([
                        teamIdMetis,
                        camp.metisId,
                        camp.nomeCampionato,
                        teamName,
                        item.team_short || '',
                        teamSlug,
                        item.team_logo || '',
                        camp.paese,
                        '', '', '', '', '',
                        `https://sportscore.com/teams/${teamSlug}`,
                        'SPORTSCORE_V1_OFFICIAL',
                        dataDiOggi
                    ]);

                    sheetClassifiche.addRow([
                        camp.metisId,
                        camp.nomeCampionato,
                        groupName,
                        pos,
                        teamName,
                        teamIdMetis,
                        teamSlug,
                        played,
                        won,
                        draws,
                        lost,
                        goalsFor,
                        goalsAgainst,
                        gd,
                        points,
                        dataDiOggi
                    ]);
                });
            });

            console.log(`   ✅ Sincronizzato con successo: ${totalTeams} squadre trovate! (slug: "${targetSlug}")`);
            successCount++;

            sheetReport.addRow([
                dataDiOggi, 'COLLECT_STANDINGS', 'SUCCESS', camp.paese, camp.nomeCampionato, 'OK', totalTeams, 'Sincronizzato'
            ]);
        } else {
            console.log(`   ⚠️ Nessuna classifica disponibile per "${camp.nomeCampionato}".`);
            skipCount++;

            sheetReport.addRow([
                dataDiOggi, 'COLLECT_STANDINGS', 'WARNING', camp.paese, camp.nomeCampionato, 'NO_DATA', 0, 'Classifica non disponibile'
            ]);
        }

        await sleep(250);
    }

    await workbook.xlsx.writeFile(EXCEL_PATH);

    console.log(`\n======================================================`);
    console.log(`🎯 COMPLETATO!`);
    console.log(`✅ Campionati sincronizzati: ${successCount}`);
    console.log(`⚠️ Campionati saltati/senza classifica: ${skipCount}`);
    console.log(`📁 File Excel aggiornato in: ${EXCEL_PATH}`);
    console.log(`======================================================\n`);
}

avviaRaccoltaDefinitiva().catch((err) => {
    console.error('❌ Errore fatale durante l\'esecuzione:', err);
});