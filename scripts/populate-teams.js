import fs from 'fs';
import path from 'path';
import axios from 'axios';
import ExcelJS from 'exceljs';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const EXCEL_PATH = path.join(__dirname, '../data/metis_database.xlsx');
const HISTORICAL_DIR = path.join(__dirname, '../data/historical');
const MAIN_DIR = path.join(HISTORICAL_DIR, 'main_leagues');
const EXTRA_DIR = path.join(HISTORICAL_DIR, 'extra_leagues');

const http = axios.create({
    headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'application/json'
    },
    timeout: 8000
});

function cleanSlug(str) {
    if (!str) return '';
    return str.toLowerCase().replace(/[^a-z0-9]/g, '-').replace(/-+/g, '-').replace(/^-|-$/g, '');
}

function cleanStr(str) {
    if (!str) return '';
    return str.toLowerCase().replace(/[^a-z0-9]/g, '');
}

async function safeWriteWorkbook(workbook, filePath, retries = 3) {
    for (let i = 0; i < retries; i++) {
        try {
            await workbook.xlsx.writeFile(filePath);
            console.log(`💾 File Excel salvato: ${filePath}`);
            return;
        } catch (err) {
            if (err.code === 'EBUSY') {
                console.warn(`⚠️ File bloccato. Tentativo ${i + 1}/${retries}...`);
                await new Promise(r => setTimeout(r, 2000));
            } else throw err;
        }
    }
}

// Dizionario per correlare i nomi di Excel con i file CSV
const LEAGUE_FILE_MAP = {
    'premierleague': ['premier_league'],
    'championship': ['championship'],
    'leagueone': ['league_one'],
    'leaguetwo': ['league_two'],
    'laliga': ['la_liga'],
    'laliga2': ['segunda_division'],
    'ligue1': ['ligue_1'],
    'ligue2': ['ligue_2'],
    'bundesliga': ['bundesliga'],
    '2bundesliga': ['2_bundesliga'],
    'italianseriea': ['serie_a'],
    'italianserieb': ['serie_b'],
    'ligaportugal': ['liga_portugal'],
    'eredivisie': ['eredivisie_olanda'],
    'jupiler': ['jupiler_league_belgio'],
    'premiership': ['premiership_scozia'],
    'superlig': ['super_lig_turchia'],
    'superleague': ['super_league_grecia', 'svizzera_super_league'],
    'danishsuperliga': ['danimarca_superliga'],
    'swisssuperleague': ['svizzera_super_league'],
    'eliteserien': ['norvegia_eliteserien'],
    'allsvenskan': ['svezia_allsvenskan'],
    'ekstraklasa': ['polonia_ekstraklasa'],
    'superliga': ['romania_superliga', 'danimarca_superliga'],
    'chinesesuperleague': ['cina_super_league'],
    'j1league': ['giappone_j1_league'],
    'mls': ['usa_mls'],
    'ligamx': ['messico_liga_mx'],
    'brazilianseriea': ['brasile_serie_a'],
    'argentina': ['argentina_primera']
};

function getTeamsFromFiles() {
    const teamsByFile = new Map();
    const dirs = [MAIN_DIR, EXTRA_DIR];

    dirs.forEach(dir => {
        if (!fs.existsSync(dir)) return;
        const files = fs.readdirSync(dir);
        files.forEach(f => {
            if (!f.endsWith('.csv')) return;
            const normF = cleanStr(f.replace('.csv', ''));
            try {
                const content = fs.readFileSync(path.join(dir, f), 'utf-8');
                const lines = content.split('\n');
                if (lines.length < 2) return;
                const headers = lines[0].split(',').map(h => h.trim().replace(/^"|"$/g, ''));
                const homeIdx = headers.findIndex(h => ['HomeTeam', 'Home', 'home'].includes(h));
                if (homeIdx === -1) return;

                if (!teamsByFile.has(normF)) teamsByFile.set(normF, new Set());

                for (let i = 1; i < lines.length; i++) {
                    const parts = lines[i].split(',');
                    const team = parts[homeIdx] ? parts[homeIdx].trim().replace(/^"|"$/g, '') : '';
                    if (team && team.length >= 3 && isNaN(team) && !team.includes('?')) {
                        teamsByFile.get(normF).add(team);
                    }
                }
            } catch {}
        });
    });
    return teamsByFile;
}

async function popolaSquadre() {
    console.log('🚀 [METIS ENGINE] Popolamento completo anagrafica squadre...\n');

    if (!fs.existsSync(EXCEL_PATH)) return;

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(EXCEL_PATH);

    const sheetCampionati = workbook.getWorksheet('Campionati');
    if (!sheetCampionati) return;

    const campionati = [];
    sheetCampionati.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;
        campionati.push({
            id: row.getCell(1).value,
            paese: row.getCell(2).value,
            nome: row.getCell(3).value,
            slug: row.getCell(4).value
        });
    });

    let sheetSquadre = workbook.getWorksheet('Squadre');
    if (sheetSquadre) workbook.removeWorksheet(sheetSquadre.id);
    sheetSquadre = workbook.addWorksheet('Squadre');

    sheetSquadre.addRow([
        'ID Squadra', 'ID Campionato', 'Campionato', 'Nome', 'Nome Breve', 'Slug', 'Logo URL', 'Paese'
    ]);

    const csvData = getTeamsFromFiles();
    const registeredTeams = new Set();
    let teamCounter = 1;

    for (const camp of campionati) {
        let leagueTeams = [];
        const cleanNomeCamp = cleanStr(camp.nome);

        // Fallback incrociato sui CSV storici
        for (const [keyToken, fileAliases] of Object.entries(LEAGUE_FILE_MAP)) {
            if (cleanNomeCamp.includes(keyToken) || cleanStr(camp.slug).includes(keyToken)) {
                for (const alias of fileAliases) {
                    const normAlias = cleanStr(alias);
                    for (const [fileName, teamSet] of csvData.entries()) {
                        if (fileName.includes(normAlias)) {
                            teamSet.forEach(t => {
                                leagueTeams.push({
                                    nome: t,
                                    short: t.slice(0, 3).toUpperCase(),
                                    slug: cleanSlug(t),
                                    logo: ''
                                });
                            });
                        }
                    }
                }
                break;
            }
        }

        // Se non trovato con il dizionario, scansione su corrispondenza libera
        if (leagueTeams.length === 0) {
            for (const [fileName, teamSet] of csvData.entries()) {
                if (cleanNomeCamp.length >= 6 && fileName.includes(cleanNomeCamp)) {
                    teamSet.forEach(t => {
                        leagueTeams.push({
                            nome: t,
                            short: t.slice(0, 3).toUpperCase(),
                            slug: cleanSlug(t),
                            logo: ''
                        });
                    });
                }
            }
        }

        leagueTeams.forEach(t => {
            const uniqueKey = `${camp.id}_${cleanStr(t.nome)}`;
            if (!registeredTeams.has(uniqueKey)) {
                registeredTeams.add(uniqueKey);
                sheetSquadre.addRow([
                    `TEAM_${String(teamCounter).padStart(4, '0')}`,
                    camp.id,
                    camp.nome,
                    t.nome,
                    t.short,
                    t.slug,
                    t.logo,
                    camp.paese
                ]);
                teamCounter++;
            }
        });

        console.log(`🛡️ [${camp.paese} - ${camp.nome}] Club censiti: ${leagueTeams.length}`);
    }

    sheetSquadre.views = [{ state: 'frozen', ySplit: 1 }];
    sheetSquadre.columns.forEach(col => {
        let maxLen = 12;
        col.eachCell({ includeEmpty: false }, cell => {
            const len = cell.value ? cell.value.toString().length : 0;
            if (len > maxLen) maxLen = Math.min(len + 3, 30);
        });
        col.width = maxLen;
    });

    await safeWriteWorkbook(workbook, EXCEL_PATH);
    console.log(`\n======================================================`);
    console.log(`✅ SQUADRE ALINEATE! Totale registrati: ${teamCounter - 1}`);
    console.log(`======================================================\n`);
}

popolaSquadre().catch(console.error);