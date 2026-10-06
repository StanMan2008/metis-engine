import axios from 'axios';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const HISTORICAL_DIR = path.join(__dirname, '../data/historical');
const MAIN_DIR = path.join(HISTORICAL_DIR, 'main_leagues');
const EXTRA_DIR = path.join(HISTORICAL_DIR, 'extra_leagues');

// 1. Campionati Principali Europei (file suddivisi per stagione)
const MAIN_LEAGUES = [
    // Italia
    { code: 'I1', name: 'Serie_A' },
    { code: 'I2', name: 'Serie_B' },
    // Inghilterra
    { code: 'E0', name: 'Premier_League' },
    { code: 'E1', name: 'Championship' },
    { code: 'E2', name: 'League_One' },
    { code: 'E3', name: 'League_Two' },
    // Spagna
    { code: 'SP1', name: 'La_Liga' },
    { code: 'SP2', name: 'Segunda_Division' },
    // Germania
    { code: 'D1', name: 'Bundesliga' },
    { code: 'D2', name: '2_Bundesliga' },
    // Francia
    { code: 'F1', name: 'Ligue_1' },
    { code: 'F2', name: 'Ligue_2' },
    // Altri campionati con quote complete
    { code: 'P1', name: 'Liga_Portugal' },
    { code: 'T1', name: 'Super_Lig_Turchia' },
    { code: 'G1', name: 'Super_League_Grecia' },
    { code: 'N1', name: 'Eredivisie_Olanda' },
    { code: 'B1', name: 'Jupiler_League_Belgio' },
    { code: 'SC0', name: 'Premiership_Scozia' }
];

// Tutte le 11 stagioni dal 2016 al 2026/2027
const SEASONS = [
    '1617', '1718', '1819', '1920', '2021',
    '2122', '2223', '2324', '2425', '2526', '2627'
];

// 2. Extra Leagues mondiali (archivio storico completo aggregato in unico CSV)
const EXTRA_LEAGUES = [
    { file: 'ARG.csv', name: 'Argentina_Primera' },
    { file: 'BRA.csv', name: 'Brasile_Serie_A' },
    { file: 'USA.csv', name: 'USA_MLS' },
    { file: 'MEX.csv', name: 'Messico_Liga_MX' },
    { file: 'JPN.csv', name: 'Giappone_J1_League' },
    { file: 'CHN.csv', name: 'Cina_Super_League' },
    { file: 'NOR.csv', name: 'Norvegia_Eliteserien' },
    { file: 'SWE.csv', name: 'Svezia_Allsvenskan' },
    { file: 'DNK.csv', name: 'Danimarca_Superliga' },
    { file: 'POL.csv', name: 'Polonia_Ekstraklasa' },
    { file: 'SWZ.csv', name: 'Svizzera_Super_League' },
    { file: 'ROU.csv', name: 'Romania_SuperLiga' },
    { file: 'AUT.csv', name: 'Austria_Bundesliga' }
];

const sleep = (ms) => new Promise(r => setTimeout(r, ms));

async function downloadFile(url, destPath) {
    try {
        const res = await axios.get(url, {
            headers: {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
            },
            timeout: 10000
        });

        if (res.status === 200 && typeof res.data === 'string' && res.data.length > 200) {
            fs.writeFileSync(destPath, res.data, 'utf-8');
            return true;
        }
        return false;
    } catch {
        return false;
    }
}

async function avviaDownloadStorico() {
    console.log('🚀 [METIS ENGINE] Inizio download storico decennale (2016-2026/27)...\n');

    fs.mkdirSync(MAIN_DIR, { recursive: true });
    fs.mkdirSync(EXTRA_DIR, { recursive: true });

    let scaricatiMain = 0;
    let scaricatiExtra = 0;

    // --- PARTE 1: Download Main Leagues per stagione ---
    console.log('📦 1/2 Scaricamento campionati europei principali (11 stagioni)...');
    for (const league of MAIN_LEAGUES) {
        for (const season of SEASONS) {
            const fileName = `${league.name}_${season}.csv`;
            const destPath = path.join(MAIN_DIR, fileName);

            if (fs.existsSync(destPath)) {
                scaricatiMain++;
                continue;
            }

            const url = `https://www.football-data.co.uk/mmz4281/${season}/${league.code}.csv`;
            const ok = await downloadFile(url, destPath);

            if (ok) {
                console.log(`   ✅ [${season}] ${league.name}`);
                scaricatiMain++;
            }
            await sleep(100); // Evita picchi di connessioni
        }
    }

    // --- PARTE 2: Download Extra Leagues archivi storici ---
    console.log('\n🌎 2/2 Scaricamento campionati Extra/Mondiali (archivio cumulativo 10+ anni)...');
    for (const extra of EXTRA_LEAGUES) {
        const destPath = path.join(EXTRA_DIR, `${extra.name}.csv`);

        const url = `https://www.football-data.co.uk/new/${extra.file}`;
        const ok = await downloadFile(url, destPath);

        if (ok) {
            // Conta quante righe/partite contiene il file
            const lineCount = fs.readFileSync(destPath, 'utf-8').split('\n').length - 1;
            console.log(`   ✅ ${extra.name} -> ${lineCount} partite storiche archiviate!`);
            scaricatiExtra++;
        } else {
            console.log(`   ⚠️ Archivio per ${extra.name} non reperibile al link diretto.`);
        }
        await sleep(150);
    }

    console.log('\n======================================================');
    console.log('🎉 DOWNLOAD STORICO DECENNALE COMPLETATO!');
    console.log(`📂 Cartella file principali: ${MAIN_DIR} (${scaricatiMain} file stagionali)`);
    console.log(`📂 Cartella campionati mondiali: ${EXTRA_DIR} (${scaricatiExtra} file storici completi)`);
    console.log('======================================================\n');
}

avviaDownloadStorico();