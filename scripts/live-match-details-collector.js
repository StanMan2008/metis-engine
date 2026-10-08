import 'dotenv/config';
import axios from 'axios';
import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const EXCEL_PATH = path.join(__dirname, '../data', 'metis_database.xlsx');

const API_KEY = process.env.SPORTSCORE_API_KEY;
const API_HOST = 'sportscore1.p.rapidapi.com';
const API_BASE_URL = `https://${API_HOST}`;

// Compromesso ideale: intervallo di 2 minuti (120000 ms) per proteggere le quote API
const POLLING_INTERVAL_MS = 120000; 

const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));

async function raccogliDettagliLiveTarget() {
    console.log(`\n🔍 [METIS ENGINE] Avvio sincronizzazione mirata (competizioni target)...`);

    try {
        const workbook = new ExcelJS.Workbook();
        if (!fs.existsSync(EXCEL_PATH)) {
            console.error('❌ Errore: File metis_database.xlsx non trovato!');
            return;
        }
        await workbook.xlsx.readFile(EXCEL_PATH);

        // 1. Estraiamo le competizioni monitorate dal foglio Excel "Campionati" (Colonna 3 = Nome)
        const sheetCampionati = workbook.getWorksheet('Campionati');
        const allowedTournaments = new Set();
        
        if (sheetCampionati) {
            sheetCampionati.eachRow((row, rowNumber) => {
                if (rowNumber === 1) return; // Salta l'intestazione
                const nomeCampionato = row.getCell(3).value; 
                if (nomeCampionato) allowedTournaments.add(nomeCampionato.toString().trim().toLowerCase());
            });
        }
        console.log(`📋 Competizioni configurate lette dall'Excel: ${allowedTournaments.size}`);

        const sheetDettagli = workbook.getWorksheet('PartiteDettagli');
        const sheetStatSquadre = workbook.getWorksheet('StatisticheSquadre');

        if (!sheetDettagli || !sheetStatSquadre) {
            console.error('❌ Errore: Schede Excel "PartiteDettagli" o "StatisticheSquadre" mancanti.');
            return;
        }

        console.log(`📡 Interrogazione API SportScore (/events/live)...`);
        const responseMatches = await axios.get(`${API_BASE_URL}/events/live`, {
            headers: { 
                'x-rapidapi-host': API_HOST, 
                'x-rapidapi-key': API_KEY 
            },
            params: { page: 1 }
        });

        const allMatches = responseMatches.data?.data || [];
        console.log(`⚽ Partite live globali trovate: ${allMatches.length}`);

        // 2. Filtraggio rigoroso: Solo Calcio E presenti nel nostro Excel delle competizioni target
        const targetMatches = allMatches.filter(match => {
            const isFootball = match.sport_id === 1 || match.tournament?.sport_id === 1;
            const tournamentName = (match.tournament?.name || '').trim().toLowerCase();
            
            const isAllowed = allowedTournaments.size === 0 || allowedTournaments.has(tournamentName);

            return isFootball && isAllowed;
        });

        console.log(`🎯 Partite live corrispondenti alle tue competizioni: ${targetMatches.length}`);

        if (targetMatches.length === 0) {
            console.log(`ℹ️ Nessun match live attivo in questo momento per le leghe monitorate.`);
            return;
        }

        for (const match of targetMatches) {
            const matchId = match.id;
            const homeName = match.home_team?.name || 'Casa';
            const awayName = match.away_team?.name || 'Ospite';
            const leagueName = match.tournament?.name || 'Sconosciuta';
            
            console.log(`⚡ In elaborazione: [${leagueName}] ${homeName} vs ${awayName} (ID: ${matchId})`);

            try {
                // Recupero Incidenti / Timeline in tempo reale
                const responseEvents = await axios.get(`${API_BASE_URL}/events/${matchId}/incidents`, {
                    headers: { 'x-rapidapi-host': API_HOST, 'x-rapidapi-key': API_KEY }
                });

                const events = responseEvents.data?.data || [];
                for (const ev of events) {
                    sheetDettagli.addRow({
                        1: `EV_${ev.id || Math.floor(Math.random() * 100000)}`,
                        2: `MATCH_${matchId}`,
                        3: ev.minute || 0,
                        4: ev.second || 0,
                        5: ev.period || '1T',
                        6: ev.team_id ? `TEAM_${ev.team_id}` : '',
                        7: ev.team_name || '',
                        8: ev.player_id ? `PLAYER_${ev.player_id}` : '',
                        9: ev.player_name || '',
                        10: ev.type || 'evento'
                    });
                }

                // Recupero Statistiche di Partita (Possesso, Tiri, Corner, ecc.)
                const responseStats = await axios.get(`${API_BASE_URL}/events/${matchId}/statistics`, {
                    headers: { 'x-rapidapi-host': API_HOST, 'x-rapidapi-key': API_KEY }
                });

                const stats = responseStats.data?.data || [];
                for (const statTeam of stats) {
                    sheetStatSquadre.addRow({
                        1: statTeam.team_id ? `TEAM_${statTeam.team_id}` : '',
                        2: statTeam.team_name || '',
                        3: match.tournament?.id ? `COMP_${match.tournament.id}` : '',
                        4: leagueName,
                        5: match.season?.name || '2026-2027',
                        6: statTeam.possession_percentage || 0,
                        7: statTeam.shots_total || 0,
                        8: statTeam.shots_on_target || 0,
                        9: statTeam.fouls || 0,
                        10: statTeam.corners || 0
                    });
                }

                // Pausa breve tra una partita e l'altra per rispettare le soglie di rate limit
                await delay(1000);

            } catch (err) {
                console.log(`⚠️ Impossibile recuperare i dettagli per il match ${matchId}: ${err.message}`);
            }
        }

        // Salvataggio sicuro delle modifiche su Excel
        await workbook.xlsx.writeFile(EXCEL_PATH);
        console.log(`💾 [EXCEL] Database aggiornato con successo.`);

    } catch (error) {
        console.error("❌ Errore critico nel ciclo di sincronizzazione:", error.message);
    }
}

// 🔄 DEMONE CONTINUO (LOOP INFINITO CON INTERVALLO DI 2 MINUTI)
async function startLiveDaemon() {
    console.log(`🚀 [METIS DAEMON] Avviato con successo.`);
    console.log(`⏱️ Frequenza di aggiornamento: ogni 2 minuti.`);
    
    while (true) {
        const timestamp = new Date().toLocaleTimeString();
        console.log(`\n==================================================`);
        console.log(`⏰ [${timestamp}] Controllo ciclico in corso...`);
        
        await raccogliDettagliLiveTarget();
        
        console.log(`💤 In attesa del prossimo ciclo (tra 2 minuti)...`);
        await delay(POLLING_INTERVAL_MS);
    }
}

// Avvio automatico del demone
startLiveDaemon();

// ============================================================================
// 📌 METIS ENGINE — NOTE & COMANDI UTILI (PROMEMORIA RAPIDO)
// ============================================================================
//
// 🚀 GESTIONE PROCESSI IN BACKGROUND (PM2):
//   • Avviare il collector live: 
//     pm2 start scripts/live-match-details-collector.js
//
//   • Controllare lo stato dei processi attivi: 
//     pm2 status
//
//   • Visualizzare i log in tempo reale (per vedere cosa sta scaricando): 
//     pm2 logs live-match-details-collector
//
//   • Fermare il processo in background: 
//     pm2 stop live-match-details-collector
//
//   • Riavviare il processo: 
//     pm2 restart live-match-details-collector
//
//   • Eliminare completamente il processo dalla lista di PM2: 
//     pm2 delete live-match-details-collector
//
// ----------------------------------------------------------------------------
// 🛠️ COMANDI NPM DI BASE:
//   • Avviare il server backend principale: 
//     npm start (o npm run dev)
//
// ============================================================================