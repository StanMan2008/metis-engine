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

// Funzione di pausa
const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));

async function popolaTuttiICampionati() {
    console.log("🚀 Avvio scansione con gestione intelligente del Rate Limit (429)...\n");

    try {
        const workbook = new ExcelJS.Workbook();
        if (fs.existsSync(EXCEL_PATH)) {
            await workbook.xlsx.readFile(EXCEL_PATH);
        } else {
            console.error('❌ Errore: File metis_database.xlsx non trovato nella cartella data!');
            return;
        }

        const worksheet = workbook.getWorksheet('Campionati');
        if (!worksheet) {
            console.error('❌ Errore: Foglio "Campionati" non trovato nel file Excel!');
            return;
        }

        // Pulisce le vecchie righe lasciando l'intestazione
        if (worksheet.rowCount > 1) {
            worksheet.spliceRows(2, worksheet.rowCount - 1);
        }

        let page = 1;
        let totalAdded = 0;
        let hasMorePages = true;

        while (hasMorePages && page <= 10) {
            let success = false;
            let response;

            // Ciclo di retry automatico per la singola pagina in caso di 429
            while (!success) {
                try {
                    console.log(`📡 Scaricando la pagina ${page} da SportScore...`);
                    response = await axios.get(`${API_BASE_URL}/leagues`, {
                        params: { page: page },
                        headers: { 
                            'x-rapidapi-host': API_HOST, 
                            'x-rapidapi-key': API_KEY 
                        }
                    });
                    success = true;
                } catch (err) {
                    if (err.response && err.response.status === 429) {
                        console.log("⚠️ Rilevato limite di velocità (429). Attendo 10 secondi e riprovo la pagina...");
                        await delay(10000); // 10 secondi di raffreddamento
                    } else {
                        throw err; // Se è un altro errore, interrompe
                    }
                }
            }

            const leghe = response.data.data || [];
            
            if (leghe.length === 0) {
                hasMorePages = false;
                break;
            }

            for (const camp of leghe) {
                worksheet.addRow([
                    `COMP_${camp.id}`,
                    camp.host?.country || 'Internazionale',
                    camp.name,
                    camp.slug || '',
                    `https://sportscore.io/leagues/${camp.slug}/${camp.id}`,
                    camp.logo || '',
                    '2024-2025',
                    camp.id,
                    'League',
                    1,
                    camp.facts?.find(f => f.name === 'Number of rounds')?.value || 38,
                    380,
                    0,
                    'SportScore',
                    new Date().toISOString()
                ]);
                totalAdded++;
            }

            page++;
            
            // Pausa di sicurezza di 3 secondi tra una pagina e l'altra per non sovraccaricare il piano Basic
            await delay(3000);
        }

        // Salvataggio finale del file Excel
        await workbook.xlsx.writeFile(EXCEL_PATH);
        console.log(`\n✅ SUCCESSO! Aggiunte ${totalAdded} competizioni nel database Excel. File salvato correttamente.`);

    } catch (error) {
        console.error("❌ Errore critico durante il popolamento:", error.message);
        if (error.response) {
            console.error("Dettaglio API:", error.response.data);
        }
    }
}

popolaTuttiICampionati();