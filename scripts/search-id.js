import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import ExcelJS from 'exceljs';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const filePath = path.join(__dirname, '../data', 'metis_database.xlsx');

// Master List pulita e mappata con precisione millimetrica
const COMPETIZIONI = [
    // ITALIA (4)
    { id: 32, paese: 'Italia', nome: 'Italian Serie A', tipo: 'League', livello: 1 },
    { id: 33, paese: 'Italia', nome: 'Italian Serie B', tipo: 'League', livello: 2 },
    { id: 34, paese: 'Italia', nome: 'Italian Serie C', tipo: 'League', livello: 3 },
    { id: 35, paese: 'Italia', nome: 'Coppa Italia', tipo: 'Cup', livello: 1 },
    
    // GERMANIA (4)
    { id: 36, paese: 'Germania', nome: 'Bundesliga', tipo: 'League', livello: 1 },
    { id: 37, paese: 'Germania', nome: '2. Bundesliga', tipo: 'League', livello: 2 },
    { id: 38, paese: 'Germania', nome: '3. Liga', tipo: 'League', livello: 3 },
    { id: 39, paese: 'Germania', nome: 'DFB-Pokal', tipo: 'Cup', livello: 1 },
    
    // SPAGNA (3)
    { id: 40, paese: 'Spagna', nome: 'La Liga', tipo: 'League', livello: 1 },
    { id: 'esp_2', paese: 'Spagna', nome: 'La Liga 2', tipo: 'League', livello: 2 },
    { id: 'esp_cup', paese: 'Spagna', nome: 'Copa del Rey', tipo: 'Cup', livello: 1 },
    
    // PORTOGALLO (2)
    { id: 41, paese: 'Portogallo', nome: 'Liga Portugal', tipo: 'League', livello: 1 },
    { id: 'por_2', paese: 'Portogallo', nome: 'Liga Portugal 2', tipo: 'League', livello: 2 },
    
    // FRANCIA (3)
    { id: 42, paese: 'Francia', nome: 'Ligue 1', tipo: 'League', livello: 1 },
    { id: 43, paese: 'Francia', nome: 'Ligue 2', tipo: 'League', livello: 2 },
    { id: 'fra_cup', paese: 'Francia', nome: 'Coupe de France', tipo: 'Cup', livello: 1 },
    
    // INGHILTERRA (6)
    { id: 44, paese: 'Inghilterra', nome: 'Premier League', tipo: 'League', livello: 1 },
    { id: 45, paese: 'Inghilterra', nome: 'Championship', tipo: 'League', livello: 2 },
    { id: 'eng_3', paese: 'Inghilterra', nome: 'League One', tipo: 'League', livello: 3 },
    { id: 'eng_4', paese: 'Inghilterra', nome: 'League Two', tipo: 'League', livello: 4 },
    { id: 'eng_cup1', paese: 'Inghilterra', nome: 'FA Cup', tipo: 'Cup', livello: 1 },
    { id: 'eng_cup2', paese: 'Inghilterra', nome: 'Carabao Cup', tipo: 'Cup', livello: 1 },
    
    // EUROPA (8)
    { id: 'den_1', paese: 'Danimarca', nome: 'Danish Superliga', tipo: 'League', livello: 1 },
    { id: 'sui_1', paese: 'Svizzera', nome: 'Swiss Super League', tipo: 'League', livello: 1 },
    { id: 'nor_1', paese: 'Norvegia', nome: 'Eliteserien', tipo: 'League', livello: 1 },
    { id: 'swe_1', paese: 'Svezia', nome: 'Allsvenskan', tipo: 'League', livello: 1 },
    { id: 'pol_1', paese: 'Polonia', nome: 'Ekstraklasa', tipo: 'League', livello: 1 },
    { id: 'rom_1', paese: 'Romania', nome: 'SuperLiga', tipo: 'League', livello: 1 },
    { id: 'tur_1', paese: 'Turchia', nome: 'Super Lig', tipo: 'League', livello: 1 },
    { id: 'gre_1', paese: 'Grecia', nome: 'Super League', tipo: 'League', livello: 1 },
    
    // ASIA (4)
    { id: 'ksa_1', paese: 'Arabia Saudita', nome: 'Saudi Pro League', tipo: 'League', livello: 1 },
    { id: 'chn_1', paese: 'Cina', nome: 'Chinese Super League', tipo: 'League', livello: 1 },
    { id: 'jpn_1', paese: 'Giappone', nome: 'J1 League', tipo: 'League', livello: 1 },
    { id: 'kor_1', paese: 'Corea del Sud', nome: 'K League 1', tipo: 'League', livello: 1 },
    
    // AMERICHE (7)
    { id: 'usa_1', paese: 'USA', nome: 'MLS', tipo: 'League', livello: 1 },
    { id: 'mex_1', paese: 'Messico', nome: 'Liga MX', tipo: 'League', livello: 1 },
    { id: 'chi_1', paese: 'Cile', nome: 'Chilean Primera Division', tipo: 'League', livello: 1 },
    { id: 'bra_1', paese: 'Brasile', nome: 'Brazilian Serie A', tipo: 'League', livello: 1 },
    { id: 'bra_2', paese: 'Brasile', nome: 'Brazilian Serie B', tipo: 'League', livello: 2 },
    { id: 'arg_1', paese: 'Argentina', nome: 'Argentine Primera Division', tipo: 'League', livello: 1 },
    { id: 'ecu_1', paese: 'Ecuador', nome: 'LigaPro Serie A', tipo: 'League', livello: 1 },
    
    // AFRICA (1)
    { id: 'egy_1', paese: 'Egitto', nome: 'Egyptian Premier League', tipo: 'League', livello: 1 }
];

async function main() {
    console.log("🚀 Ricostruzione pulita del foglio Campionati...");
    
    const workbook = new ExcelJS.Workbook();
    if (fs.existsSync(filePath)) {
        await workbook.xlsx.readFile(filePath);
    }

    let sheet = workbook.getWorksheet('Campionati');
    if (sheet) {
        workbook.removeWorksheet(sheet.id);
    }
    
    sheet = workbook.addWorksheet('Campionati');

    // Impostiamo direttamente le intestazioni usando le chiavi di colonna per evitare disallineamenti
    sheet.columns = [
        { header: 'ID', key: 'id', width: 15 },
        { header: 'Paese', key: 'paese', width: 15 },
        { header: 'Nome', key: 'nome', width: 30 },
        { header: 'Slug SportScore', key: 'slug', width: 25 },
        { header: 'URL SportScore', key: 'url', width: 40 },
        { header: 'Logo', key: 'logo', width: 10 },
        { header: 'Stagione', key: 'stagione', width: 15 },
        { header: 'Stagione ID', key: 'stagioneId', width: 15 },
        { header: 'Tipo', key: 'tipo', width: 12 },
        { header: 'Livello', key: 'livello', width: 10 },
        { header: 'Squadre', key: 'squadre', width: 12 },
        { header: 'Partite previste', key: 'partitePreviste', width: 18 },
        { header: 'Partite giocate', key: 'partiteGiocate', width: 15 },
        { header: 'Fonte', key: 'fonte', width: 15 },
        { header: 'Ultimo aggiornamento', key: 'ultimoAggiornamento', width: 25 }
    ];

    // Inserimento riga per riga tramite oggetti chiave-valore (impossibile sbagliare colonna)
    COMPETIZIONI.forEach(c => {
        const slug = c.nome.toLowerCase().replace(/\s+/g, '-').replace(/[^a-z0-9-]/g, '');
        sheet.addRow({
            id: `COMP_${c.id}`,
            paese: c.paese,
            nome: c.nome,
            slug: slug,
            url: `https://sportscore.io/leagues/${c.id}`,
            logo: '',
            stagione: '2026-2027',
            stagioneId: '',
            tipo: c.tipo,
            livello: c.livello,
            squadre: c.tipo === 'League' ? 20 : 0,
            partitePreviste: c.tipo === 'League' ? 380 : 0,
            partiteGiocate: 0,
            fonte: 'METIS_INIT',
            ultimoAggiornamento: new Date().toISOString()
        });
    });

    await workbook.xlsx.writeFile(filePath);
    console.log(`✅ Fatto! Tabella Campionati riscritta perfettamente senza sovrapposizioni.`);
    console.log(`📊 Totale competizioni inserite: ${COMPETIZIONI.length}`);
}

main().catch(err => console.error("❌ Errore critico:", err));