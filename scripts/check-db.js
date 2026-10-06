import fs from 'fs';
import path from 'path';
import ExcelJS from 'exceljs';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const EXCEL_PATH = path.join(__dirname, '../data/metis_database.xlsx');

async function checkDatabase() {
    if (!fs.existsSync(EXCEL_PATH)) {
        console.log('❌ File Excel non trovato!');
        return;
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(EXCEL_PATH);

    console.log('\n📊 === DIAGNOSTICA METIS DATABASE ===');
    console.log(`📁 File: ${EXCEL_PATH}\n`);

    workbook.worksheets.forEach(sh => {
        console.log(`📑 Foglio: [${sh.name}] -> ${sh.rowCount} righe totali`);
    });

    // Mostra le prime righe del foglio Classifiche se presente
    const sheetClass = workbook.getWorksheet('Classifiche');
    if (sheetClass) {
        console.log('\n🏆 Anteprima Classifiche (Top 5 righe):');
        let count = 0;
        sheetClass.eachRow((row, rowNumber) => {
            if (rowNumber <= 6) {
                console.log(`   Riga ${rowNumber}:`, row.values.slice(1));
            }
        });
    }

    console.log('\n=====================================\n');
}

checkDatabase().catch(console.error);