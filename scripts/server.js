import express from 'express';
import path from 'path';
import ExcelJS from 'exceljs';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = process.env.PORT || 3000;
const EXCEL_PATH = path.join(__dirname, 'data/metis_database.xlsx');

app.use(express.static(path.join(__dirname, 'public')));

// API per leggere i dati dal database Excel e mandarli al sito
apiData(app, EXCEL_PATH);

app.listen(PORT, () => {
    console.log(`🚀 Metis Web App attiva su http://localhost:${PORT}`);
});

function apiData(app, filePath) {
    app.get('/api/data', async (req, res) => {
        try {
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.readFile(filePath);

            const result = {};
            workbook.worksheets.forEach(sh => {
                const rows = [];
                let headers = [];
                sh.eachRow((row, rowNumber) => {
                    if (rowNumber === 1) {
                        headers = row.values.slice(1);
                    } else {
                        const obj = {};
                        row.values.slice(1).forEach((val, idx) => {
                            const h = headers[idx] || `col_${idx}`;
                            obj[h] = val;
                        });
                        rows.push(obj);
                    }
                });
                result[sh.name] = rows;
            });

            res.json(result);
        } catch (err) {
            res.status(500).json({ error: err.message });
        }
    });
}