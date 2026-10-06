import express from 'express';
import path from 'path';
import ExcelJS from 'exceljs';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.static(path.join(__dirname, 'public')));

app.get('/api/data', async (req, res) => {
    try {
        const excelPath = path.join(__dirname, 'data/metis_database.xlsx');
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(excelPath);
        
        const data = {};
        workbook.worksheets.forEach(sheet => {
            const sheetName = sheet.name;
            const rows = [];
            let headers = [];

            sheet.eachRow((row, rowNumber) => {
                if (rowNumber === 1) {
                    headers = row.values.slice(1);
                } else {
                    const obj = {};
                    row.values.slice(1).forEach((val, index) => {
                        const headerKey = headers[index] || `col_${index}`;
                        obj[headerKey] = val !== null && val !== undefined ? val : '';
                    });
                    rows.push(obj);
                }
            });
            data[sheetName] = rows;
        });

        res.json(data);
    } catch (error) {
        console.error("Errore lettura Excel su Cloud:", error);
        res.status(500).json({ error: 'Impossibile leggere il database Excel' });
    }
});

app.listen(PORT, () => {
    console.log(`Server attivo sulla porta ${PORT}`);
});

export default app;