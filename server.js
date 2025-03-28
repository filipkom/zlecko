const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const ExcelJS = require('exceljs');
const path = require('path');

const app = express();
const port = process.env.PORT || 3000;

app.use(cors());
app.use(bodyParser.json());
app.use(express.static(path.join(__dirname, 'public')));

let daneFirm = [];

// Wczytaj dane firm z pliku Excel przy starcie serwera
async function loadFirmy() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(path.join(__dirname, 'public', 'dane_firm.xlsx'));
    const worksheet = workbook.getWorksheet(1);

    worksheet.eachRow((row, rowNumber) => {
        if (rowNumber > 1) { // Pomijamy nagłówki
            daneFirm.push({
                nazwa: row.getCell(1).value,
                adres: row.getCell(2).value,
                nip: row.getCell(3).value,
            });
        }
    });
}

// Endpoint do pobierania listy firm na podstawie wpisanych liter
app.get('/api/firmy', (req, res) => {
    const { nazwa } = req.query;

    if (!nazwa || nazwa.length < 3) {
        return res.status(400).json({ message: 'Wpisz co najmniej 3 znaki' });
    }

    const filtered = daneFirm.filter(f =>
        f.nazwa.toLowerCase().startsWith(nazwa.toLowerCase())
    );

    res.json(filtered);
});

// Endpoint do generowania pliku Excel
app.post('/api/generuj-excel', async (req, res) => {
    const { nazwa } = req.body;

    try {
        const firma = daneFirm.find(f => f.nazwa === nazwa);
        if (!firma) throw new Error('Nie znaleziono firmy');

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(path.join(__dirname, 'public', 'szablon.xlsx'));
        const worksheet = workbook.getWorksheet(1);

        worksheet.getCell('A5').value = firma.nazwa;
        worksheet.getCell('A6').value = firma.adres;
        worksheet.getCell('A7').value = firma.nip;

        res.setHeader(
            'Content-Type',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        );
        res.setHeader(
            'Content-Disposition',
            `attachment; filename=Zlecenie_${firma.nazwa}.xlsx`
        );

        await workbook.xlsx.write(res);
        res.end();
    } catch (error) {
        console.error(error);
        res.status(500).json({ message: error.message });
    }
});

// Obsługa błędów dla nieznanych ścieżek
app.use((req, res) => {
    res.status(404).send('Nie znaleziono strony');
});

// Obsługa błędów globalnych
app.use((err, req, res, next) => {
    console.error(err.stack);
    res.status(500).send('Wystąpił błąd serwera');
});

loadFirmy()
    .then(() => {
        app.listen(port, () => {
            console.log(`Serwer działa na porcie ${port}`);
        });
    })
    .catch(error => {
        console.error('Błąd inicjalizacji:', error);
        process.exit(1);
    });
