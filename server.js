const ExcelJS = require('exceljs');

// Wczytaj dane firm z pliku Excel
async function loadFirmy() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(path.join(__dirname, 'public', 'dane_firm.xlsx'));
    const worksheet = workbook.getWorksheet(1);
    const firmy = [];

    worksheet.eachRow((row, rowNumber) => {
        if (rowNumber > 1) { // Pomijamy nagłówki
            firmy.push({
                nazwa: row.getCell(1).value,
                adres: row.getCell(2).value,
                nip: row.getCell(3).value,
            });
        }
    });
    return firmy;
}

// Zmodyfikuj endpoint do pobierania danych firmy
app.post('/api/firmy', async (req, res) => {
    const { nazwa } = req.body;
    const firmy = await loadFirmy();
    const firma = firmy.find(f => f.nazwa.toLowerCase().startsWith(nazwa.toLowerCase()));
    
    if (firma) {
        res.json(firma);
    } else {
        res.status(404).json({ message: 'Nie znaleziono firmy' });
    }
});

// Zmodyfikuj endpoint do generowania pliku Excel
app.post('/api/generuj-excel', async (req, res) => {
    const { nazwa } = req.body;
    const firmy = await loadFirmy();
    const firma = firmy.find(f => f.nazwa.toLowerCase() === nazwa.toLowerCase());

    if (firma) {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(path.join(__dirname, 'public', 'szablon.xlsx'));
        const worksheet = workbook.getWorksheet(1);

        worksheet.getCell('A5').value = firma.nazwa;
        worksheet.getCell('A6').value = firma.adres;
        worksheet.getCell('A7').value = firma.nip;

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename=Zlecenie_${firma.nazwa}.xlsx`);

        await workbook.xlsx.write(res);
        res.end();
    } else {
        res.status(404).json({ message: 'Nie znaleziono firmy' });
    }
});
