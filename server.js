const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const ExcelJS = require('exceljs');
const path = require('path');

const app = express();
const port = process.env.PORT || 3000;

app.use(cors());
app.use(bodyParser.json());
app.use(express.static('public'));

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.post('/api/generuj-excel', async (req, res) => {
  const { nazwa } = req.body;
  
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(path.join(__dirname, 'public', 'szablon.xlsx'));
    const worksheet = workbook.getWorksheet(1);

    // Wprowadź nazwę firmy do komórki A5
    const cell = worksheet.getCell('A5');
    const originalStyle = cell.style;
    cell.value = nazwa;
    cell.style = originalStyle;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=Zlecenie_${nazwa}.xlsx`);

    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    console.error("Błąd podczas generowania pliku Excel:", error);
    res.status(500).json({ message: 'Wystąpił błąd podczas generowania pliku Excel' });
  }
});

app.listen(port, () => {
  console.log(`Server działa na http://localhost:${port}`);
});
