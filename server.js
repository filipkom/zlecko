const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const XLSX = require('xlsx');
const path = require('path');

const app = express();
const port = process.env.PORT || 3000;

app.use(cors());
app.use(bodyParser.json());
app.use(express.static('public'));

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.post('/api/generuj-excel', (req, res) => {
  const { nazwa } = req.body;
  
  try {
    const workbook = XLSX.readFile(path.join(__dirname, 'public', 'szablon.xlsx'));
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];

    // Wprowadź nazwę firmy do komórki A5
    XLSX.utils.sheet_add_aoa(worksheet, [[nazwa]], { origin: 'A5' });

    const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });

    res.setHeader('Content-Disposition', `attachment; filename=Zlecenie_${nazwa}.xlsx`);
    res.type('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(buffer);
  } catch (error) {
    console.error("Błąd podczas generowania pliku Excel:", error);
    res.status(500).json({ message: 'Wystąpił błąd podczas generowania pliku Excel' });
  }
});

app.listen(port, () => {
  console.log(`Server działa na http://localhost:${port}`);
});
