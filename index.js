const express = require('express');
const app = express();
const { google } = require('googleapis');
const bodyParser = require('body-parser');
const path = require('path');

//const PORT = process.env.PORT || 3000;
const spreadsheetId = '1LJ3-aUgOrBK1wPCeJWUojxkajZh0lCmsamu6NC3Hle4';
const sheets = google.sheets('v4');


const auth = new google.auth.GoogleAuth({
        keyFile: 'keys.json',
        scopes: ['https://www.googleapis.com/auth/spreadsheets']
    });
// Body parser middleware
app.use(express.static('public'));

app.use(bodyParser.urlencoded({ extended: true }));

// Serve the index.html page when visiting the root URL
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

async function findData(spreadsheetId, range, searchValue) {
    const auth = new google.auth.GoogleAuth({
        keyFile: 'keys.json',
        scopes: ['https://www.googleapis.com/auth/spreadsheets']
    }); // Fungsi untuk melakukan autorisasi Google Sheets API
    const response = await sheets.spreadsheets.values.get({
      auth,
      spreadsheetId,
      range,
    });
    const rows = response.data.values;
    const rowIndex = rows.findIndex(row => row.includes(searchValue));
    return rowIndex !== -1;
  
}
  async function sendData(spreadsheetId, range, data) {
    const auth = new google.auth.GoogleAuth({
        keyFile: 'keys.json',
        scopes: ['https://www.googleapis.com/auth/spreadsheets']
    }); // Fungsi untuk melakukan autorisasi Google Sheets API
    const response = await sheets.spreadsheets.values.append({
      auth,
      spreadsheetId,
      range,
      valueInputOption: 'USER_ENTERED',
      resource: {
        values: [data],
      },
    });
    console.log(`${response.data.updates.updatedCells} cells appended.`);
  }


// Handle form submission
app.post('/submit-form',async (req, res) => {
    const nis = req.body.NIS;
    const name = req.body.NAMA;
    const jenis_kelamin = req.body.JenisKelamin;
    const Mapel1 = req.body.Mapel1;
    const Nilai1 = req.body.Nilai1;
    const Mapel2 = req.body.Mapel2;
    const Nilai2 = req.body.Nilai2;
    const Mapel3 = req.body.Mapel1;
    const Nilai3 = req.body.Nilai3;
    const Mapel4 = req.body.Mapel1;
    const Nilai4 = req.body.Nilai4;
    const auth = new google.auth.GoogleAuth({
      keyFile: 'keys.json',
      scopes: ['https://www.googleapis.com/auth/spreadsheets']});

    const alreadyExists = await findData(spreadsheetId, 'A:K', nis); 
    if (alreadyExists) {
      return res.send(`<script>alert("Data telah pernah dikirim!"); window.location='/';</script>`);
    }
    await sendData(spreadsheetId, 'A:K', [nis, name, jenis_kelamin, Mapel1, Nilai1, Mapel2, Nilai2, Mapel3, Nilai3, Mapel4, Nilai4,]); 
    res.send(`<script>alert('Data berhasil ditambahkan!'); window.location='/';</script>`);
  });

//app.listen(PORT, () => {
   // console.log(`Server started on port ${PORT}`);
//});





