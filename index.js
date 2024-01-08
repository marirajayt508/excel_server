const express = require('express');
const ExcelJS = require('exceljs');
const cors = require('cors')
const app = express();
var path = require('path');

const port = 5000;

app.use(express.static('public'));
app.use(express.static('public'));
app.use(cors())
app.use(express.json({ limit: '10mb' }))

app.get('/', (req, res) => {
  // Assuming your React build files are in the 'public' folder
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.post('/download', (req, res) => {
  // Generate Excel file with dropdown menu data
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('B2B');

  const data = req.body.datas
// console.log(data)
  const columns = [
    { header: 'GSTIN of Supplier', key: 'gstin', width: 25 },
    { header: 'Invoice Number', key: 'inv', width: 20 },
    { header: 'Invoice date', key: 'ind', width: 15 },
    { header: 'Supplier Name', key: 'sn', width: 25 },
    { header: 'Invoice Value', key: 'invoiceNo', width: 15 },
    { header: 'Place Of Supply', key: 'pos', width: 15 },
    { header: 'Reverse Charge', key: 'rc', width: 15, dataValidation : {
      type: 'list',
      formula: ['Yes', 'No'],
    } },
    { header: 'Invoice Type', key: 'it', width: 15 },
    { header: 'Rate', key: 'rate', width: 5 },
    { header: 'Taxable Value', key: 'tv', width: 15 },
    { header: 'Integrated Tax Paid', key: 'itp', width: 20 },
    { header: 'Central Tax Paid', key: 'ctp', width: 15 },
    { header: 'State/UT Tax Paid', key: 'stp', width: 15 },
    { header: 'Cess Paid', key: 'cp', width: 15 },
    { header: 'Eligibility For ITC', key: 'eft', width: 15 },
    { header: 'Availed ITC Integrated Tax', key: 'aiit', width: 15 },
    { header: 'Availed ITC Central Tax', key: 'aict', width: 15 },
    { header: 'Availed ITC State/UT Tax', key: 'aist', width: 15 },
    { header: 'Availed ITC Cess', key: 'atc', width: 15 },
  ];

  worksheet.columns = columns;

  worksheet.addTable({
    name: 'Elixir',
    ref: 'A1:S4',
    headerRow: true,
    style: {
      theme: 'TableStyleMedium9',
    },
    columns: columns.map(column => ({ name: column.header })),
    rows: data,
  });

  // Send the generated Excel file to the client
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename=myFile.xlsx');
  workbook.xlsx.write(res).then(() => {
    res.end();
  });
});

app.listen(port, () => {
  console.log(`Server is running at http://localhost:${port}`);
});
