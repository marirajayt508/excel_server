const XLSX = require('xlsx');
const fs = require('fs');

// Create a new workbook
const workbook = XLSX.utils.book_new();

// Create a worksheet with data
const worksheet = XLSX.utils.aoa_to_sheet([
  ['Name', 'Age', 'Gender'],
  ['John', 25, 'Male'],
  ['Jane', 30, 'Female'],
]);

// Add a drop-down list to a cell (B2 in this example)
const dropDownList = ['Male', 'Female'];
const dataValidation = {
  type: 'list',
  formula1: dropDownList.join(','),
  allowBlank: true,
  showDropDown: true,
  showInputMessage: true,
  sqref: 'B2',
};

worksheet['!dataValidation'] = [dataValidation];

// Add the worksheet to the workbook
XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

// Save the workbook to a file
const excelFilePath = 'output.xlsx';
XLSX.writeFile(workbook, excelFilePath);

console.log(`Excel file with drop-down list created: ${excelFilePath}`);
