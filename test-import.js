const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

// Create a mock trial balance file
const wb = XLSX.utils.book_new();
const wsData = [
  ['Ledger Code', 'Ledger Name', 'Opening', 'Debit', 'Credit', 'Closing'],
  ['1000', 'Cash', '100', '50', '20', '130'],
  ['2000', 'Accounts Payable', '-50', '20', '60', '-90']
];
const ws = XLSX.utils.aoa_to_sheet(wsData);
XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
XLSX.writeFile(wb, 'mock_tb.xlsx');
console.log('Created mock_tb.xlsx');
