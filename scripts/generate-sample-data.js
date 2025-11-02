// Script to generate a sample Excel file for testing SheetGrid
const XLSX = require('xlsx');

// Create workbook
const wb = XLSX.utils.book_new();

// Create sample data
const sampleData = [
  ['Product', 'Price', 'Quantity', 'Total'],
  ['Apple', 1.50, 10, '=B2*C2'],
  ['Banana', 0.75, 15, '=B3*C3'],
  ['Orange', 2.00, 8, '=B4*C4'],
  ['Grapes', 3.50, 5, '=B5*C5'],
  ['', '', 'Total', '=SUM(D2:D5)'],
  ['Revenue', 250, '', ''],
  ['Costs', 180, '', ''],
  ['Profit', '', '', '=B7-B8'],
];

// Convert to worksheet
const ws = XLSX.utils.aoa_to_sheet(sampleData);

// Add some styling (column widths)
ws['!cols'] = [
  { wch: 15 },  // Product
  { wch: 10 },  // Price
  { wch: 10 },  // Quantity
  { wch: 10 },  // Total
];

// Add worksheet to workbook
XLSX.utils.book_append_sheet(wb, ws, 'Sales');

// Write file
XLSX.writeFile(wb, 'examples/sample-data.xlsx');
console.log('Sample Excel file created: examples/sample-data.xlsx');

