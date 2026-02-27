const XLSX = require('xlsx');

// Create sample data for testing
const gad16Data = [
  { MRN: '12345', Scheduler: 'REGINALD JEROME APRN', 'Primary Insurance Name': 'AvMed' },
  { MRN: '12346', Scheduler: 'REGINALD JEROME APRN', 'Primary Insurance Name': 'AvMed' }
];

const healthData = [
  { MRN: '12347', Scheduler: 'REGINALD JEROME APRN', 'Primary Insurance Name': 'AvMed' },
  { MRN: '12348', Scheduler: 'REGINALD JEROME APRN', 'Primary Insurance Name': 'AvMed' }
];

const appointmentData = [
  { 'Chart  #': '12349', 'Appointment Date': new Date(), Scheduler: 'REGINALD JEROME APRN', 'Primary Insurance Name': 'AvMed' },
  { 'Chart  #': '12350', 'Appointment Date': new Date(), Scheduler: 'REGINALD JEROME APRN', 'Primary Insurance Name': 'AvMed' }
];

// Create workbook
const workbook = XLSX.utils.book_new();

// Add sheets
XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(gad16Data), 'GAD 16');
XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(healthData), 'Health assessment');
XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(appointmentData), 'Appointment_Report_20106_638937');

// Write file
XLSX.writeFile(workbook, 'clients/09.19.2025.xlsx');
console.log('✅ Sample Excel file created: clients/09.19.2025.xlsx');
