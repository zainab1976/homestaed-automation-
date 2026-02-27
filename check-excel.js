const XLSX = require('xlsx');
const path = require('path');

async function checkExcelFile() {
  try {
    const filePath = path.resolve(__dirname, 'clients', 'FMG 09.22.2025.xlsx');
    console.log(`📖 Reading Excel file: ${filePath}`);
    
    // Read the workbook
    const workbook = XLSX.readFile(filePath);
    
    console.log('\n📋 Available sheets:');
    workbook.SheetNames.forEach((sheetName, index) => {
      console.log(`  ${index + 1}. "${sheetName}"`);
    });
    
    // Check each sheet
    workbook.SheetNames.forEach(sheetName => {
      console.log(`\n📄 Sheet: "${sheetName}"`);
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet);
      
      console.log(`  Rows: ${jsonData.length}`);
      
      if (jsonData.length > 0) {
        console.log('  Columns:', Object.keys(jsonData[0]));
        console.log('  Sample row:', jsonData[0]);
        
        // Check for MRN/Chart columns
        const mrnColumns = Object.keys(jsonData[0]).filter(key => 
          key.toLowerCase().includes('mrn') || 
          key.toLowerCase().includes('chart') ||
          key.toLowerCase().includes('patient')
        );
        console.log('  MRN/Chart columns:', mrnColumns);
        
        // Check for insurance columns
        const insuranceColumns = Object.keys(jsonData[0]).filter(key => 
          key.toLowerCase().includes('insurance')
        );
        console.log('  Insurance columns:', insuranceColumns);
        
        // Check for scheduler columns
        const schedulerColumns = Object.keys(jsonData[0]).filter(key => 
          key.toLowerCase().includes('scheduler') || 
          key.toLowerCase().includes('provider')
        );
        console.log('  Scheduler columns:', schedulerColumns);
      }
    });
    
  } catch (error) {
    console.error('❌ Error reading Excel file:', error.message);
  }
}

checkExcelFile();
