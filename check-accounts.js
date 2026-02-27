const XLSX = require('xlsx');
const path = require('path');

async function checkAccountNames() {
  try {
    const filePath = path.resolve(__dirname, 'clients', 'Estrella Medical Servicess 09.19.2025 1.xlsx');
    console.log(`📖 Reading Excel file: ${filePath}`);
    
    const workbook = XLSX.readFile(filePath);
    
    // Check all sheets for account/facility names
    workbook.SheetNames.forEach(sheetName => {
      console.log(`\n📄 Sheet: "${sheetName}"`);
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet);
      
      if (jsonData.length > 0) {
        // Look for account/facility name columns
        const accountColumns = Object.keys(jsonData[0]).filter(key => 
          key.toLowerCase().includes('facility') || 
          key.toLowerCase().includes('account') ||
          key.toLowerCase().includes('practice')
        );
        
        console.log('  Account/Facility columns:', accountColumns);
        
        if (accountColumns.length > 0) {
          // Get unique account names
          const uniqueAccounts = [...new Set(jsonData.map(row => row[accountColumns[0]]).filter(Boolean))];
          console.log(`  Unique account names (${uniqueAccounts.length}):`);
          uniqueAccounts.forEach((account, index) => {
            console.log(`    ${index + 1}. "${account}"`);
          });
        }
      }
    });
    
  } catch (error) {
    console.error('❌ Error reading Excel file:', error.message);
  }
}

checkAccountNames();
