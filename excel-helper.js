const XLSX = require('xlsx');
const path = require('path');

/**
 * Read Excel file and return data organized by sheet names
 * @param {string} filePath - Path to the Excel file
 * @returns {Object} Object with sheet names as keys and data arrays as values
 */
async function readExcel(filePath) {
  try {
    console.log(`📖 Reading Excel file: ${filePath}`);
    
    // Check if file exists
    const fs = require('fs');
    if (!fs.existsSync(filePath)) {
      throw new Error(`Excel file not found: ${filePath}`);
    }
    
    // Read the workbook
    const workbook = XLSX.readFile(filePath);
    const result = {};
    
    // Process each sheet
    workbook.SheetNames.forEach(sheetName => {
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet);
      result[sheetName] = jsonData;
      console.log(`📄 Loaded sheet "${sheetName}" with ${jsonData.length} rows`);
    });
    
    return result;
  } catch (error) {
    console.error('❌ Error reading Excel file:', error.message);
    throw error;
  }
}

/**
 * Mark/update a specific cell in Excel file
 * @param {string} filePath - Path to the Excel file
 * @param {string} sheetName - Name of the sheet
 * @param {string} searchColumn - Column name to search for the row
 * @param {string} searchValue - Value to search for in the search column
 * @param {string} updateValue - Value to update
 * @param {string} updateColumn - Column name to update (optional, defaults to 'Status')
 */
async function markExcel(filePath, sheetName, searchColumn, searchValue, updateValue, updateColumn = 'Status') {
  try {
    console.log(`\n📝 ===== EXCEL UPDATE START =====`);
    console.log(`📁 File: ${filePath}`);
    console.log(`📄 Sheet: ${sheetName}`);
    console.log(`🔍 Search Column: ${searchColumn}`);
    console.log(`🔍 Search Value: "${searchValue}"`);
    console.log(`✏️ Update Column: ${updateColumn}`);
    console.log(`✏️ Update Value: "${updateValue}"`);
    
    // Check if file exists
    const fs = require('fs');
    if (!fs.existsSync(filePath)) {
      console.error(`❌ Excel file not found: ${filePath}`);
      throw new Error(`Excel file not found: ${filePath}`);
    }
    
    // Check file permissions
    try {
      fs.accessSync(filePath, fs.constants.W_OK);
      console.log(`✅ File is writable`);
    } catch (permError) {
      console.error(`❌ File is not writable or locked: ${permError.message}`);
      console.error(`💡 Make sure Excel file is closed before running the script`);
      throw new Error(`File is not writable. Please close the Excel file and try again.`);
    }
    
    // Read the workbook
    const workbook = XLSX.readFile(filePath);
    const worksheet = workbook.Sheets[sheetName];
    
    if (!worksheet) {
      const availableSheets = workbook.SheetNames.join(', ');
      throw new Error(`Sheet "${sheetName}" not found. Available sheets: ${availableSheets}`);
    }
    
    // Convert to JSON to work with data
    const jsonData = XLSX.utils.sheet_to_json(worksheet);
    
    if (jsonData.length === 0) {
      console.log(`⚠️ Sheet "${sheetName}" is empty`);
      return false;
    }
    
    // Show available columns for debugging
    if (jsonData.length > 0) {
      const availableColumns = Object.keys(jsonData[0]);
      console.log(`📊 Available columns in sheet: ${availableColumns.join(', ')}`);
      console.log(`📊 Total rows in sheet: ${jsonData.length}`);
      
      // Check if search column exists
      if (!availableColumns.includes(searchColumn)) {
        console.error(`❌ Column "${searchColumn}" not found!`);
        console.error(`   Available columns: ${availableColumns.join(', ')}`);
        return false;
      }
      
      // Check if update column exists, if not, we'll add it
      if (!availableColumns.includes(updateColumn)) {
        console.log(`⚠️ Column "${updateColumn}" not found. Will be added to the sheet.`);
      }
    }
    
    // Find the row to update - try multiple matching strategies
    let rowIndex = -1;
    const searchValueStr = String(searchValue).trim();
    if (!searchValueStr) {
      console.error(`❌ Empty search value for column "${searchColumn}". Refusing to update to avoid wrong row.`);
      console.error(`   Tip: Ensure the identifier (e.g., Custom ID or MRN) is present in Excel.`);
      console.error(`   ===== EXCEL UPDATE FAILED =====\n`);
      return false;
    }
    console.log(`🔍 Searching for row with ${searchColumn} = "${searchValueStr}"`);
    
    // Strategy 1: Exact match (case-insensitive)
    rowIndex = jsonData.findIndex(row => {
      const cellValue = String(row[searchColumn] || '').trim();
      return cellValue.toLowerCase() === searchValueStr.toLowerCase();
    });
    
    // Strategy 2: Try matching just the date part if it's a date
    if (rowIndex === -1 && searchValueStr.includes('/')) {
      const dateParts = searchValueStr.split('/');
      if (dateParts.length === 3) {
        rowIndex = jsonData.findIndex(row => {
          const cellValue = String(row[searchColumn] || '').trim();
          if (cellValue.includes('/')) {
            const cellParts = cellValue.split('/');
            // Match MM/DD/YYYY or DD/MM/YYYY
            return (cellParts.length === 3 && 
                   ((cellParts[0] === dateParts[0] && cellParts[1] === dateParts[1] && cellParts[2] === dateParts[2]) ||
                    (cellParts[1] === dateParts[0] && cellParts[0] === dateParts[1] && cellParts[2] === dateParts[2])));
          }
          return false;
        });
      }
    }
    
    // Strategy 3: Partial match (contains)
    if (rowIndex === -1) {
      rowIndex = jsonData.findIndex(row => {
        const cellValue = String(row[searchColumn] || '').trim();
        return cellValue.includes(searchValueStr) || searchValueStr.includes(cellValue);
      });
    }
    
    if (rowIndex === -1) {
      // Log available values for debugging
      const availableValues = jsonData.slice(0, 10).map((row, idx) => {
        const val = String(row[searchColumn] || '');
        return `Row ${idx + 2}: "${val}"`;
      }).join('\n   ');
      console.error(`❌ Row not found for ${searchColumn}: "${searchValue}"`);
      console.error(`   Available ${searchColumn} values (first 10 rows):`);
      console.error(`   ${availableValues}`);
      console.error(`   Total rows in sheet: ${jsonData.length}`);
      console.error(`   ===== EXCEL UPDATE FAILED =====\n`);
      return false;
    }
    
    console.log(`✅ Row found at index: ${rowIndex + 2} (actual row number)`);
    console.log(`   Current value in ${updateColumn}: "${jsonData[rowIndex][updateColumn] || '(empty)'}"`);
    
    // Update the value
    jsonData[rowIndex][updateColumn] = updateValue;
    
    // Convert back to worksheet - preserve existing structure
    const newWorksheet = XLSX.utils.json_to_sheet(jsonData, {
      header: Object.keys(jsonData[0]),
      skipHeader: false
    });
    
    // Preserve column widths if possible
    if (worksheet['!cols']) {
      newWorksheet['!cols'] = worksheet['!cols'];
    }
    
    // Preserve merge cells if any
    if (worksheet['!merges']) {
      newWorksheet['!merges'] = worksheet['!merges'];
    }
    
    workbook.Sheets[sheetName] = newWorksheet;
    
    // Write back to file
    try {
      XLSX.writeFile(workbook, filePath);
      console.log(`✅ Successfully wrote to file: ${filePath}`);
    } catch (writeError) {
      console.error(`❌ Error writing to file: ${writeError.message}`);
      throw writeError;
    }
    
    console.log(`✅ Successfully updated Excel:`);
    console.log(`   Sheet: ${sheetName}`);
    console.log(`   Row: ${rowIndex + 2}`);
    console.log(`   Column: ${updateColumn}`);
    console.log(`   New Value: "${updateValue}"`);
    console.log(`📝 ===== EXCEL UPDATE SUCCESS =====\n`);
    return true;
    
  } catch (error) {
    console.error(`\n❌ ===== EXCEL UPDATE ERROR =====`);
    console.error(`   File: ${filePath}`);
    console.error(`   Sheet: ${sheetName}`);
    console.error(`   Search Column: ${searchColumn}`);
    console.error(`   Search Value: "${searchValue}"`);
    console.error(`   Update Column: ${updateColumn}`);
    console.error(`   Update Value: "${updateValue}"`);
    console.error(`   Error: ${error.message}`);
    if (error.stack) {
      console.error(`   Stack: ${error.stack}`);
    }
    console.error(`❌ ===== EXCEL UPDATE FAILED =====\n`);
    return false;
  }
}

module.exports = {
  readExcel,
  markExcel
};
