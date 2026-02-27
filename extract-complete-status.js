const XLSX = require('xlsx');
const path = require('path');

// Complete status data extracted from the FULL terminal output
const allProcessedPatients = [
  // From the terminal output - all patients that were processed
  { mrn: "13775", status: "Sent", lastOrderDate: "10/07/2025" },
  { mrn: "17156", status: "need to add demo", lastOrderDate: "" },
  { mrn: "13238", status: "Sent", lastOrderDate: "10/07/2025" },
  { mrn: "17702", status: "Sent", lastOrderDate: "10/07/2025" },
  { mrn: "15243", status: "Sent", lastOrderDate: "10/07/2025" },
  { mrn: "13320", status: "Sent", lastOrderDate: "10/07/2025" },
  { mrn: "15805", status: "Sent", lastOrderDate: "10/07/2025" },
  { mrn: "13200", status: "Already", lastOrderDate: "" },
  { mrn: "VIP13267", status: "need to add demo", lastOrderDate: "" },
  { mrn: "12706", status: "Sent", lastOrderDate: "10/07/2025" },
  { mrn: "14458", status: "Unable", lastOrderDate: "" },
  { mrn: "17475", status: "Already", lastOrderDate: "" }
];

async function extractCompleteStatus() {
  try {
    const excelPath = path.join(__dirname, 'clients', 'FMG 09.22.2025.xlsx');
    
    console.log('🔍 Extracting complete status from terminal output...');
    console.log(`📊 Found ${allProcessedPatients.length} patients with status data`);
    
    // Read the workbook
    const workbook = XLSX.readFile(excelPath);
    const worksheet = workbook.Sheets['Health assessment'];
    
    if (!worksheet) {
      throw new Error('Health assessment sheet not found');
    }
    
    // Convert to JSON
    const jsonData = XLSX.utils.sheet_to_json(worksheet);
    console.log(`📊 Total patients in Excel: ${jsonData.length}`);
    
    // Update ALL patients with correct status
    const updatedData = jsonData.map(row => {
      const mrn = String(row.MRN || '').trim();
      const statusInfo = allProcessedPatients.find(s => s.mrn === mrn);
      
      if (statusInfo) {
        // This patient was processed
        return {
          ...row,
          'Status': statusInfo.status,
          'Last Order Date': statusInfo.lastOrderDate || ''
        };
      } else {
        // This patient wasn't in the processed list - mark as not processed
        return {
          ...row,
          'Status': 'Not Processed',
          'Last Order Date': ''
        };
      }
    });
    
    // Create new worksheet with updated data
    const newWorksheet = XLSX.utils.json_to_sheet(updatedData);
    
    // Set column widths for better visibility
    const colWidths = [
      { wch: 10 }, // Custom ID
      { wch: 30 }, // Appointment Facility Name
      { wch: 25 }, // Appointment Provider Name
      { wch: 15 }, // MRN
      { wch: 25 }, // Patient Name
      { wch: 25 }, // Appointment Insurance Name
      { wch: 15 }, // Status
      { wch: 15 }  // Last Order Date
    ];
    newWorksheet['!cols'] = colWidths;
    
    // Update the workbook
    workbook.Sheets['Health assessment'] = newWorksheet;
    
    // Write back to file
    XLSX.writeFile(workbook, excelPath);
    
    console.log('✅ Complete status update completed!');
    
    // Verify the update
    console.log('\n🔍 Verifying updates...');
    const updatedWorkbook = XLSX.readFile(excelPath);
    const updatedWorksheet = updatedWorkbook.Sheets['Health assessment'];
    const verifyData = XLSX.utils.sheet_to_json(updatedWorksheet);
    
    // Count status types
    const statusCounts = {};
    verifyData.forEach(patient => {
      const status = patient.Status || 'Unknown';
      statusCounts[status] = (statusCounts[status] || 0) + 1;
    });
    
    console.log('\n📊 Final Status Summary:');
    Object.keys(statusCounts).forEach(status => {
      console.log(`   ${status}: ${statusCounts[status]} patients`);
    });
    
    console.log('\n✅ Processed patients:');
    allProcessedPatients.forEach(patient => {
      const found = verifyData.find(p => String(p.MRN || '').trim() === patient.mrn);
      if (found) {
        console.log(`   MRN: ${patient.mrn} - Status: "${found.Status}" - Last Order Date: "${found['Last Order Date'] || 'EMPTY'}"`);
      }
    });
    
    console.log('\n🎉 Excel file updated with complete status data!');
    
  } catch (error) {
    console.error('❌ Error extracting complete status:', error.message);
  }
}

extractCompleteStatus();

