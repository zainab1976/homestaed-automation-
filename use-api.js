// Example: How to use your Medical Assessment API

const API_BASE = 'http://localhost:3000';

async function useAPI() {
  console.log('🚀 Using Medical Assessment API...\n');

  try {
    // 1. Check API health
    console.log('1. Checking API health...');
    const healthResponse = await fetch(`${API_BASE}/health`);
    const healthData = await healthResponse.json();
    console.log('✅ API Status:', healthData.status);
    console.log('📅 Timestamp:', healthData.timestamp);

    // 2. List available files
    console.log('\n2. Listing available files...');
    const filesResponse = await fetch(`${API_BASE}/api/files`);
    const filesData = await filesResponse.json();
    console.log('📁 Available files:');
    filesData.files.forEach(file => {
      console.log(`   - ${file.name} (${file.type})`);
    });

    // 3. Read Excel file data
    console.log('\n3. Reading Excel file data...');
    const excelResponse = await fetch(`${API_BASE}/api/excel/FMG%2009.22.2025.xlsx`);
    const excelData = await excelResponse.json();
    
    if (excelData.success) {
      console.log('📊 Excel file sheets:');
      Object.keys(excelData.data).forEach(sheetName => {
        const rowCount = excelData.data[sheetName].length;
        console.log(`   - ${sheetName}: ${rowCount} rows`);
      });
    }

    // 4. Process assessments (uncomment to run automation)
    console.log('\n4. Ready to process assessments...');
    console.log('⚠️  To start processing, uncomment the code below:');
    console.log(`
    const processResponse = await fetch(\`\${API_BASE}/api/process/FMG%2009.22.2025.xlsx\`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ headless: true, slowMo: 2000 })
    });
    const processData = await processResponse.json();
    console.log('🔄 Processing started:', processData);
    `);

  } catch (error) {
    console.error('❌ Error:', error.message);
  }
}

// Run the example
useAPI();




