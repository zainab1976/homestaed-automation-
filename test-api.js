const http = require('http');

// Test the API endpoints
async function testAPI() {
  console.log('🧪 Testing Medical Assessment Automation API...\n');
  
  // Test health endpoint
  try {
    console.log('1. Testing health endpoint...');
    const healthResponse = await makeRequest('GET', '/health');
    console.log('✅ Health check:', healthResponse);
  } catch (error) {
    console.log('❌ Health check failed:', error.message);
  }
  
  // Test files endpoint
  try {
    console.log('\n2. Testing files endpoint...');
    const filesResponse = await makeRequest('GET', '/api/files');
    console.log('✅ Files list:', filesResponse);
  } catch (error) {
    console.log('❌ Files list failed:', error.message);
  }
  
  console.log('\n🎉 API testing completed!');
}

function makeRequest(method, path, data = null) {
  return new Promise((resolve, reject) => {
    const options = {
      hostname: 'localhost',
      port: 3000,
      path: path,
      method: method,
      headers: {
        'Content-Type': 'application/json'
      }
    };
    
    const req = http.request(options, (res) => {
      let body = '';
      res.on('data', (chunk) => {
        body += chunk;
      });
      res.on('end', () => {
        try {
          const jsonBody = JSON.parse(body);
          resolve(jsonBody);
        } catch (e) {
          resolve(body);
        }
      });
    });
    
    req.on('error', (error) => {
      reject(error);
    });
    
    if (data) {
      req.write(JSON.stringify(data));
    }
    
    req.end();
  });
}

// Run the test
testAPI().catch(console.error);



