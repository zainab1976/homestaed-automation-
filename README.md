# Playwright Setup

# Medical Assessment Automation API

This project provides a REST API wrapper around Playwright automation for processing medical assessments. The API allows you to upload Excel files, read patient data, and automatically process assessments through a web interface.

## Features

- **Excel File Management**: Upload, read, and update Excel files containing patient data
- **Automated Assessment Processing**: Process GAD16 and Health Assessment patients automatically
- **RESTful API**: Easy-to-use HTTP endpoints for integration
- **Real-time Updates**: Update Excel files with processing status in real-time
- **Error Handling**: Comprehensive error handling and logging

## Quick Start

### 1. Install Dependencies
```bash
npm install
```

### 2. Set Up Environment Variables
Create a `.env` file in the project root:
```
QHSLAB_EMAIL=your-email@example.com
QHSLAB_PASSWORD=your-password
PORT=3000
```

### 3. Start the API Server
```bash
# Using npm
npm start

# Or using the provided scripts
# Windows Command Prompt
start-api.bat

# Windows PowerShell
.\start-api.ps1
```

### 4. Test the API
```bash
# Test health endpoint
curl http://localhost:3000/health

# List available files
curl http://localhost:3000/api/files
```

## API Endpoints

| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/health` | Health check |
| GET | `/api/files` | List available Excel files |
| POST | `/api/upload` | Upload Excel file |
| GET | `/api/excel/:filename` | Read Excel file data |
| POST | `/api/excel/:filename/update` | Update Excel file |
| POST | `/api/process/:filename` | Process assessments |
| GET | `/api/status/:jobId` | Get processing status |

## Usage Examples

### Upload and Process an Excel File

1. **Upload the file:**
```bash
curl -X POST -F "excelFile=@patients.xlsx" http://localhost:3000/api/upload
```

2. **Process assessments:**
```bash
curl -X POST http://localhost:3000/api/process/patients.xlsx \
  -H "Content-Type: application/json" \
  -d '{"headless": true, "slowMo": 2000}'
```

### Using JavaScript/Fetch

```javascript
// Upload file
const formData = new FormData();
formData.append('excelFile', fileInput.files[0]);

const uploadResponse = await fetch('http://localhost:3000/api/upload', {
  method: 'POST',
  body: formData
});

const uploadData = await uploadResponse.json();
console.log('File uploaded:', uploadData.file.name);

// Process assessments
const processResponse = await fetch(`http://localhost:3000/api/process/${uploadData.file.name}`, {
  method: 'POST',
  headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify({ headless: true, slowMo: 2000 })
});

const processData = await processResponse.json();
console.log('Processing started:', processData);
```

## Excel File Format

The API expects Excel files with the following sheets:
- **GAD 16**: Patients requiring PHQ-GAD16 assessments
- **Health assessment**: Patients requiring Health Assessments

Required columns:
- `DOB`: Patient Date of Birth
- `Appointment Facility Name`: Account/Facility name
- `Custom ID`: Custom identifier for the account
- `Appointment Provider Name` or `Scheduler`: Provider name
- `Primary Insurance Name`: Insurance information

## Configuration

The API uses the following configuration (can be overridden via environment variables):

- `QHSLAB_EMAIL`: Login email for the web application
- `QHSLAB_PASSWORD`: Login password for the web application
- `PORT`: API server port (default: 3000)

## Error Handling

All endpoints return JSON responses with a `success` field indicating success/failure:

```json
{
  "success": true,
  "data": { ... }
}
```

Error responses:
```json
{
  "success": false,
  "error": "Error message"
}
```

## Development

### Project Structure
```
├── server.js                 # Main API server
├── assessment-processor.js   # Playwright automation logic
├── excel-helper.js          # Excel file operations
├── index.js                 # Original automation script
├── clients/                 # Default Excel files directory
├── uploads/                 # Uploaded files directory
└── API_DOCUMENTATION.md     # Detailed API documentation
```

### Testing
```bash
# Run the test script
node test-api.js
```

## Troubleshooting

1. **Port already in use**: Change the PORT in your `.env` file
2. **Login failed**: Verify your QHSLAB_EMAIL and QHSLAB_PASSWORD
3. **File not found**: Ensure Excel files are in the correct directory
4. **Browser issues**: Check if Playwright browsers are installed: `npx playwright install`

## License

ISC

## Prerequisites

- Node.js (version 14 or higher)
- npm or yarn

## Installation

1. Install dependencies:
```bash
npm install
```

2. Install Playwright browsers:
```bash
npm run install:browsers
```

## Usage

### Running the basic example

Run the main example script:
```bash
node index.js
```

This will:
- Launch a Chromium browser
- Navigate to Google
- Take a screenshot
- Perform a search
- Take another screenshot of results
- Close the browser

### Running tests

Run all tests:
```bash
npm test
```

Run tests in headed mode (visible browser):
```bash
npm run test:headed
```

Run tests with UI mode:
```bash
npm run test:ui
```

Debug tests:
```bash
npm run test:debug
```

## Project Structure

- `index.js` - Main example script with basic Playwright automation
- `playwright.config.js` - Playwright configuration
- `package.json` - Project dependencies and scripts
- `tests/` - Directory for test files (create this if you want to add tests)

## Features

- ✅ Basic browser automation
- ✅ Screenshot capture
- ✅ Form interaction
- ✅ Multi-browser support (Chromium, Firefox, WebKit)
- ✅ Test configuration
- ✅ Example scripts

## Customization

You can modify `index.js` to:
- Change the target website
- Add more automation steps
- Customize browser settings
- Add error handling

## Documentation

For more information, visit the [Playwright documentation](https://playwright.dev/).
