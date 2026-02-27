# Medical Assessment Automation API

This API provides endpoints to manage Excel files and process medical assessments using Playwright automation.

## Base URL
```
http://localhost:3000
```

## Endpoints

### 1. Health Check
**GET** `/health`

Check if the API is running.

**Response:**
```json
{
  "status": "OK",
  "message": "Medical Assessment Automation API is running",
  "timestamp": "2024-01-15T10:30:00.000Z"
}
```

### 2. List Available Excel Files
**GET** `/api/files`

Get a list of all available Excel files in the system.

**Response:**
```json
{
  "success": true,
  "files": [
    {
      "name": "FMG 09.22.2025.xlsx",
      "path": "/path/to/clients/FMG 09.22.2025.xlsx",
      "type": "client"
    }
  ]
}
```

### 3. Upload Excel File
**POST** `/api/upload`

Upload a new Excel file to the system.

**Request:**
- Content-Type: `multipart/form-data`
- Body: `excelFile` (file)

**Response:**
```json
{
  "success": true,
  "message": "File uploaded successfully",
  "file": {
    "name": "1234567890-newfile.xlsx",
    "originalName": "newfile.xlsx",
    "path": "/path/to/uploads/1234567890-newfile.xlsx",
    "size": 12345
  }
}
```

### 4. Read Excel File Data
**GET** `/api/excel/:filename`

Read and parse an Excel file, returning its data organized by sheets.

**Parameters:**
- `filename` (string): Name of the Excel file

**Response:**
```json
{
  "success": true,
  "data": {
    "GAD 16": [
      {
        "DOB": "01/15/1980",
        "Patient Name": "John Doe",
        "Appointment Provider Name": "Dr. Smith",
        "Primary Insurance Name": "AvMed"
      }
    ],
    "Health assessment": [
      {
        "DOB": "03/22/1975",
        "Patient Name": "Jane Smith",
        "Appointment Provider Name": "Dr. Johnson",
        "Primary Insurance Name": "BlueCross"
      }
    ]
  },
  "file": {
    "name": "FMG 09.22.2025.xlsx",
    "path": "/path/to/file"
  }
}
```

### 5. Update Excel File
**POST** `/api/excel/:filename/update`

Update a specific cell in an Excel file.

**Parameters:**
- `filename` (string): Name of the Excel file

**Request Body:**
```json
{
  "sheetName": "GAD 16",
  "searchColumn": "DOB",
  "searchValue": "12345",
  "updateValue": "Sent",
  "updateColumn": "Status"
}
```

**Response:**
```json
{
  "success": true,
  "message": "Excel file updated successfully"
}
```

### 6. Process Assessments
**POST** `/api/process/:filename`

Start the automated assessment processing for all patients in the specified Excel file.

**Parameters:**
- `filename` (string): Name of the Excel file to process

**Request Body (optional):**
```json
{
  "headless": true,
  "slowMo": 2000
}
```

**Response:**
```json
{
  "success": true,
  "message": "Assessment processing started",
  "file": "FMG 09.22.2025.xlsx",
  "config": {
    "headless": true,
    "slowMo": 2000
  }
}
```

### 7. Get Processing Status
**GET** `/api/status/:jobId`

Get the status of a processing job (placeholder for future implementation).

**Parameters:**
- `jobId` (string): Job identifier

**Response:**
```json
{
  "success": true,
  "message": "Status endpoint - to be implemented",
  "jobId": "job123"
}
```

## Usage Examples

### Using curl

1. **Check API health:**
```bash
curl http://localhost:3000/health
```

2. **List available files:**
```bash
curl http://localhost:3000/api/files
```

3. **Upload an Excel file:**
```bash
curl -X POST -F "excelFile=@/path/to/your/file.xlsx" http://localhost:3000/api/upload
```

4. **Read Excel file data:**
```bash
curl http://localhost:3000/api/excel/FMG%2009.22.2025.xlsx
```

5. **Update Excel file:**
```bash
curl -X POST http://localhost:3000/api/excel/FMG%2009.22.2025.xlsx/update \
  -H "Content-Type: application/json" \
  -d '{
    "sheetName": "GAD 16",
    "searchColumn": "DOB",
    "searchValue": "12345",
    "updateValue": "Sent",
    "updateColumn": "Status"
  }'
```

6. **Process assessments:**
```bash
curl -X POST http://localhost:3000/api/process/FMG%2009.22.2025.xlsx \
  -H "Content-Type: application/json" \
  -d '{
    "headless": true,
    "slowMo": 2000
  }'
```

### Using JavaScript/Fetch

```javascript
// Upload a file
const formData = new FormData();
formData.append('excelFile', fileInput.files[0]);

fetch('http://localhost:3000/api/upload', {
  method: 'POST',
  body: formData
})
.then(response => response.json())
.then(data => console.log(data));

// Process assessments
fetch('http://localhost:3000/api/process/FMG%2009.22.2025.xlsx', {
  method: 'POST',
  headers: {
    'Content-Type': 'application/json'
  },
  body: JSON.stringify({
    headless: true,
    slowMo: 2000
  })
})
.then(response => response.json())
.then(data => console.log(data));
```

## Error Responses

All endpoints return error responses in the following format:

```json
{
  "success": false,
  "error": "Error message describing what went wrong"
}
```

Common HTTP status codes:
- `200` - Success
- `400` - Bad Request (missing or invalid parameters)
- `404` - Not Found (file or endpoint not found)
- `500` - Internal Server Error

## Environment Variables

Make sure to set these environment variables in your `.env` file:

```
QHSLAB_EMAIL=your-email@example.com
QHSLAB_PASSWORD=your-password
PORT=3000
```

## Starting the API

1. Install dependencies:
```bash
npm install
```

2. Start the server:
```bash
npm start
```

The API will be available at `http://localhost:3000`
