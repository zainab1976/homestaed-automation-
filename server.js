const express = require('express');
const cors = require('cors');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const { readExcel, markExcel } = require('./excel-helper');
const { processAllPatients, loadExcelData } = require('./assessment-processor');

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Configure multer for file uploads
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    const uploadDir = path.join(__dirname, 'uploads');
    if (!fs.existsSync(uploadDir)) {
      fs.mkdirSync(uploadDir, { recursive: true });
    }
    cb(null, uploadDir);
  },
  filename: (req, file, cb) => {
    cb(null, Date.now() + '-' + file.originalname);
  }
});

const upload = multer({ 
  storage: storage,
  fileFilter: (req, file, cb) => {
    if (file.mimetype === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || 
        file.mimetype === 'application/vnd.ms-excel') {
      cb(null, true);
    } else {
      cb(new Error('Only Excel files are allowed!'), false);
    }
  }
});

// Health check endpoint
app.get('/health', (req, res) => {
  res.json({ 
    status: 'OK', 
    message: 'Medical Assessment Automation API is running',
    timestamp: new Date().toISOString()
  });
});

// Get available Excel files
app.get('/api/files', (req, res) => {
  try {
    const clientsDir = path.join(__dirname, 'clients');
    const uploadsDir = path.join(__dirname, 'uploads');
    
    let files = [];
    
    // Check clients directory
    if (fs.existsSync(clientsDir)) {
      const clientFiles = fs.readdirSync(clientsDir)
        .filter(file => file.endsWith('.xlsx') || file.endsWith('.xls'))
        .map(file => ({
          name: file,
          path: path.join(clientsDir, file),
          type: 'client'
        }));
      files = files.concat(clientFiles);
    }
    
    // Check uploads directory
    if (fs.existsSync(uploadsDir)) {
      const uploadFiles = fs.readdirSync(uploadsDir)
        .filter(file => file.endsWith('.xlsx') || file.endsWith('.xls'))
        .map(file => ({
          name: file,
          path: path.join(uploadsDir, file),
          type: 'upload'
        }));
      files = files.concat(uploadFiles);
    }
    
    res.json({
      success: true,
      files: files
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// Upload Excel file
app.post('/api/upload', upload.single('excelFile'), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({
        success: false,
        error: 'No file uploaded'
      });
    }
    
    res.json({
      success: true,
      message: 'File uploaded successfully',
      file: {
        name: req.file.filename,
        originalName: req.file.originalname,
        path: req.file.path,
        size: req.file.size
      }
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// Read Excel file data
app.get('/api/excel/:filename', async (req, res) => {
  try {
    const { filename } = req.params;
    const filePath = path.join(__dirname, 'uploads', filename);
    
    // Check if file exists in uploads, otherwise check clients
    let actualPath = filePath;
    if (!fs.existsSync(filePath)) {
      actualPath = path.join(__dirname, 'clients', filename);
      if (!fs.existsSync(actualPath)) {
        return res.status(404).json({
          success: false,
          error: 'File not found'
        });
      }
    }
    
    const data = await readExcel(actualPath);
    
    res.json({
      success: true,
      data: data,
      file: {
        name: filename,
        path: actualPath
      }
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// Update Excel file
app.post('/api/excel/:filename/update', async (req, res) => {
  try {
    const { filename } = req.params;
    const { sheetName, searchColumn, searchValue, updateValue, updateColumn } = req.body;
    
    if (!sheetName || !searchColumn || !searchValue || !updateValue) {
      return res.status(400).json({
        success: false,
        error: 'Missing required fields: sheetName, searchColumn, searchValue, updateValue'
      });
    }
    
    const filePath = path.join(__dirname, 'uploads', filename);
    let actualPath = filePath;
    
    if (!fs.existsSync(filePath)) {
      actualPath = path.join(__dirname, 'clients', filename);
      if (!fs.existsSync(actualPath)) {
        return res.status(404).json({
          success: false,
          error: 'File not found'
        });
      }
    }
    
    await markExcel(actualPath, sheetName, searchColumn, searchValue, updateValue, updateColumn);
    
    res.json({
      success: true,
      message: 'Excel file updated successfully'
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// Process assessments for a specific Excel file
app.post('/api/process/:filename', async (req, res) => {
  try {
    const { filename } = req.params;
    const { headless = true, slowMo = 2000 } = req.body;
    
    const filePath = path.join(__dirname, 'uploads', filename);
    let actualPath = filePath;
    
    if (!fs.existsSync(filePath)) {
      actualPath = path.join(__dirname, 'clients', filename);
      if (!fs.existsSync(actualPath)) {
        return res.status(404).json({
          success: false,
          error: 'File not found'
        });
      }
    }
    
    // Update the config to use the specified file
    process.env.CLIENT_FILE_PATH = actualPath;
    
    // Start processing in background
    processAllPatients(actualPath, { headless, slowMo })
      .then(result => {
        console.log('Assessment processing completed:', result);
      })
      .catch(error => {
        console.error('Assessment processing failed:', error);
      });
    
    res.json({
      success: true,
      message: 'Assessment processing started',
      file: filename,
      config: { headless, slowMo }
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// Get processing status (placeholder for future implementation)
app.get('/api/status/:jobId', (req, res) => {
  res.json({
    success: true,
    message: 'Status endpoint - to be implemented',
    jobId: req.params.jobId
  });
});

// Error handling middleware
app.use((error, req, res, next) => {
  if (error instanceof multer.MulterError) {
    if (error.code === 'LIMIT_FILE_SIZE') {
      return res.status(400).json({
        success: false,
        error: 'File too large'
      });
    }
  }
  
  res.status(500).json({
    success: false,
    error: error.message
  });
});

// 404 handler
app.use('*', (req, res) => {
  res.status(404).json({
    success: false,
    error: 'Endpoint not found'
  });
});

// Start server
app.listen(PORT, () => {
  console.log(`🚀 Medical Assessment Automation API running on port ${PORT}`);
  console.log(`📊 Health check: http://localhost:${PORT}/health`);
  console.log(`📁 API endpoints:`);
  console.log(`   GET  /api/files - List available Excel files`);
  console.log(`   POST /api/upload - Upload Excel file`);
  console.log(`   GET  /api/excel/:filename - Read Excel file data`);
  console.log(`   POST /api/excel/:filename/update - Update Excel file`);
  console.log(`   POST /api/process/:filename - Process assessments`);
});

module.exports = app;



