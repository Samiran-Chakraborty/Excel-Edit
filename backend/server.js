import express from 'express';
import cors from 'cors';
import multer from 'multer';
import * as xlsx from 'xlsx';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const port = process.env.PORT || 3000;

app.use(cors({ origin: '*' }));
app.use(express.json({ limit: '50mb' }));

// Use multer in memory so we don't save files to disk
const storage = multer.memoryStorage();
const upload = multer({ storage });

// Serve built React frontend
const frontendDist = path.join(__dirname, '..', 'frontend', 'dist');
app.use(express.static(frontendDist));

// Upload Endpoint
app.post('/api/upload', upload.single('file'), (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: 'No file uploaded' });
  }

  try {
    const workbook = xlsx.read(req.file.buffer, { type: 'buffer' });
    const sheets = [];

    workbook.SheetNames.forEach(sheetName => {
      const sheet = workbook.Sheets[sheetName];
      const rawData = xlsx.utils.sheet_to_json(sheet, { header: 1 });
      
      if (rawData.length > 0) {
        const headers = rawData[0].map((h, i) => h ? String(h) : `Column${i + 1}`);
        const rows = [];

        for (let i = 1; i < rawData.length; i++) {
          const rowArr = rawData[i];
          const rowObj = {};
          let hasData = false;
          for (let j = 0; j < headers.length; j++) {
            rowObj[headers[j]] = rowArr[j] !== undefined ? rowArr[j] : '';
            if (rowArr[j] !== undefined && rowArr[j] !== '') hasData = true;
          }
          if (hasData) {
            rows.push(rowObj);
          }
        }
        
        sheets.push({ sheetName, headers, rows });
      } else {
        // Empty sheet
        sheets.push({ sheetName, headers: [], rows: [] });
      }
    });

    if (sheets.length === 0) {
      return res.status(400).json({ error: 'Excel file is completely empty' });
    }

    res.json({ sheets });
  } catch (err) {
    console.error('Error parsing Excel file:', err);
    res.status(500).json({ error: 'Error parsing Excel file' });
  }
});

// Download Endpoint
app.post('/api/download', (req, res) => {
  try {
    const { sheets, bookType = 'xlsx', filename = 'edited_data.xlsx' } = req.body;
    
    if (!sheets || !Array.isArray(sheets)) {
      return res.status(400).json({ error: 'Invalid data format: Expected an array of sheets' });
    }

    const workbook = xlsx.utils.book_new();

    sheets.forEach(({ sheetName, rows }) => {
      // Ensure there's a valid sheetName and fallback if not
      const safeSheetName = sheetName || 'Sheet1';
      // xlsx json_to_sheet creates an empty sheet if rows is empty or not provided
      const worksheet = xlsx.utils.json_to_sheet(rows || []);
      xlsx.utils.book_append_sheet(workbook, worksheet, safeSheetName);
    });

    // Generate buffer
    const buffer = xlsx.write(workbook, { type: 'buffer', bookType });

    res.set({
      'Content-Disposition': `attachment; filename="${filename}"`,
      'Content-Type': 'application/octet-stream',
    });

    res.send(buffer);
  } catch (err) {
    console.error('Error generating Excel file:', err);
    res.status(500).json({ error: 'Error generating Excel file' });
  }
});

// Catch-all: serve React app for any non-API route
app.get('*', (req, res) => {
  res.sendFile(path.join(frontendDist, 'index.html'));
});

app.listen(port, () => {
  console.log(`Server running on http://localhost:${port}`);
});
