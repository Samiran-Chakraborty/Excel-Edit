import express from 'express';
import cors from 'cors';
import multer from 'multer';
import * as xlsx from 'xlsx';

const app = express();
const port = process.env.PORT || 3000;

app.use(cors({ origin: '*' }));
app.use(express.json({ limit: '50mb' }));

// Use multer in memory so we don't save files to disk
const storage = multer.memoryStorage();
const upload = multer({ storage });

// Health check
app.get('/', (req, res) => {
  res.json({ status: 'ok', message: 'Excel Editor API is running 🚀' });
});

// Upload Endpoint
app.post('/api/upload', upload.single('file'), (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: 'No file uploaded' });
  }

  try {
    const workbook = xlsx.read(req.file.buffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Convert sheet to JSON array of arrays to preserve column order if needed, 
    // or array of objects. Let's use array of objects for easier AG Grid integration.
    // header: 1 gives us an array of arrays
    const rawData = xlsx.utils.sheet_to_json(sheet, { header: 1 });

    if (rawData.length === 0) {
      return res.status(400).json({ error: 'Excel file is empty' });
    }

    // Process first row as headers, handle empty headers
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

    res.json({
      headers,
      rows,
      sheetName
    });
  } catch (err) {
    console.error('Error parsing Excel file:', err);
    res.status(500).json({ error: 'Error parsing Excel file' });
  }
});

// Download Endpoint
app.post('/api/download', (req, res) => {
  try {
    const { rows, sheetName, bookType = 'xlsx', filename = 'edited_data.xlsx' } = req.body;
    
    if (!rows || !Array.isArray(rows)) {
      return res.status(400).json({ error: 'Invalid data format' });
    }

    // Convert JSON back to sheet
    const worksheet = xlsx.utils.json_to_sheet(rows);
    const workbook = xlsx.utils.book_new();
    
    xlsx.utils.book_append_sheet(workbook, worksheet, sheetName || 'Sheet1');

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

app.listen(port, () => {
  console.log(`Backend server running on http://localhost:${port}`);
});
