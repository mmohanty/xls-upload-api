const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const path = require('path');

var cors = require('cors')

const app = express();
const port = 3500;

app.use(cors());

// Set up Multer to store uploaded files in a 'uploads' directory
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, './uploads');
  },
  filename: (req, file, cb) => {
    cb(null, `${Date.now()}-${file.originalname}`);
  }
});

const upload = multer({ storage });

// Create the 'uploads' directory if it doesn't exist
const fs = require('fs');
const dir = './uploads';
if (!fs.existsSync(dir)) {
  fs.mkdirSync(dir);
}

// API endpoint to upload an XLS file
app.post('/upload', upload.single('file'), (req, res) => {
  if (!req.file) {
    return res.status(400).send('No file uploaded.');
  }

  const filePath = path.join(__dirname, "../"+req.file.path);
  
  try {
    // Read the Excel file
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    
    // Convert the worksheet to JSON
    const data = xlsx.utils.sheet_to_json(worksheet);
    
    // Send back the extracted data
    res.json({ data });
  } catch (error) {
    console.log(error);
    res.status(500).send('Error processing file.');
  }
});

// Start the server
app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});