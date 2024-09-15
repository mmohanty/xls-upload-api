const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const path = require('path');

var cors = require('cors')

const app = express();
const port = 3500;

app.use(cors());

const { decode_range, encode_cell } = xlsx.utils;

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


function log_all_cells(ws) {
  var range = decode_range(ws["!ref"]);
  var dense = ws["!data"] != null; // test if sheet is dense
  for(var R = 0; R <= range.e.r; ++R) {
    for(var C = 0; C <= range.e.c; ++C) {
      var cell = dense ? ws["!data"]?.[R]?.[C] : ws[encode_cell({r:R, c:C})];
      console.log("Row");
      console.log(R);
      console.log("Column");
      console.log( C);
      console.log("Cell;")
      console.log(cell);
    }
  }
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
    
    fs.unlinkSync(filePath);
    //iterate of data
    //log_all_cells(worksheet);
    
    //Array.from(data).map(row => {console.log(row['Loan Status'])});

  
    Array.from(data).forEach(row => {
      //console.log(row['Loan Status'])
      console.log(row);
    });

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