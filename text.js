// json-to-csv.js
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx'); // Thêm thư viện xlsx
const os = require('os');

// Function to flatten nested JSON objects
function flattenObject(obj, prefix = '') {
  return Object.keys(obj).reduce((acc, key) => {
    const prefixedKey = prefix ? `${prefix}.${key}` : key;

    if (typeof obj[key] === 'object' && obj[key] !== null && !Array.isArray(obj[key])) {
      Object.assign(acc, flattenObject(obj[key], prefixedKey));
    } else {
      let value = obj[key];
      if (typeof value === 'string') {
        value = value.replace(/\n/g, '\\n');
      }
      acc[prefixedKey] = value;
    }

    return acc;
  }, {});
}

// Function to convert JSON to XLSX
function convertJsonToXlsx(inputFile, outputFile) {
  try {
    // Read the JSON file
    const jsonContent = fs.readFileSync(inputFile, 'utf8');

    // Remove comments and parse JSON
    const cleanedJson = jsonContent.replace(/\/\/.*$/gm, '');
    const jsonData = JSON.parse(cleanedJson);

    // Flatten the JSON structure
    const flattenedData = flattenObject(jsonData);

    // Prepare data for XLSX
    const rows = Object.entries(flattenedData).map(([key, value]) => ({ Key: key, Value: value }));

    // Create a new workbook and worksheet
    const workbook = xlsx.utils.book_new();
    const worksheet = xlsx.utils.json_to_sheet(rows);

    // Set column widths
    worksheet['!cols'] = [
      { wch: 50 }, // Width for 'Key' column
      { wch: 50 }  // Width for 'Value' column
    ];

    // Append the worksheet to the workbook
    xlsx.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

    // Write the XLSX file
    xlsx.writeFile(workbook, outputFile);

    console.log(`Successfully converted ${inputFile} to ${outputFile}`);
  } catch (error) {
    console.error('Error converting JSON to XLSX:', error.message);
  }
}

function convertXlsxToJson(inputFile, outputFile) {
  try {
    // Read the XLSX file
    const workbook = xlsx.readFile(inputFile);

    // Get the first sheet
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    // Convert sheet to JSON
    const jsonData = xlsx.utils.sheet_to_json(worksheet, { defval: null });

    // Write JSON data to file
    fs.writeFileSync(outputFile, JSON.stringify(jsonData, null, 2), 'utf8');

    console.log(`Successfully converted ${inputFile} to ${outputFile}`);
  } catch (error) {
    console.error('Error converting XLSX to JSON:', error.message);
  }
}

function convertXlsxToJsonWithStructure(inputFile, outputFile) {
  try {
    // Read the XLSX file
    const workbook = xlsx.readFile(inputFile);

    // Get the first sheet
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    // Convert sheet to JSON (raw data)
    const rawData = xlsx.utils.sheet_to_json(worksheet, { defval: null });

    // Transform raw data into the desired JSON structure
    const structuredData = rawData.reduce((acc, row) => {
      const keys = row.Key.split('.'); // Split nested keys by '.'
      let current = acc;

      keys.forEach((key, index) => {
        if (index === keys.length - 1) {
          // If it's the last key, assign the value
          current[key] = row.Value;
        } else {
          // If it's not the last key, ensure the object exists
          current[key] = current[key] || {};
          current = current[key];
        }
      });

      return acc;
    }, {});

    // Write structured JSON data to file
    fs.writeFileSync(outputFile, JSON.stringify(structuredData, null, 2), 'utf8');

    console.log(`Successfully converted ${inputFile} to ${outputFile}`);
  } catch (error) {
    console.error('Error converting XLSX to JSON:', error.message);
  }
}

function getUniqueFileName(filePath) {
  let counter = 1;
  let uniqueFilePath = filePath;

  // Kiểm tra nếu file đã tồn tại
  while (fs.existsSync(uniqueFilePath)) {
    const parsedPath = path.parse(filePath);
    uniqueFilePath = path.join(
      parsedPath.dir,
      `${parsedPath.name}(${counter})${parsedPath.ext}`
    );
    counter++;
  }

  return uniqueFilePath;
}

// Check if file arguments are provided
if (process.argv.length < 3) {
  console.log('Usage: node text.js <input-file> [output-file]');
  process.exit(1);
}

const inputFile = process.argv[2];
const inputExt = path.extname(inputFile).toLowerCase();
let outputFile = process.argv[3];

// Determine output file name and conversion type
if (!outputFile) {
  if (inputExt === '.json') {
    outputFile = path.join(os.homedir(), path.basename(inputFile, '.json') + '.xlsx');
  } else if (inputExt === '.xlsx') {
    outputFile = path.join(os.homedir(), path.basename(inputFile, '.xlsx') + '.json');
  } else {
    console.error('Unsupported file format. Please provide a .json or .xlsx file.');
    process.exit(1);
  }
}

// Ensure unique file name
outputFile = getUniqueFileName(outputFile);

// Perform conversion based on input file type
if (inputExt === '.json') {
  convertJsonToXlsx(inputFile, outputFile);
} else if (inputExt === '.xlsx') {
  convertXlsxToJsonWithStructure(inputFile, outputFile);
} else {
  console.error('Unsupported file format. Please provide a .json or .xlsx file.');
  process.exit(1);
}