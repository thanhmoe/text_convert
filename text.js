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
  console.log('Usage: node json-to-xlsx.js <input-json-file> [output-xlsx-file]');
  process.exit(1);
}

const inputFile = process.argv[2];
// Nếu output file không được chỉ định, lưu vào thư mục Home với tên từ input file
let outputFile = process.argv[3] || path.join(os.homedir(), path.basename(inputFile, '.json') + '.xlsx');

// Đảm bảo tên file là duy nhất
outputFile = getUniqueFileName(outputFile);

convertJsonToXlsx(inputFile, outputFile);