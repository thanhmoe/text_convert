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
    const jsonContent = fs.readFileSync(inputFile, 'utf8');
    const cleanedJson = jsonContent.replace(/\/\/.*$/gm, '');
    const jsonData = JSON.parse(cleanedJson);

    // Flatten the JSON structure, giữ cả key có value là object rỗng
    function flatten(obj, prefix = '') {
      return Object.keys(obj).reduce((acc, key) => {
        const prefixedKey = prefix ? `${prefix}.${key}` : key;
        if (typeof obj[key] === 'object' && obj[key] !== null && !Array.isArray(obj[key])) {
          // Nếu là object rỗng thì vẫn ghi ra
          if (Object.keys(obj[key]).length === 0) {
            acc[prefixedKey] = '';
          } else {
            Object.assign(acc, flatten(obj[key], prefixedKey));
          }
        } else {
          let value = obj[key];
          if (typeof value === 'string') value = value.replace(/\n/g, '\\n');
          acc[prefixedKey] = value === undefined ? '' : value;
        }
        return acc;
      }, {});
    }

    const flattenedData = flatten(jsonData);
    const rows = Object.entries(flattenedData).map(([key, value]) => ({ key, value }));

    const workbook = xlsx.utils.book_new();
    const worksheet = xlsx.utils.json_to_sheet(rows);

    worksheet['!cols'] = [
      { wch: 50 }, // Key
      { wch: 50 }  // Value
    ];

    xlsx.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
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
    const workbook = xlsx.readFile(inputFile);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const rawData = xlsx.utils.sheet_to_json(worksheet, { defval: '' }); // giá trị mặc định là rỗng

    if (rawData.length === 0) throw new Error('No data found in the Excel file');

    // Tìm tên cột key và value không phân biệt hoa thường
    const firstRow = rawData[0];
    const columnNames = Object.keys(firstRow);
    const keyCol = columnNames.find(col => col.toLowerCase() === 'key');
    const valueCol = columnNames.find(col => col.toLowerCase() === 'value');
    if (!keyCol || !valueCol) throw new Error('Excel file must contain columns named "key" and "value" (case-insensitive)');

    const structuredData = rawData.reduce((acc, row) => {
      const keyPath = String(row[keyCol] || '').trim();
      if (!keyPath) {
        // Nếu không có key, lưu vào mảng đặc biệt
        if (!acc._no_key) acc._no_key = [];
        acc._no_key.push(row[valueCol] === undefined ? '' : row[valueCol]);
        return acc;
      }
      const keys = keyPath.split('.');
      let current = acc;
      keys.forEach((key, idx) => {
        if (idx === keys.length - 1) {
          current[key] = row[valueCol] === undefined ? '' : row[valueCol];
        } else {
          current[key] = current[key] || {};
          current = current[key];
        }
      });
      return acc;
    }, {});

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

function getTodayFolder() {
  const now = new Date();
  const dd = String(now.getDate()).padStart(2, '0');
  const mm = String(now.getMonth() + 1).padStart(2, '0');
  const yyyy = now.getFullYear();
  return `${dd}.${mm}.${yyyy}`;
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
  const todayFolder = getTodayFolder();
  const baseName =
    inputExt === '.json'
      ? path.basename(inputFile, '.json') + '.xlsx'
      : inputExt === '.xlsx'
      ? path.basename(inputFile, '.xlsx') + '.json'
      : null;

  if (!baseName) {
    console.error('Unsupported file format. Please provide a .json or .xlsx file.');
    process.exit(1);
  }

  // Đường dẫn thư mục lưu file
  const outputDir = path.join(os.homedir(), 'text_convert_output', todayFolder);
  if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });

  outputFile = path.join(outputDir, baseName);
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