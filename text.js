// json-to-csv.js
const fs = require('fs');
const path = require('path');

// Function to flatten nested JSON objects
function flattenObject(obj, prefix = '') {
  return Object.keys(obj).reduce((acc, key) => {
    const prefixedKey = prefix ? `${prefix}.${key}` : key;
    
    if (typeof obj[key] === 'object' && obj[key] !== null && !Array.isArray(obj[key])) {
      Object.assign(acc, flattenObject(obj[key], prefixedKey));
    } else {
      // Replace line breaks with a placeholder to avoid CSV issues
      let value = obj[key];
      if (typeof value === 'string') {
        value = value.replace(/\n/g, '\\n');
      }
      acc[prefixedKey] = value;
    }
    
    return acc;
  }, {});
}

// Function to convert JSON to CSV
function convertJsonToCsv(inputFile, outputFile) {
  try {
    // Read the JSON file
    const jsonContent = fs.readFileSync(inputFile, 'utf8');
    
    // Remove comments and parse JSON
    const cleanedJson = jsonContent.replace(/\/\/.*$/gm, '');
    const jsonData = JSON.parse(cleanedJson);
    
    // Flatten the JSON structure
    const flattenedData = flattenObject(jsonData);
    
    // Convert to CSV format
    const keys = Object.keys(flattenedData);
    let csvContent = 'Key,Value\n';
    
    for (const key of keys) {
      // Handle values with commas by wrapping them in quotes
      let value = flattenedData[key];
      if (value === undefined || value === null) {
        value = '';
      } else if (typeof value === 'string') {
        // Escape double quotes by doubling them
        value = value.replace(/"/g, '""');
        
        // Wrap in quotes if contains commas, quotes, or newlines
        if (value.includes(',') || value.includes('"') || value.includes('\n')) {
          value = `"${value}"`;
        }
      }
      
      csvContent += `${key},${value}\n`;
    }
    
    // Write the CSV file
    fs.writeFileSync(outputFile, csvContent, 'utf8');
    
    console.log(`Successfully converted ${inputFile} to ${outputFile}`);
  } catch (error) {
    console.error('Error converting JSON to CSV:', error.message);
  }
}

// Check if file arguments are provided
if (process.argv.length < 3) {
  console.log('Usage: node json-to-csv.js <input-json-file> [output-csv-file]');
  process.exit(1);
}

const inputFile = process.argv[2];
// If output file is not specified, generate it from input filename
const outputFile = process.argv[3] || path.basename(inputFile, '.json') + '.csv';

convertJsonToCsv(inputFile, outputFile);