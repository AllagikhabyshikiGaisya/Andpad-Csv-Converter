// Test script for debugging the converter
const fs = require('fs')
const path = require('path')
const Papa = require('papaparse')
const { detectVendor } = require('../utils/vendorDetector')
const { generateExcel } = require('../utils/excelGenerator')

console.log('=== CONVERTER TEST SCRIPT ===\n')

// Test file path (update this to your test file)
const testFilePath =
  process.argv[2] || './test-files/【請求書】2025.7クリーン産業.csv'

if (!fs.existsSync(testFilePath)) {
  console.error('❌ Test file not found:', testFilePath)
  console.log('Usage: node test/test-converter.js <path-to-csv-file>')
  process.exit(1)
}

const filename = path.basename(testFilePath)
console.log('Testing file:', filename)
console.log('Full path:', testFilePath)

// Read and parse CSV
const fileContent = fs.readFileSync(testFilePath, 'utf-8')
console.log('\nFile size:', fileContent.length, 'bytes')
console.log('First 300 characters:')
console.log(fileContent.substring(0, 300))
console.log('...\n')

const parseResult = Papa.parse(fileContent, {
  header: true,
  skipEmptyLines: true,
  dynamicTyping: false,
  encoding: 'utf-8',
})

console.log('Parse result:')
console.log('- Total rows:', parseResult.data.length)
console.log('- Headers:', parseResult.meta.fields)
console.log('- Parse errors:', parseResult.errors.length)

if (parseResult.errors.length > 0) {
  console.log('\nParse errors:', parseResult.errors.slice(0, 3))
}

console.log('\nFirst 3 rows:')
parseResult.data.slice(0, 3).forEach((row, i) => {
  console.log(`Row ${i}:`, Object.values(row).slice(0, 10))
})

// Test vendor detection
console.log('\n--- Testing Vendor Detection ---')
const detection = detectVendor(filename, parseResult.meta.fields)

if (!detection.detected) {
  console.error('❌ Vendor detection FAILED')
  console.error('Error:', detection.error)
  process.exit(1)
}

console.log('✓ Vendor detected:', detection.mapping.vendor)
console.log('Method:', detection.method)
console.log('Custom parser:', detection.mapping.customParser)
console.log(
  'Mapping has',
  Object.keys(detection.mapping.map || {}).length,
  'column mappings'
)

// Test Excel generation
console.log('\n--- Testing Excel Generation ---')
const result = generateExcel(parseResult.data, detection.mapping)

if (!result.success) {
  console.error('❌ Excel generation FAILED')
  console.error('Error:', result.error)
  process.exit(1)
}

console.log('✓ Excel generated successfully')
console.log('Output rows:', result.rowCount)

// Save output file
const outputPath = path.join(__dirname, `test-output-${Date.now()}.xlsx`)
fs.writeFileSync(outputPath, result.buffer)
console.log('✓ Output saved to:', outputPath)

console.log('\n=== ALL TESTS PASSED ===')
console.log('You can now open the Excel file to verify the output.')
