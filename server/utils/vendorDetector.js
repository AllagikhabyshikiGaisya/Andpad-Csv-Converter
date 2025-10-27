const fs = require('fs')
const path = require('path')

function detectVendor(filename, headers) {
  const mappingsDir = path.join(__dirname, '../mappings')
  const mappingFiles = fs
    .readdirSync(mappingsDir)
    .filter(f => f.endsWith('.json'))

  // Try to detect by filename first
  for (const file of mappingFiles) {
    const mappingPath = path.join(mappingsDir, file)
    const mapping = JSON.parse(fs.readFileSync(mappingPath, 'utf-8'))

    if (
      mapping.filePattern &&
      filename.toLowerCase().includes(mapping.filePattern.toLowerCase())
    ) {
      return {
        detected: true,
        mapping: mapping,
        method: 'filename',
      }
    }
  }

  // Try to detect by headers
  for (const file of mappingFiles) {
    const mappingPath = path.join(mappingsDir, file)
    const mapping = JSON.parse(fs.readFileSync(mappingPath, 'utf-8'))

    const sourceColumns = Object.keys(mapping.map)
    const matchCount = sourceColumns.filter(col => headers.includes(col)).length

    // If more than 50% of expected columns match, consider it a match
    if (matchCount / sourceColumns.length > 0.5) {
      return {
        detected: true,
        mapping: mapping,
        method: 'headers',
      }
    }
  }

  return {
    detected: false,
    mapping: null,
    method: null,
  }
}

module.exports = { detectVendor }
