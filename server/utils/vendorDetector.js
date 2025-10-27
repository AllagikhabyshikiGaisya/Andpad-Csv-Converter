const fs = require('fs')
const path = require('path')

function detectVendor(filename, headers) {
  const mappingsDir = path.join(__dirname, '../mappings')

  // Check if mappings directory exists
  if (!fs.existsSync(mappingsDir)) {
    console.error('Mappings directory not found:', mappingsDir)
    return { detected: false, mapping: null, method: null }
  }

  const mappingFiles = fs
    .readdirSync(mappingsDir)
    .filter(f => f.endsWith('.json'))

  console.log('=== VENDOR DETECTION START ===')
  console.log('Filename:', filename)
  console.log('CSV Headers:', headers)
  console.log('Available mappings:', mappingFiles)

  // Normalize filename for better matching
  const normalizedFilename = filename
    .toLowerCase()
    .replace(/\s+/g, '') // Remove spaces
    .replace(/[()（）\[\]【】]/g, '') // Remove ALL brackets - parentheses AND square brackets
    .replace(/\.csv$/i, '') // Remove .csv extension
  // Try to detect by filename first
  for (const file of mappingFiles) {
    try {
      const mappingPath = path.join(mappingsDir, file)
      const mapping = JSON.parse(fs.readFileSync(mappingPath, 'utf-8'))

      if (mapping.filePattern) {
        const patterns = Array.isArray(mapping.filePattern)
          ? mapping.filePattern
          : [mapping.filePattern]

        // Add alternate patterns if they exist
        if (mapping.alternatePatterns) {
          patterns.push(...mapping.alternatePatterns)
        }

        for (const pattern of patterns) {
          const normalizedPattern = pattern
            .toLowerCase()
            .replace(/\s+/g, '')
            .replace(/[()（）\[\]【】]/g, '') // Same here

          if (normalizedFilename.includes(normalizedPattern)) {
            console.log(`✓ Vendor detected by filename: ${mapping.vendor}`)
            console.log(`  Pattern matched: "${pattern}"`)
            return {
              detected: true,
              mapping: mapping,
              method: 'filename',
            }
          }
        }
      }
    } catch (error) {
      console.error(`Error reading mapping ${file}:`, error.message)
    }
  }

  console.log('✗ No filename pattern matched')

  // Try to detect by headers with flexible matching
  let bestMatch = null
  let bestScore = 0
  let bestMatchDetails = null

  for (const file of mappingFiles) {
    try {
      const mappingPath = path.join(mappingsDir, file)
      const mapping = JSON.parse(fs.readFileSync(mappingPath, 'utf-8'))

      const sourceColumns = Object.keys(mapping.map)
      let matchCount = 0
      const matchedColumns = []

      // Count matching columns
      sourceColumns.forEach(col => {
        const normalizedCol = col.trim().toLowerCase()
        const matchFound = headers.some(header => {
          const normalizedHeader = header.trim().toLowerCase()
          return (
            normalizedHeader === normalizedCol ||
            normalizedHeader.includes(normalizedCol) ||
            normalizedCol.includes(normalizedHeader)
          )
        })
        if (matchFound) {
          matchCount++
          matchedColumns.push(col)
        }
      })

      const matchScore = matchCount / sourceColumns.length

      console.log(`Mapping ${file}:`)
      console.log(
        `  Matched: ${matchCount}/${sourceColumns.length} (${(
          matchScore * 100
        ).toFixed(1)}%)`
      )
      console.log(`  Columns: ${matchedColumns.join(', ')}`)

      if (matchScore > bestScore) {
        bestScore = matchScore
        bestMatch = mapping
        bestMatchDetails = {
          file: file,
          matchCount: matchCount,
          totalColumns: sourceColumns.length,
          matchedColumns: matchedColumns,
        }
      }
    } catch (error) {
      console.error(`Error processing mapping ${file}:`, error.message)
    }
  }

  // Accept if more than 50% of columns match
  if (bestScore >= 0.5) {
    console.log(`✓ Vendor detected by headers: ${bestMatch.vendor}`)
    console.log(`  Best match: ${bestMatchDetails.file}`)
    console.log(`  Score: ${(bestScore * 100).toFixed(1)}%`)
    return {
      detected: true,
      mapping: bestMatch,
      method: 'headers',
    }
  }

  console.log('✗ No sufficient header match found')
  console.log('=== VENDOR DETECTION END ===')

  return {
    detected: false,
    mapping: null,
    method: null,
  }
}

module.exports = { detectVendor }
