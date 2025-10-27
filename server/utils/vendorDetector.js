const fs = require('fs')
const path = require('path')

function detectVendor(filename, headers) {
  const mappingsDir = path.join(__dirname, '../mappings')

  if (!fs.existsSync(mappingsDir)) {
    console.error('Mappings directory not found:', mappingsDir)
    return { detected: false, mapping: null, method: null }
  }

  const mappingFiles = fs
    .readdirSync(mappingsDir)
    .filter(f => f.endsWith('.json'))

  console.log('=== VENDOR DETECTION START ===')
  console.log('Original filename:', filename)
  console.log('CSV Headers:', headers.slice(0, 10))
  console.log('Available mappings:', mappingFiles)

  // More aggressive normalization
  const normalizedFilename = filename
    .toLowerCase()
    .replace(/\s+/g, '')
    .replace(/[()（）\[\]【】「」『』]/g, '') // Remove ALL types of brackets
    .replace(/[_\-\.]/g, '') // Remove separators
    .replace(/\.csv$/i, '')
    .replace(/\.xlsx$/i, '')

  console.log('Normalized filename:', normalizedFilename)

  // Try filename detection first
  for (const file of mappingFiles) {
    try {
      const mappingPath = path.join(mappingsDir, file)
      const mapping = JSON.parse(fs.readFileSync(mappingPath, 'utf-8'))

      if (mapping.filePattern) {
        const patterns = Array.isArray(mapping.filePattern)
          ? mapping.filePattern
          : [mapping.filePattern]

        if (mapping.alternatePatterns) {
          patterns.push(...mapping.alternatePatterns)
        }

        for (const pattern of patterns) {
          const normalizedPattern = pattern
            .toLowerCase()
            .replace(/\s+/g, '')
            .replace(/[()（）\[\]【】「」『』]/g, '')
            .replace(/[_\-\.]/g, '')

          console.log(`  Checking: "${pattern}" -> "${normalizedPattern}"`)
          console.log(`  Against: "${normalizedFilename}"`)
          console.log(
            `  Match: ${normalizedFilename.includes(normalizedPattern)}`
          )

          if (normalizedFilename.includes(normalizedPattern)) {
            console.log(`✓✓✓ VENDOR DETECTED: ${mapping.vendor} ✓✓✓`)
            console.log(`  Method: filename`)
            console.log(`  Pattern: "${pattern}"`)
            console.log(`  Custom Parser: ${mapping.customParser}`)
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

  console.log('✗ No filename match found')

  // Header detection (skip for custom parsers with empty maps)
  let bestMatch = null
  let bestScore = 0

  for (const file of mappingFiles) {
    try {
      const mappingPath = path.join(mappingsDir, file)
      const mapping = JSON.parse(fs.readFileSync(mappingPath, 'utf-8'))

      // Skip custom parsers with no map in header detection
      if (
        mapping.customParser &&
        (!mapping.map || Object.keys(mapping.map).length === 0)
      ) {
        console.log(`Skipping ${file} - custom parser with no map`)
        continue
      }

      const sourceColumns = Object.keys(mapping.map || {})
      if (sourceColumns.length === 0) continue

      let matchCount = 0
      sourceColumns.forEach(col => {
        if (headers.some(h => h.toLowerCase().includes(col.toLowerCase()))) {
          matchCount++
        }
      })

      const matchScore = matchCount / sourceColumns.length
      if (matchScore > bestScore) {
        bestScore = matchScore
        bestMatch = mapping
      }
    } catch (error) {
      console.error(`Error processing ${file}:`, error.message)
    }
  }

  if (bestScore >= 0.5) {
    console.log(`✓ Vendor detected by headers: ${bestMatch.vendor}`)
    return { detected: true, mapping: bestMatch, method: 'headers' }
  }

  console.log('✗✗✗ NO VENDOR DETECTED ✗✗✗')
  console.log('=== VENDOR DETECTION END ===')
  return { detected: false, mapping: null, method: null }
}

module.exports = { detectVendor }
