const fs = require('fs')
const path = require('path')

function detectVendor(filename, headers) {
  const mappingsDir = path.join(__dirname, '../mappings')

  if (!fs.existsSync(mappingsDir)) {
    console.error('❌ Mappings directory not found:', mappingsDir)
    return {
      detected: false,
      mapping: null,
      method: null,
      error: 'Mappings directory not found',
    }
  }

  const mappingFiles = fs
    .readdirSync(mappingsDir)
    .filter(f => f.endsWith('.json'))

  console.log('=== VENDOR DETECTION START ===')
  console.log('Original filename:', filename)
  console.log('CSV Headers (first 10):', headers.slice(0, 10))
  console.log('Available mapping files:', mappingFiles)

  // Enhanced normalization for Japanese filenames
  const normalizedFilename = filename
    .toLowerCase()
    .replace(/\s+/g, '') // Remove all whitespace
    .replace(/[()（）\[\]【】「」『』]/g, '') // Remove brackets
    .replace(/[_\-\.]/g, '') // Remove separators
    .replace(/\.csv$/i, '')
    .replace(/\.xlsx$/i, '')
    .replace(/請求書/g, '') // Remove common prefix
    .replace(/invoice/gi, '')

  console.log('Normalized filename:', normalizedFilename)

  // Load all mappings and separate them by priority
  const customParserMappings = []
  const standardMappings = []

  for (const file of mappingFiles) {
    try {
      const mappingPath = path.join(mappingsDir, file)
      const mapping = JSON.parse(fs.readFileSync(mappingPath, 'utf-8'))
      mapping._filename = file // Store filename for debugging

      if (mapping.customParser === true) {
        customParserMappings.push(mapping)
      } else {
        standardMappings.push(mapping)
      }
    } catch (error) {
      console.error(`❌ Error reading mapping ${file}:`, error.message)
    }
  }

  console.log(
    `Priority: ${customParserMappings.length} custom parsers, ${standardMappings.length} standard mappings`
  )

  // PRIORITY 1: Check custom parsers FIRST (they have specific formats)
  console.log('\n--- Phase 1: Checking Custom Parsers ---')
  for (const mapping of customParserMappings) {
    const match = checkFilenamePatterns(mapping, normalizedFilename)
    if (match) {
      console.log('✓✓✓ VENDOR DETECTED BY FILENAME (CUSTOM PARSER) ✓✓✓')
      console.log(`  Vendor: ${mapping.vendor}`)
      console.log(`  Method: filename (custom parser)`)
      console.log(`  Pattern matched: "${match.pattern}"`)
      console.log(`  File: ${mapping._filename}`)
      return {
        detected: true,
        mapping: mapping,
        method: 'filename-custom',
      }
    }
  }

  // PRIORITY 2: Check standard mappings by filename
  console.log('\n--- Phase 2: Checking Standard Mappings (Filename) ---')
  for (const mapping of standardMappings) {
    const match = checkFilenamePatterns(mapping, normalizedFilename)
    if (match) {
      console.log('✓✓✓ VENDOR DETECTED BY FILENAME ✓✓✓')
      console.log(`  Vendor: ${mapping.vendor}`)
      console.log(`  Method: filename`)
      console.log(`  Pattern matched: "${match.pattern}"`)
      console.log(`  File: ${mapping._filename}`)
      return {
        detected: true,
        mapping: mapping,
        method: 'filename',
      }
    }
  }

  console.log('⚠ No filename match found, trying header detection...')

  // PRIORITY 3: Header detection (only for standard mappings with actual column maps)
  console.log('\n--- Phase 3: Checking Headers ---')
  let bestMatch = null
  let bestScore = 0
  let bestMatchDetails = null

  for (const mapping of standardMappings) {
    // Skip if no map or empty map
    if (!mapping.map || Object.keys(mapping.map).length === 0) {
      console.log(`  Skipping ${mapping.vendor} - no column map`)
      continue
    }

    const sourceColumns = Object.keys(mapping.map)
    let matchCount = 0
    const matchedCols = []

    sourceColumns.forEach(col => {
      const found = headers.some(
        h =>
          h.toLowerCase().includes(col.toLowerCase()) ||
          col.toLowerCase().includes(h.toLowerCase())
      )
      if (found) {
        matchCount++
        matchedCols.push(col)
      }
    })

    const matchScore = matchCount / sourceColumns.length

    console.log(
      `  ${mapping.vendor}: ${matchCount}/${
        sourceColumns.length
      } columns matched (${Math.round(matchScore * 100)}%)`
    )

    if (matchScore > bestScore) {
      bestScore = matchScore
      bestMatch = mapping
      bestMatchDetails = {
        vendor: mapping.vendor,
        matchedColumns: matchedCols,
        totalColumns: sourceColumns.length,
        score: matchScore,
      }
    }
  }

  // Accept match if >= 50% columns matched
  if (bestScore >= 0.5) {
    console.log('✓✓✓ VENDOR DETECTED BY HEADERS ✓✓✓')
    console.log(`  Vendor: ${bestMatch.vendor}`)
    console.log(`  Method: headers`)
    console.log(`  Match score: ${Math.round(bestScore * 100)}%`)
    console.log(
      `  Matched columns: ${bestMatchDetails.matchedColumns.join(', ')}`
    )

    return {
      detected: true,
      mapping: bestMatch,
      method: 'headers',
      score: bestScore,
    }
  }

  console.log('✗✗✗ NO VENDOR DETECTED ✗✗✗')
  console.log('Tried:')
  console.log('  - Custom parsers by filename')
  console.log('  - Standard mappings by filename')
  console.log('  - Header column matching (min 50% required)')
  console.log(
    'Suggestion: Check if filename contains vendor name or verify CSV headers'
  )
  console.log('=== VENDOR DETECTION END ===')

  return {
    detected: false,
    mapping: null,
    method: null,
    error: 'No matching vendor found. Please check filename or file format.',
  }
}

// Helper function to check filename patterns
function checkFilenamePatterns(mapping, normalizedFilename) {
  if (!mapping.filePattern) return null

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
      .replace(/[()（）\[\]【】「」『』]/g, '')
      .replace(/[_\-\.]/g, '')

    console.log(`  Testing: "${pattern}" → "${normalizedPattern}"`)

    if (normalizedFilename.includes(normalizedPattern)) {
      return { pattern: pattern, normalizedPattern: normalizedPattern }
    }
  }

  return null
}

module.exports = { detectVendor }
