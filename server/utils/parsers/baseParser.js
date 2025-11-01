// ============================================
// BASE PARSER CLASS (for baseParser.js)
// ============================================

class BaseParser {
  constructor(vendorName) {
    this.vendorName = vendorName
    this.processedCount = 0
    this.skippedCount = 0
  }

  /**
   * Main parse method - to be overridden by subclasses
   */
  parse(csvData) {
    throw new Error('Parse method must be implemented by subclass')
  }

  /**
   * Log parsing start
   */
  logStart(csvData) {
    console.log(`=== ${this.vendorName.toUpperCase()} PARSER START ===`)
    console.log('Input rows:', csvData.length)
  }

  /**
   * Log parsing completion
   */
  logComplete(results) {
    console.log(`=== ${this.vendorName.toUpperCase()} PARSING COMPLETE ===`)
    console.log(`✓ Processed: ${this.processedCount} items`)
    console.log(`⊘ Skipped: ${this.skippedCount} rows`)
    console.log(`Total output rows: ${results.length}`)
  }

  /**
   * Validate results
   */
  validateResults(results) {
    if (!results || results.length === 0) {
      throw new Error(
        '有効なデータ行が見つかりませんでした。ファイル形式を確認してください。'
      )
    }
  }

  /**
   * Find data start index by looking for header patterns
   */
  findDataStart(csvData, headerPatterns, maxRows = 20) {
    for (let i = 0; i < Math.min(csvData.length, maxRows); i++) {
      const row = csvData[i]
      const rowText = [...Object.keys(row), ...Object.values(row)].join('|')
      const matchedPatterns = headerPatterns.filter(p => rowText.includes(p))

      if (matchedPatterns.length >= 2) {
        console.log(`✓ Header found at row ${i}`)
        return i + 1
      }
    }
    return -1
  }

  /**
   * Extract metadata from header rows
   */
  extractMetadata(csvData, maxRows = 15) {
    const metadata = {
      siteName: '',
      clientName: '',
      invoiceDate: '',
      projectName: '',
    }

    for (let i = 0; i < Math.min(csvData.length, maxRows); i++) {
      const row = csvData[i]
      const values = Object.values(row)
      const firstValue = String(values[0] || '').trim()

      // Look for site/project name
      if (firstValue.includes('様邸') || firstValue.includes('工事')) {
        metadata.siteName = firstValue
        metadata.projectName = firstValue
      }

      // Look for client name
      if (firstValue.includes('得意先') || firstValue.includes('ALLAGI')) {
        metadata.clientName = firstValue.replace(/得意先名?:|様/g, '').trim()
        if (metadata.clientName.includes('ALLAGI')) {
          metadata.clientName = 'ALLAGI株式会社'
        }
      }

      // Look for date
      for (const val of values) {
        const str = String(val).trim()
        if (str.match(/\d{4}\/\d{1,2}\/\d{1,2}/) || str.match(/\d+月\d+日/)) {
          metadata.invoiceDate = str
          break
        }
      }
    }

    return metadata
  }
}

module.exports = BaseParser
