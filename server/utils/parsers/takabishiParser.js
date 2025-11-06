// ============================================
// TAKABISHI PARSER - FIXED FOR ACTUAL COLUMN STRUCTURE
// ============================================
const BaseParser = require('./baseParser')
const { createMasterRow, cleanNumber, formatDate } = require('../excelUtils')

class TakabishiParser extends BaseParser {
  constructor() {
    super('高菱管理')
  }

  parse(csvData) {
    this.logStart(csvData)
    const results = []

    console.log('✓ Parsing Takabishi invoice format (H/M rows)')

    // Debug: Show first row structure
    if (csvData.length > 0) {
      const firstRow = csvData[0]
      console.log('First row columns:', Object.keys(firstRow).slice(0, 10))
      console.log('Row type column value:', firstRow[''])
      console.log('H column value:', firstRow['H'])
    }

    // Extract default 案件管理ID from header if exists
    let defaultProjectId = ''
    for (let i = 0; i < Math.min(csvData.length, 10); i++) {
      const row = csvData[i]

      // Check for project ID columns
      if (row['案件管理ID']) {
        const id = String(row['案件管理ID']).trim()
        if (id && id !== '案件管理ID' && id !== '') {
          defaultProjectId = id
          console.log(`✅ Found 案件管理ID in header: ${defaultProjectId}`)
          break
        }
      }
    }

    if (!defaultProjectId) {
      console.warn('⚠️ WARNING: 案件管理ID not found in CSV header')
    }

    // Group rows by invoice (H rows)
    const invoiceGroups = []
    let currentInvoice = null

    for (let i = 0; i < csvData.length; i++) {
      const row = csvData[i]

      // ✅ CRITICAL FIX: The row type is in the EMPTY STRING column, not 'H'
      const rowType = String(row[''] || '').trim()
      const hColumn = String(row['H'] || '').trim()

      // Debug first few rows
      if (i < 5) {
        console.log(`Row ${i}: rowType="${rowType}", H="${hColumn}"`)
      }

      // Skip completely empty rows
      if (!rowType && !hColumn) {
        this.skippedCount++
        continue
      }

      // Check if this is a header row (invoice summary)
      // The actual row type indicator is in the empty string column
      // and should be "INV" or the H column should be "H"
      if (rowType === 'INV' || hColumn === 'H') {
        // Header row - contains invoice summary
        if (currentInvoice) {
          invoiceGroups.push(currentInvoice)
        }

        currentInvoice = {
          header: row,
          details: [],
          rowIndex: i
        }

        console.log(`✓ Found invoice header at row ${i}`)
      } else if (hColumn === 'M' && currentInvoice) {
        // Detail row - belongs to current invoice
        currentInvoice.details.push(row)
      }
    }

    // Add last invoice
    if (currentInvoice) {
      invoiceGroups.push(currentInvoice)
    }

    console.log(`✓ Found ${invoiceGroups.length} invoice groups`)

    if (invoiceGroups.length === 0) {
      console.error('❌ No invoice groups found')
      console.error('This might indicate a CSV parsing or structure issue')
    }

    // Process each invoice group
    for (let groupIndex = 0; groupIndex < invoiceGroups.length; groupIndex++) {
      const invoice = invoiceGroups[groupIndex]
      const header = invoice.header

      // Debug: Show header structure
      if (groupIndex === 0) {
        console.log('\nHeader row structure:')
        console.log('  請求日付:', header['請求日付'])
        console.log('  得意先番号:', header['得意先番号'])
        console.log('  配置先番号:', header['配置先番号'])
        console.log('  請求金額:', header['請求金額'])
        console.log('  消費税:', header['消費税'])
        console.log('  請求合計:', header['請求合計'])
      }

      // Extract header data
      const invoiceDate = this.parseDate(header['請求日付'] || header['請求日'])
      const customerNumber = String(header['得意先番号'] || '').trim()
      const placementNumber = String(header['配置先番号'] || '').trim()
      const invoiceAmount = cleanNumber(header['請求金額'] || '0')
      const taxAmount = cleanNumber(header['消費税'] || header['外税'] || '0')
      const totalAmount = cleanNumber(header['請求合計'] || '0')

      // Calculate tax-excluded amount
      let amountExcludingTax = invoiceAmount
      if (!amountExcludingTax || amountExcludingTax === '0') {
        // If 請求金額 is missing, calculate from total - tax
        if (totalAmount && taxAmount) {
          const total = parseFloat(totalAmount)
          const tax = parseFloat(taxAmount)
          amountExcludingTax = String(total - tax)
        }
      }

      if (!amountExcludingTax || amountExcludingTax === '0') {
        console.log(`⊘ Skipping invoice group ${groupIndex + 1}: No amount (請求金額: ${invoiceAmount}, 消費税: ${taxAmount}, 請求合計: ${totalAmount})`)
        this.skippedCount++
        continue
      }

      // Extract project ID from header or details
      let projectId = defaultProjectId

      // Check header for project ID
      if (header['案件管理ID']) {
        projectId = String(header['案件管理ID']).trim()
      } else if (header['工事番号']) {
        projectId = String(header['工事番号']).trim()
      }

      // If still no project ID, use customer/placement combo
      if (!projectId) {
        projectId = `CUST${customerNumber}_PLACE${placementNumber}`
        console.log(`⚠️ No 案件管理ID found, using generated ID: ${projectId}`)
      }

      // Create site name from customer and placement numbers
      const siteName = `得意先${customerNumber}_配置先${placementNumber}`

      // Invoice description
      const itemDescription = `高菱管理 ${this.formatDateForDisplay(invoiceDate)} 請求分`

      console.log(`✓ Processing invoice ${groupIndex + 1}: ¥${amountExcludingTax} (${invoice.details.length} detail rows)`)

      const masterRow = createMasterRow({
        vendor: '高菱管理',
        site: siteName,
        date: invoiceDate,
        item: itemDescription,
        qty: '1',
        unit: '式',
        price: amountExcludingTax,
        amount: amountExcludingTax,
        workNo: customerNumber,
        remarks: `配置先:${placementNumber}`,
        projectId: projectId,
        result: '承認',
      })

      results.push(masterRow)
      this.processedCount++
    }

    this.logComplete(results)
    this.validateResults(results)
    return results
  }

  /**
   * Parse date from YYYYMMDD format to YYYY/MM/DD
   */
  parseDate(dateStr) {
    if (!dateStr) return ''

    const cleaned = String(dateStr).trim()

    // Format: YYYYMMDD (e.g., "20250801")
    if (cleaned.match(/^\d{8}$/)) {
      const year = cleaned.substring(0, 4)
      const month = cleaned.substring(4, 6)
      const day = cleaned.substring(6, 8)
      return `${year}/${month}/${day}`
    }

    // Already formatted
    if (cleaned.match(/^\d{4}\/\d{2}\/\d{2}$/)) {
      return cleaned
    }

    return cleaned
  }

  /**
   * Format date for display in item description
   */
  formatDateForDisplay(dateStr) {
    if (!dateStr) return ''

    // Extract YYYY年MM月 from YYYY/MM/DD
    const match = dateStr.match(/^(\d{4})\/(\d{2})/)
    if (match) {
      return `${match[1]}年${match[2]}月`
    }

    return dateStr
  }
}

module.exports = TakabishiParser
