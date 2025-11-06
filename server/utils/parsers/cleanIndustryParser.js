const BaseParser = require('./baseParser')
const { createMasterRow, cleanNumber, formatDate } = require('../excelUtils')

class CleanIndustryParser extends BaseParser {
  constructor() {
    super('クリーン産業')
  }

  parse(csvData) {
    this.logStart(csvData)

    if (this.isAlreadyANDPADFormat(csvData)) {
      console.log('⚠️ WARNING: File is already in ANDPAD format')
      throw new Error(
        'このファイルは既に変換済みです。元の業者ファイルをアップロードしてください。'
      )
    }

    const results = []
    let metadata = {
      clientName: 'ALLAGI株式会社',
      invoiceDate: '',
      siteName: '',
      projectId: '', // Will be extracted from CSV
    }

    const headerDates = []

    // ============================================
    // EXTRACT METADATA INCLUDING 案件管理ID
    // ============================================
    for (let i = 0; i < Math.min(csvData.length, 15); i++) {
      const row = csvData[i]
      const values = Object.values(row)
      const rowText = values.join('|')

      if (rowText.includes('ALLAGI')) {
        metadata.clientName = 'ALLAGI株式会社'
      }

      // ✅ CRITICAL: Extract 案件管理ID from header
      if (
        rowText.includes('案件管理ID') ||
        rowText.includes('工事番号') ||
        rowText.includes('現場No') ||
        rowText.includes('物件No')
      ) {
        console.log(`✓ Found row with project ID keyword: Row ${i}`)

        // Look for ID in the same row or next row
        for (const val of values) {
          const cleaned = String(val).trim()
          if (
            cleaned &&
            cleaned.length > 0 &&
            cleaned !== '案件管理ID' &&
            cleaned !== '工事番号' &&
            cleaned !== '現場No' &&
            cleaned !== '物件No'
          ) {
            // Check if it looks like a project ID (not a date or other metadata)
            if (
              !cleaned.match(/^\d{4}\/\d{1,2}\/\d{1,2}$/) &&
              !cleaned.includes('請求') &&
              !cleaned.includes('株式会社') &&
              !cleaned.includes('ALLAGI') &&
              !cleaned.includes('TEL') &&
              !cleaned.includes('FAX')
            ) {
              metadata.projectId = cleaned
              console.log(
                `✅ Found 案件管理ID in header: ${metadata.projectId}`
              )
              break
            }
          }
        }

        // Also check next row
        if (!metadata.projectId && i + 1 < csvData.length) {
          const nextRow = csvData[i + 1]
          const nextValues = Object.values(nextRow)
          for (const val of nextValues) {
            const cleaned = String(val).trim()
            if (cleaned && cleaned.length > 0) {
              if (
                !cleaned.match(/^\d{4}\/\d{1,2}\/\d{1,2}$/) &&
                !cleaned.includes('請求') &&
                !cleaned.includes('株式会社')
              ) {
                metadata.projectId = cleaned
                console.log(
                  `✅ Found 案件管理ID in next row: ${metadata.projectId}`
                )
                break
              }
            }
          }
        }

        if (metadata.projectId) break
      }

      // PRIORITY 1: Look for invoice date with context keywords
      if (
        rowText.includes('請求年月日') ||
        rowText.includes('請求日') ||
        rowText.includes('発行日')
      ) {
        console.log(`✓ Found row with invoice date keyword: Row ${i}`)

        // Look for date in the NEXT row
        if (i + 1 < csvData.length) {
          const nextRow = csvData[i + 1]
          const nextValues = Object.values(nextRow)
          const nextRowText = nextValues.join('|')

          const dateMatch = nextRowText.match(/(\d{4})\/(\d{1,2})\/(\d{1,2})/)
          if (dateMatch) {
            const year = dateMatch[1]
            const month = String(dateMatch[2]).padStart(2, '0')
            const day = String(dateMatch[3]).padStart(2, '0')
            metadata.invoiceDate = `${year}/${month}/${day}`
            console.log(
              `✓ Found invoice date in next row: ${metadata.invoiceDate}`
            )
            break
          }
        }

        // Also try to find date in the SAME row
        const dateMatch = rowText.match(/(\d{4})\/(\d{1,2})\/(\d{1,2})/)
        if (dateMatch && !metadata.invoiceDate) {
          const year = dateMatch[1]
          const month = String(dateMatch[2]).padStart(2, '0')
          const day = String(dateMatch[3]).padStart(2, '0')
          metadata.invoiceDate = `${year}/${month}/${day}`
          console.log(
            `✓ Found invoice date with keyword in same row: ${metadata.invoiceDate}`
          )
          break
        }
      }

      // Collect all dates for analysis
      const allMatches = rowText.matchAll(/(\d{4})\/(\d{1,2})\/(\d{1,2})/g)
      for (const match of allMatches) {
        const year = match[1]
        const month = String(match[2]).padStart(2, '0')
        const day = String(match[3]).padStart(2, '0')
        headerDates.push(`${year}/${month}/${day}`)
      }
    }

    console.log('Extracted metadata (initial):', metadata)
    console.log('Header dates found:', headerDates)

    // Find header row that contains column names
    let headerRowIndex = -1
    let dataStartIndex = -1

    for (let i = 0; i < Math.min(csvData.length, 20); i++) {
      const row = csvData[i]
      const values = Object.values(row)
      const rowText = values.join('|')

      // Look for header row with column names
      if (
        rowText.includes('業者名') &&
        rowText.includes('現場名') &&
        rowText.includes('品名')
      ) {
        headerRowIndex = i
        dataStartIndex = i + 1
        console.log(
          `✓ Found header row at ${i}, data starts at ${dataStartIndex}`
        )
        break
      }
    }

    if (dataStartIndex === -1) {
      dataStartIndex = 8
      console.log('⚠️ Could not find header row, using fallback start at row 8')
    }

    // CRITICAL: Collect transaction dates for smart invoice date detection
    const transactionDates = []

    // First pass: Collect all transaction dates
    for (let i = dataStartIndex; i < csvData.length; i++) {
      const row = csvData[i]
      const values = Object.values(row).map(v => String(v || '').trim())

      const dateStr = values[2] || ''
      const itemName = values[4] || ''

      // Skip summary rows
      if (itemName.startsWith('【')) continue

      // Extract transaction date
      if (dateStr && dateStr.match(/^\d{1,2}\/\d{1,2}$/)) {
        const parts = dateStr.split('/')
        const year = new Date().getFullYear()
        const month = String(parts[0]).padStart(2, '0')
        const day = String(parts[1]).padStart(2, '0')
        transactionDates.push(`${year}/${month}/${day}`)
      } else if (dateStr && dateStr.match(/^\d{4}\/\d{1,2}\/\d{1,2}$/)) {
        const parts = dateStr.split('/')
        const year = parts[0]
        const month = String(parts[1]).padStart(2, '0')
        const day = String(parts[2]).padStart(2, '0')
        transactionDates.push(`${year}/${month}/${day}`)
      }
    }

    console.log('Transaction dates found:', transactionDates.slice(0, 5))

    // SMART DETECTION: Determine correct invoice date
    if (!metadata.invoiceDate) {
      if (headerDates.length > 0) {
        const uniqueDates = [...new Set(headerDates)].sort()
        metadata.invoiceDate = uniqueDates[uniqueDates.length - 1]
        console.log(
          `✓ Using latest header date as invoice date: ${metadata.invoiceDate}`
        )
        console.log(`  (Other header dates: ${uniqueDates.join(', ')})`)
      } else if (transactionDates.length > 0) {
        const uniqueTransDates = [...new Set(transactionDates)].sort()
        metadata.invoiceDate = uniqueTransDates[uniqueTransDates.length - 1]
        console.log(
          `✓ Using latest transaction date as invoice date: ${metadata.invoiceDate}`
        )
      }
    }

    // CRITICAL: Final validation
    if (!metadata.invoiceDate) {
      console.warn('⚠️ WARNING: No invoice date found, using current date')
      const today = new Date()
      const year = today.getFullYear()
      const month = String(today.getMonth() + 1).padStart(2, '0')
      const day = String(today.getDate()).padStart(2, '0')
      metadata.invoiceDate = `${year}/${month}/${day}`
    }

    console.log('Final metadata:', metadata)

    // ✅ CRITICAL: Check if projectId was found in header
    if (!metadata.projectId) {
      console.warn('⚠️ WARNING: 案件管理ID not found in CSV header')
      console.warn(
        '⚠️ Each row MUST have its own 案件管理ID column, or one will be auto-generated per site'
      )
    }

    // Process data rows (Second pass: actual processing)
    for (let i = dataStartIndex; i < csvData.length; i++) {
      const row = csvData[i]
      const values = Object.values(row).map(v => String(v || '').trim())

      // Skip completely empty rows
      if (values.every(v => !v)) {
        this.skippedCount++
        continue
      }

      // Debug: Show all rows being processed
      if (i < dataStartIndex + 15) {
        console.log(`Row ${i}: [${values.slice(0, 11).join(' | ')}]`)
      }

      // Extract fields by position
      const vendorName = values[0] || ''
      const siteName = values[1] || ''
      const dateStr = values[2] || '' // Transaction date (for reference only)
      const salesNo = values[3] || ''
      const itemName = values[4] || ''
      const quantity = values[5] || '1'
      const unit = values[6] || '式'
      const unitPrice = values[7] || ''
      const subtotal = values[8] || ''
      const taxAmount = values[9] || ''
      const totalAmount = values[10] || ''

      // CRITICAL: Skip summary rows that start with 【
      if (itemName.startsWith('【') || vendorName.startsWith('【')) {
        console.log(`Skipping summary row ${i}: ${itemName || vendorName}`)
        this.skippedCount++
        continue
      }

      // Skip header/info rows
      if (this.shouldSkipHeaderRow(vendorName, itemName, values)) {
        console.log(`Skipping header row ${i}: ${vendorName}`)
        this.skippedCount++
        continue
      }

      // Must have item name and it must be meaningful
      if (!itemName || itemName.length < 2) {
        console.log(`Row ${i}: No item name, skipping`)
        this.skippedCount++
        continue
      }

      // Item name should not be just vendor/company name
      if (itemName === vendorName || itemName.includes('クリーン産業')) {
        console.log(`Row ${i}: Item name is company name, skipping`)
        this.skippedCount++
        continue
      }

      // Must have amount
      const cleanAmount = cleanNumber(subtotal) || cleanNumber(totalAmount)
      if (!cleanAmount || cleanAmount === '0') {
        console.log(`Row ${i}: No amount, skipping`)
        this.skippedCount++
        continue
      }

      // Use actual quantity from source data
      const cleanQty = cleanNumber(quantity) || '1'
      const cleanUnitPrice = cleanNumber(unitPrice)

      // Calculate unit price if missing
      let finalUnitPrice = cleanUnitPrice
      if (!finalUnitPrice && cleanAmount && cleanQty) {
        const qty = parseFloat(cleanQty)
        const amt = parseFloat(cleanAmount)
        if (qty > 0) {
          finalUnitPrice = Math.round(amt / qty).toString()
        }
      }

      // CRITICAL: Always use metadata invoice date
      const invoiceDate = metadata.invoiceDate

      console.log(
        `✓ Processing row ${i}: Site="${siteName}" Item="${itemName}" SalesNo="${salesNo}" Qty="${cleanQty}" Unit="${unit}" - ${cleanQty}${unit} × ¥${finalUnitPrice} = ¥${cleanAmount}`
      )

      // ✅ CRITICAL: Check for row-specific 案件管理ID in CSV columns
      let rowProjectId = ''

      // Check common column names for project ID
      if (row['案件管理ID']) {
        rowProjectId = String(row['案件管理ID']).trim()
      } else if (row['工事番号']) {
        rowProjectId = String(row['工事番号']).trim()
      } else if (row['現場No']) {
        rowProjectId = String(row['現場No']).trim()
      } else if (row['物件No']) {
        rowProjectId = String(row['物件No']).trim()
      }

      // Use priority: row-specific > header metadata > generate by site
      let finalProjectId = ''
      if (rowProjectId) {
        finalProjectId = rowProjectId
        console.log(`  ✅ Using 案件管理ID from row: ${finalProjectId}`)
      } else if (metadata.projectId) {
        finalProjectId = metadata.projectId
        console.log(`  ✅ Using 案件管理ID from header: ${finalProjectId}`)
      } else {
        // Fallback: Generate per site (will be same for all items in same site)
        finalProjectId = `SITE_${siteName.replace(/\s+/g, '_')}`
        console.warn(
          `  ⚠️ No 案件管理ID in CSV, using site-based ID: ${finalProjectId}`
        )
      }

      const masterRow = createMasterRow({
        vendor: 'クリーン産業',
        site: siteName || 'ALLAGI株式会社',
        date: invoiceDate,
        item: itemName,
        qty: cleanQty,
        unit: unit || '式',
        price: finalUnitPrice || cleanAmount,
        amount: cleanAmount,
        workNo: salesNo,
        remarks: vendorName !== 'ALLAGI株式会社' ? vendorName : '',
        projectId: finalProjectId, // ✅ Use extracted project ID from CSV
        result: '承認',
      })

      results.push(masterRow)
      this.processedCount++
    }

    this.logComplete(results)
    this.validateResults(results)

    return results
  }

  isAlreadyANDPADFormat(csvData) {
    if (!csvData || csvData.length === 0) return false

    const firstRow = csvData[0]
    const headers = Object.keys(firstRow)

    const andpadColumns = [
      '請求管理ID',
      '取引先',
      '取引設定',
      '担当者(発注側)',
      '請求名',
      '案件管理ID',
    ]

    const matchCount = andpadColumns.filter(col => headers.includes(col)).length

    return matchCount >= 3
  }

  shouldSkipHeaderRow(vendorName, itemName, allValues) {
    // Skip if this is a column header row
    if (vendorName === '業者名' || itemName === '品名') {
      return true
    }

    // Skip if vendor name contains these patterns (but NOT just "ALLAGI株式会社")
    const skipVendorPatterns = [
      '株式会社クリーン産業',
      'TEL',
      'FAX',
      '〒',
      '530-',
      '100-',
      '600-',
      '振込先',
      '登録番号',
      '大阪府',
      '東京都',
      '京都府',
    ]

    for (const pattern of skipVendorPatterns) {
      if (vendorName.includes(pattern)) {
        return true
      }
    }

    // Skip if row contains mostly empty values (likely separator)
    const nonEmptyCount = allValues.filter(v => v && v.trim()).length
    if (nonEmptyCount <= 2) {
      return true
    }

    // Skip rows with these keywords
    const skipKeywords = [
      '請求年月日',
      '今回御請求額',
      '今回取引額',
      'コード',
      '御中',
    ]

    const rowText = allValues.join('|')
    for (const keyword of skipKeywords) {
      if (rowText.includes(keyword)) {
        return true
      }
    }

    return false
  }
}

module.exports = CleanIndustryParser
