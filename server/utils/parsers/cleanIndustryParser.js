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
    }

    // Extract metadata from header rows
    for (let i = 0; i < Math.min(csvData.length, 15); i++) {
      const row = csvData[i]
      const values = Object.values(row)
      const rowText = values.join('|')

      if (rowText.includes('ALLAGI')) {
        metadata.clientName = 'ALLAGI株式会社'
      }

      // Extract invoice date
      const dateMatch = rowText.match(/(\d{4})\/(\d{1,2})\/(\d{1,2})/)
      if (dateMatch) {
        metadata.invoiceDate = `${dateMatch[1]}/${dateMatch[2]}/${dateMatch[3]}`
      }
    }

    console.log('Extracted metadata:', metadata)

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

    // クリーン産業 CSV format (based on logs):
    // Column positions: [0:業者名, 1:現場名, 2:月日, 3:売上No, 4:品名, 5:数量, 6:単位, 7:単価, 8:小計, 9:消費税, 10:金額]

    // Process data rows
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
      const dateStr = values[2] || ''
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

      // Must have item name and it must be meaningful (not just vendor name)
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

      // FIXED: Use actual quantity from source data
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

      // Parse date properly
      let formattedDate = ''
      if (dateStr && dateStr.match(/^\d{1,2}\/\d{1,2}$/)) {
        // Format: "7/1" or "07/01"
        const parts = dateStr.split('/')
        const year = metadata.invoiceDate
          ? metadata.invoiceDate.split('/')[0]
          : new Date().getFullYear()
        formattedDate = `${year}/${parts[0]}/${parts[1]}`
      } else if (dateStr && dateStr.match(/^\d{4}\/\d{1,2}\/\d{1,2}$/)) {
        // Already in correct format
        formattedDate = dateStr
      } else {
        formattedDate = formatDate(dateStr) || metadata.invoiceDate
      }

      console.log(
        `✓ Processing row ${i}: Site="${siteName}" Item="${itemName}" SalesNo="${salesNo}" Qty="${cleanQty}" Unit="${unit}" - ${cleanQty}${unit} × ¥${finalUnitPrice} = ¥${cleanAmount}`
      )

      const masterRow = createMasterRow({
        vendor: 'クリーン産業',
        site: siteName || 'ALLAGI株式会社',
        date: formattedDate,
        item: itemName,
        qty: cleanQty, // FIXED: Now uses actual quantity
        unit: unit || '式',
        price: finalUnitPrice || cleanAmount,
        amount: cleanAmount,
        workNo: salesNo, // This populates 請求納品明細備考
        remarks: vendorName !== 'ALLAGI株式会社' ? vendorName : '', // Additional info
        projectId: '', // Let system generate
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

  // CRITICAL: Separate method to detect header/info rows
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
