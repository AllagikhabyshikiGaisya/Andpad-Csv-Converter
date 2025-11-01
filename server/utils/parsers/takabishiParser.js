const BaseParser = require('./baseParser')
const { createMasterRow, cleanNumber, formatDate } = require('../excelUtils')

class TakabishiParser extends BaseParser {
  constructor() {
    super('髙菱管理')
  }

  parse(csvData) {
    this.logStart(csvData)
    const results = []

    // Extract metadata from early rows
    const metadata = this.extractMetadata(csvData, 15)

    // Find data start by looking for header patterns
    const headerPatterns = ['品名', '金額', '数量', '単価']
    let dataStartIndex = this.findDataStart(csvData, headerPatterns, 20)

    if (dataStartIndex === -1) {
      // Fallback: start after first 10 rows
      dataStartIndex = 10
      console.log('⚠️ Header not found, using fallback start index')
    }

    console.log(`Starting data processing from row ${dataStartIndex}`)

    // Process data rows
    for (let i = dataStartIndex; i < csvData.length; i++) {
      const row = csvData[i]
      const values = Object.values(row).map(v => String(v || '').trim())

      // Skip empty rows
      if (values.every(v => !v)) {
        this.skippedCount++
        continue
      }

      const firstCol = values[0] || ''

      // Skip summary/header rows
      if (this.shouldSkipRow(firstCol)) {
        console.log(`Skipping row ${i}: ${firstCol}`)
        this.skippedCount++
        continue
      }

      // Extract item name and amount
      let itemName = ''
      let amount = ''
      let qty = '1'
      let unit = ''
      let price = ''
      let site = metadata.siteName || ''
      let date = metadata.invoiceDate || ''

      // Try to find amount (usually last numeric column)
      for (let j = values.length - 1; j >= 0; j--) {
        const cleanAmt = cleanNumber(values[j])
        if (cleanAmt && cleanAmt !== '0') {
          amount = values[j]
          break
        }
      }

      if (!amount) {
        this.skippedCount++
        continue
      }

      // Find item name (usually longest text field)
      for (const val of values) {
        if (val && val.length > 2 && !val.match(/^\d+$/)) {
          if (!itemName || val.length > itemName.length) {
            itemName = val
          }
        }
      }

      if (!itemName) {
        this.skippedCount++
        continue
      }

      const cleanAmount = cleanNumber(amount)
      if (!cleanAmount || cleanAmount === '0') {
        this.skippedCount++
        continue
      }

      console.log(`✓ Processing row ${i}: ${itemName} - ¥${cleanAmount}`)

      const masterRow = createMasterRow({
        vendor: '髙菱管理',
        site: site,
        date: formatDate(date),
        item: itemName,
        qty: qty,
        unit: unit || '式',
        price: price || cleanAmount,
        amount: cleanAmount,
        workNo: '',
        remarks: '',
      })

      results.push(masterRow)
      this.processedCount++
    }

    this.logComplete(results)
    this.validateResults(results)
    return results
  }

  /**
   * Determine if a row should be skipped
   */
  shouldSkipRow(firstCol) {
    if (!firstCol) return false

    const skipPatterns = [
      '合計',
      '小計',
      '消費税',
      '請求書',
      '株式会社',
      '御中',
      'TEL',
      'FAX',
      '〒',
      '振込先',
      '登録番号',
      '品名',
      '金額',
      '数量',
    ]

    return skipPatterns.some(pattern => firstCol.includes(pattern))
  }
}

module.exports = TakabishiParser
