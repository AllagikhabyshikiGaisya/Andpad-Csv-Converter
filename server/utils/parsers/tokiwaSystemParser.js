const BaseParser = require('./baseParser')
const { createMasterRow, cleanNumber, formatDate } = require('../excelUtils')

class TokiwaSystemParser extends BaseParser {
  constructor() {
    super('トキワシステム')
  }

  parse(csvData) {
    this.logStart(csvData)
    const results = []

    // Extract metadata
    const metadata = this.extractMetadata(csvData, 15)
    console.log('Extracted metadata:', metadata)

    // Find header row
    const headerPatterns = ['品名', '商品', '金額', '数量']
    let dataStartIndex = this.findDataStart(csvData, headerPatterns, 20)

    if (dataStartIndex === -1) {
      console.log('⚠️ Header not found, using fallback')
      dataStartIndex = 10
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

      // Skip header/summary rows
      if (this.shouldSkipRow(firstCol)) {
        console.log(`Skipping row ${i}: ${firstCol}`)
        this.skippedCount++
        continue
      }

      // Extract fields
      let itemName = ''
      let amount = ''
      let qty = '1'
      let unit = ''
      let price = ''

      // Find amount (last numeric value)
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

      // Find item name (longest non-numeric text)
      for (const val of values) {
        if (val && val.length > 2 && !val.match(/^\d+$/) && !cleanNumber(val)) {
          if (!itemName || val.length > itemName.length) {
            itemName = val
          }
        }
      }

      if (!itemName) {
        this.skippedCount++
        continue
      }

      // Look for quantity and unit
      for (let j = 0; j < values.length; j++) {
        const val = values[j]
        const numVal = cleanNumber(val)

        if (numVal && parseFloat(numVal) > 0 && val !== amount) {
          if (!qty || qty === '1') {
            qty = numVal
          } else if (!price) {
            price = numVal
          }
        } else if (val && val.length < 5 && !numVal && val !== itemName) {
          if (!unit) unit = val
        }
      }

      const cleanAmount = cleanNumber(amount)
      if (!cleanAmount || cleanAmount === '0') {
        this.skippedCount++
        continue
      }

      console.log(`✓ Processing row ${i}: ${itemName} - ¥${cleanAmount}`)

      const masterRow = createMasterRow({
        vendor: 'トキワシステム',
        site: metadata.siteName || metadata.projectName || '',
        date: formatDate(metadata.invoiceDate || ''),
        item: itemName,
        qty: qty,
        unit: unit || '個',
        price: price || cleanAmount,
        amount: cleanAmount,
        workNo: '',
        remarks: '',
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
      '商品',
      '金額',
      '数量',
      'トキワ',
    ]

    return skipPatterns.some(pattern => firstCol.includes(pattern))
  }
}

module.exports = TokiwaSystemParser
