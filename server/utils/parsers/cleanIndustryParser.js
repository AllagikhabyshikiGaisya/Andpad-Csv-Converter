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

      const dateMatch = rowText.match(/(\d{4})\/(\d{1,2})\/(\d{1,2})/)
      if (dateMatch) {
        metadata.invoiceDate = `${dateMatch[1]}/${dateMatch[2]}/${dateMatch[3]}`
      }
    }

    console.log('Extracted metadata:', metadata)

    // Find data start - look for rows with actual item data
    let dataStartIndex = -1
    for (let i = 0; i < Math.min(csvData.length, 25); i++) {
      const row = csvData[i]
      const values = Object.values(row)

      // Look for rows that have numeric amounts and item descriptions
      const hasNumericData = values.some(v => {
        const cleaned = cleanNumber(String(v))
        return cleaned && parseFloat(cleaned) > 0
      })

      const hasItemData = values.some(v => {
        const str = String(v || '').trim()
        return (
          str.length > 3 &&
          !str.match(/^[\d\/\-]+$/) &&
          !str.includes('請求') &&
          !str.includes('株式会社')
        )
      })

      if (hasNumericData && hasItemData) {
        dataStartIndex = i
        console.log(`✓ Data starts at row ${i}`)
        break
      }
    }

    if (dataStartIndex === -1) {
      // Fallback: start after header section
      dataStartIndex = 7
      console.log('⚠️ Could not detect data start, using row 7')
    }

    // Process data rows
    for (let i = dataStartIndex; i < csvData.length; i++) {
      const row = csvData[i]
      const values = Object.values(row).map(v => String(v || '').trim())

      if (values.every(v => !v)) {
        this.skippedCount++
        continue
      }

      const firstCol = values[0] || ''
      if (this.shouldSkipRow(firstCol)) {
        this.skippedCount++
        continue
      }

      // Find date (usually in first few columns)
      let date = metadata.invoiceDate
      for (let j = 0; j < Math.min(values.length, 3); j++) {
        if (values[j].match(/\d{4}\/\d{1,2}\/\d{1,2}/)) {
          date = values[j]
          break
        }
      }

      // Find amount (look for largest numeric value)
      let amount = ''
      let maxAmount = 0
      for (const val of values) {
        const cleaned = cleanNumber(val)
        if (cleaned) {
          const num = parseFloat(cleaned)
          if (num > maxAmount && num < 100000000) {
            maxAmount = num
            amount = cleaned
          }
        }
      }

      if (!amount || amount === '0') {
        this.skippedCount++
        continue
      }

      // Find item name (longest non-numeric text)
      let itemName = ''
      let maxLength = 0
      for (const val of values) {
        const str = String(val).trim()
        if (
          str.length > maxLength &&
          str.length > 2 &&
          !str.match(/^[\d,¥]+$/) &&
          !str.match(/^\d{4}\/\d{1,2}/)
        ) {
          maxLength = str.length
          itemName = str
        }
      }

      if (!itemName) {
        itemName = '資材'
      }

      // Find quantity
      let qty = '1'
      for (const val of values) {
        const cleaned = cleanNumber(val)
        if (
          cleaned &&
          parseFloat(cleaned) > 0 &&
          parseFloat(cleaned) < 1000 &&
          cleaned !== amount
        ) {
          qty = cleaned
          break
        }
      }

      console.log(`✓ Processing row ${i}: ${itemName} - ¥${amount}`)

      const masterRow = createMasterRow({
        vendor: 'クリーン産業',
        site: metadata.siteName || '',
        date: formatDate(date),
        item: itemName,
        qty: qty,
        unit: '㎥',
        price: '',
        amount: amount,
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

  shouldSkipRow(firstCol) {
    if (!firstCol) return false

    const skipPatterns = [
      '【',
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
      '大阪',
      '東京',
      '京都',
      '530-',
      '100-',
      '600-',
      '請求年月日',
      'コード',
      '今回御請求額',
      '今回取引額',
    ]

    return skipPatterns.some(pattern => firstCol.includes(pattern))
  }
}

module.exports = CleanIndustryParser
