// ============================================
// TAKABISHI PARSER - UPDATED WITH 案件管理ID EXTRACTION
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

    // Extract metadata from early rows
    const metadata = this.extractMetadata(csvData, 15)

    // ✅ Look for 案件管理ID in header (first 15 rows)
    let defaultProjectId = ''
    for (let i = 0; i < Math.min(csvData.length, 15); i++) {
      const row = csvData[i]
      const values = Object.values(row)
      const rowText = values.join('|')

      if (
        rowText.includes('案件管理ID') ||
        rowText.includes('工事番号') ||
        rowText.includes('現場No') ||
        rowText.includes('物件No')
      ) {
        console.log(`✓ Found project ID keyword in row ${i}`)

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
            if (
              !cleaned.match(/^\d{4}\/\d{1,2}\/\d{1,2}$/) &&
              !cleaned.includes('請求') &&
              !cleaned.includes('株式会社') &&
              !cleaned.includes('TEL')
            ) {
              defaultProjectId = cleaned
              console.log(`✅ Found 案件管理ID in header: ${defaultProjectId}`)
              break
            }
          }
        }

        if (!defaultProjectId && i + 1 < csvData.length) {
          const nextRow = csvData[i + 1]
          const nextValues = Object.values(nextRow)
          for (const val of nextValues) {
            const cleaned = String(val).trim()
            if (
              cleaned &&
              cleaned.length > 0 &&
              !cleaned.match(/^\d{4}\/\d{1,2}\/\d{1,2}$/)
            ) {
              defaultProjectId = cleaned
              console.log(
                `✅ Found 案件管理ID in next row: ${defaultProjectId}`
              )
              break
            }
          }
        }

        if (defaultProjectId) break
      }

      // Check column data
      if (
        row['案件管理ID'] &&
        String(row['案件管理ID']).trim() !== '案件管理ID'
      ) {
        defaultProjectId = String(row['案件管理ID']).trim()
        console.log(`✅ Found 案件管理ID in column: ${defaultProjectId}`)
        break
      }
      if (row['工事番号'] && String(row['工事番号']).trim() !== '工事番号') {
        defaultProjectId = String(row['工事番号']).trim()
        console.log(`✅ Found 工事番号 in column: ${defaultProjectId}`)
        break
      }
    }

    if (!defaultProjectId) {
      console.warn('⚠️ WARNING: 案件管理ID not found in CSV header')
      console.warn('⚠️ Will check each data row for 案件管理ID column')
    }

    // Find data start by looking for header patterns
    const headerPatterns = ['品名', '金額', '数量', '単価']
    let dataStartIndex = this.findDataStart(csvData, headerPatterns, 20)

    if (dataStartIndex === -1) {
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

      // ✅ CRITICAL: Extract 案件管理ID from row
      let rowProjectId = ''
      if (row['案件管理ID']) {
        rowProjectId = String(row['案件管理ID']).trim()
      } else if (row['工事番号']) {
        rowProjectId = String(row['工事番号']).trim()
      } else if (row['現場No']) {
        rowProjectId = String(row['現場No']).trim()
      } else if (row['物件No']) {
        rowProjectId = String(row['物件No']).trim()
      }

      // Priority: row > header > error
      let finalProjectId = ''
      if (rowProjectId) {
        finalProjectId = rowProjectId
        if (i < dataStartIndex + 3) {
          console.log(`  ✅ Using 案件管理ID from row: ${finalProjectId}`)
        }
      } else if (defaultProjectId) {
        finalProjectId = defaultProjectId
        if (i < dataStartIndex + 3) {
          console.log(`  ✅ Using 案件管理ID from header: ${finalProjectId}`)
        }
      } else {
        finalProjectId = `MISSING_ID_${site || 'UNKNOWN'}_ROW${i}`.replace(
          /\s+/g,
          '_'
        )
        console.error(`  ❌ ERROR Row ${i}: No 案件管理ID in CSV`)
      }

      const masterRow = createMasterRow({
        vendor: '高菱管理',
        site: site,
        date: formatDate(date),
        item: itemName,
        qty: qty,
        unit: unit || '式',
        price: price || cleanAmount,
        amount: cleanAmount,
        workNo: '',
        remarks: '',
        projectId: finalProjectId, // ✅ From CSV
        result: '承認',
      })

      results.push(masterRow)
      this.processedCount++
    }

    this.logComplete(results)
    this.validateResults(results)
    return results
  }

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
