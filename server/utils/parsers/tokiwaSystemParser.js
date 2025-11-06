// ============================================
// TOKIWA SYSTEM PARSER - UPDATED WITH 案件管理ID EXTRACTION
// ============================================
const BaseParser = require('./baseParser')
const { createMasterRow, cleanNumber, formatDate } = require('../excelUtils')

class TokiwaSystemParser extends BaseParser {
  constructor() {
    super('トキワシステム')
  }

  parse(csvData) {
    this.logStart(csvData)
    const results = []

    // ✅ Extract metadata including project ID
    const metadata = this.extractMetadata(csvData, 15)
    console.log('Extracted metadata:', metadata)

    // ✅ Look for 案件管理ID in metadata/header (first 15 rows)
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
              !cleaned.includes('株式会社')
            ) {
              defaultProjectId = cleaned
              console.log(`✅ Found 案件管理ID in header: ${defaultProjectId}`)
              break
            }
          }
        }
        if (defaultProjectId) break
      }
    }

    // ⚠️ Warning if no project ID found in header
    if (!defaultProjectId) {
      console.warn('⚠️ WARNING: 案件管理ID not found in CSV header')
      console.warn('⚠️ Will check each data row for 案件管理ID column')
    }

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

      // ✅ CRITICAL: Extract 案件管理ID from row columns
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

      // Priority: row-specific > header default > error
      let finalProjectId = ''
      if (rowProjectId) {
        finalProjectId = rowProjectId
        console.log(`  ✅ Using 案件管理ID from row: ${finalProjectId}`)
      } else if (defaultProjectId) {
        finalProjectId = defaultProjectId
        console.log(`  ✅ Using 案件管理ID from header: ${finalProjectId}`)
      } else {
        // ❌ This should not happen - CSV must provide project ID
        const siteName = metadata.siteName || metadata.projectName || 'UNKNOWN'
        finalProjectId = `MISSING_ID_${siteName.replace(/\s+/g, '_')}_ROW${i}`
        console.error(`  ❌ ERROR: No 案件管理ID in CSV for row ${i}`)
        console.error(
          `  ❌ This violates the rule: "原本CSVの案件管理IDを入力"`
        )
      }

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
