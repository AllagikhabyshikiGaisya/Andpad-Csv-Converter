// ============================================
// TAIMAN PARSER - UPDATED WITH 案件管理ID EXTRACTION
// ============================================
const BaseParser = require('./baseParser')
const { createMasterRow, cleanNumber, formatDate } = require('../excelUtils')

class TaimanParser extends BaseParser {
  constructor() {
    super('大萬')
  }

  parse(csvData) {
    this.logStart(csvData)
    const results = []

    // ✅ Extract 案件管理ID from header (first 15 rows)
    let defaultProjectId = ''
    for (let i = 0; i < Math.min(csvData.length, 15); i++) {
      const row = csvData[i]
      const values = Object.values(row)
      const rowText = values.join('|')

      if (
        rowText.includes('案件管理ID') ||
        rowText.includes('工事番号') ||
        rowText.includes('現場No') ||
        rowText.includes('物件No') ||
        rowText.includes('伝票番号')
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
            cleaned !== '物件No' &&
            cleaned !== '伝票番号'
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

    // Process data rows
    for (let i = 0; i < csvData.length; i++) {
      const row = csvData[i]

      const productName = String(row['商品名'] || '').trim()
      const customerName = String(row['請求先名'] || '').trim()
      const purchaseAmount = String(row['仕入金額'] || '').trim()

      if (
        !productName ||
        productName.includes('商品名') ||
        productName.includes('合計')
      ) {
        this.skippedCount++
        continue
      }

      const cleanAmount = cleanNumber(purchaseAmount)
      if (!cleanAmount || cleanAmount === '0') {
        this.skippedCount++
        continue
      }

      // ✅ CRITICAL: Extract 案件管理ID from row
      let rowProjectId = ''
      if (row['案件管理ID']) {
        rowProjectId = String(row['案件管理ID']).trim()
      } else if (row['工事番号']) {
        rowProjectId = String(row['工事番号']).trim()
      } else if (row['伝票番号']) {
        rowProjectId = String(row['伝票番号']).trim()
      } else if (row['現場No']) {
        rowProjectId = String(row['現場No']).trim()
      } else if (row['物件No']) {
        rowProjectId = String(row['物件No']).trim()
      }

      // Priority: row > header > error
      let finalProjectId = ''
      if (rowProjectId) {
        finalProjectId = rowProjectId
        if (i < 5) {
          console.log(
            `  ✅ Row ${i}: Using 案件管理ID from row: ${finalProjectId}`
          )
        }
      } else if (defaultProjectId) {
        finalProjectId = defaultProjectId
        if (i < 5) {
          console.log(
            `  ✅ Row ${i}: Using 案件管理ID from header: ${finalProjectId}`
          )
        }
      } else {
        finalProjectId = `MISSING_ID_TAIMAN_ROW${i}`
        console.error(`  ❌ ERROR Row ${i}: No 案件管理ID in CSV`)
      }

      const masterRow = createMasterRow({
        vendor: customerName || '大萬',
        site: '',
        date: formatDate(row['出荷日'] || ''),
        item: productName,
        qty: cleanNumber(row['数量'] || '') || '1',
        unit: row['単位'] || '個',
        price: cleanNumber(row['仕入単価'] || ''),
        amount: cleanAmount,
        workNo: row['商品コード'] || '',
        remarks: row['備考'] || '',
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
}

module.exports = TaimanParser
