// ============================================
// NANSEI PARSER - UPDATED WITH 案件管理ID EXTRACTION
// ============================================
const BaseParser = require('./baseParser')
const { createMasterRow, cleanNumber, formatDate } = require('../excelUtils')

class NanseiParser extends BaseParser {
  constructor() {
    super('ナンセイ')
  }

  parse(csvData) {
    this.logStart(csvData)
    const results = []

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

    // Process data rows
    for (let i = 0; i < csvData.length; i++) {
      const row = csvData[i]

      const customerName = String(row['取引先名'] || '').trim()
      const transactionAmount = String(
        row['今回取引額(税抜)'] || row['今回取引額'] || ''
      ).trim()

      if (!customerName || customerName.includes('取引先名')) {
        this.skippedCount++
        continue
      }

      const cleanAmount = cleanNumber(transactionAmount)
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
        finalProjectId = `MISSING_ID_${customerName}_ROW${i}`.replace(
          /\s+/g,
          '_'
        )
        console.error(`  ❌ ERROR Row ${i}: No 案件管理ID in CSV`)
      }

      const masterRow = createMasterRow({
        vendor: customerName,
        site: row['振込銀行名'] || '',
        date: formatDate(row['請求日付'] || ''),
        item: '請求',
        qty: '1',
        unit: '式',
        price: cleanAmount,
        amount: cleanAmount,
        workNo: row['請求番号'] || '',
        remarks: row['取引先CD'] || '',
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

module.exports = NanseiParser
