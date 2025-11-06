// ============================================
// NAKAZAWA KENHAN PARSER - UPDATED WITH 案件管理ID EXTRACTION
// ============================================
const BaseParser = require('./baseParser')
const { createMasterRow, cleanNumber, formatDate } = require('../excelUtils')

class NakazawaKenhanParser extends BaseParser {
  constructor() {
    super('ナカザワ建販')
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
        rowText.includes('現場コード')
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
            cleaned !== '現場コード'
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

      const siteName = String(row['現場名'] || '').trim()
      const productName = String(row['商品名'] || '').trim()
      const salesAmount = String(row['売上金額'] || '').trim()

      if (!siteName && !productName) {
        this.skippedCount++
        continue
      }

      if (siteName.includes('現場名') || productName.includes('商品名')) {
        this.skippedCount++
        continue
      }

      const cleanAmount = cleanNumber(salesAmount)
      if (!cleanAmount || cleanAmount === '0') {
        this.skippedCount++
        continue
      }

      const makerName = String(row['メーカー名'] || '').trim()
      const spec = String(row['規格'] || '').trim()
      const productCode = String(row['商品コード'] || '').trim()

      let itemDescription = productName
      if (makerName) itemDescription = `${makerName} ${productName}`
      if (spec) itemDescription += ` [${spec}]`
      if (productCode) itemDescription += ` (${productCode})`

      const quantity = String(row['数量'] || '').trim()
      const unitPrice = String(row['売上単価'] || '').trim()
      const formattedDate = formatDate(String(row['売上日'] || '').trim())

      let finalUnitPrice = cleanNumber(unitPrice)
      if (!finalUnitPrice && quantity && cleanAmount) {
        const qty = parseFloat(cleanNumber(quantity)) || 1
        const amt = parseFloat(cleanAmount) || 0
        if (qty > 0) finalUnitPrice = Math.round(amt / qty).toString()
      }

      const pieceCount = String(row['個数'] || '').trim()
      const finalQty = cleanNumber(quantity) || cleanNumber(pieceCount) || '1'
      const finalUnit = String(row['売上単価単位名'] || '').trim() || '個'

      const remarksParts = [
        row['売上Ｎｏ'] ? `伝票:${row['売上Ｎｏ']}` : '',
        row['区分名'] ? `区分:${row['区分名']}` : '',
        row['担当者名'] ? `担当:${row['担当者名']}` : '',
        row['入数'] && row['入数'] !== '0' ? `入数:${row['入数']}` : '',
        row['備考'] || '',
      ].filter(Boolean)

      // ✅ CRITICAL: Extract 案件管理ID from row
      let rowProjectId = ''
      if (row['案件管理ID']) {
        rowProjectId = String(row['案件管理ID']).trim()
      } else if (row['工事番号']) {
        rowProjectId = String(row['工事番号']).trim()
      } else if (row['現場コード']) {
        rowProjectId = String(row['現場コード']).trim()
      } else if (row['現場No']) {
        rowProjectId = String(row['現場No']).trim()
      } else if (row['物件No']) {
        rowProjectId = String(row['物件No']).trim()
      }

      // Priority: row > header > error
      let finalProjectId = ''
      if (rowProjectId) {
        finalProjectId = rowProjectId
        if (this.processedCount < 5) {
          console.log(
            `  ✅ Row ${i}: Using 案件管理ID from row: ${finalProjectId}`
          )
        }
      } else if (defaultProjectId) {
        finalProjectId = defaultProjectId
        if (this.processedCount < 5) {
          console.log(
            `  ✅ Row ${i}: Using 案件管理ID from header: ${finalProjectId}`
          )
        }
      } else {
        finalProjectId = `MISSING_ID_${siteName || 'UNKNOWN'}_ROW${i}`.replace(
          /\s+/g,
          '_'
        )
        console.error(`  ❌ ERROR Row ${i}: No 案件管理ID in CSV`)
      }

      const masterRow = createMasterRow({
        vendor: 'ナカザワ建販',
        site: siteName,
        date: formattedDate,
        item: itemDescription.trim(),
        qty: finalQty,
        unit: finalUnit,
        price: finalUnitPrice,
        amount: cleanAmount,
        workNo: productCode || row['売上Ｎｏ'] || '',
        remarks: remarksParts.join(' '),
        projectId: finalProjectId, // ✅ From CSV
        result: '承認',
      })

      results.push(masterRow)
      this.processedCount++

      if (this.processedCount <= 5) {
        console.log(
          `  ✓ Row ${i}: ${siteName} - ${productName} - ¥${salesAmount}`
        )
      }
    }

    this.logComplete(results)
    this.validateResults(results)
    return results
  }
}

module.exports = NakazawaKenhanParser
