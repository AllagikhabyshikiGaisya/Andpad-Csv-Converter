// ============================================
// OMEGA JAPAN PARSER - UPDATED WITH 案件管理ID EXTRACTION
// ============================================
const BaseParser = require('./baseParser')
const { createMasterRow, cleanNumber, formatDate } = require('../excelUtils')

class OmegaJapanParser extends BaseParser {
  constructor() {
    super('オメガジャパン')
  }

  parse(csvData) {
    this.logStart(csvData)

    // Detect format
    const firstRowKeys = Object.keys(csvData[0] || {})

    // Check if headers are meaningful (not auto-generated like _1, _2, etc.)
    const hasMeaningfulHeaders = firstRowKeys.some(
      key =>
        key &&
        key.length > 2 &&
        !key.startsWith('_') &&
        !key.match(/^[一二三四五六七八九]+$/) // Japanese numbers
    )

    // Also check if we have proper column names like '品名', '金額', etc.
    const hasProperColumns = firstRowKeys.some(
      k =>
        k.includes('品名') ||
        k.includes('金額') ||
        k.includes('単価') ||
        k.includes('数量')
    )

    if (hasMeaningfulHeaders && hasProperColumns) {
      console.log('✓ Detected: CSV FORMAT (with proper headers)')
      return this.parseCSVFormat(csvData)
    } else {
      console.log('✓ Detected: INVOICE FORMAT (no proper headers)')
      return this.parseInvoiceFormat(csvData)
    }
  }

  parseCSVFormat(csvData) {
    const results = []
    const metadata = this.extractMetadata(csvData)
    let siteName = metadata.siteName
    let invoiceDate = metadata.invoiceDate

    // ✅ Extract 案件管理ID from header (first 30 rows)
    let defaultProjectId = ''
    for (let i = 0; i < Math.min(csvData.length, 30); i++) {
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
              !cleaned.includes('ALLAGI')
            ) {
              defaultProjectId = cleaned
              console.log(`✅ Found 案件管理ID: ${defaultProjectId}`)
              break
            }
          }
        }
        if (defaultProjectId) break
      }

      if (
        row['案件管理ID'] &&
        String(row['案件管理ID']).trim() !== '案件管理ID'
      ) {
        defaultProjectId = String(row['案件管理ID']).trim()
        console.log(`✅ Found 案件管理ID in column: ${defaultProjectId}`)
        break
      }
    }

    if (!defaultProjectId) {
      console.warn('⚠️ WARNING: 案件管理ID not found in CSV header')
    }

    // Find data start
    let dataStartIndex = -1
    for (let i = 0; i < Math.min(csvData.length, 30); i++) {
      const row = csvData[i]
      const values = Object.values(row)
      const rowText = values.join('|')

      if (
        values.some(v => String(v).match(/^[箱本缶個枚式セットボトル㎥]$/)) ||
        rowText.includes('外断熱') ||
        rowText.includes('塗装工事')
      ) {
        dataStartIndex = i
        break
      }
    }

    if (dataStartIndex === -1) dataStartIndex = 15

    let currentSection = ''

    for (let i = dataStartIndex; i < csvData.length; i++) {
      const row = csvData[i]
      const values = Object.values(row).map(v => String(v || '').trim())

      if (values.every(v => !v)) {
        this.skippedCount++
        continue
      }

      const colA = values[0] || ''
      const colB = values[1] || ''
      const colC = values[2] || ''
      const colD = values[3] || ''
      const colE = values[4] || ''
      const colF = values[5] || ''
      const colG = values[6] || ''
      const colH = values[7] || ''
      const colI = values[8] || ''

      // Section header
      if (
        colA.includes('工事部') ||
        colA.includes('オプション') ||
        colA === '内訳'
      ) {
        currentSection = colA
        this.skippedCount++
        continue
      }

      // Skip header rows
      if (
        colA.includes('ALLAGI') ||
        colA.includes('御中') ||
        colA.includes('請求書') ||
        colA.includes('合計') ||
        colB.includes('オメガジャパン')
      ) {
        this.skippedCount++
        continue
      }

      let itemName = colB || colA
      let specs = [colC, colD, colE].filter(Boolean).join(' ')
      let unit = colF
      let quantity = colG
      let unitPrice = colH
      let amount = colI

      if (!itemName) {
        this.skippedCount++
        continue
      }

      const cleanAmount = cleanNumber(amount)
      if (!cleanAmount || cleanAmount === '0') {
        this.skippedCount++
        continue
      }

      const formattedDate = formatDate(invoiceDate)

      let finalUnitPrice = cleanNumber(unitPrice)
      if (!finalUnitPrice && quantity && cleanAmount) {
        const qty = parseFloat(cleanNumber(quantity)) || 1
        const amt = parseFloat(cleanAmount) || 0
        if (qty > 0) finalUnitPrice = Math.round(amt / qty).toString()
      }

      let fullItemName = itemName
      if (specs) fullItemName += ` ${specs}`

      const isDiscount =
        fullItemName.includes('値引') || fullItemName.includes('割引')
      const isShipping = fullItemName.includes('送料')

      const remarksParts = [
        currentSection ? `工事区分:${currentSection}` : '',
        isDiscount ? '値引' : '',
        isShipping ? '配送料' : '',
      ].filter(Boolean)

      // ✅ Extract 案件管理ID from row
      let rowProjectId = ''
      if (row['案件管理ID']) rowProjectId = String(row['案件管理ID']).trim()
      else if (row['工事番号']) rowProjectId = String(row['工事番号']).trim()
      else if (row['現場No']) rowProjectId = String(row['現場No']).trim()

      let finalProjectId =
        rowProjectId ||
        defaultProjectId ||
        `MISSING_ID_${siteName}_ROW${i}`.replace(/\s+/g, '_')

      if (!rowProjectId && !defaultProjectId && i < dataStartIndex + 3) {
        console.error(`  ❌ ERROR Row ${i}: No 案件管理ID in CSV`)
      }

      const masterRow = createMasterRow({
        vendor: 'オメガジャパン',
        site: siteName || 'ALLAGI株式会社',
        date: formattedDate,
        item: fullItemName.trim(),
        qty: cleanNumber(quantity) || '1',
        unit: unit || '式',
        price: finalUnitPrice,
        amount: cleanAmount,
        workNo: '',
        remarks: remarksParts.join(' '),
        projectId: finalProjectId, // ✅ From CSV
        result: '承認',
      })

      results.push(masterRow)
      this.processedCount++

      if (this.processedCount <= 10) {
        console.log(`  ✓ Row ${i}: ${itemName} - ¥${amount}`)
      }
    }

    this.logComplete(results)
    this.validateResults(results)
    return results
  }

  parseInvoiceFormat(csvData) {
    const results = []
    let dataStartIndex = -1
    let currentMonth = 8
    let currentYear = 2025
    let defaultProjectId = ''

    console.log('Parsing INVOICE FORMAT...')

    // Extract period and project ID
    for (let i = 0; i < Math.min(csvData.length, 15); i++) {
      const row = csvData[i]
      const values = Object.values(row)
      const rowText = values.join('|')

      const periodMatch = rowText.match(/(\d{4})年(\d{1,2})月/)
      if (periodMatch) {
        currentYear = parseInt(periodMatch[1])
        currentMonth = parseInt(periodMatch[2])
      }

      // Look for project ID
      if (
        rowText.includes('案件管理ID') ||
        rowText.includes('工事番号') ||
        rowText.includes('現場No')
      ) {
        for (const val of values) {
          const cleaned = String(val).trim()
          if (
            cleaned &&
            cleaned.length > 0 &&
            !cleaned.match(/^\d{4}\/\d{1,2}\/\d{1,2}$/) &&
            !cleaned.includes('請求')
          ) {
            defaultProjectId = cleaned
            console.log(`✅ Found 案件管理ID: ${defaultProjectId}`)
            break
          }
        }
      }

      // Look for header row with "納品日" and "現場名"
      if (rowText.includes('納品日') && rowText.includes('現場名')) {
        dataStartIndex = i + 1
        break
      }
    }

    if (dataStartIndex === -1) dataStartIndex = 10

    for (let i = dataStartIndex; i < csvData.length; i++) {
      const row = csvData[i]
      const values = Object.values(row)

      const nonEmptyValues = values
        .map((v, idx) => ({ idx, val: String(v || '').trim() }))
        .filter(item => item.val !== '')

      if (nonEmptyValues.length === 0) {
        this.skippedCount++
        continue
      }

      let deliveryDate = nonEmptyValues[0]?.val || ''
      let siteName = nonEmptyValues[1]?.val || ''
      let taxIncludedAmount = ''

      for (let j = nonEmptyValues.length - 1; j >= 0; j--) {
        const val = nonEmptyValues[j].val
        const cleaned = cleanNumber(val)
        if (cleaned && cleaned !== '0') {
          taxIncludedAmount = val
          break
        }
      }

      // Skip header/summary rows
      if (
        deliveryDate.includes('納品日') ||
        deliveryDate.includes('合計') ||
        siteName.includes('現場名') ||
        deliveryDate.includes('請求')
      ) {
        this.skippedCount++
        continue
      }

      if (!siteName || !siteName.trim()) {
        this.skippedCount++
        continue
      }

      const cleanAmount = cleanNumber(taxIncludedAmount)
      if (!cleanAmount || cleanAmount === '0') {
        this.skippedCount++
        continue
      }

      // Parse date
      let formattedDate = ''
      const dateMatch = deliveryDate.match(/^(\d{1,2})月(\d{1,2})日?/)
      if (dateMatch) {
        formattedDate = `${currentYear}/${dateMatch[1]}/${dateMatch[2]}`
      } else {
        formattedDate = `${currentYear}/${currentMonth}/1`
      }

      const taxIncluded = parseFloat(cleanAmount) || 0
      const taxExcluded = Math.round(taxIncluded / 1.1)

      // Extract project ID from row
      let rowProjectId = ''
      if (row['案件管理ID']) rowProjectId = String(row['案件管理ID']).trim()
      else if (row['工事番号']) rowProjectId = String(row['工事番号']).trim()

      let finalProjectId =
        rowProjectId ||
        defaultProjectId ||
        `MISSING_ID_${siteName}_ROW${i}`.replace(/\s+/g, '_')

      const masterRow = createMasterRow({
        vendor: 'オメガジャパン',
        site: siteName.trim(),
        date: formattedDate,
        item: `${siteName.trim()} 工事`,
        qty: '1',
        unit: '式',
        price: taxExcluded.toString(),
        amount: taxExcluded.toString(),
        workNo: '',
        remarks: deliveryDate.includes('追加') ? '追加工事' : '',
        projectId: finalProjectId, // ✅ From CSV
        result: '承認',
      })

      masterRow['金額(税込)'] = cleanAmount
      masterRow['単価(税込)'] = cleanAmount

      results.push(masterRow)
      this.processedCount++
    }

    this.logComplete(results)
    this.validateResults(results)
    return results
  }
}

module.exports = OmegaJapanParser
