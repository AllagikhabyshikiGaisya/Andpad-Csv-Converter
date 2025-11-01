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
        !key.match(/^[一二三四五六七八九十]+$/) // Japanese numbers
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

    // Find data start
    let dataStartIndex = -1
    for (let i = 0; i < Math.min(csvData.length, 30); i++) {
      const row = csvData[i]
      const values = Object.values(row)
      const rowText = values.join('|')

      if (
        values.some(v => String(v).match(/^[箱本缶個枚式セットボトル㎡]$/)) ||
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

    console.log('Parsing INVOICE FORMAT...')
    console.log('Sample rows:')
    for (let i = 0; i < Math.min(5, csvData.length); i++) {
      const values = Object.values(csvData[i])
      console.log(`  Row ${i}:`, values.slice(0, 10))
    }

    // Extract period info from early rows
    for (let i = 0; i < Math.min(csvData.length, 15); i++) {
      const row = csvData[i]
      const values = Object.values(row)
      const rowText = values.join('|')

      const periodMatch = rowText.match(/(\d{4})年(\d{1,2})月/)
      if (periodMatch) {
        currentYear = parseInt(periodMatch[1])
        currentMonth = parseInt(periodMatch[2])
        console.log(`  Found period: ${currentYear}年${currentMonth}月`)
      }

      // Look for header row with "納品日" and "現場名"
      if (rowText.includes('納品日') && rowText.includes('現場名')) {
        dataStartIndex = i + 1
        console.log(`  Found data start at row ${dataStartIndex}`)
        break
      }
    }

    if (dataStartIndex === -1) {
      // If no header found, start after first 10 rows
      dataStartIndex = 10
      console.log(`  No header found, starting at row ${dataStartIndex}`)
    }

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

      // First non-empty value should be delivery date
      let deliveryDate = nonEmptyValues[0]?.val || ''
      // Second should be site name
      let siteName = nonEmptyValues[1]?.val || ''
      // Last numeric value should be amount
      let taxIncludedAmount = ''

      for (let j = nonEmptyValues.length - 1; j >= 0; j--) {
        const val = nonEmptyValues[j].val
        const cleaned = cleanNumber(val)
        if (cleaned && cleaned !== '0') {
          taxIncludedAmount = val
          break
        }
      }

      if (i < dataStartIndex + 5) {
        console.log(
          `  Row ${i}: date="${deliveryDate}", site="${siteName}", amount="${taxIncludedAmount}"`
        )
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
