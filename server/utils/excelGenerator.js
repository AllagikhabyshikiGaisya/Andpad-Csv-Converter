// ============================================
// EXCEL GENERATOR - FIXED VERSION
// ============================================

const XLSX = require('xlsx') // ⭐ CRITICAL FIX: Added missing import

// Master ANDPAD columns
const MASTER_COLUMNS = [
  '請求管理ID',
  '取引先',
  '取引設定',
  '担当者(発注側)',
  '請求名',
  '案件管理ID',
  '請求納品金額(税抜)',
  '請求納品金額(税込)',
  '現場監督',
  '納品実績日',
  '支払予定日',
  '請求納品明細名',
  '数量',
  '単位',
  '単価(税抜)',
  '単価(税込)',
  '金額(税抜)',
  '金額(税込)',
  '工事種類',
  '課税フラグ',
  '請求納品明細備考',
  '結果',
]

function generateExcel(csvData, mapping) {
  try {
    console.log('=== EXCEL GENERATION START ===')
    console.log('Vendor:', mapping.vendor)
    console.log('Has custom parser:', mapping.customParser)
    console.log('Input rows:', csvData.length)

    let transformedData = []

    if (mapping.customParser === true) {
      console.log('✓ Using custom parser for:', mapping.vendor)
      transformedData = parseWithCustomLogic(csvData, mapping.vendor)

      if (!transformedData || transformedData.length === 0) {
        console.error('Custom parser returned no data')
        throw new Error(
          'カスタムパーサーがデータを返しませんでした。ファイル形式を確認してください。'
        )
      }
    } else {
      if (!mapping.map || Object.keys(mapping.map).length === 0) {
        console.error('No mapping configuration found')
        throw new Error('マッピング設定が見つかりません。')
      }

      console.log('Using standard mapping')
      transformedData = parseWithMapping(csvData, mapping)

      if (!transformedData || transformedData.length === 0) {
        console.error('Standard parser returned no data')
        throw new Error(
          'データ行が見つかりませんでした。ファイル形式を確認してください。'
        )
      }
    }

    console.log('✓ Transformed rows:', transformedData.length)
    console.log('Sample row:', JSON.stringify(transformedData[0], null, 2))

    // Create workbook
    const workbook = XLSX.utils.book_new()
    const worksheet = XLSX.utils.json_to_sheet(transformedData, {
      header: MASTER_COLUMNS,
    })

    // Set column widths
    worksheet['!cols'] = MASTER_COLUMNS.map(col => {
      if (col.includes('ID')) return { wch: 12 }
      if (col.includes('明細名') || col.includes('請求名')) return { wch: 30 }
      if (col.includes('取引先') || col.includes('監督')) return { wch: 20 }
      if (col.includes('日')) return { wch: 12 }
      return { wch: 15 }
    })

    XLSX.utils.book_append_sheet(workbook, worksheet, 'ANDPAD Import')
    const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' })

    console.log('=== EXCEL GENERATION SUCCESS ===')
    return {
      success: true,
      buffer: buffer,
      rowCount: transformedData.length,
    }
  } catch (error) {
    console.error('❌ Excel generation error:', error.message)
    console.error('Stack:', error.stack)
    return {
      success: false,
      error: error.message || '生成に失敗しました',
    }
  }
}

function parseWithCustomLogic(csvData, vendorName) {
  console.log('=== CUSTOM PARSER CALLED ===')
  console.log('Vendor:', vendorName)

  if (vendorName === 'クリーン産業') {
    return parseCleanIndustry(csvData)
  }

  if (vendorName === '三高産業') {
    return parseSankoSangyo(csvData)
  }

  if (vendorName === '北恵株式会社') {
    return parseHokukei(csvData)
  }

  if (vendorName === 'ナンセイ') {
    return parseNansei(csvData)
  }

  if (vendorName === '大萬') {
    return parseTaiman(csvData)
  }

  if (vendorName === '髙菱管理' || vendorName === '高菱管理') {
    return parseTakabishi(csvData)
  }

  if (vendorName === 'オメガジャパン') {
    return parseOmegaJapan(csvData)
  }

  if (vendorName === 'ナカザワ建販') {
    return parseNakazawaKenhan(csvData)
  }

  if (vendorName === 'トキワシステム') {
    return parseTokiwaSystem(csvData)
  }

  console.error('❌ No custom parser found for:', vendorName)
  throw new Error(`カスタムパーサーが見つかりません: ${vendorName}`)
}

// ============================================
// TOKIWA SYSTEM PARSER - IMPROVED VERSION
// ============================================
function parseTokiwaSystem(csvData) {
  const results = []

  console.log('=== TOKIWA SYSTEM PARSER START ===')
  console.log('Input rows:', csvData.length)

  let processedCount = 0
  let skippedCount = 0
  let currentDeliveryDate = ''
  let currentLocation = ''

  // Process each data row
  for (let i = 0; i < csvData.length; i++) {
    const row = csvData[i]
    const values = Object.values(row).map(v => String(v || '').trim())

    // Skip completely empty rows
    if (values.every(v => !v)) {
      skippedCount++
      continue
    }

    const firstValue = values[0] || ''
    const secondValue = values[1] || ''

    // Pattern 1: "8月4日 納品    岸本光平様邸" in first column
    const deliveryPattern1 = /^(\d{1,2}月\d{1,2}日)\s*納品\s+(.+)$/
    const match1 = firstValue.match(deliveryPattern1)

    if (match1) {
      currentDeliveryDate = match1[1]
      currentLocation = match1[2].trim()
      console.log(
        `  Found delivery info at row ${i}: ${currentDeliveryDate} - ${currentLocation}`
      )
      skippedCount++
      continue
    }

    // Pattern 2: Date in first column, location in second column
    const dateOnlyPattern = /^(\d{1,2}月\d{1,2}日)$/
    const match2 = firstValue.match(dateOnlyPattern)

    if (match2 && secondValue && secondValue.includes('納品')) {
      currentDeliveryDate = match2[1]
      currentLocation = secondValue.replace(/納品/g, '').trim()
      console.log(
        `  Found delivery info at row ${i}: ${currentDeliveryDate} - ${currentLocation}`
      )
      skippedCount++
      continue
    }

    // Skip header/company info rows
    if (
      firstValue.includes('請求書') ||
      firstValue.includes('ALLAGI') ||
      firstValue.includes('株式会社') ||
      firstValue.includes('〒') ||
      firstValue.includes('TEL') ||
      firstValue.includes('FAX') ||
      firstValue.includes('登録番号') ||
      firstValue.includes('お振込先') ||
      firstValue.includes('信用金庫') ||
      firstValue.includes('銀行') ||
      firstValue.includes('大阪府') ||
      firstValue.includes('静岡県') ||
      firstValue.includes('品名') ||
      firstValue.includes('品　名') ||
      firstValue.includes('日付') ||
      firstValue.includes('###') ||
      firstValue.includes('小計') ||
      firstValue.includes('合計') ||
      firstValue.includes('ベルコリーヌ') ||
      firstValue.includes('担当') ||
      firstValue.match(/^\d{4}\/\d{1,2}\/\d{1,2}$/) || // Skip date-only rows
      firstValue.match(/^\d+\/\d+$/) // Skip page numbers
    ) {
      skippedCount++
      continue
    }

    // Now try to extract item data
    let itemName = firstValue
    let quantity = ''
    let unit = ''
    let unitPrice = ''
    let amount = ''

    // Look for numeric values and units in the row
    const numericPattern = /^[\d,]+$/
    const unitPattern = /^[本ケース式個セット袋枚台缶]$/

    for (let j = 1; j < values.length; j++) {
      const val = values[j]
      if (!val) continue

      if (unitPattern.test(val)) {
        unit = val
      } else if (numericPattern.test(val.replace(/,/g, ''))) {
        const cleaned = cleanNumber(val)
        if (cleaned && cleaned !== '0') {
          // Collect all numeric values
          if (!quantity) {
            quantity = val
          } else if (!unitPrice) {
            unitPrice = val
          } else if (!amount) {
            amount = val
          }
        }
      }
    }

    // If we found multiple numbers, the last one is usually the amount
    // Work backwards: amount is last number, unitPrice is second-to-last, quantity is third-to-last
    const allNumbers = values
      .map((v, idx) => ({ val: v, idx }))
      .filter(item => {
        const cleaned = cleanNumber(item.val)
        return cleaned && cleaned !== '0'
      })

    if (allNumbers.length >= 1) {
      amount = allNumbers[allNumbers.length - 1].val
      if (allNumbers.length >= 2) {
        unitPrice = allNumbers[allNumbers.length - 2].val
      }
      if (allNumbers.length >= 3) {
        quantity = allNumbers[allNumbers.length - 3].val
      }
    }

    // Validate amount
    const cleanAmount = cleanNumber(amount)
    if (!cleanAmount || cleanAmount === '0') {
      skippedCount++
      continue
    }

    // Skip if no item name
    if (!itemName || itemName === '') {
      skippedCount++
      continue
    }

    // Parse date
    let formattedDate = ''
    if (currentDeliveryDate) {
      const dateMatch = currentDeliveryDate.match(/^(\d{1,2})月(\d{1,2})日$/)
      if (dateMatch) {
        const year = new Date().getFullYear()
        formattedDate = `${year}/${dateMatch[1]}/${dateMatch[2]}`
      }
    }

    // Calculate unit price if not provided
    let finalUnitPrice = cleanNumber(unitPrice)
    if (!finalUnitPrice && quantity && cleanAmount) {
      const qty = parseFloat(cleanNumber(quantity)) || 1
      const amt = parseFloat(cleanAmount) || 0
      if (qty > 0) {
        finalUnitPrice = Math.round(amt / qty).toString()
      }
    }

    // Handle special items
    let isDiscount = itemName.includes('値引') || itemName.includes('割引')
    let isShipping = itemName.includes('送料')

    // Determine final quantity and unit
    let finalQty = cleanNumber(quantity) || '1'
    let finalUnit = unit || (isShipping ? 'ケース' : '本')

    // Build item description
    let itemDescription = itemName
    if (currentLocation && !isDiscount && !isShipping) {
      itemDescription = `${itemName} (${currentLocation})`
    }

    // Build remarks
    const remarksParts = [
      currentDeliveryDate ? `納品日:${currentDeliveryDate}` : '',
      currentLocation ? `現場:${currentLocation}` : '',
      isDiscount ? '値引' : '',
      isShipping ? '配送料' : '',
    ].filter(Boolean)

    const finalRemarks = remarksParts.join(' ')

    // Create master row
    const masterRow = createMasterRow({
      vendor: 'トキワシステム',
      site: currentLocation || 'ALLAGI株式会社',
      date: formattedDate,
      item: itemDescription.trim(),
      qty: finalQty,
      unit: finalUnit,
      price: finalUnitPrice,
      amount: cleanAmount,
      workNo: '',
      remarks: finalRemarks,
    })

    results.push(masterRow)
    processedCount++

    // Log first few items
    if (processedCount <= 10) {
      console.log(
        `  ✓ Row ${i}: ${itemName} - Qty:${finalQty} ${finalUnit} - ¥${amount} (${
          currentLocation || 'no location'
        })`
      )
    }
  }

  console.log('=== TOKIWA SYSTEM PARSING COMPLETE ===')
  console.log(`✓ Processed: ${processedCount} items`)
  console.log(`⊘ Skipped: ${skippedCount} rows`)
  console.log(`Total output rows: ${results.length}`)

  if (results.length === 0) {
    throw new Error(
      '有効なデータ行が見つかりませんでした。ファイル形式を確認してください。'
    )
  }

  return results
}

// ============================================
// OMEGA JAPAN PARSER - IMPROVED VERSION
// ============================================
function parseOmegaJapan(csvData) {
  console.log('=== OMEGA JAPAN PARSER START ===')
  console.log('Input rows:', csvData.length)

  if (csvData.length > 0) {
    console.log('First row keys:', Object.keys(csvData[0]))
    console.log('Total keys:', Object.keys(csvData[0]).length)
  }

  // Detect which format we're dealing with
  const firstRowKeys = Object.keys(csvData[0] || {})
  const hasProperColumns =
    firstRowKeys.length >= 8 &&
    firstRowKeys.some(k => k && k.length > 2 && !k.startsWith('_'))

  if (hasProperColumns) {
    console.log('✓ Detected: CSV FORMAT (with proper columns)')
    return parseOmegaJapanCSV(csvData)
  } else {
    console.log('✓ Detected: INVOICE FORMAT (headerless layout)')
    return parseOmegaJapanInvoice(csvData)
  }
}

// CSV Format Parser
function parseOmegaJapanCSV(csvData) {
  const results = []

  console.log('=== OMEGA JAPAN CSV PARSER START ===')
  console.log('Input rows:', csvData.length)

  let processedCount = 0
  let skippedCount = 0
  let currentSection = ''
  let siteName = ''
  let invoiceDate = ''

  // Extract site name and date from header rows
  for (let i = 0; i < Math.min(csvData.length, 15); i++) {
    const row = csvData[i]
    const values = Object.values(row)
    const firstValue = String(values[0] || '').trim()

    // Look for site name pattern
    if (firstValue.includes('様邸') || firstValue.includes('工事')) {
      siteName = firstValue
      console.log(`  Found site name: ${siteName}`)
    }

    // Look for date
    for (const val of values) {
      const str = String(val).trim()
      if (str.match(/\d{4}\/\d{1,2}\/\d{1,2}/) || str.match(/\d+月\d+日/)) {
        invoiceDate = str
        console.log(`  Found date: ${invoiceDate}`)
        break
      }
    }
  }

  // Find data start
  let dataStartIndex = -1
  for (let i = 0; i < Math.min(csvData.length, 30); i++) {
    const row = csvData[i]
    const values = Object.values(row)
    const rowText = values.join('|')

    if (
      values.some(v => String(v).match(/^[箱本缶個枚式セットボトル㎡]$/)) ||
      rowText.includes('外断熱') ||
      rowText.includes('塗装工事') ||
      rowText.includes('オプション')
    ) {
      dataStartIndex = i
      console.log(`✓ Data area detected at row ${i}`)
      break
    }
  }

  if (dataStartIndex === -1) {
    dataStartIndex = 15
  }

  console.log(`Processing CSV data from row ${dataStartIndex}`)

  // Process each row
  for (let i = dataStartIndex; i < csvData.length; i++) {
    const row = csvData[i]
    const values = Object.values(row).map(v => String(v || '').trim())

    if (values.every(v => !v)) {
      skippedCount++
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

    // Check if section header
    if (
      colA.includes('工事部') ||
      colA.includes('オプション') ||
      colA.includes('納品') ||
      colA === '内訳'
    ) {
      currentSection = colA
      console.log('  Section: ' + currentSection)
      skippedCount++
      continue
    }

    // Skip header/summary rows
    if (
      colA.includes('ALLAGI') ||
      colA.includes('株式会社') ||
      colA.includes('御中') ||
      colA.includes('〒') ||
      colA.includes('請求書') ||
      colA.includes('登録番号') ||
      colA.includes('振込先') ||
      colA.includes('銀行') ||
      colA.includes('合計') ||
      colA.includes('小計') ||
      colA.includes('御請求金') ||
      colA.includes('消費税') ||
      colA.includes('税別合計') ||
      colB.includes('ご請求書') ||
      colB.includes('オメガジャパン')
    ) {
      skippedCount++
      continue
    }

    // Extract item data
    let itemName = colB || colA
    let specs = [colC, colD, colE].filter(Boolean).join(' ')
    let unit = colF
    let quantity = colG
    let unitPrice = colH
    let amount = colI

    // If no item name, skip
    if (!itemName || itemName === '') {
      skippedCount++
      continue
    }

    // Validate amount
    const cleanAmount = cleanNumber(amount)
    if (!cleanAmount || cleanAmount === '0') {
      skippedCount++
      continue
    }

    // Format date
    let formattedDate = formatDate(invoiceDate)

    // Calculate unit price if missing
    let finalUnitPrice = cleanNumber(unitPrice)
    if (!finalUnitPrice && quantity && cleanAmount) {
      const qty = parseFloat(cleanNumber(quantity)) || 1
      const amt = parseFloat(cleanAmount) || 0
      if (qty > 0) {
        finalUnitPrice = Math.round(amt / qty).toString()
      }
    }

    // Build full item description
    let fullItemName = itemName
    if (specs) {
      fullItemName += ` ${specs}`
    }

    // Check if discount or shipping
    let isDiscount =
      fullItemName.includes('値引') || fullItemName.includes('割引')
    let isShipping = fullItemName.includes('送料')

    // Build remarks
    const remarksParts = [
      currentSection ? `工事区分:${currentSection}` : '',
      isDiscount ? '値引' : '',
      isShipping ? '配送料' : '',
    ].filter(Boolean)

    // Create master row
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
    processedCount++

    if (processedCount <= 10) {
      console.log(
        `  ✓ Row ${i}: ${itemName} - Qty:${quantity} ${unit} - ¥${amount}`
      )
    }
  }

  console.log('=== OMEGA JAPAN CSV PARSING COMPLETE ===')
  console.log(`✓ Processed: ${processedCount} items`)
  console.log(`⊘ Skipped: ${skippedCount} rows`)

  if (results.length === 0) {
    throw new Error(
      '有効なデータ行が見つかりませんでした。ファイル形式を確認してください。'
    )
  }

  return results
}

// Invoice Format Parser
function parseOmegaJapanInvoice(csvData) {
  const results = []

  console.log('=== OMEGA JAPAN INVOICE PARSER START ===')
  console.log('Input rows:', csvData.length)

  let dataStartIndex = -1
  let currentMonth = 8
  let currentYear = 2025

  // Extract period info
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

    if (
      rowText.includes('納品日') &&
      rowText.includes('現場名') &&
      rowText.includes('税込計')
    ) {
      dataStartIndex = i + 1
      console.log(
        `✓ Data header found at row ${i}, data starts at ${dataStartIndex}`
      )
      break
    }
  }

  if (dataStartIndex === -1) {
    dataStartIndex = 15
  }

  console.log(`Processing invoice data from row ${dataStartIndex}`)

  let processedCount = 0
  let skippedCount = 0

  for (let i = dataStartIndex; i < csvData.length; i++) {
    const row = csvData[i]
    const values = Object.values(row)

    const nonEmptyValues = values
      .map((v, idx) => ({ idx, val: String(v || '').trim() }))
      .filter(item => item.val !== '')

    if (nonEmptyValues.length === 0) {
      skippedCount++
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

    if (
      deliveryDate.includes('納品日') ||
      deliveryDate.includes('合計') ||
      siteName.includes('合計') ||
      siteName.includes('現場名')
    ) {
      skippedCount++
      continue
    }

    if (!siteName || !siteName.trim()) {
      skippedCount++
      continue
    }

    const cleanAmount = cleanNumber(taxIncludedAmount)
    if (!cleanAmount || cleanAmount === '0') {
      skippedCount++
      continue
    }

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
    processedCount++

    if (processedCount <= 5) {
      console.log(
        `  ✓ Row ${i}: ${siteName} - ¥${taxIncludedAmount} (${formattedDate})`
      )
    }
  }

  console.log('=== OMEGA JAPAN INVOICE PARSING COMPLETE ===')
  console.log(`✓ Processed: ${processedCount} items`)

  if (results.length === 0) {
    throw new Error(
      '有効なデータ行が見つかりませんでした。ファイル形式を確認してください。'
    )
  }

  return results
}

// ============================================
// NAKAZAWA KENHAN PARSER
// ============================================
function parseNakazawaKenhan(csvData) {
  const results = []

  console.log('=== NAKAZAWA KENHAN PARSER START ===')
  console.log('Input rows:', csvData.length)

  if (csvData.length > 0) {
    console.log('First row keys:', Object.keys(csvData[0]))
  }

  let processedCount = 0
  let skippedCount = 0

  for (let i = 0; i < csvData.length; i++) {
    const row = csvData[i]

    const siteName = String(row['現場名'] || '').trim()
    const staffCode = String(row['担当者コード'] || '').trim()
    const staffName = String(row['担当者名'] || '').trim()
    const salesNo = String(row['売上Ｎｏ'] || '').trim()
    const salesDate = String(row['売上日'] || '').trim()
    const remarks = String(row['備考'] || '').trim()
    const categoryName = String(row['区分名'] || '').trim()
    const makerName = String(row['メーカー名'] || '').trim()
    const productCode = String(row['商品コード'] || '').trim()
    const productName = String(row['商品名'] || '').trim()
    const spec = String(row['規格'] || '').trim()
    const packSize = String(row['入数'] || '').trim()
    const pieceCount = String(row['個数'] || '').trim()
    const quantity = String(row['数量'] || '').trim()
    const priceUnitName = String(row['売上単価単位名'] || '').trim()
    const unitPrice = String(row['売上単価'] || '').trim()
    const salesAmount = String(row['売上金額'] || '').trim()

    // Skip empty rows
    if (!siteName && !productName) {
      skippedCount++
      continue
    }

    // Skip header rows
    if (siteName.includes('現場名') || productName.includes('商品名')) {
      skippedCount++
      continue
    }

    // Validate amount
    const cleanAmount = cleanNumber(salesAmount)
    if (!cleanAmount || cleanAmount === '0') {
      skippedCount++
      continue
    }

    // Build item description
    let itemDescription = productName
    if (makerName) {
      itemDescription = `${makerName} ${productName}`
    }
    if (spec) {
      itemDescription += ` [${spec}]`
    }
    if (productCode) {
      itemDescription += ` (${productCode})`
    }

    // Format date
    const formattedDate = formatDate(salesDate)

    // Calculate unit price if not provided
    let finalUnitPrice = cleanNumber(unitPrice)
    if (!finalUnitPrice && quantity && cleanAmount) {
      const qty = parseFloat(cleanNumber(quantity)) || 1
      const amt = parseFloat(cleanAmount) || 0
      if (qty > 0) {
        finalUnitPrice = Math.round(amt / qty).toString()
      }
    }

    // Determine quantity and unit
    let finalQty = cleanNumber(quantity) || cleanNumber(pieceCount) || '1'
    let finalUnit = priceUnitName || '個'

    // Build remarks
    const remarksParts = [
      salesNo ? `伝票:${salesNo}` : '',
      categoryName ? `区分:${categoryName}` : '',
      staffName ? `担当:${staffName}` : '',
      packSize && packSize !== '0' ? `入数:${packSize}` : '',
      remarks,
    ].filter(Boolean)

    const finalRemarks = remarksParts.join(' ')

    // Create master row
    const masterRow = createMasterRow({
      vendor: 'ナカザワ建販',
      site: siteName,
      date: formattedDate,
      item: itemDescription.trim(),
      qty: finalQty,
      unit: finalUnit,
      price: finalUnitPrice,
      amount: cleanAmount,
      workNo: productCode || salesNo,
      remarks: finalRemarks,
    })

    results.push(masterRow)
    processedCount++

    if (processedCount <= 5) {
      console.log(
        `  ✓ Row ${i}: ${siteName} - ${productName} - ¥${salesAmount}`
      )
    }
  }

  console.log('=== NAKAZAWA KENHAN PARSING COMPLETE ===')
  console.log(`✓ Processed: ${processedCount} items`)
  console.log(`⊘ Skipped: ${skippedCount} rows`)

  if (results.length === 0) {
    throw new Error(
      '有効なデータ行が見つかりませんでした。ファイル形式を確認してください。'
    )
  }

  return results
}

// ============================================
// OTHER PARSERS
// ============================================

function parseTakabishi(csvData) {
  const results = []
  console.log('=== TAKABISHI PARSER START ===')
  console.log('Input rows:', csvData.length)

  let dataStartIndex = 0
  const headerPatterns = ['請求日付', '支店コード', '得意先番号', '請求全額']

  for (let i = 0; i < Math.min(csvData.length, 20); i++) {
    const row = csvData[i]
    const rowText = [...Object.keys(row), ...Object.values(row)].join('|')
    const matchedPatterns = headerPatterns.filter(p => rowText.includes(p))

    if (matchedPatterns.length >= 3) {
      dataStartIndex = i + 1
      console.log(`✓ Header found at row ${i}`)
      break
    }
  }

  let processedCount = 0
  let skippedCount = 0

  for (let i = dataStartIndex; i < csvData.length; i++) {
    const row = csvData[i]
    const values = Object.values(row)

    let invoiceType = String(values[0] || '').trim()
    let groupFlag = String(values[1] || '').trim()
    let invoiceDate = String(values[2] || '').trim()
    let branchCode = String(values[3] || '').trim()
    let customerCode = String(values[4] || '').trim()
    let locationCode = String(values[5] || '').trim()
    let requestAmount = String(values[11] || '').trim()
    let requestNumber = String(values[17] || '').trim()

    if (invoiceType !== 'INV' || groupFlag !== 'H') {
      skippedCount++
      continue
    }

    const cleanAmount = cleanNumber(requestAmount)
    if (!cleanAmount || cleanAmount === '0') {
      skippedCount++
      continue
    }

    const masterRow = createMasterRow({
      vendor: `髙菱管理 ${customerCode}`,
      site: locationCode,
      date: formatDate(invoiceDate),
      item: `請求 (配置先: ${locationCode})`,
      qty: '1',
      unit: '式',
      price: cleanAmount,
      amount: cleanAmount,
      workNo: requestNumber,
      remarks: branchCode ? `支店:${branchCode}` : '',
    })

    results.push(masterRow)
    processedCount++
  }

  console.log('=== TAKABISHI PARSING COMPLETE ===')
  console.log(`✓ Processed: ${processedCount} items`)

  if (results.length === 0) {
    throw new Error('有効なデータ行が見つかりませんでした。')
  }

  return results
}

function parseTaiman(csvData) {
  const results = []
  console.log('=== TAIMAN PARSER START ===')
  console.log('Input rows:', csvData.length)

  let processedCount = 0
  let skippedCount = 0

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
      skippedCount++
      continue
    }

    const cleanAmount = cleanNumber(purchaseAmount)
    if (!cleanAmount || cleanAmount === '0') {
      skippedCount++
      continue
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
    })

    results.push(masterRow)
    processedCount++
  }

  console.log('=== TAIMAN PARSING COMPLETE ===')
  console.log(`✓ Processed: ${processedCount} items`)

  if (results.length === 0) {
    throw new Error('有効なデータ行が見つかりませんでした。')
  }

  return results
}

function parseNansei(csvData) {
  const results = []
  console.log('=== NANSEI PARSER START ===')
  console.log('Input rows:', csvData.length)

  let processedCount = 0
  let skippedCount = 0

  for (let i = 0; i < csvData.length; i++) {
    const row = csvData[i]

    const customerName = String(row['取引先名'] || '').trim()
    const transactionAmount = String(
      row['今回取引額(税抜)'] || row['今回取引額'] || ''
    ).trim()

    if (!customerName || customerName.includes('取引先名')) {
      skippedCount++
      continue
    }

    const cleanAmount = cleanNumber(transactionAmount)
    if (!cleanAmount || cleanAmount === '0') {
      skippedCount++
      continue
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
    })

    results.push(masterRow)
    processedCount++
  }

  console.log('=== NANSEI PARSING COMPLETE ===')
  console.log(`✓ Processed: ${processedCount} items`)

  if (results.length === 0) {
    throw new Error('有効なデータ行が見つかりませんでした。')
  }

  return results
}

function parseHokukei(csvData) {
  const results = []
  console.log('=== HOKUKEI PARSER START ===')
  console.log('Input rows:', csvData.length)

  let processedCount = 0
  let skippedCount = 0

  for (let i = 0; i < csvData.length; i++) {
    const row = csvData[i]

    const productName = String(row['品名'] || '').trim()
    const salesAmount = String(row['売上金額'] || '').trim()

    if (
      !productName ||
      productName.includes('品名') ||
      productName.includes('合計')
    ) {
      skippedCount++
      continue
    }

    const cleanAmount = cleanNumber(salesAmount)
    if (!cleanAmount || cleanAmount === '0') {
      skippedCount++
      continue
    }

    const masterRow = createMasterRow({
      vendor: row['得意先名略称'] || '北恵株式会社',
      site: row['下店名3'] || '',
      date: formatDate(row['伝票日付'] || ''),
      item: `${row['メーカー名'] || ''} ${productName}`,
      qty: cleanNumber(row['数量'] || '') || '1',
      unit: row['単位名'] || '',
      price: cleanNumber(row['単価'] || ''),
      amount: cleanAmount,
      workNo: row['品番'] || '',
      remarks: row['メーカー名'] || '',
    })

    results.push(masterRow)
    processedCount++
  }

  console.log('=== HOKUKEI PARSING COMPLETE ===')
  console.log(`✓ Processed: ${processedCount} items`)

  if (results.length === 0) {
    throw new Error('有効なデータ行が見つかりませんでした。')
  }

  return results
}

function parseSankoSangyo(csvData) {
  const results = []
  console.log('=== SANKO SANGYO PARSER START ===')
  console.log('Input rows:', csvData.length)

  let clientName = ''
  let projectName = ''

  for (let i = 0; i < Math.min(csvData.length, 10); i++) {
    const row = csvData[i]
    const values = Object.values(row)
    const firstCol = String(values[0] || '').trim()

    if (firstCol.includes('得意先') || firstCol.includes('ALLAGI')) {
      clientName = firstCol.replace(/得意先名?:|様/g, '').trim()
      if (clientName.includes('ALLAGI')) clientName = 'ALLAGI株式会社'
    }

    if (firstCol.includes('様邸') || firstCol.includes('工事')) {
      projectName = firstCol.trim()
    }
  }

  let dataStartIndex = -1

  for (let i = 0; i < Math.min(csvData.length, 15); i++) {
    const row = csvData[i]
    const values = Object.values(row)
    const rowText = values.join('|').toLowerCase()

    if (rowText.includes('日付') || rowText.includes('商品名')) {
      dataStartIndex = i + 1
      break
    }
  }

  if (dataStartIndex === -1) {
    dataStartIndex = 10
  }

  let processedCount = 0
  let skippedCount = 0

  for (let i = dataStartIndex; i < csvData.length; i++) {
    const row = csvData[i]
    const values = Object.values(row)

    const date = String(values[0] || '').trim()
    const itemName = String(values[3] || '').trim()
    const amount = String(values[6] || '').trim()

    if (!itemName || date.includes('合計') || itemName.includes('合計')) {
      skippedCount++
      continue
    }

    const cleanAmount = cleanNumber(amount)
    if (!cleanAmount || cleanAmount === '0') {
      skippedCount++
      continue
    }

    const masterRow = createMasterRow({
      vendor: clientName || '三高産業',
      site: projectName,
      date: date,
      item: itemName,
      qty: values[4] || '',
      unit: '',
      price: values[5] || '',
      amount: amount,
      workNo: values[1] || '',
      remarks: values[2] || '',
    })

    results.push(masterRow)
    processedCount++
  }

  console.log('=== SANKO SANGYO PARSING COMPLETE ===')
  console.log(`✓ Processed: ${processedCount} items`)

  if (results.length === 0) {
    throw new Error('有効なデータ行が見つかりませんでした。')
  }

  return results
}

function parseCleanIndustry(csvData) {
  const results = []
  console.log('=== CLEAN INDUSTRY PARSER START ===')
  console.log('Input rows:', csvData.length)

  let dataStartIndex = -1

  for (let i = 0; i < Math.min(csvData.length, 15); i++) {
    const row = csvData[i]
    const values = Object.values(row)
    const rowText = values.join('|').toLowerCase()

    if (rowText.includes('業者名') || rowText.includes('品名')) {
      dataStartIndex = i + 1
      break
    }
  }

  if (dataStartIndex === -1) {
    dataStartIndex = 10
  }

  let processedCount = 0
  let skippedCount = 0

  for (let i = dataStartIndex; i < csvData.length; i++) {
    const row = csvData[i]
    const values = Object.values(row)
    const cols = values.map(v => String(v || '').trim())

    if (!cols[0] || cols[0].includes('【') || cols[0].includes('合計')) {
      skippedCount++
      continue
    }

    const amount = cols[8] || ''
    const amountNum = cleanNumber(amount)
    if (!amountNum || amountNum === '0') {
      skippedCount++
      continue
    }

    const masterRow = createMasterRow({
      vendor: cols[0] || '',
      site: cols[1] || '',
      date: cols[2] || '',
      item: cols[4] || '',
      qty: cols[5] || '',
      unit: cols[6] || '',
      price: cols[7] || '',
      amount: amount,
      workNo: cols[3] || '',
      remarks: '',
    })

    results.push(masterRow)
    processedCount++
  }

  console.log('=== CLEAN INDUSTRY PARSING COMPLETE ===')
  console.log(`✓ Processed: ${processedCount} items`)

  if (results.length === 0) {
    throw new Error('有効なデータ行が見つかりませんでした。')
  }

  return results
}

function parseWithMapping(csvData, mapping) {
  const results = []
  const csvHeaders = Object.keys(csvData[0] || {})

  console.log('=== STANDARD MAPPING PARSER ===')
  console.log('CSV Headers:', csvHeaders)

  const requiredCols = Object.keys(mapping.map || {})
  const missingCols = requiredCols.filter(col => !csvHeaders.includes(col))

  if (missingCols.length > 0) {
    console.error('❌ Missing required columns:', missingCols)
    throw new Error(`必要な列が見つかりません: ${missingCols.join(', ')}`)
  }

  let processedCount = 0

  for (let i = 0; i < csvData.length; i++) {
    const row = csvData[i]
    const firstValue = String(Object.values(row)[0] || '').trim()

    if (!firstValue) continue

    if (mapping.skipRows?.some(pattern => firstValue.includes(pattern))) {
      continue
    }

    const mapped = {}
    for (const [sourceCol, targetCol] of Object.entries(mapping.map)) {
      let value = row[sourceCol] || ''

      if (targetCol === '納品実績日') {
        value = formatDate(value)
      } else if (targetCol.includes('金額') || targetCol.includes('単価')) {
        value = cleanNumber(value)
      }

      mapped[targetCol] = String(value).trim()
    }

    const masterRow = {}
    MASTER_COLUMNS.forEach(col => {
      masterRow[col] = mapped[col] || ''
    })

    if (masterRow['金額(税抜)'] && !masterRow['金額(税込)']) {
      const tax = parseFloat(masterRow['金額(税抜)']) || 0
      masterRow['金額(税込)'] = Math.round(tax * 1.1).toString()
    }

    if (masterRow['単価(税抜)'] && !masterRow['単価(税込)']) {
      const price = parseFloat(masterRow['単価(税抜)']) || 0
      masterRow['単価(税込)'] = Math.round(price * 1.1).toString()
    }

    if (!masterRow['課税フラグ']) {
      masterRow['課税フラグ'] = '課税'
    }

    results.push(masterRow)
    processedCount++
  }

  console.log(`✓ Processed ${processedCount} rows using standard mapping`)
  return results
}

// ============================================
// UTILITY FUNCTIONS
// ============================================

function createMasterRow(data) {
  const row = {}

  MASTER_COLUMNS.forEach(col => {
    row[col] = ''
  })

  row['取引先'] = String(data.vendor || '').trim()
  row['請求名'] = String(data.site || '').trim()
  row['納品実績日'] = formatDate(data.date || '')
  row['請求納品明細名'] = String(data.item || '').trim()
  row['数量'] = String(data.qty || '').trim() || '1'
  row['単位'] = String(data.unit || '').trim()
  row['単価(税抜)'] = cleanNumber(data.price || '')
  row['金額(税抜)'] = cleanNumber(data.amount || '')
  row['課税フラグ'] = '課税'
  row['請求納品明細備考'] = String(data.workNo || '').trim()

  if (data.remarks) {
    const currentRemarks = row['請求納品明細備考']
    row['請求納品明細備考'] = currentRemarks
      ? `${currentRemarks} ${data.remarks}`
      : data.remarks
  }

  if (row['金額(税抜)']) {
    const amount = parseFloat(row['金額(税抜)']) || 0
    row['金額(税込)'] = Math.round(amount * 1.1).toString()
  }

  if (row['単価(税抜)']) {
    const price = parseFloat(row['単価(税抜)']) || 0
    row['単価(税込)'] = Math.round(price * 1.1).toString()
  }

  return row
}

function formatDate(dateStr) {
  if (!dateStr) return ''
  dateStr = String(dateStr).trim()

  if (dateStr.match(/^\d{4}\/\d{1,2}\/\d{1,2}$/)) {
    return dateStr
  }

  if (dateStr.match(/^\d{8}$/)) {
    const year = dateStr.substring(0, 4)
    const month = dateStr.substring(4, 6)
    const day = dateStr.substring(6, 8)
    return `${year}/${parseInt(month)}/${parseInt(day)}`
  }

  if (!isNaN(dateStr) && dateStr.length > 4) {
    try {
      const date = new Date((parseFloat(dateStr) - 25569) * 86400 * 1000)
      return `${date.getFullYear()}/${date.getMonth() + 1}/${date.getDate()}`
    } catch (e) {
      return dateStr
    }
  }

  if (dateStr.match(/^\d{1,2}\/\d{1,2}$/)) {
    const year = new Date().getFullYear()
    return `${year}/${dateStr}`
  }

  return dateStr
}

function cleanNumber(numStr) {
  if (!numStr) return ''
  const cleaned = String(numStr)
    .replace(/[¥,円]/g, '')
    .replace(/\s+/g, '')
    .trim()

  if (cleaned && !isNaN(cleaned)) {
    return cleaned
  }

  return ''
}

module.exports = { generateExcel }
