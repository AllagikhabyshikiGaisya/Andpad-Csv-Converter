// ============================================
// COMPLETE excelGenerator.js - FIXED VERSION
// ============================================
const XLSX = require('xlsx')

// Master ANDPAD columns
const MASTER_COLUMNS = [
  '請求管理ID',
  '取引先',
  '取引設定',
  '担当者（発注側）',
  '請求名',
  '案件管理ID',
  '請求納品金額（税抜）',
  '請求納品金額（税込）',
  '現場監督',
  '納品実績日',
  '支払予定日',
  '請求納品明細名',
  '数量',
  '単位',
  '単価（税抜）',
  '単価（税込）',
  '金額（税抜）',
  '金額（税込）',
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
    console.log('Mapping object:', JSON.stringify(mapping, null, 2))

    let transformedData = []

    // CRITICAL FIX: Check customParser flag correctly
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
      // Standard mapping - validate map exists
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
  console.log('Available parsers: クリーン産業')

  if (vendorName === 'クリーン産業') {
    return parseCleanIndustry(csvData)
  }

  console.error('❌ No custom parser found for:', vendorName)
  throw new Error(`カスタムパーサーが見つかりません: ${vendorName}`)
}

function parseCleanIndustry(csvData) {
  const results = []

  console.log('=== CLEAN INDUSTRY PARSER START ===')
  console.log('Input rows:', csvData.length)

  // Debug: show first few rows
  if (csvData.length > 0) {
    console.log('First row keys:', Object.keys(csvData[0]))
    console.log('First row values:', Object.values(csvData[0]).slice(0, 15))
    if (csvData.length > 5) {
      console.log('Row 5 values:', Object.values(csvData[5]).slice(0, 15))
      console.log('Row 10 values:', Object.values(csvData[10]).slice(0, 15))
    }
  }

  let dataStartIndex = -1
  let headerRow = null

  // Find data header row (業者名, 品名, etc.)
  for (let i = 0; i < Math.min(csvData.length, 15); i++) {
    const row = csvData[i]
    const values = Object.values(row)
    const rowText = values.join('|').toLowerCase()

    console.log(`Row ${i}: ${values.slice(0, 5).join(' | ')}`)

    // Look for header indicators
    if (
      rowText.includes('業者名') ||
      rowText.includes('品名') ||
      rowText.includes('現場名')
    ) {
      dataStartIndex = i + 1
      headerRow = i
      console.log(
        `✓ Header found at row ${i}, data starts at ${dataStartIndex}`
      )
      break
    }
  }

  // If no header found, try from row 10
  if (dataStartIndex === -1) {
    console.log('⚠ No header found, starting from row 10')
    dataStartIndex = 10
  }

  console.log(`Processing data from row ${dataStartIndex}`)

  // Process data rows
  let processedCount = 0
  let skippedCount = 0

  for (let i = dataStartIndex; i < csvData.length; i++) {
    const row = csvData[i]
    const values = Object.values(row)

    // Get all values as strings
    const cols = values.map(v => String(v || '').trim())
    const firstCol = cols[0]

    // Skip empty rows
    if (!firstCol) {
      skippedCount++
      continue
    }

    // Skip summary/special rows
    if (
      firstCol.includes('【') ||
      firstCol.includes('】') ||
      firstCol.includes('現場計') ||
      firstCol.includes('業者計') ||
      firstCol.includes('登録番号') ||
      firstCol.includes('対象') ||
      firstCol.startsWith('T1') ||
      firstCol.startsWith('T0') ||
      firstCol === '###' ||
      firstCol.startsWith('#') ||
      firstCol.includes('合計') ||
      firstCol.includes('小計')
    ) {
      console.log(`  Skipping summary row ${i}: ${firstCol}`)
      skippedCount++
      continue
    }

    // Extract columns based on Clean Industry format
    // Expected format: 業者名, 現場名, 月日, 完工No, 品名, 数量, 単位, 単価, 金額
    const vendor = cols[0] || ''
    const site = cols[1] || ''
    const date = cols[2] || ''
    const workNo = cols[3] || ''
    const item = cols[4] || ''
    const qty = cols[5] || ''
    const unit = cols[6] || ''
    const price = cols[7] || ''
    const amount = cols[8] || ''

    // Validation: Must have item name and non-zero amount
    if (!item) {
      console.log(`  Skipping row ${i}: No item name`)
      skippedCount++
      continue
    }

    const amountNum = cleanNumber(amount)
    if (!amountNum || amountNum === '0') {
      console.log(`  Skipping row ${i}: Zero or invalid amount`)
      skippedCount++
      continue
    }

    // Create master row
    const masterRow = createMasterRow({
      vendor: vendor,
      site: site,
      date: date,
      item: item,
      qty: qty,
      unit: unit,
      price: price,
      amount: amount,
      workNo: workNo,
    })

    results.push(masterRow)
    processedCount++

    // Log first few items
    if (processedCount <= 3) {
      console.log(`  ✓ Row ${i}: ${item} - ¥${amount}`)
    }
  }

  console.log('=== CLEAN INDUSTRY PARSING COMPLETE ===')
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

function parseWithMapping(csvData, mapping) {
  const results = []
  const csvHeaders = Object.keys(csvData[0] || {})

  console.log('=== STANDARD MAPPING PARSER ===')
  console.log('CSV Headers:', csvHeaders)
  console.log('Expected columns:', Object.keys(mapping.map || {}))

  // Validate required columns
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

    // Skip empty rows
    if (!firstValue) continue

    // Skip summary rows
    if (mapping.skipRows?.some(pattern => firstValue.includes(pattern))) {
      continue
    }

    // Map columns
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

    // Create master row
    const masterRow = {}
    MASTER_COLUMNS.forEach(col => {
      masterRow[col] = mapped[col] || ''
    })

    // Calculate tax
    if (masterRow['金額（税抜）'] && !masterRow['金額（税込）']) {
      const tax = parseFloat(masterRow['金額（税抜）']) || 0
      masterRow['金額（税込）'] = Math.round(tax * 1.1).toString()
    }

    if (masterRow['単価（税抜）'] && !masterRow['単価（税込）']) {
      const price = parseFloat(masterRow['単価（税抜）']) || 0
      masterRow['単価（税込）'] = Math.round(price * 1.1).toString()
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
  row['単価（税抜）'] = cleanNumber(data.price || '')
  row['金額（税抜）'] = cleanNumber(data.amount || '')
  row['課税フラグ'] = '課税'
  row['請求納品明細備考'] = String(data.workNo || '').trim()

  // Calculate tax-included amounts
  if (row['金額（税抜）']) {
    const amount = parseFloat(row['金額（税抜）']) || 0
    row['金額（税込）'] = Math.round(amount * 1.1).toString()
  }

  if (row['単価（税抜）']) {
    const price = parseFloat(row['単価（税抜）']) || 0
    row['単価（税込）'] = Math.round(price * 1.1).toString()
  }

  return row
}

function formatDate(dateStr) {
  if (!dateStr) return ''
  dateStr = String(dateStr).trim()

  // Already in correct format
  if (dateStr.match(/^\d{4}\/\d{1,2}\/\d{1,2}$/)) {
    return dateStr
  }

  // Excel serial date
  if (!isNaN(dateStr) && dateStr.length > 4) {
    try {
      const date = new Date((parseFloat(dateStr) - 25569) * 86400 * 1000)
      return `${date.getFullYear()}/${date.getMonth() + 1}/${date.getDate()}`
    } catch (e) {
      console.error('Date conversion error:', e)
      return dateStr
    }
  }

  // MM/DD format - assume current year
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

  // Validate it's a number
  if (cleaned && !isNaN(cleaned)) {
    return cleaned
  }

  return ''
}

module.exports = { generateExcel }
