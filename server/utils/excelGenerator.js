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

    let transformedData = []

    // Use custom parser if specified
    if (mapping.customParser === true) {
      console.log('Using custom parser for:', mapping.vendor)
      transformedData = parseWithCustomLogic(csvData, mapping.vendor)
    } else {
      // Standard mapping
      transformedData = parseWithMapping(csvData, mapping)
    }

    if (!transformedData || transformedData.length === 0) {
      console.error('No data after transformation')
      return {
        success: false,
        error: 'No valid data rows found. Please check file format.',
      }
    }

    console.log('Transformed rows:', transformedData.length)

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
    console.error('Excel generation error:', error)
    console.error('Stack:', error.stack)
    return {
      success: false,
      error: `Generation failed: ${error.message}`,
    }
  }
}

function parseWithCustomLogic(csvData, vendorName) {
  console.log('Custom parser for:', vendorName)

  if (vendorName === 'クリーン産業') {
    return parseCleanIndustry(csvData)
  }

  console.error('No custom parser found for:', vendorName)
  return []
}

function parseCleanIndustry(csvData) {
  const results = []
  let foundDataHeader = false

  console.log('Parsing Clean Industry invoice...')

  for (let i = 0; i < csvData.length; i++) {
    const row = csvData[i]
    const values = Object.values(row)
    const rowText = values.join('|')

    // Look for data header
    if (!foundDataHeader) {
      if (rowText.includes('業者名') && rowText.includes('品名')) {
        foundDataHeader = true
        console.log('Found data header at row:', i)
        continue
      }
      continue
    }

    // Skip empty or summary rows
    const firstCol = values[0] || ''
    if (
      !firstCol.trim() ||
      firstCol.includes('【') ||
      firstCol.includes('登録番号') ||
      firstCol.includes('請求')
    ) {
      continue
    }

    // Extract data (positions based on header order)
    const vendor = values[0] || ''
    const site = values[1] || ''
    const date = values[2] || ''
    const salesNo = values[3] || ''
    const item = values[4] || ''
    const qty = values[5] || ''
    const unit = values[6] || ''
    const price = values[7] || ''
    const amount = values[8] || ''

    // Only add if has item name and amount
    if (item.trim() && amount.trim() && amount !== '0') {
      results.push(
        createMasterRow({
          vendor: vendor,
          site: site,
          date: date,
          item: item,
          qty: qty,
          unit: unit,
          price: price,
          amount: amount,
        })
      )
    }
  }

  console.log('Clean Industry parsed items:', results.length)
  return results
}

function parseWithMapping(csvData, mapping) {
  const results = []
  const csvHeaders = Object.keys(csvData[0] || {})

  console.log('Using standard mapping')
  console.log('CSV Headers:', csvHeaders)
  console.log('Expected columns:', Object.keys(mapping.map))

  // Check for missing columns
  const requiredCols = Object.keys(mapping.map)
  const missingCols = requiredCols.filter(col => !csvHeaders.includes(col))

  if (missingCols.length > 0) {
    console.error('Missing required columns:', missingCols)
    throw new Error(`Missing columns: ${missingCols.join(', ')}`)
  }

  for (const row of csvData) {
    const firstValue = Object.values(row)[0] || ''

    // Skip summary rows
    if (
      mapping.skipRows?.some(pattern => String(firstValue).includes(pattern))
    ) {
      continue
    }

    if (!firstValue.trim()) {
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
  }

  return results
}

function createMasterRow(data) {
  const row = {}

  MASTER_COLUMNS.forEach(col => {
    row[col] = ''
  })

  row['取引先'] = data.vendor || ''
  row['請求名'] = data.site || ''
  row['納品実績日'] = formatDate(data.date || '')
  row['請求納品明細名'] = data.item || ''
  row['数量'] = data.qty || ''
  row['単位'] = data.unit || ''
  row['単価（税抜）'] = cleanNumber(data.price || '')
  row['金額（税抜）'] = cleanNumber(data.amount || '')
  row['課税フラグ'] = '課税'

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

  if (dateStr.match(/^\d{4}\/\d{1,2}\/\d{1,2}$/)) {
    return dateStr
  }

  if (!isNaN(dateStr) && dateStr.length > 4) {
    try {
      const date = new Date((parseFloat(dateStr) - 25569) * 86400 * 1000)
      return `${date.getFullYear()}/${date.getMonth() + 1}/${date.getDate()}`
    } catch (e) {
      return dateStr
    }
  }

  return dateStr
}

function cleanNumber(numStr) {
  if (!numStr) return ''
  return String(numStr)
    .replace(/[¥,円]/g, '')
    .trim()
}

module.exports = { generateExcel }
