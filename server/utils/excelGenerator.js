const XLSX = require('xlsx')
const Papa = require('papaparse')
const {
  MASTER_COLUMNS,
  setColumnWidths,
  createMasterRow,
  formatDate,
  cleanNumber,
  addInvoiceTotalsToRows,
  applyVendorSpecificRules,
  resetSequenceCounter,
  getVendorSystemId,
} = require('./excelUtils')
const { getParser, hasCustomParser } = require('./parsers')

/**
 * Generate Excel or CSV file from CSV data using vendor-specific parser or mapping
 */
function generateExcel(csvData, mapping, outputFormat = 'xlsx') {
  try {
    console.log('=== FILE GENERATION START ===')
    console.log('Vendor:', mapping.vendor)
    console.log('Output format:', outputFormat)
    console.log('Has custom parser:', mapping.customParser)
    console.log('Input rows:', csvData.length)

    // Reset sequence counter for each file
    resetSequenceCounter()

    let transformedData = []

    // Route to custom parser or standard mapping
    if (mapping.customParser === true) {
      console.log('✓ Using custom parser for:', mapping.vendor)
      transformedData = parseWithCustomParser(csvData, mapping.vendor)
    } else {
      console.log('✓ Using standard mapping')
      transformedData = parseWithMapping(csvData, mapping)
    }

    // Validate results
    if (!transformedData || transformedData.length === 0) {
      console.error('Parser returned no data')
      throw new Error(
        'データ行が見つかりませんでした。ファイル形式を確認してください。'
      )
    }

    console.log('✓ Transformed rows:', transformedData.length)

    // CRITICAL: Add invoice totals per vendor group AND consolidate by project ID
    transformedData = addInvoiceTotalsToRows(transformedData)

    // CRITICAL: Apply vendor-specific rules AFTER consolidation (e.g., 大萬 1% discount)
    // This ensures the discount is applied to the final consolidated amounts
    transformedData = applyVendorSpecificRules(transformedData, mapping.vendor)

    // Generate file based on format
    let buffer
    let fileExtension

    if (outputFormat === 'csv') {
      buffer = createCSVFile(transformedData)
      fileExtension = 'csv'
      console.log('✓ CSV file generated')
    } else {
      buffer = createExcelWorkbook(transformedData)
      fileExtension = 'xlsx'
      console.log('✓ Excel file generated')
    }

    console.log('=== FILE GENERATION SUCCESS ===')
    return {
      success: true,
      buffer: buffer,
      rowCount: transformedData.length,
      fileExtension: fileExtension,
    }
  } catch (error) {
    console.error('✖ File generation error:', error.message)
    console.error('Stack:', error.stack)
    return {
      success: false,
      error: error.message || '生成に失敗しました',
    }
  }
}

function parseWithCustomParser(csvData, vendorName) {
  console.log('=== CUSTOM PARSER CALLED ===')
  console.log('Vendor:', vendorName)

  if (!hasCustomParser(vendorName)) {
    throw new Error(`カスタムパーサーが見つかりません: ${vendorName}`)
  }

  const parser = getParser(vendorName)
  return parser.parse(csvData)
}

function parseWithMapping(csvData, mapping) {
  const results = []
  const csvHeaders = Object.keys(csvData[0] || {})

  console.log('=== STANDARD MAPPING PARSER ===')
  console.log('CSV Headers:', csvHeaders)

  if (!mapping.map || Object.keys(mapping.map).length === 0) {
    throw new Error('マッピング設定が見つかりません。')
  }

  const requiredCols = Object.keys(mapping.map)
  const missingCols = requiredCols.filter(col => !csvHeaders.includes(col))

  if (missingCols.length > 0) {
    console.error('✖ Missing required columns:', missingCols)
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

  if (results.length === 0) {
    throw new Error(
      'データ行が見つかりませんでした。ファイル形式を確認してください。'
    )
  }

  return results
}

// Create Excel with merged cells for same vendor IDs
function createExcelWorkbook(transformedData) {
  const workbook = XLSX.utils.book_new()

  const worksheet = XLSX.utils.json_to_sheet(transformedData, {
    header: MASTER_COLUMNS,
  })

  setColumnWidths(worksheet, MASTER_COLUMNS)

  worksheet['!merges'] = []

  // CRITICAL: Merge cells for columns that have duplicate consecutive values
  // Columns to merge: 取引先 (B), 請求名 (E), 請求納品明細名 (L)
  const columnsToMerge = [
    { col: 1, name: '取引先' }, // Column B
    { col: 4, name: '請求名' }, // Column E
    { col: 11, name: '請求納品明細名' }, // Column L
  ]

  columnsToMerge.forEach(({ col, name }) => {
    let currentValue = null
    let startRow = 1 // Start from row 1 (after header row 0)

    for (let i = 0; i < transformedData.length; i++) {
      const cellValue = transformedData[i][name]

      if (cellValue !== currentValue) {
        // New value found, merge previous cells if needed
        if (currentValue !== null && i > startRow) {
          worksheet['!merges'].push({
            s: { r: startRow, c: col }, // Start row, column
            e: { r: i, c: col }, // End row, column
          })
        }
        currentValue = cellValue
        startRow = i + 1 // +1 because row 0 is header
      }
    }

    // Merge the last group
    if (transformedData.length > startRow) {
      worksheet['!merges'].push({
        s: { r: startRow, c: col },
        e: { r: transformedData.length, c: col },
      })
    }
  })

  // Center align merged cells for all merged columns
  columnsToMerge.forEach(({ col }) => {
    for (let i = 1; i <= transformedData.length; i++) {
      const cellRef = XLSX.utils.encode_cell({ r: i, c: col })
      if (!worksheet[cellRef]) continue

      worksheet[cellRef].s = {
        alignment: {
          vertical: 'center',
          horizontal: 'center',
        },
      }
    }
  })

  XLSX.utils.book_append_sheet(workbook, worksheet, 'ANDPAD Import')

  const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' })

  return buffer
}

function createCSVFile(transformedData) {
  const csv = Papa.unparse(transformedData, {
    columns: MASTER_COLUMNS,
    header: true,
  })

  const BOM = '\uFEFF'
  const buffer = Buffer.from(BOM + csv, 'utf-8')

  return buffer
}

/**
 * Generate combined Excel/CSV file from multiple vendor files
 */
function generateCombinedExcel(filesData, outputFormat = 'xlsx') {
  try {
    console.log('=== COMBINED FILE GENERATION START ===')
    console.log('Number of files:', filesData.length)
    console.log('Output format:', outputFormat)

    // Reset sequence counter for batch processing
    resetSequenceCounter()

    const allTransformedData = []
    const vendorCounts = {}

    // Process each file
    for (let i = 0; i < filesData.length; i++) {
      const { csvData, mapping } = filesData[i]
      console.log(`\nProcessing file ${i + 1}: ${mapping.vendor}`)

      let transformedData = []

      if (mapping.customParser === true) {
        transformedData = parseWithCustomParser(csvData, mapping.vendor)
      } else {
        transformedData = parseWithMapping(csvData, mapping)
      }

      if (transformedData && transformedData.length > 0) {
        const vendorName = mapping.vendor
        vendorCounts[vendorName] =
          (vendorCounts[vendorName] || 0) + transformedData.length

        allTransformedData.push(...transformedData)
        console.log(`✓ Added ${transformedData.length} rows from ${vendorName}`)
      }
    }

    if (allTransformedData.length === 0) {
      throw new Error('No valid data to combine')
    }

    console.log(`\n✓ Total combined rows: ${allTransformedData.length}`)

    // CRITICAL: Group by vendor and add totals per vendor, then consolidate
    allTransformedData = addInvoiceTotalsToRows(allTransformedData)

    // CRITICAL: Apply vendor-specific rules for each vendor in the combined file
    // Group by vendor and apply rules to each vendor's data
    const vendorGroups = {}

    allTransformedData.forEach(row => {
      const vendorId = row['取引先']
      if (!vendorGroups[vendorId]) {
        vendorGroups[vendorId] = []
      }
      vendorGroups[vendorId].push(row)
    })

    // Apply vendor-specific rules to each vendor group
    Object.keys(vendorGroups).forEach(vendorId => {
      // Find vendor name from system ID
      let vendorName = null
      for (const [name, id] of Object.entries(
        require('./excelUtils').VENDOR_SYSTEM_IDS
      )) {
        if (id === vendorId) {
          vendorName = name
          break
        }
      }

      if (vendorName) {
        console.log(`\nApplying rules for vendor: ${vendorName}`)
        applyVendorSpecificRules(vendorGroups[vendorId], vendorName)
      }
    })

    // Flatten back to single array
    allTransformedData = []
    Object.values(vendorGroups).forEach(group => {
      allTransformedData.push(...group)
    })

    console.log('✓ Vendors processed:', Object.keys(vendorCounts).join(', '))

    // Generate file based on format
    let buffer
    let fileExtension

    if (outputFormat === 'csv') {
      buffer = createCSVFile(allTransformedData)
      fileExtension = 'csv'
      console.log('✓ Combined CSV file generated')
    } else {
      buffer = createExcelWorkbook(allTransformedData)
      fileExtension = 'xlsx'
      console.log('✓ Combined Excel file generated WITH MERGED CELLS')
    }

    console.log('=== COMBINED FILE GENERATION SUCCESS ===')
    return {
      success: true,
      buffer: buffer,
      rowCount: allTransformedData.length,
      vendorCount: Object.keys(vendorCounts).length,
      fileExtension: fileExtension,
      vendorBreakdown: vendorCounts,
    }
  } catch (error) {
    console.error('✖ Combined file generation error:', error.message)
    console.error('Stack:', error.stack)
    return {
      success: false,
      error: error.message || '結合ファイルの生成に失敗しました',
    }
  }
}

module.exports = { generateExcel, generateCombinedExcel }
