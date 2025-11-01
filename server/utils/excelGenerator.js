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

    // CRITICAL: Add invoice totals to ALL rows
    transformedData = addInvoiceTotalsToRows(transformedData)

    // Apply vendor-specific rules (e.g., 大萬 1% discount)
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

function createExcelWorkbook(transformedData) {
  const workbook = XLSX.utils.book_new()

  const worksheet = XLSX.utils.json_to_sheet(transformedData, {
    header: MASTER_COLUMNS,
  })

  setColumnWidths(worksheet, MASTER_COLUMNS)

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
        // CRITICAL: Add invoice totals for THIS vendor's data
        transformedData = addInvoiceTotalsToRows(transformedData)

        // Apply vendor-specific rules
        transformedData = applyVendorSpecificRules(
          transformedData,
          mapping.vendor
        )

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
      console.log('✓ Combined Excel file generated')
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
