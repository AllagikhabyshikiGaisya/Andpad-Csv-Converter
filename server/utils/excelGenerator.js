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
 * Enhanced test logging for all vendors
 */
function logTestResults(vendor, csvData, transformedData, stage = 'COMPLETE') {
  console.log('\n' + '='.repeat(100))
  console.log(`TEST RESULTS FOR: ${vendor} [${stage}]`)
  console.log('='.repeat(100))

  console.log('\n📥 INPUT DATA:')
  console.log(`  Total input rows: ${csvData.length}`)
  if (csvData.length > 0) {
    console.log(`  Input columns: ${Object.keys(csvData[0]).length}`)
    console.log(
      `  First 3 column headers:`,
      Object.keys(csvData[0]).slice(0, 3)
    )
  }

  console.log('\n📤 OUTPUT DATA:')
  console.log(`  Total output rows: ${transformedData.length}`)

  if (transformedData.length > 0) {
    const firstRow = transformedData[0]
    console.log('\n  📋 First Row Data:')
    console.log(`    請求管理ID: ${firstRow['請求管理ID']}`)
    console.log(`    取引先 (System ID): ${firstRow['取引先']}`)
    console.log(`    取引設定: ${firstRow['取引設定']}`)
    console.log(`    担当者(発注側): ${firstRow['担当者(発注側)']}`)
    console.log(`    請求名: ${firstRow['請求名']}`)
    console.log(`    案件管理ID: ${firstRow['案件管理ID']}`)
    console.log(`    請求納品金額(税抜): ¥${firstRow['請求納品金額(税抜)']}`)
    console.log(`    請求納品金額(税込): ¥${firstRow['請求納品金額(税込)']}`)
    console.log(`    現場監督: ${firstRow['現場監督']}`)
    console.log(`    納品実績日: ${firstRow['納品実績日']}`)
    console.log(`    支払予定日: ${firstRow['支払予定日']}`)
    console.log(`    請求納品明細名: ${firstRow['請求納品明細名']}`)
    console.log(`    工事種類: ${firstRow['工事種類']}`)
    console.log(`    数量: ${firstRow['数量']}`)
    console.log(`    単位: ${firstRow['単位']}`)
  }

  console.log('\n💰 FINANCIAL SUMMARY:')
  const totalTaxExcluded = transformedData.reduce((sum, row) => {
    return sum + (parseFloat(row['金額(税抜)']) || 0)
  }, 0)
  const totalTaxIncluded = transformedData.reduce((sum, row) => {
    return sum + (parseFloat(row['金額(税込)']) || 0)
  }, 0)

  const invoiceTotalExcluded = transformedData.reduce((sum, row) => {
    return sum + (parseFloat(row['請求納品金額(税抜)']) || 0)
  }, 0)
  const invoiceTotalIncluded = transformedData.reduce((sum, row) => {
    return sum + (parseFloat(row['請求納品金額(税込)']) || 0)
  }, 0)

  console.log(
    `  Line Items Total (税抜): ¥${totalTaxExcluded.toLocaleString()}`
  )
  console.log(
    `  Line Items Total (税込): ¥${totalTaxIncluded.toLocaleString()}`
  )
  console.log(
    `  Invoice Total (税抜): ¥${invoiceTotalExcluded.toLocaleString()}`
  )
  console.log(
    `  Invoice Total (税込): ¥${invoiceTotalIncluded.toLocaleString()}`
  )

  if (totalTaxExcluded > 0) {
    const taxRate = ((totalTaxIncluded / totalTaxExcluded - 1) * 100).toFixed(1)
    console.log(`  Calculated Tax Rate: ${taxRate}%`)
  }

  console.log('\n✓ VALIDATION CHECKS:')

  // Check 1: 請求名 = 請求納品明細名
  const nameMatch = transformedData.every(
    row => row['請求名'] === row['請求納品明細名']
  )
  console.log(`  [${nameMatch ? '✓' : '✗'}] 請求名 = 請求納品明細名`)

  // Check 2: Tax calculation accuracy
  let taxCorrectCount = 0
  transformedData.forEach(row => {
    const taxExcluded = parseFloat(row['金額(税抜)']) || 0
    const taxIncluded = parseFloat(row['金額(税込)']) || 0
    const expected = Math.round(taxExcluded * 1.1)
    if (Math.abs(taxIncluded - expected) <= 1) {
      taxCorrectCount++
    }
  })
  console.log(
    `  [${
      taxCorrectCount === transformedData.length ? '✓' : '⚠'
    }] Tax calculation (×1.10): ${taxCorrectCount}/${
      transformedData.length
    } correct`
  )

  // Check 3: Date format YYYY/MM/DD
  const dateFormatCorrect = transformedData.every(row => {
    const date = row['納品実績日']
    return !date || date.match(/^\d{4}\/\d{2}\/\d{2}$/)
  })
  console.log(`  [${dateFormatCorrect ? '✓' : '✗'}] Date format (YYYY/MM/DD)`)

  // Check 4: System ID present
  const systemIdCorrect = transformedData.every(row => {
    const id = row['取引先']
    return id && id.length > 0
  })
  console.log(`  [${systemIdCorrect ? '✓' : '✗'}] System ID present`)

  // Check 5: 取引設定 = '紙発注'
  const settingCorrect = transformedData.every(
    row => row['取引設定'] === '紙発注'
  )
  console.log(`  [${settingCorrect ? '✓' : '✗'}] 取引設定 = '紙発注'`)

  // Check 6: 担当者 and 現場監督 = '925646'
  const personInChargeCorrect = transformedData.every(
    row => row['担当者(発注側)'] === '925646' && row['現場監督'] === '925646'
  )
  console.log(
    `  [${personInChargeCorrect ? '✓' : '✗'}] 担当者 & 現場監督 = '925646'`
  )

  // Check 7: Project ID consolidation
  const uniqueProjectIds = new Set(transformedData.map(r => r['案件管理ID']))
  console.log(
    `  [✓] Unique Project IDs: ${uniqueProjectIds.size} (${[
      ...uniqueProjectIds,
    ].join(', ')})`
  )

  // Check 8: 工事種類 values
  const constructionTypes = new Set(transformedData.map(r => r['工事種類']))
  console.log(`  [✓] Construction types: ${[...constructionTypes].join(', ')}`)

  // Check 9: 請求管理ID format (YYYYMMDDNNN)
  const idFormatCorrect = transformedData.every(row => {
    const id = row['請求管理ID']
    return id && id.match(/^\d{8}\d{3}$/)
  })
  console.log(
    `  [${idFormatCorrect ? '✓' : '✗'}] 請求管理ID format (YYYYMMDDNNN)`
  )

  // Check 10: Invoice name format (YYYYMM_業者名_請求書) - UPDATED FORMAT
  const invoiceNameFormat = transformedData.every(row => {
    const name = row['請求名']
    return name && name.match(/^\d{6}_.*_請求書$/)
  })
  console.log(
    `  [${
      invoiceNameFormat ? '✓' : '✗'
    }] Invoice name format (YYYYMM_業者名_請求書)`
  )

  // Check 11: 数量 = 1 for consolidated rows
  const quantityCorrect = transformedData.every(row => row['数量'] === '1')
  console.log(`  [${quantityCorrect ? '✓' : '✗'}] 数量 = 1 (consolidated)`)

  // Check 12: 単位 = 式 for consolidated rows
  const unitCorrect = transformedData.every(row => row['単位'] === '式')
  console.log(`  [${unitCorrect ? '✓' : '✗'}] 単位 = 式 (consolidated)`)

  console.log('\n📊 SAMPLE OUTPUT ROWS (First 2):')
  transformedData.slice(0, 2).forEach((row, idx) => {
    console.log(`\n  --- Row ${idx + 1} ---`)
    // Show only defined columns (no 備考 or 結果)
    const displayColumns = [
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
    ]

    displayColumns.forEach(col => {
      if (row[col] !== undefined && row[col] !== '') {
        console.log(`    ${col}: ${row[col]}`)
      }
    })
  })

  console.log('\n' + '='.repeat(100))
  console.log(`END TEST RESULTS FOR: ${vendor}`)
  console.log('='.repeat(100) + '\n')
}

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

    resetSequenceCounter()

    let transformedData = []

    if (mapping.customParser === true) {
      console.log('✓ Using custom parser for:', mapping.vendor)
      transformedData = parseWithCustomParser(csvData, mapping.vendor)
    } else {
      console.log('✓ Using standard mapping')
      transformedData = parseWithMapping(csvData, mapping)
    }

    if (!transformedData || transformedData.length === 0) {
      console.error('Parser returned no data')
      throw new Error(
        'データ行が見つかりませんでした。ファイル形式を確認してください。'
      )
    }

    console.log('✓ Transformed rows:', transformedData.length)

    // Log after initial transformation
    logTestResults(mapping.vendor, csvData, transformedData, 'AFTER PARSING')

    // Add invoice totals per vendor group AND consolidate by project ID
    transformedData = addInvoiceTotalsToRows(transformedData)

    // Log after consolidation
    logTestResults(
      mapping.vendor,
      csvData,
      transformedData,
      'AFTER CONSOLIDATION'
    )

    // Apply vendor-specific rules AFTER consolidation
    transformedData = applyVendorSpecificRules(transformedData, mapping.vendor)

    // Log final results
    logTestResults(
      mapping.vendor,
      csvData,
      transformedData,
      'FINAL (AFTER RULES)'
    )

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
    console.error('✗ File generation error:', error.message)
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
    console.error('✗ Missing required columns:', missingCols)
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

  worksheet['!merges'] = []

  const columnsToMerge = [
    { col: 1, name: '取引先' },
    { col: 4, name: '請求名' },
    { col: 11, name: '請求納品明細名' },
  ]

  columnsToMerge.forEach(({ col, name }) => {
    let currentValue = null
    let startRow = 1

    for (let i = 0; i < transformedData.length; i++) {
      const cellValue = transformedData[i][name]

      if (cellValue !== currentValue) {
        if (currentValue !== null && i > startRow) {
          worksheet['!merges'].push({
            s: { r: startRow, c: col },
            e: { r: i, c: col },
          })
        }
        currentValue = cellValue
        startRow = i + 1
      }
    }

    if (transformedData.length > startRow) {
      worksheet['!merges'].push({
        s: { r: startRow, c: col },
        e: { r: transformedData.length, c: col },
      })
    }
  })

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

    resetSequenceCounter()

    const allTransformedData = []
    const vendorCounts = {}

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

    allTransformedData = addInvoiceTotalsToRows(allTransformedData)

    const vendorGroups = {}

    allTransformedData.forEach(row => {
      const vendorId = row['取引先']
      if (!vendorGroups[vendorId]) {
        vendorGroups[vendorId] = []
      }
      vendorGroups[vendorId].push(row)
    })

    Object.keys(vendorGroups).forEach(vendorId => {
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

    allTransformedData = []
    Object.values(vendorGroups).forEach(group => {
      allTransformedData.push(...group)
    })

    console.log('✓ Vendors processed:', Object.keys(vendorCounts).join(', '))

    // Log combined results
    logTestResults('COMBINED FILE', [], allTransformedData, 'FINAL COMBINED')

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
    console.error('✗ Combined file generation error:', error.message)
    console.error('Stack:', error.stack)
    return {
      success: false,
      error: error.message || '結合ファイルの生成に失敗しました',
    }
  }
}

module.exports = { generateExcel, generateCombinedExcel }
