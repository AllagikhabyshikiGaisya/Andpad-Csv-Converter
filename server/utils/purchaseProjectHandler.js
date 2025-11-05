// ============================================
// PURCHASE PROJECT HANDLER (仕入案件作成)
// Complete Integration for Land Purchase Projects
// ============================================

const XLSX = require('xlsx')
const Papa = require('papaparse')
const {
  PURCHASE_PROJECT_COLUMNS,
  createPurchaseProjectRow,
} = require('./excelUtils')

/**
 * Validate purchase project row data
 */
function validatePurchaseProjectRow(row, rowIndex) {
  const errors = []

  // Required fields
  if (!row.customerName || row.customerName.trim() === '') {
    errors.push(`Row ${rowIndex}: 顧客名 is required`)
  }

  if (!row.propertyName || row.propertyName.trim() === '') {
    errors.push(`Row ${rowIndex}: 物件名 is required`)
  }

  // Type validation
  if (row.type && !['個人', '法人'].includes(row.type)) {
    errors.push(`Row ${rowIndex}: 種別 must be either '個人' or '法人'`)
  }

  // ID format validation (if provided)
  if (row.projectManagementId && row.projectManagementId.length > 50) {
    errors.push(`Row ${rowIndex}: 案件管理ID too long (max 50 characters)`)
  }

  return errors
}

/**
 * Parse CSV file for purchase project import
 */
function parsePurchaseProjectCSV(csvData) {
  console.log('=== PURCHASE PROJECT CSV PARSING START ===')
  console.log(`Input rows: ${csvData.length}`)

  const results = []
  const errors = []
  let processedCount = 0
  let skippedCount = 0

  for (let i = 0; i < csvData.length; i++) {
    const row = csvData[i]

    // Skip header rows
    const firstValue = String(Object.values(row)[0] || '').trim()
    if (
      firstValue === '顧客名' ||
      firstValue === 'お客様名' ||
      firstValue === '物件名' ||
      firstValue === ''
    ) {
      skippedCount++
      continue
    }

    // Extract data from CSV row
    const data = {
      type: String(row['種別'] || row['個人/法人'] || '個人').trim(),
      customerName: String(row['顧客名'] || row['お客様名'] || '').trim(),
      propertyName: String(row['物件名'] || row['現場名'] || '').trim(),
      projectName: String(
        row['案件名'] || row['物件名'] || row['現場名'] || ''
      ).trim(),
      projectType: String(row['案件種別'] || row['種別'] || '土地仕入').trim(),
      projectManagementId: String(
        row['案件管理ID'] || row['工事番号'] || ''
      ).trim(),
      propertyManagementId: String(
        row['物件管理ID'] || row['工事番号'] || ''
      ).trim(),
      customerManagementId: String(row['顧客管理ID'] || '').trim(),
      projectManager: String(row['案件管理者'] || row['担当者'] || '').trim(),
    }

    // Validate
    const validationErrors = validatePurchaseProjectRow(data, i + 1)
    if (validationErrors.length > 0) {
      errors.push(...validationErrors)
      skippedCount++
      continue
    }

    // Create master row
    const projectRow = createPurchaseProjectRow(data)
    results.push(projectRow)
    processedCount++

    if (processedCount <= 5) {
      console.log(
        `  ✓ Row ${i + 1}: ${data.customerName} - ${data.propertyName} (${
          data.type
        })`
      )
    }
  }

  console.log(`\n✓ Processed: ${processedCount} projects`)
  console.log(`⊘ Skipped: ${skippedCount} rows`)

  if (errors.length > 0) {
    console.warn('\n⚠️  Validation errors:')
    errors.forEach(err => console.warn(`  - ${err}`))
  }

  if (results.length === 0) {
    throw new Error('有効なデータ行が見つかりませんでした。')
  }

  console.log('=== PURCHASE PROJECT CSV PARSING COMPLETE ===\n')

  return results
}

/**
 * Generate purchase project Excel/CSV file
 */
function generatePurchaseProjectFile(csvData, outputFormat = 'xlsx') {
  try {
    console.log('=== PURCHASE PROJECT FILE GENERATION START ===')
    console.log(`Output format: ${outputFormat}`)

    // Parse and validate data
    const transformedData = parsePurchaseProjectCSV(csvData)

    // Log results
    logPurchaseProjectResults(transformedData)

    // Generate file
    let buffer
    let fileExtension

    if (outputFormat === 'csv') {
      buffer = createPurchaseProjectCSV(transformedData)
      fileExtension = 'csv'
      console.log('✓ Purchase project CSV file generated')
    } else {
      buffer = createPurchaseProjectExcel(transformedData)
      fileExtension = 'xlsx'
      console.log('✓ Purchase project Excel file generated')
    }

    console.log('=== PURCHASE PROJECT FILE GENERATION SUCCESS ===')
    return {
      success: true,
      buffer: buffer,
      rowCount: transformedData.length,
      fileExtension: fileExtension,
    }
  } catch (error) {
    console.error('✗ Purchase project file generation error:', error.message)
    return {
      success: false,
      error: error.message || '仕入案件ファイルの生成に失敗しました',
    }
  }
}

/**
 * Create Excel workbook for purchase projects
 */
function createPurchaseProjectExcel(transformedData) {
  const workbook = XLSX.utils.book_new()

  const worksheet = XLSX.utils.json_to_sheet(transformedData, {
    header: PURCHASE_PROJECT_COLUMNS,
  })

  // Set column widths
  worksheet['!cols'] = PURCHASE_PROJECT_COLUMNS.map(col => {
    if (col.includes('管理ID')) return { wch: 20 }
    if (col.includes('名')) return { wch: 30 }
    if (col.includes('種別')) return { wch: 10 }
    if (col.includes('敬称')) return { wch: 8 }
    if (col.includes('フロー')) return { wch: 12 }
    return { wch: 15 }
  })

  XLSX.utils.book_append_sheet(workbook, worksheet, '仕入案件作成')

  const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' })

  return buffer
}

/**
 * Create CSV file for purchase projects
 */
function createPurchaseProjectCSV(transformedData) {
  const csv = Papa.unparse(transformedData, {
    columns: PURCHASE_PROJECT_COLUMNS,
    header: true,
  })

  const BOM = '\uFEFF'
  const buffer = Buffer.from(BOM + csv, 'utf-8')

  return buffer
}

/**
 * Log purchase project results
 */
function logPurchaseProjectResults(transformedData) {
  console.log('\n' + '='.repeat(100))
  console.log('PURCHASE PROJECT TEST RESULTS (仕入案件作成)')
  console.log('='.repeat(100))

  console.log('\n📊 SUMMARY:')
  console.log(`  Total projects: ${transformedData.length}`)

  // Count by type
  const typeCount = {}
  transformedData.forEach(row => {
    const type = row['種別'] || '個人'
    typeCount[type] = (typeCount[type] || 0) + 1
  })

  console.log(`  Type breakdown:`)
  Object.keys(typeCount).forEach(type => {
    console.log(`    - ${type}: ${typeCount[type]}`)
  })

  console.log('\n✓ VALIDATION CHECKS:')

  // Check 1: All required fields present
  let requiredFieldsOk = true
  transformedData.forEach((row, idx) => {
    if (!row['顧客名'] || !row['物件名']) {
      console.log(`  [✗] Row ${idx + 1}: Missing required fields`)
      requiredFieldsOk = false
    }
  })
  console.log(
    `  [${
      requiredFieldsOk ? '✓' : '✗'
    }] All required fields present (顧客名, 物件名)`
  )

  // Check 2: 敬称 correct based on type
  const honorificCorrect = transformedData.every(row => {
    if (row['種別'] === '法人') {
      return row['顧客名 敬称'] === '御中'
    } else {
      return row['顧客名 敬称'] === '様'
    }
  })
  console.log(
    `  [${
      honorificCorrect ? '✓' : '✗'
    }] 顧客名 敬称 correct (法人=御中, 個人=様)`
  )

  // Check 3: 案件フロー = '契約前'
  const flowCorrect = transformedData.every(
    row => row['案件フロー'] === '契約前'
  )
  console.log(`  [${flowCorrect ? '✓' : '✗'}] 案件フロー = '契約前'`)

  // Check 4: 案件種別 present
  const projectTypePresent = transformedData.every(row => row['案件種別'])
  console.log(`  [${projectTypePresent ? '✓' : '✗'}] 案件種別 present`)

  console.log('\n📋 SAMPLE OUTPUT ROWS (First 3):')
  transformedData.slice(0, 3).forEach((row, idx) => {
    console.log(`\n  --- Row ${idx + 1} ---`)
    PURCHASE_PROJECT_COLUMNS.forEach(col => {
      if (row[col]) {
        console.log(`    ${col}: ${row[col]}`)
      }
    })
  })

  console.log('\n' + '='.repeat(100))
  console.log('END PURCHASE PROJECT TEST RESULTS')
  console.log('='.repeat(100) + '\n')
}

module.exports = {
  generatePurchaseProjectFile,
  parsePurchaseProjectCSV,
  validatePurchaseProjectRow,
}
