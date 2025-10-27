const XLSX = require('xlsx')
const { parseCleanIndustryInvoice } = require('./cleanIndustryParser')

// Master ANDPAD columns in correct order
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
    console.log('Input rows:', csvData.length)

    let transformedData

    // Check if this vendor needs a custom parser
    if (mapping.vendor === 'クリーン産業') {
      console.log('Using custom Clean Industry parser')
      transformedData = parseCleanIndustryInvoice(csvData)

      if (transformedData.length === 0) {
        return {
          success: false,
          error: 'No valid line items found in Clean Industry invoice',
        }
      }
    } else {
      // Standard mapping for other vendors
      const csvHeaders = Object.keys(csvData[0] || {})
      console.log('CSV Headers:', csvHeaders)

      const requiredSourceColumns = Object.keys(mapping.map)
      const missingColumns = requiredSourceColumns.filter(
        col => !csvHeaders.includes(col)
      )

      if (missingColumns.length > 0) {
        console.log('Missing columns:', missingColumns)
        return {
          success: false,
          missingColumns: missingColumns,
        }
      }

      // Transform data to ANDPAD master format
      transformedData = []

      for (const row of csvData) {
        // Skip summary rows
        const firstValue = Object.values(row)[0] || ''
        const shouldSkip = mapping.skipRows?.some(pattern =>
          String(firstValue).includes(pattern)
        )

        if (shouldSkip || !firstValue.trim()) {
          continue
        }

        // Create new row with MASTER_COLUMNS structure
        const newRow = {}

        // Initialize all master columns
        MASTER_COLUMNS.forEach(col => {
          newRow[col] = ''
        })

        // Map source columns to master columns
        for (const [sourceCol, targetCol] of Object.entries(mapping.map)) {
          let value = row[sourceCol] || ''
          value = String(value).trim()

          if (targetCol === '納品実績日') {
            value = formatDate(value)
          }

          if (targetCol.includes('金額') || targetCol.includes('単価')) {
            value = cleanNumber(value)
          }

          newRow[targetCol] = value
        }

        // Calculate tax-included amounts
        if (newRow['金額（税抜）'] && !newRow['金額（税込）']) {
          const taxExcluded = parseFloat(newRow['金額（税抜）']) || 0
          newRow['金額（税込）'] = Math.round(taxExcluded * 1.1).toString()
        }

        if (newRow['単価（税抜）'] && !newRow['単価（税込）']) {
          const unitPrice = parseFloat(newRow['単価（税抜）']) || 0
          newRow['単価（税込）'] = Math.round(unitPrice * 1.1).toString()
        }

        if (!newRow['課税フラグ']) {
          newRow['課税フラグ'] = '課税'
        }

        transformedData.push(newRow)
      }
    }

    if (transformedData.length === 0) {
      return {
        success: false,
        error: 'No valid data rows found after filtering',
      }
    }

    console.log('Transformed rows:', transformedData.length)

    // Create workbook with MASTER_COLUMNS order
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

    // Generate buffer
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
      error: error.message,
    }
  }
}

function formatDate(dateStr) {
  if (!dateStr) return ''

  dateStr = String(dateStr).trim()

  if (dateStr.match(/^\d{4}\/\d{1,2}\/\d{1,2}$/)) {
    return dateStr
  }

  if (dateStr.match(/^\d{1,2}\/\d{1,2}$/)) {
    const year = new Date().getFullYear()
    return `${year}/${dateStr}`
  }

  if (!isNaN(dateStr) && dateStr.length > 4) {
    try {
      const date = XLSX.SSF.parse_date_code(parseFloat(dateStr))
      return `${date.y}/${date.m}/${date.d}`
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
