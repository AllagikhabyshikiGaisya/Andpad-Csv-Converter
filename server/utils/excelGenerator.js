const XLSX = require('xlsx')
const { parseOmegaInvoice } = require('./omegaParser')

function generateExcel(csvData, mapping) {
  try {
    let transformedData

    // Check if this is Omega Japan and needs special parsing
    if (mapping.vendor === 'オメガジャパン') {
      console.log('Using custom Omega Japan parser')
      const csvHeaders = Object.keys(csvData[0] || {})
      transformedData = parseOmegaInvoice(csvData, csvHeaders)

      if (transformedData.length === 0) {
        return {
          success: false,
          error: 'No valid line items found in Omega invoice',
        }
      }
    } else {
      // Standard mapping for other vendors
      const csvHeaders = Object.keys(csvData[0] || {})
      const requiredSourceColumns = Object.keys(mapping.map)
      const missingColumns = requiredSourceColumns.filter(
        col => !csvHeaders.includes(col)
      )

      if (missingColumns.length > 0) {
        return {
          success: false,
          missingColumns: missingColumns,
        }
      }

      // Transform data according to mapping
      transformedData = csvData.map(row => {
        const newRow = {}
        for (const [sourceCol, targetCol] of Object.entries(mapping.map)) {
          newRow[targetCol] = row[sourceCol] || ''
        }
        return newRow
      })
    }

    // Create workbook
    const workbook = XLSX.utils.book_new()
    const worksheet = XLSX.utils.json_to_sheet(transformedData)

    // Set column widths
    worksheet['!cols'] = [
      { wch: 30 }, // 請求納品明細名
      { wch: 10 }, // 数量
      { wch: 15 }, // 単価（税抜）
      { wch: 15 }, // 金額（税抜）
      { wch: 15 }, // 納品実績日
      { wch: 20 }, // 取引先
    ]

    XLSX.utils.book_append_sheet(workbook, worksheet, 'ANDPAD Import')

    // Generate buffer
    const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' })

    return {
      success: true,
      buffer: buffer,
      rowCount: transformedData.length,
    }
  } catch (error) {
    console.error('Excel generation error:', error)
    return {
      success: false,
      error: error.message,
    }
  }
}

module.exports = { generateExcel }
