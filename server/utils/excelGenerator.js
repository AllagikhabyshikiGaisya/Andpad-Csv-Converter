const XLSX = require('xlsx')

function generateExcel(csvData, mapping) {
  try {
    // Validate required columns exist in CSV
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
    const transformedData = csvData.map(row => {
      const newRow = {}
      for (const [sourceCol, targetCol] of Object.entries(mapping.map)) {
        newRow[targetCol] = row[sourceCol] || ''
      }
      return newRow
    })

    // Create workbook
    const workbook = XLSX.utils.book_new()
    const worksheet = XLSX.utils.json_to_sheet(transformedData)

    // Set column widths
    const targetColumns = Object.values(mapping.map)
    worksheet['!cols'] = targetColumns.map(() => ({ wch: 15 }))

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
