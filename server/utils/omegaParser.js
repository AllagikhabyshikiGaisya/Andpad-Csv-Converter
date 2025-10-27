function parseOmegaInvoice(csvData, headers) {
  const results = []
  let currentSection = ''
  let deliveryDate = ''

  // Extract delivery date from header
  const headerRow = csvData.find(row => row['日付：'])
  if (headerRow && headerRow['日付：']) {
    deliveryDate = headerRow['日付：'].replace(/日付：/, '').trim()
  }

  // Extract project/client name
  let clientName = ''
  for (const row of csvData) {
    const firstCol = Object.values(row)[0]
    if (
      firstCol &&
      (firstCol.includes('様邸') || firstCol.includes('新築工事'))
    ) {
      clientName = firstCol.split('\t')[0].trim()
      break
    }
  }

  // Process each row
  for (const row of csvData) {
    const firstCol = Object.values(row)[0] || ''

    // Skip header and summary rows
    if (
      firstCol.includes('ALLAGI') ||
      firstCol.includes('ご請求書') ||
      firstCol.includes('内訳') ||
      firstCol.includes('御請求金額') ||
      firstCol.trim() === ''
    ) {
      continue
    }

    // Detect section headers
    if (firstCol.includes('工事部') || firstCol.includes('オプション')) {
      currentSection = firstCol
      continue
    }

    // Extract item details
    const itemName = Object.values(row)[1] || ''
    const quantity = row['発注\n数量'] || row['発注数量'] || ''
    const unitPrice = row['単価(税抜）'] || row['単価'] || ''
    const total = row['計'] || ''

    // Only add rows with actual data
    if (
      itemName &&
      itemName.trim() !== '' &&
      total &&
      total !== '¥-' &&
      total !== '' &&
      total !== '¥0'
    ) {
      // Clean up values
      const cleanQuantity = String(quantity).trim()
      const cleanUnitPrice = String(unitPrice).replace(/[¥,]/g, '').trim()
      const cleanTotal = String(total).replace(/[¥,]/g, '').trim()

      // Skip if total is negative (discounts)
      if (!cleanTotal.startsWith('-')) {
        results.push({
          請求納品明細名: `${currentSection} ${itemName}`.trim(),
          数量: cleanQuantity || '1',
          '単価（税抜）': cleanUnitPrice,
          '金額（税抜）': cleanTotal,
          納品実績日: deliveryDate,
          取引先: clientName,
        })
      }
    }
  }

  return results
}

module.exports = { parseOmegaInvoice }
