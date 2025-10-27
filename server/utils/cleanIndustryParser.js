function parseCleanIndustryInvoice(csvData) {
  const results = []

  console.log('=== CLEAN INDUSTRY PARSER START ===')
  console.log('Total rows received:', csvData.length)

  // Find where the actual data starts (look for "業者名" header)
  let dataStartIndex = -1
  let headers = []

  for (let i = 0; i < csvData.length; i++) {
    const row = csvData[i]
    const values = Object.values(row)
    const rowText = values.join('')

    // Look for the data header row
    if (
      rowText.includes('業者名') &&
      rowText.includes('現場名') &&
      rowText.includes('品名')
    ) {
      dataStartIndex = i
      // Extract headers from this row
      headers = values.filter(v => v && v.trim() !== '')
      console.log('Data header found at row:', i)
      console.log('Headers:', headers)
      break
    }
  }

  if (dataStartIndex === -1) {
    console.error('Could not find data header row with 業者名, 現場名, 品名')
    return results
  }

  // Parse data rows (start from row after headers)
  for (let i = dataStartIndex + 1; i < csvData.length; i++) {
    const row = csvData[i]
    const values = Object.values(row)

    // Skip empty rows
    if (values.every(v => !v || v.trim() === '')) {
      continue
    }

    // Skip summary rows
    const firstValue = values[0] || ''
    if (
      firstValue.includes('【') ||
      firstValue.includes('登録番号') ||
      firstValue.trim() === ''
    ) {
      console.log('Skipping summary row:', firstValue)
      continue
    }

    // Map the columns based on expected positions
    // Expected order: 業者名, 現場名, 月日, 売上No, 品名, 数量, 単位, 単価, 金額, 消費税, 備考
    const vendor = values[0] || ''
    const siteName = values[1] || ''
    const date = values[2] || ''
    const salesNo = values[3] || ''
    const itemName = values[4] || ''
    const quantity = values[5] || ''
    const unit = values[6] || ''
    const unitPrice = values[7] || ''
    const amount = values[8] || ''
    const tax = values[9] || ''
    const notes = values[10] || ''

    // Only add rows with actual item data
    if (itemName && itemName.trim() !== '' && amount && amount.trim() !== '') {
      results.push({
        請求管理ID: '',
        取引先: vendor.trim(),
        取引設定: '',
        '担当者（発注側）': '',
        請求名: siteName.trim(),
        案件管理ID: '',
        '請求納品金額（税抜）': '',
        '請求納品金額（税込）': '',
        現場監督: '',
        納品実績日: formatDate(date),
        支払予定日: '',
        請求納品明細名: itemName.trim(),
        数量: quantity.trim(),
        単位: unit.trim(),
        '単価（税抜）': cleanNumber(unitPrice),
        '単価（税込）': '',
        '金額（税抜）': cleanNumber(amount),
        '金額（税込）': '',
        工事種類: '',
        課税フラグ: '課税',
        請求納品明細備考: notes.trim(),
        結果: '',
      })
    }
  }

  console.log('Parsed items:', results.length)
  console.log('=== CLEAN INDUSTRY PARSER END ===')

  return results
}

function formatDate(dateStr) {
  if (!dateStr) return ''

  // Handle YYYY/M/D format
  dateStr = String(dateStr).trim()
  if (dateStr.match(/^\d{4}\/\d{1,2}\/\d{1,2}$/)) {
    return dateStr
  }

  // Handle Excel date serial
  if (!isNaN(dateStr) && dateStr.length > 4) {
    // Simple conversion (not perfect but works for most cases)
    const date = new Date((parseFloat(dateStr) - 25569) * 86400 * 1000)
    const year = date.getFullYear()
    const month = date.getMonth() + 1
    const day = date.getDate()
    return `${year}/${month}/${day}`
  }

  return dateStr
}

function cleanNumber(numStr) {
  if (!numStr) return ''
  return String(numStr)
    .replace(/[¥,円]/g, '')
    .trim()
}

module.exports = { parseCleanIndustryInvoice }
