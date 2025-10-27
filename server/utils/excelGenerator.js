function parseCleanIndustry(csvData) {
  const results = []

  console.log('=== CLEAN INDUSTRY PARSER START ===')
  console.log('Input rows:', csvData.length)

  // Debug: show first 15 rows
  console.log('First 15 rows:')
  for (let i = 0; i < Math.min(15, csvData.length); i++) {
    const values = Object.values(csvData[i])
    console.log(`Row ${i}:`, values.slice(0, 5).join(' | '))
  }

  let dataStartIndex = -1

  // Find data header row
  for (let i = 0; i < csvData.length; i++) {
    const row = csvData[i]
    const values = Object.values(row)
    const rowText = values.join(' ').toLowerCase()

    if (rowText.includes('業者名') && rowText.includes('品名')) {
      dataStartIndex = i + 1
      console.log(`Found header at row ${i}, data starts at ${dataStartIndex}`)
      console.log('Header values:', values)
      break
    }
  }

  // If no header found, assume row 10 (typical for this format)
  if (dataStartIndex === -1) {
    console.log('No header found, assuming data starts at row 10')
    dataStartIndex = 10
  }

  // Process data rows
  for (let i = dataStartIndex; i < csvData.length; i++) {
    const row = csvData[i]
    const values = Object.values(row)

    const firstCol = String(values[0] || '').trim()

    // Skip empty rows
    if (!firstCol) continue

    // Skip summary/special rows
    if (
      firstCol.includes('【') ||
      firstCol.includes('】') ||
      firstCol.includes('現場計') ||
      firstCol.includes('業者計') ||
      firstCol.includes('登録番号') ||
      firstCol.includes('対象') ||
      firstCol.startsWith('T1') ||
      firstCol === '###'
    ) {
      console.log(`Row ${i}: SKIP - ${firstCol}`)
      continue
    }

    // Extract columns based on Excel structure
    const vendor = values[0] || ''
    const site = values[1] || ''
    const date = values[2] || ''
    const workNo = values[3] || ''
    const item = values[4] || ''
    const qty = values[5] || ''
    const unit = values[6] || ''
    const price = values[7] || ''
    const amount = values[8] || ''

    const itemStr = String(item).trim()
    const amountStr = String(amount).trim()

    // Must have item and amount
    if (!itemStr || !amountStr || amountStr === '0') {
      console.log(`Row ${i}: SKIP - no item/amount`)
      continue
    }

    console.log(`Row ${i}: PROCESS - ${itemStr} = ${amountStr}`)

    // Create master row
    const masterRow = {}
    const COLUMNS = [
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

    COLUMNS.forEach(col => {
      masterRow[col] = ''
    })

    masterRow['取引先'] = String(vendor).trim()
    masterRow['請求名'] = String(site).trim()
    masterRow['納品実績日'] = formatDate(date)
    masterRow['請求納品明細名'] = itemStr
    masterRow['数量'] = String(qty).trim() || '1'
    masterRow['単位'] = String(unit).trim()
    masterRow['単価（税抜）'] = cleanNumber(price)
    masterRow['金額（税抜）'] = cleanNumber(amount)
    masterRow['課税フラグ'] = '課税'
    masterRow['請求納品明細備考'] = String(workNo).trim()

    // Calculate tax
    const amountNum = parseFloat(masterRow['金額（税抜）']) || 0
    const priceNum = parseFloat(masterRow['単価（税抜）']) || 0

    if (amountNum > 0) {
      masterRow['金額（税込）'] = Math.round(amountNum * 1.1).toString()
    }
    if (priceNum > 0) {
      masterRow['単価（税込）'] = Math.round(priceNum * 1.1).toString()
    }

    results.push(masterRow)
  }

  console.log('=== CLEAN INDUSTRY PARSED ===')
  console.log('Total items:', results.length)
  return results
}

function formatDate(dateStr) {
  if (!dateStr) return ''
  const str = String(dateStr).trim()

  // Already formatted
  if (str.match(/^\d{4}\/\d{1,2}\/\d{1,2}$/)) {
    return str
  }

  // Excel serial number
  if (!isNaN(str) && str.length > 4) {
    try {
      const date = new Date((parseFloat(str) - 25569) * 86400 * 1000)
      return `${date.getFullYear()}/${date.getMonth() + 1}/${date.getDate()}`
    } catch (e) {
      return str
    }
  }

  return str
}

function cleanNumber(numStr) {
  if (!numStr) return ''
  return String(numStr)
    .replace(/[¥,円]/g, '')
    .trim()
}
