// ============================================
// EXCEL UTILITIES - FIXED FOR MANAGER REQUIREMENTS
// ============================================

const MASTER_COLUMNS = [
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
  '請求納品明細備考',
  '結果',
]

// VENDOR SYSTEM IDS
const VENDOR_SYSTEM_IDS = {
  クリーン産業: '599239',
  三高産業: '563866',
  北恵株式会社: '563913',
  ナンセイ: '563829',
  大萬: '564361',
  髙菱管理: '調整中',
  高菱管理: '調整中',
  オメガジャパン: '598454',
  ナカザワ建販: '566232',
  トキワシステム: '598417',
  ALLAGI株式会社: 'ALLAGI01',
  ALLAGI: 'ALLAGI01',
  'ＡＬＬＡＧＩ㈱': 'ALLAGI01',
}

const ANDPAD_DEFAULTS = {
  取引設定: '紙発注',
  担当者_発注側: '925646',
  現場監督: '925646',
}

// Global counters
let dailySequenceCounter = 1
let projectIdCounter = 1

function getVendorSystemId(vendorName) {
  const systemId = VENDOR_SYSTEM_IDS[vendorName]
  if (!systemId) {
    console.warn(`⚠️ No System ID found for vendor: ${vendorName}`)
    return vendorName
  }
  return systemId
}

// CRITICAL FIX: Generate Case Management ID in format K20251101001 (3-digit sequence with leading zeros)
function generateInvoiceManagementId(sequenceNumber = 1) {
  const today = new Date()
  const year = today.getFullYear()
  const month = String(today.getMonth() + 1).padStart(2, '0')
  const day = String(today.getDate()).padStart(2, '0')
  const seq = String(sequenceNumber).padStart(3, '0') // 3-digit with leading zeros

  // Format: K20251101001 (no underscore, no dash)
  return `K${year}${month}${day}${seq}`
}

function generateProjectId() {
  const today = new Date()
  const year = today.getFullYear()
  const month = String(today.getMonth() + 1).padStart(2, '0')
  const day = String(today.getDate()).padStart(2, '0')
  const sequence = String(projectIdCounter++).padStart(3, '0')

  return `PRJ-${year}${month}${day}-${sequence}`
}

// Invoice name format: 202511_業者名_請求書
function generateInvoiceName(vendorName, invoiceDate = null) {
  let year, month

  if (invoiceDate) {
    const dateMatch = invoiceDate.match(/(\d{4})[\/-](\d{1,2})/)
    if (dateMatch) {
      year = dateMatch[1]
      month = String(dateMatch[2]).padStart(2, '0')
    } else {
      const today = new Date()
      year = today.getFullYear()
      month = String(today.getMonth() + 1).padStart(2, '0')
    }
  } else {
    const today = new Date()
    year = today.getFullYear()
    month = String(today.getMonth() + 1).padStart(2, '0')
  }

  return `${year}${month}${vendorName}_請求書`
}
function createMasterRow(data) {
  const row = {}

  MASTER_COLUMNS.forEach(col => {
    row[col] = ''
  })

  row['請求管理ID'] = generateInvoiceManagementId(dailySequenceCounter++)

  const vendorName = String(data.vendor || '').trim()
  row['取引先'] = getVendorSystemId(vendorName)

  row['取引設定'] = ANDPAD_DEFAULTS.取引設定
  row['担当者(発注側)'] = ANDPAD_DEFAULTS.担当者_発注側
  row['現場監督'] = ANDPAD_DEFAULTS.現場監督

  const invoiceDate = data.date || ''
  const invoiceName = generateInvoiceName(vendorName, invoiceDate)

  row['請求名'] = invoiceName

  // CRITICAL: Handle 案件管理ID - use provided ID or generate fallback
  const providedProjectId = String(data.projectId || '').trim()
  row['案件管理ID'] = providedProjectId || generateProjectId()

  row['納品実績日'] = formatDate(invoiceDate)
  row['支払予定日'] = calculatePaymentDueDate(invoiceDate)

  // CRITICAL FIX: 請求納品明細名 MUST match 請求名 exactly
  row['請求納品明細名'] = invoiceName

  // Always set quantity to 1
  row['数量'] = '1'

  // Default unit to 式
  row['単位'] = String(data.unit || '').trim() || '式'

  row['単価(税抜)'] = cleanNumber(data.price || '')
  row['金額(税抜)'] = cleanNumber(data.amount || '')
  row['課税フラグ'] = '課税'

  // Use 建材関係 not just 建材
  row['工事種類'] = determineConstructionType(data.item || '', vendorName)

  row['請求納品明細備考'] = String(data.workNo || '').trim()

  if (data.remarks) {
    const currentRemarks = row['請求納品明細備考']
    row['請求納品明細備考'] = currentRemarks
      ? `${currentRemarks} ${data.remarks}`
      : data.remarks
  }

  // Calculate tax-inclusive amounts
  if (row['金額(税抜)']) {
    const amount = parseFloat(row['金額(税抜)']) || 0
    row['金額(税込)'] = Math.round(amount * 1.1).toString()
  }

  if (row['単価(税抜)']) {
    const price = parseFloat(row['単価(税抜)']) || 0
    row['単価(税込)'] = Math.round(price * 1.1).toString()
  }

  // CRITICAL FIX: Populate 結果 field
  // Based on requirements, this should be populated with specific status or result
  // Default to "承認" (approved) for imported invoices
  row['結果'] = data.result || '承認'

  return row
}

function calculatePaymentDueDate(invoiceDate) {
  if (!invoiceDate) return ''

  try {
    const formattedDate = formatDate(invoiceDate)
    if (!formattedDate) return ''

    const parts = formattedDate.split('/')
    if (parts.length !== 3) return ''

    const year = parseInt(parts[0])
    const month = parseInt(parts[1]) - 1
    const day = parseInt(parts[2])

    const date = new Date(year, month, day)
    date.setDate(date.getDate() + 30)

    return `${date.getFullYear()}/${date.getMonth() + 1}/${date.getDate()}`
  } catch (e) {
    console.warn('Could not calculate payment due date:', e.message)
    return ''
  }
}

function determineConstructionType(itemDescription, vendorName) {
  const buildingMaterialKeywords = [
    '建材',
    '資材',
    '木材',
    '鋼材',
    '断熱',
    'ボード',
    'テープ',
    '塗料',
    '塗装',
    'コンクリート',
    'セメント',
    '石膏',
    'サイディング',
    '防水',
    '屋根',
    '外壁',
    '床',
    '壁',
    '天井',
    'クロス',
    'タイル',
    '配管',
    'パイプ',
    '電線',
    'ケーブル',
    '金物',
    'ビス',
    'ネジ',
    '接着剤',
    'シール',
    'コーキング',
    'シート',
    'ダンパー',
    '工事',
    '材料',
    '部材',
  ]

  const otherKeywords = [
    '送料',
    '配送',
    '運賃',
    '値引',
    '割引',
    '手数料',
    'サービス',
  ]

  const lowerItem = itemDescription.toLowerCase()

  if (
    otherKeywords.some(keyword => lowerItem.includes(keyword.toLowerCase()))
  ) {
    return 'その他'
  }

  if (
    buildingMaterialKeywords.some(keyword =>
      lowerItem.includes(keyword.toLowerCase())
    )
  ) {
    return '建材関係'
  }

  return '建材関係'
}

function applyVendorSpecificRules(rows, vendorName) {
  if (vendorName === '大萬') {
    console.log('✓ Applying 大萬 1% discount rule')

    rows.forEach(row => {
      if (row['金額(税抜)']) {
        const originalAmount = parseFloat(row['金額(税抜)']) || 0
        const discountedAmount = Math.round(originalAmount * 0.99)
        row['金額(税抜)'] = discountedAmount.toString()
        row['金額(税込)'] = Math.round(discountedAmount * 1.1).toString()
      }

      if (row['単価(税抜)']) {
        const originalPrice = parseFloat(row['単価(税抜)']) || 0
        const discountedPrice = Math.round(originalPrice * 0.99)
        row['単価(税抜)'] = discountedPrice.toString()
        row['単価(税込)'] = Math.round(discountedPrice * 1.1).toString()
      }

      const currentRemarks = row['請求納品明細備考']
      row['請求納品明細備考'] = currentRemarks
        ? `${currentRemarks} [1%割引適用]`
        : '[1%割引適用]'
    })
  }

  return rows
}

// CRITICAL FIX: Calculate invoice totals per vendor group
function calculateInvoiceTotals(rows) {
  let totalTaxExcluded = 0
  let totalTaxIncluded = 0

  rows.forEach(row => {
    const amountExcluded = parseFloat(row['金額(税抜)']) || 0
    const amountIncluded = parseFloat(row['金額(税込)']) || 0

    totalTaxExcluded += amountExcluded
    totalTaxIncluded += amountIncluded
  })

  return {
    totalTaxExcluded: Math.round(totalTaxExcluded).toString(),
    totalTaxIncluded: Math.round(totalTaxIncluded).toString(),
  }
}

// CRITICAL FIX: Add invoice totals per vendor group (not all rows together)
function addInvoiceTotalsToRows(rows) {
  if (!rows || rows.length === 0) return rows

  // Group rows by vendor (取引先)
  const vendorGroups = {}

  rows.forEach(row => {
    const vendor = row['取引先'] || 'unknown'
    if (!vendorGroups[vendor]) {
      vendorGroups[vendor] = []
    }
    vendorGroups[vendor].push(row)
  })

  console.log(
    `✓ Grouped into ${Object.keys(vendorGroups).length} vendor groups`
  )

  // Calculate totals for each vendor group
  Object.keys(vendorGroups).forEach(vendor => {
    const vendorRows = vendorGroups[vendor]
    const totals = calculateInvoiceTotals(vendorRows)

    console.log(
      `  ${vendor}: ${vendorRows.length} rows, Total: ¥${totals.totalTaxExcluded}`
    )

    // Apply totals to all rows in this vendor group
    vendorRows.forEach(row => {
      row['請求納品金額(税抜)'] = totals.totalTaxExcluded
      row['請求納品金額(税込)'] = totals.totalTaxIncluded
    })
  })

  return rows
}

function formatDate(dateStr) {
  if (!dateStr) return ''
  dateStr = String(dateStr).trim()
  if (dateStr.match(/^\d{4}\/\d{1,2}\/\d{1,2}$/)) return dateStr
  if (dateStr.match(/^\d{8}$/)) {
    const year = dateStr.substring(0, 4)
    const month = dateStr.substring(4, 6)
    const day = dateStr.substring(6, 8)
    return `${year}/${parseInt(month)}/${parseInt(day)}`
  }
  if (!isNaN(dateStr) && dateStr.length > 4) {
    try {
      const date = new Date((parseFloat(dateStr) - 25569) * 86400 * 1000)
      return `${date.getFullYear()}/${date.getMonth() + 1}/${date.getDate()}`
    } catch (e) {
      return dateStr
    }
  }
  if (dateStr.match(/^\d{1,2}\/\d{1,2}$/)) {
    const year = new Date().getFullYear()
    return `${year}/${dateStr}`
  }
  return dateStr
}

function cleanNumber(numStr) {
  if (!numStr) return ''
  const cleaned = String(numStr)
    .replace(/[¥,円]/g, '')
    .replace(/\s+/g, '')
    .trim()
  if (cleaned && !isNaN(cleaned)) return cleaned
  return ''
}

function shouldSkipRow(values, additionalPatterns = []) {
  const firstValue = String(values[0] || '').trim()
  const commonSkipPatterns = [
    '請求書',
    '株式会社',
    '御中',
    '〒',
    'TEL',
    'FAX',
    '登録番号',
    '振込先',
    '銀行',
    '合計',
    '小計',
    '消費税',
    ...additionalPatterns,
  ]
  return commonSkipPatterns.some(pattern => firstValue.includes(pattern))
}

function extractNumbers(values) {
  return values
    .map((v, idx) => ({ val: v, idx }))
    .filter(item => {
      const cleaned = cleanNumber(item.val)
      return cleaned && cleaned !== '0'
    })
}

function calculateUnitPrice(amount, quantity) {
  const qty = parseFloat(cleanNumber(quantity)) || 1
  const amt = parseFloat(cleanNumber(amount)) || 0
  if (qty > 0) return Math.round(amt / qty).toString()
  return ''
}

function setColumnWidths(worksheet, columns) {
  worksheet['!cols'] = columns.map(col => {
    if (col.includes('管理ID')) return { wch: 15 } // K20251104001
    if (col.includes('取引先')) return { wch: 12 } // System ID
    if (col.includes('請求名')) return { wch: 35 } // 202507クリーン産業_請求書
    if (col.includes('案件管理ID')) return { wch: 18 } // PRJ-20251104-001
    if (col.includes('明細名')) return { wch: 35 } // Same as 請求名
    if (col.includes('担当者') || col.includes('監督')) return { wch: 12 }
    if (col.includes('日')) return { wch: 12 } // Dates
    if (col.includes('金額') || col.includes('単価')) return { wch: 12 }
    if (col.includes('備考')) return { wch: 20 }
    return { wch: 10 }
  })
}

function resetSequenceCounter() {
  dailySequenceCounter = 1
  projectIdCounter = 1
}

module.exports = {
  MASTER_COLUMNS,
  VENDOR_SYSTEM_IDS,
  ANDPAD_DEFAULTS,
  createMasterRow,
  formatDate,
  cleanNumber,
  shouldSkipRow,
  extractNumbers,
  calculateUnitPrice,
  setColumnWidths,
  resetSequenceCounter,
  applyVendorSpecificRules,
  calculateInvoiceTotals,
  addInvoiceTotalsToRows,
  getVendorSystemId,
  generateInvoiceManagementId,
  generateInvoiceName,
  generateProjectId,
}
