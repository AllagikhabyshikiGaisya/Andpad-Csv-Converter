// ============================================
// EXCEL UTILITIES - COMPLETE FIXED VERSION WITH PROJECT ID HANDLING
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
  '請求納品明細名',
  '数量',
  '単位',
  '単価(税抜)',
  '単価(税込)',
  '金額(税抜)',
  '金額(税込)',
  '工事種類',
]

// PURCHASE PROJECT COLUMNS (仕入案件作成)
const PURCHASE_PROJECT_COLUMNS = [
  '種別',
  '顧客名',
  '顧客名 敬称',
  '物件名',
  '案件名',
  '案件種別',
  '案件管理ID',
  '物件管理ID',
  '顧客管理ID',
  '案件フロー',
  '案件管理者',
]

// VENDOR SYSTEM IDS - CRITICAL: These must match ANDPAD master exactly
const VENDOR_SYSTEM_IDS = {
  クリーン産業: '599239',
  株式会社クリーン産業: '599239',
  三高産業: '563866',
  '株式会社　三高産業': '563866',
  '株式会社 三高産業': '563866',
  北恵株式会社: '563913',
  北恵: '563913',
  ナンセイ: '563829',
  '㈱ナンセイ': '563829',
  大萬: '564361',
  '株式会社　大萬': '564361',
  '株式会社 大萬': '564361',
  髙菱管理: '632782',
  高菱管理: '632782',
  オメガジャパン: '598454',
  オメガジャパン株式会社: '598454',
  ナカザワ建販: '566232',
  ナカザワ建販株式会社: '566232',
  トキワシステム: '598417',
  '㈱トキワシステム': '598417',
  ALLAGI株式会社: 'ALLAGI01',
  ALLAGI: 'ALLAGI01',
  'ＡＬＬＡＧＩ㈱': 'ALLAGI01',
}

// VENDOR INVOICE NAME FORMATS - Based on requirements
const VENDOR_INVOICE_NAMES = {
  トキワシステム: '㈱トキワシステム',
  '㈱トキワシステム': '㈱トキワシステム',
  オメガジャパン: 'オメガジャパン株式会社',
  オメガジャパン株式会社: 'オメガジャパン株式会社',
  ナカザワ建販: 'ナカザワ建販株式会社',
  ナカザワ建販株式会社: 'ナカザワ建販株式会社',
  大萬: '株式会社 大萬',
  '株式会社　大萬': '株式会社 大萬',
  '株式会社 大萬': '株式会社 大萬',
  三高産業: '株式会社 三高産業',
  '株式会社　三高産業': '株式会社 三高産業',
  '株式会社 三高産業': '株式会社 三高産業',
  クリーン産業: '株式会社クリーン産業',
  株式会社クリーン産業: '株式会社クリーン産業',
  ナンセイ: '㈱ナンセイ',
  '㈱ナンセイ': '㈱ナンセイ',
  北恵株式会社: '北恵株式会社',
  北恵: '北恵株式会社',
  髙菱管理: '高菱管理株式会社',
  高菱管理: '高菱管理株式会社',
  高菱管理株式会社: '高菱管理株式会社',
  髙菱管理株式会社: '高菱管理株式会社',
}

const ANDPAD_DEFAULTS = {
  取引設定: '紙発注',
  担当者_発注側: '925646',
  現場監督: '925646',
}

// Global counters
let dailySequenceCounter = 1
let projectIdCounter = 1
let siteToProjectIdMap = {}

function getVendorSystemId(vendorName) {
  const systemId = VENDOR_SYSTEM_IDS[vendorName]
  if (!systemId) {
    console.warn(`⚠️ No System ID found for vendor: ${vendorName}`)
    return vendorName
  }
  return systemId
}

function generateInvoiceManagementId(sequenceNumber = 1) {
  const today = new Date()
  const lastDayOfLastMonth = new Date(today.getFullYear(), today.getMonth(), 0)

  const year = lastDayOfLastMonth.getFullYear()
  const month = String(lastDayOfLastMonth.getMonth() + 1).padStart(2, '0')
  const day = String(lastDayOfLastMonth.getDate()).padStart(2, '0')
  const sequenceNum = String(sequenceNumber)

  return `K${year}${month}${day}_${sequenceNum}`
}

function generateProjectId() {
  const today = new Date()
  const year = today.getFullYear()
  const month = String(today.getMonth() + 1).padStart(2, '0')
  const day = String(today.getDate()).padStart(2, '0')
  const sequence = String(projectIdCounter++).padStart(3, '0')

  return `PRJ-${year}${month}${day}-${sequence}`
}

function getProjectIdForSite(vendorName, siteName) {
  const siteKey = `${vendorName}__${siteName}`

  if (siteToProjectIdMap[siteKey]) {
    return siteToProjectIdMap[siteKey]
  }

  const newProjectId = generateProjectId()
  siteToProjectIdMap[siteKey] = newProjectId

  console.log(`  📋 New Project ID for "${siteName}": ${newProjectId}`)

  return newProjectId
}

// FIXED: Invoice name format YYYYMM_業者名_請求書 (with underscores!)
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

  const displayName = VENDOR_INVOICE_NAMES[vendorName] || vendorName

  return `${year}${month}_${displayName}_請求書`
}

// ============================================
// CRITICAL FIX: createMasterRow with proper 案件管理ID handling
// ============================================
// ============================================
// UPDATED createMasterRow FUNCTION
// Place this in excelUtils.js to replace the existing function
// ============================================

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

  const invoiceDate = data.date || ''
  const invoiceName = generateInvoiceName(vendorName, invoiceDate)
  row['請求名'] = invoiceName

  row['_vendorName'] = vendorName

  const siteName = String(data.site || '').trim()
  const providedProjectId = String(data.projectId || '').trim()

  // ============================================
  // CRITICAL: ALWAYS use provided 案件管理ID from source CSV
  // Rule: "原本CSVの案件管理IDを入力"
  // Rule: "同じ案件管理IDのものはインポート後に1行に情報を集約"
  // ============================================
  if (providedProjectId && providedProjectId !== '') {
    row['案件管理ID'] = providedProjectId

    // Only log for non-MISSING IDs
    if (!providedProjectId.startsWith('MISSING_ID_')) {
      // Success - using CSV data
    }
  } else {
    // ❌ This should RARELY happen - log as error
    console.error(`❌ ERROR: No 案件管理ID provided for vendor "${vendorName}"`)
    console.error(`   Site: "${siteName}"`)
    console.error(`   This violates the rule: "原本CSVの案件管理IDを入力"`)
    console.error(`   The source CSV MUST contain 案件管理ID column or header`)

    // Emergency fallback only - clearly marked
    row['案件管理ID'] = `ERROR_NO_ID_${siteName.replace(/\s+/g, '_')}`
  }

  row['現場監督'] = ANDPAD_DEFAULTS.現場監督

  row['納品実績日'] = formatDate(invoiceDate)

  row['請求納品明細名'] = invoiceName

  row['数量'] = cleanNumber(data.qty || '') || '1'
  row['単位'] = String(data.unit || '').trim() || '式'

  row['単価(税抜)'] = cleanNumber(data.price || '')
  row['金額(税抜)'] = cleanNumber(data.amount || '')

  if (row['金額(税抜)']) {
    const amount = parseFloat(row['金額(税抜)']) || 0
    row['金額(税込)'] = Math.round(amount * 1.1).toString()
  }

  if (row['単価(税抜)']) {
    const price = parseFloat(row['単価(税抜)']) || 0
    row['単価(税込)'] = Math.round(price * 1.1).toString()
  }

  row['工事種類'] = determineConstructionType(data.item || '', vendorName)

  // Metadata fields for consolidation
  row['_siteName'] = siteName
  row['_itemName'] = data.item || ''
  row['_vendorName'] = vendorName
  row['_invoiceDate'] = invoiceDate

  return row
}

function createPurchaseProjectRow(data) {
  const row = {}

  PURCHASE_PROJECT_COLUMNS.forEach(col => {
    row[col] = ''
  })

  row['種別'] = data.type || '個人'
  row['顧客名'] = String(data.customerName || '').trim()
  row['顧客名 敬称'] = data.type === '法人' ? '御中' : '様'
  row['物件名'] = String(data.propertyName || '').trim()
  row['案件名'] = String(data.projectName || data.propertyName || '').trim()
  row['案件種別'] = data.projectType || '土地仕入'
  row['案件管理ID'] = String(data.projectManagementId || '').trim()
  row['物件管理ID'] = String(data.propertyManagementId || '').trim()
  row['顧客管理ID'] = String(data.customerManagementId || '').trim()
  row['案件フロー'] = '契約前'
  row['案件管理者'] = String(data.projectManager || '').trim()

  return row
}

// FIXED: Calculate payment due date - end of next month
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

    // Move to next month
    date.setMonth(date.getMonth() + 1)

    // Set to last day of that month
    date.setMonth(date.getMonth() + 1)
    date.setDate(0)

    return `${date.getFullYear()}/${String(date.getMonth() + 1).padStart(
      2,
      '0'
    )}/${String(date.getDate()).padStart(2, '0')}`
  } catch (e) {
    console.warn('Could not calculate payment due date:', e.message)
    return ''
  }
}

function determineConstructionType(itemDescription, vendorName) {
  if (
    vendorName === 'クリーン産業' ||
    vendorName === '株式会社クリーン産業' ||
    vendorName.includes('クリーン')
  ) {
    return 'その他'
  }

  if (vendorName === 'ナンセイ' || vendorName === '㈱ナンセイ') {
    return 'その他'
  }

  if (
    vendorName === 'オメガジャパン' ||
    vendorName === 'オメガジャパン株式会社'
  ) {
    return '建材'
  }

  if (vendorName === 'ナカザワ建販' || vendorName === 'ナカザワ建販株式会社') {
    return '建材'
  }

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
  ]

  const otherKeywords = [
    '送料',
    '配送',
    '運賃',
    '値引',
    '割引',
    '手数料',
    'サービス',
    '廃棄物',
    '収集運搬',
    '処理費',
    'アスベスト',
    '石綿',
    '回収',
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

  return 'その他'
}

// CRITICAL: Apply vendor-specific rules (大萬 1% discount)
function applyVendorSpecificRules(rows, vendorName) {
  if (
    vendorName === '大萬' ||
    vendorName === '株式会社　大萬' ||
    vendorName === '株式会社 大萬'
  ) {
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

      if (row['請求納品金額(税抜)']) {
        const originalInvoiceAmount = parseFloat(row['請求納品金額(税抜)']) || 0
        const discountedInvoiceAmount = Math.round(originalInvoiceAmount * 0.99)
        row['請求納品金額(税抜)'] = discountedInvoiceAmount.toString()
        row['請求納品金額(税込)'] = Math.round(
          discountedInvoiceAmount * 1.1
        ).toString()
      }
    })
  }

  return rows
}

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

// FIXED: Validation for totals (±1% tolerance)
function validateTotals(rows, vendorName) {
  rows.forEach((row, index) => {
    const invoiceAmountExcluded = parseFloat(row['請求納品金額(税抜)']) || 0
    const totalAmountExcluded = parseFloat(row['金額(税抜)']) || 0

    const difference = Math.abs(invoiceAmountExcluded - totalAmountExcluded)
    const percentDiff =
      totalAmountExcluded > 0 ? (difference / totalAmountExcluded) * 100 : 0

    if (percentDiff > 1) {
      console.warn(`⚠️ Row ${index}: Total mismatch > 1% for ${vendorName}`)
      console.warn(`  請求納品金額: ¥${invoiceAmountExcluded}`)
      console.warn(`  金額合計: ¥${totalAmountExcluded}`)
      console.warn(`  Difference: ${percentDiff.toFixed(2)}%`)
    }
  })
}

// ============================================
// CRITICAL FIX: consolidateByProjectId
// Rule: "同じ案件管理IDのものはインポート後に1行に情報を集約"
// ============================================
function consolidateByProjectId(rows) {
  if (!rows || rows.length === 0) return rows

  console.log('\n=== CONSOLIDATING ROWS BY 案件管理ID ===')
  console.log(`Input: ${rows.length} rows`)

  // Group by 案件管理ID (the actual project ID from source)
  const projectGroups = {}

  rows.forEach(row => {
    const projectId = row['案件管理ID']
    if (!projectGroups[projectId]) {
      projectGroups[projectId] = []
    }
    projectGroups[projectId].push(row)
  })

  console.log(
    `Found ${Object.keys(projectGroups).length} unique 案件管理ID values`
  )

  const consolidatedRows = []
  let consolidatedSequence = 1

  Object.keys(projectGroups).forEach(projectId => {
    const groupRows = projectGroups[projectId]

    console.log(`\n📋 案件管理ID: ${projectId}`)
    console.log(`   Items to consolidate: ${groupRows.length}`)

    const consolidatedRow = { ...groupRows[0] }

    let totalAmountExcluded = 0
    let totalAmountIncluded = 0

    const itemNames = []

    const metadataInvoiceDate = groupRows[0]['_invoiceDate'] || ''

    groupRows.forEach(row => {
      const amountExcluded = parseFloat(row['金額(税抜)']) || 0
      const amountIncluded = parseFloat(row['金額(税込)']) || 0

      totalAmountExcluded += amountExcluded
      totalAmountIncluded += amountIncluded

      if (row['_itemName']) {
        itemNames.push(row['_itemName'])
      }
    })

    // Generate new 請求管理ID for consolidated row
    consolidatedRow['請求管理ID'] = generateInvoiceManagementId(
      consolidatedSequence++
    )
    console.log(`   ✓ New 請求管理ID: ${consolidatedRow['請求管理ID']}`)

    // CRITICAL: Keep the original 案件管理ID from source CSV
    consolidatedRow['案件管理ID'] = projectId
    console.log(`   ✓ Keeping original 案件管理ID: ${projectId}`)

    consolidatedRow['請求納品金額(税抜)'] = totalAmountExcluded.toString()
    consolidatedRow['請求納品金額(税込)'] = totalAmountIncluded.toString()

    consolidatedRow['金額(税抜)'] = totalAmountExcluded.toString()
    consolidatedRow['金額(税込)'] = totalAmountIncluded.toString()

    consolidatedRow['数量'] = '1'
    consolidatedRow['単位'] = '式'
    consolidatedRow['単価(税抜)'] = totalAmountExcluded.toString()
    consolidatedRow['単価(税込)'] = totalAmountIncluded.toString()

    const vendorName = consolidatedRow['_vendorName'] || 'クリーン産業'
    const correctInvoiceName = generateInvoiceName(
      vendorName,
      metadataInvoiceDate
    )

    consolidatedRow['請求名'] = correctInvoiceName
    consolidatedRow['請求納品明細名'] = correctInvoiceName

    if (metadataInvoiceDate) {
      consolidatedRow['納品実績日'] = formatDate(metadataInvoiceDate)
      console.log(`   ✓ Using metadata date: ${consolidatedRow['納品実績日']}`)
    }

    console.log(
      `   ✓ Consolidated total: ¥${totalAmountExcluded} (tax-excluded)`
    )
    console.log(
      `   ✓ Items: ${itemNames.slice(0, 5).join(', ')}${
        itemNames.length > 5 ? '...' : ''
      }`
    )

    consolidatedRows.push(consolidatedRow)
  })

  console.log(`\nOutput: ${consolidatedRows.length} consolidated rows`)
  console.log('=== CONSOLIDATION COMPLETE ===\n')

  return consolidatedRows
}

function addInvoiceTotalsToRows(rows) {
  if (!rows || rows.length === 0) return rows

  const siteGroups = {}

  rows.forEach(row => {
    const vendor = row['取引先'] || 'unknown'
    const siteName = row['_siteName'] || 'default'
    const groupKey = `${vendor}___${siteName}`

    if (!siteGroups[groupKey]) {
      siteGroups[groupKey] = []
    }
    siteGroups[groupKey].push(row)
  })

  console.log(
    `✓ Grouped into ${
      Object.keys(siteGroups).length
    } site groups (vendor + site)`
  )

  Object.keys(siteGroups).forEach(groupKey => {
    const siteRows = siteGroups[groupKey]
    const totals = calculateInvoiceTotals(siteRows)

    const [vendor, siteName] = groupKey.split('___')
    console.log(
      `  ${vendor} / ${siteName}: ${siteRows.length} rows, Total: ¥${totals.totalTaxExcluded} (tax-excluded)`
    )

    siteRows.forEach(row => {
      row['請求納品金額(税抜)'] = totals.totalTaxExcluded
      row['請求納品金額(税込)'] = totals.totalTaxIncluded
    })
  })

  // Consolidate by 案件管理ID
  const consolidatedRows = consolidateByProjectId(rows)

  // Validate totals after consolidation
  const vendorName = consolidatedRows[0]?._vendorName || 'Unknown'
  validateTotals(consolidatedRows, vendorName)

  // Clean up temporary fields
  consolidatedRows.forEach(row => {
    delete row['_siteName']
    delete row['_itemName']
    delete row['_vendorName']
    delete row['_invoiceDate']
  })

  return consolidatedRows
}

function formatDate(dateStr) {
  if (!dateStr) return ''
  dateStr = String(dateStr).trim()
  dateStr = dateStr.replace(/\([月火水木金土日]\)/g, '')
  if (dateStr.match(/^\d{4}\/\d{2}\/\d{2}$/)) return dateStr

  if (dateStr.match(/^\d{4}\/\d{1,2}\/\d{1,2}$/)) {
    const parts = dateStr.split('/')
    const year = parts[0]
    const month = String(parts[1]).padStart(2, '0')
    const day = String(parts[2]).padStart(2, '0')
    return `${year}/${month}/${day}`
  }

  if (dateStr.match(/^\d{8}$/)) {
    const year = dateStr.substring(0, 4)
    const month = dateStr.substring(4, 6)
    const day = dateStr.substring(6, 8)
    return `${year}/${month}/${day}`
  }

  if (!isNaN(dateStr) && dateStr.length > 4) {
    try {
      const date = new Date((parseFloat(dateStr) - 25569) * 86400 * 1000)
      const year = date.getFullYear()
      const month = String(date.getMonth() + 1).padStart(2, '0')
      const day = String(date.getDate()).padStart(2, '0')
      return `${year}/${month}/${day}`
    } catch (e) {
      return dateStr
    }
  }

  if (dateStr.match(/^\d{1,2}\/\d{1,2}$/)) {
    const year = new Date().getFullYear()
    const parts = dateStr.split('/')
    const month = String(parts[0]).padStart(2, '0')
    const day = String(parts[1]).padStart(2, '0')
    return `${year}/${month}/${day}`
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
    if (col.includes('管理ID')) return { wch: 18 }
    if (col.includes('取引先')) return { wch: 12 }
    if (col.includes('請求名')) return { wch: 40 }
    if (col.includes('案件管理ID')) return { wch: 20 }
    if (col.includes('明細名')) return { wch: 40 }
    if (col.includes('担当者') || col.includes('監督')) return { wch: 12 }
    if (col.includes('日')) return { wch: 14 }
    if (col.includes('金額') || col.includes('単価')) return { wch: 14 }
    return { wch: 12 }
  })
}

function resetSequenceCounter() {
  dailySequenceCounter = 1
  projectIdCounter = 1
  siteToProjectIdMap = {}
}

module.exports = {
  MASTER_COLUMNS,
  PURCHASE_PROJECT_COLUMNS,
  VENDOR_SYSTEM_IDS,
  VENDOR_INVOICE_NAMES,
  ANDPAD_DEFAULTS,
  createMasterRow,
  createPurchaseProjectRow,
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
  consolidateByProjectId,
  getVendorSystemId,
  generateInvoiceManagementId,
  generateInvoiceName,
  generateProjectId,
  getProjectIdForSite,
  validateTotals,
}
