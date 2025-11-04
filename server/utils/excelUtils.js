// ============================================
// EXCEL UTILITIES - COMPLETE & PERFECTED
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
  担当者_発注側: '925646', // システム担当 (as per requirements)
  現場監督: '925646', // システム担当 (as per requirements)
}

// Global counters
let dailySequenceCounter = 1
let projectIdCounter = 1

// CRITICAL FIX: Track project IDs by site to ensure same site = same project ID
let siteToProjectIdMap = {}

function getVendorSystemId(vendorName) {
  const systemId = VENDOR_SYSTEM_IDS[vendorName]
  if (!systemId) {
    console.warn(`⚠️ No System ID found for vendor: ${vendorName}`)
    return vendorName
  }
  return systemId
}

// FIXED: Generate Invoice Management ID without "K" prefix - Format: 20251104001 (3-digit sequence with leading zeros)
function generateInvoiceManagementId(sequenceNumber = 1) {
  const today = new Date()
  const year = today.getFullYear()
  const month = String(today.getMonth() + 1).padStart(2, '0')
  const day = String(today.getDate()).padStart(2, '0')
  const seq = String(sequenceNumber).padStart(3, '0') // 3-digit with leading zeros

  // FIXED: Format is now 20251104001 (no "K" prefix as per requirements)
  return `${year}${month}${day}${seq}`
}

function generateProjectId() {
  const today = new Date()
  const year = today.getFullYear()
  const month = String(today.getMonth() + 1).padStart(2, '0')
  const day = String(today.getDate()).padStart(2, '0')
  const sequence = String(projectIdCounter++).padStart(3, '0')

  return `PRJ-${year}${month}${day}-${sequence}`
}

// CRITICAL FIX: Get or create project ID for a specific site
function getProjectIdForSite(vendorName, siteName) {
  // Create a unique key for this vendor+site combination
  const siteKey = `${vendorName}__${siteName}`

  // If we've already assigned a project ID to this site, reuse it
  if (siteToProjectIdMap[siteKey]) {
    return siteToProjectIdMap[siteKey]
  }

  // Otherwise, generate a new one and store it
  const newProjectId = generateProjectId()
  siteToProjectIdMap[siteKey] = newProjectId

  console.log(`  📋 New Project ID for "${siteName}": ${newProjectId}`)

  return newProjectId
}

// Invoice name format: YYYYMM_業者名_請求書
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

  // CRITICAL FIX: Use site-based project ID (same site = same project ID)
  const siteName = String(data.site || '').trim()
  const providedProjectId = String(data.projectId || '').trim()

  if (providedProjectId) {
    // If explicitly provided, use it
    row['案件管理ID'] = providedProjectId
  } else {
    // Otherwise, get or create project ID for this vendor+site combination
    row['案件管理ID'] = getProjectIdForSite(vendorName, siteName)
  }

  row['納品実績日'] = formatDate(invoiceDate)
  row['支払予定日'] = calculatePaymentDueDate(invoiceDate)

  // CRITICAL FIX: 請求納品明細名 MUST match 請求名 exactly
  row['請求納品明細名'] = invoiceName

  // FIXED: Use actual quantity from data, not hardcoded "1"
  row['数量'] = cleanNumber(data.qty || '') || '1'

  // Default unit to 式
  row['単位'] = String(data.unit || '').trim() || '式'

  row['単価(税抜)'] = cleanNumber(data.price || '')
  row['金額(税抜)'] = cleanNumber(data.amount || '')

  // FIXED: Set 課税フラグ to "課税" (taxable) as default for all imported invoices
  row['課税フラグ'] = '課税'

  // Use 建材関係 for construction materials, その他 for other items
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

  // FIXED: Leave 結果 empty (will be filled in by accounting/manager in ANDPAD)
  row['結果'] = ''

  // Store site name for grouping later (will be removed before export)
  row['_siteName'] = siteName

  // Store item name for consolidation
  row['_itemName'] = data.item || ''

  return row
}

// NEW: Create purchase project row (仕入案件作成)
function createPurchaseProjectRow(data) {
  const row = {}

  PURCHASE_PROJECT_COLUMNS.forEach(col => {
    row[col] = ''
  })

  // 種別 - Individual or Corporate (default to 個人 if not specified)
  row['種別'] = data.type || '個人'

  // 顧客名 - Customer name
  row['顧客名'] = String(data.customerName || '').trim()

  // 顧客名 敬称 - Honorific (様 for individual, 御中 for corporate)
  row['顧客名 敬称'] = data.type === '法人' ? '御中' : '様'

  // 物件名 - Property name
  row['物件名'] = String(data.propertyName || '').trim()

  // 案件名 - Project name (same as property name for land purchases)
  row['案件名'] = String(data.projectName || data.propertyName || '').trim()

  // 案件種別 - Project type
  row['案件種別'] = data.projectType || '土地仕入'

  // 案件管理ID - From Taterole construction number
  row['案件管理ID'] = String(data.projectManagementId || '').trim()

  // 物件管理ID - From Taterole construction number (Lark auto-number)
  row['物件管理ID'] = String(data.propertyManagementId || '').trim()

  // 顧客管理ID - Customer ID from Taterole (can be blank initially)
  row['顧客管理ID'] = String(data.customerManagementId || '').trim()

  // 案件フロー - Project flow status (契約前 as per manager's response)
  row['案件フロー'] = '契約前'

  // 案件管理者 - Project manager (optional)
  row['案件管理者'] = String(data.projectManager || '').trim()

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
    '廃棄物',
    '収集運搬',
    '処理費',
    'アスベスト',
    '石綿',
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

// CRITICAL: Apply vendor-specific rules (大萬 1% discount)
function applyVendorSpecificRules(rows, vendorName) {
  if (vendorName === '大萬') {
    console.log('✓ Applying 大萬 1% discount rule')

    rows.forEach(row => {
      // Apply discount to 金額
      if (row['金額(税抜)']) {
        const originalAmount = parseFloat(row['金額(税抜)']) || 0
        const discountedAmount = Math.round(originalAmount * 0.99)
        row['金額(税抜)'] = discountedAmount.toString()
        row['金額(税込)'] = Math.round(discountedAmount * 1.1).toString()
      }

      // Apply discount to 単価
      if (row['単価(税抜)']) {
        const originalPrice = parseFloat(row['単価(税抜)']) || 0
        const discountedPrice = Math.round(originalPrice * 0.99)
        row['単価(税抜)'] = discountedPrice.toString()
        row['単価(税込)'] = Math.round(discountedPrice * 1.1).toString()
      }

      // Apply discount to 請求納品金額
      if (row['請求納品金額(税抜)']) {
        const originalInvoiceAmount = parseFloat(row['請求納品金額(税抜)']) || 0
        const discountedInvoiceAmount = Math.round(originalInvoiceAmount * 0.99)
        row['請求納品金額(税抜)'] = discountedInvoiceAmount.toString()
        row['請求納品金額(税込)'] = Math.round(
          discountedInvoiceAmount * 1.1
        ).toString()
      }

      // Add note to remarks
      const currentRemarks = row['請求納品明細備考']
      row['請求納品明細備考'] = currentRemarks
        ? `${currentRemarks} [1%割引適用]`
        : '[1%割引適用]'
    })
  }

  return rows
}

// CRITICAL FIX: Calculate invoice totals per site group (not vendor group)
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

// NEW FUNCTION: Consolidate rows by Project Management ID (案件管理ID)
// This implements the requirement: 同じ案件管理IDの案件は一行に内容を集約したい
function consolidateByProjectId(rows) {
  if (!rows || rows.length === 0) return rows

  console.log('\n=== CONSOLIDATING ROWS BY PROJECT ID ===')
  console.log(`Input: ${rows.length} rows`)

  const projectGroups = {}

  // Group all rows by 案件管理ID
  rows.forEach(row => {
    const projectId = row['案件管理ID']
    if (!projectGroups[projectId]) {
      projectGroups[projectId] = []
    }
    projectGroups[projectId].push(row)
  })

  console.log(`Found ${Object.keys(projectGroups).length} unique project IDs`)

  const consolidatedRows = []

  // Consolidate each project group into a single row
  Object.keys(projectGroups).forEach(projectId => {
    const groupRows = projectGroups[projectId]

    console.log(`\n📋 Project ID: ${projectId}`)
    console.log(`   Items to consolidate: ${groupRows.length}`)

    // Use the first row as base
    const consolidatedRow = { ...groupRows[0] }

    // Calculate total amounts for this project
    let totalAmountExcluded = 0
    let totalAmountIncluded = 0

    // Collect all item names and remarks
    const itemNames = []
    const remarks = []

    groupRows.forEach(row => {
      const amountExcluded = parseFloat(row['金額(税抜)']) || 0
      const amountIncluded = parseFloat(row['金額(税込)']) || 0

      totalAmountExcluded += amountExcluded
      totalAmountIncluded += amountIncluded

      // Collect item names
      if (row['_itemName']) {
        itemNames.push(row['_itemName'])
      }

      // Collect remarks
      if (row['請求納品明細備考']) {
        remarks.push(row['請求納品明細備考'])
      }
    })

    // Update consolidated row with totals
    consolidatedRow['請求納品金額(税抜)'] = totalAmountExcluded.toString()
    consolidatedRow['請求納品金額(税込)'] = totalAmountIncluded.toString()
    consolidatedRow['金額(税抜)'] = totalAmountExcluded.toString()
    consolidatedRow['金額(税込)'] = totalAmountIncluded.toString()

    // For consolidated rows, set quantity to 1 and unit to 式
    consolidatedRow['数量'] = '1'
    consolidatedRow['単位'] = '式'
    consolidatedRow['単価(税抜)'] = totalAmountExcluded.toString()
    consolidatedRow['単価(税込)'] = totalAmountIncluded.toString()

    // Combine all remarks (unique values only)
    const uniqueRemarks = [...new Set(remarks.filter(r => r))]
    consolidatedRow['請求納品明細備考'] = uniqueRemarks.join('; ')

    console.log(
      `   ✓ Consolidated total: ¥${totalAmountExcluded} (tax-excluded)`
    )
    console.log(`   ✓ Items: ${itemNames.join(', ')}`)

    consolidatedRows.push(consolidatedRow)
  })

  console.log(`\nOutput: ${consolidatedRows.length} consolidated rows`)
  console.log('=== CONSOLIDATION COMPLETE ===\n')

  return consolidatedRows
}

// MODIFIED: Add invoice totals per SITE group AND consolidate by project ID
function addInvoiceTotalsToRows(rows) {
  if (!rows || rows.length === 0) return rows

  // STEP 1: Group by vendor + site and calculate totals
  const siteGroups = {}

  rows.forEach(row => {
    const vendor = row['取引先'] || 'unknown'
    const siteName = row['_siteName'] || 'default'
    const groupKey = `${vendor}___${siteName}` // Composite key

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

  // Calculate totals for each site group
  Object.keys(siteGroups).forEach(groupKey => {
    const siteRows = siteGroups[groupKey]
    const totals = calculateInvoiceTotals(siteRows)

    const [vendor, siteName] = groupKey.split('___')
    console.log(
      `  ${vendor} / ${siteName}: ${siteRows.length} rows, Total: ¥${totals.totalTaxExcluded} (tax-excluded)`
    )

    // Apply totals to all rows in this site group
    siteRows.forEach(row => {
      row['請求納品金額(税抜)'] = totals.totalTaxExcluded
      row['請求納品金額(税込)'] = totals.totalTaxIncluded
    })
  })

  // STEP 2: Consolidate rows by Project ID (案件管理ID)
  const consolidatedRows = consolidateByProjectId(rows)

  // Clean up temporary fields
  consolidatedRows.forEach(row => {
    delete row['_siteName']
    delete row['_itemName']
  })

  return consolidatedRows
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
    if (col.includes('管理ID')) return { wch: 15 } // 20251104001
    if (col.includes('取引先')) return { wch: 12 } // System ID
    if (col.includes('請求名')) return { wch: 35 } // 202507クリーン産業_請求書
    if (col.includes('案件管理ID')) return { wch: 18 } // PRJ-20251104-001
    if (col.includes('明細名')) return { wch: 35 } // Same as 請求名
    if (col.includes('担当者') || col.includes('監督')) return { wch: 12 }
    if (col.includes('日')) return { wch: 12 } // Dates
    if (col.includes('金額') || col.includes('単価')) return { wch: 12 }
    if (col.includes('備考')) return { wch: 30 } // Wider for consolidated remarks
    return { wch: 10 }
  })
}

function resetSequenceCounter() {
  dailySequenceCounter = 1
  projectIdCounter = 1
  siteToProjectIdMap = {} // CRITICAL: Also reset site-to-project mapping
}

module.exports = {
  MASTER_COLUMNS,
  PURCHASE_PROJECT_COLUMNS,
  VENDOR_SYSTEM_IDS,
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
}
