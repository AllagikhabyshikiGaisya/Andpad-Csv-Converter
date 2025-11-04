// ============================================
// PURCHASE PROJECT PARSER (仕入案件作成)
// ============================================

const BaseParser = require('./baseParser')
const { createPurchaseProjectRow, formatDate } = require('../excelUtils')

/**
 * Parser for land purchase project imports
 * Handles: 土地仕入案件の作成
 */
class PurchaseProjectParser extends BaseParser {
  constructor() {
    super('仕入案件')
  }

  /**
   * Parse land purchase data into ANDPAD project format
   * @param {Array} csvData - Parsed CSV data
   * @returns {Array} Transformed purchase project rows
   */
  parse(csvData) {
    this.logStart(csvData)

    const results = []

    // Expected columns for land purchase CSV:
    // - 顧客名 (Customer name)
    // - 物件名 (Property name)
    // - 案件管理ID (Project management ID from Taterole)
    // - 物件管理ID (Property management ID from Taterole)
    // - 種別 (Type: 個人 or 法人)
    // - その他オプションフィールド

    for (let i = 0; i < csvData.length; i++) {
      const row = csvData[i]

      // Skip header rows
      if (this.shouldSkipHeaderRow(row)) {
        console.log(`Skipping header row ${i}`)
        this.skippedCount++
        continue
      }

      // Extract required fields
      const customerName = String(row['顧客名'] || row['お客様名'] || '').trim()
      const propertyName = String(row['物件名'] || row['現場名'] || '').trim()
      const projectManagementId = String(
        row['案件管理ID'] || row['工事番号'] || ''
      ).trim()
      const propertyManagementId = String(
        row['物件管理ID'] || row['工事番号'] || ''
      ).trim()

      // Must have at least customer name and property name
      if (!customerName || !propertyName) {
        console.log(`Row ${i}: Missing required fields, skipping`)
        this.skippedCount++
        continue
      }

      // Determine type (個人 or 法人)
      let type = String(row['種別'] || '').trim()
      if (!type) {
        // Default to 個人 (individual) if not specified
        type = '個人'
      }

      // Extract optional fields
      const projectType = String(row['案件種別'] || '').trim() || '土地仕入'
      const customerManagementId = String(row['顧客管理ID'] || '').trim()
      const projectManager = String(row['案件管理者'] || '').trim()

      // Create project name (same as property name for land purchases)
      const projectName = propertyName

      console.log(
        `✓ Processing row ${i}: Customer="${customerName}" Property="${propertyName}" Type="${type}"`
      )

      const projectRow = createPurchaseProjectRow({
        type: type,
        customerName: customerName,
        propertyName: propertyName,
        projectName: projectName,
        projectType: projectType,
        projectManagementId: projectManagementId,
        propertyManagementId: propertyManagementId,
        customerManagementId: customerManagementId,
        projectManager: projectManager,
      })

      results.push(projectRow)
      this.processedCount++
    }

    this.logComplete(results)
    this.validateResults(results)

    return results
  }

  /**
   * Check if row should be skipped (headers, empty rows, etc.)
   */
  shouldSkipHeaderRow(row) {
    const values = Object.values(row)
    const firstValue = String(values[0] || '').trim()

    // Skip if completely empty
    if (values.every(v => !v || String(v).trim() === '')) {
      return true
    }

    // Skip if it's a header row
    const headerKeywords = [
      '顧客名',
      'お客様名',
      '物件名',
      '案件管理ID',
      '種別',
      '個人/法人',
    ]

    return headerKeywords.some(keyword => firstValue.includes(keyword))
  }
}

module.exports = PurchaseProjectParser
