const BaseParser = require('./baseParser')
const { createMasterRow, cleanNumber, formatDate } = require('../excelUtils')

class TaimanParser extends BaseParser {
  constructor() {
    super('大萬')
  }

  parse(csvData) {
    this.logStart(csvData)
    const results = []

    for (let i = 0; i < csvData.length; i++) {
      const row = csvData[i]

      const productName = String(row['商品名'] || '').trim()
      const customerName = String(row['請求先名'] || '').trim()
      const purchaseAmount = String(row['仕入金額'] || '').trim()

      if (
        !productName ||
        productName.includes('商品名') ||
        productName.includes('合計')
      ) {
        this.skippedCount++
        continue
      }

      const cleanAmount = cleanNumber(purchaseAmount)
      if (!cleanAmount || cleanAmount === '0') {
        this.skippedCount++
        continue
      }

      const masterRow = createMasterRow({
        vendor: customerName || '大萬',
        site: '',
        date: formatDate(row['出荷日'] || ''),
        item: productName,
        qty: cleanNumber(row['数量'] || '') || '1',
        unit: row['単位'] || '個',
        price: cleanNumber(row['仕入単価'] || ''),
        amount: cleanAmount,
        workNo: row['商品コード'] || '',
        remarks: row['備考'] || '',
        projectId: '', // Let system generate
        result: '承認',
      })

      results.push(masterRow)
      this.processedCount++
    }

    this.logComplete(results)
    this.validateResults(results)
    return results
  }
}

module.exports = TaimanParser
