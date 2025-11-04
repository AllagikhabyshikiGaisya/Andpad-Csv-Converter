const BaseParser = require('./baseParser')
const { createMasterRow, cleanNumber, formatDate } = require('../excelUtils')

class HokukeiParser extends BaseParser {
  constructor() {
    super('北恵株式会社')
  }

  parse(csvData) {
    this.logStart(csvData)
    const results = []

    for (let i = 0; i < csvData.length; i++) {
      const row = csvData[i]

      const productName = String(row['品名'] || '').trim()
      const salesAmount = String(row['売上金額'] || '').trim()

      if (
        !productName ||
        productName.includes('品名') ||
        productName.includes('合計')
      ) {
        this.skippedCount++
        continue
      }

      const cleanAmount = cleanNumber(salesAmount)
      if (!cleanAmount || cleanAmount === '0') {
        this.skippedCount++
        continue
      }

      const masterRow = createMasterRow({
        vendor: row['得意先名略称'] || '北恵株式会社',
        site: row['下店名3'] || '',
        date: formatDate(row['伝票日付'] || ''),
        item: `${row['メーカー名'] || ''} ${productName}`,
        qty: cleanNumber(row['数量'] || '') || '1',
        unit: row['単位名'] || '',
        price: cleanNumber(row['単価'] || ''),
        amount: cleanAmount,
        workNo: row['品番'] || '',
        remarks: row['メーカー名'] || '',
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

module.exports = HokukeiParser
