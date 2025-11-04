const BaseParser = require('./baseParser')
const { createMasterRow, cleanNumber, formatDate } = require('../excelUtils')

class NanseiParser extends BaseParser {
  constructor() {
    super('ナンセイ')
  }

  parse(csvData) {
    this.logStart(csvData)
    const results = []

    for (let i = 0; i < csvData.length; i++) {
      const row = csvData[i]

      const customerName = String(row['取引先名'] || '').trim()
      const transactionAmount = String(
        row['今回取引額(税抜)'] || row['今回取引額'] || ''
      ).trim()

      if (!customerName || customerName.includes('取引先名')) {
        this.skippedCount++
        continue
      }

      const cleanAmount = cleanNumber(transactionAmount)
      if (!cleanAmount || cleanAmount === '0') {
        this.skippedCount++
        continue
      }

      const masterRow = createMasterRow({
        vendor: customerName,
        site: row['振込銀行名'] || '',
        date: formatDate(row['請求日付'] || ''),
        item: '請求',
        qty: '1',
        unit: '式',
        price: cleanAmount,
        amount: cleanAmount,
        workNo: row['請求番号'] || '',
        remarks: row['取引先CD'] || '',
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

module.exports = NanseiParser
