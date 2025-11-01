const BaseParser = require('./baseParser')
const { createMasterRow, cleanNumber } = require('../excelUtils')

class SankoSangyoParser extends BaseParser {
  constructor() {
    super('三高産業')
  }

  parse(csvData) {
    this.logStart(csvData)
    const results = []

    const metadata = this.extractMetadata(csvData, 10)
    let clientName = metadata.clientName || '三高産業'
    let projectName = metadata.projectName

    let dataStartIndex = this.findDataStart(csvData, ['日付', '商品名'], 15)
    if (dataStartIndex === -1) dataStartIndex = 10

    for (let i = dataStartIndex; i < csvData.length; i++) {
      const row = csvData[i]
      const values = Object.values(row)

      const date = String(values[0] || '').trim()
      const itemName = String(values[3] || '').trim()
      const amount = String(values[6] || '').trim()

      if (!itemName || date.includes('合計') || itemName.includes('合計')) {
        this.skippedCount++
        continue
      }

      const cleanAmount = cleanNumber(amount)
      if (!cleanAmount || cleanAmount === '0') {
        this.skippedCount++
        continue
      }

      const masterRow = createMasterRow({
        vendor: clientName,
        site: projectName,
        date: date,
        item: itemName,
        qty: values[4] || '',
        unit: '',
        price: values[5] || '',
        amount: amount,
        workNo: values[1] || '',
        remarks: values[2] || '',
      })

      results.push(masterRow)
      this.processedCount++
    }

    this.logComplete(results)
    this.validateResults(results)
    return results
  }
}

module.exports = SankoSangyoParser
