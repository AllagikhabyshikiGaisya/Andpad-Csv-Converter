const BaseParser = require('./baseParser')
const { createMasterRow, cleanNumber, formatDate } = require('../excelUtils')

class NakazawaKenhanParser extends BaseParser {
  constructor() {
    super('ナカザワ建販')
  }

  parse(csvData) {
    this.logStart(csvData)
    const results = []

    for (let i = 0; i < csvData.length; i++) {
      const row = csvData[i]

      const siteName = String(row['現場名'] || '').trim()
      const productName = String(row['商品名'] || '').trim()
      const salesAmount = String(row['売上金額'] || '').trim()

      if (!siteName && !productName) {
        this.skippedCount++
        continue
      }

      if (siteName.includes('現場名') || productName.includes('商品名')) {
        this.skippedCount++
        continue
      }

      const cleanAmount = cleanNumber(salesAmount)
      if (!cleanAmount || cleanAmount === '0') {
        this.skippedCount++
        continue
      }

      const makerName = String(row['メーカー名'] || '').trim()
      const spec = String(row['規格'] || '').trim()
      const productCode = String(row['商品コード'] || '').trim()

      let itemDescription = productName
      if (makerName) itemDescription = `${makerName} ${productName}`
      if (spec) itemDescription += ` [${spec}]`
      if (productCode) itemDescription += ` (${productCode})`

      const quantity = String(row['数量'] || '').trim()
      const unitPrice = String(row['売上単価'] || '').trim()
      const formattedDate = formatDate(String(row['売上日'] || '').trim())

      let finalUnitPrice = cleanNumber(unitPrice)
      if (!finalUnitPrice && quantity && cleanAmount) {
        const qty = parseFloat(cleanNumber(quantity)) || 1
        const amt = parseFloat(cleanAmount) || 0
        if (qty > 0) finalUnitPrice = Math.round(amt / qty).toString()
      }

      const pieceCount = String(row['個数'] || '').trim()
      const finalQty = cleanNumber(quantity) || cleanNumber(pieceCount) || '1'
      const finalUnit = String(row['売上単価単位名'] || '').trim() || '個'

      const remarksParts = [
        row['売上Ｎｏ'] ? `伝票:${row['売上Ｎｏ']}` : '',
        row['区分名'] ? `区分:${row['区分名']}` : '',
        row['担当者名'] ? `担当:${row['担当者名']}` : '',
        row['入数'] && row['入数'] !== '0' ? `入数:${row['入数']}` : '',
        row['備考'] || '',
      ].filter(Boolean)

      const masterRow = createMasterRow({
        vendor: 'ナカザワ建販',
        site: siteName,
        date: formattedDate,
        item: itemDescription.trim(),
        qty: finalQty,
        unit: finalUnit,
        price: finalUnitPrice,
        amount: cleanAmount,
        workNo: productCode || row['売上Ｎｏ'] || '',
        remarks: remarksParts.join(' '),
        projectId: '', // Let system generate
        result: '承認',
      })

      results.push(masterRow)
      this.processedCount++

      if (this.processedCount <= 5) {
        console.log(
          `  ✓ Row ${i}: ${siteName} - ${productName} - ¥${salesAmount}`
        )
      }
    }

    this.logComplete(results)
    this.validateResults(results)
    return results
  }
}

module.exports = NakazawaKenhanParser
