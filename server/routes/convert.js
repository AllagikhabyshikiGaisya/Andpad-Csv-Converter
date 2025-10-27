const express = require('express')
const multer = require('multer')
const Papa = require('papaparse')
const XLSX = require('xlsx')
const { detectVendor } = require('../utils/vendorDetector')
const { generateExcel } = require('../utils/excelGenerator')

const router = express.Router()

const storage = multer.memoryStorage()
const upload = multer({
  storage: storage,
  limits: { fileSize: 10 * 1024 * 1024 },
})

router.post('/convert', (req, res) => {
  upload.single('file')(req, res, async err => {
    if (err) {
      console.error('Upload error:', err)
      return res.status(400).json({
        success: false,
        message: {
          en: `Upload error: ${err.message}`,
          ja: `アップロードエラー: ${err.message}`,
        },
      })
    }

    try {
      console.log('=== CONVERSION START ===')

      if (!req.file) {
        return res.status(400).json({
          success: false,
          message: {
            en: 'No file uploaded',
            ja: 'ファイルが選択されていません',
          },
        })
      }

      console.log('File:', req.file.originalname, 'Size:', req.file.size)

      let csvData, headers

      // Parse file
      if (req.file.originalname.toLowerCase().endsWith('.csv')) {
        const content = req.file.buffer.toString('utf-8')
        const result = Papa.parse(content, {
          header: true,
          skipEmptyLines: true,
        })
        csvData = result.data
        headers = Object.keys(csvData[0] || {})
      } else {
        const workbook = XLSX.read(req.file.buffer, { type: 'buffer' })
        const sheet = workbook.Sheets[workbook.SheetNames[0]]
        csvData = XLSX.utils.sheet_to_json(sheet, { defval: '' })
        headers = Object.keys(csvData[0] || {})
      }

      console.log('Rows:', csvData.length, 'Headers:', headers.slice(0, 5))

      if (!csvData || csvData.length === 0) {
        return res.status(400).json({
          success: false,
          message: {
            en: 'File is empty',
            ja: 'ファイルが空です',
          },
        })
      }

      // Detect vendor
      const vendorResult = detectVendor(req.file.originalname, headers)

      if (!vendorResult.detected) {
        return res.status(400).json({
          success: false,
          message: {
            en: 'Vendor not recognized. Please check filename.',
            ja: '業者を認識できません。ファイル名を確認してください。',
          },
        })
      }

      console.log('Vendor:', vendorResult.mapping.vendor)

      // Generate Excel
      const result = generateExcel(csvData, vendorResult.mapping)

      if (!result.success) {
        return res.status(400).json({
          success: false,
          message: {
            en: result.error || 'Conversion failed',
            ja: result.error || '変換に失敗しました',
          },
        })
      }

      // Send file
      const date = new Date()
      const ym = `${date.getFullYear()}${String(date.getMonth() + 1).padStart(
        2,
        '0'
      )}`
      const filename = `${ym}_${vendorResult.mapping.vendor}_ANDPAD.xlsx`

      res.setHeader(
        'Content-Type',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      )
      res.setHeader(
        'Content-Disposition',
        `attachment; filename="${encodeURIComponent(filename)}"`
      )
      res.send(result.buffer)

      console.log('=== SUCCESS ===', result.rowCount, 'rows')
    } catch (error) {
      console.error('Error:', error)
      res.status(500).json({
        success: false,
        message: {
          en: `Server error: ${error.message}`,
          ja: `サーバーエラー: ${error.message}`,
        },
      })
    }
  })
})

module.exports = router
