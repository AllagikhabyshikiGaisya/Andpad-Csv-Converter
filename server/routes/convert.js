const express = require('express')
const multer = require('multer')
const Papa = require('papaparse')
const XLSX = require('xlsx')
const { detectVendor } = require('../utils/vendorDetector')
const { generateExcel } = require('../utils/excelGenerator')

const router = express.Router()

// Configure multer
const storage = multer.memoryStorage()
const upload = multer({
  storage: storage,
  limits: {
    fileSize: 10 * 1024 * 1024, // 10MB
  },
  fileFilter: (req, file, cb) => {
    const ext = file.originalname.toLowerCase()
    if (ext.endsWith('.csv') || ext.endsWith('.xlsx') || ext.endsWith('.xls')) {
      cb(null, true)
    } else {
      cb(new Error('Only CSV and Excel files are allowed'))
    }
  },
})

router.post('/convert', upload.single('file'), async (req, res) => {
  try {
    console.log('=== CONVERSION REQUEST START ===')

    if (!req.file) {
      return res.status(400).json({
        success: false,
        message: {
          en: 'No file uploaded.',
          ja: 'ファイルがアップロードされていません。',
        },
      })
    }

    console.log('File received:', req.file.originalname)
    console.log('File size:', req.file.size)

    let csvData
    let headers

    // Check file type and parse accordingly
    if (req.file.originalname.toLowerCase().endsWith('.csv')) {
      // Parse CSV
      const fileContent = req.file.buffer.toString('utf-8')
      const parseResult = Papa.parse(fileContent, {
        header: true,
        skipEmptyLines: true,
        encoding: 'utf-8',
      })

      if (parseResult.errors.length > 0) {
        console.error('CSV parse errors:', parseResult.errors)
      }

      csvData = parseResult.data
      headers = Object.keys(csvData[0] || {})
    } else {
      // Parse Excel (XLSX/XLS)
      const workbook = XLSX.read(req.file.buffer, { type: 'buffer' })
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]]
      csvData = XLSX.utils.sheet_to_json(firstSheet, { defval: '' })
      headers = Object.keys(csvData[0] || {})
    }

    console.log('Parsed rows:', csvData.length)
    console.log('Headers:', headers)

    if (!csvData || csvData.length === 0) {
      return res.status(400).json({
        success: false,
        message: {
          en: 'File is empty or unreadable.',
          ja: 'ファイルが空か読み取れません。',
        },
      })
    }

    // Detect vendor
    const vendorResult = detectVendor(req.file.originalname, headers)

    if (!vendorResult.detected) {
      return res.status(400).json({
        success: false,
        message: {
          en: 'Vendor not recognized. Please check the filename or file content.',
          ja: '業者を特定できませんでした。ファイル名またはファイル内容を確認してください。',
        },
      })
    }

    console.log('Vendor detected:', vendorResult.mapping.vendor)

    // Generate Excel in ANDPAD format
    const excelResult = generateExcel(csvData, vendorResult.mapping)

    if (!excelResult.success) {
      return res.status(400).json({
        success: false,
        message: {
          en: `Conversion failed: ${
            excelResult.error || excelResult.missingColumns?.join(', ')
          }`,
          ja: `変換に失敗しました: ${
            excelResult.error || excelResult.missingColumns?.join(', ')
          }`,
        },
      })
    }

    // Generate filename
    const date = new Date()
    const yearMonth = `${date.getFullYear()}${String(
      date.getMonth() + 1
    ).padStart(2, '0')}`
    const filename = `${yearMonth}_${vendorResult.mapping.vendor}_ANDPAD請求書.xlsx`

    console.log('Sending file:', filename)
    console.log('Rows converted:', excelResult.rowCount)

    // Send file
    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    res.setHeader(
      'Content-Disposition',
      `attachment; filename="${encodeURIComponent(filename)}"`
    )
    res.send(excelResult.buffer)

    console.log('=== CONVERSION SUCCESS ===')
  } catch (error) {
    console.error('Conversion error:', error)
    res.status(500).json({
      success: false,
      message: {
        en: 'Internal server error during conversion.',
        ja: '変換中に内部サーバーエラーが発生しました。',
      },
      error: error.message,
    })
  }
})

module.exports = router
