const express = require('express')
const multer = require('multer')
const Papa = require('papaparse')
const { detectVendor } = require('../utils/vendorDetector')
const { generateExcel } = require('../utils/excelGenerator')

const router = express.Router()

// Configure multer for memory storage (Multer 2.x compatible)
const storage = multer.memoryStorage()
const upload = multer({
  storage: storage,
  limits: {
    fileSize: 10 * 1024 * 1024, // 10MB limit
  },
  fileFilter: (req, file, cb) => {
    // Accept only CSV files
    if (file.originalname.toLowerCase().endsWith('.csv')) {
      cb(null, true)
    } else {
      cb(new Error('Only CSV files are allowed'))
    }
  },
})

router.post('/convert', upload.single('file'), async (req, res) => {
  try {
    // Validate file upload
    if (!req.file) {
      return res.status(400).json({
        success: false,
        message: {
          en: 'No file uploaded. Please select a CSV file.',
          ja: 'ファイルがアップロードされていません。CSVファイルを選択してください。',
        },
      })
    }

    // Parse CSV
    const fileContent = req.file.buffer.toString('utf-8')
    const parseResult = Papa.parse(fileContent, {
      header: true,
      skipEmptyLines: true,
      encoding: 'utf-8',
    })

    if (parseResult.errors.length > 0) {
      return res.status(400).json({
        success: false,
        message: {
          en: 'Error parsing CSV file. Please check the file format.',
          ja: 'CSVファイルの解析エラー。ファイル形式を確認してください。',
        },
        details: parseResult.errors,
      })
    }

    const csvData = parseResult.data
    const headers = Object.keys(csvData[0] || {})

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

    // Generate Excel file
    const excelResult = generateExcel(csvData, vendorResult.mapping)

    if (!excelResult.success) {
      return res.status(400).json({
        success: false,
        message: {
          en: `Column mismatch: ${excelResult.missingColumns.join(', ')}`,
          ja: `カラムが一致しません: ${excelResult.missingColumns.join(', ')}`,
        },
      })
    }

    // Generate filename
    const date = new Date()
    const yearMonth = `${date.getFullYear()}${String(
      date.getMonth() + 1
    ).padStart(2, '0')}`
    const filename = `${yearMonth}_${vendorResult.mapping.vendor}_請求書.xlsx`

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
  } catch (error) {
    console.error('Conversion error:', error)

    // Handle multer errors
    if (error.message === 'Only CSV files are allowed') {
      return res.status(400).json({
        success: false,
        message: {
          en: 'Invalid file type. Please upload a CSV file.',
          ja: '無効なファイル形式です。CSVファイルをアップロードしてください。',
        },
      })
    }

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
