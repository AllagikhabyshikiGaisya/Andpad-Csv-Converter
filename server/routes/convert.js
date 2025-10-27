const express = require('express')
const multer = require('multer')
const Papa = require('papaparse')
const { detectVendor } = require('../utils/vendorDetector')
const { generateExcel } = require('../utils/excelGenerator')

const router = express.Router()

// Configure multer for Vercel (memory storage with size limits)
const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 10 * 1024 * 1024, // 10MB limit
    files: 1,
  },
})

// Add error handling middleware for multer
const multerErrorHandler = (err, req, res, next) => {
  if (err instanceof multer.MulterError) {
    if (err.code === 'LIMIT_FILE_SIZE') {
      return res.status(400).json({
        success: false,
        message: {
          en: 'File too large. Maximum size is 10MB.',
          ja: 'ファイルが大きすぎます。最大サイズは10MBです。',
        },
      })
    }
    return res.status(400).json({
      success: false,
      message: {
        en: 'File upload error.',
        ja: 'ファイルアップロードエラー。',
      },
    })
  }
  next(err)
}

router.post('/convert', (req, res, next) => {
  upload.single('file')(req, res, err => {
    if (err) {
      return multerErrorHandler(err, req, res, next)
    }
    handleConvert(req, res, next)
  })
})

async function handleConvert(req, res, next) {
  try {
    console.log('Request received')
    console.log('File:', req.file ? 'Present' : 'Missing')
    console.log('Body:', req.body)

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

    // Validate file type
    if (!req.file.originalname.toLowerCase().endsWith('.csv')) {
      return res.status(400).json({
        success: false,
        message: {
          en: 'Invalid file type. Please upload a CSV file.',
          ja: '無効なファイル形式です。CSVファイルをアップロードしてください。',
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
      console.error('CSV Parse errors:', parseResult.errors)
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

    if (!csvData || csvData.length === 0) {
      return res.status(400).json({
        success: false,
        message: {
          en: 'CSV file is empty.',
          ja: 'CSVファイルが空です。',
        },
      })
    }

    const headers = Object.keys(csvData[0] || {})
    console.log('CSV Headers:', headers)

    // Detect vendor
    const vendorResult = detectVendor(req.file.originalname, headers)
    console.log('Vendor detection:', vendorResult)

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

    console.log('Sending file:', filename)

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
    res.status(500).json({
      success: false,
      message: {
        en: 'Internal server error during conversion.',
        ja: '変換中に内部サーバーエラーが発生しました。',
      },
      error: error.message,
    })
  }
}

module.exports = router
