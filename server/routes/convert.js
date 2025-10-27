const express = require('express')
const multer = require('multer')
const Papa = require('papaparse')
const XLSX = require('xlsx')
const { detectVendor } = require('../utils/vendorDetector')
const { generateExcel } = require('../utils/excelGenerator')

const router = express.Router()

// Configure multer for file uploads with proper encoding
const storage = multer.memoryStorage()
const upload = multer({
  storage: storage,
  limits: {
    fileSize: 10 * 1024 * 1024, // 10MB limit
  },
  fileFilter: (req, file, cb) => {
    // Fix filename encoding issues
    if (file.originalname) {
      try {
        // Attempt to fix mojibake (文字化け) - garbled Japanese characters
        // This happens when UTF-8 bytes are interpreted as Latin-1
        const buffer = Buffer.from(file.originalname, 'latin1')
        file.originalname = buffer.toString('utf8')
        console.log('Fixed filename encoding:', file.originalname)
      } catch (e) {
        console.log('Filename encoding fix not needed or failed:', e.message)
      }
    }

    const allowedTypes = [
      'text/csv',
      'application/vnd.ms-excel',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    ]

    if (
      allowedTypes.includes(file.mimetype) ||
      file.originalname.toLowerCase().endsWith('.csv') ||
      file.originalname.toLowerCase().endsWith('.xlsx')
    ) {
      cb(null, true)
    } else {
      cb(new Error('Invalid file type. Only CSV and Excel files are allowed.'))
    }
  },
})

router.post('/convert', upload.single('file'), async (req, res) => {
  console.log('\n=== NEW CONVERSION REQUEST ===')
  console.log('Timestamp:', new Date().toISOString())

  try {
    // Validate file
    if (!req.file) {
      console.error('❌ No file uploaded')
      return res.status(400).json({
        success: false,
        message: {
          en: 'No file uploaded',
          ja: 'ファイルがアップロードされていません',
        },
      })
    }

    let filename = req.file.originalname
    const fileBuffer = req.file.buffer
    const fileSize = req.file.size

    // Additional filename encoding fix if still garbled
    if (
      filename.includes('ã') ||
      filename.includes('¯') ||
      filename.includes('â')
    ) {
      console.log('⚠️ Detected garbled filename, attempting to fix...')
      console.log('Original:', filename)
      try {
        const buffer = Buffer.from(filename, 'latin1')
        const fixed = buffer.toString('utf8')
        if (
          fixed.includes('ク') ||
          fixed.includes('産') ||
          fixed.includes('請')
        ) {
          filename = fixed
          console.log('✓ Fixed filename:', filename)
        }
      } catch (e) {
        console.log('Could not fix filename encoding:', e.message)
      }
    }

    console.log('File received:', filename)
    console.log('File size:', fileSize, 'bytes')
    console.log('MIME type:', req.file.mimetype)

    let csvData = []
    let headers = []

    // Parse based on file type
    if (filename.toLowerCase().endsWith('.csv')) {
      console.log('Processing as CSV file...')

      const fileContent = fileBuffer.toString('utf-8')
      console.log('File content length:', fileContent.length)
      console.log('First 200 chars:', fileContent.substring(0, 200))

      const parseResult = Papa.parse(fileContent, {
        header: true,
        skipEmptyLines: true,
        dynamicTyping: false,
        encoding: 'utf-8',
      })

      if (parseResult.errors.length > 0) {
        console.error('CSV parsing errors:', parseResult.errors)
      }

      csvData = parseResult.data
      headers = parseResult.meta.fields || []

      console.log('✓ CSV parsed successfully')
      console.log('Total rows:', csvData.length)
      console.log('Headers:', headers)
    } else if (filename.toLowerCase().endsWith('.xlsx')) {
      console.log('Processing as Excel file...')

      const workbook = XLSX.read(fileBuffer, { type: 'buffer' })
      const sheetName = workbook.SheetNames[0]
      const worksheet = workbook.Sheets[sheetName]

      csvData = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
        defval: '',
        raw: false,
      })

      console.log('✓ Excel parsed successfully')
      console.log('Sheet name:', sheetName)
      console.log('Total rows:', csvData.length)

      // Convert to object format
      if (csvData.length > 0) {
        headers = csvData[0]
        csvData = csvData.slice(1).map(row => {
          const obj = {}
          headers.forEach((header, index) => {
            obj[header] = row[index] || ''
          })
          return obj
        })
      }

      console.log('Headers:', headers)
    } else {
      console.error('❌ Unsupported file type')
      return res.status(400).json({
        success: false,
        message: {
          en: 'Unsupported file type',
          ja: 'サポートされていないファイル形式です',
        },
      })
    }

    // Validate data
    if (!csvData || csvData.length === 0) {
      console.error('❌ No data found in file')
      return res.status(400).json({
        success: false,
        message: {
          en: 'File is empty or invalid',
          ja: 'ファイルが空か無効です',
        },
      })
    }

    console.log('Sample data (first row):', JSON.stringify(csvData[0], null, 2))

    // Detect vendor
    console.log('\n--- Starting Vendor Detection ---')
    const detection = detectVendor(filename, headers)

    if (!detection.detected) {
      console.error('❌ Vendor detection failed')
      console.error('Reason:', detection.error)

      return res.status(400).json({
        success: false,
        message: {
          en: 'Could not identify vendor. Please check filename or file format.',
          ja: '業者を識別できませんでした。ファイル名またはファイル形式を確認してください。',
        },
        details: {
          filename: filename,
          headersSample: headers.slice(0, 5),
        },
      })
    }

    console.log('✓✓✓ Vendor detected:', detection.mapping.vendor)
    console.log('Detection method:', detection.method)
    console.log('Custom parser:', detection.mapping.customParser)

    // Generate Excel
    console.log('\n--- Starting Excel Generation ---')
    const result = generateExcel(csvData, detection.mapping)

    if (!result.success) {
      console.error('❌ Excel generation failed:', result.error)

      return res.status(400).json({
        success: false,
        message: {
          en: `Generation failed: ${result.error}`,
          ja: `生成失敗: ${result.error}`,
        },
      })
    }

    console.log('✓✓✓ Excel generated successfully')
    console.log('Output rows:', result.rowCount)

    // Send file
    const outputFilename = `ANDPAD_${
      detection.mapping.vendor
    }_${Date.now()}.xlsx`

    // Encode filename for Content-Disposition header (support non-ASCII characters)
    const encodedFilename = encodeURIComponent(outputFilename)
    const asciiSafeFilename = outputFilename.replace(/[^\x00-\x7F]/g, '_') // Fallback for old browsers

    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    res.setHeader(
      'Content-Disposition',
      `attachment; filename="${asciiSafeFilename}"; filename*=UTF-8''${encodedFilename}`
    )
    res.setHeader('X-Vendor', encodeURIComponent(detection.mapping.vendor))
    res.setHeader('X-Row-Count', result.rowCount)

    console.log('Sending file:', outputFilename)
    console.log('=== CONVERSION COMPLETED SUCCESSFULLY ===\n')

    res.send(result.buffer)
  } catch (error) {
    console.error('❌ CONVERSION ERROR:', error.message)
    console.error('Stack trace:', error.stack)

    res.status(500).json({
      success: false,
      message: {
        en: `Error: ${error.message}`,
        ja: `エラー: ${error.message}`,
      },
      error: process.env.NODE_ENV === 'development' ? error.stack : undefined,
    })
  }
})

module.exports = router
