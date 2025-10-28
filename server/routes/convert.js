const express = require('express')
const multer = require('multer')
const Papa = require('papaparse')
const XLSX = require('xlsx')
const iconv = require('iconv-lite')
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
    console.log('File filter - Original name:', file.originalname)
    console.log('File filter - MIME type:', file.mimetype)

    // Fix filename encoding issues for display purposes only
    if (file.originalname) {
      try {
        const buffer = Buffer.from(file.originalname, 'latin1')
        const fixed = buffer.toString('utf8')
        // Only apply fix if it results in valid Japanese characters
        if (/[ぁ-ん]|[ァ-ヶ]|[一-龯]/.test(fixed)) {
          file.originalname = fixed
          console.log('Fixed filename:', file.originalname)
        }
      } catch (e) {
        console.log('Filename encoding fix skipped')
      }
    }

    const allowedTypes = [
      'text/csv',
      'application/vnd.ms-excel',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    ]

    const fileName = file.originalname.toLowerCase()
    const isValidExtension =
      fileName.endsWith('.csv') ||
      fileName.endsWith('.xlsx') ||
      fileName.endsWith('.xls')

    const isValidMimeType = allowedTypes.includes(file.mimetype)

    console.log('Validation - Extension valid:', isValidExtension)
    console.log('Validation - MIME valid:', isValidMimeType)

    if (isValidMimeType || isValidExtension) {
      console.log('✓ File accepted')
      cb(null, true)
    } else {
      console.log('✗ File rejected')
      cb(new Error('Invalid file type. Only CSV and Excel files are allowed.'))
    }
  },
})

// Helper function to detect and decode CSV encoding
function decodeCSV(buffer) {
  const encodings = [
    'utf-8',
    'shift-jis',
    'euc-jp',
    'iso-2022-jp',
    'windows-1252',
  ]

  console.log('Attempting to detect CSV encoding...')

  // Try UTF-8 first
  let content = buffer.toString('utf-8')

  // Check for mojibake indicators
  const hasMojibake =
    content.includes('�') ||
    (content.includes('ã') && content.includes('â') && content.includes('¯'))

  if (!hasMojibake) {
    console.log('✓ UTF-8 encoding detected')
    return content
  }

  console.log('⚠️ Mojibake detected in UTF-8, trying other encodings...')

  // Try other encodings
  for (const encoding of encodings.slice(1)) {
    try {
      const decoded = iconv.decode(buffer, encoding)

      // Check if decoded content looks valid (has common Japanese characters)
      const hasValidJapanese = /[あ-ん]|[ア-ン]|[一-龯]/.test(decoded)
      const noMojibake = !decoded.includes('�')

      if (hasValidJapanese && noMojibake) {
        console.log(`✓ Successfully decoded as ${encoding}`)
        return decoded
      }
    } catch (e) {
      console.log(`  ${encoding} failed:`, e.message)
    }
  }

  // Fallback to Shift-JIS (most common for Japanese Windows)
  console.log('⚠️ Using Shift-JIS as fallback')
  try {
    return iconv.decode(buffer, 'shift-jis')
  } catch (e) {
    console.log('⚠️ All encodings failed, using UTF-8')
    return content
  }
}

router.post('/convert', upload.single('file'), async (req, res) => {
  console.log('\n=== NEW CONVERSION REQUEST ===')
  console.log('Timestamp:', new Date().toISOString())

  try {
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

    // Additional filename encoding fix
    if (
      filename.includes('ã') ||
      filename.includes('¯') ||
      filename.includes('â')
    ) {
      console.log('⚠️ Detected garbled filename, attempting to fix...')
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
        console.log('Could not fix filename encoding')
      }
    }

    console.log('File received:', filename)
    console.log('File size:', fileSize, 'bytes')
    console.log('MIME type:', req.file.mimetype)

    let csvData = []
    let headers = []

    if (filename.toLowerCase().endsWith('.csv')) {
      console.log('Processing as CSV file...')

      // Decode with proper encoding detection
      const fileContent = decodeCSV(fileBuffer)

      console.log('File content length:', fileContent.length)
      console.log('First 200 chars:', fileContent.substring(0, 200))

      const parseResult = Papa.parse(fileContent, {
        header: true,
        skipEmptyLines: true,
        dynamicTyping: false,
      })

      if (parseResult.errors.length > 0) {
        console.error('CSV parsing errors:', parseResult.errors)
      }

      csvData = parseResult.data
      headers = parseResult.meta.fields || []

      console.log('✓ CSV parsed successfully')
      console.log('Total rows:', csvData.length)
      console.log('Headers:', headers)
    } else if (
      filename.toLowerCase().endsWith('.xlsx') ||
      filename.toLowerCase().endsWith('.xls')
    ) {
      console.log('Processing as Excel file...')

      const workbook = XLSX.read(fileBuffer, {
        type: 'buffer',
        cellDates: true,
        cellNF: false,
        cellText: false,
      })

      const sheetName = workbook.SheetNames[0]
      const worksheet = workbook.Sheets[sheetName]

      console.log('✓ Excel parsed successfully')
      console.log('Sheet name:', sheetName)

      // Convert to JSON with all data
      csvData = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
        defval: '',
        raw: false,
        dateNF: 'yyyy/mm/dd',
      })

      console.log('Total rows (including header):', csvData.length)

      if (csvData.length > 0) {
        headers = csvData[0]

        // Convert array format to object format
        csvData = csvData.slice(1).map(row => {
          const obj = {}
          headers.forEach((header, index) => {
            obj[header] = row[index] || ''
          })
          return obj
        })

        console.log('Data rows after header removal:', csvData.length)
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

    const outputFilename = `ANDPAD_${
      detection.mapping.vendor
    }_${Date.now()}.xlsx`

    const encodedFilename = encodeURIComponent(outputFilename)
    const asciiSafeFilename = outputFilename.replace(/[^\x00-\x7F]/g, '_')

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
