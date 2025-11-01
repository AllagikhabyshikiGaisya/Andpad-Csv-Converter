const express = require('express')
const multer = require('multer')
const Papa = require('papaparse')
const XLSX = require('xlsx')
const iconv = require('iconv-lite')
const { detectVendor } = require('../utils/vendorDetector')
const {
  generateExcel,
  generateCombinedExcel,
} = require('../utils/excelGenerator')

const router = express.Router()

// Configure multer for MULTIPLE file uploads
const storage = multer.memoryStorage()
const upload = multer({
  storage: storage,
  limits: {
    fileSize: 10 * 1024 * 1024, // 10MB limit per file
    files: 10, // Max 10 files at once
  },
  fileFilter: (req, file, cb) => {
    console.log('File filter - Original name:', file.originalname)
    console.log('File filter - MIME type:', file.mimetype)

    // Fix filename encoding issues
    if (file.originalname) {
      try {
        const buffer = Buffer.from(file.originalname, 'latin1')
        const fixed = buffer.toString('utf8')
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

  let content = buffer.toString('utf-8')
  const hasMojibake =
    content.includes('�') ||
    (content.includes('ã') && content.includes('â') && content.includes('¯'))

  if (!hasMojibake) {
    console.log('✓ UTF-8 encoding detected')
    return content
  }

  console.log('⚠️ Mojibake detected in UTF-8, trying other encodings...')

  for (const encoding of encodings.slice(1)) {
    try {
      const decoded = iconv.decode(buffer, encoding)
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

  console.log('⚠️ Using Shift-JIS as fallback')
  try {
    return iconv.decode(buffer, 'shift-jis')
  } catch (e) {
    console.log('⚠️ All encodings failed, using UTF-8')
    return content
  }
}

// Helper function to parse a single file
async function parseFile(file) {
  let filename = file.originalname
  const fileBuffer = file.buffer

  // Fix filename encoding
  if (
    filename.includes('ã') ||
    filename.includes('¯') ||
    filename.includes('â')
  ) {
    try {
      const buffer = Buffer.from(filename, 'latin1')
      const fixed = buffer.toString('utf8')
      if (
        fixed.includes('ク') ||
        fixed.includes('産') ||
        fixed.includes('請')
      ) {
        filename = fixed
      }
    } catch (e) {
      // Ignore
    }
  }

  console.log('Parsing file:', filename)

  let csvData = []
  let headers = []

  if (filename.toLowerCase().endsWith('.csv')) {
    const fileContent = decodeCSV(fileBuffer)
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
    console.log('✓ CSV parsed:', csvData.length, 'rows')
  } else if (
    filename.toLowerCase().endsWith('.xlsx') ||
    filename.toLowerCase().endsWith('.xls')
  ) {
    const workbook = XLSX.read(fileBuffer, {
      type: 'buffer',
      cellDates: true,
      cellNF: false,
      cellText: false,
    })

    const sheetName = workbook.SheetNames[0]
    const worksheet = workbook.Sheets[sheetName]

    csvData = XLSX.utils.sheet_to_json(worksheet, {
      header: 1,
      defval: '',
      raw: false,
      dateNF: 'yyyy/mm/dd',
    })

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
    console.log('✓ Excel parsed:', csvData.length, 'rows')
  }

  return { filename, csvData, headers }
}

// Main conversion route - supports SINGLE or MULTIPLE files
router.post('/convert', upload.array('files', 10), async (req, res) => {
  console.log('\n=== NEW CONVERSION REQUEST ===')
  console.log('Timestamp:', new Date().toISOString())

  try {
    // Check if files were uploaded
    const files = req.files || (req.file ? [req.file] : [])

    if (!files || files.length === 0) {
      console.error('❌ No files uploaded')
      return res.status(400).json({
        success: false,
        message: {
          en: 'No files uploaded',
          ja: 'ファイルがアップロードされていません',
        },
      })
    }

    // Get output format from request body (default: xlsx)
    const outputFormat = req.body.outputFormat || 'xlsx'
    console.log('Output format:', outputFormat)
    console.log('Number of files:', files.length)

    // SINGLE FILE PROCESSING
    if (files.length === 1) {
      console.log('\n--- SINGLE FILE MODE ---')
      const file = files[0]

      const { filename, csvData, headers } = await parseFile(file)

      if (!csvData || csvData.length === 0) {
        return res.status(400).json({
          success: false,
          message: {
            en: 'File is empty or invalid',
            ja: 'ファイルが空か無効です',
          },
        })
      }

      console.log('Sample data:', JSON.stringify(csvData[0], null, 2))

      // Detect vendor
      const detection = detectVendor(filename, headers)

      if (!detection.detected) {
        console.error('❌ Vendor detection failed')
        return res.status(400).json({
          success: false,
          message: {
            en: 'Could not identify vendor',
            ja: '業者を識別できませんでした',
          },
          details: {
            filename: filename,
            headersSample: headers.slice(0, 5),
          },
        })
      }

      console.log('✓ Vendor:', detection.mapping.vendor)

      // Generate Excel/CSV
      const result = generateExcel(csvData, detection.mapping, outputFormat)

      if (!result.success) {
        console.error('❌ Generation failed:', result.error)
        return res.status(400).json({
          success: false,
          message: {
            en: `Generation failed: ${result.error}`,
            ja: `生成失敗: ${result.error}`,
          },
        })
      }

      console.log('✓ File generated successfully')
      console.log('Output rows:', result.rowCount)

      // Send file
      const outputFilename = `ANDPAD_${
        detection.mapping.vendor
      }_${Date.now()}.${result.fileExtension}`
      const encodedFilename = encodeURIComponent(outputFilename)
      const asciiSafeFilename = outputFilename.replace(/[^\x00-\x7F]/g, '_')

      const contentType =
        result.fileExtension === 'csv'
          ? 'text/csv; charset=utf-8'
          : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

      res.setHeader('Content-Type', contentType)
      res.setHeader(
        'Content-Disposition',
        `attachment; filename="${asciiSafeFilename}"; filename*=UTF-8''${encodedFilename}`
      )
      res.setHeader('X-Vendor', encodeURIComponent(detection.mapping.vendor))
      res.setHeader('X-Row-Count', result.rowCount)

      console.log('Sending file:', outputFilename)
      console.log('=== CONVERSION COMPLETED ===\n')

      return res.send(result.buffer)
    }

    // MULTIPLE FILES PROCESSING (BATCH MODE)
    console.log('\n--- BATCH MODE: Processing', files.length, 'files ---')

    const filesData = []

    // Parse all files
    for (let i = 0; i < files.length; i++) {
      const file = files[i]
      console.log(
        `\n[${i + 1}/${files.length}] Processing: ${file.originalname}`
      )

      try {
        const { filename, csvData, headers } = await parseFile(file)

        if (!csvData || csvData.length === 0) {
          console.warn(`⚠️ Skipping empty file: ${filename}`)
          continue
        }

        // Detect vendor
        const detection = detectVendor(filename, headers)

        if (!detection.detected) {
          console.warn(`⚠️ Could not detect vendor for: ${filename}`)
          continue
        }

        console.log(`✓ Detected: ${detection.mapping.vendor}`)

        filesData.push({
          csvData: csvData,
          mapping: detection.mapping,
        })
      } catch (error) {
        console.error(
          `❌ Error processing ${file.originalname}:`,
          error.message
        )
        continue
      }
    }

    if (filesData.length === 0) {
      return res.status(400).json({
        success: false,
        message: {
          en: 'No valid files could be processed',
          ja: '有効なファイルが処理できませんでした',
        },
      })
    }

    console.log(`\n✓ Successfully parsed ${filesData.length} files`)

    // Generate combined file
    const result = generateCombinedExcel(filesData, outputFormat)

    if (!result.success) {
      console.error('❌ Combined generation failed:', result.error)
      return res.status(400).json({
        success: false,
        message: {
          en: `Generation failed: ${result.error}`,
          ja: `生成失敗: ${result.error}`,
        },
      })
    }

    console.log('✓ Combined file generated successfully')
    console.log('Total rows:', result.rowCount)
    console.log('Vendors processed:', result.vendorCount)

    // Send combined file
    const outputFilename = `ANDPAD_Combined_${Date.now()}.${
      result.fileExtension
    }`
    const encodedFilename = encodeURIComponent(outputFilename)
    const asciiSafeFilename = outputFilename.replace(/[^\x00-\x7F]/g, '_')

    const contentType =
      result.fileExtension === 'csv'
        ? 'text/csv; charset=utf-8'
        : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

    res.setHeader('Content-Type', contentType)
    res.setHeader(
      'Content-Disposition',
      `attachment; filename="${asciiSafeFilename}"; filename*=UTF-8''${encodedFilename}`
    )
    res.setHeader('X-Vendor-Count', result.vendorCount)
    res.setHeader('X-Row-Count', result.rowCount)

    console.log('Sending combined file:', outputFilename)
    console.log('=== BATCH CONVERSION COMPLETED ===\n')

    return res.send(result.buffer)
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
