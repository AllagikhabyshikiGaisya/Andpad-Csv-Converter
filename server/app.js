const express = require('express')
const cors = require('cors')
const path = require('path')
const convertRoute = require('./routes/convert')

const app = express()
const PORT = process.env.PORT || 3001

// CORS - Must be BEFORE routes
app.use(
  cors({
    origin: '*', // Allow all origins (restrict in production)
    methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization'],
    credentials: true,
  })
)

// Body parsing middleware
app.use(express.json())
app.use(express.urlencoded({ extended: true }))

// Request logging middleware
app.use((req, res, next) => {
  console.log(`${new Date().toISOString()} - ${req.method} ${req.path}`)
  next()
})

// Routes
app.use('/api', convertRoute)

// Health check
app.get('/api/health', (req, res) => {
  res.json({
    status: 'ok',
    message: 'ANDPAD CSV Converter API is running',
    timestamp: new Date().toISOString(),
    version: '1.0.0',
  })
})

// Root endpoint
app.get('/', (req, res) => {
  res.json({
    message: 'ANDPAD CSV Converter API',
    endpoints: {
      health: '/api/health',
      convert: '/api/convert (POST)',
    },
    status: 'running',
  })
})

// 404 handler
app.use((req, res) => {
  console.log(`404 - Route not found: ${req.method} ${req.path}`)
  res.status(404).json({
    success: false,
    message: {
      en: 'Route not found',
      ja: 'ルートが見つかりません',
    },
  })
})

// Error handling middleware
app.use((err, req, res, next) => {
  console.error('=== ERROR HANDLER ===')
  console.error('Error:', err.message)
  console.error('Stack:', err.stack)
  console.error('Path:', req.path)
  console.error('Method:', req.method)

  // Handle multer errors
  if (err.code === 'LIMIT_FILE_SIZE') {
    return res.status(400).json({
      success: false,
      message: {
        en: 'File too large. Maximum size is 10MB.',
        ja: 'ファイルが大きすぎます。最大サイズは10MBです。',
      },
    })
  }

  if (err.message.includes('Invalid file type')) {
    return res.status(400).json({
      success: false,
      message: {
        en: 'Invalid file type. Only CSV and Excel files are allowed.',
        ja: '無効なファイル形式です。CSVまたはExcelファイルのみ許可されています。',
      },
    })
  }

  res.status(500).json({
    success: false,
    message: {
      en: 'Internal server error',
      ja: '内部サーバーエラー',
    },
    error: process.env.NODE_ENV === 'development' ? err.message : undefined,
  })
})

// Start server (only if not in Vercel serverless environment)
if (require.main === module) {
  app.listen(PORT, () => {
    console.log('=================================')
    console.log('ANDPAD CSV Converter API')
    console.log(`Server running on port ${PORT}`)
    console.log(`http://localhost:${PORT}`)
    console.log('=================================')
  })
}

module.exports = app
