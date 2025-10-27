const express = require('express')
const cors = require('cors')
const path = require('path')
const convertRoute = require('./routes/convert')

const app = express()
const PORT = process.env.PORT || 3001

// CORS - Must be BEFORE routes
app.use(
  cors({
    origin: '*', // Allow all origins for now (we'll restrict later)
    methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization'],
    credentials: true,
  })
)

// Body parsing middleware
app.use(express.json())
app.use(express.urlencoded({ extended: true }))

// Routes
app.use('/api', convertRoute)

// Health check
app.get('/api/health', (req, res) => {
  res.json({
    status: 'ok',
    message: 'ANDPAD CSV Converter API is running',
    timestamp: new Date().toISOString(),
  })
})

// Root endpoint
app.get('/', (req, res) => {
  res.json({
    message: 'ANDPAD CSV Converter API',
    endpoints: ['/api/health', '/api/convert'],
  })
})

// Error handling middleware
app.use((err, req, res, next) => {
  console.error('Error:', err)
  res.status(500).json({
    success: false,
    message: {
      en: 'Internal server error',
      ja: '内部サーバーエラー',
    },
    error: process.env.NODE_ENV === 'development' ? err.message : undefined,
  })
})

module.exports = app
