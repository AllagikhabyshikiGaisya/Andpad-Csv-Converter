const express = require('express')
const cors = require('cors')
const path = require('path')
const convertRoute = require('./routes/convert')

const app = express()
const PORT = process.env.PORT || 3001

// CORS Configuration - MUST BE BEFORE ROUTES
const allowedOrigins = process.env.ALLOWED_ORIGINS
  ? process.env.ALLOWED_ORIGINS.split(',')
  : ['*']

app.use(
  cors({
    origin: function (origin, callback) {
      // Allow requests with no origin (like mobile apps or curl requests)
      if (!origin) return callback(null, true)

      if (
        allowedOrigins.indexOf('*') !== -1 ||
        allowedOrigins.indexOf(origin) !== -1
      ) {
        callback(null, true)
      } else {
        callback(new Error('Not allowed by CORS'))
      }
    },
    credentials: true,
    methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization'],
  })
)

// Middleware
app.use(express.json())
app.use(express.urlencoded({ extended: true }))

// Routes
app.use('/api', convertRoute)

// Health check
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', message: 'ANDPAD CSV Converter API is running' })
})

// Root endpoint
app.get('/', (req, res) => {
  res.json({ message: 'ANDPAD CSV Converter API' })
})

// Start server (only when not in Vercel)
if (process.env.NODE_ENV !== 'production') {
  app.listen(PORT, () => {
    console.log(`🚀 Server running on http://localhost:${PORT}`)
  })
}

module.exports = app
