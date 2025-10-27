const express = require('express')
const cors = require('cors')
const path = require('path')
const convertRoute = require('./routes/convert')

const app = express()
const PORT = process.env.PORT || 3001

// Middleware
app.use(cors())
app.use(express.json())
app.use(express.urlencoded({ extended: true }))

// Routes
app.use('/api', convertRoute)

// Health check
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', message: 'ANDPAD CSV Converter API is running' })
})

// Start server (only when not in Vercel)
if (process.env.NODE_ENV !== 'production') {
  app.listen(PORT, () => {
    console.log(`🚀 Server running on http://localhost:${PORT}`)
  })
}

module.exports = app
