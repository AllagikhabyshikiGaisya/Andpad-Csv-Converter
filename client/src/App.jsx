import axios from 'axios'
import { useState } from 'react'
import DownloadButton from './components/DownloadButton'
import StatusMessage from './components/StatusMessage'
import UploadZone from './components/UploadZone'

function App() {
  const [file, setFile] = useState(null)
  const [status, setStatus] = useState('idle') // idle, processing, success, error
  const [message, setMessage] = useState({ en: '', ja: '' })
  const [downloadData, setDownloadData] = useState(null)

  const API_URL = import.meta.env.VITE_API_URL || '/api'

  const handleFileSelect = selectedFile => {
    setFile(selectedFile)
    setStatus('idle')
    setMessage({ en: '', ja: '' })
    setDownloadData(null)
  }

  const handleConvert = async () => {
    if (!file) {
      setStatus('error')
      setMessage({
        en: 'Please select a file first.',
        ja: 'まずファイルを選択してください。',
      })
      return
    }

    setStatus('processing')
    setMessage({
      en: 'Processing your file...',
      ja: 'ファイルを処理中...',
    })

    const formData = new FormData()
    formData.append('file', file)
    console.log('File to upload:', file.name, file.size, file.type)
    console.log('API URL:', `${API_URL}/convert`)
    try {
      const response = await axios.post(`${API_URL}/convert`, formData, {
        headers: {
          'Content-Type': 'multipart/form-data',
        },
        responseType: 'blob',
      })

      // Check if response is actually an error (JSON instead of blob)
      const contentType = response.headers['content-type']
      if (contentType && contentType.includes('application/json')) {
        const text = await response.data.text()
        const errorData = JSON.parse(text)
        throw new Error(errorData.message.ja)
      }

      setStatus('success')
      setMessage({
        en: 'Conversion successful!',
        ja: '変換に成功しました！',
      })

      // Prepare download
      const blob = new Blob([response.data], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      })
      const url = window.URL.createObjectURL(blob)
      const filename =
        response.headers['content-disposition']
          ?.split('filename=')[1]
          ?.replace(/"/g, '') || 'converted.xlsx'

      setDownloadData({ url, filename })
    } catch (error) {
      setStatus('error')

      if (error.response && error.response.data) {
        // Try to parse error message
        try {
          const text = await error.response.data.text()
          const errorData = JSON.parse(text)
          setMessage(errorData.message)
        } catch {
          setMessage({
            en: 'Error: Please check file format.',
            ja: 'エラー：ファイル形式を確認してください。',
          })
        }
      } else {
        setMessage({
          en: 'Network error. Please try again.',
          ja: 'ネットワークエラー。もう一度お試しください。',
        })
      }
    }
  }

  return (
    <div className="min-h-screen bg-white flex flex-col items-center justify-center p-4">
      <div className="w-full max-w-2xl">
        {/* Header */}
        <div className="text-center mb-8">
          <h1 className="text-4xl font-bold text-black mb-2">
            ANDPAD CSV Converter
          </h1>
          <p className="text-lg text-gray-700">
            ANDPADインポート用CSVコンバーター
          </p>
        </div>

        {/* Main Card */}
        <div className="bg-white border-2 border-black rounded-lg p-8 shadow-lg">
          <UploadZone onFileSelect={handleFileSelect} selectedFile={file} />

          <button
            onClick={handleConvert}
            disabled={!file || status === 'processing'}
            className={`w-full mt-6 py-3 px-6 rounded-lg font-semibold text-lg transition-all duration-200
              ${
                !file || status === 'processing'
                  ? 'bg-gray-300 text-gray-500 cursor-not-allowed'
                  : 'bg-black text-white hover:bg-gray-800 active:scale-95'
              }`}
          >
            {status === 'processing' ? (
              <>
                <span className="inline-block mr-2">⏳</span>
                Processing / 処理中...
              </>
            ) : (
              <>Convert / 変換する</>
            )}
          </button>

          {status !== 'idle' && (
            <StatusMessage status={status} message={message} />
          )}

          {downloadData && (
            <DownloadButton
              url={downloadData.url}
              filename={downloadData.filename}
            />
          )}
        </div>

        {/* Supported Vendors */}
        <div className="mt-8 text-center">
          <p className="text-sm text-gray-600 mb-2">
            Supported Vendors / 対応業者:
          </p>
          <div className="flex flex-wrap justify-center gap-2">
            {[
              'オメガジャパン',
              'トキワシステム',
              'カネカ建材',
              'ミヤコ電設',
              'ナカジマ設備',
              'ハタケヤマ商会',
              'エムテック',
              'サンリツ工業',
              'リョウエイ建材',
            ].map(vendor => (
              <span
                key={vendor}
                className="px-3 py-1 bg-gray-100 text-black text-xs rounded-full border border-gray-300"
              >
                {vendor}
              </span>
            ))}
          </div>
        </div>
      </div>
    </div>
  )
}

export default App
