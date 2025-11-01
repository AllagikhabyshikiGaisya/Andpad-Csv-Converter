import axios from 'axios'
import { useState } from 'react'
import DownloadButton from './components/DownloadButton'
import FormatSelector from './components/FormatSelector'
import StatusMessage from './components/StatusMessage'
import UploadZone from './components/UploadZone'

function App() {
  const [files, setFiles] = useState([]) // Changed to array for multiple files
  const [outputFormat, setOutputFormat] = useState('xlsx') // Phase 3: Format selection
  const [status, setStatus] = useState('idle') // idle, processing, success, error
  const [message, setMessage] = useState({ en: '', ja: '' })
  const [downloadData, setDownloadData] = useState(null)

  const API_URL = import.meta.env.VITE_API_URL || '/api'

  const handleFileSelect = selectedFiles => {
    // Support both single and multiple files
    const fileArray = Array.isArray(selectedFiles)
      ? selectedFiles
      : [selectedFiles]
    setFiles(fileArray)
    setStatus('idle')
    setMessage({ en: '', ja: '' })
    setDownloadData(null)
  }

  const handleConvert = async () => {
    if (!files || files.length === 0) {
      setStatus('error')
      setMessage({
        en: 'Please select at least one file.',
        ja: '少なくとも1つのファイルを選択してください。',
      })
      return
    }

    setStatus('processing')
    setMessage({
      en:
        files.length > 1
          ? `Processing ${files.length} files...`
          : 'Processing your file...',
      ja:
        files.length > 1
          ? `${files.length}個のファイルを処理中...`
          : 'ファイルを処理中...',
    })

    const formData = new FormData()

    // Add all files
    files.forEach(file => {
      formData.append('files', file)
    })

    // Add output format selection
    formData.append('outputFormat', outputFormat)

    console.log('Files to upload:', files.length)
    console.log('Output format:', outputFormat)
    console.log('API URL:', `${API_URL}/convert`)

    try {
      const response = await axios.post(`${API_URL}/convert`, formData, {
        headers: {
          'Content-Type': 'multipart/form-data',
        },
        responseType: 'blob',
      })

      // Check if response is actually an error
      const contentType = response.headers['content-type']
      if (contentType && contentType.includes('application/json')) {
        const text = await response.data.text()
        const errorData = JSON.parse(text)
        throw new Error(errorData.message.ja)
      }

      setStatus('success')
      setMessage({
        en:
          files.length > 1
            ? `Successfully converted ${files.length} files!`
            : 'Conversion successful!',
        ja:
          files.length > 1
            ? `${files.length}個のファイルの変換に成功しました！`
            : '変換に成功しました！',
      })

      // Prepare download
      const mimeType =
        outputFormat === 'csv'
          ? 'text/csv; charset=utf-8'
          : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

      const blob = new Blob([response.data], { type: mimeType })
      const url = window.URL.createObjectURL(blob)

      // Extract filename from Content-Disposition header
      let filename = `converted.${outputFormat}`
      const contentDisposition = response.headers['content-disposition']

      if (contentDisposition) {
        const filenameStarMatch = contentDisposition.match(
          /filename\*=UTF-8''(.+?)(?:;|$)/
        )
        if (filenameStarMatch) {
          filename = decodeURIComponent(filenameStarMatch[1])
        } else {
          const filenameMatch = contentDisposition.match(/filename="?([^"]+)"?/)
          if (filenameMatch) {
            filename = filenameMatch[1]
          }
        }
      }

      setDownloadData({ url, filename })
    } catch (error) {
      setStatus('error')

      if (error.response && error.response.data) {
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
          {/* Phase 3 Badge */}
          <div className="mt-2 inline-flex items-center gap-2">
            <span className="px-3 py-1 bg-blue-100 text-blue-800 text-xs font-semibold rounded-full">
              ✨ Multi-file Support
            </span>
            <span className="px-3 py-1 bg-green-100 text-green-800 text-xs font-semibold rounded-full">
              📤 Excel & CSV Export
            </span>
          </div>
        </div>

        {/* Main Card */}
        <div className="bg-white border-2 border-black rounded-lg p-8 shadow-lg">
          <UploadZone onFileSelect={handleFileSelect} selectedFiles={files} />

          {/* Phase 3: Format Selector */}
          <FormatSelector
            selectedFormat={outputFormat}
            onFormatChange={setOutputFormat}
          />

          <button
            onClick={handleConvert}
            disabled={files.length === 0 || status === 'processing'}
            className={`w-full mt-6 py-3 px-6 rounded-lg font-semibold text-lg transition-all duration-200
              ${
                files.length === 0 || status === 'processing'
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
              <>
                {files.length > 1
                  ? `Convert ${files.length} Files / ${files.length}個のファイルを変換`
                  : 'Convert / 変換する'}
              </>
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
              'ナンセイ',
              'クリーン産業',
              '三高産業',
              '北恵株式会社',
              'オメガジャパン',
              'トキワシステム',
              'ALLAGI',
              '大萬',
              '髙菱管理',
              'ナカザワ建販',
            ]
              .sort()
              .map(vendor => (
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
