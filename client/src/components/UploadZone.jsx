import { useRef, useState } from 'react'

function UploadZone({ onFileSelect, selectedFile }) {
  const [isDragging, setIsDragging] = useState(false)
  const fileInputRef = useRef(null)

  const isValidFile = file => {
    const fileName = file.name.toLowerCase()
    return (
      fileName.endsWith('.csv') ||
      fileName.endsWith('.xlsx') ||
      fileName.endsWith('.xls')
    )
  }

  const handleDragEnter = e => {
    e.preventDefault()
    e.stopPropagation()
    setIsDragging(true)
  }

  const handleDragLeave = e => {
    e.preventDefault()
    e.stopPropagation()
    setIsDragging(false)
  }

  const handleDragOver = e => {
    e.preventDefault()
    e.stopPropagation()
  }

  const handleDrop = e => {
    e.preventDefault()
    e.stopPropagation()
    setIsDragging(false)

    const files = e.dataTransfer.files
    if (files && files.length > 0) {
      const file = files[0]
      if (isValidFile(file)) {
        onFileSelect(file)
      } else {
        alert(
          'Please upload a CSV or Excel file\nCSVまたはExcelファイルをアップロードしてください'
        )
      }
    }
  }

  const handleFileChange = e => {
    const files = e.target.files
    if (files && files.length > 0) {
      const file = files[0]
      if (isValidFile(file)) {
        onFileSelect(file)
      } else {
        alert(
          'Please upload a CSV or Excel file\nCSVまたはExcelファイルをアップロードしてください'
        )
      }
    }
  }

  const handleClick = () => {
    fileInputRef.current?.click()
  }

  const getFileIcon = () => {
    if (!selectedFile) return '📄'
    const fileName = selectedFile.name.toLowerCase()
    if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
      return '📊'
    }
    return '📄'
  }

  return (
    <div
      onClick={handleClick}
      onDragEnter={handleDragEnter}
      onDragOver={handleDragOver}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
      className={`border-2 border-dashed rounded-lg p-12 text-center cursor-pointer transition-all duration-200
        ${
          isDragging
            ? 'border-black bg-gray-50 scale-105'
            : 'border-gray-400 hover:border-black hover:bg-gray-50'
        }`}
    >
      <input
        ref={fileInputRef}
        type="file"
        accept=".csv,.xlsx,.xls,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        onChange={handleFileChange}
        className="hidden"
      />

      <div className="text-5xl mb-4">{getFileIcon()}</div>

      {selectedFile ? (
        <div>
          <p className="text-lg font-semibold text-black mb-1">
            {selectedFile.name}
          </p>
          <p className="text-sm text-gray-600">
            {(selectedFile.size / 1024).toFixed(2)} KB
          </p>
          <p className="text-xs text-gray-500 mt-2">
            Click to change file / クリックして変更
          </p>
        </div>
      ) : (
        <div>
          <p className="text-lg font-semibold text-black mb-2">
            Upload CSV or Excel / CSV・Excelをアップロード
          </p>
          <p className="text-sm text-gray-600">
            Drag & drop or click to browse
          </p>
          <p className="text-sm text-gray-600">
            ドラッグ&ドロップまたはクリックして選択
          </p>
          <p className="text-xs text-gray-500 mt-3">
            Supported: .csv, .xlsx, .xls
          </p>
        </div>
      )}
    </div>
  )
}

export default UploadZone
