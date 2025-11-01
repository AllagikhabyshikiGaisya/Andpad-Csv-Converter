import { useRef, useState } from 'react'

function UploadZone({ onFileSelect, selectedFiles }) {
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

    const files = Array.from(e.dataTransfer.files)
    const validFiles = files.filter(isValidFile)

    if (validFiles.length > 0) {
      onFileSelect(validFiles)
    } else {
      alert(
        'Please upload CSV or Excel files only\nCSVまたはExcelファイルのみをアップロードしてください'
      )
    }
  }

  const handleFileChange = e => {
    const files = Array.from(e.target.files)
    const validFiles = files.filter(isValidFile)

    if (validFiles.length > 0) {
      onFileSelect(validFiles)
    } else {
      alert(
        'Please upload CSV or Excel files only\nCSVまたはExcelファイルのみをアップロードしてください'
      )
    }
  }

  const handleClick = () => {
    fileInputRef.current?.click()
  }

  const getFileIcon = () => {
    if (!selectedFiles || selectedFiles.length === 0) return '📄'
    if (selectedFiles.length > 1) return '📚'

    const fileName = selectedFiles[0].name.toLowerCase()
    if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
      return '📊'
    }
    return '📄'
  }

  const getTotalSize = () => {
    if (!selectedFiles || selectedFiles.length === 0) return 0
    return selectedFiles.reduce((sum, file) => sum + file.size, 0)
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
        multiple // Phase 3: Enable multiple file selection
        className="hidden"
      />

      <div className="text-5xl mb-4">{getFileIcon()}</div>

      {selectedFiles && selectedFiles.length > 0 ? (
        <div>
          {selectedFiles.length === 1 ? (
            <>
              <p className="text-lg font-semibold text-black mb-1">
                {selectedFiles[0].name}
              </p>
              <p className="text-sm text-gray-600">
                {(selectedFiles[0].size / 1024).toFixed(2)} KB
              </p>
            </>
          ) : (
            <>
              <p className="text-lg font-semibold text-black mb-2">
                {selectedFiles.length} Files Selected
              </p>
              <div className="max-h-32 overflow-y-auto mb-2">
                {selectedFiles.map((file, index) => (
                  <p key={index} className="text-sm text-gray-600 truncate">
                    {index + 1}. {file.name} ({(file.size / 1024).toFixed(1)}{' '}
                    KB)
                  </p>
                ))}
              </div>
              <p className="text-sm font-semibold text-gray-700">
                Total: {(getTotalSize() / 1024).toFixed(2)} KB
              </p>
            </>
          )}
          <p className="text-xs text-gray-500 mt-3">
            Click to change files / クリックして変更
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
          <p className="text-xs text-blue-600 font-semibold mt-3">
            ✨ Multiple files supported! / 複数ファイル対応！
          </p>
          <p className="text-xs text-gray-500 mt-1">
            Supported: .csv, .xlsx, .xls
          </p>
        </div>
      )}
    </div>
  )
}

export default UploadZone
