function FormatSelector({ selectedFormat, onFormatChange }) {
  return (
    <div className="mt-6 p-4 bg-gray-50 rounded-lg border border-gray-200">
      <p className="text-sm font-semibold text-gray-700 mb-3">
        📤 Export Format / 出力形式:
      </p>

      <div className="flex gap-4">
        <label className="flex items-center cursor-pointer">
          <input
            type="radio"
            name="format"
            value="xlsx"
            checked={selectedFormat === 'xlsx'}
            onChange={e => onFormatChange(e.target.value)}
            className="mr-2 w-4 h-4 cursor-pointer"
          />
          <div className="flex items-center">
            <span className="text-2xl mr-2">📊</span>
            <div>
              <p className="font-semibold text-black">Excel (.xlsx)</p>
              <p className="text-xs text-gray-600">Recommended / 推奨</p>
            </div>
          </div>
        </label>

        <label className="flex items-center cursor-pointer">
          <input
            type="radio"
            name="format"
            value="csv"
            checked={selectedFormat === 'csv'}
            onChange={e => onFormatChange(e.target.value)}
            className="mr-2 w-4 h-4 cursor-pointer"
          />
          <div className="flex items-center">
            <span className="text-2xl mr-2">📄</span>
            <div>
              <p className="font-semibold text-black">CSV (UTF-8)</p>
              <p className="text-xs text-gray-600">
                Text format / テキスト形式
              </p>
            </div>
          </div>
        </label>
      </div>
    </div>
  )
}

export default FormatSelector
