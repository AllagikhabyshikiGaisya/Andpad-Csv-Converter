function DownloadButton({ url, filename }) {
  const handleDownload = () => {
    const link = document.createElement('a')
    link.href = url
    link.download = decodeURIComponent(filename)
    document.body.appendChild(link)
    link.click()
    document.body.removeChild(link)
  }

  return (
    <div className="mt-6">
      <button
        onClick={handleDownload}
        className="w-full py-3 px-6 bg-green-600 text-white rounded-lg font-semibold text-lg hover:bg-green-700 active:scale-95 transition-all duration-200 flex items-center justify-center"
      >
        <span className="text-xl mr-2">⬇️</span>
        Download Converted File / 変換後ファイルをダウンロード
      </button>
      <p className="text-xs text-gray-500 text-center mt-2">File: {filename}</p>
    </div>
  )
}

export default DownloadButton
