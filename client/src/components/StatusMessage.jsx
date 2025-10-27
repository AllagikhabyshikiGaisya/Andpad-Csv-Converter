function StatusMessage({ status, message }) {
  const getStatusConfig = () => {
    switch (status) {
      case 'processing':
        return {
          icon: '⏳',
          bgColor: 'bg-blue-50',
          borderColor: 'border-blue-500',
          textColor: 'text-blue-800',
        }
      case 'success':
        return {
          icon: '✅',
          bgColor: 'bg-green-50',
          borderColor: 'border-green-500',
          textColor: 'text-green-800',
        }
      case 'error':
        return {
          icon: '⚠️',
          bgColor: 'bg-red-50',
          borderColor: 'border-red-500',
          textColor: 'text-red-800',
        }
      default:
        return {
          icon: 'ℹ️',
          bgColor: 'bg-gray-50',
          borderColor: 'border-gray-500',
          textColor: 'text-gray-800',
        }
    }
  }

  const config = getStatusConfig()

  return (
    <div
      className={`mt-6 p-4 rounded-lg border-2 ${config.bgColor} ${config.borderColor} transition-all duration-300`}
    >
      <div className="flex items-start">
        <span className="text-2xl mr-3">{config.icon}</span>
        <div className="flex-1">
          <p className={`font-semibold ${config.textColor} mb-1`}>
            {message.en}
          </p>
          <p className={`text-sm ${config.textColor}`}>{message.ja}</p>
        </div>
      </div>
    </div>
  )
}

export default StatusMessage
