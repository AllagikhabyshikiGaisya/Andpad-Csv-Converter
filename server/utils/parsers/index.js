const { PARSER_REGISTRY } = require('./registry')

/**
 * Get parser instance for a vendor
 * @param {string} vendorName - Name of the vendor
 * @returns {BaseParser} Parser instance
 */
function getParser(vendorName) {
  const ParserClass = PARSER_REGISTRY[vendorName]

  if (!ParserClass) {
    throw new Error(`カスタムパーサーが見つかりません: ${vendorName}`)
  }

  // DEBUG: Check if it's a constructor
  console.log('DEBUG - Parser for:', vendorName)
  console.log('DEBUG - ParserClass type:', typeof ParserClass)
  console.log('DEBUG - Is constructor:', typeof ParserClass === 'function')

  if (typeof ParserClass !== 'function') {
    console.error('DEBUG - ParserClass value:', ParserClass)
    throw new Error(
      `ParserClass is not a constructor for vendor: ${vendorName}`
    )
  }

  return new ParserClass()
}

/**
 * Check if a vendor has a custom parser
 * @param {string} vendorName - Name of the vendor
 * @returns {boolean}
 */
function hasCustomParser(vendorName) {
  return vendorName in PARSER_REGISTRY
}

/**
 * Get list of all supported vendors
 * @returns {string[]} Array of vendor names
 */
function getSupportedVendors() {
  return Object.keys(PARSER_REGISTRY)
}

module.exports = {
  getParser,
  hasCustomParser,
  getSupportedVendors,
}
