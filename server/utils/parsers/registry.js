const CleanIndustryParser = require('./cleanIndustryParser')
const SankoSangyoParser = require('./sankoSangyoParser')
const HokukeiParser = require('./hokukeiParser')
const NanseiParser = require('./nanseiParser')
const TaimanParser = require('./taimanParser')
const TakabishiParser = require('./takabishiParser')
const OmegaJapanParser = require('./omegaJapanParser')
const NakazawaKenhanParser = require('./nakazawaKenhanParser')
const TokiwaSystemParser = require('./tokiwaSystemParser')

// Parser registry - maps vendor names to their parser classes
const PARSER_REGISTRY = {
  クリーン産業: CleanIndustryParser,
  三高産業: SankoSangyoParser,
  北恵株式会社: HokukeiParser,
  ナンセイ: NanseiParser,
  大萬: TaimanParser,
  髙菱管理: TakabishiParser,
  高菱管理: TakabishiParser,
  オメガジャパン: OmegaJapanParser,
  ナカザワ建販: NakazawaKenhanParser,
  トキワシステム: TokiwaSystemParser,
}

module.exports = { PARSER_REGISTRY }
