/** Excel's compatibility names: the functions of the other files, under the old names and argument lists. */
import { functions as descriptive } from './descriptive.js'
import type { FnDef } from '../../args.js'
import { E, type V } from '../../values.js'
import { badDf, betaDist, functions as distributions, hypgeomDist, lognormDist, negbinomDist, tCdf } from './distributions.js'
import { functions as inverse } from './inverse.js'
import { functions as regression } from './regression.js'
import { functions as tests } from './tests.js'
import { normCdf } from './special.js'

function tdist(x: number, df: number, tails: number): V {
  if (x < 0 || badDf(df) || (tails !== 1 && tails !== 2) || !Number.isFinite(x * x)) return E.NUM
  return tails * tCdf(x, df)[1]
}

export const functions: Record<string, FnDef> = {
  BETADIST: { args: 'n n n n? n?', fn: (x: number, a: number, b: number, lo?: number, hi?: number) => betaDist(x, a, b, true, lo, hi) },
  BETAINV: inverse['BETA.INV'],
  BINOMDIST: distributions['BINOM.DIST'],
  CHIDIST: distributions['CHISQ.DIST.RT'],
  CHIINV: inverse['CHISQ.INV.RT'],
  CHITEST: tests['CHISQ.TEST'],
  CONFIDENCE: inverse['CONFIDENCE.NORM'],
  COVAR: regression['COVARIANCE.P'],
  CRITBINOM: inverse['BINOM.INV'],
  EXPONDIST: distributions['EXPON.DIST'],
  FDIST: distributions['F.DIST.RT'],
  FINV: inverse['F.INV.RT'],
  FTEST: tests['F.TEST'],
  GAMMADIST: distributions['GAMMA.DIST'],
  GAMMAINV: inverse['GAMMA.INV'],
  HYPGEOMDIST: { args: 'i i i i', fn: (k: number, n: number, K: number, N: number) => hypgeomDist(k, n, K, N, false) },
  LOGINV: inverse['LOGNORM.INV'],
  LOGNORMDIST: { args: 'n n n', fn: (x: number, m: number, sd: number) => lognormDist(x, m, sd, true) },
  MODE: descriptive['MODE.SNGL'],
  NEGBINOMDIST: { args: 'i i n', fn: (f: number, s: number, p: number) => negbinomDist(f, s, p, false) },
  NORMDIST: distributions['NORM.DIST'],
  NORMINV: inverse['NORM.INV'],
  NORMSDIST: { args: 'n', fn: (z: number) => normCdf(z)[0] },
  NORMSINV: inverse['NORM.S.INV'],
  PERCENTILE: descriptive['PERCENTILE.INC'],
  PERCENTRANK: descriptive['PERCENTRANK.INC'],
  POISSON: distributions['POISSON.DIST'],
  QUARTILE: descriptive['QUARTILE.INC'],
  RANK: descriptive['RANK.EQ'],
  STDEV: descriptive['STDEV.S'],
  STDEVP: descriptive['STDEV.P'],
  TDIST: { args: 'n i i', fn: tdist },
  TINV: inverse['T.INV.2T'],
  TTEST: tests['T.TEST'],
  VAR: descriptive['VAR.S'],
  VARP: descriptive['VAR.P'],
  WEIBULL: distributions['WEIBULL.DIST'],
  ZTEST: tests['Z.TEST'],
}
