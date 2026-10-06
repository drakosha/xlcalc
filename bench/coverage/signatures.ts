/**
 * The signatures of the Excel functions formulajs 4.6.1 has: argument kinds and typical calls (from Excel's
 * function reference, mostly its own examples). The measurement's cases are generated from it.
 *
 * Argument kinds (one character; `?` after it means optional):
 * - `n` a number, `i` a whole number (Excel truncates), `p` a probability or fraction, `d` a date (serial
 *   number), `t` a time (fraction of a day), `s` text, `b` a boolean, `a` any value;
 * - `N` numbers to fold (a range: numbers only, a direct argument converted; the SUM rule);
 * - `V` values as they are (COUNTA, AVERAGEA, AND, CONCAT): a range with all its contents;
 * - `M` a numeric array with its shape (paired and matrix functions), an error in it is the answer's error;
 * - `R` a range as it is, with its shape (lookups, databases, criteria);
 * - `D` a range of dates (holidays, payment dates);
 * - `c` a criterion.
 * `rep` says how many of the last kinds repeat (SUM: 1, SUMIFS: 2).
 *
 * Values in calls: a number, a string, a boolean; a string starting with `@` refers to a place of the data
 * sheet (`data.ts`). Dates are written as serial numbers (`dt`).
 */
import { dt } from './data.js'

export type Arg = number | string | boolean

/** What to take from an array answer: an element (row, column from 1) or a size. */
export type Pick = [number, number] | 'rows' | 'cols'

export interface Sig {
  name: string
  kinds: string
  rep?: number
  calls: Arg[][]
  /** The answer is an array: the call is wrapped in INDEX/ROWS/COLUMNS for each element of the list. */
  picks?: Pick[]
  /** Hand-written edge calls: the tag, the arguments and (for an array answer) what to take. */
  extra?: [string, Arg[], Pick?][]
  /** No automatic edge cases (criteria, databases: their edge cases are written by hand). */
  noAuto?: boolean
  /**
   * Size and count arguments (arrays, repeats, orders, periods): "large" is CAP instead of 1E10, and a date as
   * a size and "1.5" (the date 46143 in ru-RU) are not put in. Otherwise Excel risks running out of resources.
   */
  cap?: boolean
}

const S: Sig[] = []
const s = (name: string, kinds: string, calls: Arg[][], o: Partial<Sig> = {}) => S.push({ name, kinds, calls, ...o })
const each = (names: string[], kinds: string, calls: Arg[][], o: Partial<Sig> = {}) => names.forEach((n) => s(n, kinds, calls, o))

/** CHISQ.TEST and CHITEST: edge cases. */
const CHI: [string, Arg[]][] = [
  ['range: mixed', ['@actualmix', '@expected']],
  ['range with an error', ['@actualerr', '@expected']],
  ['different size', ['@actual', '@m1']],
]

// ---------- math and trigonometry ----------
s('ABS', 'n', [[-4], [2.5]])
s('ACOS', 'n', [[0.5], [-1]])
s('ACOSH', 'n', [[1], [10]])
s('ACOT', 'n', [[2]])
s('ACOTH', 'n', [[6]])
s('AGGREGATE', 'i i R n?', [[9, 4, '@nums1'], [14, 6, '@witherr', 2], [1, 6, '@witherr']])
s('ARABIC', 's', [['LVII'], ['mcmxii']])
s('ASIN', 'n', [[-0.5]])
s('ASINH', 'n', [[-2.5], [10]])
s('ATAN', 'n', [[1]])
s('ATAN2', 'n n', [[1, 1], [-1, -1]])
s('ATANH', 'n', [[0.76159416], [-0.1]])
s('BASE', 'i i i?', [[7, 2], [100, 16], [15, 2, 10]])
s('CEILING', 'n n', [[2.5, 1], [-2.5, -2], [-2.5, 2], [1.5, 0.1], [0.234, 0.01]])
s('CEILING.MATH', 'n n? i?', [[24.3, 5], [6.7], [-8.1, 2], [-5.5, 2, -1]])
s('CEILING.PRECISE', 'n n?', [[4.3], [-4.3, 2], [4.3, -2]])
s('ISO.CEILING', 'n n?', [[4.3], [-4.3, 2], [4.3, -2]])
s('COMBIN', 'i i', [[8, 2]])
s('COMBINA', 'i i', [[4, 3], [10, 3]])
s('COS', 'n', [[1.047], [0]])
s('COSH', 'n', [[4]])
s('COT', 'n', [[30]])
s('COTH', 'n', [[2]])
s('CSC', 'n', [[15]])
s('CSCH', 'n', [[1.5]])
s('DECIMAL', 's i', [['FF', 16], ['zap', 36], ['111', 2]])
s('DEGREES', 'n', [[3.14159265358979]])
s('EVEN', 'n', [[1.5], [3], [-1]])
s('EXP', 'n', [[1], [2]])
s('FACT', 'n', [[5], [1.9]])
s('FACTDOUBLE', 'n', [[6], [7]])
s('FLOOR', 'n n', [[3.7, 2], [-2.5, -2], [2.5, -2], [1.58, 0.1], [0.234, 0.01]])
s('FLOOR.MATH', 'n n? i?', [[24.3, 5], [6.7], [-8.1, 2], [-5.5, 2, -1]])
s('FLOOR.PRECISE', 'n n?', [[-3.2, -1], [3.2, 1], [-3.2, 1], [3.2]])
s('GCD', 'N', [[5, 2], [24, 36], [7, 1]], { rep: 1 })
s('INT', 'n', [[8.9], [-8.9]])
s('LCM', 'N', [[5, 2], [24, 36]], { rep: 1 })
s('LN', 'n', [[86], [2.7182818]])
s('LOG', 'n n?', [[10], [8, 2], [86, 2.7182818]])
s('LOG10', 'n', [[86], [1e5]])
s('MMULT', 'M M', [['@m1', '@m2'], ['@m23', '@m1']], { picks: [[1, 1], [2, 2], 'rows', 'cols'] })
s('MOD', 'n n', [[3, 2], [-3, 2], [3, -2], [-3, -2]])
s('MROUND', 'n n', [[10, 3], [-10, -3], [1.3, 0.2], [5, -2]])
s('MULTINOMIAL', 'N', [[2, 3, 4]], { rep: 1 })
s('MUNIT', 'i', [[3]], { picks: [[1, 1], [2, 3], 'rows'] })
s('ODD', 'n', [[1.5], [3], [2], [-1], [-2]])
s('PI', '', [[]])
s('POWER', 'n n', [[5, 2], [98.6, 3.2], [4, 1.25]])
s('PRODUCT', 'N', [['@nums1'], [5, 15, 30]], { rep: 1 })
s('QUOTIENT', 'n n', [[5, 2], [4.5, 3.1], [-10, 3]])
s('RADIANS', 'n', [[270]])
s('ROMAN', 'i i?', [[499, 0], [499, 1], [499, 4], [2013]])
s('ROUND', 'n i', [[2.15, 1], [2.149, 1], [-1.475, 2], [21.5, -1], [626.3, -3], [1.005, 2], [-50.55, -2]])
s('ROUNDDOWN', 'n i', [[3.2, 0], [76.9, 0], [3.14159, 3], [-3.14159, 1], [31415.92654, -2]])
s('ROUNDUP', 'n i', [[3.2, 0], [76.9, 0], [3.14159, 3], [-3.14159, 1], [31415.92654, -2]])
s('SEC', 'n', [[45]])
s('SECH', 'n', [[45]])
s('SERIESSUM', 'n n n M', [[0.785398163, 0, 2, '@coef']])
s('SIGN', 'n', [[10], [0], [-0.00001]])
s('SIN', 'n', [[3.14159265358979], [0.5235987755982988]])
s('SINH', 'n', [[1], [-0.5]])
s('SQRT', 'n', [[16], [-16]])
s('SQRTPI', 'n', [[1], [2]])
s('SUBTOTAL', 'i R', [[9, '@nums1'], [1, '@nums1'], [109, '@nums1'], [2, '@mixed']], { rep: 1 })
s('SUM', 'N', [['@nums1'], [3, 2], ['@nums1', '@nums2', 5]], { rep: 1 })
s('SUMPRODUCT', 'M', [['@nums1', '@nums2'], ['@m1', '@m2']], { rep: 1 })
s('SUMSQ', 'N', [[3, 4], ['@nums1']], { rep: 1 })
each(['SUMX2MY2', 'SUMX2PY2', 'SUMXMY2'], 'M M', [['@nums1', '@nums2'], ['@m1', '@m2']])
s('TAN', 'n', [[0.785], [0.7853981633974483]])
s('TANH', 'n', [[-2], [0.5]])
s('TRUNC', 'n i?', [[8.9], [-8.9], [0.45], [3.14159, 2]])

// ---------- statistical ----------
s('AVEDEV', 'N', [['@nums1'], [4, 5, 6, 7, 5, 4, 3]], { rep: 1 })
s('AVERAGE', 'N', [['@nums1'], [10, 7, 9, 27, 2]], { rep: 1 })
s('AVERAGEA', 'V', [['@nums1'], [10, 7, 9, 2]], { rep: 1 })
s('BETA.DIST', 'n n n b n? n?', [[2, 8, 10, true, 1, 3], [2, 8, 10, false, 1, 3], [0.4, 2, 3, true]])
s('BETA.INV', 'p n n n? n?', [[0.685470581, 8, 10, 1, 3], [0.5, 2, 3]])
s('BINOM.DIST', 'i i p b', [[6, 10, 0.5, false], [6, 10, 0.5, true]])
s('BINOM.DIST.RANGE', 'i p i i?', [[60, 0.75, 48], [60, 0.75, 45, 50]])
s('BINOM.INV', 'i p p', [[6, 0.5, 0.75]])
s('CHISQ.DIST', 'n n b', [[0.5, 1, true], [2, 3, false]])
s('CHISQ.DIST.RT', 'n n', [[18.307, 10]])
s('CHISQ.INV', 'p n', [[0.93, 1], [0.6, 2]])
s('CHISQ.INV.RT', 'p n', [[0.050001, 10]])
s('CHISQ.TEST', 'M M', [['@actual', '@expected']], { extra: CHI })
s('CONFIDENCE.NORM', 'p n i', [[0.05, 2.5, 50]])
s('CONFIDENCE.T', 'p n i', [[0.05, 1, 50]])
s('CORREL', 'M M', [['@nums1', '@nums2']])
s('COUNT', 'V', [['@nums1'], [1, 2, 3]], { rep: 1 })
s('COUNTA', 'V', [['@nums1'], [1, 'x', true]], { rep: 1 })
s('COUNTBLANK', 'R', [['@mixed'], ['@blank']], {
  extra: [['range with an error', ['@witherr']], ['one blank cell', ['@empty']], ['empty string from a formula', ['@emptystr']], ['texts', ['@txt']], ['one cell with a date', ['@date']]],
})
each(['COVARIANCE.P', 'COVARIANCE.S'], 'M M', [['@nums1', '@nums2']])
s('DEVSQ', 'N', [['@nums1'], [4, 5, 8, 7, 11, 4, 3]], { rep: 1 })
s('EXPON.DIST', 'n n b', [[0.2, 10, true], [0.2, 10, false]])
s('F.DIST', 'n n n b', [[15.2069, 6, 4, true], [15.2069, 6, 4, false]])
s('F.DIST.RT', 'n n n', [[15.2069, 6, 4]])
s('F.INV', 'p n n', [[0.01, 6, 4]])
s('F.INV.RT', 'p n n', [[0.01, 6, 4]])
s('F.TEST', 'M M', [['@nums1', '@nums2']])
s('FISHER', 'n', [[0.75]])
s('FISHERINV', 'n', [[0.972955]])
s('FORECAST', 'n M M', [[30, '@nums2', '@nums1']])
s('FREQUENCY', 'M M', [['@nums1', '@bins']], { picks: [[1, 1], [2, 1], [4, 1], 'rows'] })
s('GAMMA', 'n', [[2.5], [-3.75], [0], [-2]])
s('GAMMA.DIST', 'n n n b', [[10.00001131, 9, 2, false], [10.00001131, 9, 2, true]])
s('GAMMA.INV', 'p n n', [[0.068094, 9, 2]])
s('GAMMALN', 'n', [[4]])
s('GAMMALN.PRECISE', 'n', [[4]])
s('GAUSS', 'n', [[2]])
s('GEOMEAN', 'N', [['@pos'], [4, 5, 8, 7, 11, 4, 3]], { rep: 1 })
s('GROWTH', 'M M? M? b?', [['@gy', '@gx', '@gnew'], ['@gy', '@gx']], {
  picks: [[1, 1], [2, 1], 'rows'],
  extra: [['without known_x', ['@gy']], ['without the constant', ['@gy', '@gx', '@gnew', false]], ['different size', ['@gy', '@lx']], ['zero in known_y', ['@probx', '@lx']]],
})
s('HARMEAN', 'N', [['@pos'], [4, 5, 8, 7, 11, 4, 3]], { rep: 1 })
s('HYPGEOM.DIST', 'i i i i b', [[1, 4, 8, 20, true], [1, 4, 8, 20, false]])
s('INTERCEPT', 'M M', [['@nums2', '@nums1']])
s('KURT', 'N', [['@nums1'], [3, 4, 5, 2, 3, 4, 5, 6, 4, 7]], { rep: 1 })
s('LARGE', 'N i', [['@nums1', 3], ['@nums1', 1]])
s('LINEST', 'M M? b? b?', [['@ly', '@lx'], ['@ly', '@lx', true, true]], { picks: [[1, 1], [1, 2], [3, 1], 'rows', 'cols'] })
s('LOGEST', 'M M? b? b?', [['@gy', '@gx'], ['@gy', '@gx', true, true]], { picks: [[1, 1], [1, 2], [3, 1], 'rows', 'cols'] })
s('LOGNORM.DIST', 'n n n b', [[4, 3.5, 1.2, true], [4, 3.5, 1.2, false]])
s('LOGNORM.INV', 'p n n', [[0.039084, 3.5, 1.2]])
s('MAX', 'N', [['@nums1'], [10, 7, 9, 27, 2]], { rep: 1 })
s('MAXA', 'V', [['@nums1'], [10, 7, 9, 27, 2]], { rep: 1 })
s('MEDIAN', 'N', [['@nums1'], [1, 2, 3, 4, 5]], { rep: 1 })
s('MIN', 'N', [['@nums1'], [10, 7, 9, 27, 2]], { rep: 1 })
s('MINA', 'V', [['@nums1'], [10, 7, 9, 27, 2]], { rep: 1 })
s('MODE.MULT', 'N', [['@modes']], { rep: 1, picks: [[1, 1], [2, 1], [3, 1], 'rows'] })
s('MODE.SNGL', 'N', [['@dup'], [5.6, 4, 4, 3, 2, 4]], { rep: 1 })
s('NEGBINOM.DIST', 'i i p b', [[10, 5, 0.25, true], [10, 5, 0.25, false]])
s('NORM.DIST', 'n n n b', [[42, 40, 1.5, true], [42, 40, 1.5, false]])
s('NORM.INV', 'p n n', [[0.908789, 40, 1.5]])
s('NORM.S.DIST', 'n b', [[1.333333, true], [1.333333, false]])
s('NORM.S.INV', 'p', [[0.908789]])
s('PEARSON', 'M M', [['@nums1', '@nums2']])
s('PERCENTILE.EXC', 'N p', [['@nums1', 0.25], ['@nums1', 0.5], ['@nums1', 0.05]])
s('PERCENTILE.INC', 'N p', [['@nums1', 0.3], ['@nums1', 1]])
s('PERCENTRANK.EXC', 'N n i?', [['@nums1', 8], ['@nums1', 8, 2], ['@nums1', 5.5]])
s('PERCENTRANK.INC', 'N n i?', [['@nums1', 8], ['@nums1', 5.5, 2], ['@nums1', 100]])
s('PERMUT', 'i i', [[100, 3], [3, 2]])
s('PERMUTATIONA', 'i i', [[3, 2], [2, 2]])
s('PHI', 'n', [[0.75]])
s('POISSON.DIST', 'i n b', [[2, 5, true], [2, 5, false]])
s('PROB', 'M M n n?', [['@probx', '@probp', 2], ['@probx', '@probp', 1, 3]])
s('QUARTILE.EXC', 'N i', [['@nums1', 1], ['@nums1', 3]])
s('QUARTILE.INC', 'N i', [['@nums1', 1], ['@nums1', 0], ['@nums1', 4]])
s('RANK.AVG', 'n R b?', [[94, '@ties'], [2, '@dup'], [2, '@dup', 1]])
s('RANK.EQ', 'n R b?', [[94, '@ties'], [2, '@dup'], [2, '@dup', 1]])
s('RSQ', 'M M', [['@nums1', '@nums2']])
s('SKEW', 'N', [['@nums1'], [3, 4, 5, 2, 3, 4, 5, 6, 4, 7]], { rep: 1 })
s('SKEW.P', 'N', [['@nums1'], [3, 4, 5, 2, 3, 4, 5, 6, 4, 7]], { rep: 1 })
s('SLOPE', 'M M', [['@nums2', '@nums1']])
s('SMALL', 'N i', [['@nums1', 4], ['@nums1', 1]])
s('STANDARDIZE', 'n n n', [[42, 40, 1.5]])
each(['STDEV.P', 'STDEV.S', 'VAR.P', 'VAR.S'], 'N', [['@nums1'], [1345, 1301, 1368, 1322, 1310]], { rep: 1 })
each(['STDEVA', 'STDEVPA', 'VARA', 'VARPA'], 'V', [['@nums1'], [1345, 1301, 1368, 1322, 1310]], { rep: 1 })
s('STEYX', 'M M', [['@nums2', '@nums1']])
s('T.DIST', 'n n b', [[60, 1, true], [8, 3, false]])
s('T.DIST.2T', 'n n', [[1.959999998, 60]])
s('T.DIST.RT', 'n n', [[1.959999998, 60]])
s('T.INV', 'p n', [[0.75, 2]])
s('T.INV.2T', 'p n', [[0.546449, 60]])
s('T.TEST', 'M M i i', [['@nums1', '@nums2', 2, 1], ['@nums1', '@nums2', 1, 2], ['@nums1', '@nums2', 2, 3]])
s('TREND', 'M M? M? b?', [['@ly', '@lx', '@lnew'], ['@ly', '@lx']], {
  picks: [[1, 1], [2, 1], 'rows'],
  extra: [['without known_x', ['@ly']], ['without the constant', ['@ly', '@lx', '@lnew', false]], ['different size', ['@ly', '@gx']], ['text in known_y', ['@cfmix', '@lx']]],
})
s('TRIMMEAN', 'N p', [['@nums1', 0.2], ['@nums1', 0.25]])
s('WEIBULL.DIST', 'n n n b', [[105, 20, 100, true], [105, 20, 100, false]])
s('Z.TEST', 'M n n?', [['@nums1', 4], ['@nums1', 6, 2]])

// ---------- compatibility (old names) ----------
s('BETADIST', 'n n n n? n?', [[2, 8, 10, 1, 3]])
s('BETAINV', 'p n n n? n?', [[0.685470581, 8, 10, 1, 3]])
s('BINOMDIST', 'i i p b', [[6, 10, 0.5, false], [6, 10, 0.5, true]])
s('CHIDIST', 'n n', [[18.307, 10]])
s('CHIINV', 'p n', [[0.050001, 10]])
s('CHITEST', 'M M', [['@actual', '@expected']], { extra: CHI })
s('COVAR', 'M M', [['@nums1', '@nums2']])
s('CRITBINOM', 'i p p', [[6, 0.5, 0.75]])
s('EXPONDIST', 'n n b', [[0.2, 10, true], [0.2, 10, false]])
s('FDIST', 'n n n', [[15.2069, 6, 4]])
s('FINV', 'p n n', [[0.01, 6, 4]])
s('FTEST', 'M M', [['@nums1', '@nums2']])
s('GAMMADIST', 'n n n b', [[10.00001131, 9, 2, false], [10.00001131, 9, 2, true]])
s('GAMMAINV', 'p n n', [[0.068094, 9, 2]])
s('HYPGEOMDIST', 'i i i i', [[1, 4, 8, 20]])
s('LOGINV', 'p n n', [[0.039084, 3.5, 1.2]])
s('LOGNORMDIST', 'n n n', [[4, 3.5, 1.2]])
s('NEGBINOMDIST', 'i i p', [[10, 5, 0.25]])
s('NORMDIST', 'n n n b', [[42, 40, 1.5, true], [42, 40, 1.5, false]])
s('NORMINV', 'p n n', [[0.908789, 40, 1.5]])
s('NORMSDIST', 'n', [[1.333333]])
s('NORMSINV', 'p', [[0.908789]])
s('STDEVP', 'N', [['@nums1'], [1345, 1301, 1368, 1322, 1310]], { rep: 1 })
s('VARP', 'N', [['@nums1'], [1345, 1301, 1368, 1322, 1310]], { rep: 1 })
s('TDIST', 'n n i', [[1.959999998, 60, 2], [1.959999998, 60, 1]])
s('TINV', 'p n', [[0.05464493, 60]])
s('TTEST', 'M M i i', [['@nums1', '@nums2', 2, 1], ['@nums1', '@nums2', 1, 2]])
s('ZTEST', 'M n n?', [['@nums1', 4], ['@nums1', 6, 2]])

// ---------- date and time ----------
s('DATE', 'i i i', [[2025, 1, 31], [2008, 14, 2], [2008, 1, -15], [108, 1, 2], [1900, 2, 29]])
s('DATEDIF', 'd d s', [
  [dt('2001-01-01'), dt('2003-01-01'), 'Y'],
  [dt('2001-06-01'), dt('2002-08-15'), 'D'],
  [dt('2001-06-01'), dt('2002-08-15'), 'YD'],
  [dt('2001-01-31'), dt('2001-03-01'), 'MD'],
  [dt('2001-06-01'), dt('2002-08-15'), 'YM'],
  [dt('2024-01-31'), dt('2024-02-29'), 'M'],
])
s('DATEVALUE', 's', [['2025-01-31'], ['8/22/2011'], ['22-MAY-2011'], ['2011/02/23']])
s('DAY', 'd', [[dt('2011-04-15')], [dt('2024-02-29')]])
s('DAYS', 'd d', [[dt('2021-03-15'), dt('2021-02-01')], [dt('2021-12-31'), dt('2021-01-01')]])
s('DAYS360', 'd d b?', [[dt('2011-01-30'), dt('2011-12-31')], [dt('2011-01-30'), dt('2011-12-31'), true], [dt('2024-02-29'), dt('2025-02-28')]])
s('EDATE', 'd i', [[dt('2011-01-15'), 1], [dt('2011-01-15'), -1], [dt('2025-01-31'), 1], [dt('2024-02-29'), 12]])
s('EOMONTH', 'd i', [[dt('2011-01-01'), 1], [dt('2011-01-01'), -3], [dt('2024-01-31'), 1]])
s('HOUR', 't', [[0.75], [45688.78], [0.7843]])
s('ISOWEEKNUM', 'd', [[dt('2012-03-09')], [dt('2021-01-01')], [dt('2024-12-30')]])
s('MINUTE', 't', [[0.5104166667], [45688.78]])
s('MONTH', 'd', [[dt('2011-04-15')], [dt('2025-12-31')]])
s('NETWORKDAYS', 'd d D?', [[dt('2012-10-01'), dt('2013-03-01')], [dt('2012-10-01'), dt('2013-03-01'), '@hol'], [dt('2025-01-31'), dt('2025-01-01')]])
s('NETWORKDAYS.INTL', 'd d a? D?', [[dt('2006-01-01'), dt('2006-01-31')], [dt('2006-01-01'), dt('2006-02-01'), 7, '@hol'], [dt('2006-01-01'), dt('2006-02-01'), '0010001']])
s('SECOND', 't', [[0.70020833], [45688.7802]])
s('TIME', 'i i i', [[12, 0, 0], [16, 48, 10], [27, 0, 0], [0, 750, 0]])
s('TIMEVALUE', 's', [['2:24 AM'], ['22-Aug-2011 6:35 AM'], ['18:30']])
s('WEEKDAY', 'd i?', [[dt('2008-02-14')], [dt('2008-02-14'), 2], [dt('2008-02-14'), 3], [dt('2008-02-14'), 11], [dt('2008-02-14'), 17]])
s('WEEKNUM', 'd i?', [[dt('2012-03-09')], [dt('2012-03-09'), 2], [dt('2025-01-01'), 21]])
s('WORKDAY', 'd i D?', [[dt('2008-10-01'), 151], [dt('2008-10-01'), 151, '@hol'], [dt('2025-01-31'), -5]])
s('WORKDAY.INTL', 'd i a? D?', [[dt('2012-01-01'), 90, 11], [dt('2012-01-01'), 30, 17], [dt('2012-01-01'), 30, '0000011', '@hol']])
s('YEAR', 'd', [[dt('2008-07-05')], [dt('2010-07-05')]])
s('YEARFRAC', 'd d i?', [[dt('2012-01-01'), dt('2012-07-30')], [dt('2012-01-01'), dt('2012-07-30'), 1], [dt('2012-01-01'), dt('2012-07-30'), 3], [dt('2024-02-29'), dt('2025-03-01'), 1], [dt('2024-02-29'), dt('2025-03-01'), 4]])

// ---------- engineering ----------
s('BESSELI', 'n i', [[1.5, 1]])
s('BESSELJ', 'n i', [[1.9, 2]])
s('BESSELK', 'n i', [[1.5, 1]])
s('BESSELY', 'n i', [[2.5, 1]])
s('BIN2DEC', 's', [['1100100'], ['1111111111'], [1100100]])
s('BIN2HEX', 's i?', [['11111011', 4], ['1110'], ['1111111111']])
s('BIN2OCT', 's i?', [['1001', 3], ['1100100'], ['1111111111']])
s('BITAND', 'i i', [[1, 5], [13, 25]])
s('BITLSHIFT', 'i i', [[4, 2], [4, -1]])
s('BITOR', 'i i', [[23, 10]])
s('BITRSHIFT', 'i i', [[13, 2]])
s('BITXOR', 'i i', [[5, 3]])
s('COMPLEX', 'n n s?', [[3, 4], [3, 4, 'j'], [0, 1], [1, 0]])
s('CONVERT', 'n s s', [[1, 'lbm', 'kg'], [68, 'F', 'C'], [2.5, 'ft', 'sec'], [100, 'km', 'mi'], [1, 'day', 'hr']])
s('DEC2BIN', 'i i?', [[9, 4], [-100]])
s('DEC2HEX', 'i i?', [[100, 4], [-54], [28]])
s('DEC2OCT', 'i i?', [[58, 3], [-100]])
s('DELTA', 'n n?', [[5, 4], [5, 5], [0.5]])
s('ERF', 'n n?', [[0.745], [0, 1]])
s('ERFC', 'n', [[1]])
s('GESTEP', 'n n?', [[5, 4], [5, 5], [-4, -5], [-1]])
s('HEX2BIN', 's i?', [['F', 8], ['B7'], ['FFFFFFFFFF']])
s('HEX2DEC', 's', [['A5'], ['FFFFFFFF5B'], ['3DA408B9']])
s('HEX2OCT', 's i?', [['F', 3], ['3B4E'], ['FFFFFFFF00']])
s('IMABS', 's', [['5+12i']])
s('IMAGINARY', 's', [['3+4i'], ['0-j'], ['4']])
s('IMARGUMENT', 's', [['3+4i']])
s('IMCONJUGATE', 's', [['3+4i']])
s('IMCOS', 's', [['1+i']])
each(['IMCOSH', 'IMCOT', 'IMCSC', 'IMCSCH', 'IMSEC', 'IMSECH', 'IMSIN', 'IMSINH', 'IMTAN'], 's', [['4+3i']])
s('IMDIV', 's s', [['-238+240i', '10+24i']])
s('IMEXP', 's', [['1+i']])
each(['IMLN', 'IMLOG10', 'IMLOG2'], 's', [['3+4i']])
s('IMPOWER', 's n', [['2+3i', 3]])
s('IMPRODUCT', 's', [['3+4i', '5-3i'], ['1+2i', 30]], { rep: 1 })
s('IMREAL', 's', [['6-9i']])
s('IMSQRT', 's', [['1+i']])
s('IMSUB', 's s', [['13+4i', '5+3i']])
s('IMSUM', 's', [['3+4i', '5-3i']], { rep: 1 })
s('OCT2BIN', 's i?', [['3', 3], ['7777777000']])
s('OCT2DEC', 's', [['54'], ['7777777533']])
s('OCT2HEX', 's i?', [['100', 4], ['7777777533']])

// ---------- financial ----------
s('ACCRINT', 'd d d n n i i? b?', [
  [dt('2008-03-01'), dt('2008-08-31'), dt('2008-05-01'), 0.1, 1000, 2, 0],
  [dt('2008-03-01'), dt('2008-08-31'), dt('2008-05-01'), 0.1, 1000, 2, 0, false],
  [dt('2008-03-05'), dt('2008-08-31'), dt('2008-05-01'), 0.1, 1000, 2, 0, true],
])
s('COUPDAYS', 'd d i i?', [[dt('2011-01-25'), dt('2011-11-15'), 2, 1], [dt('2011-01-25'), dt('2011-11-15'), 4]])
s('CUMIPMT', 'n i n i i i', [[0.0075, 360, 125000, 13, 24, 0], [0.0075, 360, 125000, 1, 1, 0]])
s('CUMPRINC', 'n i n i i i', [[0.0075, 360, 125000, 13, 24, 0], [0.0075, 360, 125000, 1, 1, 0]])
s('DB', 'n n i i i?', [[1000000, 100000, 6, 1, 7], [1000000, 100000, 6, 2, 7], [1000000, 100000, 6, 7, 7]])
s('DDB', 'n n n n n?', [[2400, 300, 3650, 1], [2400, 300, 120, 1, 2], [2400, 300, 10, 1, 2], [2400, 300, 10, 10]])
s('DISC', 'd d n n i?', [[dt('2018-07-01'), dt('2048-01-01'), 97.975, 100, 1], [dt('2018-07-01'), dt('2048-01-01'), 97.975, 100]])
s('DOLLARDE', 'n i', [[1.02, 16], [1.1, 32]])
s('DOLLARFR', 'n i', [[1.125, 16], [1.125, 32]])
s('EFFECT', 'n i', [[0.0525, 4]])
s('FV', 'n n n n? i?', [[0.005, 10, -200, -500, 1], [0.01, 12, -1000], [0.11 / 12, 35, -2000, 0, 1]])
s('FVSCHEDULE', 'n M', [[1, '@sched']])
s('IPMT', 'n i n n n? i?', [[0.1 / 12, 1, 36, 8000], [0.1, 3, 3, 8000]])
s('IRR', 'M n?', [['@cf'], ['@cf', 0.1]], {
  extra: [['range: mixed', ['@cfmix']], ['range with an error', ['@cferr']], ['no sign change', ['@xcfpos']], ['guess -0.5', ['@cf', -0.5]], ['blank range', ['@blank']]],
})
s('ISPMT', 'n i i n', [[0.1 / 12, 1, 36, 8000000], [0.1, 1, 3, 8000000]])
s('MIRR', 'M n n', [['@cf', 0.1, 0.12]], { extra: [['range: mixed', ['@cfmix', 0.1, 0.12]], ['range with an error', ['@cferr', 0.1, 0.12]], ['no sign change', ['@xcfpos', 0.1, 0.12]]] })
s('NOMINAL', 'n i', [[0.053543, 4]])
s('NPER', 'n n n n? i?', [[0.01, -100, -1000, 10000, 1], [0.01, -100, -1000, 10000], [0.01, -100, -1000]])
s('NPV', 'n N', [[0.1, -10000, 3000, 4200, 6800], [0.08, '@cf']], { rep: 1 })
s('PDURATION', 'n n n', [[0.025, 2000, 2200], [0.025 / 12, 1000, 1200]])
s('PMT', 'n n n n? i?', [[0.08 / 12, 10, 10000], [0.08 / 12, 10, 10000, 0, 1], [0.005, 216, 0, 50000]])
s('PPMT', 'n i n n n? i?', [[0.1 / 12, 1, 24, 2000], [0.08, 10, 10, 200000]])
s('PRICEDISC', 'd d n n i?', [[dt('2008-02-16'), dt('2008-03-01'), 0.0525, 100, 2]])
s('PV', 'n n n n? i?', [[0.08 / 12, 240, 500, 0, 0], [0.1, 5, -100]])
s('RATE', 'n n n n? i? n?', [[48, -200, 8000], [48, -200, 8000, 0, 0, 0.1]])
s('RRI', 'n n n', [[96, 10000, 11000]])
s('SLN', 'n n n', [[30000, 7500, 10]])
s('SYD', 'n n n n', [[30000, 7500, 10, 1], [30000, 7500, 10, 10]])
s('TBILLEQ', 'd d n', [[dt('2008-03-31'), dt('2008-06-01'), 0.0914]])
s('TBILLPRICE', 'd d n', [[dt('2008-03-31'), dt('2008-06-01'), 0.09]])
s('TBILLYIELD', 'd d n', [[dt('2008-03-31'), dt('2008-06-01'), 98.45]])
s('XIRR', 'M D n?', [['@xcf', '@xdates'], ['@xcf', '@xdates', 0.1]], {
  extra: [['no sign change', ['@xcfpos', '@xdates']], ['dates out of order', ['@xcf', '@xdatesmix']], ['different size', ['@xcf', '@hol']], ['guess -0.5', ['@xcf', '@xdates', -0.5]], ['cash flow with an error', ['@cferr', '@xdates']], ['cash flow: mixed', ['@cfmix', '@xdates']]],
})
s('XNPV', 'n M D', [[0.09, '@xcf', '@xdates']], { extra: [['dates out of order', [0.09, '@xcf', '@xdatesmix']], ['different size', [0.09, '@xcf', '@hol']], ['cash flow: mixed', [0.09, '@cfmix', '@xdates']]] })

// ---------- information ----------
s('ERROR.TYPE', 'a', [['@div0'], ['@na']])
each(['ISBLANK', 'ISERR', 'ISERROR', 'ISLOGICAL', 'ISNA', 'ISNONTEXT', 'ISNUMBER', 'ISTEXT'], 'a', [[2.5], ['abc']])
s('ISEVEN', 'n', [[-1], [2.5], [5]])
s('ISODD', 'n', [[-1], [2.5], [5]])
s('N', 'a', [[7], ['Even'], [true], ['@date']])
s('NA', '', [[]])
s('TYPE', 'a', [[2], ['abc'], [true]], { extra: [['range', ['@nums1']]] })

// ---------- logical ----------
s('AND', 'V', [[true, true], [true, false], [2, 1]], { rep: 1, extra: [['range of booleans', ['@alltrue']], ['range with text and blanks', ['@logic']]] })
s('OR', 'V', [[false, true], [false, false]], { rep: 1, extra: [['range with text and blanks', ['@logic']]] })
s('XOR', 'V', [[true, false], [true, true]], { rep: 1, extra: [['range with text and blanks', ['@logic']]] })
s('FALSE', '', [[]])
s('TRUE', '', [[]])
s('IF', 'b a a?', [[true, 'yes', 'no'], [false, 1, 2], [false, 1]])
s('IFERROR', 'a a', [['@div0', 'err'], [5, 'err'], ['@empty', 'x']])
s('IFNA', 'a a', [['@na', 'nf'], ['@div0', 'nf'], [3, 'nf']])
s('IFS', 'b a', [[false, 1, true, 2], [false, 1, false, 2]], { rep: 2 })
s('NOT', 'b', [[false], [true]])
s('SWITCH', 'a a a', [[2, 1, 'one', 2, 'two', 'other'], [3, 1, 'one', 2, 'two', 'other'], [3, 1, 'one', 2, 'two']], {
  rep: 2,
  extra: [['case', ['b', 'a', 'A', 'B', 'BB']], ['Cyrillic and case', ['мир', 'МИР', 1, 0]]],
})

// ---------- lookup and reference ----------
s('CHOOSE', 'i a', [[2, 'a', 'b', 'c'], [1.9, 'a', 'b'], [4, 'a', 'b']], { rep: 1 })
s('CHOOSECOLS', 'R i', [['@tbl', 1, 3], ['@tbl', -1]], { rep: 1, picks: [[1, 1], [2, 2], 'cols'] })
s('CHOOSEROWS', 'R i', [['@tbl', 1, 3], ['@tbl', -1]], { rep: 1, picks: [[1, 1], [2, 2], 'rows'] })
s('COLUMN', 'R', [['@nums1'], ['@tbl']], { extra: [['one cell', ['@date']], ['table by rows', ['@htbl']]] })
s('COLUMNS', 'R', [['@tbl'], ['@nums1']], { extra: [['one cell', ['@date']], ['table by rows', ['@htbl']], ['blank range', ['@blank']]] })
s('DROP', 'R i i?', [['@tbl', 2], ['@tbl', 0, 1], ['@tbl', -1, -1]], { picks: [[1, 1], [2, 2], 'rows', 'cols'] })
s('EXPAND', 'R i i? a?', [['@m1', 3, 3, 0], ['@m1', 3, 3]], { picks: [[1, 1], [3, 3], 'rows'] })
s('HLOOKUP', 'a R i b?', [[5, '@htbl', 2, false], [4, '@htbl', 3], [4, '@htbl', 3, false], [0, '@htbl', 2]], {
  extra: [['text, any case', ['FIVE', '@htbl', 3, false]], ['wildcard *', ['th*', '@htbl', 3, false]]],
})
s('HSTACK', 'R', [['@m1', '@m2']], {
  rep: 1,
  picks: [[1, 1], [2, 4], 'cols'],
  extra: [['different height: #N/A padding', ['@m1', '@lkeys'], [3, 1]], ['different height: rows', ['@m1', '@lkeys'], 'rows'], ['blank cells', ['@m1', '@mixed'], [5, 3]], ['text', ['@m1', '@lnames'], [2, 3]]],
})
s('INDEX', 'R i i?', [['@tbl', 2, 3], ['@nums1', 3], ['@tbl', 7, 1], ['@htbl', 2, 4]])
s('LOOKUP', 'a R R?', [[4.19, '@lkeys', '@lvals'], [5, '@lkeys', '@lnames'], [0, '@lkeys', '@lvals'], [100, '@lkeys', '@lvals'], [5, '@tbl']])
s('MATCH', 'a R i?', [[5, '@lkeys', 0], [4, '@lkeys'], [100, '@lkeys', 1]], {
  extra: [['text, any case', ['THREE', '@lnames', 0]], ['wildcard *', ['th*', '@lnames', 0]], ['wildcard ?', ['f?ve', '@lnames', 0]], ['descending -1', [4, '@lvals', -1]]],
})
s('ROW', 'R', [['@nums2'], ['@tbl']], { extra: [['one cell', ['@date']], ['table by rows', ['@htbl']]] })
s('ROWS', 'R', [['@tbl'], ['@nums1']], { extra: [['one cell', ['@date']], ['table by rows', ['@htbl']], ['blank range', ['@blank']]] })
s('SORT', 'R i? i? b?', [['@nums1'], ['@nums1', 1, -1], ['@txt'], ['@tbl', 2, 1]], { picks: [[1, 1], [2, 1], [7, 1], 'rows'] })
s('TAKE', 'R i i?', [['@tbl', 2], ['@tbl', -2, 2]], { picks: [[1, 1], [2, 2], 'rows', 'cols'] })
s('TRANSPOSE', 'R', [['@m23'], ['@nums1']], { picks: [[1, 2], [2, 1], 'rows', 'cols'] })
s('UNIQUE', 'R b? b?', [['@uniq'], ['@uniq', false, true]], {
  picks: [[1, 1], [4, 1], 'rows'],
  extra: [['case: a and A one value', ['@uniq'], [5, 1]], ['by columns', ['@htbl', true], 'cols'], ['Cyrillic and case', ['@txt'], 'rows'], ['blank cells', ['@mixed'], 'rows']],
})
s('VLOOKUP', 'a R i b?', [[5, '@tbl', 2, false], [4, '@tbl', 3], [4, '@tbl', 3, false], [0, '@tbl', 2]], {
  extra: [['text, any case', ['FIVE', '@ntbl', 2, false]], ['wildcard *', ['th*', '@ntbl', 2, false]], ['wildcard ?', ['f?ve', '@ntbl', 2, false]]],
})
s('VSTACK', 'R', [['@m1', '@m2']], {
  rep: 1,
  picks: [[1, 1], [4, 2], 'rows'],
  extra: [['different width: #N/A padding', ['@m1', '@m23'], [1, 3]], ['different width: columns', ['@m1', '@m23'], 'cols'], ['blank cells', ['@lkeys', '@mixed'], [11, 1]], ['text', ['@m1', '@htbl'], [4, 1]]],
})

// ---------- text ----------
s('CHAR', 'i', [[65], [33], [200], [0], [256]])
s('CLEAN', 's', [['@ctrl'], ['abc']])
s('CODE', 's', [['A'], ['!'], ['Ж'], ['']])
s('CONCAT', 'V', [['a', 'b', 'c'], ['@texts3']], { rep: 1, extra: [['range: mixed', ['@mixed']]] })
s('CONCATENATE', 's', [['a', 'b'], ['Привет', ' ', 'Мир'], [1.5, true]], { rep: 1 })
s('DOLLAR', 'n i?', [[1234.567, 2], [1234.567, -2], [-1234.567, -2], [-0.123, 4], [99.888]])
s('EXACT', 's s', [['word', 'word'], ['Word', 'word'], ['w ord', 'word'], ['Привет', 'привет']])
s('FIND', 's s i?', [['M', 'Miriam McGovern'], ['m', 'Miriam McGovern'], ['M', 'Miriam McGovern', 3], ['и', 'Привет мир'], ['', 'abc'], ['z', 'abc']])
s('FIXED', 'n i? b?', [[1234.567, 1], [1234.567, -1], [-1234.567, -1, true], [44.332]])
s('LEFT', 's i?', [['Sale Price', 4], ['Sweden'], ['Привет', 3], ['abc', 0]])
s('LEN', 's', [['Phoenix, AZ'], [''], ['  One   '], ['Привет']])
s('LOWER', 's', [['E. E. Cummings'], ['ПРИВЕТ Мир']])
s('MID', 's i i', [['Fluid Flow', 1, 5], ['Fluid Flow', 7, 20], ['Fluid Flow', 20, 5], ['Привет мир', 3, 4]])
s('NUMBERVALUE', 's s? s?', [['2.500,27', ',', '.'], ['3.5%'], ['1 234.5'], ['1,5', ','], ['2.5']])
s('PROPER', 's', [['this is a TITLE'], ['2-way street'], ['76BudGet'], ['привет мир']])
s('REPLACE', 's i i s', [['abcdefghijk', 6, 5, '*'], ['2009', 3, 2, '10'], ['123456', 1, 3, '#'], ['Привет', 1, 1, 'п']])
s('REPT', 's i', [['*-', 3], ['-', 10], ['ab', 0]])
s('RIGHT', 's i?', [['Sale Price', 5], ['Stock Number'], ['Привет', 2]])
s('SEARCH', 's s i?', [['e', 'Statements', 6], ['margin', 'Profit Margin'], ['МИР', 'Привет мир']], {
  extra: [['wildcard ?', ['p?o', 'Profit']], ['wildcard *', ['M*n', 'Profit Margin']], ['escaped ~?', ['~?', 'a?b']]],
})
s('SUBSTITUTE', 's s s i?', [['Sales Data', 'Sales', 'Cost'], ['Quarter 1, 2008', '1', '2', 1], ['Quarter 1, 2011', '1', '2', 3], ['aaa', 'a', 'b', 2], ['Привет', 'п', 'П']])
s('T', 'a', [['Rainfall'], [19], [true]])
s('TEXT', 'a s', [[1234.567, '$#,##0.00'], [45658, 'YYYY-MM-DD'], [0.285, '0.0%'], [45658, 'dd.mm.yyyy'], [1234.5, '0.00'], [-5, '0'], [45688.75, 'h:mm AM/PM'], [4.34, '# ?/?']])
s('TEXTJOIN', 's b V', [[', ', true, 'a', '', 'b'], [', ', false, 'a', '', 'b'], ['-', true, '@texts3']], { rep: 1, extra: [['range: mixed', [' ', true, '@mixed']]] })
s('TRIM', 's', [['  First Quarter Earnings  '], ['a   b'], ['Привет  мир']])
s('UNICHAR', 'i', [[66], [32], [1046], [0]])
s('UNICODE', 's', [['B'], [' '], ['Ж'], ['']])
s('UPPER', 's', [['total'], ['привет Мир'], ['ß']])
s('VALUE', 's', [['$1,000'], ['16:48:00'], ['1.5'], ['1,5'], ['10%'], ['1e3'], ['2025-01-31'], [' 12 ']])

// ---------- criteria (SUMIF and its kin): a shared template ----------
/** Triples (criteria range, criterion, value range) with a tag; the first two are typical. */
const CRIT: [string, Arg[]][] = [
  ['typical', ['@amt', '>4', '@amt']],
  ['typical', ['@txt', 'Аренда', '@amt']],
  ['wildcard *', ['@txt', 'Ар*', '@amt']],
  ['escaped ~*', ['@txt', 'x~*y', '@amt']],
  ['wildcard ?', ['@txt', 'Ча?', '@amt']],
  ['Cyrillic, any case', ['@txt', 'аренда', '@amt']],
  ['й and и differ', ['@txt', 'Чай', '@amt']],
  ['date text', ['@dates', '>=2025-01-01', '@amt']],
  ['date as a number', ['@dates', '>45658', '@amt']],
  ['not equal', ['@txt', '<>Аренда', '@amt']],
  ['empty criterion', ['@mixed', '', '@nums1']],
  ['number against numeric text', ['@mixed', 2, '@nums1']],
  ['boolean criterion', ['@mixed', true, '@nums1']],
  ['error in the criteria range', ['@witherr', '>1', '@nums1']],
  ['nothing matches', ['@txt', 'нет', '@amt']],
]
const critCase = (shape: (r: Arg, c: Arg, v: Arg) => Arg[]) => CRIT.map(([tag, [r, c, v]]) => [tag, shape(r, c, v)] as [string, Arg[]])
const critSig = (name: string, kinds: string, shape: (r: Arg, c: Arg, v: Arg) => Arg[], rep?: number) => {
  const all = critCase(shape)
  S.push({ name, kinds, rep, calls: all.slice(0, 2).map((x) => x[1]), extra: all.slice(2), noAuto: true })
}
critSig('SUMIF', 'R c R?', (r, c, v) => [r, c, v])
critSig('AVERAGEIF', 'R c R?', (r, c, v) => [r, c, v])
critSig('COUNTIF', 'R c', (r, c) => [r, c])
critSig('SUMIFS', 'R R c', (r, c, v) => [v, r, c], 2)
critSig('AVERAGEIFS', 'R R c', (r, c, v) => [v, r, c], 2)
critSig('MAXIFS', 'R R c', (r, c, v) => [v, r, c], 2)
critSig('MINIFS', 'R R c', (r, c, v) => [v, r, c], 2)
critSig('COUNTIFS', 'R c', (r, c) => [r, c], 2)
for (const n of ['SUMIFS', 'AVERAGEIFS', 'MAXIFS', 'MINIFS']) S.find((x) => x.name === n)!.extra!.push(['two criteria', ['@amt', '@txt', 'Ар*', '@amt', '>1']])
S.find((x) => x.name === 'COUNTIFS')!.extra!.push(['two criteria', ['@txt', 'Ар*', '@amt', '>1']])

// ---------- databases: a shared template ----------
const DB: [string, Arg[]][] = [
  ['typical', ['@db', 'Yield', '@critApple']],
  ['typical', ['@db', 4, '@critPear']],
  ['field, any case', ['@db', 'profit', '@critApple']],
  ['wildcard in the criterion', ['@db', 'Profit', '@critWild']],
  ['one row', ['@db', 'Age', '@critCherry']],
  ['nothing matches', ['@db', 'Yield', '@critNone']],
  ['empty criterion (all rows)', ['@db', 'Height', '@critAll']],
  ['field number 0', ['@db', 0, '@critPear']],
  ['no such field', ['@db', 'Weight', '@critPear']],
  ['field as header text', ['@db', 'Tree', '@critPear']],
]
for (const name of ['DAVERAGE', 'DCOUNT', 'DCOUNTA', 'DGET', 'DMAX', 'DMIN', 'DPRODUCT', 'DSTDEV', 'DSTDEVP', 'DSUM', 'DVAR', 'DVARP']) {
  S.push({ name, kinds: 'R a R', calls: DB.slice(0, 2).map((x) => x[1]), extra: DB.slice(2), noAuto: true })
}

// sizes and counts: a ceiling for edge values
for (const n of ['CHOOSECOLS', 'CHOOSEROWS', 'REPT', 'BESSELI', 'BESSELJ', 'BESSELK', 'BESSELY', 'CUMIPMT', 'CUMPRINC', 'IPMT', 'PPMT', 'ISPMT', 'DB', 'DDB', 'SYD', 'FV', 'PV', 'PMT', 'NPER', 'RATE', 'BASE', 'ROMAN', 'WORKDAY', 'WORKDAY.INTL']) {
  const sig = S.find((x) => x.name === n)
  if (!sig) throw new Error(`cap: no signature ${n}`)
  sig.cap = true
}

export const SIGNATURES: readonly Sig[] = S

/**
 * Functions that build an array of a size taken from an argument: the positions of the size arguments. Those
 * positions take only number literals of at most SIZE_LIMIT in absolute value: no references, text, dates or
 * booleans. MUNIT(45688) (a date as the size) and MUNIT("1.5") in ru-RU (the text became the date 46143) ran
 * Excel out of resources and raised a modal dialog; in ru-RU any text like "1.5", "1,5", "2025-01-01" can
 * become a number near 45,000–46,000. A guard checks every case (`cases.ts`, `cases.test.ts`). SEQUENCE,
 * RANDARRAY and MAKEARRAY are not in the set (formulajs lacks them).
 */
export const ARRAY_SIZE: Readonly<Record<string, readonly number[]>> = { MUNIT: [0], EXPAND: [1, 2], TAKE: [1, 2], DROP: [1, 2] }
export const SIZE_LIMIT = 100

/**
 * Functions written with a prefix in OOXML (newer than Excel 2007). The check: Excel gives `#NAME?` on every
 * case of a function it has when the prefix is missing.
 */
export const XLFN: Record<string, string> = Object.fromEntries(
  [
    'ACOT', 'ACOTH', 'AGGREGATE', 'ARABIC', 'BASE', 'BETA.DIST', 'BETA.INV', 'BINOM.DIST', 'BINOM.DIST.RANGE', 'BINOM.INV', 'BITAND',
    'BITLSHIFT', 'BITOR', 'BITRSHIFT', 'BITXOR', 'CEILING.MATH', 'CEILING.PRECISE', 'CHISQ.DIST', 'CHISQ.DIST.RT', 'CHISQ.INV',
    'CHISQ.INV.RT', 'CHISQ.TEST', 'CHOOSECOLS', 'CHOOSEROWS', 'COMBINA', 'CONCAT', 'CONFIDENCE.NORM', 'CONFIDENCE.T', 'COT', 'COTH',
    'COVARIANCE.P', 'COVARIANCE.S', 'CSC', 'CSCH', 'DAYS', 'DECIMAL', 'DROP', 'ERF.PRECISE', 'ERFC.PRECISE', 'EXPAND', 'EXPON.DIST',
    'F.DIST', 'F.DIST.RT', 'F.INV', 'F.INV.RT', 'F.TEST', 'FLOOR.MATH', 'FLOOR.PRECISE', 'GAMMA', 'GAMMA.DIST', 'GAMMA.INV',
    'GAMMALN.PRECISE', 'GAUSS', 'HSTACK', 'HYPGEOM.DIST', 'IFNA', 'IFS', 'IMCOSH', 'IMCOT', 'IMCSC', 'IMCSCH', 'IMSEC', 'IMSECH',
    'IMSINH', 'IMTAN', 'ISO.CEILING', 'ISOWEEKNUM', 'LOGNORM.DIST', 'LOGNORM.INV', 'MAXIFS', 'MINIFS', 'MODE.MULT', 'MODE.SNGL',
    'MUNIT', 'NEGBINOM.DIST', 'NETWORKDAYS.INTL', 'NORM.DIST', 'NORM.INV', 'NORM.S.DIST', 'NORM.S.INV', 'NUMBERVALUE', 'PDURATION',
    'PERCENTILE.EXC', 'PERCENTILE.INC', 'PERCENTRANK.EXC', 'PERCENTRANK.INC', 'PERMUTATIONA', 'PHI', 'POISSON.DIST', 'QUARTILE.EXC',
    'QUARTILE.INC', 'RANK.AVG', 'RANK.EQ', 'RRI', 'SEC', 'SECH', 'SKEW.P', 'STDEV.P', 'STDEV.S', 'SWITCH', 'T.DIST', 'T.DIST.2T',
    'T.DIST.RT', 'T.INV', 'T.INV.2T', 'T.TEST', 'TAKE', 'TEXTJOIN', 'UNICHAR', 'UNICODE', 'UNIQUE', 'VAR.P', 'VAR.S', 'VSTACK',
    'WEIBULL.DIST', 'WORKDAY.INTL', 'XOR', 'Z.TEST',
  ].map((n) => [n, `_xlfn.${n}`] as const).concat([['SORT', '_xlfn._xlws.SORT']]),
)
