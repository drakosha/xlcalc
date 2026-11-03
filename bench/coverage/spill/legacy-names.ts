/**
 * A held-out set of formulas of the old kind over names (`legacy-index-held*`): one case on sheet S over B1:B8, each
 * formula's cell of the old kind; the text Excel shows (`shown`) and the values (`names`) as predicted before Excel ran.
 */
import { one, type CellSpec, type SpillCase } from './cases.js'
import { fromLegacy } from '../../../src/xlsx.js'
import type { NameInput } from '../../../src/contract.js'
import type { Val } from '../formulajs.js'

const DATA: Record<string, CellSpec> = Object.fromEntries([4, 8, 15, 16, 23, 42, -3.5, 10.25].map((v, i) => [`B${i + 1}`, v]))

export function legacyNames(about: string, names: NameInput[], formulas: Record<string, string>, predicted: { names: unknown; shown: Record<string, string> }) {
  const at = Object.keys(formulas)
  return {
    buildSpillCases: (): SpillCase[] => [
      one('names', about, { ...DATA, ...formulas }, at, { book: { sheets: [], names }, legacy: at.map((a) => `S!${a}`), predicted: predicted.names as Record<string, Val> }),
    ],
    /** The text Excel 365 was predicted to show for each formula (canonical, without `=`), committed before Excel ran. */
    SHOWN: predicted.shown as Readonly<Record<string, string>>,
    /** The converter's text of each formula, written into the prediction. */
    converted: () => Object.fromEntries(at.map((a) => [`S!${a}`, fromLegacy(formulas[a].slice(1), { names, sheet: 'S' })])),
  }
}
