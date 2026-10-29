/**
 * All Merging Conditional Formats in Excel for Mac, through its Paste Special window: the cases of `merging` and
 * `merging-held` of that kind, which Excel for Windows ran and the Mac's AppleScript could not. Predicted by the engine
 * before Excel ran them (`merging-mac.predicted.json`), as the rules Windows gave.
 */
import { buildPasteCases as merging } from './merging.js'
import { buildPasteCases as held } from './merging-held.js'
import type { PasteCases } from './cases.js'

export function buildPasteCases(): PasteCases {
  const sets = [merging(), held()]
  return {
    books: Object.assign({}, ...sets.map((s) => s.books)),
    cases: sets.flatMap((s) => s.cases.filter((c) => c.what === 'allMergingConditionalFormats')),
  }
}
