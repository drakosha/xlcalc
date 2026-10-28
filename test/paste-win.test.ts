/** Paste Special against the answers of Excel for Windows (`bench/windows/paste.ts`), as `paste-excel.test.ts` checks the Mac's. */
import { enCode } from '../src/numfmt/builtin.js'
import type { Shown } from '../bench/coverage/paste/shown.js'
import { asRu, INTO_MERGED, PASTE_KNOWN, PASTE_RU_SETS, PASTE_SETS, pasteSets, withFormats } from './paste-sets.js'

/** Where the engine pastes otherwise than Excel for Windows, by set and case: the Mac's, as both Excels paste alike, and All Merging Conditional Formats, which Windows alone has. */
const WIN_KNOWN: Record<string, string> = { ...PASTE_KNOWN, 'merging-held/into-merged': INTO_MERGED }

/** A row's height on Excel for Windows's grid of whole pixels at 96 dpi, 0.75 pt: it saves a row fitted to 26 pt as 26.25. */
const onGrid = (h: number) => Math.round(h / 0.75) * 0.75

/** Heights on Excel for Windows's pixel grid, as it saves them. */
const gridded = (shown: Shown): Shown => Object.fromEntries(Object.entries(shown).map(([name, s]) => [name, { ...s, heights: Object.fromEntries(Object.entries(s.heights).map(([k, h]) => [k, onGrid(h)])) }]))

/**
 * A book as Excel for Windows saves it, which is not the paste: a built-in number format by its code (it saves
 * `m/d/yyyy` as format 14), heights on its pixel grid; in ru-RU the codes as ru-RU shows them (`asRu`).
 */
const asWindows = (shown: Shown): Shown => withFormats(gridded(shown), enCode)

pasteSets('Excel for Windows', 'win/paste', PASTE_SETS, WIN_KNOWN, asWindows)
pasteSets('Excel for Windows ru-RU', 'win/paste/ru-RU', PASTE_RU_SETS, {}, (x) => asRu(gridded(x)))
