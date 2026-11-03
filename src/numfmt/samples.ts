/**
 * The samples of the Number Format list on Excel's Home tab (`numberFormatSamples`), measured in Excel 16.0 for
 * Windows in en-US and ru-RU (`test/excel/win/samples/`): under General the list writes that the item has no specific
 * format, whatever the cell holds; under each other item the value in the item's format, whatever the cell's own
 * format and its column's width; a number the format cannot show (a date past the calendar) as ten `#`; a sample past
 * 50 characters cut to 50 and three dots; nothing for an empty cell.
 */
import type { CellValue, FormatSample, FormatSampleOptions } from '../contract.js'
import { epochOf } from '../dates.js'
import { checkCellValue, checkSampleOptions } from '../input.js'
import { optionalLocale } from '../locale/load.js'
import { displayRules } from '../locale/rules.js'
import { choiceFormats, noFormatText } from './choices.js'

/** The longest sample the list writes whole. */
const LONGEST = 50

/**
 * Under each item of the Number Format list in `opts.locale`, in list order, what Excel shows for the active cell's
 * `value` in the workbook's `opts.dateSystem`: `No specific format` (`Формат не указан`) under General, the value in
 * the item's format under the others (`1,234.50`, `$1,234.50`, `5/18/1903`), `##########` for a number the format cannot
 * show, a sample past 50 characters cut to 50 and `...`, nothing for an empty cell.
 */
export function numberFormatSamples(value: CellValue, opts?: FormatSampleOptions): FormatSample[] {
  const v = checkCellValue(value, 'value')
  const o = checkSampleOptions(opts)
  const loc = optionalLocale(o.locale, 'opts.locale')
  const r = displayRules(loc, epochOf(o.dateSystem))
  return choiceFormats(loc).map(([key, format]) => {
    if (key === 'general') return { key, sample: noFormatText(loc) }
    const text = r.display(v, format, Infinity).text
    // a number its format cannot show fills a column without a width with `#`: the list writes ten
    const sample = typeof v === 'number' && /^#+$/.test(text) ? '#'.repeat(10) : text
    return { key, sample: sample.length > LONGEST ? `${sample.slice(0, LONGEST)}...` : sample }
  })
}
