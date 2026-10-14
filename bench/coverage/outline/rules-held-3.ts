/**
 * Held out after `rules-held-2`, predicted by the engine before Excel ran it: the rules that set gave. Group and Ungroup
 * of columns that take in XFD, which Excel's object model refused once, and of rows at the sheet's last row; level buttons
 * over columns grouped up to XFD, beside and inside other groups, and at the left edge; the marks a file's rows give
 * when a summary row shows over a group hidden by hand, as Excel opens such a file and as it keeps them while lines are
 * hidden and shown by hand.
 */
import type { ColumnInput, RowInput } from '../../../src/contract.js'
import { collector, H, L, levels, type OutlineBook, type OutlineCase } from './cases.js'
import { BOOKS as HELD } from './rules-held.js'

export const BOOKS: Record<string, OutlineBook> = {
  empty: HELD.empty,
  nested: HELD.nested,
  right: HELD.right,
  bottom: HELD.bottom,
  // columns XFA:XFD (1) and XFA:XFB inside (2, summary column XFC)
  xfdInner: { columns: { 'XFA:XFB': L(2), 'XFC:XFD': L(1) } },
  // columns XEY:XFB (1, summary column XFC) and XFD (1) apart
  xfdApart: { columns: { 'XEY:XFB': L(1), XFD: L(1) } },
  // columns A:D (1) and A:B inside (2), summaries left: neither has a summary column
  leftNested: { columns: { 'A:B': L(2), 'C:D': L(1) }, outline: { summaryRight: false } },
  // the rows of `nested` hidden but the summary row 8, with no mark, as a file holds them
  shownSummary: { rows: { '3:4': H(1), '5:7': H(2), 8: L(1), '9:10': H(1) } },
}

export function buildOutlineCases(): OutlineCase[] {
  const { out, add } = collector()
  const right = { rows: '1:3', columns: 'XEX:XFD' }
  const bottom = { rows: '1048570:1048576', columns: 'A:C' }

  add('xfd', 'empty', [{ group: 'XFD' }], right)
  add('xfd', 'empty', [{ group: 'XFC:XFD' }], right)
  add('xfd', 'empty', [{ group: 'XFC' }], right)
  add('xfd', 'right', [{ ungroup: 'XFC:XFD' }], right)
  add('xfd', 'right', [{ ungroup: 'XFC' }], right)
  add('xfd', 'right', [{ ungroup: 'XFB:XFC' }, levels(0, 1)], right)
  add('xfd', 'bottom', [{ ungroup: '1048576' }], bottom)
  add('xfd', 'bottom', [{ group: '1048575:1048576' }, levels(2)], bottom)

  add('levels', 'xfdInner', [levels(0, 2)], right)
  add('levels', 'xfdInner', [levels(0, 1)], right)
  add('levels', 'xfdApart', [levels(0, 1)], right)
  add('levels', 'xfdApart', [{ hide: 'XFC' }, levels(0, 2)], right)
  add('levels', 'leftNested', [levels(0, 1), levels(0, 2)], { rows: '1:3', columns: 'A:F' })

  add('hand', 'shownSummary', [])
  add('hand', 'shownSummary', [{ show: '11' }])
  add('hand', 'nested', [{ hidden: '3:10' }, { shown: '3:4' }])
  add('hand', 'nested', [{ hidden: '5:7' }, { hidden: '8' }])
  add('hand', 'nested', [{ hidden: '3:10' }, { shown: '8' }, { hide: '11' }, { show: '11' }])
  return out
}
