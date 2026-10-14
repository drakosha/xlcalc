/**
 * Held out after `rules-held`, predicted by the engine (`editStructure`'s outline operations, `outlineGroups`,
 * `writeXlsx`) before Excel ran it: each rule `rules-held` gave, tried where it was not measured. Group over lines at the
 * deepest level beside others; Show and Hide Detail at a line of a group and at a summary line three groups share, at a
 * line of a group without a summary line, refused where the detail is shown or hidden already; the marks Excel keeps
 * when lines are hidden and shown by hand, a summary line hidden by hand among them, and a mark a file holds inside a
 * group whose lines are all hidden; level buttons over three levels in turn, and on rows and columns at once; column groups
 * at either edge of the sheet, with and without a summary column.
 */
import type { ColumnInput, RowInput } from '../../../src/contract.js'
import { H, hide, L, levels, outlineAdder, type OutlineBook, type OutlineCase, show } from './cases.js'
import { BOOKS as HELD } from './rules-held.js'

export const BOOKS: Record<string, OutlineBook> = {
  nested: HELD.nested,
  deep: HELD.deep,
  top: HELD.top,
  topBelow: HELD.topBelow,
  bottom: HELD.bottom,
  right: HELD.right,
  // groups 2:9 (1, summary row 10), 4:7 (2, summary 8) and 6 (3, summary 7)
  three: { rows: { '2:3': L(1), '4:5': L(2), 6: L(3), 7: L(2), '8:9': L(1) } },
  // groups 3:7 (1), 5:7 (2) and 7 (3) share the summary row 8
  deepShared: { rows: { '3:4': L(1), '5:6': L(2), 7: L(3) } },
  // the groups of `nested` all hidden, the inner one's summary row 8 marked
  nestedMarked: { rows: { '3:4': H(1), '5:7': H(2), 8: { outlineLevel: 1, hidden: true, collapsed: true }, '9:10': H(1) } },
  // rows of `nested` and columns C:J (1) and E:G (2)
  both: { rows: HELD.nested.rows, columns: HELD.cols.columns },
  // columns A:C with summaries left: no summary column
  leftEdge: { columns: { 'A:C': L(1) }, outline: { summaryRight: false } },
  // columns XFB:XFD with summaries left: the summary column XFA
  rightLeft: { columns: { 'XFB:XFD': L(1) }, outline: { summaryRight: false } },
  // columns XFA:XFC: the summary column XFD
  rightInner: { columns: { 'XFA:XFC': L(1) } },
  // columns XFA:XFD (1) and XFC:XFD (2), at the edge
  rightNested: { columns: { 'XFA:XFB': L(1), 'XFC:XFD': L(2) } },
}

export function buildOutlineCases(): OutlineCase[] {
  const { out, add, g } = outlineAdder()

  g('group', 'deep', { group: '3:6' })
  g('group', 'deep', { group: '5:6' })
  g('group', 'deep', { group: '3:4' })
  g('group', 'deep', { group: '2:3' })

  g('detail', 'nested', hide('9'))
  g('detail', 'nested', hide('5'))
  g('detail', 'nested', show('9'))
  g('detail', 'nested', hide('8'), hide('6'))
  g('detail', 'nested', hide('11'), show('4'))
  g('detail', 'deepShared', hide('7'))
  g('detail', 'deepShared', hide('8'), show('8'))
  g('detail', 'deepShared', hide('6'), show('7'))
  g('detail', 'three', hide('7'), hide('8'), show('8'))
  g('detail', 'top', hide('2'))
  g('detail', 'topBelow', hide('2'), show('2'))
  add('detail', 'bottom', [hide('1048575')], { rows: '1048570:1048576', columns: 'A:C' })
  const right = { rows: '1:3', columns: 'XEX:XFD' }
  add('detail', 'right', [hide('XFC')], right)
  add('detail', 'leftEdge', [hide('B')], { rows: '1:3', columns: 'A:E' })

  g('hand', 'nested', { hidden: '8' }, hide('8'))
  g('hand', 'nested', { hidden: '11' }, { hidden: '3:10' })
  g('hand', 'nested', { hidden: '3:10' }, { shown: '8' })
  g('hand', 'nested', hide('8'), { hidden: '3:4' }, { hidden: '8:10' }, show('11'))
  g('hand', 'nested', { hidden: '6' }, hide('8'), show('8'))
  g('hand', 'nestedMarked')
  g('hand', 'nestedMarked', show('11'))
  g('hand', 'nestedMarked', show('11'), show('8'))

  g('levels', 'three', levels(3), levels(2), show('8'), levels(1), show('10'))
  g('levels', 'three', levels(2), hide('10'), levels(8))
  g('levels', 'both', levels(2, 1))
  g('levels', 'both', levels(1, 2), levels(8, 8))
  g('levels', 'topBelow', levels(1))
  add('levels', 'leftEdge', [levels(0, 1)], { rows: '1:3', columns: 'A:E' })
  add('levels', 'rightLeft', [levels(0, 1)], right)
  add('levels', 'rightInner', [levels(0, 1)], right)
  add('levels', 'rightNested', [levels(0, 2)], right)
  add('levels', 'right', [{ group: 'XFA' }, levels(0, 1)], right)
  add('levels', 'right', [{ ungroup: 'XFD' }, levels(0, 1)], right)

  g('ungroup', 'three', { ungroup: '2:9' })
  g('ungroup', 'nested', hide('11'), { ungroup: '5:7' })
  g('ungroup', 'nested', hide('8'), { ungroup: '3:4' })
  g('ungroup', 'both', hide('K'), { ungroup: 'C:J' })
  return out
}
