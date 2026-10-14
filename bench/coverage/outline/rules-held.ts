/**
 * Held out, predicted before Excel ran it by the grid's own model (`xlcalc-sheet/src/local/outline.ts`, ported in
 * `predict.ts`). The questions: what Group and Ungroup do to lines at different levels, at the deepest level and with
 * none; which group ShowDetail on a summary line, on a group's lines and on one of its lines hides and shows, nested
 * groups hidden and shown in each order, groups sharing a summary line; what ShowLevels hides, shows and marks; rows
 * hidden by hand inside a group; summaries above and left; groups at the sheet's edges with no summary line; the
 * collapsed marks Excel saves, for groups it collapsed and for groups whose lines a book holds hidden without a mark
 * (as `writeXlsx` writes lines hidden through `setLayout`).
 */
import type { ColumnInput, RowInput } from '../../../src/contract.js'
import { group, H, hide, L, levels, outlineAdder, type OutlineBook, type OutlineCase, show, ungroup } from './cases.js'

export const BOOKS: Record<string, OutlineBook> = {
  empty: {},
  // groups 3:10 (level 1, summary row 11) and 5:7 (level 2, summary row 8): the grid's example
  nested: { rows: { '3:4': L(1), '5:7': L(2), '8:10': L(1) } },
  // the grid's first example: groups 2:10 (1), 2:4 (2) and 6:9 (2), the last collapsed (summary row 10)
  example: { rows: { '2:4': L(2), 5: L(1), '6:9': H(2), 10: { outlineLevel: 1, collapsed: true } } },
  // groups 3:7 (1) and 5:7 (2) share the summary row 8
  shared: { rows: { '3:4': L(1), '5:7': L(2) } },
  deep: { rows: { '3:5': L(7) } },
  // groups 3:5 and 7:9 apart, row 6 the first's summary
  apart: { rows: { '3:5': L(1), '7:9': L(1) } },
  // row 5 hidden by hand inside the group 3:10, row 13 outside any group
  hand: { rows: { '3:4': L(1), 5: H(1), '6:10': L(1), 13: { hidden: true } } },
  // a group's rows hidden with no mark, as `writeXlsx` writes rows hidden through `setLayout`
  allHidden: { rows: { '2:4': H(1) } },
  allHiddenNested: { rows: { '3:4': H(1), '5:7': H(2), '8:10': H(1) } },
  // a mark on the summary row of a group whose rows are shown
  markShown: { rows: { '2:4': L(1), 5: { collapsed: true } } },
  // summaries above: 3:10's is row 2, 5:7's row 4
  above: { rows: { '3:4': L(1), '5:7': L(2), '8:10': L(1) }, outline: { summaryBelow: false } },
  // rows 1:3 with summaries above: no summary row; below: row 4
  top: { rows: { '1:3': L(1) }, outline: { summaryBelow: false } },
  topBelow: { rows: { '1:3': L(1) } },
  bottom: { rows: { '1048574:1048576': L(1) } },
  // columns C:J (1, summary K) and E:G (2, summary H)
  cols: { columns: { 'C:D': L(1), 'E:G': L(2), 'H:J': L(1) } },
  // summaries left: C:J's is B, E:G's is D
  colsLeft: { columns: { 'C:D': L(1), 'E:G': L(2), 'H:J': L(1) }, outline: { summaryRight: false } },
  right: { columns: { 'XFB:XFD': L(1) } },
}

export function buildOutlineCases(): OutlineCase[] {
  const { out, add, g } = outlineAdder()

  g('group', 'empty', group('3:5'))
  g('group', 'empty', group('3:5'), group('6:8'))
  g('group', 'empty', group('3:5'), group('7:9'))
  g('group', 'nested', group('4:6'))
  g('group', 'nested', group('9:12'))
  g('group', 'nested', group('5:7'))
  g('group', 'empty', group('3:5'), group('2:6'))
  g('group', 'deep', group('3:5'))
  g('group', 'deep', group('4'))
  g('group', 'deep', group('2:6'))
  g('group', 'empty', group('B:D'), group('C'))
  g('group', 'nested', hide('8'), group('5:7'))
  g('group', 'nested', hide('8'), group('4:9'))

  g('ungroup', 'nested', ungroup('5:6'))
  g('ungroup', 'nested', ungroup('3:10'))
  g('ungroup', 'nested', ungroup('1:2'))
  g('ungroup', 'nested', ungroup('2:4'))
  g('ungroup', 'nested', hide('11'), ungroup('3:10'))
  g('ungroup', 'nested', hide('8'), ungroup('5:7'))
  g('ungroup', 'cols', ungroup('A'))
  g('ungroup', 'nested', ungroup('11'))

  g('detail', 'nested', hide('8'))
  g('detail', 'nested', hide('11'))
  g('detail', 'nested', hide('8'), hide('11'), show('11'))
  g('detail', 'nested', hide('11'), show('11'))
  g('detail', 'nested', hide('11'), hide('8'), show('11'))
  g('detail', 'nested', hide('8'), hide('11'), show('8'))
  g('detail', 'nested', hide('5:7'))
  g('detail', 'nested', hide('3:10'))
  g('detail', 'nested', hide('6'))
  g('detail', 'nested', hide('4'))
  g('detail', 'nested', show('8'))
  g('detail', 'nested', hide('12'))
  g('detail', 'shared', hide('8'), hide('8'), show('8'), show('8'))
  g('detail', 'apart', hide('6'), hide('10'), show('6'))
  g('detail', 'empty', group('3:5'), group('6:8'), hide('9'))
  g('detail', 'example', show('10'), hide('11'), show('11'))
  g('detail', 'nested', hide('8'), show('11'))
  g('detail', 'nested', hide('11'), show('8'))

  g('hand', 'hand', hide('11'), show('11'))
  g('hand', 'hand', show('11'))
  g('hand', 'hand', levels(1), levels(2))
  g('hand', 'empty', group('2:4'), { hidden: '2:4' }, show('5'))
  g('hand', 'empty', group('2:4'), { hidden: '2:4' })
  g('hand', 'nested', { hidden: '5:7' }, show('8'))
  g('hand', 'nested', { hidden: '3:10' }, show('11'))
  g('hand', 'nested', hide('8'), { shown: '5:7' })

  g('file', 'allHidden')
  g('file', 'allHidden', show('5'))
  g('file', 'allHiddenNested', show('11'))
  g('file', 'allHiddenNested', show('8'))
  g('file', 'markShown')
  g('file', 'markShown', hide('5'), show('5'))
  g('file', 'example')

  g('levels', 'nested', levels(1))
  g('levels', 'nested', levels(2))
  g('levels', 'nested', levels(8))
  g('levels', 'nested', hide('8'), levels(8))
  g('levels', 'nested', hide('11'), levels(2))
  g('levels', 'nested', levels(1), levels(2))
  g('levels', 'nested', levels(3))
  g('levels', 'empty', levels(1))
  g('levels', 'cols', levels(0, 1))
  g('levels', 'nested', levels(1), show('11'))
  g('levels', 'hand', levels(8))
  g('levels', 'nested', levels(9))
  g('levels', 'example', levels(8))

  g('above', 'above', hide('4'))
  g('above', 'above', hide('2'))
  g('above', 'above', hide('4'), hide('2'), show('2'))
  g('above', 'above', levels(1))
  g('above', 'above', levels(2))
  g('above', 'above', hide('5:7'))
  g('left', 'colsLeft', hide('D'))
  g('left', 'colsLeft', hide('D'), hide('B'), show('B'))
  g('left', 'colsLeft', levels(0, 2))

  g('edge', 'top', levels(1))
  g('edge', 'top', hide('1:3'))
  g('edge', 'top', hide('1:3'), show('1:3'))
  g('edge', 'topBelow', hide('4'), show('4'))
  const bottom = { rows: '1048570:1048576', columns: 'A:C' }
  add('edge', 'bottom', [levels(1)], bottom)
  add('edge', 'bottom', [hide('1048574:1048576'), show('1048574:1048576')], bottom)
  const right = { rows: '1:3', columns: 'XEZ:XFD' }
  add('edge', 'right', [levels(0, 1)], right)
  add('edge', 'right', [hide('XFB:XFD')], right)

  g('cols', 'cols', hide('H'), hide('K'), show('K'))
  g('cols', 'cols', ungroup('E'))
  g('cols', 'cols', hide('E:G'))
  return out
}
