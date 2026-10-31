/**
 * `pivotCells`, `pivotCellAt`, `renamedValueField`, `pivotBlankLabel` and the sizes of a new slicer and timeline as a host
 * calls them, from the root and from `xlcalc/text`: over a session's `getLayout()` and a workbook, by sheet and PivotTable in any case; null outside a PivotTable and between its
 * report filters; refusing input of the wrong shape.
 */
import { describe, expect, it } from 'vitest'
import * as root from '../src/index.js'
import * as text from '../src/text.js'
import { createSession, InputError, type BookInput } from '../src/index.js'
import { ruRU } from '../src/locales/ru-RU.js'

const book = (): BookInput => ({
  sheets: [
    { name: 'Data', cells: { A1: { value: 'K' }, B1: { value: 'N' }, C1: { value: 'M' }, A2: { value: 'a' }, B2: { value: 2 }, C2: { value: 'x' } } },
    {
      name: 'P',
      cells: { A1: { value: 'M' }, B1: { value: '(All)' }, A3: { value: 'Row Labels' }, B3: { value: 'Sum of N' }, A4: { value: 'a' }, B4: { value: 2 }, A5: { value: 'Grand Total' }, B5: { value: 2 } },
      pivotTables: [
        {
          name: 'PT',
          cache: 1,
          location: { range: 'A3:B5', firstHeaderRow: 1, firstDataRow: 1, firstDataColumn: 1, pageRows: 1, pageColumns: 1 },
          fields: [{ items: [{ item: 0 }] }, {}, { items: [{ item: 0 }] }],
          rows: [0],
          pages: [{ field: 2 }],
          values: [{ field: 1, name: 'Sum of N' }],
          rowLines: [{ items: [0] }, { type: 'grand', items: [0] }],
          columnLines: [{ items: [] }],
        },
      ],
    },
  ],
  pivotCaches: [{ id: 1, source: { type: 'range', sheet: 'Data', range: 'A1:C2' }, fields: [{ name: 'K', items: ['a'] }, { name: 'N' }, { name: 'M', items: ['x'] }], records: [[0, 2, 0]] }],
})

describe.each([['the root', root], ['xlcalc/text', text]] as const)('the cells of a PivotTable from %s', (_, m) => {
  it('tells each cell over a workbook and a layout, the Filters area too', () => {
    const layout = createSession(book()).getLayout()
    for (const b of [book(), layout]) {
      // a report filter's button is drawn on its item, Row Labels' filters the rows
      expect(m.pivotCellAt(b, 'p', 'A1')).toEqual({ pivot: 'PT', kind: 'field', axis: 'pages', field: 2 })
      expect(m.pivotCellAt(b, 'P', 'B1')).toEqual({ pivot: 'PT', kind: 'pageItem', axis: 'pages', field: 2, button: true })
      expect(m.pivotCellAt(b, 'P', 'A3')).toEqual({ pivot: 'PT', kind: 'blank', axis: 'rows', button: true })
      expect(m.pivotCellAt(b, 'P', 'B3')).toEqual({ pivot: 'PT', kind: 'valueName', value: 0 })
      expect(m.pivotCellAt(b, 'P', 'B4')).toEqual({ pivot: 'PT', kind: 'value', value: 0, rowItems: [{ field: 0, item: 0 }] })
      expect(m.pivotCellAt(b, 'P', 'A5')).toEqual({ pivot: 'PT', kind: 'grand', axis: 'rows', value: 0 })
      // the blank row under the filters, a cell beside them and one outside are of no PivotTable
      expect([m.pivotCellAt(b, 'P', 'A2'), m.pivotCellAt(b, 'P', 'C1'), m.pivotCellAt(b, 'P', 'D4'), m.pivotCellAt(b, 'Data', 'A1'), m.pivotCellAt(b, 'Nope', 'A1')]).toEqual([null, null, null, null, null])
      const all = m.pivotCells(b, 'P', 'pt')!
      expect([Object.keys(all.cells), all.buttons]).toEqual([['A3', 'B3', 'A4', 'B4', 'A5', 'B5', 'A1', 'B1'], ['B1', 'A3']])
      expect(m.pivotCells(b, 'P', 'Other')).toBeNull()
    }
  })

  it('refuses input of the wrong shape', () => {
    expect(() => m.pivotCellAt(book(), 'P', 'A1:B2')).toThrow(InputError)
    expect(() => m.pivotCellAt(null as never, 'P', 'A1')).toThrow(InputError)
    expect(() => m.pivotCells(book(), 'P', 3 as never)).toThrow(InputError)
  })

  it('names a value field by its new function as Excel does', () => {
    expect(m.renamedValueField('Sum of Qty', 'sum', 'count', 'Qty')).toBe('Count of Qty')
    expect(m.renamedValueField('sum of qty2', 'sum', 'average', 'Qty')).toBe('Average of qty2')
    expect(m.renamedValueField('Quantity', 'sum', 'count', 'Qty')).toBe('Count of Qty')
    expect(m.renamedValueField(undefined, 'sum', 'max', 'Qty')).toBe('Max of Qty')
    expect(m.renamedValueField('Сумма по полю Qty', 'sum', 'count', 'Qty')).toBe('Count of Qty')
    expect(m.renamedValueField('Сумма по полю Qty2', 'sum', 'count', 'Qty', ruRU)).toBe('Количество по полю Qty2')
    expect(() => m.renamedValueField('x', 'sum', 'median' as never, 'Qty')).toThrow(InputError)
    // a locale by the tag of a loaded one or as an object, as the other functions take it: not a language alone
    expect(() => m.renamedValueField('Сумма по полю Qty2', 'sum', 'count', 'Qty', 'ru')).toThrow(InputError)
  })

  it("gives a blank item's label in the workbook's language and the size of a new slicer and timeline", () => {
    expect([m.pivotBlankLabel(), m.pivotBlankLabel('en-US'), m.pivotBlankLabel(ruRU)]).toEqual(['(blank)', '(blank)', '(пусто)'])
    expect(() => m.pivotBlankLabel('ru')).toThrow(InputError)
    expect(() => m.pivotBlankLabel('RU-ru')).toThrow(InputError)
    expect(() => m.pivotBlankLabel(3 as never)).toThrow(InputError)
    // in EMU, as a drawing's place takes its offsets
    expect([m.DEFAULT_SLICER_SIZE, m.DEFAULT_TIMELINE_SIZE]).toEqual([{ width: 1_828_800, height: 2_524_125 }, { width: 3_333_750, height: 1_371_600 }])
    expect(Object.isFrozen(m.DEFAULT_SLICER_SIZE) && Object.isFrozen(m.DEFAULT_TIMELINE_SIZE)).toBe(true)
  })
})
