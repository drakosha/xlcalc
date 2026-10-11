/**
 * `patchLines`: a map of `SheetInput.columns` or `rows` patched on a span without a session, the map `setLayout` writes
 * for the same edit: spans the patch cuts split, the patch's fields over each entry's, null removing a field and an entry
 * left with none dropped, gaps given the patch alone; the same function in `xlcalc/a1`.
 */
import { describe, expect, it } from 'vitest'
import * as a1 from '../src/a1.js'
import { createSession, InputError, patchLines, type ColumnInput, type LinePatch, type RowInput } from '../src/index.js'

/** The map `setLayout` leaves for the edit, columns or rows by the span's shape. */
function bySession(map: Record<string, ColumnInput & RowInput> | undefined, span: string, set: LinePatch<ColumnInput & RowInput>) {
  const field = /^[A-Z]/.test(span) ? 'columns' : 'rows'
  const session = createSession({ styles: [{ font: { bold: true } }], sheets: [{ name: 'S', cells: { A1: { value: 1 } }, ...(map ? { [field]: map } : {}) }] })
  session.setLayout([{ sheet: 'S', [field]: span, set } as never])
  return session.getLayout().sheets[0][field] ?? {}
}

const CASES: [string, Record<string, ColumnInput & RowInput> | undefined, string, LinePatch<ColumnInput & RowInput>, Record<string, ColumnInput & RowInput>][] = [
  ['a span cut by the patch', { 'B:F': { width: 10 } }, 'D', { hidden: true }, { 'B:C': { width: 10 }, D: { width: 10, hidden: true }, 'E:F': { width: 10 } }],
  ['the fields merged over each entry, gaps given the patch alone', { B: { width: 10 }, 'D:E': { width: 12, hidden: true } }, 'A:D', { hidden: true }, { A: { hidden: true }, B: { width: 10, hidden: true }, C: { hidden: true }, D: { width: 12, hidden: true }, E: { width: 12, hidden: true } }],
  ['null removes a field, an entry left with none is dropped', { 'B:D': { width: 10 }, F: { width: 8, hidden: true } }, 'C:F', { width: null }, { B: { width: 10 }, F: { hidden: true } }],
  ['adjacent equal entries stay apart, as setLayout writes them', { B: { width: 10 } }, 'C', { width: 10 }, { B: { width: 10 }, C: { width: 10 } }],
  ['no map', undefined, 'XFD', { width: 3 }, { XFD: { width: 3 } }],
  ['rows, a gap a span of its own', { '5:9': { height: 20 }, 12: { outlineLevel: 1 } }, '7:12', { hidden: true, height: null }, { '5:6': { height: 20 }, '7:9': { hidden: true }, '10:11': { hidden: true }, 12: { outlineLevel: 1, hidden: true } }],
  ['one row of a span, a style and a format', { '1:3': { style: 0 } }, '2', { numberFormat: '0%', style: null }, { 1: { style: 0 }, 2: { numberFormat: '0%' }, 3: { style: 0 } }],
]

describe('patchLines', () => {
  it.each(CASES)('%s: the map setLayout writes', (_, map, span, patch, want) => {
    const given = structuredClone(map)
    const got = patchLines(map, span, patch)
    expect(got).toEqual(want)
    expect(got).toEqual(bySession(map, span, patch))
    expect(map).toEqual(given)
  })

  it('the same function in xlcalc/a1', () => {
    expect(a1.patchLines).toBe(patchLines)
  })

  it('a wrong span, map or patch is an InputError naming it', () => {
    for (const span of ['b:d', 'D:B', 'C:C', '0', 'XFE', '5:1048577', 'B2'])
      expect(() => patchLines({}, span, { width: 1 }), span).toThrow(new InputError('span', `expected columns or rows ("B:D", "5:9"), got the string ${JSON.stringify(span)}`))
    expect(() => patchLines({ 5: { height: 2 } }, 'B', { width: 1 } as never)).toThrow(new InputError('map["5"]', 'expected an object under a span of the axis of "B"'))
    expect(() => patchLines({ B: 7 } as never, 'B', { width: 1 })).toThrow(new InputError('map["B"]', 'expected an object under a span of the axis of "B"'))
    expect(() => patchLines({ 'B:D': {}, C: {} }, 'B', { width: 1 })).toThrow(new InputError('map', 'two spans overlap'))
    expect(() => patchLines([] as never, 'B', { width: 1 })).toThrow(new InputError('map', 'expected lines by span, got an array'))
    expect(() => patchLines({}, 'B', null as never)).toThrow(new InputError('patch', 'expected fields, got null'))
  })
})
