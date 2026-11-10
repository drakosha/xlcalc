/**
 * `{ autoFilter, showRows: true }`, an AutoFilter removed or moved as Excel's Filter button does, against what Excel for Mac and
 * for Windows did to the same rows (`bench/coverage/filter-acts/`). A filter's criteria are not in the model, so a case's sheet
 * starts as Excel left it after them. Its own range is the file's, or what Excel read after applying criteria: a host sets that,
 * but for a last row of SUBTOTAL, which Excel reads and does not filter.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { ActCase } from '../bench/coverage/filter-acts/cases.js'
import { buildActCases as moved } from '../bench/coverage/filter-acts/moved.js'
import { buildActCases as held } from '../bench/coverage/filter-acts/removal-held.js'
import { buildActCases as removal } from '../bench/coverage/filter-acts/removal.js'
import { createSession, type BookInput, type RowInput } from '../src/index.js'

const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
interface Seen {
  range: string
  filterMode: boolean
  hidden: number[]
  total: { formula: string } | null
}
type Answers = Record<string, { opened: Seen; after: Seen[] }>
const answers = (dir: string, set: string) => (JSON.parse(readFileSync(new URL(`./excel/${dir}filter-acts/${set}.json`, import.meta.url), 'utf8')) as { answers: Answers }).answers

/** The sheet as Excel left it just before the filter was removed or moved, and the rows hidden then. */
function before(c: ActCase, seen: Seen[]): BookInput {
  const state = seen.at(-2)!
  const rows: Record<string, RowInput> = {}
  for (const r of state.hidden) rows[String(r)] = c.zero?.includes(r) ? { height: 0 } : { hidden: true }
  const applied = c.steps.some((x) => 'criteria' in x) && !state.total?.formula.startsWith('=SUBTOTAL')
  const range = applied ? state.range : c.filter!
  return { sheets: [{ name: 'S', cells: c.cells, rows, autoFilter: { range }, ...(state.filterMode ? { filtered: true } : {}) }] }
}

/** The rows a session's sheet does not show. */
function notShown(rows: Record<string, RowInput> | undefined): number[] {
  return Object.entries(rows ?? {})
    .filter(([, x]) => x.hidden || x.height === 0)
    .map(([r]) => Number(r))
    .sort((a, b) => a - b)
}

describe.each([
  ['Excel for Mac', ''],
  ['Excel for Windows', 'win/'],
])('%s', (_, dir) => {
  const REMOVED = [
    ...removal().map((c) => ['removal', c] as const),
    ...held().map((c) => ['removal-held', c] as const),
  ]
  it.each(REMOVED.map(([set, c]) => [`${set} ${c.id}`, set, c] as const))('removed, %s', (_, set, c) => {
    const seen = answers(dir, set)[c.id]
    const stages = [seen.opened, ...seen.after]
    const s = createSession(before(c, stages), { clock })
    s.setLayout([{ sheet: 'S', autoFilter: null, showRows: true }])
    const sheet = s.getBook().sheets[0]
    expect(notShown(sheet.rows)).toEqual(stages.at(-1)!.hidden)
    expect(sheet.filtered).toBeUndefined()
    expect(sheet.autoFilter).toBeUndefined()
    s.close()
  })

  it.each(moved().filter((c) => c.id.endsWith('-off-model') || c.id.endsWith('-off-button')).map((c) => [c.id, c] as const))('moved, %s', (_, c) => {
    const seen = answers(dir, 'moved')[c.id]
    const stages = [seen.opened, ...seen.after]
    const criteria = stages.find((x) => x.filterMode)!
    const s = createSession(before(c, [criteria, criteria]), { clock })
    const range = c.steps.flatMap((x) => ('over' in x ? [x.over] : []))[0]!
    s.setLayout([{ sheet: 'S', autoFilter: { range }, showRows: true }])
    const sheet = s.getBook().sheets[0]
    expect(notShown(sheet.rows)).toEqual(stages.at(-1)!.hidden)
    expect(sheet.autoFilter).toEqual({ range })
    expect(sheet.filtered).toBeUndefined()
    s.close()
  })
})

describe('showRows', () => {
  const book = (): BookInput => ({
    sheets: [{ name: 'S', cells: { A1: { value: 'h' }, A2: { value: 1 }, A3: { value: 2 }, B2: { formula: 'SUBTOTAL(9,A2:A3)' }, A9: { value: 3 } }, rows: { '3': { hidden: true }, '9': { hidden: true } }, autoFilter: { range: 'A1:A3' }, filtered: true }],
  })

  it('is one undo step, and recalculates what the rows change', () => {
    const s = createSession(book(), { clock })
    expect(s.getCell('S', 'B2')?.value).toBe(1)
    s.setLayout([{ sheet: 'S', autoFilter: null, showRows: true }])
    expect(notShown(s.getBook().sheets[0].rows)).toEqual([9])
    expect(s.getCell('S', 'B2')?.value).toBe(3)
    s.undo()
    expect(s.getBook().sheets[0]).toMatchObject({ autoFilter: { range: 'A1:A3' }, filtered: true })
    expect(notShown(s.getBook().sheets[0].rows)).toEqual([3, 9])
    expect(s.getCell('S', 'B2')?.value).toBe(1)
    s.close()
  })

  it('sees the edits before it in the same call', () => {
    const s = createSession(book(), { clock })
    s.setLayout([{ sheet: 'S', autoFilter: { range: 'A1:A9' } }, { sheet: 'S', autoFilter: null, showRows: true }])
    expect(notShown(s.getBook().sheets[0].rows)).toEqual([])
    s.close()
  })

  it('without it, the rows and the filter mode stay; on a sheet not filtered, it moves no row', () => {
    const s = createSession(book(), { clock })
    s.setLayout([{ sheet: 'S', autoFilter: null }])
    expect(s.getBook().sheets[0].filtered).toBe(true)
    expect(notShown(s.getBook().sheets[0].rows)).toEqual([3, 9])
    s.setLayout([{ sheet: 'S', filtered: false }, { sheet: 'S', autoFilter: { range: 'A1:A3' }, showRows: true }])
    expect(notShown(s.getBook().sheets[0].rows)).toEqual([3, 9])
    s.close()
  })

  it.each([
    ['A1', [2]],
    ['A1:B1', [2]],
    ['A1:A2', []],
  ])('over %s, a header alone or one row of data, shows its rows and no other', (range, hidden) => {
    const s = createSession({ sheets: [{ name: 'S', cells: { A1: { value: 'h' }, A2: { value: 1 } }, rows: { '2': { hidden: true } }, autoFilter: { range }, filtered: true }] }, { clock })
    s.setLayout([{ sheet: 'S', autoFilter: null, showRows: true }])
    const sheet = s.getBook().sheets[0]
    expect(notShown(sheet.rows)).toEqual(hidden)
    expect(sheet.filtered).toBeUndefined()
    expect(sheet.autoFilter).toBeUndefined()
    s.close()
  })

  it('takes true alone, and only with an AutoFilter', () => {
    const s = createSession(book(), { clock })
    expect(() => s.setLayout([{ sheet: 'S', autoFilter: null, showRows: false } as never])).toThrow(/showRows/)
    expect(() => s.setLayout([{ sheet: 'S', frozen: null, showRows: true } as never])).toThrow(/showRows/)
    s.close()
  })
})
