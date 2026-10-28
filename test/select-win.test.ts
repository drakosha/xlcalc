/**
 * Create Names from Selection against Excel for Windows (`test/excel/win/select/`, `bench/coverage/select/`): the names
 * `getNamesFromSelection` gives are those `Range.CreateNames` made, but where Excel refused the range (`REFUSED`); and
 * the method in every kind of session, its input checked.
 */
import { readFileSync } from 'node:fs'
import { afterAll, describe, expect, it } from 'vitest'
import { InputError } from '../src/index.js'
import { selectedInBook, selectFile, type SelectAnswer, type SelectSet } from '../bench/coverage/select/cases.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'

afterAll(closeCalcs)

/** Ranges Excel refuses ("This selection isn't valid"), where the engine gives no names: labels and nothing past them. */
const REFUSED: Record<string, string> = { 'names:labels-alone': 'a range of labels alone, nothing past them' }

describe.each(['names', 'names-held'])('the set %s', (set) => {
  const stored = JSON.parse(readFileSync(selectFile(set), 'utf8')) as { answers: Record<string, SelectAnswer> }
  it('names what Excel named', async () => {
    const { BOOKS, buildSelectCases } = (await import(`../bench/coverage/select/${set}.ts`)) as SelectSet
    const cases = buildSelectCases()
    expect(Object.keys(stored.answers).sort()).toEqual(cases.map((c) => c.id).sort())
    for (const [k, b] of Object.entries(BOOKS))
      for (const [id, got] of Object.entries(selectedInBook(b, cases.filter((c) => c.book === k)))) {
        const excel = stored.answers[id]
        if (REFUSED[`${set}:${id}`]) expect([excel, got], id).toEqual([{ error: expect.stringContaining("This selection isn't valid") }, { names: [] }])
        else expect(got, id).toEqual(excel)
      }
  })
})

describe.each(KINDS)('getNamesFromSelection, %s', (kind) => {
  it('names the cells past the labels and checks its input', async () => {
    const s = await (await opener(kind))({ sheets: [{ name: 'Data 1', cells: { A1: { value: 'x' }, B1: { value: 'Q1' }, A2: { value: 'north' }, B2: { value: 1 } } }] }, {})
    expect(await s.getNamesFromSelection('Data 1', 'A1:B2', { top: true, left: true })).toEqual([{ name: 'Q1_', ref: "'Data 1'!$B$2" }, { name: 'north', ref: "'Data 1'!$B$2" }, { name: 'x', ref: "'Data 1'!$B$2" }])
    await expect(s.getNamesFromSelection('Data 1', 'A1:B2', {})).rejects.toThrow(new InputError('labels', 'expected at least one side of labels: top, left, bottom or right'))
    await expect(s.getNamesFromSelection('Data 1', 'A1:B2', { up: true } as never)).rejects.toThrow(InputError)
    await expect(s.getNamesFromSelection('Nope', 'A1', { top: true })).rejects.toThrow(InputError)
    expect(s.revision).toBe(0)
  })
})
