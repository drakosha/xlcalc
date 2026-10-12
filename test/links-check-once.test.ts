/**
 * A session checks its book's links once, as it opens: its window's formula text, its explanations and its typed text
 * do not check them again for each formula (a book's links can hold thousands of sheets; 0.22.0 checked them, with
 * their stored values, once for each formula a window showed). The public calls check their options once a call.
 */
import { describe, expect, it, vi } from 'vitest'
import * as input from '../src/input.js'
import { createSession, toLocal, type BookInput } from '../src/index.js'

vi.mock('../src/input.js', async (original) => {
  const m = await original<typeof input>()
  return { ...m, checkExternalBooks: vi.fn(m.checkExternalBooks) }
})
const checks = vi.mocked(input.checkExternalBooks)

const externalBooks = Array.from({ length: 20 }, (_, k) => ({ target: `book${k + 1}.xlsx`, sheets: Array.from({ length: 5 }, (_, j) => ({ name: `S${j + 1}`, cells: { A1: j } })) }))
const cells = Object.fromEntries(Array.from({ length: 40 }, (_, i) => [`A${i + 1}`, { formula: `'[${1 + (i % 20)}]S${1 + (i % 5)}'!A1*2` }]))
const book: BookInput = { sheets: [{ name: 'S', cells }], externalBooks }

describe('links checked once', () => {
  it('a session: its window with formula text, an explanation, typed formulas and a paste check no link again', () => {
    const s = createSession(book)
    checks.mockClear()
    s.getViewport('S', 'A1:A40', { formulaLocal: true, editText: true })
    s.explain('S', 'A1')
    s.setCellTexts(Array.from({ length: 10 }, (_, i) => ({ sheet: 'S', cell: `B${i + 1}`, text: `=[book${i + 1}.xlsx]S1!A1+1` })))
    s.pasteText('S', 'C1', '=[book1.xlsx]S2!A1\t=[book2.xlsx]S3!A1')
    expect(checks).not.toHaveBeenCalled()
  })

  it('a public call checks its options once', () => {
    checks.mockClear()
    expect(toLocal("'[1]S1'!A1", 'en-US', { externalBooks }).ok).toBe(true)
    expect(checks).toHaveBeenCalledTimes(1)
  })
})
