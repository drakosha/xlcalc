/**
 * Formulas that begin with `#REF!` against Excel for Mac and Windows (`bench/coverage/errstore`): both compute the book
 * as it opens, `#REF!-B1` as `#REF!` whatever value the file stores (`#VALUE!`, 7), and deleting what a formula reads
 * makes it `#REF!-…` with the value `#REF!`, which the file Excel saves holds. The engine computes the file as Excel opened it, and its
 * structure operations write the formulas Excel for Windows's object model shows.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { bookFile } from '../bench/coverage/errstore/cases.js'
import { partsHash } from '../bench/coverage/xlsx/sent.js'
import { compute, editStructure } from '../src/index.js'
import { readXlsx } from '../src/xlsx.js'

const json = (path: string) => JSON.parse(readFileSync(new URL(path, import.meta.url), 'utf8'))
const mac = json('./excel/errstore/mac.json')
const win = json('./excel/win/errstore/win.json')
/** Excel for Windows's `Value2` errors, by their HRESULT. */
const WIN_ERROR: Record<number, string> = { [-2146826273]: '#VALUE!', [-2146826265]: '#REF!', [-2146826246]: '#N/A' }
const winShown = (v: unknown) => (typeof v === 'object' && v && 'e' in v ? WIN_ERROR[(v as { e: number }).e] : String(v))
const rows = (n: number, col: string) => Array.from({ length: n }, (_, i) => `${col}${i + 1}`)

describe('formulas that begin with #REF!', async () => {
  const data = await bookFile()
  const { book } = await readXlsx(data)
  const shown = (b: typeof book, cells: string[]) => {
    const got = compute(b).sheets[0].cells
    return cells.map((c) => got[c].text)
  }

  it('the files sent are the ones Excel read', async () => {
    expect(mac.sent).toBe(await partsHash(data))
    expect(win.sent).toBe(await partsHash(data))
  })

  it('the engine computes the book as both Excels opened it, not the values stored (the controls C11 and C12 among them)', () => {
    const engine = shown(book, rows(12, 'C'))
    expect(engine).toEqual(mac.opened.map((x: string[]) => x[1]))
    expect(engine).toEqual(win.opening[1].map((r: unknown[]) => winShown(r[0])))
  })

  it("deleting column A and the sheet T gives Excel's formulas and the values it saved", () => {
    const b = editStructure(book, { op: 'deleteColumns', sheet: 'S', column: 'A' })
    if (!b.ok) throw new Error(JSON.stringify(b))
    const t = editStructure(b.book, { op: 'deleteSheet', sheet: 'T' })
    if (!t.ok) throw new Error(JSON.stringify(t))
    const cells = t.book.sheets[0].cells
    const saved = new Map<string, string>(mac.saved.map((c: string) => [/r="(\w+)"/.exec(c)![1], /<v>(.*)<\/v>/.exec(c)![1]]))
    expect(win.saved).toEqual(mac.saved)
    rows(12, 'B').forEach((c, i) => {
      // Excel for Windows's object model keeps a reference to the deleted sheet as #REF!A1, as the engine does; the file both save holds #REF!
      expect([c, `=${cells[c]?.formula}`, shown(t.book, [c])[0]]).toEqual([c, win.deleted[0][i][0], saved.get(c)])
      expect(mac.full[i][1]).toBe(saved.get(c))
    })
  })
})
