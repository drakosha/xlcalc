/**
 * Formulas over functions the engine computes no value for, in files Excel for Windows computed and saved
 * (`bench/coverage/xlsx/stored.ts`, `test/excel/win/xlsx/stored/`): read with the values the file stores, the engine
 * shows them (`unknownFunctions: 'held'`) and computes the cells reading them from them, as Excel shows a file it
 * opens; an edit that reaches such a formula computes it. `STORED` and `BLOCKED`, Excel's values, were written before Excel ran.
 */
import { existsSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { Val } from '../bench/coverage/formulajs.js'
import { compute, createSession, FUNCTION_NAMES, type CellValue } from '../src/index.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { openArchive } from '../src/xlsx/zip/read.js'
import { show, toVal } from './excel-cases.js'

const dir = new URL('./excel/win/xlsx/', import.meta.url)

/** The values Excel was predicted to store, by cell of `outside`. */
const STORED: Record<string, Val> = {
  B1: { t: 'n', v: 20 },
  C1: { t: 'n', v: 21 },
  D1: { t: 'e', v: 'NAME' },
  E1: { t: 's', v: 'none' },
  F1: { t: 'e', v: 'NA' },
  G1: { t: 'e', v: 'VALUE' },
  H1: { t: 'n', v: 3 },
}

/** The values Excel was predicted to store in `blocked`: #BLOCKED! as a rich value of `errorType` 10, ERROR.TYPE 11. */
const BLOCKED: Record<string, CellValue> = { B1: { error: 'BLOCKED' }, C1: { error: 'BLOCKED' }, D1: 11, E1: 11 }

/** The `errorType` of each rich value a file's errors keep. */
async function richTypes(data: Uint8Array): Promise<string[]> {
  const xml = (await openArchive(data).text('xl/richData/rdrichvalue.xml')) ?? ''
  return [...xml.matchAll(/<rv s="\d+"><v>(\d+)<\/v>/g)].map((m) => m[1])
}

describe.skipIf(!existsSync(new URL('stored.json', dir)))('formulas keeping the values a file Excel for Windows saved stores', () => {
  it('outside: the engine shows what Excel stored, as predicted, and computes from it', async () => {
    const r = await readXlsx(readFileSync(new URL('stored/outside.xlsx', dir)))
    const stored = r.cached.sheets[0].cells
    expect(Object.fromEntries(Object.keys(STORED).map((a) => [a, show(toVal(stored[a].value))]))).toEqual(Object.fromEntries(Object.entries(STORED).map(([a, v]) => [a, show(v)])))
    const res = compute(r.book, { cached: r.cached, unknownFunctions: 'held' })
    for (const a of Object.keys(STORED)) expect(res.sheets[0].cells[a].value, a).toEqual(stored[a].value)
    // OFFSET keeps its value while the engine does not know it
    const marked = res.diagnostics.flatMap((d) => (d.kind === 'cell' && d.fromFile ? [d.cell] : []))
    expect(marked).toEqual([...(FUNCTION_NAMES.includes('OFFSET') ? [] : ['B1']), 'D1', 'E1', 'F1', 'G1'])
  })

  it("blocked: #BLOCKED! is stored as a rich value, read as itself and written back as Excel writes it", async () => {
    const data = readFileSync(new URL('stored/blocked.xlsx', dir))
    const r = await readXlsx(data)
    expect(Object.fromEntries(Object.keys(BLOCKED).map((a) => [a, r.cached.sheets[0].cells[a].value]))).toEqual(BLOCKED)
    expect(r.issues).toEqual([])
    const again = await writeXlsx({ book: r.book, result: r.cached })
    expect([again.issues, await richTypes(again.data), await richTypes(data)]).toEqual([[], ['10', '10'], ['10', '10']])
    const back = (await readXlsx(again.data)).cached.sheets[0].cells
    expect(Object.fromEntries(Object.keys(BLOCKED).map((a) => [a, back[a].value]))).toEqual(BLOCKED)
  })

  it("outside: a change that reaches the add-in's formulas computes them (#NAME?, caught by IFERROR); the cube's and the web's keep their values", async () => {
    const r = await readXlsx(readFileSync(new URL('stored/outside.xlsx', dir)))
    const s = createSession(r.book, { cached: r.cached, unknownFunctions: 'held' })
    s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 3 } }])
    expect(['D1', 'E1', 'F1', 'G1', 'H1'].map((a) => s.getCell('S', a)?.value)).toEqual([{ error: 'NAME' }, 'none', { error: 'NA' }, { error: 'VALUE' }, 3])
  })
})
