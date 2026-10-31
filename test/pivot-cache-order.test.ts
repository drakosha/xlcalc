import { readFileSync } from 'node:fs'
import { expect, it } from 'vitest'
import { createState } from '../src/engine.js'
import { clockOptions } from '../src/result.js'
import { refreshPivots } from '../src/pivot/engine.js'
import { stateValues } from '../src/structure/values.js'
import { readXlsx } from '../src/xlsx.js'

const root = new URL('./excel/win/pivot/', import.meta.url)
for (const [beforeFile, afterFile, field] of [['sort/refresh-label-sort-new-1', 'sort/refresh-label-sort-new', 0], ['basic/price-sum', 'sort/months', 8]] as const)
  it(`refreshes cache items in source order: ${afterFile}`, async () => {
    const before = (await readXlsx(readFileSync(new URL(`${beforeFile}.xlsx`, root)))).book
    const after = (await readXlsx(readFileSync(new URL(`${afterFile}.xlsx`, root)))).book
    const book = { ...before, sheets: before.sheets.map(s => s.name === 'D' ? after.sheets.find(s => s.name === 'D')! : s) }
    const r = refreshPivots(book, undefined, stateValues(createState(book, clockOptions({ now: '2026-10-30', seed: 0 }))), 0, true)
    expect(r.ok).toBe(true)
    if (r.ok) expect(r.book.pivotCaches![0].fields[field].items).toEqual(after.pivotCaches![0].fields[field].items)
  })
