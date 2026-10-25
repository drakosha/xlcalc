/**
 * `xlcalc/xlsx/keep` against Excel. `npx tsx bench/coverage/run.ts keep made`: each book of `keep/books.ts` (`MADE`)
 * goes to Excel, which gives it charts, pictures, formatted runs, theme colours and the rest (`excel-keep.applescript`)
 * and saves it into `test/excel/keep/made/<id>.xlsx`. `keep predict` writes what Excel is predicted to do with each
 * made file and each edit of `EDITS` written over it (`keep/open.predicted.json`: it opens the file without its repair
 * window and keeps every part counted), before `keep open` sends them: `writeKept` writes each edit over its made file,
 * Excel opens, computes and saves it into `test/excel/keep/open/<made>-<edit>.xlsx`, and `test/excel/keep/open.json`
 * records what was sent (a hash of its parts), the repair window's text where Excel refused one, and the parts of the
 * file written and of the file Excel saved. `keep struct-predict` writes what the parts refer to after each operation of
 * the structure of `STRUCT` that `writeKept` writes over a made file (`keep/struct.predicted.json`), before `keep struct`
 * has Excel do the same operation on it and records what the parts it saved refer to (`test/excel/keep/struct.json`).
 * Test: `test/xlsx-keep-excel.test.ts`.
 */
import { createHash } from 'node:crypto'
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { homedir } from 'node:os'
import { join } from 'node:path'
import { createSession } from '../../src/index.js'
import { openKept, writeKept } from '../../src/xlsx-keep.js'
import { ensureExcel, excelVersion, runSavingScript } from './excel.js'
import { baseFile, EDITS, MADE, MADE_STRUCT, partsOf, refsOf, STRUCT } from './keep/books.js'
import { saveThroughExcel } from './xlsx/excel.js'
import { partsHash } from './xlsx/sent.js'
import { HERE, ROOT } from './paths.js'
import { loadStored, secondsSince, writeAnswers, writePrediction } from './stored.js'

const DATA = join(ROOT, 'test/excel/keep')
const ONLY = process.env.FKEEP_ONLY?.split(',').filter(Boolean)

/** A file Excel made: the hash of what was sent, what its script answered, the parts its file holds. */
interface MadeRun {
  meta: Record<string, string | number>
  books: Record<string, { sent: string; made: string[]; parts: Record<string, number> }>
}

/** A file written over a made one: the hash of its parts, the parts written, and Excel's refusal or the parts of the file it saved. */
export interface OpenRun {
  meta: Record<string, string | number>
  cases: Record<string, { sent: string; written: Record<string, number>; rejected?: string; saved?: Record<string, number> }>
}

/** A picture Excel can insert: an 8 by 8 PNG in its sandbox. */
function picture(): string {
  const file = join(homedir(), 'Library/Containers/com.microsoft.Excel/Data/tmp/keep-picture.png')
  mkdirSync(join(file, '..'), { recursive: true })
  writeFileSync(file, Buffer.from('89504e470d0a1a0a0000000d49484452000000080000000808020000004b6d29dc0000001549444154789c63f8cfc0f09f8181e13f031a6000005ff10fa1f2b1b6cf0000000049454e44ae426082', 'hex'))
  return file
}

async function made(): Promise<void> {
  const file = join(DATA, 'made.json')
  const run = loadStored<MadeRun>(file) ?? { meta: {}, books: {} }
  mkdirSync(join(DATA, 'made'), { recursive: true })
  const t0 = Date.now()
  await ensureExcel('en-US')
  const png = picture()
  for (const m of [...MADE, ...MADE_STRUCT]) {
    if (ONLY && !ONLY.includes(m.id)) continue
    const base = await baseFile(m)
    const sent = createHash('sha256').update(JSON.stringify(m)).update(await partsHash(base)).digest('hex')
    if (run.books[m.id]?.sent === sent && existsSync(join(DATA, 'made', `${m.id}.xlsx`))) continue
    const r = await runSavingScript(`made-${m.id}`, join(HERE, 'excel-keep.applescript'), base, [m.features.join(','), png])
    if ('rejected' in r) throw new Error(`${m.id}: Excel refused the book it was to make: ${r.rejected}`)
    writeFileSync(join(DATA, 'made', `${m.id}.xlsx`), r.saved)
    run.books[m.id] = { sent, made: r.out.split('\n').filter(Boolean), parts: await partsOf(r.saved) }
    console.log(`${m.id}: ${run.books[m.id].made.join(', ')}; ${JSON.stringify(run.books[m.id].parts)}`)
    run.meta = { version: excelVersion(), seconds: secondsSince(t0), date: new Date().toISOString().slice(0, 10) }
    writeAnswers(file, run)
  }
}

/** Each edit written over each made file: its id, the file written. */
export async function written(): Promise<{ id: string; data: Uint8Array }[]> {
  const out: { id: string; data: Uint8Array }[] = []
  for (const m of MADE) {
    const original = new Uint8Array(readFileSync(join(DATA, 'made', `${m.id}.xlsx`)))
    for (const e of EDITS) {
      const read = await openKept(original)
      const s = createSession(read.book, { cached: read.cached })
      e.edit(s)
      out.push({ id: `${m.id}-${e.id}`, data: (await writeKept(original, s.getSnapshot())).data })
      s.close()
    }
  }
  return out
}

/** The prediction, written before `open` runs: Excel opens each file written without its repair window and keeps every part the made file has, but its pivot table's cache and notes as the edit leaves them. */
async function predict(): Promise<void> {
  const run = loadStored<MadeRun>(join(DATA, 'made.json'))
  if (!run) throw new Error('no made files: run keep made first')
  const cases: Record<string, { opens: true; parts: Record<string, number> }> = {}
  for (const w of await written()) {
    const parts = await partsOf(w.data)
    cases[w.id] = { opens: true, parts }
  }
  writePrediction(new URL('./keep/books.ts', import.meta.url).href, 'open', { cases })
}

async function open(): Promise<void> {
  const file = join(DATA, 'open.json')
  const run = loadStored<OpenRun>(file) ?? { meta: {}, cases: {} }
  mkdirSync(join(DATA, 'open'), { recursive: true })
  const t0 = Date.now()
  await ensureExcel('en-US')
  for (const w of await written()) {
    if (ONLY && !ONLY.includes(w.id)) continue
    const sent = await partsHash(w.data)
    if (run.cases[w.id]?.sent === sent && (run.cases[w.id].rejected || existsSync(join(DATA, 'open', `${w.id}.xlsx`)))) continue
    const r = await saveThroughExcel(w.data, w.id)
    const parts = await partsOf(w.data)
    if ('rejected' in r) {
      run.cases[w.id] = { sent, written: parts, rejected: r.rejected }
      console.log(`${w.id}: Excel refused the file: ${r.rejected}`)
    } else {
      writeFileSync(join(DATA, 'open', `${w.id}.xlsx`), r.saved)
      run.cases[w.id] = { sent, written: parts, saved: await partsOf(r.saved) }
      console.log(`${w.id}: opened; written ${JSON.stringify(parts)}, saved ${JSON.stringify(run.cases[w.id].saved)}`)
    }
    run.meta = { version: excelVersion(), seconds: secondsSince(t0), date: new Date().toISOString().slice(0, 10) }
    writeAnswers(file, run)
  }
}

/** A made file with an operation of the structure: Excel's answer, what it refers to, and the hash of what was sent. */
export interface StructRun {
  meta: Record<string, string | number>
  cases: Record<string, { sent: string; out: string; refs?: Record<string, unknown> }>
}

/** Each operation of `STRUCT` written over its made file by `writeKept`: what its parts refer to then, or the refusal. */
export async function structWritten(): Promise<Record<string, { refs: Record<string, unknown> } | { refused: string }>> {
  const out: Record<string, { refs: Record<string, unknown> } | { refused: string }> = {}
  for (const c of STRUCT) {
    const original = new Uint8Array(readFileSync(join(DATA, 'made', `${c.made}.xlsx`)))
    const read = await openKept(original)
    const s = createSession(read.book, { cached: read.cached })
    const r = s.editStructure(c.op)
    if (r.ok === false) throw new Error(`${c.id}: the session refused the operation: ${r.message}`)
    try {
      out[c.id] = { refs: await refsOf((await writeKept(original, s.getSnapshot())).data) }
    } catch (e) {
      if ((e as Error).name !== 'KeptRefusal') throw e
      out[c.id] = { refused: (e as Error).message }
    }
    s.close()
  }
  return out
}

/** The operations of the structure in Excel over the made files: what the parts Excel saved refer to. */
async function struct(): Promise<void> {
  const file = join(DATA, 'struct.json')
  const run = loadStored<StructRun>(file) ?? { meta: {}, cases: {} }
  mkdirSync(join(DATA, 'struct'), { recursive: true })
  const t0 = Date.now()
  await ensureExcel('en-US')
  for (const c of STRUCT) {
    if (ONLY && !ONLY.includes(c.id)) continue
    const data = new Uint8Array(readFileSync(join(DATA, 'made', `${c.made}.xlsx`)))
    const sent = createHash('sha256').update(JSON.stringify(c.excel)).update(await partsHash(data)).digest('hex')
    if (run.cases[c.id]?.sent === sent) continue
    const r = await runSavingScript(`struct-${c.id}`, join(HERE, 'excel-keep-struct.applescript'), data, c.excel)
    if ('rejected' in r) throw new Error(`${c.id}: Excel refused the made file: ${r.rejected}`)
    writeFileSync(join(DATA, 'struct', `${c.id}.xlsx`), r.saved)
    run.cases[c.id] = { sent, out: r.out.trim(), refs: await refsOf(r.saved) }
    console.log(`${c.id}: ${r.out.trim()} ${JSON.stringify(run.cases[c.id].refs)}`)
    run.meta = { version: excelVersion(), seconds: secondsSince(t0), date: new Date().toISOString().slice(0, 10) }
    writeAnswers(file, run)
  }
}

export async function stepKeep(what: string | undefined): Promise<void> {
  if (what === 'made') return made()
  if (what === 'predict') return predict()
  if (what === 'open') return open()
  if (what === 'struct-predict') return writePrediction(new URL('./keep/books.ts', import.meta.url).href, 'struct', { cases: await structWritten() })
  if (what === 'struct') return struct()
  throw new Error('keep: made | predict | open | struct-predict | struct')
}
