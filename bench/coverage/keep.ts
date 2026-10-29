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
 * `keep struct-open-predict` writes what Excel is predicted to do with each of those files `writeKept` wrote
 * (`keep/struct-open.predicted.json`: it opens it without its repair window, keeps every part counted and what they
 * refer to), before `keep struct-open` has Excel open, compute and save them (`test/excel/keep/struct-open.json`,
 * `struct-open/<case>.xlsx`). `keep made` also asks Excel what it does with an AutoFilter beside a table (`MADE_PROBES`),
 * its answer in `made.json` beside the engine's committed before it ran (`PROBES`). `keep formats-predict` writes the
 * cell formats `writeKept` gives the cells, rows and columns of the made file `formats` after each operation of `FORMATS`
 * (`keep/formats.predicted.json`), before `keep formats` has Excel do each operation and open the file `writeKept` wrote
 * for it, and records what both files Excel saved give them (`test/excel/keep/formats.json`, `formats/<case>.xlsx`,
 * `formats-open/<case>.xlsx`). Test: `test/xlsx-keep-excel.test.ts`.
 * The conditional formats the model does not read (`keep/rules.ts`): `keep rules-made-predict` and `keep rules-predict`
 * write what Excel for Windows is predicted to save (`bench/windows/run.ts keep rules-made|rules`; `rules-held-predict`
 * and `rules-held-2-predict` for the held-out sets, `keep rules-held`, `keep rules-held-2`), `keep
 * rules-open-predict` and `keep rules-open` what Excel for Mac does with the files `writeKept` wrote for the cases
 * (`test/excel/keep/rules-open.json`, `rules-open/<case>.xlsx`). Test: `test/xlsx-keep-rules-excel.test.ts`.
 * `keep win-predict` writes what Excel for Windows is predicted to save of the files of `open`, `struct-open` and
 * `formats` from the Mac's answers (`keep/win.predicted.json`), before `bench/windows/run.ts keep <set>` runs them.
 * `keep looks-predict` and `keep looks`: what Excel saves of a file `writeXlsx` wrote with a look over a whole column or
 * row, a cell for each of its cells (`LOOKS`, `test/excel/keep/looks.json`).
 */
import { execFileSync } from 'node:child_process'
import { createHash } from 'node:crypto'
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { homedir } from 'node:os'
import { join } from 'node:path'
import { createSession } from '../../src/index.js'
import { openKept, writeKept } from '../../src/xlsx-keep.js'
import { writeXlsx } from '../../src/xlsx.js'
import { openArchive } from '../../src/xlsx/zip/read.js'
import type { BookInput } from '../../src/contract.js'
import { dialogText, ensureExcel, excelVersion, runSavingScript, runWatched, settleRejected } from './excel.js'
import { baseFile, EDITS, MADE, MADE_PROBES, MADE_STRUCT, partsOf, refsOf, STRUCT } from './keep/books.js'
import { RULES_BOOKS, RULES_SETS, rulesFile, rulesOf, rulesWritten } from './keep/rules.js'
import { FORMATS, formatsOf, MADE_FORMATS } from './keep/formats.js'
import { saveThroughExcel } from './xlsx/excel.js'
import { partsHash } from './xlsx/sent.js'
import { HERE, ROOT } from './paths.js'
import { loadStored, secondsSince, writeAnswers, writePrediction } from './stored.js'

const DATA = join(ROOT, 'test/excel/keep')
const ONLY = process.env.FKEEP_ONLY?.split(',').filter(Boolean)

/** A file Excel made: the hash of what was sent, what its script answered, the parts its file holds. */
export interface MadeRun {
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
  for (const m of [...MADE, ...MADE_STRUCT, ...MADE_PROBES, ...MADE_FORMATS]) {
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
      if (e.made && !e.made.includes(m.id)) continue
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

/** A case of `STRUCT` written by `writeKept` and opened by Excel: the hash of the parts sent, Excel's repair window or the parts of the file it saved and what they refer to. */
export interface StructOpenRun {
  meta: Record<string, string | number>
  cases: Record<string, { sent: string; rejected?: string; saved?: Record<string, number>; refs?: Record<string, unknown> }>
}

/** Each operation of `STRUCT` written over its made file by `writeKept`: the file, or the refusal. */
export async function structFiles(): Promise<{ id: string; written: { data: Uint8Array } | { refused: string } }[]> {
  const out: { id: string; written: { data: Uint8Array } | { refused: string } }[] = []
  for (const c of STRUCT) {
    const original = new Uint8Array(readFileSync(join(DATA, 'made', `${c.made}.xlsx`)))
    const read = await openKept(original)
    const s = createSession(read.book, { cached: read.cached })
    const r = s.editStructure(c.op)
    if (r.ok === false) throw new Error(`${c.id}: the session refused the operation: ${r.message}`)
    try {
      out.push({ id: c.id, written: { data: (await writeKept(original, s.getSnapshot())).data } })
    } catch (e) {
      if ((e as Error).name !== 'KeptRefusal') throw e
      out.push({ id: c.id, written: { refused: (e as Error).message } })
    }
    s.close()
  }
  return out
}

/** Each operation of `STRUCT` written over its made file by `writeKept`: what its parts refer to then, or the refusal. */
export async function structWritten(): Promise<Record<string, { refs: Record<string, unknown> } | { refused: string }>> {
  const out: Record<string, { refs: Record<string, unknown> } | { refused: string }> = {}
  for (const f of await structFiles()) out[f.id] = 'data' in f.written ? { refs: await refsOf(f.written.data) } : f.written
  return out
}

/** Excel's warning that a formula refers to cells it lost, which deleting a chart's data shows once its book is saved. */
const LOST_REFERENCES = /problem with one or more formula references/i

/**
 * A file through an AppleScript that saves it (`excel-keep-struct.applescript`, `excel-save.applescript`): its answer
 * and the file it saved, or the repair window's text, the book closed unsaved. Excel's warning of lost references is
 * answered OK and its text goes with the answer; any other window stops the run.
 */
async function throughExcel(id: string, script: string, data: Uint8Array, args: string[]): Promise<{ out: string; saved: Uint8Array } | { rejected: string }> {
  const r = await runWatched<{ out: string; saved: Uint8Array } | { rejected: string }>({
    script: join(HERE, script),
    label: id,
    folder: 'fkeep',
    book: (file) => writeFileSync(file, data),
    args: (dir) => [join(dir, 'saved.xlsx'), 'saved.xlsx', ...args],
    timeout: 180_000,
    poll: 1000,
    fatal: false,
    repair: async (stop, text) => {
      await stop()
      await settleRejected()
      return { rejected: text }
    },
    done: (out, dir) => ({ out: out.trim(), saved: new Uint8Array(readFileSync(join(dir, 'saved.xlsx'))) }),
  })
  const shown = dialogText()
  if (!shown || 'rejected' in r) return r
  if (!LOST_REFERENCES.test(shown)) throw new Error(`${id}: Excel window after the case: ${shown}`)
  execFileSync('osascript', ['-e', 'tell application "System Events" to tell process "Microsoft Excel"\nrepeat with w in windows\ntry\nif subrole of w is "AXDialog" then\nrepeat with b in (buttons of w)\nif name of b is "OK" then\nclick b\nreturn\nend if\nend repeat\nend if\nend try\nend repeat\nend tell'])
  return { ...r, out: `${r.out}; Excel warned: ${shown.replace(/^AXDialog «|»;.*$/g, '')}` }
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
    const r = await throughExcel(c.id, 'excel-keep-struct.applescript', data, c.excel)
    if ('rejected' in r) throw new Error(`${c.id}: Excel refused the made file: ${r.rejected}`)
    writeFileSync(join(DATA, 'struct', `${c.id}.xlsx`), r.saved)
    run.cases[c.id] = { sent, out: r.out, refs: await refsOf(r.saved) }
    console.log(`${c.id}: ${r.out.trim()} ${JSON.stringify(run.cases[c.id].refs)}`)
    run.meta = { version: excelVersion(), seconds: secondsSince(t0), date: new Date().toISOString().slice(0, 10) }
    writeAnswers(file, run)
  }
}

/** The prediction, written before `struct-open` runs: Excel opens each file `writeKept` wrote for `STRUCT` without its repair window, and keeps its parts and what they refer to. */
async function structOpenPredict(): Promise<void> {
  const cases: Record<string, { opens: true; parts: Record<string, number>; refs: Record<string, unknown> }> = {}
  for (const f of await structFiles()) if ('data' in f.written) cases[f.id] = { opens: true, parts: await partsOf(f.written.data), refs: await refsOf(f.written.data) }
  writePrediction(new URL('./keep/books.ts', import.meta.url).href, 'struct-open', { cases })
}

/** The files `writeKept` wrote for `STRUCT` through Excel: opened, computed and saved. */
async function structOpen(): Promise<void> {
  const file = join(DATA, 'struct-open.json')
  const run = loadStored<StructOpenRun>(file) ?? { meta: {}, cases: {} }
  mkdirSync(join(DATA, 'struct-open'), { recursive: true })
  const t0 = Date.now()
  await ensureExcel('en-US')
  for (const f of await structFiles()) {
    if ((ONLY && !ONLY.includes(f.id)) || !('data' in f.written)) continue
    const sent = await partsHash(f.written.data)
    if (run.cases[f.id]?.sent === sent && (run.cases[f.id].rejected || existsSync(join(DATA, 'struct-open', `${f.id}.xlsx`)))) continue
    const r = await throughExcel(f.id, 'excel-save.applescript', f.written.data, [])
    if ('rejected' in r) {
      run.cases[f.id] = { sent, rejected: r.rejected }
      console.log(`${f.id}: Excel refused the file: ${r.rejected}`)
    } else {
      writeFileSync(join(DATA, 'struct-open', `${f.id}.xlsx`), r.saved)
      run.cases[f.id] = { sent, saved: await partsOf(r.saved), refs: await refsOf(r.saved) }
      console.log(`${f.id}: opened${r.out ? ` (${r.out})` : ''}; saved ${JSON.stringify(run.cases[f.id].saved)} ${JSON.stringify(run.cases[f.id].refs)}`)
    }
    run.meta = { version: excelVersion(), seconds: secondsSince(t0), date: new Date().toISOString().slice(0, 10) }
    writeAnswers(file, run)
  }
}

/** The saved `mixed` of Excel for Windows (`bench/windows/run.ts keep rules-made`), the file the rules' cases are written over. */
const rulesMixed = () => new Uint8Array(readFileSync(join(ROOT, 'test/excel/win/keep/rules-made/mixed.xlsx')))

/** Each case of `RULES_SETS` (or of `cases`) written by `writeKept` over the saved `mixed`: the file, or the refusal. */
export async function rulesFiles(cases = Object.values(RULES_SETS).flat()): Promise<{ id: string; written: { data: Uint8Array } | { refused: string } }[]> {
  const data = rulesMixed()
  return Promise.all(cases.map(async (c) => ({ id: c.id, written: await rulesWritten(data, c) })))
}

/**
 * The predictions of the rules' steps, written before each runs: Excel for Windows opens each book of `RULES_BOOKS`
 * and saves its rules as they are, priorities counted from 1 (`rules-made`); after each case Excel saves the rules
 * `writeKept` writes for the same case, refused or not (`rules`); Excel for Mac opens each file `writeKept` wrote
 * without its repair window and saves its rules as written (`rules-open`).
 */
async function rulesPredict(what: string): Promise<void> {
  const cases: Record<string, unknown> = {}
  if (what === 'rules-made') for (const id of Object.keys(RULES_BOOKS)) cases[id] = { opens: true, ...(await rulesOf(await rulesFile(id))) }
  else for (const f of await rulesFiles(RULES_SETS[what])) cases[f.id] = 'data' in f.written ? { ...(what === 'rules-open' ? { opens: true } : {}), ...(await rulesOf(f.written.data)) } : f.written
  writePrediction(new URL('./keep/rules.ts', import.meta.url).href, what, { cases })
}

/** The files `writeKept` wrote for the rules' cases through Excel for Mac: opened, computed and saved (`test/excel/keep/rules-open.json`). */
async function rulesOpen(): Promise<void> {
  const file = join(DATA, 'rules-open.json')
  const run = loadStored<StructOpenRun>(file) ?? { meta: {}, cases: {} }
  mkdirSync(join(DATA, 'rules-open'), { recursive: true })
  const t0 = Date.now()
  await ensureExcel('en-US')
  for (const f of await rulesFiles()) {
    if ((ONLY && !ONLY.includes(f.id)) || !('data' in f.written)) continue
    const sent = await partsHash(f.written.data)
    if (run.cases[f.id]?.sent === sent && (run.cases[f.id].rejected || existsSync(join(DATA, 'rules-open', `${f.id}.xlsx`)))) continue
    const r = await saveThroughExcel(f.written.data, f.id)
    if ('rejected' in r) {
      run.cases[f.id] = { sent, rejected: r.rejected }
      console.log(`${f.id}: Excel refused the file: ${r.rejected}`)
    } else {
      writeFileSync(join(DATA, 'rules-open', `${f.id}.xlsx`), r.saved)
      run.cases[f.id] = { sent, saved: await partsOf(r.saved), refs: await rulesOf(r.saved) }
      console.log(`${f.id}: opened; ${JSON.stringify(run.cases[f.id].refs)}`)
    }
    run.meta = { version: excelVersion(), seconds: secondsSince(t0), date: new Date().toISOString().slice(0, 10) }
    writeAnswers(file, run)
  }
}

/** A case of `FORMATS`: what Excel did and what its file gives, and whether it opened the file `writeKept` wrote and what the file it saved gives. */
export interface FormatsRun {
  meta: Record<string, string | number>
  cases: Record<string, { sent: string; out: string; formats: Awaited<ReturnType<typeof formatsOf>>; open: { sent: string; rejected?: string; formats?: Awaited<ReturnType<typeof formatsOf>> } }>
}

/** Each operation of `FORMATS` written over the made file by `writeKept`. */
export async function formatsFiles(): Promise<{ id: string; data: Uint8Array }[]> {
  const original = new Uint8Array(readFileSync(join(DATA, 'made', 'formats.xlsx')))
  const out: { id: string; data: Uint8Array }[] = []
  for (const c of FORMATS) {
    const read = await openKept(original)
    const s = createSession(read.book, { cached: read.cached })
    const r = s.editStructure(c.op)
    if (r.ok === false) throw new Error(`${c.id}: the session refused the operation: ${r.message}`)
    out.push({ id: c.id, data: (await writeKept(original, s.getSnapshot())).data })
    s.close()
  }
  return out
}

/** The prediction, written before `formats` runs: Excel gives the cells, rows and columns the formats `writeKept` gives them, opens its file and keeps them. */
async function formatsPredict(): Promise<void> {
  const cases: Record<string, { formats: Awaited<ReturnType<typeof formatsOf>>; opens: true }> = {}
  for (const f of await formatsFiles()) cases[f.id] = { formats: await formatsOf(f.data), opens: true }
  writePrediction(new URL('./keep/formats.ts', import.meta.url).href, 'formats', { cases })
}

/** The operations of `FORMATS` in Excel over the made file, and the files `writeKept` wrote for them opened and saved by Excel. */
async function formats(): Promise<void> {
  const file = join(DATA, 'formats.json')
  const run = loadStored<FormatsRun>(file) ?? { meta: {}, cases: {} }
  for (const d of ['formats', 'formats-open']) mkdirSync(join(DATA, d), { recursive: true })
  const t0 = Date.now()
  await ensureExcel('en-US')
  const data = new Uint8Array(readFileSync(join(DATA, 'made', 'formats.xlsx')))
  const ours = new Map((await formatsFiles()).map((f) => [f.id, f.data]))
  for (const c of FORMATS) {
    if (ONLY && !ONLY.includes(c.id)) continue
    const sent = createHash('sha256').update(JSON.stringify(c.excel)).update(await partsHash(data)).digest('hex')
    const written = ours.get(c.id)!
    const openSent = await partsHash(written)
    if (run.cases[c.id]?.sent === sent && run.cases[c.id].open.sent === openSent) continue
    const r = await throughExcel(c.id, 'excel-keep-struct.applescript', data, c.excel)
    if ('rejected' in r) throw new Error(`${c.id}: Excel refused the made file: ${r.rejected}`)
    writeFileSync(join(DATA, 'formats', `${c.id}.xlsx`), r.saved)
    const o = await throughExcel(`${c.id}-open`, 'excel-save.applescript', written, [])
    if (!('rejected' in o)) writeFileSync(join(DATA, 'formats-open', `${c.id}.xlsx`), o.saved)
    run.cases[c.id] = { sent, out: r.out, formats: await formatsOf(r.saved), open: { sent: openSent, ...('rejected' in o ? { rejected: o.rejected } : { formats: await formatsOf(o.saved) }) } }
    console.log(`${c.id}: ${r.out}; ${JSON.stringify(run.cases[c.id].formats.at)}; ${'rejected' in o ? `refused: ${o.rejected}` : 'opened'}`)
    run.meta = { version: excelVersion(), seconds: secondsSince(t0), date: new Date().toISOString().slice(0, 10) }
    writeAnswers(file, run)
  }
}

/**
 * The prediction of Excel for Windows over the files `writeKept` wrote for the Mac's sets (`bench/windows/keep.ts`),
 * written before it runs: it opens each without repairing it and saves what Excel for Mac saved, the parts counted and
 * their references (`open`, `struct-open`) or the cell formats (`formats-open`).
 */
async function winPredict(): Promise<void> {
  const answers = <T>(name: string) => JSON.parse(readFileSync(join(DATA, `${name}.json`), 'utf8')) as T
  const open = answers<OpenRun>('open').cases
  const struct = answers<StructOpenRun>('struct-open').cases
  const formats = answers<FormatsRun>('formats').cases
  const sets = {
    open: Object.fromEntries((await written()).map((f) => [f.id, { opens: true, parts: open[f.id].saved }])),
    'struct-open': Object.fromEntries((await structFiles()).flatMap((f) => ('data' in f.written ? [[f.id, { opens: true, parts: struct[f.id].saved, refs: struct[f.id].refs }]] : []))),
    'formats-open': Object.fromEntries((await formatsFiles()).map((f) => [f.id, { opens: true, formats: formats[f.id].open.formats }])),
  }
  writePrediction(new URL('./keep/books.ts', import.meta.url).href, 'win', { sets })
}

/** Books with a look over a whole column and over a whole row, which `writeXlsx` writes as a cell for each of their cells. */
export const LOOKS: Record<string, BookInput> = {
  column: { styles: [{ font: { bold: true } }], sheets: [{ name: 'S', cells: { A1: { value: 1 } }, looks: [{ range: 'C:C', style: 0 }] }] },
  row: { styles: [{ font: { bold: true } }], sheets: [{ name: 'S', cells: { A1: { value: 1 } }, looks: [{ range: '3:3', style: 0 }] }] },
}

/** What a file's first sheet holds of cells and of formatted lines: its cells, its columns and rows with a cell format of their own. */
export async function looksOf(data: Uint8Array): Promise<{ cells: number; cols: string[]; rows: string[] }> {
  const xml = (await openArchive(data).text('xl/worksheets/sheet1.xml')) ?? ''
  const cols = [...xml.matchAll(/<col\b[^>]*\bstyle="\d+"[^>]*>/g)].map((m) => m[0])
  const rows = [...xml.matchAll(/<row\b[^>]*\bcustomFormat="(?:1|true)"[^>]*>/g)].map((m) => m[0].replace(/\s(?:spans|x14ac:dyDescent)="[^"]*"/g, ''))
  return { cells: xml.split(/<c[\s>]/).length - 1, cols, rows }
}

/** A book of `LOOKS` written, and what its sheet holds. */
const looksFile = async (id: string) => (await writeXlsx({ book: LOOKS[id] })).data

/** The prediction, written before `looks` runs: Excel opens each file without its repair window and saves the cells as they were written, a cell format on each, and no formatted line. */
async function looksPredict(): Promise<void> {
  const cases: Record<string, { opens: true; saved: Awaited<ReturnType<typeof looksOf>> }> = {}
  for (const id of Object.keys(LOOKS)) cases[id] = { opens: true, saved: await looksOf(await looksFile(id)) }
  writePrediction(new URL('./keep/books.ts', import.meta.url).href, 'looks', { cases })
}

/** The books of `LOOKS` through Excel: opened, computed and saved (`test/excel/keep/looks.json`, `looks/<id>.xlsx`): what the file it saved holds, and how long it took. */
async function looks(): Promise<void> {
  const file = join(DATA, 'looks.json')
  const run = loadStored<{ meta: Record<string, string | number>; cases: Record<string, { sent: string; rejected?: string; seconds: number; bytes?: number; saved?: Awaited<ReturnType<typeof looksOf>> }> }>(file) ?? { meta: {}, cases: {} }
  mkdirSync(join(DATA, 'looks'), { recursive: true })
  const t0 = Date.now()
  await ensureExcel('en-US')
  for (const id of Object.keys(LOOKS)) {
    const data = await looksFile(id)
    const sent = await partsHash(data)
    if (run.cases[id]?.sent === sent) continue
    const t = Date.now()
    const r = await saveThroughExcel(data, `looks-${id}`)
    if ('rejected' in r) run.cases[id] = { sent, rejected: r.rejected, seconds: secondsSince(t) }
    else {
      writeFileSync(join(DATA, 'looks', `${id}.xlsx`), r.saved)
      run.cases[id] = { sent, seconds: secondsSince(t), bytes: r.saved.length, saved: await looksOf(r.saved) }
    }
    console.log(`${id}: ${JSON.stringify(run.cases[id])}`)
    run.meta = { version: excelVersion(), seconds: secondsSince(t0), date: new Date().toISOString().slice(0, 10) }
    writeAnswers(file, run)
  }
}

export async function stepKeep(what: string | undefined): Promise<void> {
  const rules = what?.match(/^(rules(?:-[a-z0-9-]+)?)-predict$/)?.[1]
  if (rules && (rules === 'rules-made' || rules === 'rules-open' || RULES_SETS[rules])) return rulesPredict(rules)
  if (what === 'rules-open') return rulesOpen()
  if (what === 'made') return made()
  if (what === 'predict') return predict()
  if (what === 'open') return open()
  if (what === 'struct-predict') return writePrediction(new URL('./keep/books.ts', import.meta.url).href, 'struct', { cases: await structWritten() })
  if (what === 'struct') return struct()
  if (what === 'struct-open-predict') return structOpenPredict()
  if (what === 'struct-open') return structOpen()
  if (what === 'formats-predict') return formatsPredict()
  if (what === 'formats') return formats()
  if (what === 'win-predict') return winPredict()
  if (what === 'looks-predict') return looksPredict()
  if (what === 'looks') return looks()
  throw new Error('keep: made | predict | open | struct-predict | struct | struct-open-predict | struct-open | formats-predict | formats | win-predict | looks-predict | looks | rules-made-predict | rules[-held[-2]]-predict | rules-open-predict | rules-open')
}
