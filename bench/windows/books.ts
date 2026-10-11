/**
 * The Mac harness's book steps in Excel for Windows, each in the very books the Mac step writes: what cells show
 * (`display`), spilling (`spill`), book-level features (`workbook`) and structure operations (`struct`). A case's book
 * goes to the machine with its steps (computations, edits, reads) and the answer is stored as the Mac step stores it,
 * in `test/excel/win/<kind>/`. Rectangles are read through COM directly (`Value2`, `Formula`): the mirror sheets stay
 * in the books, so the files are the Mac's, but are not read.
 */
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname } from 'node:path'
import type { RangeRef } from '../../src/contract.js'
import { cellsOf as rectCells, MIRRORS, writeCaseBook, type RectsAnswer } from '../coverage/case-book.js'
import { place, writeBook as writeDisplayBook, type StoredDisplay } from '../coverage/display.js'
import type { DisplayAnswer, DisplayCase, DisplaySet } from '../coverage/display/cases.js'
import { partsOf, type SpillAnswer, type StoredSpill } from '../coverage/spill.js'
import type { SpillCase, SpillSet } from '../coverage/spill/cases.js'
import { localeDir } from '../coverage/locales.js'
import { inCaseOrder } from '../coverage/stored.js'
import { cellsOf as structCells, LOOK, lookArgs, lookOf, opArgs, READ, styled, writeBook as writeStructBook, type ExcelStructAnswer, type RawLook, type StoredStruct } from '../coverage/struct.js'
import type { StructCase, StructSet } from '../coverage/struct/cases.js'
import { addsNames, stepLines, withMirrors, type StoredTables } from '../coverage/tables.js'
import type { TableAnswer, TableCase, TableSet } from '../coverage/tables/cases.js'
import { checkCycles, type StoredWorkbook } from '../coverage/workbook.js'
import type { Step as BookStep, WorkbookAnswer, WorkbookCase, WorkbookSet } from '../coverage/workbook/cases.js'
import { formatAddress } from '../../src/model.js'
import { writeXlsx } from '../../src/xlsx.js'
import { valOf, winFile, type WinCell } from './functions.js'
import { inJobs, metaOf, type Answer, type Step, type Tag, type WinEnv } from './remote.js'

const PART = 400

/** An answer file of a step: the stored answers of unchanged cases are kept (FCOV_FRESH=1 keeps none). */
function stored<T extends { answers: Record<string, unknown>; cases: { id: string }[] }>(file: string, same: (c: T['cases'][number]) => string): Map<string, unknown> {
  if (!existsSync(file) || process.env.FCOV_FRESH) return new Map()
  const prev = JSON.parse(readFileSync(file, 'utf8')) as T
  return new Map(prev.cases.flatMap((c) => (prev.answers[c.id] ? [[same(c), prev.answers[c.id]]] : [])))
}

const write = (file: string, data: unknown) => {
  mkdirSync(dirname(file), { recursive: true })
  writeFileSync(file, JSON.stringify(data, null, 1) + '\n')
}

/** `#RRGGBB` of a colour as COM gives it (BGR in one number). */
const rgb = (bgr: number) => [bgr & 0xff, (bgr >> 8) & 0xff, (bgr >> 16) & 0xff]
const hex = (bgr: number) => '#' + rgb(bgr).map((x) => x.toString(16).padStart(2, '0').toUpperCase()).join('')

// ---- display

/** A display part's book: computed fully, then each case cell's shown text and colour (and its format codes for built-in formats). */
async function displayBook(cases: DisplayCase[], file: string) {
  await writeDisplayBook(cases, file)
  const at = place(cases)
  const cells = cases.map((c) => formatAddress(at.get(c.id)!.row, at.get(c.id)!.col))
  return { file, steps: [{ do: 'full' }, { do: 'shown', sheet: 'S', cells, codes: cases.some((c) => c.numFmtId !== undefined) }] }
}

export async function stepDisplay(set: string, tag: Tag, e: WinEnv): Promise<void> {
  const mod = (await import(`../coverage/display/${set}.js`)) as DisplaySet
  const cases = mod.buildDisplayCases()
  const file = winFile(`display/${localeDir(tag)}${set}.json`)
  const same = (c: DisplayCase) => JSON.stringify({ ...c, predicted: undefined })
  const prev = stored<StoredDisplay>(file, same)
  const answers: Record<string, DisplayAnswer> = {}
  for (const c of cases) if (prev.has(same(c))) answers[c.id] = prev.get(same(c)) as DisplayAnswer
  const rejected: string[] = []
  const t0 = Date.now()
  const save = () => write(file, { meta: { ...metaOf(e, tag), seconds: Math.round((Date.now() - t0) / 100) / 10 }, cases, answers: inCaseOrder(cases, answers), ...(rejected.length ? { rejected } : {}) })
  const todo = cases.filter((c) => !answers[c.id])
  const parts = Array.from({ length: Math.ceil(todo.length / PART) }, (_, k) => todo.slice(k * PART, (k + 1) * PART))
  await inJobs(`display-${set}-${tag}`, parts, displayBook, (part, a, cutShort) => {
    if (a && 'answers' in a) {
      const shown = a.answers[1] as [string, number, string?, string?][]
      part.forEach((c, i) => {
        const [text, color, code, local] = shown[i]
        answers[c.id] = { text, ...(color ? { color: hex(color) } : {}), ...(c.numFmtId !== undefined ? { code, local } : {}) }
      })
      return []
    }
    if (!a && cutShort) return [part]
    if (part.length > 1) return [part.slice(0, part.length >> 1), part.slice(part.length >> 1)]
    rejected.push(part[0].id)
    return []
  }, save)
  save()
  console.log(`display ${set} ${tag}: ${todo.length} cases run, ${Object.keys(answers).length} of ${cases.length} answered${rejected.length ? `, not loaded: ${rejected.join(', ')}` : ''}`)
}

// ---- rectangles: spill and workbook

/** Reading the rectangles: each one's values and formulas, and with `f2` its formulas as Excel 365 shows them. */
const readSteps = (read: RangeRef[], f2 = false): Step[] =>
  read.flatMap((g) => [{ do: 'values', sheet: g.sheet, range: g.range }, { do: 'formulas', sheet: g.sheet, range: g.range }, ...(f2 ? [{ do: 'formulas2', sheet: g.sheet, range: g.range }] : [])])

/** The rectangles' answer from the answers of `readSteps`, as the Mac's mirror reading gives it: blank cells left out, formulas where a cell has one. */
function rectsOf(read: RangeRef[], raw: unknown[], f2 = false): RectsAnswer & { formulas2?: Record<string, string> } {
  const out: RectsAnswer & { formulas2?: Record<string, string> } = { values: {}, formulas: {}, ...(f2 ? { formulas2: {} } : {}) }
  const per = f2 ? 3 : 2
  read.forEach((g, i) => {
    const [vals, fs, f2s] = raw.slice(i * per, (i + 1) * per) as [WinCell[][], string[][], string[][]?]
    const { cols, addrs } = rectCells(g.range)
    addrs.forEach((a, k) => {
      const [r, j] = [Math.floor(k / cols), k % cols]
      const key = `${g.sheet}!${a}`
      const f = String(fs[r][j])
      if (f.startsWith('=')) {
        out.formulas[key] = f
        if (out.formulas2) out.formulas2[key] = String(f2s![r][j])
      }
      const v = valOf(vals[r][j])
      if (v) out.values[key] = v
    })
  })
  return out
}

/** Cases one book each, in jobs; a case whose book Excel did not open, or which hangs, gets `failed`. */
export async function caseJobs<C extends { id: string }>(label: string, cases: C[], book: (c: C, file: string) => Promise<{ file: string; steps: Step[] }>, take: (c: C, a: Extract<Answer, { answers: unknown[] }>) => void, failed: (c: C, why: string) => void, save: () => void) {
  await inJobs(label, cases, book, (c, a, cutShort) => {
    if (a && 'answers' in a) take(c, a)
    else if (!a && cutShort) return [c]
    else failed(c, a ? a.rejected : 'Excel hangs on the book')
    return []
  }, save)
}

export async function stepSpill(set: string, tag: Tag, e: WinEnv): Promise<void> {
  const cases = ((await import(`../coverage/spill/${set}.js`)) as SpillSet).buildSpillCases()
  const file = winFile(`spill/${set}.json`)
  const same = (c: SpillCase) => JSON.stringify({ ...c, predicted: undefined })
  const prev = stored<StoredSpill>(file, same)
  const answers: Record<string, SpillAnswer> = {}
  for (const c of cases) if (prev.has(same(c))) answers[c.id] = prev.get(same(c)) as SpillAnswer
  const t0 = Date.now()
  const save = () => write(file, { meta: { ...metaOf(e, tag), seconds: Math.round((Date.now() - t0) / 100) / 10 }, cases, answers: inCaseOrder(cases, answers) })
  const edit = (v: NonNullable<SpillCase['edits']>[number]['value']) =>
    v === null ? { kind: 'x' } : typeof v === 'number' ? { kind: 'n', value: String(v) } : typeof v === 'string' ? { kind: 's', value: v } : { kind: 'f2', value: `=${v.formula}` }
  const todo = cases.filter((c) => !answers[c.id])
  await caseJobs(
    `spill-${set}`,
    todo,
    async (c, f) => {
      await writeCaseBook(c.book, c.read, f, partsOf(c))
      const edits = (c.edits ?? []).flatMap((x): Step[] => [{ do: 'set', sheet: x.sheet, cell: x.cell, ...edit(x.value) }, { do: 'calc' }])
      return { file: f, steps: [{ do: 'full' }, ...edits, ...readSteps(c.read, !!c.legacy)] }
    },
    (c, a) => {
      const n = (c.edits ?? []).length * 2
      const answer: SpillAnswer = rectsOf(c.read, a.answers.slice(1 + n), !!c.legacy)
      for (let i = 0; i < (c.edits ?? []).length; i++) {
        const r = a.answers[1 + 2 * i] as { error?: string } | null
        if (r?.error) (answer.refused ??= {})[i] = r.error
      }
      answers[c.id] = answer
    },
    (c, why) => {
      console.log(`${c.id}: Excel did not load the book (${why})`)
      answers[c.id] = { values: {}, formulas: {}, rejected: true }
    },
    save,
  )
  save()
  console.log(`spill ${set}: ${todo.length} cases run, ${Object.keys(answers).length} of ${cases.length} answered`)
}

/** A workbook step as steps of `xl.ps1`; a read takes the rectangles and the text of the cells to show. */
function bookSteps(s: BookStep, c: WorkbookCase, shows: { sheet: string; cell: string }[]): Step[] {
  if (s === 'full' || s === 'calc') return [{ do: s }]
  if (s === 'read') return [...readSteps(c.read), ...shows.map((x) => ({ do: 'shown', sheet: x.sheet, cells: [x.cell] }))]
  if ('set' in s) {
    const { sheet, cell, value } = s.set
    return [{ do: 'set', sheet, cell, ...(value === null ? { kind: 'x' } : typeof value === 'number' ? { kind: 'n', value: String(value) } : { kind: 's', value }) }]
  }
  if ('formula' in s) return [{ do: 'set', sheet: s.formula.sheet, cell: s.formula.cell, kind: 'f', value: `=${s.formula.formula.replace(/^=/, '')}` }]
  if ('hide' in s) {
    const { sheet, rows, columns, hidden } = s.hide
    return rows !== undefined ? [{ do: 'hide', sheet, range: `A${rows.replace(':', ':A')}`, hidden }] : [{ do: 'hideColumns', sheet, range: `${columns!.split(':')[0]}1:${columns!.split(':').at(-1)}1`, hidden }]
  }
  return [{ do: 'filter', ...s.filter }]
}

export async function stepWorkbook(set: string, tag: Tag, e: WinEnv): Promise<void> {
  const cases = ((await import(`../coverage/workbook/${set}.js`)) as WorkbookSet).buildWorkbookCases()
  const file = winFile(`workbook/${set}.json`)
  const same = (c: WorkbookCase) => JSON.stringify({ ...c, predicted: undefined })
  const prev = stored<StoredWorkbook>(file, same)
  const answers: Record<string, WorkbookAnswer> = {}
  for (const c of cases) if (prev.has(same(c))) answers[c.id] = prev.get(same(c)) as WorkbookAnswer
  const t0 = Date.now()
  const failed: string[] = []
  const save = () => write(file, { meta: { ...metaOf(e, tag), seconds: Math.round((Date.now() - t0) / 100) / 10 }, cases, answers: inCaseOrder(cases, answers) })
  const todo = cases.filter((c) => !answers[c.id])
  const plan = (c: WorkbookCase) => {
    const shows = (c.show ?? []).flatMap((g) => rectCells(g.range).addrs.map((cell) => ({ sheet: g.sheet, cell })))
    return { shows, steps: (c.steps ?? ['full', 'read']).map((s) => ({ s, xl: bookSteps(s, c, shows) })) }
  }
  await caseJobs(
    `workbook-${set}`,
    todo,
    async (c, f) => {
      checkCycles(c)
      // the mirror sheets stay, under the case's names: they change a cycle's first computation (`FileParts.mirrors`)
      await writeCaseBook(c.book, c.read, f, c.file)
      return { file: f, steps: [{ do: 'settings' }, ...plan(c).steps.flatMap((x) => x.xl)] }
    },
    (c, a) => {
      const [iteration, count, delta] = a.answers[0] as [boolean, number, number]
      const { shows, steps } = plan(c)
      const reads: WorkbookAnswer['reads'] = []
      let at = 1
      for (const { s, xl } of steps) {
        const got = a.answers.slice(at, at + xl.length)
        at += xl.length
        if (s !== 'read') continue
        const rects = rectsOf(c.read, got.slice(0, c.read.length * 2))
        const texts = got.slice(c.read.length * 2).map((x) => String((x as [string][])[0][0]))
        reads.push({ ...rects, ...(shows.length ? { shown: Object.fromEntries(shows.map((x, i) => [`${x.sheet}!${x.cell}`, texts[i]])) } : {}) })
      }
      answers[c.id] = { settings: { iteration, count, delta }, reads }
    },
    (c, why) => failed.push(`${c.id}: ${why}`),
    save,
  )
  save()
  console.log(`workbook ${set}: ${todo.length} cases run, ${Object.keys(answers).length} of ${cases.length} answered`)
  if (failed.length) throw new Error(`workbook ${set}: Excel did not load ${failed.join('; ')}`)
}

// ---- tables

/**
 * A table case: its book as `writeXlsx` writes it with the three empty mirror sheets (the Mac's file), Excel computing
 * it fully and after each step; then the rectangles' values and formulas (in the user's language too for a local set)
 * and the tables of every sheet.
 */
export async function stepTables(set: string, tag: Tag, e: WinEnv): Promise<void> {
  const mod = (await import(`../coverage/tables/${set}.js`)) as TableSet
  const cases = mod.buildTableCases()
  // formulas typed and read in the user's language in every locale but en-US, as on the Mac
  const local = mod.local === true || tag !== 'en-US'
  const file = winFile(`tables/${localeDir(tag)}${set}.json`)
  const same = (c: TableCase) => JSON.stringify(c)
  const prev = stored<StoredTables>(file, same)
  const answers: Record<string, TableAnswer> = {}
  for (const c of cases) if (prev.has(same(c))) answers[c.id] = prev.get(same(c)) as TableAnswer
  const t0 = Date.now()
  const save = () => write(file, { meta: { ...metaOf(e, tag), seconds: Math.round((Date.now() - t0) / 100) / 10 }, cases, answers: inCaseOrder(cases, answers) })
  const todo = cases.filter((c) => !answers[c.id])
  const reads = (c: TableCase): Step[] => c.read.flatMap((g) => [{ do: 'values', ...g }, { do: 'formulas', ...g }, ...(local ? [{ do: 'formulasLocal', ...g }] : [])])
  await caseJobs(
    `tables-${set}`,
    todo,
    async (c, f) => {
      const { data, issues } = await writeXlsx({ book: withMirrors(c.book) }, { values: 'omit' })
      if (issues.length) throw new Error(`${c.id}: the book does not go into a file whole: ${JSON.stringify(issues)}`)
      writeFileSync(f, data)
      const names: Step[] = addsNames(c) ? [{ do: 'names', visible: true }] : []
      return { file: f, steps: [{ do: 'full' }, ...stepLines(c).flatMap((line): Step[] => [{ do: 'table', line }, { do: 'calc' }]), ...reads(c), ...names, { do: 'tables', skip: [...MIRRORS] }] }
    },
    (c, a) => {
      const n = stepLines(c).length
      const per = local ? 3 : 2
      const raw = a.answers.slice(1 + 2 * n)
      const rect = rectsOf(c.read, raw.slice(0, c.read.length * per).filter((_, i) => i % per < 2))
      const localText: Record<string, string> = {}
      if (local)
        c.read.forEach((g, i) => {
          const fs = raw[i * per + 2] as string[][]
          const { cols, addrs } = rectCells(g.range)
          addrs.forEach((x, k) => {
            const t = String(fs[Math.floor(k / cols)][k % cols])
            if (t.startsWith('=')) localText[`${g.sheet}!${x}`] = t
          })
        })
      const tables = (raw.at(-1) as [string, string, string, boolean, boolean, string[]][]).map(([sheet, name, range, headers, totals, columns]) => ({ sheet, name, range: range.replaceAll('$', ''), headers, totals, columns }))
      const answer: TableAnswer = { values: rect.values, formulas: rect.formulas, ...(local ? { local: localText } : {}), tables }
      if (addsNames(c)) answer.names = (raw.at(-2) as [string, string][]).map(([name]) => name)
      for (let i = 0; i < n; i++) {
        const r = a.answers[1 + 2 * i] as { error?: string } | null
        if (r?.error) (answer.refused ??= {})[i] = r.error
      }
      answers[c.id] = answer
    },
    (c, why) => {
      console.log(`${c.id}: Excel did not load the book (${why})`)
      answers[c.id] = { values: {}, formulas: {}, tables: [], rejected: true }
    },
    save,
  )
  save()
  console.log(`tables ${set} ${tag}: ${todo.length} cases run, ${Object.keys(answers).length} of ${cases.length} answered`)
}

// ---- structure operations

export async function stepStruct(set: string, tag: Tag, e: WinEnv): Promise<void> {
  const { books, cases, dynamic = false, look = [], widths, files } = ((await import(`../coverage/struct/${set}.js`)) as StructSet).buildStructCases()
  const lookRects = [LOOK, ...look].join(',')
  const file = winFile(`struct/${set}.json`)
  const key = (c: StructCase) => JSON.stringify([c, books[c.book], files?.[c.book]])
  const prev = existsSync(file) && !process.env.FCOV_FRESH ? (JSON.parse(readFileSync(file, 'utf8')) as StoredStruct) : null
  const was = new Map(prev?.cases.map((c) => [c.id, JSON.stringify([c, prev.books[c.book], prev.files?.[c.book]])]) ?? [])
  const answers: Record<string, ExcelStructAnswer> = {}
  for (const c of cases) if (prev?.answers[c.id] && was.get(c.id) === key(c)) answers[c.id] = prev.answers[c.id]
  const t0 = Date.now()
  const failed: string[] = []
  const save = () => {
    const meta = { ...metaOf(e, tag), read: READ, ...(Object.values(books).some(styled) ? { look: lookRects } : {}), seconds: Math.round((Date.now() - t0) / 100) / 10 }
    write(file, { meta, books, cases, ...(dynamic ? { dynamic } : {}), ...(widths === false ? { widths } : {}), ...(files ? { files } : {}), answers: inCaseOrder(cases, answers) })
  }
  const todo = cases.filter((c) => !answers[c.id])
  await caseJobs(
    `struct-${set}`,
    todo,
    async (c, f) => {
      const book = books[c.book]
      await writeStructBook(c, book, f, dynamic, files?.[c.book])
      const [sheet, range, target, to] = opArgs(c)
      const [cells, columns, rows] = styled(book) ? lookArgs(lookRects).map((x) => x.split('\n')) : [[], [], []]
      const looked = styled(book) ? { look: { cells, columns, rows } } : {}
      return {
        file: f,
        steps: [{ do: 'struct', op: c.via ?? c.op.op, sheet, range, target, to }, { do: 'sheets', range: READ, ...looked }, { do: 'formulaAt', cells: c.extra ?? [] }, { do: 'names' }],
      }
    },
    (c, a) => {
      const [op, sheets, extra, names] = a.answers as [{ error?: string } | null, [string, string, WinCell[][], [unknown[][], unknown[][], unknown[][]] | ''][], string[], [string, string][]]
      answers[c.id] = {
        ...(op?.error ? { error: `op: ${op.error}` } : {}),
        sheets: sheets.map(([name, used, f, lk]) => ({ name, used, cells: structCells(f), ...(lk ? { look: lookOf(rawLook(lk)) } : {}) })),
        extra: (c.extra ?? []).map((x, i) => ({ sheet: x.sheet, cell: x.cell, text: extra[i] })),
        names: names.map(([name, ref]) => ({ name, ref })),
      }
    },
    (c, why) => failed.push(`${c.id}: ${why}`),
    save,
  )
  save()
  console.log(`struct ${set}: ${todo.length} cases run, ${Object.keys(answers).length} of ${cases.length} answered`)
  if (failed.length) throw new Error(`struct ${set}: Excel did not load ${failed.join('; ')}`)
}

/** The look as `xl.ps1` reads it, in the Mac's shape: a fill as `[r, g, b]`. */
function rawLook([cells, cols, rows]: [unknown[][], unknown[][], unknown[][]]): RawLook {
  return [cells.map(([a, bold, fill, format, merge]) => [String(a), bold === true, rgb(Number(fill)), String(format), String(merge)]), cols as RawLook[1], rows as RawLook[2]]
}
