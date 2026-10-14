/**
 * The Mac harness's book steps in Excel for Windows, each in the very books the Mac step writes: what cells show
 * (`display`), spilling (`spill`), book-level features (`workbook`) and structure operations (`struct`). A case's book
 * goes to the machine with its steps (computations, edits, reads) and the answer is stored as the Mac step stores it,
 * in `test/excel/win/<kind>/`. Rectangles are read through COM directly (`Value2`, `Formula`): the mirror sheets stay
 * in the books, so the files are the Mac's, but are not read.
 */
import { writeFileSync } from 'node:fs'
import type { CellValue, RangeRef } from '../../src/contract.js'
import { inParts } from '../coverage/book.js'
import { fromBgr } from '../coverage/cond/cases.js'
import { cellsOf as rectCells, eachRectCell, editOf, MIRRORS, refusedOf, writeCaseBook, type RectsAnswer } from '../coverage/case-book.js'
import { place, writeBook as writeDisplayBook, type StoredDisplay } from '../coverage/display.js'
import type { DisplayCase, DisplaySet } from '../coverage/display/cases.js'
import { partsOf, type SpillAnswer, type StoredSpill } from '../coverage/spill.js'
import type { SpillSet } from '../coverage/spill/cases.js'
import { localeDir } from '../coverage/locales.js'
import { answeredLine, inCaseOrder, keptSame, loadStored, secondsSince, setModule, unpredicted, writeAnswers, type FileAnswers } from '../coverage/stored.js'
import { cellsOf as structCells, LOOK, lookArgs, lookOf, opArgs, READ, structKey, styled, writeBook as writeStructBook, type RawLook, type StoredStruct } from '../coverage/struct.js'
import type { StructCase, StructSet } from '../coverage/struct/cases.js'
import { addsNames, stepLines, tableShapes, withMirrors, type RawTable, type StoredTables } from '../coverage/tables.js'
import type { TableAnswer, TableCase, TableSet } from '../coverage/tables/cases.js'
import { checkCycles, hideRange, type StoredWorkbook } from '../coverage/workbook.js'
import type { Step as BookStep, WorkbookAnswer, WorkbookCase, WorkbookSet } from '../coverage/workbook/cases.js'
import { formatAddress } from '../../src/model.js'
import { readXlsx, writeXlsx } from '../../src/xlsx.js'
import { valOf, winFile, type WinCell } from './functions.js'
import { again, inJobs, metaOf, type Answer, type Step, type Tag, type WinEnv } from './remote.js'

/** A colour as COM gives it (BGR in one number) as `[r, g, b]`. */
const rgb = (bgr: number) => [bgr & 0xff, (bgr >> 8) & 0xff, (bgr >> 16) & 0xff]

// ---- display

/** A display part's book: computed fully, then each case cell's shown text and colour (and its format codes for built-in formats). */
async function displayBook(cases: DisplayCase[], file: string) {
  await writeDisplayBook(cases, file)
  const at = place(cases)
  const cells = cases.map((c) => formatAddress(at.get(c.id)!.row, at.get(c.id)!.col))
  return { file, steps: [{ do: 'full' }, { do: 'shown', sheet: 'S', cells, codes: cases.some((c) => c.numFmtId !== undefined) }] }
}

export async function stepDisplay(set: string, tag: Tag, e: WinEnv): Promise<void> {
  const cases = (await setModule<DisplaySet>('display', set)).buildDisplayCases()
  const file = winFile(`display/${localeDir(tag)}${set}.json`)
  const answers = keptSame({ cases }, loadStored<StoredDisplay>(file), unpredicted)
  const rejected: string[] = []
  const t0 = Date.now()
  const save = () => writeAnswers(file, { meta: { ...metaOf(e, tag), seconds: secondsSince(t0) }, cases, answers: inCaseOrder(cases, answers), ...(rejected.length ? { rejected } : {}) })
  const todo = cases.filter((c) => !answers[c.id])
  await inJobs(`display-${set}-${tag}`, inParts(todo), displayBook, (part, a, cutShort) => {
    if (a && 'answers' in a) {
      const shown = a.answers[1] as [string, number, string?, string?][]
      part.forEach((c, i) => {
        const [text, color, code, local] = shown[i]
        answers[c.id] = { text, ...(color ? { color: fromBgr(color) } : {}), ...(c.numFmtId !== undefined ? { code, local } : {}) }
      })
      return []
    }
    return again(part, a, cutShort, (c) => rejected.push(c.id))
  }, save)
  save()
  console.log(`${answeredLine(`display ${set} ${tag}`, todo.length, answers, cases)}${rejected.length ? `, not loaded: ${rejected.join(', ')}` : ''}`)
}

// ---- rectangles: spill and workbook

/** Reading the rectangles: each one's values and formulas, and with `f2` its formulas as Excel 365 shows them. */
const readSteps = (read: RangeRef[], f2 = false): Step[] =>
  read.flatMap((g) => [{ do: 'values', sheet: g.sheet, range: g.range }, { do: 'formulas', sheet: g.sheet, range: g.range }, ...(f2 ? [{ do: 'formulas2', sheet: g.sheet, range: g.range }] : [])])

/** The rectangles' answer from the answers of `readSteps`, as the Mac's mirror reading gives it: blank cells left out, formulas where a cell has one. */
function rectsOf(read: RangeRef[], raw: unknown[], f2 = false): RectsAnswer & { formulas2?: Record<string, string> } {
  const out: RectsAnswer & { formulas2?: Record<string, string> } = { values: {}, formulas: {}, ...(f2 ? { formulas2: {} } : {}) }
  const per = f2 ? 3 : 2
  eachRectCell(read, (key, i, r, j) => {
    const [vals, fs, f2s] = raw.slice(i * per, (i + 1) * per) as [WinCell[][], string[][], string[][]?]
    const f = String(fs[r][j])
    if (f.startsWith('=')) {
      out.formulas[key] = f
      if (out.formulas2) out.formulas2[key] = String(f2s![r][j])
    }
    const v = valOf(vals[r][j])
    if (v) out.values[key] = v
  })
  return out
}

/** A value typed into a cell as `xl.ps1` takes it: `editOf`'s kind, a formula by `Formula2` (f2), and its text but for a cell cleared. */
export function comEdit(v: Parameters<typeof editOf>[0]): { kind: string; value?: string } {
  const [kind, value] = editOf(v)
  return kind === 'x' ? { kind } : { kind: kind === 'f' ? 'f2' : kind, value }
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

/** A case's structure operation as a step of `xl.ps1` (`op` its name there), none without one. */
export function opSteps(c: { id: string; book: string; op?: StructCase['op'] }, op: string | undefined = c.op?.op): Step[] {
  if (!c.op) return []
  const [sheet, range, target, to] = opArgs({ id: c.id, book: c.book, op: c.op })
  return [{ do: 'struct', op, sheet, range, target, to }]
}

/** The errors of the steps whose answers these are (null, or `{ error }`), each after the name `name` gives its step. */
export const stepErrors = (answers: unknown[], name: (i: number) => string): string[] => (answers as ({ error?: string } | null)[]).flatMap((x, i) => (x?.error ? [`${name(i)}: ${x.error}`] : []))

/** The errors of what a case makes and of its operation, whose steps come after the first `skip`; `n` the answers up to their end. */
export function makeErrors(c: { make?: unknown[]; op?: unknown }, answers: unknown[], skip = 0): { n: number; errors: string[] } {
  const makes = (c.make ?? []).length
  const n = skip + makes + (c.op ? 1 : 0)
  return { n, errors: stepErrors(answers.slice(skip, n), (i) => (i < makes ? 'make' : 'op')) }
}

/** The refusals of a case's `n` edits, whose answers come after the full calculation, each before its calculation. */
const editsRefused = (answers: unknown[], n: number) => refusedOf(Array.from({ length: n }, (_, i) => (answers[1 + 2 * i] as { error?: string } | null)?.error))

/** What a step records of a case whose book Excel did not load: `empty` answers, marked rejected. */
const rejected =
  <A>(answers: Record<string, A>, empty: Omit<A, 'rejected'>) =>
  (c: { id: string }, why: string) => {
    console.log(`${c.id}: Excel did not load the book (${why})`)
    answers[c.id] = { ...empty, rejected: true } as A
  }

/**
 * The cases to do of a step whose cases go to Excel as files (`files`) and come back saved, one book each, in jobs
 * (`job` names them): `steps` are a case's steps, `answer` its answer from theirs. A case whose book Excel did not load,
 * hung on or did not save is refused; the answers are saved after every job, and the step's line (`label`) printed.
 */
export async function fileJobs<C extends { id: string }, A>(label: string, job: string, store: FileAnswers<C, A>, files: Map<string, Uint8Array>, steps: (c: C) => Step[], answer: (c: C, a: Extract<Answer, { answers: unknown[] }>) => A, save: () => void): Promise<void> {
  const todo = store.todo
  await caseJobs(
    job,
    todo,
    async (c, f) => {
      writeFileSync(f, files.get(c.id)!)
      return { file: f, steps: steps(c) }
    },
    (c, a) => {
      const got = answer(c, a)
      store.got(c.id, a.saved ? { answer: got, saved: a.saved } : { rejected: 'Excel did not save the book' })
    },
    (c, why) => {
      console.log(`${c.id}: Excel did not load the book (${why})`)
      store.got(c.id, { rejected: why })
    },
    save,
  )
  save()
  console.log(answeredLine(label, todo.length, store.answers, store.cases, store.rejected))
}

export async function stepSpill(set: string, tag: Tag, e: WinEnv): Promise<void> {
  const cases = (await setModule<SpillSet>('spill', set)).buildSpillCases()
  const file = winFile(`spill/${set}.json`)
  const answers = keptSame({ cases }, loadStored<StoredSpill>(file), unpredicted)
  const t0 = Date.now()
  const save = () => writeAnswers(file, { meta: { ...metaOf(e, tag), seconds: secondsSince(t0) }, cases, answers: inCaseOrder(cases, answers) })
  const todo = cases.filter((c) => !answers[c.id])
  await caseJobs(
    `spill-${set}`,
    todo,
    async (c, f) => {
      await writeCaseBook(c.book, c.read, f, partsOf(c))
      const edits = (c.edits ?? []).flatMap((x): Step[] => [{ do: 'set', sheet: x.sheet, cell: x.cell, ...comEdit(x.value) }, { do: 'calc' }])
      return { file: f, steps: [{ do: 'full' }, ...edits, ...readSteps(c.read, !!c.legacy)] }
    },
    (c, a) => {
      const n = (c.edits ?? []).length * 2
      const answer: SpillAnswer = rectsOf(c.read, a.answers.slice(1 + n), !!c.legacy)
      const refused = editsRefused(a.answers, (c.edits ?? []).length)
      answers[c.id] = refused ? { ...answer, refused } : answer
    },
    rejected(answers, { values: {}, formulas: {} }),
    save,
  )
  save()
  console.log(answeredLine(`spill ${set}`, todo.length, answers, cases))
}

/** A workbook step as steps of `xl.ps1`; a read takes the rectangles and the text of the cells to show. */
function bookSteps(s: BookStep, c: WorkbookCase, shows: { sheet: string; cell: string }[]): Step[] {
  if (s === 'full' || s === 'calc') return [{ do: s }]
  if (s === 'save') return [{ do: 'saveAs' }]
  if (s === 'read') return [...readSteps(c.read), ...shows.map((x) => ({ do: 'shown', sheet: x.sheet, cells: [x.cell] }))]
  if ('set' in s) {
    const { sheet, cell, value } = s.set
    return [{ do: 'set', sheet, cell, ...comEdit(value) }]
  }
  if ('formula' in s) return [{ do: 'set', sheet: s.formula.sheet, cell: s.formula.cell, kind: 'f', value: `=${s.formula.formula.replace(/^=/, '')}` }]
  if ('hide' in s) {
    const { sheet, rows, hidden } = s.hide
    return [{ do: rows !== undefined ? 'hide' : 'hideColumns', sheet, range: hideRange(s.hide), hidden }]
  }
  return [{ do: 'filter', ...s.filter }]
}

/** What a file Excel saved stores for the cells of the rectangles, by `Sheet!A1`. */
async function savedValues(read: RangeRef[], saved: Uint8Array): Promise<Record<string, CellValue>> {
  const { cached } = await readXlsx(saved)
  const out: Record<string, CellValue> = {}
  for (const g of read) {
    const cells = cached.sheets.find((x) => x.name === g.sheet)?.cells ?? {}
    for (const a of rectCells(g.range).addrs) if (a in cells) out[`${g.sheet}!${a}`] = cells[a].value
  }
  return out
}

export async function stepWorkbook(set: string, tag: Tag, e: WinEnv): Promise<void> {
  const cases = (await setModule<WorkbookSet>('workbook', set)).buildWorkbookCases()
  const file = winFile(`workbook/${set}.json`)
  const answers = keptSame({ cases }, loadStored<StoredWorkbook>(file), unpredicted)
  const t0 = Date.now()
  const failed: string[] = []
  const save = () => writeAnswers(file, { meta: { ...metaOf(e, tag), seconds: secondsSince(t0) }, cases, answers: inCaseOrder(cases, answers) })
  const todo = cases.filter((c) => !answers[c.id])
  // the files Excel saved at a case's `save` step, read once the jobs are done
  const saved = new Map<string, Uint8Array>()
  const plan = (c: WorkbookCase) => {
    const shows = (c.show ?? []).flatMap((g) => rectCells(g.range).addrs.map((cell) => ({ sheet: g.sheet, cell })))
    return { shows, steps: (c.steps ?? ['full', 'read']).map((s) => ({ s, xl: bookSteps(s, c, shows) })) }
  }
  await caseJobs(
    `workbook-${set}`,
    todo,
    async (c, f) => {
      if (!c.cycle) checkCycles(c)
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
      if (a.saved) saved.set(c.id, a.saved)
    },
    (c, why) => failed.push(`${c.id}: ${why}`),
    save,
  )
  for (const c of todo) if (saved.has(c.id)) answers[c.id].saved = await savedValues(c.read, saved.get(c.id)!)
  save()
  console.log(answeredLine(`workbook ${set}`, todo.length, answers, cases))
  if (failed.length) throw new Error(`workbook ${set}: Excel did not load ${failed.join('; ')}`)
}

// ---- tables

/**
 * A table case: its book as `writeXlsx` writes it with the three empty mirror sheets (the Mac's file), Excel computing
 * it fully and after each step; then the rectangles' values and formulas (in the user's language too for a local set)
 * and the tables of every sheet.
 */
export async function stepTables(set: string, tag: Tag, e: WinEnv): Promise<void> {
  const mod = await setModule<TableSet>('tables', set)
  const cases = mod.buildTableCases()
  // formulas typed and read in the user's language in every locale but en-US, as on the Mac
  const local = mod.local === true || tag !== 'en-US'
  const file = winFile(`tables/${localeDir(tag)}${set}.json`)
  const answers = keptSame({ cases }, loadStored<StoredTables>(file))
  const t0 = Date.now()
  const save = () => writeAnswers(file, { meta: { ...metaOf(e, tag), seconds: secondsSince(t0) }, cases, answers: inCaseOrder(cases, answers) })
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
      // a cell's formula in the user's language where that text is a formula
      if (local)
        eachRectCell(c.read, (key, i, r, j) => {
          const t = String((raw[i * per + 2] as string[][])[r][j])
          if (t.startsWith('=')) localText[key] = t
        })
      const answer: TableAnswer = { values: rect.values, formulas: rect.formulas, ...(local ? { local: localText } : {}), tables: tableShapes(raw.at(-1) as RawTable[]) }
      if (addsNames(c)) answer.names = (raw.at(-2) as [string, string][]).map(([name]) => name)
      const refused = editsRefused(a.answers, n)
      answers[c.id] = refused ? { ...answer, refused } : answer
    },
    rejected(answers, { values: {}, formulas: {}, tables: [] }),
    save,
  )
  save()
  console.log(answeredLine(`tables ${set} ${tag}`, todo.length, answers, cases))
}

// ---- structure operations

export async function stepStruct(set: string, tag: Tag, e: WinEnv): Promise<void> {
  const built = (await setModule<StructSet>('struct', set)).buildStructCases()
  const { books, cases, dynamic = false, look = [], widths, files } = built
  const lookRects = [LOOK, ...look].join(',')
  const file = winFile(`struct/${set}.json`)
  const answers = keptSame(built, loadStored<StoredStruct>(file), structKey)
  const t0 = Date.now()
  const failed: string[] = []
  const save = () => {
    const meta = { ...metaOf(e, tag), read: READ, ...(Object.values(books).some(styled) ? { look: lookRects } : {}), seconds: secondsSince(t0) }
    writeAnswers(file, { meta, books, cases, ...(dynamic ? { dynamic } : {}), ...(widths === false ? { widths } : {}), ...(files ? { files } : {}), answers: inCaseOrder(cases, answers) })
  }
  const todo = cases.filter((c) => !answers[c.id])
  await caseJobs(
    `struct-${set}`,
    todo,
    async (c, f) => {
      const book = books[c.book]
      await writeStructBook(c, book, f, dynamic, files?.[c.book])
      const [cells, columns, rows] = styled(book) ? lookArgs(lookRects).map((x) => x.split('\n')) : [[], [], []]
      const looked = styled(book) ? { look: { cells, columns, rows } } : {}
      const before = c.before ? opSteps({ ...c, op: c.before }) : []
      return { file: f, steps: [...before, ...opSteps(c, c.via ?? c.op.op), { do: 'sheets', range: READ, ...looked }, { do: 'formulaAt', cells: c.extra ?? [] }, { do: 'names' }] }
    },
    (c, a) => {
      const all = a.answers as unknown[]
      const first = c.before ? (all[0] as { error?: string } | null) : null
      const [op, sheets, extra, names] = (c.before ? all.slice(1) : all) as [{ error?: string } | null, [string, string, WinCell[][], [unknown[][], unknown[][], unknown[][]] | ''][], string[], [string, string][]]
      answers[c.id] = {
        ...(first?.error ? { error: `before: ${first.error}` } : op?.error ? { error: `op: ${op.error}` } : {}),
        sheets: sheets.map(([name, used, f, lk]) => ({ name, used, cells: structCells(f), ...(lk ? { look: lookOf(rawLook(lk)) } : {}) })),
        extra: (c.extra ?? []).map((x, i) => ({ sheet: x.sheet, cell: x.cell, text: extra[i] })),
        names: names.map(([name, ref]) => ({ name, ref })),
      }
    },
    (c, why) => failed.push(`${c.id}: ${why}`),
    save,
  )
  save()
  console.log(answeredLine(`struct ${set}`, todo.length, answers, cases))
  if (failed.length) throw new Error(`struct ${set}: Excel did not load ${failed.join('; ')}`)
}

/** Excel's `HorizontalAlignment` constants by name (`AlignmentStyle.horizontal`); a range of mixed alignments reads null. */
const ALIGN: Record<number, string> = { 1: 'general', [-4131]: 'left', [-4108]: 'center', [-4152]: 'right', 5: 'fill', [-4130]: 'justify', 7: 'centerContinuous', [-4117]: 'distributed' }

/** A cell's alignment as `xl.ps1` writes the constant (an integer as `{ e }`), by name. */
function alignOf(v: unknown): string {
  const n = typeof v === 'object' && v !== null && 'e' in v ? Number(v.e) : Number(v)
  return ALIGN[n] ?? `unknown ${JSON.stringify(v)}`
}

/** The look as `xl.ps1` reads it, in the Mac's shape: a fill as `[r, g, b]`. */
function rawLook([cells, cols, rows]: [unknown[][], unknown[][], unknown[][]]): RawLook {
  return [cells.map(([a, bold, fill, format, merge, align]) => [String(a), bold === true, rgb(Number(fill)), String(format), String(merge), alignOf(align)]), cols as RawLook[1], rows as RawLook[2]]
}
