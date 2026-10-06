/**
 * Structure operations against Excel en-US: each case is a small book and one operation (`StructureOp`);
 * Excel applies the operation to the book and the formulas of every sheet and the defined names are read
 * back (`excel-struct.applescript`). `npx tsx bench/coverage/run.ts struct <set>` runs the cases of
 * `struct/<set>.ts` and stores the answers in `docs/excel-struct/<set>.json`;
 * `test/structure-excel.test.ts` compares `restructure` with them without Excel. `FSTRUCT_ONLY=id,id` runs
 * only those cases, `FCOV_FRESH=1` runs every case again.
 */
import { execFileSync, spawn } from 'node:child_process'
import { existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { homedir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import ExcelJS from 'exceljs'
import { declineRepair, ensureExcel, excelDialogs, excelVersion } from './excel.js'
import type { BookInput } from '../../src/contract.js'
import type { StructCase, StructCases, StructSet } from './struct/cases.js'

const HERE = fileURLToPath(new URL('.', import.meta.url))
const ROOT = join(HERE, '../..')
const OUT = join(HERE, 'out')
const SCRIPT = join(HERE, 'excel-struct.applescript')
const SANDBOX = join(homedir(), 'Library/Containers/com.microsoft.Excel/Data/tmp')

/** The rectangle read on every sheet after the operation; a case keeps its cells inside it or lists them in `extra`. */
export const READ = 'A1:Z60'

export interface ExcelStructAnswer {
  /** Excel's error when it refused the operation (AppleScript error number and text). */
  error?: string
  /** What Excel shows in each non-empty cell of `READ` (`=SUM(A1:A3)`, `5`), and the sheet's used range. */
  sheets: { name: string; used: string; cells: Record<string, string> }[]
  extra: { sheet: string; cell: string; text: string }[]
  /** `name` is `Sheet!name` for a local name; `ref` as Excel shows it (`=S!$A$5`). */
  names: { name: string; ref: string }[]
}

export interface StoredStruct extends StructCases {
  meta: Record<string, string | number>
  answers: Record<string, ExcelStructAnswer>
}

/** The book as xlsx: values and formulas (no cached results), names with their scope. */
async function writeBook(c: StructCase, book: BookInput, file: string): Promise<void> {
  const wb = new ExcelJS.Workbook()
  for (const s of book.sheets) {
    const ws = wb.addWorksheet(s.name)
    for (const [addr, cell] of Object.entries(s.cells)) {
      if (cell.value !== null && typeof cell.value === 'object') throw new Error(`${c.id}: dates are not used in structure cases`)
      ws.getCell(addr).value = (cell.formula !== undefined ? { formula: cell.formula } : (cell.value ?? null)) as ExcelJS.CellValue
    }
  }
  // ExcelJS keeps names in a cell matrix (no whole columns, no scope): its model is replaced by ours
  const index = new Map(book.sheets.map((s, i) => [s.name, i]))
  const names = (book.names ?? []).map((n) => ({ name: n.name, ranges: [n.ref], ...(n.scope !== undefined ? { localSheetId: index.get(n.scope) } : {}) }))
  Object.defineProperty(wb.definedNames, 'model', { get: () => names })
  await wb.xlsx.writeFile(file)
}

/**
 * The AppleScript arguments of an operation: sheet, range, target sheet or new name, target range. Excel's
 * `range` takes cell ranges only: whole columns and rows are spelled as cells (`B1:B1048576`), rows and
 * columns to insert or delete as the cells of their first column or row.
 */
function opArgs(c: StructCase): string[] {
  const o = c.op
  const cells = (r: string) => r.replace(/^([A-Z]+):([A-Z]+)$/, (_, a, b) => `${a}1:${b}1048576`).replace(/^(\d+):(\d+)$/, (_, a, b) => `A${a}:XFD${b}`)
  switch (o.op) {
    case 'insertRows':
    case 'deleteRows':
      return [o.sheet, `A${o.row}:A${o.row + (o.count ?? 1) - 1}`, '', '']
    case 'insertColumns':
    case 'deleteColumns':
      return [o.sheet, `${o.column}1:${colAt(o.column, (o.count ?? 1) - 1)}1`, '', '']
    case 'moveRange':
      return [o.from.sheet, cells(o.from.range), o.to.sheet, o.to.cell]
    case 'copyRange':
      return [o.from.sheet, cells(o.from.range), o.to.sheet, cells(o.to.range)]
    case 'renameSheet':
      return [o.sheet, '', o.name, '']
    case 'deleteSheet':
      return [o.sheet, '', '', '']
  }
}

function colAt(letters: string, plus: number): string {
  let n = 0
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64)
  n += plus
  let s = ''
  while (n > 0) {
    s = String.fromCharCode(65 + ((n - 1) % 26)) + s
    n = Math.floor((n - 1) / 26)
  }
  return s
}

/** Formulas of `READ` as Excel returns them (a list of rows) by address, empty cells left out. */
function cellsOf(rows: unknown): Record<string, string> {
  const out: Record<string, string> = {}
  const grid = Array.isArray(rows) ? (rows as unknown[][]) : [[rows]]
  const [, c0, r0] = /^([A-Z]+)(\d+)/.exec(READ)!
  grid.forEach((row, i) =>
    row.forEach((v, j) => {
      const text = String(v)
      if (text !== '') out[`${colAt(c0, j)}${Number(r0) + i}`] = text
    }),
  )
  return out
}

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms))
const dialogText = () => excelDialogs().map((d) => `${d.subrole} «${d.text}»`).join('; ')

/** One case in Excel. A window that is not a book stops the run (nothing is clicked); the repair window is declined. */
async function runCase(c: StructCase, book: BookInput): Promise<ExcelStructAnswer> {
  const before = dialogText()
  if (before) throw new Error(`${c.id}: Excel has a window that is not a book: ${before}`)
  if (Number(execFileSync('osascript', ['-e', 'tell application "Microsoft Excel" to count workbooks'], { encoding: 'utf8' }).trim()) !== 0) throw new Error('Excel has open books')
  const dir = join(SANDBOX, `fstruct-${process.pid}`)
  mkdirSync(dir, { recursive: true })
  try {
    const file = join(dir, 'book.xlsx')
    await writeBook(c, book, file)
    const extra = (c.extra ?? []).map((e) => `${e.sheet}\t${e.cell}`).join('\n')
    const child = spawn('osascript', [SCRIPT, file, c.op.op, ...opArgs(c), READ, extra])
    let out = ''
    let err = ''
    child.stdout.on('data', (b: Buffer) => (out += b.toString('utf8')))
    child.stderr.on('data', (b: Buffer) => (err += b.toString('utf8')))
    const exited = new Promise<number | null>((res) => child.on('close', res))
    const until = Date.now() + 90_000
    for (;;) {
      const code = await Promise.race([exited, sleep(1000).then(() => undefined)])
      if (code !== undefined) {
        const after = dialogText()
        if (after) throw new Error(`${c.id}: Excel window after the case: ${after}`)
        if (code !== 0) throw new Error(`${c.id}: osascript ${code}: ${err.trim().split('\n')[0]}`)
        const raw = JSON.parse(out) as { error: string; sheets: [string, string, unknown][]; extra: [string, string, string][]; names: [string, string][] }
        return {
          ...(raw.error ? { error: raw.error } : {}),
          sheets: raw.sheets.map(([name, used, f]) => ({ name, used, cells: cellsOf(f) })),
          extra: raw.extra.map(([sheet, cell, text]) => ({ sheet, cell, text })),
          names: raw.names.map(([name, ref]) => ({ name, ref })),
        }
      }
      if (declineRepair()) {
        child.kill()
        await exited
        throw new Error(`${c.id}: Excel did not accept the book (repair window declined)`)
      }
      const open = dialogText()
      if (open) {
        child.kill()
        throw new Error(`${c.id}: Excel shows a window, stopped without clicking: ${open}`)
      }
      if (Date.now() > until) {
        child.kill()
        throw new Error(`${c.id}: Excel did not answer in 90 s`)
      }
    }
  } finally {
    rmSync(dir, { recursive: true, force: true })
  }
}

/** Runs the set's cases that have no stored answer for the same case (FCOV_FRESH=1 runs all) and stores the answers. */
export async function stepStruct(set: string): Promise<void> {
  if (!/^[a-z0-9-]+$/.test(set)) throw new Error(`set name: [a-z0-9-], not ${set}`)
  const mod = (await import(`./struct/${set}.js`)) as StructSet
  const { books, cases } = mod.buildStructCases()
  const target = join(ROOT, `docs/excel-struct/${set}.json`)
  const prev = existsSync(target) && !process.env.FCOV_FRESH ? (JSON.parse(readFileSync(target, 'utf8')) as StoredStruct) : null
  // an answer is kept when its case and the case's book are unchanged
  const key = (c: StructCase, bs: Record<string, BookInput>) => JSON.stringify([c, bs[c.book]])
  const same = new Map(prev?.cases.map((c) => [c.id, key(c, prev.books)]) ?? [])
  const answers: Record<string, ExcelStructAnswer> = {}
  for (const c of cases) if (prev?.answers[c.id] && same.get(c.id) === key(c, books)) answers[c.id] = prev.answers[c.id]
  const t0 = Date.now()
  mkdirSync(join(ROOT, 'docs/excel-struct'), { recursive: true })
  mkdirSync(OUT, { recursive: true })
  const save = () => {
    const meta = { version: excelVersion(), locale: 'en-US (-AppleLanguages (en) -AppleLocale en_US)', read: READ, seconds: Math.round((Date.now() - t0) / 100) / 10 }
    writeFileSync(target, JSON.stringify({ meta, books, cases, answers }, null, 1) + '\n')
  }
  await ensureExcel('en')
  let done = 0
  try {
    const only = process.env.FSTRUCT_ONLY?.split(',')
    for (const c of cases) {
      if (answers[c.id] || (only && !only.includes(c.id))) continue
      answers[c.id] = await runCase(c, books[c.book])
      done++
    }
  } finally {
    save()
    console.log(`struct ${set}: ${done} cases run, ${Object.keys(answers).length} of ${cases.length} answered, ${((Date.now() - t0) / 1000).toFixed(1)} s`)
  }
}
