/**
 * Spilling against Excel en-US: each case of `spill/<set>.ts` is a small book whose formulas are written as
 * dynamic array formulas (Excel 365); Excel computes it and the values of the case's rectangles are read
 * back. `npx tsx bench/coverage/run.ts spill <set>` stores the answers in `test/excel/spill/<set>.json`;
 * `test/spill-excel.test.ts` compares the engine with them without Excel.
 *
 * Reading: the `value` of a range with errors in it is unsafe (errors drop out of the list, a two-dimensional
 * read crashed Excel), so every cell of a rectangle has three mirror formulas that never give an error, on
 * the sheets `WT` (a type code: -5 blank, 0 number, -1 TRUE, -2 FALSE, -3 text, > 0 `ERROR.TYPE`), `WN`
 * (the number) and `WX` (the text), and those are read; the rectangles' formulas are read as Excel shows them.
 * `FSPILL_ONLY=id,id` runs only those cases, `FCOV_FRESH=1` runs every case again.
 */
import { spawn } from 'node:child_process'
import { existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { homedir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import ExcelJS from 'exceljs'
import { markDynamic } from './book.js'
import { dialogText, ensureExcel, excelDialogs, excelVersion } from './excel.js'
import type { Val } from './formulajs.js'
import { xl } from './sets/arrays/common.js'
import type { BookInput, RangeRef } from '../../src/contract.js'
import { formatAddress, parseAddress } from '../../src/model.js'
import { quoteSheet } from '../../src/syntax.js'
import type { SpillCase, SpillSet } from './spill/cases.js'

const HERE = fileURLToPath(new URL('.', import.meta.url))
const ROOT = join(HERE, '../..')
const SCRIPT = join(HERE, 'excel-spill.applescript')
const SANDBOX = join(homedir(), 'Library/Containers/com.microsoft.Excel/Data/tmp')
const MIRRORS = ['WT', 'WN', 'WX'] as const

export interface SpillAnswer {
  /** Non-blank cells of the case's rectangles (`S!A1`). */
  values: Record<string, Val>
  /** Formulas as Excel shows them, for the cells that have one (`=SEQUENCE(3)`). */
  formulas: Record<string, string>
}

export interface StoredSpill {
  meta: Record<string, string | number>
  cases: SpillCase[]
  answers: Record<string, SpillAnswer>
}

/** A formula in the engine's spelling as Excel stores it: newer functions prefixed, `E2#` as `ANCHORARRAY(E2)`. */
export const fileFormula = (f: string) => xl(f).replace(/((?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!)?\$?[A-Z]{1,3}\$?\d+#/g, (m) => `_xlfn.ANCHORARRAY(${m.slice(0, -1)})`)

const ERROR_TYPE: Record<number, Extract<Val, { t: 'e' }>['v']> = { 1: 'NULL', 2: 'DIV0', 3: 'VALUE', 4: 'REF', 5: 'NAME', 6: 'NUM', 7: 'NA', 8: 'GETTING_DATA', 9: 'SPILL', 14: 'CALC' }

/** A rectangle's cells row by row. */
function cellsOf(range: string): { rows: number; cols: number; addrs: string[] } {
  const [a, b = a] = range.split(':').map(parseAddress)
  const addrs: string[] = []
  for (let r = a.row; r <= b.row; r++) for (let c = a.col; c <= b.col; c++) addrs.push(formatAddress(r, c))
  return { rows: b.row - a.row + 1, cols: b.col - a.col + 1, addrs }
}

/** Where each rectangle's mirrors are: stacked from row 1 of the mirror sheets. */
function mirrorRanges(read: RangeRef[]): string[] {
  let row = 0
  return read.map((g) => {
    const { rows, cols } = cellsOf(g.range)
    const out = `A${row + 1}:${formatAddress(row + rows - 1, cols - 1)}`
    row += rows
    return out
  })
}

const MIRROR: Record<(typeof MIRRORS)[number], (x: string) => string> = {
  WT: (x) => `IF(ISBLANK(${x}),-5,IF(ISERROR(${x}),ERROR.TYPE(${x}),IF(ISNUMBER(${x}),0,IF(ISLOGICAL(${x}),IF(${x},-1,-2),IF(ISTEXT(${x}),-3,-4)))))`,
  WN: (x) => `IF(ISNUMBER(${x}),${x},0)`,
  WX: (x) => `IF(ISTEXT(${x}),${x},"")`,
}

/** The case's book as xlsx: its formulas as dynamic array formulas, its names, merged areas and looks, then the mirror sheets. */
async function writeBook(c: SpillCase, file: string): Promise<void> {
  const wb = new ExcelJS.Workbook()
  const book: BookInput = c.book
  for (const s of book.sheets) {
    const ws = wb.addWorksheet(s.name)
    for (const [addr, cell] of Object.entries(s.cells)) {
      const x = ws.getCell(addr)
      if (cell.formula !== undefined) x.value = { formula: fileFormula(cell.formula), shareType: 'array', ref: addr } as ExcelJS.CellValue
      else if (cell.value !== undefined) x.value = cell.value as ExcelJS.CellValue
      if (cell.format !== undefined) x.numFmt = cell.format
      const fill = cell.style === undefined ? undefined : book.styles![cell.style].fill
      if (fill) x.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${fill.color.slice(1)}` } }
    }
    for (const m of s.merges ?? []) ws.mergeCells(m)
  }
  const at = mirrorRanges(c.read)
  for (const m of MIRRORS) {
    const ws = wb.addWorksheet(m)
    c.read.forEach((g, i) => {
      const src = cellsOf(g.range)
      const dst = cellsOf(at[i])
      src.addrs.forEach((a, k) => (ws.getCell(dst.addrs[k]).value = { formula: MIRROR[m](`${quoteSheet(g.sheet)}!${a}`) } as ExcelJS.CellValue))
    })
  }
  // ExcelJS keeps names in a cell matrix (no `#`): its model is replaced by the case's names
  const names = (book.names ?? []).map((n) => ({ name: n.name, ranges: [fileFormula(n.ref)] }))
  Object.defineProperty(wb.definedNames, 'model', { get: () => names })
  await wb.xlsx.writeFile(file)
  await markDynamic(file, book.sheets.length)
}

/** A range's `value` or `formula` as AppleScript returns it (a cell alone is not a list) as rows. */
const grid = (x: unknown): unknown[][] => (Array.isArray(x) ? (Array.isArray(x[0]) ? (x as unknown[][]) : [x]) : [[x]])

function answerOf(c: SpillCase, raw: [unknown, unknown, unknown, unknown][]): SpillAnswer {
  const out: SpillAnswer = { values: {}, formulas: {} }
  c.read.forEach((g, i) => {
    const [t, nums, strs, fs] = raw[i].map(grid)
    const { cols, addrs } = cellsOf(g.range)
    addrs.forEach((a, k) => {
      const [r, j] = [Math.floor(k / cols), k % cols]
      const key = `${g.sheet}!${a}`
      const type = Number(t[r][j])
      const f = String(fs[r][j])
      if (f.startsWith('=')) out.formulas[key] = f
      if (type === -5) return
      out.values[key] =
        type === 0 ? { t: 'n', v: Number(nums[r][j]) } : type === -1 ? { t: 'b', v: true } : type === -2 ? { t: 'b', v: false } : type === -3 ? { t: 's', v: String(strs[r][j]) } : type > 0 ? { t: 'e', v: ERROR_TYPE[type] ?? 'ERROR' } : { t: 'other', v: `type ${type}` }
    })
  })
  return out
}

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms))

/** One case in Excel. A window that is not a book stops the run and is left open (nothing is clicked). */
async function runCase(c: SpillCase): Promise<SpillAnswer> {
  const before = dialogText()
  if (before) throw new Error(`${c.id}: Excel has a window that is not a book: ${before}`)
  const dir = join(SANDBOX, `fspill-${process.pid}`)
  mkdirSync(dir, { recursive: true })
  try {
    const file = join(dir, 'book.xlsx')
    await writeBook(c, file)
    const at = mirrorRanges(c.read)
    const lines = c.read.map((g, i) => `${g.sheet}\t${g.range}\t${at[i]}`).join('\n')
    const edits = (c.edits ?? []).map((e) => `${e.sheet}\t${e.cell}\t${e.value === null ? 'x' : typeof e.value === 'number' ? 'n' : 's'}\t${e.value ?? ''}`).join('\n')
    const child = spawn('osascript', [SCRIPT, file, lines, edits])
    let out = ''
    let err = ''
    child.stdout.on('data', (d: Buffer) => (out += d.toString('utf8')))
    child.stderr.on('data', (d: Buffer) => (err += d.toString('utf8')))
    const exited = new Promise<number | null>((res) => child.on('close', res))
    const until = Date.now() + 120_000
    for (;;) {
      const code = await Promise.race([exited, sleep(1000).then(() => undefined)])
      if (code !== undefined) {
        const after = dialogText()
        if (after) throw new Error(`${c.id}: Excel window after the case: ${after}`)
        if (code !== 0) throw new Error(`${c.id}: osascript ${code}: ${err.trim().split('\n')[0]}`)
        return answerOf(c, JSON.parse(out) as [unknown, unknown, unknown, unknown][])
      }
      if (excelDialogs().length) {
        child.kill()
        throw new Error(`${c.id}: Excel shows a window, stopped without clicking: ${dialogText()}`)
      }
      if (Date.now() > until) {
        child.kill()
        throw new Error(`${c.id}: Excel did not answer in 120 s`)
      }
    }
  } finally {
    rmSync(dir, { recursive: true, force: true })
  }
}

/** Runs the set's cases without a stored answer for the same case (FCOV_FRESH=1 runs all) and stores the answers. */
export async function stepSpill(set: string): Promise<void> {
  if (!/^[a-z0-9-]+$/.test(set)) throw new Error(`set name: [a-z0-9-], not ${set}`)
  const mod = (await import(`./spill/${set}.js`)) as SpillSet
  const cases = mod.buildSpillCases()
  const target = join(ROOT, `test/excel/spill/${set}.json`)
  const prev = existsSync(target) && !process.env.FCOV_FRESH ? (JSON.parse(readFileSync(target, 'utf8')) as StoredSpill) : null
  const was = new Map(prev?.cases.map((c) => [c.id, JSON.stringify({ ...c, predicted: undefined })]) ?? [])
  const answers: Record<string, SpillAnswer> = {}
  for (const c of cases) if (prev?.answers[c.id] && was.get(c.id) === JSON.stringify({ ...c, predicted: undefined })) answers[c.id] = prev.answers[c.id]
  const t0 = Date.now()
  mkdirSync(join(ROOT, 'test/excel/spill'), { recursive: true })
  const save = () => {
    const meta = { version: excelVersion(), locale: 'en-US (-AppleLanguages (en) -AppleLocale en_US)', seconds: Math.round((Date.now() - t0) / 100) / 10 }
    writeFileSync(target, JSON.stringify({ meta, cases, answers }, null, 1) + '\n')
  }
  await ensureExcel('en')
  let done = 0
  try {
    const only = process.env.FSPILL_ONLY?.split(',')
    for (const c of cases) {
      if (answers[c.id] || (only && !only.includes(c.id))) continue
      answers[c.id] = await runCase(c)
      done++
    }
  } finally {
    save()
    console.log(`spill ${set}: ${done} cases run, ${Object.keys(answers).length} of ${cases.length} answered, ${((Date.now() - t0) / 1000).toFixed(1)} s`)
  }
}
