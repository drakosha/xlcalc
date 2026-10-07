/**
 * Cell display against Excel: each case of `display/<set>.ts` is a cell with a number format in a column of
 * a given width; Excel shows the book and the text of every case's cell (`string value`, VBA's `Range.Text`)
 * is read back with the colour of its display format's font. `npx tsx bench/coverage/run.ts display <set>`
 * stores the answers of Excel en-US in `test/excel/display/<set>.json`, `display <set> ru` those of Excel ru-RU
 * in `test/excel/display/ru-RU/<set>.json`; `test/display-excel.test.ts` compares the engine with them.
 * Cases go into books of at most `PART` cells, one column per width; a rerun keeps the answers of unchanged
 * cases (`FCOV_FRESH=1` runs every case again). A book Excel refuses to load (its repair window is declined, as
 * the `probes` step does) is halved down to the refused cases, which the step names.
 */
import JSZip from 'jszip'
import { existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { createRequire } from 'node:module'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import ExcelJS from 'exceljs'
import { dialogText, ensureExcel, excelVersion, runScriptWatched, type Locale } from './excel.js'
import { xl } from './sets/arrays/common.js'
import { formatAddress, inputScalar } from '../../src/model.js'
import type { DisplayAnswer, DisplayCase, DisplaySet } from './display/cases.js'

const HERE = fileURLToPath(new URL('.', import.meta.url))
const ROOT = join(HERE, '../..')
const SCRIPT = join(HERE, 'excel-display.applescript')
const OUT = join(HERE, 'out')
const PART = 400

/**
 * ExcelJS writes a format that equals one of Excel's built-in formats by its number, and Excel shows some built-in
 * formats by the regional settings (in ru-RU `d-mmm-yy` as `dd.mmm.yy`): a case's format is written as a custom
 * format unless the case is `builtin`.
 */
const NumFmt = createRequire(import.meta.url)('exceljs/lib/xlsx/xform/style/numfmt-xform.js') as { getDefaultFmtId(code: string): number | undefined }
const builtinId = NumFmt.getDefaultFmtId

/** The automatic font colour reads as black, as `[Black]` does: neither is stored. */
const BLACK = '#000000'

export interface StoredDisplay {
  meta: Record<string, string | number>
  cases: DisplayCase[]
  answers: Record<string, DisplayAnswer>
}

/** Where the answers of a set are stored for a locale. */
export const displayFile = (set: string, locale: Locale) => join(ROOT, `test/excel/display/${locale === 'en' ? '' : 'ru-RU/'}${set}.json`)

/** The address of each case: columns by width and span (from B, one empty column between them), cases down from row 1. */
export function place(cases: DisplayCase[]): Map<string, { row: number; col: number }> {
  const cols = new Map<string, number>()
  const rows = new Map<string, number>()
  let next = 1
  const out = new Map<string, { row: number; col: number }>()
  for (const c of cases) {
    const k = `${c.width ?? ''}|${c.span ?? 1}`
    if (!cols.has(k)) {
      cols.set(k, next)
      next += (c.span ?? 1) + 1
    }
    const row = rows.get(k) ?? 0
    rows.set(k, row + 1)
    out.set(c.id, { row, col: cols.get(k)! })
  }
  return out
}

/** The book of a part: sheet `S` with the cases' cells, column widths and merged areas; sheet `L` lists their addresses. */
async function writeBook(cases: DisplayCase[], file: string): Promise<void> {
  const wb = new ExcelJS.Workbook()
  const ws = wb.addWorksheet('S')
  const list = wb.addWorksheet('L')
  const at = place(cases)
  const builtins = new Set(cases.flatMap((c) => (c.builtin ? [c.cell.format] : [])))
  if (cases.some((c) => !c.builtin && builtins.has(c.cell.format))) throw new Error('a format is both built-in and custom in one book')
  NumFmt.getDefaultFmtId = (code) => (builtins.has(code) ? builtinId(code) : undefined)
  cases.forEach((c, i) => {
    const { row, col } = at.get(c.id)!
    const addr = formatAddress(row, col)
    const x = ws.getCell(addr)
    x.value = (c.cell.formula !== undefined ? { formula: xl(c.cell.formula) } : inputScalar(c.cell.value)) as ExcelJS.CellValue
    if (c.cell.format !== undefined) x.numFmt = c.cell.format
    else if (c.numFmtId !== undefined) x.numFmt = builtinMark(c.numFmtId)
    for (let k = 0; k < (c.span ?? 1); k++) if (c.width !== undefined) ws.getColumn(col + 1 + k).width = c.width
    if ((c.span ?? 1) > 1) ws.mergeCells(`${addr}:${formatAddress(row, col + c.span! - 1)}`)
    list.getCell(`A${i + 1}`).value = addr
  })
  await wb.xlsx.writeFile(file)
  if (cases.some((c) => c.numFmtId !== undefined)) await builtinIds(file)
}

/** A placeholder code that marks a built-in format by number until `builtinIds` writes the number alone. */
const builtinMark = (id: number) => `"builtin ${id}"`

/** The cells' formats marked by `builtinMark` become built-in numbers without a code, as Excel stores them. */
async function builtinIds(file: string) {
  const zip = await JSZip.loadAsync(readFileSync(file))
  let xml = await zip.file('xl/styles.xml')!.async('string')
  const ids = new Map<string, string>()
  xml = xml.replace(/<numFmt numFmtId="(\d+)" formatCode="&quot;builtin (\d+)&quot;"\/>/g, (_, custom: string, id: string) => {
    ids.set(custom, id)
    return ''
  })
  xml = xml.replace(/<numFmts count="\d+">/, () => `<numFmts count="${(xml.match(/<numFmt /g) ?? []).length}">`).replace(/<numFmts count="0"><\/numFmts>/, '')
  xml = xml.replace(/(<xf [^>]*numFmtId=")(\d+)"/g, (m, head: string, id: string) => (ids.has(id) ? `${head}${ids.get(id)}"` : m))
  zip.file('xl/styles.xml', xml)
  writeFileSync(file, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }))
}

/** One part in Excel: its answers, or 'rejected' when Excel would not load the book (its repair window is declined). */
async function runPart(cases: DisplayCase[], label: string): Promise<DisplayAnswer[] | 'rejected'> {
  const before = dialogText()
  if (before) throw new Error(`${label}: Excel has a window that is not a book: ${before}`)
  mkdirSync(OUT, { recursive: true })
  const file = join(OUT, `display-${process.pid}.xlsx`)
  await writeBook(cases, file)
  const codes = cases.some((c) => c.numFmtId !== undefined)
  const out = await runScriptWatched(SCRIPT, file, cases.length, codes ? ['codes'] : []).finally(() => rmSync(file))
  if (out === 'rejected') return out
  const after = dialogText()
  if (after) throw new Error(`${label}: Excel window after the part: ${after}`)
  const raw = JSON.parse(out) as [string, number[] | '', string?, string?][]
  if (raw.length !== cases.length) throw new Error(`${label}: Excel gave ${raw.length} cells of ${cases.length}`)
  return raw.map(([text, rgb, code, local], i) => {
    if (!Array.isArray(rgb)) throw new Error(`${label}: no colour for ${text}`)
    const color = '#' + rgb.map((x) => x.toString(16).padStart(2, '0').toUpperCase()).join('')
    const codesOf = cases[i].numFmtId !== undefined ? { code, local } : {}
    return color === BLACK ? { text, ...codesOf } : { text, color, ...codesOf }
  })
}

/** Runs the set's cases without a stored answer for the same case and stores the answers after each part. */
export async function stepDisplay(set: string, locale: Locale): Promise<void> {
  if (!/^[a-z0-9-]+$/.test(set)) throw new Error(`set name: [a-z0-9-], not ${set}`)
  const mod = (await import(`./display/${set}.js`)) as DisplaySet
  const cases = mod.buildDisplayCases()
  const target = displayFile(set, locale)
  const prev = existsSync(target) && !process.env.FCOV_FRESH ? (JSON.parse(readFileSync(target, 'utf8')) as StoredDisplay) : null
  const same = (c: DisplayCase) => JSON.stringify({ ...c, predicted: undefined })
  const was = new Map(prev?.cases.map((c) => [c.id, same(c)]) ?? [])
  const answers: Record<string, DisplayAnswer> = {}
  for (const c of cases) if (prev?.answers[c.id] && was.get(c.id) === same(c)) answers[c.id] = prev.answers[c.id]
  const t0 = Date.now()
  mkdirSync(join(target, '..'), { recursive: true })
  const save = () => {
    const meta = { version: excelVersion(), locale: locale === 'en' ? 'en-US (-AppleLanguages (en) -AppleLocale en_US)' : 'ru-RU (system language and locale)', seconds: Math.round((Date.now() - t0) / 100) / 10 }
    writeFileSync(target, JSON.stringify({ meta, cases, answers }, null, 1) + '\n')
  }
  await ensureExcel(locale)
  const todo = cases.filter((c) => !answers[c.id])
  const rejected: string[] = []
  try {
    // a book Excel does not load is halved down to the cases it refuses
    const run = async (part: DisplayCase[], label: string): Promise<void> => {
      const got = await runPart(part, label)
      if (got !== 'rejected') return void part.forEach((c, i) => (answers[c.id] = got[i]))
      if (part.length === 1) return void rejected.push(part[0].id)
      await run(part.slice(0, part.length >> 1), `${label}a`)
      await run(part.slice(part.length >> 1), `${label}b`)
    }
    for (let k = 0; k * PART < todo.length; k++) {
      await run(todo.slice(k * PART, (k + 1) * PART), `${set} ${locale} part ${k + 1}`)
      save()
    }
  } finally {
    save()
    console.log(`display ${set} ${locale}: ${todo.length} cases run, ${Object.keys(answers).length} of ${cases.length} answered, ${((Date.now() - t0) / 1000).toFixed(1)} s`)
  }
  if (rejected.length) throw new Error(`Excel does not load the cases ${rejected.join(', ')}: remove them from the set`)
}
