/**
 * Excel's answers to structure operations (`docs/excel-struct/<set>.json`, from `run.ts struct <set>`)
 * and the comparison with `restructure`: what each cell shows (`=formula` or the value) in the read
 * rectangle of every sheet, the extra cells, the sheets and the names. Formulas are compared in the form
 * Excel writes references (`A1:A1048576` is `A:A`), with sheet names quoted.
 */
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import type { ExcelStructAnswer, StoredStruct } from '../bench/coverage/struct.js'
import type { BookInput, CellInput } from '../src/contract.js'
import { formatAddress, parseAddress } from '../src/model.js'
import { decode, encode } from '../src/structure/area.js'
import { canonical, parse, sheetKey, walk } from '../src/syntax.js'

const dir = new URL('../docs/excel-struct/', import.meta.url)

export const STRUCT_SETS: (StoredStruct & { name: string })[] = existsSync(dir)
  ? readdirSync(dir)
      .filter((f) => f.endsWith('.json'))
      .sort()
      .map((f) => ({ name: f.slice(0, -5), ...(JSON.parse(readFileSync(new URL(f, dir), 'utf8')) as StoredStruct) }))
  : []

/** Formula text (without `=`) with references as Excel writes them; as written when it does not parse. */
export function normalFormula(f: string): string {
  let ast
  try {
    ast = parse(f)
  } catch {
    return f
  }
  walk(ast, (n) => {
    if (n.k === 'ref') n.ref = { ...n.ref, text: encode(decode(n.ref)) }
  })
  return canonical(ast)
}

/** What Excel's `formula` of a cell shows for our cell. */
function shown(c: CellInput): string {
  if (c.formula !== undefined) return `=${normalFormula(c.formula)}`
  const v = c.value
  if (v === undefined || v === null) return ''
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'
  return typeof v === 'object' ? v.date : String(v)
}

const excelShown = (text: string) => (text.startsWith('=') ? `=${normalFormula(text.slice(1))}` : text)

/** Excel's `Sheet!name` as scope and name. */
function nameKey(excel: string): string {
  const i = excel.lastIndexOf('!')
  if (i < 0) return `|${excel.toUpperCase()}`
  const scope = excel.slice(0, i).replace(/^'(.*)'$/, '$1').replaceAll("''", "'")
  return `${sheetKey(scope)}|${excel.slice(i + 1).toUpperCase()}`
}

/** Differences between a book and Excel's answer, as `place: Excel … , engine …` lines. */
export function differences(book: BookInput, a: ExcelStructAnswer, read: string): string[] {
  const out: string[] = []
  const [r0, r1] = read.split(':').map(parseAddress)
  const inRead = (addr: string) => {
    const p = parseAddress(addr)
    return p.row >= r0.row && p.row <= r1.row && p.col >= r0.col && p.col <= r1.col
  }
  const names = book.sheets.map((s) => s.name)
  if (JSON.stringify(names) !== JSON.stringify(a.sheets.map((s) => s.name))) out.push(`sheets: Excel ${JSON.stringify(a.sheets.map((s) => s.name))}, engine ${JSON.stringify(names)}`)
  for (const s of a.sheets) {
    const ours = book.sheets.find((x) => x.name === s.name)
    if (!ours) continue
    const addrs = new Set([...Object.keys(s.cells), ...Object.keys(ours.cells).filter(inRead)])
    for (const addr of addrs) {
      const excel = s.cells[addr] === undefined ? '' : excelShown(s.cells[addr])
      const engine = ours.cells[addr] ? shown(ours.cells[addr]) : ''
      if (excel !== engine) out.push(`'${s.name}'!${addr}: Excel ${JSON.stringify(excel)}, engine ${JSON.stringify(engine)}`)
    }
    // Excel's used range must lie in the read rectangle unless the case reads extra cells
    const used = s.used.replaceAll('$', '').split(':')
    if (!a.extra.length && !used.every((x) => /^[A-Z]+\d+$/.test(x) && inRead(x))) out.push(`'${s.name}': Excel's used range ${s.used} is outside ${read}`)
  }
  for (const e of a.extra) {
    const c = book.sheets.find((x) => x.name === e.sheet)?.cells[e.cell]
    const engine = c ? shown(c) : ''
    if (excelShown(e.text) !== engine) out.push(`'${e.sheet}'!${e.cell}: Excel ${JSON.stringify(e.text)}, engine ${JSON.stringify(engine)}`)
  }
  const excelNames = new Map(a.names.filter((n) => !n.name.startsWith('_xlfn.')).map((n) => [nameKey(n.name), normalFormula(n.ref.slice(1))]))
  const ourNames = new Map((book.names ?? []).map((n) => [`${n.scope === undefined ? '' : sheetKey(n.scope)}|${n.name.toUpperCase()}`, normalFormula(n.ref)]))
  for (const k of new Set([...excelNames.keys(), ...ourNames.keys()]))
    if (excelNames.get(k) !== ourNames.get(k)) out.push(`name ${k}: Excel ${JSON.stringify(excelNames.get(k) ?? null)}, engine ${JSON.stringify(ourNames.get(k) ?? null)}`)
  return out
}

/** Number of formulas Excel showed in a case (for the report). */
export const formulaCount = (a: ExcelStructAnswer) => a.sheets.reduce((n, s) => n + Object.values(s.cells).filter((t) => t.startsWith('=')).length, 0) + a.names.filter((n) => !n.name.startsWith('_xlfn.')).length

export { formatAddress }
