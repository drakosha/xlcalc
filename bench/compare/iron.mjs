// IronCalc (WebAssembly) behind the benchmark's calls (`adapters` of bench/hyperformula/shared.mjs): a book opened cell
// by cell with its calculation paused, edits, and values read back. It reads values only as shown text, so the cells
// get a 15-digit scientific format before they are read; that is outside every timed call.
import { readFileSync } from 'node:fs'
import { createRequire } from 'node:module'
import { adapters, hfScalar } from '../hyperformula/shared.mjs'
import { parseAddress } from '../../dist/model.js'
import { ERRORS } from './errors.mjs'

const dir = createRequire(import.meta.url).resolve('@ironcalc/wasm/package.json').replace(/package\.json$/, '')
const wasm = await import(dir + 'wasm.js')
export const ironMemory = wasm.initSync({ module: readFileSync(dir + 'wasm_bg.wasm') }).memory
export const IRONCALC = JSON.parse(readFileSync(dir + 'package.json', 'utf8')).version

const SHOWN = '0.00000000000000E+00'
const text = (v) => (typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v))

/** A cell's input as typed into IronCalc: HyperFormula's scalar (a formula with `=`, a text after an apostrophe) as text. */
function inputs(book, cells) {
  const out = []
  book.sheets.forEach((sheet, s) => {
    for (const [addr, cell] of Object.entries(cells(sheet))) {
      const { row, col } = parseAddress(addr)
      const v = cell === null ? '' : hfScalar(cell)
      if (v !== null && v !== undefined) out.push([s, row + 1, col + 1, text(v)])
    }
  })
  return out
}

/** The populated cells of each sheet given the 15-digit format, by their runs down each column, so that their numbers read back whole. */
function shown(model, book) {
  book.sheets.forEach((sheet, s) => {
    const columns = new Map()
    for (const addr of Object.keys(sheet.cells)) {
      const { row, col } = parseAddress(addr)
      if (!columns.has(col)) columns.set(col, [])
      columns.get(col).push(row + 1)
    }
    for (const [col, rows] of columns) {
      rows.sort((a, b) => a - b)
      for (let i = 0, j = 0; i < rows.length; i = j) {
        for (j = i + 1; j < rows.length && rows[j] === rows[j - 1] + 1; j++);
        model.updateRangeStyle({ sheet: s, row: rows[i], column: col + 1, width: 1, height: rows[j - 1] - rows[i] + 1 }, 'num_fmt', SHOWN)
      }
    }
  })
}

export function ironOpen(book, prepared) {
  const model = new wasm.Model('Book', 'en', 'UTC', 'en')
  book.sheets.forEach((s, i) => {
    if (i) model.newSheet()
    model.renameSheet(i, s.name)
  })
  model.pauseEvaluation()
  for (const n of book.names ?? []) model.newDefinedName(n.name, n.scope === undefined ? null : book.sheets.findIndex((s) => s.name === n.scope), n.ref.replace(/^=/, ''))
  for (const [s, r, c, v] of prepared) model.setUserInput(s, r, c, v)
  model.resumeEvaluation()
  model.evaluate()
  return model
}

export function ironValue(model, s, row, col) {
  const type = model.getCellType(s, row + 1, col + 1)
  const v = model.getFormattedCellValue(s, row + 1, col + 1)
  if (v === '') return null
  if (type === 1) return Number(v)
  if (type === 4) return v === 'TRUE'
  if (type === 16) return { error: ERRORS[v] ?? v }
  return v
}

/** Each open model's book; a model gets the shown format once, before its values are first read. */
const books = new WeakMap()
const formatted = new WeakSet()

adapters.set('ironcalc', {
  prepare: (book) => inputs(book, (sheet) => sheet.cells),
  open(book, prepared) {
    const model = ironOpen(book, prepared)
    books.set(model, book)
    return model
  },
  edit(model, edits) {
    const book = books.get(model)
    const typed = inputs(book, (sheet) => Object.fromEntries(edits.filter((e) => e.sheet === sheet.name).map((e) => [e.cell, e.input])))
    if (typed.length === 1) return () => model.setUserInput(...typed[0])
    return () => {
      model.pauseEvaluation()
      for (const t of typed) model.setUserInput(...t)
      model.resumeEvaluation()
      model.evaluate()
    }
  },
  close: (model) => model.free(),
  value(model, s, row, col) {
    if (!formatted.has(model)) {
      formatted.add(model)
      shown(model, books.get(model))
    }
    return ironValue(model, s, row, col)
  },
})
