// node --max-old-space-size=8192 read-bench.mjs <base> <branch> <dir with node_modules/hyperformula>: reading cells,
// both builds and HyperFormula in one process, the order changed each round, each build with display false and true.
// First the values of the synthetic data sheet alone, then the synthetic book with its formulas (values computed).
import { createRequire } from 'node:module'
import { hfInput, hfScalar, HF_OPTIONS, CLOCK as clock } from '../hyperformula/shared.mjs'
import { median } from '../stats.mjs'
import { load } from './pair.mjs'
const [base, branch, hfDir] = process.argv.slice(2)
const { HyperFormula } = createRequire(hfDir + '/package.json')('hyperformula')
const libs = await load(base, branch, true)
const letters = libs[0].colLetters
const f = (v) => (v < 10 ? v.toFixed(2) : v.toFixed(0))
const all = (s, sheet, range) => { let cursor, k = 0; do { const p = s.readCells([{ sheet, range }], { fields: ['value'], maxCells: 10000, ...(cursor ? { cursor } : {}) }); k += p.cells.length; cursor = p.cursor } while (cursor); return k }
const box = (sheet, c1, r1, c2, r2) => ({ start: { sheet, col: c1, row: r1 }, end: { sheet, col: c2, row: r2 } })
// a build without `getValues` (a base before it) does not take a case of values
const NOT = Symbol('not measured')
const values = (s, sheet, range, o) => (s.getValues ? s.getValues(sheet, range, o) : NOT)

/** Rows of values against HyperFormula's: the same size, numbers within 1e-9 relatively, an error against an error; the cells that differ. */
function differing(ours, theirs) {
  let n = ours.length === theirs.length ? 0 : Math.abs(ours.length - theirs.length)
  for (let r = 0; r < Math.min(ours.length, theirs.length); r++)
    for (let c = 0; c < Math.max(ours[r].length, theirs[r].length); c++) {
      const [a, b] = [ours[r][c] ?? null, theirs[r][c] ?? null]
      const same = typeof a === 'number' && typeof b === 'number' ? Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(a), Math.abs(b)) : typeof a === 'object' && a !== null ? typeof b === 'object' && b !== null : a === b
      if (!same) n++
    }
  return n
}

/** The sessions of both builds, display false and true, opened on `book`, after HyperFormula's `hf` (null: not measured). */
function engines(book, hf) {
  const out = hf ? [['HyperFormula', hf]] : []
  for (const [name, l] of [['base', libs[0]], ['branch', libs[1]]])
    for (const display of [false, true]) {
      const t = performance.now()
      const s = l.createSession(book, { clock, display })
      s.readCells([{ sheet: book.sheets[0].name, range: 'A1:A2' }], { fields: ['value'] })
      out.push([`${name} display ${display}`, s, performance.now() - t])
    }
  console.log(`opened (with the first read): ${out.filter((e) => e.length > 2).map(([n, , ms]) => `${n} ${f(ms)} ms`).join(', ')}`)
  return out
}

/** Each case: nine rounds, the engines in turn, median; a case without HyperFormula's call is the builds' alone. */
function measure(list, cases) {
  for (const [what, h, x] of cases) {
    const ms = list.map(() => [])
    for (let round = 0; round < 9; round++) {
      const order = list.map((_, i) => i)
      if (round % 2) order.reverse()
      for (const i of order) {
        const [name, s] = list[i]
        if (name === 'HyperFormula' && !h) continue
        const t = performance.now()
        const got = name === 'HyperFormula' ? h(s) : x(s)
        if (got !== NOT) ms[i].push(performance.now() - t)
      }
    }
    // the rows of values each build gives, against HyperFormula's, outside the timer
    const hf = h && list.find(([n]) => n === 'HyperFormula')
    const checked = hf && what.includes('as values') ? list.filter(([n]) => n !== 'HyperFormula').map(([n, s]) => [n, x(s)]).filter(([, v]) => v !== NOT).map(([n, v]) => `${n} ${differing(v, h(hf[1]))}`) : []
    console.log(`${what}: ${list.map(([n], i) => (ms[i].length ? `${n} ${f(median(ms[i]))} ms` : `${n} not measured`)).join(', ')}${checked.length ? ` (cells differing from HyperFormula: ${checked.join(', ')})` : ''}`)
  }
}

// the values of the data sheet alone
{
  const { book } = libs[0].synthBook({ seed: 1, rows: 34000 })
  const out = libs[0].compute(book, { clock }).sheets[0].cells
  const cells = {}
  const grid = []
  let width = 0
  for (const [addr, c] of Object.entries(out)) {
    if (c.value === null || c.value === undefined || typeof c.value === 'object') continue
    cells[addr] = { value: c.value }
    const m = /^([A-Z]+)(\d+)$/.exec(addr)
    let col = 0
    for (const ch of m[1]) col = col * 26 + ch.charCodeAt(0) - 64
    // a text as HyperFormula takes one, after an apostrophe: else '25%' or '001' would be read as a number
    ;(grid[Number(m[2]) - 1] ??= [])[col - 1] = hfScalar({ value: c.value })
    width = Math.max(width, col)
  }
  for (let r = 0; r < grid.length; r++) grid[r] ??= []
  const rows = grid.length
  const at = (i) => [(i * 7) % width, (i * 3391) % rows]
  console.log(`values of the data sheet, ${rows} rows`)
  measure(engines({ sheets: [{ name: 'Data', cells }] }, HyperFormula.buildFromSheets({ Data: grid }, { licenseKey: 'gpl-v3' })), [
    ['window A1:P100', (hf) => hf.getRangeValues(box(0, 0, 0, 15, 99)), (s) => s.readCells([{ sheet: 'Data', range: 'A1:P100' }], { fields: ['value'] })],
    ['window A1:P100, value, text and formula', null, (s) => s.readCells([{ sheet: 'Data', range: 'A1:P100' }])],
    ['column B, its rows', (hf) => hf.getRangeValues(box(0, 1, 0, 1, rows - 1)), (s) => all(s, 'Data', `B1:B${rows}`)],
    ['column B whole', (hf) => hf.getRangeValues(box(0, 1, 0, 1, 1048575)), (s) => all(s, 'Data', 'B:B')],
    ['the sheet by pages of 10,000', (hf) => hf.getSheetValues(0), (s) => all(s, 'Data', `A1:${letters(width)}${rows}`)],
    ['10,000 single cells', (hf) => { for (let i = 0; i < 10000; i++) { const [c, r] = at(i); hf.getCellValue({ sheet: 0, col: c, row: r }) } }, (s) => { for (let i = 0; i < 10000; i++) { const [c, r] = at(i); s.getCell('Data', letters(c + 1) + (r + 1)) } }],
    ['window A1:P100 as values', (hf) => hf.getRangeValues(box(0, 0, 0, 15, 99)), (s) => values(s, 'Data', 'A1:P100')],
    ['column B, its rows, as values', (hf) => hf.getRangeValues(box(0, 1, 0, 1, rows - 1)), (s) => values(s, 'Data', `B1:B${rows}`)],
    ['column B whole as values', (hf) => hf.getRangeValues(box(0, 1, 0, 1, 1048575)), (s) => values(s, 'Data', 'B:B', { maxCells: 1048576 })],
    ['the sheet as values', (hf) => hf.getSheetValues(0), (s) => values(s, 'Data', `A1:${letters(width)}${rows}`)],
  ])
}

// the synthetic book with its formulas: HyperFormula opens 3,400 rows in seconds, 34,000 not in 30 s (bench/hyperformula)
for (const rows of [3400, 34000]) {
  const { book } = libs[0].synthBook({ seed: 1, rows })
  const input = rows <= 3400 ? hfInput(book) : null
  const hf = input && HyperFormula.buildFromSheets(input.sheets, HF_OPTIONS, input.names)
  const [data, sheet] = [book.sheets[0].name, book.sheets.at(-1).name]
  const formulas = book.sheets.slice(1).map((s) => s.name)
  const id = (name) => book.sheets.findIndex((s) => s.name === name)
  const at = (i) => [(i * 7) % 26, (i * 13) % 51]
  console.log(`the synthetic book with formulas, ${rows} rows${hf ? '' : ' (HyperFormula not measured)'}`)
  measure(engines(book, hf), [
    [`window ${sheet}!A1:P50`, (hf) => hf.getRangeValues(box(id(sheet), 0, 0, 15, 49)), (s) => s.readCells([{ sheet, range: 'A1:P50' }], { fields: ['value'] })],
    ['the sheets of formulas by pages of 10,000', (hf) => { for (const n of formulas) hf.getSheetValues(id(n)) }, (s) => { for (const n of formulas) all(s, n, 'A1:Z60') }],
    ['the data sheet by pages of 10,000', (hf) => hf.getSheetValues(0), (s) => all(s, data, `A1:O${rows + 1}`)],
    [`10,000 single cells of ${sheet}`, (hf) => { for (let i = 0; i < 10000; i++) { const [c, r] = at(i); hf.getCellValue({ sheet: id(sheet), col: c, row: r }) } }, (s) => { for (let i = 0; i < 10000; i++) { const [c, r] = at(i); s.getCell(sheet, letters(c + 1) + (r + 1)) } }],
    [`window ${sheet}!A1:P50 as values`, (hf) => hf.getRangeValues(box(id(sheet), 0, 0, 15, 49)), (s) => values(s, sheet, 'A1:P50')],
    ['the data sheet as values', (hf) => hf.getSheetValues(0), (s) => values(s, data, `A1:O${rows + 1}`)],
  ])
}
